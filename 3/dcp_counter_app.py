import os
import re
import subprocess
import sys
from collections import defaultdict
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd


# ==============================================================================
# REGEX PATTERNS & TEXT NORMALIZATION
# ==============================================================================
CPL_TYPE_PATTERN = re.compile(
    r"_(FTR|TSR|SHR|TLR|PSA|ADV|INT|PRO|RTG|POL|XSN)(?:[-_]|$)", re.IGNORECASE
)
VERSION_SUFFIX_PATTERN = re.compile(r"_(OV|VF|CC)(?:[-_(\s]|$)", re.IGNORECASE)
PART_TOKEN_PATTERN = re.compile(
    r"[-_](?:P|PT|PART)[-_]?([1-9])(?=[-_])", re.IGNORECASE
)
LANG_SUB_PATTERN = re.compile(
    r"_(?:S|F|C-\d+)_([A-Z]{2,3})-([A-Z]{2,3})(?:-([A-Z0-9]+))?_",
    re.IGNORECASE,
)


def clean_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val)
    s = re.sub(r"[\u00a0\u2000-\u200f\u202f\u205f\u3000]", " ", s)
    return s.strip()


def is_cpl_filename(text: str) -> bool:
    if not text:
        return False
    return bool(CPL_TYPE_PATTERN.search(text)) and bool(
        VERSION_SUFFIX_PATTERN.search(text)
    )


def classify_package_flags(cpl: str, release_title: str = "", fmt: str = "") -> dict:
    """
    Inspect CPL + metadata to determine:
      - is_adcc: True if ADCC / CCAP accessibility track
      - is_subtitled: True if burned-in/open subtitle version (e.g., TE-EN, TE-AR)
      - sub_code: Detected subtitle code if present
    """
    combined = f"{cpl} {release_title} {fmt}"
    is_adcc = bool(re.search(r"\bADCC\b|-CCAP(?:[-_]|$)", combined, re.IGNORECASE))

    sub_code = ""
    is_subtitled = False

    m = LANG_SUB_PATTERN.search(cpl)
    if m:
        code = m.group(2).upper()
        if code != "XX":
            sub_code = code
            is_subtitled = True

    if not is_subtitled and not is_adcc:
        if re.search(r"\b(subtitles?|subs?)\b", f"{release_title} {fmt}", re.IGNORECASE):
            is_subtitled = True
            sub_code = "META"

    # Note: ADCC packages often carry -TE-EN-CCAP_; treat ADCC separately from standard subtitled DCPs
    if is_adcc:
        is_subtitled = False

    return {
        "is_adcc": is_adcc,
        "is_subtitled": is_subtitled,
        "sub_code": sub_code,
    }


def extract_version_type(cpl: str) -> str:
    m = VERSION_SUFFIX_PATTERN.search(cpl)
    if m:
        return m.group(1).upper()
    return "UNKNOWN"


def normalize_cpl_base(cpl: str) -> str:
    s = re.sub(r"\s*\(.*?\)\s*$", "", cpl).strip()
    s = PART_TOKEN_PATTERN.sub("", s)
    s = re.sub(r"_(?:OV|VF|CC)$", "", s, flags=re.IGNORECASE)
    return s


def normalize_lab_name(raw_lab: str, is_trailer: bool = False) -> str:
    lab = clean_str(raw_lab)
    if not lab or lab.lower() in ("nan", "none", "na", "n/a", "-"):
        if is_trailer:
            return "FLAGGED - Missing Lab (Trailers Section)"
        return "FLAGGED - Missing/Unclear Lab"
    if re.search(r"\bqube\b", lab, re.IGNORECASE):
        return "Qube Lab"
    return lab


# ==============================================================================
# CORE PARSING & COUNTING ENGINE
# ==============================================================================
def process_mastered_sheet(
    filepath: str,
    sheet_name: str,
    include_trailers: bool = False,
    only_trailers: bool = False,
    include_adcc: bool = True,
    include_subtitles: bool = False,
) -> dict:
    df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)

    header_map = {}
    in_trailers = False
    in_ingesting = False
    current_block = None
    block_counter = 0

    audit_rows = []
    primary_entries = []
    seen_exact_cpls_in_block = set()

    for r_idx in range(len(df)):
        excel_row_num = r_idx + 1
        row_vals = [clean_str(df.iat[r_idx, c]) for c in range(df.shape[1])]
        if not any(row_vals):
            continue

        # Detect TRAILERS section banner
        if any(v.upper() == "TRAILERS" for v in row_vals):
            in_trailers = True
            in_ingesting = False
            current_block = None
            audit_rows.append(
                {
                    "row_num": excel_row_num,
                    "section": "Trailers Banner",
                    "raw_text": "TRAILERS",
                    "norm_pkg": "-",
                    "lab": "-",
                    "ver_type": "Non-Package",
                    "status": "Excluded",
                    "reason": "Section Banner",
                }
            )
            continue

        # Detect Table Header Row (Main or Trailers)
        upper_vals = [v.upper() for v in row_vals]
        if ("MOVIE" in upper_vals or "MOVIE NAME" in upper_vals) and (
            "LANGUAGE" in upper_vals or "FORMAT" in upper_vals
        ):
            header_map = {}
            for c_idx, v in enumerate(upper_vals):
                if v in ("SNO", "S:NO", "S.NO", "SL NO", "SL.NO"):
                    header_map["sno"] = c_idx
                elif "DATE" in v:
                    header_map["date"] = c_idx
                elif v in ("MOVIE", "MOVIE NAME", "PACKAGE", "CPL"):
                    header_map["movie"] = c_idx
                elif "LANG" in v:
                    header_map["lang"] = c_idx
                elif "FORMAT" in v:
                    header_map["format"] = c_idx
                elif "MRN" in v:
                    header_map["mrn"] = c_idx
                elif "LAB" in v:
                    header_map["lab"] = c_idx
            in_ingesting = False
            audit_rows.append(
                {
                    "row_num": excel_row_num,
                    "section": "Trailers Header" if in_trailers else "Main Header",
                    "raw_text": " | ".join([v for v in row_vals if v]),
                    "norm_pkg": "-",
                    "lab": "-",
                    "ver_type": "Non-Package",
                    "status": "Excluded",
                    "reason": "Table Header Row",
                }
            )
            continue

        if not header_map or "movie" not in header_map:
            audit_rows.append(
                {
                    "row_num": excel_row_num,
                    "section": "Top Metadata",
                    "raw_text": " | ".join([v for v in row_vals if v]),
                    "norm_pkg": "-",
                    "lab": "-",
                    "ver_type": "Non-Package",
                    "status": "Excluded",
                    "reason": "Title / Metadata Row",
                }
            )
            continue

        m_col = header_map["movie"]
        movie_cell = row_vals[m_col] if m_col < len(row_vals) else ""
        date_cell = (
            row_vals[header_map["date"]]
            if "date" in header_map and header_map["date"] < len(row_vals)
            else ""
        )
        sno_cell = (
            row_vals[header_map["sno"]]
            if "sno" in header_map and header_map["sno"] < len(row_vals)
            else ""
        )
        lab_cell = (
            row_vals[header_map["lab"]]
            if "lab" in header_map and header_map["lab"] < len(row_vals)
            else ""
        )
        fmt_cell = (
            row_vals[header_map["format"]]
            if "format" in header_map and header_map["format"] < len(row_vals)
            else ""
        )

        if not movie_cell:
            continue

        if re.search(r"ingesting\s*details", movie_cell, re.IGNORECASE):
            in_ingesting = True
            audit_rows.append(
                {
                    "row_num": excel_row_num,
                    "section": current_block["title"] if current_block else "Main",
                    "raw_text": movie_cell,
                    "norm_pkg": "-",
                    "lab": current_block["lab"] if current_block else "-",
                    "ver_type": "Non-Package",
                    "status": "Excluded",
                    "reason": "Ingesting Details Section Header",
                }
            )
            continue

        # ======================================================================
        # MAIN FEATURES SECTION
        # ======================================================================
        if not in_trailers:
            if (date_cell or sno_cell or lab_cell) and not is_cpl_filename(movie_cell):
                block_counter += 1
                in_ingesting = False
                seen_exact_cpls_in_block = set()
                lab_norm = normalize_lab_name(lab_cell, is_trailer=False)
                current_block = {
                    "block_id": block_counter,
                    "title": movie_cell,
                    "lab": lab_norm,
                    "format": fmt_cell,
                    "is_trailer": False,
                }
                audit_rows.append(
                    {
                        "row_num": excel_row_num,
                        "section": f"Block #{block_counter} Header",
                        "raw_text": movie_cell,
                        "norm_pkg": movie_cell,
                        "lab": lab_norm,
                        "ver_type": "Non-Package",
                        "status": "Excluded",
                        "reason": "Release Title Header Row (Context Captured)",
                    }
                )
                continue

            if is_cpl_filename(movie_cell):
                ver = extract_version_type(movie_cell)
                norm_base = normalize_cpl_base(movie_cell)
                ctx_title = current_block["title"] if current_block else norm_base
                ctx_fmt = current_block["format"] if current_block else ""
                ctx_lab = (
                    current_block["lab"]
                    if current_block
                    else "FLAGGED - Missing/Unclear Lab"
                )

                if only_trailers:
                    audit_rows.append(
                        {
                            "row_num": excel_row_num,
                            "section": ctx_title,
                            "raw_text": movie_cell,
                            "norm_pkg": norm_base,
                            "lab": ctx_lab,
                            "ver_type": ver,
                            "status": "Excluded",
                            "reason": "Main Feature Excluded ('Only Trailers' Mode Active)",
                        }
                    )
                    continue

                if in_ingesting:
                    audit_rows.append(
                        {
                            "row_num": excel_row_num,
                            "section": ctx_title,
                            "raw_text": movie_cell,
                            "norm_pkg": norm_base,
                            "lab": ctx_lab,
                            "ver_type": ver,
                            "status": "Excluded",
                            "reason": "Ingesting Details / Step Instruction",
                        }
                    )
                    continue

                flags = classify_package_flags(movie_cell, ctx_title, ctx_fmt)

                if flags["is_adcc"] and not include_adcc:
                    audit_rows.append(
                        {
                            "row_num": excel_row_num,
                            "section": ctx_title,
                            "raw_text": movie_cell,
                            "norm_pkg": norm_base,
                            "lab": ctx_lab,
                            "ver_type": f"{ver} (ADCC/CCAP)",
                            "status": "Excluded",
                            "reason": "ADCC / CCAP Package Excluded (Checkbox Unchecked)",
                        }
                    )
                    continue

                if flags["is_subtitled"] and not include_subtitles:
                    audit_rows.append(
                        {
                            "row_num": excel_row_num,
                            "section": ctx_title,
                            "raw_text": movie_cell,
                            "norm_pkg": norm_base,
                            "lab": ctx_lab,
                            "ver_type": f"{ver} (Subtitled)",
                            "status": "Excluded",
                            "reason": f"Subtitled Version Excluded ({flags['sub_code']} — Checkbox Unchecked)",
                        }
                    )
                    continue

                dup_key = (current_block["block_id"] if current_block else 0, movie_cell)
                if dup_key in seen_exact_cpls_in_block:
                    audit_rows.append(
                        {
                            "row_num": excel_row_num,
                            "section": ctx_title,
                            "raw_text": movie_cell,
                            "norm_pkg": norm_base,
                            "lab": ctx_lab,
                            "ver_type": ver,
                            "status": "Excluded",
                            "reason": "Duplicate Entry",
                        }
                    )
                    continue
                seen_exact_cpls_in_block.add(dup_key)

                group_id = (
                    f"BLK_{current_block['block_id']}"
                    if current_block
                    else f"ORPHAN_{norm_base}_{ctx_lab}"
                )
                primary_entries.append(
                    {
                        "audit_idx": len(audit_rows),
                        "row_num": excel_row_num,
                        "group_id": group_id,
                        "movie_title": ctx_title,
                        "norm_base": norm_base,
                        "raw_cpl": movie_cell,
                        "lab": ctx_lab,
                        "ver_type": ver,
                        "is_trailer": False,
                    }
                )
                audit_rows.append(
                    {
                        "row_num": excel_row_num,
                        "section": ctx_title,
                        "raw_text": movie_cell,
                        "norm_pkg": norm_base,
                        "lab": ctx_lab,
                        "ver_type": ver,
                        "status": "PENDING",
                        "reason": "",
                    }
                )
            else:
                audit_rows.append(
                    {
                        "row_num": excel_row_num,
                        "section": current_block["title"] if current_block else "Main",
                        "raw_text": movie_cell,
                        "norm_pkg": "-",
                        "lab": current_block["lab"] if current_block else "-",
                        "ver_type": "Non-Package",
                        "status": "Excluded",
                        "reason": "Instruction / Note Row",
                    }
                )

        # ======================================================================
        # TRAILERS SUB-TABLE SECTION
        # ======================================================================
        else:
            if is_cpl_filename(movie_cell):
                ver = extract_version_type(movie_cell)
                norm_base = normalize_cpl_base(movie_cell)
                lab_norm = normalize_lab_name(lab_cell, is_trailer=True)
                trailer_title = f"{movie_cell.split('_')[0]} (Trailer)"

                if not (include_trailers or only_trailers):
                    audit_rows.append(
                        {
                            "row_num": excel_row_num,
                            "section": "Trailers",
                            "raw_text": movie_cell,
                            "norm_pkg": norm_base,
                            "lab": lab_norm,
                            "ver_type": ver,
                            "status": "Excluded",
                            "reason": "Trailers Section Excluded (Checkbox Unchecked)",
                        }
                    )
                    continue

                flags = classify_package_flags(movie_cell, trailer_title, fmt_cell)

                if flags["is_adcc"] and not include_adcc:
                    audit_rows.append(
                        {
                            "row_num": excel_row_num,
                            "section": "Trailers",
                            "raw_text": movie_cell,
                            "norm_pkg": norm_base,
                            "lab": lab_norm,
                            "ver_type": f"{ver} (ADCC/CCAP)",
                            "status": "Excluded",
                            "reason": "ADCC / CCAP Package Excluded (Checkbox Unchecked)",
                        }
                    )
                    continue

                if flags["is_subtitled"] and not include_subtitles:
                    audit_rows.append(
                        {
                            "row_num": excel_row_num,
                            "section": "Trailers",
                            "raw_text": movie_cell,
                            "norm_pkg": norm_base,
                            "lab": lab_norm,
                            "ver_type": f"{ver} (Subtitled)",
                            "status": "Excluded",
                            "reason": f"Subtitled Version Excluded ({flags['sub_code']} — Checkbox Unchecked)",
                        }
                    )
                    continue

                block_counter += 1
                group_id = f"TRL_{norm_base}_{lab_norm}"
                primary_entries.append(
                    {
                        "audit_idx": len(audit_rows),
                        "row_num": excel_row_num,
                        "group_id": group_id,
                        "movie_title": trailer_title,
                        "norm_base": norm_base,
                        "raw_cpl": movie_cell,
                        "lab": lab_norm,
                        "ver_type": ver,
                        "is_trailer": True,
                    }
                )
                audit_rows.append(
                    {
                        "row_num": excel_row_num,
                        "section": "Trailers",
                        "raw_text": movie_cell,
                        "norm_pkg": norm_base,
                        "lab": lab_norm,
                        "ver_type": ver,
                        "status": "PENDING",
                        "reason": "",
                    }
                )
            else:
                audit_rows.append(
                    {
                        "row_num": excel_row_num,
                        "section": "Trailers",
                        "raw_text": movie_cell,
                        "norm_pkg": "-",
                        "lab": "-",
                        "ver_type": "Non-Package",
                        "status": "Excluded",
                        "reason": "Non-Package Row",
                    }
                )

    # ==========================================================================
    # GROUPING, P1+P2 CONSOLIDATION & VF-OVER-OV PRIORITY
    # ==========================================================================
    grouped_blocks = defaultdict(list)
    for item in primary_entries:
        grouped_blocks[item["group_id"]].append(item)

    seen_canonical_signatures = {}
    counted_packages = []

    for gid, items in grouped_blocks.items():
        first = items[0]
        lab = first["lab"]
        title = first["movie_title"]
        norm_bases = sorted(set(x["norm_base"] for x in items))
        canonical_sig = (tuple(norm_bases), lab)

        if canonical_sig in seen_canonical_signatures:
            for it in items:
                audit_rows[it["audit_idx"]]["status"] = "Excluded"
                audit_rows[it["audit_idx"]]["reason"] = (
                    f"Duplicate Package Signature (Already counted in {seen_canonical_signatures[canonical_sig]})"
                )
            continue

        versions_present = [x["ver_type"] for x in items]
        chosen_version = "VF" if "VF" in versions_present else "OV"
        has_both_ov_and_vf = ("OV" in versions_present) and ("VF" in versions_present)
        is_multi_part = len(items) > 1

        seen_canonical_signatures[canonical_sig] = title

        included_cpls = []
        for it in items:
            if has_both_ov_and_vf and it["ver_type"] == "OV":
                audit_rows[it["audit_idx"]]["status"] = "Excluded"
                audit_rows[it["audit_idx"]]["reason"] = (
                    "VF Priority Override (Paired VF counted for this package)"
                )
            else:
                included_cpls.append(it["raw_cpl"])
                audit_rows[it["audit_idx"]]["status"] = "Included"
                if is_multi_part and len(included_cpls) > 1:
                    audit_rows[it["audit_idx"]]["reason"] = (
                        f"Counted as {chosen_version} (Merged P1+P2 into 1 package count)"
                    )
                elif is_multi_part:
                    audit_rows[it["audit_idx"]]["reason"] = (
                        f"Counted as {chosen_version} (Primary part of P1+P2 single count)"
                    )
                else:
                    audit_rows[it["audit_idx"]]["reason"] = (
                        f"Counted as {chosen_version}"
                    )

        counted_packages.append(
            {
                "movie_name": title,
                "norm_base": " / ".join(norm_bases),
                "lab": lab,
                "version": chosen_version,
                "ov_count": 1 if chosen_version == "OV" else 0,
                "vf_count": 1 if chosen_version == "VF" else 0,
                "total_count": 1,
                "package_ids": "\n".join(included_cpls),
                "is_trailer": first["is_trailer"],
            }
        )

    # ==========================================================================
    # AGGREGATION & VALIDATION CHECKS
    # ==========================================================================
    total_ov = sum(p["ov_count"] for p in counted_packages)
    total_vf = sum(p["vf_count"] for p in counted_packages)
    grand_total = len(counted_packages)

    qube_items = [p for p in counted_packages if p["lab"] == "Qube Lab"]
    other_items = [
        p
        for p in counted_packages
        if p["lab"] != "Qube Lab" and not p["lab"].startswith("FLAGGED")
    ]
    flagged_items = [p for p in counted_packages if p["lab"].startswith("FLAGGED")]

    qube_ov = sum(p["ov_count"] for p in qube_items)
    qube_vf = sum(p["vf_count"] for p in qube_items)
    qube_total = len(qube_items)

    other_ov = sum(p["ov_count"] for p in other_items)
    other_vf = sum(p["vf_count"] for p in other_items)
    other_total = len(other_items)

    flagged_ov = sum(p["ov_count"] for p in flagged_items)
    flagged_vf = sum(p["vf_count"] for p in flagged_items)
    flagged_total = len(flagged_items)

    lab_breakdown = defaultdict(lambda: {"OV": 0, "VF": 0, "Total": 0})
    for p in counted_packages:
        lab_breakdown[p["lab"]]["OV"] += p["ov_count"]
        lab_breakdown[p["lab"]]["VF"] += p["vf_count"]
        lab_breakdown[p["lab"]]["Total"] += 1

    validations = []
    v1 = (total_ov + total_vf) == grand_total
    validations.append(
        ("OV + VF == Grand Total", v1, f"{total_ov} + {total_vf} = {grand_total}")
    )

    v2 = (qube_total + other_total + flagged_total) == grand_total
    validations.append(
        (
            "Qube + Other Labs + Flagged == Grand Total",
            v2,
            f"{qube_total} + {other_total} + {flagged_total} = {grand_total}",
        )
    )

    v3 = all(p["version"] in ("OV", "VF") for p in counted_packages)
    validations.append(("Filter & Version Integrity", v3, "Verified cleanly"))

    v4 = len(set((p["norm_base"], p["lab"]) for p in counted_packages)) == grand_total
    validations.append(
        (
            "Zero Duplicate Packages / P1+P2 Double-Counts",
            v4,
            f"{grand_total} unique counted packages",
        )
    )

    return {
        "total_ov": total_ov,
        "total_vf": total_vf,
        "grand_total": grand_total,
        "qube_ov": qube_ov,
        "qube_vf": qube_vf,
        "qube_total": qube_total,
        "other_ov": other_ov,
        "other_vf": other_vf,
        "other_total": other_total,
        "flagged_ov": flagged_ov,
        "flagged_vf": flagged_vf,
        "flagged_total": flagged_total,
        "lab_breakdown": dict(lab_breakdown),
        "counted_packages": counted_packages,
        "qube_items": qube_items,
        "other_items": other_items,
        "flagged_items": flagged_items,
        "audit_rows": audit_rows,
        "validations": validations,
        "all_passed": all(v[1] for v in validations),
        "mode_label": (
            "TRAILERS ONLY"
            if only_trailers
            else ("FEATURES + TRAILERS" if include_trailers else "FEATURES ONLY")
        ),
        "active_filters": {
            "include_trailers": include_trailers,
            "only_trailers": only_trailers,
            "include_adcc": include_adcc,
            "include_subtitles": include_subtitles,
        },
    }


# ==============================================================================
# EXCEL WORKBOOK GENERATOR (7 FORMATTED SHEETS)
# ==============================================================================
def style_sheet_table(ws, header_row: int, max_row: int, max_col: int, wrap_cols=None):
    wrap_cols = wrap_cols or set()
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=10)
    thin_side = Side(border_style="thin", color="D9D9D9")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for c in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    ws.row_dimensions[header_row].height = 26

    for r in range(header_row + 1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = body_font
            cell.border = border
            is_num = isinstance(cell.value, (int, float))
            cell.alignment = Alignment(
                horizontal="center" if is_num else "left",
                vertical="center",
                wrap_text=(c in wrap_cols),
            )

    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max_col)}{max_row}"
    ws.freeze_panes = f"A{header_row + 1}"

    for c in range(1, max_col + 1):
        col_letter = get_column_letter(c)
        max_len = 0
        for r in range(header_row, max_row + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None:
                lines = str(val).split("\n")
                longest = max(len(line) for line in lines)
                if longest > max_len:
                    max_len = longest
        ws.column_dimensions[col_letter].width = min(max(max_len + 5, 14), 65)


def export_excel_report(results: dict, output_path: str, source_filename: str):
    wb = openpyxl.Workbook()
    bold_font = Font(name="Calibri", size=11, bold=True)
    title_font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    flag_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    pass_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    excl_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    f = results["active_filters"]
    filter_summary = (
        f"Mode: {results['mode_label']} | "
        f"ADCC/CCAP: {'Included' if f['include_adcc'] else 'Excluded'} | "
        f"Subtitled Versions: {'Included' if f['include_subtitles'] else 'Excluded'}"
    )

    # SHEET 1: Overall Count
    ws1 = wb.active
    ws1.title = "Overall Count"
    ws1["A1"] = f"DCP OV & VF Overall Count Report ({results['mode_label']})"
    ws1["A1"].font = title_font
    ws1["A2"] = (
        f"Source: {source_filename} | {filter_summary} | "
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    ws1["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")

    ws1.append([])
    ws1.append(["Metric", "Package Count"])
    ws1.append(["Total OV Count", results["total_ov"]])
    ws1.append(["Total VF Count", results["total_vf"]])
    ws1.append(["Grand Total (OV + VF)", results["grand_total"]])
    ws1.append(["Qube Lab Total", results["qube_total"]])
    ws1.append(["Other Labs Total (Verified External Labs)", results["other_total"]])
    if results["flagged_total"] > 0:
        ws1.append(["Flagged / Missing Lab Total", results["flagged_total"]])

    style_sheet_table(ws1, header_row=4, max_row=ws1.max_row, max_col=2)
    ws1.cell(row=7, column=1).font = bold_font
    ws1.cell(row=7, column=2).font = bold_font
    ws1.cell(row=7, column=1).fill = total_fill
    ws1.cell(row=7, column=2).fill = total_fill

    v_start = ws1.max_row + 2
    ws1.cell(row=v_start, column=1, value="Automated Validation Checks").font = bold_font
    for idx, (label, passed, detail) in enumerate(results["validations"], start=1):
        r = v_start + idx
        ws1.cell(row=r, column=1, value=label)
        c2 = ws1.cell(row=r, column=2, value=f"{'PASSED' if passed else 'FAILED'} ({detail})")
        c2.fill = pass_fill if passed else excl_fill
        c2.font = bold_font

    ws1.column_dimensions["A"].width = 44
    ws1.column_dimensions["B"].width = 36

    # SHEET 2: Lab-wise Overall
    ws2 = wb.create_sheet("Lab-wise Overall")
    ws2["A1"] = "Lab-wise OV & VF Breakdown"
    ws2["A1"].font = title_font
    ws2.append([])
    ws2.append(["Lab Category", "Lab Name", "OV Count", "VF Count", "Grand Total"])

    q_data = results["lab_breakdown"].get("Qube Lab", {"OV": 0, "VF": 0, "Total": 0})
    ws2.append(["Qube Lab", "Qube Lab", q_data["OV"], q_data["VF"], q_data["Total"]])

    other_labs = sorted(
        k
        for k in results["lab_breakdown"].keys()
        if k != "Qube Lab" and not k.startswith("FLAGGED")
    )
    for lab_name in other_labs:
        d = results["lab_breakdown"][lab_name]
        ws2.append(["Other Labs", lab_name, d["OV"], d["VF"], d["Total"]])

    flagged_labs = sorted(
        k for k in results["lab_breakdown"].keys() if k.startswith("FLAGGED")
    )
    for lab_name in flagged_labs:
        d = results["lab_breakdown"][lab_name]
        ws2.append(["Flagged / Unclear", lab_name, d["OV"], d["VF"], d["Total"]])

    ws2.append(
        [
            "ALL LABS",
            "Overall Grand Total",
            results["total_ov"],
            results["total_vf"],
            results["grand_total"],
        ]
    )

    style_sheet_table(ws2, header_row=3, max_row=ws2.max_row, max_col=5)
    for r in range(4, ws2.max_row):
        if str(ws2.cell(row=r, column=2).value).startswith("FLAGGED"):
            for c in range(1, 6):
                ws2.cell(row=r, column=c).fill = flag_fill
    for c in range(1, 6):
        ws2.cell(row=ws2.max_row, column=c).font = bold_font
        ws2.cell(row=ws2.max_row, column=c).fill = total_fill

    # SHEET 3: Movie-wise Count
    ws3 = wb.create_sheet("Movie-wise Count")
    ws3["A1"] = "Movie / Package-wise Count"
    ws3["A1"].font = title_font
    ws3.append([])
    ws3.append(
        [
            "Movie / Package Name",
            "Canonical CPL Base",
            "Lab Name",
            "OV Count",
            "VF Count",
            "Total Count",
        ]
    )

    for pkg in results["counted_packages"]:
        ws3.append(
            [
                pkg["movie_name"],
                pkg["norm_base"],
                pkg["lab"],
                pkg["ov_count"],
                pkg["vf_count"],
                pkg["total_count"],
            ]
        )

    ws3.append(
        [
            "Grand Total",
            "",
            "",
            results["total_ov"],
            results["total_vf"],
            results["grand_total"],
        ]
    )
    style_sheet_table(ws3, header_row=3, max_row=ws3.max_row, max_col=6)
    for c in range(1, 7):
        ws3.cell(row=ws3.max_row, column=c).font = bold_font
        ws3.cell(row=ws3.max_row, column=c).fill = total_fill

    # SHEET 4: Qube Lab Details
    ws4 = wb.create_sheet("Qube Lab Details")
    ws4["A1"] = "Qube Lab Detailed Package Breakdown"
    ws4["A1"].font = title_font
    ws4.append([])
    ws4.append(
        [
            "Movie / Package Name",
            "Counted Version (OV/VF)",
            "Package Identification (Counted CPLs)",
            "Individual Count",
        ]
    )

    for pkg in results["qube_items"]:
        ws4.append(
            [pkg["movie_name"], pkg["version"], pkg["package_ids"], pkg["total_count"]]
        )

    ws4.append(["Total Qube Lab Packages", "", "", results["qube_total"]])
    style_sheet_table(
        ws4, header_row=3, max_row=ws4.max_row, max_col=4, wrap_cols={3}
    )
    for r in range(4, ws4.max_row):
        if "\n" in str(ws4.cell(row=r, column=3).value):
            ws4.row_dimensions[r].height = 30
    for c in range(1, 5):
        ws4.cell(row=ws4.max_row, column=c).font = bold_font
        ws4.cell(row=ws4.max_row, column=c).fill = total_fill

    # SHEET 5: Other Labs Details
    ws5 = wb.create_sheet("Other Labs Details")
    ws5["A1"] = "Other Labs (& Flagged) Detailed Package Breakdown"
    ws5["A1"].font = title_font
    ws5.append([])
    ws5.append(
        [
            "Lab Name",
            "Movie / Package Name",
            "Counted Version (OV/VF)",
            "Package Identification (Counted CPLs)",
            "Individual Count",
        ]
    )

    combined_non_qube = results["other_items"] + results["flagged_items"]
    for pkg in combined_non_qube:
        ws5.append(
            [
                pkg["lab"],
                pkg["movie_name"],
                pkg["version"],
                pkg["package_ids"],
                pkg["total_count"],
            ]
        )

    ws5.append(
        [
            "Total Non-Qube Packages",
            "",
            "",
            "",
            results["other_total"] + results["flagged_total"],
        ]
    )
    style_sheet_table(
        ws5, header_row=3, max_row=ws5.max_row, max_col=5, wrap_cols={4}
    )
    for r in range(4, ws5.max_row):
        if "\n" in str(ws5.cell(row=r, column=4).value):
            ws5.row_dimensions[r].height = 30
        if str(ws5.cell(row=r, column=1).value).startswith("FLAGGED"):
            for c in range(1, 6):
                ws5.cell(row=r, column=c).fill = flag_fill
    for c in range(1, 6):
        ws5.cell(row=ws5.max_row, column=c).font = bold_font
        ws5.cell(row=ws5.max_row, column=c).fill = total_fill

    # SHEET 6: Filename Audit
    ws6 = wb.create_sheet("Filename Audit")
    ws6["A1"] = "Complete Row-by-Row Filename & Exclusion Audit"
    ws6["A1"].font = title_font
    ws6.append([])
    ws6.append(
        [
            "Excel Row #",
            "Section / Release Context",
            "Original Filename / Cell Text",
            "Normalized Package Base",
            "Lab Name",
            "Identified Version",
            "Status",
            "Reason for Inclusion / Exclusion",
        ]
    )

    for a in results["audit_rows"]:
        ws6.append(
            [
                a["row_num"],
                a["section"],
                a["raw_text"],
                a["norm_pkg"],
                a["lab"],
                a["ver_type"],
                a["status"],
                a["reason"],
            ]
        )

    style_sheet_table(ws6, header_row=3, max_row=ws6.max_row, max_col=8)
    for r in range(4, ws6.max_row + 1):
        status_val = ws6.cell(row=r, column=7).value
        ws6.cell(row=r, column=7).fill = (
            pass_fill if status_val == "Included" else excl_fill
        )

    # SHEET 7: Counting Rules
    ws7 = wb.create_sheet("Counting Rules")
    ws7["A1"] = "DCP OV & VF Counting and Classification Rules"
    ws7["A1"].font = title_font
    ws7.append([])
    ws7.append(["Rule Category", "Rule Description & Implementation Logic"])

    rules_doc = [
        (
            "1. Active Extraction Mode",
            f"Run executed in '{results['mode_label']}' mode. Non-package metadata, headers, notes, and all rows inside 'Ingesting Details :' blocks are excluded.",
        ),
        (
            "2. ADCC / CCAP & Subtitle Filters",
            f"ADCC / CCAP accessibility tracks (-CCAP_): {'INCLUDED' if f['include_adcc'] else 'EXCLUDED'}. "
            f"Subtitled versions (ISDCF subtitle code != 'XX', e.g., TE-EN, TE-AR): {'INCLUDED' if f['include_subtitles'] else 'EXCLUDED'}.",
        ),
        (
            "3. P1 & P2 Multi-Part Consolidation",
            "Part 1 (-P1_ / _P1_) and Part 2 (-P2_ / _P2_) of the same mastered version are normalized to a single canonical package and counted as ONE (1) package, never two.",
        ),
        (
            "4. OV vs. VF Priority Rule",
            "If both OV and VF exist for the same mastered package (including mixed pairs where P1 is OV and P2 is VF), ONLY VF is counted (Count = 1) and the corresponding OV is excluded.",
        ),
        (
            "5. Duplicate Prevention",
            "Exact duplicate CPL rows within a block and duplicate package signatures across blocks are deduplicated so no package is ever counted twice.",
        ),
        (
            "6. Lab-wise Classification",
            "Packages assigned to QUBE / Qube Cinema are grouped under 'Qube Lab'. All external labs are preserved under their exact individual lab names. Missing Labs are explicitly flagged.",
        ),
    ]
    for cat, desc in rules_doc:
        ws7.append([cat, desc])

    style_sheet_table(
        ws7, header_row=3, max_row=ws7.max_row, max_col=2, wrap_cols={2}
    )
    ws7.column_dimensions["A"].width = 34
    ws7.column_dimensions["B"].width = 95
    for r in range(4, ws7.max_row + 1):
        ws7.row_dimensions[r].height = 38

    wb.save(output_path)


# ==============================================================================
# PORTABLE WINDOWS GUI APPLICATION (TKINTER)
# ==============================================================================
class DCPCounterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DCP OV & VF Mastered List Counter")
        self.root.geometry("800x650")
        self.root.minsize(740, 580)

        self.filepath = tk.StringVar()
        self.selected_sheet = tk.StringVar()

        # Filter Checkboxes (Sensible defaults matching main feature workflow)
        self.include_trailers = tk.BooleanVar(value=False)
        self.only_trailers = tk.BooleanVar(value=False)
        self.include_adcc = tk.BooleanVar(value=True)
        self.include_subtitles = tk.BooleanVar(value=False)

        self.last_output_path = None
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 14, "pady": 6}

        # Title Banner
        header_frame = tk.Frame(self.root, bg="#1F4E78", height=58)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        tk.Label(
            header_frame,
            text="DCP OV & VF Mastered List Analyzer",
            bg="#1F4E78",
            fg="white",
            font=("Segoe UI", 14, "bold"),
        ).pack(side="left", padx=16)

        # 1. File Selection Frame
        file_frame = tk.LabelFrame(
            self.root,
            text="1. Select Mastered List Excel File",
            font=("Segoe UI", 10, "bold"),
        )
        file_frame.pack(fill="x", **pad)

        entry = tk.Entry(
            file_frame, textvariable=self.filepath, font=("Segoe UI", 10), state="readonly"
        )
        entry.pack(side="left", fill="x", expand=True, padx=(10, 6), pady=10)

        tk.Button(
            file_frame,
            text="Browse Excel...",
            command=self.browse_file,
            font=("Segoe UI", 10),
            width=15,
        ).pack(side="right", padx=(0, 10), pady=10)

        # 2. Worksheet & Filter Options Frame
        opt_frame = tk.LabelFrame(
            self.root,
            text="2. Worksheet & Extraction Filter Options",
            font=("Segoe UI", 10, "bold"),
        )
        opt_frame.pack(fill="x", **pad)

        # Row 1: Worksheet Dropdown
        row1 = tk.Frame(opt_frame)
        row1.pack(fill="x", padx=10, pady=(8, 4))
        tk.Label(row1, text="Worksheet:", font=("Segoe UI", 10)).pack(
            side="left", padx=(0, 6)
        )
        self.sheet_combo = ttk.Combobox(
            row1,
            textvariable=self.selected_sheet,
            state="readonly",
            width=28,
            font=("Segoe UI", 10),
        )
        self.sheet_combo.pack(side="left")

        # Row 2: Trailers Options
        row2 = tk.Frame(opt_frame)
        row2.pack(fill="x", padx=10, pady=3)

        self.chk_inc_trailers = tk.Checkbutton(
            row2,
            text="Include TRAILERS section along with Features",
            variable=self.include_trailers,
            command=self._on_toggle_include_trailers,
            font=("Segoe UI", 10),
        )
        self.chk_inc_trailers.pack(side="left", padx=(0, 20))

        self.chk_only_trailers = tk.Checkbutton(
            row2,
            text="Extract ONLY Trailers separately (ignore Features)",
            variable=self.only_trailers,
            command=self._on_toggle_only_trailers,
            font=("Segoe UI", 10, "bold"),
            fg="#1F4E78",
        )
        self.chk_only_trailers.pack(side="left")

        # Row 3: ADCC & Subtitles Options
        row3 = tk.Frame(opt_frame)
        row3.pack(fill="x", padx=10, pady=(3, 8))

        tk.Checkbutton(
            row3,
            text="Include ADCC / CCAP packages (-CCAP_)",
            variable=self.include_adcc,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=(0, 24))

        tk.Checkbutton(
            row3,
            text="Include Subtitled DCP versions (EN / AR / MS etc.)",
            variable=self.include_subtitles,
            font=("Segoe UI", 10),
        ).pack(side="left")

        # Action Buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill="x", padx=14, pady=8)

        self.run_btn = tk.Button(
            btn_frame,
            text="Generate 7-Sheet Excel Report",
            command=self.run_analysis,
            bg="#1F4E78",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            height=2,
            width=28,
            state="disabled",
        )
        self.run_btn.pack(side="left")

        self.open_folder_btn = tk.Button(
            btn_frame,
            text="Open Output Folder",
            command=self.open_output_folder,
            font=("Segoe UI", 10),
            height=2,
            width=20,
            state="disabled",
        )
        self.open_folder_btn.pack(side="right")

        # 3. Live Summary & Validation Log
        log_frame = tk.LabelFrame(
            self.root, text="3. Summary & Validation Log", font=("Segoe UI", 10, "bold")
        )
        log_frame.pack(fill="both", expand=True, padx=14, pady=(4, 14))

        self.log_text = tk.Text(
            log_frame, font=("Consolas", 10), wrap="word", state="disabled", bg="#FAFAFA"
        )
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True, padx=6, pady=6)

        self.write_log("Ready. Click 'Browse Excel...' to load a Hyderabad Mastered List workbook.")

    def _on_toggle_only_trailers(self):
        if self.only_trailers.get():
            self.include_trailers.set(True)
            self.chk_inc_trailers.configure(state="disabled")
        else:
            self.chk_inc_trailers.configure(state="normal")

    def _on_toggle_include_trailers(self):
        if not self.include_trailers.get():
            self.only_trailers.set(False)

    def write_log(self, msg: str, clear: bool = False):
        self.log_text.configure(state="normal")
        if clear:
            self.log_text.delete("1.0", tk.END)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Hyderabad Mastered List Excel File",
            filetypes=[("Excel Workbooks", "*.xlsx *.xls")],
        )
        if not path:
            return
        try:
            xls = pd.ExcelFile(path)
            sheets = xls.sheet_names
            if not sheets:
                messagebox.showerror("Error", "No worksheets found in the selected file.")
                return
            self.filepath.set(path)
            self.sheet_combo["values"] = sheets
            self.selected_sheet.set(sheets[0])
            self.run_btn.configure(state="normal")
            self.open_folder_btn.configure(state="disabled")
            self.write_log(
                f"Loaded: {os.path.basename(path)}\nAvailable sheets: {', '.join(sheets)}",
                clear=True,
            )
        except Exception as exc:
            messagebox.showerror("File Read Error", f"Could not open Excel file:\n{exc}")

    def run_analysis(self):
        in_path = self.filepath.get()
        sheet = self.selected_sheet.get()
        if not in_path or not sheet:
            return

        try:
            self.write_log(f"Processing '{os.path.basename(in_path)}' [{sheet}]...", clear=True)
            results = process_mastered_sheet(
                in_path,
                sheet,
                include_trailers=self.include_trailers.get(),
                only_trailers=self.only_trailers.get(),
                include_adcc=self.include_adcc.get(),
                include_subtitles=self.include_subtitles.get(),
            )

            base_name = os.path.splitext(os.path.basename(in_path))[0]
            mode_tag = "TrailersOnly_" if self.only_trailers.get() else ""
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_name = f"DCP_OV_VF_Report_{mode_tag}{base_name}_{timestamp}.xlsx"
            out_path = os.path.join(os.path.dirname(in_path), out_name)

            export_excel_report(results, out_path, os.path.basename(in_path))
            self.last_output_path = out_path
            self.open_folder_btn.configure(state="normal")

            self.write_log("=" * 64)
            self.write_log(f"REPORT GENERATED   : {out_name}")
            self.write_log(f"Extraction Mode    : {results['mode_label']}")
            self.write_log("=" * 64)
            self.write_log(f"Overall OV Count   : {results['total_ov']}")
            self.write_log(f"Overall VF Count   : {results['total_vf']}")
            self.write_log(f"Grand Total        : {results['grand_total']}")
            self.write_log("-" * 64)
            self.write_log(
                f"Qube Lab           : {results['qube_total']}  (OV: {results['qube_ov']}, VF: {results['qube_vf']})"
            )
            self.write_log(
                f"Other Labs (Total) : {results['other_total']}  (OV: {results['other_ov']}, VF: {results['other_vf']})"
            )
            if results["flagged_total"] > 0:
                self.write_log(
                    f"Flagged / No Lab   : {results['flagged_total']}  (OV: {results['flagged_ov']}, VF: {results['flagged_vf']})"
                )
            self.write_log("-" * 64)
            self.write_log("Individual Lab Breakdown:")
            for lab_name, stats in results["lab_breakdown"].items():
                self.write_log(
                    f"  • {lab_name}: Total={stats['Total']} (OV={stats['OV']}, VF={stats['VF']})"
                )
            self.write_log("-" * 64)
            self.write_log("Validation Checks:")
            for label, passed, detail in results["validations"]:
                status_tag = "[PASS]" if passed else "[FAIL]"
                self.write_log(f"  {status_tag} {label} -> {detail}")

            messagebox.showinfo(
                "Report Complete",
                f"7-Sheet DCP Report saved successfully:\n\n{out_path}",
            )
        except Exception as exc:
            self.write_log(f"\n[ERROR] {exc}")
            messagebox.showerror("Processing Error", f"An error occurred:\n{exc}")

    def open_output_folder(self):
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            return
        folder = os.path.dirname(self.last_output_path)
        if sys.platform.startswith("win"):
            subprocess.run(["explorer", "/select,", os.path.normpath(self.last_output_path)])
        else:
            os.startfile(folder) if hasattr(os, "startfile") else None


if __name__ == "__main__":
    root = tk.Tk()
    app = DCPCounterApp(root)
    root.mainloop()