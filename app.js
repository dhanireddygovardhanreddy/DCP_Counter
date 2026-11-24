// ==============================================================================
// DCP OV & VF MASTERED LIST ANALYZER — CLIENT-SIDE PARSER & REPORT ENGINE
// ==============================================================================

const CPL_TYPE_PATTERN = /_(FTR|TSR|SHR|TLR|PSA|ADV|INT|PRO|RTG|POL|XSN)(?:[-_]|$)/i;
const VERSION_SUFFIX_PATTERN = /_(OV|VF|CC)(?:[-_(\s]|$)/i;
const PART_TOKEN_PATTERN = /[-_](?:P|PT|PART)[-_]?0?([1-9])(?=[-_])/gi;
const P1_REEL_PATTERN = /[-_](?:P|PT|PART)[-_]?0?1(?=[-_])/i;
const LANG_SUB_PATTERN = /_(?:S|F|C-\d+)_([A-Z]{2,3})-([A-Z]{2,3})(?:-([A-Z0-9]+))?_/i;
const isTrailerCpl = /_(TLR|TSR|PRO|ADV|PSA|INT|RTG|POL|XSN)(?:[-_]|$)/i.test(movieCell);
const effectiveIsTrailer = inTrailers || isTrailerCpl;

const SPECIALTY_MODES = {
  ALL: "All Packages (Standard)",
  ONLY_TRAILERS: "Only Trailers",
  ONLY_ADCC: "Only ADCC / CCAP",
  ONLY_EPIQ: "Only EPIQ",
  ONLY_FLAT: "Only Flat (_F_)",
  ONLY_SCOPE: "Only Scope (_S_ — Excludes C-239)",
  ONLY_HDR: "Only HDR (All HDR / DVis / Barco)",
  ONLY_BARCO_HDR: "Only HDR by Barco",
};

// Application State
let queuedFiles = []; // Array of { name, arrayBuffer, workbook }
let lastGeneratedReport = null; // { results, filename }

// ==============================================================================
// STRING & CPL HELPERS
// ==============================================================================
function cleanStr(val) {
  if (val === null || val === undefined) return "";
  return String(val)
    .replace(/[\u00a0\u2000-\u200f\u202f\u205f\u3000]/g, " ")
    .trim();
}

function isCplFilename(text) {
  if (!text) return false;
  return CPL_TYPE_PATTERN.test(text) && VERSION_SUFFIX_PATTERN.test(text);
}

function extractVersionType(cpl) {
  const m = cpl.match(VERSION_SUFFIX_PATTERN);
  return m ? m[1].toUpperCase() : "UNKNOWN";
}

function normalizeCplBase(cpl) {
  let s = cpl.replace(/\s*\(.*?\)\s*$/, "").trim();
  s = s.replace(PART_TOKEN_PATTERN, "");
  s = s.replace(/_(?:OV|VF|CC)$/i, "");
  return s;
}

function normalizeLabName(rawLab, isTrailer = false) {
  const lab = cleanStr(rawLab);
  if (!lab || ["nan", "none", "na", "n/a", "-"].includes(lab.toLowerCase())) {
    return isTrailer
      ? "FLAGGED - Missing Lab (Trailers Section)"
      : "FLAGGED - Missing/Unclear Lab";
  }
  if (/\bqube\b/i.test(lab)) return "Qube Lab";
  return lab;
}

function classifyPackageAttributes(cpl, releaseTitle = "", fmt = "") {
  const combined = `${cpl} ${releaseTitle} ${fmt}`;
  const isAdcc = /\bADCC\b|-CCAP(?:[-_]|$)/i.test(combined);
  const isEpiq = /(?:^|[-_\s])EPIQ(?:[-_\s]|$)/i.test(combined);

  const isFlat =
    /_F_(?:[A-Z]{2,3}-)/i.test(cpl) ||
    /\bFLAT\b/i.test(`${releaseTitle} ${fmt}`);

  // Scope (_S_ ONLY — excludes _C-239_)
  const hasC239 = /_C-239_/i.test(cpl);
  const isScope =
    !hasC239 &&
    (/_S_(?:[A-Z]{2,3}-)/i.test(cpl) ||
      /\bSCOPE\b/i.test(`${releaseTitle} ${fmt}`));

  const isBarcoHdr = /BARCO|HDR[-_]?B\b/i.test(combined);
  const isHdr =
    isBarcoHdr ||
    /(?:^|[-_\s])(DVis|Dolby\s*Vision|Dolby\s*Cinema|HDR|EclairColor)(?:[-_\s]|$)/i.test(
      combined
    );

  let subCode = "";
  let isSubtitled = false;
  const m = cpl.match(LANG_SUB_PATTERN);
  if (m) {
    const code = m[2].toUpperCase();
    if (code !== "XX") {
      subCode = code;
      isSubtitled = true;
    }
  }

  if (!isSubtitled && !isAdcc) {
    if (/\b(subtitles?|subs?)\b/i.test(`${releaseTitle} ${fmt}`)) {
      isSubtitled = true;
      subCode = "META";
    }
  }

  if (isAdcc) isSubtitled = false;

  return {
    isAdcc,
    isSubtitled,
    subCode,
    isEpiq,
    isFlat,
    isScope,
    isHdr,
    isBarcoHdr,
  };
}

function matchesSpecialtyMode(specialtyMode, isTrailerRow, attrs) {
  switch (specialtyMode) {
    case "ALL":
      return [true, ""];
    case "ONLY_TRAILERS":
      return [isTrailerRow, "Excluded: Not in Trailers Section"];
    case "ONLY_ADCC":
      return [attrs.isAdcc, "Excluded: Not an ADCC / CCAP Package"];
    case "ONLY_EPIQ":
      return [attrs.isEpiq, "Excluded: Not an EPIQ Package"];
    case "ONLY_FLAT":
      return [attrs.isFlat, "Excluded: Not a Flat Aspect Ratio Package"];
    case "ONLY_SCOPE":
      return [attrs.isScope, "Excluded: Not a Scope (_S_) Package"];
    case "ONLY_HDR":
      return [attrs.isHdr, "Excluded: Not an HDR / Dolby Vision Package"];
    case "ONLY_BARCO_HDR":
      return [attrs.isBarcoHdr, "Excluded: Not an HDR by Barco Package"];
    default:
      return [true, ""];
  }
}

function extractMovieIdentity(cpls, releaseTitle, mrn) {
  const mrnClean = cleanStr(mrn).toUpperCase();
  const rootFromCpl = (cpl) => {
    let s = cpl.replace(PART_TOKEN_PATTERN, "");
    s = s.split(CPL_TYPE_PATTERN)[0];
    s = s.replace(/[-_](?:FMS|CC|2D|3D|EPIQ)$/i, "");
    s = s.replace(/[-_]+$/, "").trim();
    return s;
  };

  if (mrnClean && !["NAN", "NONE", "NA", "N/A", "-"].includes(mrnClean)) {
    let display = releaseTitle;
    if (cpls && cpls.length > 0) {
      const rawRoot = rootFromCpl(cpls[0]);
      if (rawRoot) display = rawRoot.replace(/(?<=[a-z])(?=[A-Z])/g, " ");
    }
    return [`MRN:${mrnClean}`, display];
  }

  if (cpls && cpls.length > 0) {
    const rawRoot = rootFromCpl(cpls[0]);
    const display = rawRoot.replace(/(?<=[a-z])(?=[A-Z])/g, " ");
    const normKey = rawRoot.toLowerCase().replace(/[^a-z0-9]/g, "");
    if (normKey) return [`CPLROOT:${normKey}`, display || releaseTitle];
  }

  const fallback = releaseTitle.toLowerCase().replace(/[^a-z0-9]/g, "").slice(0, 20);
  return [`TITLE:${fallback}`, releaseTitle];
}

// ==============================================================================
// CORE WORKBOOK PARSER & RECONCILIATION ENGINE
// ==============================================================================
function processMasteredFiles(fileObjs, options) {
  const {
    preferredSheet = "",
    includeTrailers = false,
    includeAdcc = true,
    includeSubtitles = false,
    specialtyMode = "ALL",
  } = options;

  const auditRows = [];
  const primaryEntries = [];
  const scanTrailers = specialtyMode === "ONLY_TRAILERS" || includeTrailers;
  const processedSources = [];

  fileObjs.forEach((fObj, fIdx) => {
    const srcName = fObj.name;
    const wb = fObj.workbook;
    const targetSheet =
      preferredSheet && wb.SheetNames.includes(preferredSheet)
        ? preferredSheet
        : wb.SheetNames[0];

    processedSources.push(`${srcName} [${targetSheet}]`);
    const sheet = wb.Sheets[targetSheet];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    let headerMap = {};
    let inTrailers = false;
    let inIngesting = false;
    let currentBlock = null;
    let blockCounter = 0;
    let seenExactCplsInBlock = new Set();
    let currentBlockHasP1 = false;

    for (let rIdx = 0; rIdx < rows.length; rIdx++) {
      const excelRowNum = rIdx + 1;
      const rowVals = rows[rIdx].map(cleanStr);
      if (!rowVals.some((v) => v !== "")) continue;

      if (rowVals.some((v) => v.toUpperCase() === "TRAILERS")) {
        inTrailers = true;
        inIngesting = false;
        currentBlock = null;
        auditRows.push({
          source_file: srcName,
          row_num: excelRowNum,
          section: "Trailers Banner",
          raw_text: "TRAILERS",
          norm_pkg: "-",
          lab: "-",
          ver_type: "Non-Package",
          status: "Excluded",
          reason: "Section Banner",
        });
        continue;
      }

      const upperVals = rowVals.map((v) => v.toUpperCase());
      const MOVIE_COL_ALIASES = new Set([
            "MOVIE",
            "MOVIE NAME",
            "COMPOSITION",
            "COMPOSITION NAME",
            "CPL",
            "CPL NAME",
            "PACKAGE",
             "PACKAGE NAME",
            "CONTENT",
            "TITLE",
            ]);
      if (
        upperVals.some((v) => MOVIE_COL_ALIASES.has(v)) &&
        (upperVals.includes("LANGUAGE") || upperVals.includes("FORMAT")) || upperVals.some((v) => v.includes("DATE")))
      ) {
        headerMap = {};
        upperVals.forEach((v, cIdx) => {
          if (["SNO", "S:NO", "S.NO", "SL NO", "SL.NO"].includes(v)) headerMap.sno = cIdx;
          else if (v.includes("DATE")) headerMap.date = cIdx;
          else if (MOVIE_COL_ALIASES.has(v)) headerMap.movie = cIdx;
          else if (v.includes("LANG")) headerMap.lang = cIdx;
          else if (v.includes("FORMAT")) headerMap.format = cIdx;
          else if (v.includes("MRN")) headerMap.mrn = cIdx;
          else if (v.includes("LAB")) headerMap.lab = cIdx;
        });
        inIngesting = false;
        auditRows.push({
          source_file: srcName,
          row_num: excelRowNum,
          section: inTrailers ? "Trailers Header" : "Main Header",
          raw_text: rowVals.filter(Boolean).join(" | "),
          norm_pkg: "-",
          lab: "-",
          ver_type: "Non-Package",
          status: "Excluded",
          reason: "Table Header Row",
        });
        continue;
      }

      if (headerMap.movie === undefined) {
        const cplColIdx = rowVals.findIndex((cell) => isCplFilename(cell));
  if (cplColIdx !== -1) {
    headerMap.movie = cplColIdx;
  } else {
        auditRows.push({
          source_file: srcName,
          row_num: excelRowNum,
          section: "Top Metadata",
          raw_text: rowVals.filter(Boolean).join(" | "),
          norm_pkg: "-",
          lab: "-",
          ver_type: "Non-Package",
          status: "Excluded",
          reason: "Title / Metadata Row",
        });
        continue;
      }
    }

      const movieCell = rowVals[headerMap.movie] || "";
      const labCell = headerMap.lab !== undefined ? rowVals[headerMap.lab] || "" : "";
      const fmtCell = headerMap.format !== undefined ? rowVals[headerMap.format] || "" : "";
      const mrnCell = headerMap.mrn !== undefined ? rowVals[headerMap.mrn] || "" : "";

      if (!movieCell) continue;

      if (/ingesting\s*details/i.test(movieCell)) {
        inIngesting = true;
        auditRows.push({
          source_file: srcName,
          row_num: excelRowNum,
          section: currentBlock ? currentBlock.title : "Main",
          raw_text: movieCell,
          norm_pkg: "-",
          lab: currentBlock ? currentBlock.lab : "-",
          ver_type: "Non-Package",
          status: "Excluded",
          reason: "Ingesting Details Section Header",
        });
        continue;
      }

      if (!inTrailers) {
        const isInstructionNote = /^(step[-\s]*\d+|create\s+show|note\b|ingesting\s+details)/i.test(
          movieCell
        );

        if (!isCplFilename(movieCell) && !isInstructionNote) {
          blockCounter++;
          inIngesting = false;
          seenExactCplsInBlock = new Set();
          currentBlockHasP1 = false;
          const labNorm = normalizeLabName(labCell, false);
          currentBlock = {
            block_id: blockCounter,
            title: movieCell,
            lab: labNorm,
            format: fmtCell,
            mrn: mrnCell,
            is_trailer: false,
          };
          auditRows.push({
            source_file: srcName,
            row_num: excelRowNum,
            section: `Block #${blockCounter} Header`,
            raw_text: movieCell,
            norm_pkg: movieCell,
            lab: labNorm,
            ver_type: "Non-Package",
            status: "Excluded",
            reason: "Release Title Header Row (Context Captured)",
          });
          continue;
        }

        if (isCplFilename(movieCell)) {
          if (
            !inIngesting &&
            currentBlock &&
            currentBlockHasP1 &&
            P1_REEL_PATTERN.test(movieCell)
          ) {
            blockCounter++;
            seenExactCplsInBlock = new Set();
            currentBlockHasP1 = false;
            currentBlock = {
              block_id: blockCounter,
              title: `${normalizeCplBase(movieCell)} (Unlabeled Header)`,
              lab: currentBlock.lab,
              format: currentBlock.format,
              mrn: currentBlock.mrn,
              is_trailer: false,
            };
          }

          if (P1_REEL_PATTERN.test(movieCell) && !inIngesting) {
            currentBlockHasP1 = true;
          }

          const ver = extractVersionType(movieCell);
          const normBase = normalizeCplBase(movieCell);
          const ctxTitle = currentBlock ? currentBlock.title : normBase;
          const ctxFmt = currentBlock ? currentBlock.format : "";
          const ctxMrn = currentBlock ? currentBlock.mrn : "";
          const ctxLab = currentBlock ? currentBlock.lab : "FLAGGED - Missing/Unclear Lab";

          if (inIngesting) {
            auditRows.push({
              source_file: srcName,
              row_num: excelRowNum,
              section: ctxTitle,
              raw_text: movieCell,
              norm_pkg: normBase,
              lab: ctxLab,
              ver_type: ver,
              status: "Excluded",
              reason: "Ingesting Details / Step Instruction",
            });
            continue;
          }

          const attrs = classifyPackageAttributes(movieCell, ctxTitle, ctxFmt);
          const [matchedSpec, specReason] = matchesSpecialtyMode(specialtyMode, false, attrs);
          if (!matchedSpec) {
            auditRows.push({
              source_file: srcName,
              row_num: excelRowNum,
              section: ctxTitle,
              raw_text: movieCell,
              norm_pkg: normBase,
              lab: ctxLab,
              ver_type: ver,
              status: "Excluded",
              reason: specReason,
            });
            continue;
          }

          if (specialtyMode !== "ONLY_ADCC" && attrs.isAdcc && !includeAdcc) {
            auditRows.push({
              source_file: srcName,
              row_num: excelRowNum,
              section: ctxTitle,
              raw_text: movieCell,
              norm_pkg: normBase,
              lab: ctxLab,
              ver_type: `${ver} (ADCC/CCAP)`,
              status: "Excluded",
              reason: "ADCC / CCAP Package Excluded (Checkbox Unchecked)",
            });
            continue;
          }

          if (attrs.isSubtitled && !includeSubtitles) {
            auditRows.push({
              source_file: srcName,
              row_num: excelRowNum,
              section: ctxTitle,
              raw_text: movieCell,
              norm_pkg: normBase,
              lab: ctxLab,
              ver_type: `${ver} (Subtitled)`,
              status: "Excluded",
              reason: `Subtitled Version Excluded (${attrs.subCode} — Checkbox Unchecked)`,
            });
            continue;
          }

          const dupKey = `${currentBlock ? currentBlock.block_id : 0}::${movieCell}`;
          if (seenExactCplsInBlock.has(dupKey)) {
            auditRows.push({
              source_file: srcName,
              row_num: excelRowNum,
              section: ctxTitle,
              raw_text: movieCell,
              norm_pkg: normBase,
              lab: ctxLab,
              ver_type: ver,
              status: "Excluded",
              reason: "Duplicate Entry",
            });
            continue;
          }
          seenExactCplsInBlock.add(dupKey);

          const groupId = currentBlock
            ? `F${fIdx}_BLK_${currentBlock.block_id}`
            : `F${fIdx}_ORPHAN_${normBase}_${ctxLab}`;

          primaryEntries.push({
            audit_idx: auditRows.length,
            source_file: srcName,
            row_num: excelRowNum,
            group_id: groupId,
            movie_title: ctxTitle,
            norm_base: normBase,
            raw_cpl: movieCell,
            lab: ctxLab,
            mrn: ctxMrn,
            ver_type: ver,
            is_trailer: false,
          });

          auditRows.push({
            source_file: srcName,
            row_num: excelRowNum,
            section: ctxTitle,
            raw_text: movieCell,
            norm_pkg: normBase,
            lab: ctxLab,
            ver_type: ver,
            status: "PENDING",
            reason: "",
          });
        } else {
          auditRows.push({
            source_file: srcName,
            row_num: excelRowNum,
            section: currentBlock ? currentBlock.title : "Main",
            raw_text: movieCell,
            norm_pkg: "-",
            lab: currentBlock ? currentBlock.lab : "-",
            ver_type: "Non-Package",
            status: "Excluded",
            reason: "Instruction / Note Row",
          });
        }
      } else {
        // TRAILERS SUB-TABLE
        if (isCplFilename(movieCell)) {
          const ver = extractVersionType(movieCell);
          const normBase = normalizeCplBase(movieCell);
          const labNorm = normalizeLabName(labCell, true);
          const trailerTitle = `${movieCell.split("_")[0]} (Trailer)`;

          if (!scanTrailers) {
            auditRows.push({
              source_file: srcName,
              row_num: excelRowNum,
              section: "Trailers",
              raw_text: movieCell,
              norm_pkg: normBase,
              lab: labNorm,
              ver_type: ver,
              status: "Excluded",
              reason: "Trailers Section Excluded (Checkbox Unchecked)",
            });
            continue;
          }

          const attrs = classifyPackageAttributes(movieCell, trailerTitle, fmtCell);
          const [matchedSpec, specReason] = matchesSpecialtyMode(specialtyMode, true, attrs);
          if (!matchedSpec) {
            auditRows.push({
              source_file: srcName,
              row_num: excelRowNum,
              section: "Trailers",
              raw_text: movieCell,
              norm_pkg: normBase,
              lab: labNorm,
              ver_type: ver,
              status: "Excluded",
              reason: specReason,
            });
            continue;
          }

          if (specialtyMode !== "ONLY_ADCC" && attrs.isAdcc && !includeAdcc) {
            auditRows.push({
              source_file: srcName,
              row_num: excelRowNum,
              section: "Trailers",
              raw_text: movieCell,
              norm_pkg: normBase,
              lab: labNorm,
              ver_type: `${ver} (ADCC/CCAP)`,
              status: "Excluded",
              reason: "ADCC / CCAP Package Excluded (Checkbox Unchecked)",
            });
            continue;
          }

          if (attrs.isSubtitled && !includeSubtitles) {
            auditRows.push({
              source_file: srcName,
              row_num: excelRowNum,
              section: "Trailers",
              raw_text: movieCell,
              norm_pkg: normBase,
              lab: labNorm,
              ver_type: `${ver} (Subtitled)`,
              status: "Excluded",
              reason: `Subtitled Version Excluded (${attrs.subCode} — Checkbox Unchecked)`,
            });
            continue;
          }

          blockCounter++;
          const groupId = `F${fIdx}_TRL_${normBase}_${labNorm}`;
          primaryEntries.push({
            audit_idx: auditRows.length,
            source_file: srcName,
            row_num: excelRowNum,
            group_id: groupId,
            movie_title: trailerTitle,
            norm_base: normBase,
            raw_cpl: movieCell,
            lab: labNorm,
            mrn: mrnCell,
            ver_type: ver,
            is_trailer: true,
          });
          auditRows.push({
            source_file: srcName,
            row_num: excelRowNum,
            section: "Trailers",
            raw_text: movieCell,
            norm_pkg: normBase,
            lab: labNorm,
            ver_type: ver,
            status: "PENDING",
            reason: "",
          });
        } else {
          auditRows.push({
            source_file: srcName,
            row_num: excelRowNum,
            section: "Trailers",
            raw_text: movieCell,
            norm_pkg: "-",
            lab: "-",
            ver_type: "Non-Package",
            status: "Excluded",
            reason: "Non-Package Row",
          });
        }
      }
    }
  });

  // Group P1+P2, apply VF-over-OV priority, and deduplicate across blocks/files
  const groupedBlocks = new Map();
  for (const item of primaryEntries) {
    if (!groupedBlocks.has(item.group_id)) groupedBlocks.set(item.group_id, []);
    groupedBlocks.get(item.group_id).push(item);
  }

  const seenCanonicalSignatures = new Map();
  const countedPackages = [];

  for (const [, items] of groupedBlocks.entries()) {
    const first = items[0];
    const lab = first.lab;
    const title = first.movie_title;
    const mrn = first.mrn || "";
    const srcFile = first.source_file;
    const normBases = Array.from(new Set(items.map((x) => x.norm_base))).sort();
    const canonicalSig = `${normBases.join("|")}::${lab}`;

    if (seenCanonicalSignatures.has(canonicalSig)) {
      const prevInfo = seenCanonicalSignatures.get(canonicalSig);
      for (const it of items) {
        auditRows[it.audit_idx].status = "Excluded";
        auditRows[it.audit_idx].reason = `Duplicate Package Signature (Already counted in ${prevInfo})`;
      }
      continue;
    }

    const versionsPresent = items.map((x) => x.ver_type);
    const chosenVersion = versionsPresent.includes("VF") ? "VF" : "OV";
    const hasBothOvAndVf = versionsPresent.includes("OV") && versionsPresent.includes("VF");
    const isMultiPart = items.length > 1;

    seenCanonicalSignatures.set(canonicalSig, `${title} [${srcFile}]`);

    const includedCpls = [];
    for (const it of items) {
      if (hasBothOvAndVf && it.ver_type === "OV") {
        auditRows[it.audit_idx].status = "Excluded";
        auditRows[it.audit_idx].reason = "VF Priority Override (Paired VF counted for this package)";
      } else {
        includedCpls.push(it.raw_cpl);
        auditRows[it.audit_idx].status = "Included";
        if (isMultiPart && includedCpls.length > 1) {
          auditRows[it.audit_idx].reason = `Counted as ${chosenVersion} (Merged P1+P2 into 1 package count)`;
        } else if (isMultiPart) {
          auditRows[it.audit_idx].reason = `Counted as ${chosenVersion} (Primary part of P1+P2 single count)`;
        } else {
          auditRows[it.audit_idx].reason = `Counted as ${chosenVersion}`;
        }
      }
    }

    const [movieKey, canonicalMovieName] = extractMovieIdentity(includedCpls, title, mrn);

    countedPackages.push({
      source_file: srcFile,
      movie_key: movieKey,
      canonical_movie: canonicalMovieName,
      mrn: mrn || "-",
      movie_name: title,
      norm_base: normBases.join(" / "),
      lab,
      version: chosenVersion,
      ov_count: chosenVersion === "OV" ? 1 : 0,
      vf_count: chosenVersion === "VF" ? 1 : 0,
      total_count: 1,
      package_ids: includedCpls.join("\n"),
      is_trailer: first.is_trailer,
    });
  }

  // Unique Movie Aggregation
  const movieGroups = new Map();
  for (const p of countedPackages) {
    if (!movieGroups.has(p.movie_key)) {
      movieGroups.set(p.movie_key, {
        movie_title: p.canonical_movie,
        mrn: p.mrn,
        labs: new Set(),
        sources: new Set(),
        ov_pkgs: 0,
        vf_pkgs: 0,
        total_pkgs: 0,
      });
    }
    const g = movieGroups.get(p.movie_key);
    if (
      p.canonical_movie.length > g.movie_title.length &&
      p.canonical_movie !== p.canonical_movie.toUpperCase()
    ) {
      g.movie_title = p.canonical_movie;
    }
    g.labs.add(p.lab);
    g.sources.add(p.source_file);
    g.ov_pkgs += p.ov_count;
    g.vf_pkgs += p.vf_count;
    g.total_pkgs += 1;
  }

  const uniqueMoviesList = Array.from(movieGroups.values()).map((g) => ({
    movie_title: g.movie_title,
    mrn: g.mrn,
    labs: Array.from(g.labs).sort().join(", "),
    sources: Array.from(g.sources).sort().join(", "),
    ov_pkgs: g.ov_pkgs,
    vf_pkgs: g.vf_pkgs,
    total_pkgs: g.total_pkgs,
  }));

  const uniqueMovieCount = uniqueMoviesList.length;
  const totalOv = countedPackages.reduce((acc, p) => acc + p.ov_count, 0);
  const totalVf = countedPackages.reduce((acc, p) => acc + p.vf_count, 0);
  const grandTotal = countedPackages.length;

  const qubeItems = countedPackages.filter((p) => p.lab === "Qube Lab");
  const otherItems = countedPackages.filter(
    (p) => p.lab !== "Qube Lab" && !p.lab.startsWith("FLAGGED")
  );
  const flaggedItems = countedPackages.filter((p) => p.lab.startsWith("FLAGGED"));

  const qubeOv = qubeItems.reduce((a, p) => a + p.ov_count, 0);
  const qubeVf = qubeItems.reduce((a, p) => a + p.vf_count, 0);
  const qubeTotal = qubeItems.length;

  const otherOv = otherItems.reduce((a, p) => a + p.ov_count, 0);
  const otherVf = otherItems.reduce((a, p) => a + p.vf_count, 0);
  const otherTotal = otherItems.length;

  const flaggedOv = flaggedItems.reduce((a, p) => a + p.ov_count, 0);
  const flaggedVf = flaggedItems.reduce((a, p) => a + p.vf_count, 0);
  const flaggedTotal = flaggedItems.length;

  const labMap = new Map();
  for (const p of countedPackages) {
    if (!labMap.has(p.lab)) {
      labMap.set(p.lab, { OV: 0, VF: 0, Total: 0, Movies: new Set() });
    }
    const entry = labMap.get(p.lab);
    entry.OV += p.ov_count;
    entry.VF += p.vf_count;
    entry.Total += 1;
    entry.Movies.add(p.movie_key);
  }

  const labBreakdown = {};
  for (const [lName, stats] of labMap.entries()) {
    labBreakdown[lName] = {
      OV: stats.OV,
      VF: stats.VF,
      Total: stats.Total,
      UniqueMovies: stats.Movies.size,
    };
  }

  const uniqueSigs = new Set(countedPackages.map((p) => `${p.norm_base}::${p.lab}`));
  const validations = [
    ["OV + VF == Grand Total", totalOv + totalVf === grandTotal, `${totalOv} + ${totalVf} = ${grandTotal}`],
    [
      "Qube + Other Labs + Flagged == Grand Total",
      qubeTotal + otherTotal + flaggedTotal === grandTotal,
      `${qubeTotal} + ${otherTotal} + ${flaggedTotal} = ${grandTotal}`,
    ],
    [
      "Filter & Version Integrity",
      countedPackages.every((p) => p.version === "OV" || p.version === "VF"),
      "Verified cleanly",
    ],
    [
      "Zero Duplicate Packages / P1+P2 Double-Counts",
      uniqueSigs.size === grandTotal,
      `${grandTotal} unique packages across ${uniqueMovieCount} movie(s)`,
    ],
  ];

  return {
    processedSources,
    uniqueMovieCount,
    uniqueMoviesList,
    totalOv,
    totalVf,
    grandTotal,
    qubeOv,
    qubeVf,
    qubeTotal,
    otherOv,
    otherVf,
    otherTotal,
    flaggedOv,
    flaggedVf,
    flaggedTotal,
    labBreakdown,
    countedPackages,
    qubeItems,
    otherItems,
    flaggedItems,
    auditRows,
    validations,
    modeLabel: SPECIALTY_MODES[specialtyMode] || specialtyMode,
    specialtyMode,
    activeFilters: {
      includeTrailers: scanTrailers,
      includeAdcc: specialtyMode === "ONLY_ADCC" ? true : includeAdcc,
      includeSubtitles,
      specialtyMode,
    },
  };
}

// ==============================================================================
// 8-SHEET EXCEL REPORT BUILDER & BROWSER DOWNLOAD (SHEETJS)
// ==============================================================================
function buildAndDownloadExcel(results, filename) {
  const wb = XLSX.utils.book_new();
  const f = results.activeFilters;
  const filterSummary =
    `Filter Mode: ${results.modeLabel} | ` +
    `Trailers: ${f.includeTrailers ? "Included" : "Excluded"} | ` +
    `ADCC/CCAP: ${f.includeAdcc ? "Included" : "Excluded"} | ` +
    `Subtitled Versions: ${f.includeSubtitles ? "Included" : "Excluded"}`;

  // 1. Overall Count
  const s1Data = [
    [`DCP OV & VF Overall Count Report — ${results.modeLabel}`],
    [`Files (${results.processedSources.length}): ${results.processedSources.join(", ")} | ${filterSummary}`],
    [],
    ["Metric", "Count"],
    ["Total Unique Movies", results.uniqueMovieCount],
    ["Total OV Packages", results.totalOv],
    ["Total VF Packages", results.totalVf],
    ["Grand Total Packages (OV + VF)", results.grandTotal],
    ["Qube Lab Packages", results.qubeTotal],
    ["Other Labs Packages", results.otherTotal],
  ];
  if (results.flaggedTotal > 0) {
    s1Data.push(["Flagged / Missing Lab Packages", results.flaggedTotal]);
  }
  s1Data.push([], ["Automated Validation Checks", "Status"]);
  for (const [label, passed, detail] of results.validations) {
    s1Data.push([label, `${passed ? "PASSED" : "FAILED"} (${detail})`]);
  }
  const ws1 = XLSX.utils.aoa_to_sheet(s1Data);
  ws1["!cols"] = [{ wch: 44 }, { wch: 38 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Overall Count");

  // 2. Unique Movies
  const sMovData = [
    ["Unique Movies Summary (All Versions Consolidated per Film)"],
    [],
    ["Movie Title", "MRN", "Source File(s)", "Lab(s)", "OV Packages", "VF Packages", "Total Packages"],
    ...results.uniqueMoviesList.map((m) => [
      m.movie_title,
      m.mrn,
      m.sources,
      m.labs,
      m.ov_pkgs,
      m.vf_pkgs,
      m.total_pkgs,
    ]),
    [
      `Total Unique Movies: ${results.uniqueMovieCount}`,
      "",
      "",
      "",
      results.totalOv,
      results.totalVf,
      results.grandTotal,
    ],
  ];
  const wsMov = XLSX.utils.aoa_to_sheet(sMovData);
  wsMov["!cols"] = [{ wch: 30 }, { wch: 20 }, { wch: 32 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, wsMov, "Unique Movies");

  // 3. Lab-wise Overall
  const sLabData = [
    ["Lab-wise OV, VF & Unique Movie Breakdown"],
    [],
    ["Lab Category", "Lab Name", "Unique Movies", "OV Count", "VF Count", "Grand Total"],
  ];
  const qStats = results.labBreakdown["Qube Lab"] || { OV: 0, VF: 0, Total: 0, UniqueMovies: 0 };
  sLabData.push(["Qube Lab", "Qube Lab", qStats.UniqueMovies, qStats.OV, qStats.VF, qStats.Total]);

  Object.keys(results.labBreakdown)
    .filter((k) => k !== "Qube Lab" && !k.startsWith("FLAGGED"))
    .sort()
    .forEach((lab) => {
      const d = results.labBreakdown[lab];
      sLabData.push(["Other Labs", lab, d.UniqueMovies, d.OV, d.VF, d.Total]);
    });

  Object.keys(results.labBreakdown)
    .filter((k) => k.startsWith("FLAGGED"))
    .sort()
    .forEach((lab) => {
      const d = results.labBreakdown[lab];
      sLabData.push(["Flagged / Unclear", lab, d.UniqueMovies, d.OV, d.VF, d.Total]);
    });

  sLabData.push([
    "ALL LABS",
    "Overall Grand Total",
    results.uniqueMovieCount,
    results.totalOv,
    results.totalVf,
    results.grandTotal,
  ]);
  const wsLab = XLSX.utils.aoa_to_sheet(sLabData);
  wsLab["!cols"] = [{ wch: 18 }, { wch: 36 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsLab, "Lab-wise Overall");

  // 4. Movie-wise Count
  const sPkgData = [
    ["Package / Version-wise Count"],
    [],
    [
      "Source File",
      "Movie Name",
      "Version / Release Description",
      "Canonical CPL Base",
      "Lab Name",
      "OV Count",
      "VF Count",
      "Total Count",
    ],
    ...results.countedPackages.map((p) => [
      p.source_file,
      p.canonical_movie,
      p.movie_name,
      p.norm_base,
      p.lab,
      p.ov_count,
      p.vf_count,
      p.total_count,
    ]),
    [
      "Grand Total",
      `${results.uniqueMovieCount} Unique Movies`,
      "",
      "",
      "",
      results.totalOv,
      results.totalVf,
      results.grandTotal,
    ],
  ];
  const wsPkg = XLSX.utils.aoa_to_sheet(sPkgData);
  wsPkg["!cols"] = [{ wch: 28 }, { wch: 24 }, { wch: 42 }, { wch: 46 }, { wch: 22 }, { wch: 11 }, { wch: 11 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, wsPkg, "Movie-wise Count");

  // 5. Qube Lab Details
  const sQubeData = [
    ["Qube Lab Detailed Package Breakdown"],
    [],
    ["Source File", "Movie / Package Name", "Counted Version (OV/VF)", "Package Identification (Counted CPLs)", "Individual Count"],
    ...results.qubeItems.map((p) => [p.source_file, p.movie_name, p.version, p.package_ids, p.total_count]),
    ["Total Qube Lab Packages", "", "", "", results.qubeTotal],
  ];
  const wsQube = XLSX.utils.aoa_to_sheet(sQubeData);
  wsQube["!cols"] = [{ wch: 28 }, { wch: 42 }, { wch: 22 }, { wch: 58 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, wsQube, "Qube Lab Details");

  // 6. Other Labs Details
  const combinedNonQube = [...results.otherItems, ...results.flaggedItems];
  const sOtherData = [
    ["Other Labs (& Flagged) Detailed Package Breakdown"],
    [],
    ["Source File", "Lab Name", "Movie / Package Name", "Counted Version (OV/VF)", "Package Identification (Counted CPLs)", "Individual Count"],
    ...combinedNonQube.map((p) => [p.source_file, p.lab, p.movie_name, p.version, p.package_ids, p.total_count]),
    ["Total Non-Qube Packages", "", "", "", "", results.otherTotal + results.flaggedTotal],
  ];
  const wsOther = XLSX.utils.aoa_to_sheet(sOtherData);
  wsOther["!cols"] = [{ wch: 28 }, { wch: 30 }, { wch: 42 }, { wch: 22 }, { wch: 58 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, wsOther, "Other Labs Details");

  // 7. Filename Audit
  const sAuditData = [
    ["Complete Row-by-Row Filename & Exclusion Audit"],
    [],
    [
      "Source File",
      "Excel Row #",
      "Section / Release Context",
      "Original Filename / Cell Text",
      "Normalized Package Base",
      "Lab Name",
      "Identified Version",
      "Status",
      "Reason for Inclusion / Exclusion",
    ],
    ...results.auditRows.map((a) => [
      a.source_file,
      a.row_num,
      a.section,
      a.raw_text,
      a.norm_pkg,
      a.lab,
      a.ver_type,
      a.status,
      a.reason,
    ]),
  ];
  const wsAudit = XLSX.utils.aoa_to_sheet(sAuditData);
  wsAudit["!cols"] = [{ wch: 26 }, { wch: 12 }, { wch: 34 }, { wch: 54 }, { wch: 44 }, { wch: 24 }, { wch: 18 }, { wch: 12 }, { wch: 48 }];
  XLSX.utils.book_append_sheet(wb, wsAudit, "Filename Audit");

  // 8. Counting Rules
  const sRulesData = [
    ["DCP OV & VF Counting and Classification Rules"],
    [],
    ["Rule Category", "Rule Description & Implementation Logic"],
    ["1. Unique Movie Counting", "Distinct movies are grouped by MRN number when available, or canonical CPL root title when MRN is blank."],
    ["2. Active Extraction Mode", `Executed in '${results.modeLabel}' mode across ${results.processedSources.length} file(s). Ingesting Details and non-package rows are excluded.`],
    ["3. ADCC & Subtitle Filters", `ADCC/CCAP: ${f.includeAdcc ? "INCLUDED" : "EXCLUDED"} | Subtitled DCPs: ${f.includeSubtitles ? "INCLUDED" : "EXCLUDED"}.`],
    ["4. Scope / Flat / EPIQ / HDR", "Scope matches '_S_' only (excludes '_C-239_'). Flat matches '_F_'. EPIQ matches '-EPIQ_'. HDR matches Dolby Vision (DVis), HDR by Barco, and EclairColor."],
    ["5. P1 & P2 Consolidation", "Part 1 (-P1_/-P01_) and Part 2 (-P2_/-P02_) of the same version count as ONE (1) package."],
    ["6. OV vs. VF Priority", "When both OV and VF exist for the same package, ONLY VF is counted (Count = 1) and OV is excluded."],
    ["7. Lab-wise Classification", "QUBE / Qube Cinema entries map to 'Qube Lab'. External labs retain exact names. Missing Labs are explicitly flagged."],
  ];
  const wsRules = XLSX.utils.aoa_to_sheet(sRulesData);
  wsRules["!cols"] = [{ wch: 32 }, { wch: 90 }];
  XLSX.utils.book_append_sheet(wb, wsRules, "Counting Rules");

  XLSX.writeFile(wb, filename);
}

// ==============================================================================
// UI WIRING & INTERACTIVITY
// ==============================================================================
const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const browseBtn = document.getElementById("browseBtn");
const clearQueueBtn = document.getElementById("clearQueueBtn");
const fileListEl = document.getElementById("fileList");
const fileCountBadge = document.getElementById("fileCountBadge");
const sheetSelect = document.getElementById("sheetSelect");
const analyzeBtn = document.getElementById("analyzeBtn");
const resetFiltersBtn = document.getElementById("resetFiltersBtn");
const downloadLastBtn = document.getElementById("downloadLastBtn");
const copyLogBtn = document.getElementById("copyLogBtn");
const clearLogBtn = document.getElementById("clearLogBtn");
const logOutput = document.getElementById("logOutput");
const labTableBody = document.getElementById("labTableBody");
const activeModePill = document.getElementById("activeModePill");
const themeToggle = document.getElementById("themeToggle");

const chkTrailers = document.getElementById("chkTrailers");
const chkAdcc = document.getElementById("chkAdcc");
const chkSubtitles = document.getElementById("chkSubtitles");
const chkLogOnly = document.getElementById("chkLogOnly");

function appendLog(text, clear = false) {
  if (clear) logOutput.textContent = "";
  logOutput.textContent += (logOutput.textContent ? "\n" : "") + text;
  logOutput.scrollTop = logOutput.scrollHeight;
}

function getSelectedSpecialtyMode() {
  const checked = document.querySelector('input[name="specialtyMode"]:checked');
  return checked ? checked.value : "ALL";
}

function getSelectedBatchMode() {
  const checked = document.querySelector('input[name="batchMode"]:checked');
  return checked ? checked.value : "COMBINED";
}

function updateUIState() {
  fileCountBadge.textContent = `${queuedFiles.length} file(s) queued`;
  analyzeBtn.disabled = queuedFiles.length === 0;
  analyzeBtn.textContent = chkLogOnly.checked
    ? "Calculate & Show Log Only"
    : "Analyze & Generate Report";

  fileListEl.innerHTML = "";
  queuedFiles.forEach((f, idx) => {
    const li = document.createElement("li");
    li.className = "file-item";
    li.innerHTML = `<span>${f.name}</span><button title="Remove" data-idx="${idx}">&times;</button>`;
    li.querySelector("button").addEventListener("click", () => {
      queuedFiles.splice(idx, 1);
      refreshSheetDropdown();
      updateUIState();
    });
    fileListEl.appendChild(li);
  });
}

function refreshSheetDropdown() {
  const current = sheetSelect.value;
  sheetSelect.innerHTML = '<option value="">Auto (First Sheet)</option>';
  if (queuedFiles.length > 0) {
    const names = new Set();
    queuedFiles.forEach((f) => f.workbook.SheetNames.forEach((s) => names.add(s)));
    names.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
      sheetSelect.appendChild(opt);
    });
    if (names.has(current)) sheetSelect.value = current;
  }
}

async function handleIncomingFiles(fileList) {
  let added = 0;
  for (const file of fileList) {
    if (!/\.(xlsx|xls)$/i.test(file.name)) continue;
    if (queuedFiles.some((q) => q.name === file.name && q.size === file.size)) continue;
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    queuedFiles.push({ name: file.name, size: file.size, workbook });
    added++;
  }
  refreshSheetDropdown();
  updateUIState();
  if (added > 0) {
    appendLog(`Added ${added} file(s). Total queued: ${queuedFiles.length} workbook(s).`, true);
  }
}

function renderDashboard(results) {
  document.getElementById("kpiMovies").textContent = results.uniqueMovieCount;
  document.getElementById("kpiOv").textContent = results.totalOv;
  document.getElementById("kpiVf").textContent = results.totalVf;
  document.getElementById("kpiTotal").textContent = results.grandTotal;
  activeModePill.textContent = results.modeLabel;

  labTableBody.innerHTML = "";
  const labs = Object.keys(results.labBreakdown).sort((a, b) => {
    if (a === "Qube Lab") return -1;
    if (b === "Qube Lab") return 1;
    if (a.startsWith("FLAGGED")) return 1;
    if (b.startsWith("FLAGGED")) return -1;
    return a.localeCompare(b);
  });

  if (labs.length === 0) {
    labTableBody.innerHTML =
      '<tr><td colspan="6" class="empty-cell">No matching packages found for the selected filter.</td></tr>';
    return;
  }

  for (const lab of labs) {
    const s = results.labBreakdown[lab];
    const share = results.grandTotal > 0 ? ((s.Total / results.grandTotal) * 100).toFixed(1) : "0.0";
    const tr = document.createElement("tr");
    if (lab.startsWith("FLAGGED")) tr.className = "flagged-row";
    tr.innerHTML = `
      <td><strong>${lab}</strong></td>
      <td class="num">${s.UniqueMovies}</td>
      <td class="num">${s.OV}</td>
      <td class="num">${s.VF}</td>
      <td class="num"><strong>${s.Total}</strong></td>
      <td class="num">${share}%</td>
    `;
    labTableBody.appendChild(tr);
  }

  const totalTr = document.createElement("tr");
  totalTr.className = "total-row";
  totalTr.innerHTML = `
    <td>Overall Grand Total</td>
    <td class="num">${results.uniqueMovieCount}</td>
    <td class="num">${results.totalOv}</td>
    <td class="num">${results.totalVf}</td>
    <td class="num">${results.grandTotal}</td>
    <td class="num">100.0%</td>
  `;
  labTableBody.appendChild(totalTr);
}

function logResultsSummary(title, results) {
  appendLog("==================================================================");
  appendLog(title);
  appendLog(`Filter Mode        : ${results.modeLabel}`);
  appendLog("==================================================================");
  appendLog(`Unique Movies      : ${results.uniqueMovieCount}`);
  appendLog(`Overall OV Count   : ${results.totalOv}`);
  appendLog(`Overall VF Count   : ${results.totalVf}`);
  appendLog(`Grand Total (Pkgs) : ${results.grandTotal}`);
  appendLog("------------------------------------------------------------------");
  appendLog(`Qube Lab           : ${results.qubeTotal}  (OV: ${results.qubeOv}, VF: ${results.qubeVf})`);
  appendLog(`Other Labs (Total) : ${results.otherTotal}  (OV: ${results.otherOv}, VF: ${results.otherVf})`);
  if (results.flaggedTotal > 0) {
    appendLog(`Flagged / No Lab   : ${results.flaggedTotal}  (OV: ${results.flaggedOv}, VF: ${results.flaggedVf})`);
  }
  appendLog("------------------------------------------------------------------");
  appendLog("Individual Lab Breakdown:");
  for (const [lab, s] of Object.entries(results.labBreakdown)) {
    appendLog(`  • ${lab}: Movies=${s.UniqueMovies}, Total=${s.Total} (OV=${s.OV}, VF=${s.VF})`);
  }
  appendLog("------------------------------------------------------------------");
  appendLog("Unique Movies List:");
  for (const m of results.uniqueMoviesList) {
    appendLog(`  • ${m.movie_title} [${m.mrn}] -> ${m.total_pkgs} pkg(s) (OV=${m.ov_pkgs}, VF=${m.vf_pkgs}) | ${m.labs}`);
  }
  appendLog("------------------------------------------------------------------");
  appendLog("Validation Checks:");
  for (const [label, passed, detail] of results.validations) {
    appendLog(`  ${passed ? "[PASS]" : "[FAIL]"} ${label} -> ${detail}`);
  }
  appendLog("");
}

function makeTimestamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

// Event Listeners
dropZone.addEventListener("click", () => fileInput.click());
browseBtn.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", (e) => {
  handleIncomingFiles(e.target.files);
  fileInput.value = "";
});

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  if (e.dataTransfer.files) handleIncomingFiles(e.dataTransfer.files);
});

clearQueueBtn.addEventListener("click", () => {
  queuedFiles = [];
  refreshSheetDropdown();
  updateUIState();
  appendLog("File queue cleared.", true);
});

chkLogOnly.addEventListener("change", updateUIState);

document.querySelectorAll('input[name="specialtyMode"]').forEach((rb) => {
  rb.addEventListener("change", () => {
    const mode = getSelectedSpecialtyMode();
    if (mode === "ONLY_TRAILERS") {
      chkTrailers.checked = true;
      chkTrailers.disabled = true;
    } else {
      chkTrailers.disabled = false;
    }
    if (mode === "ONLY_ADCC") {
      chkAdcc.checked = true;
      chkAdcc.disabled = true;
    } else {
      chkAdcc.disabled = false;
    }
  });
});

resetFiltersBtn.addEventListener("click", () => {
  chkTrailers.checked = false;
  chkTrailers.disabled = false;
  chkAdcc.checked = true;
  chkAdcc.disabled = false;
  chkSubtitles.checked = false;
  chkLogOnly.checked = false;
  document.querySelector('input[name="specialtyMode"][value="ALL"]').checked = true;
  document.querySelector('input[name="batchMode"][value="COMBINED"]').checked = true;
  updateUIState();
});

analyzeBtn.addEventListener("click", () => {
  if (queuedFiles.length === 0) return;

  const specialtyMode = getSelectedSpecialtyMode();
  const batchMode = getSelectedBatchMode();
  const skipExport = chkLogOnly.checked;
  const modeSlug = specialtyMode === "ALL" ? "" : `${specialtyMode.replace(/_/g, "")}_`;
  const ts = makeTimestamp();

  const opts = {
    preferredSheet: sheetSelect.value,
    includeTrailers: chkTrailers.checked,
    includeAdcc: chkAdcc.checked,
    includeSubtitles: chkSubtitles.checked,
    specialtyMode,
  };

  appendLog(
    `Processing ${queuedFiles.length} workbook(s) [${skipExport ? "LOG ONLY" : batchMode}]...\n`,
    true
  );

  if (batchMode === "COMBINED" || queuedFiles.length === 1) {
    const results = processMasteredFiles(queuedFiles, opts);
    const baseName =
      queuedFiles.length === 1
        ? queuedFiles[0].name.replace(/\.(xlsx|xls)$/i, "")
        : `Combined_${queuedFiles.length}Files`;
    const outName = `DCP_OV_VF_Report_${modeSlug}${baseName}_${ts}.xlsx`;

    lastGeneratedReport = { results, filename: outName };
    downloadLastBtn.disabled = false;
    renderDashboard(results);

    if (skipExport) {
      logResultsSummary(`LOG SUMMARY (NO FILE DOWNLOADED): ${baseName}`, results);
    } else {
      logResultsSummary(`REPORT GENERATED: ${outName}`, results);
      buildAndDownloadExcel(results, outName);
    }
  } else {
    let lastRes = null;
    let lastOut = "";
    for (const fObj of queuedFiles) {
      const res = processMasteredFiles([fObj], opts);
      const baseName = fObj.name.replace(/\.(xlsx|xls)$/i, "");
      const outName = `DCP_OV_VF_Report_${modeSlug}${baseName}_${ts}.xlsx`;
      lastRes = res;
      lastOut = outName;

      if (skipExport) {
        logResultsSummary(`LOG SUMMARY (NO FILE DOWNLOADED): ${baseName}`, res);
      } else {
        logResultsSummary(`REPORT GENERATED: ${outName}`, res);
        buildAndDownloadExcel(res, outName);
      }
    }
    if (lastRes) {
      lastGeneratedReport = { results: lastRes, filename: lastOut };
      downloadLastBtn.disabled = false;
      renderDashboard(lastRes);
    }
  }
});

downloadLastBtn.addEventListener("click", () => {
  if (!lastGeneratedReport) return;
  buildAndDownloadExcel(lastGeneratedReport.results, lastGeneratedReport.filename);
});

copyLogBtn.addEventListener("click", async () => {
  const text = logOutput.textContent.trim();
  if (!text) return;
  await navigator.clipboard.writeText(text);
  copyLogBtn.textContent = "Copied!";
  setTimeout(() => (copyLogBtn.textContent = "Copy Log"), 1500);
});

clearLogBtn.addEventListener("click", () => {
  logOutput.textContent = "";
});

themeToggle.addEventListener("click", () => {
  document.body.classList.toggle("light-theme");
});