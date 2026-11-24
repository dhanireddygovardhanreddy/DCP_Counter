"""resolvepatch - patch DaVinci Resolve/Fusion Studio on Windows or Linux.

Verified on:
  - Windows: Resolve 21.0.0.28, 21.0.4.5, 21.1.0.0014, 21.1.1.10, and 21.0.0 beta
  - Windows: Fusion Studio 21.0.0.25 and 21.1.0
  - Linux: Resolve 21.0.0.28, 21.0.0b.0033, 21.0.4.0005, 21.1.0.0014, and 21.1.1.0010
  - Linux: Fusion Studio 21.0 Beta 3, 21.0.4, and 21.1.0

Usage:
    python resolvepatch.py                                # patch every detected target
    python resolvepatch.py --restore                      # restore everything from .bak
    python resolvepatch.py --targets resolve              # only Resolve bundle
    python resolvepatch.py --targets fusion               # only Fusion DLL(s)
    python resolvepatch.py --targets all                  # everything detected
    python resolvepatch.py --path <p>                     # explicit Resolve binary path

The "resolve" target is a bundle: it patches Resolve.exe plus every companion
binary present in the install folder — fraunhoferdcp.dll (EasyDCP cinema-
package watermarks) and the Remote Monitor exe (Studio-only viewer app,
renamed from "DaVinci Remote Monitor.exe" to "Blackmagic Remote Monitor.exe"
in v21 beta — both names are probed).
"""

import argparse
import ctypes
import logging
import os
import re
import shutil
import struct
import subprocess
import sys
import time
from pathlib import Path
from typing import Callable, Optional, Sequence, Union

try:
    import msvcrt
except ImportError:
    msvcrt = None

try:
    import winreg
except ImportError:
    winreg = None

try:
    import termios
    import tty
except ImportError:
    termios = None
    tty = None

logger = logging.getLogger("resolvepatch")


# --------------------------------------------------------------------- constants

DEFAULT_PATH = r"C:\Program Files\Blackmagic Design\DaVinci Resolve\Resolve.exe"

DEFAULT_LINUX_RESOLVE_PATHS = (
    "/opt/resolve/bin/resolve",
    "/usr/bin/davinci-resolve",
)

DEFAULT_LINUX_FUSION_PATHS = (
    "/opt/BlackmagicDesign/Fusion21/libfusionsystem.so",
    "/opt/resolve/libs/Fusion/libfusionsystem.so",
)

__version__ = "2026.10.04.1"

# Standard install paths for Fusion Studio 21.
DEFAULT_FUSION_PATHS = (
    r"C:\Program Files\Blackmagic Design\Fusion 21\fusionsystem.dll",
)

# Companion binaries (fraunhoferdcp.dll for EasyDCP cinema-package export,
# Blackmagic/DaVinci Remote Monitor.exe for the Studio-only viewer app) live
# next to Resolve.exe and are auto-discovered from its install directory by
# `_resolve_companion_paths()`. v21 beta renamed Remote Monitor from "DaVinci
# Remote Monitor.exe" to "Blackmagic Remote Monitor.exe" — both names are
# probed.

# File extensions whose ShellOpen command points at Resolve.exe; used by the
# auto-locator when the standard install path doesn't exist.
SHELLOPEN_EXTENSION_KEYS = (
    "ResolveBinFile",
    "ResolveDrpFile",
    "ResolveDBKeyFile",
    "ResolveTimelineFile",
    "ResolveTemplateBundle",
)

# Fake RLM license file written next to Resolve.exe; HKLM env var
# RLM_LICENSE=blackmagic.lic points the RLM library at this name.
LICENSE_FILE_CONTENTS = (
    'LICENSE blackmagic davinciresolvestudio 999999 permanent uncounted\n'
    '  hostid=ANY issuer=ANY customer=ANY issued=14-Aug-2025\n'
    '  akey=0000-0000-0000-0000-0000 _ck=00 sig="00"'
)

LINUX_LICENSE_FILE_CONTENTS = (
    'LICENSE blackmagic davinciresolvestudio 999999 permanent uncounted\n'
    ' hostid=ANY issuer=CGP customer=CGP issued=28-dec-2023\n'
    ' akey=0000-0000-0000-0000 _ck=00 sig="00"'
)

LINUX_FUSION_LICENSE_FILE_CONTENTS = (
    'LICENSE blackmagic fusionstudio 999999 permanent uncounted\n'
    ' hostid=ANY issuer=CGP customer=CGP issued=28-dec-2023\n'
    ' akey=0000-0000-0000-0000 _ck=00 sig="00"'
)

# Pattern entry: int (0..255) for an exact byte, None for a single-byte wildcard.
PatternByte = Optional[int]
Pattern = Sequence[PatternByte]
# Replacement: either fixed bytes, or a callable computing bytes from context.
ReplacementFn = Callable[[bytearray, int, Pattern], bytes]
Replacement = Union[bytes, ReplacementFn]
PATCHABLE_STATES = ("UNPATCHED", "PATCHED", "PARTIAL", "UNKNOWN", "UNSUPPORTED")
RESTORABLE_STATES = ("PATCHED", "PARTIAL")


class PatchError(Exception):
    """Anything that prevents the patch from proceeding cleanly."""


# --------------------------------------------------------------------- pattern matching

_compiled_pattern_cache: dict = {}


def _compile_pattern(pattern: Pattern) -> "re.Pattern[bytes]":
    key = tuple(pattern)
    rx = _compiled_pattern_cache.get(key)
    if rx is None:
        parts = [b'.' if b is None else re.escape(bytes([b])) for b in pattern]
        rx = re.compile(b''.join(parts), re.DOTALL)
        _compiled_pattern_cache[key] = rx
    return rx


def find_all(data: bytes, pattern: Pattern, start: int = 0,
             end: Optional[int] = None) -> list:
    """Return absolute offsets where `pattern` matches inside `data[start:end]`."""
    if end is None:
        end = len(data)
    if all(b is not None for b in pattern):
        needle = bytes(pattern)  # type: ignore[arg-type]
        hits = []
        pos = data.find(needle, start, end)
        while pos != -1:
            hits.append(pos)
            pos = data.find(needle, pos + 1, end)
        return hits
    return [m.start() for m in _compile_pattern(pattern).finditer(data, start, end)]


def _pattern_search(data: bytes, pattern: Pattern) -> bool:
    if all(b is not None for b in pattern):
        return data.find(bytes(pattern)) != -1  # type: ignore[arg-type]
    return _compile_pattern(pattern).search(data) is not None


def _patched_pattern(sig: Pattern, replacement: Replacement) -> Pattern:
    """Build a search pattern for the already-patched bytes.

    Replacement callables usually preserve version-specific displacement bytes.
    To discover which output positions are preserved, run them twice with
    different wildcard values; positions that change stay wildcarded.
    """
    if not callable(replacement):
        return list(replacement)

    def make_sample(seed: int) -> bytearray:
        sample = bytearray(len(sig))
        for i, b in enumerate(sig):
            if b is None:
                sample[i] = (seed + (i * 37)) & 0xFF
            else:
                sample[i] = b
        return sample

    out_a = replacement(make_sample(0x21), 0, sig)
    out_b = replacement(make_sample(0xB7), 0, sig)
    if len(out_a) != len(out_b):
        raise PatchError("replacement callable returned unstable output length")
    return [None if a != b else a for a, b in zip(out_a, out_b)]


def _classify_patch_hits(original_hits: set[int],
                         patched_hits: set[int],
                         patch_count: int) -> str:
    """Map hit sets to a state. Patches that match neither form are treated
    as "doesn't apply to this build" (e.g., version-specific variants in the
    same table), not as failures — they're excluded from classification.

    `patch_count` is kept on the signature for backwards compat but is no
    longer used: a "complete" state is now defined as `every applicable
    pattern is in the patched form`, which is independent of total count."""
    if original_hits and patched_hits:
        return "PARTIAL"
    if original_hits:
        return "UNPATCHED"
    if patched_hits:
        return "PATCHED"
    return "UNKNOWN"


def _classify_patch_state_data(data: bytes,
                               patches: "list[tuple[Pattern, Replacement]]") -> str:
    originals = [sig for sig, _ in patches]
    patched = [_patched_pattern(sig, replacement) for sig, replacement in patches]
    original_hits = {i for i, sig in enumerate(originals) if _pattern_search(data, sig)}
    patched_hits = {i for i, sig in enumerate(patched) if _pattern_search(data, sig)}
    return _classify_patch_hits(original_hits, patched_hits, len(patches))


def _scan_patch_state_file(path: str,
                           patches: "list[tuple[Pattern, Replacement]]",
                           chunk_size: int = 16 * 1024 * 1024) -> str:
    originals = [sig for sig, _ in patches]
    patched = [_patched_pattern(sig, replacement) for sig, replacement in patches]
    all_patterns = originals + patched
    max_pat_len = max(len(sig) for sig in all_patterns)
    original_hits: set[int] = set()
    patched_hits: set[int] = set()
    tail = b""

    with open(path, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            data = tail + chunk
            for i, sig in enumerate(originals):
                if i not in original_hits and _pattern_search(data, sig):
                    original_hits.add(i)
            for i, sig in enumerate(patched):
                if i not in patched_hits and _pattern_search(data, sig):
                    patched_hits.add(i)
            if not original_hits and len(patched_hits) == len(patches):
                break
            tail = data[-(max_pat_len + 64):]

    return _classify_patch_hits(original_hits, patched_hits, len(patches))


# --------------------------------------------------------------------- replacement callables

def _force_first_jne_to_jmp(data, addr, _sig):
    """v21 chooser bypass — see PATCHES_21 entry.

    Pattern is 28 bytes; bytes [22..27] are `0F 85 disp32` (jne early_exit).
    Convert to `90 E9 disp32` (nop+jmp, same target via preserved displacement).
    The license-check function then unconditionally takes the early-exit
    success branch and never reaches dialog construction."""
    return (bytes(data[addr:addr + 22])
            + bytes([0x90, 0xE9])
            + bytes(data[addr + 24:addr + 28]))


def _fusion_dolicensing_je_to_jmp_epilogue(data, addr, _sig):
    """Fusion Studio activation-dialog bypass — see PATCHES_FUSION entry.

    Pattern is 21 bytes; the last 2 are `74 11` (je rel8 +0x11), routing to
    a THREAD_SPAWN block inside Fusion::FusionApp::DoLicensing(). Replace
    those 2 bytes with `EB 71` (jmp rel8 +0x71) so control falls through to
    the function's `mov al, 1; ret` epilogue instead. The first 19 bytes
    (which include the version-specific [g_IsResolve] disp32) are preserved
    verbatim. The `+0x71` displacement has been stable across observed
    versions because the THREAD_SPAWN block has identical MSVC codegen."""
    return bytes(data[addr:addr + 19]) + bytes([0xEB, 0x71])


def _linux_nop_je(data: bytearray, addr: int, _sig) -> bytes:
    """Linux Resolve v21 dispatcher bypass.

    Pattern is 18 bytes; bytes [12..18] are `0f 84 disp32` (je error/dialog).
    Replace those 6 bytes with NOPs so execution falls through to success.
    """
    result = bytearray(data[addr:addr + 18])
    result[12:18] = b"\x90\x90\x90\x90\x90\x90"
    return bytes(result)


def _linux_jne_to_jmp(data: bytearray, addr: int, _sig) -> bytes:
    """Linux Resolve v21 secondary bypass.

    Pattern is 15 bytes; bytes [9..15] are `0f 85 disp32` (jne success).
    Convert to `90 e9 disp32`, preserving the original displacement.
    """
    result = bytearray(data[addr:addr + 15])
    result[9] = 0x90
    result[10] = 0xE9
    return bytes(result)


# --- generic replacement helper for "copy pattern_len bytes, splice at offsets" ---

def _make_subst(pattern_len: int,
                substitutions: "list[tuple[int, bytes]]") -> ReplacementFn:
    """Build a replacement callable that copies `pattern_len` bytes starting at
    `addr` and splices the given `(offset, replacement_bytes)` pairs in. All
    bytes outside the splice ranges are preserved (so wildcards in the search
    pattern remain themselves, including rel32 displacements, register
    encodings, etc.)."""
    def fn(data: bytearray, addr: int, _sig) -> bytes:
        result = bytearray(data[addr:addr + pattern_len])
        for off, repl in substitutions:
            result[off:off + len(repl)] = repl
        return bytes(result)
    return fn


# Linux Resolve 21.0.4+ / modern dispatcher and helper replacers
_linux_nop_je_anchored = _make_subst(
    pattern_len=21,
    substitutions=[(15, b"\x90\x90\x90\x90\x90\x90")],
)

_linux_jne_to_jmp_anchored = _make_subst(
    pattern_len=20,
    substitutions=[(14, b"\x90\xE9")],
)

_linux_helper_check1_replace = _make_subst(
    pattern_len=30,
    substitutions=[(7, b"\xEB")],
)

_linux_helper_check2_replace = _make_subst(
    pattern_len=33,
    substitutions=[(7, b"\xEB")],
)


# Windows Resolve 21.0.4+ chooser replacer
_v2104_chooser_replace = _make_subst(
    pattern_len=23,
    substitutions=[(9, b"\x90\xE9")],
)



# --- Resolve.exe: License A/B/C (universal across v20 / v21) ---

# License A: function prologue `40 53 ... 84 C0 ... 84 C0` is an inline license
# checker that returns its own al via two `test al, al` instructions feeding a
# je and a final return. Replace both `84 C0` (test al, al) with `B0 01`
# (mov al, 1) so the test always yields ZF=0, the conditional is skipped, and
# the final return delivers al=1 (success).
_license_a_replace = _make_subst(
    pattern_len=45,
    substitutions=[(22, b"\xB0\x01"), (43, b"\xB0\x01")],
)

# License B: a stack-frame-loading function whose entry and tail-call are
# guarded by `74` (je rel8) short jumps. Flip both to `EB` (jmp rel8) so the
# function always takes the "license valid" branch through the body and
# bypasses the failure path at the tail.
_license_b_replace = _make_subst(
    pattern_len=55,
    substitutions=[(0, b"\xEB"), (54, b"\xEB")],
)

# License C: short helper whose first instruction is `40 84 F6` (test sil, sil)
# followed by `74` (je) over the body. Flip `74` -> `75` (jne) so the body is
# executed exactly when the original code would have skipped it — net effect:
# the success-side code path runs unconditionally.
#
# Two pattern variants share the same splice (position 3, 74 -> 75):
#   * v20.x compatible: anchors on the full body including the
#     `48 8B 5C 24 ?? 48 8B 74 24 ?? 48 83 C4 ?? 5F` epilogue.
#   * v21-beta+ compatible: shortened to the immutable prefix only, because
#     v21 beta tail-jumps to a shared exit (E9 disp32) instead of inlining
#     the cleanup. The shorter prefix is still unique in both v20.3.2 and
#     v21.0.0-beta installs.
# License C: short helper `test sil, sil; je; mov edx, imm; mov ecx, imm;
# call rel32; mov eax, imm`. Flip 74 (je) -> 75 (jne) at position 3 so the body
# always runs.
_license_c_replace_v21 = _make_subst(
    pattern_len=25,
    substitutions=[(3, b"\x75")],
)


# --- fraunhoferdcp.dll: EasyDCP watermark/branding bypass (8 patches) ---
# Each EasyDCP patch wires around a `test bl, bl` (84 DB) or related check that
# gates the cinema-package watermark/branding code path. The replacements come
# straight from `Patches.txt`:
#   * `84 DB` -> `FE C3` = `inc bl` (always sets bl != 0)
#   * `84 DB` -> `32 DB` = `xor bl, bl` (always sets bl = 0)
#   * `75 15` -> `75 00` = short jne with displacement 0 (effective NOP)
# Lengths and splice offsets were derived from the original/patched pattern
# pairs in Patches.txt; each pattern matches exactly once in fraunhoferdcp.dll
# (verified).

_easydcp_1_replace = _make_subst(21, [(12, b"\x00")])
_easydcp_2_replace = _make_subst(38, [(27, b"\xFE\xC3")])
_easydcp_3_replace = _make_subst(37, [(0,  b"\xFE\xC3")])
_easydcp_4_replace = _make_subst(37, [(0,  b"\xFE\xC3")])
_easydcp_5_replace = _make_subst(48, [(0,  b"\xFE\xC3")])
_easydcp_6_replace = _make_subst(48, [(0,  b"\xFE\xC3")])
_easydcp_7_replace = _make_subst(49, [(0,  b"\xFE\xC3")])
_easydcp_8_replace = _make_subst(39, [(0,  b"\x32\xDB")])


# --- DaVinci Remote Monitor.exe ---
# After a `jg` (0F 8F) and a `mov [m], bl` (88 1D), the binary tests `bl` and
# branches on Studio license state. Replace `test bl, bl` with `mov bl, 1`
# (B3 01) so the test reflects "licensed" regardless.
_remote_monitor_replace = _make_subst(28, [(12, b"\xB3\x01")])


# --------------------------------------------------------------------- patch tables
#
# Each patch is `(pattern, replacement)`.

# Resolve 21.0.0.x chooser patch
PATCHES_21: "list[tuple[Pattern, Replacement]]" = [
    (
        [0x48, 0x89, 0x5C, 0x24, 0x10, 0x57,
         0x48, 0x81, 0xEC, 0x80, 0x00, 0x00, 0x00,
         0x33, 0xDB,
         0xE8, None, None, None, None,
         0x84, 0xC0,
         0x0F, 0x85, None, None, None, None],
        _force_first_jne_to_jmp,
    ),
]

# Resolve 21.0.4+ / 21.1+ / 21.1.1+ Windows chooser patch
PATCHES_21_2104: "list[tuple[Pattern, Replacement]]" = [
    # Chooser bypass: in v21.0.4+ the enclosing function prologue changed, but
    # the core chooser dispatch remains: `33 DB; call check; test al, al; jne early_exit`.
    # Convert `0F 85` (jne) at index 9 to `90 E9` (jmp) so Resolve always takes
    # the early-exit success path.
    (
        [0x33, 0xDB,
         0xE8, None, None, None, None,
         0x84, 0xC0,
         0x0F, 0x85, None, None, None, None,
         0xC7, 0x44, 0x24, 0x20, 0xFF, 0xFF, 0xFF, 0xFF],
        _v2104_chooser_replace,
    ),
]
PATCHES_21_MODERN = PATCHES_21_2104

# Universal Resolve 21.x patches (License A/B/C).
_PATCHES_UNIVERSAL_RESOLVE: "list[tuple[Pattern, Replacement]]" = [
    # License A: function prologue at file offset 0x3C5BE80 (v21 install).
    #   40 53                     push rbx
    #   48 83 EC ??               sub rsp, ??
    #   48 8B D9                  mov rbx, rcx
    #   E8 ?? ?? ?? ??            call <inner_check_1>
    #   48 8B C8                  mov rcx, rax
    #   E8 ?? ?? ?? ??            call <inner_check_2>
    #   84 C0                     test al, al        <- patched to B0 01 (mov al,1)
    #   0F 84 ?? ?? ?? ??         je  <error>
    #   E8 ?? ?? ?? FF            call <log>
    #   48 8B C8                  mov rcx, rax
    #   E8 ?? ?? ?? ??            call <inner_check_3>
    #   84 C0                     test al, al        <- patched to B0 01 (mov al,1)
    (
        [0x40, 0x53, 0x48, 0x83, 0xEC, None,
         0x48, 0x8B, 0xD9,
         0xE8, None, None, None, None,
         0x48, 0x8B, 0xC8,
         0xE8, None, None, None, None,
         0x84, 0xC0,
         0x0F, 0x84, None, None, None, None,
         0xE8, None, None, None, 0xFF,
         0x48, 0x8B, 0xC8,
         0xE8, None, None, None, None,
         0x84, 0xC0],
        _license_a_replace,
    ),
    # License B: nested `74 ... 74` short-jump pair (entry/tail guard) inside
    # a stack-frame-loading function. Flipping both `74` -> `EB` removes
    # both the entry and tail failure branches.
    (
        [0x74, None,
         0x48, 0x8B, 0x44, 0x24, None,
         0x8B, 0x4C, 0x24, None,
         0x89, 0x48, None,
         0x33, 0xC0, 0xEB, 0x3B,
         0x48, 0x8B, 0x44, 0x24, None,
         0x8B, 0x90, None, None, 0x00, 0x00,
         0x48, 0x8B, 0x44, 0x24, None,
         0x8B, 0x88, None, None, 0x00, 0x00,
         0xE8, None, None, None, None,
         0x89, 0x44, 0x24, None,
         0x83, 0x7C, 0x24, None, 0x00,
         0x74],
        _license_b_replace,
    ),
    # License C (v21 layout): anchor on immutable prefix (test sil, sil; je;
    # mov edx, imm; mov ecx, imm; call rel32; mov eax, imm). Flip 74 -> 75.
    (
        [0x40, 0x84, 0xF6, 0x74, None,
         0xBA, None, None, 0x00, 0x00,
         0xB9, None, 0x00, 0x00, 0x00,
         0xE8, None, None, None, None,
         0xB8, None, None, None, None],
        _license_c_replace_v21,
    ),
]


# Patch tables for fraunhoferdcp.dll (EasyDCP cinema-package export) and
# DaVinci Remote Monitor.exe. Both binaries live in the same Resolve install
# folder, so locating them is trivial.
PATCHES_EASYDCP: "list[tuple[Pattern, Replacement]]" = [
    # #1: short jne with 0x15 displacement; flipping disp to 0 makes the
    # jne a no-op step.
    (
        [0xC6, 0x00, None,
         0xE9, None, None, None, None,
         0x83, 0xF8, None,
         0x75, 0x15,
         0x48, 0x8D, 0x54, 0x24, None,
         0x48, 0x8B, 0xCE],
        _easydcp_1_replace,
    ),
    # #2: `test bl, bl; je 0x930` — replace test with `inc bl` so bl != 0.
    (
        [0x0F, 0xB6, 0xD8,
         0x48, 0x8D, 0x4C, 0x24, None,
         0xFF, 0x15, None, None, None, None,
         0x48, 0x8D, 0x54, 0x24, None,
         None, None, None,
         0xE8, None, None, 0x00, 0x00,
         0x84, 0xDB,
         0x0F, 0x84, 0x30, 0x09, 0x00, 0x00,
         0x48, 0x8B, 0xC8],
        _easydcp_2_replace,
    ),
    # #3-#7: `84 DB 0F 85 ...` (test bl, bl; jne ...) at the top of various
    # cinema-watermark code paths. Replace `84 DB` with `FE C3` (inc bl) so
    # the test result is always non-zero — the jne is always taken, skipping
    # the watermark/branding code that lives in the fall-through path.
    (
        [0x84, 0xDB, 0x0F, 0x85, None, None, None, None,
         0xBF, None, None, None, 0x00,
         0x48, 0x89, 0x75, 0xC0,
         0x48, 0x89, 0x75, 0xD0,
         0xBB, 0x0F, 0x00, 0x00, 0x00,
         0x48, 0x89, 0x5D, 0xD8,
         0xC6, 0x45, 0xC0, 0x00,
         0x8D, 0x4B, 0x31],
        _easydcp_3_replace,
    ),
    (
        [0x84, 0xDB, 0x0F, 0x85, 0x13, 0x01, 0x00, 0x00,
         0xBF, 0x95, 0x5F, 0x01, 0x00,
         0x48, 0x89, 0x75, 0xC0,
         0x48, 0x89, 0x75, 0xD0,
         0xBB, 0x0F, 0x00, 0x00, 0x00,
         0x48, 0x89, 0x5D, 0xD8,
         0xC6, 0x45, 0xC0, 0x00,
         0x8D, 0x4B, 0x11],
        _easydcp_4_replace,
    ),
    (
        [0x84, 0xDB, 0x0F, 0x85, 0xEC, 0x01, 0x00, 0x00,
         0xBF, 0x96, 0x5F, 0x01, 0x00,
         0x48, 0x89, 0x75, 0xC0,
         0x48, 0x89, 0x75, 0xD0,
         0xB8, 0x0F, 0x00, 0x00, 0x00,
         0x48, 0x89, 0x45, 0xD8,
         0x88, 0x5D, 0xC0,
         0x44, 0x8D, 0x40, 0x31,
         0x48, 0x8D, 0x15, None, None, None, None,
         0x48, 0x8D, 0x4D, 0xC0],
        _easydcp_5_replace,
    ),
    (
        [0x84, 0xDB, 0x0F, 0x85, 0xAB, 0x00, 0x00, 0x00,
         0xBF, None, None, None, 0x00,
         0x48, 0x89, 0x75, None,
         0x48, 0x89, 0x75, None,
         0xBB, None, 0x00, 0x00, 0x00,
         0x48, 0x89, 0x5D, None,
         0xC6, 0x45, None, 0x00,
         0x44, 0x8D, 0x43, None,
         0x48, 0x8D, 0x15, None, None, None, None,
         0x48, 0x8D, 0x4D],
        _easydcp_6_replace,
    ),
    (
        [0x84, 0xDB,
         0xBB, None, 0x00, 0x00, 0x00,
         0x0F, 0x85, 0xA8, 0x00, 0x00, 0x00,
         0xBF, None, None, None, 0x00,
         0x48, 0x89, 0x75, None,
         0x48, 0x89, 0x75, None,
         0x48, 0x89, 0x5D, None,
         0xC6, 0x45, None, None,
         0x44, 0x8D, 0x43, None,
         0x48, 0x8D, 0x15, None, None, None, None,
         0x48, 0x8D, 0x4D, 0xC0],
        _easydcp_7_replace,
    ),
    # #8: alternate branch — `84 DB` -> `32 DB` (xor bl, bl). The matching
    # jne immediately afterward is always taken, skipping the watermarking
    # code (mirror of the inc-bl approach used by #2-#7, just on the opposite
    # branch direction).
    (
        [0x84, 0xDB, 0x0F, 0x85, 0x67, 0x02, 0x00, 0x00,
         0xBF, None, None, None, 0x00,
         0x41, 0xBF, None, 0x00, 0x00, 0x00,
         0xBB, None, 0x00, 0x00, 0x00,
         0x48, 0x89, 0x5D, None,
         0x45, 0x8D, None, None,
         0x48, 0x8D, 0x15, None, None, None, None],
        _easydcp_8_replace,
    ),
]

PATCHES_REMOTE_MONITOR: "list[tuple[Pattern, Replacement]]" = [
    # Studio-feature gate just after the panel-id check: `test bl, bl;
    # jne <feature_off>`. Replace `test bl, bl` (84 DB) with `mov bl, 1`
    # (B3 01) so bl is forced to non-zero, the jne is always taken, and
    # Studio-only features stay on.
    (
        [0x0F, 0x8F, None, None, None, None,
         0x88, 0x1D, None, None, None, None,
         0x84, 0xDB,
         0x0F, 0x85, None, None, None, None,
         0xBA, None, 0x00, 0x00, 0x00,
         0x48, 0x8D, 0x0D],
        _remote_monitor_replace,
    ),
]


# Fusion Studio activation-dialog bypass (covers v20.3.x and v21.0.x). Inside
# Fusion::FusionApp::DoLicensing() the relevant code is:
#     mov  rcx, rbx                ; 48 8B CB
#     call qword [rax+0x220]       ; FF 90 20 02 00 00     (Log "Checking for licenses...")
#     mov  rax, [g_IsResolve]      ; 48 8B 05 disp32        (disp32 is version-specific)
#     cmp  byte [rax], 0           ; 80 38 00
#     je   THREAD_SPAWN            ; 74 11
# The je conditionally enters a block that allocates Events, registers a
# callback in [g_handler], spawns a worker thread, and ResumeThread()s it.
# That worker thread is what eventually requests the activation-key dialog.
# Flipping `74 11` -> `EB 71` retargets the jmp to the function's
# `mov al, 1; ret` epilogue, skipping the THREAD_SPAWN block entirely. As a
# bonus, InitInstance's wait-for-license-message loop is gated by
# `cmp [g_handler], <addr>` and skips itself when the handler is unset.
#
# Pattern length is 21 bytes (4 wildcards). An earlier 12-byte version
# (mov rax,[mem]; cmp byte [rax],0; je +0x11) had two coincidental matches
# elsewhere in the binary — likely other "is feature flag X enabled?"
# checks — so the patcher's "matched > 1 time" guard would refuse to apply.
# Anchoring on the preceding `48 8B CB FF 90 20 02 00 00` (mov rcx, rbx +
# call vtable[0x220]) makes the match unique on both v20 and v21.
PATCHES_FUSION: "list[tuple[Pattern, Replacement]]" = [
    (
        [0x48, 0x8B, 0xCB,
         0xFF, 0x90, 0x20, 0x02, 0x00, 0x00,
         0x48, 0x8B, 0x05, None, None, None, None,
         0x80, 0x38, 0x00,
         0x74, 0x11],
        _fusion_dolicensing_je_to_jmp_epilogue,
    ),
]

# Append universal License A/B/C patches to v21 tables.
PATCHES_21.extend(_PATCHES_UNIVERSAL_RESOLVE)
PATCHES_21_2104.extend(_PATCHES_UNIVERSAL_RESOLVE)


PATCHES_LINUX_21_2104: "list[tuple[Pattern, Replacement]]" = [
    # Resolve 21.0.4.0005+, 21.1.0.0014+, and 21.1.1.0010+ Linux ELF x86-64.
    # Patch 0: final license-check dispatcher branch.
    #   48 89 c7                   mov rdi, rax
    #   be 05 00 00 00             mov esi, 5
    #   e8 XX XX XX XX             call <check_fn>
    #   84 c0                      test al, al
    #   0f 84 disp32               je <error_path>  <- NOP this (6 bytes)
    (
        [0x48, 0x89, 0xC7,
         0xBE, 0x05, 0x00, 0x00, 0x00,
         0xE8, None, None, None, None,
         0x84, 0xC0,
         0x0F, 0x84, None, None, None, None],
        _linux_nop_je_anchored,
    ),
    # Patch 1: secondary license check.
    #   53                         push rbx
    #   48 83 ec 20                sub rsp, 0x20
    #   b3 01                      mov bl, 1
    #   e8 XX XX XX XX             call <check_fn>
    #   84 c0                      test al, al
    #   0f 85 disp32               jne <success>    <- force jmp
    (
        [0x53,
         0x48, 0x83, 0xEC, 0x20,
         0xB3, 0x01,
         0xE8, None, None, None, None,
         0x84, 0xC0,
         0x0F, 0x85, None, None, None, None],
        _linux_jne_to_jmp_anchored,
    ),
    # Patch 2: licensing helper check 1.
    #   89 45 fc                   mov [rbp-4], eax
    #   83 7d fc 00                cmp [rbp-4], 0
    #   74 11                      je +0x11         <- force jmp rel8 (eb 11)
    (
        [0x89, 0x45, 0xFC,
         0x83, 0x7D, 0xFC, 0x00,
         0x74, 0x11,
         0x48, 0x8B, 0x45, 0xC8,
         0x8B, 0x55, 0xFC,
         0x89, 0x50, 0x58,
         0xB8, 0x00, 0x00, 0x00, 0x00,
         0xEB, None,
         0x48, 0x8B, 0x45, 0xE0],
        _linux_helper_check1_replace,
    ),
    # Patch 3: licensing helper check 2.
    #   89 45 fc                   mov [rbp-4], eax
    #   83 7d fc 00                cmp [rbp-4], 0
    #   74 11                      je +0x11         <- force jmp rel8 (eb 11)
    (
        [0x89, 0x45, 0xFC,
         0x83, 0x7D, 0xFC, 0x00,
         0x74, 0x11,
         0x48, 0x8B, 0x45, 0xC8,
         0x8B, 0x55, 0xFC,
         0x89, 0x50, 0x58,
         0xB8, 0x00, 0x00, 0x00, 0x00,
         0xEB, None,
         0xB8, 0x01, 0x00, 0x00, 0x00,
         0xC9, 0xC3],
        _linux_helper_check2_replace,
    ),
]
PATCHES_LINUX_21_MODERN = PATCHES_LINUX_21_2104


PATCHES_LINUX_21: "list[tuple[Pattern, Replacement]]" = [
    # Resolve 21.0.0.28 Linux ELF x86-64.
    # Patch 0: final license-check dispatcher branch.
    #   be 05 00 00 00             mov esi, 5
    #   e8 XX XX XX XX             call <check_fn>
    #   84 c0                      test al, al
    #   0f 84 d3 00 00 00          je <error_path>  <- NOP this
    (
        [0xBE, 0x05, 0x00, 0x00, 0x00,
         0xE8, None, None, None, None,
         0x84, 0xC0,
         0x0F, 0x84, 0xD3, 0x00, 0x00, 0x00],
        _linux_nop_je,
    ),
    # Patch 1: secondary license check.
    #   b3 01                      mov bl, 1
    #   e8 XX XX XX XX             call <check_fn>
    #   84 c0                      test al, al
    #   0f 85 cc 00 00 00          jne <success>    <- force jmp
    (
        [0xB3, 0x01,
         0xE8, None, None, None, None,
         0x84, 0xC0,
         0x0F, 0x85, 0xCC, 0x00, 0x00, 0x00],
        _linux_jne_to_jmp,
    ),
]


PATCHES_LINUX_21_B0033: "list[tuple[Pattern, Replacement]]" = [
    # Resolve 21.0.0b.0033 Linux ELF x86-64.
    #
    # This build still has the 21.x "Checking Licenses" dispatcher, but the
    # branch displacements changed from build 0028. Patch 0 NOPs the final
    # `je failure_path`; patch 1 turns the early-success `jne` into `jmp`.
    (
        [0xBE, 0x05, 0x00, 0x00, 0x00,
         0xE8, 0x1B, 0x8A, 0x01, 0x00,
         0x84, 0xC0,
         0x0F, 0x84, 0xCA, 0x00, 0x00, 0x00],
        _linux_nop_je,
    ),
    (
        [0xB3, 0x01,
         0xE8, 0x24, 0x6C, 0x94, 0x03,
         0x84, 0xC0,
         0x0F, 0x85, 0xC9, 0x00, 0x00, 0x00],
        _linux_jne_to_jmp,
    ),
    (
        [0x89, 0x45, 0xFC,
         0x83, 0x7D, 0xFC, 0x00,
         0x74, 0x11,
         0x48, 0x8B, 0x45, 0xC8,
         0x8B, 0x55, 0xFC,
         0x89, 0x50, 0x58,
         0xB8, 0x00, 0x00, 0x00, 0x00,
         0xEB, None,
         0x48, 0x8B, 0x45, 0xE0],
        _linux_helper_check1_replace,
    ),
    (
        [0x89, 0x45, 0xFC,
         0x83, 0x7D, 0xFC, 0x00,
         0x74, 0x11,
         0x48, 0x8B, 0x45, 0xC8,
         0x8B, 0x55, 0xFC,
         0x89, 0x50, 0x58,
         0xB8, 0x00, 0x00, 0x00, 0x00,
         0xEB, None,
         0xB8, 0x01, 0x00, 0x00, 0x00,
         0xC9, 0xC3],
        _linux_helper_check2_replace,
    ),
]


PATCHES_LINUX_FUSION: "list[tuple[Pattern, Replacement]]" = [
    # Fusion Studio 20.3.2 / 21 Linux libfusionsystem.so.
    # Verified on standalone Fusion Studio 21 Beta 3 and 21.0.4 bundled.
    # Same small helper shape as Resolve: two `cmp result,0; je continue`
    # branches guard success in a licensing helper. Force both short branches
    # to unconditional jumps.
    (
        [0x89, 0x45, 0xFC,
         0x83, 0x7D, 0xFC, 0x00,
         0x74, 0x11,
         0x48, 0x8B, 0x45, 0xC8,
         0x8B, 0x55, 0xFC,
         0x89, 0x50, 0x58,
         0xB8, 0x00, 0x00, 0x00, 0x00,
         0xEB, None,
         0x48, 0x8B, 0x45, 0xE0],
        _linux_helper_check1_replace,
    ),
    (
        [0x89, 0x45, 0xFC,
         0x83, 0x7D, 0xFC, 0x00,
         0x74, 0x11,
         0x48, 0x8B, 0x45, 0xC8,
         0x8B, 0x55, 0xFC,
         0x89, 0x50, 0x58,
         0xB8, 0x00, 0x00, 0x00, 0x00,
         0xEB, None,
         0xB8, 0x01, 0x00, 0x00, 0x00,
         0xC9, 0xC3],
        _linux_helper_check2_replace,
    ),
]


PATCHES_LINUX_REMOTE_MONITOR: "list[tuple[Pattern, Replacement]]" = [
    # Blackmagic / DaVinci Remote Monitor Linux ELF x86-64.
    (
        [0x89, 0x45, 0xFC,
         0x83, 0x7D, 0xFC, 0x00,
         0x74, 0x11,
         0x48, 0x8B, 0x45, 0xC8,
         0x8B, 0x55, 0xFC,
         0x89, 0x50, 0x58,
         0xB8, 0x00, 0x00, 0x00, 0x00,
         0xEB, None,
         0x48, 0x8B, 0x45, 0xE0],
        _linux_helper_check1_replace,
    ),
    (
        [0x89, 0x45, 0xFC,
         0x83, 0x7D, 0xFC, 0x00,
         0x74, 0x11,
         0x48, 0x8B, 0x45, 0xC8,
         0x8B, 0x55, 0xFC,
         0x89, 0x50, 0x58,
         0xB8, 0x00, 0x00, 0x00, 0x00,
         0xEB, None,
         0xB8, 0x01, 0x00, 0x00, 0x00,
         0xC9, 0xC3],
        _linux_helper_check2_replace,
    ),
]



# --------------------------------------------------------------------- PE version

def determine_version(data: bytes) -> "tuple[int, int, int]":
    """Read VS_FIXEDFILEINFO directly via its 0xFEEF04BD signature — unique
    enough that we don't need to walk the full PE header."""
    idx = data.find(b'\xBD\x04\xEF\xFE')
    if idx < 0:
        raise PatchError("Failed to parse PE header for main executable.")
    file_version_ms = struct.unpack_from('<I', data, idx + 8)[0]
    file_version_ls = struct.unpack_from('<I', data, idx + 12)[0]
    return (
        (file_version_ms >> 16) & 0xFFFF,
        file_version_ms & 0xFFFF,
        (file_version_ls >> 16) & 0xFFFF,
    )


# --------------------------------------------------------------------- inner-function patch

def patch_4func(data: bytearray) -> None:
    """Locate a specific call (matched by an outer pattern), follow its rel32
    target, and flip a few `je` -> `jne` inside the destination function
    (the Dolby Vision license validator). Indices `[0]` and `[0, 1, 2]` were
    chosen against v20.x and remain stable on v21.0.x.

    Raises PatchError without mutating `data` if any pattern is missing —
    we resolve every write site up front and only apply once everything is
    found, so partial mutation can't escape this function."""
    outer_pattern = [0xE8, None, None, None, None, 0x88, 0x83, None, None, None, None,
                     0x48, 0x8D, 0x4C, 0x24, None, 0xFF, 0x15]
    occs = find_all(data, outer_pattern)
    if len(occs) != 1:
        raise PatchError(
            f"patch_4func: outer pattern matched {len(occs)} times instead of 1."
        )

    call_addr = occs[0]
    rel32 = struct.unpack_from('<i', data, call_addr + 1)[0]
    fn_start = call_addr + 5 + rel32

    inner_patches = [
        ([0x84, 0xC0, 0x0F, 0x84], bytes([0x84, 0xC0, 0x0F, 0x85]), [0]),
        ([0x85, 0xDB, 0x0F, 0x84], bytes([0x85, 0xDB, 0x0F, 0x85]), [0, 1, 2]),
    ]
    # Resolve every write site up front so a missing inner pattern aborts
    # before any byte is touched.
    writes: list = []
    for sub_pat, repl, idxs in inner_patches:
        occs = find_all(data, sub_pat, fn_start, fn_start + 0x1000)
        for i in idxs:
            if i >= len(occs):
                raise PatchError(
                    f"patch_4func: inner pattern {bytes(b for b in sub_pat).hex()} "
                    f"index {i} not found (only {len(occs)} hits in window)."
                )
            writes.append((occs[i], repl))
    for off, repl in writes:
        data[off:off + len(repl)] = repl


# --------------------------------------------------------------------- license file & env var

def configure_license_file(resolve_path: str) -> None:
    """Write the fake RLM license next to Resolve.exe and set HKLM
    RLM_LICENSE=blackmagic.lic. Broadcasts WM_SETTINGCHANGE so newly-spawned
    processes pick up the env var without a logout/reboot."""
    lic_path = Path(resolve_path).with_name("blackmagic.lic")
    try:
        lic_path.write_text(LICENSE_FILE_CONTENTS)
        logger.info("wrote license file: %s", lic_path)
    except OSError as e:
        raise PatchError(f"failed to write license file: {e}") from e

    if winreg is None:
        logger.warning("winreg module unavailable (non-Windows); skipping registry configuration")
        return

    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"System\CurrentControlSet\Control\Session Manager\Environment",
            0, winreg.KEY_SET_VALUE,
        ) as key:
            winreg.SetValueEx(key, "RLM_LICENSE", 0, winreg.REG_SZ, "blackmagic.lic")
        logger.info("set HKLM RLM_LICENSE=blackmagic.lic")
    except OSError as e:
        raise PatchError(
            f"failed to set HKLM RLM_LICENSE (set it manually if needed): {e}"
        ) from e

    HWND_BROADCAST = 0xFFFF
    WM_SETTINGCHANGE = 0x001A
    SMTO_ABORTIFHUNG = 0x0002
    try:
        result = ctypes.c_long()
        ctypes.windll.user32.SendMessageTimeoutW(
            HWND_BROADCAST, WM_SETTINGCHANGE, 0,
            "Environment", SMTO_ABORTIFHUNG, 5000, ctypes.byref(result),
        )
        logger.info("broadcast WM_SETTINGCHANGE so new shells/Explorer see RLM_LICENSE")
    except (OSError, AttributeError) as e:
        logger.warning("failed to broadcast WM_SETTINGCHANGE: %s", e)


# --------------------------------------------------------------------- atomic write

def _atomic_replace_with_retry(target: Union[str, Path], payload: bytes,
                               action: str, hint: str = "") -> None:
    """Write via sibling .new + os.replace, retrying transient file locks."""
    target_path = Path(target)
    tmp_path = target_path.with_name(target_path.name + ".new")
    last_err: Optional[OSError] = None
    for attempt in range(5):
        try:
            with open(tmp_path, "wb") as f:
                f.write(payload)
            os.replace(tmp_path, target_path)
            return
        except OSError as e:
            last_err = e
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass
            logger.warning("%s attempt %d/5 failed: %s - retrying in 2s",
                           action, attempt + 1, e)
            time.sleep(2)
    extra = f" {hint}" if hint else ""
    raise PatchError(f"Unable to {action} {target_path}: {last_err}.{extra}")


def _keep_backup(binary_path: str, label: str) -> None:
    """Create a .bak copy of binary_path if one does not already exist.
    Preserves existing backups so repeated runs cannot overwrite clean originals."""
    bak_path = binary_path + ".bak"
    try:
        if Path(bak_path).exists():
            logger.info("%s backup already exists, keeping it: %s", label, bak_path)
        else:
            shutil.copy2(binary_path, bak_path)
            logger.info("backup written to %s", bak_path)
    except OSError as e:
        raise PatchError(f"Cannot create {label} backup: {e}") from e


def _read_clean_source(target_path: str, label: str) -> bytearray:
    """Read clean original binary data for patching. If <target>.bak exists,
    read from .bak (ensuring that re-patching an already-patched file never
    stacks or leaves obsolete patches behind). If no backup exists yet, create
    <target>.bak from target_path first."""
    bak_path = target_path + ".bak"
    if Path(bak_path).exists():
        logger.info("%s: using clean original backup as patch source (%s)", label, bak_path)
        try:
            with open(bak_path, "rb") as f:
                return bytearray(f.read())
        except OSError as e:
            raise PatchError(f"Cannot read {label} backup ({bak_path}): {e}") from e

    _keep_backup(target_path, label)
    try:
        with open(target_path, "rb") as f:
            return bytearray(f.read())
    except OSError as e:
        raise PatchError(f"Cannot read {label} binary ({target_path}): {e}") from e


# --------------------------------------------------------------------- main patch / restore

def _select_patches(version: "tuple[int, int, int]") -> "list[tuple[Pattern, Replacement]]":
    """Pick the patch table for the detected version. Only v21.x is supported."""
    major, minor, micro = version
    if major == 21:
        if (minor, micro) >= (0, 4):
            return PATCHES_21_2104
        return PATCHES_21
    raise PatchError(
        f"Resolve {major}.{minor}.{micro} is unsupported. "
        "This patcher only works on v21.x."
    )


def _apply_memory_patch_table(data: bytearray,
                              patches: "list[tuple[Pattern, Replacement]]",
                              label: str) -> bool:
    modified = False
    for i, (sig, replacement) in enumerate(patches):
        occs = find_all(data, sig)
        if not occs:
            logger.info("%s patch[%d]: no match (already patched or layout differs)", label, i)
            continue
        if len(occs) > 1:
            logger.warning("%s patch[%d]: matched %d times - skipping", label, i, len(occs))
            continue
        addr = occs[0]
        repl_bytes = replacement(data, addr, sig) if callable(replacement) else replacement
        logger.info("%s patch[%d]: applying at file offset 0x%08X (%d bytes)",
                    label, i, addr, len(repl_bytes))
        data[addr:addr + len(repl_bytes)] = repl_bytes
        modified = True
    return modified


def patch(resolve_path: str) -> None:
    """Patch Resolve.exe in place. Uses the clean .bak backup as the source
    if one exists, ensuring fresh patches are applied onto unpatched bytes
    and removing any obsolete/stale patches."""
    data = _read_clean_source(resolve_path, "Resolve.exe")

    version = determine_version(data)
    logger.info("detected Resolve version %d.%d.%d", *version)

    patches = _select_patches(version)
    modified = _apply_memory_patch_table(data, patches, "resolve")

    if version[0] == 21:
        try:
            patch_4func(data)
            logger.info("patch_4func: applied")
            modified = True
        except PatchError as e:
            logger.warning("patch_4func: %s (continuing anyway)", e)

    if not modified:
        raise PatchError(
            "No patches applied. Either this version is unsupported, the "
            "binary is already patched, or the byte offsets have shifted."
        )

    # bytearray is bytes-like — pass directly to avoid copying ~640 MB.
    _atomic_replace_with_retry(
        resolve_path,
        data,
        action="write",
        hint=("Make sure Resolve isn't running, close any Explorer window showing "
              "the folder, and run this script as Administrator."),
    )

    try:
        configure_license_file(resolve_path)
    except PatchError as e:
        # License file is best-effort; the binary patch already succeeded.
        logger.warning("configure_license_file failed: %s", e)


def restore(resolve_path: str) -> None:
    """Restore Resolve.exe from <path>.bak. Warns if the .bak's PE version
    differs from the current binary's, since restoring across an upgrade
    would replace the new install with the old one."""
    bak = resolve_path + ".bak"
    if not Path(bak).exists():
        raise PatchError(f"No backup found at {bak}")
    with open(bak, "rb") as f:
        bak_data = f.read()

    # Cross-version sanity check: same major.minor.micro on both sides. The
    # VS_FIXEDFILEINFO signature lives in the resource section near the end
    # of the binary, so we have to scan the whole file (transient — released
    # after this block).
    try:
        bak_version = determine_version(bak_data)
    except PatchError:
        bak_version = None
    cur_version = None
    try:
        with open(resolve_path, "rb") as f:
            cur_data = f.read()
        cur_version = determine_version(cur_data)
        del cur_data
    except (OSError, PatchError):
        pass
    if bak_version and cur_version and bak_version != cur_version:
        logger.warning(
            "backup is %d.%d.%d but current Resolve.exe is %d.%d.%d — "
            "restore would downgrade. Aborting. Delete %s manually if "
            "you really want to overwrite.",
            *bak_version, *cur_version, bak,
        )
        raise PatchError("backup version mismatch")

    _atomic_replace_with_retry(resolve_path, bak_data, action="restore")
    logger.info("restored %s from %s", resolve_path, bak)


# --------------------------------------------------------------------- state detection (no writes)

def state_of_resolve(resolve_path: str) -> str:
    """Classify a Resolve.exe install without modifying it.

    Returns one of: MISSING (file not present), UNSUPPORTED (version not v20/v21),
    UNPATCHED, PATCHED, PARTIAL, or UNKNOWN."""
    if not Path(resolve_path).exists():
        return "MISSING"
    try:
        with open(resolve_path, "rb") as f:
            data = f.read()
        version = determine_version(data)
    except (OSError, PatchError):
        return "UNSUPPORTED"
    try:
        patches = _select_patches(version)
    except PatchError:
        return "UNSUPPORTED"
    return _classify_patch_state_data(data, patches)


def state_of_fusion(dll_path: str) -> str:
    """Classify a fusionsystem.dll install without modifying it.

    Returns one of: MISSING / UNPATCHED / PATCHED / PARTIAL / UNKNOWN."""
    if not Path(dll_path).exists():
        return "MISSING"
    try:
        with open(dll_path, "rb") as f:
            data = f.read()
    except OSError:
        return "UNSUPPORTED"
    return _classify_patch_state_data(data, PATCHES_FUSION)


# --------------------------------------------------------------------- fusion patch / restore / locate

def patch_fusion(dll_path: str) -> None:
    """Patch fusionsystem.dll in place. Uses the clean .bak backup as the source
    if one exists, ensuring fresh patches are applied onto unpatched bytes."""
    data = _read_clean_source(dll_path, "Fusion")

    modified = _apply_memory_patch_table(data, PATCHES_FUSION, "fusion")

    if not modified:
        raise PatchError(
            f"No Fusion patches applied to {dll_path}. Either it's already "
            "patched, this Fusion version is unsupported, or offsets shifted."
        )

    _atomic_replace_with_retry(dll_path, bytes(data), action="write")


def restore_fusion(dll_path: str) -> None:
    """Restore fusionsystem.dll from <path>.bak. Unlike Resolve.exe there is
    no embedded VS_FIXEDFILEINFO version on the DLL itself, so we skip the
    cross-version sanity check — version is encoded in the install path and
    each Fusion version has its own .bak."""
    bak = dll_path + ".bak"
    if not Path(bak).exists():
        raise PatchError(f"No backup found at {bak}")
    with open(bak, "rb") as f:
        bak_data = f.read()
    _atomic_replace_with_retry(dll_path, bak_data, action="restore")
    logger.info("restored %s from %s", dll_path, bak)


def locate_fusion() -> "list[str]":
    """Return all Fusion fusionsystem.dll paths actually present on disk."""
    return [p for p in DEFAULT_FUSION_PATHS if Path(p).exists()]


def _kill_fusion() -> None:
    """Best-effort kill of any running Fusion processes that would lock the
    DLL. Silently ignores missing processes."""
    for image in ("Fusion.exe", "FusionServer.exe"):
        try:
            subprocess.run(["taskkill", "/F", "/IM", image],
                           capture_output=True, text=True)
        except OSError:
            pass


# --------------------------------------------------------------------- easydcp / remote_monitor patch / restore
#
# Both targets share Fusion's lifecycle (small auxiliary file, no version
# detection on the binary itself, identical atomic-write + .bak pattern). The
# helpers below are intentionally near-clones of patch_fusion/restore_fusion
# so changes to one are easy to mirror to the other. They're invoked only
# through the Resolve bundle (`patch_resolve_bundle` / `restore_resolve_bundle`),
# which finds the companion paths by scanning Resolve's install folder.

def state_of_easydcp(dll_path: str) -> str:
    """Classify fraunhoferdcp.dll without modifying it.
    Returns MISSING / UNPATCHED / PATCHED / PARTIAL / UNKNOWN / UNSUPPORTED."""
    if not Path(dll_path).exists():
        return "MISSING"
    try:
        with open(dll_path, "rb") as f:
            data = f.read()
    except OSError:
        return "UNSUPPORTED"
    return _classify_patch_state_data(data, PATCHES_EASYDCP)


def state_of_remote_monitor(exe_path: str) -> str:
    """Classify DaVinci Remote Monitor.exe without modifying it."""
    if not Path(exe_path).exists():
        return "MISSING"
    try:
        with open(exe_path, "rb") as f:
            data = f.read()
    except OSError:
        return "UNSUPPORTED"
    return _classify_patch_state_data(data, PATCHES_REMOTE_MONITOR)


def patch_easydcp(dll_path: str) -> None:
    """Patch fraunhoferdcp.dll in place. Uses the clean .bak backup as the source
    if one exists, ensuring fresh patches are applied onto unpatched bytes."""
    data = _read_clean_source(dll_path, "EasyDCP")

    modified = _apply_memory_patch_table(data, PATCHES_EASYDCP, "easydcp")

    if not modified:
        raise PatchError(
            f"No EasyDCP patches applied to {dll_path}. Either it's already "
            "patched, the DLL version is unsupported, or offsets shifted."
        )

    _atomic_replace_with_retry(dll_path, bytes(data), action="write")


def restore_easydcp(dll_path: str) -> None:
    """Restore fraunhoferdcp.dll from <path>.bak."""
    bak = dll_path + ".bak"
    if not Path(bak).exists():
        raise PatchError(f"No backup found at {bak}")
    with open(bak, "rb") as f:
        bak_data = f.read()
    _atomic_replace_with_retry(dll_path, bak_data, action="restore")
    logger.info("restored %s from %s", dll_path, bak)


def patch_remote_monitor(exe_path: str) -> None:
    """Patch DaVinci Remote Monitor.exe in place. Uses the clean .bak backup as the source
    if one exists, ensuring fresh patches are applied onto unpatched bytes."""
    data = _read_clean_source(exe_path, "Remote Monitor")

    modified = _apply_memory_patch_table(data, PATCHES_REMOTE_MONITOR, "remote_monitor")

    if not modified:
        raise PatchError(
            f"No Remote Monitor patches applied to {exe_path}. Either it's "
            "already patched or the binary version is unsupported."
        )

    _atomic_replace_with_retry(exe_path, bytes(data), action="write")


def restore_remote_monitor(exe_path: str) -> None:
    """Restore DaVinci Remote Monitor.exe from <path>.bak."""
    bak = exe_path + ".bak"
    if not Path(bak).exists():
        raise PatchError(f"No backup found at {bak}")
    with open(bak, "rb") as f:
        bak_data = f.read()
    _atomic_replace_with_retry(exe_path, bak_data, action="restore")
    logger.info("restored %s from %s", exe_path, bak)


def _kill_remote_monitor() -> None:
    """Best-effort kill of any running Remote Monitor process. Tries both
    image names (v21 beta renamed the binary)."""
    for image in ("Blackmagic Remote Monitor.exe", "DaVinci Remote Monitor.exe"):
        try:
            subprocess.run(["taskkill", "/F", "/IM", image],
                           capture_output=True, text=True)
        except OSError:
            pass


# --------------------------------------------------------------------- resolve bundle (resolve.exe + companions)
#
# A "Resolve install" is really three binaries that live in the same folder:
#   * Resolve.exe                   — main app + license / DV
#   * fraunhoferdcp.dll             — EasyDCP cinema-package watermark gates
#   * DaVinci Remote Monitor.exe    — Studio-feature gate in the viewer app
#
# Each has its own .bak and its own patch table. Exposing them as separate
# rows in the menu was confusing, so the helpers below bundle all three under
# the single "Resolve" target. Each companion is patched only if the file
# exists, so portable / partial installs still work.

def _resolve_companion_paths(resolve_path: Optional[str]) -> "list[tuple[str, str, Callable, Callable, Callable]]":
    """For a given resolve_path, return [(label, companion_path, state_fn,
    patch_fn, restore_fn), ...] for each companion file that actually exists.
    Empty list if the install dir can't be derived.

    Remote Monitor has two possible filenames depending on the Resolve major
    version (`Blackmagic Remote Monitor.exe` in v21+, `DaVinci Remote
    Monitor.exe` in pre-v21). We try both in order and stop on the first
    match — a stray .bak / .exe with the old name shouldn't cause us to
    double-patch."""
    if not resolve_path:
        return []
    install_dir = Path(resolve_path).parent
    result: list[tuple[str, str, Callable, Callable, Callable]] = []

    easydcp = install_dir / "fraunhoferdcp.dll"
    if easydcp.exists():
        result.append((
            "EasyDCP (fraunhoferdcp.dll)",
            str(easydcp),
            state_of_easydcp, patch_easydcp, restore_easydcp,
        ))

    for rm_name in ("Blackmagic Remote Monitor.exe", "DaVinci Remote Monitor.exe"):
        rm_path = install_dir / rm_name
        if rm_path.exists():
            result.append((
                f"Remote Monitor ({rm_name})",
                str(rm_path),
                state_of_remote_monitor, patch_remote_monitor, restore_remote_monitor,
            ))
            break

    return result


def state_of_resolve_bundle(resolve_path: Optional[str]) -> str:
    """Aggregate state across Resolve.exe + every installed companion.
    The main binary's state wins on MISSING/UNSUPPORTED. Otherwise:
      * all PATCHED  -> PATCHED
      * all UNPATCHED -> UNPATCHED
      * mixed         -> PARTIAL
    """
    if not resolve_path:
        return "MISSING"
    main = state_of_resolve(resolve_path)
    if main in ("MISSING", "UNSUPPORTED"):
        return main
    companion_states = [s(p) for _, p, s, _, _ in _resolve_companion_paths(resolve_path)]
    # MISSING for a companion is treated as "not part of this install" rather
    # than a failure, so we drop it from aggregation.
    states = [main] + [s for s in companion_states if s != "MISSING"]
    if all(s == "PATCHED" for s in states):
        return "PATCHED"
    if all(s == "UNPATCHED" for s in states):
        return "UNPATCHED"
    return "PARTIAL"


def _kill_resolve_bundle() -> None:
    """Kill every process that locks any of the bundle files."""
    _kill_resolve()
    _kill_remote_monitor()


def patch_resolve_bundle(resolve_path: str) -> None:
    """Patch Resolve.exe + every installed companion. Companion failures are
    logged but don't abort the main patch — partial coverage is better than
    none, and Resolve.exe is the file that matters most."""
    patch(resolve_path)
    for label, companion_path, _state_fn, patch_fn, _restore_fn in \
            _resolve_companion_paths(resolve_path):
        try:
            logger.info("attempting to patch %s: %s", label, companion_path)
            patch_fn(companion_path)
            logger.info("successfully patched %s", label)
        except PatchError as e:
            logger.warning("%s patch skipped: %s", label, e)


def restore_resolve_bundle(resolve_path: str) -> None:
    """Restore Resolve.exe + every installed companion from .bak. Same
    best-effort behaviour as patch_resolve_bundle."""
    restore(resolve_path)
    for label, companion_path, _state_fn, _patch_fn, restore_fn in \
            _resolve_companion_paths(resolve_path):
        try:
            logger.info("attempting to restore %s: %s", label, companion_path)
            restore_fn(companion_path)
            logger.info("successfully restored %s", label)
        except PatchError as e:
            logger.warning("%s restore skipped: %s", label, e)


# --------------------------------------------------------------------- linux patch / restore / locate

def detect_linux_build(data: bytes) -> "tuple[int, int, int, int]":
    """Extract Resolve Linux version from embedded build strings.

    Observed formats include:
      - 21.1.1.0010 (release/21.1.1\x00\x00DaVinci Resolve Studio\x00\x0021.1.1\x00\x0021.1.1.0010)
      - 21.1.0.0014 (release/21.1\x0021.1.0\x0021.1.0.0014)
      - 21.0.4.0005 (release/21.0.4\x0021.0.4.0005)
      - 21.0.0b.0028
      - 20.3.2.0009 Linux/Clang x86_64
    """
    patterns = (
        re.compile(rb"release/\d+(?:\.\d+)+(?:[^\x00]*\x00+)*?(\d+)\.(\d+)\.(\d+)\.(\d+)\x00"),
        re.compile(rb"release/\d+(?:\.\d+)+\x00(?:\d+\.\d+\.\d+\x00)?(\d+)\.(\d+)\.(\d+)\.(\d+)\x00"),
        re.compile(rb"release/(\d+)\.(\d+)\.(\d+)\x00(\d+)\.(\d+)\.(\d+)\.(\d+)\x00"),
        re.compile(rb"(?:release/|\x00)(2[0-9])\.(\d+)\.(\d+)(?:b|\.)(\d+)\x00"),
        re.compile(rb"(\d+)\.(\d+)\.(\d+)\.(\d+)\s+Linux/Clang"),
        re.compile(rb"\x00(2[0-9])\.(\d+)\.(\d+)\.(\d+)\x00"),
    )
    for pattern in patterns:
        m = pattern.search(data)
        if m:
            if len(m.groups()) == 7:
                return (int(m.group(4)), int(m.group(5)),
                        int(m.group(6)), int(m.group(7)))
            return (int(m.group(1)), int(m.group(2)),
                    int(m.group(3)), int(m.group(4)))
    return (0, 0, 0, 0)


def detect_linux_version(data: bytes) -> "tuple[int, int, int]":
    major, minor, micro, _build = detect_linux_build(data)
    return major, minor, micro


def _linux_version_label(version: "tuple[int, ...]") -> str:
    if len(version) >= 4 and version[3]:
        return f"{version[0]}.{version[1]}.{version[2]} build {version[3]:04d}"
    if len(version) >= 3:
        return f"{version[0]}.{version[1]}.{version[2]}"
    return "0.0.0"


def select_linux_patches(version: "tuple[int, ...]") -> "list[tuple[Pattern, Replacement]]":
    major = version[0]
    minor = version[1] if len(version) >= 2 else 0
    micro = version[2] if len(version) >= 3 else 0
    build = version[3] if len(version) >= 4 else 0
    if major == 21:
        if (minor, micro) >= (0, 4) or build >= 100 or build in (5, 10, 14):
            return PATCHES_LINUX_21_2104
        if build >= 33:
            return PATCHES_LINUX_21_B0033
        return PATCHES_LINUX_21
    raise PatchError(
        f"Resolve {_linux_version_label(version)} is not supported on Linux. "
        "Only v21.x is supported by the Linux patcher."
    )


def _atomic_write_linux(target: str, payload: bytes, action: str) -> None:
    tmp = target + ".new"
    try:
        original_mode = os.stat(target).st_mode
    except OSError:
        original_mode = None

    last_err: Optional[OSError] = None
    for attempt in range(5):
        try:
            with open(tmp, "wb") as f:
                f.write(payload)
            if original_mode is not None:
                os.chmod(tmp, original_mode)
            os.replace(tmp, target)
            return
        except OSError as e:
            last_err = e
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
            logger.warning("%s attempt %d/5 failed: %s - retrying in 2s",
                           action, attempt + 1, e)
            time.sleep(2)
    raise PatchError(f"Unable to {action} Linux target after 5 attempts: {last_err}")


def _linux_install_root(resolve_path: str) -> Path:
    """Return the Resolve install root for a Linux binary path.

    `/opt/resolve/bin/resolve` -> `/opt/resolve`. If a custom path does not
    follow that layout, use the binary's directory.
    """
    p = Path(resolve_path).resolve()
    if p.parent.name == "bin":
        return p.parent.parent
    return p.parent


def _write_linux_license(root: Path, contents: str, label: str) -> None:
    lic_dir = root / ".license"
    lic_path = lic_dir / "blackmagic.lic"
    try:
        lic_dir.mkdir(parents=True, exist_ok=True)
        lic_path.write_text(contents, encoding="utf-8")
        logger.info("wrote %s license file: %s", label, lic_path)
    except OSError as e:
        raise PatchError(f"failed to write {label} license file {lic_path}: {e}") from e


def configure_linux_license_file(resolve_path: str) -> None:
    _write_linux_license(
        _linux_install_root(resolve_path),
        LINUX_LICENSE_FILE_CONTENTS,
        "Linux Resolve",
    )


def configure_linux_fusion_license_file(lib_path: str) -> None:
    _write_linux_license(
        Path(lib_path).resolve().parent,
        LINUX_FUSION_LICENSE_FILE_CONTENTS,
        "Linux Fusion",
    )


def _linux_read_binary(binary_path: str, label: str) -> bytearray:
    try:
        with open(binary_path, "rb") as f:
            return bytearray(f.read())
    except OSError as e:
        raise PatchError(f"Cannot read {label} binary: {e}") from e


def _linux_keep_backup(binary_path: str, label: str) -> None:
    _keep_backup(binary_path, label)


def _validate_linux_restore_backup(target: str,
                                   bak_data: bytes,
                                   label: str,
                                   check_version: bool = False) -> None:
    if not bak_data.startswith(b"\x7fELF"):
        raise PatchError(f"{label} backup is not an ELF binary; refusing restore")
    try:
        target_stat = os.stat(target)
        with open(target, "rb") as f:
            target_magic = f.read(4)
    except OSError as e:
        raise PatchError(f"Cannot inspect current {label} binary before restore: {e}") from e
    if target_magic != b"\x7fELF":
        raise PatchError(f"current {label} target is not an ELF binary; refusing restore")
    if len(bak_data) != target_stat.st_size:
        raise PatchError(
            f"{label} backup size mismatch: backup has {len(bak_data):,} bytes, "
            f"current target has {target_stat.st_size:,} bytes. Refusing cross-version restore."
        )
    if check_version:
        bak_version = detect_linux_build(bak_data)
        cur_version = detect_linux_build(_linux_read_binary(target, label))
        if bak_version != (0, 0, 0, 0) and cur_version != (0, 0, 0, 0) and bak_version != cur_version:
            raise PatchError(
                f"{label} backup version mismatch: backup is "
                f"{_linux_version_label(bak_version)}, current is "
                f"{_linux_version_label(cur_version)}"
            )


def _linux_read_clean_source(binary_path: str, label: str) -> bytearray:
    """Read clean original binary data for patching on Linux. Uses .bak if it
    exists, otherwise creates .bak first."""
    bak_path = binary_path + ".bak"
    if Path(bak_path).exists():
        logger.info("%s: using clean original backup as patch source (%s)", label, bak_path)
        try:
            with open(bak_path, "rb") as f:
                return bytearray(f.read())
        except OSError as e:
            raise PatchError(f"Cannot read {label} backup ({bak_path}): {e}") from e

    _linux_keep_backup(binary_path, label)
    return _linux_read_binary(binary_path, label)


def _restore_linux_target(path: str, label: str, check_version: bool) -> None:
    bak = path + ".bak"
    if not Path(bak).exists():
        raise PatchError(f"No backup found at {bak}")
    try:
        with open(bak, "rb") as f:
            bak_data = f.read()
    except OSError as e:
        raise PatchError(f"Cannot read {label} backup: {e}") from e
    _validate_linux_restore_backup(path, bak_data, label, check_version=check_version)
    _atomic_write_linux(path, bak_data, action="restore")
    logger.info("restored %s from %s", path, bak)


def state_of_linux_remote_monitor(exe_path: str) -> str:
    if not Path(exe_path).exists():
        return "MISSING"
    try:
        return _scan_patch_state_file(exe_path, PATCHES_LINUX_REMOTE_MONITOR, chunk_size=8 * 1024 * 1024)
    except (OSError, PatchError):
        return "UNSUPPORTED"


def patch_linux_remote_monitor(exe_path: str) -> None:
    data = _linux_read_clean_source(exe_path, "Linux Remote Monitor")
    modified = _apply_memory_patch_table(data, PATCHES_LINUX_REMOTE_MONITOR, "linux remote monitor")
    if not modified:
        raise PatchError(
            "No Linux Remote Monitor patches applied. The binary may already be patched, "
            "or the byte layout differs from the supported version."
        )
    _atomic_write_linux(exe_path, bytes(data), action="patch")
    logger.info("patched Linux Remote Monitor successfully")


def restore_linux_remote_monitor(exe_path: str) -> None:
    _restore_linux_target(exe_path, "Linux Remote Monitor", check_version=False)


def _linux_resolve_companion_paths(resolve_path: Optional[str]) -> "list[tuple[str, str, Callable, Callable, Callable]]":
    if not resolve_path:
        return []
    bin_dir = Path(resolve_path).resolve().parent
    result: list[tuple[str, str, Callable, Callable, Callable]] = []
    for rm_name in ("Blackmagic Remote Monitor", "DaVinci Remote Monitor"):
        rm_path = bin_dir / rm_name
        if rm_path.exists():
            result.append((
                f"Remote Monitor ({rm_name})",
                str(rm_path),
                state_of_linux_remote_monitor,
                patch_linux_remote_monitor,
                restore_linux_remote_monitor,
            ))
            break
    return result


def patch_linux(resolve_path: str) -> None:
    data = _linux_read_clean_source(resolve_path, "Linux Resolve")
    version = detect_linux_build(bytes(data))
    logger.info("detected Linux Resolve version %s", _linux_version_label(version))
    patches = select_linux_patches(version)
    modified = _apply_memory_patch_table(data, patches, "linux resolve")

    if not modified:
        raise PatchError(
            "No Linux patches applied. The binary may already be patched, "
            "or the byte layout differs from the supported version."
        )

    _atomic_write_linux(resolve_path, bytes(data), action="patch")
    configure_linux_license_file(resolve_path)
    logger.info("patched Linux Resolve successfully")

    for label, companion_path, _state_fn, patch_fn, _restore_fn in \
            _linux_resolve_companion_paths(resolve_path):
        try:
            logger.info("attempting to patch %s: %s", label, companion_path)
            patch_fn(companion_path)
            logger.info("successfully patched %s", label)
        except PatchError as e:
            logger.warning("%s patch skipped: %s", label, e)


def patch_linux_fusion(lib_path: str) -> None:
    data = _linux_read_clean_source(lib_path, "Linux Fusion")
    modified = _apply_memory_patch_table(data, PATCHES_LINUX_FUSION, "linux fusion")
    if not modified:
        raise PatchError(
            "No Linux Fusion patches applied. The library may already be patched, "
            "or the byte layout differs from the supported version."
        )
    _atomic_write_linux(lib_path, bytes(data), action="patch")
    configure_linux_fusion_license_file(lib_path)
    logger.info("patched Linux Fusion successfully")


def restore_linux(resolve_path: str) -> None:
    _restore_linux_target(resolve_path, "Linux Resolve", check_version=True)
    for label, companion_path, _state_fn, _patch_fn, restore_fn in \
            _linux_resolve_companion_paths(resolve_path):
        try:
            logger.info("attempting to restore %s: %s", label, companion_path)
            restore_fn(companion_path)
            logger.info("successfully restored %s", label)
        except PatchError as e:
            logger.warning("%s restore skipped: %s", label, e)


def restore_linux_fusion(lib_path: str) -> None:
    _restore_linux_target(lib_path, "Linux Fusion", check_version=False)


def locate_linux_resolve() -> str:
    for candidate in DEFAULT_LINUX_RESOLVE_PATHS:
        p = Path(candidate)
        if p.is_symlink():
            real = p.resolve()
            if real.exists():
                logger.info("using resolved symlink: %s -> %s", candidate, real)
                return str(real)
        elif p.exists():
            return candidate
    raise PatchError(
        "Could not locate the Linux Resolve binary. "
        f"Searched: {', '.join(DEFAULT_LINUX_RESOLVE_PATHS)}. "
        "Pass --path explicitly."
    )


def locate_linux_fusion() -> "list[str]":
    """Return standalone Fusion installs.

    `/opt/resolve/libs/Fusion/libfusionsystem.so` belongs to Resolve's bundled
    Fusion page, not the standalone Fusion app. Showing it next to
    `/opt/BlackmagicDesign/Fusion20/...` is confusing and can lead to patching
    the wrong target. Keep it as a fallback only when no standalone Fusion
    install is present.
    """
    seen: set[str] = set()
    standalone: list[str] = []
    bundled: list[str] = []
    for candidate in DEFAULT_LINUX_FUSION_PATHS:
        p = Path(candidate)
        if not p.exists():
            continue
        real = str(p.resolve())
        if real in seen:
            continue
        seen.add(real)
        if "/opt/resolve/libs/Fusion/" in candidate:
            bundled.append(candidate)
        else:
            standalone.append(candidate)
    return standalone if standalone else bundled


def state_of_linux_resolve(resolve_path: str) -> str:
    if not Path(resolve_path).exists():
        return "MISSING"
    try:
        data = _linux_read_binary(resolve_path, "Linux Resolve")
        version = detect_linux_build(bytes(data))
        if version == (0, 0, 0, 0):
            return "UNSUPPORTED"
        patches = select_linux_patches(version)
        main_state = _classify_patch_state_data(bytes(data), patches)
    except (OSError, PatchError):
        return "UNSUPPORTED"

    if main_state in ("MISSING", "UNSUPPORTED"):
        return main_state

    companion_states = [s(p) for _, p, s, _, _ in _linux_resolve_companion_paths(resolve_path)]
    states = [main_state] + [s for s in companion_states if s != "MISSING"]
    if all(s == "PATCHED" for s in states):
        return "PATCHED"
    if all(s == "UNPATCHED" for s in states):
        return "UNPATCHED"
    return "PARTIAL"


def state_of_linux_fusion(lib_path: str) -> str:
    if not Path(lib_path).exists():
        return "MISSING"
    try:
        return _scan_patch_state_file(lib_path, PATCHES_LINUX_FUSION, chunk_size=8 * 1024 * 1024)
    except (OSError, PatchError):
        return "UNSUPPORTED"


def _require_root() -> None:
    if not hasattr(os, "geteuid"):
        logger.error("root check is only available on POSIX systems.")
        raise SystemExit(1)
    if os.geteuid() != 0:
        logger.error("This script must be run as root (use sudo).")
        raise SystemExit(1)


# --------------------------------------------------------------------- locate

def _extract_path_from_command(cmd: str) -> Optional[str]:
    """Extract executable path from a shell open command string like:
        "C:\\Program Files\\...\\Resolve.exe" "%1"
    or
        C:\\Path\\Resolve.exe %1
    """
    cmd = cmd.strip()
    if cmd.startswith('"'):
        end_quote = cmd.find('"', 1)
        if end_quote != -1:
            candidate = cmd[1:end_quote]
            if Path(candidate).exists():
                return candidate
    parts = cmd.split()
    if parts and Path(parts[0]).exists():
        return parts[0]
    return None


def _path_from_shellopen() -> Optional[str]:
    """Look up Resolve.exe via the registered shell open-command for known
    Resolve file types across HKCR, HKCU, and HKLM."""
    if winreg is None:
        return None

    for typ in SHELLOPEN_EXTENSION_KEYS:
        # Check HKCR (merged classes)
        try:
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT,
                rf"{typ}\shell\open\command",
            ) as key:
                value, _ = winreg.QueryValueEx(key, "")
                path = _extract_path_from_command(value)
                if path:
                    return path
        except OSError:
            pass

        # Check HKCU and HKLM
        for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            try:
                with winreg.OpenKey(
                    root,
                    rf"Software\Classes\{typ}\shell\open\command",
                ) as key:
                    value, _ = winreg.QueryValueEx(key, "")
                    path = _extract_path_from_command(value)
                    if path:
                        return path
            except OSError:
                continue
    return None


def _path_from_install_location() -> Optional[str]:
    """Look up Resolve.exe via the official InstallLocation registry key."""
    if winreg is None:
        return None
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(
                root,
                r"SOFTWARE\Blackmagic Design\DaVinci Resolve",
            ) as key:
                loc, _ = winreg.QueryValueEx(key, "InstallLocation")
                if loc:
                    candidate = Path(loc) / "Resolve.exe"
                    if candidate.exists():
                        return str(candidate)
        except OSError:
            continue
    return None


def locate() -> str:
    """Find Resolve.exe via the registry, falling back to the standard install path."""
    path = _path_from_install_location()
    if path is not None:
        logger.info("Resolve found via InstallLocation regkey: %s", path)
        return path
    path = _path_from_shellopen()
    if path is not None:
        logger.info("Resolve found via shellopen regkey: %s", path)
        return path
    if Path(DEFAULT_PATH).exists():
        return DEFAULT_PATH
    raise PatchError("Resolve could not be located.")


# --------------------------------------------------------------------- CLI

def _build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Patch DaVinci Resolve/Fusion on the current OS. "
                    "Windows: Resolve.exe + Fusion DLLs. Linux: Resolve + Fusion ELF targets.",
    )
    p.add_argument("--restore", action="store_true",
                   help="restore from .bak files and exit (applies to all selected targets)")
    p.add_argument("--path", default=None,
                   help="explicit path to Resolve binary (skips auto-locate)")
    p.add_argument("--targets", default=None,
                   help="non-interactive: comma-separated targets "
                        "(resolve, fusion, all). 'resolve' patches Resolve.exe + "
                        "fraunhoferdcp.dll + DaVinci Remote Monitor.exe together. "
                        "If omitted on a TTY, an interactive menu is shown instead.")
    return p


def _require_admin() -> None:
    """Exit immediately if not running as Administrator on Windows.
    On non-Windows, fail fast — the patcher uses winreg/ctypes-windll
    further down and would crash anyway."""
    if sys.platform != "win32":
        logger.error("This script only runs on Windows.")
        raise SystemExit(1)
    if not ctypes.windll.shell32.IsUserAnAdmin():
        logger.error("This script must be run as Administrator.")
        logger.error("Right-click your terminal/PowerShell and choose 'Run as administrator'.")
        raise SystemExit(1)


def _kill_resolve() -> bool:
    """Kill any running Resolve.exe. Returns True if a process was killed,
    False if there was nothing to kill or the kill failed."""
    try:
        result = subprocess.run(
            ["taskkill", "/F", "/IM", "Resolve.exe"],
            capture_output=True,
            text=True,
        )
    except OSError as e:
        logger.debug("taskkill unavailable: %s", e)
        return False
    if result.returncode == 0:
        logger.info("killed running Resolve.exe")
        return True
    # taskkill returns 128 when no matching process exists — that's fine.
    return False


def _parse_targets(spec: str) -> "set[str]":
    """`--targets all` expands to {resolve, fusion}; comma-separated names are
    validated. The "resolve" target is a bundle covering Resolve.exe plus
    its companion binaries (fraunhoferdcp.dll, DaVinci Remote Monitor.exe).
    Unknown names raise SystemExit via the arg parser convention."""
    valid = {"resolve", "fusion"}
    raw = {t.strip().lower() for t in spec.split(",") if t.strip()}
    if "all" in raw:
        return valid
    bad = raw - valid
    if bad:
        logger.error(
            "unknown --targets value(s): %s (valid: %s, all)",
            ", ".join(sorted(bad)), ", ".join(sorted(valid)),
        )
        raise SystemExit(2)
    return raw


# --------------------------------------------------------------------- interactive menu

# A target row in the menu: (kind, label, path, state).
# kind ∈ {"resolve", "fusion21"} — used by _execute to dispatch.
def _build_target_rows(resolve_path: Optional[str]) -> "list[tuple[str, str, str, str]]":
    rows: list[tuple[str, str, str, str]] = []
    if resolve_path is not None:
        rows.append(("resolve",
                     "DaVinci Resolve (Resolve.exe + companion files)",
                     resolve_path,
                     state_of_resolve_bundle(resolve_path)))
    else:
        rows.append(("resolve",
                     "DaVinci Resolve (Resolve.exe + companion files)",
                     DEFAULT_PATH,
                     "MISSING"))
    for fp in DEFAULT_FUSION_PATHS:
        rows.append(("fusion21",
                     "Fusion Studio 21 (fusionsystem.dll)",
                     fp,
                     state_of_fusion(fp)))
    return rows


# ----- arrow-key UI primitives (Windows msvcrt, Linux termios, ANSI redraw) -----

def _enable_ansi() -> None:
    """Best-effort: turn on ENABLE_VIRTUAL_TERMINAL_PROCESSING so ANSI escape
    codes (cursor up, line clear) are interpreted instead of printed literally.
    Modern Windows terminals have this on by default; this is a safety net for
    older console hosts."""
    try:
        STD_OUTPUT_HANDLE = -11
        ENABLE_VT = 0x0004
        kernel32 = ctypes.windll.kernel32
        h = kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
        mode = ctypes.c_ulong()
        if kernel32.GetConsoleMode(h, ctypes.byref(mode)):
            kernel32.SetConsoleMode(h, mode.value | ENABLE_VT)
    except Exception:
        pass


def _read_key() -> str:
    """Block on a single keypress. Returns a normalized token:
        'up' / 'down' / 'left' / 'right' / 'enter' / 'space' / 'escape',
    or a lowercased single character (e.g. 'a', 'q'), or '' for unknown.
    Raises KeyboardInterrupt on Ctrl-C."""
    if sys.platform == "win32" and msvcrt is not None:
        ch = msvcrt.getch()
        # Special keys are two-byte sequences: \xe0 (or \x00) followed by a code.
        if ch in (b'\xe0', b'\x00'):
            ch2 = msvcrt.getch()
            return {b'H': 'up', b'P': 'down', b'K': 'left', b'M': 'right'}.get(ch2, '')
        if ch in (b'\r', b'\n'):
            return 'enter'
        if ch == b'\x1b':
            return 'escape'
        if ch == b' ':
            return 'space'
        if ch == b'\x03':
            raise KeyboardInterrupt
        try:
            return ch.decode('utf-8', errors='replace').lower()
        except UnicodeDecodeError:
            return ''

    if termios is None or tty is None:
        # Last-resort fallback for unusual environments.
        try:
            ch = input().strip().lower()
        except (EOFError, KeyboardInterrupt):
            return 'escape'
        return 'enter' if not ch else ch[:1]

    fd = sys.stdin.fileno()
    old_settings = termios.tcgetattr(fd)
    try:
        tty.setraw(fd)
        ch = sys.stdin.read(1)
        if ch == '\x03':
            raise KeyboardInterrupt
        if ch in ('\r', '\n'):
            return 'enter'
        if ch == ' ':
            return 'space'
        if ch == '\x1b':
            ch2 = sys.stdin.read(1)
            if ch2 == '[':
                ch3 = sys.stdin.read(1)
                return {'A': 'up', 'B': 'down', 'C': 'right', 'D': 'left'}.get(ch3, '')
            return 'escape'
        return ch.lower()
    finally:
        termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)


class _RedrawRegion:
    """In-place menu redraw via ANSI cursor-up + line-clear. The first
    `render()` just prints; subsequent calls overwrite the previous render
    by going up N lines, clearing each, and re-printing.

    Multi-line strings (i.e. items in `lines` that contain '\\n') are split
    so each visual line gets its own `\\x1b[2K` clear and is counted in
    `self.lines`. Without this split, the cursor-up would undercount and
    drift downward by one line per redraw — leaving a growing gap above
    the menu."""

    def __init__(self) -> None:
        self.lines: int = 0

    def render(self, lines: "list[str]") -> None:
        flat: list[str] = []
        for line in lines:
            flat.extend(line.split('\n'))
        if self.lines:
            sys.stdout.write(f'\r\x1b[{self.lines}A')
        for line in flat:
            sys.stdout.write('\x1b[2K' + line + '\n')
        self.lines = len(flat)
        sys.stdout.flush()


def _arrow_single_select(header: str, options: "list[str]",
                         footer: str = "") -> Optional[int]:
    """Single-select arrow-key menu. Returns the chosen index, or None
    if the user pressed q/escape."""
    cursor = 0
    region = _RedrawRegion()
    while True:
        lines = [header, ""]
        for i, opt in enumerate(options):
            mark = ">" if i == cursor else " "
            lines.append(f" {mark} {opt}")
        if footer:
            lines.extend(["", footer])
        region.render(lines)
        try:
            key = _read_key()
        except KeyboardInterrupt:
            return None
        if key == 'up' and cursor > 0:
            cursor -= 1
        elif key == 'down' and cursor < len(options) - 1:
            cursor += 1
        elif key == 'enter':
            return cursor
        elif key in ('escape', 'q'):
            return None


def _arrow_multi_select(header: str, options: "list[str]",
                        preselected: "Optional[set[int]]" = None,
                        footer: str = "") -> "Optional[set[int]]":
    """Multi-select arrow-key menu. SPACE toggles the current row, A toggles
    all-vs-none, ENTER confirms. Returns set of chosen indices, or None on
    cancel (q/escape) or if the user confirms with nothing selected."""
    cursor = 0
    selected: set[int] = set(preselected) if preselected else set()
    region = _RedrawRegion()
    while True:
        lines = [header, ""]
        for i, opt in enumerate(options):
            arrow = ">" if i == cursor else " "
            box = "[x]" if i in selected else "[ ]"
            lines.append(f" {arrow} {box} {opt}")
        if footer:
            lines.extend(["", footer])
        region.render(lines)
        try:
            key = _read_key()
        except KeyboardInterrupt:
            return None
        if key == 'up' and cursor > 0:
            cursor -= 1
        elif key == 'down' and cursor < len(options) - 1:
            cursor += 1
        elif key == 'space':
            if cursor in selected:
                selected.discard(cursor)
            else:
                selected.add(cursor)
        elif key == 'a':
            selected = set() if len(selected) == len(options) else set(range(len(options)))
        elif key == 'enter':
            return selected if selected else None
        elif key in ('escape', 'q'):
            return None


# ----- the menu itself --------------------------------------------------------

def _choose_primary_action(action_filter: Optional[str] = None) -> Optional[str]:
    if action_filter is not None:
        return action_filter
    idx = _arrow_single_select(
        header="\nWhat do you want to do?",
        options=["Patch installed targets",
                 "Restore from .bak",
                 "Quit"],
        footer="(Up/Down to move, Enter to confirm, q/Esc to cancel)",
    )
    if idx is None or idx == 2:
        return None
    return "patch" if idx == 0 else "restore"


def _choose_targets_for_action(action: str,
                               rows: "list[tuple[str, str, str, str]]",
                               option_formatter: Callable[[tuple[str, str, str, str]], str],
                               allow_bak_restore: bool = False
                               ) -> "list[tuple[str, str]]":
    if action == "patch":
        relevant = [r for r in rows if r[3] in PATCHABLE_STATES]
        preselect = {i for i, r in enumerate(relevant) if r[3] == "UNPATCHED"}
    else:
        relevant = [r for r in rows if r[3] in RESTORABLE_STATES
                    or (allow_bak_restore and Path(r[2] + ".bak").exists())]
        preselect = set(range(len(relevant)))

    if not relevant:
        print("\nNothing to do for this action.\n")
        return []

    verb = action.upper()
    chosen_idx = _arrow_multi_select(
        header=f"\nSelect targets to {verb}:",
        options=[option_formatter(r) for r in relevant],
        preselected=preselect,
        footer="(Up/Down move, Space toggle, A all/none, Enter confirm, q/Esc cancel)",
    )
    if not chosen_idx:
        print("\nCancelled (no targets selected).\n")
        return []

    print()
    print(f"About to {verb}:")
    for i in sorted(chosen_idx):
        _kind, label, path, state = relevant[i]
        note = ""
        if action == "patch" and state == "PATCHED":
            note = "  (already patched - will re-apply cleanly from backup)"
        elif action == "patch" and state in ("PARTIAL", "UNKNOWN"):
            note = f"  ({state.lower()} state - will re-apply cleanly from backup)"
        elif action == "restore" and allow_bak_restore and not Path(path + ".bak").exists():
            note = "  (no .bak found - restore will fail)"
        elif action == "restore" and state != "PATCHED":
            note = f"  (current state: {state})"
        print(f"  - {label}{note}")
    print()

    confirm_idx = _arrow_single_select(
        header="Continue?",
        options=["Yes, proceed", "No, cancel"],
        footer="(Up/Down + Enter, or q/Esc to cancel)",
    )
    if confirm_idx != 0:
        print("\nCancelled.\n")
        return []
    return [(relevant[i][0], relevant[i][2]) for i in sorted(chosen_idx)]


def interactive_menu(resolve_path: Optional[str],
                     action_filter: Optional[str] = None
                     ) -> "tuple[Optional[str], list[tuple[str, str]]]":
    """Arrow-key menu. If `action_filter` is None, the user picks patch vs
    restore first. If 'patch' or 'restore' is passed, that step is skipped
    (used by the smart `--restore` path).

    For 'patch': lists all installed targets; UNPATCHED rows are pre-checked.
    For 'restore': lists PATCHED/PARTIAL targets; all are pre-checked.

    Returns (action, [(kind, path), ...]) or (None, []) on cancel."""
    rows = _build_target_rows(resolve_path)

    print()
    print("=" * 72)
    print("        Blackmagic Design Patcher  -  DaVinci Resolve & Fusion")
    print("=" * 72)

    action = _choose_primary_action(action_filter)
    if action is None:
        return None, []
    chosen = _choose_targets_for_action(
        action,
        rows,
        lambda r: f"{r[1]:42s}  {r[3]}",
    )
    return action, chosen


def _build_linux_target_rows(resolve_path: Optional[str]) -> "list[tuple[str, str, str, str]]":
    rows: list[tuple[str, str, str, str]] = []
    if resolve_path is not None:
        rows.append(("linux_resolve",
                     "DaVinci Resolve Linux",
                     resolve_path,
                     state_of_linux_resolve(resolve_path)))
    for lib_path in locate_linux_fusion():
        rows.append(("linux_fusion21",
                     "Fusion Studio 21 Linux",
                     lib_path,
                     state_of_linux_fusion(lib_path)))
    return rows


def interactive_linux_menu(resolve_path: Optional[str]) -> "tuple[Optional[str], list[tuple[str, str]]]":
    """Arrow-key menu for Linux Resolve/Fusion targets."""
    print()
    print("=" * 72)
    print("        Blackmagic Design Patcher  -  Resolve & Fusion Linux")
    print("=" * 72)
    print()
    print("Checking installed Linux targets...")
    sys.stdout.flush()
    rows = _build_linux_target_rows(resolve_path)
    sys.stdout.write("\x1b[1A\x1b[2K")
    if not rows:
        print("No Linux Resolve/Fusion targets found.")
        return None, []
    print("Detected targets:")
    for _kind, label, path, state in rows:
        print(f"  - {label:28s} {state:10s} {path}")
    sys.stdout.flush()

    action = _choose_primary_action()
    if action is None:
        return None, []

    chosen = _choose_targets_for_action(
        action,
        rows,
        lambda r: f"{r[1]:28s} {r[3]:10s} {r[2]}",
        allow_bak_restore=True,
    )
    return action, chosen


def _execute(action: str, chosen: "list[tuple[str, str]]") -> int:
    """Run `action` ('patch'|'restore') against the chosen [(kind, path), ...].
    Each kill is deferred until just before its first relevant work — so we
    never kill Fusion if only Resolve was selected, and vice versa."""
    groups = (
        # "resolve" target is now a bundle: Resolve.exe + fraunhoferdcp.dll +
        # DaVinci Remote Monitor.exe (any companion that exists in the
        # install dir). _kill_resolve_bundle also kills Remote Monitor so
        # the file isn't locked when we try to write it.
        ("Resolve", lambda k: k == "resolve", _kill_resolve_bundle,
         patch_resolve_bundle, restore_resolve_bundle),
        ("Fusion", lambda k: k.startswith("fusion"), _kill_fusion, patch_fusion, restore_fusion),
        ("Linux Resolve", lambda k: k == "linux_resolve", None, patch_linux, restore_linux),
        ("Linux Fusion", lambda k: k.startswith("linux_fusion"), None,
         patch_linux_fusion, restore_linux_fusion),
    )
    rc = 0
    for label, matches, before, patch_fn, restore_fn in groups:
        picked = [(k, p) for k, p in chosen if matches(k)]
        if not picked:
            continue
        if before is not None:
            before()
        fn = restore_fn if action == "restore" else patch_fn
        done = "restored" if action == "restore" else "patched"
        for _kind, path in picked:
            try:
                logger.info("attempting to %s %s: %s", action, label, path)
                fn(path)
                logger.info("successfully %s %s", done, label)
            except PatchError as e:
                logger.error("%s failed (%s): %s", label, path, e)
                rc = 1
    return rc


def _resolve_chosen_from_cli(targets_spec: str, resolve_path: Optional[str]) -> "list[tuple[str, str]]":
    """Translate a `--targets` string into the `chosen` list shape used by
    _execute. Skips MISSING installs silently — same behaviour as the menu.
    The "resolve" kind is a bundle (handled inside _execute); we only emit
    the main Resolve.exe path here."""
    targets = _parse_targets(targets_spec)
    chosen: list[tuple[str, str]] = []
    if "resolve" in targets and resolve_path is not None:
        chosen.append(("resolve", resolve_path))
    if "fusion" in targets:
        for fp in locate_fusion():
            chosen.append(("fusion21", fp))
    return chosen


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    args = _build_arg_parser().parse_args()  # exits cleanly on --help

    _enable_ansi()

    if sys.platform.startswith("linux"):
        _require_root()
        path: Optional[str] = None
        try:
            path = args.path or locate_linux_resolve()
        except PatchError:
            path = None

        # Same interaction model as Windows: no flags + real terminal => menu.
        if args.targets is None and not args.restore and sys.stdin.isatty():
            action, chosen = interactive_linux_menu(path)
            if action is None or not chosen:
                return 0
            return _execute(action, chosen)

        action = "restore" if args.restore else "patch"
        targets = _parse_targets(args.targets or "all")
        chosen: list[tuple[str, str]] = []
        if "resolve" in targets and path is not None:
            logger.info("target Linux Resolve binary: %s", path)
            chosen.append(("linux_resolve", path))
        if "fusion" in targets:
            for lib_path in locate_linux_fusion():
                logger.info("target Linux Fusion library: %s", lib_path)
                chosen.append(("linux_fusion21", lib_path))
        if not chosen:
            logger.info("nothing to do (no Linux targets resolved)")
            return 0
        return _execute(action, chosen)

    if sys.platform != "win32":
        logger.error("unsupported operating system: %s", sys.platform)
        return 1

    _require_admin()

    # Auto-locate Resolve. Missing-install isn't fatal at this point — the
    # menu / CLI dispatch handles it.
    resolve_path: Optional[str] = None
    try:
        resolve_path = args.path or locate()
    except PatchError:
        resolve_path = None

    # ---- explicit CLI mode (scripted / unattended) --------------------------
    # `--targets` given, OR stdin isn't a TTY: bypass menus entirely.
    if args.targets is not None or not sys.stdin.isatty():
        targets_spec = args.targets or "all"
        chosen = _resolve_chosen_from_cli(targets_spec, resolve_path)
        if not chosen:
            logger.info("nothing to do (no targets resolved)")
            return 0
        action = "restore" if args.restore else "patch"
        return _execute(action, chosen)

    # ---- interactive flow ---------------------------------------------------

    if args.restore:
        # Smart --restore: skip the menu when 0 or 1 targets are patched.
        rows = _build_target_rows(resolve_path)
        patched = [r for r in rows if r[3] in ("PATCHED", "PARTIAL")]

        if not patched:
            logger.info("nothing to restore (no targets are currently patched)")
            return 0

        if len(patched) == 1:
            kind, label, path, _state = patched[0]
            print()
            print(f"Single patched install detected: {label}")
            print(f"  {path}")
            print()
            confirm_idx = _arrow_single_select(
                header=f"Restore {label} from .bak?",
                options=["Yes, restore", "No, cancel"],
                footer="(Up/Down + Enter, or q/Esc to cancel)",
            )
            if confirm_idx != 0:
                print("\nCancelled.\n")
                return 0
            return _execute("restore", [(kind, path)])

        # 2+ patched: full menu, but skip the patch/restore choice (filtered
        # to restore action only — and only patched rows are listed).
        action, chosen = interactive_menu(resolve_path, action_filter="restore")
    else:
        action, chosen = interactive_menu(resolve_path, action_filter=None)

    if action is None or not chosen:
        return 0
    return _execute(action, chosen)


if __name__ == "__main__":
    sys.exit(main())
