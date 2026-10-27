#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tdoc2md.py - Convert 3GPP TDocs (pCRs, CRs, LSs, discussion papers) from Word to Markdown.

Standalone conversion core using only the Python standard library (Python >= 3.8).
Office integration is optional and can be supplied through callbacks.
The .docx (OpenXML) package is parsed directly with ``zipfile`` + ``xml.etree``,
so Microsoft Word is NOT required for .docx files.

Features
--------
* Track changes rendered inline as ``[INSERTED: ...]`` / ``[DELETED: ...]``
  (handles fragmented runs and <w:del> nested inside <w:ins>), or accepted /
  rejected views via ``--revisions accept|reject``.
* Tables -> GitHub-flavoured Markdown tables (merged cells, nested tables,
  1x1 "boxed text" tables, 3GPP CR cover sheet detection).
* Images: <w:drawing>, <w:pict> and <w:object> are exported to ``media/`` in
  document order. EMF/WMF previews are converted to PNG when a converter is
  available (Windows PowerShell/System.Drawing, Pillow, or LibreOffice).
* Embedded Visio (and other OLE) objects: the rendered preview is shown inline
  and the editable source (.vsdx/.vsd/.bin) is extracted and linked.
* Word field codes (DOCPROPERTY, HYPERLINK, ...) are dropped; only visible text is kept.
* Headings, automatic list numbering, 3GPP change markers ("Start of changes",
  "Next change", ...), "all text is new" blocks and placeholder clauses.
* Metadata from a 3GPP TDoc list (.xlsx, read with the stdlib) or from the
  TDoc cover page (Source:, Title:, Agenda item:, ...).
* Optional bulk corpus: per Agenda Item / category mega-files with chunking.

Optional helpers (used automatically if present, never required)
----------------------------------------------------------------
* Windows PowerShell (System.Drawing) or Pillow or LibreOffice -> EMF/WMF to PNG.
* MS Word via pywin32, or LibreOffice -> legacy binary .doc to .docx.

Usage
-----
    python tdoc2md.py <file.docx | tdoc folder | meeting folder> [-o OUT_DIR]
                      [--excel TDocs.xlsx] [--bulk] [--max-chars 200000]
                      [--revisions markup|accept|reject] [--image-dpi 150]
"""
from __future__ import annotations

import argparse
import gzip
import importlib
import logging
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import zlib
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, Iterator, List, Optional, Tuple, Union

__version__ = "2.1.0"
log = logging.getLogger("tdoc2md")

# Path handling is part of the standalone converter; no companion imports are required.
_GetShortPathNameW = None
if os.name == "nt":
    try:
        import ctypes
        from ctypes import wintypes
        _GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
        _GetShortPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
        _GetShortPathNameW.restype = wintypes.DWORD
    except Exception:
        _GetShortPathNameW = None


def to_extended_path(p: Path | str) -> str:
    """Convert a path to Windows extended-length format (\\\\?\\ or \\\\?\\UNC\\\\)

    Bypasses the Win32 MAX_PATH (260 character) limit for Python standard library file I/O.
    On non-Windows platforms, returns the string path unchanged.
    """
    p_str = os.path.abspath(str(p))
    if os.name == "nt" and not p_str.startswith("\\\\?\\"):
        if p_str.startswith("\\\\"):
            return "\\\\?\\UNC\\" + p_str[2:]
        return "\\\\?\\" + p_str
    return p_str


def strip_extended_path(p: Path | str) -> str:
    """Remove Windows \\\\?\\ or \\\\?\\UNC\\\\ prefix for clean relative paths, display, and external tools."""
    s = str(p)
    if s.startswith("\\\\?\\UNC\\"):
        return "\\\\" + s[8:]
    if s.startswith("\\\\?\\"):
        return s[4:]
    return s


def to_tool_path(p: Path | str) -> str:
    """Return a path safe for external tools (PowerShell, LibreOffice, Word COM).

    Strips Windows \\\\?\\ prefixes, and on Windows converts paths >= 250 characters
    to 8.3 short paths using GetShortPathNameW to avoid MAX_PATH failures.
    """
    p_str = os.path.abspath(str(p))
    if os.name != "nt":
        return p_str
    p_clean = strip_extended_path(p_str)
    if len(p_clean) < 250 or _GetShortPathNameW is None:
        return p_clean

    def _shorten(path_str: str) -> str:
        buf = ctypes.create_unicode_buffer(1024)
        res = _GetShortPathNameW(to_extended_path(path_str), buf, len(buf))
        if res > 0:
            return strip_extended_path(buf.value)
        return path_str

    short = _shorten(p_clean)
    if short != p_clean:
        return short

    parts: List[str] = []
    curr = p_clean
    while curr and not os.path.exists(to_extended_path(curr)):
        parent, tail = os.path.split(curr)
        if parent == curr or not tail:
            break
        parts.append(tail)
        curr = parent

    short_parent = _shorten(curr)
    for part in reversed(parts):
        short_parent = os.path.join(short_parent, part)
    return short_parent


def _office_backend():
    """Load the skill's optional office adapter only when it is installed beside us.

    Embedding applications supply callbacks and only need to copy this file.
    """
    if not Path(__file__).with_name("tdoc2md_office.py").is_file():
        return None
    name = (__package__ + "." if __package__ else "") + "tdoc2md_office"
    return importlib.import_module(name)


def __getattr__(name):
    """Preserve historical office helper imports in the skill installation."""
    if name in {"load_word_config", "MsoAssignmentMethod", "apply_word_sensitivity_label",
                "find_libreoffice_exe", "convert_doc_to_docx", "_convert_with_soffice"}:
        backend = _office_backend()
        if backend is not None:
            return getattr(backend, name)
    raise AttributeError("module %r has no attribute %r" % (__name__, name))


DOC_EXTENSIONS = {".docx", ".docm", ".doc"}
METAFILE_EXTENSIONS = {".emf", ".wmf"}
COMPRESSED_METAFILES = {".emz": ".emf", ".wmz": ".wmf"}
RASTER_TO_CONVERT = {".tif", ".tiff", ".bmp"}

TDOC_RE = re.compile(r"(?<![A-Za-z0-9])([A-Z][A-Z0-9]{0,3}-\d{6,8}(?:r\d{1,2})?)(?![0-9])", re.I)
REVISION_SUFFIX_RE = re.compile(r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", re.I)

# 3GPP structural heuristics (kept from the legacy llm_exporter, extended)
BOUNDARY_RE = re.compile(
    r"(?i)^[\W_]*(?:"
    r"(?:start|end|beginning|begin)\s+of\s+(?:the\s+)?(?:\w+\s+){0,3}?(?:changes?|modifications?|proposals?)"
    r"|(?:next|first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|last|\d+(?:st|nd|rd|th))"
    r"\s+(?:\w+\s+)?(?:changes?|modifications?)"
    r")(?:\s+(?:\d+|[ivx]+))?[\W_]*$"
)
LEGACY_BOUNDARY_RE = re.compile(r"(?i)<[-\s]*(next|end of)\s*change")
ALL_NEW_RE = re.compile(
    r"(?i)\ball\s+(?:new\s+text|(?:the\s+)?(?:\w+\s+){0,2}?text\s+"
    r"(?:\w+\s+){0,3}?(?:is\s+)?(?:new|added))\b"
)
PLACEHOLDER_CLAUSE_RE = re.compile(r"^(?:\d+\.)+[A-Za-z]{1,2}$")
NUMERIC_CLAUSE_RE = re.compile(r"^(?:[A-Z]\.)?\d+(?:\.\d+)*\.?$")


def change_boundary_qualifier(text: str) -> Optional[str]:
    """Return a boundary's qualifier, or None when this is ordinary prose.

    Keep the complete marker for rendering: qualifiers can identify a separate
    contribution baseline, not just whether text is new to the target spec.
    """
    candidate = re.sub(r"^[\s*]+|[\s*]+$", "", text)
    match = re.fullmatch(r"(.*?)\s*\((.*)\)", candidate)
    boundary, qualifier = (match.group(1), match.group(2)) if match else (candidate, "")
    return qualifier if BOUNDARY_RE.fullmatch(boundary) else None

COVER_FIELDS = {
    "source": "source", "source to wg": "source", "title": "title", "agenda item": "agenda item",
    "agenda": "agenda item", "document for": "document for", "work item": "work item",
    "work item / release": "work item", "work item code": "work item", "spec": "spec",
    "version": "version", "abstract": "abstract", "abstract of the contribution": "abstract",
    "to": "to", "cc": "cc", "response to": "response to", "release": "release",
    "contact person": "contact", "category": "category", "date": "date",
}
COVER_LINE_RE = re.compile(r"^\s*([A-Za-z][A-Za-z /]{0,40}?)\s*:\s*(.+)$", re.S)
REVISION_OF_RE = re.compile(r"(?i)revision\s+of\s+([A-Z][A-Z0-9]{0,3}-\d{6,8})")

# Common Symbol-font private-use code points (w:sym)
SYMBOL_MAP = {
    0xAE: "\u2192", 0xAC: "\u2190", 0xAD: "\u2191", 0xAF: "\u2193", 0xAB: "\u2194", 0xDE: "\u21d2",
    0xDC: "\u21d0", 0xDB: "\u21d4", 0xB3: "\u2265", 0xA3: "\u2264", 0xB9: "\u2260", 0xB4: "\u00d7",
    0xB1: "\u00b1", 0xB7: "\u2022", 0xD7: "\u00b7", 0x2D: "\u2212", 0xA5: "\u221e", 0xBB: "\u2248",
    0x6C: "\u03bb", 0x6D: "\u03bc", 0x61: "\u03b1", 0x62: "\u03b2", 0x44: "\u0394", 0xE6: "(", 0xF6: ")",
}

DEFAULT_CORPUS_PROMPT = (
    "This file contains a programmatic compilation of 3GPP Technical Documents (TDocs). "
    "These documents represent telecommunications standards proposals, revisions, and working group agreements.\n\n"
    "**Structural Rules for parsing this text:**\n"
    "- `[ADDED BLOCK]:` Denotes text new to the target specification; this does not identify edits relative to an earlier contribution.\n"
    "- `[INSERTED: <text>]`: Denotes specific inline text additions explicitly marked via Word Track Changes.\n"
    "- `[DELETED: <text>]`: Denotes specific inline text removals explicitly marked via Word Track Changes.\n"
    "- `> **[...change...]**`: Change boundary markers of a (p)CR (e.g. 'Start of changes', 'Next change').\n"
    "- `![...](media/...)`: Figures (call flows, architecture diagrams). Embedded Visio sources are linked next to them.\n\n"
    "**Your Task:** Please use this corpus to analyze technical agreements, architectural changes, "
    "or contradictions within this specific Agenda Item."
)

MARKUP_LEGEND = (
    "> **Markup legend:** `[INSERTED: ...]` / `[DELETED: ...]` = Word track changes; "
    "`[ADDED BLOCK]:` = text new to the target specification, separate from tracked contribution edits; "
    "`> **[...]**` = change boundary markers; figures are exported to `media/`."
)


# --------------------------------------------------------------------------------------
# Generic helpers
# --------------------------------------------------------------------------------------
def ln(tag) -> str:
    """Local name of an XML tag (namespace stripped)."""
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def get_attr(el, name: str, default=None):
    """Attribute lookup by local name (namespace-agnostic)."""
    if el is None:
        return default
    for k, v in el.attrib.items():
        if k == name or k.endswith("}" + name):
            return v
    return default


def rel_attr(el, *names: str) -> Optional[str]:
    """Relationship-id attribute (r:id, r:embed, r:link, o:relid)."""
    if el is None:
        return None
    for k, v in el.attrib.items():
        local = ln(k)
        if local in names and ("relationships" in k or local == "relid"):
            return v
    return None


def iter_skip(el, skip: Tuple[str, ...] = ("txbxContent",)) -> Iterator:
    """Depth-first iteration over descendants without entering elements whose local name is in `skip`."""
    for child in el:
        if ln(child.tag) in skip:
            continue
        yield child
        yield from iter_skip(child, skip)


def sanitize_filename(name, max_len: int = 120) -> str:
    s = re.sub(r'[\\/*?:"<>|\x00-\x1f]', "_", str(name or "").strip())
    s = re.sub(r"_+", "_", s).strip(" ._")
    return (s or "Unknown")[:max_len]


def safe_media_name(name: str) -> str:
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", name)
    return re.sub(r"_+", "_", s).strip("._") or "file"


def is_within(path: Path, parent: Optional[Path]) -> bool:
    if parent is None:
        return False
    try:
        p1 = Path(strip_extended_path(path.resolve()))
        p2 = Path(strip_extended_path(parent.resolve()))
        p1.relative_to(p2)
        return True
    except (ValueError, OSError):
        return False


def guess_tdoc_id(path: Path) -> str:
    folder = path.parent.name.strip()
    m = TDOC_RE.fullmatch(folder)
    if m:
        return m.group(1).upper()
    m = TDOC_RE.search(path.stem)
    if m:
        return m.group(1).upper()
    m = TDOC_RE.search(folder)
    if m:
        return m.group(1).upper()
    return sanitize_filename(path.stem)


def base_tdoc_id(tdoc_id: str) -> str:
    m = REVISION_SUFFIX_RE.match(tdoc_id)
    return (m.group(1) if m else tdoc_id).upper()


def _roman(n: int) -> str:
    vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
            (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out or "0"


def _letters(n: int) -> str:
    if n <= 0:
        return "0"
    s = ""
    while n > 0:
        n -= 1
        s = chr(97 + n % 26) + s
        n //= 26
    return s


def format_number(n: int, fmt: str) -> str:
    if fmt == "lowerLetter":
        return _letters(n)
    if fmt == "upperLetter":
        return _letters(n).upper()
    if fmt == "lowerRoman":
        return _roman(n)
    if fmt == "upperRoman":
        return _roman(n).upper()
    if fmt == "decimalZero":
        return "%02d" % n
    return str(n)


# --------------------------------------------------------------------------------------
# Styles & numbering
# --------------------------------------------------------------------------------------
class Styles:
    def __init__(self, root):
        self.styles: Dict[str, dict] = {}
        self.default_para: Optional[str] = None
        self._cache: Dict[Optional[str], Tuple[str, Optional[int], Optional[Tuple[str, int]]]] = {}
        if root is None:
            return
        for s in root:
            if ln(s.tag) != "style":
                continue
            if get_attr(s, "type") not in (None, "paragraph"):
                continue
            sid = get_attr(s, "styleId")
            if not sid:
                continue
            entry = {"name": sid, "based": None, "outline": None, "numId": None, "ilvl": None}
            for c in s:
                l = ln(c.tag)
                if l == "name":
                    entry["name"] = get_attr(c, "val", sid)
                elif l == "basedOn":
                    entry["based"] = get_attr(c, "val")
                elif l == "pPr":
                    for pc in c:
                        pl = ln(pc.tag)
                        if pl == "outlineLvl":
                            try:
                                entry["outline"] = int(get_attr(pc, "val", "9"))
                            except ValueError:
                                pass
                        elif pl == "numPr":
                            for nc in pc:
                                if ln(nc.tag) == "numId":
                                    entry["numId"] = get_attr(nc, "val")
                                elif ln(nc.tag) == "ilvl":
                                    try:
                                        entry["ilvl"] = int(get_attr(nc, "val", "0"))
                                    except ValueError:
                                        pass
            self.styles[sid] = entry
            if get_attr(s, "default") in ("1", "true", "on"):
                self.default_para = sid

    def info(self, style_id: Optional[str]):
        """Return (style name, heading level or None, (numId, ilvl) or None)."""
        sid = style_id or self.default_para
        if sid in self._cache:
            return self._cache[sid]
        name, heading, num = "", None, None
        cur, depth, first = sid, 0, True
        while cur and cur in self.styles and depth < 15:
            st = self.styles[cur]
            if first:
                name, first = st["name"], False
            if heading is None:
                m = re.match(r"(?i)^heading\s*([1-9])$", st["name"] or "")
                if m:
                    heading = int(m.group(1))
                elif st["outline"] is not None and st["outline"] < 9:
                    heading = st["outline"] + 1
            if num is None and st["numId"] is not None:
                num = (st["numId"], st["ilvl"] or 0)
            cur = st["based"]
            depth += 1
        self._cache[sid] = (name, heading, num)
        return self._cache[sid]


class Numbering:
    def __init__(self, root):
        self.abstract: Dict[str, Dict[int, Tuple[str, str, int]]] = {}
        self.nums: Dict[str, Tuple[str, Dict[int, int]]] = {}
        self.counters: Dict[str, List[Optional[int]]] = {}
        if root is None:
            return
        for el in root:
            l = ln(el.tag)
            if l == "abstractNum":
                aid = get_attr(el, "abstractNumId")
                levels = {}
                for lvl in el:
                    if ln(lvl.tag) != "lvl":
                        continue
                    try:
                        ilvl = int(get_attr(lvl, "ilvl", "0"))
                    except ValueError:
                        continue
                    fmt, text, start = "decimal", "", 1
                    for c in lvl:
                        cl = ln(c.tag)
                        if cl == "numFmt":
                            fmt = get_attr(c, "val", "decimal")
                        elif cl == "lvlText":
                            text = get_attr(c, "val", "")
                        elif cl == "start":
                            try:
                                start = int(get_attr(c, "val", "1"))
                            except ValueError:
                                pass
                    levels[ilvl] = (fmt, text, start)
                self.abstract[aid] = levels
            elif l == "num":
                nid = get_attr(el, "numId")
                aid, overrides = None, {}
                for c in el:
                    cl = ln(c.tag)
                    if cl == "abstractNumId":
                        aid = get_attr(c, "val")
                    elif cl == "lvlOverride":
                        try:
                            ilvl = int(get_attr(c, "ilvl", "0"))
                        except ValueError:
                            continue
                        for oc in c:
                            if ln(oc.tag) == "startOverride":
                                try:
                                    overrides[ilvl] = int(get_attr(oc, "val", "1"))
                                except ValueError:
                                    pass
                self.nums[nid] = (aid, overrides)

    def label(self, num_id: Optional[str], ilvl: int) -> Optional[str]:
        if not num_id or num_id == "0" or num_id not in self.nums:
            return None
        aid, overrides = self.nums[num_id]
        levels = self.abstract.get(aid, {})
        if ilvl not in levels:
            return None
        cnt = self.counters.setdefault(num_id, [None] * 10)
        fmt, text, start = levels[ilvl]
        start = overrides.get(ilvl, start)
        cnt[ilvl] = start if cnt[ilvl] is None else cnt[ilvl] + 1
        for d in range(ilvl + 1, 10):
            cnt[d] = None
        if fmt == "bullet":
            return "-"

        def repl(m):
            k = int(m.group(1)) - 1
            lvl = levels.get(k, ("decimal", "", 1))
            val = cnt[k] if 0 <= k < 10 and cnt[k] is not None else overrides.get(k, lvl[2])
            return "" if lvl[0] == "none" else format_number(val, lvl[0])

        out = re.sub(r"%(\d)", repl, text).strip()
        return out or None


# --------------------------------------------------------------------------------------
# Conversion result containers
# --------------------------------------------------------------------------------------
@dataclass
class ParaRender:
    text: str
    plain: str
    has_rev: bool
    all_ins: bool
    has_content: bool


@dataclass
class DocResult:
    markdown: str
    cover: Dict[str, str] = field(default_factory=dict)
    is_cr: bool = False
    has_revisions: bool = False
    detected_type: str = "Other"
    media_files: List[Path] = field(default_factory=list)
    image_count: int = 0
    warnings: List[str] = field(default_factory=list)


# --------------------------------------------------------------------------------------
# DOCX -> Markdown converter
# --------------------------------------------------------------------------------------
class DocxConverter:
    def __init__(self, docx_path: Path, tdoc_id: str, media_dir: Path, media_rel: str = "media",
                 prefix: Optional[str] = None, revision_mode: str = "markup", image_dpi: int = 150,
                 convert_images: bool = True, keep_metafiles: bool = False,
                 cancellation_check: Optional[Callable[[], bool]] = None,
                 metafile_converter: Optional[Callable[[List[Tuple[Path, Path]], int], None]] = None):
        self.path = Path(to_extended_path(docx_path))
        self.tdoc_id = tdoc_id
        self.media_dir = Path(to_extended_path(media_dir))
        self.media_rel = media_rel.rstrip("/")
        self.prefix = safe_media_name(prefix or tdoc_id)
        self.revision_mode = revision_mode
        self.image_dpi = image_dpi
        self.convert_images = convert_images
        self.keep_metafiles = keep_metafiles
        self.cancellation_check = cancellation_check
        self.metafile_converter = metafile_converter
        # state
        self.W = ""
        self.z: Optional[zipfile.ZipFile] = None
        self.names_lower: Dict[str, str] = {}
        self.rels: Dict[str, Tuple[str, str, bool]] = {}
        self.styles = Styles(None)
        self.numbering = Numbering(None)
        self.saved: Dict[str, Tuple[str, Path]] = {}
        self.metafile_jobs: List[Tuple[Path, Path, str]] = []
        self.metafile_tokens: Dict[Path, str] = {}
        self.in_new_block = False
        self.explicit_new_block = False
        self.cover_done = False
        self.cover_seen_table = False
        self.result = DocResult(markdown="")
        self.figure_no = 0
        self.top_texts: List[str] = []
        self.marker_count = 0

    # ---------------- package access ----------------
    def _read(self, name: str) -> Optional[bytes]:
        real = self.names_lower.get(name.lower())
        if real is None:
            return None
        try:
            return self.z.read(real)
        except KeyError:
            return None

    def _xml(self, name: str):
        data = self._read(name)
        if data is None:
            return None
        try:
            return ET.fromstring(data)
        except ET.ParseError as e:
            self.result.warnings.append("XML parse error in %s: %s" % (name, e))
            return None

    def _part_rels(self, part: str) -> Dict[str, Tuple[str, str, bool]]:
        d, n = posixpath.split(part)
        root = self._xml(posixpath.join(d, "_rels", n + ".rels"))
        rels = {}
        if root is None:
            return rels
        for r in root:
            rid, tgt = r.get("Id"), r.get("Target", "")
            typ = r.get("Type", "").rsplit("/", 1)[-1]
            if (r.get("TargetMode") or "").lower() == "external":
                rels[rid] = (tgt, typ, True)
                continue
            full = tgt.lstrip("/") if tgt.startswith("/") else posixpath.join(d, tgt)
            rels[rid] = (posixpath.normpath(full), typ, False)
        return rels

    def _main_part(self) -> str:
        root = self._xml("_rels/.rels")
        if root is not None:
            for r in root:
                if r.get("Type", "").endswith("/officeDocument"):
                    return r.get("Target", "word/document.xml").lstrip("/")
        return "word/document.xml"

    # ---------------- public ----------------
    def convert(self) -> DocResult:
        with zipfile.ZipFile(to_extended_path(self.path)) as z:
            self.z = z
            self.names_lower = {n.lower(): n for n in z.namelist()}
            main = self._main_part()
            root = self._xml(main)
            if root is None:
                raise ValueError("No main document part found in %s" % self.path.name)
            self.W = root.tag[: root.tag.index("}") + 1] if root.tag.startswith("{") else ""
            self.rels = self._part_rels(main)
            for _rid, (target, typ, ext) in self.rels.items():
                if ext:
                    continue
                if typ == "styles":
                    self.styles = Styles(self._xml(target))
                elif typ == "numbering":
                    self.numbering = Numbering(self._xml(target))
            body = root.find(self.W + "body")
            blocks = self._render_blocks(body, top=True, in_cell=False) if body is not None else []
        md = "\n\n".join(b for b in blocks if b and b.strip())
        md = self._finalize_media(md)
        md = re.sub(r"\n{3,}", "\n\n", md).strip() + "\n"
        self.result.markdown = md
        self._parse_cover()
        self.result.detected_type = self._detect_type()
        return self.result

    # ---------------- blocks ----------------
    def _render_blocks(self, container, top: bool, in_cell: bool) -> List[str]:
        blocks: List[str] = []
        W = self.W
        for el in container:
            if self.cancellation_check and self.cancellation_check():
                break
            tag, l = el.tag, ln(el.tag)
            if tag == W + "p":
                b = self._render_para_block(el, top, in_cell)
                if b:
                    blocks.append(b)
            elif tag == W + "tbl":
                b = self._render_table(el, top, in_cell)
                if b:
                    blocks.append(b)
            elif l == "sdt":
                sc = el.find(W + "sdtContent")
                if sc is not None:
                    blocks.extend(self._render_blocks(sc, top, in_cell))
            elif l in ("customXml", "sdtContent", "ins", "moveTo", "smartTag"):
                blocks.extend(self._render_blocks(el, top, in_cell))
            elif l == "AlternateContent":
                choice = next((c for c in el if ln(c.tag) == "Choice"), None)
                if choice is not None:
                    blocks.extend(self._render_blocks(choice, top, in_cell))
        return blocks

    def _para_props(self, p):
        W = self.W
        pPr = p.find(W + "pPr")
        style_id, num, outline = None, None, None
        if pPr is not None:
            ps = pPr.find(W + "pStyle")
            if ps is not None:
                style_id = get_attr(ps, "val")
            numPr = pPr.find(W + "numPr")
            if numPr is not None:
                nid = get_attr(numPr.find(W + "numId"), "val")
                il = get_attr(numPr.find(W + "ilvl"), "val", "0")
                num = (nid, int(il) if il and il.isdigit() else 0)
            ol = pPr.find(W + "outlineLvl")
            if ol is not None:
                try:
                    outline = int(get_attr(ol, "val", "9"))
                except ValueError:
                    pass
        name, heading, style_num = self.styles.info(style_id)
        if heading is None and outline is not None and outline < 9:
            heading = outline + 1
        if num is None:
            num = style_num
        return name or (style_id or ""), heading, num

    def _render_para_block(self, p, top: bool, in_cell: bool) -> Optional[str]:
        style_name, heading, num = self._para_props(p)
        segs: List[Tuple[str, str, str]] = []
        for c in p:
            if ln(c.tag) == "pPr":
                continue
            self._walk(c, "NORMAL", segs)
        para = self._render_segments(segs)
        label = self.numbering.label(num[0], num[1]) if num else None
        if not para.has_content:
            return None
        if para.has_rev:
            self.result.has_revisions = True
        # Retain internal whitespace in every view so projected markup and a
        # direct accept/reject export agree, including adjacent revision spans.
        text = para.text.strip()
        plain = re.sub(r"\s+", " ", para.plain).strip()
        if in_cell:
            heading = None
        if top and not self.cover_done:
            if heading:
                self.cover_done = True
            elif len(self.top_texts) < 40:
                self.top_texts.append(plain)

        note = ""
        if top:
            short = len(plain) <= 120
            qualifier = change_boundary_qualifier(plain)
            if qualifier is not None or (short and LEGACY_BOUNDARY_RE.search(plain)):
                self.in_new_block = bool(qualifier and ALL_NEW_RE.search(qualifier))
                self.explicit_new_block = self.in_new_block
                self.marker_count += 1
                marker = "> **[%s]**" % plain
                if self.in_new_block:
                    marker += "\n\n> **[Change scope: new text for the target specification; inline revisions retain their separate edit layer]**"
                return marker
            words = plain.split()
            first = words[0] if words else ""
            if heading and self.in_new_block and not self.explicit_new_block and NUMERIC_CLAUSE_RE.match(first):
                self.in_new_block = False
            if short and ALL_NEW_RE.search(plain):
                self.in_new_block = True
                self.explicit_new_block = True
                return "> **[Change scope: new text for the target specification; inline revisions retain their separate edit layer]** %s" % text
            if heading and PLACEHOLDER_CLAUSE_RE.match(first):
                self.in_new_block = True
                note = "> **[Note to LLM: entering new placeholder clause '%s']**\n\n" % first

        if heading:
            level = max(1, min(heading, 6))
            htxt = text.replace("\n", " ")
            if label and label != "-":
                htxt = "%s %s" % (label, htxt)
            return note + "%s %s" % ("#" * level, htxt)

        if (top and self.in_new_block and not para.has_rev and self.revision_mode == "markup"):
            text = "[ADDED BLOCK]: " + text

        if label:
            indent = "  " * (num[1] if num else 0)
            text = "%s%s %s" % (indent, label, text)
        elif style_name in ("TH", "TF") and "![" not in text and "\n" not in text:
            text = "**%s**" % text
        return note + text

    # ---------------- inline walking ----------------
    def _walk(self, el, state: str, out: List[Tuple[str, str, str]]):
        W = self.W
        tag = el.tag
        l = ln(tag)
        if tag.startswith(W) or not W:
            if l == "t":
                out.append(("text", state, el.text or ""))
                return
            if l == "delText":
                out.append(("text", "DEL", el.text or ""))
                return
            if l in ("tab", "ptab"):
                out.append(("text", state, "\t"))
                return
            if l in ("br", "cr"):
                out.append(("text", state, "\n"))
                return
            if l == "noBreakHyphen":
                out.append(("text", state, "-"))
                return
            if l == "sym":
                try:
                    code = int(get_attr(el, "char", "0"), 16)
                    if code >= 0xF000:
                        code -= 0xF000
                    ch = SYMBOL_MAP.get(code, chr(code) if 0x20 <= code < 0x7F else "")
                    if ch:
                        out.append(("text", state, ch))
                except ValueError:
                    pass
                return
            if l in ("pPr", "rPr", "instrText", "delInstrText", "proofErr", "bookmarkStart", "bookmarkEnd",
                     "commentRangeStart", "commentRangeEnd", "commentReference", "annotationRef",
                     "lastRenderedPageBreak", "fldChar", "rPrChange", "pPrChange", "sectPr",
                     "footnoteReference", "endnoteReference", "softHyphen", "permStart", "permEnd"):
                return
            if l == "r":
                rPr = el.find(W + "rPr")
                if rPr is not None:
                    v = rPr.find(W + "vanish")
                    if v is not None and get_attr(v, "val", "true") not in ("0", "false", "off"):
                        return
            if l in ("ins", "moveTo"):
                state = "DEL" if state == "DEL" else "INS"
            elif l in ("del", "moveFrom"):
                state = "DEL"
            elif l == "object":
                self._handle_object(el, state, out)
                return
            elif l == "drawing":
                self._handle_drawing(el, state, out)
                return
            elif l == "pict":
                self._handle_pict(el, state, out)
                return
        if l == "AlternateContent":
            choice = next((c for c in el if ln(c.tag) == "Choice"), None)
            target = choice if choice is not None else next((c for c in el if ln(c.tag) == "Fallback"), None)
            if target is not None:
                for c in target:
                    self._walk(c, state, out)
            return
        for c in el:
            self._walk(c, state, out)

    def _render_segments(self, segs: List[Tuple[str, str, str]]) -> ParaRender:
        merged: List[List[str]] = []
        for kind, state, val in segs:
            if kind == "text" and merged and merged[-1][0] == "text" and merged[-1][1] == state:
                merged[-1][2] += val
            else:
                merged.append([kind, state, val])
        out, plain = [], []
        has_rev, ins_content, other_content, has_content = False, False, False, False
        for kind, state, val in merged:
            if kind == "text":
                val = val.replace("\xa0", " ").replace("\t", " ")
                visible = bool(val) if state != "NORMAL" else bool(val.strip())
            else:
                visible = True
            if visible:
                if state != "NORMAL":
                    has_rev = True
                if state == "INS":
                    ins_content = True
                elif state == "NORMAL":
                    other_content = True
            if state != "DEL" and kind == "text":
                plain.append(val)
            rendered = self._fmt_segment(kind, state, val)
            if rendered.strip():
                has_content = True
            out.append(rendered)
        return ParaRender(text="".join(out), plain="".join(plain), has_rev=has_rev,
                          all_ins=ins_content and not other_content, has_content=has_content)

    def _fmt_segment(self, kind: str, state: str, val: str) -> str:
        mode = self.revision_mode
        if state == "NORMAL":
            return val
        if mode == "accept":
            return val if state == "INS" else ""
        if mode == "reject":
            return val if state == "DEL" else ""
        label = "INSERTED" if state == "INS" else "DELETED"
        if kind != "text":
            return "[%s: %s]" % (label, val)
        return "[%s: %s]" % (label, val) if val else ""

    # ---------------- images / objects ----------------
    def _alt(self, hint: Optional[str] = None) -> str:
        self.figure_no += 1
        alt = (hint or "").strip().replace("\n", " ")
        alt = re.sub(r"[\[\]]", "", alt)[:100]
        return alt or "%s figure %d" % (self.tdoc_id, self.figure_no)

    def _handle_drawing(self, el, state, out):
        hint = None
        for d in iter_skip(el):
            if ln(d.tag) == "docPr":
                hint = get_attr(d, "descr") or get_attr(d, "title")
                break
        for d in iter_skip(el):
            l = ln(d.tag)
            if l == "blip":
                rid = rel_attr(d, "embed", "link")
                if rid:
                    out.append(("image", state, self._image_md(rid, self._alt(hint))))
            elif l == "chart":
                out.append(("image", state, "*[Chart omitted: %s]*" % self._alt(hint)))
            elif l == "relIds" and "diagram" in d.tag:
                out.append(("image", state, "*[SmartArt diagram omitted]*"))
        self._textboxes(el, state, out)

    def _handle_pict(self, el, state, out):
        for d in iter_skip(el):
            if ln(d.tag) == "imagedata":
                rid = rel_attr(d, "id", "relid")
                if rid:
                    out.append(("image", state, self._image_md(rid, self._alt(get_attr(d, "title")))))
        self._textboxes(el, state, out)

    def _handle_object(self, el, state, out):
        img_rid, ole_rid, prog = None, None, None
        for d in iter_skip(el):
            l = ln(d.tag)
            if l == "imagedata" and img_rid is None:
                img_rid = rel_attr(d, "id", "relid")
            elif l == "blip" and img_rid is None:
                img_rid = rel_attr(d, "embed", "link")
            elif l == "OLEObject":
                ole_rid = rel_attr(d, "id")
                prog = get_attr(d, "ProgID")
        if img_rid:
            out.append(("image", state, self._image_md(img_rid, self._alt(prog), ole_rid, prog)))
        elif ole_rid:
            link = self._ole_link(ole_rid, prog)
            if link:
                out.append(("image", state, link.strip()))

    def _textboxes(self, el, state, out):
        for d in iter_skip(el):
            if ln(d.tag) == "txbxContent":
                parts = self._render_blocks(d, top=False, in_cell=True)
                txt = " / ".join(p.replace("\n", " ") for p in parts if p.strip())
                if txt:
                    out.append(("image", state, "[Text box: %s]" % txt))

    def _extract_rel(self, rid: str) -> Optional[Tuple[str, Path]]:
        if rid in self.saved:
            return self.saved[rid]
        info = self.rels.get(rid)
        if not info:
            return None
        full, _typ, external = info
        if external:
            res = ("EXTERNAL", Path(full))
            self.saved[rid] = res
            return res
        data = self._read(full)
        if data is None:
            return None
        base = posixpath.basename(full)
        ext = posixpath.splitext(base)[1].lower()
        if ext in COMPRESSED_METAFILES:
            try:
                data = gzip.decompress(data)
                base = posixpath.splitext(base)[0] + COMPRESSED_METAFILES[ext]
            except OSError:
                pass
        fname = safe_media_name("%s_%s" % (self.prefix, base))
        self.media_dir.mkdir(parents=True, exist_ok=True)
        path = self.media_dir / fname
        path.write_bytes(data)
        self.result.media_files.append(path)
        self.saved[rid] = (fname, path)
        return self.saved[rid]

    def _ole_link(self, ole_rid: Optional[str], prog: Optional[str]) -> str:
        if not ole_rid:
            return ""
        o = self._extract_rel(ole_rid)
        if not o or o[0] == "EXTERNAL":
            return ""
        fname, path = o
        prog_l = (prog or "").lower()
        if "visio" in prog_l or path.suffix.lower() in (".vsdx", ".vsd", ".vsdm"):
            label = "Visio"
        elif prog:
            label = prog
        else:
            label = "OLE object"
        return " *(Embedded %s source: [%s](%s/%s))*" % (label, fname, self.media_rel, fname)

    def _image_md(self, rid: str, alt: str, ole_rid: Optional[str] = None, prog: Optional[str] = None) -> str:
        res = self._extract_rel(rid)
        if res is None:
            md = "*[Image not found in package: %s]*" % alt
        elif res[0] == "EXTERNAL":
            md = "*[Linked external image: %s]*" % res[1]
        else:
            fname, path = res
            suffix = path.suffix.lower()
            if suffix in METAFILE_EXTENSIONS and self.convert_images:
                token = self.metafile_tokens.get(path)
                if token is None:
                    token = "\x00META%d\x00" % len(self.metafile_jobs)
                    self.metafile_tokens[path] = token
                    self.metafile_jobs.append((path, path.with_suffix(".png"), token))
                link = token
            elif suffix in RASTER_TO_CONVERT and self.convert_images:
                link = _convert_raster_to_png(path) or fname
            else:
                link = fname
            self.result.image_count += 1
            md = "![%s](%s/%s)" % (alt, self.media_rel, link)
        return md + self._ole_link(ole_rid, prog)

    def _finalize_media(self, md: str) -> str:
        if not self.metafile_jobs:
            return md
        converter = self.metafile_converter or convert_metafiles
        converter([(s, d) for s, d, _ in self.metafile_jobs], self.image_dpi)
        for src, dst, token in self.metafile_jobs:
            if dst.exists() and dst.stat().st_size > 0:
                md = md.replace(token, dst.name)
                self.result.media_files.append(dst)
                if not self.keep_metafiles:
                    try:
                        src.unlink()
                        self.result.media_files.remove(src)
                    except (OSError, ValueError):
                        pass
            else:
                md = md.replace(token, src.name)
                self.result.warnings.append(
                    "Could not convert %s to PNG (no converter available); linked the original metafile." % src.name)
        return md

    # ---------------- tables ----------------
    def _iter_children(self, el, wanted: str):
        W = self.W
        for c in el:
            l = ln(c.tag)
            if l == wanted:
                yield c
            elif l == "sdt":
                sc = c.find(W + "sdtContent")
                if sc is not None:
                    yield from self._iter_children(sc, wanted)
            elif l in ("customXml", "sdtContent"):
                yield from self._iter_children(c, wanted)

    def _table_rows(self, tbl) -> List[List[List[str]]]:
        W = self.W
        rows: List[List[List[str]]] = []
        for tr in self._iter_children(tbl, "tr"):
            row_state = "NORMAL"
            trPr = tr.find(W + "trPr")
            if trPr is not None:
                if trPr.find(W + "del") is not None:
                    row_state = "DEL"
                elif trPr.find(W + "ins") is not None:
                    row_state = "INS"
            if row_state == "DEL" and self.revision_mode == "accept":
                continue
            if row_state == "INS" and self.revision_mode == "reject":
                continue
            cells: List[List[str]] = []
            for tc in self._iter_children(tr, "tc"):
                span = 1
                tcPr = tc.find(W + "tcPr")
                if tcPr is not None:
                    gs = tcPr.find(W + "gridSpan")
                    if gs is not None:
                        try:
                            span = max(1, int(get_attr(gs, "val", "1")))
                        except ValueError:
                            span = 1
                # Vertically merged continuation cells are normally empty; if they do carry
                # content (e.g. a figure), render it rather than silently dropping it.
                parts = self._render_blocks(tc, top=False, in_cell=True)
                cells.append(parts)
                cells.extend([[] for _ in range(span - 1)])
            if row_state != "NORMAL" and self.revision_mode == "markup" and cells:
                self.result.has_revisions = True
                cells[0] = ["[%s ROW]" % ("INSERTED" if row_state == "INS" else "DELETED")] + cells[0]
            rows.append(cells)
        return rows

    @staticmethod
    def _cell_md(parts: List[str]) -> str:
        txt = "<br>".join(p.strip().replace("\n", "<br>") for p in parts if p.strip())
        return txt.replace("|", "\\|")

    def _render_table(self, tbl, top: bool, in_cell: bool) -> str:
        rows = self._table_rows(tbl)
        rows = [r for r in rows if any(any(p.strip() for p in c) for c in r)]
        if not rows:
            return ""
        if in_cell:
            return "\n".join(" ; ".join(self._cell_md(c) for c in r if any(p.strip() for p in c)) for r in rows)
        flat = " ".join(" ".join(" ".join(c) for c in r) for r in rows)
        if not self.cover_seen_table and "CHANGE REQUEST" in flat.upper():
            self.cover_seen_table = True
            self.result.is_cr = True
            return self._render_cr_cover(rows)
        ncols = max(len(r) for r in rows)
        if ncols == 1:
            # boxed text (1-column tables): render as block quote
            lines = []
            for r in rows:
                for part in (r[0] if r else []):
                    if part.strip():
                        lines.extend("> " + ln_ for ln_ in part.split("\n"))
                        lines.append(">")
            return "\n".join(lines[:-1]) if lines else ""
        md_rows = [[self._cell_md(c) for c in r] + [""] * (ncols - len(r)) for r in rows]
        out = ["| " + " | ".join(md_rows[0]) + " |", "| " + " | ".join(["---"] * ncols) + " |"]
        for r in md_rows[1:]:
            out.append("| " + " | ".join(r) + " |")
        return "\n".join(out)

    def _render_cr_cover(self, rows: List[List[List[str]]]) -> str:
        lines = ["## CR Cover Page", ""]
        for r in rows:
            vals = [" ".join(p.strip() for p in c if p.strip()).strip() for c in r]
            vals = [v for v in vals if v]
            if not vals:
                continue
            pairs, i = [], 0
            while i < len(vals):
                v = vals[i]
                if v.endswith(":") and i + 1 < len(vals) and not vals[i + 1].endswith(":"):
                    key, value = v[:-1].strip(), vals[i + 1]
                    pairs.append("**%s:** %s" % (key, value))
                    norm = COVER_FIELDS.get(key.lower())
                    if norm and norm not in self.result.cover:
                        self.result.cover[norm] = value
                    i += 2
                else:
                    pairs.append(v)
                    i += 1
            lines.append("- " + " | ".join(pairs).replace("\n", " "))
        return "\n".join(lines)

    # ---------------- metadata heuristics ----------------
    def _parse_cover(self):
        cover = self.result.cover
        for txt in self.top_texts:
            m = REVISION_OF_RE.search(txt)
            if m and "revision of" not in cover:
                cover["revision of"] = m.group(1).upper()
            m = COVER_LINE_RE.match(txt)
            if m:
                key = COVER_FIELDS.get(re.sub(r"\s+", " ", m.group(1).strip().lower()))
                if key and key not in cover:
                    cover[key] = m.group(2).strip()

    def _detect_type(self) -> str:
        title = (self.result.cover.get("title") or "") + " " + self.path.name
        if self.result.is_cr:
            return "CR"
        if re.search(r"(?i)\bLS\b|liaison", title) or "to" in self.result.cover:
            return "LS"
        if re.search(r"(?i)\bp-?CR\b", title) or self.marker_count or "spec" in self.result.cover:
            return "pCR"
        if self.result.has_revisions:
            return "pCR"
        return "Discussion"


# --------------------------------------------------------------------------------------
# Image conversion helpers (all optional)
# --------------------------------------------------------------------------------------
_PS_SCRIPT = r"""
param([string]$ListFile, [int]$Dpi)
Add-Type -AssemblyName System.Drawing
foreach ($line in [System.IO.File]::ReadAllLines($ListFile)) {
  if (-not $line) { continue }
  $parts = $line.Split('|'); $src = $parts[0]; $dst = $parts[1]
  try {
    $mf = New-Object System.Drawing.Imaging.Metafile($src)
    $w = [int][Math]::Ceiling($mf.Width * $Dpi / $mf.HorizontalResolution)
    $h = [int][Math]::Ceiling($mf.Height * $Dpi / $mf.VerticalResolution)
    $max = 4000
    if ($w -gt $max -or $h -gt $max) { $s = [Math]::Min($max / $w, $max / $h); $w = [int]($w * $s); $h = [int]($h * $s) }
    if ($w -lt 1) { $w = 1 }; if ($h -lt 1) { $h = 1 }
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $bmp.SetResolution($Dpi, $Dpi)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::White)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.DrawImage($mf, 0, 0, $w, $h)
    $g.Dispose()
    $bmp.Save($dst, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose(); $mf.Dispose()
    Write-Output "OK|$dst"
  } catch { Write-Output "FAIL|$dst|$($_.Exception.Message)" }
}
"""


def _convert_with_powershell(jobs: List[Tuple[Path, Path]], dpi: int) -> None:
    exe = shutil.which("powershell") or shutil.which("pwsh")
    if not exe or not jobs:
        return
    tmp = Path(to_extended_path(tempfile.mkdtemp(prefix="tdoc2md_ps_")))
    try:
        script = tmp / "emf2png.ps1"
        script.write_text(_PS_SCRIPT, encoding="utf-8-sig")
        lst = tmp / "jobs.txt"
        lst.write_text("\n".join("%s|%s" % (to_tool_path(s), to_tool_path(d)) for s, d in jobs), encoding="utf-8-sig")
        proc = subprocess.run([exe, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File",
                               to_tool_path(script), "-ListFile", to_tool_path(lst), "-Dpi", str(dpi)],
                              capture_output=True, text=True, timeout=600)
        for line in (proc.stdout or "").splitlines():
            if line.startswith("FAIL|"):
                log.debug("PowerShell conversion failed: %s", line)
    except (OSError, subprocess.SubprocessError) as e:
        log.debug("PowerShell conversion unavailable: %s", e)
    finally:
        shutil.rmtree(to_extended_path(tmp), ignore_errors=True)


def _convert_with_pillow(jobs: List[Tuple[Path, Path]], dpi: int) -> None:
    try:
        from PIL import Image  # type: ignore
    except ImportError:
        return
    for src, dst in jobs:
        try:
            with Image.open(to_extended_path(src)) as im:
                try:
                    im.load(dpi=dpi)
                except TypeError:
                    im.load()
                im.save(to_extended_path(dst), "PNG")
        except Exception as e:  # noqa: BLE001 - Pillow raises many types
            log.debug("Pillow could not convert %s: %s", src.name, e)


def convert_metafiles(jobs: List[Tuple[Path, Path]], dpi: int = 150) -> None:
    """Convert EMF/WMF files to PNG using the best available converter."""
    for _s, d in jobs:
        d_ext = Path(to_extended_path(d))
        if d_ext.exists():
            try:
                d_ext.unlink()
            except OSError:
                pass

    def pending():
        return [(s, d) for s, d in jobs if not (Path(to_extended_path(d)).exists() and Path(to_extended_path(d)).stat().st_size > 0)]

    if os.name == "nt":
        _convert_with_powershell(pending(), dpi)
    if pending():
        _convert_with_pillow(pending(), dpi)
    if pending():
        backend = _office_backend()
        if backend is not None:
            backend._convert_with_soffice(pending())


def _convert_raster_to_png(path: Path) -> Optional[str]:
    try:
        from PIL import Image  # type: ignore
    except ImportError:
        return None
    try:
        dst = path.with_suffix(".png")
        with Image.open(to_extended_path(path)) as im:
            im.save(to_extended_path(dst), "PNG")
        return dst.name
    except Exception:  # noqa: BLE001
        return None


# --------------------------------------------------------------------------------------
# ZIP document extraction
# --------------------------------------------------------------------------------------
def extract_docs_from_zip(zip_path: Path, tmpdir: Path) -> List[Path]:
    out = []
    zip_path_ext = to_extended_path(zip_path)
    tmpdir_ext = Path(to_extended_path(tmpdir))
    target_dir = tmpdir_ext / zip_path.stem
    try:
        with zipfile.ZipFile(zip_path_ext) as z:
            for name in z.namelist():
                base = posixpath.basename(name)
                if not base or base.startswith("~$") or name.startswith("__MACOSX"):
                    continue
                if posixpath.splitext(base)[1].lower() in DOC_EXTENSIONS:
                    target = target_dir / base
                    target.parent.mkdir(parents=True, exist_ok=True)
                    target.write_bytes(z.read(name))
                    out.append(target)
    except (zipfile.BadZipFile, zlib.error, EOFError, OSError) as e:
        shutil.rmtree(to_extended_path(target_dir), ignore_errors=True)
        raise zipfile.BadZipFile(f"Corrupt or unreadable zip file ({e})") from e
    return out


def discover_documents(input_path: Path, output_dir: Optional[Path]) -> List[Path]:
    input_path = Path(to_extended_path(input_path))
    if output_dir is not None:
        output_dir = Path(to_extended_path(output_dir))
    if input_path.is_file():
        return [input_path]
    found = []
    for p in sorted(input_path.rglob("*")):
        if not p.is_file() or p.name.startswith("~$") or is_within(p, output_dir):
            continue
        suffix = p.suffix.lower()
        if suffix in DOC_EXTENSIONS:
            found.append(p)
        elif suffix == ".zip":
            siblings = [q for q in p.parent.iterdir() if q.is_file() and q.suffix.lower() in DOC_EXTENSIONS]
            if not siblings:
                found.append(p)
    return found


# --------------------------------------------------------------------------------------
# Excel TDoc list (stdlib xlsx reader)
# --------------------------------------------------------------------------------------
def _col_index(ref: str) -> int:
    idx = 0
    for ch in ref:
        if ch.isalpha():
            idx = idx * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return idx - 1


def read_xlsx_records(path: Path) -> List[Dict[str, str]]:
    """Read the first worksheet of an .xlsx into a list of {lower-case header: value} dicts."""
    with zipfile.ZipFile(to_extended_path(path)) as z:
        names = set(z.namelist())
        shared: List[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root:
                if ln(si.tag) == "si":
                    shared.append("".join(n.text or "" for n in iter_skip(si, ("rPh",)) if ln(n.tag) == "t")
                                  if len(si) else (si.text or ""))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheet = next((e for e in wb.iter() if ln(e.tag) == "sheet"), None)
        rid = rel_attr(sheet, "id") if sheet is not None else None
        sheet_path = "xl/worksheets/sheet1.xml"
        if rid and "xl/_rels/workbook.xml.rels" in names:
            for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
                if r.get("Id") == rid:
                    t = r.get("Target", "")
                    sheet_path = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))
        sh = ET.fromstring(z.read(sheet_path))
        raw_rows: List[Dict[int, str]] = []
        for row in sh.iter():
            if ln(row.tag) != "row":
                continue
            vals: Dict[int, str] = {}
            for c in row:
                if ln(c.tag) != "c":
                    continue
                ref = c.get("r", "")
                col = _col_index(ref) if ref else len(vals)
                t, v = c.get("t"), None
                for cc in c:
                    if ln(cc.tag) == "v":
                        v = cc.text
                    elif ln(cc.tag) == "is":
                        v = "".join(x.text or "" for x in cc.iter() if ln(x.tag) == "t")
                if t == "s" and v is not None:
                    try:
                        v = shared[int(v)]
                    except (ValueError, IndexError):
                        pass
                elif t == "b":
                    v = "TRUE" if v == "1" else "FALSE"
                vals[col] = (v or "").strip()
            raw_rows.append(vals)
    raw_rows = [r for r in raw_rows if any(r.values())]
    if not raw_rows:
        return []
    header = {i: h.strip().lower() for i, h in raw_rows[0].items() if h.strip()}
    return [{header[i]: r.get(i, "") for i in header} for r in raw_rows[1:]]


def load_tdoc_metadata(excel: Optional[Path], input_path: Path) -> Dict[str, Dict[str, str]]:
    candidates: List[Path] = []
    if excel:
        candidates = [Path(to_extended_path(excel))]
    elif input_path.is_dir():
        input_ext = Path(to_extended_path(input_path))
        candidates = sorted(p for p in input_ext.glob("*.xlsx") if not p.name.startswith("~$"))
    for cand in candidates:
        try:
            records = read_xlsx_records(cand)
        except Exception as e:  # noqa: BLE001
            log.warning("Could not read %s: %s", cand.name, e)
            continue
        if records and "tdoc" in records[0]:
            log.info("Using TDoc metadata from %s (%d rows)", cand.name, len(records))
            return {r["tdoc"].strip().upper(): r for r in records if r.get("tdoc")}
    return {}


# --------------------------------------------------------------------------------------
# Markdown assembly, index and bulk corpus
# --------------------------------------------------------------------------------------
@dataclass
class ExportedDoc:
    tdoc_id: str
    name: str
    md_path: Path
    text: str
    meta: Dict[str, str]
    doc_type: str
    agenda_item: str
    title: str
    status: str
    images: int
    warnings: List[str]


def _pick(meta: Dict[str, str], cover: Dict[str, str], *keys: str) -> str:
    for k in keys:
        v = (meta.get(k) or "").strip()
        if v:
            return v
    for k in keys:
        v = (cover.get(k) or "").strip()
        if v:
            return v
    return ""


def build_header(tdoc_id: str, meta: Dict[str, str], res: DocResult, original: str, revisions: str) -> str:
    cover = res.cover
    title = _pick(meta, cover, "title")
    is_tdoc = bool(
        TDOC_RE.search(tdoc_id)
        or (meta and any(k in meta for k in ("tdoc", "tdoc_id", "agenda item", "work item", "spec")))
        or any(k in cover for k in ("source", "agenda item", "work item", "spec"))
    )
    if is_tdoc:
        doc_type = _pick(meta, {}, "type") or "%s (detected)" % res.detected_type
        fields = [
            ("TDoc", tdoc_id), ("Title", title), ("Source", _pick(meta, cover, "source")),
            ("Type", doc_type), ("Agenda Item", _pick(meta, cover, "agenda item")),
            ("Status", _pick(meta, cover, "tdoc status", "status")),
            ("Document for", _pick(meta, cover, "for", "document for")),
            ("Work Item", _pick(meta, cover, "work item")), ("Spec", _pick(meta, cover, "spec")),
            ("Version", _pick(meta, cover, "version")), ("Revision of", _pick(meta, cover, "revision of")),
            ("Related TDocs", _pick(meta, cover, "related tdocs")),
            ("Original file", "`%s`" % original),
            ("Exporter", "tdoc2md v%s (revisions: %s)" % (__version__, revisions)),
        ]
        lines = ["# %s%s" % (tdoc_id, (": " + title) if title else ""), ""]
        lines += ["- **%s:** %s" % (k, v.replace("\n", " ")) for k, v in fields if v]
        lines.append("")
        if revisions == "markup":
            lines += [MARKUP_LEGEND, ""]
        abstract = _pick(meta, cover, "abstract")
        if abstract:
            lines += ["**Abstract:** " + abstract.replace("\n", " "), ""]
        lines += ["---", ""]
        return "\n".join(lines)
    else:
        doc_name = title or tdoc_id.replace("_", " ")
        fields = [
            ("Original file", "`%s`" % original),
            ("Exporter", "Word-to-Markdown v%s (revisions: %s)" % (__version__, revisions)),
        ]
        lines = ["# %s" % doc_name, ""]
        lines += ["- **%s:** %s" % (k, v.replace("\n", " ")) for k, v in fields if v]
        lines.append("")
        if revisions == "markup":
            lines += [MARKUP_LEGEND, ""]
        lines += ["---", ""]
        return "\n".join(lines)


def category_for(doc_type: str) -> str:
    t = doc_type or ""
    if "CR" in t:
        return "CRs"
    if re.search(r"\bLS\b", t):
        return "LSs"
    return "Discussion_Papers"


def write_index(out_dir: Path, docs: List[ExportedDoc]) -> Path:
    out_dir = Path(to_extended_path(out_dir))
    out_dir.mkdir(parents=True, exist_ok=True)
    lines = ["# TDoc Markdown Export", "", "Generated by tdoc2md v%s. %d document(s)." % (__version__, len(docs)), "",
             "| TDoc | Title | Type | Agenda Item | Status | Images |", "| --- | --- | --- | --- | --- | --- |"]
    for d in sorted(docs, key=lambda x: x.name):
        title = (d.title or "").replace("|", "\\|")
        lines.append("| [%s](%s) | %s | %s | %s | %s | %d |" % (
            d.name, d.md_path.name, title, d.doc_type, d.agenda_item, d.status, d.images))
    path = out_dir / "index.md"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def write_bulk_corpus(out_dir: Path, docs: List[ExportedDoc], max_chars: int, prompt: str) -> List[Path]:
    out_dir = Path(to_extended_path(out_dir))
    out_dir.mkdir(parents=True, exist_ok=True)
    groups: Dict[str, Dict[str, List[ExportedDoc]]] = {}
    for d in docs:
        ai = sanitize_filename((d.agenda_item or "Unknown").replace(" ", "_"))
        groups.setdefault(ai, {}).setdefault(category_for(d.doc_type), []).append(d)
    written: List[Path] = []
    for ai, cats in sorted(groups.items()):
        for cat, items in sorted(cats.items()):
            header = ("# 3GPP LLM Corpus\n**Agenda Item:** %s\n**Document Category:** %s\n\n"
                      "## Context Guide for LLM\n%s\n\n---\n\n" % (ai, cat, prompt))
            chunks: List[str] = []
            current = header
            has = False
            for d in items:
                piece = d.text.rstrip() + "\n\n---\n\n"
                if len(piece) + len(header) > max_chars:
                    log.warning("%s (%d chars) exceeds --max-chars; written to its own part.", d.name, len(piece))
                if has and len(current) + len(piece) > max_chars:
                    chunks.append(current)
                    current, has = header, False
                current += piece
                has = True
            if has:
                chunks.append(current)
            for i, text in enumerate(chunks, 1):
                suffix = "_Part%d" % i if len(chunks) > 1 else ""
                path = out_dir / ("AI_%s_%s%s.md" % (ai, sanitize_filename(cat), suffix))
                path.write_text(text, encoding="utf-8")
                written.append(path)
    return written


# --------------------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------------------
def convert_docx_file(docx_path: Path, out_dir: Path, tdoc_id: Optional[str] = None, name: Optional[str] = None,
                      meta: Optional[Dict[str, str]] = None, original_name: Optional[str] = None,
                      revisions: str = "markup", image_dpi: int = 150, convert_images: bool = True,
                      keep_metafiles: bool = False, media_dir: Optional[Path] = None,
                      media_rel: str = "media",
                      cancellation_check: Optional[Callable[[], bool]] = None,
                      metafile_converter: Optional[Callable[[List[Tuple[Path, Path]], int], None]] = None,
                      include_header: bool = True) -> ExportedDoc:
    """Convert a single .docx to <out_dir>/<name>.md (+ media in <out_dir>/media). Library entry point."""
    docx_path = Path(to_extended_path(docx_path))
    out_dir = Path(to_extended_path(out_dir))
    out_dir.mkdir(parents=True, exist_ok=True)
    tdoc_id = tdoc_id or guess_tdoc_id(docx_path)
    name = name or tdoc_id
    meta = meta or {}
    conv = DocxConverter(docx_path, tdoc_id, media_dir if media_dir is not None else out_dir / "media", media_rel, prefix=name, revision_mode=revisions,
                         image_dpi=image_dpi, convert_images=convert_images, keep_metafiles=keep_metafiles,
                         cancellation_check=cancellation_check, metafile_converter=metafile_converter)
    res = conv.convert()
    header_text = build_header(tdoc_id, meta, res, original_name or docx_path.name, revisions) if include_header else ""
    text = header_text + res.markdown
    md_path = out_dir / ("%s.md" % sanitize_filename(name))
    md_path.write_text(text, encoding="utf-8")
    doc_type = _pick(meta, {}, "type") or res.detected_type
    return ExportedDoc(
        tdoc_id=tdoc_id, name=name, md_path=md_path, text=text, meta=meta, doc_type=doc_type,
        agenda_item=_pick(meta, res.cover, "agenda item"), title=_pick(meta, res.cover, "title"),
        status=_pick(meta, res.cover, "tdoc status", "status"), images=res.image_count, warnings=res.warnings)


def run(input_path: Path, output_dir: Optional[Path] = None, excel: Optional[Path] = None, bulk: bool = False,
        max_chars: int = 200000, revisions: str = "markup", image_dpi: int = 150, convert_images: bool = True,
        keep_metafiles: bool = False, prompt: str = DEFAULT_CORPUS_PROMPT,
        sensitivity_label_id: Optional[str] = None, sensitivity_label_name: Optional[str] = None,
        word_config: Optional[Union[dict, Path, str]] = None,
        progress_callback: Optional[Callable[[str], None]] = None,
        cancellation_check: Optional[Callable[[], bool]] = None,
        document_converter: Optional[Callable[[Path, Path, dict], Optional[Path]]] = None,
        metafile_converter: Optional[Callable[[List[Tuple[Path, Path]], int], None]] = None) -> Tuple[List[ExportedDoc], List[str]]:
    input_path = Path(to_extended_path(input_path))
    if not input_path.exists():
        raise FileNotFoundError(strip_extended_path(input_path))
    if output_dir is None:
        base = input_path if input_path.is_dir() else input_path.parent
        output_dir = base / "Markdown"
    output_dir = Path(to_extended_path(output_dir))
    output_dir.mkdir(parents=True, exist_ok=True)

    metadata = load_tdoc_metadata(excel, input_path)
    sources = discover_documents(input_path, output_dir)
    log.info("Found %d document(s) under %s", len(sources), strip_extended_path(input_path))

    backend = _office_backend() if document_converter is None else None
    if backend is not None:
        w_cfg = backend.load_word_config(word_config, label_id_override=sensitivity_label_id,
                                         label_name_override=sensitivity_label_name)
    else:
        if word_config is not None and not isinstance(word_config, dict):
            raise ValueError("A configuration file requires the optional office adapter; supply a dict with callbacks")
        w_cfg = dict(word_config or {})
        if sensitivity_label_id is not None:
            w_cfg["sensitivity_level_label_id"] = sensitivity_label_id
        if sensitivity_label_name is not None:
            w_cfg["sensitivity_level_label_name"] = sensitivity_label_name
    exported: List[ExportedDoc] = []
    failures: List[str] = []
    used: Dict[str, int] = {}
    tmp = Path(to_extended_path(tempfile.mkdtemp(prefix="tdoc2md_")))
    try:
        work: List[Tuple[Path, Path]] = []  # (original path, path to convert)
        for src in sources:
            if cancellation_check and cancellation_check():
                break
            if src.suffix.lower() == ".zip":
                try:
                    docs = extract_docs_from_zip(src, tmp)
                    work.extend((src, p) for p in docs)
                except Exception as e:
                    tdoc_id = guess_tdoc_id(src.with_suffix(".docx"))
                    msg = "%s (%s): %s" % (tdoc_id, src.name, e)
                    log.error("Failed: %s", msg)
                    failures.append(msg)
            else:
                work.append((src, src))
        for original, path in work:
            if cancellation_check and cancellation_check():
                break
            tdoc_id = guess_tdoc_id(original if original.suffix.lower() != ".zip" else original.with_suffix(".docx"))
            if progress_callback:
                progress_callback("Processing %s..." % tdoc_id)
            try:
                if path.suffix.lower() == ".doc":
                    converter = document_converter or (backend.convert_doc_to_docx if backend else None)
                    converted = converter(path, tmp, w_cfg) if converter else None
                    if not converted:
                        raise RuntimeError("legacy .doc requires MS Word (pywin32) or LibreOffice for conversion")
                    path = converted
                name = tdoc_id
                if name in used:
                    used[name] += 1
                    name = "%s_%s" % (tdoc_id, safe_media_name(Path(path).stem)[:40])
                else:
                    used[name] = 1
                meta = metadata.get(tdoc_id.upper()) or metadata.get(base_tdoc_id(tdoc_id)) or {}
                log.info("Converting %s (%s)", tdoc_id, original.name)
                doc = convert_docx_file(path, output_dir, tdoc_id=tdoc_id, name=name, meta=meta,
                                        original_name=original.name, revisions=revisions, image_dpi=image_dpi,
                                        convert_images=convert_images, keep_metafiles=keep_metafiles,
                                        cancellation_check=cancellation_check, metafile_converter=metafile_converter)
                for w in doc.warnings:
                    log.warning("%s: %s", tdoc_id, w)
                exported.append(doc)
            except Exception as e:  # noqa: BLE001 - keep batch going
                msg = "%s (%s): %s" % (tdoc_id, original.name, e)
                log.error("Failed: %s", msg)
                failures.append(msg)
    finally:
        shutil.rmtree(to_extended_path(tmp), ignore_errors=True)

    if exported:
        write_index(output_dir, exported)
        if bulk:
            for p in write_bulk_corpus(output_dir, exported, max_chars, prompt):
                log.info("Wrote corpus %s", p.name)
    return exported, failures


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Convert 3GPP TDocs (.docx/.doc/.zip) to Markdown with track changes, "
                                             "tables, images and embedded Visio diagrams.")
    ap.add_argument("input", help="A .docx/.doc/.zip file, a TDoc folder, or a meeting folder with TDoc sub-folders")
    ap.add_argument("-o", "--output-dir", help="Output directory (default: <input>/Markdown)")
    ap.add_argument("--excel", help="3GPP TDoc list (.xlsx) with columns TDoc, Title, Source, Type, Agenda Item, ... "
                                    "(default: auto-detect *.xlsx in the input folder)")
    ap.add_argument("--bulk", action="store_true", help="Also write per Agenda Item / category corpus files")
    ap.add_argument("--max-chars", type=int, default=200000, help="Max characters per bulk corpus file")
    ap.add_argument("--revisions", choices=["markup", "accept", "reject"], default="markup",
                    help="markup: show [INSERTED]/[DELETED]; accept: final text; reject: original text")
    ap.add_argument("--image-dpi", type=int, default=150, help="Resolution for EMF/WMF to PNG conversion")
    ap.add_argument("--no-image-convert", action="store_true", help="Keep EMF/WMF/TIFF images as extracted")
    ap.add_argument("--keep-metafiles", action="store_true", help="Keep original EMF/WMF next to converted PNGs")
    ap.add_argument("--sensitivity-label-id", help="Corporate sensitivity label GUID for Word COM save operations")
    ap.add_argument("--sensitivity-label-name", help="Corporate sensitivity label Name (e.g. OFFEN)")
    ap.add_argument("--word-config", help="Path to custom word_config.json file")
    ap.add_argument("-q", "--quiet", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else (logging.WARNING if args.quiet else logging.INFO),
                        format="%(levelname)s: %(message)s")
    try:
        exported, failures = run(Path(args.input), Path(args.output_dir) if args.output_dir else None,
                                 Path(args.excel) if args.excel else None, args.bulk, args.max_chars,
                                 args.revisions, args.image_dpi, not args.no_image_convert, args.keep_metafiles,
                                 sensitivity_label_id=args.sensitivity_label_id,
                                 sensitivity_label_name=args.sensitivity_label_name,
                                 word_config=args.word_config)
    except FileNotFoundError as e:
        log.error("Input not found: %s", e)
        return 2
    out = exported[0].md_path.parent if exported else None
    display_out = strip_extended_path(out) if out else ""
    print("Converted %d document(s)%s%s" % (len(exported), (" -> %s" % display_out) if display_out else "",
                                            ("; %d failed" % len(failures)) if failures else ""))
    for f in failures:
        print("  FAILED: " + f)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
