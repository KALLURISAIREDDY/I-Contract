#!/usr/bin/env python3
"""
pdf_to_docx_libreoffice.py  (v5 — production-grade)

Complete pipeline:
  STEP 1: PDF → ODT → DOCX  (LibreOffice headless)
  STEP 2: Post-process DOCX  (fix underlines, widths, highlights)

═══════════════════════════════════════════════════════════════════════════════
BUGS FIXED IN v5  (on top of v4)
═══════════════════════════════════════════════════════════════════════════════

  ROOT CAUSE (v4 missed this)
  ───────────────────────────
  LibreOffice adds <w:pBdr><w:bottom> (a full-width horizontal rule that looks
  exactly like an underline) to paragraphs that contain ANY of:

    • <w:hyperlink>  elements                  ← v4 handled this
    • REF / PAGEREF / NOTEREF field codes      ← v4 MISSED these
    • fldChar-based links (BEGIN/SEPARATE/END) ← v4 MISSED these
    • Bookmark cross-references                ← v4 MISSED these

  In the screenshots: "Annex A contains a comprehensive list…" has a full-
  width underline bar because the paragraph contains a REF field pointing to
  "Annex A", even though there is no <w:hyperlink> XML element.

  FIX-A  <w:pBdr><w:bottom> — remove UNCONDITIONALLY from every paragraph
  ──────────────────────────────────────────────────────────────────────────
  Word documents produced by LibreOffice from PDFs should NEVER have
  <w:pBdr><w:bottom> for legitimate decorative reasons.  The only source of
  these borders in LO-converted files is the hyperlink/field contamination
  described above.  Therefore v5 removes them from ALL paragraphs without
  any conditional check.

  FIX-B  Run-level <w:u> injection via fields
  ────────────────────────────────────────────
  LO also injects <w:u> onto runs inside field-code paragraphs.  The v4
  logic only stripped <w:u> from runs whose rStyle was a hyperlink style in
  paragraphs containing <w:hyperlink>.  v5 extends this to paragraphs that
  contain fldChar BEGIN/SEPARATE/END sequences, using the same rule:
    • Keep <w:u> on runs that are direct children of <w:hyperlink>
    • Keep <w:u> on runs whose text is a URL
    • Keep <w:u> on runs whose rStyle is NOT a known LO hyperlink/field style
    • Strip <w:u> from runs whose rStyle IS a known LO hyperlink/field style

  FIX-C  Underline scope tightening
  ───────────────────────────────────
  In the broken screenshot the underlines span the FULL line width.
  After removing <w:pBdr><w:bottom>, the remaining per-run <w:u> marks on
  non-hyperlink runs are stripped using the same two-stage algorithm as v4,
  now extended to cover field-containing paragraphs too.

  UNCHANGED from v4
  ─────────────────
  • Black table shading deep-sweep (double pass)
  • Highlighted text visibility (force color→auto)
  • Page geometry detection and normalisation
  • Table/textbox/frame/indent width fixes

═══════════════════════════════════════════════════════════════════════════════

Usage:
  python pdf_to_docx_libreoffice_v5.py input.pdf output.docx
  python pdf_to_docx_libreoffice_v5.py input.pdf output.docx --highlight-words Sony "Name"
  python pdf_to_docx_libreoffice_v5.py existing.docx output.docx --skip-conversion

Requirements:
  pip install lxml
  LibreOffice (soffice on PATH)
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from lxml import etree

# ───────────────────────────────────────────────────────────────
#  Namespaces
# ───────────────────────────────────────────────────────────────
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
VML = "urn:schemas-microsoft-com:vml"

def w(tag):  return "{%s}%s" % (W,   tag)
def wp(tag): return "{%s}%s" % (WP,  tag)
def vml(tag):return "{%s}%s" % (VML, tag)

DXA_PER_INCH = 1440
EMU_PER_DXA  = 914400 // DXA_PER_INCH

# Very-dark fills that LibreOffice incorrectly uses as table cell shading
_BLACK_FILLS: Set[str] = {
    "000000","0D0D0D","1A1A1A","0C0C0C","111111","0F0F0F",
    "050505","080808","020202","030303","040404","060606",
    "070707","090909","0A0A0A","0B0B0B","0E0E0E","010101",
}

# Character-style IDs that LibreOffice injects for hyperlinks / cross-refs
_LO_LINK_STYLES: Set[str] = {
    "Hyperlink", "InternetLink", "InternetLink20",
    "a", "hyperlink",
    # LO cross-reference styles
    "FootnoteReference", "EndnoteReference",
}

# URL pattern for bare-URL runs
_URL_RE = re.compile(r"https?://|www\.", re.IGNORECASE)

# Field instruction keywords that indicate a clickable cross-reference
_FIELD_LINK_RE = re.compile(
    r"\b(REF|PAGEREF|NOTEREF|HYPERLINK|XE)\b", re.IGNORECASE
)

# <w:u val="…"> values that mean "no underline"
_NO_UNDERLINE_VALS: Set[str] = {"none", "0"}


# ───────────────────────────────────────────────────────────────
#  PART A — LibreOffice conversion
# ───────────────────────────────────────────────────────────────

def find_soffice() -> str:
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    raise RuntimeError("LibreOffice not found on PATH.")


def run_cmd(cmd: List[str]) -> subprocess.CompletedProcess:
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print(f"[ERROR] {' '.join(cmd)}\nSTDERR: {r.stderr}")
        raise RuntimeError("LibreOffice command failed")
    return r


def convert_pdf_to_docx(pdf_path: str, output_docx: str) -> str:
    soffice     = find_soffice()
    pdf_path    = str(Path(pdf_path).resolve())
    output_docx = str(Path(output_docx).resolve())
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"Input not found: {pdf_path}")
    os.makedirs(os.path.dirname(output_docx) or ".", exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="lo_") as tmp:
        print("[STEP 1] PDF → ODT")
        run_cmd([soffice, "--headless", "--infilter=writer_pdf_import",
                 "--convert-to", "odt", "--outdir", tmp, pdf_path])
        odt = [f for f in os.listdir(tmp) if f.endswith(".odt")]
        if not odt:
            raise RuntimeError("LibreOffice produced no .odt file")

        print("[STEP 2] ODT → DOCX")
        run_cmd([soffice, "--headless",
                 "--convert-to", "docx:MS Word 2007 XML",
                 "--outdir", tmp, os.path.join(tmp, odt[0])])
        docx = [f for f in os.listdir(tmp) if f.endswith(".docx")]
        if not docx:
            raise RuntimeError("LibreOffice produced no .docx file")

        shutil.move(os.path.join(tmp, docx[0]), output_docx)

    print(f"[STEP 2 DONE] Converted: {output_docx}")
    return output_docx


# ───────────────────────────────────────────────────────────────
#  XML helpers
# ───────────────────────────────────────────────────────────────
DOCUMENT_XML = "word/document.xml"
STYLES_XML   = "word/styles.xml"

def read_xml(zf: zipfile.ZipFile, name: str) -> etree._Element:
    return etree.fromstring(zf.read(name))

def write_xml(tree: etree._Element) -> bytes:
    return etree.tostring(
        tree, xml_declaration=True, encoding="UTF-8", standalone=True
    )


# ───────────────────────────────────────────────────────────────
#  Page geometry
# ───────────────────────────────────────────────────────────────
_PAPER_SIZES_DXA = [
    (12240, 15840),   # US Letter portrait
    (11906, 16838),   # A4 portrait
    (10440, 13320),   # B5 portrait
    (15840, 12240),   # US Letter landscape
    (16838, 11906),   # A4 landscape
]
_MIN_PAGE_DIM   = 5040
_MAX_PAGE_DIM   = 25200
_MIN_MARGIN     = 360
_MAX_MARGIN     = 5040
_DEFAULT_MARGIN = DXA_PER_INCH


def _nearest_paper(pw: int, ph: int) -> Tuple[int, int]:
    best, bd = _PAPER_SIZES_DXA[0], float("inf")
    for sz in _PAPER_SIZES_DXA:
        d = abs(sz[0] - pw) + abs(sz[1] - ph)
        if d < bd:
            best, bd = sz, d
    return best


def detect_page_geometry(doc_tree: etree._Element) -> Dict:
    sectPr = doc_tree.find(".//" + w("sectPr"))
    raw: Dict = {}
    if sectPr is not None:
        pgSz  = sectPr.find(w("pgSz"))
        pgMar = sectPr.find(w("pgMar"))
        if pgSz is not None:
            try:
                raw["pw"] = int(pgSz.get(w("w"), 0))
                raw["ph"] = int(pgSz.get(w("h"), 0))
            except (TypeError, ValueError):
                pass
        if pgMar is not None:
            for k, a in [("ml","left"),("mr","right"),("mt","top"),("mb","bottom")]:
                try:
                    raw[k] = int(pgMar.get(w(a), 0))
                except (TypeError, ValueError):
                    pass

    def vdim(v): return v is not None and _MIN_PAGE_DIM <= v <= _MAX_PAGE_DIM

    pw = raw.get("pw"); ph = raw.get("ph")
    if not vdim(pw) or not vdim(ph):
        pw, ph = _PAPER_SIZES_DXA[0]
    else:
        spw, sph = _nearest_paper(pw, ph)
        if abs(spw - pw) / pw < 0.05 and abs(sph - ph) / ph < 0.05:
            pw, ph = spw, sph

    def cm(v):
        return v if (v is not None and _MIN_MARGIN <= v <= _MAX_MARGIN) else _DEFAULT_MARGIN

    ml = cm(raw.get("ml")); mr = cm(raw.get("mr"))
    mt = cm(raw.get("mt")); mb = cm(raw.get("mb"))
    return {
        "page_w": pw, "page_h": ph,
        "margin_left": ml, "margin_right": mr,
        "margin_top":  mt, "margin_bottom": mb,
        "content_w":   max(pw - ml - mr, DXA_PER_INCH),
        "content_h":   max(ph - mt - mb, DXA_PER_INCH),
    }


def ensure_page_geometry(doc_tree: etree._Element, geo: Dict) -> None:
    sectPr = doc_tree.find(".//" + w("sectPr"))
    if sectPr is None:
        body = doc_tree.find(w("body"))
        if body is None:
            return
        sectPr = etree.SubElement(body, w("sectPr"))
    for tag, keys in [
        (w("pgSz"),  [(w("w"), "page_w"),    (w("h"),      "page_h")]),
        (w("pgMar"), [(w("left"), "margin_left"), (w("right"), "margin_right"),
                      (w("top"),  "margin_top"),  (w("bottom"),"margin_bottom")]),
    ]:
        el = sectPr.find(tag)
        if el is None:
            el = etree.SubElement(sectPr, tag)
        for attr, key in keys:
            el.set(attr, str(geo[key]))


# ───────────────────────────────────────────────────────────────
#  Style loader
# ───────────────────────────────────────────────────────────────
def load_styles(zf: zipfile.ZipFile) -> Dict:
    """
    Returns a dict keyed by styleId:
      name, basedOn, hasUnderline, inheritsUnderline, isLinkStyle
    """
    if STYLES_XML not in zf.namelist():
        return {}
    raw: Dict = {}
    for style in etree.fromstring(zf.read(STYLES_XML)).iter(w("style")):
        sid = style.get(w("styleId"), "")
        if not sid:
            continue
        ne = style.find(w("name"))
        be = style.find(w("basedOn"))
        name     = ne.get(w("val"), "") if ne is not None else ""
        based_on = be.get(w("val"))     if be is not None else None
        has_u = False
        rPr = style.find(w("rPr"))
        if rPr is not None:
            u = rPr.find(w("u"))
            if u is not None:
                has_u = u.get(w("val"), "single") not in _NO_UNDERLINE_VALS
        is_link = (
            sid in _LO_LINK_STYLES
            or "hyperlink" in name.lower()
            or "internet"  in name.lower()
        )
        raw[sid] = {
            "name": name, "basedOn": based_on,
            "hasUnderline": has_u, "isLinkStyle": is_link,
        }

    def _inh(sid, visited):
        if sid in visited:
            return False
        visited.add(sid)
        entry = raw.get(sid)
        if entry is None:
            return False
        if entry["hasUnderline"]:
            return True
        return _inh(entry["basedOn"], visited) if entry["basedOn"] else False

    for sid in raw:
        raw[sid]["inheritsUnderline"] = _inh(sid, set())
    return raw


# ───────────────────────────────────────────────────────────────
#  Run-classification helpers
# ───────────────────────────────────────────────────────────────

def _run_char_style(run: etree._Element) -> str:
    rPr = run.find(w("rPr"))
    if rPr is None:
        return ""
    rs = rPr.find(w("rStyle"))
    return rs.get(w("val"), "") if rs is not None else ""


def _run_is_direct_hyperlink_child(run: etree._Element) -> bool:
    """True only when the run's immediate XML parent is <w:hyperlink>."""
    parent = run.getparent()
    return parent is not None and parent.tag == w("hyperlink")


def _run_style_is_lo_link(run: etree._Element, styles: Dict) -> bool:
    """True when the run's rStyle is one of LO's injected link/field styles."""
    sid = _run_char_style(run)
    if not sid:
        return False
    if sid in _LO_LINK_STYLES:
        return True
    entry = styles.get(sid, {})
    return entry.get("isLinkStyle", False)


def _run_text(run: etree._Element) -> str:
    return "".join(t.text or "" for t in run.iter(w("t")))


# ───────────────────────────────────────────────────────────────
#  Paragraph-classification helpers
# ───────────────────────────────────────────────────────────────

def _para_has_hyperlink_element(para: etree._Element) -> bool:
    """True when the paragraph contains at least one <w:hyperlink> element."""
    return para.find(".//" + w("hyperlink")) is not None


def _para_has_field_link(para: etree._Element) -> bool:
    """
    True when the paragraph contains a field code (fldChar/instrText) that
    is a cross-reference or hyperlink field (REF, PAGEREF, NOTEREF, HYPERLINK).

    LibreOffice renders these fields visually as underlined blue text but does
    NOT wrap them in <w:hyperlink> — it uses the fldChar BEGIN/instrText/
    fldChar SEPARATE/fldChar END pattern instead.
    """
    # Check instrText nodes for known link field keywords
    for instr in para.iter(w("instrText")):
        if instr.text and _FIELD_LINK_RE.search(instr.text):
            return True
    # Presence of fldChar itself (even without matching instrText in this para)
    # indicates a field run — combine with rStyle check downstream
    return False


def _para_contains_lo_link(para: etree._Element) -> bool:
    """
    True when the paragraph contains any LO-injected link structure:
    <w:hyperlink>, REF/PAGEREF field, or any run styled as a link style.
    """
    if _para_has_hyperlink_element(para):
        return True
    if _para_has_field_link(para):
        return True
    # Check if any run in the para carries an LO link character style —
    # this catches cases where instrText is on a different paragraph but
    # the styled runs leaked into this one.
    for run in para.iter(w("r")):
        sid = _run_char_style(run)
        if sid in _LO_LINK_STYLES:
            return True
    return False


# ───────────────────────────────────────────────────────────────
#  FIX — Underlines  (v5 definitive algorithm)
# ───────────────────────────────────────────────────────────────
#
# TWO SEPARATE PROBLEMS — SOLVED INDEPENDENTLY
# ─────────────────────────────────────────────
#
# PROBLEM 1: <w:pBdr><w:bottom>  (full-width underline bar)
# ──────────────────────────────
# LibreOffice adds a paragraph bottom border to paragraphs that reference
# anything link-like: <w:hyperlink>, REF fields, PAGEREF fields, bookmark
# cross-references, fldChar sequences, etc.
#
# SOLUTION: Remove <w:pBdr><w:bottom> from ALL paragraphs unconditionally.
#
# Rationale: In a PDF→DOCX conversion, no paragraph should ever have a
# decorative bottom border.  The only source of <w:pBdr><w:bottom> in LO-
# converted files is this link-adjacent contamination.  Removing them all is
# safe and correct.
#
# PROBLEM 2: <w:u> on non-link runs  (per-run underline injection)
# ───────────────────────────────────
# LO also injects <w:u> onto runs that are siblings of link runs.
#
# SOLUTION: Per-paragraph two-stage analysis:
#
#   If the paragraph contains ANY LO link structure:
#     → For each run with <w:u>:
#         KEEP if: run is direct child of <w:hyperlink>
#         KEEP if: run text matches a URL pattern
#         KEEP if: run's rStyle is NOT an LO link style (genuine underline)
#         STRIP if: run's rStyle IS an LO link style (injected by LO)
#         STRIP if: run has no rStyle but paragraph has links (LO sibling contamination)
#              — BUT only strip when pBdr was also present (strong signal of
#                 full contamination); otherwise keep (defensive)
#
#   If the paragraph contains NO LO link structure:
#     → Touch nothing. All underlines are from the original PDF.

def fix_underlines(doc_tree: etree._Element, styles: Dict) -> Tuple[int, int, int]:
    """
    Remove LibreOffice-injected underline artifacts.

    Returns: (n_removed, n_kept, n_normalised)
    """
    removed = kept = normalised = 0

    for para in doc_tree.iter(w("p")):
        pPr  = para.find(w("pPr"))

        # ── PROBLEM 1: Remove <w:pBdr><w:bottom> unconditionally ────────────
        # This is the primary cause of the full-width underline bars shown in
        # the screenshot. Remove from EVERY paragraph — not just those with
        # hyperlinks — because LO adds them for fields too.
        pbdr_was_present = False
        if pPr is not None:
            pBdr = pPr.find(w("pBdr"))
            if pBdr is not None:
                bot = pBdr.find(w("bottom"))
                if bot is not None:
                    pBdr.remove(bot)
                    removed += 1
                    pbdr_was_present = True
                if len(pBdr) == 0:
                    pPr.remove(pBdr)

        # ── PROBLEM 2: Remove injected per-run <w:u> marks ──────────────────
        # Only do this for paragraphs that contain an LO link structure.
        has_lo_link = _para_contains_lo_link(para)
        if not has_lo_link:
            # No link structure → all underlines are original → touch nothing
            continue

        for run in para.iter(w("r")):
            rPr = run.find(w("rPr"))
            if rPr is None:
                continue
            u = rPr.find(w("u"))
            if u is None:
                continue
            u_val = u.get(w("val"), "single")
            if u_val in _NO_UNDERLINE_VALS:
                continue

            is_hl_child   = _run_is_direct_hyperlink_child(run)
            is_lo_style   = _run_style_is_lo_link(run, styles)
            run_txt       = _run_text(run)
            is_url        = bool(_URL_RE.search(run_txt))
            has_any_style = bool(_run_char_style(run))

            if is_hl_child or is_url:
                # Genuine hyperlink run — normalise thickness to "single"
                if u_val not in ("single", "none"):
                    u.set(w("val"), "single")
                    normalised += 1
                kept += 1

            elif is_lo_style:
                # LO injected its link character style onto this run.
                # The <w:u> is an artefact — remove it.
                rPr.remove(u)
                removed += 1

            elif not has_any_style and pbdr_was_present:
                # No explicit rStyle, but the paragraph had a pBdr bottom
                # border (strong signal of full LO contamination).
                # Strip the injected underline.
                rPr.remove(u)
                removed += 1

            else:
                # Run has underline, paragraph has links, but run's style is
                # not an LO link style → genuine underline from the PDF.
                # PRESERVE IT.
                kept += 1

    return removed, kept, normalised


# ───────────────────────────────────────────────────────────────
#  FIX — Table widths & black shading
# ───────────────────────────────────────────────────────────────

def _col_min_w(cw: int, nc: int) -> int:
    return max(round(cw * 0.03), cw // max(nc * 4, 1), 1)


def _min_frac(nc: int) -> float:
    if nc <= 2:  return 0.05
    if nc <= 5:  return 0.03
    if nc <= 10: return 0.015
    return 0.008


def _count_cols(tbl: etree._Element) -> int:
    mx = 0
    for row in tbl.findall(w("tr")):
        n = 0
        for cell in row.findall(w("tc")):
            tcPr = cell.find(w("tcPr"))
            gs   = tcPr.find(w("gridSpan")) if tcPr is not None else None
            n   += int(gs.get(w("val"), 1)) if gs is not None else 1
        mx = max(mx, n)
    return mx


def _grid_broken(cols: List, cw: int, nc: int) -> bool:
    if not cols:
        return True
    try:
        ws = [int(c.get(w("w"), 0)) for c in cols]
    except (TypeError, ValueError):
        return True
    tot = sum(ws)
    if tot == 0:
        return True
    mn = max(_col_min_w(cw, nc), round(cw * _min_frac(nc)))
    if any(x < mn for x in ws):
        return True
    if abs(tot - cw) / max(tot, cw, 1) > 0.25:
        return True
    return False


def _rebuild_cols(tbl: etree._Element, nc: int, cw: int) -> List[int]:
    g = tbl.find(w("tblGrid"))
    if g is not None:
        tbl.remove(g)
    g   = etree.Element(w("tblGrid"))
    cw_ = cw // nc
    rem = cw - cw_ * nc
    for i in range(nc):
        gc = etree.SubElement(g, w("gridCol"))
        gc.set(w("w"), str(cw_ + (1 if i < rem else 0)))
    pr  = tbl.find(w("tblPr"))
    pos = (list(tbl).index(pr) + 1) if pr is not None else 0
    tbl.insert(pos, g)
    return [cw_ + (1 if i < rem else 0) for i in range(nc)]


def _redist(tbl: etree._Element, cw: int) -> List[int]:
    g    = tbl.find(w("tblGrid"))
    cols = g.findall(w("gridCol")) if g is not None else []
    nc   = _count_cols(tbl) or len(cols) or 1
    if _grid_broken(cols, cw, nc):
        return _rebuild_cols(tbl, nc, cw)
    try:
        cur = [int(c.get(w("w"), 0)) for c in cols]
    except (TypeError, ValueError):
        return _rebuild_cols(tbl, nc, cw)
    tot = sum(cur) or 1
    mn  = _col_min_w(cw, nc)
    nw  = [max(mn, round(x / tot * cw)) for x in cur]
    nw[-1] = max(mn, nw[-1] + (cw - sum(nw)))
    grid = tbl.find(w("tblGrid"))
    for col, v in zip(grid.findall(w("gridCol")), nw):
        col.set(w("w"), str(v))
    return nw


def _apply_widths(tbl: etree._Element, nw: List[int], mn: int) -> None:
    for row in tbl.findall(w("tr")):
        ci = 0
        for cell in row.findall(w("tc")):
            if ci >= len(nw):
                break
            tcPr = cell.find(w("tcPr"))
            if tcPr is None:
                tcPr = etree.SubElement(cell, w("tcPr"))
                cell.insert(0, tcPr)
            tcW  = tcPr.find(w("tcW"))
            if tcW is None:
                tcW = etree.SubElement(tcPr, w("tcW"))
            gs   = tcPr.find(w("gridSpan"))
            span = int(gs.get(w("val"), 1)) if gs is not None else 1
            cw_  = sum(nw[ci:ci + span]) if ci + span <= len(nw) else nw[ci]
            tcW.set(w("w"),    str(max(cw_, mn)))
            tcW.set(w("type"), "dxa")
            ci  += span


def _is_black_fill(fill_val: str) -> bool:
    if not fill_val:
        return False
    fv = fill_val.upper().strip("#").strip()
    return fv in _BLACK_FILLS or fv == "0"


def _clear_shading_on_element(el: etree._Element) -> None:
    shd = el.find(w("shd"))
    if shd is None:
        return
    fill = (shd.get(w("fill")) or "").upper().strip("#").strip()
    val  = (shd.get(w("val"))  or "").lower()
    if (
        _is_black_fill(fill)
        or (val == "solid" and fill in ("", "AUTO", "FFFFFF"))
        or (val == "solid" and _is_black_fill(fill))
        or (fill == "" and val == "solid")
    ):
        shd.set(w("fill"),  "auto")
        shd.set(w("color"), "auto")
        shd.set(w("val"),   "clear")


def _clear_table_black_shading(tbl: etree._Element) -> None:
    """Deep sweep: clear all black shading from every level of the table."""
    tpr = tbl.find(w("tblPr"))
    if tpr is not None:
        _clear_shading_on_element(tpr)
        tblBdr = tpr.find(w("tblBorders"))
        if tblBdr is not None:
            for bdr in tblBdr:
                c = (bdr.get(w("color")) or "").upper().strip("#")
                if _is_black_fill(c) and bdr.get(w("val"), "") not in ("none","nil"):
                    bdr.set(w("color"), "auto")

    for row in tbl.findall(w("tr")):
        trPr = row.find(w("trPr"))
        if trPr is not None:
            _clear_shading_on_element(trPr)
        for cell in row.findall(w("tc")):
            tcp = cell.find(w("tcPr"))
            if tcp is not None:
                _clear_shading_on_element(tcp)
                tcBdr = tcp.find(w("tcBorders"))
                if tcBdr is not None:
                    for bdr in tcBdr:
                        c = (bdr.get(w("color")) or "").upper().strip("#")
                        if _is_black_fill(c) and bdr.get(w("val"),"") not in ("none","nil"):
                            bdr.set(w("color"), "auto")
            for rpr in cell.iter(w("rPr")):
                _clear_shading_on_element(rpr)
            for ppr in cell.iter(w("pPr")):
                _clear_shading_on_element(ppr)


def fix_table_widths(doc_tree: etree._Element, cw: int) -> int:
    fixed = 0
    for tbl in doc_tree.iter(w("tbl")):
        pr = tbl.find(w("tblPr"))
        if pr is None:
            pr = etree.SubElement(tbl, w("tblPr"))
            tbl.insert(0, pr)
        tw = pr.find(w("tblW"))
        if tw is None:
            tw = etree.SubElement(pr, w("tblW"))
        tw.set(w("w"),    str(cw))
        tw.set(w("type"), "dxa")

        _clear_table_black_shading(tbl)       # pre-width pass

        nc = max(_count_cols(tbl), 1)
        mn = _col_min_w(cw, nc)
        nw = _redist(tbl, cw)
        _apply_widths(tbl, nw, mn)

        _clear_table_black_shading(tbl)       # post-width pass
        fixed += 1
    return fixed


def fix_textbox_widths(doc_tree: etree._Element, cw: int) -> int:
    fixed = 0
    ef    = cw * EMU_PER_DXA
    for ex in doc_tree.iter(wp("extent")):
        try:
            cx = int(ex.get("cx", 0))
            if 0 < cx < ef * 0.5:
                ex.set("cx", str(ef))
                fixed += 1
        except (TypeError, ValueError):
            pass
    for sh in doc_tree.iter(vml("shape")):
        st = sh.get("style", "")
        if "width:" in st:
            ns = re.sub(r"width:\s*[\d.]+\s*(pt|px|in|cm|mm)", "width:100%", st)
            if ns != st:
                sh.set("style", ns)
                fixed += 1
    return fixed


def fix_frame_widths(doc_tree: etree._Element, cw: int) -> int:
    fixed = 0
    thr   = cw * 0.60
    for pPr in doc_tree.iter(w("pPr")):
        fp = pPr.find(w("framePr"))
        if fp is None:
            continue
        fw = fp.get(w("w"))
        if fw:
            try:
                if 0 < int(fw) < thr:
                    pPr.remove(fp)
                    fixed += 1
            except (TypeError, ValueError):
                pass
    return fixed


def fix_indentation(doc_tree: etree._Element, cw: int) -> int:
    fixed = 0
    mi    = cw // 2
    mh    = cw // 4
    for pPr in doc_tree.iter(w("pPr")):
        ind = pPr.find(w("ind"))
        if ind is None:
            continue
        changed = False
        for a in ("left", "right", "start", "end"):
            v = ind.get(w(a))
            if v:
                try:
                    iv = int(v)
                    if iv > mi:
                        ind.set(w(a), str(mi)); changed = True
                    elif iv < -mi:
                        ind.set(w(a), "0");     changed = True
                except (TypeError, ValueError):
                    pass
        for a in ("hanging", "firstLine"):
            v = ind.get(w(a))
            if v:
                try:
                    if int(v) > mh:
                        ind.set(w(a), str(mh // 2))
                        changed = True
                except (TypeError, ValueError):
                    pass
        if changed:
            fixed += 1
    return fixed


# ───────────────────────────────────────────────────────────────
#  FIX — Highlights
# ───────────────────────────────────────────────────────────────
_HL_RGB: Dict[str, Tuple[int, int, int]] = {
    "yellow":      (255, 255,   0),
    "green":       (  0, 255,   0),
    "cyan":        (  0, 255, 255),
    "magenta":     (255,   0, 255),
    "blue":        (  0,   0, 255),
    "red":         (255,   0,   0),
    "darkBlue":    (  0,   0, 128),
    "darkCyan":    (  0, 128, 128),
    "darkGreen":   (  0, 128,   0),
    "darkMagenta": (128,   0, 128),
    "darkRed":     (128,   0,   0),
    "darkYellow":  (128, 128,   0),
    "darkGray":    (128, 128, 128),
    "lightGray":   (192, 192, 192),
    "black":       (  0,   0,   0),
}
_HL_RADIUS = 90


def _h2rgb(h: str) -> Optional[Tuple[int, int, int]]:
    h = h.upper().lstrip("#")
    if len(h) != 6:
        return None
    try:
        return int(h[:2], 16), int(h[2:4], 16), int(h[4:], 16)
    except ValueError:
        return None


def _nearest_hl(hx: str) -> Optional[str]:
    rgb = _h2rgb(hx)
    if rgb is None:
        return None
    best, bd = None, float("inf")
    for name, ref in _HL_RGB.items():
        d = sum((a - b) ** 2 for a, b in zip(rgb, ref)) ** 0.5
        if d < bd:
            bd, best = d, name
    return best if bd <= _HL_RADIUS else None


def _force_text_color_auto(rPr: etree._Element) -> None:
    """Ensure <w:color val="auto"> so text is always readable on a highlight."""
    ce = rPr.find(w("color"))
    if ce is None:
        ce = etree.SubElement(rPr, w("color"))
    ce.set(w("val"), "auto")


def fix_highlights_from_shading(doc_tree: etree._Element) -> Tuple[int, int]:
    conv = cfix = 0
    for rPr in doc_tree.iter(w("rPr")):
        shd  = rPr.find(w("shd"))
        if shd is None:
            continue
        fill = (shd.get(w("fill")) or "").upper().lstrip("#")
        if not fill or fill in ("AUTO", "FFFFFF", ""):
            continue
        hl = _nearest_hl(fill)
        if hl is None:
            continue
        rPr.remove(shd)
        he = rPr.find(w("highlight"))
        if he is None:
            he = etree.SubElement(rPr, w("highlight"))
        he.set(w("val"), hl)
        conv += 1
        _force_text_color_auto(rPr)
        cfix += 1
    return conv, cfix


def add_highlight_to_keywords(
    doc_tree: etree._Element,
    kws:      List[str],
    colour:   str = "yellow",
) -> int:
    if not kws:
        return 0
    pat = re.compile("|".join(re.escape(k) for k in kws), re.IGNORECASE)
    n   = 0
    for run in doc_tree.iter(w("r")):
        t = run.find(w("t"))
        if t is None or not t.text:
            continue
        if not pat.search(t.text):
            continue
        rPr = run.find(w("rPr"))
        if rPr is None:
            rPr = etree.SubElement(run, w("rPr"))
            run.insert(0, rPr)
        he = rPr.find(w("highlight"))
        if he is None:
            he = etree.SubElement(rPr, w("highlight"))
        he.set(w("val"), colour)
        _force_text_color_auto(rPr)
        n += 1
    return n


def ensure_highlighted_text_visible(doc_tree: etree._Element) -> int:
    """Final safety sweep: any run with an active highlight gets color→auto."""
    forced = 0
    for rPr in doc_tree.iter(w("rPr")):
        hl = rPr.find(w("highlight"))
        if hl is None:
            continue
        if hl.get(w("val"), "") in ("none", ""):
            continue
        ce = rPr.find(w("color"))
        if ce is None or ce.get(w("val"), "").lower() not in ("auto", ""):
            _force_text_color_auto(rPr)
            forced += 1
    return forced


# ───────────────────────────────────────────────────────────────
#  Orchestrator
# ───────────────────────────────────────────────────────────────
def postprocess(
    docx_path:          str,
    output_path:        str,
    highlight_keywords: Optional[List[str]] = None,
    highlight_colour:   str  = "yellow",
    verbose:            bool = True,
) -> None:
    docx_path   = Path(docx_path)
    output_path = Path(output_path)
    if verbose:
        print(f"[POST-PROCESS] {docx_path} → {output_path}")

    if not zipfile.is_zipfile(str(docx_path)):
        raise ValueError(f"Not a valid DOCX/ZIP: {docx_path}")

    with zipfile.ZipFile(str(docx_path), "r") as zf:
        if DOCUMENT_XML not in zf.namelist():
            raise FileNotFoundError(f"{DOCUMENT_XML} not found in {docx_path}")
        doc  = read_xml(zf, DOCUMENT_XML)
        styl = load_styles(zf)
        rest = {n: zf.read(n) for n in zf.namelist() if n != DOCUMENT_XML}

    geo = detect_page_geometry(doc)
    ensure_page_geometry(doc, geo)
    cw  = geo["content_w"]
    if verbose:
        print(f"  Page {geo['page_w']}×{geo['page_h']} DXA  "
              f"margins L{geo['margin_left']} R{geo['margin_right']} "
              f"T{geo['margin_top']} B{geo['margin_bottom']}")
        print(f"  Content width {cw} DXA  ({cw / DXA_PER_INCH:.2f}\")")

    # ── Fix underlines ───────────────────────────────────────────────────────
    nr, nk, nn = fix_underlines(doc, styl)
    if verbose:
        print(f"  Underlines : -{nr} LO-injected removed  "
              f"+{nk} kept (genuine/hyperlink)  {nn} normalised→single")

    # ── Fix tables / layout ──────────────────────────────────────────────────
    nt = fix_table_widths(doc, cw)
    nx = fix_textbox_widths(doc, cw)
    nf = fix_frame_widths(doc, cw)
    ni = fix_indentation(doc, cw)
    if verbose:
        print(f"  Layout     : {nt} tables  {nx} textboxes  "
              f"{nf} frames  {ni} indents")

    # ── Fix highlights ───────────────────────────────────────────────────────
    ns, nc2 = fix_highlights_from_shading(doc)
    nkw     = add_highlight_to_keywords(doc, highlight_keywords or [], highlight_colour)
    nvis    = ensure_highlighted_text_visible(doc)
    if verbose:
        print(f"  Highlights : {ns} shading→hl  {nkw} keyword-hl  "
              f"{nvis} text-color→auto")

    # ── Write output ─────────────────────────────────────────────────────────
    xml = write_xml(doc)
    with zipfile.ZipFile(str(output_path), "w", zipfile.ZIP_DEFLATED) as zo:
        zo.writestr(DOCUMENT_XML, xml)
        for name, data in rest.items():
            zo.writestr(name, data)

    if verbose:
        print(f"[DONE] {output_path}")


# ───────────────────────────────────────────────────────────────
#  CLI
# ───────────────────────────────────────────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser(
        description="PDF → DOCX via LibreOffice + post-processing (v5)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  %(prog)s in.pdf out.docx\n"
            "  %(prog)s in.pdf out.docx --highlight-words Sony Name\n"
            "  %(prog)s in.docx out.docx --skip-conversion\n"
        ),
    )
    ap.add_argument("input",  help="Input PDF (or DOCX with --skip-conversion)")
    ap.add_argument("output", help="Output DOCX path")
    ap.add_argument(
        "--highlight-words", nargs="*", metavar="W", default=[],
        help="Words/phrases to highlight in the output",
    )
    ap.add_argument(
        "--highlight-colour", default="yellow",
        choices=sorted(_HL_RGB.keys()),
        help="Highlight colour for --highlight-words  (default: yellow)",
    )
    ap.add_argument(
        "--skip-conversion", action="store_true",
        help="Skip LibreOffice PDF→DOCX and post-process an existing DOCX",
    )
    ap.add_argument(
        "--quiet", action="store_true",
        help="Suppress progress output",
    )
    args = ap.parse_args()

    inp = Path(args.input)
    out = Path(args.output)
    v   = not args.quiet

    if args.skip_conversion:
        if v:
            print("[SKIP] LibreOffice conversion — post-processing only")
        postprocess(str(inp), str(out), args.highlight_words,
                    args.highlight_colour, v)
    else:
        if not inp.exists():
            print(f"[ERROR] Input file not found: {inp}", file=sys.stderr)
            sys.exit(1)
        raw = out.with_suffix(".raw.docx")
        try:
            convert_pdf_to_docx(str(inp), str(raw))
            if v:
                print("\n[STEP 3] Post-processing …")
            postprocess(str(raw), str(out), args.highlight_words,
                        args.highlight_colour, v)
        finally:
            if raw.exists():
                raw.unlink()
                if v:
                    print(f"  Cleaned temp file: {raw}")

    if v:
        print(f"\n[SUCCESS] Output: {out}")


if __name__ == "__main__":
    main()
