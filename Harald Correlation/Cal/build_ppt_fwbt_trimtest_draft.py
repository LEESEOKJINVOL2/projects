"""Draft-only structure check for the FWBT_TRIMTEST_DP6P1P3 report deck.

Purpose (per user request 2026-08-15): produce a layout draft FAST, with NO
images and NO Bench/ATE data -- only slide titles + item-name placeholders,
so the slide composition/grouping/left-right placement can be reviewed and
approved before any plotting work starts. Once approved, a follow-up script
fills the same slide grid with real Bench/ATE plots + stats (same
interleaved-table style as Cal/build_ppt_cal_distribution.py).

Grouping rule (derived from the renamed Test Item names in
FWBT_TRIMTEST_DP6P1P3_renamed.xlsx's "Test Items" sheet, see that file's
naming cleanup): strip the "BT2G_"/"BT5G_" prefix and normalize any
"BAND<n>" token to "BAND" so 2G and 5G variants of the *same* measurement
collapse onto one group key. Checked against all 2737 items:
  - every 2G side has at most 1 item per group (single BAND0)
  - 5G side has 1, 2, 6, or 7 items per group (one per 5G band, BAND1..7)
  - 41 items carry no BT2G_/BT5G_ prefix at all ("common" items, e.g.
    BT_XO_AMP_*, BT_CLKMEAS_*) -- these get their own single-column slide.
So each slide is: label column, then one 2G column (if any), then up to 7
5G columns (BAND1..BAND7, if any) -- exactly the "2G left / 5G right"
placement the user asked for.
"""
from __future__ import annotations

import datetime
import re
from pathlib import Path

import openpyxl
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE_DIR = Path(__file__).resolve().parent.parent
XLSX_PATH = BASE_DIR / "FWBT_TRIMTEST_DP6P1P3_renamed.xlsx"
TEMPLATE_PATH = BASE_DIR / "PPT-Template.pptx"
REPORT_DIR = BASE_DIR / "report" / "BT_TX_TRIMTEST"

REPORT_DATE = datetime.date.today().isoformat()
TITLE_TEXT = "FWBT TRIMTEST Report (DRAFT -- layout check only, no data)"
SUBTITLE_LINES = ["RF LAB1 Team", REPORT_DATE, "Draft: item placement/grouping check -- no plots yet"]

SLIDE_WIDTH_IN = 13.338542213473316
LEFT_MARGIN = 0.3
RIGHT_MARGIN = 0.3
CELL_GAP_H = 0.05
LABEL_WIDTH = 1.3
TABLE_TOP = 1.3
HEADER_ROW_H = 0.35
NAME_ROW_H = 1.4
BAND_ORDER = [None, "BAND1", "BAND2", "BAND3", "BAND4", "BAND5", "BAND6", "BAND7"]


PREFIX_RE = re.compile(r"^BT(2G|5G)(TTR)?_")


def base_key(name: str):
    """Return (side, key, band_label) for one cleaned Test Item name.

    Covers all 4 observed prefix variants -- BT2G_/BT5G_ (580/1853 items)
    and BT2GTTR_/BT5GTTR_ (78/185 items, a distinct LBRX sub-block: CTRIM/
    TIA2C/HD3NOTCH/BEST_TRIM items with their own LBRXMODE/GINDX dimension).
    TTR is kept in the group key (not stripped) so it never collides with a
    plain BT2G_/BT5G_ group of a similar-looking name -- it still lands on
    the 2G/5G side of the slide, just as its own group.
    """
    m = PREFIX_RE.match(name)
    if not m:
        return "common", name, None
    side = m.group(1)
    ttr_tag = m.group(2) or ""
    rest = ttr_tag + ("_" if ttr_tag else "") + name[m.end():]
    band_m = re.search(r"BAND(\d+)", rest)
    band_label = f"BAND{band_m.group(1)}" if band_m else None
    key = re.sub(r"BAND\d+", "BAND", rest)
    return side, key, band_label


def load_groups():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb["Test Items"]
    names = [row[0] for row in ws.iter_rows(min_row=2, values_only=True) if row[0]]

    groups = {}
    for n in names:
        side, key, band_label = base_key(n)
        g = groups.setdefault(key, {"2G": None, "5G": {}, "common": None})
        if side == "2G":
            g["2G"] = n
        elif side == "5G":
            g["5G"][band_label] = n
        else:
            g["common"] = n
    return groups


def _find_layout(prs, prefix, fallback):
    for i, layout in enumerate(prs.slide_layouts):
        if layout.name.startswith(prefix):
            return i
    return fallback


def _get_ph(slide, idx):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None


def _remove_placeholder(slide, idx: int) -> None:
    ph = _get_ph(slide, idx)
    if ph is not None:
        ph._element.getparent().remove(ph._element)


def _set_cell(cell, text, size=8, bold=True, color=None):
    """text may contain "\n" -- each segment becomes its own paragraph
    (a plain "\n" inside a single run does not render as a line break)."""
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Pt(2)
    cell.margin_right = Pt(2)
    cell.margin_top = Pt(1)
    cell.margin_bottom = Pt(1)
    tf = cell.text_frame
    tf.word_wrap = True
    lines = (text if text else " ").split("\n")
    tf.text = lines[0]
    for line in lines[1:]:
        tf.add_paragraph().text = line
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0] if p.runs else p.add_run()
        run.font.size = Pt(size)
        run.font.bold = bold
        if color is not None:
            run.font.color.rgb = color


def add_title_slide(prs):
    layout_idx = _find_layout(prs, "2_", 0)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    ph_title = _get_ph(slide, 0)
    if ph_title:
        ph_title.text = TITLE_TEXT
    ph_subtitle = _get_ph(slide, 1)
    if ph_subtitle:
        tf = ph_subtitle.text_frame
        tf.text = SUBTITLE_LINES[0]
        for line in SUBTITLE_LINES[1:]:
            p = tf.add_paragraph()
            p.text = line
    return slide


def _columns_for_group(key, g):
    """Return ordered list of (header_label, item_name) -- 2G first, then
    5G bands in BAND1..BAND7 order, then a lone "common" column.

    Header label is 2 lines: "2G"/"5G" then "(BANDn)" below -- 2G is
    always (BAND0); 5G shows whichever band number that item actually
    carries. Items with no BAND token at all get just "5G" (no 2nd line).
    """
    cols = []
    if g["2G"]:
        cols.append(("2G\n(BAND0)", g["2G"]))
    if None in g["5G"]:
        cols.append(("5G", g["5G"][None]))
    for band in BAND_ORDER[1:]:
        if band in g["5G"]:
            cols.append((f"5G\n({band})", g["5G"][band]))
    if g["common"]:
        cols.append(("--", g["common"]))
    return cols


def clean_title(key: str) -> str:
    """Drop the literal "BAND" placeholder token from a group key so the
    slide title reads e.g. "A4P_EXPAN" instead of "A4P_BAND_EXPAN" --
    the band number itself now lives in the column headers instead."""
    parts = [p for p in key.split("_") if p != "BAND"]
    return "_".join(parts)


def add_group_slide(prs, layout_idx, title_text, columns):
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    ph_title = _get_ph(slide, 1) or _get_ph(slide, 0)
    if ph_title:
        ph_title.text = title_text
    _remove_placeholder(slide, 12)

    n_cols = len(columns)
    usable_left = LEFT_MARGIN
    usable_right = SLIDE_WIDTH_IN - RIGHT_MARGIN
    grid_width = usable_right - usable_left - LABEL_WIDTH
    cell_width = (grid_width - (n_cols - 1) * CELL_GAP_H) / n_cols

    n_rows = 2
    col_widths = [LABEL_WIDTH] + [cell_width] * n_cols
    row_heights = [HEADER_ROW_H, NAME_ROW_H]
    total_width = sum(col_widths)
    total_height = sum(row_heights)

    gframe = slide.shapes.add_table(n_rows, 1 + n_cols, Inches(usable_left), Inches(TABLE_TOP),
                                     Inches(total_width), Inches(total_height))
    table = gframe.table
    table.first_row = False
    for i, w in enumerate(col_widths):
        table.columns[i].width = Inches(w)
    for r, h in enumerate(row_heights):
        table.rows[r].height = Inches(h)

    _set_cell(table.cell(0, 0), "Band", size=8)
    _set_cell(table.cell(1, 0), "Item / [plot]", size=7)

    for ci, (label, item_name) in enumerate(columns):
        c = 1 + ci
        _set_cell(table.cell(0, c), label, size=8)
        _set_cell(table.cell(1, c), f"{item_name}\n[Bench/ATE plot here]", size=6, bold=False)

    return slide


def _reorder_slides(prs, ordered_slides) -> None:
    from pptx.oxml.ns import qn
    xml_slides = prs.slides._sldIdLst
    id_by_rid = {e.get(qn("r:id")): e for e in list(xml_slides)}
    part_to_rid = {prs.part.related_part(rid): rid for rid in id_by_rid}
    for e in list(xml_slides):
        xml_slides.remove(e)
    for slide in ordered_slides:
        rid = part_to_rid[slide.part]
        xml_slides.append(id_by_rid[rid])


def main() -> None:
    prs = Presentation(TEMPLATE_PATH)
    content_layout_idx = _find_layout(prs, "1_", 5)

    print("Loading Test Item names + grouping by 2G/5G-band...")
    groups = load_groups()
    print(f"{len(groups)} group(s) from {sum(1 for _ in groups)} keys")

    ordered = []
    title_slide = add_title_slide(prs)
    ordered.append(title_slide)

    n_both = n_2g_only = n_5g_only = n_common = 0
    for key in sorted(groups):
        g = groups[key]
        columns = _columns_for_group(key, g)
        if g["common"]:
            n_common += 1
            title = clean_title(g["common"])
        elif g["2G"] and g["5G"]:
            n_both += 1
            title = clean_title(key)
        elif g["2G"]:
            n_2g_only += 1
            title = clean_title(key)
        else:
            n_5g_only += 1
            title = clean_title(key)
        slide = add_group_slide(prs, content_layout_idx, title, columns)
        ordered.append(slide)

    print(f"  both 2G&5G: {n_both}, 2G-only: {n_2g_only}, 5G-only: {n_5g_only}, common: {n_common}")
    print(f"Total content slides: {len(ordered) - 1}")

    _reorder_slides(prs, ordered)

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = REPORT_DIR / f"{REPORT_DATE}_FWBT_TRIMTEST_Draft_v2.pptx"
    prs.save(out_path)
    print(f"\nSaved: {out_path}")
    print(f"Total slides: {len(prs.slides._sldIdLst)}")


if __name__ == "__main__":
    main()
