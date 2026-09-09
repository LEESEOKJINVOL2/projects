"""Build the Harald_260820 PPT deck (newer retest pull, 2026-08-20) -- same
slide layout as build_ppt_harald_distribution.py: Cal-style interleaved
table with USL/LSL as two separate red-text rows, plus a "Bench"/"ATE" row
label (no numeric shared span -- see that file's docstring for why: a
PATNAME-grouped page's up-to-4 items are often unrelated measurements, so a
shared axis was tried and reverted for the original Harald report; same
grouping mechanism here, same reasoning applies).

Data comes from make_harald_260820_distribution.py (Cal's own read_bench/
read_ate against the flat Bench_40pcs/Harald_260820 + single ATE file
layout this retest pull uses, unlike the original pull's vendor-nested one).
"""
from __future__ import annotations

import datetime
import statistics
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

_THIS_DIR = Path(__file__).resolve().parent
_BASE_DIR = _THIS_DIR.parent
sys.path.insert(0, str(_BASE_DIR / "Cal"))
sys.path.insert(0, str(_THIS_DIR))

import make_cal_distribution as cal  # _find_layout/add_item_slide pieces reused, not modified
import make_harald_260820_distribution as harald

PATHS = harald.PATHS
TEMPLATE_PATH = PATHS.ppt_template

SLIDE_WIDTH_IN = 13.338542213473316
LEFT_MARGIN = 0.05
RIGHT_MARGIN = 0.05
CELL_GAP_H = 0.03
LABEL_WIDTH = 0.9

TABLE_TOP = 1.05
HEADER_ROW_H = 0.55
TABLE_ROW_H = 0.2
TABLE_PLOT_GAP = 0.35
IMAGE_GAP = 0.15
CONTENT_AREA_TOP = 1.5908508311461067
CONTENT_AREA_BOTTOM = CONTENT_AREA_TOP + 5.240927384076991

ATE_BORDER_COLOR = RGBColor(0xC0, 0x00, 0x00)
ATE_TEXT_COLOR = RGBColor(0xC0, 0x00, 0x00)
USL_COLOR = RGBColor(0xB0, 0x00, 0x00)

REPORT_DATE = datetime.date.today().isoformat()
TITLE_TEXT = f"{PATHS.title_prefix} {PATHS.project} {PATHS.title_suffix}"
SUBTITLE_LINES = ["RF LAB1 Team", REPORT_DATE]


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


def _set_cell(cell, text, size=7, bold=True, color=None):
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


def _add_label(slide, left, top, width, height, text, size=9, bold=True):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    return box


def _add_ruler(slide, left, top, width, height, vmin=None, vmax=None, row_label=None):
    """Row label only (no numeric span) -- see module docstring / build_ppt_
    harald_distribution.py's own _add_ruler for why: this project's
    PATNAME-section pages mix unrelated measurements far more often than
    FWBT_TRIMTEST's band-grouped pages, so a shared numeric axis would
    misrepresent most charts."""
    if vmin is not None and vmax is not None:
        tick_h = 0.2
        _add_label(slide, left, top, width, tick_h, f"{vmax:.0f}", size=9, bold=True)
        _add_label(slide, left, top + height - tick_h, width, tick_h, f"{vmin:.0f}", size=9, bold=True)
    if row_label:
        label_h = 0.4
        label_top = top + (height - label_h) / 2
        _add_label(slide, left, label_top, width, label_h, row_label, size=16, bold=True)


def _add_no_ate_box(slide, left, top, width, height):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    box.fill.background()
    box.line.color.rgb = ATE_BORDER_COLOR
    box.line.width = Pt(1)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "No ATE match"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = ATE_TEXT_COLOR


def _stats(values):
    if not values:
        return None
    mean = statistics.mean(values)
    std = statistics.stdev(values) if len(values) > 1 else 0.0
    return mean, min(values), max(values), std


def _fmt_stat(values):
    s = _stats(values)
    if s is None:
        return "", ""
    mean, vmin, vmax, std = s
    return f"{mean:.4g}±{std:.4g}", f"{vmin:.4g} / {vmax:.4g}"


def _fmt_limit(v):
    return f"{v:g}" if v is not None else ""


def _add_interleaved_table(slide, left, top, table_width, jobs):
    n_items = len(jobs)
    n_data_cols = 2 * n_items
    n_cols = 1 + n_data_cols
    n_rows = 6  # item header + Bench/ATE subheader + Mean+-Std + Min/Max + USL + LSL

    grid_width = table_width - LABEL_WIDTH
    cell_width = (grid_width - (n_data_cols - 1) * CELL_GAP_H) / n_data_cols
    col_widths = [LABEL_WIDTH] + [cell_width] * n_data_cols
    row_heights = [HEADER_ROW_H] + [TABLE_ROW_H] * (n_rows - 1)
    total_width = sum(col_widths)
    total_height = sum(row_heights)

    gframe = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top),
                                     Inches(total_width), Inches(total_height))
    table = gframe.table
    table.first_row = False
    for i, w in enumerate(col_widths):
        table.columns[i].width = Inches(w)
    for r, h in enumerate(row_heights):
        table.rows[r].height = Inches(h)

    _set_cell(table.cell(0, 0), "Item", size=7)
    _set_cell(table.cell(1, 0), "", size=7)
    _set_cell(table.cell(2, 0), "Mean±Std", size=7)
    _set_cell(table.cell(3, 0), "Min/Max", size=7)
    _set_cell(table.cell(4, 0), "USL", size=7, color=USL_COLOR)
    _set_cell(table.cell(5, 0), "LSL", size=7, color=USL_COLOR)

    for ci, job in enumerate(jobs):
        c1 = 1 + ci * 2
        merged = table.cell(0, c1)
        merged.merge(table.cell(0, c1 + 1))
        _set_cell(merged, cal.wrap_label(job["cleaned_name"]), size=6)
        _set_cell(table.cell(1, c1), "Bench", size=7)
        _set_cell(table.cell(1, c1 + 1), "ATE", size=7)

        b_mean, b_minmax = _fmt_stat(job["bench_values"])
        a_mean, a_minmax = _fmt_stat(job["ate_values"] or [])
        _set_cell(table.cell(2, c1), b_mean, size=6, bold=True)
        _set_cell(table.cell(2, c1 + 1), a_mean, size=6, bold=True)
        _set_cell(table.cell(3, c1), b_minmax, size=6, bold=True)
        _set_cell(table.cell(3, c1 + 1), a_minmax, size=6, bold=True)

        llim, ulim = cal.effective_spec(job)
        usl_cell = table.cell(4, c1)
        usl_cell.merge(table.cell(4, c1 + 1))
        _set_cell(usl_cell, _fmt_limit(ulim), size=6, bold=True, color=USL_COLOR)
        lsl_cell = table.cell(5, c1)
        lsl_cell.merge(table.cell(5, c1 + 1))
        _set_cell(lsl_cell, _fmt_limit(llim), size=6, bold=True, color=USL_COLOR)

    return total_height


def _add_image_stretched(slide, path: Path, left, top, width, height):
    slide.shapes.add_picture(str(path), Inches(left), Inches(top), Inches(width), Inches(height))


def add_item_slide(prs, layout_idx, title_text, jobs):
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    ph_title = _get_ph(slide, 1) or _get_ph(slide, 0)
    if ph_title:
        ph_title.text = title_text
    _remove_placeholder(slide, 12)

    usable_left = LEFT_MARGIN
    usable_right = SLIDE_WIDTH_IN - RIGHT_MARGIN
    n_cols = len(jobs)
    cell_width = (usable_right - usable_left - LABEL_WIDTH - (n_cols - 1) * CELL_GAP_H) / n_cols

    table_h = _add_interleaved_table(slide, usable_left, TABLE_TOP, usable_right - usable_left, jobs)

    plots_top = TABLE_TOP + table_h + TABLE_PLOT_GAP
    plots_bottom = CONTENT_AREA_BOTTOM
    image_h = (plots_bottom - plots_top - IMAGE_GAP) / 2
    bench_top = plots_top
    ate_top = bench_top + image_h + IMAGE_GAP

    _add_ruler(slide, usable_left, bench_top, LABEL_WIDTH, image_h, row_label="Bench")
    _add_ruler(slide, usable_left, ate_top, LABEL_WIDTH, image_h, row_label="ATE")

    for ci, job in enumerate(jobs):
        cell_left = usable_left + LABEL_WIDTH + ci * (cell_width + CELL_GAP_H)
        if job["bench_png"].exists():
            _add_image_stretched(slide, job["bench_png"], cell_left, bench_top, cell_width, image_h)
        if job["ate_values"] is not None and job["ate_png"].exists():
            _add_image_stretched(slide, job["ate_png"], cell_left, ate_top, cell_width, image_h)
        else:
            _add_no_ate_box(slide, cell_left, ate_top, cell_width, image_h)

    return slide


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


def _reorder_slides(prs, ordered_slides) -> None:
    xml_slides = prs.slides._sldIdLst
    id_by_rid = {e.get(qn("r:id")): e for e in list(xml_slides)}
    part_to_rid = {prs.part.related_part(rid): rid for rid in id_by_rid}
    for e in list(xml_slides):
        xml_slides.remove(e)
    for slide in ordered_slides:
        rid = part_to_rid[slide.part]
        xml_slides.append(id_by_rid[rid])


def build_all_jobs():
    bench_items = cal.read_bench(PATHS.bench_dir)
    ate_items = cal.read_ate_auto(harald.ATE_PATH)  # ATE_PATH may be one file or a folder of per-config files
    jobs, n_no_ate = cal.build_jobs(bench_items, ate_items)
    for job in jobs:
        job["cleaned_name"] = harald.clean(job["name"])
    return jobs, n_no_ate, len(bench_items)


def main() -> None:
    prs = Presentation(TEMPLATE_PATH)
    content_layout_idx = _find_layout(prs, "1_", 5)

    print("Reading Bench_Harald_260820 CSVs + ATE log...")
    jobs, n_no_ate, n_bench_items = build_all_jobs()
    print(f"=== items with >=1 Bench limit: {len(jobs)}/{n_bench_items} "
          f"(ATE match: {len(jobs) - n_no_ate}, no ATE match: {n_no_ate}) ===")

    pages = cal.build_pages(jobs)
    n_sections = len({p["section"] for p in pages})
    print(f"{n_sections} PATNAME section(s) with a qualifying item")

    bench_dir = harald.bench_png_dir()
    ate_dir = harald.ate_png_dir() if PATHS.extract_mode != "bench" else Path("__no_such_dir__")

    for job in jobs:
        fname = cal.safe_filename(job["cleaned_name"])
        job["bench_png"] = bench_dir / f"{fname}.png"
        job["ate_png"] = ate_dir / f"{fname}.png"

    ordered = []
    for page in pages:
        title = f"Harald_{page['section']}"
        if page["n_pages"] > 1:
            title += f" ({page['page']}/{page['n_pages']})"
        slide = add_item_slide(prs, content_layout_idx, title, page["jobs"])
        ordered.append(slide)

    print(f"{len(pages)} content slide(s) ({cal.COLS_PER_SLIDE} item(s)/slide max)")

    title_slide = add_title_slide(prs)
    ordered = [title_slide, *ordered]
    _reorder_slides(prs, ordered)

    report_dir = PATHS.report_dir
    report_dir.mkdir(parents=True, exist_ok=True)
    out_path = report_dir / f"{REPORT_DATE}_Harald_260820_Distribution.pptx"
    prs.save(out_path)
    print(f"\nSaved: {out_path}")
    print(f"Total slides: {len(prs.slides._sldIdLst)}")


if __name__ == "__main__":
    main()
