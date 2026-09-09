"""Build the BT_RX PPT deck: RSSI / Sensitivity(PER) / Sensitivity(BER)
sections, one slide per packet_type per section.

Layout (deliberately simpler than BT_TX's Power/DEVM/ACP builder -- this
project has no pa_supply/pa_gain dimension and every packet_type here has
at most 5 matched (band, channel, power) columns, so there's no need for
BT_TX's band-2.4G-vs-5G slide-splitting logic):
  * ONE table under the title: each matched (band, channel, power) combo
    gets a merged 2-column header ("band{N}\\nCH{ch:03d}\\n{pwr}dBm") over
    a Bench/ATE sub-header pair, same interleaved-table convention as every
    other project's Bench/ATE correlation slide (see md/CLAUDE.md's
    "BT_TX: Bench/ATE correlation" section).
  * Two stacked image rows below the table (Bench on top, ATE below), one
    column per combo -- reusing the PNGs make_bt_rx_*.py already wrote.
  * NO left-side ruler. Every other project's ruler assumes one shared
    Y-range across a slide's columns; here each combo's chart auto-ranges
    to its OWN data independently (see bt_rx_ate_lookup.draw_png), so
    there's no single shared min/max a ruler could meaningfully show. Each
    PNG already carries its own axis at a large, legible font -- same
    reasoning UWB uses to hide its ruler for small (<=2) column counts,
    just extended here since BT_RX's columns never share an axis at all.

Only combos with >=1 matched Bench DUT are included (2026-08-14 user
decision -- an ATE test point with no Bench data nearby, e.g. the -17dBm
"M17" items, is excluded entirely rather than shown as a per-column "No
Bench match" hole).
"""
from __future__ import annotations

import datetime
from collections import defaultdict
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

import bt_rx_ate_lookup as rx
import make_bt_rx_rssi_revision1 as rssi_mod
import make_bt_rx_sensitivity_per_revision1 as per_mod
import make_bt_rx_sensitivity_ber_revision1 as ber_mod

PATHS = rx.PATHS
TEMPLATE_PATH = PATHS.ppt_template

# 2026-08-14: these all now come from bt_rx_ate_lookup.py (single source of
# truth, shared with figure_size_for() -- see that function's docstring for
# why keeping two independent copies of this geometry was a real bug: this
# builder's OWN layout is what figure_size_for() must match, so a local copy
# here could silently drift from what the chart scripts actually drew).
SLIDE_WIDTH_IN = rx._SLIDE_WIDTH_IN
LEFT_MARGIN = rx._LEFT_MARGIN
RIGHT_MARGIN = rx._RIGHT_MARGIN
CELL_GAP_H = rx._CELL_GAP_H
# Table row-label column ("Condition"/"Mean±Std"/"Min/Max") AND the image
# grid below both start this far right of usable_left, so the table's data
# columns line up exactly with their chart image below (same alignment
# purpose as every other project's RULER_WIDTH, just holding a row-label
# column instead of a min/max axis ruler -- see module docstring for why
# there's no ruler here).
LABEL_WIDTH = rx._LABEL_WIDTH

TABLE_TOP = rx._TABLE_TOP
TABLE_ROW_H = rx._TABLE_ROW_H
# 3-line header cell ("band{N}\nCH{ch}\n{pwr}dBm" at 7pt) needs more room
# than one plain data row -- PowerPoint silently grows a row past its
# declared height to fit wrapped text (real bug hit here 2026-08-14: with
# HEADER_ROW_H == TABLE_ROW_H, the header row rendered taller than
# table_h accounted for, so the image grid below was positioned too high
# and overlapped the table's own Mean±Std/Min-Max rows). Sized generously
# (3 lines * ~1.2*7pt + margins) rather than exactly, since this table's
# total height is a small fraction of the slide either way.
HEADER_ROW_H = rx._HEADER_ROW_H
TABLE_PLOT_GAP = rx._TABLE_PLOT_GAP
IMAGE_GAP = rx._IMAGE_GAP
CONTENT_AREA_TOP = rx._CONTENT_AREA_TOP
CONTENT_AREA_BOTTOM = rx._CONTENT_AREA_BOTTOM

REPORT_DATE = datetime.date.today().isoformat()
TITLE_TEXT = f"{PATHS.title_prefix} {PATHS.project} {PATHS.title_suffix}"
SUBTITLE_LINES = ["RF LAB1 Team", REPORT_DATE]
SECTION_FONT_SIZE = Pt(40)

SECTIONS = [
    ("RSSI", rssi_mod, "rssi_revision1", "RSSI (dBm)"),
    ("Sensitivity_PER", per_mod, "sensitivity_per_revision1", "PER"),
    ("Sensitivity_BER", ber_mod, "sensitivity_ber_revision1", "BER"),
]


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
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Pt(2)
    cell.margin_right = Pt(2)
    cell.margin_top = Pt(1)
    cell.margin_bottom = Pt(1)
    tf = cell.text_frame
    tf.word_wrap = True
    tf.text = text if text else " "
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0] if p.runs else p.add_run()
        run.font.size = Pt(size)
        run.font.bold = bold
        if color is not None:
            run.font.color.rgb = color


def _stats(values):
    import statistics
    return (statistics.mean(values), min(values), max(values),
            statistics.stdev(values) if len(values) > 1 else 0.0)


def _fmt_stat(values, fmt):
    if not values:
        return "", ""
    mean, vmin, vmax, std = _stats(values)
    return f"{mean:{fmt}}±{std:{fmt}}", f"{vmin:{fmt}} / {vmax:{fmt}}"


SPEC_COLOR = RGBColor(0xB0, 0x00, 0x00)  # dark red -- same USL_COLOR as BT_TX's build_ppt_pwr_evm_revision2.py


POWER_REF_COLOR = RGBColor(0xFF, 0x00, 0x00)  # same red as draw_png's power_ref axhline


def _add_interleaved_table(slide, left, top, table_width, header_label, cond_labels, jobs, value_fmt,
                            show_power_level=False):
    """One combined table, Bench/ATE columns alternating per condition --
    same convention as every other project's Bench/ATE slide (see
    md/CLAUDE.md), reimplemented self-contained here (BT_RX doesn't import
    chart-building code from other projects).

    2026-08-14: USL/LSL rows added, one merged Bench+ATE cell per condition
    (same "spec is one physical value regardless of which side measured
    it" reasoning as BT_TX's colvar_rows) -- sourced from each job's own
    ATE per-item Lower/Upper Limit (see bt_rx_ate_lookup.parse_ate_items),
    NOT a Bench spec-sheet lookup (this project has none). A row is added
    only if at least one column in this slide actually has that side of
    the spec; columns without it show a blank cell rather than dropping
    the row/shifting others out of alignment.

    show_power_level (2026-08-15, RSSI only): adds an UNMERGED "Power
    Level" row -- Bench's own matched sweep power under the Bench column,
    ATE's own nominal target power under the ATE column (these can differ
    slightly, see bt_rx_ate_lookup.condition_title's docstring), matching
    the red dashed reference line drawn on each chart (see draw_png's
    power_ref param). Unlike USL/LSL this is never merged: it's two
    genuinely different numbers, one per side, not one shared spec value."""
    n_cond = len(cond_labels)
    n_data_cols = 2 * n_cond
    n_cols = 1 + n_data_cols

    any_usl = any(job.get("spec") and job["spec"][1] is not None for job in jobs)
    any_lsl = any(job.get("spec") and job["spec"][0] is not None for job in jobs)
    spec_row_labels = ([("USL",)] if any_usl else []) + ([("LSL",)] if any_lsl else [])
    n_spec_rows = len(spec_row_labels)
    n_power_rows = 1 if show_power_level else 0
    n_rows = 4 + n_spec_rows + n_power_rows  # cond header + Bench/ATE subheader + Mean+-Std + Min/Max [+ USL] [+ LSL] [+ Power Level]

    grid_width = table_width - LABEL_WIDTH
    cell_width = (grid_width - (n_data_cols - 1) * CELL_GAP_H) / n_data_cols
    col_widths = [LABEL_WIDTH] + [cell_width] * n_data_cols
    total_width = sum(col_widths)
    row_heights = [HEADER_ROW_H] + [TABLE_ROW_H] * (n_rows - 1)
    total_height = sum(row_heights)

    gframe = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top),
                                     Inches(total_width), Inches(total_height))
    table = gframe.table
    table.first_row = False
    for i, w in enumerate(col_widths):
        table.columns[i].width = Inches(w)
    for r, h in enumerate(row_heights):
        table.rows[r].height = Inches(h)

    _set_cell(table.cell(0, 0), header_label, size=7)
    _set_cell(table.cell(1, 0), "", size=7)
    for ci, (bench_label, ate_label) in enumerate(cond_labels):
        c1 = 1 + ci * 2
        # 2026-08-15 (user request): Bench and ATE each get their OWN
        # header cell (their own matched power can differ slightly, see
        # bt_rx_ate_lookup.condition_title's docstring) -- no longer one
        # merged cell spanning both, which could only show a single number.
        _set_cell(table.cell(0, c1), bench_label, size=7)
        _set_cell(table.cell(0, c1 + 1), ate_label, size=7)
        _set_cell(table.cell(1, c1), "Bench", size=7)
        _set_cell(table.cell(1, c1 + 1), "ATE", size=7)

    _set_cell(table.cell(2, 0), "Mean±Std", size=7)
    _set_cell(table.cell(3, 0), "Min/Max", size=7)
    for ci, job in enumerate(jobs):
        c1 = 1 + ci * 2
        b_mean, b_minmax = _fmt_stat(job["bench_values"], value_fmt)
        a_mean, a_minmax = _fmt_stat(job["ate_values"], value_fmt)
        _set_cell(table.cell(2, c1), b_mean, size=7, bold=True)
        _set_cell(table.cell(2, c1 + 1), a_mean, size=7, bold=True)
        _set_cell(table.cell(3, c1), b_minmax, size=7, bold=True)
        _set_cell(table.cell(3, c1 + 1), a_minmax, size=7, bold=True)

    for row_i, (label,) in enumerate(spec_row_labels, start=4):
        spec_idx = 1 if label == "USL" else 0  # (lsl, usl)
        _set_cell(table.cell(row_i, 0), label, size=7, color=SPEC_COLOR)
        for ci, job in enumerate(jobs):
            c1 = 1 + ci * 2
            spec = job.get("spec")
            val = spec[spec_idx] if spec else None
            merged = table.cell(row_i, c1)
            merged.merge(table.cell(row_i, c1 + 1))
            _set_cell(merged, f"{val:{value_fmt}}" if val is not None else "", size=7,
                       bold=True, color=SPEC_COLOR)

    if show_power_level:
        row_i = 4 + n_spec_rows
        _set_cell(table.cell(row_i, 0), "Power Level", size=7, color=POWER_REF_COLOR)
        for ci, job in enumerate(jobs):
            c1 = 1 + ci * 2
            bench_pwr, ate_pwr = job.get("bench_pwr"), job.get("pwr")
            bench_text = f"{bench_pwr:{value_fmt}}dBm" if bench_pwr is not None else ""
            ate_text = f"{ate_pwr:{value_fmt}}dBm" if ate_pwr is not None else ""
            _set_cell(table.cell(row_i, c1), bench_text, size=7, bold=True, color=POWER_REF_COLOR)
            _set_cell(table.cell(row_i, c1 + 1), ate_text, size=7, bold=True, color=POWER_REF_COLOR)

    return total_height


def _add_image_stretched(slide, path: Path, left, top, width, height):
    slide.shapes.add_picture(str(path), Inches(left), Inches(top), Inches(width), Inches(height))


# 2026-08-27 (user report): widened from 0.35in to match Harald/Cal's own
# LABEL_WIDTH (0.9in) -- see BT_TX's build_ppt_pwr_evm_revision2.py's
# identical comment for the full story (a narrower column needed the
# row_label text shrunk/wrap-disabled to fit; widening it lets it stay a
# normal 16pt like Harald/Cal already do).
RULER_W = 0.9


def _add_label(slide, left, top, width, height, text, size=8, bold=True, shrink_to_fit=False, wrap=True):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    if shrink_to_fit:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    return box


def _add_ruler(slide, left, top, width, height, vmin, vmax, value_fmt, row_label=None):
    """Axis-reference strip, ONE per row (2026-08-15 customer request, "TX
    처럼 행 전체가 하나의 값을 공유하도록" -- match BT_TX's Power/DEVM
    ruler exactly, see build_ppt_pwr_evm_revision2.py's _add_ruler): Max
    flush with the image row's top edge, Min flush with its bottom edge.
    Meaningful here because rx.finalize_shared_ranges() now gives every
    column on a packet's slide the SAME Y-range (superseding an earlier
    per-column-ruler attempt, back when each column still had its own
    independent range).

    row_label (2026-08-26, user request): "Bench"/"ATE" text centered on
    this ruler -- this project already labels "Bench"/"ATE" in the table's
    own sub-header row (see add_packet_slide's own docstring for why that
    was judged sufficient at the time), but the user now wants the SAME
    label next to the chart row too, matching BT_TX/Cal/Harald's ruler."""
    tick_h = 0.18
    _add_label(slide, left, top, width, tick_h, f"{vmax:{value_fmt}}")
    _add_label(slide, left, top + height - tick_h, width, tick_h, f"{vmin:{value_fmt}}")
    if row_label:
        # 2026-08-27: RULER_W widened to 0.9in (was 0.35in) so this can
        # stay a plain 16pt, matching Harald/Cal -- see RULER_W's own
        # comment and BT_TX's build_ppt_pwr_evm_revision2.py's _add_ruler
        # docstring for the narrower-column workarounds this replaces.
        label_h = 0.4
        label_top = top + (height - label_h) / 2
        _add_label(slide, left, label_top, width, label_h, row_label, size=16, bold=True)


def add_packet_slide(prs, layout_idx, title_text, jobs, header_label, value_fmt,
                      bench_png_dir, ate_png_dir, show_power_level=False):
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    ph_title = _get_ph(slide, 1) or _get_ph(slide, 0)
    if ph_title:
        ph_title.text = title_text
    _remove_placeholder(slide, 12)

    usable_left = LEFT_MARGIN
    usable_right = SLIDE_WIDTH_IN - RIGHT_MARGIN
    n_cols = len(jobs)
    grid_left = usable_left + LABEL_WIDTH + RULER_W
    cell_width = (usable_right - grid_left - (n_cols - 1) * CELL_GAP_H) / n_cols

    # 2026-08-15 (user request): separate Bench/ATE header labels, each
    # showing THAT side's own matched power (see bt_rx_ate_lookup.
    # condition_title's docstring for why Bench's actual matched sweep
    # power and ATE's own nominal target power can differ slightly). No
    # "Bench"/"ATE" prefix on the number itself (2026-08-15, later same
    # day) -- the row right below already labels each column "Bench"/
    # "ATE", so repeating it here was redundant; the chart images' own
    # titles keep the prefix (see condition_title) since they have no such
    # adjacent subheader row of their own.
    cond_labels = [
        (f"band{job['band']}\nCH{job['ch']:03d}\n{rx.fmt_num(job['bench_pwr'])}dBm",
         f"band{job['band']}\nCH{job['ch']:03d}\n{rx.fmt_num(job['pwr'])}dBm")
        for job in jobs
    ]
    table_h = _add_interleaved_table(slide, usable_left, TABLE_TOP, usable_right - usable_left,
                                      header_label, cond_labels, jobs, value_fmt,
                                      show_power_level=show_power_level)

    plots_top = TABLE_TOP + table_h + TABLE_PLOT_GAP
    plots_bottom = CONTENT_AREA_BOTTOM
    image_h = (plots_bottom - plots_top - IMAGE_GAP) / 2
    bench_top = plots_top
    ate_top = bench_top + image_h + IMAGE_GAP

    # ONE ruler per row (2026-08-15) -- every job on this slide now shares
    # the identical value_range (see rx.finalize_shared_ranges), so any
    # job's own value_range speaks for the whole slide.
    shared_vr = jobs[0].get("value_range") if jobs else None
    if shared_vr:
        _add_ruler(slide, usable_left + LABEL_WIDTH, bench_top, RULER_W, image_h, shared_vr[0], shared_vr[1],
                   value_fmt, row_label="Bench")
        _add_ruler(slide, usable_left + LABEL_WIDTH, ate_top, RULER_W, image_h, shared_vr[0], shared_vr[1],
                   value_fmt, row_label="ATE")

    for ci, job in enumerate(jobs):
        cell_left = grid_left + ci * (cell_width + CELL_GAP_H)
        fname = rx.safe_filename(rx.combo_name(job["packet"], job["band"], job["ch"], job["pwr"])) + ".png"
        bench_path = bench_png_dir / fname
        ate_path = ate_png_dir / fname
        if bench_path.exists():
            _add_image_stretched(slide, bench_path, cell_left, bench_top, cell_width, image_h)
        if ate_path.exists():
            _add_image_stretched(slide, ate_path, cell_left, ate_top, cell_width, image_h)

    return slide


def _duplicate_slide(prs, source_slide):
    slide = prs.slides.add_slide(source_slide.slide_layout)
    for shape in list(slide.shapes):
        shape._element.getparent().remove(shape._element)
    for shape in source_slide.shapes:
        slide.shapes._spTree.append(deepcopy(shape._element))
    return slide


def _set_section_text(ph, text: str) -> None:
    p = ph.text_frame.paragraphs[0]
    for run in list(p.runs):
        run._r.getparent().remove(run._r)
    run = p.add_run()
    run.text = text
    run.font.bold = True
    run.font.size = SECTION_FONT_SIZE
    p.alignment = PP_ALIGN.CENTER


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


TOC_LINK_COLOR = RGBColor(0x1F, 0x4E, 0x79)  # dark blue, reads as a link -- same as BT_TX's build_ppt_combined_revision2.py
TOC_BLOCK_WIDTH_IN = 9.0  # centered column, NOT the full usable_left..usable_right
# chart-grid width -- pinning label/page range to opposite slide edges read
# as "spread across both sides" (2026-08-15 customer feedback, v2).
TOC_RULE_COLOR = RGBColor(0xD9, 0xD9, 0xD9)  # light gray row divider


def _add_toc_slide(prs, layout_idx, section_toc):
    """One Table-of-Contents slide, one row per section: "{label}" left,
    "{start}-{end}" page range right, BOTH clickable (2026-08-15 customer
    request, applied to RX to match BT_TX's build_ppt_combined_revision2.py).
    Each row is its own textbox so python-pptx's shape-level
    click_action.target_slide can jump straight to that section's own
    divider slide -- reordering slides afterward doesn't break this, since
    the link targets the slide PART, not its position.

    Laid out in a centered TOC_BLOCK_WIDTH_IN-wide column (not the full
    chart-grid usable width) with a thin rule under each row -- see
    build_ppt_combined_revision2.py's identical layout, v2 (2026-08-15)."""
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    _remove_placeholder(slide, 12)
    ph_title = _get_ph(slide, 1) or _get_ph(slide, 0)
    if ph_title:
        ph_title.text = "Table of Contents"

    usable_left = LEFT_MARGIN
    usable_right = SLIDE_WIDTH_IN - RIGHT_MARGIN
    block_width = min(TOC_BLOCK_WIDTH_IN, usable_right - usable_left)
    block_left = usable_left + (usable_right - usable_left - block_width) / 2

    top = CONTENT_AREA_TOP
    available_h = CONTENT_AREA_BOTTOM - top
    row_h = min(0.42, available_h / max(1, len(section_toc)))
    label_w = block_width * 0.7
    page_w = block_width * 0.3

    for i, (label, start, end, divider) in enumerate(section_toc):
        row_top = top + i * row_h

        label_box = slide.shapes.add_textbox(Inches(block_left), Inches(row_top), Inches(label_w), Inches(row_h))
        tf = label_box.text_frame
        tf.margin_left = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        run = tf.paragraphs[0].add_run()
        run.text = label
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.underline = True
        run.font.color.rgb = TOC_LINK_COLOR

        page_box = slide.shapes.add_textbox(Inches(block_left + label_w), Inches(row_top),
                                             Inches(page_w), Inches(row_h))
        tf2 = page_box.text_frame
        tf2.margin_left = 0
        tf2.margin_top = 0
        tf2.margin_bottom = 0
        tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.RIGHT
        run2 = p2.add_run()
        run2.text = f"{start}-{end}" if end != start else f"{start}"
        run2.font.size = Pt(14)
        run2.font.bold = True
        run2.font.color.rgb = TOC_LINK_COLOR

        label_box.click_action.target_slide = divider
        page_box.click_action.target_slide = divider

        rule = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(block_left), Inches(row_top + row_h - 0.02),
                                           Inches(block_left + block_width), Inches(row_top + row_h - 0.02))
        rule.line.color.rgb = TOC_RULE_COLOR
        rule.line.width = Pt(0.75)

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


def main() -> None:
    prs = Presentation(TEMPLATE_PATH)
    content_layout_idx = _find_layout(prs, "1_", 5)

    divider_source = prs.slides[0]
    first_ph = _get_ph(divider_source, 1) or _get_ph(divider_source, 0)

    sections = []  # (label, divider, packet_slides) -- see build_ppt_combined_revision2.py's `sections`
    for i, (label, mod, subdir, value_fmt_label) in enumerate(SECTIONS):
        print(f"Building {label} slides...")
        jobs, audit = mod.build_jobs()
        n_ate = len(audit)
        n_drawn = len(jobs)
        print(f"  {n_ate} ATE item(s), {n_drawn} with a Bench match (drawn)")

        if i == 0:
            _set_section_text(first_ph, label)
            divider = divider_source
        else:
            divider = _duplicate_slide(prs, divider_source)
            ph = _get_ph(divider, 1) or _get_ph(divider, 0)
            if ph:
                _set_section_text(ph, label)

        by_packet = defaultdict(list)
        for job in jobs:
            by_packet[job["packet"]].append(job)

        bench_dir = mod.bench_png_dir()
        # extract_mode="bench" (see paths.py's docstring / md/CLAUDE.md's
        # extract_mode section) means ATE is treated as genuinely absent,
        # not just "didn't re-run" -- pointing at a nonexistent dir makes
        # add_packet_slide's own `if ate_path.exists()` check naturally
        # skip every ATE image, same effect as every other project's
        # explicit "force ATE to an empty dict" gate.
        ate_dir = mod.ate_png_dir() if PATHS.extract_mode != "bench" else Path("__no_such_dir__")
        # PER is percent scale (0-100, see bt_rx_ate_lookup.BENCH_COLS'
        # `per-%` comment) -- ".2f" reads as e.g. "2.00", not the old
        # fraction-scale ".3f"'s "0.020".
        value_fmt = ".2e" if label == "Sensitivity_BER" else ".2f"

        # 2026-08-15 (user request, RSSI only): a "Power Level" table row +
        # matching red dashed reference line on each chart (see draw_png's
        # power_ref param / bt_rx_ate_lookup.finalize_shared_ranges'
        # include_power_ref param) -- RSSI is the only metric sharing dBm
        # units with the applied power level itself, so PER/BER never get
        # this row.
        show_power_level = label == "RSSI"

        packet_slides = []
        for packet in sorted(by_packet.keys()):
            packet_jobs = sorted(by_packet[packet], key=lambda j: (j["band"], j["ch"]))
            title = f"BT_RX_{label}_{packet}"
            slide = add_packet_slide(prs, content_layout_idx, title, packet_jobs,
                                      "Condition", value_fmt, bench_dir, ate_dir,
                                      show_power_level=show_power_level)
            packet_slides.append(slide)
        print(f"  {len(by_packet)} {label} slide(s)")
        sections.append((label, divider, packet_slides))

    title_slide = add_title_slide(prs)

    # 2026-08-15 (customer request, applied to RX to match BT_TX's
    # build_ppt_combined_revision2.py): Table of Contents, one hyperlinked
    # row per section with its page range. 2026-08-27 (customer request,
    # both projects): TOC moved to page 2 (was page 1) -- title=page 1,
    # TOC=page 2, each section's own divider slide IS that section's first
    # page -- unaffected by the title/TOC swap since content still starts
    # at page 3 either way.
    print("\nComputing Table of Contents page ranges...")
    section_toc = []
    page = 3
    for label, divider, packet_slides in sections:
        start = page
        end = page + len(packet_slides)
        section_toc.append((label, start, end, divider))
        page = end + 1
    toc_slide = _add_toc_slide(prs, content_layout_idx, section_toc)
    for label, start, end, _divider in section_toc:
        print(f"  {label}: {start}-{end}")

    ordered = [title_slide, toc_slide]
    for label, divider, packet_slides in sections:
        ordered.extend([divider, *packet_slides])
    _reorder_slides(prs, ordered)

    report_dir = PATHS.report_dir
    report_dir.mkdir(parents=True, exist_ok=True)
    out_path = report_dir / f"{REPORT_DATE}_BT_RX_RSSI_Sensitivity_Revision1.pptx"
    prs.save(out_path)
    print(f"\nSaved: {out_path}")
    print(f"Total slides: {len(prs.slides._sldIdLst)}  ({len(sections)} sections + 1 title + 1 TOC)")


if __name__ == "__main__":
    main()
