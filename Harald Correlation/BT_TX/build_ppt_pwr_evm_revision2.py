"""Build the Power/EVM revision2 PPT slides.

Layout rules (per customer spec, 2026-07-16; band merge + HPA titles 2026-08-14):
  * Band is derived from frequency (band_of): <3000 -> 2.4, else -> 5 (5G and
    6G share one "Band5G" bucket -- fewer slides, every HPA case for the
    combined high-freq range visible at a glance).
  * 2.4GHz band (only 3 frequencies): ONE slide per (metric, packet_type),
    covering all pa_supply 0/1/2 as three stacked row-blocks. Each row is
    [axis-ruler label][freq1 image][freq2 image][freq3 image], frequencies
    in ascending order. Frequency string centered above each image.
    Title: "BT_TX_{Power|EVM}_{packet_type}_Band2.4G"
  * 5G band (many frequencies, 6G merged in): too many columns to fit 3
    supply rows on one slide, so split by pa_supply -- ONE slide per
    (metric, packet_type, band, pa_supply), a single row of all that band's
    frequency images in ascending order, same ruler+freq-label treatment.
    Title: "BT_TX_{Power|EVM}_{packet_type}_Band5G_HPA_{1.5V|1.2V|0.73V}"
  * Axis ruler: since the embedded chart's own tick text becomes illegible
    once shrunk into a grid cell, a small manual ruler (5 evenly spaced
    values, max at top .. min at bottom, matching the chart's own
    max-at-top orientation) is drawn immediately left of each row's
    images, using that row's encompassing (min of mins, max of maxes)
    fixed Y-range across its frequencies.

This is a fresh script for revision2 -- revision1's build_ppt.py is no
longer used.
"""

from __future__ import annotations

import datetime
import re
import statistics
from collections import defaultdict
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

import make_bt_tx_pwr_revision2 as pwr_mod
import make_bt_tx_devm_rms_revision2 as devm_rms_mod
import make_bt_tx_devm_peak_revision2 as devm_peak_mod
import make_bt_tx_devm_99pct_revision2 as devm_99pct_mod
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")
BASE_DIR = PATHS.base_dir
TEMPLATE_PATH = PATHS.ppt_template
BT_TX_DIR = PATHS.bench_dir

PWR_PNG_DIR = PATHS.result_png_dir / "pwr_revision2"
DEVM_RMS_PNG_DIR = PATHS.result_png_dir / "devm_rms_revision2"
DEVM_PEAK_PNG_DIR = PATHS.result_png_dir / "devm_peak_revision2"
DEVM_99PCT_PNG_DIR = PATHS.result_png_dir / "devm_99pct_revision2"

FNAME_RE = re.compile(
    r"packet_type\((?P<packet>[^)]+)\)_frequency-MHz\((?P<freq>[^)]+)\)_"
    r"pa_supply\((?P<supply>[^)]+)\)_pa_slices\((?P<slices>[^)]+)\)_"
    r"pa_gain\((?P<gain>[^)]+)\)_dig_gain\((?P<dig>[^)]+)\)\.png"
)

SLIDE_WIDTH_IN = 13.338542213473316
LEFT_MARGIN = 0.05
RIGHT_MARGIN = 0.05
# 2026-08-27 (user report): was 0.6in -- too narrow to hold a readable
# "Bench"/"ATE" row_label on one line (see _add_ruler's docstring for the
# two failed narrower-column attempts). Widened to match Harald/Cal's own
# LABEL_WIDTH (0.9in, see Cal/make_bt_tx_trimtest_distribution.py), where
# the identical row_label at the same 16pt never needed shrinking.
RULER_WIDTH = 0.9
CELL_GAP_H = 0.03

# Format standard (established 2026-07-18 on the WL_TX deck, ported here
# 2026-07-19, see share/PPT_FORMAT_STANDARD.md): a stats TABLE sits under
# the title (replacing the old per-image freq/supply + inline USL/LSL
# labels), the chart grid fills only the top half of the real content area,
# and the bottom half is a reserved "ATE Data location" placeholder box.
#
# CONTENT_AREA_TOP/BOTTOM come from PPT-Template.pptx's own content
# placeholder (layout "1_", idx=12): top=1.5908508311461067in,
# height=5.240927384076991in -- verified identical to the WL_TX project's
# template (same file). Re-derive from that placeholder if the template
# ever changes; a guessed bottom margin WILL collide with the slide
# master's footer, which sits below this real content-area bottom.
CONTENT_AREA_TOP = 1.5908508311461067
CONTENT_AREA_BOTTOM = CONTENT_AREA_TOP + 5.240927384076991
# Bottom half of the CONTENT area (i.e. everything below the title) is left
# empty, reserved for ATE data to be placed there later.
GRID_CONTENT_BOTTOM = CONTENT_AREA_TOP + (CONTENT_AREA_BOTTOM - CONTENT_AREA_TOP) / 2

TABLE_TOP = 1.05
TABLE_ROW_H = 0.16  # 6pt table font -- single-line cells, no wasted padding
TABLE_GRID_GAP = 0.08
ATE_GAP = 0.1
# Gap between the Bench+ATE table row and the chart grid below it
# (2026-07-20: bumped up after the Bench chart's top edge was found
# clipping into the table above it in an exported slide -- PowerPoint's
# own render rounds row heights slightly taller than the XML-declared
# value in practice).
TABLE_PLOT_GAP = 0.35


# 2026-08-14: merged 5G+6G into one "Band5G" bucket -- see make_bt_tx_pwr_
# revision2.band_of's identical comment for the full rationale.
def band_of(freq: float) -> float:
    if freq < 3000:
        return 2.4
    return 5


def fmt_band(band: float) -> str:
    return f"{band:g}"


def discover(png_dir: Path):
    """Return data[packet][band][supply] = [(freq, path), ...] sorted by freq."""
    data: dict = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    for p in sorted(png_dir.glob("*.png")):
        m = FNAME_RE.match(p.name)
        if not m:
            print(f"  WARN: no match {p.name}")
            continue
        freq = float(m["freq"])
        band = band_of(freq)
        data[m["packet"]][band][m["supply"]].append((freq, p))
    for packet in data:
        for band in data[packet]:
            for supply in data[packet][band]:
                data[packet][band][supply].sort()
    return data


def _stats(values):
    """(mean, min, max, std) -- std is 0.0 for a single-point sample."""
    return (statistics.mean(values), min(values), max(values),
            statistics.stdev(values) if len(values) > 1 else 0.0)


def get_ranges_and_stats(mod, spec_by_packet=None):
    """ranges[(packet, gain, freq)] = (min, max), via the exact same
    computation the chart script itself uses, PLUS combo_stats[(packet,
    supply, gain, freq)] = (mean, min, max, std) from that same combo's
    outlier-cleaned values, so the stats table always matches what's
    actually histogrammed. Takes the chart module directly (pwr_mod /
    devm_rms_mod / devm_peak_mod / devm_99pct_mod).

    Power's build_packet_ranges(data) returns (ranges, outliers_by_combo)
    and takes no spec argument, with no kept-values already computed --
    kept values are recomputed here via the module's own
    find_combo_outliers. The DEVM family's build_packet_ranges(data,
    spec_by_packet=None) returns a longer tuple (ranges, kept_by_combo,
    outliers_by_combo, include_outliers_group, step_by_group) and needs
    spec_by_packet passed through so its Y-axis is the spec-anchored fixed
    window (2026-07-19); its kept_by_combo is reused directly instead of
    recomputing outliers a second time."""
    raw, pa_slices_seen, dig_gain_seen = mod.read_all(BT_TX_DIR)
    pa_slices_max = max(pa_slices_seen)
    dig_gain_max = max(dig_gain_seen)
    fixed = mod.build_fixed_combos(raw, pa_slices_max, dig_gain_max)
    if spec_by_packet is not None:
        result = mod.build_packet_ranges(fixed, spec_by_packet)
    else:
        result = mod.build_packet_ranges(fixed)
    ranges = result[0]

    combo_stats = {}
    if len(result) >= 4:  # DEVM family: kept_by_combo already computed
        kept_by_combo = result[1]
        for combo, kept in kept_by_combo.items():
            vals = [v for v, _dut in kept]
            if vals:
                combo_stats[combo] = _stats(vals)
    else:  # Power (and FreqAcc, via its own module): recompute kept values
        for combo, values in fixed.items():
            kept, _outliers = mod.find_combo_outliers(values)
            vals = [v for v, _dut in kept]
            if vals:
                combo_stats[combo] = _stats(vals)
    return ranges, combo_stats


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


def _add_label(slide, left, top, width, height, text, size=10, bold=True,
               shrink_to_fit=False, wrap=True):
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


def _set_cell(cell, text, size=6, bold=True, color=None):
    """text may contain '\\n' for a multi-line header (e.g. "S0\\n2412MHz")
    -- python-pptx's text-frame.text setter turns each '\\n'-separated
    piece into its own paragraph, so every paragraph must be styled, not
    just the first.

    A truly empty string (blank spec cell, e.g. USL/LSL columns with no
    limit) makes PowerPoint render that ROW much taller than TABLE_ROW_H --
    confirmed via win32com (11.52pt baseline vs 23.6pt with one empty
    cell): a run with no characters has nothing to measure, so PowerPoint
    falls back to a much larger default line height for that row instead
    of the 6pt we set on the (now-empty) run. A single space keeps the
    cell visually blank but gives PowerPoint an actual 6pt glyph to size
    the row against."""
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


def _grid_col_widths(cell_width, n_cols, gap=CELL_GAP_H):
    """Data-column widths for the stats table so each column's LEFT edge
    lines up with its corresponding chart image below it. The image grid
    has a `gap` gap between images (cell_left advances by cell_width+gap
    each column); a plain pptx table has no such inter-column gap, so
    without this every column after the first drifts left of its image by
    an accumulating amount (e.g. ACP's 12 columns drift by up to
    11*CELL_GAP_H =~ 0.33in by the last column). Widening every column but
    the last by `gap` reproduces that same per-column advance; the last
    column stays at the plain cell_width so the table's total width still
    matches the grid's total width exactly (n*cell_width + (n-1)*gap)."""
    if n_cols <= 1:
        return [cell_width] * n_cols
    return [cell_width + gap] * (n_cols - 1) + [cell_width]


def _add_stats_table(slide, left, col_widths, top, header_label, headers, data_rows,
                      colvar_rows=None, constant_rows=None, row_h=None):
    """Stats table directly under the title, columns aligned with the chart
    grid below it (col_widths[0] = RULER_WIDTH, the rest match each chart
    cell's width).

    data_rows: [(row_label, [value_str,...])] -- one cell per column, plain
      (Mean+-Std / Min/Max).
    colvar_rows: [(row_label, [value_str,...], color)] -- like data_rows but
      bold+colored, one cell per column, NOT merged -- for a spec limit
      (USL/LSL) that genuinely varies by column, e.g. BT_TX Power (differs
      by pa_supply/frequency block) and ACP (differs by offset).
    constant_rows: [(row_label, text, color)] -- a spec limit whose value is
      the SAME for every column on this slide (e.g. DEVM/FreqAcc, whose spec
      is keyed by packet only), so the data cells are MERGED into one
      instead of repeating the same text across every column.
    row_h (optional): per-row height in inches, default TABLE_ROW_H.

    Returns the table's total height in inches."""
    colvar_rows = colvar_rows or []
    constant_rows = constant_rows or []
    row_h = row_h if row_h is not None else TABLE_ROW_H
    n_cols = 1 + len(headers)
    n_rows = 1 + len(data_rows) + len(colvar_rows) + len(constant_rows)
    total_width = sum(col_widths)
    total_height = row_h * n_rows

    gframe = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top),
                                     Inches(total_width), Inches(total_height))
    table = gframe.table
    table.first_row = False
    for i, w in enumerate(col_widths):
        table.columns[i].width = Inches(w)
    for r in range(n_rows):
        table.rows[r].height = Inches(row_h)

    _set_cell(table.cell(0, 0), header_label, size=6)
    for ci, h in enumerate(headers, start=1):
        _set_cell(table.cell(0, ci), h, size=6)

    row_i = 1
    for label, values in data_rows:
        _set_cell(table.cell(row_i, 0), label, size=6)
        for ci, v in enumerate(values, start=1):
            _set_cell(table.cell(row_i, ci), v, size=6, bold=True)
        row_i += 1

    for label, values, color in colvar_rows:
        _set_cell(table.cell(row_i, 0), label, size=6, color=color)
        for ci, v in enumerate(values, start=1):
            _set_cell(table.cell(row_i, ci), v, size=6, bold=True, color=color)
        row_i += 1

    for label, text, color in constant_rows:
        _set_cell(table.cell(row_i, 0), label, size=6, color=color)
        merged_cell = table.cell(row_i, 1)
        if n_cols > 2:
            merged_cell.merge(table.cell(row_i, n_cols - 1))
        _set_cell(merged_cell, text, size=6, color=color)
        row_i += 1

    return total_height


ATE_BORDER_COLOR = RGBColor(0xC0, 0x00, 0x00)
ATE_TEXT_COLOR = RGBColor(0xC0, 0x00, 0x00)
USL_COLOR = RGBColor(0xB0, 0x00, 0x00)


def _add_ate_placeholder(slide, left, top, width, height):
    """Empty bordered box labeled "ATE Data location" -- reserved space for
    ATE data to be added later, per customer requirement (kept as a clearly
    marked empty region rather than being handed to the chart grid)."""
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.line.color.rgb = ATE_BORDER_COLOR
    box.line.width = Pt(1.5)
    box.fill.background()
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(6)
    tf.margin_top = Pt(4)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = "ATE Data location"
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = ATE_TEXT_COLOR
    return box


def _add_no_ate_box(slide, left, top, width, height):
    """Small per-column 'hole' for a combo with no ATE match -- unlike
    _add_ate_placeholder (one big box spanning the whole row), this only
    covers ONE column's cell so the other matched columns on the same
    slide still show their ATE chart (2026-07-20, see _add_ate_section)."""
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


def _add_interleaved_table(slide, left, top, table_width, header_label, cond_labels,
                            bench_data_rows, ate_lookup, spec_rows=None):
    """ONE combined table with Bench/ATE columns ALTERNATING per condition
    (2026-07-20 customer request, replacing the earlier side-by-side
    two-table layout: "Table이 Bench랑 ATE랑 번갈아가면서 나오도록"). Each
    condition gets a merged 2-column header cell (its freq/supply label)
    spanning a [Bench][ATE] sub-header pair; every data/spec row repeats
    that Bench-then-ATE pairing.

    cond_labels: [str, ...] one per condition, e.g. "S0\\n2402MHz".
    bench_data_rows: [(row_label, [bench_value_str, ...])] one value per
      condition, e.g. [("Mean±Std", [...]), ("Min/Max", [...])].
    ate_lookup: callable(row_label, condition_index) -> ate_value_str for
      that same row/condition (blank string where there's no ATE match).
    spec_rows: [(row_label, [value_or_blank, ...], color)] one value per
      condition -- merged across that condition's Bench+ATE pair since a
      spec limit is the same physical value regardless of which side
      measured it.

    Returns the table's total height in inches."""
    spec_rows = spec_rows or []
    n_cond = len(cond_labels)
    n_data_cols = 2 * n_cond
    n_cols = 1 + n_data_cols
    n_rows = 2 + len(bench_data_rows) + len(spec_rows)  # cond header + Bench/ATE subheader + data + spec

    grid_width = table_width - RULER_WIDTH
    cell_width = (grid_width - (n_data_cols - 1) * CELL_GAP_H) / n_data_cols
    col_widths = [RULER_WIDTH] + _grid_col_widths(cell_width, n_data_cols)
    total_width = sum(col_widths)
    total_height = TABLE_ROW_H * n_rows

    gframe = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top),
                                     Inches(total_width), Inches(total_height))
    table = gframe.table
    table.first_row = False
    for i, w in enumerate(col_widths):
        table.columns[i].width = Inches(w)
    for r in range(n_rows):
        table.rows[r].height = Inches(TABLE_ROW_H)

    _set_cell(table.cell(0, 0), header_label, size=6)
    _set_cell(table.cell(1, 0), "", size=6)
    for ci, label in enumerate(cond_labels):
        c1 = 1 + ci * 2
        merged = table.cell(0, c1)
        merged.merge(table.cell(0, c1 + 1))
        _set_cell(merged, label, size=6)
        _set_cell(table.cell(1, c1), "Bench", size=6)
        _set_cell(table.cell(1, c1 + 1), "ATE", size=6)

    row_i = 2
    for label, bench_vals in bench_data_rows:
        _set_cell(table.cell(row_i, 0), label, size=6)
        for ci in range(n_cond):
            c1 = 1 + ci * 2
            _set_cell(table.cell(row_i, c1), bench_vals[ci], size=6, bold=True)
            _set_cell(table.cell(row_i, c1 + 1), ate_lookup(label, ci), size=6, bold=True)
        row_i += 1

    for label, values, color in spec_rows:
        _set_cell(table.cell(row_i, 0), label, size=6, color=color)
        for ci in range(n_cond):
            c1 = 1 + ci * 2
            merged = table.cell(row_i, c1)
            merged.merge(table.cell(row_i, c1 + 1))
            _set_cell(merged, values[ci], size=6, bold=True, color=color)
        row_i += 1

    return total_height


def _add_ate_images(slide, usable_left, cell_width, top, height, cols, packet, gain, ate_data):
    """Per column: the ATE chart image (matched) or a small _add_no_ate_box
    (not matched -- a "hole", never blocking the other columns' charts).
    cols: [(supply, freq), ...] in column order."""
    for ci, (supply, freq) in enumerate(cols):
        cell_left = usable_left + RULER_WIDTH + ci * (cell_width + CELL_GAP_H)
        info = ate_data.get((packet, supply, gain, freq))
        png = info["png"] if info else None
        if png is not None:
            _add_image_stretched(slide, png, cell_left, top, cell_width, height)
        else:
            _add_no_ate_box(slide, cell_left, top, cell_width, height)


def _spec_row_values(get_spec, packet, gain, cols):
    """cols: [(freq, supply), ...] in column order. Returns (usl_strs,
    lsl_strs, any_usl, any_lsl) -- per-column string lists (blank where that
    column has no spec side), for the colvar_rows (per-column, unmerged)
    case."""
    usl_vals, lsl_vals = [], []
    any_usl = any_lsl = False
    for freq, supply in cols:
        spec = get_spec(packet, gain, freq, supply)
        lsl, usl = spec if spec else (None, None)
        usl_vals.append(f"{usl:g}" if usl is not None else "")
        lsl_vals.append(f"{lsl:g}" if lsl is not None else "")
        any_usl = any_usl or usl is not None
        any_lsl = any_lsl or lsl is not None
    return usl_vals, lsl_vals, any_usl, any_lsl


def _spec_rows(get_spec, packet, gain, cols, spec_constant):
    """Returns (colvar_rows, constant_rows) for _add_stats_table, given
    get_spec(packet, gain, freq, supply) -> (lsl, usl) or None.
    spec_constant=True (DEVM/FreqAcc: spec keyed by packet [+band], which
    doesn't vary across one slide's columns) merges into one cell;
    spec_constant=False (Power/ACP: spec genuinely differs column-to-column
    within a single slide) keeps one value per column instead."""
    if get_spec is None:
        return [], []
    if spec_constant:
        freq0, supply0 = cols[0]
        spec = get_spec(packet, gain, freq0, supply0)
        lsl, usl = spec if spec else (None, None)
        constant_rows = []
        if usl is not None:
            constant_rows.append(("USL", f"{usl:g}", USL_COLOR))
        if lsl is not None:
            constant_rows.append(("LSL", f"{lsl:g}", USL_COLOR))
        return [], constant_rows
    usl_vals, lsl_vals, any_usl, any_lsl = _spec_row_values(get_spec, packet, gain, cols)
    colvar_rows = []
    if any_usl:
        colvar_rows.append(("USL", usl_vals, USL_COLOR))
    if any_lsl:
        colvar_rows.append(("LSL", lsl_vals, USL_COLOR))
    return colvar_rows, []


def _add_ruler(slide, left, top, width, height, vmin, vmax, row_label=None, value_fmt="{:.0f}"):
    """Vertical axis-reference strip: Max flush with the image's top edge,
    Min flush with its bottom edge (chart itself now draws max-on-top).

    row_label (2026-08-26, user request -- add the "Bench"/"ATE" text
    Cal/Harald's own _add_ruler already shows, ported back here): centered
    between the vmax/vmin numbers, identifying which stacked image row this
    ruler belongs to. Cal/Harald's _add_ruler was originally COPIED from
    this exact function, then had row_label added on their side only --
    this brings BT_TX's own copy back in sync.

    value_fmt (2026-08-27, additive -- default unchanged): every OTHER
    metric family's Y-axis spans many whole units, so the old hardcoded
    ":.0f" was fine. BW06DB/BW20DB's per-group spreads run as small as
    ~0.005MHz (see make_bt_tx_bw_revision1.py's module docstring) -- ":.0f"
    would print the SAME integer for both vmin and vmax, silently erasing
    the axis. Callers pass a metric-appropriate format string instead of
    every other family's behavior changing."""
    tick_h = 0.2
    _add_label(slide, left, top, width, tick_h, value_fmt.format(vmax), size=9, bold=True)
    _add_label(slide, left, top + height - tick_h, width, tick_h, value_fmt.format(vmin), size=9, bold=True)
    if row_label:
        # 2026-08-27: two earlier attempts at the old RULER_WIDTH=0.6in
        # (shrink_to_fit autofit, then a smaller fixed size + wrap=False)
        # both worked around a column too narrow for "Bench" at a normal
        # size. User asked for the same look Harald/Cal already have there
        # (plain 16pt, never shrunk) -- fixed at the root instead by
        # widening RULER_WIDTH itself to match Harald/Cal's own 0.9in (see
        # RULER_WIDTH's own comment above), so 16pt fits on one line
        # without any of the narrower-column workarounds.
        label_h = 0.4
        label_top = top + (height - label_h) / 2
        _add_label(slide, left, label_top, width, label_h, row_label, size=16, bold=True)


def _add_image_stretched(slide, path: Path, cell_left, cell_top, cell_width, cell_height):
    slide.shapes.add_picture(
        str(path), Inches(cell_left), Inches(cell_top), Inches(cell_width), Inches(cell_height)
    )


def add_band24_slide(prs, layout_idx, title_text, band_data, ranges, packet, gain, combo_stats,
                      get_spec=None, spec_constant=True, ate_data=None, ruler_fmt="{:.0f}"):
    """ONE row: pa_supply0's 3 freq images, then pa_supply1's 3, then
    pa_supply2's 3 -- 9 images total in sequence, not a 3x3 grid. Stats
    table (Mean/Min/Max/Std per column, header "S{supply}\\n{freq}MHz",
    plus a USL/LSL row if get_spec is given) sits under the title; ATE
    placeholder box fills the reserved bottom half (format standard
    established 2026-07-18 on the WL_TX deck, see
    share/PPT_FORMAT_STANDARD.md).

    ate_data (2026-07-20, optional): {(packet,supply,gain,freq): {"png":
    Path|None, "stats": (mean,min,max,std)|None}}. When given, the bottom
    half gets its OWN stats table (mirroring the Bench one above, not
    merged with it) plus one ATE image per column where matched, or a
    small per-column "No ATE match" box where not -- a single unmatched
    column never blocks the other columns' ATE charts on the same slide.
    When None (the default -- DEVM/FreqAcc callers that haven't been wired
    up yet), falls back to the original single empty placeholder box."""
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    ph_title = _get_ph(slide, 1) or _get_ph(slide, 0)
    if ph_title:
        ph_title.text = title_text
    _remove_placeholder(slide, 12)

    supplies = sorted(band_data.keys())
    block_size = max(len(v) for v in band_data.values())
    n_cols = block_size * len(supplies)

    usable_left = LEFT_MARGIN
    usable_right = SLIDE_WIDTH_IN - RIGHT_MARGIN
    grid_width = usable_right - usable_left - RULER_WIDTH
    cell_width = (grid_width - (n_cols - 1) * CELL_GAP_H) / n_cols

    ordered = [(supply, freq, path) for supply in supplies for freq, path in band_data[supply]]
    # 2026-08-14 (customer request): "S{N}" column header -> HPA voltage
    # label (e.g. "HPA 1.5V"), same customer-confirmed mapping used for
    # slide/chart titles -- see make_bt_tx_pwr_revision2.SUPPLY_TO_HPA.
    headers = [f"HPA {pwr_mod.hpa_label(supply)}\n{freq:g}MHz" for supply, freq, _path in ordered]
    mean_std, min_max = [], []
    for supply, freq, _path in ordered:
        s = combo_stats.get((packet, supply, gain, freq))
        if s:
            mean_std.append(f"{s[0]:.2f}±{s[3]:.2f}")
            min_max.append(f"{s[1]:.2f} / {s[2]:.2f}")
        else:
            mean_std.append(""); min_max.append("")
    data_rows = [("Mean±Std", mean_std), ("Min/Max", min_max)]
    colvar_rows, constant_rows = _spec_rows(
        get_spec, packet, gain, [(freq, supply) for supply, freq, _path in ordered], spec_constant)

    cols = [(supply, freq) for supply, freq, _path in ordered]

    if ate_data is not None:
        # 2026-07-20 customer request: ONE combined table with Bench/ATE
        # columns ALTERNATING per condition ("Table이 Bench랑 ATE랑
        # 번갈아가면서 나오도록") -- see _add_interleaved_table's docstring.
        # The chart grid below is unaffected (still the original two
        # full-width stacked rows, Bench row on top / ATE row below).
        def ate_lookup(label, ci, _cols=cols):
            supply, freq = _cols[ci]
            info = ate_data.get((packet, supply, gain, freq))
            stats = info["stats"] if info else None
            if not stats:
                return ""
            return f"{stats[0]:.2f}±{stats[3]:.2f}" if label == "Mean±Std" else f"{stats[1]:.2f} / {stats[2]:.2f}"

        spec_rows = list(colvar_rows) + [(label, [text] * n_cols, color) for label, text, color in constant_rows]
        table_h = _add_interleaved_table(slide, usable_left, TABLE_TOP, usable_right - usable_left,
                                          "HPA/Freq", headers, data_rows, ate_lookup, spec_rows=spec_rows)
    else:
        col_widths = [RULER_WIDTH] + _grid_col_widths(cell_width, n_cols)
        table_h = _add_stats_table(slide, usable_left, col_widths, TABLE_TOP, "HPA/Freq", headers,
                                    data_rows, colvar_rows=colvar_rows, constant_rows=constant_rows)

    vmin, vmax = None, None
    for freq_list in band_data.values():
        for freq, _path in freq_list:
            r = ranges.get((packet, gain, band_of(freq)))
            if r is None:
                continue
            vmin = r[0] if vmin is None else min(vmin, r[0])
            vmax = r[1] if vmax is None else max(vmax, r[1])
    if vmin is None:
        vmin, vmax = 0, 1

    if ate_data is not None:
        plots_top = TABLE_TOP + table_h + TABLE_PLOT_GAP
        plots_bottom = CONTENT_AREA_BOTTOM
        bench_image_h = (plots_bottom - plots_top - ATE_GAP) / 2
        bench_image_top = plots_top
        ate_image_top = bench_image_top + bench_image_h + ATE_GAP
        ate_image_h = plots_bottom - ate_image_top

        _add_ruler(slide, usable_left, bench_image_top, RULER_WIDTH, bench_image_h, vmin, vmax, row_label="Bench",
                   value_fmt=ruler_fmt)
        col = 0
        for supply in supplies:
            for freq, path in band_data[supply]:
                cell_left = usable_left + RULER_WIDTH + col * (cell_width + CELL_GAP_H)
                _add_image_stretched(slide, path, cell_left, bench_image_top, cell_width, bench_image_h)
                col += 1

        _add_ruler(slide, usable_left, ate_image_top, RULER_WIDTH, ate_image_h, vmin, vmax, row_label="ATE",
                   value_fmt=ruler_fmt)
        _add_ate_images(slide, usable_left, cell_width, ate_image_top, ate_image_h, cols, packet, gain, ate_data)
    else:
        content_top = TABLE_TOP + table_h + TABLE_GRID_GAP
        content_bottom = GRID_CONTENT_BOTTOM
        image_h = content_bottom - content_top

        _add_ruler(slide, usable_left, content_top, RULER_WIDTH, image_h, vmin, vmax, value_fmt=ruler_fmt)
        col = 0
        for supply in supplies:
            for freq, path in band_data[supply]:
                cell_left = usable_left + RULER_WIDTH + col * (cell_width + CELL_GAP_H)
                _add_image_stretched(slide, path, cell_left, content_top, cell_width, image_h)
                col += 1

        _add_ate_placeholder(slide, usable_left, GRID_CONTENT_BOTTOM + ATE_GAP,
                              usable_right - usable_left, CONTENT_AREA_BOTTOM - (GRID_CONTENT_BOTTOM + ATE_GAP))

    return slide


def add_highfreq_supply_slide(prs, layout_idx, title_text, freq_list, ranges, packet, gain, combo_stats,
                               supply=None, get_spec=None, spec_constant=True, ate_data=None, ruler_fmt="{:.0f}"):
    """Single row (one pa_supply), all of that band's frequency images in
    order. Stats table (Mean/Min/Max/Std per frequency, plus a USL/LSL row
    if get_spec is given) sits under the title; ATE placeholder box fills
    the reserved bottom half (or, when ate_data is given, its own separate
    ATE stats table + per-column chart/no-match -- see add_band24_slide's
    docstring for the same 2026-07-20 addition)."""
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    ph_title = _get_ph(slide, 1) or _get_ph(slide, 0)
    if ph_title:
        ph_title.text = title_text
    _remove_placeholder(slide, 12)

    n_cols = len(freq_list)
    usable_left = LEFT_MARGIN
    usable_right = SLIDE_WIDTH_IN - RIGHT_MARGIN
    grid_width = usable_right - usable_left - RULER_WIDTH
    cell_width = (grid_width - (n_cols - 1) * CELL_GAP_H) / n_cols

    headers = [f"{freq:g}" for freq, _path in freq_list]
    mean_std, min_max = [], []
    for freq, _path in freq_list:
        s = combo_stats.get((packet, supply, gain, freq))
        if s:
            mean_std.append(f"{s[0]:.2f}±{s[3]:.2f}")
            min_max.append(f"{s[1]:.2f} / {s[2]:.2f}")
        else:
            mean_std.append(""); min_max.append("")
    data_rows = [("Mean±Std", mean_std), ("Min/Max", min_max)]
    colvar_rows, constant_rows = _spec_rows(
        get_spec, packet, gain, [(freq, supply) for freq, _path in freq_list], spec_constant)

    cols = [(supply, freq) for freq, _path in freq_list]

    if ate_data is not None:
        # 2026-07-20 customer request: ONE combined table with Bench/ATE
        # columns ALTERNATING per condition -- see add_band24_slide's
        # identical restructuring for the full rationale.
        def ate_lookup(label, ci, _cols=cols):
            supply_, freq = _cols[ci]
            info = ate_data.get((packet, supply_, gain, freq))
            stats = info["stats"] if info else None
            if not stats:
                return ""
            return f"{stats[0]:.2f}±{stats[3]:.2f}" if label == "Mean±Std" else f"{stats[1]:.2f} / {stats[2]:.2f}"

        spec_rows = list(colvar_rows) + [(label, [text] * n_cols, color) for label, text, color in constant_rows]
        table_h = _add_interleaved_table(slide, usable_left, TABLE_TOP, usable_right - usable_left,
                                          "Freq (MHz)", headers, data_rows, ate_lookup, spec_rows=spec_rows)
    else:
        col_widths = [RULER_WIDTH] + _grid_col_widths(cell_width, n_cols)
        table_h = _add_stats_table(slide, usable_left, col_widths, TABLE_TOP, "Freq (MHz)", headers,
                                    data_rows, colvar_rows=colvar_rows, constant_rows=constant_rows)

    vmin, vmax = None, None
    for freq, _path in freq_list:
        # The DEVM family (2026-07-18) can key 5G/6G ranges per-supply (one
        # axis per slide, see e.g. make_bt_tx_devm_rms_revision2.group_of)
        # -- try that first, falling back to the plain (packet,gain,band)
        # key Power always uses.
        r = ranges.get((packet, gain, band_of(freq), supply))
        if r is None:
            r = ranges.get((packet, gain, band_of(freq)))
        if r is None:
            continue
        vmin = r[0] if vmin is None else min(vmin, r[0])
        vmax = r[1] if vmax is None else max(vmax, r[1])
    if vmin is None:
        vmin, vmax = 0, 1

    if ate_data is not None:
        plots_top = TABLE_TOP + table_h + TABLE_PLOT_GAP
        plots_bottom = CONTENT_AREA_BOTTOM
        bench_image_h = (plots_bottom - plots_top - ATE_GAP) / 2
        bench_image_top = plots_top
        ate_image_top = bench_image_top + bench_image_h + ATE_GAP
        ate_image_h = plots_bottom - ate_image_top

        _add_ruler(slide, usable_left, bench_image_top, RULER_WIDTH, bench_image_h, vmin, vmax, row_label="Bench",
                   value_fmt=ruler_fmt)
        for ci, (freq, path) in enumerate(freq_list):
            cell_left = usable_left + RULER_WIDTH + ci * (cell_width + CELL_GAP_H)
            _add_image_stretched(slide, path, cell_left, bench_image_top, cell_width, bench_image_h)

        _add_ruler(slide, usable_left, ate_image_top, RULER_WIDTH, ate_image_h, vmin, vmax, row_label="ATE",
                   value_fmt=ruler_fmt)
        _add_ate_images(slide, usable_left, cell_width, ate_image_top, ate_image_h, cols, packet, gain, ate_data)
    else:
        content_top = TABLE_TOP + table_h + TABLE_GRID_GAP
        content_bottom = GRID_CONTENT_BOTTOM
        image_h = content_bottom - content_top

        _add_ruler(slide, usable_left, content_top, RULER_WIDTH, image_h, vmin, vmax, value_fmt=ruler_fmt)
        for ci, (freq, path) in enumerate(freq_list):
            cell_left = usable_left + RULER_WIDTH + ci * (cell_width + CELL_GAP_H)
            _add_image_stretched(slide, path, cell_left, content_top, cell_width, image_h)

        _add_ate_placeholder(slide, usable_left, GRID_CONTENT_BOTTOM + ATE_GAP,
                              usable_right - usable_left, CONTENT_AREA_BOTTOM - (GRID_CONTENT_BOTTOM + ATE_GAP))

    return slide


# 2026-08-14: derived from the actual data (same convention as pa_slices_max/
# dig_gain_max elsewhere in this project), NOT hardcoded. The literal "0"
# this used to be assumed the OLD 40-DUT pull's own pa_gain value -- once a
# pull with a different gain came in (this one: '1' throughout), every
# combo_stats/spec/ATE lookup keyed by (packet,supply,gain,freq) silently
# stopped matching, because the gain component of every lookup key never
# matched a real key. Root cause of a real bug found 2026-08-14: every
# "Mean+/-Std"/"Min/Max" table row AND every ATE image across the WHOLE
# combined deck (Power/DEVM/ACP-shared-code/FreqAcc, everything that goes
# through build_metric_slides()) was silently blank/"No ATE match" even
# though the underlying data was correctly matched -- the bug was purely in
# this one stale lookup key, not in any of the matching logic itself.
_gains_seen = {k[2] for k in pwr_mod.read_all(BT_TX_DIR)[0].keys()}
GAIN = max(_gains_seen, key=float) if _gains_seen else "0"
# Computed at run time (not hardcoded) so every rebuild's filename/subtitle
# reflects the actual build date instead of going stale (2026-07-18 fix).
REPORT_DATE = datetime.date.today().isoformat()
# Default only -- build_ppt_full_revision2.py and build_ppt_freqacc_
# revision2.py both overwrite this (PE.TITLE_TEXT = ...) with the
# config.xml-driven "{title_prefix} {project} {title_suffix}" pattern
# before calling add_title_slide(); this module is never run standalone.
TITLE_TEXT = f"{PATHS.title_prefix} {PATHS.project} {PATHS.title_suffix}"
SUBTITLE_LINES = ["RF LAB1 Team", REPORT_DATE]
SECTION_FONT_SIZE = Pt(40)


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


def make_divider_factory(prs):
    """Return a `new_divider(text)` that stamps and hands back one section
    divider slide. 2026-07-28.

    The template's slide 0 IS a divider: the first section consumes it in
    place and every later section gets a duplicate. Getting that wrong doesn't
    look broken -- _reorder_slides() below simply omits an unreused slide 0
    from the deck while its part stays in the package as dead weight.

    Lives here (rather than in one of the two deck builders) because
    build_ppt_full_revision2.py, build_ppt_freqacc_revision2.py, and
    build_ppt_combined_revision2.py all need the same bookkeeping, and all
    three already import this module. Same closure pattern UWB's and WIFI's
    deck builders use.
    """
    divider_source = prs.slides[0]
    state = {"source_used": False}

    def new_divider(text: str):
        if state["source_used"]:
            divider = _duplicate_slide(prs, divider_source)
        else:
            divider = divider_source
            state["source_used"] = True
        ph = _get_ph(divider, 1) or _get_ph(divider, 0)
        if ph:
            _set_section_text(ph, text)
        return divider

    return new_divider


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


def build_metric_slides(prs, layout_idx, metric_label, png_dir, ranges, combo_stats, get_spec=None,
                         spec_constant=True, ate_data=None, ruler_fmt="{:.0f}"):
    """ate_data (2026-07-20, optional): {(packet,supply,gain,freq): {"png":.., "stats":..}}
    -- see add_band24_slide's docstring. None (default) preserves the
    original single-empty-placeholder behavior for metrics that don't have
    an ATE counterpart generator yet (DEVM/FreqAcc/ACP).

    ruler_fmt (2026-08-27, additive, default unchanged): format string for
    the ruler's vmin/vmax labels -- see _add_ruler's docstring. Every
    existing caller keeps the old ":.0f" behavior; BW06DB/BW20DB pass a
    decimal-aware format since their axis spans can be well under 1 unit."""
    data = discover(png_dir)
    slides = []
    for packet in sorted(data.keys()):
        band24 = data[packet].get(2.4)
        if band24:
            # 2026-08-27 (customer request via LEE SeokJin): every slide
            # title must show its HPA voltage -- for a single combined
            # Band2.4G slide (multiple supplies stacked as row-blocks) that
            # isn't possible, since no single HPA value describes the
            # whole slide. Rule confirmed with the customer: 2 supplies
            # present -> keep them combined on ONE slide (as before, no
            # HPA suffix, since there's more than one value to name); ANY
            # other count (1, or 3+) -> split into one slide PER supply
            # instead, reusing add_highfreq_supply_slide (already proven
            # for Band5G, which already titles per-HPA the same way)
            # rather than writing new slide-drawing code. The count==1
            # case is NOT an edge case to special-handle separately -- most
            # BT_TX packets (e.g. 8DH5, 4DH5, every HDRP family) only ever
            # sweep ONE HPA rail at 2.4G, so this is the common case, not
            # a rare one; a first version of this fix only checked ">= 3"
            # and silently left every 1-supply packet's title unlabeled
            # (caught via a real screenshot, BT_TX_BW06DB_8DH5_Band2.4G
            # missing its "_HPA_1.5V" suffix even though the whole slide
            # was that one voltage).
            if len(band24) != 2:
                for supply in sorted(band24.keys()):
                    freq_list = band24[supply]
                    title = f"BT_TX_{metric_label}_{packet}_Band2.4G_HPA_{pwr_mod.hpa_label(supply)}"
                    slides.append(add_highfreq_supply_slide(
                        prs, layout_idx, title, freq_list, ranges, packet, GAIN, combo_stats,
                        supply=supply, get_spec=get_spec, spec_constant=spec_constant, ate_data=ate_data,
                        ruler_fmt=ruler_fmt))
            else:
                title = f"BT_TX_{metric_label}_{packet}_Band2.4G"
                slides.append(add_band24_slide(prs, layout_idx, title, band24, ranges, packet, GAIN, combo_stats,
                                                get_spec=get_spec, spec_constant=spec_constant, ate_data=ate_data,
                                                ruler_fmt=ruler_fmt))
        for band in (5,):
            band_data = data[packet].get(band)
            if not band_data:
                continue
            for supply in sorted(band_data.keys()):
                freq_list = band_data[supply]
                title = f"BT_TX_{metric_label}_{packet}_Band{fmt_band(band)}G_HPA_{pwr_mod.hpa_label(supply)}"
                slides.append(add_highfreq_supply_slide(prs, layout_idx, title, freq_list, ranges, packet, GAIN,
                                                          combo_stats, supply=supply, get_spec=get_spec,
                                                          spec_constant=spec_constant, ate_data=ate_data,
                                                          ruler_fmt=ruler_fmt))
    return slides


DEVM_SECTIONS = [
    ("rms", "DEVM RMS", devm_rms_mod, DEVM_RMS_PNG_DIR),
    ("peak", "DEVM Peak", devm_peak_mod, DEVM_PEAK_PNG_DIR),
    ("99pct", "DEVM99pct", devm_99pct_mod, DEVM_99PCT_PNG_DIR),
]


def main() -> None:
    prs = Presentation(TEMPLATE_PATH)
    content_layout_idx = _find_layout(prs, "1_", 5)

    print("Computing Power Y-axis ranges and stats...")
    pwr_ranges, pwr_stats = get_ranges_and_stats(pwr_mod)

    print("Parsing Power spec limits...")
    import bt_tx_power_spec_lookup as power_spec_lookup
    power_spec_table = power_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    def get_power_spec(packet, gain, freq, supply):
        return power_spec_lookup.lookup(power_spec_table, packet, freq, supply)

    print("Parsing DEVM spec limits...")
    import bt_tx_devm_spec_lookup as devm_spec_lookup
    devm_spec_table = devm_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    divider_source = prs.slides[0]
    first_ph = _get_ph(divider_source, 1) or _get_ph(divider_source, 0)
    if first_ph:
        _set_section_text(first_ph, "Power")
    power_divider = divider_source

    print("Building Power slides...")
    # Power's USL/LSL genuinely differs column-to-column within one slide
    # (e.g. by pa_supply/frequency block) -- spec_constant=False shows it
    # per-column instead of merging (see share/PPT_FORMAT_STANDARD.md and
    # the 2026-07-19 clarification this project needed beyond that doc).
    power_slides = build_metric_slides(prs, content_layout_idx, "Power", PWR_PNG_DIR, pwr_ranges, pwr_stats,
                                        get_spec=get_power_spec, spec_constant=False)
    print(f"  {len(power_slides)} Power slides")

    ordered = [power_divider, *power_slides]
    for devm_metric, label, mod, png_dir in DEVM_SECTIONS:
        print(f"Computing {label} Y-axis ranges and stats...")
        ranges, stats = get_ranges_and_stats(mod, devm_spec_table.get(devm_metric, {}))

        def get_spec(packet, gain, freq, supply, _metric=devm_metric):
            return devm_spec_lookup.lookup(devm_spec_table, _metric, packet)

        divider = _duplicate_slide(prs, divider_source)
        ph = _get_ph(divider, 1) or _get_ph(divider, 0)
        if ph:
            _set_section_text(ph, label)

        print(f"Building {label} slides...")
        # DEVM's spec is keyed by packet only -- constant across every
        # column of a given slide, so spec_constant=True merges it into one
        # cell (matches the WL_TX/share reference exactly).
        slides = build_metric_slides(prs, content_layout_idx, label, png_dir, ranges, stats,
                                      get_spec=get_spec, spec_constant=True)
        print(f"  {len(slides)} {label} slides")
        ordered.extend([divider, *slides])

    title_slide = add_title_slide(prs)
    ordered = [title_slide, *ordered]
    _reorder_slides(prs, ordered)

    report_dir = PATHS.report_dir
    report_dir.mkdir(parents=True, exist_ok=True)
    out_path = report_dir / f"{REPORT_DATE}_BT_TX_PWR_DEVM_Revision2.pptx"
    prs.save(out_path)
    print(f"\nSaved: {out_path}")
    print(f"Total slides: {len(prs.slides._sldIdLst)}")


if __name__ == "__main__":
    main()
