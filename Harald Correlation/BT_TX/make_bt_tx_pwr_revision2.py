"""BT TX Power revision2 distribution histograms.

Same pa_slices/dig_gain-fixed-at-max rule as revision1, but band merging
is DROPPED per customer feedback: each EXACT frequency-MHz gets its own
histogram again (no band-level merge, no per-frequency color breakdown
needed since each chart now has only one frequency).

The chart itself is also rotated 90 degrees from the original vertical
detail histogram: Count is on the X axis, value bins are on the Y axis
(horizontal bars).

For every (packet_type, pa_supply, pa_gain, frequency-MHz) combination,
collect every power_aver-dBm value from every DUT and draw a horizontal
bar-chart histogram:
  X axis = Count
  Y axis = Power value bins

Two outputs per combo:
  * Excel workbook: horizontal bar chart on top, bin/count table + raw
    values below.
  * PNG image.

File name = packet_type(..)_frequency-MHz(..)_pa_supply(..)_pa_slices(..)_pa_gain(..)_dig_gain(..)
Chart title is prefixed with "BT_TX_".
"""

from __future__ import annotations

import argparse
import csv
import functools
import math
import re
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties

import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")
from openpyxl.styles import Alignment, Font, PatternFill

import bt_tx_power_spec_lookup as spec_lookup


RAW_KEY_COLS = [
    "cfg-carrier_01-packet_type",
    "cfg-carrier_01-pa_supply",
    "cfg-carrier_01-pa_gain",
    "cfg-carrier_01-pa_slices",
    "cfg-carrier_01-dig_gain",
    "cfg-carrier_01-frequency-MHz",
]
Y_COL = "power_aver-dBm"
# 2026-08-15 real bug found (customer-confirmed months ago, Slack -- "PA
# Slice variants... and data_pattern=PRBS9 fixed" -- but never actually
# implemented): cfg-carrier_01-data_pattern was missing from RAW_KEY_COLS
# entirely. Bench sweeps 3 data patterns (PRBS9/PAT1/PAT2) for 1DH5/LE1M/
# LE2M specifically (every other packet only ever has PRBS9 rows) -- all 3
# were silently pooled into one histogram, tripling the real sample count
# (30 instead of 10 DUTs) and skewing Mean/Std/Min/Max, while ATE only
# ever measures the PRBS9 condition. Filtering to PRBS9-only here makes
# Bench's real sample count match ATE's again.
DATA_PATTERN = "PRBS9"
BINS = 40
DUT_RE = re.compile(r"#(\d+)")
IQR_MULT = 5.0


# 2026-08-14 (customer request): 5G and 6G share one slide/title bucket
# ("Band5G") -- fewer pages, and every HPA case for the combined high-freq
# range is visible at a glance. Real frequencies are untouched; only this
# bucketing label collapses.
def band_of(freq: float) -> float:
    if freq < 3000:
        return 2.4
    return 5


# pa_supply index -> HPA rail voltage, confirmed by the customer (Slack
# message, 2026-08-13): 0->1.5V, 1->1.2V, 2->0.73V. See bt_tx_power_spec_
# lookup.PA_SUPPLY_VOLT_TO_INDEX for the inverse mapping used in spec
# lookups. 2026-08-14: used to replace the "supply(N)"/"pa_supplyN" chart
# and slide title tokens with the customer's preferred "HPA_X.XV" label
# (Testcoverage workbook terminology) -- filenames are untouched (see
# combo_title_short()'s docstring for why).
SUPPLY_TO_HPA = {"0": "1.5V", "1": "1.2V", "2": "0.73V"}


def hpa_label(supply) -> str:
    return SUPPLY_TO_HPA.get(str(supply), f"supply({supply})")


def hpa_token(supply) -> str:
    """2026-08-27 (customer request via LEE SeokJin): Summary Test Item
    names need the HPA voltage too, e.g. "BT_TX_Power_1DH5_HPA1P5_2402MHz"
    (was "..._PaSupply0_..."). Same voltage as hpa_label(), reformatted to
    a bare identifier token (no space/dot/unit) matching the customer's
    own example: "1.5V" -> "1P5", "1.2V" -> "1P2", "0.73V" -> "0P73" --
    same style as the ATE log's own NEW_SUPPLY_LABEL_TO_INDEX tokens (see
    make_ate_tx_pwr_revision2.py), just the inverse direction."""
    return hpa_label(supply).replace(".", "P").replace("V", "")


def to_float(text: str):
    text = (text or "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def fmt_num(value: float) -> str:
    return f"{value:g}"


def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:\[\]<>|"]', "_", name)


def dut_id_of(path: Path) -> str:
    m = DUT_RE.search(path.name)
    return f"#{m.group(1)}" if m else path.stem


@functools.lru_cache(maxsize=None)
def read_all(bt_tx_dir: Path):
    """Single pass: build raw[(packet,supply,gain,pa_slices,dig_gain,freq)] = [(value,dut_id)].

    Also returns the full set of pa_slices/dig_gain values seen so the
    caller can determine the dataset-wide maximum.

    2026-07-20: switched from csv.DictReader (allocates a full dict of
    every column per row) to csv.reader + a one-time header->index map --
    only ~7 of this CSV's many columns are ever used, so per-row dict
    construction was pure overhead. @lru_cache means a build script that
    calls this twice in one process (once for its own stats, once via the
    sibling ATE module's get_ate_data()) only actually reads the 40 files
    once -- safe because nothing downstream mutates the returned raw
    dict/sets, only reads them."""
    raw: dict[tuple, list[tuple[float, str]]] = defaultdict(list)
    pa_slices_seen: set[float] = set()
    dig_gain_seen: set[float] = set()

    # 2026-08-25: rglob + suffix filter, not a flat glob("*.csv") -- a
    # newer Bench pull nests each DUT's file under its own #NN_<serial>/
    # folder (Bench/#01_725HW.../harald_bt_tx_..._cns_temp_data.csv)
    # instead of 40 files directly in bt_tx_dir, and that folder ALSO has
    # a companion .xlsx report plus a summary/summary_<timestamp>.csv (a
    # completely different file -- test pass-rate/timing metadata, not
    # measurement rows) that a blanket "*.csv" would wrongly sweep in too.
    # Matching the real data file's own "_cns_temp_data.csv" suffix finds
    # it at any depth while skipping both of those, and still finds the
    # exact same files as before for the older flat layout (which has
    # nothing else in it to filter out).
    files = sorted(bt_tx_dir.rglob("*_cns_temp_data.csv"))
    for path in files:
        dut = dut_id_of(path)
        with path.open(encoding="utf-8-sig", newline="") as fh:
            reader = csv.reader(fh)
            header = next(reader)
            col = {name: i for i, name in enumerate(header)}
            i_packet = col.get(RAW_KEY_COLS[0])
            i_supply = col.get(RAW_KEY_COLS[1])
            i_gain = col.get(RAW_KEY_COLS[2])
            i_slices = col.get(RAW_KEY_COLS[3])
            i_dig = col.get(RAW_KEY_COLS[4])
            i_freq = col.get(RAW_KEY_COLS[5])
            i_y = col.get(Y_COL)
            i_pattern = col.get("cfg-carrier_01-data_pattern")

            def cell(row, idx):
                return row[idx] if idx is not None and idx < len(row) else ""

            for row in reader:
                if i_pattern is not None and (cell(row, i_pattern) or "").strip() != DATA_PATTERN:
                    continue
                val = to_float(cell(row, i_y))
                if val is None:
                    continue
                packet = (cell(row, i_packet) or "").strip()
                supply = (cell(row, i_supply) or "").strip()
                gain = (cell(row, i_gain) or "").strip()
                pa_slices = to_float(cell(row, i_slices))
                dig_gain = to_float(cell(row, i_dig))
                freq = to_float(cell(row, i_freq))
                if pa_slices is None or dig_gain is None or freq is None:
                    continue
                pa_slices_seen.add(pa_slices)
                dig_gain_seen.add(dig_gain)
                key = (packet, supply, gain, pa_slices, dig_gain, freq)
                raw[key].append((val, dut))
        print(f"  read {path.name}")
    return raw, pa_slices_seen, dig_gain_seen


def build_fixed_combos(raw, pa_slices_max, dig_gain_max):
    """Filter to pa_slices==max & dig_gain==max, drop those two from the key."""
    data: dict[tuple, list[tuple[float, str]]] = defaultdict(list)
    for (packet, supply, gain, pa_slices, dig_gain, freq), values in raw.items():
        if pa_slices != pa_slices_max or dig_gain != dig_gain_max:
            continue
        combo = (packet, supply, gain, freq)
        data[combo].extend(values)
    return data


def find_combo_outliers(values):
    """IQR rule (Q1-IQR_MULT*IQR .. Q3+IQR_MULT*IQR) computed on a SINGLE combo's own
    ~40 DUT values (not pooled across pa_supply). Pooling across pa_supply
    first doesn't work: each pa_supply has its own distinct center, so the
    combined set is multi-modal and IQR on it clips the tails of otherwise
    normal, tightly-clustered per-supply data. Per-combo IQR only flags a
    genuine within-condition anomaly (e.g. one DUT spiking far from the
    other 39 units measured under the exact same setting).

    Returns (kept, outliers) where each is a list of (value, dut_id).
    """
    vals = np.array([v for v, _ in values])
    q1, q3 = np.percentile(vals, [25, 75])
    iqr = q3 - q1
    if iqr <= 0:
        return list(values), []
    lower = q1 - IQR_MULT * iqr
    upper = q3 + IQR_MULT * iqr
    kept = [(v, d) for v, d in values if lower <= v <= upper]
    outliers = [(v, d) for v, d in values if v < lower or v > upper]
    return kept, outliers


SPEC_MARGIN_STEP = 0.5


def _snap_range(vmin_raw, vmax_raw, step=SPEC_MARGIN_STEP):
    """Floor/ceil (vmin_raw, vmax_raw) to the nearest `step` grid, then push
    ONE more step out if a bound already sits exactly on that grid --
    2026-08-15 customer request: an LSL/USL is often already a round number
    (e.g. USL=16, LSL=13.5), so a naive floor/ceil leaves it exactly on the
    plotted axis edge where its dashed line + text label become invisible/
    clipped (confirmed real case: USL=16 with the span capped at 16). The
    fix: span becomes 13/16.5 for that example -- LSL/USL land 0.5 inside
    the edge, not on it. Also doubles as the "round the whole span to a
    nice grid" mechanism (e.g. FreqAcc's 142969/152669 -> 142000/153000,
    ACP's -52/1 -> -60/10) when called with a coarser step -- ordinary
    floor/ceil already gives clearance there since real data rarely lands
    exactly on a multiple of 1000 or 10."""
    vmin = math.floor(vmin_raw / step) * step
    vmax = math.ceil(vmax_raw / step) * step
    if vmin == vmin_raw:
        vmin -= step
    if vmax == vmax_raw:
        vmax += step
    return vmin, vmax


def build_packet_ranges(data):
    """Per packet condition (packet_type + pa_gain + BAND, i.e. every
    frequency within the same band, and every key dimension except
    pa_supply), combine pa_supply 0~2's outlier-cleaned values across ALL
    of that band's frequencies and find min/max, rounded to integer
    (max=ceil, min=floor). This makes every frequency's chart within the
    same band share one common Y-axis scale.

    Returns:
      ranges[(packet_type, gain, band)] = (int_min, int_max)
      outliers_by_combo[(packet_type, supply, gain, freq)] = [(value, dut_id), ...]
    """
    cleaned_by_combo: dict[tuple, list[tuple[float, str]]] = {}
    outliers_by_combo: dict[tuple, list[tuple[float, str]]] = {}
    for combo, values in data.items():
        kept, outliers = find_combo_outliers(values)
        cleaned_by_combo[combo] = kept
        outliers_by_combo[combo] = outliers

    per_group_values: dict[tuple, list[float]] = defaultdict(list)
    for (packet, supply, gain, freq), kept in cleaned_by_combo.items():
        per_group_values[(packet, gain, band_of(freq))].extend(v for v, _ in kept)

    ranges: dict[tuple, tuple[int, int]] = {}
    for group, values in per_group_values.items():
        ranges[group] = (math.floor(min(values)), math.ceil(max(values)))

    return ranges, outliers_by_combo


def combo_name(combo, pa_slices_max, dig_gain_max) -> str:
    packet, supply, gain, freq = combo
    return (
        f"packet_type({packet})_frequency-MHz({fmt_num(freq)})_pa_supply({supply})_"
        f"pa_slices({fmt_num(pa_slices_max)})_pa_gain({gain})_"
        f"dig_gain({fmt_num(dig_gain_max)})"
    )


def combo_title_short(combo, pa_slices_max, dig_gain_max) -> str:
    """Short plot-title format (2026-07-19 user request) -- the old
    combo_name()-based title (e.g. "BT_TX_packet_type(8DH5)_frequency-
    MHz(5152)_pa_supply(1)_pa_slices(7)_pa_gain(0)_dig_gain(128)") was so
    long that _fit_title_fontsize had to shrink it far below a readable
    size just to keep it inside the plot frame. Trims every wrapper/prefix
    down to the essentials: "8DH5_5152MHz_supply(1)_slices(7)_gain(0)_
    dig_gain(128)". Does NOT change the PNG/XLSX FILENAME (still
    combo_name(), since build_ppt_*.py's discover() regexes parse that
    exact verbose format) -- only the text drawn as the chart's own title."""
    packet, supply, gain, freq = combo
    return (f"{packet}_{fmt_num(freq)}MHz_HPA_{hpa_label(supply)}_"
            f"slices({fmt_num(pa_slices_max)})_gain({gain})_dig_gain({fmt_num(dig_gain_max)})")


def excluded_text(excluded) -> str:
    if not excluded:
        return ""
    lines = [f"{v:.3f} ({dut})" for v, dut in excluded]
    return "Excluded outlier(s):\n" + "\n".join(lines)


def _fit_title_fontsize(fig, ax, text, max_width_frac=0.98, start=18, min_size=7):
    """Largest bold fontsize (<=start) whose rendered width stays within the
    chart's own PLOT FRAME (the axes box itself, not the whole figure) --
    long auto-generated titles no longer spill past the frame (customer
    request, 2026-07-19: bold + as large as the frame allows). Measuring
    against the full figure width was wrong: the y-axis label/tick text
    eats into the left margin, so the actual plot frame is narrower than
    the figure -- caller must call fig.tight_layout() first so the axes'
    pixel position/width is already finalized when this measures it."""
    renderer = fig.canvas.get_renderer()
    max_width_px = ax.get_window_extent(renderer=renderer).width * max_width_frac
    for size in range(start, min_size - 1, -1):
        t = ax.text(0, 0, text, fontsize=size, fontweight="bold")
        width = t.get_window_extent(renderer=renderer).width
        t.remove()
        if width <= max_width_px:
            return size
    return min_size


# -- PPT grid geometry, mirrored from build_ppt_pwr_evm_revision2.py, so a
# source figure's aspect matches the actual PPT cell it lands in instead of
# being stretched into a different shape (confirmed real bug 2026-08-14: a
# fixed tall (9,14) figure stretched into a 1-2 column row -- wide, short
# cells, especially once ATE stacking halves the height -- smeared the
# title/axis text horizontally; same distortion class as the UWB/WL_TX
# "squish factor" bug this project already fixed elsewhere. Both PPT slide
# layouts (add_band24_slide's 2.4G row, add_highfreq_supply_slide's 5/6G
# row) place every column in ONE row at this same cell_width/image_h, so
# one figure_size_for(n_cols) covers both.
_SLIDE_WIDTH_IN = 13.338542213473316
_LEFT_MARGIN = 0.05
_RIGHT_MARGIN = 0.05
_RULER_WIDTH = 0.6
_CELL_GAP_H = 0.03
_GRID_WIDTH = _SLIDE_WIDTH_IN - _LEFT_MARGIN - _RIGHT_MARGIN - _RULER_WIDTH

_TABLE_TOP = 1.05
_TABLE_ROW_H = 0.16
_TABLE_PLOT_GAP = 0.35
_ATE_GAP = 0.1
_CONTENT_AREA_TOP = 1.5908508311461067
_CONTENT_AREA_BOTTOM = _CONTENT_AREA_TOP + 5.240927384076991
# header + Bench/ATE subheader + Mean+/-Std + Min/Max + USL + LSL -- Power's
# spec_constant=False always shows both rows (see build_ppt_full_revision2's
# get_power_spec comment). An approximation (a real slide's table_h can vary
# by a fraction of a row), close enough to eliminate the SEVERE aspect
# mismatch this is fixing -- not meant to be pixel-exact.
_TABLE_N_ROWS = 6


def figure_size_for(n_cols: int) -> tuple[float, float]:
    cell_width = _GRID_WIDTH / n_cols if n_cols == 1 else (_GRID_WIDTH - (n_cols - 1) * _CELL_GAP_H) / n_cols
    table_h = _TABLE_N_ROWS * _TABLE_ROW_H
    plots_top = _TABLE_TOP + table_h + _TABLE_PLOT_GAP
    image_h = (_CONTENT_AREA_BOTTOM - plots_top - _ATE_GAP) / 2
    return (cell_width, image_h)


_REF_FIGSIZE = (9, max(4, 0.3 * BINS + 2))  # the shape every font size below was tuned against


def _font_scale(figsize) -> float:
    """Every fontsize in this function (25/20/8/30/10pt) was tuned for
    _REF_FIGSIZE's tall 9x14 shape. figure_size_for() now hands back much
    shorter figures for narrow-column PPT rows (as low as ~2.2in tall for a
    1-column row) -- at a FIXED point size, the same tick-label text takes
    up a much bigger fraction of a shorter figure, and 25pt y-tick labels
    on a 2.2in-tall axis collide into unreadable overlapping mush (real bug
    found 2026-08-14, screenshot from the customer). Scale every font by
    the figure's actual height relative to the reference height it was
    tuned at, floored so text never shrinks past legibility."""
    if figsize is None:
        return 1.0
    return max(0.35, min(1.0, figsize[1] / _REF_FIGSIZE[1]))


def draw_png(values, title, xlabel, png_path: Path, value_range=None, excluded=None, spec=None, figsize=None) -> None:
    counts, edges = np.histogram(values, bins=BINS, range=value_range)
    fs = _font_scale(figsize)

    fig, ax = plt.subplots(figsize=figsize or _REF_FIGSIZE)
    max_count = max(counts.max(), 1)
    for i in range(BINS):
        lo, hi = edges[i], edges[i + 1]
        bin_h = hi - lo
        c = counts[i]
        ax.barh(lo, c, height=bin_h * 0.9, align="edge", color="#4472C4",
                edgecolor="black", linewidth=0.6)
        if c > 0:
            ax.text(c + max_count * 0.01, lo + bin_h / 2, f"{int(c)}", va="center", ha="left", fontsize=8 * fs)

    vmin, vmax = value_range if value_range else (edges[0], edges[-1])
    ax.set_ylim(vmin, vmax)
    # 2026-08-14 (customer request): cap Y-axis ticks at 8 with a nice
    # (whole-number) step, instead of the old unconditional step=1 -- Power's
    # range can now exceed 8 units once ATE widening (see make_ate_tx_pwr_
    # revision2.py's Y-range-widening block) pushes vmin/vmax past what a
    # single-unit-per-tick axis can show cleanly. Same min-gap boundary
    # guard as make_bt_tx_freqacc_revision2.py's identical block, so a
    # forced-in exact vmin/vmax never lands right next to a step tick.
    y_lo, y_hi = math.floor(vmin), math.ceil(vmax)
    span = max(1, y_hi - y_lo)
    step = max(1, math.ceil(span / 8))
    n_steps = max(1, round(span / step))
    yticks = sorted({int(round(y_lo + i * step)) for i in range(n_steps + 1)} | {y_lo, y_hi})
    max_ticks = 8
    if len(yticks) > max_ticks:
        min_gap = (y_hi - y_lo) / max_ticks * 0.5
        inner = [t for t in yticks if t not in (y_lo, y_hi)
                 and (t - y_lo) >= min_gap and (y_hi - t) >= min_gap]
        keep_n = max(0, max_ticks - 2)
        if keep_n > 0 and inner:
            idx_step = len(inner) / keep_n
            thinned = sorted({inner[min(len(inner) - 1, round(i * idx_step))] for i in range(keep_n)})
        else:
            thinned = []
        yticks = sorted({y_lo, y_hi} | set(thinned))
    ax.set_yticks(yticks)
    # Explicit X (Count) range -- 2026-08-14 real bug: ATE values sometimes
    # fall entirely outside value_range (the Bench-derived Y-window this ATE
    # chart reuses for comparability), so every np.histogram bin comes back
    # 0 and matplotlib's autoscale, with no real bar extent to anchor on,
    # picks an arbitrary tiny window around 0 -- MaxNLocator(integer=True)
    # then has no integers to place inside that sub-1-wide window and falls
    # back to garbled overlapping fractional tick labels. Anchoring xlim to
    # (0, max_count) unconditionally keeps the axis sane (0..1 when the
    # histogram is genuinely empty) regardless of why it's empty.
    ax.set_xlim(0, max_count * 1.08)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True))
    plt.setp(ax.get_yticklabels(), fontsize=25 * fs, fontweight="bold")
    plt.setp(ax.get_xticklabels(), fontsize=25 * fs, fontweight="bold")
    ax.set_ylabel(xlabel, fontsize=20 * fs, fontweight="bold")
    ax.set_xlabel("Count", fontsize=20 * fs, fontweight="bold")
    ax.grid(True, axis="both", color="0.85", linewidth=0.6)
    ax.set_axisbelow(True)
    # Finalize the axes' pixel position before sizing the title against it.
    fig.tight_layout()
    ax.set_title(title, fontsize=_fit_title_fontsize(fig, ax, title, start=int(18 * fs), min_size=int(max(5, 7 * fs))),
                 fontweight="bold")

    if spec is not None:
        lsl, usl = spec
        # NOTE: va="top" visually overlaps the axhline in tall/wide-range
        # figures (confirmed matplotlib rendering quirk, 2026-07-17) -- use
        # va="bottom" uniformly with a manual data-space pad for the
        # "place below the line" case instead.
        pad = (vmax - vmin) * 0.035
        for spec_val, spec_label in ((lsl, "LSL"), (usl, "USL")):
            if not (vmin <= spec_val <= vmax):
                continue  # outside this chart's own axis range -- nothing to draw
            # 2026-07-20 customer request: Limit line is blue dashed (was red
            # solid-looking/thin) -- red dashed is reserved for a "Target"
            # value if one is ever added, same linewidth as this Limit line.
            ax.axhline(y=spec_val, color="blue", linestyle="--", linewidth=2.5 * fs, zorder=5)
            above = spec_val >= (vmin + vmax) / 2
            text_y = spec_val + pad if above else spec_val - pad
            ax.text(ax.get_xlim()[1] * 0.99, text_y, f"{spec_label} {spec_val:g}",
                    color="blue", fontsize=10 * fs, fontweight="bold", ha="right",
                    va="bottom", clip_on=True)

    note = excluded_text(excluded)
    if note:
        ax.text(0.98, 0.02, note, transform=ax.transAxes, ha="right", va="bottom",
                 fontsize=30 * fs, fontweight="bold", color="#B00000",
                 bbox=dict(boxstyle="round", facecolor="#FFF2F2", edgecolor="#B00000"))

    fig.tight_layout()
    fig.savefig(png_path, dpi=450)
    plt.close(fig)


def write_xlsx(xlsx_path: Path, title: str, xlabel: str, values, value_range=None, excluded=None) -> None:
    counts, edges = np.histogram(values, bins=BINS, range=value_range)
    labels = [f"{edges[i]:.3f} ~ {edges[i+1]:.3f}" for i in range(BINS)]

    wb = Workbook()
    ws = wb.active
    ws.title = "Detail"
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    chart_row_span = 22
    data_header_row = 1 + chart_row_span

    ws.cell(data_header_row, 1, f"{xlabel} Bin")
    ws.cell(data_header_row, 2, "Count")
    for i, (label, count) in enumerate(zip(labels, counts), data_header_row + 1):
        ws.cell(i, 1, label)
        ws.cell(i, 2, int(count))
    for cell in ws[data_header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 10

    ws.cell(data_header_row, 4, "Raw Values")
    for i, v in enumerate(values, data_header_row + 1):
        ws.cell(i, 4, v)
    ws.column_dimensions["D"].width = 14

    if excluded:
        ws.cell(data_header_row, 6, "Excluded Outlier(s)")
        ws.cell(data_header_row, 6).font = Font(bold=True, color="B00000")
        for i, (v, dut) in enumerate(excluded, data_header_row + 1):
            ws.cell(i, 6, f"{v:.3f} ({dut})")
        ws.column_dimensions["F"].width = 22

    last_row = data_header_row + len(counts)
    chart = BarChart()
    chart.type = "bar"  # horizontal
    chart.title = title
    chart.style = 10
    chart.height = 12
    chart.width = 26
    chart.x_axis.title = f"{xlabel} Bin"
    chart.y_axis.title = "Count"
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.y_axis.majorGridlines = ChartLines()
    chart.y_axis.majorGridlines.spPr = GraphicalProperties()
    chart.y_axis.majorGridlines.graphicalProperties.line.solidFill = "808080"
    chart.y_axis.majorGridlines.graphicalProperties.line.width = 12700
    chart.legend = None
    chart.gapWidth = 0

    data_ref = Reference(ws, min_col=2, min_row=data_header_row, max_row=last_row)
    cats_ref = Reference(ws, min_col=1, min_row=data_header_row + 1, max_row=last_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    series = chart.series[-1]
    series.graphicalProperties.solidFill = "4472C4"
    ws.add_chart(chart, "A1")

    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--bt-tx-dir", type=Path, default=PATHS.bench_dir)
    p.add_argument("--out-dir", type=Path, default=PATHS.result_xlsx_dir / "pwr_revision2")
    p.add_argument("--png-dir", type=Path, default=PATHS.result_png_dir / "pwr_revision2")
    p.add_argument("--limit", type=int, default=0,
                   help="Generate only the first N combos (0 = all).")
    p.add_argument("--count-only", action="store_true",
                   help="Only print combo/sample-size audit, generate nothing.")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    print("Reading BT_Tx CSVs...")
    raw, pa_slices_seen, dig_gain_seen = read_all(args.bt_tx_dir)
    pa_slices_max = max(pa_slices_seen)
    dig_gain_max = max(dig_gain_seen)
    print(f"pa_slices seen: {sorted(pa_slices_seen)} -> fixed at max {fmt_num(pa_slices_max)}")
    print(f"dig_gain seen: {sorted(dig_gain_seen)} -> fixed at max {fmt_num(dig_gain_max)}")

    data = build_fixed_combos(raw, pa_slices_max, dig_gain_max)
    combos = list(data.keys())
    print(f"combos (packet,supply,gain,frequency): {len(combos)}")

    packet_ranges, outliers_by_combo = build_packet_ranges(data)
    n_outlier_combos = sum(1 for v in outliers_by_combo.values() if v)
    n_outlier_points = sum(len(v) for v in outliers_by_combo.values())
    print(f"\n=== per packet-condition Y-axis scale (pa_supply 0~2 combined, floor/ceil): {len(packet_ranges)} groups ===")
    print(f"    IQR outlier points excluded from scale: {n_outlier_points} (across {n_outlier_combos} combo(s))")

    spec_workbook = PATHS.bt_tx_spec_workbook or PATHS.spec_workbook
    print(f"Parsing TX Power Specification (LSL/USL) from {spec_workbook.name}...")
    spec_table = spec_lookup.parse_spec(spec_workbook)
    n_spec_hits = sum(
        1 for combo in combos
        if spec_lookup.lookup(spec_table, combo[0], combo[3], combo[1]) is not None
    )
    print(f"    {n_spec_hits}/{len(combos)} combos have a matching LSL/USL spec")

    # 2026-08-15 (customer request): Bench and ATE charts must share the
    # EXACT same Y-axis span for direct top/bottom comparison. Lazy/
    # deferred import (not at module top-level) to avoid a circular import
    # -- make_ate_tx_pwr_revision2 imports this module as pwr_mod; by the
    # time main() runs, this module is already fully initialized, so
    # Python just returns it from sys.modules. get_ate_group_extents() is
    # called from BOTH this main() and the ATE module's own draw_all(), so
    # both compute the identical widened range independently from the same
    # two raw data sources -- no file-based coordination needed.
    import make_ate_tx_pwr_revision2 as ate_pwr_mod
    ate_extents = ate_pwr_mod.get_ate_group_extents(data, pa_slices_max, dig_gain_max)
    for group, (lo, hi) in ate_extents.items():
        if group in packet_ranges:
            vmin, vmax = packet_ranges[group]
            packet_ranges[group] = (min(vmin, lo), max(vmax, hi))

    # Fold every combo's own LSL/USL into its group's span (so the Limit
    # line is never outside the chart), then snap the whole span to the
    # nice 0.5 grid with clearance -- see _snap_range's docstring.
    for group in list(packet_ranges.keys()):
        packet, gain, band = group
        vmin_raw, vmax_raw = packet_ranges[group]
        for p, supply, g, freq in combos:
            if p != packet or g != gain or band_of(freq) != band:
                continue
            spec = spec_lookup.lookup(spec_table, packet, freq, supply)
            if spec is None:
                continue
            lsl, usl = spec
            if lsl is not None:
                vmin_raw = min(vmin_raw, lsl)
            if usl is not None:
                vmax_raw = max(vmax_raw, usl)
        packet_ranges[group] = _snap_range(vmin_raw, vmax_raw)

    if args.count_only:
        sizes = defaultdict(int)
        anomalies = []
        for combo, values in data.items():
            n = len(values)
            sizes[n] += 1
            if n % 40 != 0:
                anomalies.append((combo, n))
        print("\n=== sample-size audit (expected = 40 per combo) ===")
        print("  sample sizes seen:", dict(sorted(sizes.items())))
        print(f"\nnon-multiple-of-40 anomalies: {len(anomalies)}")
        for combo, n in anomalies[:30]:
            print(f"  {combo} n={n}")
        if len(anomalies) > 30:
            print(f"  ... and {len(anomalies) - 30} more")
        return

    if args.limit > 0:
        combos = combos[: args.limit]
        print(f"Limiting to first {len(combos)} combo(s)")

    args.out_dir.mkdir(parents=True, exist_ok=True)
    args.png_dir.mkdir(parents=True, exist_ok=True)

    # Precompute each combo's PPT column count -- mirrors build_ppt_pwr_evm_
    # revision2.py's discover()/add_band24_slide()/add_highfreq_supply_slide()
    # grouping exactly, so figure_size_for() matches the actual cell each
    # PNG lands in (see figure_size_for's docstring).
    band24_freqs = defaultdict(set)   # (packet, supply) -> {freq, ...}, band 2.4 only
    highfreq_freqs = defaultdict(set)  # (packet, band, supply) -> {freq, ...}, band 5/6
    for packet, supply, _gain, freq in combos:
        if band_of(freq) == 2.4:
            band24_freqs[(packet, supply)].add(freq)
        else:
            highfreq_freqs[(packet, band_of(freq), supply)].add(freq)
    band24_supplies = defaultdict(set)
    for (packet, supply) in band24_freqs:
        band24_supplies[packet].add(supply)
    band24_n_cols = {
        packet: max(len(band24_freqs[(packet, s)]) for s in supplies) * len(supplies)
        for packet, supplies in band24_supplies.items()
    }

    def n_cols_for(packet, supply, freq):
        if band_of(freq) == 2.4:
            return band24_n_cols[packet]
        return len(highfreq_freqs[(packet, band_of(freq), supply)])

    jobs = []
    for combo in combos:
        packet, supply, gain, freq = combo
        name = combo_name(combo, pa_slices_max, dig_gain_max)
        fname = safe_filename(name)
        title = combo_title_short(combo, pa_slices_max, dig_gain_max)
        values = [v for v, _dut in data[combo]]
        value_range = packet_ranges[(packet, gain, band_of(freq))]
        excluded = outliers_by_combo.get(combo)
        spec = spec_lookup.lookup(spec_table, packet, freq, combo[1])
        figsize = figure_size_for(n_cols_for(packet, supply, freq))
        jobs.append((values, title, args.png_dir / f"{fname}.png", value_range, excluded, spec, figsize,
                     args.out_dir / f"{fname}.xlsx", fname))

    n = 0
    with ProcessPoolExecutor(max_workers=PATHS.png_workers) as pool:
        for _ in pool.map(_draw_one, jobs, chunksize=4):
            n += 1
            if n % 200 == 0 or n == len(jobs):
                print(f"  [{n}/{len(jobs)}]")

    print(f"\nDone. XLSX -> {args.out_dir}")
    print(f"      PNG  -> {args.png_dir}")


def _draw_one(job) -> None:
    """Worker for ProcessPoolExecutor -- must be a module-level function so
    it's picklable on Windows (spawn-only, no fork)."""
    values, title, png_path, value_range, excluded, spec, figsize, xlsx_path, _fname = job
    draw_png(values, title, "Power (dBm)", png_path, value_range, excluded, spec, figsize)
    if PATHS.write_xlsx:
        write_xlsx(xlsx_path, title, "Power (dBm)", values, value_range, excluded)


if __name__ == "__main__":
    main()
