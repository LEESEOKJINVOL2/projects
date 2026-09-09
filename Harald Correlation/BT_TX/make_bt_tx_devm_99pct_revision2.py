"""BT TX DEVM99pct revision2 distribution histograms.

Sibling script to make_bt_tx_devm_rms_revision2.py -- identical pipeline,
different data column (evm_99pct_aver-%). New metric added 2026-07-19,
covering the same 13-packet DEVM breadth as DEVM RMS/Peak.

For every (packet_type, pa_supply, pa_gain, frequency-MHz) combination,
collect every evm_99pct_aver-% value from every DUT and draw a horizontal
bar-chart histogram:
  X axis = Count
  Y axis = DEVM99pct value bins

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
from openpyxl.styles import Alignment, Font, PatternFill

import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")


RAW_KEY_COLS = [
    "cfg-carrier_01-packet_type",
    "cfg-carrier_01-pa_supply",
    "cfg-carrier_01-pa_gain",
    "cfg-carrier_01-pa_slices",
    "cfg-carrier_01-dig_gain",
    "cfg-carrier_01-frequency-MHz",
]
Y_COL = "evm_99pct_aver-%"
XLABEL = "DEVM99pct (%)"
SPEC_METRIC = "99pct"
# 2026-08-15 real bug -- see make_bt_tx_pwr_revision2.py's identical
# comment for the full rationale: cfg-carrier_01-data_pattern was missing
# from RAW_KEY_COLS, silently pooling 3 data-pattern sweeps (PRBS9/PAT1/
# PAT2) for 1DH5/LE1M/LE2M into one histogram (tripling the real sample
# count) while ATE only ever measures PRBS9.
DATA_PATTERN = "PRBS9"
BINS = 40
DUT_RE = re.compile(r"#(\d+)")
IQR_MULT = 5.0
# If a group's biggest excluded-outlier value is below this, extend the axis
# to include it (scatter it instead of excluding it, 2026-07-18). At/above
# this the outlier is too extreme -- keep the old exclude+annotate behavior
# to avoid distorting the chart (e.g. HDT8_pa_supply0 hits this exception).
OUTLIER_SCATTER_MAX_THRESHOLD = 5


# 2026-08-14: 5G/6G merged into one "Band5G" bucket -- see make_bt_tx_pwr_
# revision2.band_of's identical comment for the full rationale.
def band_of(freq: float) -> float:
    if freq < 3000:
        return 2.4
    return 5


# See make_bt_tx_pwr_revision2.SUPPLY_TO_HPA/hpa_label for the full
# rationale (customer-confirmed voltage mapping + title-only scope).
SUPPLY_TO_HPA = {"0": "1.5V", "1": "1.2V", "2": "0.73V"}


def hpa_label(supply) -> str:
    return SUPPLY_TO_HPA.get(str(supply), f"supply({supply})")


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

    2026-07-20: csv.DictReader -> csv.reader + one-time header->index map
    (only ~7 of many columns are used); @lru_cache so a process that calls
    this twice (e.g. a build script consuming both this module and its
    ATE-sibling module) only reads the 40 files once -- safe, nothing
    downstream mutates the returned raw dict/sets."""
    raw: dict[tuple, list[tuple[float, str]]] = defaultdict(list)
    pa_slices_seen: set[float] = set()
    dig_gain_seen: set[float] = set()

    # 2026-08-25: rglob + suffix filter, not a flat glob -- see
    # make_bt_tx_pwr_revision2.py's read_all() for why (a newer Bench pull
    # nests each DUT's file under #NN_<serial>/, alongside an unrelated
    # summary/summary_<timestamp>.csv that a blanket "*.csv" would wrongly
    # sweep in too).
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


def group_of(packet, gain, freq, supply):
    """The axis-sharing unit ("slide"): 2.4G merges all 3 supplies into one
    slide (so its group ignores supply, as before); 5G/6G puts each supply
    on its own slide, so its group MUST include supply."""
    band = band_of(freq)
    if band == 2.4:
        return (packet, gain, 2.4)
    return (packet, gain, band, supply)


def nice_step(value, target_ticks=10):
    """'Nice' gridline step (1/2/5 x 10^k) so a 0..value axis gets roughly
    target_ticks major divisions -- e.g. value=20 -> step=2, value=35 ->
    step=5 (confirmed against the customer's own EDR2 DEVM RMS/Peak
    examples, 2026-07-19: "0-21% line, major step 2%, limit line at 20%"
    and "0-36% line, major step 5%, limit line at 35%")."""
    if value <= 0:
        return 1
    raw_step = value / target_ticks
    # 2026-08-14 (customer request): never drop below a whole-number step --
    # see make_bt_tx_freqacc_revision2.nice_step's identical comment for the
    # full rationale.
    magnitude = max(1, 10 ** math.floor(math.log10(raw_step)))
    residual = raw_step / magnitude
    if residual <= 1:
        nice = 1
    elif residual <= 2:
        nice = 2
    elif residual <= 5:
        nice = 5
    else:
        nice = 10
    return nice * magnitude


def spec_window(spec_max, target_ticks=10):
    """(window_min, window_max, step) for a spec-anchored axis: 0 to just
    above the ceil-to-step spec value (a small fixed +1 margin so the
    limit line isn't glued to the plot's top edge) -- matches the
    customer's own EDR2 examples (limit 20 -> window 0-21 step 2; limit
    35 -> window 0-36 step 5)."""
    step = nice_step(spec_max, target_ticks)
    # round first -- floating-point noise (e.g. 0.07*100 == 7.000000000000001)
    # otherwise pushes ceil() up an unwanted extra step.
    ceil_to_step = math.ceil(round(spec_max / step, 6)) * step
    return 0, ceil_to_step + 1, step


def build_packet_ranges(data, spec_by_packet=None):
    """Per axis-sharing group (see group_of):

    Packets WITH a DEVM spec (2026-07-19, customer request -- "open up the
    window to see the trends", e.g. EDR2 DEVM RMS 0-21% by 2%, limit at
    20%): the Y-axis is a FIXED window anchored on the spec value (see
    spec_window()), the SAME for every combo of that packet regardless of
    band/supply -- not data-driven. Any DUT value landing outside that
    fixed window is a genuine anomaly: excluded and text-annotated (never
    folded back in -- the window is already generous by design, so a value
    that still doesn't fit is worth flagging on its own, per the
    customer's "values that deviate significantly...shown separately as
    outliers").

    Packets with NO spec keep the original data-driven design unchanged:
    combine pa_supply-appropriate outlier-cleaned values, floor/ceil to
    integer, then extend-and-fold-in or exclude-and-annotate based on
    (biggest outlier - base_vmax) vs OUTLIER_SCATTER_MAX_THRESHOLD.

    Returns:
      ranges[group] = (min, max) -- max may now be a non-integer for
        spec-anchored groups (e.g. window_max = 21.0), unlike the old
        integer-only floor/ceil ranges.
      kept_by_combo[(packet_type, supply, gain, freq)] = [(value, dut_id), ...]
      outliers_by_combo[(packet_type, supply, gain, freq)] = [(value, dut_id), ...]
      include_outliers_group[group] = bool -- always False for spec-anchored
        groups; for non-spec groups, True means fold outliers back into
        the histogram for every combo in this group.
      step_by_group[group] = gridline step (1 for non-spec groups, the
        spec_window()-derived step for spec-anchored ones).
    """
    spec_by_packet = spec_by_packet or {}
    spec_window_by_packet = {p: spec_window(v) for p, v in spec_by_packet.items() if v is not None}

    kept_by_combo: dict[tuple, list[tuple[float, str]]] = {}
    outliers_by_combo: dict[tuple, list[tuple[float, str]]] = {}
    for combo, values in data.items():
        packet = combo[0]
        if packet in spec_window_by_packet:
            vmin, vmax, _step = spec_window_by_packet[packet]
            kept = [(v, d) for v, d in values if vmin <= v <= vmax]
            outliers = [(v, d) for v, d in values if not (vmin <= v <= vmax)]
        else:
            kept, outliers = find_combo_outliers(values)
        kept_by_combo[combo] = kept
        outliers_by_combo[combo] = outliers

    per_group_kept: dict[tuple, list[float]] = defaultdict(list)
    per_group_outliers: dict[tuple, list[float]] = defaultdict(list)
    for (packet, supply, gain, freq), kept in kept_by_combo.items():
        if packet in spec_window_by_packet:
            continue
        per_group_kept[group_of(packet, gain, freq, supply)].extend(v for v, _ in kept)
    for (packet, supply, gain, freq), outliers in outliers_by_combo.items():
        if packet in spec_window_by_packet:
            continue
        if outliers:
            per_group_outliers[group_of(packet, gain, freq, supply)].extend(v for v, _ in outliers)

    ranges: dict[tuple, tuple[float, float]] = {}
    include_outliers_group: dict[tuple, bool] = {}
    step_by_group: dict[tuple, float] = {}
    for group, values in per_group_kept.items():
        base_vmin, base_vmax = math.floor(min(values)), math.ceil(max(values))
        group_outliers = per_group_outliers.get(group)
        if not group_outliers:
            ranges[group] = (base_vmin, base_vmax)
            include_outliers_group[group] = False
        else:
            excl_max = max(group_outliers)
            if (excl_max - base_vmax) >= OUTLIER_SCATTER_MAX_THRESHOLD:
                ranges[group] = (base_vmin, base_vmax)
                include_outliers_group[group] = False
            else:
                ranges[group] = (base_vmin, max(base_vmax, math.ceil(excl_max) + 1))
                include_outliers_group[group] = True
        step_by_group[group] = 1

    for combo in data:
        packet, supply, gain, freq = combo
        if packet not in spec_window_by_packet:
            continue
        group = group_of(packet, gain, freq, supply)
        if group in ranges:
            continue
        vmin, vmax, step = spec_window_by_packet[packet]
        ranges[group] = (vmin, vmax)
        include_outliers_group[group] = False
        step_by_group[group] = step

    return ranges, kept_by_combo, outliers_by_combo, include_outliers_group, step_by_group

    return ranges, kept_by_combo, outliers_by_combo, include_outliers_group


def combo_name(combo, pa_slices_max, dig_gain_max) -> str:
    packet, supply, gain, freq = combo
    return (
        f"packet_type({packet})_frequency-MHz({fmt_num(freq)})_pa_supply({supply})_"
        f"pa_slices({fmt_num(pa_slices_max)})_pa_gain({gain})_"
        f"dig_gain({fmt_num(dig_gain_max)})"
    )


def combo_title_short(combo, pa_slices_max, dig_gain_max) -> str:
    """Short plot-title format (2026-07-19 user request), same as
    make_bt_tx_pwr_revision2.combo_title_short: "8DH5_5152MHz_supply(1)_
    slices(7)_gain(0)_dig_gain(128)". Does NOT change the PNG/XLSX
    FILENAME (still combo_name(), since build_ppt_*.py's discover() regex
    parses that exact verbose format) -- only the chart's own title text."""
    packet, supply, gain, freq = combo
    return (f"{packet}_{fmt_num(freq)}MHz_HPA_{hpa_label(supply)}_"
            f"slices({fmt_num(pa_slices_max)})_gain({gain})_dig_gain({fmt_num(dig_gain_max)})")


def excluded_text(excluded) -> str:
    if not excluded:
        return ""
    lines = [f"{v:.3f} ({dut})" for v, dut in excluded]
    return "Excluded outlier(s):\n" + "\n".join(lines)


def _place_excluded_note(ax, edges, counts, note, fontsize=28) -> None:
    """Place the excluded-outlier annotation INSIDE the axes, in a bar-free
    gap so it never covers real data -- moving it outside the axes changes
    each chart's subplot margins individually, which throws off alignment
    against sibling charts once every image in a PPT row gets stretched to
    the same fixed cell size.

    Finds the largest contiguous run of zero-count bins and centers the
    note there. Falls back to the top-right corner (may then overlap a bar)
    only if no run is long enough to hold the note -- i.e. the histogram
    has no real gap anywhere."""
    n_lines = note.count("\n") + 1
    # required_bins scales with fontsize (baseline 14pt ~ 1 bin/line) so a
    # bigger font still gets a gap that's actually large enough to hold it.
    required_bins = math.ceil(n_lines * fontsize / 14) + 1
    runs = []
    i, n = 0, len(counts)
    while i < n:
        if counts[i] == 0:
            j = i
            while j < n and counts[j] == 0:
                j += 1
            runs.append((i, j))
            i = j
        else:
            i += 1
    big_enough = [r for r in runs if (r[1] - r[0]) >= required_bins]
    xlim = ax.get_xlim()
    if big_enough:
        start, end = max(big_enough, key=lambda r: r[1] - r[0])
        y_center = (edges[start] + edges[end]) / 2
        ax.text(xlim[1] * 0.98, y_center, note, ha="right", va="center",
                 fontsize=fontsize, fontweight="bold", color="#B00000",
                 bbox=dict(boxstyle="round", facecolor="#FFF2F2", edgecolor="#B00000"))
    else:
        ax.text(xlim[1] * 0.98, edges[-1], note, ha="right", va="top",
                 fontsize=fontsize, fontweight="bold", color="#B00000",
                 bbox=dict(boxstyle="round", facecolor="#FFF2F2", edgecolor="#B00000"))


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


# -- PPT grid geometry, mirrored from build_ppt_pwr_evm_revision2.py, see
# make_bt_tx_pwr_revision2.figure_size_for's docstring for the full
# rationale (matches source figure aspect to the actual PPT cell). --
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
# header + Bench/ATE subheader + Mean+/-Std + Min/Max + USL -- an
# approximation, see make_bt_tx_pwr_revision2.py's _TABLE_N_ROWS note.
_TABLE_N_ROWS = 5


def figure_size_for(n_cols: int) -> tuple[float, float]:
    cell_width = _GRID_WIDTH / n_cols if n_cols == 1 else (_GRID_WIDTH - (n_cols - 1) * _CELL_GAP_H) / n_cols
    table_h = _TABLE_N_ROWS * _TABLE_ROW_H
    plots_top = _TABLE_TOP + table_h + _TABLE_PLOT_GAP
    image_h = (_CONTENT_AREA_BOTTOM - plots_top - _ATE_GAP) / 2
    return (cell_width, image_h)


_REF_FIGSIZE = (9, max(4, 0.3 * BINS + 2))


def _font_scale(figsize) -> float:
    """See make_bt_tx_pwr_revision2._font_scale's docstring."""
    if figsize is None:
        return 1.0
    return max(0.35, min(1.0, figsize[1] / _REF_FIGSIZE[1]))


def draw_png(values, title, xlabel, png_path: Path, value_range=None, excluded=None, spec=None,
             step=1, figsize=None) -> None:
    """`values` is the data to histogram. For a combo whose group applies the
    outlier-fold-in rule, the caller already folds the outlier value(s) back
    into `values` (plotted as an ordinary bar, no special marker) and passes
    excluded=None; otherwise `values` is kept-only and `excluded` carries
    the old text-annotation outliers -- see build_packet_ranges / main().
    `spec` is (lsl, usl) from bt_tx_devm_spec_lookup, either side possibly
    None; only drawn if it falls within this chart's own axis range.
    `step` is the gridline step (2026-07-19: variable for spec-anchored
    windows, e.g. 2 or 5, vs the old fixed 1 for data-driven windows)."""
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
    # Ticks at vmin, vmin+step, ... up to vmax, always including the exact
    # boundary values even if they don't land on a step multiple (avoids a
    # PPT ruler Min/Max mismatch -- same convention as the fixed step=1 case).
    n_steps = int(round((vmax - vmin) / step))
    yticks = sorted({round(vmin + i * step, 6) for i in range(n_steps + 1)} | {round(vmin, 6), round(vmax, 6)})
    # See make_bt_tx_freqacc_revision2.py's identical block for the full
    # rationale -- thin the tick SET (not the step) down to whatever fits
    # this figure's actual (possibly font-scaled-short) height.
    y_lo_t, y_hi_t = round(vmin, 6), round(vmax, 6)
    # 2026-08-14 (customer request): cap at 8 regardless of figure height --
    # even a tall figure's "up to 15" budget read as too dense.
    max_ticks = max(4, min(8, round(15 * fs)))
    if len(yticks) > max_ticks:
        # Exclude candidates too close to a forced boundary BEFORE thinning
        # picks from them -- otherwise thinning's index-0 pick can be the
        # step-multiple immediately next to vmin/vmax, visually overlapping
        # it (2026-08-14: hit when a widened ATE axis, see make_ate_tx_pwr_
        # revision2.py's Y-range-widening block, puts vmin/vmax off the
        # regular step grid).
        min_gap = (y_hi_t - y_lo_t) / max_ticks * 0.5
        inner = [t for t in yticks if t not in (y_lo_t, y_hi_t)
                 and (t - y_lo_t) >= min_gap and (y_hi_t - t) >= min_gap]
        keep_n = max(0, max_ticks - 2)
        if keep_n > 0 and inner:
            idx_step = len(inner) / keep_n
            thinned = sorted({inner[min(len(inner) - 1, round(i * idx_step))] for i in range(keep_n)})
        else:
            thinned = []
        yticks = sorted({y_lo_t, y_hi_t} | set(thinned))
    ax.set_yticks(yticks)
    # See make_bt_tx_pwr_revision2.py's identical line for the full
    # rationale -- an out-of-range ATE histogram (all bins 0) otherwise
    # leaves matplotlib to autoscale X to a garbled sub-1-wide window.
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
        spec_min, spec_max = spec
        pad = (vmax - vmin) * 0.035
        for spec_val, spec_label in ((spec_min, "LSL"), (spec_max, "USL")):
            if spec_val is None or not (vmin <= spec_val <= vmax):
                continue
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
        _place_excluded_note(ax, edges, counts, note, fontsize=28 * fs)

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
    p.add_argument("--out-dir", type=Path, default=PATHS.result_xlsx_dir / "devm_99pct_revision2")
    p.add_argument("--png-dir", type=Path, default=PATHS.result_png_dir / "devm_99pct_revision2")
    p.add_argument("--limit", type=int, default=0,
                   help="Generate only the first N combos (0 = all).")
    p.add_argument("--count-only", action="store_true",
                   help="Only print combo/sample-size audit, generate nothing.")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    print("Parsing DEVM spec limits...")
    import bt_tx_devm_spec_lookup as spec_lookup
    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    print("Reading BT_Tx CSVs...")
    raw, pa_slices_seen, dig_gain_seen = read_all(args.bt_tx_dir)
    pa_slices_max = max(pa_slices_seen)
    dig_gain_max = max(dig_gain_seen)
    print(f"pa_slices seen: {sorted(pa_slices_seen)} -> fixed at max {fmt_num(pa_slices_max)}")
    print(f"dig_gain seen: {sorted(dig_gain_seen)} -> fixed at max {fmt_num(dig_gain_max)}")

    data = build_fixed_combos(raw, pa_slices_max, dig_gain_max)
    combos = list(data.keys())
    print(f"combos (packet,supply,gain,frequency): {len(combos)}")

    spec_by_packet = spec_table.get(SPEC_METRIC, {})
    packet_ranges, kept_by_combo, outliers_by_combo, include_outliers_group, step_by_group = build_packet_ranges(
        data, spec_by_packet)

    # 2026-08-14 (customer request): Bench and ATE must share one Y-axis
    # span. Lazy/deferred import (not at module top-level) to avoid a
    # circular import -- make_ate_tx_devm_99pct_revision2 imports this
    # module as devm_mod; by the time main() runs this module is fully
    # initialized, so Python just returns the already-loaded module. See
    # get_ate_group_extents_if_new_format's docstring for why this is safe
    # to union (new item format only, never the legacy 99-100 pass-score
    # one) and why computing it independently here matches what the ATE
    # script computes for itself -- same two raw data sources, same result.
    import make_ate_tx_devm_99pct_revision2 as ate99_mod
    ate_extents = ate99_mod.get_ate_group_extents_if_new_format(data, pa_slices_max, dig_gain_max)
    for group, (ate_lo, ate_hi) in ate_extents.items():
        if group in packet_ranges:
            vmin, vmax = packet_ranges[group]
            packet_ranges[group] = (min(vmin, math.floor(ate_lo)), max(vmax, math.ceil(ate_hi)))
    n_outlier_combos = sum(1 for v in outliers_by_combo.values() if v)
    n_outlier_points = sum(len(v) for v in outliers_by_combo.values())
    n_extended_groups = sum(1 for v in include_outliers_group.values() if v)
    n_spec_anchored_groups = sum(1 for g in step_by_group if g[0] in spec_by_packet)
    print(f"\n=== per axis-sharing-group Y-axis scale: {len(packet_ranges)} groups ===")
    print(f"    Spec-anchored (fixed 0..spec+margin) groups: {n_spec_anchored_groups}")
    print(f"    IQR outlier points found: {n_outlier_points} (across {n_outlier_combos} combo(s))")
    print(f"    Non-spec groups with axis extended + outliers folded into the histogram "
          f"((Excluded-Outlier-Max - base axis Max) < {OUTLIER_SCATTER_MAX_THRESHOLD}): {n_extended_groups}")

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

    # Precompute each combo's PPT column count -- see make_bt_tx_pwr_
    # revision2.py's identical block for the full rationale.
    band24_freqs = defaultdict(set)
    highfreq_freqs = defaultdict(set)
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
        all_values = [v for v, _dut in data[combo]]
        kept_values = [v for v, _dut in kept_by_combo[combo]]
        group = group_of(packet, gain, freq, supply)
        value_range = packet_ranges[group]
        step = step_by_group[group]
        combo_outliers = outliers_by_combo.get(combo) or None
        if include_outliers_group.get(group) and combo_outliers:
            hist_values, excluded = all_values, None
        else:
            hist_values, excluded = kept_values, combo_outliers
        spec = spec_lookup.lookup(spec_table, SPEC_METRIC, packet)
        figsize = figure_size_for(n_cols_for(packet, supply, freq))
        jobs.append((hist_values, title, args.png_dir / f"{fname}.png", value_range, excluded, spec, step,
                     figsize, all_values, args.out_dir / f"{fname}.xlsx"))

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
    hist_values, title, png_path, value_range, excluded, spec, step, figsize, all_values, xlsx_path = job
    draw_png(hist_values, title, XLABEL, png_path, value_range, excluded, spec, step, figsize)
    if PATHS.write_xlsx:
        write_xlsx(xlsx_path, title, XLABEL, all_values, value_range, excluded)


if __name__ == "__main__":
    main()
