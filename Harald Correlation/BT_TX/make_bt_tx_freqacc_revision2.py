"""BT TX Frequency-Accuracy / Modulation-Accuracy metric distributions, revision2.

New metrics requested by the customer on top of the existing Power/EVM/ACP
work, split by packet_type category (each category needs a different metric
set, and HDT's "drift rate" is a DIFFERENT underlying column than BDR/BLE's
"drift rate" despite the same display name):

  BDR (1DH5) / BLE (LE1M, LE2M):
    ICFT             initial_frequency_offset_aver-Hz
    CFO              freq_offset_aver-Hz
    Drift rate       max_freq_drift_rate_aver-Hz
    Df1 average      delta_f1_avg_aver-Hz
    Df2 average      delta_f2_avg_aver-Hz
    Df2 max          delta_f2_max_access_aver-Hz

  EDR (3DH5) / HDR (4DH5, 8DH5) -- same metric set, packets kept distinct:
    omega_i          omega_i_aver-Hz
    omega_o          omega_0_aver-Hz
    omega_i+omega_o  omega_max_i0_aver-Hz
    DEVM peak        evm_peak_aver-%

  HDT (HDT3, HDT4, HDT8):
    ICFT             initial_frequency_offset_aver-Hz  (shared column w/ BDR/BLE)
    CFO              freq_offset_aver-Hz                (shared column w/ BDR/BLE)
    Drift rate       frequency_drift_aver-Hz             (DIFFERENT column!)

Out of scope (per customer, 2026-07-17): HDRPL16/32, HDRPM16/8, HDRPS2,
UHDR32/48 -- these packet_types get none of the new metrics.

Otherwise identical pipeline to make_bt_tx_pwr_revision2.py: grouping key
(packet_type, pa_supply, pa_gain, frequency-MHz), pa_slices/dig_gain fixed
at dataset-wide max, Y-axis (value bins) fixed per (packet_type, gain, band)
combining pa_supply 0~2, floor/ceil to integer. No value is ever excluded
from the axis (2026-07-20 customer correction) -- every FreqAcc metric is a
frequency-unit (Hz) value, so the axis always widens to cover the group's
full data extent instead of dropping outliers; this does not apply to
Power/ACP/DEVM. band_of(freq): <3000->2.4, <5900->5, else->6.

Output folder per metric: Result_XLSX/PNG/BT_TX/{metric_key}_revision2/
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
BINS = 40
DUT_RE = re.compile(r"#(\d+)")
# 2026-08-15 real bug -- see make_bt_tx_pwr_revision2.py's identical
# comment for the full rationale: cfg-carrier_01-data_pattern was missing
# from RAW_KEY_COLS, silently pooling 3 data-pattern sweeps (PRBS9/PAT1/
# PAT2) for 1DH5/LE1M/LE2M into one histogram (tripling the real sample
# count) while ATE only ever measures PRBS9.
DATA_PATTERN = "PRBS9"

# key, column, xlabel, applicable packet_types
METRICS = [
    # HDT6 added 2026-08-14 -- a real Bench packet_type this list never
    # covered (chosen against an earlier pull that had HDT3/HDT4 instead);
    # bt_tx_freqacc_spec_lookup.py now has HDT6's own icft spec row too.
    ("icft", "initial_frequency_offset_aver-Hz", "ICFT (Hz)",
     {"1DH5", "LE1M", "LE2M", "HDT3", "HDT4", "HDT6", "HDT8"}),
    ("cfo", "freq_offset_aver-Hz", "CFO (Hz)",
     {"1DH5", "LE1M", "LE2M", "HDT3", "HDT4", "HDT6", "HDT8"}),
    ("drift_rate", "max_freq_drift_rate_aver-Hz", "Drift Rate (Hz)",
     {"1DH5", "LE1M", "LE2M"}),
    ("df1_avg", "delta_f1_avg_aver-Hz", "Df1 Average (Hz)",
     {"1DH5", "LE1M", "LE2M"}),
    ("df2_avg", "delta_f2_avg_aver-Hz", "Df2 Average (Hz)",
     {"1DH5", "LE1M", "LE2M"}),
    ("df2_max", "delta_f2_max_access_aver-Hz", "Df2 Max (Hz)",
     {"1DH5", "LE1M", "LE2M"}),
    ("omega_i", "omega_i_aver-Hz", "ωi (Hz)",
     {"3DH5", "4DH5", "8DH5"}),
    ("omega_o", "omega_0_aver-Hz", "ωo (Hz)",
     {"3DH5", "4DH5", "8DH5"}),
    ("omega_io", "omega_max_i0_aver-Hz", "ωi+ωo (Hz)",
     {"3DH5", "4DH5", "8DH5"}),
    # devm_peak moved out (2026-07-19) into its own standalone script,
    # make_bt_tx_devm_peak_revision2.py, expanded to the full 13-packet
    # DEVM breadth (this pipeline's 3-packet-only scope no longer applies).
    ("freq_drift", "frequency_drift_aver-Hz", "Drift Rate (Hz)",
     {"HDT3", "HDT4", "HDT6", "HDT8"}),
]


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
def _read_all_cached(bt_tx_dir: Path, metric_items: tuple):
    """Single pass over every CSV: for EACH metric column, build
    raw[metric_key][(packet,supply,gain,pa_slices,dig_gain,freq)] = [(value,dut_id)].
    Reading every metric's column in one pass avoids re-scanning the 40
    CSVs once per metric. 2026-07-20: csv.DictReader -> csv.reader + a
    one-time header->index map (both for the RAW_KEY_COLS and for every
    metric's own data column)."""
    metric_cols = dict(metric_items)
    raw: dict[str, dict[tuple, list[tuple[float, str]]]] = {key: defaultdict(list) for key in metric_cols}
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
            metric_idx = {mk: col.get(mc) for mk, mc in metric_cols.items()}
            i_pattern = col.get("cfg-carrier_01-data_pattern")

            def cell(row, idx):
                return row[idx] if idx is not None and idx < len(row) else ""

            for row in reader:
                if i_pattern is not None and (cell(row, i_pattern) or "").strip() != DATA_PATTERN:
                    continue
                packet = (cell(row, i_packet) or "").strip()
                supply = (cell(row, i_supply) or "").strip()
                gain = (cell(row, i_gain) or "").strip()
                pa_slices = to_float(cell(row, i_slices))
                dig_gain = to_float(cell(row, i_dig))
                freq = to_float(cell(row, i_freq))
                if pa_slices is None or dig_gain is None or freq is None:
                    continue
                key_tuple = (packet, supply, gain, pa_slices, dig_gain, freq)
                for metric_key, idx in metric_idx.items():
                    val = to_float(cell(row, idx))
                    if val is None:
                        continue
                    raw[metric_key][key_tuple].append((val, dut))
                pa_slices_seen.add(pa_slices)
                dig_gain_seen.add(dig_gain)
        print(f"  read {path.name}")
    return raw, pa_slices_seen, dig_gain_seen


def read_all(bt_tx_dir: Path, metric_cols: dict):
    """Public wrapper: dicts aren't hashable, so delegate to the
    @lru_cache'd _read_all_cached() keyed by a sorted tuple-of-items form.
    A build script that reaches this via two paths (its own stats pass +
    the ATE sibling module's get_ate_data_all()) with the same metric set
    only reads the 40 files once."""
    return _read_all_cached(bt_tx_dir, tuple(sorted(metric_cols.items())))


def build_fixed_combos(raw_for_metric, pa_slices_max, dig_gain_max, packets):
    """Filter to pa_slices==max & dig_gain==max & packet in the metric's
    applicable set, drop pa_slices/dig_gain from the key."""
    data: dict[tuple, list[tuple[float, str]]] = defaultdict(list)
    for (packet, supply, gain, pa_slices, dig_gain, freq), values in raw_for_metric.items():
        if pa_slices != pa_slices_max or dig_gain != dig_gain_max:
            continue
        if packet not in packets:
            continue
        combo = (packet, supply, gain, freq)
        data[combo].extend(values)
    return data


def nice_step(value, target_ticks=10):
    """'Nice' gridline step (1/2/5 x 10^k) so a spec-anchored axis gets
    roughly target_ticks major divisions -- same convention as
    make_bt_tx_devm_rms_revision2.py's spec_window (2026-07-19), reused
    here for FreqAcc's Hz-scale limits (2026-07-20 customer request)."""
    if value <= 0:
        return 1
    raw_step = value / target_ticks
    # 2026-08-14 (customer request): never drop below a whole-number step --
    # sub-1 magnitudes (0.1/0.2/0.5) produced dense fractional tick labels
    # (e.g. 0, 0.5, 1, ..., 4.5) that are hard to read; forcing magnitude>=1
    # collapses that to integer steps (e.g. step=1 -> 0, 1, 2, 3, 4) while
    # leaving already-integer steps (value=20 -> step=2, value=35 -> step=5)
    # unchanged.
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


def spec_window(spec_min, spec_max, data_min, data_max, target_ticks=10):
    """Spec-anchored axis window (2026-07-20 customer request: every chart
    except Power/ACP must show the Limit line -- when a real USL/LSL
    exists, the axis has to be sized around it, not purely from the data).

    Generalizes DEVM's 0-anchored spec_window to FreqAcc's possibly
    negative, possibly one-sided (LSL-only or USL-only), possibly
    asymmetric Hz-scale limits:
      * on whichever side (min/max) has a real spec value, the window is
        anchored one nice step beyond it, so the limit line isn't glued to
        the plot edge;
      * on a side with NO spec value (e.g. df2_max is LSL-only -- "must be
        > 115kHz", no upper limit), that side falls back to the group's
        own data extent, same as the old pure-data-driven behavior on that
        side.
    Window boundaries are always whole integers, even when the spec value
    itself isn't (customer: "USL, LSL이 실수면...세로축 MAX/MIN을 정수로
    표시해줘"). Returns (vmin, vmax, step); caller must not call this when
    both spec_min and spec_max are None."""
    raw_min = spec_min if spec_min is not None else data_min
    raw_max = spec_max if spec_max is not None else data_max
    span = raw_max - raw_min
    if span <= 0:
        span = abs(raw_max) or abs(raw_min) or 1
    step = nice_step(span, target_ticks)
    if spec_min is not None:
        vmin = math.floor(round((spec_min - step) / step, 6)) * step
    else:
        vmin = math.floor(data_min)
    if spec_max is not None:
        vmax = math.ceil(round((spec_max + step) / step, 6)) * step
    else:
        vmax = math.ceil(data_max)
    return int(vmin), int(vmax), step


FREQACC_SNAP_STEP = 1000


def _snap_range(vmin_raw, vmax_raw, step=FREQACC_SNAP_STEP):
    """Floor/ceil to the nearest `step` Hz grid, then push ONE more step
    out if a bound already sits exactly on that grid -- same rationale as
    make_bt_tx_pwr_revision2._snap_range. 2026-08-15 customer request:
    "span이 난잡해서 1000의 배수로 조절" (e.g. 142969/152669 ->
    142000/153000) -- applied to non-spec-anchored groups only; spec-
    anchored groups already get a clean margin from spec_window's own
    nice-step anchoring."""
    vmin = math.floor(vmin_raw / step) * step
    vmax = math.ceil(vmax_raw / step) * step
    if vmin == vmin_raw:
        vmin -= step
    if vmax == vmax_raw:
        vmax += step
    return int(vmin), int(vmax)


def group_of(packet, gain, freq, supply, scatter_metric=False):
    """The axis-sharing unit. Default (old behavior, all metrics except
    devm_peak): (packet, gain, band) -- merges all 3 supplies together
    regardless of band. scatter_metric=True (devm_peak, 2026-07-18, same
    rule as EVM): 2.4G still merges all supplies into one slide, but 5G/6G
    puts each supply on its own slide, so its group must include supply."""
    band = band_of(freq)
    if not scatter_metric or band == 2.4:
        return (packet, gain, band)
    return (packet, gain, band, supply)


def build_packet_ranges(data, scatter_metric=False, spec_fn=None):
    """Per axis-sharing group (see group_of): floor/ceil integer Y-range
    covering the group's FULL data extent -- no value is ever excluded from
    the axis (2026-07-20 customer correction: a prior revision excluded
    IQR/spec-window outliers and annotated them as text instead, but a
    combo with 30+ outliers made that note box overflow the plot; the
    customer's actual preference is to just widen the axis so every point
    is visible as a normal bar -- exclude+annotate is gone for every
    FreqAcc metric, since all of them are frequency-unit (Hz) values. This
    does NOT apply to Power/ACP/DEVM, which keep their own IQR-based
    exclusion in their own scripts).

    spec_fn(packet, band) -> (min_hz, max_hz) or None. When a real spec
    value exists for a group's packet/band, spec_window() still anchors the
    axis one nice step beyond the spec (so the Limit line reads cleanly)
    -- but if the group's actual data extends further than that anchor on
    either side, the axis widens again to cover it. Also returns
    step_by_group: the nice gridline step for spec-anchored groups, or the
    old ceil(range/15) auto-step otherwise.

    Returns (ranges, step_by_group).
    """
    group_values: dict[tuple, list[float]] = defaultdict(list)
    group_packet_band: dict[tuple, tuple] = {}
    for combo, values in data.items():
        packet, supply, gain, freq = combo
        group = group_of(packet, gain, freq, supply, scatter_metric)
        group_values[group].extend(v for v, _d in values)
        group_packet_band[group] = (packet, band_of(freq))

    ranges: dict[tuple, tuple[int, int]] = {}
    step_by_group: dict[tuple, float] = {}
    for group, values in group_values.items():
        data_min, data_max = min(values), max(values)
        packet, band = group_packet_band[group]
        spec = spec_fn(packet, band) if spec_fn else None
        if spec is not None and (spec[0] is not None or spec[1] is not None):
            vmin, vmax, step = spec_window(spec[0], spec[1], data_min, data_max)
            vmin = min(vmin, math.floor(data_min))
            vmax = max(vmax, math.ceil(data_max))
        else:
            vmin, vmax = _snap_range(data_min, data_max)
            step = max(1, math.ceil((vmax - vmin) / 15))
        ranges[group] = (vmin, vmax)
        step_by_group[group] = step

    return ranges, step_by_group


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
# header + Bench/ATE subheader + Mean+/-Std + Min/Max + USL + LSL -- most
# FreqAcc sub-metrics have a symmetric +-spec (2 rows); cfo/drift_rate/
# df2_avg have none, so this over-estimates table_h slightly for those --
# an approximation, see make_bt_tx_pwr_revision2.py's _TABLE_N_ROWS note.
_TABLE_N_ROWS = 6


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


def draw_png(values, title, xlabel, png_path: Path, value_range=None, spec=None,
             step=1, figsize=None) -> None:
    """`values` is every value for the combo -- nothing is excluded, the
    caller (build_packet_ranges) already widened `value_range` to cover the
    group's full data extent (2026-07-20 customer correction). `step` is
    the gridline step decided by build_packet_ranges -- the spec_window()-
    derived nice step for spec-anchored groups, or the old ceil(range/15)
    auto-step otherwise; always passed in now rather than recomputed here,
    so every group's step is decided in exactly one place."""
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
    # Integer gridlines, but step scaled to the range -- Power/EVM ranges are
    # only ~10-90 units so step=1 was fine there; these Hz-scale metrics can
    # span thousands of units, where step=1 would create thousands of tick
    # labels and make rendering pathologically slow.
    y_lo, y_hi = math.floor(vmin), math.ceil(vmax)
    # Always include the exact boundary values, even if they don't fall on
    # a step multiple -- otherwise the topmost/bottommost tick label can
    # read a bit short of the true min/max, which mismatches the PPT
    # ruler's Min/Max text and confuses the customer.
    n_steps = max(1, int(round((y_hi - y_lo) / step)))
    yticks = sorted({int(round(y_lo + i * step)) for i in range(n_steps + 1)} | {y_lo, y_hi})
    # 2026-08-14 real bug: `step` targets ~15 gridlines sized for the OLD
    # fixed-tall (9,14) figure -- figure_size_for() now hands back much
    # shorter figures for narrow-column PPT rows, where 15 tick labels at
    # even the font-scaled size still vertically collide into an
    # unreadable stack (screenshot from the customer, e.g. Df2 Average).
    # Thin the tick SET (not the step used for gridline spacing/data
    # binning) down to a count that actually fits the available height,
    # always keeping the exact y_lo/y_hi boundary values.
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
        # va="top" visually overlaps the axhline in tall/wide-range figures
        # (confirmed matplotlib rendering quirk, 2026-07-17) -- use
        # va="bottom" uniformly with a manual data-space pad instead.
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

    fig.tight_layout()
    fig.savefig(png_path, dpi=450)
    plt.close(fig)


def write_xlsx(xlsx_path: Path, title: str, xlabel: str, values, value_range=None) -> None:
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
    p.add_argument("--out-base", type=Path, default=PATHS.result_xlsx_dir)
    p.add_argument("--png-base", type=Path, default=PATHS.result_png_dir)
    p.add_argument("--metric", type=str, default="",
                   help="Only generate this metric key (0 = all). See METRICS for keys.")
    p.add_argument("--limit", type=int, default=0,
                   help="Generate only the first N combos PER METRIC (0 = all).")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    metrics = [m for m in METRICS if not args.metric or m[0] == args.metric]
    if not metrics:
        raise SystemExit(f"Unknown --metric {args.metric!r}. Valid keys: {[m[0] for m in METRICS]}")

    print("Parsing FreqAcc spec limits...")
    import bt_tx_freqacc_spec_lookup as spec_lookup
    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    metric_cols = {key: col for key, col, _xlabel, _packets in metrics}
    print(f"Reading BT_Tx CSVs for {len(metric_cols)} metric column(s)...")
    raw_by_metric, pa_slices_seen, dig_gain_seen = read_all(args.bt_tx_dir, metric_cols)
    pa_slices_max = max(pa_slices_seen)
    dig_gain_max = max(dig_gain_seen)
    print(f"pa_slices seen: {sorted(pa_slices_seen)} -> fixed at max {fmt_num(pa_slices_max)}")
    print(f"dig_gain seen: {sorted(dig_gain_seen)} -> fixed at max {fmt_num(dig_gain_max)}")

    for key, col, xlabel, packets in metrics:
        print(f"\n=== {key} ({col}) -- packets: {sorted(packets)} ===")
        data = build_fixed_combos(raw_by_metric[key], pa_slices_max, dig_gain_max, packets)
        combos = list(data.keys())
        print(f"combos (packet,supply,gain,frequency): {len(combos)}")
        if not combos:
            print("  (no data -- skipping)")
            continue

        def spec_fn(packet, band, _key=key):
            return spec_lookup.lookup(spec_table, _key, packet, band)

        packet_ranges, step_by_group = build_packet_ranges(data, spec_fn=spec_fn)
        n_spec_anchored = sum(1 for g in step_by_group if spec_fn(g[0], g[2]) not in (None, (None, None)))
        print(f"Spec-anchored (Limit-visible) groups: {n_spec_anchored}/{len(step_by_group)}")

        # 2026-08-15 (customer request): Bench and ATE charts must share
        # the EXACT same Y-axis span. Lazy/deferred import (not at module
        # top-level) to avoid a circular import -- make_ate_tx_freqacc_
        # revision2 imports this module as FA; by the time main() runs,
        # this module is already fully initialized, so Python just
        # returns it from sys.modules. get_ate_group_extents() is called
        # from BOTH this loop and the ATE module's own draw_all(), so both
        # compute the identical widened+snapped range independently from
        # the same two raw data sources -- no file-based coordination
        # needed.
        import make_ate_tx_freqacc_revision2 as ate_fa_mod
        for group, (lo, hi) in ate_fa_mod.get_ate_group_extents(key, data).items():
            if group in packet_ranges:
                vmin, vmax = packet_ranges[group]
                packet_ranges[group] = _snap_range(min(vmin, lo), max(vmax, hi))

        if args.limit > 0:
            combos = combos[: args.limit]
            print(f"Limiting to first {len(combos)} combo(s)")

        out_dir = args.out_base / f"{key}_revision2"
        png_dir = args.png_base / f"{key}_revision2"
        out_dir.mkdir(parents=True, exist_ok=True)
        png_dir.mkdir(parents=True, exist_ok=True)

        # Precompute each combo's PPT column count -- see make_bt_tx_pwr_
        # revision2.py's identical block for the full rationale. Computed
        # per metric key since each sub-metric's PNGs are discovered/laid
        # out on their own PPT slides, independent of every other key.
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
            spec = spec_lookup.lookup(spec_table, key, packet, band_of(freq))

            group = (packet, gain, band_of(freq))
            value_range = packet_ranges[group]
            step = step_by_group.get(group, 1)
            figsize = figure_size_for(n_cols_for(packet, supply, freq))
            jobs.append((all_values, title, xlabel, png_dir / f"{fname}.png", value_range, spec, step,
                         figsize, out_dir / f"{fname}.xlsx"))

        n = 0
        with ProcessPoolExecutor(max_workers=PATHS.png_workers) as pool:
            for _ in pool.map(_draw_one, jobs, chunksize=4):
                n += 1
                if n % 200 == 0 or n == len(jobs):
                    print(f"  [{n}/{len(jobs)}]")

        print(f"Done: {key} -> XLSX {out_dir}  PNG {png_dir}")


def _draw_one(job) -> None:
    """Worker for ProcessPoolExecutor -- must be a module-level function so
    it's picklable on Windows (spawn-only, no fork)."""
    all_values, title, xlabel, png_path, value_range, spec, step, figsize, xlsx_path = job
    draw_png(all_values, title, xlabel, png_path, value_range, spec, step, figsize)
    if PATHS.write_xlsx:
        write_xlsx(xlsx_path, title, xlabel, all_values, value_range)


if __name__ == "__main__":
    main()
