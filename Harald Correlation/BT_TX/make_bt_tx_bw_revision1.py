"""BT TX 6dB/20dB Bandwidth -- Bench-side reader.

New items added to the TX report per user request 2026-08-26:
  6dB BW    bandwidth_6db_aver-MHz   (has an ATE counterpart, see
                                       make_ate_tx_bw_revision1.py)
  20dB BW   bandwidth_20db_aver-MHz  (Bench-only -- confirmed no ATE item
                                       exists for this one)

Same grouping key/conventions as every other BT_TX metric (see
make_bt_tx_pwr_revision2.py's module docstring for the fuller rationale):
  * key = (packet_type, pa_supply, pa_gain, pa_slices, dig_gain,
    frequency-MHz), pa_slices/dig_gain fixed at the dataset-wide max
    afterward via build_fixed_combos().
  * PRBS9-only (cfg-carrier_01-data_pattern) -- Bench sweeps 3 data
    patterns for some packets, ATE only ever measures PRBS9; pooling all 3
    would triple the real sample count.
  * DUT identity per value carried via dut_id_of() (the "#NN" index parsed
    from the Bench filename), same as every other bench module.

Unlike Power, BW is read for EVERY packet_type Bench has (no applicable-
packet allowlist) -- confirmed via real data: bandwidth_6db_aver-MHz has
non-blank values for all 15 packet_types in the 01.Tx pull, and the value
genuinely varies by packet_type (e.g. 1DH5=0.376MHz vs LE1M=0.664MHz at
the same channel/supply) -- this is real signal, not noise to be pooled
away.

2026-08-27: chart-drawing/PPT wiring added -- BW06DB/BW20DB are now their
own PPT sections (Option A, user-confirmed), reusing build_ppt_pwr_evm_
revision2.py's generic add_band24_slide/add_highfreq_supply_slide exactly
like DEVM RMS/Peak/99pct already do. No spec/USL-LSL exists for either BW
item (user confirmed no spec sheet was provided), so build_packet_ranges()
below never takes/uses a spec_by_packet argument -- every group goes
through the plain data-driven (IQR-outlier + nice-step axis) path.

One thing this family needs that DEVM/Power/FreqAcc don't: a genuinely
scale-aware gridline step. Confirmed against the real 01.Tx pull, BW
values are sub-1 to low-single-digit MHz with per-axis-sharing-group
spreads as small as ~0.005MHz (e.g. HDRPS2 2.012-2.017) up to ~0.16MHz
(e.g. 2DH5 0.376-0.532) -- DEVM's own floor/ceil-to-INTEGER-with-step-1
convention (tuned for %/dBm-scale metrics per an explicit 2026-08-14
customer request to "never drop below a whole-number step") would collapse
almost every BW histogram to a 2-tick 0..1 (or N..N+1) axis while the real
data occupies a sliver of it -- see nice_step()'s own docstring below.
"""

from __future__ import annotations

import argparse
import csv
import functools
import math
import re
import statistics
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
DATA_PATTERN = "PRBS9"
DUT_RE = re.compile(r"#(\d+)")

# metric_key -> Bench column name
METRIC_COLS = {
    "bw6": "bandwidth_6db_aver-MHz",
    "bw20": "bandwidth_20db_aver-MHz",
}

# metric_key -> (chart-title metric label, Y-axis label, PNG/XLSX subfolder
# name, section/TOC label) -- one small table instead of scattering these
# four strings across main()/build_ppt_combined_revision2.py.
METRIC_META = {
    "bw6": {"metric_label": "BW06DB", "xlabel": "6dB BW (MHz)",
            "dirname": "bw06db_revision1", "section_label": "6dB BW"},
    "bw20": {"metric_label": "BW20DB", "xlabel": "20dB BW (MHz)",
             "dirname": "bw20db_revision1", "section_label": "20dB BW"},
}

BINS = 40
# Same per-combo IQR outlier rule as make_bt_tx_devm_rms_revision2.py's
# find_combo_outliers (IQR_MULT=5.0) -- copied verbatim, see that module's
# docstring for the full rationale. Reused here even though BW has no spec
# to anchor against: still useful for flagging a genuine within-condition
# anomaly (one DUT far from the other 39 measured under the same setting).
IQR_MULT = 5.0
# Same fold-in-vs-exclude-and-annotate decision DEVM uses, just expressed in
# STEP units instead of DEVM's fixed absolute-5 (which was sized for its own
# integer-step %/dBm scale) -- see nice_step()'s docstring for why BW needs
# its own scale-aware step in the first place.
OUTLIER_SCATTER_STEPS = 5
# 2026-08-27 (customer request via LEE SeokJin): axis min/max floor/ceil
# step for BW specifically -- fixed at 0.1MHz (1 decimal place), NOT the
# scale-aware nice_step() above. See build_packet_ranges's own comment for
# the full rationale (the customer wants a coarse, always-1-decimal axis
# even though it visually squeezes tight real spreads into ~one bar).
BW_AXIS_STEP = 0.1


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


@functools.lru_cache(maxsize=None)
def read_all(bt_tx_dir: Path):
    """Single pass over every CSV: for EACH of bw6/bw20, build
    raw[metric_key][(packet,supply,gain,pa_slices,dig_gain,freq)] =
    [(value,dut_id), ...]. Also returns the pa_slices/dig_gain sets seen,
    for the caller's own dataset-wide-max fixing (see build_fixed_combos)."""
    raw: dict[str, dict[tuple, list[tuple[float, str]]]] = {
        key: defaultdict(list) for key in METRIC_COLS
    }
    pa_slices_seen: set[float] = set()
    dig_gain_seen: set[float] = set()

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
            i_pattern = col.get("cfg-carrier_01-data_pattern")
            metric_idx = {mk: col.get(mc) for mk, mc in METRIC_COLS.items()}

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
                key = (packet, supply, gain, pa_slices, dig_gain, freq)
                for metric_key, idx in metric_idx.items():
                    val = to_float(cell(row, idx))
                    if val is None:
                        continue
                    raw[metric_key][key].append((val, dut))
                pa_slices_seen.add(pa_slices)
                dig_gain_seen.add(dig_gain)
        print(f"  read {path.name}")
    return raw, pa_slices_seen, dig_gain_seen


def build_fixed_combos(raw_for_metric, pa_slices_max, dig_gain_max):
    """Filter to pa_slices==max & dig_gain==max, drop those two from the
    key -- same convention as every other BT_TX metric module."""
    data: dict[tuple, list[tuple[float, str]]] = defaultdict(list)
    for (packet, supply, gain, pa_slices, dig_gain, freq), values in raw_for_metric.items():
        if pa_slices != pa_slices_max or dig_gain != dig_gain_max:
            continue
        combo = (packet, supply, gain, freq)
        data[combo].extend(values)
    return data


def find_combo_outliers(values):
    """IQR rule on a SINGLE combo's own ~40 DUT values -- copied verbatim
    from make_bt_tx_devm_rms_revision2.find_combo_outliers, see that
    function's docstring for the full rationale. Returns (kept, outliers)
    where each is a list of (value, dut_id)."""
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
    """The axis-sharing unit ("slide") -- copied verbatim from
    make_bt_tx_devm_rms_revision2.group_of: 2.4G merges all 3 supplies into
    one slide (group ignores supply); 5G/6G puts each supply on its own
    slide (group includes supply)."""
    band = band_of(freq)
    if band == 2.4:
        return (packet, gain, 2.4)
    return (packet, gain, band, supply)


def nice_step(value, target_ticks=10):
    """'Nice' gridline step (1/2/5 x 10^k), roughly target_ticks divisions
    over 0..value -- the plain generic form, WITHOUT DEVM/FreqAcc's own
    nice_step's 2026-08-14 "never drop below a whole-number step" floor.
    That floor was a customer request specific to DEVM's %-scale and
    FreqAcc's Hz-scale spec-anchored axes; BW has no spec at all and its
    real per-group data spreads run from ~0.005MHz to ~0.16MHz (confirmed
    against the 01.Tx pull -- see this module's docstring), so forcing
    step>=1 would put every histogram on a 2-tick 0..1MHz (or N..N+1MHz)
    axis with the real data squeezed into a sliver at one end."""
    if value <= 0:
        return 0.001
    raw_step = value / target_ticks
    magnitude = 10 ** math.floor(math.log10(raw_step))
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


def _floor_to_step(value, step):
    return math.floor(round(value / step, 6)) * step


def _ceil_to_step(value, step):
    return math.ceil(round(value / step, 6)) * step


def build_packet_ranges(data):
    """Per axis-sharing group (see group_of): per-combo IQR outlier removal
    (find_combo_outliers), pooled per group, then a step-grid-rounded
    Y-range sized to the group's OWN spread via nice_step() -- see that
    function's docstring for why BW can't reuse DEVM's integer-only rule.

    No spec exists for either BW item (user-confirmed), so -- unlike
    make_bt_tx_devm_rms_revision2.build_packet_ranges -- this never takes a
    spec_by_packet argument; every group goes through what DEVM calls its
    "non-spec" data-driven path. Same widen-and-fold-in vs
    exclude-and-annotate outlier decision DEVM uses, just measured in STEP
    units (OUTLIER_SCATTER_STEPS) instead of DEVM's fixed absolute-5.

    Returns the same 5-tuple shape build_ppt_pwr_evm_revision2.
    get_ranges_and_stats() expects from a "DEVM-family" module (len(result)
    >= 4, kept_by_combo at index 1):
      ranges[group] = (min, max)
      kept_by_combo[(packet, supply, gain, freq)] = [(value, dut_id), ...]
      outliers_by_combo[(packet, supply, gain, freq)] = [(value, dut_id), ...]
      include_outliers_group[group] = bool
      step_by_group[group] = gridline step
    """
    kept_by_combo: dict[tuple, list[tuple[float, str]]] = {}
    outliers_by_combo: dict[tuple, list[tuple[float, str]]] = {}
    for combo, values in data.items():
        kept, outliers = find_combo_outliers(values)
        kept_by_combo[combo] = kept
        outliers_by_combo[combo] = outliers

    per_group_kept: dict[tuple, list[float]] = defaultdict(list)
    per_group_outliers: dict[tuple, list[float]] = defaultdict(list)
    for (packet, supply, gain, freq), kept in kept_by_combo.items():
        per_group_kept[group_of(packet, gain, freq, supply)].extend(v for v, _ in kept)
    for (packet, supply, gain, freq), outliers in outliers_by_combo.items():
        if outliers:
            per_group_outliers[group_of(packet, gain, freq, supply)].extend(v for v, _ in outliers)

    ranges: dict[tuple, tuple[float, float]] = {}
    include_outliers_group: dict[tuple, bool] = {}
    step_by_group: dict[tuple, float] = {}
    for group, values in per_group_kept.items():
        data_min, data_max = min(values), max(values)
        # 2026-08-27 (customer request via LEE SeokJin): the scale-aware
        # nice_step() above was originally added so BW's own tight real
        # spreads (~0.005-0.16MHz) wouldn't collapse onto DEVM's whole-
        # unit axis -- but the customer explicitly wants the OPPOSITE for
        # BW specifically: a coarse, fixed 0.1MHz-rounded axis (e.g. data
        # 0.371-0.376 -> axis 0.3-0.4), even if that visually squeezes
        # every real value into what looks like a single bar/line. Fixed
        # step (not nice_step's dynamic one) matches that exactly.
        step = BW_AXIS_STEP
        base_vmin = _floor_to_step(data_min, step)
        base_vmax = _ceil_to_step(data_max, step)
        if base_vmax <= base_vmin:
            base_vmax = base_vmin + step

        group_outliers = per_group_outliers.get(group)
        if not group_outliers:
            ranges[group] = (base_vmin, base_vmax)
            include_outliers_group[group] = False
        else:
            excl_max = max(group_outliers)
            threshold = step * OUTLIER_SCATTER_STEPS
            if (excl_max - base_vmax) >= threshold:
                ranges[group] = (base_vmin, base_vmax)
                include_outliers_group[group] = False
            else:
                widened_vmax = max(base_vmax, _ceil_to_step(excl_max, step) + step)
                ranges[group] = (base_vmin, widened_vmax)
                include_outliers_group[group] = True
        step_by_group[group] = step

    return ranges, kept_by_combo, outliers_by_combo, include_outliers_group, step_by_group


def combo_name(combo, pa_slices_max, dig_gain_max) -> str:
    packet, supply, gain, freq = combo
    return (
        f"packet_type({packet})_frequency-MHz({fmt_num(freq)})_pa_supply({supply})_"
        f"pa_slices({fmt_num(pa_slices_max)})_pa_gain({gain})_"
        f"dig_gain({fmt_num(dig_gain_max)})"
    )


def combo_title_short(combo, pa_slices_max, dig_gain_max) -> str:
    """Short plot-title format, same convention as every other BT_TX metric
    module's combo_title_short. Does NOT change the PNG/XLSX FILENAME
    (still combo_name(), since build_ppt_pwr_evm_revision2.discover()'s
    regex parses that exact verbose format) -- only the chart's own title
    text."""
    packet, supply, gain, freq = combo
    return (f"{packet}_{fmt_num(freq)}MHz_HPA_{hpa_label(supply)}_"
            f"slices({fmt_num(pa_slices_max)})_gain({gain})_dig_gain({fmt_num(dig_gain_max)})")


def excluded_text(excluded) -> str:
    if not excluded:
        return ""
    lines = [f"{v:.3f} ({dut})" for v, dut in excluded]
    return "Excluded outlier(s):\n" + "\n".join(lines)


def _place_excluded_note(ax, edges, counts, note, fontsize=28) -> None:
    """Copied verbatim from make_bt_tx_devm_rms_revision2._place_excluded_
    note -- see that function's docstring for the full rationale (places
    the excluded-outlier annotation inside the axes, in a bar-free gap)."""
    n_lines = note.count("\n") + 1
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
    """Copied verbatim from make_bt_tx_devm_rms_revision2._fit_title_
    fontsize -- largest bold fontsize (<=start) whose rendered width stays
    within the chart's own plot frame."""
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
# header + Bench/ATE subheader + Mean+/-Std + Min/Max -- BW has NO spec at
# all (unlike DEVM RMS's _TABLE_N_ROWS=5, which budgets one extra row for
# its USL-only spec line), so this is 4 for both bw6 (interleaved
# Bench/ATE table) and bw20 (plain Bench-only table, which is actually
# only 3 rows) -- an over-estimate for bw20 is harmless, same approximation
# convention as every other module's _TABLE_N_ROWS (see make_bt_tx_pwr_
# revision2.py's identical note).
_TABLE_N_ROWS = 4


def figure_size_for(n_cols: int) -> tuple[float, float]:
    cell_width = _GRID_WIDTH / n_cols if n_cols == 1 else (_GRID_WIDTH - (n_cols - 1) * _CELL_GAP_H) / n_cols
    table_h = _TABLE_N_ROWS * _TABLE_ROW_H
    plots_top = _TABLE_TOP + table_h + _TABLE_PLOT_GAP
    image_h = (_CONTENT_AREA_BOTTOM - plots_top - _ATE_GAP) / 2
    return (cell_width, image_h)


_REF_FIGSIZE = (9, max(4, 0.3 * BINS + 2))  # the shape every font size below was tuned against


def _font_scale(figsize) -> float:
    """See make_bt_tx_pwr_revision2._font_scale's docstring -- same fix,
    same reasoning."""
    if figsize is None:
        return 1.0
    return max(0.35, min(1.0, figsize[1] / _REF_FIGSIZE[1]))


def draw_png(values, title, xlabel, png_path: Path, value_range=None, excluded=None, spec=None,
             step=1, figsize=None) -> None:
    """Copied verbatim from make_bt_tx_devm_rms_revision2.draw_png (already
    scale-agnostic -- every tick/step computation below works the same for
    a fractional step as it does for DEVM's integer one). `spec` is always
    None for BW (no spec exists) but the parameter is kept for interface
    parity with every other metric module's draw_png."""
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
    n_steps = int(round((vmax - vmin) / step))
    yticks = sorted({round(vmin + i * step, 6) for i in range(n_steps + 1)} | {round(vmin, 6), round(vmax, 6)})
    y_lo_t, y_hi_t = round(vmin, 6), round(vmax, 6)
    max_ticks = max(4, min(8, round(15 * fs)))
    if len(yticks) > max_ticks:
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
    ax.set_xlim(0, max_count * 1.08)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True))
    plt.setp(ax.get_yticklabels(), fontsize=25 * fs, fontweight="bold")
    plt.setp(ax.get_xticklabels(), fontsize=25 * fs, fontweight="bold")
    ax.set_ylabel(xlabel, fontsize=20 * fs, fontweight="bold")
    ax.set_xlabel("Count", fontsize=20 * fs, fontweight="bold")
    ax.grid(True, axis="both", color="0.85", linewidth=0.6)
    ax.set_axisbelow(True)
    fig.tight_layout()
    ax.set_title(title, fontsize=_fit_title_fontsize(fig, ax, title, start=int(18 * fs), min_size=int(max(5, 7 * fs))),
                 fontweight="bold")

    if spec is not None:
        spec_min, spec_max = spec
        pad = (vmax - vmin) * 0.035
        for spec_val, spec_label in ((spec_min, "LSL"), (spec_max, "USL")):
            if spec_val is None or not (vmin <= spec_val <= vmax):
                continue
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


def _stats(values):
    return (statistics.mean(values), min(values), max(values),
            statistics.stdev(values) if len(values) > 1 else 0.0)


def get_ranges_and_stats(metric_key, bt_tx_dir=None):
    """(ranges, combo_stats, pa_slices_max, dig_gain_max) for ONE metric key
    (bw6/bw20) -- same contract as build_ppt_pwr_evm_revision2.PE.
    get_ranges_and_stats(mod) (ranges[group]=(min,max), combo_stats[combo]=
    (mean,min,max,std) from kept_by_combo), but reads via THIS module's own
    two-metric read_all() instead of going through PE's generic
    mod.read_all()/mod.build_fixed_combos(raw, ...) interface -- PE's
    version assumes read_all() returns a single flat raw dict ready for
    build_fixed_combos, but this module's read_all() returns raw['bw6']/
    raw['bw20'] (one pass covers both metrics). Same reason make_bt_tx_
    freqacc_revision2.py's deck builder has its own get_ranges_and_stats_
    by_metric() instead of reusing PE's generic one -- see that module's
    docstring."""
    bt_tx_dir = bt_tx_dir or PATHS.bench_dir
    raw, pa_slices_seen, dig_gain_seen = read_all(bt_tx_dir)
    pa_slices_max = max(pa_slices_seen)
    dig_gain_max = max(dig_gain_seen)
    data = build_fixed_combos(raw[metric_key], pa_slices_max, dig_gain_max)
    ranges, kept_by_combo, _outliers, _incl, _step = build_packet_ranges(data)
    combo_stats = {}
    for combo, kept in kept_by_combo.items():
        vals = [v for v, _dut in kept]
        if vals:
            combo_stats[combo] = _stats(vals)
    return ranges, combo_stats, pa_slices_max, dig_gain_max


def _n_cols_for_metric(combos):
    """Precompute each combo's PPT column count -- same block as every other
    BT_TX metric module's main(), factored into a helper here since main()
    below runs it once per metric key (bw6, bw20)."""
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

    return n_cols_for


def _draw_one_metric(bt_tx_dir: Path, metric_key: str, out_base: Path, png_base: Path,
                      limit: int = 0, count_only: bool = False) -> None:
    meta = METRIC_META[metric_key]
    xlabel = meta["xlabel"]
    out_dir = out_base / meta["dirname"]
    png_dir = png_base / meta["dirname"]

    raw, pa_slices_seen, dig_gain_seen = read_all(bt_tx_dir)
    pa_slices_max = max(pa_slices_seen)
    dig_gain_max = max(dig_gain_seen)
    print(f"pa_slices seen: {sorted(pa_slices_seen)} -> fixed at max {fmt_num(pa_slices_max)}")
    print(f"dig_gain seen: {sorted(dig_gain_seen)} -> fixed at max {fmt_num(dig_gain_max)}")

    data = build_fixed_combos(raw[metric_key], pa_slices_max, dig_gain_max)
    combos = list(data.keys())
    print(f"[{metric_key}] combos (packet,supply,gain,frequency): {len(combos)}")

    packet_ranges, kept_by_combo, outliers_by_combo, include_outliers_group, step_by_group = build_packet_ranges(data)
    n_outlier_combos = sum(1 for v in outliers_by_combo.values() if v)
    n_outlier_points = sum(len(v) for v in outliers_by_combo.values())
    n_extended_groups = sum(1 for v in include_outliers_group.values() if v)
    print(f"[{metric_key}] per axis-sharing-group Y-axis scale: {len(packet_ranges)} groups")
    print(f"    IQR outlier points found: {n_outlier_points} (across {n_outlier_combos} combo(s))")
    print(f"    Groups with axis extended + outliers folded in: {n_extended_groups}")

    # 2026-08-27: bw6 (only) has an ATE counterpart -- Bench and ATE charts
    # must share the exact same Y-axis span, same "get_ate_group_extents()
    # called from both sides" convention as DEVM RMS (see make_bt_tx_devm_
    # rms_revision2.main()'s identical block for the full rationale). bw20
    # has no ATE module at all, so this is a no-op for it.
    if metric_key == "bw6":
        import make_ate_tx_bw_revision1 as ate_bw_mod
        for group, (lo, hi) in ate_bw_mod.get_ate_group_extents(data, pa_slices_max, dig_gain_max).items():
            if group in packet_ranges:
                vmin, vmax = packet_ranges[group]
                step = step_by_group.get(group) or nice_step(max(hi - lo, 1e-9))
                packet_ranges[group] = (min(vmin, _floor_to_step(lo, step)), max(vmax, _ceil_to_step(hi, step)))

    if count_only:
        return

    if limit > 0:
        combos = combos[:limit]
        print(f"[{metric_key}] limiting to first {len(combos)} combo(s)")

    out_dir.mkdir(parents=True, exist_ok=True)
    png_dir.mkdir(parents=True, exist_ok=True)

    n_cols_for = _n_cols_for_metric(list(data.keys()))

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
        figsize = figure_size_for(n_cols_for(packet, supply, freq))
        jobs.append((hist_values, title, xlabel, png_dir / f"{fname}.png", value_range, excluded, step,
                     figsize, all_values, out_dir / f"{fname}.xlsx"))

    n = 0
    with ProcessPoolExecutor(max_workers=PATHS.png_workers) as pool:
        for _ in pool.map(_draw_one, jobs, chunksize=4):
            n += 1
            if n % 200 == 0 or n == len(jobs):
                print(f"  [{metric_key}] [{n}/{len(jobs)}]")

    print(f"[{metric_key}] Done. XLSX -> {out_dir}")
    print(f"[{metric_key}]       PNG  -> {png_dir}")


def _draw_one(job) -> None:
    """Worker for ProcessPoolExecutor -- must be a module-level function so
    it's picklable on Windows (spawn-only, no fork)."""
    hist_values, title, xlabel, png_path, value_range, excluded, step, figsize, all_values, xlsx_path = job
    draw_png(hist_values, title, xlabel, png_path, value_range, excluded, spec=None, step=step, figsize=figsize)
    if PATHS.write_xlsx:
        write_xlsx(xlsx_path, title, xlabel, all_values, value_range, excluded)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--bt-tx-dir", type=Path, default=PATHS.bench_dir)
    p.add_argument("--out-base", type=Path, default=PATHS.result_xlsx_dir)
    p.add_argument("--png-base", type=Path, default=PATHS.result_png_dir)
    p.add_argument("--metric", type=str, default="",
                   help="Only generate this metric key (bw6/bw20; default = both).")
    p.add_argument("--limit", type=int, default=0,
                   help="Generate only the first N combos PER METRIC (0 = all).")
    p.add_argument("--count-only", action="store_true",
                   help="Only print combo/sample-size audit, generate nothing.")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    metric_keys = [args.metric] if args.metric else list(METRIC_META.keys())
    for mk in metric_keys:
        if mk not in METRIC_META:
            raise SystemExit(f"Unknown --metric {mk!r}. Valid keys: {list(METRIC_META.keys())}")
    for mk in metric_keys:
        print(f"\n=== {mk} ({METRIC_META[mk]['metric_label']}) ===")
        _draw_one_metric(args.bt_tx_dir, mk, args.out_base, args.png_base,
                          limit=args.limit, count_only=args.count_only)


if __name__ == "__main__":
    main()
