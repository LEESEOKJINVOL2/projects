"""BT TX ACP revision2 stacked-bar distribution plots.

Changes from acp_revision1, per customer feedback (several rounds of
sample review before landing here):
  * Only pa_supply=0 is used (not 0/1/2).
  * All 21 offsets (-10..+10), which used to be 21 separate charts per
    (packet_type, pa_supply, pa_gain, band) combo, are now merged into ONE
    chart per combo.
  * band is still merged (multiple exact frequencies sharing a band feed
    one chart), and each frequency is distinguishable by color.
  * Y axis = Offset, -10 topmost .. +10 bottommost. Each offset owns a
    block of BINS rows: the offset's own combined (all-frequency) ACP
    values are histogrammed locally (auto range, not shared across
    offsets), and each bin becomes one horizontal bar, stacked by
    frequency color. The bin's value range + total count floats as text
    just past the end of its bar (e.g. "58.57~58.96 n=3").
  * X axis = Count.

pa_slices/dig_gain remain fixed at the dataset-wide max, same as revision1.

Two outputs per (packet_type, band, pa_gain) combo:
  * Excel workbook: a stacked horizontal BarChart mirroring the PNG
    (categories = offset+bin, one series per frequency), plus the same
    bin/count/raw detail table.
  * PNG image.

File name = packet_type(..)_band(..)_pa_supply(0)_pa_slices(..)_pa_gain(..)_dig_gain(..)
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
from matplotlib import cm

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
    "cfg-carrier_01-band",
]
OFFSET_COL = "meas-offset_frequency-MHz"
FREQ_COL = "cfg-carrier_01-frequency-MHz"
Y_COL = "acp_avg-dBm"
FIXED_SUPPLY = "0"
# 2026-08-15 real bug -- see make_bt_tx_pwr_revision2.py's identical
# comment for the full rationale: cfg-carrier_01-data_pattern was missing
# from RAW_KEY_COLS, silently pooling 3 data-pattern sweeps (PRBS9/PAT1/
# PAT2) for 1DH5/LE1M/LE2M into one histogram (tripling the real sample
# count) while ATE only ever measures PRBS9.
DATA_PATTERN = "PRBS9"


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


def color_hex(i: int, n: int) -> str:
    r, g, b, _ = cm.get_cmap("hsv")(i / max(1, n))
    return f"{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}"


# 2026-08-25: added for the Raw Values sheet (add_bt_tx_raw_values_sheet.py)
# -- ACP's own read_all() below has never needed per-DUT identity, unlike
# every other BT_TX bench module (make_bt_tx_pwr_revision2.py etc.), which
# already tag every value with dut_id_of(path) via this exact DUT_RE. read_all()
# itself is untouched; read_all_with_dut() below is the only consumer.
DUT_RE = re.compile(r"#(\d+)")


def dut_id_of(path: Path) -> str:
    m = DUT_RE.search(path.name)
    return f"#{m.group(1)}" if m else path.stem


@functools.lru_cache(maxsize=None)
def read_all(bt_tx_dir: Path):
    """Single pass: build raw[(packet,gain,pa_slices,dig_gain,band)][offset][freq] = [values].
    Only pa_supply == FIXED_SUPPLY rows are kept.

    2026-07-20: csv.DictReader -> csv.reader + one-time header->index map;
    @lru_cache so a process calling this twice (e.g. a build script
    consuming both this module and its ATE-sibling module) only reads the
    40 files once -- safe, nothing downstream mutates the returned raw
    nested-dict structure."""
    raw: dict[tuple, dict[float, dict[float, list[float]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(list))
    )
    pa_slices_seen: set[float] = set()
    dig_gain_seen: set[float] = set()

    # 2026-08-25: rglob + suffix filter, not a flat glob -- see
    # make_bt_tx_pwr_revision2.py's read_all() for why (a newer Bench pull
    # nests each DUT's file under #NN_<serial>/, alongside an unrelated
    # summary/summary_<timestamp>.csv that a blanket "*.csv" would wrongly
    # sweep in too).
    files = sorted(bt_tx_dir.rglob("*_cns_temp_data.csv"))
    for path in files:
        with path.open(encoding="utf-8-sig", newline="") as fh:
            reader = csv.reader(fh)
            header = next(reader)
            col = {name: i for i, name in enumerate(header)}
            i_packet = col.get(RAW_KEY_COLS[0])
            i_supply = col.get(RAW_KEY_COLS[1])
            i_gain = col.get(RAW_KEY_COLS[2])
            i_slices = col.get(RAW_KEY_COLS[3])
            i_dig = col.get(RAW_KEY_COLS[4])
            i_band = col.get(RAW_KEY_COLS[5])
            i_offset = col.get(OFFSET_COL)
            i_freq = col.get(FREQ_COL)
            i_y = col.get(Y_COL)
            i_pattern = col.get("cfg-carrier_01-data_pattern")

            def cell(row, idx):
                return row[idx] if idx is not None and idx < len(row) else ""

            for row in reader:
                supply = (cell(row, i_supply) or "").strip()
                if supply != FIXED_SUPPLY:
                    continue
                if i_pattern is not None and (cell(row, i_pattern) or "").strip() != DATA_PATTERN:
                    continue
                val = to_float(cell(row, i_y))
                if val is None:
                    continue
                offset = to_float(cell(row, i_offset))
                if offset is None:
                    continue
                freq = to_float(cell(row, i_freq))
                if freq is None:
                    continue
                packet = (cell(row, i_packet) or "").strip()
                gain = (cell(row, i_gain) or "").strip()
                pa_slices = to_float(cell(row, i_slices))
                dig_gain = to_float(cell(row, i_dig))
                band = to_float(cell(row, i_band))
                if pa_slices is None or dig_gain is None or band is None:
                    continue
                # 2026-08-14 (customer request): 5G and 6G share one "Band5G"
                # bucket, same merge as Power/DEVM/FreqAcc's band_of() -- see
                # make_bt_tx_pwr_revision2.band_of's comment for the full
                # rationale. ACP reads band as a literal CSV column (not
                # derived from frequency), so it's normalized right here.
                if band == 6:
                    band = 5.0
                pa_slices_seen.add(pa_slices)
                dig_gain_seen.add(dig_gain)
                key = (packet, gain, pa_slices, dig_gain, band)
                raw[key][offset][freq].append(val)
        print(f"  read {path.name}")
    return raw, pa_slices_seen, dig_gain_seen


@functools.lru_cache(maxsize=None)
def read_all_with_dut(bt_tx_dir: Path):
    """Same as read_all() above -- identical loop/filters, copied verbatim
    -- except each stored value is a (value, dut_id) tuple instead of a
    plain float, via dut_id_of()/DUT_RE (the same pattern every other
    BT_TX bench module already uses for this). 2026-08-25, added for the
    Raw Values sheet (add_bt_tx_raw_values_sheet.py) -- read_all() itself
    is untouched, and build_fixed_combos() below works unchanged on this
    function's output too (it only ever .extend()s whatever value list it
    is given, never inspects the value's own type)."""
    raw: dict[tuple, dict[float, dict[float, list[tuple[float, str]]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(list))
    )
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
            i_band = col.get(RAW_KEY_COLS[5])
            i_offset = col.get(OFFSET_COL)
            i_freq = col.get(FREQ_COL)
            i_y = col.get(Y_COL)
            i_pattern = col.get("cfg-carrier_01-data_pattern")

            def cell(row, idx):
                return row[idx] if idx is not None and idx < len(row) else ""

            for row in reader:
                supply = (cell(row, i_supply) or "").strip()
                if supply != FIXED_SUPPLY:
                    continue
                if i_pattern is not None and (cell(row, i_pattern) or "").strip() != DATA_PATTERN:
                    continue
                val = to_float(cell(row, i_y))
                if val is None:
                    continue
                offset = to_float(cell(row, i_offset))
                if offset is None:
                    continue
                freq = to_float(cell(row, i_freq))
                if freq is None:
                    continue
                packet = (cell(row, i_packet) or "").strip()
                gain = (cell(row, i_gain) or "").strip()
                pa_slices = to_float(cell(row, i_slices))
                dig_gain = to_float(cell(row, i_dig))
                band = to_float(cell(row, i_band))
                if pa_slices is None or dig_gain is None or band is None:
                    continue
                if band == 6:
                    band = 5.0
                pa_slices_seen.add(pa_slices)
                dig_gain_seen.add(dig_gain)
                key = (packet, gain, pa_slices, dig_gain, band)
                raw[key][offset][freq].append((val, dut))
        print(f"  read {path.name}")
    return raw, pa_slices_seen, dig_gain_seen


def build_fixed_combos(raw, pa_slices_max, dig_gain_max):
    """Filter to pa_slices==max & dig_gain==max, drop those two from the key."""
    data: dict[tuple, dict[float, dict[float, list[float]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(list))
    )
    for (packet, gain, pa_slices, dig_gain, band), offset_map in raw.items():
        if pa_slices != pa_slices_max or dig_gain != dig_gain_max:
            continue
        combo = (packet, gain, band)
        for offset, freq_map in offset_map.items():
            for freq, values in freq_map.items():
                data[combo][offset][freq].extend(values)
    return data


def combo_name(combo, pa_slices_max, dig_gain_max) -> str:
    packet, gain, band = combo
    return (
        f"packet_type({packet})_band({fmt_num(band)})_pa_supply({FIXED_SUPPLY})_"
        f"pa_slices({fmt_num(pa_slices_max)})_pa_gain({gain})_"
        f"dig_gain({fmt_num(dig_gain_max)})"
    )


def combo_title_short(combo, pa_slices_max, dig_gain_max) -> str:
    """Short plot-title format (2026-07-19 user request), same idea as
    make_bt_tx_pwr_revision2.combo_title_short: "1DH5_Band2.4_supply(0)_
    slices(7)_gain(0)_dig_gain(128)". Deliberately has NO "_offset(...)"
    suffix -- unlike Power/DEVM/FreqAcc, the ACP PPT's stats table already
    has an "Offset" header row identifying each column, so repeating the
    offset in every single chart's own title would be redundant (the exact
    kind of duplication this whole table format was introduced to remove).
    Does NOT change the PNG/XLSX FILENAME (still uses combo_name() +
    "_offset(...)", since build_ppt_acp_revision2.py's discover() regex
    parses that exact format) -- only the chart's own title text."""
    packet, gain, band = combo
    return (f"{packet}_Band{fmt_num(band)}_supply({FIXED_SUPPLY})_"
            f"slices({fmt_num(pa_slices_max)})_gain({gain})_dig_gain({fmt_num(dig_gain_max)})")


BINS = 10


def offset_bins(offset_map: dict, offset, freqs):
    """Histogram one offset's combined (all-frequency) values locally
    (auto range), then split each bin's count back out per frequency for
    stacking. Returns list of (bin_low, bin_high, {freq: count}, total)."""
    freq_map = offset_map[offset]
    combined = [v for values in freq_map.values() for v in values]
    _, edges = np.histogram(combined, bins=BINS)
    rows = []
    for i in range(BINS):
        lo, hi = edges[i], edges[i + 1]
        is_last = i == BINS - 1
        per_freq = {}
        for f in freqs:
            values = freq_map.get(f, [])
            if is_last:
                c = sum(1 for v in values if lo <= v <= hi)
            else:
                c = sum(1 for v in values if lo <= v < hi)
            per_freq[f] = c
        rows.append((lo, hi, per_freq, sum(per_freq.values())))
    return rows


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


def draw_png(offset_map: dict, title: str, xlabel: str, png_path: Path) -> None:
    """Y = Offset (-10 topmost .. +10 bottommost). Each offset owns a block
    of BINS rows: its own combined ACP values are histogrammed locally
    (auto range per offset), each bin drawn as one horizontal bar stacked
    by frequency color, with "{lo}~{hi} n={count}" floating past the bar
    end. X = Count."""
    offsets = sorted(offset_map.keys())
    freqs = sorted({f for om in offset_map.values() for f in om})
    n_freq = len(freqs)
    colors = [f"#{color_hex(fi, n_freq)}" for fi in range(n_freq)]

    fig, ax = plt.subplots(figsize=(12, max(8, 0.32 * len(offsets) * BINS)))

    max_count = 1
    y_tick_pos = []
    y_tick_labels = []
    row = 0
    for offset in offsets:
        rows = offset_bins(offset_map, offset, freqs)
        y_tick_pos.append(row + (BINS - 1) / 2)
        y_tick_labels.append(f"[{fmt_num(offset)}]")
        for lo, hi, per_freq, total in rows:
            if total > 0:
                left = 0
                for fi, f in enumerate(freqs):
                    c = per_freq[f]
                    if c <= 0:
                        continue
                    ax.barh(row, c, left=left, height=0.8, color=colors[fi],
                            edgecolor="black", linewidth=0.4)
                    left += c
                ax.text(left + 0.3, row, f"{lo:.2f}~{hi:.2f} n={total}",
                        va="center", ha="left", fontsize=6)
                max_count = max(max_count, left)
            row += 1

    for fi, f in enumerate(freqs):
        ax.barh(-1, 0, color=colors[fi], label=f"{fmt_num(f)} MHz")

    ax.set_yticks(y_tick_pos)
    ax.set_yticklabels(y_tick_labels, fontsize=9)
    ax.set_ylim(row, -1)
    ax.set_xlim(0, max_count * 1.6)
    ax.set_xlabel("Count")
    ax.set_ylabel("Offset")
    ax.grid(True, axis="x", color="0.85", linewidth=0.6)
    ax.set_axisbelow(True)
    ax.legend(loc="lower right", fontsize=8, frameon=True)
    # Finalize the axes' pixel position before sizing the title against it.
    fig.tight_layout()
    ax.set_title(title, fontsize=_fit_title_fontsize(fig, ax, title), fontweight="bold")
    fig.tight_layout()
    fig.savefig(png_path, dpi=450)
    plt.close(fig)


def local_bins(freq_map: dict, freqs):
    """Histogram one offset's combined (all-frequency) values locally
    (auto range, matching the confirmed merged-chart's per-offset
    binning), then split each bin's count back out per frequency."""
    combined = [v for values in freq_map.values() for v in values]
    _, edges = np.histogram(combined, bins=BINS)
    rows = []
    for i in range(BINS):
        lo, hi = edges[i], edges[i + 1]
        is_last = i == BINS - 1
        per_freq = {}
        for f in freqs:
            values = freq_map.get(f, [])
            if is_last:
                c = sum(1 for v in values if lo <= v <= hi)
            else:
                c = sum(1 for v in values if lo <= v < hi)
            per_freq[f] = c
        rows.append((lo, hi, per_freq, sum(per_freq.values())))
    return rows


FIXED_Y_MIN_DEFAULT = -50
FIXED_Y_MAX_DEFAULT = -20
FIXED_GRID_STEP = 10
# The PPT slide only ever shows offsets -6..6 (0 excluded) -- that's the
# window checked when deciding each combo's Y-axis min/max.
SLIDE_OFFSETS = [o for o in range(-6, 7) if o != 0]


def compute_axis_range(offset_map: dict, offsets=SLIDE_OFFSETS,
                        y_min_default=FIXED_Y_MIN_DEFAULT, y_max_default=FIXED_Y_MAX_DEFAULT):
    """UNUSED as of 2026-07-19 -- kept as the documented rollback baseline
    for the "adaptive per-combo axis" design (do not delete, see
    feedback_acp_rollback_baseline in project memory). main() now uses a
    single FIXED (FIXED_Y_MIN_DEFAULT, FIXED_Y_MAX_DEFAULT) for every combo
    instead: the customer reviewed the adaptive version and found the
    per-combo-widened range (e.g. down to -60) "too wide to see
    distribution trends," and asked for a fixed -20/-50 window instead,
    explicitly accepting that offset +/-1 data may clip above -20 (that's
    useful as a visual flag for gross failures).

    Per-combo Y-axis range from that combo's own offset -6..6 data:
      * Max: if the data rises above y_max_default, raise it just enough
        to include the data, rounded UP to the next 10 (23.7 -> 30).
        Otherwise keep the default (-20) -- matches the customer's stated
        baseline, only ever grows.
      * Min: always floor(actual_min/10)*10 -- e.g. actual min -8.7 -> -10.
        Unlike max this SHRINKS the axis whenever the data doesn't reach
        all the way down to the default min, since a too-low default wastes
        most of the frame as blank space and makes the (already thin) bars
        hard to see.
    Falls back to (y_min_default, y_max_default) if this combo has no data
    at all in the given offset window."""
    vals = [v for off in offsets for v in offset_map.get(off, {}).values() for v in v]
    if not vals:
        return y_min_default, y_max_default
    actual_min = min(vals)
    actual_max = max(vals)
    y_min = math.floor(actual_min / 10) * 10
    y_max = y_max_default if actual_max <= y_max_default else math.ceil(actual_max / 10) * 10
    return y_min, y_max


def compute_axis_range_per_offset(offset_map: dict, offset) -> tuple:
    """UNUSED as of 2026-08-15 -- kept as a documented rollback baseline
    (do not delete), same convention as compute_axis_range below.
    Superseded by compute_shared_span() (customer request: one shared
    span per whole slide/page, not one per offset column -- "한 페이지에
    들어가는 값들을 비교해보고 span을 고정").

    2026-08-14 (customer request): replaces the old whole-row FIXED axis
    -- each offset column now gets its OWN Y-range from just that offset's
    real (all-frequency-pooled) data, floor/ceil to the nearest integer.

    Unlike the rejected 2026-07-19 per-COMBO adaptive design
    (compute_axis_range above, kept as the documented rollback baseline),
    this is per OFFSET, not per whole (packet,band) row: real ACP data
    spreads radically differently by offset (far offsets like +/-6 sit
    almost entirely in -58..-8; near-carrier +/-1/+/-2 legitimately reach
    +20), so pooling all 12 offsets into one shared window either clips
    the near-carrier offsets' upper data or flattens the far offsets' tight
    cluster into a sliver -- confirmed against the real dataset, 2026-08-14.
    Keeping the axis PER OFFSET means each column stays exactly as tight
    (and as readable) as its own real distribution, avoiding the "too wide
    to see distribution trends" complaint that killed the per-combo
    version, while still showing every real data point (no more fixed
    -50/-20 clipping)."""
    vals = [v for values in offset_map[offset].values() for v in values]
    if not vals:
        return FIXED_Y_MIN_DEFAULT, FIXED_Y_MAX_DEFAULT
    return math.floor(min(vals)), math.ceil(max(vals))


ACP_SNAP_STEP = 10


def _snap_range(vmin_raw, vmax_raw, step=ACP_SNAP_STEP):
    """Floor/ceil to the nearest `step` grid, then push ONE more step out
    if a bound already sits exactly on that grid -- same rationale as
    make_bt_tx_pwr_revision2._snap_range. Used by compute_shared_span()."""
    vmin = math.floor(vmin_raw / step) * step
    vmax = math.ceil(vmax_raw / step) * step
    if vmin == vmin_raw:
        vmin -= step
    if vmax == vmax_raw:
        vmax += step
    return int(vmin), int(vmax)


def compute_shared_span(offset_map: dict, extra_range=None, specs=None, offsets=SLIDE_OFFSETS):
    """2026-08-15 (customer request, supersedes the 2026-08-14 per-offset-
    adaptive design -- compute_axis_range_per_offset above): ONE shared
    Y-axis for the WHOLE slide (every offset column on that page), from
    the union of every real Bench value shown, optionally widened by
    `extra_range` (a (min,max) tuple of real ATE data pooled the same way
    -- see make_ate_tx_acp_revision2.get_ate_slide_extent) and `specs`
    ([(lsl,usl) or None, ...] for every offset shown), snapped to the
    nearest 10 with clearance (see _snap_range). Real example that drove
    this: offset -6's min=-52, offset +1's max=1 -> shared span -60..10 --
    "한 페이지에 들어가는 값들을 비교해보고 span을 고정" (compare every
    value on one page, then pin one shared span for it)."""
    vals = [v for off in offsets for values in offset_map.get(off, {}).values() for v in values]
    if not vals and not extra_range:
        return FIXED_Y_MIN_DEFAULT, FIXED_Y_MAX_DEFAULT
    vmin_raw = min(vals) if vals else extra_range[0]
    vmax_raw = max(vals) if vals else extra_range[1]
    if extra_range:
        vmin_raw = min(vmin_raw, extra_range[0])
        vmax_raw = max(vmax_raw, extra_range[1])
    for spec in (specs or []):
        if spec is None:
            continue
        lsl, usl = spec
        if lsl is not None:
            vmin_raw = min(vmin_raw, lsl)
        if usl is not None:
            vmax_raw = max(vmax_raw, usl)
    return _snap_range(vmin_raw, vmax_raw)


def _nice_yticks(y_min: int, y_max: int, max_ticks: int = 8) -> list:
    """Integer tick list capped at max_ticks with a nice (whole-number)
    step, always including the exact y_min/y_max boundary, with a min-gap
    filter so a forced-in boundary never lands right next to a step tick --
    same pattern as make_bt_tx_pwr_revision2.py's identical block
    (2026-08-14 customer request), needed here now that the axis is
    per-offset data-driven instead of a single FIXED_GRID_STEP=10 window."""
    span = max(1, y_max - y_min)
    step = max(1, math.ceil(span / max_ticks))
    n_steps = max(1, round(span / step))
    yticks = sorted({int(round(y_min + i * step)) for i in range(n_steps + 1)} | {y_min, y_max})
    if len(yticks) > max_ticks:
        min_gap = (y_max - y_min) / max_ticks * 0.5
        inner = [t for t in yticks if t not in (y_min, y_max)
                 and (t - y_min) >= min_gap and (y_max - t) >= min_gap]
        keep_n = max(0, max_ticks - 2)
        if keep_n > 0 and inner:
            idx_step = len(inner) / keep_n
            thinned = sorted({inner[min(len(inner) - 1, round(i * idx_step))] for i in range(keep_n)})
        else:
            thinned = []
        yticks = sorted({y_min, y_max} | set(thinned))
    return yticks


def draw_png_single_offset(freq_map: dict, title: str, xlabel: str, png_path: Path,
                            y_min: int = FIXED_Y_MIN_DEFAULT, y_max: int = FIXED_Y_MAX_DEFAULT,
                            spec=None) -> None:
    """One offset's own chart. Bins use LOCAL auto-range (BINS=10, matching
    the confirmed merged-chart's fine per-offset detail), plotted at their
    true ACP value on a Y axis FIXED to [y_min, y_max] with gridlines every
    FIXED_GRID_STEP, at the SAME figure height as the (unfixed-axis)
    local2x baseline -- so the local cluster will render small/thin within
    the wide fixed frame; that's the expected tradeoff of a shared fixed
    scale at this figure size. y_min/y_max are computed per-combo (see
    compute_axis_range) from that combo's own offset -6..6 data, so the
    chart is neither entirely blank (max too low) nor mostly wasted blank
    space (min too low)."""
    freqs = sorted(freq_map.keys())
    n_freq = len(freqs)
    colors = [f"#{color_hex(fi, n_freq)}" for fi in range(n_freq)]

    rows = local_bins(freq_map, freqs)
    baseline_height = 6
    fig, ax = plt.subplots(figsize=(9, baseline_height * 2))

    max_count = 1
    for lo, hi, per_freq, total in rows:
        if total <= 0:
            continue
        bin_h = hi - lo
        left = 0
        for fi, f in enumerate(freqs):
            c = per_freq[f]
            if c <= 0:
                continue
            ax.barh(lo, c, left=left, height=bin_h * 0.9, align="edge",
                    color=colors[fi], edgecolor="black", linewidth=0.4)
            left += c
        max_count = max(max_count, left)

    ax.set_ylim(y_min, y_max)
    ax.set_yticks(_nice_yticks(y_min, y_max))
    ax.set_xlim(0, max_count * 1.6)
    ax.set_xlabel("Count")
    ax.set_ylabel(xlabel)
    ax.grid(True, axis="x", color="0.85", linewidth=0.6)
    ax.set_axisbelow(True)
    # Finalize the axes' pixel position before sizing the title against it.
    fig.tight_layout()
    ax.set_title(title, fontsize=_fit_title_fontsize(fig, ax, title), fontweight="bold")

    if spec is not None:
        spec_min, spec_max = spec
        # va="top" visually overlaps the axhline in tall/wide-range figures
        # (confirmed matplotlib rendering quirk, 2026-07-17) -- use
        # va="bottom" uniformly with a manual data-space pad instead.
        pad = (y_max - y_min) * 0.035
        for spec_val, spec_label in ((spec_min, "LSL"), (spec_max, "USL")):
            if spec_val is None or not (y_min <= spec_val <= y_max):
                continue
            # 2026-07-20 customer request: Limit line is blue dashed (was red
            # solid-looking/thin) -- red dashed is reserved for a "Target"
            # value if one is ever added, same linewidth as this Limit line.
            ax.axhline(y=spec_val, color="blue", linestyle="--", linewidth=2.5, zorder=5)
            above = spec_val >= (y_min + y_max) / 2
            text_y = spec_val + pad if above else spec_val - pad
            ax.text(ax.get_xlim()[1] * 0.99, text_y, f"{spec_label} {spec_val:g}",
                    color="blue", fontsize=10, fontweight="bold", ha="right",
                    va="bottom", clip_on=True)

    fig.tight_layout()
    fig.savefig(png_path, dpi=450)
    plt.close(fig)


def draw_legend_png(freqs, png_path: Path) -> None:
    """Standalone legend (color swatch + frequency label per row, no axes)
    -- generated once per (packet,band) combo since every one of that
    combo's 13 offset charts shares the same frequency/color set. The
    customer found the legend repeated inside every tiny offset chart too
    cluttered; this single image gets placed once per PPT slide instead."""
    n_freq = len(freqs)
    colors = [f"#{color_hex(fi, n_freq)}" for fi in range(n_freq)]

    fig, ax = plt.subplots(figsize=(2.4, max(1.2, 0.4 * n_freq)))
    ax.axis("off")
    for fi, f in enumerate(freqs):
        y = n_freq - 1 - fi
        ax.add_patch(plt.Rectangle((0, y), 0.3, 0.8, color=colors[fi]))
        ax.text(0.4, y + 0.4, f"{fmt_num(f)} MHz", va="center", ha="left",
                fontsize=14, fontweight="bold")
    ax.set_xlim(0, 2.2)
    ax.set_ylim(0, n_freq)
    fig.tight_layout()
    fig.savefig(png_path, dpi=450, transparent=True)
    plt.close(fig)


def write_xlsx(xlsx_path: Path, title: str, xlabel: str, offset_map: dict) -> None:
    """Mirror the PNG: a stacked horizontal BarChart with categories =
    offset+bin range, one series per frequency (stacked count)."""
    offsets = sorted(offset_map.keys())
    freqs = sorted({f for om in offset_map.values() for f in om})
    n_freq = len(freqs)

    all_rows = []  # (category_label, {freq: count}, total, raw_values)
    for offset in offsets:
        rows = offset_bins(offset_map, offset, freqs)
        freq_map = offset_map[offset]
        for lo, hi, per_freq, total in rows:
            label = f"[{fmt_num(offset)}] {lo:.3f} ~ {hi:.3f}"
            all_rows.append((label, per_freq, total))

    wb = Workbook()
    ws = wb.active
    ws.title = "Detail"
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    chart_row_span = max(20, 2 * len(offsets))
    data_header_row = 1 + chart_row_span

    ws.cell(data_header_row, 1, f"[Offset] {xlabel} Bin")
    for fi, f in enumerate(freqs):
        ws.cell(data_header_row, 2 + fi, f"{fmt_num(f)} MHz")
    ws.cell(data_header_row, 2 + n_freq, "Total N")
    for i, (label, per_freq, total) in enumerate(all_rows, data_header_row + 1):
        ws.cell(i, 1, label)
        for fi, f in enumerate(freqs):
            ws.cell(i, 2 + fi, per_freq[f])
        ws.cell(i, 2 + n_freq, total)
    for cell in ws[data_header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 28
    for j in range(2, 3 + n_freq):
        ws.column_dimensions[ws.cell(1, j).column_letter].width = 12

    last_row = data_header_row + len(all_rows)
    chart = BarChart()
    chart.type = "bar"  # horizontal
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = title
    chart.style = 10
    chart.height = max(20, 0.5 * len(all_rows))
    chart.width = 30
    chart.x_axis.title = f"[Offset] {xlabel} Bin"
    chart.y_axis.title = "Count"
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.y_axis.majorGridlines = ChartLines()
    chart.y_axis.majorGridlines.spPr = GraphicalProperties()
    chart.y_axis.majorGridlines.graphicalProperties.line.solidFill = "808080"
    chart.y_axis.majorGridlines.graphicalProperties.line.width = 12700
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.gapWidth = 0

    cats_ref = Reference(ws, min_col=1, min_row=data_header_row + 1, max_row=last_row)
    for fi, f in enumerate(freqs):
        data_ref = Reference(ws, min_col=2 + fi, min_row=data_header_row, max_row=last_row)
        chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    for fi in range(n_freq):
        chart.series[fi].graphicalProperties.solidFill = color_hex(fi, n_freq)

    ws.add_chart(chart, ws.cell(1, 3 + n_freq).coordinate)

    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)


def write_xlsx_single_offset(xlsx_path: Path, title: str, xlabel: str, freq_map: dict) -> None:
    """Mirror the single-offset PNG: a stacked horizontal BarChart with
    categories = local-auto-range bin label, one series per frequency
    (stacked count), plus a raw-values column per frequency."""
    freqs = sorted(freq_map.keys())
    n_freq = len(freqs)
    rows = local_bins(freq_map, freqs)

    wb = Workbook()
    ws = wb.active
    ws.title = "Detail"
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    chart_row_span = 22
    data_header_row = 1 + chart_row_span

    ws.cell(data_header_row, 1, f"{xlabel} Bin")
    for fi, f in enumerate(freqs):
        ws.cell(data_header_row, 2 + fi, f"{fmt_num(f)} MHz")
    ws.cell(data_header_row, 2 + n_freq, "Total N")
    for i, (lo, hi, per_freq, total) in enumerate(rows, data_header_row + 1):
        ws.cell(i, 1, f"{lo:.3f} ~ {hi:.3f}")
        for fi, f in enumerate(freqs):
            ws.cell(i, 2 + fi, per_freq[f])
        ws.cell(i, 2 + n_freq, total)
    for cell in ws[data_header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 22
    for j in range(2, 3 + n_freq):
        ws.column_dimensions[ws.cell(1, j).column_letter].width = 12

    raw_col_start = 3 + n_freq + 1
    for fi, f in enumerate(freqs):
        col = raw_col_start + fi
        ws.cell(data_header_row, col, f"{fmt_num(f)} MHz Raw")
        for i, v in enumerate(freq_map.get(f, []), data_header_row + 1):
            ws.cell(i, col, v)
        ws.column_dimensions[ws.cell(1, col).column_letter].width = 14

    last_row = data_header_row + len(rows)
    chart = BarChart()
    chart.type = "bar"  # horizontal
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = title
    chart.style = 10
    chart.height = 14
    chart.width = 26
    chart.x_axis.title = f"{xlabel} Bin"
    chart.y_axis.title = "Count"
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.y_axis.majorGridlines = ChartLines()
    chart.y_axis.majorGridlines.spPr = GraphicalProperties()
    chart.y_axis.majorGridlines.graphicalProperties.line.solidFill = "808080"
    chart.y_axis.majorGridlines.graphicalProperties.line.width = 12700
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.gapWidth = 0

    cats_ref = Reference(ws, min_col=1, min_row=data_header_row + 1, max_row=last_row)
    for fi, f in enumerate(freqs):
        data_ref = Reference(ws, min_col=2 + fi, min_row=data_header_row, max_row=last_row)
        chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    for fi in range(n_freq):
        chart.series[fi].graphicalProperties.solidFill = color_hex(fi, n_freq)

    ws.add_chart(chart, ws.cell(1, raw_col_start + n_freq + 1).coordinate)

    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--bt-tx-dir", type=Path, default=PATHS.bench_dir)
    p.add_argument("--out-dir", type=Path, default=PATHS.result_xlsx_dir / "acp_revision2")
    p.add_argument("--png-dir", type=Path, default=PATHS.result_png_dir / "acp_revision2")
    p.add_argument("--legend-dir", type=Path,
                   default=PATHS.result_png_dir / "acp_revision2_legend")
    p.add_argument("--limit", type=int, default=0,
                   help="Generate only the first N (combo,offset) pairs (0 = all).")
    p.add_argument("--count-only", action="store_true",
                   help="Only print combo/sample-size audit, generate nothing.")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    print("Parsing ACP spec limits...")
    import bt_tx_acp_spec_lookup as spec_lookup
    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    print("Reading BT_Tx CSVs (pa_supply=0 only)...")
    raw, pa_slices_seen, dig_gain_seen = read_all(args.bt_tx_dir)
    pa_slices_max = max(pa_slices_seen)
    dig_gain_max = max(dig_gain_seen)
    print(f"pa_slices seen: {sorted(pa_slices_seen)} -> fixed at max {fmt_num(pa_slices_max)}")
    print(f"dig_gain seen: {sorted(dig_gain_seen)} -> fixed at max {fmt_num(dig_gain_max)}")

    data = build_fixed_combos(raw, pa_slices_max, dig_gain_max)
    combos = list(data.keys())
    pairs = [(combo, offset) for combo in combos for offset in sorted(data[combo].keys())]
    print(f"combos (packet,gain,band): {len(combos)}  total (combo,offset) pairs: {len(pairs)}")

    if args.count_only:
        for combo in combos[:10]:
            offset_map = data[combo]
            n_offsets = len(offset_map)
            n_total = sum(len(vs) for om in offset_map.values() for vs in om.values())
            n_freqs = len({f for om in offset_map.values() for f in om})
            print(f"  {combo}: offsets={n_offsets} freqs={n_freqs} total_samples={n_total}")
        return

    if args.limit > 0:
        pairs = pairs[: args.limit]
        print(f"Limiting to first {len(pairs)} pair(s)")

    args.out_dir.mkdir(parents=True, exist_ok=True)
    args.png_dir.mkdir(parents=True, exist_ok=True)
    args.legend_dir.mkdir(parents=True, exist_ok=True)

    legend_combos = combos[: args.limit] if 0 < args.limit < len(combos) else combos
    for combo in legend_combos:
        name = combo_name(combo, pa_slices_max, dig_gain_max)
        freqs = sorted({f for om in data[combo].values() for f in om})
        draw_legend_png(freqs, args.legend_dir / f"{safe_filename(name)}_legend.png")
    print(f"Legends -> {args.legend_dir} ({len(legend_combos)})")

    # 2026-08-15 (customer request): ONE shared Y-axis for the WHOLE slide
    # (every offset column), superseding the 2026-08-14 per-offset-
    # adaptive design -- see compute_shared_span's docstring for the full
    # rationale. Lazy/deferred import (not at module top-level) to avoid a
    # circular import -- make_ate_tx_acp_revision2 imports this module as
    # ACP_CHART; by the time main() runs, this module is already fully
    # initialized, so Python just returns it from sys.modules.
    import make_ate_tx_acp_revision2 as ate_acp_mod
    ate_idx = None
    if PATHS.extract_mode != "bench":
        try:
            ate_idx = ate_acp_mod.parse_ate_acp_records(ate_acp_mod.ATE_CSV)
        except (FileNotFoundError, OSError):
            ate_idx = None
    axis_range_by_combo = {}
    for combo in combos:
        packet = combo[0]
        specs = [spec_lookup.lookup(spec_table, packet, int(off)) for off in SLIDE_OFFSETS]
        extra_range = ate_acp_mod.get_ate_slide_extent(combo, data[combo], ate_idx)
        axis_range_by_combo[combo] = compute_shared_span(data[combo], extra_range, specs)
    print(f"Y-axis: shared per-slide range for all {len(combos)} combo(s)")

    jobs = []
    for combo, offset in pairs:
        name = combo_name(combo, pa_slices_max, dig_gain_max)
        off_label = fmt_num(offset)
        fname = safe_filename(f"{name}_offset({off_label})")
        title = combo_title_short(combo, pa_slices_max, dig_gain_max)
        freq_map = data[combo][offset]
        y_min, y_max = axis_range_by_combo[combo]
        packet = combo[0]
        spec = spec_lookup.lookup(spec_table, packet, int(offset))
        jobs.append((freq_map, title, args.png_dir / f"{fname}.png", y_min, y_max, spec,
                     args.out_dir / f"{fname}.xlsx"))

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
    freq_map, title, png_path, y_min, y_max, spec, xlsx_path = job
    draw_png_single_offset(freq_map, title, "ACP (dBm)", png_path, y_min=y_min, y_max=y_max, spec=spec)
    if PATHS.write_xlsx:
        write_xlsx_single_offset(xlsx_path, title, "ACP (dBm)", freq_map)


if __name__ == "__main__":
    main()
