"""Shared BT_RX helpers: Bench sweep indexing, ATE item parsing/matching,
and the common chart-drawing function -- imported by all 3 make_bt_rx_*.py
scripts (RSSI / Sensitivity-PER / Sensitivity-BER), same shared-infrastructure
exception as bt_tx_power_spec_lookup.py (parsing/matching logic, not
per-metric chart logic -- see md/CLAUDE.md's "self-contained scripts"
convention).

## Dataset shape (2026-08-14, first BT_RX build)

10 DUTs (not 40 like every other project here) -- Bench_40pcs/Rx has 10 CSVs,
and the shared ate_log_path CSV's BTRX_* rows also carry exactly 10 DUT
columns (confirmed: same ATE_10pcs_Data.csv file BT_TX itself now reads).

Each Bench CSV is a per-DUT POWER SWEEP per (packet_type, band, sub_band,
channel): ~64 rows stepping from -20dBm down to -103dBm (5dB steps down to
-90, then 0.25dB steps -91..-103) recording rssi_avg-dBm/per/ber at each
power level, used to characterize the DUT's own sensitivity curve. This is
DIFFERENT from every other project's Bench data (one value per DUT per
condition) -- there is no single "the" Bench value for a (packet,channel)
condition until you pick a power level.

## ATE item format

`BTRX{band}_{packet}_JTAG_{metric}_CH{ch}_{pwr}_X_X_NV_VXXX`, e.g.
`BTRX0_1DH5_JTAG_AVGRSSI_CH000_M92P8_X_X_NV_VXXX`. metric in
{AVGRSSI, PER, BER} for this project's 3 requested items (many other metrics
exist in the log -- POLLDONE, TOTPKTS, GPKTS, etc. -- not used here).

`{pwr}` encodes a FIXED dBm test level identically to BTTX's power labels:
"M" + abs(value), decimal point rendered as "P" with trailing zeros
stripped (`M92P8` -> -92.8, `M17` -> -17.0, `M7` -> -7.0). This is a
PRODUCTION test level chosen by the test engineer -- it does NOT
necessarily land exactly on Bench's own 0.25dB sweep grid, so matching is
nearest-neighbor within POWER_TOL_DBM, not exact (see nearest_bench_row()).
Confirmed against real data: `M92P8` (-92.8) sits within 0.05dB of Bench's
own grid (nearest points -92.75/-93.0) -- a clean match. Some ATE items
(e.g. `M17` = -17.0dBm) fall OUTSIDE Bench's entire swept range (-20..-103)
for every one of the 10 DUTs checked -- these render "No Bench match"
exactly like an unmapped combo elsewhere in this project, not a bug.
**This tolerance is a first-pass assumption, not confirmed with the
customer yet** -- flag any oddly-large per-combo delta printed by the
audit output back to the user before trusting a full regeneration.

`{band}` = Bench's own `cfg-carrier_01-sub_band` (as an int) and `{ch}` =
Bench's own `cfg-carrier_01-channel` (as an int, e.g. `CH078` <-> channel
"78") -- DIRECTLY, no offset arithmetic. Verified against every packet type
in the 2026-08-14 dataset (2.4G sub_band=0, 5G sub_band in {1,2}, 6G
sub_band in {4,7}) by cross-referencing Bench's own (band, sub_band,
channel, frequency-MHz) tuples against every ATE (band, packet, channel)
triple seen for AVGRSSI -- every one lines up exactly. This is SIMPLER than
BTTX's own NEW_BAND_MIN_FREQ scheme (freq = band_min + channel offset) --
BT_RX's ATE item apparently encodes sub_band/channel verbatim rather than a
computed frequency offset. Packet names also pass through verbatim, no
alias table needed (unlike BTTX's "2M"->LE2M-style aliases).

Sensitivity-BER is customer-scoped to 1DH5 packets only (see the customer's
own Test Item Matching table) -- make_bt_rx_sensitivity_ber_revision1.py
enforces that filter itself, not this shared module (a BER-metric ATE item
for a non-1DH5 packet, if one ever appears in a future ATE pull, is not
filtered out here -- keep the metric-key scripts symmetric with how BT_TX's
per-metric scripts each own their own combo-set decisions).
"""

from __future__ import annotations

import csv
import functools
import math
import re
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_rx_paths

PATHS = bt_rx_paths.get_paths("BT_RX")

DUT_RE = re.compile(r"#(\d+)")
POWER_TOL_DBM = 1.0  # see module docstring -- first-pass assumption

# -- PPT grid geometry, single source of truth for both the chart scripts
# (figure_size_for below) and build_ppt_bt_rx_revision1.py's own table/image
# layout (which imports these same constants rather than keeping its own
# copy) -- 2026-08-14, real bug fixed: every make_bt_rx_*.py/make_ate_rx_*.py
# script drew every chart at a FIXED figsize=(9,6) regardless of how many
# columns that packet_type's slide would end up with. A packet with only 1
# matched combo (e.g. 3DH5) gets ONE image stretched across almost the
# entire ~12.5in-wide row -- PowerPoint stretches a 9x6in (1.5:1 aspect)
# source into a ~12.5x2.1in (~6:1 aspect) cell, badly distorting the text
# (same "squish factor" class of bug BT_TX/UWB already fixed for their own
# multi-column grids, see md/CLAUDE.md). Mirrors BT_TX's make_bt_tx_pwr_
# revision2.py figure_size_for(), just simpler: BT_RX has no ruler column
# (see build_ppt_bt_rx_revision1.py's module docstring) and only two
# possible table row-counts (4 with no spec row, 5 with one USL/LSL row --
# BT_RX's ATE items are always one-sided, never both).
_SLIDE_WIDTH_IN = 13.338542213473316
_LEFT_MARGIN = 0.05
_RIGHT_MARGIN = 0.05
_LABEL_WIDTH = 0.7
_CELL_GAP_H = 0.03
_GRID_WIDTH = _SLIDE_WIDTH_IN - _LEFT_MARGIN - _RIGHT_MARGIN - _LABEL_WIDTH

_TABLE_TOP = 1.05
_TABLE_ROW_H = 0.2
_HEADER_ROW_H = 0.45
_TABLE_PLOT_GAP = 0.35
_IMAGE_GAP = 0.15
_CONTENT_AREA_TOP = 1.5908508311461067
_CONTENT_AREA_BOTTOM = _CONTENT_AREA_TOP + 5.240927384076991


def figure_size_for(n_cols: int, has_spec: bool, show_power_level: bool = False) -> tuple[float, float]:
    """(width, height) in inches for a chart that will land in an n_cols-wide
    PPT row -- pass the SAME n_cols/has_spec/show_power_level
    build_ppt_bt_rx_revision1.py's add_packet_slide() will use for that
    packet's slide, or the resulting PNG will be stretched into a
    differently-shaped cell than it was drawn for (the exact bug this
    function fixes).

    show_power_level (2026-08-15, RSSI only): the table gets an extra
    "Power Level" row for RSSI (see _add_interleaved_table's docstring) --
    must be counted here too or this function under-estimates table_h,
    reintroducing the same "image positioned too high, overlaps the
    table's own rows" bug HEADER_ROW_H itself was originally added to fix."""
    n_cols = max(1, n_cols)
    cell_width = _GRID_WIDTH / n_cols if n_cols == 1 else (_GRID_WIDTH - (n_cols - 1) * _CELL_GAP_H) / n_cols
    # cond header + Bench/ATE subheader + Mean+-Std + Min/Max [+ USL/LSL] [+ Power Level]
    n_rows = 4 + (1 if has_spec else 0) + (1 if show_power_level else 0)
    table_h = _HEADER_ROW_H + (n_rows - 1) * _TABLE_ROW_H
    plots_top = _TABLE_TOP + table_h + _TABLE_PLOT_GAP
    image_h = (_CONTENT_AREA_BOTTOM - plots_top - _IMAGE_GAP) / 2
    return (cell_width, image_h)


def n_cols_by_packet(jobs) -> dict:
    """{packet: count} -- the PPT column count each packet_type's slide will
    have, i.e. how many matched combos share that packet (mirrors
    build_ppt_bt_rx_revision1.py's own `by_packet` grouping exactly). Pass
    the FULL job list (before any --limit slicing) so a --limit sample run
    still sizes each figure for the column count a real full run would
    produce -- same reasoning as BT_TX's identical precompute-from-full-set
    comment in make_bt_tx_pwr_revision2.py."""
    return Counter(job["packet"] for job in jobs)


def jobs_have_spec(jobs) -> bool:
    """Whether ANY job in this metric's full job list carries a spec --
    determines the PPT table's row count (4 vs 5), which figure_size_for()
    needs. In practice this is constant across a whole metric (BT_RX's ATE
    Upper Limit is the same for every row of a given metric), not per-job,
    but checking the full list is more robust than assuming that."""
    return any(job.get("spec") is not None for job in jobs)

BENCH_COLS = [
    "cfg-carrier_01-packet_type",
    "cfg-carrier_01-sub_band",
    "cfg-carrier_01-channel",
    "meas-rx_power_level-dBm",
    "rssi_avg-dBm",
    # 2026-08-14: "per-%" (percent, 0-100), not "per" (fraction, 0-1) --
    # ATE's own raw PER values are themselves percent-scale (see
    # METRIC_VALUE_SCALE's docstring below), so pulling Bench's per-%
    # column directly means NEITHER side needs a value conversion, only
    # this one column choice. Bench's own "per-%" is exactly 100x "per"
    # (confirmed: e.g. a real row read per=0.03/per-%=3.0), so this is
    # not a different measurement, just the same one already un-scaled.
    "per-%",
    "ber",
]

ATE_ITEM_RE = re.compile(
    r"^BTRX(?P<band>\d+)_(?P<packet>[A-Za-z0-9]+)_JTAG_(?P<metric>[A-Z]+)_"
    r"CH(?P<ch>\d+)_(?P<pwr>[A-Za-z0-9]+)_X_X_NV_VXXX$"
)


def to_float(text: str):
    text = (text or "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def dut_id_of(path: Path) -> str:
    m = DUT_RE.search(path.name)
    return f"#{m.group(1)}" if m else path.stem


def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:\[\]<>|"]', "_", name)


def decode_power_token(tok: str):
    """"M92P8" -> -92.8, "M17" -> -17.0, "M7" -> -7.0. None if malformed."""
    if not tok.startswith("M"):
        return None
    body = tok[1:]
    try:
        if "P" in body:
            int_part, frac_part = body.split("P", 1)
            return -float(f"{int_part}.{frac_part}")
        return -float(body)
    except ValueError:
        return None


@functools.lru_cache(maxsize=None)
def read_bench(bench_dir: Path):
    """-> {(packet, sub_band_int, channel_int): [(dut_id, power_level, rssi, per, ber), ...]}

    Every sweep row across all 10 DUT files (rows with a blank
    meas-rx_power_level-dBm -- the per-DUT computed sensitivity SUMMARY rows
    -- are skipped; those aren't part of the raw sweep grid this matches
    against)."""
    data: dict[tuple, list[tuple]] = defaultdict(list)
    # 2026-08-25: rglob (not glob) + an exact "*_cns_temp_data.csv" suffix,
    # not a bare "*.csv" -- the newer nested Bench pull (post-2026-08-24)
    # puts each DUT's real data file under its own "#NN_<serial>/" subfolder
    # alongside an .xlsx companion and a "summary/summary_<timestamp>.csv"
    # metadata file that must NOT be swept in as if it were a DUT sweep file
    # (same fix already applied to every BT_TX Bench module, e.g.
    # make_bt_tx_pwr_revision2.py's read_all()). The "not startswith('.')"
    # guard is a second, independent layer: the exact-suffix match already
    # excludes a stray ".DS_Store" (macOS junk, confirmed present in a real
    # #NN folder) on its own, but this also protects against some other
    # dotfile that happens to end in that suffix in a future pull. Verified
    # backward-compatible with the OLDER flat Bench_40pcs/Rx layout: those
    # real filenames (e.g. "..._#4_cns_temp_data.csv") already end in
    # "_cns_temp_data.csv", so rglob-with-this-suffix matches them exactly
    # the same as before (glob("*.csv") also matched them, but so does this).
    files = sorted(
        p for p in bench_dir.rglob("*_cns_temp_data.csv")
        if not p.name.startswith(".")
    )
    for path in files:
        dut = dut_id_of(path)
        with path.open(encoding="utf-8-sig", newline="") as fh:
            reader = csv.reader(fh)
            header = next(reader)
            col = {name: i for i, name in enumerate(header)}
            idx = [col.get(c) for c in BENCH_COLS]

            def cell(row, i):
                return row[i] if i is not None and i < len(row) else ""

            for row in reader:
                pwr = to_float(cell(row, idx[3]))
                if pwr is None:
                    continue
                packet = (cell(row, idx[0]) or "").strip()
                sub_band = to_float(cell(row, idx[1]))
                channel = to_float(cell(row, idx[2]))
                if sub_band is None or channel is None:
                    continue
                rssi = to_float(cell(row, idx[4]))
                per = to_float(cell(row, idx[5]))
                ber = to_float(cell(row, idx[6]))
                key = (packet, int(sub_band), int(channel))
                data[key].append((dut, pwr, rssi, per, ber))
        print(f"  read {path.name}")
    return data


def nearest_bench_rows(rows, target_pwr, tol=POWER_TOL_DBM):
    """Per DUT, pick the sweep row nearest target_pwr (within tol). ->
    {dut_id: (delta, pwr, rssi, per, ber)}, one entry per DUT that has a row
    within tolerance (a DUT can appear more than once in `rows` -- one per
    sweep step -- so this also collapses to the single closest per DUT).
    `pwr` (2026-08-15) is the ACTUAL matched sweep-grid power (e.g. -86.85),
    distinct from `target_pwr` (the ATE item's own nominal test power, e.g.
    -86.9) -- every DUT sweeps the identical grid (confirmed, see module
    docstring), so this is the same value for every DUT matched here; kept
    per-DUT anyway rather than returned separately, to stay a single
    self-contained return value."""
    best: dict[str, tuple] = {}
    for dut, pwr, rssi, per, ber in rows:
        delta = abs(pwr - target_pwr)
        if delta > tol:
            continue
        cur = best.get(dut)
        if cur is None or delta < cur[0]:
            best[dut] = (delta, pwr, rssi, per, ber)
    return best


def matched_bench_pwr(matched):
    """The actual matched sweep-grid power from a nearest_bench_rows()
    result (any DUT's -- see that function's docstring for why they're all
    the same value), or None if nothing matched."""
    for v in matched.values():
        return v[1]
    return None


def _values_from_row(row):
    # Same no-read sentinel handling as make_ate_tx_pwr_revision2.py's
    # _values_from_row -- ATE logs mix in 9.91e+37 for a no-read cell.
    out = []
    for v in row[3:]:
        v = v.strip()
        if v in ("", "NA"):
            continue
        try:
            f = float(v)
        except ValueError:
            continue
        if abs(f) < 1e6:
            out.append(f)
    return out


def _to_float_or_none(text: str):
    text = (text or "").strip()
    if not text or text.upper() == "NA":
        return None
    return float(text)


# 2026-08-14 (customer request, "TX 그리는 방법이랑 동일하게" -- follow the same
# per-item ATE Lower/Upper Limit convention BT_TX's make_ate_tx_devm_99pct_
# revision2.py uses, rather than a Bench spec-sheet lookup). Every matched
# BTRX_* row already carries its own Upper Limit (row[1]) -- confirmed
# constant across every single row for a given metric (PER always "10",
# BER always "0.1", AVGRSSI always blank), matching the "Rx Sensitivity
# Specification" sheet's own universal "10% PER" criterion.
#
# Two independent unit facts, confirmed empirically (see bt_rx_ate_lookup
# module history / conversation, not guessed):
#  - PER: ATE's raw per-DUT VALUES are themselves in PERCENT units (e.g. a
#    real matched row read 0,0,0,0,0,0,1,0,2,0). 2026-08-14 (later): rather
#    than converting ATE's values down to fraction, BENCH_COLS now pulls
#    Bench's own `per-%` column (also percent, 0-100) instead of `per`
#    (fraction) -- so PER needs NO value conversion at all, VALUE_SCALE=1.0;
#    its Upper Limit ("10") is already the same percent scale too, so
#    LIMIT_SCALE=1.0 as well. Simpler than the original fraction-side
#    conversion and equally correct (per-% is just per*100, confirmed
#    directly against real rows) -- this is a column CHOICE, not a
#    different measurement.
#  - BER: ATE's raw per-DUT VALUES are ALREADY a small fraction matching
#    Bench's `ber` column directly (e.g. 0.00073746) -- VALUE_SCALE=1.0,
#    no conversion. Its own Upper Limit ("0.1") is STILL entered in percent
#    despite the value column being fraction-scale (0.1 -> 0.1% -> 0.001
#    fraction, NOT 0.1 as a raw fraction -- a raw-fraction 0.1 would mean
#    "10% BER allowed", nonsensically loose next to real measured values
#    clustering at ~0.0007-0.0015, i.e. right up against a 0.001 boundary).
#    LIMIT_SCALE=0.01 converts BER's limit onto its own value column's
#    fraction scale. Bench has no analogous `ber-%`-for-values swap since
#    BER's value scale was never the mismatched side to begin with.
METRIC_VALUE_SCALE = {"AVGRSSI": 1.0, "PER": 1.0, "BER": 1.0}
METRIC_LIMIT_SCALE = {"AVGRSSI": 1.0, "PER": 1.0, "BER": 0.01}


@functools.lru_cache(maxsize=None)
def parse_ate_items(ate_csv: Path, metric: str):
    """-> {(packet, band_int, ch_int): [{"pwr": float, "item": str,
    "values": [float,...], "spec": (lsl, usl)|None}, ...]}

    Filters strictly to BTRX_* item names carrying the given metric
    (AVGRSSI/PER/BER). `values` and `spec` are both scaled per
    METRIC_VALUE_SCALE/METRIC_LIMIT_SCALE (see their docstring) so they're
    directly comparable to Bench's own per/ber fraction columns."""
    value_scale = METRIC_VALUE_SCALE.get(metric, 1.0)
    limit_scale = METRIC_LIMIT_SCALE.get(metric, 1.0)
    idx: dict[tuple, list[dict]] = defaultdict(list)
    with ate_csv.open(encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        next(reader)
        for row in reader:
            name = row[0]
            if not name.startswith("BTRX"):
                continue
            m = ATE_ITEM_RE.match(name)
            if not m or m.group("metric") != metric:
                continue
            pwr = decode_power_token(m.group("pwr"))
            if pwr is None:
                continue
            values = _values_from_row(row)
            if not values:
                continue
            values = [v * value_scale for v in values]
            ul = _to_float_or_none(row[1] if len(row) > 1 else "")
            ll = _to_float_or_none(row[2] if len(row) > 2 else "")
            spec = None
            if ul is not None or ll is not None:
                spec = (ll * limit_scale if ll is not None else None,
                        ul * limit_scale if ul is not None else None)
            key = (m.group("packet"), int(m.group("band")), int(m.group("ch")))
            idx[key].append({"pwr": pwr, "item": name, "values": values, "spec": spec})
    return idx


# 2026-08-25 (additive-only, mirrors every BT_TX ATE module's own
# parse_..._per_dut() sibling, e.g. make_ate_tx_pwr_revision2.py's
# parse_ate_records_per_dut): parse_ate_items() itself is left completely
# untouched -- this captures the ATE CSV's own header row (the real DUT
# serial numbers, header[3:]) once and zips it with each data row's
# row[3:] to build a {serial: value} dict per item, instead of
# _values_from_row()'s flat, identity-discarding list. Used only by the
# new Raw Values / per-config summary feature (add_bt_rx_raw_values_sheet.py
# / update_bt_rx_summary_by_vendor.py), never by the existing make_bt_rx_*.py
# / make_ate_rx_*.py scripts.
def _values_by_serial_from_row(row, header):
    out = {}
    for i in range(3, len(row)):
        serial = header[i] if i < len(header) else None
        if not serial:
            continue
        v = row[i].strip()
        if v in ("", "NA"):
            continue
        try:
            f = float(v)
        except ValueError:
            continue
        if abs(f) < 1e6:
            out[serial] = f
    return out


@functools.lru_cache(maxsize=None)
def parse_ate_items_per_dut(ate_csv: Path, metric: str):
    """-> {(packet, band_int, ch_int): [{"pwr": float, "item": str,
    "by_serial": {serial: float, ...}, "spec": (lsl, usl)|None}, ...]}

    Same filtering/scaling rules as parse_ate_items (that function is left
    untouched -- this is an additive sibling, not a replacement), except
    each record carries a {serial: value} dict (keyed by the ATE CSV's own
    header[3:] real DUT serials) instead of a flat, identity-discarding
    `values` list."""
    value_scale = METRIC_VALUE_SCALE.get(metric, 1.0)
    limit_scale = METRIC_LIMIT_SCALE.get(metric, 1.0)
    idx: dict[tuple, list[dict]] = defaultdict(list)
    with ate_csv.open(encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        header = next(reader)
        for row in reader:
            name = row[0]
            if not name.startswith("BTRX"):
                continue
            m = ATE_ITEM_RE.match(name)
            if not m or m.group("metric") != metric:
                continue
            pwr = decode_power_token(m.group("pwr"))
            if pwr is None:
                continue
            by_serial = _values_by_serial_from_row(row, header)
            if not by_serial:
                continue
            by_serial = {s: v * value_scale for s, v in by_serial.items()}
            ul = _to_float_or_none(row[1] if len(row) > 1 else "")
            ll = _to_float_or_none(row[2] if len(row) > 2 else "")
            spec = None
            if ul is not None or ll is not None:
                spec = (ll * limit_scale if ll is not None else None,
                        ul * limit_scale if ul is not None else None)
            key = (m.group("packet"), int(m.group("band")), int(m.group("ch")))
            idx[key].append({"pwr": pwr, "item": name, "by_serial": by_serial, "spec": spec})
    return idx


def nice_step(span, target_ticks=8):
    """'Nice' gridline step (1/2/5 x 10^k) so a window of this span gets
    roughly target_ticks divisions -- same magnitude-scaling idea as
    BT_TX's nice_step (make_bt_tx_freqacc_revision2.py etc.), kept
    metric-agnostic here (no floor at 1) since BT_RX's 3 metrics span very
    different natural scales (RSSI ~dBm tens, PER/BER ~0.001-0.1 fraction)
    -- a fixed floor tuned for one would break the others."""
    if span <= 0:
        return 1.0
    raw_step = span / target_ticks
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


def snap_range(vmin_raw, vmax_raw, target_ticks=8):
    """Floor/ceil (vmin_raw, vmax_raw) to a nice_step()-sized grid, then
    push ONE more step out if a bound already sits exactly on that grid --
    same rationale as BT_TX's _snap_range (2026-08-15 customer request,
    "TX 그리는 방법이랑 동일하게" -- port BT_TX's Bench/ATE-span-parity +
    spec-margin design here too): a spec Limit is often already a round
    number, so a naive floor/ceil can leave it exactly on the plotted axis
    edge where the dashed line has no visible clearance."""
    if vmax_raw <= vmin_raw:
        pad = abs(vmin_raw) * 0.1 or 1.0
        vmin_raw, vmax_raw = vmin_raw - pad, vmax_raw + pad
    step = nice_step(vmax_raw - vmin_raw, target_ticks)
    vmin = math.floor(vmin_raw / step) * step
    vmax = math.ceil(vmax_raw / step) * step
    if vmin == vmin_raw:
        vmin -= step
    if vmax == vmax_raw:
        vmax += step
    return vmin, vmax


def compute_value_range(bench_values, ate_values, spec):
    """ONE shared Y-axis span for a job's Bench+ATE PNG pair (2026-08-15
    customer request, matching BT_TX's Bench/ATE-span-parity fix): union
    of every real value on both sides plus any spec Limit, snapped to a
    nice grid with clearance (see snap_range). Computed ONCE per job by
    each make_bt_rx_*.py's build_jobs() and stored as job["value_range"],
    so the Bench script's own draw_png() call and the ATE script's
    (make_ate_rx_*.py, which reuses the identical job list via
    bench_mod.build_jobs()) both draw with the exact same span -- no
    cross-import needed here, unlike BT_TX, since this project's bench/ate
    scripts already share one job list instead of computing independently."""
    vals = list(bench_values) + list(ate_values)
    if spec is not None:
        vals.extend(v for v in spec if v is not None)
    if not vals:
        return None
    return snap_range(min(vals), max(vals))


def finalize_shared_ranges(jobs, include_power_ref=False):
    """Overwrite each job's "value_range" (initially per-job/per-column,
    see compute_value_range) with ONE shared span per packet_type -- the
    same span build_ppt_bt_rx_revision1.py's add_packet_slide() uses for
    every column's Bench+ATE chart AND its single per-row ruler on that
    packet's slide (2026-08-15 customer request: match BT_TX's Power/DEVM
    slides, where every column on a slide shares one Y-axis, superseding
    an earlier per-column-ruler attempt the customer didn't want). Pools
    every real value across ALL of that packet's matched combos (Bench,
    ATE, and any spec Limit), snapped to a nice grid with clearance (see
    snap_range).

    include_power_ref (2026-08-15, later same day, RSSI only): also folds
    in each job's "bench_pwr"/"pwr" (the actual matched Bench sweep power
    and the ATE item's own nominal target power) -- ONLY valid to set when
    the plotted quantity shares dBm units with these power values, which is
    true for RSSI but NOT for PER/BER (a percent/fraction axis), so
    make_bt_rx_rssi_revision1.py is the only caller that passes True. Keeps
    the new red power-level reference line (see draw_png's power_ref param)
    inside the shared axis even on a combo whose own RSSI data happens to
    sit far from the nominal power level."""
    by_packet = defaultdict(list)
    for job in jobs:
        by_packet[job["packet"]].append(job)
    for packet, group in by_packet.items():
        vals = []
        for job in group:
            vals.extend(job["bench_values"])
            vals.extend(job["ate_values"])
            if job["spec"] is not None:
                vals.extend(v for v in job["spec"] if v is not None)
            if include_power_ref:
                if job.get("bench_pwr") is not None:
                    vals.append(job["bench_pwr"])
                if job.get("pwr") is not None:
                    vals.append(job["pwr"])
        if not vals:
            continue
        shared = snap_range(min(vals), max(vals))
        for job in group:
            job["value_range"] = shared


def fmt_num(value: float) -> str:
    return f"{value:g}"


def combo_name(packet: str, band: int, channel: int, pwr: float) -> str:
    return f"{packet}_band{band}_CH{channel:03d}_{fmt_num(pwr)}dBm"


def condition_title(job, side: str) -> str:
    """Short, single-line, SIDE-SPECIFIC chart-title text -- side is
    "Bench" or "ATE". Shared by all 3 make_bt_rx_*.py scripts' title_for()
    (called with side="Bench") and their make_ate_rx_*.py counterparts
    (side="ATE", replacing the old "ATE_" + bench_mod.title_for(job)
    string-concat approach).

    Deliberately WITHOUT the "BT_RX_{metric}_" prefix (2026-08-14, same
    reasoning as BT_TX's combo_title_short(): the slide's own title and the
    table's condition-header already say that -- repeating it in every
    tiny chart image left nothing but the packet/condition tail actually
    visible once _fit_title_fontsize shrank a long packet name (e.g.
    HDRPS2) down to fit a narrow 4-5-column cell, genuinely clipping past
    the axes frame. combo_name() (the FILENAME) is unaffected -- still
    fully verbose/unique.

    2026-08-15 (user question, then correction): the ATE item's own
    nominal test power and the actual Bench sweep-grid power this combo
    matched against can differ slightly (e.g. ATE -86.9 vs Bench's nearest
    real grid point -86.85) since Bench's sweep is on a 0.25dB grid that
    doesn't always land exactly on ATE's target (see this module's own
    docstring / nearest_bench_rows' docstring for the matching rule). A
    first attempt showed BOTH values in one title (two lines) -- the user
    instead asked for each chart to show only ITS OWN side's power
    (Bench chart -> Bench's actual matched power, ATE chart -> ATE's own
    target power), same single-line length as the original title, so the
    narrow-column clipping problem doesn't reappear. The table's own
    per-side condition header (see build_ppt_bt_rx_revision1.py's
    _add_interleaved_table) shows the same two numbers side by side for
    direct comparison."""
    pwr = job.get("bench_pwr") if side == "Bench" else job["pwr"]
    pwr_part = f"_{side}{fmt_num(pwr)}dBm" if pwr is not None else ""
    return f"{job['packet']}_band{job['band']}_CH{job['ch']:03d}{pwr_part}"


def _fit_title_fontsize(fig, ax, text, max_width_frac=0.98, start=18, min_size=7):
    renderer = fig.canvas.get_renderer()
    max_width_px = ax.get_window_extent(renderer=renderer).width * max_width_frac
    for size in range(start, min_size - 1, -1):
        t = ax.text(0, 0, text, fontsize=size, fontweight="bold")
        width = t.get_window_extent(renderer=renderer).width
        t.remove()
        if width <= max_width_px:
            return size
    return min_size


BINS = 10  # 10 DUTs, not 40 -- see module docstring


def draw_png(values, title, xlabel, png_path: Path, value_range=None, spec=None, figsize=None,
             power_ref=None) -> None:
    """Horizontal-bar distribution histogram, same visual convention as
    make_bt_tx_pwr_revision2.draw_png (Count on X, value bins on Y).

    figsize (2026-08-14, optional): pass figure_size_for(n_cols, has_spec)
    so this PNG's own aspect ratio matches the actual PPT cell it will be
    stretched into -- see figure_size_for's docstring for the distortion
    bug this fixes. Falls back to (9, 6) when not given (e.g. a quick
    manual/standalone call). No font-size scaling by figsize is needed
    here (unlike BT_TX's _font_scale): figure_size_for's HEIGHT is
    constant regardless of n_cols (only width varies with column count),
    so fonts tuned for that one height stay legible across every slide.

    spec (2026-08-14, optional): (lsl, usl), either side possibly None --
    same per-item ATE Lower/Upper Limit convention as
    make_ate_tx_devm_99pct_revision2.py, see bt_rx_ate_lookup module
    docstring for the PER/BER unit-scale derivation. When value_range is
    auto-computed (None), the spec value is folded into the auto-range so
    the Limit line is always visible even when this combo's own data sits
    far from it -- same intent as BT_TX's spec_window().

    power_ref (2026-08-15, RSSI only): the actual input power level this
    combo was measured at (Bench's matched sweep power, or ATE's own
    nominal target -- see condition_title's docstring) drawn as a RED
    dashed reference line, distinct in both color and meaning from spec's
    blue pass/fail Limit line -- this isn't a spec, it's "how close did the
    DUT's own RSSI reading land to the actual applied signal", i.e. an
    RSSI-accuracy reference. Only meaningful when `values`/`xlabel` are
    already in dBm (RSSI), never pass this for PER/BER."""
    if not values:
        return
    vmin_data, vmax_data = min(values), max(values)
    if value_range is None:
        candidates = [vmin_data, vmax_data]
        if spec is not None:
            candidates.extend(v for v in spec if v is not None)
        if power_ref is not None:
            candidates.append(power_ref)
        vmin_c, vmax_c = min(candidates), max(candidates)
        pad = (vmax_c - vmin_c) * 0.1 or (abs(vmax_c) * 0.1 or 1.0)
        value_range = (vmin_c - pad, vmax_c + pad)

    counts, edges = np.histogram(values, bins=BINS, range=value_range)
    fig, ax = plt.subplots(figsize=figsize or (9, 6))
    max_count = max(counts.max(), 1)
    for i in range(BINS):
        lo, hi = edges[i], edges[i + 1]
        bin_h = hi - lo
        c = counts[i]
        ax.barh(lo, c, height=bin_h * 0.9, align="edge", color="#4472C4",
                edgecolor="black", linewidth=0.6)
        if c > 0:
            ax.text(c + max_count * 0.01, lo + bin_h / 2, f"{int(c)}", va="center", ha="left", fontsize=8)

    vmin, vmax = value_range
    ax.set_ylim(vmin, vmax)
    # 2026-08-14: nbins=10 (5 before that) overlapped -- figure_size_for's
    # image_h is a fixed ~2.1in regardless of n_cols (see its docstring),
    # genuinely too short for 10 evenly-spaced tick labels at a legible
    # bold size without them running into each other. nbins=5 leaves each
    # tick comfortable room; the actual bin edges (BINS=10) are unaffected,
    # this only thins the axis LABELS, not the histogram's resolution.
    ax.yaxis.set_major_locator(MaxNLocator(nbins=5))
    ax.set_xlim(0, max_count * 1.15)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True))
    # 2026-08-15 (customer request): reduced to match the customer's BT_TX
    # reference screenshot's more modest tick/label proportions -- an
    # earlier pass overshot in the opposite direction (first tried 22/18pt
    # to match BT_TX's OWN absolute point size, which badly overlapped in
    # RX's much narrower per-cell canvas -- up to 11 columns per slide vs
    # BT_TX's 3-4 -- then landed on 16/13pt, which the customer still found
    # too big). 11/10pt is smaller than even the original 13/13pt.
    plt.setp(ax.get_yticklabels(), fontsize=11, fontweight="bold")
    plt.setp(ax.get_xticklabels(), fontsize=11, fontweight="bold")
    ax.set_ylabel(xlabel, fontsize=10, fontweight="bold")
    ax.set_xlabel("Count", fontsize=10, fontweight="bold")
    ax.grid(True, axis="both", color="0.85", linewidth=0.6)
    ax.set_axisbelow(True)
    fig.tight_layout()
    ax.set_title(title, fontsize=_fit_title_fontsize(fig, ax, title, start=13, min_size=6),
                 fontweight="bold")

    if spec is not None:
        lsl, usl = spec
        # Line only, no floating text label (2026-08-14, deliberate
        # departure from BT_TX's draw_png): BT_RX's PER/BER USL is almost
        # always the value that SETS vmax in the first place (it's folded
        # into the auto-range above), leaving barely any headroom above it
        # -- combined with figure_size_for's short (~2.1in) images, every
        # text-placement heuristic tried here (above/below the line, by
        # midpoint, by available room) either clipped past the axes frame
        # or collided with the topmost y-axis tick label. The exact value
        # is already shown precisely in the PPT table's own USL/LSL row
        # (see build_ppt_bt_rx_revision1.py), so the dashed line alone
        # (still the same blue/style as BT_TX) conveys "this is the limit"
        # without a redundant, fragile label fighting for the same few
        # pixels as the axis itself.
        for spec_val in (lsl, usl):
            if spec_val is None or not (vmin <= spec_val <= vmax):
                continue
            ax.axhline(y=spec_val, color="blue", linestyle="--", linewidth=2.0, zorder=5)

    if power_ref is not None and vmin <= power_ref <= vmax:
        ax.axhline(y=power_ref, color="red", linestyle="--", linewidth=2.0, zorder=5)

    fig.tight_layout()
    png_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(png_path, dpi=450)
    plt.close(fig)
