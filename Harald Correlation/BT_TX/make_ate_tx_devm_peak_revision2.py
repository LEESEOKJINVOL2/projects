"""BT TX DEVM Peak -- ATE counterpart to make_bt_tx_devm_peak_revision2.py.

Same matching rules as make_ate_tx_pwr_revision2.py (packet alias, band
recomputed from frequency not the ATE name's 2G/5G token, freq tolerance,
no gain/pwr disambiguation needed) -- see that module's docstring for the
full rationale. Only the metric (ATE "DEVM_PEAK", Bench Y_COL
evm_peak_aver-%) and the value-range/step lookup (DEVM's spec-anchored
Y-axis via group_of()/build_packet_ranges(), not Power's simple per-band
floor/ceil) differ.

CLI: --limit N (only draw the first N matched combos), --count-only.

2026-08-14: a newer ATE pull uses a second item-name format (see
make_ate_tx_pwr_revision2.py's module docstring for the full field-by-field
derivation) -- added alongside FNAME_RE, not replacing it. For DEVM Peak
specifically: EDR-family packets (2DH5/3DH5/4DH5/8DH5) use test "DEVMPEAK";
every other DEVM-bearing family (HDRPS2/HDRPM8/12/16/HDRPL16/24/32,
HDT-family via the HDT7P5 alias, UHDR32/48) uses "EVMPEAK" instead --
confirmed empirically, a packet only ever has one of the two names.
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

import make_bt_tx_devm_peak_revision2 as devm_mod
import bt_tx_devm_spec_lookup as spec_lookup
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")

ATE_CSV = PATHS.ate_log_path
ATE_PNG_DIR = PATHS.result_png_dir.parent / "BT_TX_ATE" / "devm_peak_revision2"

FREQ_TOL = 5.0  # MHz
PACKET_ALIAS = {"2M": "LE2M", "VHDRPM16": "HDRPM16", "VHDRPL32": "HDRPL32", "HDT7P5": "HDT8"}
ATE_METRIC = "DEVM_PEAK"
DEVM_SPEC_KEY = "peak"

FNAME_RE = re.compile(
    r"^BTTX_(?P<mode>[A-Za-z0-9]+)_(?P<band>\d+G)_(?P<freq>\d+)MHZ_X_EPA_"
    r"(?P<pwr>[A-Za-z0-9]+)DBM_(?P<packet>[A-Za-z0-9]+)_PS(?P<supply>\d+)_X_"
    r"(?P<metric>[A-Za-z0-9_]+)$"
)

# -- New (2026-08-14) item-name format, see make_ate_tx_pwr_revision2.py --
NEW_FNAME_RE = re.compile(
    r"^BTTX(?P<band>\d+)_(?P<packet>[A-Za-z0-9]+)_(?P<domain>[A-Za-z]+)_(?P<test>[A-Za-z0-9]+)_"
    r"CH(?P<ch>\d+)_(?P<supply>[0-9A-Za-z]+)_(?P<c1>[A-Za-z0-9]+)_(?P<c2>[A-Za-z0-9]+)_NV_VXXX$"
)
NEW_BAND_MIN_FREQ = {0: 2402, 1: 5150, 2: 5725, 4: 5925, 5: 6051, 6: 6176, 7: 6301}
NEW_SUPPLY_LABEL_TO_INDEX = {"1P5": "0", "1P2": "1", "0P73": "2"}
NEW_DEVM_PEAK_TEST_NAMES = ("DEVMPEAK", "EVMPEAK")


def _stats(values):
    return (statistics.mean(values), min(values), max(values),
            statistics.stdev(values) if len(values) > 1 else 0.0)


def _values_from_row(row):
    # 2026-07-20: the ATE log itself contains a "no-read" sentinel
    # (9.91e+37, seen on a handful of HDRPM16/HDRPL32 DEVM_RMS DUTs)
    # mixed in among real measurements -- treat any magnitude no real
    # DEVM/Power/ACP/FreqAcc value could ever reach as invalid, same
    # as a blank cell, or it silently wrecks the pooled mean/std.
    return [f for v in row[3:] if v.strip() not in ("", "NA") and abs(f := float(v)) < 1e6]


@functools.lru_cache(maxsize=None)
def parse_ate_records(path: Path, metric: str = ATE_METRIC):
    """Tries the OLD item-name format (FNAME_RE) first, then the NEW one
    (NEW_FNAME_RE) -- see make_ate_tx_pwr_revision2.py's module docstring.

    2026-07-20: @lru_cache -- get_ate_data() and draw_all() each
    independently re-parsed this same 26k-row ATE CSV in one process run;
    caching means the second call is free. Safe: nothing downstream
    mutates the returned index dict/lists."""
    idx: dict[tuple, list[dict]] = {}
    with path.open(encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        next(reader)
        for row in reader:
            name = row[0]
            if not name.startswith("BTTX"):
                continue

            if name.startswith("BTTX_"):
                m = FNAME_RE.match(name)
                if not m or m.group("metric") != metric:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                supply = m.group("supply")
                freq = float(m.group("freq"))
            else:
                if metric != ATE_METRIC:
                    continue
                m = NEW_FNAME_RE.match(name)
                if not m or m.group("domain") != "RF" or m.group("test") not in NEW_DEVM_PEAK_TEST_NAMES:
                    continue
                if m.group("c1") != "X" or m.group("c2") != "X":
                    continue
                supply = NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
                band_min = NEW_BAND_MIN_FREQ.get(int(m.group("band")))
                if supply is None or band_min is None:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(band_min + int(m.group("ch")))

            values = _values_from_row(row)
            if not values:
                continue
            record = {
                "item": name, "freq": freq,
                "values": values,
                "ul": (row[1] or "").strip() if len(row) > 1 else "",
                "ll": (row[2] or "").strip() if len(row) > 2 else "",
            }
            idx.setdefault((packet, supply), []).append(record)
    return idx


def _nearest(freq, candidates, tol):
    best, best_diff = None, None
    for r in candidates:
        diff = abs(r["freq"] - freq)
        if diff <= tol and (best_diff is None or diff < best_diff):
            best, best_diff = r, diff
    return best


def _values_by_serial_from_row(serials, row):
    """Same value-validity filtering as _values_from_row() (blank/NA
    skipped, |v|>=1e6 no-read sentinel dropped) but keeps the DUT-serial
    pairing that function discards -- `serials` (the caller's captured
    header[3:]) lines up 1:1 with row[3:] since both come from the same
    CSV row/header in column order. 2026-08-25, added for the Raw Values
    sheet (add_bt_tx_raw_values_sheet.py); _values_from_row() itself is
    untouched."""
    out = {}
    for serial, v in zip(serials, row[3:]):
        v = (v or "").strip()
        if v in ("", "NA"):
            continue
        try:
            f = float(v)
        except ValueError:
            continue
        if abs(f) >= 1e6:
            continue
        out[serial] = f
    return out


@functools.lru_cache(maxsize=None)
def parse_ate_records_per_dut(path: Path, metric: str = ATE_METRIC):
    """Same return shape/matching rules as parse_ate_records() above --
    {(packet, supply): [record, ...]} -- mirrored verbatim, except each
    record carries "by_serial": {dut_serial: value} (via
    _values_by_serial_from_row()) instead of a flat "values" list. 2026-08-25,
    added for the Raw Values sheet (add_bt_tx_raw_values_sheet.py);
    parse_ate_records() itself is untouched."""
    idx: dict[tuple, list[dict]] = {}
    with path.open(encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        header = next(reader)
        serials = header[3:]
        for row in reader:
            name = row[0]
            if not name.startswith("BTTX"):
                continue

            if name.startswith("BTTX_"):
                m = FNAME_RE.match(name)
                if not m or m.group("metric") != metric:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                supply = m.group("supply")
                freq = float(m.group("freq"))
            else:
                if metric != ATE_METRIC:
                    continue
                m = NEW_FNAME_RE.match(name)
                if not m or m.group("domain") != "RF" or m.group("test") not in NEW_DEVM_PEAK_TEST_NAMES:
                    continue
                if m.group("c1") != "X" or m.group("c2") != "X":
                    continue
                supply = NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
                band_min = NEW_BAND_MIN_FREQ.get(int(m.group("band")))
                if supply is None or band_min is None:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(band_min + int(m.group("ch")))

            by_serial = _values_by_serial_from_row(serials, row)
            if not by_serial:
                continue
            record = {
                "item": name, "freq": freq,
                "by_serial": by_serial,
            }
            idx.setdefault((packet, supply), []).append(record)
    return idx


def get_ate_data(pa_slices_max=None, dig_gain_max=None):
    """-> {(packet, supply, gain, freq): {"png":.., "stats":.., "status":.., "item":..}}"""
    raw, slices_seen, dig_seen = devm_mod.read_all(PATHS.bench_dir)
    if pa_slices_max is None:
        pa_slices_max = max(slices_seen)
    if dig_gain_max is None:
        dig_gain_max = max(dig_seen)
    bench = devm_mod.build_fixed_combos(raw, pa_slices_max, dig_gain_max)

    ate_idx = parse_ate_records(ATE_CSV)

    out = {}
    for combo in bench.keys():
        packet, supply, gain, freq = combo
        candidates = ate_idx.get((packet, supply), [])
        rec = _nearest(freq, candidates, FREQ_TOL)
        if rec is None:
            out[combo] = {"png": None, "stats": None, "status": "NONE", "item": None}
            continue
        status = "EXACT" if rec["freq"] == freq else "APPROX"
        fname = devm_mod.combo_name(combo, pa_slices_max, dig_gain_max)
        png_path = ATE_PNG_DIR / f"{fname}.png"
        out[combo] = {
            "png": png_path if png_path.exists() else None,
            "stats": _stats(rec["values"]),
            "status": status,
            "item": rec["item"],
        }
    return out, bench, pa_slices_max, dig_gain_max


def get_ate_group_extents(bench, pa_slices_max, dig_gain_max):
    """{group: (min_ate_value, max_ate_value)} pooled from every Bench
    combo's matched ATE record. Called from BOTH this module's own
    draw_all() and the Bench module's main() (see make_bt_tx_devm_peak_
    revision2.py) so both compute the identical union independently from
    the same two raw data sources -- 2026-08-15 customer request: Bench
    and ATE charts must share the exact same Y-axis span."""
    if PATHS.extract_mode == "bench":
        return {}
    try:
        ate_idx = parse_ate_records(ATE_CSV)
    except (FileNotFoundError, OSError):
        return {}
    group_values = defaultdict(list)
    for combo in bench.keys():
        packet, supply, gain, freq = combo
        candidates = ate_idx.get((packet, supply), [])
        rec = _nearest(freq, candidates, FREQ_TOL)
        if rec is None:
            continue
        group = devm_mod.group_of(packet, gain, freq, supply)
        group_values[group].extend(rec["values"])
    return {g: (min(vs), max(vs)) for g, vs in group_values.items()}


def draw_all(limit: int = 0, count_only: bool = False):
    ate_data, bench, pa_slices_max, dig_gain_max = get_ate_data()

    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    spec_by_packet = spec_table.get(DEVM_SPEC_KEY, {})
    ranges, _kept, _outliers, _incl, step_by_group = devm_mod.build_packet_ranges(bench, spec_by_packet)

    total = len(ate_data)
    exact = sum(1 for v in ate_data.values() if v["status"] == "EXACT")
    approx = sum(1 for v in ate_data.values() if v["status"] == "APPROX")
    none_ = sum(1 for v in ate_data.values() if v["status"] == "NONE")
    print(f"Bench {ATE_METRIC} combos: {total}  ATE EXACT: {exact}  APPROX: {approx}  NONE: {none_}")

    if count_only:
        return

    matched = [(combo, v) for combo, v in ate_data.items() if v["status"] != "NONE"]
    if limit > 0:
        matched = matched[:limit]
        print(f"Limiting to first {len(matched)} matched combo(s)")

    ATE_PNG_DIR.mkdir(parents=True, exist_ok=True)
    ate_idx = parse_ate_records(ATE_CSV)

    # Precompute each combo's PPT column count from the FULL Bench combo set
    # -- see make_ate_tx_pwr_revision2.py's identical block for the full
    # rationale.
    band24_freqs = defaultdict(set)
    highfreq_freqs = defaultdict(set)
    for packet, supply, _gain, freq in bench.keys():
        if devm_mod.band_of(freq) == 2.4:
            band24_freqs[(packet, supply)].add(freq)
        else:
            highfreq_freqs[(packet, devm_mod.band_of(freq), supply)].add(freq)
    band24_supplies = defaultdict(set)
    for (packet, supply) in band24_freqs:
        band24_supplies[packet].add(supply)
    band24_n_cols = {
        packet: max(len(band24_freqs[(packet, s)]) for s in supplies) * len(supplies)
        for packet, supplies in band24_supplies.items()
    }

    def n_cols_for(packet, supply, freq):
        if devm_mod.band_of(freq) == 2.4:
            return band24_n_cols[packet]
        return len(highfreq_freqs[(packet, devm_mod.band_of(freq), supply)])

    combo_recs = {}
    for combo, v in matched:
        packet, supply, gain, freq = combo
        candidates = ate_idx.get((packet, supply), [])
        combo_recs[combo] = _nearest(freq, candidates, FREQ_TOL)

    # Widen each axis-sharing group's Y-range to cover real ATE data that
    # falls outside the Bench-derived window -- see make_ate_tx_pwr_
    # revision2.py's identical block for the full rationale (2026-08-14).
    # get_ate_group_extents() is the SAME function the Bench module's
    # main() calls (2026-08-15), so both scripts arrive at the identical
    # widened range independently -- no file-based coordination needed.
    for group, (lo, hi) in get_ate_group_extents(bench, pa_slices_max, dig_gain_max).items():
        if group in ranges:
            vmin, vmax = ranges[group]
            ranges[group] = (min(vmin, math.floor(lo)), max(vmax, math.ceil(hi)))

    jobs = []
    for combo, v in matched:
        packet, supply, gain, freq = combo
        rec = combo_recs[combo]
        group = devm_mod.group_of(packet, gain, freq, supply)
        value_range = ranges.get(group)
        step = step_by_group.get(group, 1)
        spec = spec_lookup.lookup(spec_table, DEVM_SPEC_KEY, packet)
        name = devm_mod.combo_name(combo, pa_slices_max, dig_gain_max)
        title = "ATE_" + devm_mod.combo_title_short(combo, pa_slices_max, dig_gain_max)
        png_path = ATE_PNG_DIR / f"{name}.png"
        figsize = devm_mod.figure_size_for(n_cols_for(packet, supply, freq))
        jobs.append((rec["values"], title, png_path, value_range, step, spec, figsize, name, rec["item"]))

    n = 0
    with ProcessPoolExecutor(max_workers=PATHS.png_workers) as pool:
        for name, item in pool.map(_draw_one, jobs, chunksize=4):
            n += 1
            if n % 200 == 0 or n == len(jobs):
                print(f"  [{n}/{len(jobs)}] {name}  (ATE item: {item})")

    print(f"\nDone. ATE PNG -> {ATE_PNG_DIR}")


def _draw_one(job):
    """Worker for ProcessPoolExecutor -- must be a module-level function so
    it's picklable on Windows (spawn-only, no fork)."""
    values, title, png_path, value_range, step, spec, figsize, name, item = job
    devm_mod.draw_png(values, title, "DEVM Peak (%)", png_path, value_range=value_range, step=step, spec=spec,
                       figsize=figsize)
    return name, item


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--limit", type=int, default=0, help="only draw the first N matched combos")
    p.add_argument("--count-only", action="store_true", help="print match-rate audit only")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    draw_all(limit=args.limit, count_only=args.count_only)


if __name__ == "__main__":
    main()
