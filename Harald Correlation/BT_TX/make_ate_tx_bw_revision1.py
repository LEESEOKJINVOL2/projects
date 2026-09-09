"""BT TX 6dB Bandwidth -- ATE counterpart to make_bt_tx_bw_revision1.py.

No 20dB BW ATE counterpart exists (confirmed 2026-08-26 by grepping the
real ATE log for "BW20DB" -- zero matches) -- Bench-only, no ATE module
for it.

Item name format (real example, user-confirmed 2026-08-26):
  BTTX0_4DH5_RF_BW06DB_CH077_1P5_X_X_NV_VXXX
    BTTX0  -> band 0
    4DH5   -> packet_type
    BW06DB -> the metric (6dB bandwidth)
    CH077  -> channel (freq = NEW_BAND_MIN_FREQ[band] + ch, same scheme
              already proven for Power's new-format ATE items)
    1P5    -> HPA supply voltage label, maps to Bench's pa_supply index via
              make_ate_tx_pwr_revision2.NEW_SUPPLY_LABEL_TO_INDEX (1P5->0,
              1P2->1, 0P73->2); "0P73LP" (a low-power variant seen in the
              real log) has no mapping and is deliberately left unmatched,
              same as Power's own identical convention.

This is exactly the SAME shape as Power's new-format ATE items (see
make_ate_tx_pwr_revision2.py's module docstring for the full format
history) -- reuses that module's NEW_FNAME_RE/NEW_BAND_MIN_FREQ/
NEW_SUPPLY_LABEL_TO_INDEX/PACKET_ALIAS directly rather than redefining
them, so the two stay in sync automatically if that format ever changes.

Unit note (user-confirmed 2026-08-26): the ATE log's own BW06DB values are
in kHz (e.g. 426.67), while Bench's bandwidth_6db_aver-MHz is in MHz (e.g.
0.427) -- divide every ATE value by 1000 before comparing/reporting.

2026-08-27: draw_png/PPT wiring added, mirroring make_ate_tx_devm_rms_
revision2.py's own structure (get_ate_group_extents/draw_all/main()) --
BW06DB's PPT slide shows Bench+ATE stacked rows like every other ATE-
backed family. No 20dB counterpart here at all (see module docstring
above) -- BW20DB's PPT slide is Bench-only, wired directly off
make_bt_tx_bw_revision1.py with ate_data=None, nothing in this file."""

from __future__ import annotations

import argparse
import csv
import functools
import math
import statistics
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import make_bt_tx_bw_revision1 as bw_mod
import make_ate_tx_pwr_revision2 as pwr_ate_mod
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")
ATE_CSV = PATHS.ate_log_path
ATE_PNG_DIR = PATHS.result_png_dir.parent / "BT_TX_ATE" / bw_mod.METRIC_META["bw6"]["dirname"]

FREQ_TOL = 5.0  # MHz, same tolerance as Power's ATE matching
VALUE_SCALE = 1.0 / 1000.0  # ATE's kHz -> Bench's MHz


def _stats(values):
    return (statistics.mean(values), min(values), max(values),
            statistics.stdev(values) if len(values) > 1 else 0.0)


def _values_from_row(row):
    # Same no-read sentinel handling as every other ATE module's own
    # _values_from_row (9.91e+37 mixed in for a no-read cell) -- checked
    # against the RAW (pre-scale) value, same as everywhere else, then
    # scaled down to Bench's MHz units.
    out = []
    for v in row[3:]:
        v = v.strip()
        if v in ("", "NA"):
            continue
        f = float(v)
        if abs(f) >= 1e6:
            continue
        out.append(f * VALUE_SCALE)
    return out


def _values_by_serial_from_row(serials, row):
    """Same value-validity filtering as _values_from_row() (blank/NA
    skipped, |v|>=1e6-pre-scale no-read sentinel dropped, /1000 scaled) but
    keeps the DUT-serial pairing that function discards -- same convention
    as make_ate_tx_pwr_revision2.py's own _values_by_serial_from_row(),
    added for the Raw Values sheet."""
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
        out[serial] = f * VALUE_SCALE
    return out


def _nearest(freq, candidates, tol):
    best, best_diff = None, None
    for r in candidates:
        diff = abs(r["freq"] - freq)
        if diff <= tol and (best_diff is None or diff < best_diff):
            best, best_diff = r, diff
    return best


@functools.lru_cache(maxsize=None)
def parse_ate_records(path: Path):
    """-> {(packet, supply): [record, ...]}, record has freq/values --
    BW06DB only, new-format items exclusively (no legacy "BTTX_" form seen
    for this metric)."""
    idx: dict[tuple, list[dict]] = {}
    with path.open(encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        next(reader)
        for row in reader:
            name = row[0]
            if not name.startswith("BTTX") or name.startswith("BTTX_"):
                continue  # BW06DB only ever appears in the new CH-based format
            m = pwr_ate_mod.NEW_FNAME_RE.match(name)
            if not m or m.group("domain") != "RF" or m.group("test") != "BW06DB":
                continue
            if m.group("c1") != "X" or m.group("c2") != "X":
                continue  # PortToANT/PortToCOND cal spot-checks, not the standard reading
            supply = pwr_ate_mod.NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
            band_min = pwr_ate_mod.NEW_BAND_MIN_FREQ.get(int(m.group("band")))
            if supply is None or band_min is None:
                continue
            packet = pwr_ate_mod.PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
            freq = float(band_min + int(m.group("ch")))

            values = _values_from_row(row)
            if not values:
                continue
            record = {"item": name, "freq": freq, "values": values}
            idx.setdefault((packet, supply), []).append(record)
    return idx


@functools.lru_cache(maxsize=None)
def parse_ate_records_per_dut(path: Path):
    """Same return shape/matching rules as parse_ate_records() above --
    mirrored verbatim, except each record carries "by_serial": {dut_serial:
    value} (via _values_by_serial_from_row()) instead of a flat "values"
    list. Added for the Raw Values sheet; parse_ate_records() itself is
    untouched."""
    idx: dict[tuple, list[dict]] = {}
    with path.open(encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        header = next(reader)
        serials = header[3:]
        for row in reader:
            name = row[0]
            if not name.startswith("BTTX") or name.startswith("BTTX_"):
                continue
            m = pwr_ate_mod.NEW_FNAME_RE.match(name)
            if not m or m.group("domain") != "RF" or m.group("test") != "BW06DB":
                continue
            if m.group("c1") != "X" or m.group("c2") != "X":
                continue
            supply = pwr_ate_mod.NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
            band_min = pwr_ate_mod.NEW_BAND_MIN_FREQ.get(int(m.group("band")))
            if supply is None or band_min is None:
                continue
            packet = pwr_ate_mod.PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
            freq = float(band_min + int(m.group("ch")))

            by_serial = _values_by_serial_from_row(serials, row)
            if not by_serial:
                continue
            record = {"item": name, "freq": freq, "by_serial": by_serial}
            idx.setdefault((packet, supply), []).append(record)
    return idx


def get_ate_data(pa_slices_max=None, dig_gain_max=None):
    """-> ({(packet,supply,gain,freq): {"png": Path|None, "stats":
    (mean,min,max,std)|None, "status": "EXACT"|"APPROX"|"NONE", "item":
    str|None}}, bench_bw6, pa_slices_max, dig_gain_max) -- bw6 only; caller
    handles bw20 separately (Bench-only, no ATE side at all).

    "png" (2026-08-27, added for the PPT wiring): build_ppt_pwr_evm_
    revision2._add_ate_images() reads this key directly off every ate_data
    entry (info["png"]) to place the ATE chart image, same as every other
    ATE module's get_ate_data() (e.g. make_ate_tx_devm_rms_revision2.py) --
    without it every BW06DB ATE cell would render as a false "No ATE match"
    hole even though draw_all() (this module's own PNG generator) wrote a
    real file for it."""
    raw, slices_seen, dig_seen = bw_mod.read_all(PATHS.bench_dir)
    if pa_slices_max is None:
        pa_slices_max = max(slices_seen)
    if dig_gain_max is None:
        dig_gain_max = max(dig_seen)
    bench_bw6 = bw_mod.build_fixed_combos(raw["bw6"], pa_slices_max, dig_gain_max)

    ate_idx = parse_ate_records(ATE_CSV)

    out = {}
    for combo in bench_bw6.keys():
        packet, supply, gain, freq = combo
        candidates = ate_idx.get((packet, supply), [])
        rec = _nearest(freq, candidates, FREQ_TOL)
        if rec is None:
            out[combo] = {"png": None, "stats": None, "status": "NONE", "item": None}
            continue
        status = "EXACT" if rec["freq"] == freq else "APPROX"
        fname = bw_mod.combo_name(combo, pa_slices_max, dig_gain_max)
        png_path = ATE_PNG_DIR / f"{fname}.png"
        out[combo] = {
            "png": png_path if png_path.exists() else None,
            "stats": _stats(rec["values"]),
            "status": status,
            "item": rec["item"],
        }
    return out, bench_bw6, pa_slices_max, dig_gain_max


def get_ate_group_extents(bench_bw6, pa_slices_max, dig_gain_max):
    """{group: (min_ate_value, max_ate_value)} pooled from every Bench bw6
    combo's matched ATE record -- same convention as make_ate_tx_devm_rms_
    revision2.get_ate_group_extents (see that function's docstring), called
    from BOTH this module's own draw_all() and make_bt_tx_bw_revision1.py's
    _draw_one_metric() so both compute the identical union independently
    from the same two raw data sources."""
    if PATHS.extract_mode == "bench":
        return {}
    try:
        ate_idx = parse_ate_records(ATE_CSV)
    except (FileNotFoundError, OSError):
        return {}
    group_values = defaultdict(list)
    for combo in bench_bw6.keys():
        packet, supply, gain, freq = combo
        candidates = ate_idx.get((packet, supply), [])
        rec = _nearest(freq, candidates, FREQ_TOL)
        if rec is None:
            continue
        group = bw_mod.group_of(packet, gain, freq, supply)
        group_values[group].extend(rec["values"])
    return {g: (min(vs), max(vs)) for g, vs in group_values.items()}


def draw_all(limit: int = 0, count_only: bool = False):
    ate_data, bench_bw6, pa_slices_max, dig_gain_max = get_ate_data()

    raw, _slices_seen, _dig_seen = bw_mod.read_all(PATHS.bench_dir)
    data = bw_mod.build_fixed_combos(raw["bw6"], pa_slices_max, dig_gain_max)
    ranges, _kept, _outliers, _incl, step_by_group = bw_mod.build_packet_ranges(data)
    # Widen each group's axis to also cover real ATE data -- see this
    # module's get_ate_group_extents() docstring; called independently from
    # make_bt_tx_bw_revision1.py's own bw6 pass, both arrive at the same
    # widened range from the same two raw sources.
    for group, (lo, hi) in get_ate_group_extents(bench_bw6, pa_slices_max, dig_gain_max).items():
        if group in ranges:
            vmin, vmax = ranges[group]
            step = step_by_group.get(group) or bw_mod.nice_step(max(hi - lo, 1e-9))
            ranges[group] = (min(vmin, bw_mod._floor_to_step(lo, step)),
                              max(vmax, bw_mod._ceil_to_step(hi, step)))

    total = len(ate_data)
    exact = sum(1 for v in ate_data.values() if v["status"] == "EXACT")
    approx = sum(1 for v in ate_data.values() if v["status"] == "APPROX")
    none_ = sum(1 for v in ate_data.values() if v["status"] == "NONE")
    print(f"Bench BW06DB combos: {total}  ATE EXACT: {exact}  APPROX: {approx}  NONE: {none_}")

    if count_only:
        return

    matched = [(combo, v) for combo, v in ate_data.items() if v["status"] != "NONE"]
    if limit > 0:
        matched = matched[:limit]
        print(f"Limiting to first {len(matched)} matched combo(s)")

    ATE_PNG_DIR.mkdir(parents=True, exist_ok=True)
    ate_idx = parse_ate_records(ATE_CSV)

    # Precompute each combo's PPT column count from the FULL Bench bw6 combo
    # set -- same block as every other ATE module's draw_all().
    n_cols_for = bw_mod._n_cols_for_metric(list(bench_bw6.keys()))

    combo_recs = {}
    for combo, v in matched:
        packet, supply, gain, freq = combo
        candidates = ate_idx.get((packet, supply), [])
        combo_recs[combo] = _nearest(freq, candidates, FREQ_TOL)

    jobs = []
    for combo, v in matched:
        packet, supply, gain, freq = combo
        rec = combo_recs[combo]
        group = bw_mod.group_of(packet, gain, freq, supply)
        value_range = ranges.get(group)
        step = step_by_group.get(group, 1)
        name = bw_mod.combo_name(combo, pa_slices_max, dig_gain_max)
        title = "ATE_" + bw_mod.combo_title_short(combo, pa_slices_max, dig_gain_max)
        png_path = ATE_PNG_DIR / f"{name}.png"
        figsize = bw_mod.figure_size_for(n_cols_for(packet, supply, freq))
        jobs.append((rec["values"], title, png_path, value_range, step, figsize, name, rec["item"]))

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
    values, title, png_path, value_range, step, figsize, name, item = job
    bw_mod.draw_png(values, title, bw_mod.METRIC_META["bw6"]["xlabel"], png_path,
                     value_range=value_range, step=step, spec=None, figsize=figsize)
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
