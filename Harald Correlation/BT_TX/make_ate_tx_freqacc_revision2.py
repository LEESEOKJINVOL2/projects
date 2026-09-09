"""BT TX FreqAcc -- ATE counterpart to make_bt_tx_freqacc_revision2.py.

Same matching rules as make_ate_tx_pwr_revision2.py (packet alias, band
recomputed from frequency not the ATE name's 2G/5G token, freq tolerance,
no gain/pwr disambiguation needed) -- see that module's docstring for the
full rationale. Consolidated into ONE script covering all 9 FreqAcc
sub-metrics (mirroring make_bt_tx_freqacc_revision2.py's own single-script,
single-CSV-read-pass design), rather than one file per sub-metric like the
3 DEVM scripts:

  key          ATE metric        applicable packets
  icft         FREQ_OFFHEADER    1DH5, LE1M, LE2M, HDT3, HDT4, HDT8
  cfo          FREQ_OFFSET       1DH5, LE1M, LE2M, HDT3, HDT4, HDT8
  drift_rate   FREQ_DRIFTRATE    1DH5, LE1M, LE2M
  df1_avg      DF1_AVG           1DH5, LE1M, LE2M
  df2_avg      DF2_AVG           1DH5, LE1M, LE2M
  df2_max      DF2_99MAX         1DH5, LE1M, LE2M
  omega_i      OMEGA_I           3DH5, 4DH5, 8DH5
  omega_o      OMEGA_0           3DH5, 4DH5, 8DH5
  omega_io     OMEGA_I0          3DH5, 4DH5, 8DH5
  freq_drift   FREQ_DRIFT        HDT3, HDT4, HDT8 (Bench has ZERO data for
                                  this one -- frequency_drift_aver-Hz is
                                  blank for every HDT row in this dataset,
                                  confirmed 2026-07-20 -- so it's skipped
                                  automatically, same as the Bench script.)

CLI: --metric KEY (only this sub-metric; default = all), --limit N (only
draw the first N matched combos per metric), --count-only.

2026-08-14: a newer ATE pull uses a second item-name format (see
make_ate_tx_pwr_revision2.py's module docstring for the full field-by-field
derivation) -- added alongside FNAME_RE, not replacing it. The new format
uses DIFFERENT test-name tokens than the old METRIC_ATE_MAP (which are that
old pull's own convention) -- see NEW_METRIC_ATE_MAP. cfo/drift_rate/
freq_drift have no counterpart test name anywhere in the new pull (checked
1DH5's and 3DH5/4DH5/8DH5's full RF test-name sets) and stay unmapped, a
real gap not a parsing bug. omega_i/omega_o/omega_io map to
INITFREQERR/BLOCKFREQERR/TOTFREQERR -- inferred from standard Bluetooth
terminology (initial vs block vs total carrier frequency error, matching
wi/w0/wi+w0) since these are the only remaining frequency-stability-shaped
RF test names left over for 3DH5/4DH5/8DH5 once every other sub-metric had
an exact-name match; not directly confirmed against a customer-supplied
mapping. Also: the NEW format's values are already in Hz (confirmed by its
own per-item Upper/Lower Limit matching Bench's spec directly, e.g. ICFT
UL/LL=75000/-75000, DF1AAVG UL/LL=175000/140000) -- unlike the OLD format's
kHz convention, no *1000 scaling is applied to new-format rows.
"""
from __future__ import annotations

import argparse
import csv
import math
import re
import statistics
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import make_bt_tx_freqacc_revision2 as FA
import bt_tx_freqacc_spec_lookup as spec_lookup
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")

ATE_CSV = PATHS.ate_log_path
ATE_PNG_BASE = PATHS.result_png_dir.parent / "BT_TX_ATE"

FREQ_TOL = 5.0  # MHz
PACKET_ALIAS = {"2M": "LE2M", "VHDRPM16": "HDRPM16", "VHDRPL32": "HDRPL32"}

METRIC_ATE_MAP = {
    "icft": "FREQ_OFFHEADER",
    "cfo": "FREQ_OFFSET",
    "drift_rate": "FREQ_DRIFTRATE",
    "df1_avg": "DF1_AVG",
    "df2_avg": "DF2_AVG",
    "df2_max": "DF2_99MAX",
    "omega_i": "OMEGA_I",
    "omega_o": "OMEGA_0",
    "omega_io": "OMEGA_I0",
    "freq_drift": "FREQ_DRIFT",
}

FNAME_RE = re.compile(
    r"^BTTX_(?P<mode>[A-Za-z0-9]+)_(?P<band>\d+G)_(?P<freq>\d+)MHZ_X_EPA_"
    r"(?P<pwr>[A-Za-z0-9]+)DBM_(?P<packet>[A-Za-z0-9]+)_PS(?P<supply>\d+)_X_"
    r"(?P<metric>[A-Za-z0-9_]+)$"
)

# -- New (2026-08-14) item-name format, see module docstring and
# make_ate_tx_pwr_revision2.py's -- test-name tokens differ from
# METRIC_ATE_MAP above, and cfo/drift_rate/freq_drift have no counterpart. --
NEW_FNAME_RE = re.compile(
    r"^BTTX(?P<band>\d+)_(?P<packet>[A-Za-z0-9]+)_(?P<domain>[A-Za-z]+)_(?P<test>[A-Za-z0-9]+)_"
    r"CH(?P<ch>\d+)_(?P<supply>[0-9A-Za-z]+)_(?P<c1>[A-Za-z0-9]+)_(?P<c2>[A-Za-z0-9]+)_NV_VXXX$"
)
NEW_BAND_MIN_FREQ = {0: 2402, 1: 5150, 2: 5725, 4: 5925, 5: 6051, 6: 6176, 7: 6301}
NEW_SUPPLY_LABEL_TO_INDEX = {"1P5": "0", "1P2": "1", "0P73": "2"}
NEW_METRIC_ATE_MAP = {
    "icft": "ICFT",
    "df1_avg": "DF1AAVG",
    "df2_avg": "DF2AAVG",
    "df2_max": "DF2AMAX",
    "omega_i": "INITFREQERR",
    "omega_o": "BLOCKFREQERR",
    "omega_io": "TOTFREQERR",
}
NEW_TEST_TO_KEY = {v: k for k, v in NEW_METRIC_ATE_MAP.items()}


def _stats(values):
    return (statistics.mean(values), min(values), max(values),
            statistics.stdev(values) if len(values) > 1 else 0.0)


def _values_from_row(row, scale=1.0):
    # 2026-07-20: the ATE log itself contains a "no-read" sentinel
    # (9.91e+37, seen on a handful of HDRPM16/HDRPL32 DEVM_RMS DUTs)
    # mixed in among real measurements -- filter it (pre-scale, since it's
    # just as implausible scaled as not) same as a blank cell, or it
    # silently wrecks the pooled mean/std.
    return [f * scale for v in row[3:] if v.strip() not in ("", "NA") and abs(f := float(v)) < 1e6]


def parse_all_ate_records(path: Path):
    """Single pass over the ATE log -> {(packet, supply, metric_key): [record, ...]}
    for every metric this module cares about, so all 9 sub-metrics share one
    file read instead of 9. `metric_key` is METRIC_ATE_MAP's OLD-format
    metric string for old-format rows, or the sub-metric KEY itself (icft/
    df1_avg/...) for new-format rows -- draw_all()/get_ate_data_all() look
    up by OLD_METRIC_ATE_MAP[key] first then by key, see their own lookups."""
    wanted_metrics = set(METRIC_ATE_MAP.values())
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
                if not m or m.group("metric") not in wanted_metrics:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                supply = m.group("supply")
                freq = float(m.group("freq"))
                # ATE reports every FreqAcc-family metric in kHz while Bench
                # reports in Hz (OLD format only) -- confirmed 2026-07-20 by
                # comparing each metric's ATE Upper/Lower Limit against
                # Bench's own spec value: DF1_AVG 175/175000, DF2_99MAX
                # 115/115000, FREQ_DRIFTRATE 20/20000, OMEGA_0 10/10000,
                # OMEGA_I/OMEGA_I0 75/75000 -- exactly x1000 in every case.
                values = _values_from_row(row, scale=1000.0)
                index_metric = m.group("metric")
            else:
                m = NEW_FNAME_RE.match(name)
                if not m or m.group("domain") != "RF" or m.group("test") not in NEW_TEST_TO_KEY:
                    continue
                if m.group("c1") != "X" or m.group("c2") != "X":
                    continue
                supply = NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
                band_min = NEW_BAND_MIN_FREQ.get(int(m.group("band")))
                if supply is None or band_min is None:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(band_min + int(m.group("ch")))
                # New format is already Hz -- see module docstring.
                values = _values_from_row(row, scale=1.0)
                index_metric = NEW_TEST_TO_KEY[m.group("test")]

            if not values:
                continue
            record = {
                "item": name, "freq": freq, "values": values,
                "ul": (row[1] or "").strip() if len(row) > 1 else "",
                "ll": (row[2] or "").strip() if len(row) > 2 else "",
            }
            idx.setdefault((packet, supply, index_metric), []).append(record)
    return idx


def _values_by_serial_from_row(serials, row, scale=1.0):
    """Same value-validity filtering as _values_from_row() (blank/NA
    skipped, |v|>=1e6 no-read sentinel dropped, pre-scale) but keeps the
    DUT-serial pairing that function discards -- `serials` (the caller's
    captured header[3:]) lines up 1:1 with row[3:] since both come from
    the same CSV row/header in column order. 2026-08-25, added for the Raw
    Values sheet (add_bt_tx_raw_values_sheet.py); _values_from_row() itself
    is untouched."""
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
        out[serial] = f * scale
    return out


def parse_all_ate_records_per_dut(path: Path):
    """Same return shape/matching rules as parse_all_ate_records() above --
    {(packet, supply, metric_key): [record, ...]} -- mirrored verbatim,
    except each record carries "by_serial": {dut_serial: value} (via
    _values_by_serial_from_row(), same OLD-format x1000 kHz->Hz scale)
    instead of a flat "values" list. 2026-08-25, added for the Raw Values
    sheet (add_bt_tx_raw_values_sheet.py); parse_all_ate_records() itself
    is untouched."""
    wanted_metrics = set(METRIC_ATE_MAP.values())
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
                if not m or m.group("metric") not in wanted_metrics:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                supply = m.group("supply")
                freq = float(m.group("freq"))
                by_serial = _values_by_serial_from_row(serials, row, scale=1000.0)
                index_metric = m.group("metric")
            else:
                m = NEW_FNAME_RE.match(name)
                if not m or m.group("domain") != "RF" or m.group("test") not in NEW_TEST_TO_KEY:
                    continue
                if m.group("c1") != "X" or m.group("c2") != "X":
                    continue
                supply = NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
                band_min = NEW_BAND_MIN_FREQ.get(int(m.group("band")))
                if supply is None or band_min is None:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(band_min + int(m.group("ch")))
                by_serial = _values_by_serial_from_row(serials, row, scale=1.0)
                index_metric = NEW_TEST_TO_KEY[m.group("test")]

            if not by_serial:
                continue
            record = {
                "item": name, "freq": freq,
                "by_serial": by_serial,
            }
            idx.setdefault((packet, supply, index_metric), []).append(record)
    return idx


def _ate_candidates(ate_idx, packet, supply, key):
    """Merge OLD-format (METRIC_ATE_MAP[key]) and NEW-format (key itself)
    candidates -- parse_all_ate_records() indexes old-format rows under the
    old metric string and new-format rows under the sub-metric key, see its
    own docstring. A real ATE pull only ever populates one of the two."""
    old_metric = METRIC_ATE_MAP.get(key)
    old = ate_idx.get((packet, supply, old_metric), []) if old_metric else []
    new = ate_idx.get((packet, supply, key), [])
    return old + new if old and new else (old or new)


def _nearest(freq, candidates, tol):
    best, best_diff = None, None
    for r in candidates:
        diff = abs(r["freq"] - freq)
        if diff <= tol and (best_diff is None or diff < best_diff):
            best, best_diff = r, diff
    return best


def get_ate_data_all(pa_slices_max=None, dig_gain_max=None):
    """-> (ate_data_by_metric, bench_by_metric, ranges_by_metric,
    pa_slices_max, dig_gain_max). ate_data_by_metric[key] =
    {(packet,supply,gain,freq): {"png":.., "stats":.., "status":.., "item":..}}.

    ranges_by_metric[key] is spec_fn-anchored the same way the Bench chart
    script computes it (2026-07-20), so an ATE chart drawn with this same
    range shares the Bench chart's axis exactly."""
    metric_cols = {key: col for key, col, _xlabel, _packets in FA.METRICS}
    raw, slices_seen, dig_seen = FA.read_all(PATHS.bench_dir, metric_cols)
    if pa_slices_max is None:
        pa_slices_max = max(slices_seen)
    if dig_gain_max is None:
        dig_gain_max = max(dig_seen)

    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    ate_idx = parse_all_ate_records(ATE_CSV)

    ate_data_by_metric = {}
    bench_by_metric = {}
    ranges_by_metric = {}
    for key, col, xlabel, packets in FA.METRICS:
        bench = FA.build_fixed_combos(raw[key], pa_slices_max, dig_gain_max, packets)
        bench_by_metric[key] = bench
        if not bench:
            ate_data_by_metric[key] = {}
            continue

        def spec_fn(packet, band, _key=key):
            return spec_lookup.lookup(spec_table, _key, packet, band)

        ranges_by_metric[key] = FA.build_packet_ranges(bench, spec_fn=spec_fn)[0]
        ate_metric = METRIC_ATE_MAP[key]
        out = {}
        for combo in bench.keys():
            packet, supply, gain, freq = combo
            candidates = _ate_candidates(ate_idx, packet, supply, key)
            rec = _nearest(freq, candidates, FREQ_TOL)
            if rec is None:
                out[combo] = {"png": None, "stats": None, "status": "NONE", "item": None}
                continue
            status = "EXACT" if rec["freq"] == freq else "APPROX"
            fname = FA.combo_name(combo, pa_slices_max, dig_gain_max)
            png_path = ATE_PNG_BASE / f"{key}_revision2" / f"{fname}.png"
            out[combo] = {
                "png": png_path if png_path.exists() else None,
                "stats": _stats(rec["values"]),
                "status": status,
                "item": rec["item"],
            }
        ate_data_by_metric[key] = out
    return ate_data_by_metric, bench_by_metric, ranges_by_metric, pa_slices_max, dig_gain_max


def get_ate_group_extents(key, bench, ate_idx=None):
    """{(packet,gain,band): (min_ate_value, max_ate_value)} pooled from
    every Bench combo's matched ATE record, for one FreqAcc sub-metric
    `key`. Called from BOTH this module's own draw_all() and the Bench
    module's main() (see make_bt_tx_freqacc_revision2.py) so both compute
    the identical union independently from the same two raw data sources
    -- 2026-08-15 customer request: Bench and ATE charts must share the
    exact same Y-axis span. `ate_idx` may be passed in to reuse an
    already-parsed index (draw_all() has one); left as None, it's parsed
    here (guarded, so a Bench-only run with no ATE log doesn't crash)."""
    if PATHS.extract_mode == "bench":
        return {}
    if ate_idx is None:
        try:
            ate_idx = parse_all_ate_records(ATE_CSV)
        except (FileNotFoundError, OSError):
            return {}
    group_values = defaultdict(list)
    for combo in bench.keys():
        packet, supply, gain, freq = combo
        candidates = _ate_candidates(ate_idx, packet, supply, key)
        rec = _nearest(freq, candidates, FREQ_TOL)
        if rec is None:
            continue
        group = (packet, gain, FA.band_of(freq))
        group_values[group].extend(rec["values"])
    return {g: (min(vs), max(vs)) for g, vs in group_values.items()}


def draw_all(only_metric: str = "", limit: int = 0, count_only: bool = False):
    metric_cols = {key: col for key, col, _xlabel, _packets in FA.METRICS}
    raw, slices_seen, dig_seen = FA.read_all(PATHS.bench_dir, metric_cols)
    pa_slices_max = max(slices_seen)
    dig_gain_max = max(dig_seen)

    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    ate_idx = parse_all_ate_records(ATE_CSV)

    for key, col, xlabel, packets in FA.METRICS:
        if only_metric and key != only_metric:
            continue

        bench = FA.build_fixed_combos(raw[key], pa_slices_max, dig_gain_max, packets)
        if not bench:
            print(f"=== {key}: no Bench data -- skipping ===")
            continue

        def spec_fn(packet, band, _key=key):
            return spec_lookup.lookup(spec_table, _key, packet, band)

        ranges, step_by_group = FA.build_packet_ranges(bench, spec_fn=spec_fn)

        ate_metric = METRIC_ATE_MAP[key]
        recs_by_combo, status_by_combo = {}, {}
        for combo in bench.keys():
            packet, supply, gain, freq = combo
            candidates = _ate_candidates(ate_idx, packet, supply, key)
            rec = _nearest(freq, candidates, FREQ_TOL)
            recs_by_combo[combo] = rec
            status_by_combo[combo] = "NONE" if rec is None else ("EXACT" if rec["freq"] == freq else "APPROX")

        total = len(status_by_combo)
        exact = sum(1 for s in status_by_combo.values() if s == "EXACT")
        approx = sum(1 for s in status_by_combo.values() if s == "APPROX")
        none_ = sum(1 for s in status_by_combo.values() if s == "NONE")
        print(f"=== {key} ({ate_metric}) -- Bench combos: {total}  "
              f"ATE EXACT: {exact}  APPROX: {approx}  NONE: {none_} ===")

        if count_only:
            continue

        matched = [combo for combo, s in status_by_combo.items() if s != "NONE"]
        if limit > 0:
            matched = matched[:limit]
            print(f"  Limiting to first {len(matched)} matched combo(s)")

        png_dir = ATE_PNG_BASE / f"{key}_revision2"
        png_dir.mkdir(parents=True, exist_ok=True)

        # Precompute each combo's PPT column count from the FULL Bench combo
        # set for this key -- see make_ate_tx_pwr_revision2.py's identical
        # block for the full rationale.
        band24_freqs = defaultdict(set)
        highfreq_freqs = defaultdict(set)
        for packet, supply, _gain, freq in bench.keys():
            if FA.band_of(freq) == 2.4:
                band24_freqs[(packet, supply)].add(freq)
            else:
                highfreq_freqs[(packet, FA.band_of(freq), supply)].add(freq)
        band24_supplies = defaultdict(set)
        for (packet, supply) in band24_freqs:
            band24_supplies[packet].add(supply)
        band24_n_cols = {
            packet: max(len(band24_freqs[(packet, s)]) for s in supplies) * len(supplies)
            for packet, supplies in band24_supplies.items()
        }

        def n_cols_for(packet, supply, freq):
            if FA.band_of(freq) == 2.4:
                return band24_n_cols[packet]
            return len(highfreq_freqs[(packet, FA.band_of(freq), supply)])

        # Widen each axis-sharing group's Y-range to cover real ATE data
        # that falls outside the Bench-derived window -- see make_ate_tx_
        # pwr_revision2.py's identical block for the full rationale
        # (2026-08-14). get_ate_group_extents() is the SAME function the
        # Bench module's main() calls (2026-08-15), so both scripts arrive
        # at the identical widened+snapped range independently -- no
        # file-based coordination needed.
        for group, (lo, hi) in get_ate_group_extents(key, bench, ate_idx).items():
            if group in ranges:
                vmin, vmax = ranges[group]
                ranges[group] = FA._snap_range(min(vmin, lo), max(vmax, hi))

        jobs = []
        for combo in matched:
            packet, supply, gain, freq = combo
            rec = recs_by_combo[combo]
            group = (packet, gain, FA.band_of(freq))
            value_range = ranges.get(group)
            step = step_by_group.get(group, 1)
            spec = spec_fn(packet, FA.band_of(freq))
            name = FA.combo_name(combo, pa_slices_max, dig_gain_max)
            title = "ATE_" + FA.combo_title_short(combo, pa_slices_max, dig_gain_max)
            png_path = png_dir / f"{name}.png"
            figsize = FA.figure_size_for(n_cols_for(packet, supply, freq))
            jobs.append((rec["values"], title, xlabel, png_path, value_range, spec, step, figsize, name))

        i = 0
        with ProcessPoolExecutor(max_workers=PATHS.png_workers) as pool:
            for name in pool.map(_draw_one, jobs, chunksize=4):
                i += 1
                if i % 200 == 0 or i == len(jobs):
                    print(f"  [{i}/{len(jobs)}] {name}")

        print(f"  Done: {key} -> ATE PNG {png_dir}")


def _draw_one(job):
    """Worker for ProcessPoolExecutor -- must be a module-level function so
    it's picklable on Windows (spawn-only, no fork)."""
    values, title, xlabel, png_path, value_range, spec, step, figsize, name = job
    FA.draw_png(values, title, xlabel, png_path, value_range=value_range, spec=spec, step=step, figsize=figsize)
    return name


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--metric", type=str, default="", help="only this sub-metric key (default: all)")
    p.add_argument("--limit", type=int, default=0, help="only draw the first N matched combos per metric")
    p.add_argument("--count-only", action="store_true", help="print match-rate audit only")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    draw_all(only_metric=args.metric, limit=args.limit, count_only=args.count_only)


if __name__ == "__main__":
    main()
