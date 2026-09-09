"""BT TX DEVM 99pct -- ATE counterpart to make_bt_tx_devm_99pct_revision2.py.

Same matching rules as make_ate_tx_pwr_revision2.py (packet alias, band
recomputed from frequency not the ATE name's 2G/5G token, freq tolerance,
no gain/pwr disambiguation needed) -- see that module's docstring for the
full rationale.

2026-07-20 correction: unlike every other ATE-vs-Bench chart in this
project, ATE's "DEVM_99" item is NOT the same kind of measurement as
Bench's evm_99pct_aver-% column. Bench is a small EVM percentage (a few
percent, same scale as RMS/Peak). The ATE log's raw DEVM_99 values are
clustered at 99-100 (occasionally dipping as low as ~46 for a genuine
marginal/failing DUT) with its own real Lower Limit of 99.0 recorded
per-item in the log -- this reads as a pass-score/margin metric, not an
EVM magnitude, confirmed with the customer ("Lowerlimit이 99라서 값이
이상한건 아닌것 같은데 Bench와 같은 Y축을 가질수는없네": the value itself
is legitimate, it just can't share Bench's axis). Consequences:
  - The Y-axis is computed purely from ATE's OWN pooled data per
    axis-sharing group (packet, gain, band) -- never borrowed from
    Bench's build_packet_ranges(), which is on a completely different
    scale and would put every ATE bar off-chart.
  - The Limit line uses the ATE record's OWN per-item Lower/Upper Limit
    columns (parsed into rec["ll"]/rec["ul"]) instead of a
    bt_tx_devm_spec_lookup sheet value -- Bench's EVM-percentage spec has
    no meaning on ATE's pass-score scale.

CLI: --limit N (only draw the first N matched combos), --count-only.

2026-08-14: a newer ATE pull uses a second item-name format (see
make_ate_tx_pwr_revision2.py's module docstring for the full field-by-field
derivation) -- added alongside FNAME_RE, not replacing it. Only "DEVM99"
exists for this metric (no "EVM99" counterpart anywhere in the log), so
this only covers EDR-family packets (2DH5/3DH5/4DH5/8DH5); HDRP/HDT/UHDR
stay unmatched, a real gap not a parsing bug. IMPORTANT: unlike the old
DEVM_99 item (a 99-100 pass-score, see the correction above), the NEW
DEVM99 item is a normal small %EVM value with a per-item Upper Limit that
already matches bt_tx_devm_spec_lookup's own 99pct spec (e.g. 2DH5 -> UL=17,
same as the spec sheet) -- confirmed by inspecting a raw row. No change
needed to this file's "own ATE-pooled axis + own per-item ll/ul" drawing
logic below: it already generalizes correctly to both semantics, it just
happens to draw a small-percentage axis instead of a 99-100 one now.
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

import make_bt_tx_devm_99pct_revision2 as devm_mod
import bt_tx_devm_spec_lookup as spec_lookup
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")

ATE_CSV = PATHS.ate_log_path
ATE_PNG_DIR = PATHS.result_png_dir.parent / "BT_TX_ATE" / "devm_99pct_revision2"

FREQ_TOL = 5.0  # MHz
PACKET_ALIAS = {"2M": "LE2M", "VHDRPM16": "HDRPM16", "VHDRPL32": "HDRPL32", "HDT7P5": "HDT8"}
ATE_METRIC = "DEVM_99"

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
NEW_DEVM_99_TEST_NAMES = ("DEVM99",)


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
                if not m or m.group("domain") != "RF" or m.group("test") not in NEW_DEVM_99_TEST_NAMES:
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
    parse_ate_records() itself is untouched. (Note: this is the metric
    whose ATE value is a 0-100 pass-score, not an EVM percentage -- see
    this module's own docstring -- but the Raw Values sheet just wants
    whatever raw number the ATE log recorded per DUT, same as every other
    metric here.)"""
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
                if not m or m.group("domain") != "RF" or m.group("test") not in NEW_DEVM_99_TEST_NAMES:
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


def get_ate_group_extents_if_new_format(bench, pa_slices_max, dig_gain_max):
    """{group: (min_ate_value, max_ate_value)} pooled per Bench axis-sharing
    group (devm_mod.group_of), using every Bench combo's matched ATE
    record -- ONLY if every match is the 2026-08-14 new item-name format
    (a normal small-%EVM value on Bench's own scale). Returns {} if any
    matched record is the legacy 99-100 pass-score format (genuinely a
    different physical scale, see this module's docstring -- unioning
    those with Bench's axis would put every Bench bar off-chart), or if
    extract_mode=bench, or if the ATE log is missing.

    Called from BOTH this module's own draw_all() and the Bench module's
    main() (see make_bt_tx_devm_99pct_revision2.py) so both sides compute
    the identical union independently from the same two raw data sources,
    with no file-based coordination needed -- same inputs, same groups,
    same result."""
    if PATHS.extract_mode == "bench":
        return {}
    try:
        ate_idx = parse_ate_records(ATE_CSV)
    except (FileNotFoundError, OSError):
        return {}
    group_values: dict[tuple, list] = defaultdict(list)
    for combo in bench.keys():
        packet, supply, gain, freq = combo
        candidates = ate_idx.get((packet, supply), [])
        rec = _nearest(freq, candidates, FREQ_TOL)
        if rec is None:
            continue
        if rec["item"].startswith("BTTX_"):
            return {}  # legacy pass-score format present -- never union
        group = devm_mod.group_of(packet, gain, freq, supply)
        group_values[group].extend(rec["values"])
    return {g: (min(vs), max(vs)) for g, vs in group_values.items()}


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


def _to_float_or_none(text: str):
    text = (text or "").strip()
    if not text or text.upper() == "NA":
        return None
    return float(text)


def draw_all(limit: int = 0, count_only: bool = False):
    ate_data, bench, pa_slices_max, dig_gain_max = get_ate_data()

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

    # Resolve each combo's ATE record once, and pool values per axis-sharing
    # group -- see module docstring: under the LEGACY item format, ATE's
    # data is on a completely different scale from Bench's (a 99-100
    # pass-score, not a %EVM value), so it built its own independent axis,
    # purely from ATE's own pooled data.
    #
    # 2026-08-14 (customer request): "the two spans must match." Under the
    # legacy format that's still impossible (see above -- would put every
    # Bench bar off-chart) and this code intentionally keeps the old
    # ATE-only-axis behavior there. But the NEW item format's DEVM99 value
    # is a normal small-%EVM value on Bench's own scale (confirmed by
    # comparing its Upper Limit to bt_tx_devm_spec_lookup's own 99pct spec
    # -- e.g. 2DH5 -> UL=17 either way), so for that format we union Bench's
    # own build_packet_ranges() with ATE's real extent per group and use
    # that SAME union on both sides -- get_ate_group_extents_if_new_format()
    # is called from this module's own draw_all() AND from the Bench
    # module's main(), so both compute the identical union independently.
    resolved: dict[tuple, dict] = {}
    group_values: dict[tuple, list] = defaultdict(list)
    for combo, v in matched:
        packet, supply, gain, freq = combo
        candidates = ate_idx.get((packet, supply), [])
        rec = _nearest(freq, candidates, FREQ_TOL)
        resolved[combo] = rec
        group = devm_mod.group_of(packet, gain, freq, supply)
        group_values[group].extend(rec["values"])

    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    spec_by_packet = spec_table.get(devm_mod.SPEC_METRIC, {})
    bench_ranges, _kept, _outliers, _incl, bench_step_by_group = devm_mod.build_packet_ranges(bench, spec_by_packet)
    ate_extents = get_ate_group_extents_if_new_format(bench, pa_slices_max, dig_gain_max)

    ate_ranges: dict[tuple, tuple] = {}
    step_by_group: dict[tuple, float] = {}
    for group, values in group_values.items():
        vmin, vmax = math.floor(min(values)), math.ceil(max(values))
        if group in ate_extents and group in bench_ranges:
            bvmin, bvmax = bench_ranges[group]
            vmin, vmax = min(vmin, bvmin), max(vmax, bvmax)
            step_by_group[group] = bench_step_by_group.get(group, max(1, math.ceil((vmax - vmin) / 15)))
        else:
            step_by_group[group] = max(1, math.ceil((vmax - vmin) / 15))
        ate_ranges[group] = (vmin, vmax)

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

    jobs = []
    for combo, v in matched:
        packet, supply, gain, freq = combo
        rec = resolved[combo]
        group = devm_mod.group_of(packet, gain, freq, supply)
        value_range = ate_ranges.get(group)
        step = step_by_group.get(group, 1)
        # ATE's own per-item Lower/Upper Limit -- NOT bt_tx_devm_spec_lookup
        # (that's Bench's EVM-percentage spec, meaningless on this scale).
        ll, ul = _to_float_or_none(rec["ll"]), _to_float_or_none(rec["ul"])
        spec = (ll, ul) if (ll is not None or ul is not None) else None
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
    devm_mod.draw_png(values, title, "DEVM 99pct (%)", png_path, value_range=value_range, step=step, spec=spec,
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
