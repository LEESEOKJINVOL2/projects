"""BT TX Power -- ATE counterpart to make_bt_tx_pwr_revision2.py.

Reads the ATE log (config.xml's per-project ate_log_path -- the exact log FILE, 2026-07-22)
and, for every Bench Power combo (packet_type, pa_supply, pa_gain,
frequency-MHz) that has a matching ATE item, draws an ATE histogram PNG in
the exact same visual style as make_bt_tx_pwr_revision2.draw_png (reused
directly), saved under Result_PNG/BT_TX_ATE/pwr_revision2 using the
IDENTICAL filename convention as the Bench PNG so build_ppt_pwr_evm_revision2.py
can place Bench (top) and ATE (bottom) side by side on one slide, one image
per column.

Matching rules (established by comparing Bench CSVs against
ate_40pcs/40pcs_raw_data_transposed.csv):
  - ATE item name: BTTX_{mode}_{2G|5G}_{freq}MHZ_X_EPA_{pwr}DBM_{packet}_
    PS{supply}_X_{metric}. Power uses metric == "PWR_AVG".
  - Packet alias (ATE -> Bench): 2M -> LE2M, VHDRPM16 -> HDRPM16,
    VHDRPL32 -> HDRPL32.
  - Band is NOT taken from the ATE name's 2G/5G token -- 6GHz conditions
    are mislabeled "5G" there. Matching is done purely on packet+supply+
    frequency (with tolerance), same as Bench's own band_of() would bucket
    it, so this mislabeling never matters.
  - Frequency: exact match preferred; else nearest ATE freq within
    FREQ_TOL (5MHz), covering the ~1-4MHz channel-plan near-misses seen in
    the HDRP/UHDR families.
  - gain/pwr are NOT matching keys: verified empirically that for a fixed
    (packet, supply, freq), Bench has exactly one pa_gain value and ATE
    has exactly one EPA power label (0 duplicates on either side), so no
    disambiguation beyond packet+supply+freq is needed.
  - Roughly 30% of Bench combos have no ATE counterpart at all -- mostly
    because the condition isn't in the ATE test plan (e.g. 4DH5/8DH5 are
    entirely absent). That's expected, not a bug; get_ate_data() reports
    those combos with png=None so the PPT builder can render a "No ATE
    match" gap instead of blocking the whole slide.

2026-08-14: a newer ATE pull (ate_40pcs/ATE_10pcs_Data.csv) uses a second,
completely different item-name scheme -- added ALONGSIDE the rules above
(not replacing them), since switching between an old-style and new-style
ATE pull should stay a one-line config.xml edit, not a code change. New
format: BTTX{band}_{packet}_{domain}_{test}_CH{ch}_{supply}_{c1}_{c2}_NV_VXXX,
e.g. "BTTX0_1DH5_RF_PWRA_CH000_1P5_X_X_NV_VXXX". Derived from value-
distribution analysis (no format spec was provided), cross-checked against
the spec workbook's own 'Frequency Table' sheet:
  - {band}: the customer's own Band number (0,1,2,4,5,6,7 -- no Band 3,
    consistent with every other spec sheet in this project). Confirmed by
    the per-band max CH{ch} values matching that band's channel count in
    the Frequency Table sheet.
  - CH{ch}: a channel INDEX, uniform 1MHz spacing within its band --
    freq_MHz = NEW_BAND_MIN_FREQ[band] + ch. Confirmed: band=0/ch=39 ->
    2441MHz (the Frequency Table's own Band0 "Mid" reference); band=0/ch=0
    -> 2402MHz, matching a real Bench 1DH5 condition exactly.
  - {supply}: VDD_HPA voltage label (1P5/1P2/0P73), maps 1:1 to the same
    0/1/2 pa_supply index as Bench and the new TX Power Specification sheet
    (see bt_tx_power_spec_lookup.PA_SUPPLY_VOLT_TO_INDEX). "0P73LP" (a low-
    power variant) has no counterpart anywhere in the spec workbook and is
    deliberately left unmapped, not guessed.
  - {test}: legacy modulations (1DH5/LE1M/LE2M -- BDR/LE) get a separate
    Average/Peak split (PWRA/PWRP); every other packet family (HDRP/HDT/
    UHDR/...) has only one "PWR" test, confirmed empirically to never
    co-exist with PWRA/PWRP for the same packet. This script wants the
    AVERAGE reading (Bench's power_aver-dBm): PWRA when present, else the
    packet's only PWR. PWRP (Peak) is for a future Power-Peak ATE script.

    EDR packets (2DH5/3DH5/4DH5/8DH5) have NEITHER PWRA nor plain PWR --
    Power is split into DPSKPWR (8DPSK/DQPSK payload) and GFSKPWR (GFSK
    header) instead, since EDR packets transmit both a GFSK header and a
    DPSK/8DPSK payload. Deliberately left unmapped (2026-08-14, customer
    decision: "skip for now, plan to apply DPSKPWR later") -- these 15
    combos report "No ATE match" until that's confirmed. Bench's own
    differential_power_{min,aver,max}-dB column is a plausible cross-check
    once a candidate is picked (likely GFSK-vs-DPSK delta), but nothing
    reads it yet.
  - {c1}_{c2}: almost always "X_X" (the standard conducted reading). A
    handful of items at specific validation channels (e.g. band0's Mid
    channel) instead carry 4 extra "PortToANT_ANT"/"PortToCOND_COND"/etc.
    variants -- calibration spot-checks, not the Bench-comparable reading,
    so only "X_X" is matched.

CLI: --limit N (only process the first N matched combos, for a quick
check), --count-only (print the match-rate audit without drawing PNGs).
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

import make_bt_tx_pwr_revision2 as pwr_mod
import bt_tx_power_spec_lookup as spec_lookup
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")

ATE_CSV = PATHS.ate_log_path
ATE_PNG_DIR = PATHS.result_png_dir / "pwr_revision2"
ATE_PNG_DIR = PATHS.result_png_dir.parent / "BT_TX_ATE" / "pwr_revision2"

FREQ_TOL = 5.0  # MHz
PACKET_ALIAS = {
    "2M": "LE2M", "VHDRPM16": "HDRPM16", "VHDRPL32": "HDRPL32",
    # 2026-08-14, new-format-only: "HDT7P5" is the only HDT-family label the
    # new ATE pull has (no "HDT6"/"HDT8" literal token anywhere in it), and
    # its channel coverage (2.4/5/6GHz, all 3 bands) matches Bench's HDT8
    # exactly -- inferred, not from a supplied mapping table.
    "HDT7P5": "HDT8",
}
ATE_METRIC = "PWR_AVG"

FNAME_RE = re.compile(
    r"^BTTX_(?P<mode>[A-Za-z0-9]+)_(?P<band>\d+G)_(?P<freq>\d+)MHZ_X_EPA_"
    r"(?P<pwr>[A-Za-z0-9]+)DBM_(?P<packet>[A-Za-z0-9]+)_PS(?P<supply>\d+)_X_"
    r"(?P<metric>[A-Za-z0-9_]+)$"
)

# -- New (2026-08-14) item-name format, see module docstring --
NEW_FNAME_RE = re.compile(
    r"^BTTX(?P<band>\d+)_(?P<packet>[A-Za-z0-9]+)_(?P<domain>[A-Za-z]+)_(?P<test>[A-Za-z0-9]+)_"
    r"CH(?P<ch>\d+)_(?P<supply>[0-9A-Za-z]+)_(?P<c1>[A-Za-z0-9]+)_(?P<c2>[A-Za-z0-9]+)_NV_VXXX$"
)
NEW_BAND_MIN_FREQ = {0: 2402, 1: 5150, 2: 5725, 4: 5925, 5: 6051, 6: 6176, 7: 6301}
NEW_SUPPLY_LABEL_TO_INDEX = {"1P5": "0", "1P2": "1", "0P73": "2"}
NEW_POWER_TEST_NAMES = ("PWRA", "PWR")  # tries PWRA (Average) first, falls back to the packet's only PWR


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
    """-> {(packet, supply): [record, ...]}, record has freq/pwr/item/values/ul/ll.

    Tries the OLD item-name format (FNAME_RE) first, then the NEW one
    (NEW_FNAME_RE) -- see module docstring for both formats' field meanings.
    A given ATE pull only ever matches one of the two in practice (the
    "BTTX_" vs "BTTX<digit>_" prefixes are mutually exclusive), so trying
    both here means switching between an old-style and new-style ATE pull
    stays a one-line config.xml <ate_log_path> edit, not a code change.

    2026-07-20: @lru_cache -- get_ate_data() and draw_all() each independently
    re-parsed this same 26k-row ATE CSV in one process run; caching means
    the second call is free. Safe: nothing downstream mutates the returned
    index dict/lists."""
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
                pwr = m.group("pwr")
            else:
                if metric != ATE_METRIC:
                    continue  # the new format only has a Power (this script's) counterpart so far
                m = NEW_FNAME_RE.match(name)
                if not m or m.group("domain") != "RF" or m.group("test") not in NEW_POWER_TEST_NAMES:
                    continue
                if m.group("c1") != "X" or m.group("c2") != "X":
                    continue  # PortToANT/PortToCOND cal spot-checks, not the standard conducted reading
                supply = NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
                band_min = NEW_BAND_MIN_FREQ.get(int(m.group("band")))
                if supply is None or band_min is None:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(band_min + int(m.group("ch")))
                pwr = m.group("test")

            values = _values_from_row(row)
            if not values:
                continue
            record = {
                "item": name, "freq": freq, "pwr": pwr,
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
                pwr = m.group("pwr")
            else:
                if metric != ATE_METRIC:
                    continue
                m = NEW_FNAME_RE.match(name)
                if not m or m.group("domain") != "RF" or m.group("test") not in NEW_POWER_TEST_NAMES:
                    continue
                if m.group("c1") != "X" or m.group("c2") != "X":
                    continue
                supply = NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
                band_min = NEW_BAND_MIN_FREQ.get(int(m.group("band")))
                if supply is None or band_min is None:
                    continue
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(band_min + int(m.group("ch")))
                pwr = m.group("test")

            by_serial = _values_by_serial_from_row(serials, row)
            if not by_serial:
                continue
            record = {
                "item": name, "freq": freq, "pwr": pwr,
                "by_serial": by_serial,
            }
            idx.setdefault((packet, supply), []).append(record)
    return idx


def get_ate_data(pa_slices_max=None, dig_gain_max=None):
    """-> {(packet, supply, gain, freq): {"png": Path|None, "stats": (mean,min,max,std)|None,
    "status": "EXACT"|"APPROX"|"NONE", "item": str|None}}

    Re-derives the Bench combo set directly from the CSVs (same pattern as
    build_ppt_pwr_evm_revision2.get_ranges_and_stats -- independent re-read,
    not dependent on already-generated PNGs) so this stays correct even if
    Result_PNG hasn't been (re)built yet.
    """
    raw, slices_seen, dig_seen = pwr_mod.read_all(PATHS.bench_dir)
    if pa_slices_max is None:
        pa_slices_max = max(slices_seen)
    if dig_gain_max is None:
        dig_gain_max = max(dig_seen)
    bench = pwr_mod.build_fixed_combos(raw, pa_slices_max, dig_gain_max)

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
        fname = pwr_mod.combo_name(combo, pa_slices_max, dig_gain_max)
        png_path = ATE_PNG_DIR / f"{fname}.png"
        out[combo] = {
            "png": png_path if png_path.exists() else None,
            "stats": _stats(rec["values"]),
            "status": status,
            "item": rec["item"],
        }
    return out, bench, pa_slices_max, dig_gain_max


def get_ate_group_extents(bench, pa_slices_max, dig_gain_max):
    """{(packet,gain,band): (min_ate_value, max_ate_value)} pooled from
    every Bench combo's matched ATE record. Called from BOTH this module's
    own draw_all() and the Bench module's main() (see make_bt_tx_pwr_
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
        group = (packet, gain, pwr_mod.band_of(freq))
        group_values[group].extend(rec["values"])
    return {g: (min(vs), max(vs)) for g, vs in group_values.items()}


def draw_all(limit: int = 0, count_only: bool = False):
    ate_data, bench, pa_slices_max, dig_gain_max = get_ate_data()
    packet_ranges, _outliers = pwr_mod.build_packet_ranges(bench)
    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    total = len(ate_data)
    exact = sum(1 for v in ate_data.values() if v["status"] == "EXACT")
    approx = sum(1 for v in ate_data.values() if v["status"] == "APPROX")
    none_ = sum(1 for v in ate_data.values() if v["status"] == "NONE")
    print(f"Bench Power combos: {total}  ATE EXACT: {exact}  APPROX: {approx}  NONE: {none_}")

    if count_only:
        return

    matched = [(combo, v) for combo, v in ate_data.items() if v["status"] != "NONE"]
    if limit > 0:
        matched = matched[:limit]
        print(f"Limiting to first {len(matched)} matched combo(s)")

    ATE_PNG_DIR.mkdir(parents=True, exist_ok=True)
    ate_idx = parse_ate_records(ATE_CSV)

    # Precompute each combo's PPT column count from the FULL Bench combo set
    # (bench.keys(), not just the ATE-matched subset) -- the ATE image lands
    # in the same-width column as its Bench sibling regardless of how many
    # OTHER columns on that slide have an ATE match. See make_bt_tx_pwr_
    # revision2.py's identical block for the full rationale.
    band24_freqs = defaultdict(set)
    highfreq_freqs = defaultdict(set)
    for packet, supply, _gain, freq in bench.keys():
        if pwr_mod.band_of(freq) == 2.4:
            band24_freqs[(packet, supply)].add(freq)
        else:
            highfreq_freqs[(packet, pwr_mod.band_of(freq), supply)].add(freq)
    band24_supplies = defaultdict(set)
    for (packet, supply) in band24_freqs:
        band24_supplies[packet].add(supply)
    band24_n_cols = {
        packet: max(len(band24_freqs[(packet, s)]) for s in supplies) * len(supplies)
        for packet, supplies in band24_supplies.items()
    }

    def n_cols_for(packet, supply, freq):
        if pwr_mod.band_of(freq) == 2.4:
            return band24_n_cols[packet]
        return len(highfreq_freqs[(packet, pwr_mod.band_of(freq), supply)])

    combo_recs = {}
    for combo, v in matched:
        packet, supply, gain, freq = combo
        candidates = ate_idx.get((packet, supply), [])
        combo_recs[combo] = _nearest(freq, candidates, FREQ_TOL)

    # Widen each (packet, gain, band) axis-sharing group's Y-range to cover
    # real ATE data that falls outside the Bench-derived window (2026-08-14
    # bug: np.histogram(range=value_range) silently drops out-of-range
    # values, so a combo whose ATE mean sits below/above the Bench axis
    # rendered a totally blank chart despite having real matched data).
    # get_ate_group_extents() is the SAME function the Bench module's
    # main() calls (2026-08-15), so both scripts arrive at the identical
    # widened range independently -- no file-based coordination needed.
    for group, (lo, hi) in get_ate_group_extents(bench, pa_slices_max, dig_gain_max).items():
        if group in packet_ranges:
            vmin, vmax = packet_ranges[group]
            packet_ranges[group] = (min(vmin, lo), max(vmax, hi))

    # Fold every combo's own LSL/USL into its group's span, then snap to
    # the nice 0.5 grid with clearance -- same rule and rationale as
    # make_bt_tx_pwr_revision2.main()'s identical block (2026-08-15
    # customer request: Limit line must always be visibly inside the
    # chart, and Bench/ATE must share the exact same final span).
    for group in list(packet_ranges.keys()):
        packet, gain, band = group
        vmin_raw, vmax_raw = packet_ranges[group]
        for p, supply, g, freq in bench.keys():
            if p != packet or g != gain or pwr_mod.band_of(freq) != band:
                continue
            spec = spec_lookup.lookup(spec_table, packet, freq, supply)
            if spec is None:
                continue
            lsl, usl = spec
            if lsl is not None:
                vmin_raw = min(vmin_raw, lsl)
            if usl is not None:
                vmax_raw = max(vmax_raw, usl)
        packet_ranges[group] = pwr_mod._snap_range(vmin_raw, vmax_raw)

    jobs = []
    for combo, v in matched:
        packet, supply, gain, freq = combo
        rec = combo_recs[combo]
        band = pwr_mod.band_of(freq)
        value_range = packet_ranges.get((packet, gain, band))
        spec = spec_lookup.lookup(spec_table, packet, freq, supply)
        name = pwr_mod.combo_name(combo, pa_slices_max, dig_gain_max)
        title = "ATE_" + pwr_mod.combo_title_short(combo, pa_slices_max, dig_gain_max)
        png_path = ATE_PNG_DIR / f"{name}.png"
        figsize = pwr_mod.figure_size_for(n_cols_for(packet, supply, freq))
        jobs.append((rec["values"], title, png_path, value_range, spec, figsize, name, rec["item"]))

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
    values, title, png_path, value_range, spec, figsize, name, item = job
    pwr_mod.draw_png(values, title, "Power (dBm)", png_path, value_range=value_range, spec=spec, figsize=figsize)
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
