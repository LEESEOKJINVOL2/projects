"""BT TX ACP -- ATE counterpart to make_bt_tx_acp_revision2.py.

ACP's bench chart pools ALL frequencies within a (packet, band) into one
histogram PER OFFSET (stacked bars, one color per frequency) -- see
make_bt_tx_acp_revision2.draw_png_single_offset. The ATE side mirrors that:
for each offset actually shown on the PPT slide (-6..6 excluding 0, per
build_ppt_acp_revision2.OFFSETS), every one of that (packet,band) group's
Bench frequencies is matched against the ATE log's ACP_{M|P}n items
(supply is always "0" for ACP) using the same packet-alias/freq-tolerance
rule as make_ate_tx_pwr_revision2.py, and every match's 40 DUT values are
pooled into an ATE freq_map keyed by the ATE item's own frequency -- so the
resulting ATE chart is visually comparable to the Bench one (same
per-frequency color stacking), not a flat single-color histogram.

CLI: --limit N (only draw the first N matched (combo,offset) pairs),
--count-only.

2026-08-14: a newer ATE pull uses a second item-name format (see
make_ate_tx_pwr_revision2.py's module docstring for the full field-by-field
derivation) -- added alongside FNAME_RE, not replacing it. ACP's own test
name is literally "ACP" in both formats; the offset lives in the new
format's second-to-last token (e.g. "M5"/"0"/"5" for -5/0/+5MHz, same
M-prefix-for-negative convention as the old ACP_{M|P}n metric suffix) with
the last token always "X" (no PortToANT-style cal variants seen on ACP).
Confirmed FIXED_SUPPLY still applies: the new format's own supply label
(1P5/1P2/0P73) is translated to the same 0/1/2 index before filtering.
"""
from __future__ import annotations

import argparse
import csv
import math
import re
import statistics
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import make_bt_tx_acp_revision2 as ACP_CHART
import bt_tx_acp_spec_lookup as spec_lookup
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")

ATE_CSV = PATHS.ate_log_path
ATE_PNG_DIR = PATHS.result_png_dir.parent / "BT_TX_ATE" / "acp_revision2"

FREQ_TOL = 5.0  # MHz
PACKET_ALIAS = {"2M": "LE2M", "VHDRPM16": "HDRPM16", "VHDRPL32": "HDRPL32"}
FIXED_SUPPLY = "0"
# Only the offsets the PPT slide actually shows (build_ppt_acp_revision2.OFFSETS)
# -- 0 is dropped there (its real ACP values fall outside the fixed axis).
OFFSETS = [o for o in range(-6, 7) if o != 0]

FNAME_RE = re.compile(
    r"^BTTX_(?P<mode>[A-Za-z0-9]+)_(?P<band>\d+G)_(?P<freq>\d+)MHZ_X_EPA_"
    r"(?P<pwr>[A-Za-z0-9]+)DBM_(?P<packet>[A-Za-z0-9]+)_PS(?P<supply>\d+)_X_"
    r"(?P<metric>[A-Za-z0-9_]+)$"
)
ACP_METRIC_RE = re.compile(r"^ACP_([MP])(\d+)$")

# -- New (2026-08-14) item-name format, see make_ate_tx_pwr_revision2.py --
# ACP's own offset lives where every other metric has a fixed "X" -- the
# generic NEW_FNAME_RE's c1 group captures it directly (M5/0/5, etc).
NEW_FNAME_RE = re.compile(
    r"^BTTX(?P<band>\d+)_(?P<packet>[A-Za-z0-9]+)_(?P<domain>[A-Za-z]+)_(?P<test>[A-Za-z0-9]+)_"
    r"CH(?P<ch>\d+)_(?P<supply>[0-9A-Za-z]+)_(?P<offset>[A-Za-z0-9]+)_(?P<c2>[A-Za-z0-9]+)_NV_VXXX$"
)
NEW_BAND_MIN_FREQ = {0: 2402, 1: 5150, 2: 5725, 4: 5925, 5: 6051, 6: 6176, 7: 6301}
NEW_SUPPLY_LABEL_TO_INDEX = {"1P5": "0", "1P2": "1", "0P73": "2"}
NEW_OFFSET_RE = re.compile(r"^M?(\d+)$")


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


def parse_ate_acp_records(path: Path):
    """-> {(packet, offset): [record, ...]}, record has freq/values. Only
    supply=='0' rows (ACP is FIXED_SUPPLY on both Bench and ATE sides).

    Tries the OLD item-name format (FNAME_RE) first, then the NEW one
    (NEW_FNAME_RE) -- see make_ate_tx_pwr_revision2.py's module docstring."""
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
                if not m or m.group("supply") != FIXED_SUPPLY:
                    continue
                mm = ACP_METRIC_RE.match(m.group("metric"))
                if not mm:
                    continue
                sign, num = mm.group(1), int(mm.group(2))
                offset = -num if sign == "M" else num
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(m.group("freq"))
            else:
                m = NEW_FNAME_RE.match(name)
                if not m or m.group("domain") != "RF" or m.group("test") != "ACP" or m.group("c2") != "X":
                    continue
                supply = NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
                if supply != FIXED_SUPPLY:
                    continue
                om = NEW_OFFSET_RE.match(m.group("offset"))
                band_min = NEW_BAND_MIN_FREQ.get(int(m.group("band")))
                if om is None or band_min is None:
                    continue
                offset = -int(om.group(1)) if m.group("offset").startswith("M") else int(om.group(1))
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(band_min + int(m.group("ch")))

            values = _values_from_row(row)
            if not values:
                continue
            idx.setdefault((packet, offset), []).append({
                "item": name, "freq": freq, "values": values,
                "ul": (row[1] or "").strip() if len(row) > 1 else "",
                "ll": (row[2] or "").strip() if len(row) > 2 else "",
            })
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


def parse_ate_acp_records_per_dut(path: Path):
    """Same return shape/matching rules as parse_ate_acp_records() above --
    {(packet, offset): [record, ...]} -- mirrored verbatim, except each
    record carries "by_serial": {dut_serial: value} (via
    _values_by_serial_from_row()) instead of a flat "values" list. 2026-08-25,
    added for the Raw Values sheet (add_bt_tx_raw_values_sheet.py);
    parse_ate_acp_records() itself is untouched."""
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
                if not m or m.group("supply") != FIXED_SUPPLY:
                    continue
                mm = ACP_METRIC_RE.match(m.group("metric"))
                if not mm:
                    continue
                sign, num = mm.group(1), int(mm.group(2))
                offset = -num if sign == "M" else num
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(m.group("freq"))
            else:
                m = NEW_FNAME_RE.match(name)
                if not m or m.group("domain") != "RF" or m.group("test") != "ACP" or m.group("c2") != "X":
                    continue
                supply = NEW_SUPPLY_LABEL_TO_INDEX.get(m.group("supply"))
                if supply != FIXED_SUPPLY:
                    continue
                om = NEW_OFFSET_RE.match(m.group("offset"))
                band_min = NEW_BAND_MIN_FREQ.get(int(m.group("band")))
                if om is None or band_min is None:
                    continue
                offset = -int(om.group(1)) if m.group("offset").startswith("M") else int(om.group(1))
                packet = PACKET_ALIAS.get(m.group("packet"), m.group("packet"))
                freq = float(band_min + int(m.group("ch")))

            by_serial = _values_by_serial_from_row(serials, row)
            if not by_serial:
                continue
            idx.setdefault((packet, offset), []).append({
                "item": name, "freq": freq,
                "by_serial": by_serial,
            })
    return idx


def get_ate_slide_extent(combo, offset_map, ate_idx=None):
    """(min_ate_value, max_ate_value) or None -- pooled across EVERY
    offset shown on this (packet,band) slide (OFFSETS), matched the same
    way as get_ate_data(). Called from BOTH this module's own draw_all()
    (via get_ate_data()) and the Bench module's main() (see
    make_bt_tx_acp_revision2.py) so both compute the identical union
    independently -- 2026-08-15 customer request: ONE shared page-level
    span (supersedes the 2026-08-14 per-offset design), and Bench/ATE
    must show that exact same span."""
    if PATHS.extract_mode == "bench":
        return None
    if ate_idx is None:
        try:
            ate_idx = parse_ate_acp_records(ATE_CSV)
        except (FileNotFoundError, OSError):
            return None
    packet, gain, band = combo
    vals = []
    for offset in OFFSETS:
        freq_map = offset_map.get(offset)
        if not freq_map:
            continue
        candidates = ate_idx.get((packet, offset), [])
        for bench_freq in freq_map:
            rec = _nearest(bench_freq, candidates, FREQ_TOL)
            if rec is not None:
                vals.extend(rec["values"])
    if not vals:
        return None
    return min(vals), max(vals)


def get_ate_data(pa_slices_max=None, dig_gain_max=None):
    """-> {(packet, gain, band, offset): {"png":.., "stats":.., "status":..,
    "freq_map": {ate_freq: values}}}, plus the raw bench `data` dict and the
    resolved pa_slices_max/dig_gain_max."""
    raw, slices_seen, dig_seen = ACP_CHART.read_all(PATHS.bench_dir)
    if pa_slices_max is None:
        pa_slices_max = max(slices_seen)
    if dig_gain_max is None:
        dig_gain_max = max(dig_seen)
    data = ACP_CHART.build_fixed_combos(raw, pa_slices_max, dig_gain_max)

    ate_idx = parse_ate_acp_records(ATE_CSV)

    out = {}
    for combo, offset_map in data.items():
        packet, gain, band = combo
        for offset in OFFSETS:
            freq_map = offset_map.get(offset)
            if not freq_map:
                continue
            candidates = ate_idx.get((packet, offset), [])
            ate_freq_map: dict[float, list[float]] = {}
            any_exact = any_approx = False
            for bench_freq in freq_map:
                rec = _nearest(bench_freq, candidates, FREQ_TOL)
                if rec is None:
                    continue
                ate_freq_map.setdefault(rec["freq"], []).extend(rec["values"])
                if rec["freq"] == bench_freq:
                    any_exact = True
                else:
                    any_approx = True
            key = (packet, gain, band, offset)
            if not ate_freq_map:
                out[key] = {"png": None, "stats": None, "status": "NONE", "freq_map": {}}
                continue
            status = "EXACT" if any_exact and not any_approx else "APPROX"
            pooled = [v for vs in ate_freq_map.values() for v in vs]
            fname = ACP_CHART.combo_name(combo, pa_slices_max, dig_gain_max)
            off_label = ACP_CHART.fmt_num(offset)
            png_path = ATE_PNG_DIR / f"{ACP_CHART.safe_filename(f'{fname}_offset({off_label})')}.png"
            out[key] = {
                "png": png_path if png_path.exists() else None,
                "stats": _stats(pooled),
                "status": status,
                "freq_map": ate_freq_map,
            }
    return out, data, pa_slices_max, dig_gain_max


def draw_all(limit: int = 0, count_only: bool = False):
    ate_data, data, pa_slices_max, dig_gain_max = get_ate_data()
    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    total = len(ate_data)
    exact = sum(1 for v in ate_data.values() if v["status"] == "EXACT")
    approx = sum(1 for v in ate_data.values() if v["status"] == "APPROX")
    none_ = sum(1 for v in ate_data.values() if v["status"] == "NONE")
    print(f"Bench ACP (combo,offset) pairs: {total}  ATE EXACT: {exact}  APPROX: {approx}  NONE: {none_}")

    if count_only:
        return

    matched = [(key, v) for key, v in ate_data.items() if v["status"] != "NONE"]
    if limit > 0:
        matched = matched[:limit]
        print(f"Limiting to first {len(matched)} matched (combo,offset) pair(s)")

    ATE_PNG_DIR.mkdir(parents=True, exist_ok=True)

    # 2026-08-15 (customer request): ONE shared Y-axis for the WHOLE slide
    # (every offset column), superseding the 2026-08-14 per-offset design
    # -- see ACP_CHART.compute_shared_span's docstring for the full
    # rationale. get_ate_slide_extent() is the SAME function the Bench
    # module's main() calls, so both scripts arrive at the identical
    # shared span independently -- no file-based coordination needed.
    ate_idx = parse_ate_acp_records(ATE_CSV)
    span_by_combo = {}
    for combo in {(p, g, b) for p, g, b, _off in ate_data}:
        packet = combo[0]
        specs = [spec_lookup.lookup(spec_table, packet, int(off)) for off in OFFSETS]
        extra_range = get_ate_slide_extent(combo, data[combo], ate_idx)
        span_by_combo[combo] = ACP_CHART.compute_shared_span(data[combo], extra_range, specs, offsets=OFFSETS)

    jobs = []
    for (packet, gain, band, offset), v in matched:
        combo = (packet, gain, band)
        name = ACP_CHART.combo_name(combo, pa_slices_max, dig_gain_max)
        off_label = ACP_CHART.fmt_num(offset)
        fname = ACP_CHART.safe_filename(f"{name}_offset({off_label})")
        title = "ATE_" + ACP_CHART.combo_title_short(combo, pa_slices_max, dig_gain_max)
        png_path = ATE_PNG_DIR / f"{fname}.png"
        spec = spec_lookup.lookup(spec_table, packet, int(offset))
        y_min, y_max = span_by_combo[combo]
        jobs.append((v["freq_map"], title, png_path, spec, fname, y_min, y_max))

    n = 0
    with ProcessPoolExecutor(max_workers=PATHS.png_workers) as pool:
        for fname in pool.map(_draw_one, jobs, chunksize=4):
            n += 1
            if n % 200 == 0 or n == len(jobs):
                print(f"  [{n}/{len(jobs)}] {fname}")

    print(f"\nDone. ATE PNG -> {ATE_PNG_DIR}")


def _draw_one(job):
    """Worker for ProcessPoolExecutor -- must be a module-level function so
    it's picklable on Windows (spawn-only, no fork)."""
    freq_map, title, png_path, spec, fname, y_min, y_max = job
    ACP_CHART.draw_png_single_offset(
        freq_map, title, "ACP (dBm)", png_path,
        y_min=y_min, y_max=y_max, spec=spec)
    return fname


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--limit", type=int, default=0, help="only draw the first N matched (combo,offset) pairs")
    p.add_argument("--count-only", action="store_true", help="print match-rate audit only")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    draw_all(limit=args.limit, count_only=args.count_only)


if __name__ == "__main__":
    main()
