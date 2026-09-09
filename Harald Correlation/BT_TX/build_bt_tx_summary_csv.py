"""Customer-requested summary: one row per Test Item (== one histogram
chart == one PPT table column) across every BT_TX metric family, Bench
and ATE side by side -- Mean/Std/Min/Max/USL/LSL for each.

CSV, not Excel (2026-07-22 customer request, replacing the earlier
build_xlsx_summary.py) -- a single flat header row (`Bench Mean`, `Bench
Std`, ..., `ATE Mean`, ...) instead of the xlsx version's merged 2-row
header, since CSV has no cell-merging concept. Row/column data is
otherwise identical.

Reuses the exact same stats/spec/ATE-matching functions build_ppt_full_
revision2.py and build_ppt_freqacc_revision2.py already call to build the
PPT -- this script never re-derives a number, it just tabulates the same
values into a flat file instead of a slide grid.

Covers all 12 metric families: Power, DEVM RMS/Peak/99pct, ACP, and
FreqAcc's 9 sub-metrics (ICFT/CFO/Drift Rate/Df1 Avg/Df2 Avg/Df2 Max/
omega_i/omega_o/omega_i+omega_o). DEVM_99pct's ATE USL/LSL is the one
exception that does NOT come from the Bench spec sheet -- see
make_ate_tx_devm_99pct_revision2.py's module docstring: ATE's DEVM_99 item
is a pass-score metric with its own per-item Lower/Upper Limit recorded in
the ATE log itself, not the same scale as Bench's EVM-percentage spec.

extract_mode="bench" (see paths.py) makes every ATE column blank, same
convention as the PPT.
"""

from __future__ import annotations

import csv
import datetime
import statistics

import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

import build_ppt_pwr_evm_revision2 as PE
import build_ppt_freqacc_revision2 as FREQACC_BUILD
import build_ppt_acp_revision2 as ACP_BUILD
import make_bt_tx_pwr_revision2 as pwr_mod
import make_bt_tx_freqacc_revision2 as fa_mod
import bt_tx_power_spec_lookup as power_spec_lookup
import bt_tx_devm_spec_lookup as devm_spec_lookup
import bt_tx_acp_spec_lookup as acp_spec_lookup
import bt_tx_freqacc_spec_lookup as freqacc_spec_lookup
import make_ate_tx_pwr_revision2 as ate_pwr_mod
import make_ate_tx_devm_rms_revision2 as ate_devm_rms_mod
import make_ate_tx_devm_peak_revision2 as ate_devm_peak_mod
import make_ate_tx_devm_99pct_revision2 as ate_devm_99pct_mod
import make_ate_tx_acp_revision2 as ate_acp_mod
import make_ate_tx_freqacc_revision2 as ate_fa_mod
import make_bt_tx_bw_revision1 as bw_mod
import make_ate_tx_bw_revision1 as ate_bw_mod

PATHS = bt_tx_paths.get_paths("BT_TX")
REPORT_DATE = datetime.date.today().isoformat()

SUB_HEADERS = ["Mean", "Std", "Min", "Max", "USL", "LSL"]
HEADER_ROW = ["Test Item"] + [f"Bench {h}" for h in SUB_HEADERS] + [f"ATE {h}" for h in SUB_HEADERS]


def _r(v):
    return round(v, 4) if isinstance(v, (int, float)) else v


def _stats_cells(stats):
    """stats: (mean, min, max, std) or None -> [mean, std, min, max]."""
    if not stats:
        return [None, None, None, None]
    mean, vmin, vmax, std = stats
    return [_r(mean), _r(std), _r(vmin), _r(vmax)]


def _spec_cells(spec):
    """spec: (lsl, usl) or None -> [usl, lsl] (matches SUB_HEADERS order)."""
    if not spec:
        return [None, None]
    lsl, usl = spec
    return [_r(usl), _r(lsl)]


def _row(name, bench_stats, bench_spec, ate_stats, ate_spec):
    return [name, *_stats_cells(bench_stats), *_spec_cells(bench_spec),
            *_stats_cells(ate_stats), *_spec_cells(ate_spec)]


def _ate_or_empty(mod, label: str) -> dict:
    """Mirrors build_ppt_full_revision2.py's _ate_data_or_empty -- same
    extract_mode="bench" gate, same reasoning."""
    if PATHS.extract_mode == "bench":
        print(f"  extract_mode=bench -- skipping ATE for {label}")
        return {}
    ate_data, _bench, _sm, _dm = mod.get_ate_data()
    return ate_data


def build_power_rows(power_spec_table) -> list[list]:
    print("Computing Power rows...")
    _ranges, stats = PE.get_ranges_and_stats(PE.pwr_mod)
    ate_data = _ate_or_empty(ate_pwr_mod, "Power")
    rows = []
    for combo, bench_stats in stats.items():
        packet, supply, _gain, freq = combo
        name = f"BT_TX_Power_{packet}_HPA{pwr_mod.hpa_token(supply)}_{pwr_mod.fmt_num(freq)}MHz"
        spec = power_spec_lookup.lookup(power_spec_table, packet, freq, supply)
        info = ate_data.get(combo)
        ate_stats = info["stats"] if info else None
        rows.append(_row(name, bench_stats, spec, ate_stats, spec))
    return rows


DEVM_SECTIONS = [
    ("rms", "DEVM_RMS", PE.devm_rms_mod, ate_devm_rms_mod),
    ("peak", "DEVM_Peak", PE.devm_peak_mod, ate_devm_peak_mod),
    ("99pct", "DEVM_99pct", PE.devm_99pct_mod, ate_devm_99pct_mod),
]


def _devm_99pct_ate_spec(combo):
    """The one exception: ATE's DEVM_99 Limit comes from the ATE record's
    OWN per-item Lower/Upper Limit column, not the Bench spec sheet -- see
    module docstring. Mirrors make_ate_tx_devm_99pct_revision2.draw_all()'s
    own lookup exactly."""
    packet, supply, _gain, freq = combo
    ate_idx = ate_devm_99pct_mod.parse_ate_records(ate_devm_99pct_mod.ATE_CSV)
    candidates = ate_idx.get((packet, supply), [])
    rec = ate_devm_99pct_mod._nearest(freq, candidates, ate_devm_99pct_mod.FREQ_TOL)
    if rec is None:
        return None
    ll = ate_devm_99pct_mod._to_float_or_none(rec["ll"])
    ul = ate_devm_99pct_mod._to_float_or_none(rec["ul"])
    if ll is None and ul is None:
        return None
    return (ll, ul)


def build_devm_rows(devm_spec_table) -> list[list]:
    rows = []
    for devm_metric, label, mod, ate_mod in DEVM_SECTIONS:
        print(f"Computing {label} rows...")
        _ranges, stats = PE.get_ranges_and_stats(mod, devm_spec_table.get(devm_metric, {}))
        ate_data = _ate_or_empty(ate_mod, label)
        for combo, bench_stats in stats.items():
            packet, supply, _gain, freq = combo
            name = f"BT_TX_{label}_{packet}_HPA{pwr_mod.hpa_token(supply)}_{pwr_mod.fmt_num(freq)}MHz"
            bench_spec = devm_spec_lookup.lookup(devm_spec_table, devm_metric, packet)
            info = ate_data.get(combo)
            ate_stats = info["stats"] if info else None
            if devm_metric == "99pct" and PATHS.extract_mode != "bench":
                ate_spec = _devm_99pct_ate_spec(combo)
            else:
                ate_spec = bench_spec
            rows.append(_row(name, bench_stats, bench_spec, ate_stats, ate_spec))
    return rows


def build_acp_rows(acp_spec_table) -> list[list]:
    print("Computing ACP rows...")
    data = ACP_BUILD.discover(ACP_BUILD.ACP_PNG_DIR)
    offset_stats_by_combo = ACP_BUILD.get_offset_stats_by_combo()
    ate_data = _ate_or_empty(ate_acp_mod, "ACP")
    rows = []
    for packet, band in sorted(data.keys()):
        band_f = float(band)
        offset_stats = offset_stats_by_combo.get((packet, band), {})
        for offset in sorted(offset_stats.keys()):
            bench_stats = offset_stats[offset]
            name = (f"BT_TX_ACP_{packet}_Band{ACP_BUILD.fmt_band(band)}G_"
                    f"HPA{pwr_mod.hpa_token(ACP_BUILD.ACP_CHART.FIXED_SUPPLY)}_Offset{offset}")
            spec = acp_spec_lookup.lookup(acp_spec_table, packet, int(offset))
            info = ate_data.get((packet, "0", band_f, offset))
            ate_stats = info["stats"] if info else None
            rows.append(_row(name, bench_stats, spec, ate_stats, spec))
    return rows


FA_NAME_SAFE = {
    "icft": "ICFT",
    "cfo": "CFO",
    "drift_rate": "DriftRate",
    "df1_avg": "Df1Average",
    "df2_avg": "Df2Average",
    "df2_max": "Df2Max",
    "omega_i": "OmegaI",
    "omega_o": "OmegaO",
    "omega_io": "OmegaIO",
    "freq_drift": "DriftRateHDT",
}


def build_freqacc_rows(spec_table) -> list[list]:
    print("Computing FreqAcc rows (all 9 sub-metrics)...")
    _ranges_by_metric, stats_by_metric = FREQACC_BUILD.get_ranges_and_stats_by_metric(spec_table)
    ate_data_by_metric = FREQACC_BUILD._ate_data_by_metric_or_empty()
    rows = []
    for key, _col, _xlabel, _packets in fa_mod.METRICS:
        stats = stats_by_metric.get(key)
        if not stats:
            continue
        ate_data = ate_data_by_metric.get(key, {})
        label = FA_NAME_SAFE.get(key, key)
        for combo, bench_stats in stats.items():
            packet, supply, _gain, freq = combo
            name = f"BT_TX_{label}_{packet}_HPA{pwr_mod.hpa_token(supply)}_{fa_mod.fmt_num(freq)}MHz"
            band = fa_mod.band_of(freq)
            spec = freqacc_spec_lookup.lookup(spec_table, key, packet, band)
            info = ate_data.get(combo)
            ate_stats = info["stats"] if info else None
            rows.append(_row(name, bench_stats, spec, ate_stats, spec))
    return rows


def _stats_from_raw(values) -> tuple | None:
    """[(value, dut_id), ...] -> (mean, min, max, std) or None -- same
    shape every other build_*_rows() function's bench_stats already uses."""
    if not values:
        return None
    vals = [v for v, _dut in values]
    return (statistics.mean(vals), min(vals), max(vals),
            statistics.stdev(vals) if len(vals) > 1 else 0.0)


def build_bw_rows() -> list[list]:
    """6dB/20dB Bandwidth (2026-08-26 user request). No spec/USL-LSL for
    either item -- not requested, left blank rather than guessed. 20dB has
    no ATE counterpart at all (confirmed empty on the real ATE log), so its
    ATE columns are always blank -- not extract_mode-gated like 6dB's."""
    print("Computing 6dB/20dB BW rows...")
    raw, slices_seen, dig_seen = bw_mod.read_all(PATHS.bench_dir)
    pa_slices_max = max(slices_seen)
    dig_gain_max = max(dig_seen)
    bench_bw6 = bw_mod.build_fixed_combos(raw["bw6"], pa_slices_max, dig_gain_max)
    bench_bw20 = bw_mod.build_fixed_combos(raw["bw20"], pa_slices_max, dig_gain_max)

    ate_data = _ate_or_empty(ate_bw_mod, "6dB BW")

    rows = []
    for combo, values in bench_bw6.items():
        packet, supply, _gain, freq = combo
        name = f"BT_TX_BW06DB_{packet}_HPA{pwr_mod.hpa_token(supply)}_{pwr_mod.fmt_num(freq)}MHz"
        bench_stats = _stats_from_raw(values)
        info = ate_data.get(combo)
        ate_stats = info["stats"] if info else None
        rows.append(_row(name, bench_stats, None, ate_stats, None))
    for combo, values in bench_bw20.items():
        packet, supply, _gain, freq = combo
        name = f"BT_TX_BW20DB_{packet}_HPA{pwr_mod.hpa_token(supply)}_{pwr_mod.fmt_num(freq)}MHz"
        bench_stats = _stats_from_raw(values)
        rows.append(_row(name, bench_stats, None, None, None))
    return rows


def write_csv(rows: list[list], out_path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(HEADER_ROW)
        writer.writerows(rows)


def main() -> None:
    print("Parsing spec workbooks...")
    power_spec_table = power_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    devm_spec_table = devm_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    acp_spec_table = acp_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    freqacc_spec_table = freqacc_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    rows = []
    rows += build_power_rows(power_spec_table)
    rows += build_devm_rows(devm_spec_table)
    rows += build_acp_rows(acp_spec_table)
    rows += build_freqacc_rows(freqacc_spec_table)
    rows += build_bw_rows()
    print(f"Total rows: {len(rows)}")

    report_dir = PATHS.report_dir
    out_path = report_dir / f"{REPORT_DATE}_BT_TX_Summary.csv"
    write_csv(rows, out_path)
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
