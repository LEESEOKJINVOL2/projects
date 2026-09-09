"""XLSX summary with Bench-vs-ATE delta columns for BT_RX (RSSI /
Sensitivity_PER / Sensitivity_BER). No existing summary script for this
project (only BT_TX/NB_Tx/UWB had one) -- written fresh, modeled on
BT_TX/build_bt_tx_summary_csv.py's row shape but adapted to BT_RX's own
build_jobs() convention (bt_rx_ate_lookup.py / make_bt_rx_*_revision1.py).
No BT_RX file is modified, only imported.

Per user request 2026-08-15: Mean/Std/Min/Max (Bench + ATE) plus 4 delta
columns (|Bench-ATE|, abs()'d via the shared xlsx_summary_writer.delta,
blank where either side is missing).
"""
from __future__ import annotations

import datetime
import statistics
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import paths as bt_rx_paths
import xlsx_summary_writer as xw

import bt_rx_ate_lookup as rx
import make_bt_rx_rssi_revision1 as rssi_mod
import make_bt_rx_sensitivity_per_revision1 as per_mod
import make_bt_rx_sensitivity_ber_revision1 as ber_mod

PATHS = bt_rx_paths.get_paths("BT_RX")
REPORT_DATE = datetime.date.today().isoformat()

# Same (label, module) list as build_ppt_bt_rx_revision1.SECTIONS, minus the
# PNG-subdir/axis-label fields this script has no use for.
SECTIONS = [
    ("RSSI", rssi_mod),
    ("Sensitivity_PER", per_mod),
    ("Sensitivity_BER", ber_mod),
]


def _stats(values):
    """(mean, min, max, std) -- same shape/order as build_ppt_bt_rx_
    revision1.py's own _stats (that copy lives in the PPT builder, not a
    shared module, so this is its own copy here, not an import)."""
    if not values:
        return None
    return (statistics.mean(values), min(values), max(values),
            statistics.stdev(values) if len(values) > 1 else 0.0)


def _r(v):
    return round(v, 4) if isinstance(v, (int, float)) else v


def _stats_cells(stats):
    if not stats:
        return [None, None, None, None]
    mean, vmin, vmax, std = stats
    return [_r(mean), _r(std), _r(vmin), _r(vmax)]


def _spec_cells(spec):
    """spec: (lsl, usl) or None -- same order convention bt_rx_ate_lookup's
    parse_ate_items uses (this project's spec source: each ATE row's own
    Upper/Lower Limit, no Bench spec-sheet -- see that module's docstring)."""
    if not spec:
        return [None, None]
    lsl, usl = spec
    return [_r(usl), _r(lsl)]


def build_rows():
    rows = []
    for label, mod in SECTIONS:
        print(f"Building {label} rows...")
        jobs, _audit = mod.build_jobs()
        for job in jobs:
            name = f"BT_RX_{label}_{rx.combo_name(job['packet'], job['band'], job['ch'], job['pwr'])}"
            bench_stats = _stats(job["bench_values"])
            ate_stats = _stats(job["ate_values"]) if job["ate_values"] else None
            spec = job["spec"]
            bench_cells = _stats_cells(bench_stats) + _spec_cells(spec)
            ate_cells = _stats_cells(ate_stats) + _spec_cells(spec)
            deltas = [xw.delta(bench_cells[i], ate_cells[i]) for i in range(4)]
            rows.append([name, *bench_cells, *ate_cells, *deltas])
        print(f"  {len(jobs)} row(s)")
    return rows


def main() -> None:
    rows = build_rows()
    print(f"Total rows: {len(rows)}")
    out_path = PATHS.report_dir / f"{REPORT_DATE}_BT_RX_Summary.xlsx"
    xw.write_summary_xlsx(rows, out_path, "BT_RX Summary")
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
