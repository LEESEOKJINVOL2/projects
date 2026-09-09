"""XLSX summary with Bench-vs-ATE delta columns for Harald_260820 (the
newer retest pull) -- same shape as Cal/BT_TX/BT_RX's own xlsx summaries
(build_cal_xlsx_summary.py / BT_TX/build_bt_tx_xlsx_summary.py /
BT_RX/build_bt_rx_xlsx_summary.py), reusing the shared xlsx_summary_writer.py
helper. Written fresh (no existing Harald summary script), pulling data via
Cal's own read_bench/read_ate/build_jobs/effective_spec against the flat
Bench_40pcs/Harald_260820 + single ATE file layout (see make_harald_260820_
distribution.py) -- no Cal/BT_TX/BT_RX file is modified, only imported.

Per the same convention as the other three: Mean/Std/Min/Max (Bench + ATE)
plus 4 delta columns (|Bench-ATE|, abs()'d, blank where either side is
missing). Test Item name is the CLEANED name (make_harald_260820_
distribution.clean()), matching the PPT deck's own display convention.
"""
from __future__ import annotations

import datetime
import statistics
import sys
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_BASE_DIR = _THIS_DIR.parent
sys.path.insert(0, str(_BASE_DIR))
sys.path.insert(0, str(_BASE_DIR / "Cal"))
sys.path.insert(0, str(_THIS_DIR))

import xlsx_summary_writer as xw
import make_cal_distribution as cal
import make_harald_260820_distribution as harald

PATHS = harald.PATHS
REPORT_DATE = datetime.date.today().isoformat()


def _stats(values):
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


def build_rows():
    print("Reading Bench_Harald_260820 CSVs...")
    bench_items = cal.read_bench(PATHS.bench_dir)
    print(f"  {len(bench_items)} test item(s) found")
    print(f"Parsing ATE log ({harald.ATE_PATH.name})...")
    ate_items = cal.read_ate_auto(harald.ATE_PATH)  # ATE_PATH may be one file or a folder of per-config files
    jobs, n_no_ate = cal.build_jobs(bench_items, ate_items)
    print(f"  {len(jobs)} qualifying item(s) (>=1 Bench limit), {len(jobs) - n_no_ate} with ATE match")

    rows = []
    for job in jobs:
        bench_stats = _stats(job["bench_values"])
        ate_stats = _stats(job["ate_values"]) if job["ate_values"] else None
        ate_llim, ate_ulim = cal.effective_spec(job)
        bench_cells = _stats_cells(bench_stats) + [_r(job["bench_ulim"]), _r(job["bench_llim"])]
        ate_cells = _stats_cells(ate_stats) + [_r(ate_ulim), _r(ate_llim)]
        deltas = [xw.delta(bench_cells[i], ate_cells[i]) for i in range(4)]
        rows.append([harald.clean(job["name"]), *bench_cells, *ate_cells, *deltas])
    return rows


def main() -> None:
    rows = build_rows()
    print(f"Total rows: {len(rows)}")
    out_path = PATHS.report_dir / f"{REPORT_DATE}_Harald_260820_Summary.xlsx"
    xw.write_summary_xlsx(rows, out_path, "Harald_260820 Summary")
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
