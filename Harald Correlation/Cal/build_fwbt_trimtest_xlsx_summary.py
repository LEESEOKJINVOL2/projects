"""XLSX summary for the FWBT_TRIMTEST (DP6P1P3-spec) Cal pipeline -- same
Bench/ATE/Delta Mean/Std/Min/Max/USL/LSL shape as the other projects'
summaries (build_cal_xlsx_summary.py / BT_TX/build_bt_tx_xlsx_summary.py /
BT_RX/build_bt_rx_xlsx_summary.py / Harald/build_harald_trimtest_xlsx_
summary.py), via the shared xlsx_summary_writer.py.

No existing summary script covered this specific pipeline yet --
build_cal_xlsx_summary.py summarizes the OTHER (plain, non-trimtest) Cal
report, reading make_cal_distribution directly. This one reuses make_bt_tx_
trimtest_distribution.py's own load_item_specs/build_jobs (imported, not
modified), so Mean/Std/Min/Max here already reflect the same outlier-
exclusion (find_outliers) the PPT's plots and table use -- job["bench_
values"]/["ate_values"] are outlier-filtered by build_jobs itself, this
script never re-derives a number.
"""
from __future__ import annotations

import datetime
import statistics
import sys
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
_BASE_DIR = _THIS_DIR.parent
sys.path.insert(0, str(_BASE_DIR))
sys.path.insert(0, str(_THIS_DIR))

import xlsx_summary_writer as xw
import make_cal_distribution as cal
import make_bt_tx_trimtest_distribution as data

PATHS = data.PATHS
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


def build_rows(jobs):
    rows = []
    for job in jobs:
        bench_stats = _stats(job["bench_values"])
        ate_stats = _stats(job["ate_values"]) if job["ate_values"] else None
        bench_cells = _stats_cells(bench_stats) + [_r(job["usl"]), _r(job["lsl"])]
        ate_cells = _stats_cells(ate_stats) + [_r(job["usl"]), _r(job["lsl"])]
        deltas = [xw.delta(bench_cells[i], ate_cells[i]) for i in range(4)]
        rows.append([job["cleaned_name"], *bench_cells, *ate_cells, *deltas])
    return rows


def main() -> None:
    print("Loading Test Item names + USL/LSL/Unit spec...")
    specs = data.load_item_specs()
    print(f"  {len(specs)} items")

    print("Reading Bench_Cal CSVs...")
    bench_items = cal.read_bench(PATHS.bench_dir)
    print(f"  {len(bench_items)} test item(s) found in Bench log")

    print(f"Parsing ATE log ({PATHS.ate_log_path.name})...")
    ate_items = cal.read_ate(PATHS.ate_log_path)

    jobs = data.build_jobs(specs, bench_items, ate_items)
    print(f"Total rows: {len(jobs)}")

    rows = build_rows(jobs)
    out_path = PATHS.report_dir / f"{REPORT_DATE}_FWBT_TRIMTEST_Summary.xlsx"
    xw.write_summary_xlsx(rows, out_path, "FWBT TRIMTEST Summary")
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
