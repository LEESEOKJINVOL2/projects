"""XLSX summary with Bench-vs-ATE delta columns for Cal. No existing
summary script for this project either -- written fresh, reusing Cal/
make_cal_distribution.py's own read_bench/read_ate/build_jobs/effective_spec
(imported, not modified).

Per user request 2026-08-15: Mean/Std/Min/Max (Bench + ATE) plus 4 delta
columns (|Bench-ATE|, abs()'d via the shared xlsx_summary_writer.delta,
blank where either side is missing). Same >=1-Bench-limit qualifying rule
make_cal_distribution.build_jobs() already applies (3847/6830 items, per
that module's own docstring) -- this script doesn't re-filter, it just
tabulates the same jobs into a worksheet instead of a slide grid.
"""
from __future__ import annotations

import datetime
import statistics
import sys
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(_THIS_DIR))
sys.path.insert(0, str(_THIS_DIR.parent))

import xlsx_summary_writer as xw
import make_cal_distribution as cal

PATHS = cal.PATHS
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
    print("Reading Bench_Cal CSVs...")
    bench_items = cal.read_bench(PATHS.bench_dir)
    print(f"  {len(bench_items)} test item(s) found")
    print(f"Parsing ATE log ({PATHS.ate_log_path.name})...")
    ate_items = cal.read_ate(PATHS.ate_log_path)
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
        rows.append([job["name"], *bench_cells, *ate_cells, *deltas])
    return rows


def main() -> None:
    rows = build_rows()
    print(f"Total rows: {len(rows)}")
    out_path = PATHS.report_dir / f"{REPORT_DATE}_Cal_Summary.xlsx"
    xw.write_summary_xlsx(rows, out_path, "Cal Summary")
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
