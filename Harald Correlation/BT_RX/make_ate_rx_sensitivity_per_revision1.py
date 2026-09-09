"""BT_RX Sensitivity (PER) -- ATE counterpart to
make_bt_rx_sensitivity_per_revision1.py. Must run AFTER that script
(imports its build_jobs()/title_for()/audit helpers directly).

CLI: --limit N (first N matched combos only), --count-only (audit only,
same output as the Bench script's).
"""
from __future__ import annotations

import argparse

import bt_rx_ate_lookup as rx
import make_bt_rx_sensitivity_per_revision1 as bench_mod

PATHS = bench_mod.PATHS


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--limit", type=int, default=0)
    p.add_argument("--count-only", action="store_true")
    args = p.parse_args()

    print("Reading Bench_RX CSVs...")
    jobs, audit = bench_mod.build_jobs()
    bench_mod.print_audit(audit)
    if args.count_only:
        return

    # From the FULL job list, before any --limit slicing -- see
    # n_cols_by_packet's docstring. Same n_cols/has_spec the Bench script
    # used, so this ATE PNG lands in the identical-shaped cell as its
    # Bench sibling on the same slide.
    n_cols_by_packet = rx.n_cols_by_packet(jobs)
    has_spec = rx.jobs_have_spec(jobs)

    if args.limit > 0:
        jobs = jobs[: args.limit]
        print(f"Limiting to first {len(jobs)} combo(s)")

    out_dir = bench_mod.ate_png_dir()
    out_dir.mkdir(parents=True, exist_ok=True)

    n = 0
    for job in jobs:
        fname = rx.safe_filename(rx.combo_name(job["packet"], job["band"], job["ch"], job["pwr"]))
        figsize = rx.figure_size_for(n_cols_by_packet[job["packet"]], has_spec)
        rx.draw_png(job["ate_values"], rx.condition_title(job, "ATE"), bench_mod.XLABEL, out_dir / f"{fname}.png",
                    value_range=job["value_range"], spec=job["spec"], figsize=figsize)
        n += 1
        if n % 20 == 0 or n == len(jobs):
            print(f"  [{n}/{len(jobs)}] {job['item']}  (ATE DUTs: {len(job['ate_values'])})")

    print(f"\nDone. ATE PNG -> {out_dir}")


if __name__ == "__main__":
    main()
