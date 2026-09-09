"""BT_RX Sensitivity (BER) Bench distribution histograms (bench `ber`, ATE
metric `BER`). ATE counterpart: make_ate_rx_sensitivity_ber_revision1.py
(must run after this one). Customer-scoped to **1DH5 packets only** (see
the Test Item Matching table's own "Sensitivity (cfg-carrier_01-packet_type
가 1DH5인 경우)" row) -- any BER ATE item for a different packet_type, if
one ever shows up in a future ATE pull, is intentionally skipped here, not
drawn.

Kept as a separate script from its ATE counterpart so this project's
extract_mode="ate"/"bench" split works the same way as every other
project's Bench/ATE script pair -- see md/CLAUDE.md's extract_mode section.

See bt_rx_ate_lookup.py's module docstring for the matching-rule writeup
(band/channel mapping, power-token decoding, nearest-match tolerance).

CLI: --limit N (first N matched combos only), --count-only (audit match
rate + per-combo power deltas without drawing anything).
"""
from __future__ import annotations

import argparse
from pathlib import Path

import bt_rx_ate_lookup as rx

PATHS = rx.PATHS
ATE_METRIC = "BER"
ITEM_LABEL = "Sensitivity_BER"
PACKET_FILTER = "1DH5"
XLABEL = "BER"
SUBDIR = "sensitivity_ber_revision1"


def bench_png_dir() -> Path:
    return PATHS.result_png_dir / SUBDIR


def ate_png_dir() -> Path:
    return PATHS.result_png_dir.parent / "BT_RX_ATE" / SUBDIR


def title_for(job) -> str:
    """See bt_rx_ate_lookup.condition_title's docstring for the full
    rationale (short + no metric prefix + side-specific power)."""
    return rx.condition_title(job, "Bench")


def build_jobs():
    bench = rx.read_bench(PATHS.bench_dir)
    ate_idx = rx.parse_ate_items(PATHS.ate_log_path, ATE_METRIC)

    jobs = []
    audit = []
    for (packet, band, ch), items in sorted(ate_idx.items()):
        if packet != PACKET_FILTER:
            continue
        bench_rows = bench.get((packet, band, ch), [])
        for rec in items:
            matched = rx.nearest_bench_rows(bench_rows, rec["pwr"])
            bench_values = [v[4] for v in matched.values() if v[4] is not None]  # ber
            audit.append((packet, band, ch, rec["pwr"], rec["item"], len(matched),
                          max((v[0] for v in matched.values()), default=None)))
            if not bench_values:
                continue  # excluded entirely -- user decision 2026-08-14, see bt_rx_ate_lookup.py
            jobs.append({
                "packet": packet, "band": band, "ch": ch, "pwr": rec["pwr"],
                "bench_pwr": rx.matched_bench_pwr(matched),
                "item": rec["item"], "bench_values": bench_values, "ate_values": rec["values"],
                "spec": rec["spec"],
                "value_range": rx.compute_value_range(bench_values, rec["values"], rec["spec"]),
            })
    rx.finalize_shared_ranges(jobs)
    return jobs, audit


def print_audit(audit):
    n = len(audit)
    n_bench_matched = sum(1 for a in audit if a[5] > 0)
    print(f"\n=== {ITEM_LABEL} ({PACKET_FILTER} only): ATE items found: {n}  with >=1 Bench DUT match: {n_bench_matched} ===")
    deltas = [a[6] for a in audit if a[6] is not None]
    if deltas:
        print(f"    matched-DUT power delta: max={max(deltas):.3f}dB avg={sum(deltas)/len(deltas):.3f}dB "
              f"(tolerance={rx.POWER_TOL_DBM}dB)")
    unmatched = [a for a in audit if a[5] == 0]
    if unmatched:
        print(f"    {len(unmatched)} item(s) with NO Bench DUT within tolerance (shown, first 10):")
        for packet, band, ch, pwr, item, _n, _d in unmatched[:10]:
            print(f"      {item}  (target {pwr:g}dBm)")


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--limit", type=int, default=0)
    p.add_argument("--count-only", action="store_true")
    args = p.parse_args()

    print("Reading Bench_RX CSVs...")
    jobs, audit = build_jobs()
    print_audit(audit)
    if args.count_only:
        return

    # From the FULL job list, before any --limit slicing -- see
    # n_cols_by_packet's docstring.
    n_cols_by_packet = rx.n_cols_by_packet(jobs)
    has_spec = rx.jobs_have_spec(jobs)

    if args.limit > 0:
        jobs = jobs[: args.limit]
        print(f"Limiting to first {len(jobs)} combo(s)")

    out_dir = bench_png_dir()
    out_dir.mkdir(parents=True, exist_ok=True)

    n = 0
    for job in jobs:
        fname = rx.safe_filename(rx.combo_name(job["packet"], job["band"], job["ch"], job["pwr"]))
        figsize = rx.figure_size_for(n_cols_by_packet[job["packet"]], has_spec)
        rx.draw_png(job["bench_values"], title_for(job), XLABEL, out_dir / f"{fname}.png",
                    value_range=job["value_range"], spec=job["spec"], figsize=figsize)
        n += 1
        if n % 20 == 0 or n == len(jobs):
            print(f"  [{n}/{len(jobs)}] {job['item']}  (Bench DUTs matched: {len(job['bench_values'])})")

    print(f"\nDone. Bench PNG -> {out_dir}")


if __name__ == "__main__":
    main()
