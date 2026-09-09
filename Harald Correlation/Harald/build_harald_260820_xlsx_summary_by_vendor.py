"""Per-group xlsx summaries for Harald_260820 -- one sheet per top-level
Bench subfolder (originally Skyworks/Sony; 2026-08-25: also used for a
100-DUT pull's DNNJ/DSKJ/SNKJ/SSNJ, 4 test-config folders). Same Test Item
set/columns as build_harald_260820_xlsx_summary.py's own pooled summary
(Mean/Std/Min/Max/USL/LSL for Bench+ATE, plus Delta), but Mean/Std/Min/Max
are computed from ONLY that folder's DUTs instead of every DUT pooled.

Per user request 2026-08-20: the group split originally came from a user-
provided DUT list (Downloads/"260818_2did 40pcs.txt", SKYWORKS section
then a SONY section, 20 serials each). 2026-08-21: switched to deriving
the exact same split straight from Bench_40pcs/Harald's own vendor/DUT
subfolder structure instead -- verified byte-for-byte identical to that
txt file's two lists -- since that folder is a real, always-present part
of this workspace (unlike a personal Downloads file), and Harald's Bench
folder can now be pointed at either layout (flat Harald_260820 or this
vendor-nested one) via the GUI's own Folders card, so this one already has
to exist for that to work at all. 2026-08-25: generalized from "exactly
2 vendors" to "however many top-level subfolders Bench actually has" --
see vendor_lists_from_nested_bench's own docstring.

DUT-serial ("2did") labeling for Bench values reuses add_harald_260820_
raw_values_sheet.py's hex-id -> serial mapping (imported, not modified) --
same reasoning as that file: the flat Bench_40pcs/Harald_260820/*.csv
filenames only carry an internal hex id, not the human DUT serial. ATE
per-DUT parsing also reuses that file's read_ate_per_dut (ATE's own CSV
header already has real serials, no mapping needed there).
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
import add_harald_260820_raw_values_sheet as raw_sheet  # build_hex_to_serial/read_ate_per_dut (reused, not modified)

PATHS = harald.PATHS
REPORT_DATE = datetime.date.today().isoformat()
# 2026-08-21: was a personal-Downloads-folder txt file path -- now the real
# vendor-nested Bench folder (see module docstring). Kept the same name so
# update_harald_260820_summary_by_vendor.py's call site needs no changes.
VENDOR_LIST_PATH = raw_sheet.NESTED_BENCH_DIR


def parse_vendor_lists(path: Path) -> dict:
    """-> {GROUP_LABEL: [serial, ...]}, one entry per top-level subfolder
    of `path`, named after that subfolder verbatim -- see
    vendor_lists_from_nested_bench's own docstring (2026-08-25: reverted
    to this straightforward per-folder split, per reviewer request, rather
    than collapsing multiple config folders into a smaller number of
    "real vendors"). Kept the name `parse_vendor_lists` (and its one-
    argument shape) from the txt-file version this originally replaced, so
    nothing calling it needs to change."""
    return raw_sheet.vendor_lists_from_nested_bench(path)


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


def build_rows(vendor_serials: set, hex_to_serial: dict, bench_items: dict, ate_by_item: dict, jobs: list):
    rows = []
    for job in jobs:
        raw_name = job["name"]
        item = harald.clean(raw_name)

        bench_vals = [v for v, hexid in bench_items[raw_name]["values"]
                      if hex_to_serial.get(hexid) in vendor_serials]
        ate_vals = [v for serial, v in ate_by_item.get(raw_name, {}).items()
                    if serial in vendor_serials]

        bench_stats = _stats(bench_vals)
        ate_stats = _stats(ate_vals) if ate_vals else None
        ate_llim, ate_ulim = cal.effective_spec(job)
        bench_cells = _stats_cells(bench_stats) + [_r(job["bench_ulim"]), _r(job["bench_llim"])]
        ate_cells = _stats_cells(ate_stats) + [_r(ate_ulim), _r(ate_llim)]
        deltas = [xw.delta(bench_cells[i], ate_cells[i]) for i in range(4)]
        rows.append([item, *bench_cells, *ate_cells, *deltas])
    return rows


def main() -> None:
    print(f"Parsing vendor DUT list ({VENDOR_LIST_PATH.name})...")
    vendor_lists = parse_vendor_lists(VENDOR_LIST_PATH)
    for vendor, serials in vendor_lists.items():
        print(f"  {vendor}: {len(serials)} DUT(s)")

    print("Building hex-id -> DUT-serial map from the reference Bench copy...")
    hex_to_serial = raw_sheet.build_hex_to_serial(raw_sheet.NESTED_BENCH_DIR)
    print(f"  {len(hex_to_serial)} DUT(s) mapped")

    print("Reading Bench_Harald_260820 CSVs...")
    bench_items = cal.read_bench(PATHS.bench_dir)
    print(f"  {len(bench_items)} test item(s) found")

    print(f"Parsing ATE log ({harald.ATE_PATH.name}) per-DUT...")
    ate_by_item = raw_sheet.read_ate_per_dut(harald.ATE_PATH)

    ate_items_std = cal.read_ate(harald.ATE_PATH)
    jobs, n_no_ate = cal.build_jobs(bench_items, ate_items_std)
    print(f"  {len(jobs)} qualifying item(s) (>=1 Bench limit), {len(jobs) - n_no_ate} with ATE match (all-vendor)")

    for vendor, serials in vendor_lists.items():
        vendor_set = set(serials)
        print(f"\nBuilding {vendor} rows...")
        rows = build_rows(vendor_set, hex_to_serial, bench_items, ate_by_item, jobs)
        out_path = PATHS.report_dir / f"{REPORT_DATE}_Harald_260820_Summary_{vendor}.xlsx"
        xw.write_summary_xlsx(rows, out_path, f"Harald_260820 {vendor}")
        print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
