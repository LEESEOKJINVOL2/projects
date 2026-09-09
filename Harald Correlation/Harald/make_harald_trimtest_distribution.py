"""FWBT_TRIMTEST-style distribution PNGs for the Harald project (D7 chip
variant), per user request 2026-08-19: apply the same item-name "renamed"
cleanup the DP6P1P3 file got for Cal to FWBT_TRIMTEST_D7.xlsx, then build
the same 2G/5G-band-grouped report Cal/FWBT_TRIMTEST got, but against
Harald's own Bench/ATE data, not Cal's.

2026-08-20 update, per user request: this is now the ONLY Harald pipeline
going forward, and it reads the NEWER retest pull (Bench_40pcs/Harald_260820
+ ate_40pcs/Harald_260820/CAL_Only_Bench40pcs_Retest0_transposed.csv,
PATHS = get_paths("Harald_260820")) instead of the original pull
(Bench_40pcs/Harald + two per-vendor ATE files, PATHS = get_paths("Harald")).
That original pull's raw layout was nested by vendor, needing
make_harald_distribution.py's own read_bench_harald/read_ate_harald; this
retest pull's layout is FLAT (Bench_40pcs/Harald_260820/*.csv directly, ONE
ATE file) -- the exact same shape Cal/make_cal_distribution.py's own
read_bench/read_ate already expect, so this file now reads via THOSE
directly instead. Verified 2026-08-20: all 2812 D7 TESTNAME strings still
match byte-for-byte against this new Bench log too (0 missing) -- same
device/test program, just a newer measurement run.

Item universe + spec: FWBT_TRIMTEST_D7_renamed.xlsx's "Sheet1"
(TESTNAME/LIMITLO/LIMITHI/UNITS, 2812 rows). LIMITLO/LIMITHI map onto the
same "usl"/"lsl" slots the DP6P1P3 file's "Test Items" sheet USL/LSL columns
used (cross-checked identical values for items common to both files, e.g.
BT2G_A4P_BAND0_EXPAN: -94/-18 in both) -- this project's spec source, like
FWBT_TRIMTEST/Cal's, is the item-list sheet, not Bench's own embedded
LLIM/ULIM.

Reuses Cal/make_bt_tx_trimtest_distribution.py's generic pipeline pieces
(clean/build_raw_name_map/build_jobs/build_groups/group_value_range/
figure_size_for/draw_png/_draw_one, plus the geometry constants below)
UNMODIFIED via re-export -- none of those hardcode Cal's own bench_dir/
ate_log_path/xlsx, they operate on plain dicts, so this file only supplies
its own item/spec source (load_item_specs below) and its own Bench/ATE
reading (load_bench_ate below, via Cal's read_bench/read_ate).

Grouping (2G-left/5G-right by band) is the SAME measurement-across-bands
convention FWBT_TRIMTEST/Cal already uses -- unlike the PATNAME-section
grouping build_ppt_harald_distribution.py/build_ppt_harald_260820_
distribution.py use for Harald's OTHER (full ~4000-item) report, every
column in a group here is the same measurement on a different band, so
sharing one Y-axis range per group (group_value_range) is safe here, unlike
that report (where it was tried and reverted -- see make_harald_
distribution.py's module docstring for that finding).
"""
from __future__ import annotations

import argparse
import sys
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import openpyxl

_THIS_DIR = Path(__file__).resolve().parent
_BASE_DIR = _THIS_DIR.parent
sys.path.insert(0, str(_BASE_DIR))
sys.path.insert(0, str(_BASE_DIR / "Cal"))
sys.path.insert(0, str(_THIS_DIR))

import paths as harald_paths
import make_cal_distribution as cal  # read_bench/read_ate/safe_filename (reused, not modified)
import make_bt_tx_trimtest_distribution as trimtest  # generic pipeline pieces (reused, not modified)

PATHS = harald_paths.get_paths("Harald_260820")
XLSX_PATH = _BASE_DIR / "FWBT_TRIMTEST_D7_renamed.xlsx"
# 2026-08-21: ATE re-pull replaced the old "CAL_Only_Bench40pcs_Retest0_
# transposed.csv" (no longer present on disk) with this new file -- same
# header shape (SerialNumber,Upper Limit,Lower Limit,<DUT...>), same 40
# DUT serials, just a different underlying measurement condition
# ("withoutKGD, new CW table").
ATE_PATH = Path(r"C:\Users\seokjin.lee\Desktop\Harald_Correlation\rawdata\00.PBCAL\ATE")


def load_bench_ate():
    """Bench/ATE reading for this retest pull -- Cal's own read_bench/
    read_ate directly, since the raw layout is flat/single-file (see module
    docstring). Shared by this module's main() and build_ppt_harald_
    trimtest_revision1.py so both read the exact same way."""
    bench_items = cal.read_bench(PATHS.bench_dir)
    ate_items = cal.read_ate_auto(ATE_PATH)  # ATE_PATH may be one file or a folder of per-config files
    return bench_items, ate_items

# -- re-exported from Cal/make_bt_tx_trimtest_distribution.py so build_ppt_
# harald_trimtest_revision1.py can reference them the same way Cal/
# build_ppt_fwbt_trimtest_revision1.py references its own `data.X` -- these
# are dataset-independent (pure geometry/logic), not Cal-specific. --
SLIDE_WIDTH_IN = trimtest.SLIDE_WIDTH_IN
LEFT_MARGIN = trimtest.LEFT_MARGIN
RIGHT_MARGIN = trimtest.RIGHT_MARGIN
CELL_GAP_H = trimtest.CELL_GAP_H
LABEL_WIDTH = trimtest.LABEL_WIDTH
TABLE_TOP = trimtest.TABLE_TOP
HEADER_ROW_H = trimtest.HEADER_ROW_H
TABLE_ROW_H = trimtest.TABLE_ROW_H
TABLE_PLOT_GAP = trimtest.TABLE_PLOT_GAP
IMAGE_GAP = trimtest.IMAGE_GAP
CONTENT_AREA_TOP = trimtest.CONTENT_AREA_TOP
CONTENT_AREA_BOTTOM = trimtest.CONTENT_AREA_BOTTOM

figure_size_for = trimtest.figure_size_for
build_jobs = trimtest.build_jobs
build_groups = trimtest.build_groups
_draw_one = trimtest._draw_one


def load_item_specs():
    """-> {cleaned_name: {"usl":, "lsl":, "unit":}} from FWBT_TRIMTEST_D7_
    renamed.xlsx's Sheet1 (TESTNAME/LIMITLO/LIMITHI/UNITS).

    2026-08-20 fix: LIMITLO -> lsl, LIMITHI -> usl (was swapped: LIMITLO was
    being read into "usl" and LIMITHI into "lsl"). LIMITLO/LIMITHI are
    unambiguous by name ("Limit Low"/"Limit High") -- verified LIMITLO <=
    LIMITHI for all 2443 items carrying both (0 violations), so LIMITLO is
    genuinely always the lower bound (LSL) and LIMITHI the upper bound
    (USL). The original mapping came from matching D7's column values
    positionally against FWBT_TRIMTEST_DP6P1P3_renamed.xlsx's own "USL"/
    "LSL"-labeled columns (same file position, identical values for shared
    items, e.g. BT2G_A4P_BAND0_EXPAN) -- that file's own USL/LSL header
    labels are apparently swapped from the actual low/high value ordering,
    which this fix does not touch (DP6P1P3/Cal's own pipeline is untouched;
    only this D7/Harald one reads LIMITLO/LIMITHI, where the low/high
    semantic is explicit in the column name itself, not just a label)."""
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb["Sheet1"]
    specs = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        name, limitlo, limithi, unit = row[0], row[1], row[2], row[3]
        if not name:
            continue
        specs[name] = {"usl": limithi, "lsl": limitlo, "unit": unit or ""}
    return specs


def bench_png_dir() -> Path:
    return PATHS.result_png_dir / "harald_trimtest_distribution"


def ate_png_dir() -> Path:
    return PATHS.result_png_dir.parent / "Harald_TRIMTEST_ATE" / "harald_trimtest_distribution"


def parse_args():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--limit", type=int, default=0, help="only draw the first N groups")
    p.add_argument("--count-only", action="store_true", help="print the match audit only")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    print("Loading Test Item names + LIMITLO/LIMITHI/UNITS spec...")
    specs = load_item_specs()
    print(f"  {len(specs)} items")

    print("Reading Bench_Harald_260820 CSVs...")
    bench_items, ate_items = load_bench_ate()
    print(f"  {len(bench_items)} test item(s) found in Bench log")

    jobs = build_jobs(specs, bench_items, ate_items)
    groups = build_groups(jobs)
    print(f"{len(groups)} group(s) (slides)")

    if args.limit > 0:
        keys = list(groups)[: args.limit]
        groups = {k: groups[k] for k in keys}
        jobs = [j for g in groups.values() for j in g["jobs"]]
        print(f"Limiting to first {len(groups)} group(s), {len(jobs)} job(s)")

    for g in groups.values():
        n_cols = len(g["jobs"])
        for job in g["jobs"]:
            job["n_cols"] = n_cols

    if args.count_only:
        return

    bench_dir_out = bench_png_dir()
    ate_dir_out = ate_png_dir()
    bench_dir_out.mkdir(parents=True, exist_ok=True)
    ate_dir_out.mkdir(parents=True, exist_ok=True)

    for job in jobs:
        fname = cal.safe_filename(job["cleaned_name"])
        job["bench_png"] = bench_dir_out / f"{fname}.png"
        job["ate_png"] = ate_dir_out / f"{fname}.png"

    n = 0
    with ProcessPoolExecutor(max_workers=PATHS.png_workers) as pool:
        for name, has_ate in pool.map(_draw_one, jobs, chunksize=20):
            n += 1
            if n % 200 == 0 or n == len(jobs):
                print(f"  [{n}/{len(jobs)}] {name}  (ATE: {'yes' if has_ate else 'no'})")

    print(f"\nDone. Bench PNG -> {bench_dir_out}")
    print(f"      ATE   PNG -> {ate_dir_out}")


if __name__ == "__main__":
    main()
