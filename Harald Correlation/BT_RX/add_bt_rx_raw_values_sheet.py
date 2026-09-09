"""Add a "Raw Values" sheet to the existing BT_RX xlsx summary workbook
(report/BT_RX/{date}_BT_RX_Summary.xlsx, built by build_bt_rx_xlsx_summary.py),
per user request 2026-08-25 ("TX에서 한거 동일하게 RX에도 적용해줘" -- apply what
was just done for BT_TX identically to BT_RX). Mirrors BT_TX's own
add_bt_tx_raw_values_sheet.py closely: one row per Test Item, same
name/order as the existing "BT_RX Summary" sheet, columns split into a
Bench block and an ATE block, one column per DUT in each, headed by the
real DUT serial number -- NOT the anonymous "#NN" index bt_rx_ate_lookup's
read_bench()/dut_id_of() tags a value with internally.

This ADDS a sheet only -- "BT_RX Summary" (built by build_bt_rx_xlsx_
summary.py) is left exactly as already generated.

Simpler than BT_TX's equivalent for two reasons:

  * BT_RX has only 3 metric "families" (RSSI / Sensitivity_PER /
    Sensitivity_BER), each with its own make_bt_rx_*_revision1.py owning a
    single build_jobs() -- no ACP-style "pool multiple readings per DUT
    into one Test Item" case exists here, so no _dut_means()-style
    averaging is needed; every combo contributes exactly one Bench sample
    and one ATE record per DUT.
  * bt_rx_ate_lookup.read_bench() already tags every Bench row with its own
    dut_id per (packet, sub_band, channel) key, and nearest_bench_rows()
    already returns a dut_id-keyed dict -- unlike several BT_TX bench
    modules, no NEW read_all_with_dut()-style sibling was needed on the
    Bench side; only the ATE side needed a new per-dut parser (see
    bt_rx_ate_lookup.parse_ate_items_per_dut(), added alongside the
    existing parse_ate_items(), never replacing it).

The real (post-2026-08-24) Bench pull nests each DUT's file under its own
"#NN_<serial>/" folder, e.g. "..._#01_725HW3001RE0001A6Q_cns_temp_data.csv"
-- same convention as BT_TX -- so a "#NN" -> serial map can be built once
by scanning Bench's own folder (build_num_to_serial() below, identical
regex to BT_TX's own), and used to translate read_bench()'s/
nearest_bench_rows()'s own "#NN" dut_id into the same real serial the ATE
side's header row already uses.

The OLDER flat Bench layout (files directly in bench_dir, no "#NN_<serial>/"
folder) has no serial in its filenames at all -- there is genuinely no way
to correlate Bench's anonymous index to ATE's real serial in that layout.
build_num_to_serial() then returns {}, and main() prints a message and
skips writing the sheet entirely, rather than guessing or positionally
aligning (which would silently produce wrong data) -- same rule as BT_TX.
"""
from __future__ import annotations

import csv
import datetime
import re
import sys
from pathlib import Path as _Path

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_rx_paths

import bt_rx_ate_lookup as rx
import make_bt_rx_rssi_revision1 as rssi_mod
import make_bt_rx_sensitivity_per_revision1 as per_mod
import make_bt_rx_sensitivity_ber_revision1 as ber_mod

PATHS = bt_rx_paths.get_paths("BT_RX")
REPORT_DATE = datetime.date.today().isoformat()
SUMMARY_PATH = PATHS.report_dir / f"{REPORT_DATE}_BT_RX_Summary.xlsx"

# (label, module, ate_metric, bench_tuple_index) -- bench_tuple_index is the
# position of this metric's value inside both read_bench()'s own per-row
# tuple (dut, pwr, rssi, per, ber) AND nearest_bench_rows()'s result tuple
# (delta, pwr, rssi, per, ber) -- rssi/per/ber all sit at the SAME index in
# both tuples (2, 3, 4 respectively; "dut" and "delta" both occupy index 0),
# so one index works for either.
SECTIONS = [
    ("RSSI", rssi_mod, "AVGRSSI", 2),
    ("Sensitivity_PER", per_mod, "PER", 3),
    ("Sensitivity_BER", ber_mod, "BER", 4),
]

# Matches the real (post-2026-08-24) Bench filename convention: a "#NN"
# dut-index token immediately followed by "_<real DUT serial>_cns_temp_
# data.csv" -- identical pattern to BT_TX's own SERIAL_RE. The OLDER flat
# layout's filenames (e.g. "..._#4_cns_temp_data.csv") have nothing between
# the "#NN" token and "_cns_temp_data.csv", so this simply never matches
# them -- build_num_to_serial() naturally returns {} for that layout.
SERIAL_RE = re.compile(r"#(\d+)_([^_/\\]+)_cns_temp_data\.csv$", re.IGNORECASE)
# 2026-08-25 real bug found in a live 40-DUT RX pull: DUT #30's own FOLDER
# is named "#30_725HW3001KT0001A6Q" (agrees with 40pcs_config.txt and the
# ATE log's own header), but the .csv FILE inside it is misnamed
# "..._#30_725HW30010C0001A6Q_cns_temp_data.csv" -- a different serial,
# apparently a copy/paste slip when that one test record was saved. Reading
# the serial from the filename alone (the original SERIAL_RE above) picks
# up that wrong value, which then has no entry in 40pcs_config.txt and
# trips this whole workbook's fallback-to-flat/no-config-sheets path for
# every DUT, not just #30. The per-DUT FOLDER name is set once when that
# DUT's test slot is created and is what 40pcs_config.txt/ATE both already
# agree with, so it's the more trustworthy source -- prefer it.
FOLDER_SERIAL_RE = re.compile(r"^#(\d+)_(.+)$", re.IGNORECASE)


def build_num_to_serial(bench_dir) -> dict:
    """{"#NN": dut_serial} -- scans bench_dir (any depth, via rglob, same as
    bt_rx_ate_lookup.read_bench()'s own sweep) for the real-serial filename
    pattern, skipping dotfiles (e.g. a stray ".DS_Store") the same way
    read_bench() does. The "#NN" key format matches dut_id_of()'s own output
    (f"#{digits}") exactly, since both capture the same digit run right
    after the literal "#"."""
    mapping = {}
    for p in bench_dir.rglob("*_cns_temp_data.csv"):
        if p.name.startswith("."):
            continue
        # Prefer the per-DUT FOLDER's own serial (see FOLDER_SERIAL_RE's
        # docstring above) -- falls back to the filename-embedded one only
        # when there's no per-DUT folder wrapper at all (old flat layout,
        # where the file's parent IS bench_dir itself, not a "#NN_serial/"
        # subfolder).
        fm = FOLDER_SERIAL_RE.match(p.parent.name)
        if fm:
            mapping[f"#{fm.group(1)}"] = fm.group(2)
            continue
        m = SERIAL_RE.search(p.name)
        if m:
            mapping[f"#{m.group(1)}"] = m.group(2)
    return mapping


def collect_ate_serials() -> set:
    """The ATE log's own header[3:] (real DUT serials, fixed order) -- {}
    if extract_mode=="bench" (ATE genuinely not used this run) or the log
    is missing, same degrade-gracefully rule as BT_TX's own collect_ate_
    serials()."""
    if PATHS.extract_mode == "bench" or not PATHS.ate_log_path:
        return set()
    try:
        with PATHS.ate_log_path.open(encoding="utf-8-sig", newline="") as fh:
            reader = csv.reader(fh)
            header = next(reader)
            return set(header[3:])
    except (FileNotFoundError, OSError):
        return set()


def bench_by_serial_for_job(bench_all, packet, band, ch, target_pwr, value_idx, num_to_serial) -> dict:
    """One job's Bench side -> {serial: value}, re-running nearest_bench_
    rows() (cheap: bench_all is read_bench()'s lru_cache'd result) rather
    than reusing the job's own already-flattened "bench_values" list, since
    that list has already discarded per-DUT identity by the time build_jobs()
    returns it (see this project's CLAUDE.md-documented pattern)."""
    bench_rows = bench_all.get((packet, band, ch), [])
    matched = rx.nearest_bench_rows(bench_rows, target_pwr)
    out = {}
    for dut, rec in matched.items():
        v = rec[value_idx]
        if v is None:
            continue
        serial = num_to_serial.get(dut)
        if serial:
            out[serial] = v
    return out


def ate_by_serial_for_item(ate_idx_per_dut, key, item_name) -> dict:
    """One job's ATE side -> {serial: value}, found by exact item name
    (not by re-matching power) inside parse_ate_items_per_dut()'s own
    per-(packet,band,ch) record list -- every job's own "item" name already
    identifies exactly which ATE row it came from."""
    for rec in ate_idx_per_dut.get(key, []):
        if rec["item"] == item_name:
            return rec["by_serial"]
    return {}


def make_row(name, bench_map, ate_map, all_serials) -> list:
    row = [name]
    for s in all_serials:
        row.append(bench_map.get(s))
    for s in all_serials:
        row.append(ate_map.get(s))
    return row


def collect_raw_rows(num_to_serial) -> list:
    """(name, bench_by_serial, ate_by_serial) for every Test Item row across
    all 3 BT_RX metric families -- same name/order as "BT_RX Summary". Kept
    as ONE function (rather than duplicated in update_bt_rx_summary_by_
    vendor.py) so both the flat Raw Values sheet (this script's own main(),
    via flatten_rows()) and the per-config Summary/Raw-Values sheets share
    the exact same underlying per-DUT data."""
    bench_all = rx.read_bench(PATHS.bench_dir)
    rows = []
    for label, mod, metric, value_idx in SECTIONS:
        print(f"Raw Values: {label}...")
        jobs, _audit = mod.build_jobs()
        ate_per_dut = {}
        if PATHS.extract_mode != "bench" and PATHS.ate_log_path:
            ate_per_dut = rx.parse_ate_items_per_dut(PATHS.ate_log_path, metric)
        for job in jobs:
            name = f"BT_RX_{label}_{rx.combo_name(job['packet'], job['band'], job['ch'], job['pwr'])}"
            bench_map = bench_by_serial_for_job(
                bench_all, job["packet"], job["band"], job["ch"], job["pwr"], value_idx, num_to_serial)
            ate_map = {}
            if ate_per_dut:
                key = (job["packet"], job["band"], job["ch"])
                ate_map = ate_by_serial_for_item(ate_per_dut, key, job["item"])
            rows.append((name, bench_map, ate_map))
        print(f"  {len(jobs)} row(s)")
    return rows


def flatten_rows(rows, all_serials) -> list:
    """[(name, bench_map, ate_map), ...] -> flat make_row() lists, for the
    fixed Bench-block-then-ATE-block "Raw Values" layout this script's own
    main() writes."""
    return [make_row(name, bench_map, ate_map, all_serials) for name, bench_map, ate_map in rows]


def write_flat_raw_values_sheet(wb, rows, all_serials) -> None:
    """Creates/replaces the 2-level-header ("Bench"/"ATE" -> serial) "Raw
    Values" sheet in an already-open workbook `wb` -- does NOT save.
    Identical layout to BT_TX's own write_flat_raw_values_sheet(). Hoisted
    out of main() so update_bt_rx_summary_by_vendor.py's graceful-
    degradation path (missing/incomplete 40pcs_config.txt) can build the
    exact same fallback sheet this script's own main() writes."""
    n_cols = 1 + 2 * len(all_serials)

    if "Raw Values" in wb.sheetnames:
        del wb["Raw Values"]
    ws = wb.create_sheet("Raw Values")

    ws.cell(row=1, column=1, value="Test Item")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    col = 2
    for group_label in ("Bench", "ATE"):
        start = col
        for serial in all_serials:
            ws.cell(row=2, column=col, value=serial)
            col += 1
        ws.cell(row=1, column=start, value=group_label)
        ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=col - 1)

    header_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=1, max_row=2, max_col=n_cols):
        for cell in row:
            cell.font = header_font
            cell.alignment = center
    for i in range(1, n_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 12

    for r, row in enumerate(rows, start=3):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v)

    ws.freeze_panes = "B3"


def main() -> None:
    print("Building #NN -> DUT serial map from the Bench folder...")
    num_to_serial = build_num_to_serial(PATHS.bench_dir)
    if not num_to_serial:
        print("Bench folder has no embedded DUT serials (old flat layout) -- "
              "skipping Raw Values sheet.")
        return
    print(f"  {len(num_to_serial)} DUT(s) mapped")

    ate_serials = collect_ate_serials()
    all_serials = sorted(set(num_to_serial.values()) | ate_serials)
    print(f"  {len(all_serials)} DUT serial(s) total")

    dict_rows = collect_raw_rows(num_to_serial)
    print(f"Total rows: {len(dict_rows)}")
    rows = flatten_rows(dict_rows, all_serials)

    print("Writing sheet into existing workbook...")
    wb = openpyxl.load_workbook(SUMMARY_PATH)
    write_flat_raw_values_sheet(wb, rows, all_serials)
    wb.save(SUMMARY_PATH)
    n_cols = 1 + 2 * len(all_serials)
    print(f"\nSaved: {SUMMARY_PATH}  (added sheet 'Raw Values', {len(rows)} row(s) x {n_cols} col(s))")


if __name__ == "__main__":
    main()
