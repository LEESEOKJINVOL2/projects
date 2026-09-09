"""Add a "Raw Values" sheet to the existing BT_TX xlsx summary workbook
(report/BT_TX/{date}_BT_TX_Summary.xlsx, built by build_bt_tx_xlsx_summary.py),
per user request 2026-08-25 -- mirrors Harald's own Raw Values feature
(Harald/add_harald_260820_raw_values_sheet.py): one row per Test Item, same
name/order as the existing "BT_TX Summary" sheet, columns split into a
Bench block and an ATE block, one column per DUT in each, headed by the
real DUT serial number -- NOT the anonymous "#NN" index every BT_TX bench
module's own read_all() tags a value with internally.

This ADDS a sheet only -- "BT_TX Summary" (built by build_bt_tx_xlsx_
summary.py) is left exactly as already generated. Only ONE new sheet is
added -- no per-group/per-vendor variants (that's a different, Harald-only
feature, out of scope here).

Why this needs more than a straight port of Harald's script: BT_TX's 12
metric-family "rows" (Power, DEVM RMS/Peak/99pct, ACP, FreqAcc's 9 sub-
metrics -- see build_bt_tx_summary_csv.py) each come from their own Bench
module and their own ATE module, and:

  * Bench side: every module's read_all() (freqacc's read_all(bt_tx_dir,
    metric_cols); ACP's NEW read_all_with_dut(bt_tx_dir), added alongside
    its existing dut-less read_all()) already tags each value with a
    "#NN"-style dut_id (make_bt_tx_pwr_revision2.DUT_RE et al) -- an
    anonymous per-file index, not the real serial.
  * ATE side: every module's own parse function discards per-DUT identity
    (positional row[3:] via _values_from_row()); this script uses each
    module's NEW parse_..._per_dut() sibling instead (added alongside the
    existing function, never replacing it), which captures header[3:] (the
    ATE CSV's own real DUT-serial row) once and keeps {serial: value} per
    record.
  * The real (post-2026-08-24) Bench pull nests each DUT's file under its
    own "#NN_<serial>/" folder, e.g. "..._#01_725HW3001RE0001A6Q_cns_temp_
    data.csv" -- so a "#NN" -> serial map can be built once by scanning
    Bench's own folder (see build_num_to_serial() below), and used to
    translate every bench module's own "#NN" dut_id into the same real
    serial the ATE side already uses -- that's how a Bench column and an
    ATE column end up meaning the same physical DUT.
  * The OLDER flat Bench layout (files directly in bt_tx_dir, no "#NN_
    <serial>/" folder) has no serial in its filenames at all -- there is
    genuinely no way to correlate Bench's anonymous index to ATE's real
    serial in that layout. build_num_to_serial() then returns {}, and
    main() prints a message and skips writing the sheet entirely, rather
    than guessing or positionally aligning (which would silently produce
    wrong data).

ACP is the one family whose Test Item pools MULTIPLE Bench frequencies
together (a whole band's worth of channels, per offset) -- see
build_ppt_acp_revision2.get_offset_stats_by_combo(). A single DUT can
therefore contribute more than one raw reading to one ACP row (one per
frequency in that band); every other family's Test Item is one exact
(packet, supply, freq) combo, so a DUT contributes exactly one reading.
_dut_means()/bench_by_serial()/ate_by_serial_from_records() below always
average a DUT's own contribution(s) down to one cell -- for the ordinary
one-value case this is just that value, and for ACP it's that DUT's mean
across the pooled frequencies, so every family fits the same fixed one-
column-per-DUT layout.

Row set/order/names are guaranteed identical to "BT_TX Summary" because
this script reuses build_bt_tx_summary_csv.py's own row-building helpers
(PE.get_ranges_and_stats, ACP_BUILD.discover/get_offset_stats_by_combo,
FREQACC_BUILD.get_ranges_and_stats_by_metric, S.DEVM_SECTIONS/FA_NAME_SAFE)
to get each family's exact combo set + name string, then separately fetches
that SAME combo's raw per-DUT Bench list and raw per-DUT ATE match --
never re-deriving a combo or a name independently.
"""
from __future__ import annotations

import csv
import datetime
import re
import statistics
import sys
from collections import defaultdict
from pathlib import Path as _Path

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

import build_bt_tx_summary_csv as S
import build_ppt_pwr_evm_revision2 as PE
import build_ppt_acp_revision2 as ACP_BUILD
import build_ppt_freqacc_revision2 as FREQACC_BUILD
import make_bt_tx_acp_revision2 as acp_mod
import make_ate_tx_acp_revision2 as ate_acp_mod
import make_bt_tx_bw_revision1 as bw_mod
import make_ate_tx_bw_revision1 as ate_bw_mod

PATHS = bt_tx_paths.get_paths("BT_TX")
REPORT_DATE = datetime.date.today().isoformat()
SUMMARY_PATH = PATHS.report_dir / f"{REPORT_DATE}_BT_TX_Summary.xlsx"

# Matches the real (post-2026-08-24) Bench filename convention: a "#NN"
# dut-index token immediately followed by "_<real DUT serial>_cns_temp_
# data.csv" -- e.g. "..._#01_725HW3001RE0001A6Q_cns_temp_data.csv". The
# OLDER flat layout's filenames (e.g. "..._#4_cns_temp_data.csv") have
# nothing between the "#NN" token and "_cns_temp_data.csv", so this simply
# never matches them -- build_num_to_serial() naturally returns {} for that
# layout instead of needing a separate "is this the new layout" check.
SERIAL_RE = re.compile(r"#(\d+)_([^_/\\]+)_cns_temp_data\.csv$", re.IGNORECASE)
# 2026-08-25: a live 40-DUT BT_RX pull hit a real case of the per-DUT
# FOLDER name and the .csv FILENAME inside it disagreeing on the serial (a
# copy/paste slip in one test record's filename) -- see BT_RX's own
# add_bt_rx_raw_values_sheet.py for the full story. Reading the serial from
# the filename alone (the original SERIAL_RE above) would pick up the
# wrong value there, which then has no entry in 40pcs_config.txt and trips
# the whole workbook's fallback-to-flat path for every DUT. Hasn't
# (yet) happened in a BT_TX pull, but it's the identical filename
# convention/risk, so applying the same fix here preemptively: the
# per-DUT FOLDER name is set once when that DUT's test slot is created and
# is what 40pcs_config.txt/ATE both already agree with, so prefer it.
FOLDER_SERIAL_RE = re.compile(r"^#(\d+)_(.+)$", re.IGNORECASE)


def build_num_to_serial(bench_dir) -> dict:
    """{"#NN": dut_serial} -- scans bench_dir (any depth, via rglob, same
    as every bench module's own read_all()) for the real-serial filename
    pattern. The "#NN" key format matches dut_id_of()'s own output
    (f"#{digits}") exactly, since both capture the same digit run right
    after the literal "#"."""
    mapping = {}
    for p in bench_dir.rglob("*_cns_temp_data.csv"):
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
    is missing, same degrade-gracefully rule as every other BT_TX ATE
    module's own extract_mode guard."""
    if PATHS.extract_mode == "bench" or not PATHS.ate_log_path:
        return set()
    try:
        with PATHS.ate_log_path.open(encoding="utf-8-sig", newline="") as fh:
            reader = csv.reader(fh)
            header = next(reader)
            return set(header[3:])
    except (FileNotFoundError, OSError):
        return set()


def _dut_means(values) -> dict:
    """values: [(value, key), ...] -> {key: mean(values for that key)}.
    Handles the common one-value-per-key case (Power/DEVM/FreqAcc: one
    Bench reading per DUT per exact combo, so this is just that value) and
    ACP's pooled-across-frequency case (multiple readings per DUT for the
    same offset/band Test Item, see module docstring) with the same code
    path."""
    groups = defaultdict(list)
    for v, k in values:
        groups[k].append(v)
    return {k: statistics.mean(vs) for k, vs in groups.items()}


def bench_by_serial(values, num_to_serial) -> dict:
    """[(value, "#NN"), ...] -> {dut_serial: mean value}. A "#NN" absent
    from num_to_serial is skipped rather than crashing -- shouldn't happen
    once num_to_serial is confirmed non-empty (main()'s own early-out), but
    kept defensive per this session's usual rule (degrade, don't guess)."""
    by_num = _dut_means(values)
    out = {}
    for num, v in by_num.items():
        serial = num_to_serial.get(num)
        if serial:
            out[serial] = v
    return out


def ate_by_serial_from_records(records) -> dict:
    """records: [{"by_serial": {serial: value}, ...}, ...] -> {serial: mean
    value}. Mirrors bench_by_serial's per-DUT mean -- ACP in particular can
    match more than one ATE record (one per Bench frequency) into a single
    Test Item row."""
    pooled = defaultdict(list)
    for rec in records:
        for serial, v in rec.get("by_serial", {}).items():
            pooled[serial].append(v)
    return {serial: statistics.mean(vs) for serial, vs in pooled.items()}


def make_row(name, bench_map, ate_map, all_serials) -> list:
    row = [name]
    for s in all_serials:
        row.append(bench_map.get(s))
    for s in all_serials:
        row.append(ate_map.get(s))
    return row


def power_raw_rows(num_to_serial) -> list:
    print("Raw Values: Power...")
    _ranges, stats = PE.get_ranges_and_stats(S.pwr_mod)
    raw, pa_slices_seen, dig_gain_seen = S.pwr_mod.read_all(PATHS.bench_dir)
    pa_slices_max, dig_gain_max = max(pa_slices_seen), max(dig_gain_seen)
    bench_data = S.pwr_mod.build_fixed_combos(raw, pa_slices_max, dig_gain_max)
    ate_idx = {}
    if PATHS.extract_mode != "bench":
        ate_idx = S.ate_pwr_mod.parse_ate_records_per_dut(S.ate_pwr_mod.ATE_CSV)
    rows = []
    for combo in stats.keys():
        packet, supply, _gain, freq = combo
        name = f"BT_TX_Power_{packet}_HPA{S.pwr_mod.hpa_token(supply)}_{S.pwr_mod.fmt_num(freq)}MHz"
        bench_map = bench_by_serial(bench_data.get(combo, []), num_to_serial)
        ate_map = {}
        if ate_idx:
            candidates = ate_idx.get((packet, supply), [])
            rec = S.ate_pwr_mod._nearest(freq, candidates, S.ate_pwr_mod.FREQ_TOL)
            if rec:
                ate_map = ate_by_serial_from_records([rec])
        rows.append((name, bench_map, ate_map))
    return rows


def devm_raw_rows(devm_spec_table, num_to_serial) -> list:
    rows = []
    for devm_metric, label, mod, ate_mod in S.DEVM_SECTIONS:
        print(f"Raw Values: {label}...")
        _ranges, stats = PE.get_ranges_and_stats(mod, devm_spec_table.get(devm_metric, {}))
        raw, pa_slices_seen, dig_gain_seen = mod.read_all(PATHS.bench_dir)
        pa_slices_max, dig_gain_max = max(pa_slices_seen), max(dig_gain_seen)
        bench_data = mod.build_fixed_combos(raw, pa_slices_max, dig_gain_max)
        ate_idx = {}
        if PATHS.extract_mode != "bench":
            ate_idx = ate_mod.parse_ate_records_per_dut(ate_mod.ATE_CSV)
        for combo in stats.keys():
            packet, supply, _gain, freq = combo
            name = f"BT_TX_{label}_{packet}_HPA{S.pwr_mod.hpa_token(supply)}_{S.pwr_mod.fmt_num(freq)}MHz"
            bench_map = bench_by_serial(bench_data.get(combo, []), num_to_serial)
            ate_map = {}
            if ate_idx:
                candidates = ate_idx.get((packet, supply), [])
                rec = ate_mod._nearest(freq, candidates, ate_mod.FREQ_TOL)
                if rec:
                    ate_map = ate_by_serial_from_records([rec])
            rows.append((name, bench_map, ate_map))
    return rows


def acp_raw_rows(num_to_serial) -> list:
    print("Raw Values: ACP...")
    data = ACP_BUILD.discover(ACP_BUILD.ACP_PNG_DIR)
    offset_stats_by_combo = ACP_BUILD.get_offset_stats_by_combo()

    raw, pa_slices_seen, dig_gain_seen = acp_mod.read_all_with_dut(PATHS.bench_dir)
    pa_slices_max, dig_gain_max = max(pa_slices_seen), max(dig_gain_seen)
    # build_fixed_combos() is generic (only ever .extend()s value lists, never
    # inspects their contents) so it works unchanged on read_all_with_dut()'s
    # (value, dut_id) tuples, same as on read_all()'s plain floats.
    bench_fixed = acp_mod.build_fixed_combos(raw, pa_slices_max, dig_gain_max)

    # Index by (packet, band_str) the same way ACP_BUILD.get_offset_stats_by_
    # combo() does -- gain is dropped from the key there too (verified: this
    # dataset's ACP rows are always pa_gain "0"), so mirror that exactly
    # rather than fixing a pre-existing quirk out of scope here.
    bench_by_key: dict = {}
    for (packet, _gain, band), offset_map in bench_fixed.items():
        bench_by_key[(packet, acp_mod.fmt_num(band))] = offset_map

    ate_idx = {}
    if PATHS.extract_mode != "bench":
        ate_idx = ate_acp_mod.parse_ate_acp_records_per_dut(ate_acp_mod.ATE_CSV)

    rows = []
    for packet, band in sorted(data.keys()):
        offset_stats = offset_stats_by_combo.get((packet, band), {})
        offset_map = bench_by_key.get((packet, band), {})
        for offset in sorted(offset_stats.keys()):
            name = (f"BT_TX_ACP_{packet}_Band{ACP_BUILD.fmt_band(band)}G_"
                    f"HPA{S.pwr_mod.hpa_token(acp_mod.FIXED_SUPPLY)}_Offset{offset}")
            freq_map = offset_map.get(offset, {})
            pooled_values = [vd for values in freq_map.values() for vd in values]
            bench_map = bench_by_serial(pooled_values, num_to_serial)
            ate_map = {}
            if ate_idx:
                candidates = ate_idx.get((packet, offset), [])
                # Dedup by the matched ATE record's own freq, same as
                # make_ate_tx_acp_revision2.get_ate_data()'s ate_freq_map --
                # two Bench frequencies nearest-matching the same single ATE
                # record must not double-count that record's values.
                ate_freq_map = {}
                for bench_freq in freq_map:
                    rec = ate_acp_mod._nearest(bench_freq, candidates, ate_acp_mod.FREQ_TOL)
                    if rec:
                        ate_freq_map[rec["freq"]] = rec
                ate_map = ate_by_serial_from_records(list(ate_freq_map.values()))
            rows.append((name, bench_map, ate_map))
    return rows


def freqacc_raw_rows(freqacc_spec_table, num_to_serial) -> list:
    print("Raw Values: FreqAcc (all 9 sub-metrics)...")
    _ranges_by_metric, stats_by_metric = FREQACC_BUILD.get_ranges_and_stats_by_metric(freqacc_spec_table)

    metric_cols = {key: col for key, col, _xlabel, _packets in S.fa_mod.METRICS}
    raw, pa_slices_seen, dig_gain_seen = S.fa_mod.read_all(PATHS.bench_dir, metric_cols)
    pa_slices_max, dig_gain_max = max(pa_slices_seen), max(dig_gain_seen)

    ate_idx = {}
    if PATHS.extract_mode != "bench":
        ate_idx = S.ate_fa_mod.parse_all_ate_records_per_dut(S.ate_fa_mod.ATE_CSV)

    rows = []
    for key, _col, _xlabel, packets in S.fa_mod.METRICS:
        stats = stats_by_metric.get(key)
        if not stats:
            continue
        fixed = S.fa_mod.build_fixed_combos(raw[key], pa_slices_max, dig_gain_max, packets)
        label = S.FA_NAME_SAFE.get(key, key)
        for combo in stats.keys():
            packet, supply, _gain, freq = combo
            name = f"BT_TX_{label}_{packet}_HPA{S.pwr_mod.hpa_token(supply)}_{S.fa_mod.fmt_num(freq)}MHz"
            bench_map = bench_by_serial(fixed.get(combo, []), num_to_serial)
            ate_map = {}
            if ate_idx:
                candidates = S.ate_fa_mod._ate_candidates(ate_idx, packet, supply, key)
                rec = S.ate_fa_mod._nearest(freq, candidates, S.ate_fa_mod.FREQ_TOL)
                if rec:
                    ate_map = ate_by_serial_from_records([rec])
            rows.append((name, bench_map, ate_map))
    return rows


def bw_raw_rows(num_to_serial) -> list:
    """6dB/20dB Bandwidth (2026-08-26 user request, extended 2026-08-27 to
    Raw Values/Config sheets). 20dB has no ATE counterpart at all (confirmed
    empty on the real ATE log) -- its ate_map is always {}, unlike 6dB
    which checks extract_mode like every other family."""
    print("Raw Values: 6dB/20dB BW...")
    raw, pa_slices_seen, dig_gain_seen = bw_mod.read_all(PATHS.bench_dir)
    pa_slices_max, dig_gain_max = max(pa_slices_seen), max(dig_gain_seen)
    bench_bw6 = bw_mod.build_fixed_combos(raw["bw6"], pa_slices_max, dig_gain_max)
    bench_bw20 = bw_mod.build_fixed_combos(raw["bw20"], pa_slices_max, dig_gain_max)

    ate_idx = {}
    if PATHS.extract_mode != "bench":
        ate_idx = ate_bw_mod.parse_ate_records_per_dut(ate_bw_mod.ATE_CSV)

    rows = []
    for combo, values in bench_bw6.items():
        packet, supply, _gain, freq = combo
        name = f"BT_TX_BW06DB_{packet}_HPA{S.pwr_mod.hpa_token(supply)}_{bw_mod.fmt_num(freq)}MHz"
        bench_map = bench_by_serial(values, num_to_serial)
        ate_map = {}
        if ate_idx:
            candidates = ate_idx.get((packet, supply), [])
            rec = ate_bw_mod._nearest(freq, candidates, ate_bw_mod.FREQ_TOL)
            if rec:
                ate_map = ate_by_serial_from_records([rec])
        rows.append((name, bench_map, ate_map))
    for combo, values in bench_bw20.items():
        packet, supply, _gain, freq = combo
        name = f"BT_TX_BW20DB_{packet}_HPA{S.pwr_mod.hpa_token(supply)}_{bw_mod.fmt_num(freq)}MHz"
        bench_map = bench_by_serial(values, num_to_serial)
        rows.append((name, bench_map, {}))
    return rows


def parse_spec_tables():
    """(devm_spec_table, freqacc_spec_table) -- hoisted out of main() so
    update_bt_tx_summary_by_vendor.py can call collect_raw_rows() (below)
    without duplicating this parse."""
    devm_spec_table = S.devm_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    freqacc_spec_table = S.freqacc_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    return devm_spec_table, freqacc_spec_table


def collect_raw_rows(num_to_serial, devm_spec_table, freqacc_spec_table) -> list:
    """(name, bench_by_serial, ate_by_serial) for all 911 Test Item rows
    across all 14 metric families (Power, DEVM RMS/Peak/99pct, ACP,
    FreqAcc's 9 sub-metrics, 6dB/20dB BW) -- same name/order as "BT_TX
    Summary". This is
    the traversal every per-row dict in this module is built from; kept as
    ONE function (rather than duplicated in update_bt_tx_summary_by_vendor.py)
    per 2026-08-25 refactor so both the flat Raw Values sheet (main() below,
    via flatten_rows()) and the per-group Summary/Raw-Values sheets share
    the exact same underlying per-DUT data."""
    rows = []
    rows += power_raw_rows(num_to_serial)
    rows += devm_raw_rows(devm_spec_table, num_to_serial)
    rows += acp_raw_rows(num_to_serial)
    rows += freqacc_raw_rows(freqacc_spec_table, num_to_serial)
    rows += bw_raw_rows(num_to_serial)
    return rows


def flatten_rows(rows, all_serials) -> list:
    """[(name, bench_map, ate_map), ...] -> flat make_row() lists, for the
    fixed Bench-block-then-ATE-block "Raw Values" layout this script's own
    main() writes."""
    return [make_row(name, bench_map, ate_map, all_serials) for name, bench_map, ate_map in rows]


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

    print("Parsing spec workbooks (needed for DEVM/FreqAcc's own outlier/"
          "range windows, to keep this sheet's row set identical to the "
          "summary sheet's)...")
    devm_spec_table, freqacc_spec_table = parse_spec_tables()

    dict_rows = collect_raw_rows(num_to_serial, devm_spec_table, freqacc_spec_table)
    print(f"Total rows: {len(dict_rows)}")
    rows = flatten_rows(dict_rows, all_serials)

    print("Writing sheet into existing workbook...")
    wb = openpyxl.load_workbook(SUMMARY_PATH)
    write_flat_raw_values_sheet(wb, rows, all_serials)
    wb.save(SUMMARY_PATH)
    n_cols = 1 + 2 * len(all_serials)
    print(f"\nSaved: {SUMMARY_PATH}  (added sheet 'Raw Values', {len(rows)} row(s) x {n_cols} col(s))")


def write_flat_raw_values_sheet(wb, rows, all_serials) -> None:
    """Creates/replaces the 2-level-header ("Bench"/"ATE" -> serial) "Raw
    Values" sheet in an already-open workbook `wb` -- does NOT save. Hoisted
    out of main() so update_bt_tx_summary_by_vendor.py's graceful-degradation
    path (missing/incomplete 40pcs_config.txt) can build the exact same
    fallback sheet this script's own main() writes, without duplicating the
    write logic."""
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


if __name__ == "__main__":
    main()
