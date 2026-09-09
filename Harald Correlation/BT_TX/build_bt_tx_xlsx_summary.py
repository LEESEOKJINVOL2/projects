"""XLSX summary with Bench-vs-ATE delta columns, per user request 2026-08-15
("mean std min max이랑 이거의 delta 값도 추출하는 summary파일 xlsx 형식으로
만들어줘 ... 전부 abs 처리해주면 되고").

Reuses build_bt_tx_summary_csv.py's row-building functions (Power/DEVM/ACP/
FreqAcc) verbatim -- imported, not modified, so this never re-derives a
stat, it only adds delta columns on top and writes .xlsx (merged Bench/ATE/
Delta header via the shared xlsx_summary_writer.py) instead of that
script's flat CSV.
"""
from __future__ import annotations

import datetime
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths
import xlsx_summary_writer as xw

import build_bt_tx_summary_csv as S

PATHS = bt_tx_paths.get_paths("BT_TX")
REPORT_DATE = datetime.date.today().isoformat()


def add_deltas(rows: list) -> list:
    """S's rows are [name, Bench Mean/Std/Min/Max/USL/LSL (idx 1-6), ATE
    Mean/Std/Min/Max/USL/LSL (idx 7-12)] -- delta only for Mean/Std/Min/Max
    (idx 1-4 vs 7-10), never USL/LSL (idx 5-6 vs 11-12)."""
    out = []
    for row in rows:
        deltas = [xw.delta(row[1 + i], row[7 + i]) for i in range(4)]
        out.append(row + deltas)
    return out


def main() -> None:
    print("Parsing spec workbooks...")
    power_spec_table = S.power_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    devm_spec_table = S.devm_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    acp_spec_table = S.acp_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)
    freqacc_spec_table = S.freqacc_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    rows = []
    rows += S.build_power_rows(power_spec_table)
    rows += S.build_devm_rows(devm_spec_table)
    rows += S.build_acp_rows(acp_spec_table)
    rows += S.build_freqacc_rows(freqacc_spec_table)
    rows += S.build_bw_rows()
    print(f"Total rows: {len(rows)}")

    rows = add_deltas(rows)

    out_path = PATHS.report_dir / f"{REPORT_DATE}_BT_TX_Summary.xlsx"
    xw.write_summary_xlsx(rows, out_path, "BT_TX Summary")
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()
