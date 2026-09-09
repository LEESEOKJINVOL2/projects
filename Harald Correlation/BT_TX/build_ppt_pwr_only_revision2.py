"""Build a Power-only BT_TX PPT deck (2026-08-14).

Standalone half-deck builder, same pattern this project already uses for
build_ppt_full_revision2.py / build_ppt_freqacc_revision2.py (each a
working, independently runnable partial deck, per that module's own
docstring) -- Power is the only metric wired to the new
Harald_P2_Test_Coverage_v4.4_JSCK.xlsx spec workbook and the new 10-DUT
bench pull (Bench_40pcs/Tx) so far; DEVM/ACP/FreqAcc (and Power's own ATE
counterpart) still read the old workbook/old dataset assumptions and need
the same treatment before they belong in the same deck as this one.

Reuses every slide-building/table/ruler helper from
build_ppt_pwr_evm_revision2.py (PE) unchanged -- only the spec workbook
(PATHS.bt_tx_spec_workbook override, not PE's own hardcoded
PATHS.spec_workbook) and the section set (Power only, no DEVM) differ from
that module's own main().
"""

from __future__ import annotations

from pptx import Presentation

import build_ppt_pwr_evm_revision2 as PE
import bt_tx_power_spec_lookup as power_spec_lookup

PATHS = PE.PATHS
REPORT_DATE = PE.REPORT_DATE


def main() -> None:
    prs = Presentation(PE.TEMPLATE_PATH)
    content_layout_idx = PE._find_layout(prs, "1_", 5)

    print("Computing Power Y-axis ranges and stats...")
    pwr_ranges, pwr_stats = PE.get_ranges_and_stats(PE.pwr_mod)

    spec_workbook = PATHS.bt_tx_spec_workbook or PATHS.spec_workbook
    print(f"Parsing Power spec limits from {spec_workbook.name}...")
    power_spec_table = power_spec_lookup.parse_spec(spec_workbook)

    def get_power_spec(packet, gain, freq, supply):
        return power_spec_lookup.lookup(power_spec_table, packet, freq, supply)

    divider_source = prs.slides[0]
    first_ph = PE._get_ph(divider_source, 1) or PE._get_ph(divider_source, 0)
    if first_ph:
        PE._set_section_text(first_ph, "Power")
    power_divider = divider_source

    print("Building Power slides...")
    # spec_constant=False -- Power's USL/LSL genuinely differs column-to-
    # column within one slide (by pa_supply/frequency block), same as PE's
    # own main() (see that function's comment).
    power_slides = PE.build_metric_slides(prs, content_layout_idx, "Power", PE.PWR_PNG_DIR, pwr_ranges, pwr_stats,
                                           get_spec=get_power_spec, spec_constant=False)
    print(f"  {len(power_slides)} Power slides")

    title_slide = PE.add_title_slide(prs)
    ordered = [title_slide, power_divider, *power_slides]
    PE._reorder_slides(prs, ordered)

    report_dir = PATHS.report_dir
    report_dir.mkdir(parents=True, exist_ok=True)
    out_path = report_dir / f"{REPORT_DATE}_BT_TX_Power_Only_Revision2.pptx"
    prs.save(out_path)
    print(f"\nSaved: {out_path}")
    print(f"Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
