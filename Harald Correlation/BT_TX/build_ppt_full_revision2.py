"""Build the combined Power / DEVM (RMS, Peak, 99pct) / ACP revision2 PPT
deck, in that section order. Reuses the slide-building logic already
validated in build_ppt_pwr_evm_revision2.py (Power/DEVM) and
build_ppt_acp_revision2.py (ACP) -- this script only handles the top-level
deck assembly (title slide, section dividers in order, final save).

DEVM (2026-07-19): "EVM" and "DEVM" mean the same thing in this project;
the single "EVM" section is now three DEVM sections (RMS/Peak/99pct), each
its own divider + slides, per bt_tx_devm_spec_lookup.py's packet mapping.
"""

from __future__ import annotations

import datetime
from pathlib import Path

from pptx import Presentation

import build_ppt_pwr_evm_revision2 as PE
import build_ppt_acp_revision2 as ACP
import bt_tx_power_spec_lookup as power_spec_lookup
import bt_tx_devm_spec_lookup as devm_spec_lookup
import make_ate_tx_pwr_revision2 as ate_pwr_mod
import make_ate_tx_devm_rms_revision2 as ate_devm_rms_mod
import make_ate_tx_devm_peak_revision2 as ate_devm_peak_mod
import make_ate_tx_devm_99pct_revision2 as ate_devm_99pct_mod
import make_ate_tx_acp_revision2 as ate_acp_mod
import make_bt_tx_bw_revision1 as bw_mod
import make_ate_tx_bw_revision1 as ate_bw_mod
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")
BASE_DIR = PATHS.base_dir
# Computed at run time (not hardcoded) so every rebuild's filename/subtitle
# reflects the actual build date instead of going stale (2026-07-18 fix).
REPORT_DATE = datetime.date.today().isoformat()


def _ate_data_or_empty(mod, label: str) -> dict:
    """2026-07-22: extract_mode="bench" must make build_ppt treat ATE as
    genuinely absent (an empty dict), not just "didn't re-run the ATE
    script" -- otherwise a stale ATE PNG/match from an earlier "both" run
    would silently leak into a report meant to be Bench-only. See
    paths.py's extract_mode docstring."""
    if PATHS.extract_mode == "bench":
        print(f"  extract_mode=bench -- skipping ATE for {label}")
        return {}
    ate_data, _bench, _slices_max, _dig_max = mod.get_ate_data()
    return ate_data


def build_sections(prs, content_layout_idx, new_divider) -> list:
    """Build this deck's Power / DEVM x3 / ACP sections into `prs` and return
    them as [(divider, slides), ...] in final order -- WITHOUT adding a title
    slide, reordering, or saving.

    2026-07-28: split out of main() so build_ppt_combined_revision2.py can
    append FreqAcc's sections into the SAME presentation instead of building
    two decks and merging them through PowerPoint COM afterwards. main() below
    still produces the standalone Power/DEVM/ACP deck, unchanged.

    `new_divider(text)` is supplied by the caller and owns the "first section
    reuses the template's slide 0, later ones duplicate it" bookkeeping -- that
    matters because _reorder_slides() drops any slide missing from its ordered
    list, so a never-reused slide 0 would linger as an orphaned part. Same
    closure pattern UWB's and WIFI's deck builders already use.
    """
    print("Computing Power Y-axis ranges and stats...")
    pwr_ranges, pwr_stats = PE.get_ranges_and_stats(PE.pwr_mod)

    print("Parsing Power spec limits...")
    power_spec_table = power_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    def get_power_spec(packet, gain, freq, supply):
        return power_spec_lookup.lookup(power_spec_table, packet, freq, supply)

    print("Parsing DEVM spec limits...")
    devm_spec_table = devm_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    power_divider = new_divider("Power")

    print("Matching Power combos against ATE log (BT_TX only)...")
    # 2026-07-20: real ATE data now lands directly in the reserved bottom
    # half instead of an empty placeholder, for Power only so far -- DEVM
    # and ACP still get the plain empty box below (no ATE generator script
    # exists yet for those metrics). ~30% of combos genuinely have no ATE
    # counterpart (not in the ATE test plan, or the ATE item exists but its
    # 40 DUT cells are all blank e.g. UHDR32/48) -- those render as a
    # per-column "No ATE match" hole, never blocking the slide.
    power_ate_data = _ate_data_or_empty(ate_pwr_mod, "Power")
    n_ate_matched = sum(1 for v in power_ate_data.values() if v["status"] != "NONE")
    print(f"  {n_ate_matched}/{len(power_ate_data)} Power combos have ATE data")

    print("Building Power slides...")
    # Power's USL/LSL genuinely differs column-to-column within one slide
    # (pa_supply/frequency block) -- spec_constant=False shows it per-column
    # instead of merging (see share/PPT_FORMAT_STANDARD.md).
    power_slides = PE.build_metric_slides(prs, content_layout_idx, "Power", PE.PWR_PNG_DIR, pwr_ranges, pwr_stats,
                                           get_spec=get_power_spec, spec_constant=False, ate_data=power_ate_data)
    print(f"  {len(power_slides)} Power slides")

    # 2026-07-20: one ATE matching module per DEVM submetric, same pattern
    # as Power's ate_pwr_mod above.
    DEVM_ATE_MODS = {
        "rms": ate_devm_rms_mod,
        "peak": ate_devm_peak_mod,
        "99pct": ate_devm_99pct_mod,
    }

    devm_sections = []  # [(divider, slides), ...] in RMS -> Peak -> 99pct order
    for devm_metric, label, mod, png_dir in PE.DEVM_SECTIONS:
        print(f"Computing {label} Y-axis ranges and stats...")
        ranges, stats = PE.get_ranges_and_stats(mod, devm_spec_table.get(devm_metric, {}))

        def get_spec(packet, gain, freq, supply, _metric=devm_metric):
            return devm_spec_lookup.lookup(devm_spec_table, _metric, packet)

        divider = new_divider(label)

        print(f"Matching {label} combos against ATE log...")
        devm_ate_data = _ate_data_or_empty(DEVM_ATE_MODS[devm_metric], label)
        n_matched = sum(1 for v in devm_ate_data.values() if v["status"] != "NONE")
        print(f"  {n_matched}/{len(devm_ate_data)} {label} combos have ATE data")

        print(f"Building {label} slides...")
        # DEVM's spec is keyed by packet only -- constant across every
        # column of a given slide, so spec_constant=True merges it.
        slides = PE.build_metric_slides(prs, content_layout_idx, label, png_dir, ranges, stats,
                                         get_spec=get_spec, spec_constant=True, ate_data=devm_ate_data)
        print(f"  {len(slides)} {label} slides")
        devm_sections.append((divider, slides))

    print("\n=== 6dB BW / 20dB BW sections ===")
    # 2026-08-27: new items, per-user Option A -- each its own section, same
    # pattern as Power/DEVM/FreqAcc's sub-metrics. No spec/USL-LSL exists for
    # either (get_spec=None throughout). ruler_fmt="{:.1f}" (was "{:.3f}"
    # when the axis itself was still scale-aware/fine-grained) -- the
    # customer later asked for the axis min/max floor/ceil'd to a fixed 1
    # decimal place instead (see make_bt_tx_bw_revision1.BW_AXIS_STEP), so
    # the ruler label format was loosened to match -- 3 decimals on an
    # axis that's now always an exact 1-decimal value (e.g. 0.3/0.4) would
    # just print trailing zeros.
    BW_RULER_FMT = "{:.1f}"

    bw6_meta = bw_mod.METRIC_META["bw6"]
    bw6_divider = new_divider(bw6_meta["section_label"])
    print("Computing 6dB BW Y-axis ranges and stats...")
    bw6_ranges, bw6_stats, _bw6_slices_max, _bw6_dig_max = bw_mod.get_ranges_and_stats("bw6")

    print("Matching 6dB BW combos against ATE log...")
    bw6_ate_data = _ate_data_or_empty(ate_bw_mod, "6dB BW")
    n_bw6_matched = sum(1 for v in bw6_ate_data.values() if v["status"] != "NONE")
    print(f"  {n_bw6_matched}/{len(bw6_ate_data)} 6dB BW combos have ATE data")

    print("Building 6dB BW slides...")
    bw6_slides = PE.build_metric_slides(
        prs, content_layout_idx, bw6_meta["metric_label"], PATHS.result_png_dir / bw6_meta["dirname"],
        bw6_ranges, bw6_stats, get_spec=None, ate_data=bw6_ate_data, ruler_fmt=BW_RULER_FMT)
    print(f"  {len(bw6_slides)} 6dB BW slides")

    bw20_meta = bw_mod.METRIC_META["bw20"]
    bw20_divider = new_divider(bw20_meta["section_label"])
    print("Computing 20dB BW Y-axis ranges and stats...")
    bw20_ranges, bw20_stats, _bw20_slices_max, _bw20_dig_max = bw_mod.get_ranges_and_stats("bw20")

    print("Building 20dB BW slides...")
    # No ATE counterpart exists for BW20DB at all (confirmed empty on the
    # real ATE log -- see make_ate_tx_bw_revision1.py's module docstring).
    # 2026-08-27 (user request, changed from the original ate_data=None
    # single-row layout): ate_data={} (NOT None) so build_metric_slides
    # takes the same two-row Bench+ATE layout as every ATE-backed family --
    # every combo's ate_data.get(combo) then comes back None, so each
    # column renders its normal per-column "No ATE match" placeholder
    # (same box every other family already shows for an unmatched combo),
    # just for ALL columns here since none can ever match.
    bw20_slides = PE.build_metric_slides(
        prs, content_layout_idx, bw20_meta["metric_label"], PATHS.result_png_dir / bw20_meta["dirname"],
        bw20_ranges, bw20_stats, get_spec=None, ate_data={}, ruler_fmt=BW_RULER_FMT)
    print(f"  {len(bw20_slides)} 20dB BW slides")

    bw_sections = [(bw6_divider, bw6_slides), (bw20_divider, bw20_slides)]

    acp_divider = new_divider("ACP")

    print("Building ACP slides...")
    acp_data = ACP.discover(ACP.ACP_PNG_DIR)
    acp_legends = ACP.discover_legends(ACP.ACP_LEGEND_DIR)
    print("  Computing per-combo Y-axis range...")
    acp_axis_range_by_combo = ACP.get_axis_range_by_combo()
    print("  Computing per-offset stats...")
    acp_offset_stats_by_combo = ACP.get_offset_stats_by_combo()
    print("  Parsing ACP spec limits...")
    import bt_tx_acp_spec_lookup as acp_spec_lookup
    acp_spec_table = acp_spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    print("  Matching ACP combos against ATE log...")
    # 2026-08-14 real bug: this used to look up ate_data with a hardcoded
    # gain="0" -- the exact same root cause as the Power/DEVM GAIN="0" bug
    # fixed earlier in build_ppt_pwr_evm_revision2.py (this dataset's real
    # pa_gain is '1'), except that fix never touched this file, so ACP kept
    # silently missing all 178 real ATE matches. Derive it from the actual
    # bench data instead of assuming a constant.
    if PATHS.extract_mode == "bench":
        print("  extract_mode=bench -- skipping ATE for ACP")
        acp_ate_data, acp_gain = {}, "0"
    else:
        acp_ate_data, _acp_bench_data, _acp_slices_max, _acp_dig_max = ate_acp_mod.get_ate_data()
        acp_gain = max({k[1] for k in _acp_bench_data.keys()}, key=float) if _acp_bench_data else "0"
    n_acp_matched = sum(1 for v in acp_ate_data.values() if v["status"] != "NONE")
    print(f"  {n_acp_matched}/{len(acp_ate_data)} ACP (combo,offset) pairs have ATE data")

    acp_slides = []
    for packet, band in sorted(acp_data.keys()):
        title = f"BT_TX_ACP_{packet}_{ACP.fmt_band(band)}G"
        y_min, y_max = acp_axis_range_by_combo.get(
            (packet, band), (ACP.FIXED_Y_MIN_DEFAULT, ACP.FIXED_Y_MAX_DEFAULT))
        offset_stats = acp_offset_stats_by_combo.get((packet, band), {})
        # ACP's Bench combo key includes gain (acp_gain, derived above from
        # the real data -- see the 2026-08-14 comment) -- ate_acp_mod keys
        # by (packet, gain, band, offset) with band as a FLOAT (from the
        # raw CSV column), while `band` here is discover()'s regex string
        # capture -- must convert before the lookup or it silently misses
        # every entry.
        band_f = float(band)
        acp_ate_slice = {off: acp_ate_data.get((packet, acp_gain, band_f, off)) for off in ate_acp_mod.OFFSETS}

        def get_acp_spec(off, _packet=packet):
            return acp_spec_lookup.lookup(acp_spec_table, _packet, int(off))

        acp_slides.append(ACP.add_acp_slide(prs, content_layout_idx, title, acp_data[(packet, band)],
                                             offset_stats, legend_path=acp_legends.get((packet, band)),
                                             y_min=y_min, y_max=y_max, get_spec=get_acp_spec,
                                             ate_data=acp_ate_slice))
    print(f"  {len(acp_slides)} ACP slides")

    return [(power_divider, power_slides), *devm_sections, *bw_sections, (acp_divider, acp_slides)]


def main() -> None:
    """Standalone Power/DEVM/ACP-only deck.

    2026-07-28: no longer part of run_all.py's pipeline -- that now builds one
    combined deck via build_ppt_combined_revision2.py. Kept as a working entry
    point for anyone who wants just this half.
    """
    prs = Presentation(PE.TEMPLATE_PATH)
    content_layout_idx = PE._find_layout(prs, "1_", 5)

    sections = build_sections(prs, content_layout_idx, PE.make_divider_factory(prs))

    # 2026-07-22: title_prefix/title_suffix are config.xml-driven (see
    # paths.py's docstring) -- was the inconsistent, hardcoded "Bench BT TX
    # Power / DEVM / ACP Distribution" before this, now the same
    # "{prefix} {project} {suffix}" pattern every project's title uses.
    PE.TITLE_TEXT = f"{PATHS.title_prefix} {PATHS.project} {PATHS.title_suffix}"
    PE.SUBTITLE_LINES = ["RF LAB1 Team", REPORT_DATE]
    ordered = [PE.add_title_slide(prs)]
    for divider, slides in sections:
        ordered.extend([divider, *slides])
    PE._reorder_slides(prs, ordered)

    report_dir = PATHS.report_dir
    report_dir.mkdir(parents=True, exist_ok=True)
    out_path = report_dir / f"{REPORT_DATE}_BT_TX_PWR_DEVM_ACP_Revision2.pptx"
    prs.save(out_path)
    print(f"\nSaved: {out_path}")
    print(f"Total slides: {len(prs.slides._sldIdLst)}")


if __name__ == "__main__":
    main()
