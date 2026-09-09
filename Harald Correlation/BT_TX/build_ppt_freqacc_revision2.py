"""Build the Frequency-Accuracy / Modulation-Accuracy metric PPT slides, in
the EXACT same format as build_ppt_pwr_evm_revision2.py (Power/EVM):
  * 2.4GHz band -> one slide, all (pa_supply x frequency) images in a
    single row, "pa_supply{N}" label above the center frequency of each
    3-image block.
  * 5/6GHz bands -> split by pa_supply, one slide per (band, supply).
  * Top-half-of-content-area placement, Min/Max ruler on the left.

One section (divider + slides) per metric, in the order defined by
make_bt_tx_freqacc_revision2.METRICS. Metrics with no data (e.g. HDT's own
"drift rate" column, which is empty in this dataset) are skipped.

Reuses build_ppt_pwr_evm_revision2.py's slide-building functions directly
-- this script only adds the metric-specific Y-axis range computation and
the section/deck assembly.
"""

from __future__ import annotations

import datetime
from pathlib import Path

from pptx import Presentation

import build_ppt_pwr_evm_revision2 as PE
import bt_tx_freqacc_spec_lookup as spec_lookup
import make_bt_tx_freqacc_revision2 as FA
import make_ate_tx_freqacc_revision2 as ATE_FA
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")
BASE_DIR = PATHS.base_dir
# Computed at run time (not hardcoded) so every rebuild's filename/subtitle
# reflects the actual build date instead of going stale (2026-07-18 fix).
REPORT_DATE = datetime.date.today().isoformat()


def _ate_data_by_metric_or_empty() -> dict:
    """2026-07-22: extract_mode="bench" must make every FreqAcc sub-metric
    treat ATE as genuinely absent (an empty dict per key, so
    ate_data_by_metric.get(key) still returns {} not None -- {} keeps the
    per-column "No ATE match" rendering; None would fall back to the OLD
    single big placeholder box instead). See paths.py's extract_mode
    docstring and build_ppt_full_revision2.py's _ate_data_or_empty()."""
    if PATHS.extract_mode == "bench":
        print("  extract_mode=bench -- skipping ATE for FreqAcc")
        return {key: {} for key, *_rest in FA.METRICS}
    ate_data_by_metric, _bench_by_metric, _ranges2, _sm, _dm = ATE_FA.get_ate_data_all()
    return ate_data_by_metric

LABELS = {
    "icft": "ICFT",
    "cfo": "CFO",
    "drift_rate": "Drift Rate",
    "df1_avg": "Df1 Average",
    "df2_avg": "Df2 Average",
    "df2_max": "Df2 Max",
    "omega_i": "ωi",
    "omega_o": "ωo",
    "omega_io": "ωi+ωo",
    "freq_drift": "Drift Rate (HDT)",
}


def get_ranges_and_stats_by_metric(spec_table):
    """One shared read_all() pass across every metric's column (avoids
    re-scanning the 40 CSVs once per metric), then per-metric
    build_fixed_combos + build_packet_ranges -- identical range logic to
    what make_bt_tx_freqacc_revision2.py itself used when drawing the
    charts, so the PPT ruler always matches what's baked into the PNGs
    (2026-07-20: that now means spec_fn-anchored + never-excludes-a-point,
    see make_bt_tx_freqacc_revision2.py's module docstring -- passing
    spec_table through here keeps this in sync with the chart script's own
    axis, not last week's ceil/floor-only version).

    Also computes combo_stats[key][(packet,supply,gain,freq)] =
    (mean,min,max,std) straight from every value for that combo -- FreqAcc
    no longer excludes anything (2026-07-20), so there's no more kept/
    outlier split to run through find_combo_outliers (removed from the
    chart script; this table must match the chart's own no-exclusion
    Min/Max exactly, see share/PPT_FORMAT_STANDARD.md)."""
    metric_cols = {key: col for key, col, _label, _packets in FA.METRICS}
    raw, pa_slices_seen, dig_gain_seen = FA.read_all(PATHS.bench_dir, metric_cols)
    pa_slices_max = max(pa_slices_seen)
    dig_gain_max = max(dig_gain_seen)
    ranges_result = {}
    stats_result = {}
    for key, col, xlabel, packets in FA.METRICS:
        fixed = FA.build_fixed_combos(raw[key], pa_slices_max, dig_gain_max, packets)
        if not fixed:
            print(f"  {key}: no data, skipping")
            continue

        def spec_fn(packet, band, _key=key):
            return spec_lookup.lookup(spec_table, _key, packet, band)

        ranges_result[key] = FA.build_packet_ranges(fixed, spec_fn=spec_fn)[0]
        combo_stats = {}
        for combo, values in fixed.items():
            vals = [v for v, _dut in values]
            if vals:
                combo_stats[combo] = PE._stats(vals)
        stats_result[key] = combo_stats
    return ranges_result, stats_result


def build_sections(prs, content_layout_idx, new_divider) -> list:
    """Build one section per FreqAcc sub-metric into `prs` and return them as
    [(divider, slides), ...] in METRICS order -- WITHOUT adding a title slide,
    reordering, or saving.

    2026-07-28: split out of main() so build_ppt_combined_revision2.py can
    append these sections to the SAME presentation that already holds
    Power/DEVM/ACP, replacing the PowerPoint-COM merge that used to stitch two
    finished decks together. main() below still produces the standalone FreqAcc
    deck, unchanged. See build_ppt_full_revision2.make_divider_factory() for
    what `new_divider` owns.
    """
    print("Parsing FreqAcc spec limits...")
    spec_table = spec_lookup.parse_spec(PATHS.bt_tx_spec_workbook or PATHS.spec_workbook)

    print("Computing per-metric Y-axis ranges and stats...")
    ranges_by_metric, stats_by_metric = get_ranges_and_stats_by_metric(spec_table)

    print("Matching FreqAcc combos against ATE log (all 9 sub-metrics, one pass)...")
    ate_data_by_metric = _ate_data_by_metric_or_empty()
    for key in ate_data_by_metric:
        data = ate_data_by_metric[key]
        if data:
            n_matched = sum(1 for v in data.values() if v["status"] != "NONE")
            print(f"  {key}: {n_matched}/{len(data)} combos have ATE data")

    sections = []
    for key, col, xlabel, packets in FA.METRICS:
        if key not in ranges_by_metric:
            continue
        label = LABELS.get(key, key)

        divider = new_divider(label)

        png_dir = PATHS.result_png_dir / f"{key}_revision2"
        print(f"Building {label} slides...")

        def get_spec(packet, gain, freq, supply, _key=key):
            return spec_lookup.lookup(spec_table, _key, packet, PE.band_of(freq))

        # FreqAcc's spec is keyed by (metric, packet[, band]); a given slide
        # only ever covers ONE band, so the spec is constant across all of
        # that slide's columns -- spec_constant=True merges it into one cell.
        slides = PE.build_metric_slides(prs, content_layout_idx, label, png_dir, ranges_by_metric[key],
                                         stats_by_metric[key], get_spec=get_spec, spec_constant=True,
                                         ate_data=ate_data_by_metric.get(key))
        print(f"  {len(slides)} slides")
        sections.append((divider, slides))

    return sections


def main() -> None:
    """Standalone FreqAcc-only deck.

    2026-07-28: no longer part of run_all.py's pipeline -- that now builds one
    combined deck via build_ppt_combined_revision2.py. Kept as a working entry
    point for anyone who wants just this half.
    """
    prs = Presentation(PE.TEMPLATE_PATH)
    content_layout_idx = PE._find_layout(prs, "1_", 5)

    sections = build_sections(prs, content_layout_idx, PE.make_divider_factory(prs))

    # 2026-07-22: title_prefix/title_suffix are config.xml-driven (see
    # paths.py's docstring) -- was the hardcoded "Bench BT TX Freq / Mod
    # Accuracy" before this (also the one deck with no "Distribution" suffix
    # at all), now the same "{prefix} {project} {suffix}" pattern every
    # project's title uses.
    PE.TITLE_TEXT = f"{PATHS.title_prefix} {PATHS.project} {PATHS.title_suffix}"
    PE.SUBTITLE_LINES = ["RF LAB1 Team", REPORT_DATE]
    ordered = [PE.add_title_slide(prs)]
    for divider, slides in sections:
        ordered.extend([divider, *slides])
    PE._reorder_slides(prs, ordered)

    report_dir = PATHS.report_dir
    report_dir.mkdir(parents=True, exist_ok=True)
    out_path = report_dir / f"{REPORT_DATE}_BT_TX_FreqAcc_Revision2.pptx"
    prs.save(out_path)
    print(f"\nSaved: {out_path}")
    print(f"Total slides: {len(prs.slides._sldIdLst)}")


if __name__ == "__main__":
    main()
