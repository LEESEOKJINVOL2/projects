"""FWBT_TRIMTEST distribution PNGs -- Bench + ATE, per user request
2026-08-15 to graduate the layout-only draft (build_ppt_fwbt_trimtest_draft.py)
into a real report with data.

Item universe + spec: FWBT_TRIMTEST_DP6P1P3_renamed.xlsx's "Test Items"
sheet (2737 rows, Test Item/USL/LSL/Unit) -- the cleaned display names
produced by the earlier naming cleanup this session. USL/LSL always come
from THIS sheet, never from Bench's own embedded LLIM/ULIM column (unlike
Cal/make_cal_distribution.py, which has no separate spec source and reads
its own Bench log's LLIM/ULIM instead) -- this project's "item tab" IS the
spec source, per user instruction.

Raw-name mapping: the cleaned names in that sheet do NOT appear verbatim in
Bench_40pcs/Cal or the shared ATE log -- those logs carry the ORIGINAL
padded name (e.g. "BT2G_A4P_BAND0_A0_X_X_X_X_NV_VXXX"). Reversing the exact
same clean() transform against every raw Bench-log name and keeping the
ones that land on one of our 2737 cleaned names recovers a 1:1,
collision-free mapping for all 2737 items (verified 2026-08-15) -- this
avoids depending on the ORIGINAL (pre-cleanup) xlsx file at all, which only
still exists in the user's Downloads folder, not in this workspace.

Bench/ATE reading is reused directly from Cal/make_cal_distribution.py
(read_bench, read_ate, safe_filename, wrap_label, BINS, _font_scale,
_REF_FIGSIZE) -- this project's Bench data physically lives in the exact
same Bench_40pcs/Cal folder Cal already parses (config.xml's
BT_TX_TRIMTEST project points bench_dir at it too). The histogram itself
(draw_png below) is a local reimplementation of Cal's own -- same "cal
plot" visual convention (horizontal-bar, BINS=10 for this 10-DUT dataset,
blue dashed spec lines) -- needed only because the title-fitting/placement
had to change (see draw_png's docstring: Cal's own version sizes/centers
the title against the AXES, which clips on our up-to-8-column slides; ours
sizes/centers against the full FIGURE instead, per user request
2026-08-15). No Cal/BT_TX file is modified -- only imported.

Y-axis range is shared across every column in one group/slide
(group_value_range), not per-item -- per user request 2026-08-15 to add a
"span" ruler (axis min/max shown once per Bench/ATE row, see
build_ppt_fwbt_trimtest_revision1.py's _add_ruler, matching BT_TX/
build_ppt_pwr_evm_revision2.py's own convention) -- a ruler is only
meaningful if every chart in that row is actually drawn to the same scale.

Grouping (2G-left / 5G-right by band) and slide-column ordering are reused
verbatim from the already-approved draft, build_ppt_fwbt_trimtest_draft.py
(base_key, clean_title, BAND_ORDER, _columns_for_group) -- see that
module's docstring for the exact rule (4 prefix variants, BAND<n>
normalization, 41 "common" items).

Image sizing follows the fixed-height/variable-width convention from
WIFI's make_wl_tx_pwr_spec_distribution.py's figure_size_for(n_cols) (per
user instruction to reuse WiFi's ratio rule): a PNG is drawn at exactly the
aspect of the PPT cell it will be stretched into, so add_picture's resize
is always a pure uniform scale, never a distorting stretch. n_cols here is
"how many band-columns share this item's eventual slide" (up to 8: 2G +
5G BAND1..7), the same quantity Cal's own figure_size_for(n_cols) keys off.
"""
from __future__ import annotations

import argparse
import re
import sys
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import openpyxl

_THIS_DIR = Path(__file__).resolve().parent  # Cal/ itself (2026-08-19: moved
# in from BT_TX_TRIMTEST/ per user request -- these items are Cal's own
# Bench/ATE data, so keeping this script in a separate top-level folder was
# just confusing; the project's own identity/output paths, still
# "BT_TX_TRIMTEST" below, are unaffected by where the source file lives)
_BASE_DIR = _THIS_DIR.parent
sys.path.insert(0, str(_BASE_DIR))
sys.path.insert(0, str(_THIS_DIR))

import paths as trimtest_paths
import make_cal_distribution as cal  # Bench/ATE readers + draw_png (reused, not modified)
from build_ppt_fwbt_trimtest_draft import (
    base_key, clean_title, BAND_ORDER, _columns_for_group,
)

PATHS = trimtest_paths.get_paths("Cal")  # 2026-08-19: was "BT_TX_TRIMTEST" --
# unified onto Cal's own project identity per user request, so report_dir/
# result_png_dir now nest under report/Cal/ instead of a separate
# report/BT_TX_TRIMTEST/ tree. No collision with Cal's own PNG/report output
# since bench_png_dir()/ate_png_dir() below use their own "trimtest_
# distribution"/"BT_TX_TRIMTEST_ATE" subfolder names, distinct from Cal's
# "cal_distribution"/"Cal_ATE".
XLSX_PATH = _BASE_DIR / "FWBT_TRIMTEST_DP6P1P3_renamed.xlsx"

# -- PPT grid geometry for the REAL deck (build_ppt_fwbt_trimtest_revision1.py
# mirrors these exactly) -- narrower margins/gaps than the draft since a
# slide can carry up to 8 band-columns (16 Bench/ATE sub-columns) at once,
# same tight-grid precedent as Cal's own geometry constants. --
SLIDE_WIDTH_IN = 13.338542213473316
LEFT_MARGIN = 0.05
RIGHT_MARGIN = 0.05
CELL_GAP_H = 0.03
LABEL_WIDTH = 0.9
GRID_WIDTH = SLIDE_WIDTH_IN - LEFT_MARGIN - RIGHT_MARGIN - LABEL_WIDTH

TABLE_TOP = 1.05
HEADER_ROW_H = 0.55
TABLE_ROW_H = 0.2
TABLE_N_DATA_ROWS = 5  # Bench/ATE subheader + Mean+-Std + Min/Max + USL + LSL
TABLE_PLOT_GAP = 0.35
IMAGE_GAP = 0.15
CONTENT_AREA_TOP = 1.5908508311461067
CONTENT_AREA_BOTTOM = CONTENT_AREA_TOP + 5.240927384076991


def figure_size_for(n_cols: int) -> tuple:
    """-> (cell_width, image_h) in inches, the exact size
    build_ppt_fwbt_trimtest_revision1.py will place this PNG at."""
    cell_width = GRID_WIDTH / n_cols if n_cols == 1 else (GRID_WIDTH - (n_cols - 1) * CELL_GAP_H) / n_cols
    table_h = HEADER_ROW_H + TABLE_N_DATA_ROWS * TABLE_ROW_H
    plots_top = TABLE_TOP + table_h + TABLE_PLOT_GAP
    image_h = (CONTENT_AREA_BOTTOM - plots_top - IMAGE_GAP) / 2
    return (cell_width, image_h)


DROP_TOKENS = {"X", "ATE", "NVVS"}


def clean(name: str) -> str:
    """Same transform as the earlier item-name cleanup this session (kept
    here, not imported, since it belongs to that one-off renaming step, not
    to the draft/report pipeline modules)."""
    parts = name.split("_")
    if len(parts) >= 2 and parts[-2] == "NV" and parts[-1] == "VXXX":
        parts = parts[:-2]
    parts = [p for p in parts if p not in DROP_TOKENS]
    return "_".join(parts)


def load_item_specs():
    """-> {cleaned_name: {"usl": float|None, "lsl": float|None, "unit": str}}"""
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb["Test Items"]
    specs = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        name, usl, lsl, unit = row[0], row[1], row[2], row[3]
        if not name:
            continue
        specs[name] = {"usl": usl, "lsl": lsl, "unit": unit or ""}
    return specs


def build_raw_name_map(bench_items: dict, cleaned_names: set) -> dict:
    """-> {cleaned_name: raw_name}, by reverse-applying clean() to every raw
    Bench-log item name and keeping the ones that land on a name we want.
    Verified 2026-08-15: collision-free, covers all 2737 items."""
    rev = {}
    for raw in bench_items:
        c = clean(raw)
        if c in cleaned_names:
            rev[c] = raw
    return rev


IQR_MULT = 5.0  # same threshold BT_TX/make_bt_tx_pwr_revision2.py's find_combo_outliers uses


def find_outliers(values_with_id):
    """IQR rule (Q1-IQR_MULT*IQR .. Q3+IQR_MULT*IQR) on one item's own
    values -- reused convention from BT_TX/make_bt_tx_pwr_revision2.py's
    find_combo_outliers (same IQR_MULT, same "IQR<=0 -> keep everything"
    guard for constant data), applied here per-item per-side (Bench, ATE
    computed independently) instead of per-combo.

    values_with_id: [(value, id), ...] -- id is whatever the caller wants
    shown in the excluded-outlier note (a real DUT hex id for Bench, a
    1-based position index for ATE -- see build_jobs, where ATE's own
    per-DUT identity isn't available from cal.read_ate's plain value list).

    -> (kept, outliers), each [(value, id), ...]."""
    if len(values_with_id) < 4:  # too few points for a meaningful IQR
        return list(values_with_id), []
    vals = np.array([v for v, _id in values_with_id])
    q1, q3 = np.percentile(vals, [25, 75])
    iqr = q3 - q1
    if iqr <= 0:
        return list(values_with_id), []
    lower = q1 - IQR_MULT * iqr
    upper = q3 + IQR_MULT * iqr
    kept = [(v, i) for v, i in values_with_id if lower <= v <= upper]
    outliers = [(v, i) for v, i in values_with_id if v < lower or v > upper]
    return kept, outliers


EXCLUDED_NOTE_MAX_LINES = 5  # cap so the note box never grows tall enough
# to span into the next collision-check bin below/above the one it's
# centered on -- see _outlier_note_slot's docstring.


def excluded_text(outliers) -> str:
    """Same wording/format as BT_TX/make_bt_tx_pwr_revision2.py's
    excluded_text (kept here as its own copy, not imported, since the
    numeric formatting differs slightly -- BT_TX's dBm-scale values use
    .3f; this project's values span much larger magnitudes, e.g. the
    ~30000-code A4P items, so :g reads better than a fixed 3 decimals).
    Truncated to EXCLUDED_NOTE_MAX_LINES real entries -- an item with many
    outliers used to produce a note tall enough to swallow whatever empty
    region it was centered in (see _outlier_note_slot), turning the
    collision-avoidance moot."""
    if not outliers:
        return ""
    shown = outliers[:EXCLUDED_NOTE_MAX_LINES]
    lines = [f"{v:g} ({i})" for v, i in shown]
    if len(outliers) > EXCLUDED_NOTE_MAX_LINES:
        lines.append(f"(+{len(outliers) - EXCLUDED_NOTE_MAX_LINES} more)")
    return "Excluded outlier(s):\n" + "\n".join(lines)


def _outlier_note_slot(counts, edges, max_count, vmin, vmax, spec):
    """Pick a (y_data, ha, va) placement for the outlier note that avoids
    both the bars AND the USL/LSL dashed lines -- a fixed top/bottom
    corner (tried first) collides constantly here because value_range
    always pads FROM the extreme values (which usually include usl/lsl
    themselves, see group_value_range), so a spec line sits close to
    vmin/vmax on nearly every chart; a corner heuristic can never dodge
    that. Instead: score every histogram bin by (how full it is) + (how
    close a spec line sits to it), pick the lowest-scoring bin, and center
    the note there -- if that bin is genuinely near-empty, the note has
    the FULL x-width of the axes to sit in, not just one corner's sliver.
    """
    lsl, usl = spec if spec else (None, None)
    span = vmax - vmin or 1.0
    spec_zone = span * 0.08  # a spec line "claims" this much of the axis near it
    best_i, best_score = 0, None
    for i in range(len(counts)):
        lo, hi = edges[i], edges[i + 1]
        score = counts[i] / max(max_count, 1)
        for sv in (lsl, usl):
            if sv is not None and lo - spec_zone <= sv <= hi + spec_zone:
                score += 2.0
        if best_score is None or score < best_score:
            best_score, best_i = score, i
    lo, hi = edges[best_i], edges[best_i + 1]
    return (lo + hi) / 2


def build_jobs(specs: dict, bench_items: dict, ate_items: dict):
    """-> [job, ...], one per Test Item that has Bench data. job keys:
    name (raw), cleaned_name, group_key, col_label, usl, lsl, unit,
    bench_values, ate_values (or None) -- both OUTLIER-EXCLUDED (see
    find_outliers) so histograms, table Mean/Std/Min/Max, and the shared
    group_value_range axis all agree on the same "real" data; the excluded
    points themselves live in bench_outliers/ate_outliers for draw_png's
    red annotation box, per user request 2026-08-21 (matching BT_TX's own
    "Excluded outlier(s)" convention)."""
    name_map = build_raw_name_map(bench_items, set(specs))
    jobs = []
    n_no_bench = n_no_ate = 0
    for cleaned_name, spec in specs.items():
        raw = name_map.get(cleaned_name)
        if raw is None:
            n_no_bench += 1
            continue
        bench_rec = bench_items[raw]
        ate_rec = ate_items.get(raw)
        if ate_rec is None:
            n_no_ate += 1
        side, group_key, band_label = base_key(cleaned_name)
        if side == "2G":
            col_label = "2G\n(BAND0)"
        elif side == "5G":
            col_label = f"5G\n({band_label})" if band_label else "5G"
        else:
            col_label = "--"

        bench_kept, bench_outliers = find_outliers(
            [(v, dut) for v, dut in bench_rec["values"]])
        if ate_rec is not None:
            # ATE has no per-value DUT identity by this point (cal.read_ate
            # returns a plain value list, see find_outliers' docstring) --
            # use a 1-based position index as the id shown in the note.
            ate_kept, ate_outliers = find_outliers(
                [(v, i) for i, v in enumerate(ate_rec["values"], start=1)])
        else:
            ate_kept, ate_outliers = [], []

        jobs.append({
            "name": raw,
            "cleaned_name": cleaned_name,
            "side": side,
            "group_key": group_key,
            "col_label": col_label,
            "usl": spec["usl"],
            "lsl": spec["lsl"],
            "unit": spec["unit"],
            "bench_values": [v for v, _id in bench_kept],
            "bench_outliers": bench_outliers,
            "ate_values": [v for v, _id in ate_kept] if ate_rec is not None else None,
            "ate_outliers": ate_outliers,
        })
    print(f"  {len(jobs)}/{len(specs)} items matched to Bench data "
          f"({n_no_bench} no Bench match, {len(jobs) - n_no_ate}/{len(jobs)} have ATE)")
    n_bench_outliers = sum(len(j["bench_outliers"]) for j in jobs)
    n_ate_outliers = sum(len(j["ate_outliers"]) for j in jobs)
    if n_bench_outliers or n_ate_outliers:
        print(f"  outliers excluded from plots: {n_bench_outliers} Bench, {n_ate_outliers} ATE")
    return jobs


def build_groups(jobs):
    """-> {group_key: {"title": str, "jobs": [job, ...] in column order}}
    Reuses the draft's base_key/_columns_for_group ordering (2G, then 5G
    BAND1..7, then a lone common column) by re-deriving the same {"2G":,
    "5G": {band: }, "common":} shape from our job list."""
    raw_groups = {}
    for job in jobs:
        g = raw_groups.setdefault(job["group_key"], {"2G": None, "5G": {}, "common": None})
        if job["side"] == "2G":
            g["2G"] = job
        elif job["side"] == "5G":
            _, _, band_label = base_key(job["cleaned_name"])
            g["5G"][band_label] = job
        else:
            g["common"] = job

    groups = {}
    for key, g in raw_groups.items():
        ordered_jobs = []
        if g["2G"]:
            ordered_jobs.append(g["2G"])
        if None in g["5G"]:
            ordered_jobs.append(g["5G"][None])
        for band in BAND_ORDER[1:]:
            if band in g["5G"]:
                ordered_jobs.append(g["5G"][band])
        if g["common"]:
            ordered_jobs.append(g["common"])
        title = clean_title(g["common"]["cleaned_name"]) if g["common"] else clean_title(key)
        value_range = group_value_range(ordered_jobs)
        for job in ordered_jobs:
            job["value_range"] = value_range
        groups[key] = {"title": title, "jobs": ordered_jobs, "value_range": value_range}
    return groups


def group_value_range(jobs):
    """ONE shared Y-axis range for every column on this group's eventual
    slide -- per user request 2026-08-15 to add a "span" ruler (the axis
    min/max shown once per Bench/ATE row, see BT_TX/build_ppt_pwr_evm_
    revision2.py's _add_ruler, the convention the user pointed to). A ruler
    is only meaningful if every chart on the slide is actually drawn on
    that SAME scale -- unlike the item-by-item _shared_range below (each
    column auto-ranging to its own data+spec), so this pools every column's
    bench+ate+usl+lsl values across the WHOLE group, not just one column."""
    pool = []
    for job in jobs:
        pool += job["bench_values"]
        if job["ate_values"]:
            pool += job["ate_values"]
        for lim in (job["usl"], job["lsl"]):
            if lim is not None:
                pool.append(lim)
    if not pool:
        return (0.0, 1.0)
    lo, hi = min(pool), max(pool)
    if lo == hi:
        pad = abs(lo) * 0.1 or 1.0
        return (lo - pad, hi + pad)
    pad = (hi - lo) * 0.1
    return (lo - pad, hi + pad)


def bench_png_dir() -> Path:
    return PATHS.result_png_dir / "trimtest_distribution"


def ate_png_dir() -> Path:
    return PATHS.result_png_dir.parent / "BT_TX_TRIMTEST_ATE" / "trimtest_distribution"


def _fit_title_fontsize(fig, ax, text, max_width_frac=0.98, start=14, min_size=7):
    """Same search Cal's own _fit_title_fontsize does, but measured against
    the FULL FIGURE canvas width (fig.get_window_extent), not just the
    axes' width -- per user request 2026-08-15 ("빈 가로 여백을 최대한
    사용해서 plot 제목이 안짤리게") to use the margin the y-axis tick
    labels otherwise leave unused. Paired with _center_title_on_figure
    below, which also re-centers the title on the figure (not the axes) --
    sizing against the wider figure width only helps if the title is
    actually allowed to use both sides of that width, not just centered on
    the (narrower, off-center) axes box."""
    renderer = fig.canvas.get_renderer()
    max_width_px = fig.get_window_extent(renderer=renderer).width * max_width_frac
    for size in range(start, min_size - 1, -1):
        t = fig.text(0, 0, text, fontsize=size, fontweight="bold")
        width = t.get_window_extent(renderer=renderer).width
        t.remove()
        if width <= max_width_px:
            return size
    return min_size


def _center_title_on_figure(ax):
    """Shift the axes title's x so it centers on the FIGURE's horizontal
    center instead of the axes' own center. The y-axis tick-label margin
    makes the axes narrower than, and off-center within, the figure -- on
    an 8-column slide (cell ~1.5in wide) centering on the axes wasted that
    margin on the left while starving the right, which is exactly what
    clipped the title in BT2G_A4P_BAND0_EXPAN.png before this fix (rendered
    as "...BAND0_EXPA", cut off). Called AFTER the final fig.tight_layout()
    so the axes position it reads is the one actually used at save time."""
    pos = ax.get_position()
    if pos.width <= 0:
        return
    x_center = (0.5 - pos.x0) / pos.width
    ax.title.set_x(x_center)


def draw_png(values, title, ylabel, png_path, value_range=None, spec=None, figsize=None, outliers=None) -> None:
    """Cal's own horizontal-bar histogram (BINS=10, spec dashed lines),
    reimplemented here rather than imported only because the title needs
    figure-centered sizing/placement (_fit_title_fontsize/
    _center_title_on_figure above) -- everything else (bar draw, grid, tick
    styling, spec line placement) is the same "cal plot" convention,
    unchanged on purpose.

    outliers (2026-08-21, per user request): [(value, id), ...] already
    excluded from `values` by build_jobs' find_outliers -- drawn as a red
    "Excluded outlier(s)" note (BT_TX/make_bt_tx_pwr_revision2.py's own
    convention, reused) instead of being histogrammed, so one extreme DUT
    reading no longer drags the whole axis/USL-LSL-relative scale out."""
    if not values:
        return
    BINS = cal.BINS
    vmin_data, vmax_data = min(values), max(values)
    if value_range is None:
        pad = (vmax_data - vmin_data) * 0.1 or (abs(vmax_data) * 0.1 or 1.0)
        value_range = (vmin_data - pad, vmax_data + pad)
    vmin, vmax = value_range
    fs = cal._font_scale(figsize)

    counts, edges = np.histogram(values, bins=BINS, range=value_range)
    fig, ax = plt.subplots(figsize=figsize or cal._REF_FIGSIZE)
    max_count = max(counts.max(), 1)
    for i in range(BINS):
        lo, hi = edges[i], edges[i + 1]
        bin_h = hi - lo
        c = counts[i]
        ax.barh(lo, c, height=bin_h * 0.9, align="edge", color="#4472C4",
                edgecolor="black", linewidth=0.6)
        if c > 0:
            ax.text(c + max_count * 0.01, lo + bin_h / 2, f"{int(c)}", va="center", ha="left", fontsize=9 * fs)

    ax.set_ylim(vmin, vmax)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=10))
    ax.set_xlim(0, max_count * 1.15)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True))
    plt.setp(ax.get_yticklabels(), fontsize=13 * fs, fontweight="bold")
    plt.setp(ax.get_xticklabels(), fontsize=13 * fs, fontweight="bold")
    ax.set_ylabel(ylabel, fontsize=13 * fs, fontweight="bold")
    ax.set_xlabel("Count", fontsize=13 * fs, fontweight="bold")
    ax.grid(True, axis="both", color="0.85", linewidth=0.6)
    ax.set_axisbelow(True)
    fig.tight_layout()

    if title:
        wrapped_title = cal.wrap_label(title)
        ax.set_title(wrapped_title, fontsize=_fit_title_fontsize(fig, ax, wrapped_title), fontweight="bold")

    if spec is not None:
        lsl, usl = spec
        pad = (vmax - vmin) * 0.035
        for spec_val, spec_label in ((lsl, "LSL"), (usl, "USL")):
            if spec_val is None or not (vmin <= spec_val <= vmax):
                continue
            ax.axhline(y=spec_val, color="blue", linestyle="--", linewidth=2.0, zorder=5)
            above = spec_val >= (vmin + vmax) / 2
            text_y = spec_val + pad if above else spec_val - pad
            ax.text(ax.get_xlim()[1] * 0.99, text_y, f"{spec_label} {spec_val:g}",
                    color="blue", fontsize=9 * fs, fontweight="bold", ha="right",
                    va="bottom", clip_on=True)

    note = excluded_text(outliers)
    if note:
        # Centered in the histogram's own lowest-occupancy bin (see
        # _outlier_note_slot) rather than a fixed corner -- 2026-08-21 fix:
        # a fixed top/bottom corner collided with the USL/LSL dashed line
        # almost every time (value_range pads FROM the extreme values,
        # which usually include usl/lsl, so a spec line sits close to
        # vmin/vmax on nearly every chart -- confirmed via a real rendered
        # PNG where the note text was fully illegible under the line).
        note_y = _outlier_note_slot(counts, edges, max_count, vmin, vmax, spec)
        ax.text(max_count * 0.55, note_y, note, ha="center", va="center",
                fontsize=9 * fs, fontweight="bold", color="#B00000",
                bbox=dict(boxstyle="round", facecolor="#FFF2F2", edgecolor="#B00000"))

    fig.tight_layout()
    if title:
        _center_title_on_figure(ax)  # after final layout -- see its docstring
    png_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(png_path, dpi=300)
    plt.close(fig)


def _draw_one(job):
    """Chart title is the item's cleaned name -- restored 2026-08-15 (was
    blank) now that draw_png's title fitting/centering above uses the full
    figure width instead of just the axes width. value_range is the
    GROUP-shared range (job["value_range"], set in build_groups) so every
    column on this job's eventual slide is drawn to the same scale -- the
    "span" ruler build_ppt_fwbt_trimtest_revision1.py draws per Bench/ATE
    row only makes sense if every chart in that row actually uses it."""
    name = job["cleaned_name"]
    figsize = figure_size_for(job["n_cols"])
    value_range = job["value_range"]
    ylabel = job["unit"] or "Value"

    draw_png(job["bench_values"], name, ylabel, job["bench_png"], value_range=value_range,
             spec=(job["lsl"], job["usl"]), figsize=figsize, outliers=job.get("bench_outliers"))

    if job["ate_values"] is not None:
        draw_png(job["ate_values"], name, ylabel, job["ate_png"], value_range=value_range,
                 spec=(job["lsl"], job["usl"]), figsize=figsize, outliers=job.get("ate_outliers"))
    return name, job["ate_values"] is not None


def parse_args():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--limit", type=int, default=0, help="only draw the first N groups")
    p.add_argument("--count-only", action="store_true", help="print the match audit only")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    print("Loading Test Item names + USL/LSL/Unit spec...")
    specs = load_item_specs()
    print(f"  {len(specs)} items")

    print("Reading Bench_Cal CSVs...")
    bench_items = cal.read_bench(PATHS.bench_dir)
    print(f"  {len(bench_items)} test item(s) found in Bench log")

    print(f"Parsing ATE log ({PATHS.ate_log_path.name})...")
    ate_items = cal.read_ate(PATHS.ate_log_path)

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
