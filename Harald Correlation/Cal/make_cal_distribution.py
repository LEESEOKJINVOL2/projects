"""Cal (calibration) distribution histograms -- Bench + ATE, matched by exact
TEST_NAME string.

Bench_40pcs/Cal holds 10 DUT logs in a completely different layout from
every other project's wide-CSV-per-DUT format: one row per test item
("TNAME:,<name>,<RESULT_PB>,<LLIM>,<LLIM_TYPE>,<ULIM_TYPE>,<ULIM>,<P/F>"),
same ~6830-item set/order across all 10 DUTs (confirmed 2026-08-14). The
shared ate_log_path CSV carries the exact same TEST_NAME strings for these
items (not a BTTX_/BTRX_-style encoded name) -- match is a plain exact
string compare after stripping whitespace (item names carry leading
alignment-padding spaces in column B, no trailing spaces seen), per user
instruction 2026-08-14: match on the trimmed text only.

Per user instruction: an item is only plotted if the Bench log itself
carries at least one limit (LLIM and/or ULIM non-blank) -- items with
neither are skipped entirely (3847/6830 qualify, 3846 of those have a
matching ATE row, confirmed 2026-08-14).

For every qualifying item, draws two side-by-side-comparable PNGs (same
horizontal-histogram visual convention as every other project; BINS=10
for a 10-DUT dataset, same precedent as BT_RX's bt_rx_ate_lookup.draw_png):
  * Bench: this item's 10 DUT RESULT_PB values, Bench's own LLIM/ULIM as
    the LSL/USL spec lines.
  * ATE: the matched item's 10 DUT values from the shared ATE log (only
    drawn when a match exists), using the SAME axis range as its Bench
    sibling for direct comparison, and preferring the ATE row's own
    Upper/Lower Limit for the spec lines (falls back to Bench's limit if
    the ATE row's own is blank).

CLI: --limit N (first N qualifying items only), --count-only (print the
qualify/match audit without drawing anything).
"""
from __future__ import annotations

import argparse
import csv
import io
import re
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as cal_paths

PATHS = cal_paths.get_paths("Cal")

BINS = 10  # 10 DUTs, same precedent as BT_RX's bt_rx_ate_lookup.BINS
DUT_RE = re.compile(r"logfile_([0-9A-Za-z]+)_")

# -- PPT grid geometry, mirrored from build_ppt_cal_distribution.py, so a
# source PNG's aspect matches the actual PPT cell it lands in instead of
# being stretched into a different shape (same distortion class as BT_TX's
# figure_size_for fix, see make_bt_tx_pwr_revision2.py's docstring) --
# real bug found 2026-08-14: every PNG was drawn at a fixed (9,6) regardless
# of how many items shared its eventual slide, so a section's last page
# (often just 1-3 leftover items, since sections rarely divide evenly by
# COLS_PER_SLIDE) got stretched from that fixed aspect into a MUCH wider/
# shorter cell -- non-uniform scaling that visibly smeared axis/tick text.
COLS_PER_SLIDE = 4
_SLIDE_WIDTH_IN = 13.338542213473316
_LEFT_MARGIN = 0.05
_RIGHT_MARGIN = 0.05
_CELL_GAP_H = 0.03
_LABEL_WIDTH = 0.9
_GRID_WIDTH = _SLIDE_WIDTH_IN - _LEFT_MARGIN - _RIGHT_MARGIN - _LABEL_WIDTH

_TABLE_TOP = 1.05
_HEADER_ROW_H = 0.55
_TABLE_ROW_H = 0.2
_TABLE_N_DATA_ROWS = 4  # Bench/ATE subheader + Spec + Mean+-Std + Min/Max
_TABLE_PLOT_GAP = 0.35
_IMAGE_GAP = 0.15
_CONTENT_AREA_TOP = 1.5908508311461067
_CONTENT_AREA_BOTTOM = _CONTENT_AREA_TOP + 5.240927384076991


def figure_size_for(n_cols: int) -> tuple[float, float]:
    """-> (cell_width, image_h) in the SAME inches build_ppt_cal_distribution.py
    will place the PNG at -- feeding matplotlib this exact size means
    add_picture's stretch is uniform (a pure dpi scale), never a shape-
    distorting non-uniform one."""
    cell_width = _GRID_WIDTH / n_cols if n_cols == 1 else (_GRID_WIDTH - (n_cols - 1) * _CELL_GAP_H) / n_cols
    table_h = _HEADER_ROW_H + _TABLE_N_DATA_ROWS * _TABLE_ROW_H
    plots_top = _TABLE_TOP + table_h + _TABLE_PLOT_GAP
    image_h = (_CONTENT_AREA_BOTTOM - plots_top - _IMAGE_GAP) / 2
    return (cell_width, image_h)


def build_pages(jobs, cols_per_slide: int = COLS_PER_SLIDE):
    """Group qualifying items by their Bench-log PATNAME section (first-seen
    order), then paginate each section's items into chunks of at most
    cols_per_slide -- the exact grouping build_ppt_cal_distribution.py turns
    into slides. Both make_cal_distribution.py (to size each PNG for its
    eventual cell) and build_ppt_cal_distribution.py (to lay out the actual
    slides) call this SAME function so the two never drift apart.

    -> [{"section": str, "page": int (1-based), "n_pages": int, "jobs": [...]}, ...]
    """
    by_section: dict[str, list] = {}
    for job in jobs:
        by_section.setdefault(job["section"] or "(no section)", []).append(job)

    pages = []
    for section, section_jobs in by_section.items():
        n_pages = (len(section_jobs) + cols_per_slide - 1) // cols_per_slide
        for page in range(n_pages):
            chunk = section_jobs[page * cols_per_slide:(page + 1) * cols_per_slide]
            pages.append({"section": section, "page": page + 1, "n_pages": n_pages, "jobs": chunk})
    return pages


def to_float(text: str):
    text = (text or "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def dut_id_of(path: Path) -> str:
    m = DUT_RE.search(path.name)
    return m.group(1) if m else path.stem


def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:\[\]<>|"]', "_", name)


def _read_bench_csv_rows(path: Path) -> list:
    """Every Bench log CSV verified so far has been UTF-8 (with or without
    a BOM), but two real-world surprises have shown up on colleagues'
    machines: one pull saved in the local Windows codepage instead (cp949,
    the default on a Korean-locale PC), another actually saved as UTF-16
    (common from Windows-native tools/exports). utf-8-sig can't decode
    either, and used to crash the entire run with a raw UnicodeDecodeError
    partway through whichever DUT file hit it first.

    UTF-16 needs checking FIRST and explicitly, not just added to the
    fallback chain: cp949 doesn't raise on a UTF-16 byte stream, it
    "succeeds" by decoding it into garbage riddled with literal NUL
    characters (every ASCII byte in UTF-16 is followed/preceded by a 0x00
    byte) -- which Python's csv module then refuses outright ("line
    contains NUL"), a confusing second failure past the first fix. UTF-16
    is detected up front from its BOM (\\xff\\xfe or \\xfe\\xff) so it's
    decoded correctly instead of accidentally "surviving" cp949 as noise.

    Falls back to latin-1 (which cannot raise -- every byte 0x00-0xFF maps
    to *something* in it) for anything still unrecognized, and strips any
    stray NUL that slips through regardless of path, as a last-resort
    backstop -- a handful of corrupted characters shouldn't abort an
    otherwise-readable file."""
    raw = path.read_bytes()
    candidates = ("utf-16",) if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else ("utf-8-sig", "cp949")
    text = None
    for enc in candidates:
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("latin-1")  # never raises
    text = text.replace("\x00", "")
    return list(csv.reader(io.StringIO(text, newline="")))


def read_bench(bench_dir: Path):
    """-> {item_name: {"section": str, "llim": float|None, "ulim": float|None,
                        "values": [(value, dut), ...]}}

    Only "TNAME:" rows are test items. "PATNAME:" rows are section headers
    (column B names the group the following TNAME rows belong to, e.g.
    "CLKMEASX_XXXXXXXX") -- kept per item so build_ppt_cal_distribution.py
    can group/paginate slides by section instead of one slide per item.
    Column B (item name) is stripped before use -- some rows carry leading
    alignment padding. llim/ulim/section are taken from whichever DUT file
    first has a non-blank value for that item (identical across all 10 DUTs
    in every case checked 2026-08-14, but this stays defensive rather than
    assuming file #1 always has it)."""
    # 2026-08-21: rglob, not glob -- bench_dir is now a GUI-picked path that
    # may be either layout Harald's raw pulls come in: flat (Bench_40pcs/
    # Harald_260820/*.csv, 40 files directly) or vendor-nested (Bench_40pcs/
    # Harald/{Skyworks,Sony}/{DUT serial}/*.csv). dut_id_of() already reads
    # the DUT identity out of the FILENAME itself, never the folder path, so
    # finding files recursively is enough to support both shapes with no
    # vendor-awareness needed here -- matches the existing "pool every DUT
    # together, no vendor split" convention this report already uses. A
    # flat folder has no subdirectories to recurse into, so this returns
    # the exact same files a flat pull's glob("*.csv") already did.
    #
    # 2026-08-21: skip "._"-prefixed files -- macOS AppleDouble resource-
    # fork metadata, one per real file, left behind after a vendor-nested
    # Bench pull got zipped/transferred through a Mac. Binary, not a real
    # Bench log, but it matches "*.csv" (same base name as its real
    # counterpart) -- this is what turned into the "line contains NUL" /
    # UnicodeDecodeError a colleague hit: the parser choked trying to read
    # one of these as if it were real test data. Same skip already used in
    # add_harald_260820_raw_values_sheet.py's build_hex_to_serial() for the
    # exact same reason.
    items: dict[str, dict] = {}
    files = sorted(p for p in bench_dir.rglob("*.csv") if not p.name.startswith("."))
    for path in files:
        dut = dut_id_of(path)
        section = None
        rows_iter = iter(_read_bench_csv_rows(path))
        next(rows_iter)
        for row in rows_iter:
            if not row:
                continue
            tag = row[0].strip()
            if tag == "PATNAME:":
                section = (row[1] or "").strip()
                continue
            if tag != "TNAME:":
                continue
            name = (row[1] or "").strip()
            value = to_float(row[2] if len(row) > 2 else "")
            if value is None:
                continue
            llim = to_float(row[3] if len(row) > 3 else "")
            ulim = to_float(row[6] if len(row) > 6 else "")
            rec = items.setdefault(name, {"section": section, "llim": None, "ulim": None, "values": []})
            if rec["section"] is None and section is not None:
                rec["section"] = section
            if rec["llim"] is None and llim is not None:
                rec["llim"] = llim
            if rec["ulim"] is None and ulim is not None:
                rec["ulim"] = ulim
            rec["values"].append((value, dut))
        print(f"  read {path.name}")
    return items


def _values_from_row(row):
    # Same no-read sentinel handling as the other projects' ATE parsers --
    # the ATE log mixes in 9.91e+37 for a no-read cell.
    out = []
    for v in row[3:]:
        v = v.strip()
        if v in ("", "NA"):
            continue
        try:
            f = float(v)
        except ValueError:
            continue
        if abs(f) < 1e6:
            out.append(f)
    return out


def read_ate(ate_csv: Path):
    """-> {item_name: {"llim": float|None, "ulim": float|None, "values": [float, ...]}}

    Generic pass, no item-name-prefix filter (unlike every other project's
    BTTX_/BTRX_-style parsing) -- Cal's items span several unrelated
    prefixes (ARF_, BOOT_, BT_, DC_, LBID, ...), so the only filter that
    makes sense is "does this row have >=1 real numeric DUT value". Header
    is SerialNumber,Upper Limit,Lower Limit,<DUT...>."""
    idx: dict[str, dict] = {}
    rows_iter = iter(_read_bench_csv_rows(ate_csv))  # same encoding fallback as read_bench
    next(rows_iter)
    for row in rows_iter:
        if not row:
            continue
        name = row[0].strip()
        values = _values_from_row(row)
        if not values:
            continue
        ulim = to_float(row[1] if len(row) > 1 else "")
        llim = to_float(row[2] if len(row) > 2 else "")
        idx[name] = {"llim": llim, "ulim": ulim, "values": values}
    return idx


def read_ate_auto(path: Path):
    """Same return shape as read_ate() -- {item_name: {"llim", "ulim",
    "values"}} -- but accepts EITHER a single ATE csv file (every project's
    existing model) or a FOLDER of them.

    2026-08-21: a 100-DUT Harald pull split its ATE side across 4 config-
    specific files instead of one covering every DUT (SKYWORKS(DSKJ)/
    SKYWORKS(SSNJ)/SONY(DNNJ)/SONY(SNKJ), 25 DUT columns each) -- verified
    all 4 carry the exact same 11591-item name set, just disjoint DUT
    columns, so merging is a plain per-item concatenation, the same
    across-file merge convention read_bench() already uses across DUT
    files: values lists concatenated, llim/ulim kept from whichever file
    first has them (identical across files in every case checked, same as
    read_bench's own llim/ulim note). "._"-prefixed files (macOS
    AppleDouble junk, see read_bench's own note) are skipped here too."""
    if not path.is_dir():
        return read_ate(path)
    merged: dict[str, dict] = {}
    for f in sorted(p for p in path.glob("*.csv") if not p.name.startswith(".")):
        for name, rec in read_ate(f).items():
            dest = merged.setdefault(name, {"llim": None, "ulim": None, "values": []})
            if dest["llim"] is None and rec["llim"] is not None:
                dest["llim"] = rec["llim"]
            if dest["ulim"] is None and rec["ulim"] is not None:
                dest["ulim"] = rec["ulim"]
            dest["values"].extend(rec["values"])
    return merged


def wrap_label(text: str, max_len: int = 22) -> str:
    """Greedy underscore-boundary wrap so a ~40-60 char item name stays
    legible instead of one unbroken line -- real bug found 2026-08-14:
    figure_size_for(4)'s narrow (~3.06in wide) cell can't fit even a 7pt
    single-line title for these names (title text was overflowing past both
    edges of its own PNG canvas, not just the PPT cell), so the title (and
    the table header in build_ppt_cal_distribution.py, which calls this same
    function) needs to be pre-wrapped across lines rather than shrunk
    indefinitely. get_window_extent() on a "\\n"-joined matplotlib Text
    already measures the widest LINE, not the full string, so
    _fit_title_fontsize needs no change -- it just needs to receive
    already-wrapped text."""
    tokens = text.split("_")
    lines, cur = [], ""
    for tok in tokens:
        piece = tok if not cur else f"{cur}_{tok}"
        if len(piece) > max_len and cur:
            lines.append(cur)
            cur = tok
        else:
            cur = piece
    if cur:
        lines.append(cur)
    return "\n".join(lines)


def _fit_title_fontsize(fig, ax, text, max_width_frac=0.98, start=14, min_size=7):
    renderer = fig.canvas.get_renderer()
    max_width_px = ax.get_window_extent(renderer=renderer).width * max_width_frac
    for size in range(start, min_size - 1, -1):
        t = ax.text(0, 0, text, fontsize=size, fontweight="bold")
        width = t.get_window_extent(renderer=renderer).width
        t.remove()
        if width <= max_width_px:
            return size
    return min_size


_REF_FIGSIZE = (9, 6)  # the shape every font size below was tuned against


def _font_scale(figsize) -> float:
    """Every fontsize below was tuned for a 6in-tall reference figure. Unlike
    BT_TX's cell width (which genuinely varies with n_cols, see
    figure_size_for), Cal's image_h is the SAME ~1.97in for every n_cols --
    only cell WIDTH changes -- so scaling font by width here would be
    backwards: width is already handled correctly by figure_size_for itself
    (matplotlib draws at the exact target width, so no further correction
    needed there). What still needs correcting is that every chart now
    renders at ~1.97in tall instead of the 6in the point sizes below were
    tuned against -- without this, a 13pt tick label that looked right
    stretched down from a 6in figure into a ~1.97in cell (the ORIGINAL
    figsize=(9,6)-always bug's one accidental correctness, purely because
    2.4G/... 4-column pages happened to have a near-1.5 cell aspect) would
    render ~3x too large once drawn natively at 1.97in tall with no
    subsequent stretch to shrink it back down. Same height-based-scale
    pattern as make_bt_tx_pwr_revision2.py's _font_scale, just noting here
    that height is constant across every n_cols rather than genuinely
    variable."""
    if figsize is None:
        return 1.0
    return max(0.28, min(1.0, figsize[1] / _REF_FIGSIZE[1]))


def draw_png(values, title, ylabel, png_path: Path, value_range=None, spec=None, figsize=None) -> None:
    """Horizontal-bar distribution histogram, same visual convention as
    bt_rx_ate_lookup.draw_png (Count on X, value bins on Y, BINS=10 for a
    10-DUT dataset) plus an optional LSL/USL spec line pair -- either side
    may be None since most Cal items only carry one of the two limits.
    `figsize` should be figure_size_for(n_cols) for the PPT cell this PNG
    will be stretched into -- see that function's docstring for why."""
    if not values:
        return
    vmin_data, vmax_data = min(values), max(values)
    if value_range is None:
        pad = (vmax_data - vmin_data) * 0.1 or (abs(vmax_data) * 0.1 or 1.0)
        value_range = (vmin_data - pad, vmax_data + pad)
    vmin, vmax = value_range
    fs = _font_scale(figsize)

    counts, edges = np.histogram(values, bins=BINS, range=value_range)
    fig, ax = plt.subplots(figsize=figsize or _REF_FIGSIZE)
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
    # Wrapped first (see wrap_label's docstring -- a single unbroken ~50-60
    # char title has no font size, down to min_size, that fits a 4-column
    # page's ~3.06in-wide axes). Unscaled size search (start=14/min=7) then
    # adapts to whichever WIDTH figure_size_for(n_cols) actually gave this
    # axes, no additional height-based fs correction needed on top of that.
    wrapped_title = wrap_label(title)
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

    fig.tight_layout()
    png_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(png_path, dpi=300)
    plt.close(fig)


def build_jobs(bench_items, ate_items):
    """Filter to items with >=1 Bench limit, pair with an ATE match by exact
    (stripped) name if one exists."""
    jobs = []
    n_no_ate = 0
    for name, rec in bench_items.items():
        if rec["llim"] is None and rec["ulim"] is None:
            continue
        ate_rec = ate_items.get(name)
        if ate_rec is None:
            n_no_ate += 1
        jobs.append({
            "name": name,
            "section": rec["section"],
            "bench_values": [v for v, _dut in rec["values"]],
            "bench_llim": rec["llim"], "bench_ulim": rec["ulim"],
            "ate_values": ate_rec["values"] if ate_rec else None,
            "ate_llim": ate_rec["llim"] if ate_rec else None,
            "ate_ulim": ate_rec["ulim"] if ate_rec else None,
        })
    return jobs, n_no_ate


def effective_spec(job):
    """(llim, ulim) actually drawn on the ATE chart -- ATE's own limit,
    falling back to Bench's when the ATE row's own is blank. Same rule
    _draw_one uses; factored out so build_ppt_cal_distribution.py's stats
    table can show the identical spec without re-deriving it."""
    llim = job["ate_llim"] if job["ate_llim"] is not None else job["bench_llim"]
    ulim = job["ate_ulim"] if job["ate_ulim"] is not None else job["bench_ulim"]
    return llim, ulim


def bench_png_dir() -> Path:
    return PATHS.result_png_dir / "cal_distribution"


def ate_png_dir() -> Path:
    return PATHS.result_png_dir.parent / "Cal_ATE" / "cal_distribution"


def _shared_range(bench_values, ate_values, llim, ulim):
    pool = list(bench_values)
    if ate_values:
        pool += ate_values
    for lim in (llim, ulim):
        if lim is not None:
            pool.append(lim)
    lo, hi = min(pool), max(pool)
    if lo == hi:
        pad = abs(lo) * 0.1 or 1.0
        return (lo - pad, hi + pad)
    pad = (hi - lo) * 0.1
    return (lo - pad, hi + pad)


def _draw_one(job):
    name = job["name"]
    title = f"Cal_{name}"
    figsize = figure_size_for(job["n_cols"])

    llim, ulim = effective_spec(job)
    value_range = _shared_range(job["bench_values"], job["ate_values"], llim, ulim)

    # ylabel="Value" (generic), not the item name -- the (now wrapped)
    # title already names the item; a second copy of the same ~50-60 char
    # name rotated vertically along the Y axis can't fit inside this
    # image's ~1.97in-tall canvas at any legible font size (real bug found
    # 2026-08-14, same "too long for this canvas" class as the title fix
    # in draw_png/wrap_label -- this one has no width-based fix available
    # since the axis label runs along the constrained HEIGHT, not width).
    draw_png(job["bench_values"], title, "Value", job["bench_png"], value_range=value_range,
             spec=(job["bench_llim"], job["bench_ulim"]), figsize=figsize)

    if job["ate_values"] is not None:
        draw_png(job["ate_values"], "ATE_" + title, "Value", job["ate_png"], value_range=value_range,
                 spec=(llim, ulim), figsize=figsize)
    return name, job["ate_values"] is not None


def parse_args():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--limit", type=int, default=0, help="only draw the first N qualifying items")
    p.add_argument("--count-only", action="store_true", help="print the qualify/match audit only")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    print("Reading Bench_Cal CSVs...")
    bench_items = read_bench(PATHS.bench_dir)
    print(f"  {len(bench_items)} test item(s) found")

    print(f"Parsing ATE log ({PATHS.ate_log_path.name})...")
    ate_items = read_ate(PATHS.ate_log_path)

    jobs, n_no_ate = build_jobs(bench_items, ate_items)
    print(f"\n=== items with >=1 Bench limit: {len(jobs)}/{len(bench_items)} "
          f"(ATE match: {len(jobs) - n_no_ate}, no ATE match: {n_no_ate}) ===")

    if args.count_only:
        return

    if args.limit > 0:
        jobs = jobs[: args.limit]
        print(f"Limiting to first {len(jobs)} item(s)")

    # n_cols per item = how many items will actually share its eventual PPT
    # slide (see build_pages) -- needed so figure_size_for(n_cols) draws each
    # PNG at the exact aspect its slide cell will place it at, instead of a
    # fixed shape that gets non-uniformly stretched (see figure_size_for's
    # and _font_scale's docstrings for the bug this fixes).
    for page in build_pages(jobs):
        for job in page["jobs"]:
            job["n_cols"] = len(page["jobs"])

    bench_dir_out = bench_png_dir()
    ate_dir_out = ate_png_dir()
    bench_dir_out.mkdir(parents=True, exist_ok=True)
    ate_dir_out.mkdir(parents=True, exist_ok=True)

    for job in jobs:
        fname = safe_filename(job["name"])
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
