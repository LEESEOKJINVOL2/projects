"""Centralized path configuration, shared across every project in this
workspace (BT_TX, NB_Tx, UWB, WIFI) -- 2026-07-20 restructure.

Replaces the earlier per-project loaders (bt_tx_paths.py, nb_tx_paths.py):
instead of one config.xml per project, there is now ONE config.xml at the
workspace root with two sections:
  <paths>    -- values shared by every project (spec_workbook, ppt_template,
                report_dir, result_png_dir, result_xlsx_dir, extract_mode,
                the shared ate_log_path FILE read by all four projects).
                report_dir/result_png_dir/result_xlsx_dir each get the
                calling project's name auto-appended (see those properties
                below) -- callers never assemble that path themselves.
  <projects> -- per-project overrides that genuinely differ -- as of
                2026-07-23 this is just bench_dir (a real, different raw-
                data folder per project) -- these do NOT get the project
                name appended again, since the value itself already names
                the project's own bench-data subfolder.

Each project's scripts live one directory below the workspace root (e.g.
BT_TX/make_bt_tx_pwr_revision2.py, NB_Tx/make_nb_tx_pwr_revision1.py), so
they reach this module via a small sys.path shim:

    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    import paths
    PATHS = paths.get_paths("BT_TX")   # or "NB_Tx", "UWB", "WIFI"
"""

from __future__ import annotations

import sys
import xml.etree.ElementTree as ET
from pathlib import Path

_CONFIG_FILENAME = "config.xml"


def _default_search_start() -> Path:
    """2026-07-23: where to start walking up for config.xml, by run mode.

    Un-frozen (`python run_all.py`): Path(__file__).resolve().parent -- this
    file's own real directory (the workspace root), exactly as before.

    Frozen (PyInstaller exe): `__file__` for a bundled module resolves
    inside PyInstaller's temp extraction directory (sys._MEIPASS, e.g.
    C:\\Users\\...\\AppData\\Local\\Temp\\_MEIxxxxx\\paths.py) -- walking up
    from THERE would search the Temp folder tree, never finding the real
    config.xml (which stays external, next to the .exe on disk, per this
    project's "bundle code+runtime only, data stays external" design, see
    project memory). Search from the exe's own directory instead."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _find_config(start: Path) -> Path:
    """Walk up from `start` looking for config.xml -- lets any script find
    it regardless of the working directory it's invoked from, as long as
    it lives somewhere under (or next to) the workspace root (or, when
    frozen, next to the .exe)."""
    for d in (start, *start.parents):
        candidate = d / _CONFIG_FILENAME
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        f"{_CONFIG_FILENAME} not found starting from {start} -- "
        "expected it at the workspace root, above every project folder "
        "(or, when frozen, next to the .exe)."
    )


class Paths:
    """Resolved path config for one project. Every property returns a
    pathlib.Path, resolved relative to config.xml's own folder (the
    workspace root) unless the configured value is already absolute."""

    def __init__(self, project: str, config_path: Path | None = None):
        self.project = project
        self.config_path = config_path or _find_config(_default_search_start())
        self.base_dir = self.config_path.resolve().parent
        root = ET.parse(self.config_path).getroot()

        paths_el = root.find("paths")
        if paths_el is None:
            raise ValueError(f"{self.config_path} has no <paths> section")
        self._raw = {child.tag: (child.text or "").strip() for child in paths_el}

        projects_el = root.find("projects")
        proj_el = projects_el.find(project) if projects_el is not None else None
        if proj_el is None:
            raise ValueError(
                f"{self.config_path} has no <projects><{project}> entry -- "
                f"valid projects: {[c.tag for c in projects_el] if projects_el is not None else []}"
            )
        self._proj_raw = {child.tag: (child.text or "").strip()
                           for child in proj_el if child.tag != "options"}
        # <options> is optional, per-project (currently BT_TX only, ported
        # from its own former standalone bt_tx_paths.py during the
        # 2026-07-21 BT_TX-ATE merge) -- absent for every other project, so
        # every property below has a safe fallback matching the pre-options
        # behavior.
        options_el = proj_el.find("options")
        self._options = {child.tag: (child.text or "").strip()
                          for child in options_el} if options_el is not None else {}

    def _resolve(self, key: str, raw: dict) -> Path:
        """Resolve one config value to a Path, relative to config.xml's folder.

        2026-07-28: an EMPTY tag is treated exactly like a missing one. Before
        this, `<bench_dir></bench_dir>` (or `<ate_log_path/>`, which is easy to
        leave behind while swapping ATE pulls) produced `Path("")`, and
        `base_dir / Path("")` is the WORKSPACE ROOT itself -- not an error. Two
        real consequences, both silent:
          * bench_dir -> the root gets globbed for *.csv, finds nothing, and
            every chart script reports "no data" and skips, so the run
            "succeeds" with an empty deck.
          * ate_log_path -> callers' `is None` guards (uwb_ate_lookup.py,
            make_nb_tx_ate_revision1.py, build_ppt_nb_tx_revision1.py) never
            fire because a real Path came back, and the parser then tries to
            open() a directory.
        Raising KeyError here routes an empty value into the same
        already-handled path as a missing one -- ate_log_path turns it into
        None, and everything else fails loudly with the message below.
        """
        value = raw.get(key, "").strip()
        if not value:
            raise KeyError(
                f"config.xml has no usable <{key}> entry for project "
                f"{self.project!r} (tag missing or empty) -- config: {self.config_path}"
            )
        p = Path(value)
        return p if p.is_absolute() else (self.base_dir / p)

    # -- per-project values (bench_dir differs in both location AND name
    # per project, e.g. Bench_40pcs/BT_Tx vs Bench_40pcs/NB_Tx -- the
    # project name is already baked into the configured value, not
    # appended again) --
    @property
    def bench_dir(self) -> Path:
        return self._resolve("bench_dir", self._proj_raw)

    @property
    def bt_tx_spec_workbook(self) -> Path | None:
        """BT_TX-only override of the shared <paths><spec_workbook> --
        2026-08-14, points at the customer's new
        'Harald_P2_Test_Coverage_v4.4_JSCK.xlsx'. Originally added for Power
        alone ('TX Power Specification' sheet, see bt_tx_power_spec_lookup.py)
        but DEVM/ACP/FreqAcc's own sheets ('BT - BDR & EDR Spec' / 'BT - BLE
        and HDT Spec' / 'BT - HDR HDRP UHDR Spec') turned out to live in this
        SAME new workbook too (same sheet NAMES as the old workbook, but with
        real content differences -- see bt_tx_devm_spec_lookup.py's and
        bt_tx_acp_spec_lookup.py's module docstrings for what changed), so
        this one override now covers all four BT_TX metrics. A per-project
        override rather than changing the shared <spec_workbook>, so UWB/
        WL_TX (which also read that shared value) are unaffected. Lives in
        <projects><BT_TX> (per-project), not <paths> (shared), unlike
        ate_log_path -- this workbook is BT_TX-specific, not read by every
        project the way the ATE log is. Returns None if config.xml has no
        <bt_tx_spec_workbook> entry, so callers fall back to the shared
        spec_workbook (`PATHS.bt_tx_spec_workbook or PATHS.spec_workbook`)."""
        if not self._proj_raw.get("bt_tx_spec_workbook", "").strip():
            return None
        return self._resolve("bt_tx_spec_workbook", self._proj_raw)

    def validate_bench_dir(self) -> str | None:
        """Return a human-readable reason this project's bench data is unusable,
        or None if it looks fine. 2026-07-28.

        Nothing in this workspace used to check that bench_dir exists, and
        `Path.glob()` returns an empty iterator for a missing directory instead
        of raising. The result was a SILENT empty deliverable reported as
        success: with `Bench_40pcs/NB_Tx` absent (its real state here), a
        NB_Tx run produced a deck with a title slide, empty section dividers
        and ZERO charts, a summary CSV with zero data rows, and `NB_Tx: OK`
        from run_all.py. Only make_nb_tx_psd_revision1.py failed loudly, and
        extract_mode="ate" skips that one step, removing even that.

        Deliberately a RETURNED reason rather than a raise, and deliberately
        NOT wired into the `bench_dir` property: every chart script does
        `add_argument("--bench-dir", default=PATHS.bench_dir)`, whose default is
        evaluated eagerly, so a raising property would kill a run even when the
        caller passed a perfectly good `--bench-dir` override. run_all.py calls
        this per project before running any of its steps (see run_project()).
        """
        try:
            d = self.bench_dir
        except KeyError as e:
            return str(e)
        if not d.exists():
            return f"bench_dir does not exist: {d}"
        if not d.is_dir():
            return f"bench_dir is not a directory: {d}"
        if not any(d.glob("*.csv")):
            return f"bench_dir has no *.csv files: {d}"
        return None


    # -- optional per-project <options> --
    @property
    def write_xlsx(self) -> bool:
        raw = self._options.get("write_xlsx", "false").strip().lower()
        return raw in ("1", "true", "yes", "on")

    @property
    def png_workers(self) -> int:
        """Process count for parallel PNG generation (see png_jobs.py).

        2026-07-27: capped at this machine's logical CPU count. config.xml
        carries ONE value (5) that has to be reasonable on every PC this runs
        on -- the report owner's 20-core desktop as well as the i5/i7 laptops
        it also gets run on. Oversubscribing a 2-core machine with 5 drawing
        processes only adds context-switching and ~69MB of resident memory per
        worker for no throughput, so clamp instead of trusting the file. A
        machine with MORE cores than the configured value is left alone --
        raising it there is a deliberate config decision, not something to
        infer here.
        """
        raw = self._options.get("png_workers", "1").strip()
        try:
            n = int(raw)
        except ValueError:
            return 1
        n = max(1, n)
        try:
            import multiprocessing
            cpus = multiprocessing.cpu_count()
        except (ImportError, NotImplementedError):
            return n
        return min(n, max(1, cpus))

    # -- shared values (same file for every project) --
    @property
    def spec_workbook(self) -> Path:
        return self._resolve("spec_workbook", self._raw)

    @property
    def ppt_template(self) -> Path:
        return self._resolve("ppt_template", self._raw)

    @property
    def title_prefix(self) -> str:
        """First word(s) of every project's PPT title slide -- 2026-07-22,
        default "Bench" if config.xml has no <title_prefix> entry. See
        title_suffix below; both combine with the project name as
        f"{title_prefix} {self.project} {title_suffix}"."""
        return self._raw.get("title_prefix", "Bench")

    @property
    def title_suffix(self) -> str:
        """Last word(s) of every project's PPT title slide -- 2026-07-22,
        default "Distribution" if config.xml has no <title_suffix> entry.
        Before this was added, BT_TX's two decks each hardcoded their own
        inconsistent title ("Bench BT TX Power / DEVM / ACP Distribution",
        "Bench BT TX Freq / Mod Accuracy" -- note the space in "BT TX" and
        the missing "Distribution" on the second one) instead of using the
        project name + a shared prefix/suffix; both now follow the exact
        same pattern as NB_Tx's title."""
        return self._raw.get("title_suffix", "Distribution")

    @property
    def extract_mode(self) -> str:
        """"both" (default), "ate", or "bench" -- 2026-07-22, shared across
        every project (lives in <paths>, not a per-project <options> block
        -- one project asking for "ate" while another still runs "both"
        was never a real use case). Controls which extraction steps
        run_all.py runs, and whether a project's build_ppt_*.py fetches/
        uses ATE data at all:
          "both"  -- run every Bench + ATE chart script, PPT shows both rows.
          "ate"   -- skip the Bench chart scripts (assumes their PNGs
                     already exist from a prior run and Bench data hasn't
                     changed); PPT still shows both rows.
          "bench" -- skip the ATE chart scripts AND force build_ppt to treat
                     ATE as absent (empty dict, not just "didn't re-run"),
                     so a stale ATE PNG from an earlier run never leaks into
                     a Bench-only report. PPT shows every ATE cell/column as
                     blank / "No ATE match", same rendering path as a
                     genuinely-unmatched combo.
        Invalid values fall back to "both" rather than raising -- a typo in
        config.xml should degrade to "do everything," not silently narrow
        the report."""
        raw = self._raw.get("extract_mode", "both").strip().lower()
        return raw if raw in ("both", "ate", "bench") else "both"

    @property
    def ate_log_path(self) -> Path | None:
        """The ATE log CSV itself (not a folder) -- 2026-07-22, replaces the
        old ate_log_dir + hardcoded "40pcs_raw_data_transposed.csv" filename
        convention every ATE script used to assume. That convention forced a
        rename/copy dance every time a new ATE pull arrived under a
        different filename (this happened several times in one day). Now
        config.xml's <ate_log_path> names the exact file to use -- switching
        logs is a one-line config edit, no renaming.

        Lives in <paths> (shared), not a per-project <projects><NAME>
        entry -- as of 2026-07-24, ALL FOUR projects (BT_TX, NB_Tx, UWB,
        WIFI) read this exact same physical ATE log file (BTTX_*/NBTX_*/
        UWBTX_*/WIFITX_* items all live in one CSV per the customer's ATE
        test program; confirmed by explicit user correction 2026-07-24 --
        an earlier, mistaken assumption had briefly split UWB/WIFI onto
        their own ate_log_path_uwb_wifi, since removed). A per-project copy
        would be pure duplication, same reasoning as extract_mode above.
        Every project's own ATE lookup module already filters strictly by
        its own item-name prefix, so all four safely reading the same file
        finds only its own rows. Returns None if config.xml has no
        <ate_log_path> entry at all rather than raising -- e.g. a given
        data pull covers only SOME projects' items (an 8DH5 BT-only retest,
        say); the other projects' lookups then correctly find zero rows and
        render "No ATE match" rather than crashing."""
        # 2026-07-28: an EMPTY <ate_log_path/> now yields None too, not a Path
        # pointing at the workspace root -- see _resolve()'s docstring. Leaving
        # the tag empty while swapping ATE pulls is an easy mistake, and before
        # this it defeated every caller's `is None` guard.
        if not self._raw.get("ate_log_path", "").strip():
            return None
        return self._resolve("ate_log_path", self._raw)

    # -- report_dir: the shared base ("report") with this project's name
    # auto-appended, e.g. report/BT_TX, report/NB_Tx. result_png_dir/
    # result_xlsx_dir are bare subfolder NAMES (not independently-resolved
    # paths) that nest INSIDE that same per-project report_dir, so a
    # project's charts, workbooks, and PPT all land under one
    # report/{project}/ tree, per the 2026-07-20 "one shared report/ folder"
    # request. --
    #
    # 2026-08-21: optional <report_dir_override> in a project's own
    # <projects><NAME> block (added for ReportBench GUI's Folders card --
    # its single per-row "Open" picker needs an actual place to persist a
    # picked report-output folder). When present, it replaces the whole
    # computed path above -- project name is NOT re-appended, the override
    # IS the final directory -- so picking any folder always means exactly
    # that folder, not that-folder/BT_TX. Absent for every project by
    # default; falls back to the shared-base behavior above.
    @property
    def report_dir(self) -> Path:
        override = self._proj_raw.get("report_dir_override", "").strip()
        if override:
            p = Path(override)
            return p if p.is_absolute() else (self.base_dir / p)
        return self._resolve("report_dir", self._raw) / self.project

    @property
    def result_png_dir(self) -> Path:
        return self.report_dir / self._raw["result_png_dir"]

    @property
    def result_xlsx_dir(self) -> Path:
        return self.report_dir / self._raw["result_xlsx_dir"]


_cache: dict[str, Paths] = {}


def get_paths(project: str, config_path: Path | None = None) -> Paths:
    """Returns the shared Paths instance for `project` (loaded once per
    project, cached), or a fresh one if an explicit config_path is given."""
    if config_path is not None:
        return Paths(project, config_path)
    if project not in _cache:
        _cache[project] = Paths(project)
    return _cache[project]


if __name__ == "__main__":
    import sys
    project = sys.argv[1] if len(sys.argv) > 1 else "BT_TX"
    p = get_paths(project)
    print(f"project: {p.project}")
    print(f"config: {p.config_path}")
    print(f"base_dir: {p.base_dir}")
    for name in ("bench_dir", "ate_log_path", "spec_workbook",
                 "ppt_template", "extract_mode", "result_png_dir",
                 "result_xlsx_dir", "report_dir"):
        print(f"  {name}: {getattr(p, name)}")
