"""Merge the two standalone BT_TX revision2 decks into one file, per
customer request (2026-07-20): Power/DEVM/ACP slides first, FreqAcc slides
appended after -- one deliverable instead of two separate PPTX files.

Uses PowerPoint COM's Slides.InsertFromFile (not python-pptx, which has no
built-in cross-presentation slide copy) -- both source decks were built off
the same PPT-Template.pptx, so themes/masters/layouts already match and
InsertFromFile carries every slide over with no manual XML surgery needed.

Run this AFTER both build_ppt_full_revision2.py and
build_ppt_freqacc_revision2.py have produced today's decks.

2026-07-22: PowerPoint's own COM SaveAs writes the merged file's zip
container with almost no compression -- confirmed by inspecting the zip
directly: ~1126MB compressed vs ~1144MB uncompressed content, a ~1:1 ratio,
vs. the ~65% ratio python-pptx's own writer gets on the same PNG-heavy
content. That alone was inflating the merged deck to ~1.1GB when the two
source decks summed to ~740MB. Fixed by reopening the freshly-merged file
with python-pptx and saving it right back over itself -- pure zip
re-compression, no XML/content touched, confirmed byte-for-byte identical
slide count before/after (~30s for a 530-slide/1.1GB deck).
"""
from __future__ import annotations

import datetime
import subprocess
from pathlib import Path

import win32com.client
from pptx import Presentation

import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).resolve().parent.parent))
import paths as bt_tx_paths

PATHS = bt_tx_paths.get_paths("BT_TX")
REPORT_DATE = datetime.date.today().isoformat()


def main() -> None:
    report_dir = PATHS.report_dir
    pwr_path = report_dir / f"{REPORT_DATE}_BT_TX_PWR_DEVM_ACP_Revision2.pptx"
    freqacc_path = report_dir / f"{REPORT_DATE}_BT_TX_FreqAcc_Revision2.pptx"
    out_path = report_dir / f"{REPORT_DATE}_BT_TX_Full_Revision2.pptx"

    if not pwr_path.exists():
        raise SystemExit(f"Missing {pwr_path} -- run build_ppt_full_revision2.py first")
    if not freqacc_path.exists():
        raise SystemExit(f"Missing {freqacc_path} -- run build_ppt_freqacc_revision2.py first")

    # A stale POWERPNT.EXE from a prior COM session (e.g. leftover verification)
    # would make Dispatch() attach to that instance instead of a clean one.
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
    pp = win32com.client.Dispatch("PowerPoint.Application")
    pp.Visible = True
    pres = pp.Presentations.Open(str(pwr_path), WithWindow=False)
    n_pwr = pres.Slides.Count
    print(f"Base deck (Power/DEVM/ACP): {n_pwr} slides")

    # Index = insert AFTER this slide number -- n_pwr means "at the very end".
    pres.Slides.InsertFromFile(str(freqacc_path), n_pwr)
    n_total = pres.Slides.Count
    print(f"FreqAcc slides appended: {n_total - n_pwr}")
    print(f"Total: {n_total}")

    pres.SaveAs(str(out_path))
    pres.Close()
    pp.Quit()
    raw_size = out_path.stat().st_size
    print(f"\nSaved (uncompressed COM output): {out_path}  ({raw_size / 1e6:.0f}MB)")

    print("Re-compressing via python-pptx (COM's SaveAs barely deflates the zip)...")
    Presentation(str(out_path)).save(str(out_path))
    compressed_size = out_path.stat().st_size
    print(f"Done: {out_path}  ({compressed_size / 1e6:.0f}MB, "
          f"{100 * (1 - compressed_size / raw_size):.0f}% smaller)")


if __name__ == "__main__":
    main()
