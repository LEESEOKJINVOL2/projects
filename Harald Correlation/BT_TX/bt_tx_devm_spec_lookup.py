"""Parse DEVM RMS / DEVM Peak / DEVM99pct spec limits per packet_type, from
the same three sheets as bt_tx_freqacc_spec_lookup.py
(NPI_TestCoverage_v0.4.xlsx). Supersedes bt_tx_evm_spec_lookup.py (RMS-only)
-- "EVM" and "DEVM" are the same thing in this project; all three metrics
now live under the "DEVM" name (2026-07-19).

Confirmed with user (2026-07-19):
  - 3DH5 uses 8-DPSK (3DH5 is the 8-DPSK/EDR3 packet).
  - 2026-08-15 update: 2DH5 (the pi/4-DQPSK/EDR2 packet) IS present in the
    current Harald_P2 Bench dataset, superseding the 2026-07-19 assumption
    above that it wasn't -- its spec row was being silently skipped (see
    parse_spec's EDR block) until this was caught and fixed.
  - 3DH5's "TX 99% DEVM" row (BDR&EDR sheet) has Min=99/Max=blank -- that's
    a symbol-count requirement, not the DEVM value itself. The real
    threshold (0.20 = 20%) is buried in the Notes text ("% symbols whose
    DEVM is better than 0.20"), same pattern as make_bt_tx_pwr_revision2's
    Df2max fix. All other rows below have the real value directly in the
    Max column.
  - HDR (4DH5, 8DH5) has NO "TX Peak DEVM" row in the sheet at all -- only
    RMS and 99pct exist for HDR. The sheet's "TX Differential Phase
    Encoding" row is a different measurement entirely (not a DEVM value)
    and is not used here.
  - DEVM Peak's spec is confirmed to be 3DH5-only; every other packet gets
    no Peak spec (None), even though the DEVM Peak CHART itself now covers
    all 13 packets (matching DEVM RMS/99pct's packet breadth -- expanded
    2026-07-19 from the earlier 3-packet-only devm_peak metric).
  - UHDR's sheet rows are labelled "TX Peak DEVM" in the Desc column, but
    their Units/Notes ("(99%)" / "99% EVM") mean they are actually the
    99th-percentile value, not a peak (single max) value -- confirmed with
    user to map these to DEVM99pct, not DEVM Peak.

Per-packet mapping:
  DEVM RMS ("evm_aver_aver-%" data column):
    2DH5              -> 'BT - BDR & EDR Spec', Mode=EDR,  'TX RMS DEVM' / pi/4-DQPSK   = 0.07
    3DH5              -> 'BT - BDR & EDR Spec', Mode=EDR,  'TX RMS DEVM' / 8-DPSK       = 0.13
    4DH5, 8DH5        -> 'BT - HDR HDRP UHDR Spec', Mode=HDR, 'TX RMS DEVM' / pi/4-DQPSK = 0.07
    HDRPS2            -> same sheet, Mode=HDRpS,   'TX RMS DEVM' / OQPSK                = 0.084
    HDRPM8/16,
    HDRPL16/32        -> same sheet, Mode=HDRpM/L, 'TX RMS EVM' / 16QAM                  = 0.0316
    UHDR32            -> same sheet, UHDR section, 'TX RMS DEVM' / 'UHDR 16/24/32'       = 0.084
    UHDR48            -> same sheet, UHDR section, 'TX RMS DEVM' / 'UHDR 48'             = 0.05
    HDT3/HDT4/HDT8    -> 'BT - BLE and HDT Spec', Mode=HDT3/HDT4/HDT8,
                         'EVM -- Payload RMS' in dB (own value per packet:
                         -13/-16/-22), converted via 10**(dB/20)*100
    1DH5, LE1M, LE2M  -> no spec row exists (GFSK modulation)

  DEVM Peak ("evm_peak_aver-%" data column):
    2DH5              -> 'BT - BDR & EDR Spec', Mode=EDR, 'TX Peak DEVM' / pi/4-DQPSK    = 0.35
    3DH5              -> 'BT - BDR & EDR Spec', Mode=EDR, 'TX Peak DEVM' / 8-DPSK        = 0.25
    everyone else     -> no spec (confirmed 2026-07-19 -- sheet has no Peak row for them)

  DEVM99pct ("evm_99pct_aver-%" data column):
    2DH5              -> 'BT - BDR & EDR Spec', Mode=EDR, 'TX 99% DEVM' / pi/4-DQPSK     = 0.17
    3DH5              -> 'BT - BDR & EDR Spec', Mode=EDR, 'TX 99% DEVM' / 8-DPSK,
                         value from Notes text (0.20) -- see note above
    4DH5, 8DH5        -> 'BT - HDR HDRP UHDR Spec', Mode=HDR, 'TX 99% DEVM' / pi/4-DQPSK = 0.17
    HDRPS2            -> same sheet, Mode=HDRpS,   'TX 99% DEVM' / OQPSK                 = 0.21
    HDRPM8/16,
    HDRPL16/32        -> same sheet, Mode=HDRpM/L, 'TX 99% EVM' / 16QAM                  = 0.079
    UHDR32            -> same sheet, UHDR section (labelled 'TX Peak DEVM' in the sheet,
                         but meant as 99pct -- see note above), 'UHDR 16/24/32'          = 0.21
    UHDR48            -> same sheet, UHDR section, 'UHDR 48'                             = 0.15
    HDT3/HDT4/HDT8, 1DH5, LE1M, LE2M -> no spec row exists

All non-HDT values above are fractions-of-1 in the sheet (e.g. 0.2 = 20%);
multiplied by 100 to match our %-scale data columns. Every value here is an
UPPER limit only (no min side).

2026-08-14: re-pointed at the customer's new Harald_P2_Test_Coverage_v4.4_
JSCK.xlsx (same 3 sheet NAMES, real content differences -- see
bt_tx_power_spec_lookup.py's sibling note in paths.py's bt_tx_spec_workbook
docstring). What changed here:
  - Value SCALING is now inconsistent WITHIN the sheet itself: some rows
    (HDR/HDRpS2/HDRpM4-8/HDT-family) already store the value as a percent
    number (e.g. "7" meaning 7%), while others (UHDR, HDRpL16) still store
    a fraction-of-1 (e.g. "0.084" meaning 8.4%) -- confirmed by literally
    reading both conventions off adjacent rows referencing the identical
    modulation class. `_pct()` normalizes by magnitude (a real value < 1 is
    unambiguously a fraction; no BT DEVM/EVM limit is ever under 1%), not a
    blind *100 -- the old blind *100 silently produced 700% instead of 7%
    for HDR's RMS DEVM once pointed at this workbook.
  - HDRp-family Mode rows are now split per-packet (HDRpS2, HDRpM4/6/8/12/
    16, HDRpL8/12/16/24/32) instead of v0.4's shared "HDRpS"/"HDRpM/L"
    blocks -- and HDRPM8 vs HDRPM16 turned out to have GENUINELY DIFFERENT
    values once split (HDRPM8 = OQPSK figure 8.4%, HDRPM16 = 16QAM figure
    3.16%), not the same value as the old shared-block code assumed.
  - EDR's "TX 99% DEVM" value moved from being buried in the Notes text
    into the Max column directly; HDR's got "TX 99% DEVM Peak" (with the
    literal word "Peak" now in the description). Both are now read from
    Max first, falling back to the old Notes-regex only if Max is blank.
  - HDT-family EVM rows are richer: HDT2/3/4/6/8 now have BOTH an RMS row
    AND a 99%-peak row (v0.4 had RMS only, as a dB value needing
    10**(db/20)*100 conversion; v4.4 gives both directly in %). HDT6 (a
    real Bench packet_type) is now included -- v0.4's list (HDT3/HDT4/HDT8)
    never actually matched any Bench data (Bench has no HDT3/HDT4 rows),
    while HDT6 (which Bench DOES have) was missing.
"""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl

from bt_tx_freqacc_spec_lookup import BDR_EDR_SHEET, HDR_SHEET, BLE_HDT_SHEET, _parse_records

HDT_PAYLOAD_DESC = "EVM -- Payload RMS"
NOTES_PCT_RE = re.compile(r"better than\s*([\d.]+)", re.IGNORECASE)

METRICS = ("rms", "peak", "99pct")


def _pct(v):
    """Normalize a DEVM/EVM value to percent-of-100 scale. See module
    docstring's 2026-08-14 note: the new workbook mixes both conventions
    WITHIN the same sheet (some rows already store e.g. "7" meaning 7%,
    others still store "0.084" meaning 8.4%) -- a value < 1 is
    unambiguously a fraction in this domain (no real BT DEVM/EVM limit is
    under 1%), so that's the normalization rule, not a blind *100."""
    return v * 100 if v < 1 else v


def parse_spec(xlsx_path: Path):
    """Return spec[metric][packet_type] = max_percent, metric in
    ('rms', 'peak', '99pct')."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    bdr_edr = _parse_records(wb[BDR_EDR_SHEET])
    hdr = _parse_records(wb[HDR_SHEET])
    ble_hdt = _parse_records(wb[BLE_HDT_SHEET])

    spec = {m: {} for m in METRICS}

    # --- 2DH5/3DH5 (EDR), all three metrics -- Conditions column (E)
    # determines which packet a row belongs to: EDR2 (2DH5) payload is
    # pi/4-DQPSK, EDR3 (3DH5) payload is 8-DPSK (customer-confirmed,
    # 2026-08-15). 2DH5's spec was previously missing entirely: the old
    # code only ever matched "8-DPSK" on the (wrong, dataset-specific)
    # assumption that "pi/4-DQPSK would be for a 2DH5/EDR2 packet, which
    # isn't in our dataset" -- the current Harald_P2 bench data DOES
    # include 2DH5, so that row was silently never picked up and 2DH5 got
    # no Limit line on the RMS/Peak/99pct Bench charts.
    EDR_COND_TO_PACKET = {"8-DPSK": "3DH5", "DQPSK": "2DH5"}
    for rec in bdr_edr:
        if rec["mode"] != "EDR" or not rec["cond"]:
            continue
        packet = next((p for cond_key, p in EDR_COND_TO_PACKET.items() if cond_key in rec["cond"]), None)
        if packet is None:
            continue
        if rec["desc"] == "TX RMS DEVM" and rec["max"] is not None:
            spec["rms"][packet] = _pct(rec["max"])
        elif rec["desc"] == "TX Peak DEVM" and rec["max"] is not None:
            spec["peak"][packet] = _pct(rec["max"])
        elif rec["desc"] == "TX 99% DEVM":
            if rec["max"] is not None:
                spec["99pct"][packet] = _pct(rec["max"])
            else:
                m = NOTES_PCT_RE.search(rec["notes"])
                if m:
                    spec["99pct"][packet] = float(m.group(1)) * 100

    # --- 4DH5 / 8DH5 (HDR): RMS + 99pct only ---
    for rec in hdr:
        if rec["mode"] != "HDR" or not rec["cond"] or "DQPSK" not in rec["cond"]:
            continue
        if rec["desc"] == "TX RMS DEVM" and rec["max"] is not None:
            for p in ("4DH5", "8DH5"):
                spec["rms"][p] = _pct(rec["max"])
        elif rec["desc"] in ("TX 99% DEVM", "TX 99% DEVM Peak") and rec["max"] is not None:
            for p in ("4DH5", "8DH5"):
                spec["99pct"][p] = _pct(rec["max"])

    # --- HDRp-family: 2026-08-14, each packet now has its OWN Mode row
    # (see module docstring) -- direct Mode -> packet_type map, no more
    # shared-block matching by Conditions text. desc is "TX RMS DEVM"/"TX
    # RMS EVM" (RMS) or "TX 99% DEVM Peak"/"TX 99% EVM Peak" (99pct) --
    # "DEVM" vs "EVM" just tracks which modulation-class row it came from,
    # both mean the same measurement. ---
    HDRP_MODE_TO_PACKET = {
        "HDRpS2": "HDRPS2", "HDRpM8": "HDRPM8", "HDRpM12": "HDRPM12",
        "HDRpM16": "HDRPM16", "HDRpL16": "HDRPL16", "HDRpL24": "HDRPL24",
        "HDRpL32": "HDRPL32",
    }
    for rec in hdr:
        packet = HDRP_MODE_TO_PACKET.get(rec["mode"])
        if packet is None or rec["max"] is None:
            continue
        desc = rec["desc"]
        if desc in ("TX RMS DEVM", "TX RMS EVM"):
            spec["rms"][packet] = _pct(rec["max"])
        elif desc in ("TX 99% DEVM Peak", "TX 99% EVM Peak"):
            spec["99pct"][packet] = _pct(rec["max"])

    # --- UHDR32 / UHDR48: RMS + 99pct, matched by Conditions text (Mode
    # column stays 'UHDR' for both packets) ---
    for rec in hdr:
        if rec["mode"] != "UHDR" or not rec["cond"] or rec["max"] is None:
            continue
        packet = {"UHDR 16/24/32": "UHDR32", "UHDR 48": "UHDR48"}.get(rec["cond"])
        if packet is None:
            continue
        if rec["desc"] == "TX RMS DEVM":
            spec["rms"][packet] = _pct(rec["max"])
        elif rec["desc"] == "TX Peak DEVM":
            spec["99pct"][packet] = _pct(rec["max"])

    # --- HDT2/HDT3/HDT4/HDT6/HDT8: RMS + 99pct, both given directly in %
    # now (v0.4 only had RMS, as a dB value needing 10**(dB/20)*100 --
    # v4.4's HDT2/3/4/6/8 rows are desc="EVM", cond="RMS" or "99% peak").
    # HDT6 added (a real Bench packet_type v0.4's HDT3/HDT4/HDT8 list never
    # actually covered -- Bench has no HDT3/HDT4 data at all). ---
    for packet in ("HDT3", "HDT4", "HDT6", "HDT8"):
        for rec in ble_hdt:
            if rec["mode"] != packet or rec["desc"] != "EVM" or rec["max"] is None:
                continue
            if rec["cond"] == "RMS":
                spec["rms"][packet] = _pct(rec["max"])
            elif rec["cond"] and "99%" in rec["cond"]:
                spec["99pct"][packet] = _pct(rec["max"])

    return spec


def lookup(spec: dict, metric: str, packet_type: str):
    """Return (lsl, usl) with lsl always None (upper-limit-only spec), or
    None if this packet has no spec for this metric."""
    v = spec.get(metric, {}).get(packet_type)
    if v is None:
        return None
    return (None, v)


if __name__ == "__main__":
    spec = parse_spec(Path("NPI_TestCoverage_v0.4.xlsx"))
    for metric in METRICS:
        print(f"--- {metric} ---")
        for packet in sorted(spec[metric]):
            print(" ", packet, "->", round(spec[metric][packet], 3), "%")
    print()
    print("lookup(rms, HDT3) ->", lookup(spec, "rms", "HDT3"))
    print("lookup(peak, 3DH5) ->", lookup(spec, "peak", "3DH5"))
    print("lookup(peak, 4DH5) -> (expect None)", lookup(spec, "peak", "4DH5"))
    print("lookup(99pct, 3DH5) ->", lookup(spec, "99pct", "3DH5"), "(expect ~20.0)")
    print("lookup(99pct, UHDR32) ->", lookup(spec, "99pct", "UHDR32"), "(expect ~21.0)")
    print("lookup(rms, 1DH5) -> (expect None)", lookup(spec, "rms", "1DH5"))
