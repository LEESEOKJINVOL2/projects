"""Parse per-offset ACP/In-Band-Emission spec limits for the ACP chart
overlay, from the same three sheets as bt_tx_freqacc_spec_lookup.py:
  BDR (1DH5)         -> 'BT - BDR & EDR Spec',   Mode='BDR', desc contains "ACP"
  EDR (3DH5)         -> 'BT - BDR & EDR Spec',   Mode='EDR', desc contains "In-Band Emissions"
  BLE (LE1M, LE2M)   -> 'BT - BLE and HDT Spec', Mode=packet, desc contains "In-Band Emission"
  HDR (4DH5, 8DH5)   -> 'BT - HDR HDRP UHDR Spec', Mode='HDR', desc contains
                        "In-Band Emissions", Conditions prefixed "HDR-4,"/"HDR-8,"
                        to disambiguate the two packets sharing one Mode block.
  HDRpS2, HDRPM8,
  HDRPM16, HDRPL16,
  HDRPL32            -> 'BT - HDR HDRP UHDR Spec', each packet now has its OWN
                        Mode row (HDRpS2/HDRpM8/HDRpM16/HDRpL16/HDRpL32 --
                        2026-08-14, v4.4 split what v0.4 shared as one
                        "HDRpS"/"HDRpM/L" Mode block per packet-pair). desc
                        contains "In-Band Emissions"; the old "HDRp-M (4MHz
                        BW),"/"HDRp-L (8MHz BW)," Conditions prefix is still
                        present (now redundant with Mode) and stripped.
                        2026-08-14: HDRPL16/32's +/-7..13MHz offsets are now
                        populated with real dBm values (v0.4 had them blank/
                        relative -- the note below about "no Limit line at
                        all" no longer applies to those offsets).
  HDT3, HDT4, HDT8   -> 'BT - BLE and HDT Spec', Mode='HDT2' specifically
                        (2026-07-20 customer instruction: "HDT 2 3 4 상관없이
                        HDT2의 In-Band Emission값을 찾아서 추가" -- always use
                        HDT2's own row regardless of which HDT variant is
                        being charted. HDT3 has no In-Band Emission row of
                        its own in the sheet at all; HDT4/HDT6 do, but are
                        numerically identical to HDT2's -- simplest to reuse
                        HDT2 uniformly rather than look up each variant).

UHDR32/UHDR48 ACP charts remain NOT covered -- their own ACP limits are
relative-to-EBW dBr masks that require a measured EBW per DUT, not a fixed
value we can overlay (confirmed with user 2026-07-17).

Per-offset values are only reliable when the sheet's unit is an absolute
level (dBm). Several near-carrier offsets (BDR/BLE fTX+/-1MHz, HDR-4
fTX+/-2MHz, HDR-8 fTX+/-3MHz, EDR fTX+/-1MHz, HDRpS2/HDRp-M/HDRp-L's own
near-carrier offsets) are either blank or a "Relative to power in wanted
signal" dBc/dB value -- NOT comparable to our bench data column
(acp_avg-dBm, an absolute level). Confirmed with user 2026-07-17 (and
reconfirmed 2026-07-20 for the HDRP additions): skip those offsets entirely
rather than overlay a unit-mismatched line.

Returns spec[(packet_type, offset)] = (min_dbm_or_None, max_dbm_or_None).
"""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl

from bt_tx_freqacc_spec_lookup import BDR_EDR_SHEET, BLE_HDT_SHEET, HDR_SHEET, _parse_records, _to_float

RELATIVE_RE = re.compile(r"relative to power", re.IGNORECASE)


def _offsets_from_condition(cond: str):
    """Extract the offset magnitudes from a Conditions string like
    'fTX +/- 2MHz', 'fTX +/-3/4/5 MHz', 'fTX +/-3/4/5… MHz',
    'fTX +/- 6/7/8/9/10MHz', or (after stripping a 'HDR-4,'/'HDR-8,' prefix)
    'fTX +/-5/6/7 MHz'. Returns a list of positive ints."""
    m = re.search(r"fTX\s*(.*)", cond, re.IGNORECASE)
    tail = m.group(1) if m else cond
    return [int(n) for n in re.findall(r"\d+", tail)]


def _add_offsets(spec, packet, cond, max_val, notes):
    if max_val is None:
        return
    if notes and RELATIVE_RE.search(notes):
        return  # relative dBc/dB value -- not comparable to our absolute dBm data, skip
    for off in _offsets_from_condition(cond):
        for signed in (off, -off):
            spec[(packet, signed)] = (None, max_val)


def parse_spec(xlsx_path: Path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    bdr_edr = _parse_records(wb[BDR_EDR_SHEET])
    ble_hdt = _parse_records(wb[BLE_HDT_SHEET])
    hdr = _parse_records(wb[HDR_SHEET])

    spec: dict[tuple, tuple] = {}

    # BDR (1DH5)
    for rec in bdr_edr:
        if rec["mode"] == "BDR" and "ACP" in rec["desc"] and rec["cond"]:
            _add_offsets(spec, "1DH5", rec["cond"], rec["max"], rec["notes"])

    # EDR (3DH5)
    for rec in bdr_edr:
        if rec["mode"] == "EDR" and "In-Band Emissions" in rec["desc"] and rec["cond"]:
            _add_offsets(spec, "3DH5", rec["cond"], rec["max"], rec["notes"])

    # BLE (LE1M, LE2M)
    for packet in ("LE1M", "LE2M"):
        for rec in ble_hdt:
            if rec["mode"] == packet and "In-Band Emission" in rec["desc"] and rec["cond"]:
                _add_offsets(spec, packet, rec["cond"], rec["max"], rec["notes"])

    # HDR (4DH5, 8DH5) -- shared Mode='HDR' block, disambiguated by the
    # "HDR-4,"/"HDR-8," prefix in Conditions.
    for rec in hdr:
        if rec["mode"] != "HDR" or "In-Band Emissions" not in rec["desc"] or not rec["cond"]:
            continue
        cond = rec["cond"]
        if cond.startswith("HDR-4,"):
            _add_offsets(spec, "4DH5", cond[len("HDR-4,"):], rec["max"], rec["notes"])
        elif cond.startswith("HDR-8,"):
            _add_offsets(spec, "8DH5", cond[len("HDR-8,"):], rec["max"], rec["notes"])

    # HDRp-family -- 2026-08-14: the new workbook (v4.4) splits what v0.4
    # shared as ONE "HDRpS"/"HDRpM/L" Mode block into a SEPARATE Mode row
    # per packet (HDRpS2, HDRpM4/6/8/12/16, HDRpL8/12/16/24/32) -- no more
    # prefix-parsing needed, Mode already identifies the packet directly.
    # The old "HDRp-M (4MHz BW),"/"HDRp-L (8MHz BW)," Conditions prefix is
    # still present (now redundant with Mode) and still stripped for a
    # clean offset string. Also newly usable: HDRPL16/32's +/-7..13MHz
    # offsets are populated with real dBm values in v4.4 (v0.4 had them
    # blank/relative -- see this module's old docstring note "HDRPL16/32
    # end up with no Limit line at all", no longer true).
    HDRP_MODE_TO_PACKET = {
        "HDRpS2": ("HDRPS2", "HDRpS2,"),
        "HDRpM8": ("HDRPM8", "HDRp-M (4MHz BW),"),
        "HDRpM12": ("HDRPM12", "HDRp-M (4MHz BW),"),
        "HDRpM16": ("HDRPM16", "HDRp-M (4MHz BW),"),
        "HDRpL16": ("HDRPL16", "HDRp-L (8MHz BW),"),
        "HDRpL24": ("HDRPL24", "HDRp-L (8MHz BW),"),
        "HDRpL32": ("HDRPL32", "HDRp-L (8MHz BW),"),
    }
    for rec in hdr:
        entry = HDRP_MODE_TO_PACKET.get(rec["mode"])
        if entry is None or "In-Band Emissions" not in rec["desc"] or not rec["cond"]:
            continue
        packet, prefix = entry
        cond = rec["cond"]
        tail = cond[len(prefix):] if cond.startswith(prefix) else cond
        _add_offsets(spec, packet, tail, rec["max"], rec["notes"])

    # HDT3/HDT4/HDT6/HDT8 -- always HDT2's own In-Band Emission row
    # (2026-07-20 customer instruction: "HDT 2 3 4 상관없이 HDT2의 In-Band
    # Emission값을 찾아서 추가"), never each variant's own row even where
    # one exists. HDT6 added 2026-08-14 -- a real Bench packet_type this
    # list never covered (the original three were chosen against a bench
    # pull that had HDT3/HDT4, not HDT6; the current pull is the reverse).
    for rec in ble_hdt:
        if rec["mode"] == "HDT2" and "In-Band Emission" in rec["desc"] and rec["cond"]:
            for packet in ("HDT3", "HDT4", "HDT6", "HDT8"):
                _add_offsets(spec, packet, rec["cond"], rec["max"], rec["notes"])

    return spec


def lookup(spec: dict, packet_type: str, offset: int):
    return spec.get((packet_type, offset))


if __name__ == "__main__":
    spec = parse_spec(Path("NPI_TestCoverage_v0.4.xlsx"))
    for key in sorted(spec, key=lambda k: (k[0], k[1])):
        print(key, "->", spec[key])
    print()
    print("total entries:", len(spec))
    print("lookup(1DH5, 2) ->", lookup(spec, "1DH5", 2))
    print("lookup(1DH5, 1) ->", lookup(spec, "1DH5", 1), "(expect None, no value/skip)")
    print("lookup(4DH5, 2) ->", lookup(spec, "4DH5", 2), "(expect None, dBc skip)")
    print("lookup(4DH5, 5) ->", lookup(spec, "4DH5", 5))
    print("lookup(8DH5, 6) ->", lookup(spec, "8DH5", 6))
    print("lookup(LE2M, 6) ->", lookup(spec, "LE2M", 6))
    print("lookup(LE2M, 2) ->", lookup(spec, "LE2M", 2), "(expect None, no ±1/2/3 row for LE2M)")
