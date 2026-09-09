"""Parse the 'BT - BDR & EDR Spec' and 'BT - BLE and HDT Spec' sheets
(NPI_TestCoverage_v0.4.xlsx) into per-metric (min, max) limits in Hz, for
overlaying red dashed spec lines on the FreqAcc/ModAcc charts.

Sheet layout (both sheets, same convention as 'BT Performance Specification'):
  col B = RF Test Case ID (sparse), col C = Mode (forward-filled -- only set
  on the first row of a mode/packet block), col D = Test Description
  (always explicit), col E = Conditions (explicit or None), col F = Min,
  col G = Max, col H = Units, col I = Notes.
  Section-title rows (e.g. "TRANSMITTER TESTS") have text only in col B and
  None in col D -- these are skipped and do NOT reset the forward-filled Mode.

Only metric/packet combos that (a) the customer specified and (b) actually
have a generated chart are wired up here (confirmed 2026-07-17 against
Result_PNG/BT_TX/*_revision2/ -- several customer-named combos have zero
charts because the underlying CSV column is empty for that packet, e.g.
BDR's own "cfo"/"icft" chart doesn't exist, HDT3/HDT8 have no CFO/Drift-Rate
rows in the spec sheet at all):

  cfo        : LE1M, LE2M          -- BLE and HDT Spec sheet's "Carrier
                                       Frequency Offset" row
  icft       : 1DH5, HDT4, HDT3, HDT8 -- 1DH5 from BDR&EDR sheet's "TX
                                       Carrier Frequency Offset" row; HDT4
                                       from BLE&HDT sheet's "Carrier
                                       Frequency Offset" row; HDT3/HDT8 have
                                       no such row in either sheet, so their
                                       ±125000 Hz is a literal customer-
                                       supplied override (2026-07-20: "HDT:
                                       -125~125 kHz"), not a sheet lookup --
                                       see parse_spec's icft block. NOT
                                       LE1M/LE2M -- ICFT has no chart at all
                                       for those two packets.
  drift_rate : 1DH5                -- BDR "TX Carrier Frequency Drift Rate"
  df1_avg    : 1DH5, LE1M, LE2M    -- band-dependent (2.4G vs 5/6G row)
  df2_max    : 1DH5                -- BDR row's Min/Max cells hold a 99.9%
                                       statistical criterion, not the Hz
                                       limit itself; the real value is in
                                       the Notes text ("...must be > 115
                                       kHz") -- confirmed with user
                                       2026-07-17, parsed via regex, used as
                                       a min-only line.
  omega_i    : 3DH5 (EDR sheet), 4DH5 & 8DH5 (HDR sheet, shared row)
  omega_o    : 3DH5 (EDR sheet), 4DH5 & 8DH5 (HDR sheet, shared row)
  omega_io   : 3DH5 (EDR sheet), 4DH5 & 8DH5 (HDR sheet, shared row)

df2_avg, devm_peak: customer confirmed "spec 없음" -- not looked up.
"""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl

BDR_EDR_SHEET = "BT - BDR & EDR Spec"
BLE_HDT_SHEET = "BT - BLE and HDT Spec"
HDR_SHEET = "BT - HDR HDRP UHDR Spec"

COL_MODE = 3
COL_DESC = 4
COL_COND = 5
COL_MIN = 6
COL_MAX = 7
COL_UNITS = 8
COL_NOTES = 9

NOTES_KHZ_RE = re.compile(r">\s*([\d.]+)\s*kHz", re.IGNORECASE)


def _to_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().lstrip("+")
    try:
        return float(s)
    except ValueError:
        return None


def _parse_records(ws):
    """Return [{mode, desc, cond, min, max, units, notes}, ...] -- one dict
    per data row, Mode AND Test-Description forward-filled.

    Some sheets (BT - BDR & EDR Spec) repeat both Mode and Test Description
    explicitly on every row of a multi-offset block; others (BT - HDR HDRP
    UHDR Spec) blank-continue BOTH columns, only setting them on the first
    row of the block. A true section-title row (e.g. "TRANSMITTER TESTS")
    has text ONLY in col B, with Mode/Desc/Cond/Min/Max/Units/Notes all
    None -- detected via `has_data` and skipped without disturbing the
    forward-filled state."""
    records = []
    current_mode = None
    current_desc = None
    for r in range(1, ws.max_row + 1):
        mode_cell = ws.cell(r, COL_MODE).value
        desc_cell = ws.cell(r, COL_DESC).value
        cond_cell = ws.cell(r, COL_COND).value
        min_cell = ws.cell(r, COL_MIN).value
        max_cell = ws.cell(r, COL_MAX).value
        units_cell = ws.cell(r, COL_UNITS).value
        notes_cell = ws.cell(r, COL_NOTES).value
        has_data = any(v is not None for v in (cond_cell, min_cell, max_cell, units_cell, notes_cell))
        if desc_cell is None and not has_data:
            continue  # section-title row -- skip, don't touch forward-fill state
        if mode_cell:
            current_mode = str(mode_cell).strip()
        if desc_cell:
            current_desc = str(desc_cell).strip()
        if current_desc is None:
            continue
        records.append({
            "mode": current_mode,
            "desc": current_desc,
            "cond": str(cond_cell).strip() if cond_cell else None,
            "min": _to_float(min_cell),
            "max": _to_float(max_cell),
            "units": units_cell,
            "notes": notes_cell or "",
        })
    return records


def _find(records, mode, desc_contains, cond_contains=None, cond_is_none=False, cond_exact=None):
    for rec in records:
        if rec["mode"] != mode:
            continue
        if desc_contains.lower() not in rec["desc"].lower():
            continue
        if cond_is_none:
            if rec["cond"] is not None:
                continue
        elif cond_exact is not None:
            norm = lambda s: re.sub(r"\s+", " ", s.strip().lower())
            if rec["cond"] is None or norm(rec["cond"]) != norm(cond_exact):
                continue
        elif cond_contains is not None:
            if rec["cond"] is None or cond_contains.lower() not in rec["cond"].lower():
                continue
        return rec
    return None


def parse_spec(xlsx_path: Path):
    """Return spec[(metric_key, packet_type, band_or_None)] = (min_hz, max_hz)."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    bdr_edr = _parse_records(wb[BDR_EDR_SHEET])
    ble_hdt = _parse_records(wb[BLE_HDT_SHEET])
    hdr = _parse_records(wb[HDR_SHEET])

    spec = {}

    def khz(v):
        return None if v is None else v * 1000.0

    # -- cfo: LE1M, LE2M (both sheets have the same "Carrier Frequency
    # Offset" row per packet, no band split) --
    for packet, mode in (("LE1M", "LE1M"), ("LE2M", "LE2M")):
        rec = _find(ble_hdt, mode, "Carrier Frequency Offset", cond_is_none=True)
        if rec:
            spec[("cfo", packet, None)] = (khz(rec["min"]), khz(rec["max"]))

    # -- icft: 1DH5 and HDT4 only (customer-corrected 2026-07-20, replacing
    # an earlier LE1M/LE2M version of this block). LE1M/LE2M deliberately
    # EXCLUDED even though the 'BT - BLE and HDT Spec' sheet has a
    # "Carrier Frequency Offset" row for them (same row cfo itself uses) --
    # verified at the raw bench CSV level that
    # initial_frequency_offset_aver-Hz is blank for every LE1M/LE2M row
    # (including rows where the sibling freq_offset_aver-Hz/cfo column IS
    # populated), so ICFT has no chart for those packets at all and a spec
    # there would never be visible. Customer: "원본 데이터에서 없으면
    # 없는대로... le1m, le2m 제외하고 나머지 limit 적용".
    #   1DH5 -- BDR&EDR sheet, BDR mode, "TX Carrier Frequency Offset" row
    #           (min=-75kHz, max=75kHz, customer-confirmed)
    #   HDT4 -- BLE&HDT sheet, HDT4 mode, "Carrier Frequency Offset" row
    #           (min=-125kHz, max=125kHz)
    # HDT3/HDT8 have NO such row in the sheet at all (same gap as their
    # missing CFO/Drift-Rate rows, see module docstring) -- they stay
    # spec-less, which just means their icft chart keeps the old
    # data-driven (non-spec-anchored) Y-axis, exactly as before.
    rec = _find(bdr_edr, "BDR", "TX Carrier Frequency Offset", cond_is_none=True)
    if rec:
        spec[("icft", "1DH5", None)] = (khz(rec["min"]), khz(rec["max"]))

    rec = _find(ble_hdt, "HDT4", "Carrier Frequency Offset", cond_is_none=True)
    if rec:
        spec[("icft", "HDT4", None)] = (khz(rec["min"]), khz(rec["max"]))

    # HDT6 -- 2026-08-14: the new workbook gives HDT6 its own "Carrier
    # Frequency Offset" row too (-125/+125kHz, same value as HDT4's), a
    # real Bench packet_type the old HDT3/HDT8-hardcoded-override list
    # below never covered. Read from the sheet now that it exists, same as
    # HDT4, rather than adding a third hardcode.
    rec = _find(ble_hdt, "HDT6", "Carrier Frequency Offset", cond_is_none=True)
    if rec:
        spec[("icft", "HDT6", None)] = (khz(rec["min"]), khz(rec["max"]))

    # HDT3/HDT8 have no ICFT/CFO row in either sheet at all (see docstring
    # above) -- normally that means "no spec, no Limit line" and we leave it
    # alone. But the customer explicitly directed (2026-07-20): "HDT: -125~125
    # kHz" fixed, for ICFT specifically. Unlike every other value in this
    # module, this one is NOT derived from a sheet lookup -- there is nothing
    # in the workbook to derive it from, so this is a literal customer-
    # supplied override, not a self-authored hardcode. HDT4 already lands on
    # this exact same ±125000 Hz from its own sheet row above; this only
    # fills the gap for the two packets the sheet has nothing for.
    spec[("icft", "HDT3", None)] = (-125000.0, 125000.0)
    spec[("icft", "HDT8", None)] = (-125000.0, 125000.0)

    # -- drift_rate: 1DH5 only --
    rec = _find(bdr_edr, "BDR", "TX Carrier Frequency Drift Rate", cond_is_none=True)
    if rec:
        spec[("drift_rate", "1DH5", None)] = (khz(rec["min"]), khz(rec["max"]))

    # -- df1_avg: band-dependent, 1DH5/LE1M/LE2M --
    for packet, mode, sheet in (("1DH5", "BDR", bdr_edr), ("LE1M", "LE1M", ble_hdt), ("LE2M", "LE2M", ble_hdt)):
        for band, cond_sub in ((2.4, "2"), (5, "5"), (6, "5")):
            rec = _find(sheet, mode, "Δf1 (avg)", cond_contains=cond_sub)
            if rec:
                spec[("df1_avg", packet, band)] = (khz(rec["min"]), khz(rec["max"]))

    # -- df2_max: 1DH5 only, value buried in Notes text ("...must be > 115 kHz") --
    rec = _find(bdr_edr, "BDR", "Δf2 (max)", cond_is_none=True)
    if rec:
        m = NOTES_KHZ_RE.search(rec["notes"])
        if m:
            spec[("df2_max", "1DH5", None)] = (float(m.group(1)) * 1000.0, None)

    # -- omega_i / omega_o / omega_io: EDR (3DH5) from BDR&EDR sheet, HDR (4DH5, 8DH5) shared row from HDR sheet --
    edr_omega = {
        "omega_i": _find(bdr_edr, "EDR", "TX Carrier Frequency Stability", cond_exact="ωi  (GFSK header)"),
        "omega_o": _find(bdr_edr, "EDR", "TX Carrier Frequency Stability", cond_exact="ω0"),
        "omega_io": _find(bdr_edr, "EDR", "TX Carrier Frequency Stability", cond_exact="ωi + ω0"),
    }
    for key, rec in edr_omega.items():
        if rec:
            spec[(key, "3DH5", None)] = (khz(rec["min"]), khz(rec["max"]))

    hdr_omega = {
        "omega_i": _find(hdr, "HDR", "TX Carrier Frequency Stability", cond_exact="wi (GFSK header)"),
        "omega_o": _find(hdr, "HDR", "TX Carrier Frequency Stability", cond_exact="w0"),
        "omega_io": _find(hdr, "HDR", "TX Carrier Frequency Stability", cond_exact="wi + w0"),
    }
    for key, rec in hdr_omega.items():
        if rec:
            for packet in ("4DH5", "8DH5"):
                spec[(key, packet, None)] = (khz(rec["min"]), khz(rec["max"]))

    return spec


def lookup(spec: dict, metric_key: str, packet_type: str, band: float | None = None):
    """Return (min_hz, max_hz) or None. Band-independent metrics are stored
    under band=None; band-dependent ones (currently only df1_avg) require
    the caller's band."""
    if (metric_key, packet_type, None) in spec:
        return spec[(metric_key, packet_type, None)]
    return spec.get((metric_key, packet_type, band))


if __name__ == "__main__":
    spec = parse_spec(Path("NPI_TestCoverage_v0.4.xlsx"))
    for key in sorted(spec, key=lambda k: (k[0], k[1], k[2] or 0)):
        print(key, "->", spec[key])
    print()
    print("lookup(df1_avg, 1DH5, 2.4) ->", lookup(spec, "df1_avg", "1DH5", 2.4))
    print("lookup(df1_avg, 1DH5, 5) ->", lookup(spec, "df1_avg", "1DH5", 5))
    print("lookup(df2_max, 1DH5) ->", lookup(spec, "df2_max", "1DH5"))
    print("lookup(cfo, LE1M) ->", lookup(spec, "cfo", "LE1M"))
    print("lookup(omega_i, 4DH5) ->", lookup(spec, "omega_i", "4DH5"))
    print("lookup(omega_i, 3DH5) ->", lookup(spec, "omega_i", "3DH5"))
