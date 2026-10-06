"""Units sheet: canonical display-cased unit strings.

There is no `map_unit` in the JSON - units only show up as the lowercase `unit`
field on each map_variable entry. This sheet is the source of the nicely-cased
display text (e.g. "bn USD 2005 MER") for a given lowercase unit (e.g.
"bn usd 2005 mer"), used when writing variables back into the workbook.
"""

from .sheet_utils import find_header_row


def _locate(ws):
    hdr_row, unit_col = find_header_row(ws, ["Units", "Description"])
    return hdr_row, unit_col, unit_col + 1


def read_units_display(ws):
    hdr_row, unit_col, desc_col = _locate(ws)
    display = {}
    rows_by_unit = {}
    for r in range(hdr_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=unit_col).value
        if v:
            key = str(v).strip().lower()
            display[key] = str(v).strip()
            rows_by_unit[key] = r
    return display, {"hdr_row": hdr_row, "unit_col": unit_col, "desc_col": desc_col, "rows_by_unit": rows_by_unit}


def ensure_units(ws, units_lower, warnings):
    """Make sure every unit in `units_lower` has a row in the Units sheet.

    Returns the (possibly updated) lowercase -> display-case lookup.
    """
    display, meta = read_units_display(ws)
    next_row = ws.max_row + 1
    for u in units_lower:
        if u not in display:
            ws.cell(row=next_row, column=meta["unit_col"]).value = u
            warnings.append(
                f"units: '{u}' was used by a variable but not listed in the Units sheet; "
                f"added a new row at {next_row} - please review its display casing/description"
            )
            display[u] = u
            next_row += 1
    return display
