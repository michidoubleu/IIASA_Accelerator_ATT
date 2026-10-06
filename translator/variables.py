"""Variables sheet <-> map_variable.

Layout: Variable | Description | Unit | <one column per item code...>
A cell with any truthy value (e.g. "X") under an item column means that
variable is reported for that item.

NOTE: as of 2026-10, the example workbook's item columns were human-friendly
labels used by researchers before they started reporting (not a synced data
source), and many of them don't correspond to any current item code. The
first `write_variables` call rebuilds the item-column headers to use the real
item codes from map_item, and fills the X-marks straight from each variable's
`items` list in the JSON - the matrix stops being authoritative input and
becomes a derived, human-readable view of the JSON. From then on it round-trips
normally via `read_variables`.
"""

from .sheet_utils import find_header_row


def _locate(ws):
    hdr_row, var_col = find_header_row(ws, ["Variable", "Description", "Unit"])
    desc_col, unit_col = var_col + 1, var_col + 2
    item_col_start = unit_col + 1
    return hdr_row, var_col, desc_col, unit_col, item_col_start


def _read_item_headers(ws, hdr_row, item_col_start):
    headers = {}
    c = item_col_start
    while True:
        v = ws.cell(row=hdr_row, column=c).value
        if not v:
            break
        headers[c] = str(v).strip().lower()
        c += 1
    return headers


def read_variables(ws, known_item_codes, warnings):
    hdr_row, var_col, desc_col, unit_col, item_col_start = _locate(ws)
    item_headers = _read_item_headers(ws, hdr_row, item_col_start)
    item_matrix_ready = any(code in known_item_codes for code in item_headers.values())

    map_variable = {}
    rows = {}
    for r in range(hdr_row + 1, ws.max_row + 1):
        name = ws.cell(row=r, column=var_col).value
        if not name:
            continue
        key = str(name).strip().lower()
        desc = ws.cell(row=r, column=desc_col).value or ""
        unit = ws.cell(row=r, column=unit_col).value or ""
        items = []
        for c, item_code in item_headers.items():
            val = ws.cell(row=r, column=c).value
            if val and str(val).strip():
                if item_code not in known_item_codes:
                    if item_matrix_ready:
                        warnings.append(
                            f"map_variable: row '{key}' marks unknown item column '{item_code}' (col {c}); ignored"
                        )
                    continue
                items.append(item_code)
        map_variable[key] = {
            "name": key,
            "unit": str(unit).strip().lower(),
            "items": items,
            "description": str(desc).strip(),
        }
        rows[key] = r

    meta = {"hdr_row": hdr_row, "var_col": var_col, "desc_col": desc_col, "unit_col": unit_col,
             "item_col_start": item_col_start, "item_headers": item_headers, "rows": rows,
             "item_matrix_ready": item_matrix_ready}
    return map_variable, meta


def write_variables(ws, map_variable, item_order, units_display, warnings):
    hdr_row, var_col, desc_col, unit_col, item_col_start = _locate(ws)
    old_headers = _read_item_headers(ws, hdr_row, item_col_start)
    old_rows = {}
    for r in range(hdr_row + 1, ws.max_row + 1):
        name = ws.cell(row=r, column=var_col).value
        if name:
            old_rows[str(name).strip().lower()] = r

    headers_already_current = list(old_headers.values()) == item_order
    if not headers_already_current:
        # Rebuild the item-code header columns.
        for i, code in enumerate(item_order):
            ws.cell(row=hdr_row, column=item_col_start + i).value = code
        tail_col = item_col_start + len(item_order)
        c = tail_col
        while ws.cell(row=hdr_row, column=c).value:
            ws.cell(row=hdr_row, column=c).value = None
            for r in range(hdr_row + 1, ws.max_row + 1):
                ws.cell(row=r, column=c).value = None
            c += 1
        if old_headers:
            warnings.append(
                f"map_variable: rebuilt {len(old_headers)} item column header(s) to use item codes "
                f"(was: friendly labels, or item order/set changed) - previous X-marks were discarded "
                f"and replaced from the JSON"
            )

    code_to_col = {code: item_col_start + i for i, code in enumerate(item_order)}
    next_row = ws.max_row + 1
    seen = set()
    for key, data in map_variable.items():
        seen.add(key)
        r = old_rows.get(key)
        exogenous = data.get("exogenous", False)

        if r is None and exogenous:
            # No row exists and nothing explicitly asked for one - stays hidden.
            continue
        if r is None and "exogenous" not in data:
            # Brand new variable, never seen in this sheet, no explicit decision made:
            # default to hidden/exogenous rather than silently adding a row.
            warnings.append(
                f"map_variable: '{key}' has no row in the Variables sheet and no explicit "
                f"'exogenous' flag; treated as exogenous (no row added) - set exogenous: false "
                f"in the JSON if it should be added to the workbook"
            )
            continue
        if r is not None and exogenous:
            # Previously tracked, now explicitly marked exogenous - hide it.
            ws.cell(row=r, column=var_col).value = None
            ws.cell(row=r, column=desc_col).value = None
            ws.cell(row=r, column=unit_col).value = None
            for col in code_to_col.values():
                ws.cell(row=r, column=col).value = None
            warnings.append(f"map_variable: '{key}' marked exogenous; cleared its existing row {r}")
            continue

        if r is None:
            r = next_row
            next_row += 1
            warnings.append(f"map_variable: added new row for '{key}' at row {r}")
        ws.cell(row=r, column=var_col).value = key.upper()
        ws.cell(row=r, column=desc_col).value = data["description"]
        ws.cell(row=r, column=unit_col).value = units_display.get(data["unit"], data["unit"])
        item_set = set(data["items"])
        for code, col in code_to_col.items():
            ws.cell(row=r, column=col).value = "X" if code in item_set else None

    for key, r in old_rows.items():
        if key not in seen:
            warnings.append(
                f"map_variable: variable '{key}' exists in the Variables sheet (row {r}) but not in the JSON; left untouched"
            )
