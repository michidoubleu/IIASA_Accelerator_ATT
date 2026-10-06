"""Items sheet <-> map_item.

The 'Items' sheet lists one item per row as (Code, Description, ...extra reference
columns like GTAP/CAPRI/model mappings, which are left untouched). Some codes
legitimately appear on more than one row with two different descriptions (added by
different people at different times) - those are merged into a single map_item
entry with both descriptions combined, and the duplicate row is blanked out.
"""

from .sheet_utils import find_header_row


def _locate(ws):
    hdr_row, code_col = find_header_row(ws, ["Code", "Description"])
    desc_col = code_col + 1
    return hdr_row, code_col, desc_col


def read_items(ws, warnings):
    hdr_row, code_col, desc_col = _locate(ws)
    entries = {}  # key -> {"descs": [...], "rows": [...]}
    for r in range(hdr_row + 1, ws.max_row + 1):
        code = ws.cell(row=r, column=code_col).value
        if not code:
            continue
        key = str(code).strip().lower()
        desc_raw = ws.cell(row=r, column=desc_col).value
        desc = str(desc_raw).strip() if desc_raw else ""
        entry = entries.setdefault(key, {"descs": [], "rows": []})
        entry["rows"].append(r)
        if desc and desc not in entry["descs"]:
            entry["descs"].append(desc)

    map_item = {}
    for key, entry in entries.items():
        if len(entry["rows"]) > 1:
            warnings.append(
                f"map_item: merged {len(entry['rows'])} rows for code '{key}' "
                f"(rows {entry['rows']}) into one entry with combined description"
            )
        description = " / ".join(entry["descs"]) if entry["descs"] else ""
        map_item[key] = {"name": key, "description": description}

    meta = {"hdr_row": hdr_row, "code_col": code_col, "desc_col": desc_col, "entries": entries}
    return map_item, meta


def write_items(ws, map_item, warnings):
    _, meta = read_items(ws, [])
    code_col, desc_col, entries = meta["code_col"], meta["desc_col"], meta["entries"]

    for key in entries:
        if key not in map_item:
            warnings.append(
                f"map_item: code '{key}' exists in the Items sheet but not in the JSON; left untouched"
            )

    next_row = ws.max_row + 1
    for key, data in map_item.items():
        entry = entries.get(key)
        if entry:
            rows = entry["rows"]
            ws.cell(row=rows[0], column=code_col).value = key.upper()
            ws.cell(row=rows[0], column=desc_col).value = data["description"]
            for extra_row in rows[1:]:
                ws.cell(row=extra_row, column=code_col).value = None
                ws.cell(row=extra_row, column=desc_col).value = None
        else:
            ws.cell(row=next_row, column=code_col).value = key.upper()
            ws.cell(row=next_row, column=desc_col).value = data["description"]
            warnings.append(f"map_item: added new row for code '{key}' at row {next_row}")
            next_row += 1
