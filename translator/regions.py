"""Regions sheet <-> map_region (+ map_region_group).

The 'Regions' sheet packs three region lists into one column, back to back,
separated by section marker rows:
    Code | Region | Details | Comment      <- header, starts the "non eu" list
    ...
    European Union                          <- marker, starts the "eu" list
    ...
    NUTS2                                   <- marker, starts the "nuts2" list
    ...

For "non eu"/"eu" rows: name = Region column, description = Details column
(falls back to Region if Details is blank).

For "nuts2" rows only the bare code (e.g. BE10) was ever filled in - there is no
name at all. We derive one as "{code}--{Country}", looking up Country from the
ISO2 table that already lives further right on the same sheet (Country/ISO2/
ISO3/ISO#/AggCode), via the code's first two letters. The looked-up country name
is written back into the Region column for that row, so the translation link is
visible in the template itself rather than hidden in code.
"""

from .sheet_utils import find_header_row, find_marker_row

GROUPS = ("non eu", "eu", "nuts2")

# Eurostat's NUTS classification uses two codes that differ from ISO 3166-1
# alpha-2: Greece is "EL" (ISO2 "GR") and the United Kingdom is "UK" (ISO2 "GB").
NUTS_TO_ISO2_OVERRIDES = {"EL": "GR", "UK": "GB"}


def _locate(ws):
    hdr_row, code_col = find_header_row(ws, ["Code", "Region", "Details", "Comment"])
    name_col, detail_col, comment_col = code_col + 1, code_col + 2, code_col + 3
    iso_hdr_row, iso_country_col = find_header_row(ws, ["Country", "ISO2", "ISO3", "ISO#", "AggCode"])
    iso2_col = iso_country_col + 1
    return {
        "hdr_row": hdr_row,
        "code_col": code_col,
        "name_col": name_col,
        "detail_col": detail_col,
        "comment_col": comment_col,
        "iso_hdr_row": iso_hdr_row,
        "iso_country_col": iso_country_col,
        "iso2_col": iso2_col,
    }


def _iso_lookup(ws, loc):
    lookup = {}
    for r in range(loc["iso_hdr_row"] + 1, ws.max_row + 1):
        country = ws.cell(row=r, column=loc["iso_country_col"]).value
        iso2 = ws.cell(row=r, column=loc["iso2_col"]).value
        if country and iso2:
            lookup[str(iso2).strip().upper()] = str(country).strip()
    return lookup


def _section_bounds(ws, loc):
    eu_marker_row = find_marker_row(ws, loc["code_col"], "European Union")
    nuts2_marker_row = find_marker_row(ws, loc["code_col"], "NUTS2")
    return {
        "non eu": (loc["hdr_row"] + 1, eu_marker_row - 1),
        "eu": (eu_marker_row + 1, nuts2_marker_row - 1),
        "nuts2": (nuts2_marker_row + 1, ws.max_row),
    }, eu_marker_row, nuts2_marker_row


def read_regions(ws, warnings):
    loc = _locate(ws)
    iso_lookup = _iso_lookup(ws, loc)
    bounds, eu_marker_row, nuts2_marker_row = _section_bounds(ws, loc)

    map_region = {}
    region_rows = {}
    for group, (r_start, r_end) in bounds.items():
        for r in range(r_start, r_end + 1):
            code = ws.cell(row=r, column=loc["code_col"]).value
            if not code:
                continue
            if not isinstance(code, str):
                warnings.append(
                    f"map_region: cell at row {r} in the Regions sheet is not text ({code!r}, "
                    f"likely a typo that Excel auto-converted); skipped"
                )
                continue
            key = str(code).strip().lower()
            if group == "nuts2":
                iso2 = key[:2].upper()
                iso2 = NUTS_TO_ISO2_OVERRIDES.get(iso2, iso2)
                country = iso_lookup.get(iso2)
                if not country:
                    warnings.append(
                        f"map_region: no ISO country match for NUTS2 code '{code}' "
                        f"(prefix '{key[:2].upper()}'); keeping existing Region cell as fallback"
                    )
                    country = ws.cell(row=r, column=loc["name_col"]).value or "UNKNOWN"
                name = f"{str(code).strip()}--{country}"
                description = name
            else:
                name = str(ws.cell(row=r, column=loc["name_col"]).value or "").strip()
                detail = ws.cell(row=r, column=loc["detail_col"]).value
                description = str(detail).strip() if detail else name

            if key in map_region:
                warnings.append(f"map_region: duplicate code '{key}' at row {r}; keeping first occurrence")
                continue
            map_region[key] = {"name": name, "description": description, "region_group": group}
            region_rows[key] = r

    meta = {"loc": loc, "iso_lookup": iso_lookup, "region_rows": region_rows,
             "eu_marker_row": eu_marker_row, "nuts2_marker_row": nuts2_marker_row}
    return map_region, meta


def write_regions(ws, map_region, warnings):
    _, meta = read_regions(ws, [])
    loc = meta["loc"]
    region_rows = dict(meta["region_rows"])

    for key in region_rows:
        if key not in map_region:
            warnings.append(
                f"map_region: code '{key}' exists in the Regions sheet but not in the JSON; left untouched"
            )

    by_group = {g: [] for g in GROUPS}
    for key, data in map_region.items():
        by_group.setdefault(data["region_group"], []).append((key, data))

    for group in GROUPS:
        for key, data in by_group.get(group, []):
            r = region_rows.get(key)
            if r is None:
                r = _insert_row(ws, loc, meta, group)
                region_rows[key] = r
                warnings.append(f"map_region: added new row for code '{key}' ({group}) at row {r}")
            ws.cell(row=r, column=loc["code_col"]).value = key.upper()
            if group == "nuts2":
                country = data["name"].split("--", 1)[-1] if "--" in data["name"] else data["name"]
                ws.cell(row=r, column=loc["name_col"]).value = country
            else:
                ws.cell(row=r, column=loc["name_col"]).value = data["name"]
                if data["description"] and data["description"] != data["name"]:
                    ws.cell(row=r, column=loc["detail_col"]).value = data["description"]


def _insert_row(ws, loc, meta, group):
    """Insert a fresh row for a brand-new region code, right before the next
    section marker (or at the end of the sheet for the last section)."""
    if group == "non eu":
        insert_at = find_marker_row(ws, loc["code_col"], "European Union")
    elif group == "eu":
        insert_at = find_marker_row(ws, loc["code_col"], "NUTS2")
    else:
        insert_at = ws.max_row + 1
        return insert_at
    ws.insert_rows(insert_at)
    return insert_at
