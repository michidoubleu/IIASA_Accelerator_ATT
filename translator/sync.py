"""Top-level two-way sync between the JSON validator template and the Excel workbook.

Static JSON keys that don't correspond to anything in the workbook (the JSON-schema
plumbing used by the online validator: root, mapschema_*, template_validators,
enum_map_dependencies, root_schema_declarations), and map_region_group (which carries
a `region_boundry_file_link` that isn't in the spreadsheet either), are never derived
from the spreadsheet. If no base JSON is given, a bundled default is used for them
instead - see `default_schema.json`.

Variables with no row in the Variables sheet are "exogenous": they're kept in the
JSON (tagged `exogenous: true`) but never get a row written back into the workbook,
since they aren't part of what researchers currently report there. A variable stops
being exogenous as soon as a row for it appears in the sheet again.
"""

from pathlib import Path

import openpyxl

from . import items, regions, units, variables
from .json_io import load_json, save_json

STATIC_JSON_KEYS = (
    "root",
    "mapschema_item",
    "mapschema_region",
    "mapschema_variable",
    "mapschema_region_group",
    "template_validators",
    "enum_map_dependencies",
    "root_schema_declarations",
)

DEFAULT_SCHEMA_PATH = Path(__file__).parent / "default_schema.json"


def excel_to_json(xlsx_path, base_json_path, out_json_path):
    """Read the workbook and produce an updated JSON. `base_json_path` may be None,
    in which case the bundled default schema/validator config is used instead."""
    warnings = []
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if base_json_path:
        base = load_json(base_json_path)
    else:
        base = load_json(DEFAULT_SCHEMA_PATH)
        warnings.append(
            "no base JSON given: used the bundled default for the validator schema "
            "parts (root, mapschema_*, template_validators, enum_map_dependencies, "
            "root_schema_declarations) and for map_region_group - review these if your "
            "setup differs from the bundled defaults"
        )

    map_item, _ = items.read_items(wb["Items"], warnings)
    map_region, _ = regions.read_regions(wb["Regions"], warnings)
    map_variable, var_meta = variables.read_variables(wb["Variables"], set(map_item.keys()), warnings)

    base_vars = base.get("map_variable", {})
    if var_meta["item_matrix_ready"]:
        map_variable_out = dict(map_variable)
        exogenous_kept = 0
        for key, data in base_vars.items():
            if key not in map_variable_out:
                data = dict(data)
                data["exogenous"] = True
                map_variable_out[key] = data
                exogenous_kept += 1
        if exogenous_kept:
            warnings.append(
                f"map_variable: {exogenous_kept} variable(s) from the base JSON have no row in the "
                f"Variables sheet; kept them tagged as 'exogenous: true' (they won't get a row when "
                f"writing back to Excel until you add one for them yourself)"
            )
    elif not base_vars:
        map_variable_out = map_variable
        warnings.append(
            "map_variable: no base map_variable to fall back to and the Variables sheet's item "
            "columns are not item codes yet, so items lists are likely empty for everything - "
            "run to-excel once to normalize the sheet, then re-run to-json"
        )
    else:
        added = sorted(set(map_variable) - set(base_vars))
        missing = sorted(set(base_vars) - set(map_variable))
        warnings.append(
            "map_variable: the Variables sheet's item columns are not yet item codes "
            "(run to-excel once to normalize them); the whole sheet is being treated as "
            "a non-authoritative draft for now, so the base JSON's map_variable is kept as-is. "
            f"{len(added)} sheet-only row(s) ignored ({added[:10]}{'...' if len(added) > 10 else ''}), "
            f"{len(missing)} JSON-only variable(s) would otherwise have been dropped ({missing[:10]}{'...' if len(missing) > 10 else ''})"
        )
        map_variable_out = base_vars

    out = dict(base)
    out["map_item"] = map_item
    out["map_region"] = map_region
    out["map_variable"] = map_variable_out

    mrg = dict(base.get("map_region_group", {}))
    groups_used = {v["region_group"] for v in map_region.values()}
    for g in groups_used:
        if g not in mrg:
            mrg[g] = {"name": g, "description": g, "region_boundry_file_link": ""}
            warnings.append(f"map_region_group: new group '{g}' found in Regions sheet; added placeholder entry")
    out["map_region_group"] = mrg

    for key in STATIC_JSON_KEYS:
        out[key] = base.get(key)

    save_json(out_json_path, out)
    return warnings


def json_to_excel(json_path, base_xlsx_path, out_xlsx_path):
    """Load the workbook as a template and patch in the content from the JSON,
    leaving every other sheet/column/row untouched."""
    warnings = []
    data = load_json(json_path)
    wb = openpyxl.load_workbook(base_xlsx_path)

    items.write_items(wb["Items"], data["map_item"], warnings)

    regions.write_regions(wb["Regions"], data["map_region"], warnings)

    item_order = list(data["map_item"].keys())
    units_used = {v["unit"] for v in data["map_variable"].values() if v.get("unit")}
    units_display = units.ensure_units(wb["Units"], units_used, warnings)
    variables.write_variables(wb["Variables"], data["map_variable"], item_order, units_display, warnings)

    wb.save(out_xlsx_path)
    return warnings
