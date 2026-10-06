#!/usr/bin/env python3
"""Two-way translator between the AgMIP reporting template workbook and its
JSON validator/mapping file.

Usage:
    python cli.py to-json   <workbook.xlsx> <base_template.json> <out_template.json>
    python cli.py to-excel  <template.json> <base_workbook.xlsx> <out_workbook.xlsx>

Both commands never overwrite their inputs; always write to the given output path.
A warnings log is printed to stdout and also written next to the output file as
<output>.warnings.txt.
"""

import argparse
import sys
from pathlib import Path

from translator import excel_to_json, json_to_excel


def _report(warnings, out_path):
    if warnings:
        print(f"\n{len(warnings)} warning(s):")
        for w in warnings:
            print(f"  - {w}")
    else:
        print("\nNo warnings.")
    log_path = Path(str(out_path) + ".warnings.txt")
    log_path.write_text("\n".join(warnings) + ("\n" if warnings else ""), encoding="utf-8")
    print(f"\nWarnings log: {log_path}")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    p_to_json = sub.add_parser("to-json", help="Build/update a JSON template from a workbook")
    p_to_json.add_argument("xlsx", help="Workbook to read updates from")
    p_to_json.add_argument("out_json", help="Path to write the updated JSON to")
    p_to_json.add_argument(
        "base_json", nargs="?", default=None,
        help="Existing JSON template to carry forward validator-schema parts and exogenous "
             "variables from. Omit to build from scratch using the bundled default schema.",
    )

    p_to_excel = sub.add_parser("to-excel", help="Update a workbook from a JSON template")
    p_to_excel.add_argument("json", help="JSON template to read from")
    p_to_excel.add_argument("base_xlsx", help="Workbook used as a template (other sheets/columns are preserved)")
    p_to_excel.add_argument("out_xlsx", help="Path to write the updated workbook to")

    args = parser.parse_args()

    if args.command == "to-json":
        warnings = excel_to_json(args.xlsx, args.base_json, args.out_json)
        print(f"Wrote {args.out_json}")
        _report(warnings, args.out_json)
    elif args.command == "to-excel":
        warnings = json_to_excel(args.json, args.base_xlsx, args.out_xlsx)
        print(f"Wrote {args.out_xlsx}")
        _report(warnings, args.out_xlsx)


if __name__ == "__main__":
    sys.exit(main())
