# AgMIP reporting template translator

Two-way sync between the reporting template workbook (`.xlsx`) that researchers
fill in and the JSON file used by the online validator/merge tool.

- **Workbook → JSON**: `cli.py to-json` reads a (possibly edited) workbook and
  produces an updated JSON.
- **JSON → Workbook**: `cli.py to-excel` patches a copy of the canonical
  workbook from a JSON, regenerating only the Items/Regions/Units/Variables
  sheets and leaving everything else untouched.

Requires Python with `openpyxl` installed (`pip install openpyxl`).

---

## For researchers: proposing changes to the template

You only ever need to touch the workbook — never the JSON.

1. **Download the current canonical workbook** from this repo
   (`examples/AgMIP_Reporting_template_example.xlsx`).
2. **Edit it** to propose your changes:
   - **New or renamed item**: add/edit a row in the `Items` sheet (`Code`,
     `Description` columns).
   - **New or renamed region**: add/edit a row in the `Regions` sheet, in the
     matching block (non-EU country list / EU country list / NUTS2 codes).
   - **New variable, or which items a variable is reported for**: add/edit a
     row in the `Variables` sheet. Put an `X` in an item's column if that
     variable is reported for that item.
   - Everything else in the workbook (`Overview crop definitions`,
     `Overview regional mapping`, `DM_Conversion_rates`, `Listing_template`)
     is reference material only — feel free to use it to check your work, but
     it isn't synced into the JSON.
3. **Send back your edited workbook** (e.g. as a pull request attaching the
   file, or to whoever maintains the template). Don't worry about the JSON or
   about getting the formatting perfect — the maintainer's tooling takes care
   of merging it in and will flag anything ambiguous.

A few things worth knowing about how your edits get interpreted:
- If you reuse an existing item/region code with a slightly different
  description, both descriptions are kept (combined), nothing is silently
  dropped.
- A variable with no row in the `Variables` sheet is treated as not currently
  reported ("exogenous") and is kept in the JSON but won't show up in future
  regenerated workbooks until it has a row again.

---

## For maintainers: turning a proposal into the canonical JSON + workbook

1. **Build an updated JSON from the proposed workbook:**
   ```
   python cli.py to-json <proposed.xlsx> <new.json> <current_canonical.json>
   ```
   The third argument (base JSON) is optional, but normally you want it: it
   carries forward the validator-schema config and the list of currently
   exogenous (untracked) variables. Omit it only if you're bootstrapping from
   scratch.

2. **Review `<new.json>.warnings.txt`.** It lists everything that needed a
   judgment call: merged duplicate codes, new codes/regions/variables found,
   codes that exist in the sheet but weren't in the JSON (left untouched),
   variables with no row (flagged exogenous), unresolvable NUTS2 country
   lookups, etc. Also diff `<new.json>` against the current canonical JSON to
   see the actual changes.

3. **Hand-edit `<new.json>` if needed** — e.g. to flip `"exogenous": false`
   on a variable you've decided should now be actively tracked, to fix a
   flagged data issue, or to adjust the validator-schema parts
   (`root`, `mapschema_*`, `template_validators`, `enum_map_dependencies`,
   `root_schema_declarations`, `map_region_group`) which the tool never
   touches on its own.

4. **Once you're happy with the JSON, make it the new canonical JSON**, and
   regenerate the canonical workbook from it so the two stay in sync for the
   next round of proposals:
   ```
   python cli.py to-excel <new_canonical.json> <current_canonical.xlsx> <new_canonical.xlsx>
   ```
   Use the *current* canonical workbook as the base — this preserves every
   other sheet, formatting, and reference table, and only patches the
   Items/Regions/Units/Variables content. Check `<new_canonical.xlsx>.warnings.txt`
   the same way.

5. **Commit both `<new_canonical.json>` and `<new_canonical.xlsx>`** as the
   new canonical versions and publish them.

### Command reference

```
python cli.py to-json   <xlsx> <out_json> [base_json]
python cli.py to-excel  <json> <base_xlsx> <out_xlsx>
```

Neither command overwrites its inputs — always writes to the given output
path, plus a `<output>.warnings.txt` log of everything worth a second look.
