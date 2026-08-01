# Apply Legacy Compatibility

**Location**: ribbon `EGTools++` tab → `EGTools` group

When a workbook authored in modern Excel is opened in an older version, new functions
appear as broken `_xlfn.XLOOKUP`-style tokens and evaluate to `#NAME?`. This button
batch-converts modern/native function tokens in every formula of the active workbook
to EGTools-compatible names, so the formulas keep calculating on legacy Excel.

## How it runs

1. **Saved-state check** — formulas are rewritten in place, so the workbook must be saved.
   If not, you are asked to save first (declining or cancelling aborts).
2. **Pre-scan** — all worksheet formulas and defined names are scanned read-only and a
   per-function count summary is shown. If nothing matches, it reports and exits.
3. **Backup** — on confirmation a backup copy is saved right before converting
   (same folder as the workbook, `name_egbak_<timestamp>.<ext>`). If the backup fails,
   nothing is converted.
4. **Convert & report** — shows how many locations were converted and the backup path.

## Conversion rules

| Target | Conversion |
|---|---|
| Regular compatibility functions | `_xlfn.F` / `_xlfn._xlws.F` / `EG.F` → `F` (EGTools drop-in name) |
| IMAGE | `_xlfn.IMAGE` / `EG.IMAGE` → `IMAGE` |
| GROUPBY / PIVOTBY | Function name goes bare and the aggregator argument gains quotes (`SUM` → `"SUM"`; only for whitelisted aggregators with matching positions — otherwise the formula is left untouched) |

String literals and partial identifier matches are never modified (whole-identifier replacement).

## Notes

- To undo, run [Restore Native Functions](restore-native.md) or open the backup file.
- The conversion-target list shares a single source with the add-in's function registry, so they always match.
