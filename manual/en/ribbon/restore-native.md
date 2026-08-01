# Restore Native Functions

**Location**: ribbon `EGTools++` tab → `EGTools` group

When a workbook converted for EGTools compatibility moves back to modern Excel, this
button batch-restores EGTools function names (`EG.*` etc.) in every formula of the
active workbook to the native built-ins. After restoring, the formulas calculate
without the add-in.

## How it runs

Same safety flow as [Apply Legacy Compatibility](apply-compat.md):
saved-state check → read-only pre-scan (per-function count summary) → automatic backup
(`name_egbak_<timestamp>.<ext>`) → convert → result summary.

## Conversion rules

| Target | Conversion |
|---|---|
| Regular compatibility functions | `EG.F` → `F` (native). Cells still holding the legacy drop-in name are re-tokenised so the native function takes over |
| IMAGE | `IMAGE` → `EG.IMAGE` (**exception**: always keeps the EGTools implementation; genuine native `_xlfn.IMAGE` is left alone) |
| GROUPBY / PIVOTBY | Function name reverts to native and the aggregator argument loses its quotes (`"SUM"` → `SUM`; only for whitelisted aggregators with matching positions — otherwise the formula is left untouched) |

## Notes

- If the host Excel lacks the native function (legacy versions), restored formulas become
  `#NAME?` — use this button when moving back to modern Excel.
- To convert for legacy Excel again, run [Apply Legacy Compatibility](apply-compat.md).
