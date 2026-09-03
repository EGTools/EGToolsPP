# Restore Native Functions

**Location**: ribbon `EGTools` tab → `EGTools` group → split-button dropdown menu

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
| Regular compatibility functions | `EG.F` / `xF` (xSORT/xFILTER/xLET) → `F` (native). Cells still holding the legacy drop-in name are re-tokenised so the native function takes over |
| IMAGE | **Source-aware**: if the first argument is an `https://` source the call stays native `IMAGE`; anything else (local path, `http://`) converts to `EG.IMAGE`. A reference/expression argument is judged by **its value evaluated at conversion time** (evaluation failure, errors and non-strings safely fall back to `EG.IMAGE`). Native IMAGE is https-only, so if the referenced value later changes to a local path that cell shows `#VALUE!` — `EG.IMAGE` by contrast renders every source |
| GROUPBY / PIVOTBY | Function name reverts to native and the aggregator argument loses its quotes (`"SUM"` → `SUM`; only for whitelisted aggregators with matching positions — otherwise the formula is left untouched) |
| `_xlpm.*` (LET/LAMBDA parameters) | Strips the parameter storage prefix left behind by the legacy round-trip (`xLET(_xlpm.x, …)` → `LET(x, …)`). Excel refuses to write formulas containing `_xlpm.`, so without this strip the cell cannot be restored. Counted as `_xlpm.*` in the conversion summary |
| `@` (implicit intersection) | An `@` that modern Excel prepends when loading legacy files is removed when it sits **directly in front of a name being restored to native** (`TAKE(@xSORT(…),3)` → `TAKE(SORT(…),3)`). Left in place, `@SORT` would collapse the array result to a single value. `@` before cell references or non-target functions is preserved. Counted as `@(암시적 교차)` in the conversion summary |
| `_xll.*` / `_xludf.*` (add-in UDF storage prefixes) | Strips the prefixes Excel records when the workbook was opened on a host where the add-in UDF was not registered (`_xll.xSORT(…)` → `SORT(…)`). In particular this fixes **defined names**, whose formulas keep these prefixes and previously never converted. Only prefixes in front of EGTools functions are stripped — other add-ins' UDFs are left alone. Counted as `_xll.*` in the conversion summary |
| Staged-rollout functions (IMPORTTEXT/IMPORTCSV) | Functions introduced but not yet generally available (365 preview channels only) are **probed for actual native availability on this Excel** before restoring, and convert to native only on a positive result. Unconfirmed, they keep the EGTools implementation like IMAGE does (`IMPORTTEXT` → `EG.IMPORTTEXT`). Generally-available functions always convert without this probe |

## Notes

- If the host Excel lacks the native function (legacy versions), restored formulas become
  `#NAME?` — use this button when moving back to modern Excel.
- To convert for legacy Excel again, run [Apply Legacy Compatibility](apply-compat.md).
