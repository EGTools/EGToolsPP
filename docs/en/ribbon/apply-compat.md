# Apply Legacy Compatibility

**Location**: ribbon `EGTools` tab → `EGTools` group → split-button dropdown menu

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
| Regular compatibility functions | `_xlfn.F` / `_xlfn._xlws.F` / `EG.F` → `F` (EGTools drop-in name). The conflicting names SORT/FILTER/LET convert to `xSORT`/`xFILTER`/`xLET` |
| IMAGE | `_xlfn.IMAGE` / `EG.IMAGE` → `IMAGE` |
| GROUPBY / PIVOTBY | Function name goes bare and the aggregator argument gains quotes (`SUM` → `"SUM"`; only for whitelisted aggregators with matching positions — otherwise the formula is left untouched) |
| `_xlpm.*` (LET/LAMBDA parameters) | Modern Excel's internal storage prefix for LET/LAMBDA **parameter names**; on legacy hosts it surfaces as `=_xlfn.LET(_xlpm.x, …)`. Excel refuses to write back any formula containing `_xlpm.`, so the prefix is stripped (`_xlpm.x` → `x`). Counted as `_xlpm.*` in the conversion summary |
| `_xll.*` / `_xludf.*` (add-in UDF storage prefixes) | Strips the prefixes left in workbooks (notably in **defined names**) that were opened on a host where the EGTools UDF was not registered, re-pointing the call at the registered EGTools function (`_xll.xSORT(…)` → `xSORT(…)`). Other add-ins' UDFs are left alone. Counted as `_xll.*` in the conversion summary |
| `_xlfn._ARRAY` / `_xlfn._ARRAYROW` (in-cell array wrappers) | Internal functions the Microsoft 365 beta stores when it packs a spilled result into a single cell. On hosts without the native wrappers they evaluate to `#NAME?`, so they are rewritten to `ARRAYTOTEXT(…,1)`, which yields the same values in text notation: `_xlfn._ARRAYROW(a)` → `a`, `_xlfn._ARRAYROW(a,b,…)` → `HSTACK(a,b,…)`, `_xlfn._ARRAY(r)` → `ARRAYTOTEXT(r,1)`, `_xlfn._ARRAY(r1,r2,…)` → `ARRAYTOTEXT(VSTACK(r1,r2,…),1)`. Example: `=_xlfn._ARRAY(_xlfn._ARRAYROW(FILTER(A:A,B:B=1)))` → `=ARRAYTOTEXT(FILTER(A:A,B:B=1),1)`. The resulting text matches the strict notation of [List/Array Cells to Text](list-to-text.md) (`{"a",1;"b",2}`), so HAS/HASANY/HASALL/FLATTEN read it as is. Processed recursively from the inside out; a nested `_ARRAY` becomes a quoted text item inside the outer text. Counted as `_ARRAY`/`_ARRAYROW` in the conversion summary |
| Staged-rollout functions (HAS/HASANY/HASALL/FLATTEN, IMPORTTEXT/IMPORTCSV) | If this Excel really lacks the native function (probed at run time), `_xlfn.HAS` / bare `HAS` are rewritten to **the EG. name registered on this host**, e.g. `EG.HAS` (prevents `#NAME?` on non-beta Microsoft 365 channels). On legacy versions (2024 and earlier) the target stays the bare name (`HAS`) as before. A document converted to EG. names gets bare names when it is converted again on a legacy version |

String literals and partial identifier matches are never modified (whole-identifier replacement).

## Notes

- To undo, run [Restore Native Functions](restore-native.md) or open the backup file.
- The conversion-target list shares a single source with the add-in's function registry, so they always match.
- List/array cells created in the beta are not handled by this command. Run
  [List/Array Cells to Text](list-to-text.md) **first** and then this command, so that formulas such as
  HAS that reference a list cell do not become `#VALUE!`.
