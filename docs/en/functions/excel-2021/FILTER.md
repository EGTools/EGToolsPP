# FILTER

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

> ⚠️ On some legacy hosts it registers as `xFILTER` due to a keyword conflict.

Filters an array, keeping rows/columns where include is TRUE.

## Syntax

```
=FILTER(array, include, [if_empty])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | data to filter |
| include | Required | boolean array (rows or columns) |
| if_empty | Optional | value if nothing matches (optional) |

## Returns

Returns an array keeping only the rows or columns whose include value is truthy; on dynamic-array hosts it spills. Returns #VALUE! when include does not match the array's row or column count, and #N/A when nothing remains and if_empty is omitted.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=FILTER({1;2;3;4},{1;0;1;0})` | {1;3} | Filter rows |
| `=FILTER({1,2;3,4;5,6},{0;1;1})` | {3,4;5,6} | Filter rows of two columns |
| `=FILTER({1;2},{0;0},"none")` | none | Fallback when empty |

## Notes

- Text and blank values in include are treated as FALSE (native returns #VALUE!).
- With no matches and if_empty omitted, #N/A is returned instead of #CALC!.
- If include contains an error value, the whole result is that error.
- Related functions: SORT, UNIQUE.
- Supported: Excel 2010+. Registered as `FILTER` (drop-in) on hosts without the native function, and as `EG.FILTER` on modern Excel that has it.
