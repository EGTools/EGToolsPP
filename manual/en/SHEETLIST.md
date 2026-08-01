# SHEETLIST

**Category**: EGTools-only function

Returns the visible sheet names of the calling workbook.

## Syntax

```
=SHEETLIST([start], [end], [strictly_between], [except_me])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| start | Optional | first sheet number or name (default: first) |
| end | Optional | last sheet number or name (default: last) |
| strictly_between | Optional | TRUE excludes the start/end sheets |
| except_me | Optional | TRUE excludes the calling sheet (default) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SHEETLIST(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `SHEETLIST` on every Excel version (no `EG.`/`x` prefix).
