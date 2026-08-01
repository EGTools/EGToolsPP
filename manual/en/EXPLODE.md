# EXPLODE

**Category**: EGTools-only function

Splits one column by delimiter(s), expanding into rows or columns; other columns are copied.

## Syntax

```
=EXPLODE(array, col_number, [delimiter], [by_col], [pad_with], [ignore_empty])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | source table |
| col_number | Required | column to split, 1-based |
| delimiter | Optional | delimiter(s), e.g. " " or {",",";"} (default space) |
| by_col | Optional | TRUE spreads parts across columns (default FALSE: down rows) |
| pad_with | Optional | fill value for short rows when by_col (default "") |
| ignore_empty | Optional | TRUE drops empty parts |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=EXPLODE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `EXPLODE` on every Excel version (no `EG.`/`x` prefix).
