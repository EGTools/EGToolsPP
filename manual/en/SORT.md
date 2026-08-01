# SORT

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Sorts an array by a row/column.

## Syntax

```
=SORT(array, [sort_index], [sort_order], [by_col])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | data to sort |
| sort_index | Optional | column (or row) to sort by (default 1) |
| sort_order | Optional | 1 ascending (default), -1 descending |
| by_col | Optional | TRUE to sort columns instead of rows |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SORT(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `SORT` (drop-in) on hosts without the native function, and as `EG.SORT` on modern Excel that has it. On some legacy hosts it registers as `xSORT` due to a keyword conflict.
