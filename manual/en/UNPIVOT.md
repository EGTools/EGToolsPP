# UNPIVOT

**Category**: EGTools-only function

Converts a cross-tab (pivoted) table into a normalized data table.

## Syntax

```
=UNPIVOT(array, [fixed_col], [fixed_row], [sort_order], [col_names], [has_subset])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | source table (headers included) |
| fixed_col | Optional | row-header columns on the left (default 1) |
| fixed_row | Optional | column-header rows on top (default 1) |
| sort_order | Optional | result sort: column number, negative for descending, e.g. {1,-3} |
| col_names | Optional | titles for the result columns |
| has_subset | Optional | TRUE if value columns repeat as sets |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=UNPIVOT(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `UNPIVOT` on every Excel version (no `EG.`/`x` prefix).
