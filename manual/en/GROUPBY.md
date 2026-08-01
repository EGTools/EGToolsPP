# GROUPBY

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Groups rows by key(s) and aggregates values. Aggregator is text: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/PERCENTOF.

## Syntax

```
=GROUPBY(row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array], [field_relationship])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| row_fields | Required | key column(s) |
| values | Required | value column(s) to aggregate |
| function | Required | aggregator text, e.g. "SUM" |
| field_headers | Optional | 0 none/hide, 1 has/hide, 2 none/show generated, 3 has/show (default auto) |
| total_depth | Optional | 0 none, 1 grand total (default), 2 grand+subtotals; negative = at top |
| sort_order | Optional | output column number(s), negative descending, e.g. {2,-1} |
| filter_array | Optional | TRUE/FALSE per row to include |
| field_relationship | Optional | accepted for compatibility, ignored |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=GROUPBY(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `GROUPBY` (drop-in) on hosts without the native function, and as `EG.GROUPBY` on modern Excel that has it.
