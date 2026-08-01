# PIVOTBY

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Pivots data by row and column keys. Aggregator is text (SUM, COUNT, …).

## Syntax

```
=PIVOTBY(row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order], [col_total_depth], [col_sort_order], [filter_array], [relative_to])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| row_fields | Required | row key column |
| col_fields | Required | column key column |
| values | Required | values to aggregate |
| function | Required | aggregator text, e.g. "SUM" |
| field_headers | Optional | 0 none/hide, 1 has/hide, 2 none/show, 3 has/show (default auto) |
| row_total_depth | Optional | 0 none, 1 grand total (default); negative = at top |
| row_sort_order | Optional | row sort: column number, negative descending |
| col_total_depth | Optional | 0 none, 1 total column (default); negative = at left |
| col_sort_order | Optional | column sort: 1 ascending, -1 descending |
| filter_array | Optional | TRUE/FALSE per row to include |
| relative_to | Optional | PERCENTOF base: 0 grand (default), 1 row totals, 2 column totals |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=PIVOTBY(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `PIVOTBY` (drop-in) on hosts without the native function, and as `EG.PIVOTBY` on modern Excel that has it.
