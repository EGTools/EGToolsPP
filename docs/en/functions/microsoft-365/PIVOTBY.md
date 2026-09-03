# PIVOTBY

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365

Pivots data by row and column keys. Aggregator is text (SUM, COUNT, …).

## Syntax

```
=PIVOTBY(row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order], [col_total_depth], [col_sort_order], [filter_array], [relative_to])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| row_fields | Required | row key column(s) |
| col_fields | Required | column key column(s) |
| values | Required | value column(s) to aggregate |
| function | Required | aggregator text: one of SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + extensions (param via relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) |
| field_headers | Optional | 0 none/hide, 1 has/hide, 2 none/show, 3 has/show (default auto) |
| row_total_depth | Optional | 0 none, 1 grand total (default), 2 grand+subtotals; negative = at top |
| row_sort_order | Optional | row sort: column number, negative descending |
| col_total_depth | Optional | 0 none, 1 total column (default), 2 total+subtotal columns; negative = at left |
| col_sort_order | Optional | column sort: ±field number(s), negative descending |
| filter_array | Optional | TRUE/FALSE per row to include |
| relative_to | Optional | PERCENTOF base 0-2 (0 grand default); for extension aggregators the parameter: TEXTJOIN delimiter, LARGE/SMALL k, PERCENTILE p (0-1), QUARTILE 0-4 |

## Returns

Returns a spilled 2-D cross-tab with one header row per column field, one row per row-key combination, and a total row/column by default; row_fields, col_fields and values each accept multiple columns, as in native. #VALUE! for an invalid aggregator or option value, or PERCENTOF relative_to outside 0-2; #N/A if no rows pass the filter; empty intersections become empty strings, also in subtotal rows/columns.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","Total";"a",10,30,40;"b",20,"",20;"Total",30,30,60} | pivot with row/column totals |
| `=PIVOTBY({"a";"a";"b"},{"x","p";"x","q";"y","p"},{10;20;30},"SUM")` | {"","x","x","y","Total";"","p","q","p","";"a",10,20,"",30;"b","","",30,30;"Total",10,20,30,60} | two column fields → two header rows |

## Notes

- The function argument is text instead of the native eta-lambda; the supported aggregators are SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF (16, same set as native), plus EGTools extensions TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) - the extension parameter (TEXTJOIN delimiter, LARGE/SMALL k, PERCENTILE p, QUARTILE 0-4) is passed via relative_to (any other name returns #VALUE!).
- row_fields, col_fields and values each accept multiple columns, as in native: each column field adds one header row, and each column group nests one value column per values column. Even with two or more value columns the data cells of the total/subtotal columns are filled, aggregated per value column - native leaves these cells blank, so this is the only part that differs from native output (an intentional difference). With field_headers 2 or 3 the header block is: a row of the column field names joined with ", ", the column-key rows, then a row of the row field names and value names.
- row_total_depth/col_total_depth 2 adds subtotal rows/columns per block of the first row/column field (the axis needs 2+ fields, otherwise it acts as 1); negative values place totals at the top/left. row_sort_order/col_sort_order take ±field number(s), also as an array; sorting always keeps the field hierarchy order and only sets the sort direction of the specified field.
- For PERCENTOF, this implementation supports relative_to 0 (column totals, default) / 1 (row totals) / 2 (grand total) only; 3 and 4 (parent totals) return #VALUE!. Total row/column labels follow the UI language (Korean UI: 합계/총합계, English UI: Total/Grand Total).
- Supported: Excel 2010+. Registered as `PIVOTBY` (drop-in) on hosts without the native function, and as `EG.PIVOTBY` on modern Excel that has it.
