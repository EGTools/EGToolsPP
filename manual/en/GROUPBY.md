# GROUPBY

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Groups rows by key(s) and aggregates values. Aggregator is text: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF.

## Syntax

```
=GROUPBY(row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array], [field_relationship])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| row_fields | Required | key column(s) |
| values | Required | value column(s) to aggregate |
| function | Required | aggregator text: one of SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF |
| field_headers | Optional | 0 none/hide, 1 has/hide, 2 none/show generated, 3 has/show (default auto) |
| total_depth | Optional | 0 none, 1 grand total (default), 2 grand+subtotals; negative = at top |
| sort_order | Optional | output column number(s), negative descending, e.g. {2,-1} |
| filter_array | Optional | TRUE/FALSE per row to include |
| field_relationship | Optional | accepted for compatibility, ignored |

## Returns

Returns a spilled 2-D array of group key columns plus aggregated value columns. Missing required arguments, mismatched key/value row counts, or invalid aggregator/option values give #VALUE!; if no rows pass the filter the result is #N/A; AVERAGE/PERCENTOF cells with a zero denominator show #DIV/0!.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=GROUPBY({"a";"b";"a"},{10;20;30},"SUM")` | {"a",40;"b",20;"Total",60} | sum per key plus grand total row |
| `=GROUPBY({"a";"b";"a"},{10;20;30},"COUNT",0,0)` | {"a",2;"b",1} | count per key, no totals |

## Notes

- The function argument is text instead of the native eta-lambda; the supported aggregators are SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF (16, same set as native) (anything else returns #VALUE!). field_relationship is accepted for compatibility but ignored.
- Total/subtotal labels follow the UI language (Korean UI: 합계/총합계, English UI: Total/Grand Total). Subtotal rows are labeled with the first key value.
- When field_headers is omitted, a header row is auto-detected if the first row is all text and non-text values appear below it.
- Supported: Excel 2010+. Registered as `GROUPBY` (drop-in) on hosts without the native function, and as `EG.GROUPBY` on modern Excel that has it.
