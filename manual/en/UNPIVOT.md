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

Returns the normalized table, including the title row, as a 2-D spilled array. Returns #VALUE! when the input is not an array, fixed_col/fixed_row are negative, out of range, or both 0, or a sort_order entry is 0 or outside the result columns.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=UNPIVOT({"Item","Jan","Feb";"A",10,20;"B",30,40},1,1,,{"Item","Month","Qty"})` | {"Item","Month","Qty";"A","Jan",10;"A","Feb",20;"B","Jan",30;"B","Feb",40} | Unpivot a cross-tab |
| `=UNPIVOT({"Item","Jan","Feb";"A",10,20;"B",30,40},1,1,-3,{"Item","Month","Qty"})` | {"Item","Month","Qty";"B","Feb",40;"B","Jan",30;"A","Feb",20;"A","Jan",10} | Sort by column 3 descending |

## Notes

- Works on values only; filling merged cells is not supported.
- Entries whose value cells are all empty are dropped from the result.
- When col_names is omitted, default titles depend on the UI language (Korean/English).
- Supported: Excel 2010+. Always registered as `UNPIVOT` on every Excel version.
