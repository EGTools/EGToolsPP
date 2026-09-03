# ADVANCEDFILTER

**Category**: EGTools-only function

Filters a table with an Advanced-Filter style criteria grid (rows=OR, columns=AND).

## Syntax

```
=ADVANCEDFILTER(data_array, filter_array, [unique])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| data_array | Required | data with header row |
| filter_array | Required | criteria grid with header row; supports =, <>, >, <, >=, <=, * and ? |
| unique | Optional | TRUE removes duplicate result rows |

## Returns

Returns the filtered rows, including the header row, as a 2-D spilled array. Returns #VALUE! when data/filter are not arrays with at least 2 rows or a filter header is not found in the data headers, and #N/A when no data row matches.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ADVANCEDFILTER({"Name","Qty";"A",10;"B",5;"A",7},{"Qty";">6"})` | {"Name","Qty";"A",10;"A",7} | Keep rows where Qty>6 |
| `=ADVANCEDFILTER({"Name","Qty";"Apple",10;"Banana",5;"Avocado",7},{"Name";"A*"})` | {"Name","Qty";"Apple",10;"Avocado",7} | Wildcard filter A* |

## Notes

- A text criterion without an operator is a "begins with" match, like Excel's Advanced Filter.
- Wildcards * and ? are supported with = and <> criteria.
- The data header row is always included in the result.
- Supported: Excel 2010+. Always registered as `ADVANCEDFILTER` on every Excel version.
