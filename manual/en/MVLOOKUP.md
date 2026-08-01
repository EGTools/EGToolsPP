# MVLOOKUP

**Category**: EGTools-only function

Runs VLOOKUP for many keys at once, returning one or more columns.

## Syntax

```
=MVLOOKUP(lookup_value, lookup_array, return_column, [match_mode], [if_not_found])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| lookup_value | Required | key column (or single value) |
| lookup_array | Required | data table; keys in first column |
| return_column | Required | column number(s) to return, e.g. 2 or {2,3} |
| match_mode | Optional | FALSE exact (default), TRUE approximate (largest ≤ key) |
| if_not_found | Optional | value when no match (default "") |

## Returns

Returns a spilled array of key-rows x return-columns; rows without a match get if_not_found (default ""). Returns #VALUE! when a required argument is missing, lookup_array is not an array, lookup_value has more than one column, or a return column number is invalid.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=MVLOOKUP({"a";"c"},{"a",1;"b",2},2)` | {1;""} | Exact match for many keys |
| `=MVLOOKUP("b",{"a",1,"x";"b",2,"y"},{2,3})` | {2,"y"} | Return several columns at once |
| `=MVLOOKUP(15,{10,"low";20,"high"},2,TRUE)` | low | Approximate match (largest key <= value) |

## Notes

- Exact match is hash-based; with duplicate keys the first row wins.
- Approximate match sorts internally and uses binary search, so the data need not be pre-sorted.
- Key comparison is case-insensitive.
- Supported: Excel 2010+. Always registered as `MVLOOKUP` on every Excel version.
