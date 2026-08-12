# XLOOKUP

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Searches a range or array for a match and returns the corresponding item from a second range or array.

## Syntax

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| lookup_value | Required | the value to search for |
| lookup_array | Required | the range or array to search |
| return_array | Required | the range or array to return from |
| if_not_found | Optional | value to return if no match (optional) |
| match_mode | Optional | 0 exact (default), -1 next smaller, 1 next larger |
| search_mode | Optional | 1 first-to-last (default), -1 last-to-first |

## Returns

Returns the matched row (vertical search) or column (horizontal search) of return_array as an array; on dynamic-array hosts it spills. Returns #N/A when a range is omitted, the lookup array is empty, or nothing matches and if_not_found is omitted; #REF! when the matched position lies outside return_array; #VALUE! on internal errors.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=XLOOKUP("b",{"a";"b";"c"},{10;20;30})` | 20 | Exact match lookup |
| `=XLOOKUP(2,{1;2;3},{10,11;20,21;30,31})` | {20,21} | Whole matched row spills |
| `=XLOOKUP(9,{1;2;3},{10;20;30},"none")` | none | Fallback when not found |
| `=XLOOKUP({2;3},{1;2;3},{"a";"b";"c"})` | {b;c} | Array lookup_value → element-wise lookup |

## Notes

- match_mode 2 (wildcard) and search_mode 2/-2 (binary search) are not supported.
- A one-column lookup_array triggers vertical search; otherwise the first row is searched horizontally.
- If lookup_value is an array, each element is looked up and the result spills in the same shape as lookup_value — when the return range has multiple columns, each element's result is demoted to its first value (native-compatible), and an error lookup value returns that error.
- Related functions: XMATCH, FILTER.
- Supported: Excel 2010+. Registered as `XLOOKUP` (drop-in) on hosts without the native function, and as `EG.XLOOKUP` on modern Excel that has it.
