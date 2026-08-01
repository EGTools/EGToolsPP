# XMATCH

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Returns the position of a lookup value within an array.

## Syntax

```
=XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| lookup_value | Required | value to find |
| lookup_array | Required | array to search |
| match_mode | Optional | 0 exact (default), -1 next smaller, 1 next larger |
| search_mode | Optional | 1 first-to-last (default), -1 last-to-first |

## Returns

Returns the 1-based position as a numeric scalar. Returns #N/A when the array is empty or nothing matches, and #VALUE! on internal errors.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=XMATCH(3,{1;2;3;4})` | 3 | Exact match position |
| `=XMATCH(2.5,{1;2;3},1)` | 3 | Next larger match |
| `=XMATCH(5,{1;2;3})` | #N/A | No match gives #N/A |

## Notes

- match_mode 2 (wildcard) and search_mode 2/-2 (binary search) are not supported.
- Related functions: XLOOKUP.
- Supported: Excel 2010+. Registered as `XMATCH` (drop-in) on hosts without the native function, and as `EG.XMATCH` on modern Excel that has it.
