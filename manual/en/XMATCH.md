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
| match_mode | Optional | 0 exact (default), -1 next smaller, 1 next larger, 3 regex |
| search_mode | Optional | 1 first-to-last (default), -1 last-to-first |

## Returns

Returns the 1-based position as a numeric scalar. Returns #N/A when the array is empty or nothing matches, and #VALUE! on internal errors.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=XMATCH(3,{1;2;3;4})` | 3 | Exact match position |
| `=XMATCH(2.5,{1;2;3},1)` | 3 | Next larger match |
| `=XMATCH(5,{1;2;3})` | #N/A | No match gives #N/A |
| `=XMATCH("rr[yx]$",{"apple";"Banana";"cherry"},3)` | 3 | Regex match (match_mode 3) |
| `=XMATCH("e",{"apple";"Banana";"cherry"},3,-1)` | 3 | Regex, searching last-to-first |

## Notes

- match_mode 2 (wildcard) and search_mode 2/-2 (binary search) are not supported.
- match_mode 3 (regex) interprets lookup_value as a regular-expression pattern and returns the position of the first item whose text contains a match for the pattern (partial match, REGEXTEST-style) — the same mode number as the 2024 addition to native Microsoft 365 XMATCH. Matching is case-sensitive by default; prefix the pattern with (?i) to ignore case (same usage as native). Non-text cells never match, and a non-text lookup_value, an invalid regex, or a combination with search_mode 2/-2 (binary search) returns #VALUE!.
- The regex flavor is std::wregex ECMAScript (may differ from native 365's PCRE2 in places).
- Related functions: XLOOKUP.
- Supported: Excel 2010+. Registered as `XMATCH` (drop-in) on hosts without the native function, and as `EG.XMATCH` on modern Excel that has it.
