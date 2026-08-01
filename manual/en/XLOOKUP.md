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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=XLOOKUP(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `XLOOKUP` (drop-in) on hosts without the native function, and as `EG.XLOOKUP` on modern Excel that has it.
