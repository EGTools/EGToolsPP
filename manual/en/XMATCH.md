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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=XMATCH(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `XMATCH` (drop-in) on hosts without the native function, and as `EG.XMATCH` on modern Excel that has it.
