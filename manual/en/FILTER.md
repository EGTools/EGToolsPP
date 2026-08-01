# FILTER

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Filters an array, keeping rows/columns where include is TRUE.

## Syntax

```
=FILTER(array, include, [if_empty])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | data to filter |
| include | Required | boolean array (rows or columns) |
| if_empty | Optional | value if nothing matches (optional) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=FILTER(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `FILTER` (drop-in) on hosts without the native function, and as `EG.FILTER` on modern Excel that has it. On some legacy hosts it registers as `xFILTER` due to a keyword conflict.
