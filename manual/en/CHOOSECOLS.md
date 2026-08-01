# CHOOSECOLS

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Returns the specified columns from an array.

## Syntax

```
=CHOOSECOLS(array, col_num1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| col_num1 | Required | column to return (neg = from end) |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CHOOSECOLS(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `CHOOSECOLS` (drop-in) on hosts without the native function, and as `EG.CHOOSECOLS` on modern Excel that has it.
