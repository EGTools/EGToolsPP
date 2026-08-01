# WRAPCOLS

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Wraps a vector into a 2D array, a set number of values per column.

## Syntax

```
=WRAPCOLS(vector, wrap_count, [pad_with])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| vector | Required | the vector to wrap |
| wrap_count | Required | values per column |
| pad_with | Optional | pad value (default #N/A) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=WRAPCOLS(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `WRAPCOLS` (drop-in) on hosts without the native function, and as `EG.WRAPCOLS` on modern Excel that has it.
