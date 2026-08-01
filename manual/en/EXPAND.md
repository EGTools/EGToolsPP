# EXPAND

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Expands an array to the given size, padding with a value.

## Syntax

```
=EXPAND(array, rows, [cols], [pad_with])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| rows | Required | total rows |
| cols | Optional | total columns (optional) |
| pad_with | Optional | pad value (default #N/A) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=EXPAND(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `EXPAND` (drop-in) on hosts without the native function, and as `EG.EXPAND` on modern Excel that has it.
