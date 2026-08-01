# CHOOSEROWS

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Returns the specified rows from an array.

## Syntax

```
=CHOOSEROWS(array, row_num1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| row_num1 | Required | row to return (neg = from end) |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CHOOSEROWS(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `CHOOSEROWS` (drop-in) on hosts without the native function, and as `EG.CHOOSEROWS` on modern Excel that has it.
