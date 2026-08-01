# TAKE

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Returns rows/columns from the start or end of an array.

## Syntax

```
=TAKE(array, rows, [cols])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| rows | Required | rows to take (neg = from end) |
| cols | Optional | columns to take (optional) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TAKE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `TAKE` (drop-in) on hosts without the native function, and as `EG.TAKE` on modern Excel that has it.
