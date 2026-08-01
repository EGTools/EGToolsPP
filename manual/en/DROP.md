# DROP

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Removes rows/columns from the start or end of an array.

## Syntax

```
=DROP(array, rows, [cols])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| rows | Required | rows to drop (neg = from end) |
| cols | Optional | columns to drop (optional) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DROP(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `DROP` (drop-in) on hosts without the native function, and as `EG.DROP` on modern Excel that has it.
