# TOROW

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Returns the array as a single row.

## Syntax

```
=TOROW(array, [ignore], [scan_by_column])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| ignore | Optional | 0 none, 1 blanks, 2 errors, 3 both |
| scan_by_column | Optional | TRUE to scan by column |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TOROW(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `TOROW` (drop-in) on hosts without the native function, and as `EG.TOROW` on modern Excel that has it.
