# LET

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Assigns names to values and uses them in a final calculation.

## Syntax

```
=LET(name1, value1, calculation, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| name1 | Required | name to define |
| value1 | Required | value for this name |
| calculation | Required | the calculation using the names above |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=LET(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `LET` (drop-in) on hosts without the native function, and as `EG.LET` on modern Excel that has it. On some legacy hosts it registers as `xLET` due to a keyword conflict.
