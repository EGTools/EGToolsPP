# SWITCH

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Compares an expression to values and returns the first matching result (or default).

## Syntax

```
=SWITCH(expression, value1, result1, [default], ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| expression | Required | value to compare |
| value1 | Required | value to match |
| result1 | Required | result if the value matches |
| default | Optional | result if no value matches (optional) |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SWITCH(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `SWITCH` (drop-in) on hosts without the native function, and as `EG.SWITCH` on modern Excel that has it.
