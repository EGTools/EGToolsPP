# BITRSHIFT

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns a number shifted right by the given number of bits.

## Syntax

```
=BITRSHIFT(number, shift_amount)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number | Required | the number to shift (>= 0) |
| shift_amount | Required | bits to shift right (negative shifts left) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BITRSHIFT(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `BITRSHIFT` (drop-in) on hosts without the native function, and as `EG.BITRSHIFT` on modern Excel that has it.
