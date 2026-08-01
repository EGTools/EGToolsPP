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

Returns the right-shifted value (floored) as a scalar number. Returns #NUM! if number is not an integer in [0, 2^48-1], if shift_amount is not an integer or its absolute value exceeds 53, or if the result exceeds 2^53.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BITRSHIFT(13,2)` | 3 | Shift right by 2 bits |
| `=BITRSHIFT(3,-2)` | 12 | Negative shift goes left |

## Notes

- A negative shift_amount shifts left instead.
- Related function: BITLSHIFT
- Supported: Excel 2010+. Registered as `BITRSHIFT` (drop-in) on hosts without the native function, and as `EG.BITRSHIFT` on modern Excel that has it.
