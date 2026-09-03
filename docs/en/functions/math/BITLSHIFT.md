# BITLSHIFT

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns a number shifted left by the given number of bits.

## Syntax

```
=BITLSHIFT(number, shift_amount)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number | Required | the number to shift (>= 0) |
| shift_amount | Required | bits to shift left (negative shifts right) |

## Returns

Returns the left-shifted value as a scalar number. Returns #NUM! if number is not an integer in [0, 2^48-1], if shift_amount is not an integer or its absolute value exceeds 53, or if the result exceeds 2^53.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BITLSHIFT(4,2)` | 16 | Shift left by 2 bits |
| `=BITLSHIFT(4,-1)` | 2 | Negative shift goes right |

## Notes

- A negative shift_amount shifts right instead.
- All arguments accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Related function: BITRSHIFT
- Supported: Excel 2010+. Registered as `BITLSHIFT` (drop-in) on hosts without the native function, and as `EG.BITLSHIFT` on modern Excel that has it.
