# BITAND

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns a bitwise AND of two numbers.

## Syntax

```
=BITAND(number1, number2)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number1 | Required | first number (>= 0) |
| number2 | Required | second number (>= 0) |

## Returns

Returns the bitwise AND of two numbers as a scalar number. Returns #NUM! if an argument is not a number/logical, is negative or fractional, or exceeds 2^48-1 (281,474,976,710,655).

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BITAND(13,25)` | 9 | 1101 AND 11001 = 1001 |
| `=BITAND(-1,2)` | #NUM! | Negative operand is an error |

## Notes

- Non-numeric arguments also return #NUM!, not #VALUE!.
- All arguments accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Related functions: BITOR, BITXOR
- Supported: Excel 2010+. Registered as `BITAND` (drop-in) on hosts without the native function, and as `EG.BITAND` on modern Excel that has it.
