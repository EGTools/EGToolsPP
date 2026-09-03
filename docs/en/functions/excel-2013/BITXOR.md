# BITXOR

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns a bitwise XOR of two numbers.

## Syntax

```
=BITXOR(number1, number2)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number1 | Required | first number (>= 0) |
| number2 | Required | second number (>= 0) |

## Returns

Returns the bitwise XOR of two numbers as a scalar number. Returns #NUM! if an argument is not a number/logical, is negative or fractional, or exceeds 2^48-1 (281,474,976,710,655).

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BITXOR(5,3)` | 6 | 101 XOR 011 = 110 |
| `=BITXOR(5,"a")` | #NUM! | Non-numeric operand is an error |

## Notes

- Non-numeric arguments also return #NUM!, not #VALUE!.
- All arguments accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Related functions: BITAND, BITOR
- Supported: Excel 2010+. Registered as `BITXOR` (drop-in) on hosts without the native function, and as `EG.BITXOR` on modern Excel that has it.
