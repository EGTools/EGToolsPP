# BITOR

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns a bitwise OR of two numbers.

## Syntax

```
=BITOR(number1, number2)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number1 | Required | first number (>= 0) |
| number2 | Required | second number (>= 0) |

## Returns

Returns the bitwise OR of two numbers as a scalar number. Returns #NUM! if an argument is not a number/logical, is negative or fractional, or exceeds 2^48-1 (281,474,976,710,655).

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BITOR(23,10)` | 31 | 10111 OR 01010 = 11111 |
| `=BITOR(1.5,2)` | #NUM! | Fractional operand is an error |

## Notes

- Non-numeric arguments also return #NUM!, not #VALUE!.
- Related functions: BITAND, BITXOR
- Supported: Excel 2010+. Registered as `BITOR` (drop-in) on hosts without the native function, and as `EG.BITOR` on modern Excel that has it.
