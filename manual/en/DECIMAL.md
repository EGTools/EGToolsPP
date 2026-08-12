# DECIMAL

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Converts a text representation of a number in a given base to a decimal number.

## Syntax

```
=DECIMAL(text, radix)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | the text to convert |
| radix | Required | base, 2 to 36 |

## Returns

Returns the decimal value of the base-radix text (scalar number). A radix outside 2 to 36 or a character not valid in that radix returns a #NUM! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DECIMAL("FF",16)` | 255 | Hex to decimal |
| `=DECIMAL("111",2)` | 7 | Binary to decimal |

## Notes

- Digits are case-insensitive.
- All arguments accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Related function: BASE
- Supported: Excel 2010+. Registered as `DECIMAL` (drop-in) on hosts without the native function, and as `EG.DECIMAL` on modern Excel that has it.
