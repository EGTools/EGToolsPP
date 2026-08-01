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
- Related function: BASE
- Supported: Excel 2010+. Registered as `DECIMAL` (drop-in) on hosts without the native function, and as `EG.DECIMAL` on modern Excel that has it.
