# BASE

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Converts a number into text in the given radix (base).

## Syntax

```
=BASE(number, radix, [min_length])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number | Required | the number to convert (>= 0) |
| radix | Required | base, 2 to 36 |
| min_length | Optional | minimum length of the result (optional) |

## Returns

Returns the representation in the given radix as text (scalar). A negative or non-numeric number, or a radix outside 2 to 36, returns a #NUM! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BASE(255,16)` | FF | 255 in hexadecimal |
| `=BASE(15,2,8)` | 00001111 | Zero-padded to min_length |

## Notes

- The fractional part of number is truncated (floor).
- Results shorter than min_length are left-padded with zeros.
- Related function: DECIMAL
- Supported: Excel 2010+. Registered as `BASE` (drop-in) on hosts without the native function, and as `EG.BASE` on modern Excel that has it.
