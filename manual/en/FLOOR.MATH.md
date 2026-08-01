# FLOOR.MATH

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Rounds a number down to the nearest integer or multiple of significance.

## Syntax

```
=FLOOR.MATH(number, [significance], [mode])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number | Required | the number to round |
| significance | Optional | multiple to round to (optional, default 1) |
| mode | Optional | if nonzero, rounds negatives toward zero (optional) |

## Returns

Returns the number rounded down to a multiple of significance (scalar). A non-numeric number returns a #VALUE! error; a significance of 0 returns 0.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=FLOOR.MATH(6.7)` | 6 | Round down to default multiple 1 |
| `=FLOOR.MATH(-5.5,2)` | -6 | Negatives round away from zero |
| `=FLOOR.MATH(-5.5,2,1)` | -4 | Nonzero mode: toward zero |

## Notes

- The sign of significance is ignored (absolute value used).
- Related function: CEILING.MATH
- Supported: Excel 2010+. Registered as `FLOOR.MATH` (drop-in) on hosts without the native function, and as `EG.FLOOR.MATH` on modern Excel that has it.
