# CEILING.MATH

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Rounds a number up to the nearest integer or multiple of significance.

## Syntax

```
=CEILING.MATH(number, [significance], [mode])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number | Required | the number to round |
| significance | Optional | multiple to round to (optional, default 1) |
| mode | Optional | if nonzero, rounds negatives away from zero (optional) |

## Returns

Returns the number rounded up to a multiple of significance (scalar). A non-numeric number returns a #VALUE! error; a significance of 0 returns 0.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CEILING.MATH(6.3)` | 7 | Round up to default multiple 1 |
| `=CEILING.MATH(-5.5,2)` | -4 | Negatives round toward zero |
| `=CEILING.MATH(-5.5,2,1)` | -6 | Nonzero mode: away from zero |

## Notes

- The sign of significance is ignored (absolute value used).
- Related function: FLOOR.MATH
- Supported: Excel 2010+. Registered as `CEILING.MATH` (drop-in) on hosts without the native function, and as `EG.CEILING.MATH` on modern Excel that has it.
