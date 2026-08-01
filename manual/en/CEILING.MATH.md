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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CEILING.MATH(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `CEILING.MATH` (drop-in) on hosts without the native function, and as `EG.CEILING.MATH` on modern Excel that has it.
