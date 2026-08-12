# RANDARRAY

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Returns an array of random numbers.

## Syntax

```
=RANDARRAY([rows], [cols], [min], [max], [integer])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| rows | Optional | number of rows (default 1) |
| cols | Optional | number of columns (default 1) |
| min | Optional | minimum (default 0) |
| max | Optional | maximum (default 1) |
| integer | Optional | TRUE for whole numbers |

## Returns

Returns a rows-by-cols array of random numbers; on dynamic-array hosts it spills. Returns #VALUE! when rows or cols is 0 or less, or when max is less than min.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=RANDARRAY(2,3,1,10,TRUE)` |  | Random; result varies each recalc |

## Notes

- With integer=TRUE, generates integers from min to max inclusive; otherwise reals in [min, max).
- Arrays in All arguments return an array of that argument's shape where each element is the first value of that element's result (native-compatible demotion lifting).
- Related functions: SEQUENCE.
- Supported: Excel 2010+. Registered as `RANDARRAY` (drop-in) on hosts without the native function, and as `EG.RANDARRAY` on modern Excel that has it.
