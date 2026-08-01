# CP

**Category**: EGTools-only function

Process capability index Cp from measurements and spec limits.

## Syntax

```
=CP(data, [usl], [lsl])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| data | Required | measurement range or array |
| usl | Optional | upper spec limit (optional if LSL given) |
| lsl | Optional | lower spec limit (optional if USL given) |

## Returns

Returns the process capability index Cp = (USL-LSL)/(6σ) as a scalar number. Missing data or both limits missing returns #VALUE!; fewer than 2 numeric values returns #NUM!; a zero standard deviation returns #DIV/0!.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CP({1;2;3;4;5},6,0)` | 0.632455532 | Cp for spec limits 0 to 6 |

## Notes

- Uses the sample standard deviation (n-1).
- If only one limit is given, the missing side is treated as the mean.
- Related function: CPK
- Supported: Excel 2010+. Always registered as `CP` on every Excel version.
