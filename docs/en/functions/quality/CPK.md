# CPK

**Category**: EGTools-only function

Process capability index Cpk from measurements and spec limits.

## Syntax

```
=CPK(data, [usl], [lsl])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| data | Required | measurement range or array |
| usl | Optional | upper spec limit (optional if LSL given) |
| lsl | Optional | lower spec limit (optional if USL given) |

## Returns

Returns the process capability index Cpk = min((USL-μ)/3σ, (μ-LSL)/3σ) as a scalar number. Missing data or both limits missing returns #VALUE!; fewer than 2 numeric values returns #NUM!; a zero standard deviation returns #DIV/0!.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CPK({1;2;3;4;5},5,0)` | 0.421637021 | Cpk for spec limits 0 to 5 |

## Notes

- Uses the sample standard deviation (n-1).
- If only one limit is given, only that side is evaluated.
- USL·LSL accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Related function: CP
- Supported: Excel 2010+. Always registered as `CPK` on every Excel version.
