# PDURATION

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the number of periods required by an investment to reach a value.

## Syntax

```
=PDURATION(rate, pv, fv)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| rate | Required | interest rate per period |
| pv | Required | present value |
| fv | Required | future value |

## Returns

Returns the number of periods needed to reach the target value (scalar number). If rate, pv, or fv is non-numeric or not positive, returns a #NUM! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=PDURATION(0.025,2000,2200)` | 3.859866163 | Periods for 2000 to 2200 at 2.5% |

## Notes

- All arguments accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Related function: RRI
- Supported: Excel 2010+. Registered as `PDURATION` (drop-in) on hosts without the native function, and as `EG.PDURATION` on modern Excel that has it.
