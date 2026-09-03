# RRI

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns an equivalent interest rate for the growth of an investment.

## Syntax

```
=RRI(nper, pv, fv)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| nper | Required | number of periods |
| pv | Required | present value |
| fv | Required | future value |

## Returns

Returns the equivalent interest rate per period for the growth of an investment (scalar number). Returns a #NUM! error if an argument is non-numeric, nper is not positive, or pv is 0.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=RRI(96,10000,11000)` | 0.000993307 | Rate for 10000 to 11000 over 96 periods |

## Notes

- All arguments accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Related function: PDURATION
- Supported: Excel 2010+. Registered as `RRI` (drop-in) on hosts without the native function, and as `EG.RRI` on modern Excel that has it.
