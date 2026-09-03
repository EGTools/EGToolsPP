# SAMPLINGSIZE

**Category**: EGTools-only function

Returns the ISO 2859-1 sample size (n) for a lot size and AQL.

## Syntax

```
=SAMPLINGSIZE(lot_size, aql, [sample_type], [inspection], [plan], [round])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| lot_size | Required | lot/batch quantity |
| aql | Required | acceptance quality limit, e.g. 1.0 |
| sample_type | Optional | inspection level G1/G2/G3/S1–S4 (default G2) |
| inspection | Optional | 0 normal (default), 1 tightened, -1 reduced |
| plan | Optional | 1 single (default), 2 double, 3–5 multiple; +10 forces without switching |
| round | Optional | sampling round: 1 (single), up to 2 (double) or 5 (multiple) |

## Returns

Returns the ISO 2859-1 sample size n as a scalar number. An AQL outside 0-1000 or an invalid level/plan/round returns #VALUE!; a lot size below 2 or an AQL below 0.01 returns #NUM!; #N/A when no table value can be reached.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SAMPLINGSIZE(1000,1)` | 80 | G2, normal, single, AQL 1.0 |
| `=SAMPLINGSIZE(1000,1,"G2",0,2,1)` | 50 | Double sampling, first sample |
| `=SAMPLINGSIZE(30,0.065)` | 30 | Capped at the lot size |

## Notes

- If the tabled n exceeds the lot size, the lot size is returned; a lot size of 1 returns 1.
- Adding 10 to plan forces that plan, ignoring the table's switch marks (*, ++).
- All arguments accept arrays, so a LOT list or an AQL table can be computed at once — element-wise with a same-shape spill; scalars are broadcast, a column vector × a row vector expands to their outer product, and size-mismatched elements yield #N/A.
- Related functions: SAMPLINGLABEL, SAMPLINGAC, SAMPLINGRE
- Supported: Excel 2010+. Always registered as `SAMPLINGSIZE` on every Excel version.
