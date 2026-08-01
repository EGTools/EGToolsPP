# SAMPLINGRE

**Category**: EGTools-only function

Returns the ISO 2859-1 rejection number (Re).

## Syntax

```
=SAMPLINGRE(lot_size, aql, [sample_type], [inspection], [plan], [round])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| lot_size | Required | lot/batch quantity |
| aql | Required | acceptance quality limit |
| sample_type | Optional | inspection level (default G2) |
| inspection | Optional | 0 normal, 1 tightened, -1 reduced |
| plan | Optional | 1 single, 2 double, 3–5 multiple |
| round | Optional | sampling round |

## Returns

Returns the ISO 2859-1 rejection number Re as a scalar number. An AQL outside 0-1000 or an invalid level/plan/round returns #VALUE!; a lot size below 2 or an AQL below 0.01 returns #NUM!; #N/A when no table value can be reached.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SAMPLINGRE(1000,1)` | 3 | G2, normal, single, AQL 1.0 |
| `=SAMPLINGRE(1000,1,"G2",0,2,1)` | 3 | Double sampling, first-round Re |

## Notes

- For single sampling, Re = Ac + 1.
- A lot size of 1 returns 1.
- Related functions: SAMPLINGSIZE, SAMPLINGAC, SAMPLINGLABEL
- Supported: Excel 2010+. Always registered as `SAMPLINGRE` on every Excel version.
