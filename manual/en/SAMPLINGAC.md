# SAMPLINGAC

**Category**: EGTools-only function

Returns the ISO 2859-1 acceptance number (Ac).

## Syntax

```
=SAMPLINGAC(lot_size, aql, [sample_type], [inspection], [plan], [round])
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

Returns the ISO 2859-1 acceptance number Ac as a scalar number. An AQL outside 0-1000 or an invalid level/plan/round returns #VALUE!; a lot size below 2 or an AQL below 0.01 returns #NUM!; #N/A when no table value can be reached.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SAMPLINGAC(1000,1)` | 2 | G2, normal, single, AQL 1.0 |
| `=SAMPLINGAC(1000,1,"G2",1)` | 1 | Tightened inspection |
| `=SAMPLINGAC(500,2.5)` | 3 | Lot 500, AQL 2.5 |

## Notes

- A lot size of 1 returns 0.
- Related functions: SAMPLINGSIZE, SAMPLINGRE, SAMPLINGLABEL
- Supported: Excel 2010+. Always registered as `SAMPLINGAC` on every Excel version.
