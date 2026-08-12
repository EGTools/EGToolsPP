# SAMPLINGLABEL

**Category**: EGTools-only function

Returns the ISO 2859-1 sample size code letter for a lot size and inspection level.

## Syntax

```
=SAMPLINGLABEL(lot_size, [sample_type])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| lot_size | Required | lot/batch quantity (2 or more) |
| sample_type | Optional | inspection level: G1/G2/G3 (general I,II,III) or S1–S4 (special); default G2 |

## Returns

Returns the ISO 2859-1 sample size code letter as a one-character text scalar. A lot size below 2 returns #NUM!; an inspection level other than S1-S4/G1-G3 returns #VALUE!.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SAMPLINGLABEL(1000)` | J | Default G2, lot 1000 |
| `=SAMPLINGLABEL(1000,"S3")` | E | Special level S-3 |

## Notes

- Hyphens and spaces in the level are ignored ("G-2" = "G2").
- Based on the built-in ISO 2859-1:1999 tables.
- All arguments accept arrays, so a LOT list or an AQL table can be computed at once — element-wise with a same-shape spill; scalars are broadcast, a column vector × a row vector expands to their outer product, and size-mismatched elements yield #N/A.
- Related functions: SAMPLINGSIZE, SAMPLINGAC, SAMPLINGRE
- Supported: Excel 2010+. Always registered as `SAMPLINGLABEL` on every Excel version.
