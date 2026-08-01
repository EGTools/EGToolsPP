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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SAMPLINGLABEL(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `SAMPLINGLABEL` on every Excel version (no `EG.`/`x` prefix).
