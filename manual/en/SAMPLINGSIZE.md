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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SAMPLINGSIZE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `SAMPLINGSIZE` on every Excel version.
