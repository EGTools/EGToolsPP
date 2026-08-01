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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SAMPLINGAC(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `SAMPLINGAC` on every Excel version.
