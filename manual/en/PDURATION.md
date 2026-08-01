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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=PDURATION(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `PDURATION` (drop-in) on hosts without the native function, and as `EG.PDURATION` on modern Excel that has it.
