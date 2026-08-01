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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=RRI(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `RRI` (drop-in) on hosts without the native function, and as `EG.RRI` on modern Excel that has it.
