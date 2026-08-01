# EXRATE

**Category**: EGTools-only function

Looks up the KRW standard exchange rate for a date (up to 10 days back).

## Syntax

```
=EXRATE([currency], [date])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| currency | Optional | 3-letter code, e.g. USD (default), EUR, JPY… |
| date | Optional | date to look up (default today) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=EXRATE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `EXRATE` on every Excel version.
