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

Returns the KRW standard exchange rate as a numeric scalar. Returns #VALUE! for a future date, and #N/A for an unsupported currency code or when no published rate is found within the 10-day look-back.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=EXRATE("USD")` |  | Result depends on date and network |

## Notes

- Fetches the standard rate from smbs.biz (Seoul Money Brokerage) over HTTP — a network connection is required.
- If no rate is published for the date, it looks back up to 10 days (holidays).
- CNY is mapped to CNH automatically. Registered non-macro and thread-safe for parallel recalculation.
- An array in the currency and date returns #VALUE! — for multiple items, copy the formula down row by row.
- Supported: Excel 2010+. Always registered as `EXRATE` on every Excel version.
