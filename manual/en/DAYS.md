# DAYS

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the number of days between two dates (end − start).

## Syntax

```
=DAYS(end_date, start_date)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| end_date | Required | end date |
| start_date | Required | start date |

## Returns

Returns end_date minus start_date in days as a scalar number. The time-of-day part of each date is truncated. Returns #VALUE! if an argument is omitted.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DAYS(DATE(2026,3,1),DATE(2026,1,15))` | 45 | Days between two dates |
| `=DAYS(DATE(2026,1,1),DATE(2026,1,31))` | -30 | Negative when end is earlier |

## Notes

- Date text such as "2026-01-15" is not parsed; non-numeric values are treated as 0.
- Supported: Excel 2010+. Registered as `DAYS` (drop-in) on hosts without the native function, and as `EG.DAYS` on modern Excel that has it.
