# MONTHBHYWEEK

**Category**: EGTools-only function

The month a date's week belongs to, by week-start and owner weekday.

## Syntax

```
=MONTHBHYWEEK(date, [start_of_week], [owner_of_week])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| date | Required | the date |
| start_of_week | Optional | week start day 1=Sun…7=Sat (default 2=Mon) |
| owner_of_week | Optional | owner day deciding the month (default 5=Thu) |

## Returns

Returns the month (1-12, scalar number) the date's week belongs to. An unparseable date or a weekday argument outside 1-7 returns a #VALUE! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=MONTHBHYWEEK(DATE(2026,2,1))` | 1 | Sun Feb 1's week belongs to January |
| `=MONTHBHYWEEK(DATE(2026,2,2))` | 2 | Feb 2's week belongs to February |

## Notes

- Weekday numbers are 1=Sun to 7=Sat; defaults are Monday start and Thursday owner.
- A week's month is the month containing its owner weekday.
- Related function: WEEKNUMOFMONTH
- Supported: Excel 2010+. Always registered as `MONTHBHYWEEK` on every Excel version.
