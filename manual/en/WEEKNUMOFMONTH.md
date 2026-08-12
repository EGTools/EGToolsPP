# WEEKNUMOFMONTH

**Category**: EGTools-only function

Week number within the month, by week-start and owner weekday.

## Syntax

```
=WEEKNUMOFMONTH(date, [start_of_week], [owner_of_week])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| date | Required | the date |
| start_of_week | Optional | week start day 1=Sun…7=Sat (default 2=Mon) |
| owner_of_week | Optional | owner day deciding the month (default 5=Thu) |

## Returns

Returns the week number within the month as a scalar number. An unparseable date or a weekday argument outside 1-7 returns a #VALUE! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=WEEKNUMOFMONTH(DATE(2026,1,15))` | 3 | Third week of January |
| `=WEEKNUMOFMONTH(DATE(2026,2,1))` | 5 | Feb 1 falls in week 5 of January |

## Notes

- Weekday numbers are 1=Sun to 7=Sat; defaults are Monday start and Thursday owner.
- The owning month and week are decided by the week's owner weekday.
- If the date is an array, the function is applied element-wise and spills an array of the same shape (an error element returns that error).
- Related function: MONTHBHYWEEK
- Supported: Excel 2010+. Always registered as `WEEKNUMOFMONTH` on every Excel version.
