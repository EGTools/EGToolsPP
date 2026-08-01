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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=WEEKNUMOFMONTH(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `WEEKNUMOFMONTH` on every Excel version.
