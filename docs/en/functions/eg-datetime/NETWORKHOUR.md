# NETWORKHOUR

**Category**: EGTools-only function

Calculates net working hours minus break times.

## Syntax

```
=NETWORKHOUR(start_time, end_time, [break_times])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| start_time | Required | start time(s) |
| end_time | Required | end time(s), same shape |
| break_times | Optional | break intervals: first column start, second column end |

## Returns

Spills an array sized by broadcasting start and end (scalar for single values); values are day fractions (1 = 24 hours). Start/end shapes that cannot be broadcast (differing sizes other than 1) or a break array with fewer than 2 columns return #VALUE!; non-numeric start/end cells yield empty text, and a date-bearing end earlier than its start yields #VALUE! for that cell.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=NETWORKHOUR(TIME(9,0,0),TIME(18,0,0))*24` | 9 | Working hours 9:00 to 18:00 |
| `=NETWORKHOUR(TIME(9,0,0),TIME(18,0,0),HSTACK(TIME(12,0,0),TIME(13,0,0)))*24` | 8 | Minus a one-hour lunch break |
| `=NETWORKHOUR(TIME(22,0,0),TIME(6,0,0))*24` | 8 | Overnight shift |

## Notes

- Results are in days; multiply by 24 for hours.
- With time-only values, an end smaller than its start is treated as crossing midnight.
- Overlapping break intervals are merged; breaks crossing midnight are split.
- Start and end support scalar↔array broadcasting — e.g. =NETWORKHOUR(A2:A100,$B$1) fixes the end time to one cell.
- Supported: Excel 2010+. Always registered as `NETWORKHOUR` on every Excel version.
