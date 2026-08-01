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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=NETWORKHOUR(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `NETWORKHOUR` on every Excel version.
