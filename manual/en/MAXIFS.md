# MAXIFS

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Returns the maximum of cells that meet all criteria.

## Syntax

```
=MAXIFS(max_range, criteria_range1, criteria1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| max_range | Required | cells to maximise |
| criteria_range1 | Required | range to test |
| criteria1 | Required | criterion, e.g. ">5" |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=MAXIFS(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `MAXIFS` (drop-in) on hosts without the native function, and as `EG.MAXIFS` on modern Excel that has it.
