# IFS

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Checks conditions in order and returns the value for the first TRUE condition.

## Syntax

```
=IFS(condition1, value1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| condition1 | Required | 1st condition |
| value1 | Required | result if condition1 is TRUE |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IFS(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `IFS` (drop-in) on hosts without the native function, and as `EG.IFS` on modern Excel that has it.
