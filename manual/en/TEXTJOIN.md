# TEXTJOIN

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Joins text using a delimiter, optionally ignoring empty cells.

## Syntax

```
=TEXTJOIN(delimiter, ignore_empty, text1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| delimiter | Required | separator between items |
| ignore_empty | Required | TRUE to skip empty cells |
| text1 | Required | text or range |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTJOIN(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `TEXTJOIN` (drop-in) on hosts without the native function, and as `EG.TEXTJOIN` on modern Excel that has it.
