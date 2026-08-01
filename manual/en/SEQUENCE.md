# SEQUENCE

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Generates a sequence of numbers as an array.

## Syntax

```
=SEQUENCE(rows, [cols], [start], [step])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| rows | Required | number of rows |
| cols | Optional | number of columns (default 1) |
| start | Optional | first value (default 1) |
| step | Optional | increment (default 1) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SEQUENCE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `SEQUENCE` (drop-in) on hosts without the native function, and as `EG.SEQUENCE` on modern Excel that has it.
