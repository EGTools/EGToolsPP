# RANDARRAY

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Returns an array of random numbers.

## Syntax

```
=RANDARRAY([rows], [cols], [min], [max], [integer])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| rows | Optional | number of rows (default 1) |
| cols | Optional | number of columns (default 1) |
| min | Optional | minimum (default 0) |
| max | Optional | maximum (default 1) |
| integer | Optional | TRUE for whole numbers |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=RANDARRAY(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `RANDARRAY` (drop-in) on hosts without the native function, and as `EG.RANDARRAY` on modern Excel that has it.
