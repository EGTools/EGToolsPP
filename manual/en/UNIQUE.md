# UNIQUE

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Returns the unique rows (or columns) of an array.

## Syntax

```
=UNIQUE(array, [by_col], [exactly_once])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | data |
| by_col | Optional | TRUE for unique columns |
| exactly_once | Optional | TRUE to keep items appearing once |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=UNIQUE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `UNIQUE` (drop-in) on hosts without the native function, and as `EG.UNIQUE` on modern Excel that has it.
