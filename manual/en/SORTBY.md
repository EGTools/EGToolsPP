# SORTBY

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Sorts an array by values in other arrays.

## Syntax

```
=SORTBY(array, by_array1, [sort_order1], ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | data to sort |
| by_array1 | Required | key array (aligned to rows) |
| sort_order1 | Optional | 1 asc (default), -1 desc (optional) |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SORTBY(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `SORTBY` (drop-in) on hosts without the native function, and as `EG.SORTBY` on modern Excel that has it.
