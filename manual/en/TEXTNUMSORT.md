# TEXTNUMSORT

**Category**: EGTools-only function

Sorts rows in natural order (text and embedded numbers compared separately).

## Syntax

```
=TEXTNUMSORT(array, [col_index], [sort_order], [ignore_text])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | data to sort |
| col_index | Optional | key column, 1-based (default 1) |
| sort_order | Optional | 1 ascending (default), -1 descending |
| ignore_text | Optional | TRUE compares number parts only |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTNUMSORT(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `TEXTNUMSORT` on every Excel version (no `EG.`/`x` prefix).
