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

Returns a sorted array of the same shape as the input, spilling on dynamic-array hosts. A non-array single value is returned unchanged. Returns #VALUE! when the array is missing, the key column is out of range, or sort_order is not 1/-1.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTNUMSORT({"a10";"a2";"a1"})` | {a1;a2;a10} | natural sort ascending |
| `=TEXTNUMSORT({"a1";"a10";"a2"},1,-1)` | {a10;a2;a1} | descending |

## Notes

- Text and number runs are compared separately; numbers compare numerically (a2 < a10).
- Commas inside numbers are ignored as thousands separators; '.' is treated as a decimal point.
- The sort is stable, so rows with equal keys keep their order.
- Supported: Excel 2010+. Always registered as `TEXTNUMSORT` on every Excel version.
