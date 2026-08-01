# SORT

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Sorts an array by a row/column.

## Syntax

```
=SORT(array, [sort_index], [sort_order], [by_col])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | data to sort |
| sort_index | Optional | column (or row) to sort by (default 1) |
| sort_order | Optional | 1 ascending (default), -1 descending |
| by_col | Optional | TRUE to sort columns instead of rows |

## Returns

Returns a same-size array sorted by the given column (or row); on dynamic-array hosts it spills. Returns #VALUE! when sort_index is outside the array or on internal errors.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SORT({3;1;2})` | {1;2;3} | Ascending sort |
| `=SORT({1,3;2,1;3,2},2)` | {2,1;3,2;1,3} | Sort by column 2 |
| `=SORT({5;3;9},1,-1)` | {9;5;3} | Descending sort |

## Notes

- Any negative sort_order sorts descending; anything else (including 0) sorts ascending (native accepts only 1/-1).
- Equal keys keep their original order (stable sort).
- Related functions: SORTBY, UNIQUE.
- Supported: Excel 2010+. Registered as `SORT` (drop-in) on hosts without the native function, and as `EG.SORT` on modern Excel that has it. On some legacy hosts it registers as `xSORT` due to a keyword conflict.
