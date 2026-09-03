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

Returns the array with its rows reordered by the key arrays; on dynamic-array hosts it spills. Returns #VALUE! when no valid key array (same row count as array) is given.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SORTBY({"a";"b";"c"},{3;1;2})` | {b;c;a} | Sort by key array |
| `=SORTBY({1;2;3},{2;2;1},1,{9;1;5},-1)` | {3;1;2} | Second key descending |

## Notes

- Key arrays must be column vectors with the same row count as array; only row sorting is supported (no by-column sorting).
- Key arrays with a mismatched row count are silently ignored.
- Related functions: SORT.
- Supported: Excel 2010+. Registered as `SORTBY` (drop-in) on hosts without the native function, and as `EG.SORTBY` on modern Excel that has it.
