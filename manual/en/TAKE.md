# TAKE

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Returns rows/columns from the start or end of an array.

## Syntax

```
=TAKE(array, rows, [cols])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| rows | Required | rows to take (neg = from end) |
| cols | Optional | columns to take (optional) |

## Returns

Returns the requested number of rows/columns from the start (positive) or end (negative) of the array; on dynamic-array hosts the result spills. Returns #VALUE! if the array is empty or the row/column count resolves to 0.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TAKE({1,2,3;4,5,6;7,8,9},2)` | {1,2,3;4,5,6} | First two rows |
| `=TAKE({1,2,3;4,5,6;7,8,9},-1,2)` | {7,8} | Last row, first two columns |
| `=TAKE({1,2;3,4},0)` | #VALUE! | Zero rows is an error |

## Notes

- Negative rows/cols take from the end.
- An empty result returns #VALUE! instead of the native #CALC!.
- Related function: DROP
- Supported: Excel 2010+. Registered as `TAKE` (drop-in) on hosts without the native function, and as `EG.TAKE` on modern Excel that has it.
