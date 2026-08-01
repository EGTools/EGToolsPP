# DROP

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Removes rows/columns from the start or end of an array.

## Syntax

```
=DROP(array, rows, [cols])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| rows | Required | rows to drop (neg = from end) |
| cols | Optional | columns to drop (optional) |

## Returns

Returns the array with the given number of rows/columns removed; on dynamic-array hosts the result spills. Returns #VALUE! if the array is empty or nothing remains after dropping.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DROP({1,2,3;4,5,6;7,8,9},1)` | {4,5,6;7,8,9} | Drop the first row |
| `=DROP({1,2,3;4,5,6;7,8,9},1,-1)` | {4,5;7,8} | Drop first row, last column |
| `=DROP({1,2},5)` | #VALUE! | Error when everything is dropped |

## Notes

- Negative rows/cols drop from the end.
- An empty result returns #VALUE! instead of the native #CALC!.
- Related function: TAKE
- Supported: Excel 2010+. Registered as `DROP` (drop-in) on hosts without the native function, and as `EG.DROP` on modern Excel that has it.
