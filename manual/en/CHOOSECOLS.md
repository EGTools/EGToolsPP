# CHOOSECOLS

**Category**: MS Excel compatibility function · **Native since**: Excel 2024

Returns the specified columns from an array.

## Syntax

```
=CHOOSECOLS(array, col_num1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| col_num1 | Required | column to return (neg = from end) |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns an array made of the selected columns; on dynamic-array hosts the result spills. Returns #VALUE! if the array is empty, a column number is 0 or out of range, or no column number is supplied.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CHOOSECOLS({1,2,3;4,5,6},1,3)` | {1,3;4,6} | Pick columns 1 and 3 |
| `=CHOOSECOLS({1,2,3;4,5,6},-1)` | {3;6} | Pick the last column |
| `=CHOOSECOLS({1,2;3,4},5)` | #VALUE! | Column number out of range |

## Notes

- Negative column numbers count from the end.
- Use CHOOSEROWS to select rows.
- Column numbers also accept arrays such as {1,3} or SEQUENCE(n).
- Supported: Excel 2010+. Registered as `CHOOSECOLS` (drop-in) on hosts without the native function, and as `EG.CHOOSECOLS` on modern Excel that has it.
