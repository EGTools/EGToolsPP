# CHOOSEROWS

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Returns the specified rows from an array.

## Syntax

```
=CHOOSEROWS(array, row_num1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| row_num1 | Required | row to return (neg = from end) |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns an array made of the selected rows; on dynamic-array hosts the result spills. Returns #VALUE! if the array is empty, a row number is 0 or out of range, or no row number is supplied.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CHOOSEROWS({1,2;3,4;5,6},1,-1)` | {1,2;5,6} | First and last rows |
| `=CHOOSEROWS({1,2;3,4},3)` | #VALUE! | Row number out of range |

## Notes

- Negative row numbers count from the end.
- Use CHOOSECOLS to select columns.
- Supported: Excel 2010+. Registered as `CHOOSEROWS` (drop-in) on hosts without the native function, and as `EG.CHOOSEROWS` on modern Excel that has it.
