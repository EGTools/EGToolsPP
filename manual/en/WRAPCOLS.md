# WRAPCOLS

**Category**: MS Excel compatibility function · **Native since**: Excel 2024

Wraps a vector into a 2D array, a set number of values per column.

## Syntax

```
=WRAPCOLS(vector, wrap_count, [pad_with])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| vector | Required | the vector to wrap |
| wrap_count | Required | values per column |
| pad_with | Optional | pad value (default #N/A) |

## Returns

Returns the vector wrapped into a 2-D array with wrap_count values per column, filled top to bottom; on dynamic-array hosts the result spills. Missing cells in the last column are filled with pad_with (default #N/A). Returns #VALUE! if the vector is empty or wrap_count is 0 or less.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=WRAPCOLS({1,2,3,4,5},2)` | {1,3,5;2,4,#N/A} | Two values per column |
| `=WRAPCOLS({1,2,3,4,5},3,0)` | {1,4;2,5;3,0} | Pad with 0 |

## Notes

- A 2-D array input is also accepted and flattened row-major.
- Related function: WRAPROWS
- Supported: Excel 2010+. Registered as `WRAPCOLS` (drop-in) on hosts without the native function, and as `EG.WRAPCOLS` on modern Excel that has it.
