# WRAPROWS

**Category**: MS Excel compatibility function · **Native since**: Excel 2024

Wraps a vector into a 2D array, a set number of values per row.

## Syntax

```
=WRAPROWS(vector, wrap_count, [pad_with])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| vector | Required | the vector to wrap |
| wrap_count | Required | values per row |
| pad_with | Optional | pad value (default #N/A) |

## Returns

Returns the vector wrapped into a 2-D array with wrap_count values per row, filled left to right; on dynamic-array hosts the result spills. Missing cells in the last row are filled with pad_with (default #N/A). Returns #VALUE! if the vector is empty or wrap_count is 0 or less.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=WRAPROWS({1,2,3,4,5},2)` | {1,2;3,4;5,#N/A} | Two values per row |
| `=WRAPROWS({1,2,3,4,5},2,0)` | {1,2;3,4;5,0} | Pad with 0 |

## Notes

- A 2-D array input is also accepted and flattened row-major.
- Arrays in wrap_count·pad_with return an array of that argument's shape where each element is the first value of that element's result (native-compatible demotion lifting).
- Related function: WRAPCOLS
- Supported: Excel 2010+. Registered as `WRAPROWS` (drop-in) on hosts without the native function, and as `EG.WRAPROWS` on modern Excel that has it.
