# EXPAND

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Expands an array to the given size, padding with a value.

## Syntax

```
=EXPAND(array, rows, [cols], [pad_with])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| rows | Required | total rows |
| cols | Optional | total columns (optional) |
| pad_with | Optional | pad value (default #N/A) |

## Returns

Returns the array expanded to rows x cols; on dynamic-array hosts the result spills. New cells are filled with pad_with (default #N/A). Returns #VALUE! if the array is empty or the target size is smaller than the source.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=EXPAND({1,2;3,4},3,3,0)` | {1,2,0;3,4,0;0,0,0} | Expand to 3x3, pad with 0 |
| `=EXPAND({1,2},1,4)` | {1,2,#N/A,#N/A} | Default pad value is #N/A |
| `=EXPAND({1,2;3,4},1,1)` | #VALUE! | Shrinking is not allowed |

## Notes

- Omitting rows or cols keeps the source dimension.
- Omitting pad_with pads with #N/A.
- Supported: Excel 2010+. Registered as `EXPAND` (drop-in) on hosts without the native function, and as `EG.EXPAND` on modern Excel that has it.
