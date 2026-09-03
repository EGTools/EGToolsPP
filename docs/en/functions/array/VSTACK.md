# VSTACK

**Category**: MS Excel compatibility function · **Native since**: Excel 2024

Stacks arrays vertically into one array.

## Syntax

```
=VSTACK(array1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array1 | Required | array to stack |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns the argument arrays stacked vertically; on dynamic-array hosts the result spills. Narrower arrays are padded with #N/A to the widest width. Returns #VALUE! if no valid array argument is given.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=VSTACK({1,2},{3,4;5,6})` | {1,2;3,4;5,6} | Stack two arrays vertically |
| `=VSTACK({1,2,3},{4,5})` | {1,2,3;4,5,#N/A} | Widths differ, padded with #N/A |

## Notes

- Omitted arguments are skipped.
- Related function: HSTACK
- Supported: Excel 2010+. Registered as `VSTACK` (drop-in) on hosts without the native function, and as `EG.VSTACK` on modern Excel that has it.
