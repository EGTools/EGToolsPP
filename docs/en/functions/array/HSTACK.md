# HSTACK

**Category**: MS Excel compatibility function · **Native since**: Excel 2024

Stacks arrays horizontally into one array.

## Syntax

```
=HSTACK(array1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array1 | Required | array to stack |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns the argument arrays stacked side by side; on dynamic-array hosts the result spills. Shorter arrays are padded with #N/A to the tallest height. Returns #VALUE! if no valid array argument is given.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=HSTACK({1;2},{3,4;5,6})` | {1,3,4;2,5,6} | Stack two arrays horizontally |
| `=HSTACK({1,2},{3;4;5})` | {1,2,3;#N/A,#N/A,4;#N/A,#N/A,5} | Heights differ, padded with #N/A |

## Notes

- Omitted arguments are skipped.
- Related function: VSTACK
- Supported: Excel 2010+. Registered as `HSTACK` (drop-in) on hosts without the native function, and as `EG.HSTACK` on modern Excel that has it.
