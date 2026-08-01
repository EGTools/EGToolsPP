# SEQUENCE

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Generates a sequence of numbers as an array.

## Syntax

```
=SEQUENCE(rows, [cols], [start], [step])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| rows | Required | number of rows |
| cols | Optional | number of columns (default 1) |
| start | Optional | first value (default 1) |
| step | Optional | increment (default 1) |

## Returns

Returns a rows-by-cols array of sequential numbers filled row by row; on dynamic-array hosts it spills. Returns #VALUE! when rows or cols is 0 or less.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SEQUENCE(3)` | {1;2;3} | Basic 3-row sequence |
| `=SEQUENCE(2,3,0,10)` | {0,10,20;30,40,50} | Custom start and step |

## Notes

- Omitting rows defaults it to 1 (native requires it).
- Related functions: RANDARRAY.
- Supported: Excel 2010+. Registered as `SEQUENCE` (drop-in) on hosts without the native function, and as `EG.SEQUENCE` on modern Excel that has it.
