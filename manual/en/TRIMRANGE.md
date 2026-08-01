# TRIMRANGE

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Trims blank rows/columns from the edges of an array.

## Syntax

```
=TRIMRANGE(array, [trim_rows], [trim_cols])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array or range values |
| trim_rows | Optional | 0 none, 1 leading, 2 trailing, 3 both (default) |
| trim_cols | Optional | 0 none, 1 leading, 2 trailing, 3 both (default) |

## Returns

Returns the array with blank edge rows/columns removed, spilled. Returns #VALUE! when the array is empty, a trim mode is outside 0-3, or every cell is blank.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TRIMRANGE({"","";"a",1;"",""})` | {"a",1} | Trim blank edge rows and columns |
| `=TRIMRANGE({"";1;""},1)` | {1;""} | Trim leading blank rows only |

## Notes

- Native TRIMRANGE trims a reference; this implementation is value-based, so array literals and formula results work too.
- Empty strings ("") count as blank.
- Supported: Excel 2010+. Registered as `TRIMRANGE` (drop-in) on hosts without the native function, and as `EG.TRIMRANGE` on modern Excel that has it.
