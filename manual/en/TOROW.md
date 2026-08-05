# TOROW

**Category**: MS Excel compatibility function · **Native since**: Excel 2024

Returns the array as a single row.

## Syntax

```
=TOROW(array, [ignore], [scan_by_column])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| ignore | Optional | 0 none, 1 blanks, 2 errors, 3 both |
| scan_by_column | Optional | TRUE to scan by column |

## Returns

Returns every value of the array laid out as a single row (1 x N); on dynamic-array hosts the result spills. The ignore option (1=blanks, 2=errors, 3=both) filters values out. Returns #VALUE! if the array is empty or no values remain.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TOROW({1,2;3,4})` | {1,2,3,4} | Row-major flatten to one row |
| `=TOROW({1,2;3,4},0,TRUE)` | {1,3,2,4} | Scan by column |
| `=TOROW({1,"",2},1)` | {1,2} | Ignore blanks |

## Notes

- Empty strings ("") count as blanks for ignore options 1 and 3.
- If every value is filtered out, returns #VALUE! instead of the native #CALC!.
- Related function: TOCOL
- Supported: Excel 2010+. Registered as `TOROW` (drop-in) on hosts without the native function, and as `EG.TOROW` on modern Excel that has it.
