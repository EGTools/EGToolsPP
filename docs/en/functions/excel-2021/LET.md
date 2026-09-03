# LET

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Assigns names to values and uses them in a final calculation.

## Syntax

```
=LET(name1, value1, calculation, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| name1 | Required | name to define |
| value1 | Required | value for this name |
| calculation | Required | the calculation using the names above |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns the result of evaluating the final calculation with each name substituted by its value; an array result spills. Returns #VALUE! when the argument structure is invalid or the calling cell's formula cannot be read; when evaluation fails, the substituted formula string is returned instead of an error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=LET(x,3,x*2)` | 6 | Define a name, then compute |
| `=LET(x,2,y,x+1,x*y)` | 6 | Values may use earlier names |

## Notes

- Unlike native LET, no names are bound: this macro-sheet function parses the calling cell's formula text, substitutes names with their values and evaluates via Application.Evaluate.
- The names "R" and "C" are reserved by Excel's R1C1 parser and cannot be used.
- When evaluation fails, the substituted formula string is returned for debugging.
- Supported: Excel 2010+. Registered as `LET` (drop-in) on hosts without the native function, and as `EG.LET` on modern Excel that has it. On some legacy hosts it registers as `xLET` due to a keyword conflict.
