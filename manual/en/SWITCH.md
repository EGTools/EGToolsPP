# SWITCH

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Compares an expression to values and returns the first matching result (or default).

## Syntax

```
=SWITCH(expression, value1, result1, [default], ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| expression | Required | value to compare |
| value1 | Required | value to match |
| result1 | Required | result if the value matches |
| default | Optional | result if no value matches (optional) |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns the result paired with the first value equal to the expression (or the trailing default), as-is. Returns #N/A when no value matches and no default is given.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SWITCH(2,1,"one",2,"two","other")` | two | Matching value's result |
| `=SWITCH(9,1,"one","other")` | other | Default returned |

## Notes

- Leaving a middle argument empty truncates the list; later value/result pairs are ignored.
- Related functions: IFS.
- Supported: Excel 2010+. Registered as `SWITCH` (drop-in) on hosts without the native function, and as `EG.SWITCH` on modern Excel that has it.
