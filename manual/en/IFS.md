# IFS

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Checks conditions in order and returns the value for the first TRUE condition.

## Syntax

```
=IFS(condition1, value1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| condition1 | Required | 1st condition |
| value1 | Required | result if condition1 is TRUE |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns the value paired with the first TRUE condition, as-is (scalar or array). Returns #N/A when no condition is TRUE; an error value in a condition is returned as-is.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IFS(1>2,"a",2>1,"b")` | b | First TRUE condition wins |
| `=IFS(1>2,"a")` | #N/A | No TRUE condition |

## Notes

- Text or blank conditions are treated as FALSE (native returns #VALUE!).
- Conditions and values accept arrays and are evaluated element-wise, spilling an array of the same shape; if an evaluated condition is an error, that element becomes the error (errors in unused arguments do not propagate).
- Related functions: SWITCH.
- Supported: Excel 2010+. Registered as `IFS` (drop-in) on hosts without the native function, and as `EG.IFS` on modern Excel that has it.
