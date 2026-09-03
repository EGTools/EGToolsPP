# XOR

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the logical exclusive OR of the arguments (TRUE for an odd count of TRUEs).

## Syntax

```
=XOR(logical1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| logical1 | Required | value or range to test |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns TRUE when an odd number of values are TRUE, otherwise FALSE, as a logical scalar. The first error among the arguments is returned as-is; with no logical values at all, #VALUE! is returned.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=XOR(TRUE,FALSE)` | TRUE | Odd number of TRUEs |
| `=XOR(TRUE,TRUE)` | FALSE | Even number of TRUEs |
| `=XOR(1,0,1,"text")` | FALSE | Text ignored |

## Notes

- Text and blank values are ignored (native returns #VALUE! for unconvertible text).
- Supported: Excel 2010+. Registered as `XOR` (drop-in) on hosts without the native function, and as `EG.XOR` on modern Excel that has it.
