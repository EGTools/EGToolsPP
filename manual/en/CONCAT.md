# CONCAT

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Concatenates text from values and ranges (no delimiter).

## Syntax

```
=CONCAT(text1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text1 | Required | text or range |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns a single text value (scalar) concatenating every argument (including ranges/arrays) in order. If any argument contains an error value, that error is returned.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CONCAT("A",1,"B")` | A1B | concatenate values as-is |
| `=CONCAT({1,2;3,4})` | 1234 | array joined row-major |

## Notes

- No delimiter is inserted; use TEXTJOIN when you need one.
- Up to 255 arguments are supported.
- Supported: Excel 2010+. Registered as `CONCAT` (drop-in) on hosts without the native function, and as `EG.CONCAT` on modern Excel that has it.
