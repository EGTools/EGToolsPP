# VALUETOTEXT

**Category**: MS Excel compatibility function · **Native since**: Excel 2024

Returns the text representation of a value.

## Syntax

```
=VALUETOTEXT(value, [format])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| value | Required | the value |
| format | Optional | 0 concise (default), 1 strict (quotes text) |

## Returns

Returns the text representation of the value as a single text (scalar). The implementation itself returns no error values.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=VALUETOTEXT(1.5)` | 1.5 | number unchanged |
| `=VALUETOTEXT("abc",1)` | "abc" | strict quotes text |
| `=VALUETOTEXT(TRUE)` | TRUE | logical value |

## Notes

- In strict (1) format, quotes inside text are doubled ("").
- Related function: ARRAYTOTEXT
- Supported: Excel 2010+. Registered as `VALUETOTEXT` (drop-in) on hosts without the native function, and as `EG.VALUETOTEXT` on modern Excel that has it.
