# READGS1

**Category**: EGTools-only function

Splits GS1 text into (AI), value rows.

## Syntax

```
=READGS1(gs1_text, [concat])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| gs1_text | Required | (AI)value form or a decoded barcode string |
| concat | Optional | TRUE = join into one (AI)value string |

## Returns

Splits GS1 text into a spilled N-by-2 array of (AI) and value (one joined (AI)value string with concat=TRUE). Returns #VALUE! for empty input, #NUM! when an unknown AI is met in a decoded stream, and #N/A when nothing can be split.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=READGS1("(01)04012345678901(10)LOT42")` | {"(01)","04012345678901";"(10)","LOT42"} | Split into (AI), value rows |
| `=READGS1("(01)04012345678901(10)LOT42",TRUE)` | (01)04012345678901(10)LOT42 | Join back into one string |

## Notes

- Accepts both the bracketed (AI)value form and a raw decoded stream with GS separators.
- Uses the fixed-length and variable-length GS1 AI tables to split a stream.
- Supported: Excel 2010+. Always registered as `READGS1` on every Excel version.
