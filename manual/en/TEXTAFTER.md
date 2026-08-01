# TEXTAFTER

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Returns the text after a delimiter.

## Syntax

```
=TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| delimiter | Required | delimiter to search for |
| instance_num | Optional | which occurrence, negative from the end (default 1) |
| match_mode | Optional | 1 to ignore case (default 0) |
| match_end | Optional | 1 treats the end of text as a delimiter (default 0) |
| if_not_found | Optional | value when not found (default #N/A) |

## Returns

Returns the single text (scalar) after the delimiter. Returns #VALUE! when instance_num is 0, and the if_not_found value (default #N/A) when the delimiter is not found.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTAFTER("red-blue-green","-")` | blue-green | after first delimiter |
| `=TEXTAFTER("a,b,c",",",-1)` | c | after last delimiter |
| `=TEXTAFTER("abc","x",1,0,0,"none")` | none | if_not_found value |

## Notes

- match_end=1 treats the end of the text (positive instance) or its start (negative instance) as one virtual delimiter.
- match_mode=1 ignores case.
- Related functions: TEXTBEFORE, TEXTSPLIT
- Supported: Excel 2010+. Registered as `TEXTAFTER` (drop-in) on hosts without the native function, and as `EG.TEXTAFTER` on modern Excel that has it.
