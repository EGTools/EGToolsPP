# TEXTBEFORE

**Category**: MS Excel compatibility function · **Native since**: Excel 2024

Returns the text before a delimiter.

## Syntax

```
=TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
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

Returns the single text (scalar) before the delimiter. Returns #VALUE! when instance_num is 0, and the if_not_found value (default #N/A) when the delimiter is not found.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTBEFORE("red-blue-green","-")` | red | before first delimiter |
| `=TEXTBEFORE("a,b,c",",",-1)` | a,b | before last delimiter |
| `=TEXTBEFORE("abc","x")` | #N/A | delimiter not found |

## Notes

- match_end=1 treats the end of the text (positive instance) or its start (negative instance) as one virtual delimiter.
- match_mode=1 ignores case.
- If text is an array, the function is applied element-wise and spills an array of the same shape (an error element returns that error); an array delimiter returns #VALUE!.
- Related functions: TEXTAFTER, TEXTSPLIT
- Supported: Excel 2010+. Registered as `TEXTBEFORE` (drop-in) on hosts without the native function, and as `EG.TEXTBEFORE` on modern Excel that has it.
