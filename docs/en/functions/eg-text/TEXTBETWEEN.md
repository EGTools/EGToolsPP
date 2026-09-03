# TEXTBETWEEN

**Category**: EGTools-only function

Extracts the text between StartKey and EndKey.

## Syntax

```
=TEXTBETWEEN(text, left_text, right_text, [delimiter_or_instance], [include])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| left_text | Required | start marker |
| right_text | Required | end marker |
| delimiter_or_instance | Optional | joiner for all matches (default ","), or a number N for the Nth match only |
| include | Optional | TRUE includes the markers |

## Returns

Returns the extracted text as a single scalar. Returns #VALUE! when text, start key, or end key is empty, or when the Nth-instance number is out of range; returns an empty string when no section matches.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTBETWEEN("[a] and [b]","[","]")` | a,b | join all sections with , |
| `=TEXTBETWEEN("[a] and [b]","[","]",2)` | b | 2nd section only |
| `=TEXTBETWEEN("<x><y>","<",">","/",TRUE)` | <x>/<y> | include the markers |

## Notes

- Key matching ignores case.
- If the 4th argument is a number it selects the Nth section; if text, it is the joiner for all sections (default ",").
- If the text is an array, it is processed element-wise and spills an array of the same shape (an error element returns that error); an array in start_key, end_key and the other arguments returns #VALUE!.
- Related function: TEXTREPLACE
- Supported: Excel 2010+. Always registered as `TEXTBETWEEN` on every Excel version.
