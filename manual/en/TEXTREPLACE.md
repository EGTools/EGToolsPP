# TEXTREPLACE

**Category**: EGTools-only function

Replaces every StartKey…EndKey section of the text.

## Syntax

```
=TEXTREPLACE(text, start_key, end_key, [replace], [include])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| start_key | Required | section start marker |
| end_key | Required | section end marker |
| replace | Optional | replacement text (default empty) |
| include | Optional | TRUE keeps the markers around the replacement |

## Returns

Returns the text with every StartKey..EndKey section replaced, as a single scalar. When text, start key, or end key is empty the original text is returned unchanged instead of an error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTREPLACE("a[x]b[y]c","[","]","*")` | a*b*c | replace every section |
| `=TEXTREPLACE("a<1>b","<",">","2",TRUE)` | a<2>b | keep the markers |

## Notes

- Key matching ignores case.
- A start key without a matching end key is itself replaced and the rest of the text is kept.
- If the text is an array, it is processed element-wise and spills an array of the same shape (an error element returns that error); an array in start_key, end_key and the other arguments returns #VALUE!.
- Related function: TEXTBETWEEN
- Supported: Excel 2010+. Always registered as `TEXTREPLACE` on every Excel version.
