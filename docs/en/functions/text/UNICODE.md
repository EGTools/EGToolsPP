# UNICODE

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the Unicode code point of the first character of the text.

## Syntax

```
=UNICODE(text)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text whose first character is examined |

## Returns

Returns the code point (number, scalar) of the first character; an array of texts is evaluated element-wise and spills. Returns #VALUE! for empty text.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=UNICODE("A")` | 65 | Latin letter |
| `=UNICODE("가")` | 44032 | Hangul syllable |
| `=UNICODE("😀")` | 128512 | beyond the BMP |

## Notes

- Surrogate pairs are decoded, so characters beyond the BMP return the full code point.
- Related function: UNICHAR
- Supported: Excel 2010+. Registered as `UNICODE` (drop-in) on hosts without the native function, and as `EG.UNICODE` on modern Excel that has it.
