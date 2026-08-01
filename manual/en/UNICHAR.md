# UNICHAR

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the Unicode character for the given number (code point).

## Syntax

```
=UNICHAR(number)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number | Required | Unicode code point |

## Returns

Returns the character (scalar text) for the code point; an array of numbers is evaluated element-wise and spills. Returns #VALUE! when the code is outside 1..1114111 or in the surrogate range (55296..57343).

## Examples

| Formula | Result | Description |
|---|---|---|
| `=UNICHAR(44032)` | 가 | Hangul syllable |
| `=UNICHAR(128512)` | 😀 | beyond the BMP |
| `=UNICHAR(0)` | #VALUE! | out of range |

## Notes

- Characters above 0xFFFF are returned as a UTF-16 surrogate pair.
- Related function: UNICODE
- Supported: Excel 2010+. Registered as `UNICHAR` (drop-in) on hosts without the native function, and as `EG.UNICHAR` on modern Excel that has it.
