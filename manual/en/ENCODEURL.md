# ENCODEURL

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns a URL-encoded (percent-encoded) string.

## Syntax

```
=ENCODEURL(text)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text to URL-encode |

## Returns

Returns the percent-encoded text as a scalar string; an array input is mapped element-wise and spilled as an array of the same shape. Inputs are converted to text, so no specific error values are produced.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ENCODEURL("a b&c")` | a%20b%26c | space and & percent-encoded |
| `=ENCODEURL("한글")` | %ED%95%9C%EA%B8%80 | UTF-8 bytes of Korean text |

## Notes

- Only RFC 3986 unreserved characters (letters, digits, - _ . ~) are kept; everything else is %XX-encoded per UTF-8 byte.
- Supported: Excel 2010+. Registered as `ENCODEURL` (drop-in) on hosts without the native function, and as `EG.ENCODEURL` on modern Excel that has it.
