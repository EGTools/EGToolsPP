# DECODEURL

**Category**: EGTools-only function

Decodes a percent-encoded (URL) string back to text.

## Syntax

```
=DECODEURL(text)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | the URL-encoded text |

## Returns

Returns the text decoded from UTF-8 percent-encoding. An array input is processed element-wise and spills as an array of the same size.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DECODEURL("a%20b%26c")` | a b&c | Decode percent-encoding |
| `=DECODEURL("%ED%95%9C%EA%B8%80")` | 한글 | Decode UTF-8 Korean text |

## Notes

- Invalid % sequences are passed through unchanged.
- A + sign is not converted to a space (no form encoding).
- Inverse of ENCODEURL.
- Supported: Excel 2010+. Always registered as `DECODEURL` on every Excel version.
