# QRCODE

**Category**: EGTools-only function

Encodes text as a QR code image in the calling cell.

## Syntax

```
=QRCODE(text, [margin], [gs1])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | the text to encode |
| margin | Optional | quiet-zone modules (default 0) |
| gs1 | Optional | TRUE = treat text as GS1 (AI)value pairs |

## Returns

Returns an empty string ("") on success; the QR code picture is inserted separately, sized to the calling cell. Returns #VALUE! when the text is empty or GS1 conversion/encoding fails, and #REF! when the calling cell cannot be determined.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=QRCODE("https://example.com")` |  | Insert a QR code picture (image insertion, environment-dependent) |

## Notes

- gs1=TRUE approximates GS1 with GS (0x1D) separators — not full FNC1 encoding; use CODE128 (GS1-128) for full compliance.
- The image keeps its square aspect ratio within the calling cell.
- Supported: Excel 2010+. Always registered as `QRCODE` on every Excel version.
