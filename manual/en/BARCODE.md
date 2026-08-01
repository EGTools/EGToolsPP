# BARCODE

**Category**: EGTools-only function

Encodes text as a barcode and inserts the image into the calling cell.

## Syntax

```
=BARCODE(text, [type], [option])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | the text to encode |
| type | Optional | QRCODE (default), CODE128, DATAMATRIX, EAN13/8, ITF, PDF417, AZTEC, CODE39/93, UPC-A/E, CODABAR, GS1-128/DM/QR |
| option | Optional | number = margin (2D types) or TRUE = show text below (1D types) |

## Returns

Returns an empty string ("") on success; the barcode picture is inserted separately, sized to the calling cell. Returns #VALUE! when the text is empty, the type is not recognized, or the text cannot be encoded in that format, and #REF! when the calling cell cannot be determined.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BARCODE("12345678","CODE128")` |  | Insert a Code 128 picture (image insertion, environment-dependent) |
| `=BARCODE("https://example.com")` |  | Insert the default QR code (image insertion, environment-dependent) |

## Notes

- Supported types: CODE128, CODE39/93, EAN13/8, UPC-A/E, ITF, CODABAR, PDF417, QRCODE (default), DATAMATRIX, AZTEC, GS1-128/GS1DATAMATRIX/GS1QRCODE (numeric codes also accepted).
- The option argument is a margin when numeric (2D types) or TRUE to draw the text below the bars (1D types).
- The image is sized to the calling cell; 2D types keep their aspect ratio.
- Supported: Excel 2010+. Always registered as `BARCODE` on every Excel version.
