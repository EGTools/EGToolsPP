# BARCODE

**Category**: EGTools-only function

Encodes text as a barcode and inserts the image into the calling cell (2D types keep their aspect ratio and are centered when the cell is larger).

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
- Generation uses zint; reading (READBARCODE) uses zxing-cpp. GS1-128/GS1DATAMATRIX/GS1QRCODE encode parenthesized AI text such as "(01)04012345678901(10)LOT42" to specification (FNC1 first, separators after variable-length AIs; symbology identifiers ]C1/]d2/]Q3); square-bracket [AI]value notation and a raw scanner stream such as `]C10104012345678901<GS>10LOT42` (`<GS>` = separator 0x1D — typing the literal text `<GS>` is recognized too; a leading ]C1 identifier is ignored if present) produce the same symbol. AI format/length/check-digit violations, unknown AIs and non-ASCII characters return #VALUE!.
- EAN13/EAN8/UPC-A/UPC-E accept data without the check digit (12/7/11/6–7 digits) and append it, or verify it when included (a wrong one returns #VALUE!); ITF is padded with a leading 0 when the digit count is odd.
- The option argument is a margin in module units when numeric (2D types) or TRUE to draw the text below the bars (1D types).
- 2D types keep their specified aspect ratio (square modules), fitted inside the calling cell and centered; 1D types fill the cell.
- An array in the text returns #VALUE! (one picture per cell) — for multiple items, copy the formula down row by row.
- Supported: Excel 2010+. Always registered as `BARCODE` on every Excel version.
