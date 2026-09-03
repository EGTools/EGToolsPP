# READBARCODE

**Category**: EGTools-only function

Decodes a barcode from an image file or URL.

## Syntax

```
=READBARCODE(source, [result_type])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| source | Required | local file path or http(s) URL |
| result_type | Optional | 1 text (default), 2 code text, 3 format, 4 symbology id, 5 raw hex; array like {1,3} allowed |

## Returns

Returns the decoded barcode content shaped like the result_type argument (spills for an array like {1,3}). Returns #VALUE! for an empty source or an invalid result_type item, #NULL! when a URL download fails, and #N/A when the file is missing, the image cannot be read, or no barcode is found.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=READBARCODE("C:\barcodes\qr.png")` |  | Barcode text in the image (result depends on the file) |
| `=READBARCODE("https://example.com/code.png",{1,3})` |  | Text and format together (network-dependent) |

## Notes

- Reads WIC-supported formats (PNG/JPG/BMP/GIF) from a local file path or an http(s) URL.
- result_type: 1 text (default), 2 code text (control characters shown as <GS> etc.), 3 format, 4 symbology id, 5 raw hex.
- Tries rotated/low-quality images too (TryHarder/TryRotate) but decodes only one barcode.
- An array in source returns #VALUE! — for multiple items, copy the formula down row by row.
- Supported: Excel 2010+. Always registered as `READBARCODE` on every Excel version.
