# CODE128

**Category**: EGTools-only function

Encodes text as a Code 128 barcode image in the calling cell.

## Syntax

```
=CODE128(text, [show_text], [gs1], [text_size])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | the text to encode |
| show_text | Optional | TRUE = draw the text under the bars |
| gs1 | Optional | TRUE = GS1-128 with FNC1 |
| text_size | Optional | relative label size |

## Returns

Returns an empty string ("") on success; the Code 128 barcode picture is inserted separately, sized to the calling cell. Returns #VALUE! when the text is empty or GS1 conversion/encoding fails, and #REF! when the calling cell cannot be determined.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=CODE128("ABC-1234",TRUE)` |  | Insert with the text drawn below the bars (image insertion, environment-dependent) |
| `=CODE128("(01)04012345678901(10)LOT42",FALSE,TRUE)` |  | Insert as GS1-128 with FNC1 (image insertion, environment-dependent) |

## Notes

- gs1=TRUE encodes (AI)value text as fully FNC1-compliant GS1-128 (symbology identifier ]C1).
- The image fills the calling cell (1D format).
- text_size is a relative value that scales the text band height.
- An array in the text returns #VALUE! (one picture per cell) — for multiple items, copy the formula down row by row.
- Supported: Excel 2010+. Always registered as `CODE128` on every Excel version.
