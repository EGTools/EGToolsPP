# DATAMATRIX

**Category**: EGTools-only function

Encodes text as a DataMatrix image in the calling cell.

## Syntax

```
=DATAMATRIX(text, [margin], [gs1])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | the text to encode |
| margin | Optional | quiet-zone modules (default 0) |
| gs1 | Optional | TRUE = treat text as GS1 (AI)value pairs |

## Returns

Returns an empty string ("") on success; the DataMatrix picture is inserted separately, sized to the calling cell. Returns #VALUE! when the text is empty or GS1 conversion/encoding fails, and #REF! when the calling cell cannot be determined.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DATAMATRIX("LOT-2026-08")` |  | Insert a DataMatrix picture (image insertion, environment-dependent) |

## Notes

- gs1=TRUE approximates GS1 with GS (0x1D) separators — not full FNC1 encoding; use CODE128 (GS1-128) for full compliance.
- The image keeps its square aspect ratio within the calling cell.
- An array in the text returns #VALUE! (one picture per cell) — for multiple items, copy the formula down row by row.
- Supported: Excel 2010+. Always registered as `DATAMATRIX` on every Excel version.
