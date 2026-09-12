# DATAMATRIX

**Category**: EGTools-only function

Encodes text as a DataMatrix image in the calling cell (aspect ratio kept, centered when the cell is larger).

## Syntax

```
=DATAMATRIX(text, [margin], [gs1])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | the text to encode |
| margin | Optional | quiet-zone modules (default 0) |
| gs1 | Optional | TRUE = encode as GS1 (FNC1): (AI)value or [AI]value notation, or a raw GS-separated stream |

## Returns

Returns an empty string ("") on success; the DataMatrix picture is inserted separately, sized to the calling cell. Returns #VALUE! when the text is empty or GS1 conversion/encoding fails, and #REF! when the calling cell cannot be determined.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DATAMATRIX("LOT-2026-08")` |  | Insert a DataMatrix picture (image insertion, environment-dependent) |

## Notes

- gs1=TRUE encodes parenthesized AI text such as "(01)04012345678901(10)LOT42" as GS1 DataMatrix (symbology identifier ]d2) — FNC1 is placed in the first position and separators follow variable-length AIs as the specification requires. Square-bracket notation such as "[01]04012345678901[10]LOT42" and a raw scanner stream such as `]d20104012345678901<GS>10LOT42` (`<GS>` = separator 0x1D — typing the literal text `<GS>` is recognized too; a leading ]d2 identifier is ignored if present) produce the same symbol; unknown AIs return #VALUE!.
- GS1 data is validated for AI format, length and check digits (e.g. the (01) GTIN check digit); violations return #VALUE!, and only printable ASCII is allowed (Korean etc. returns #VALUE!). Plain text is encoded as UTF-8, with an ECI inserted automatically for non-Latin-1 characters (supported by most readers).
- The picture keeps the symbol's specified aspect ratio (square modules) while being fitted inside the calling cell, and is centered in the leftover space; margin is in module units (default 0).
- An array in the text returns #VALUE! (one picture per cell) — for multiple items, copy the formula down row by row.
- Supported: Excel 2010+. Always registered as `DATAMATRIX` on every Excel version.
