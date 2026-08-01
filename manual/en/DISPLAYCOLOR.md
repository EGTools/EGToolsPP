# DISPLAYCOLOR

**Category**: EGTools-only function

Returns the displayed fill or font color of each cell in a range.

## Syntax

```
=DISPLAYCOLOR(color_range, [font])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| color_range | Required | the cell range to inspect |
| font | Optional | TRUE = font color instead of fill (default FALSE) |

## Returns

Returns the displayed fill color (font color when font=TRUE) of each cell as a spilled numeric array of the same size. Returns #VALUE! when the range exceeds 100,000 cells or the COM connection/color read fails.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DISPLAYCOLOR(A1:B2)` |  | Displayed color number of each cell (depends on cell formatting) |

## Notes

- Based on DisplayFormat, so the returned color reflects conditional formatting.
- Changing a format alone does not trigger recalculation — press F9.
- Macro-type function, excluded from multithreaded recalculation.
- Supported: Excel 2010+. Always registered as `DISPLAYCOLOR` on every Excel version.
