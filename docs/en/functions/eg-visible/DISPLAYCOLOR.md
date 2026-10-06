# DISPLAYCOLOR

**Category**: EGTools-only function

Returns the fill or font color of each cell in a range.

## Syntax

```
=DISPLAYCOLOR(color_range, [font], [conditional])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| color_range | Required | the cell range to inspect |
| font | Optional | TRUE = font color instead of fill (default FALSE) |
| conditional | Optional | TRUE = use displayed color incl. conditional formatting (accurate only while the sheet is active); default FALSE = cell's own format color |

## Returns

Returns the fill color (font color when font=TRUE) of each cell as a spilled numeric array of the same size — the cell's own format color by default, or the displayed color with conditional formatting applied when conditional=TRUE. Returns #VALUE! when the range exceeds 100,000 cells or the COM connection/color read fails.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DISPLAYCOLOR(A1:B2)` |  | Fill color number of each cell (depends on cell formatting) |
| `=DISPLAYCOLOR(A1:B2,,TRUE)` |  | Displayed color number incl. conditional formatting (depends on cell formatting) |

## Notes

- By default (conditional omitted/FALSE) only the color set directly in the cell's format (Interior.Color / Font.Color) is used; colors painted by conditional formatting are ignored. Results stay stable whichever sheet is active, even when another sheet is edited or F9/full recalculation runs, so the default is recommended unless you really need conditional-formatting colors.
- With conditional=TRUE the actual displayed color with conditional formatting applied (DisplayFormat) is used. Conditional-formatting colors are read correctly only when recalculation runs while the formula's sheet is active.
- Caution: with conditional=TRUE, recalculation while another sheet (or another workbook) is active — editing that sheet, F9, or full recalculation alike — fails to read the conditional-formatting colors and produces wrong values. Because these functions are volatile, merely editing another sheet triggers this, so it happens often in practice.
- Caution: a wrong value produced this way stays even after you return to the formula's sheet; press F9 while the formula's sheet is active to correct it.
- Caution: if you need to aggregate by conditional-formatting color, prefer aggregating by the conditional-formatting condition itself where possible (e.g. `=COUNTIF(A1:A10,">5")`).
- As a macro-type function taking cell references, Excel treats it as volatile, so it recalculates whenever a value changes anywhere in the workbook. Changing only a cell color or a conditional-formatting rule does not trigger recalculation, though — press F9. It is also excluded from multithreaded recalculation.
- Existing formulas using up to the second argument (font) remain compatible, but formulas that relied on conditional-formatting colors must pass TRUE as the third argument to get the same results as before.
- Supported: Excel 2010+. Always registered as `DISPLAYCOLOR` on every Excel version.
