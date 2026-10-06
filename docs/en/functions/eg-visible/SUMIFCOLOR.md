# SUMIFCOLOR

**Category**: EGTools-only function

Sums cells whose color matches a reference cell.

## Syntax

```
=SUMIFCOLOR(search_range, color_cell, [font], [conditional])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| search_range | Required | the range to sum |
| color_cell | Required | cell holding the reference color (first cell used) |
| font | Optional | TRUE = compare font color (default FALSE) |
| conditional | Optional | TRUE = use displayed color incl. conditional formatting (accurate only while the sheet is active); default FALSE = cell's own format color |

## Returns

Returns the sum of numeric cells with the same color as the reference cell (the cell's own format color by default, the displayed color when conditional=TRUE). Returns #VALUE! on COM failure, unreadable reference color, or more than 100,000 scanned cells; an error value in a matching cell is returned as-is.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SUMIFCOLOR(A1:A10,C1)` |  | Sum of cells with the same fill color as C1 (depends on cell formatting) |
| `=SUMIFCOLOR(A1:A10,C1,,TRUE)` |  | Sum based on conditional-formatting colors (depends on cell formatting) |

## Notes

- By default (conditional omitted/FALSE) only the color set directly in the cell's format (Interior.Color / Font.Color) is used; colors painted by conditional formatting are ignored. Results stay stable whichever sheet is active, even when another sheet is edited or F9/full recalculation runs, so the default is recommended unless you really need conditional-formatting colors.
- With conditional=TRUE the actual displayed color with conditional formatting applied (DisplayFormat) is used. Conditional-formatting colors are read correctly only when recalculation runs while the formula's sheet is active.
- Caution: with conditional=TRUE, recalculation while another sheet (or another workbook) is active — editing that sheet, F9, or full recalculation alike — fails to read the conditional-formatting colors and produces wrong values. Because these functions are volatile, merely editing another sheet triggers this, so it happens often in practice.
- Caution: a wrong value produced this way stays even after you return to the formula's sheet; press F9 while the formula's sheet is active to correct it.
- Caution: if you need to aggregate by conditional-formatting color, prefer aggregating by the conditional-formatting condition itself where possible (e.g. `=SUMIF(A1:A10,">5")`).
- As a macro-type function taking cell references, Excel treats it as volatile, so it recalculates whenever a value changes anywhere in the workbook. Changing only a cell color or a conditional-formatting rule does not trigger recalculation, though — press F9. It is also excluded from multithreaded recalculation.
- Existing formulas using up to the third argument (font) remain compatible, but formulas that relied on conditional-formatting colors must pass TRUE as the fourth argument to get the same results as before.
- A merged area is treated as one cell (only its first cell is checked; its value lives there).
- Supported: Excel 2010+. Always registered as `SUMIFCOLOR` on every Excel version.
