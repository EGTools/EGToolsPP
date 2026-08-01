# SUMIFCOLOR

**Category**: EGTools-only function

Sums cells whose displayed color matches a reference cell.

## Syntax

```
=SUMIFCOLOR(search_range, color_cell, [font])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| search_range | Required | the range to sum |
| color_cell | Required | cell holding the reference color (first cell used) |
| font | Optional | TRUE = compare font color (default FALSE) |

## Returns

Returns the sum of numeric cells whose displayed color matches the reference cell. Returns #VALUE! on COM failure, unreadable reference color, or more than 100,000 scanned cells; an error value in a matching cell is returned as-is.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SUMIFCOLOR(A1:A10,C1)` |  | Sum of cells with the same fill color as C1 (depends on cell formatting) |

## Notes

- DisplayFormat-based, so colors from conditional formatting are compared as displayed.
- Format changes alone do not trigger recalculation (press F9); macro-type, excluded from multithreaded recalculation.
- Supported: Excel 2010+. Always registered as `SUMIFCOLOR` on every Excel version.
