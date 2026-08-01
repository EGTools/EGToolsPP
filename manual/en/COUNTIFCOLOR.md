# COUNTIFCOLOR

**Category**: EGTools-only function

Counts cells whose displayed color matches a reference cell.

## Syntax

```
=COUNTIFCOLOR(search_range, color_cell, [font])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| search_range | Required | the range to check |
| color_cell | Required | cell holding the reference color (first cell used) |
| font | Optional | TRUE = compare font color (default FALSE) |

## Returns

Returns the number of cells whose displayed color matches the reference cell. Returns #VALUE! on COM failure, unreadable reference color, or more than 100,000 scanned cells, and 0 when the range does not overlap the used range.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=COUNTIFCOLOR(A1:A10,C1)` |  | Cells with the same fill color as C1 (depends on cell formatting) |

## Notes

- With font=TRUE only non-empty cells are counted.
- DisplayFormat-based, so colors from conditional formatting are compared as displayed.
- Format changes alone do not trigger recalculation (press F9); macro-type, excluded from multithreaded recalculation.
- Supported: Excel 2010+. Always registered as `COUNTIFCOLOR` on every Excel version.
