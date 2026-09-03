# ISVISIBLE

**Category**: EGTools-only function

Returns whether each cell in a range is visible (row and column not hidden).

## Syntax

```
=ISVISIBLE(range)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| range | Required | the cell range to check |

## Returns

Returns a spilled array of the same size that is TRUE where the cell is visible and FALSE where its row or column is hidden. Returns #VALUE! when the range exceeds 100,000 cells or the COM connection fails.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ISVISIBLE(A1:A5)` |  | Visibility of each cell (depends on hidden state) |

## Notes

- Hiding or unhiding rows/columns alone does not trigger recalculation — press F9.
- Macro-type function, excluded from multithreaded recalculation.
- Supported: Excel 2010+. Always registered as `ISVISIBLE` on every Excel version.
