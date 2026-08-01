# VISIBLEAGGR

**Category**: EGTools-only function

Lists or aggregates only the visible cells of a range.

## Syntax

```
=VISIBLEAGGR(range, [function])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| range | Required | the cell range |
| function | Optional | SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT; omitted = spill visible cells |

## Returns

With function omitted, spills an array of only the visible cells; with SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT returns the aggregate over visible cells. Returns #N/A when no cell is visible, #VALUE! for an unknown function name, and #DIV/0! for AVERAGE with no numbers.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=VISIBLEAGGR(A1:A10,"SUM")` |  | Sum of visible cells (depends on hidden state) |
| `=VISIBLEAGGR(A1:B10)` |  | List only visible cells (depends on hidden state) |

## Notes

- Error cells are skipped when aggregating.
- Changing hidden state alone does not trigger recalculation (press F9); macro-type, excluded from multithreaded recalculation.
- Supported: Excel 2010+. Always registered as `VISIBLEAGGR` on every Excel version.
