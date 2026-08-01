# VISIBLEAGGR

**Category**: EGTools-only function

Lists or aggregates only the visible cells of a range.

## Syntax

```
=VISIBLEAGGR(range, [function], [option])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| range | Required | the cell range |
| function | Optional | SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S·P/VAR.S·P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC); omitted = spill visible cells |
| option | Optional | TEXTJOIN delimiter, LARGE/SMALL k, PERCENTILE p (0-1), QUARTILE quart (0-4) |

## Returns

Without a function, spills the visible cells as a 2-D array; with a function, returns a scalar aggregate. No visible cells gives #N/A; a visible error cell is returned as-is except for COUNT/COUNTA; an unsupported function name gives #VALUE!.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=VISIBLEAGGR(A1:B10)` |  | spill visible cells (depends on hidden state) |
| `=VISIBLEAGGR(A1:A10,"SUM")` |  | sum of visible cells (depends on hidden state) |
| `=VISIBLEAGGR(A1:A10,"LARGE",2)` |  | 2nd largest among visible values |

## Notes

- Supported aggregators (21): SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE(.INC/.EXC)/QUARTILE(.INC/.EXC). Anything else returns #VALUE!.
- Option argument: TEXTJOIN = delimiter, LARGE/SMALL = k, PERCENTILE family = p (0-1), QUARTILE family = 0-4.
- Changing row/column visibility alone does not recalculate - press F9. Macro-type, excluded from multithreaded recalculation.
- Supported: Excel 2010+. Always registered as `VISIBLEAGGR` on every Excel version.
