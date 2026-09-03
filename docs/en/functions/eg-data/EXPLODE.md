# EXPLODE

**Category**: EGTools-only function

Splits one column by delimiter(s), expanding into rows or columns; other columns are copied.

## Syntax

```
=EXPLODE(array, col_number, [delimiter], [by_col], [pad_with], [ignore_empty])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | source table |
| col_number | Required | column to split, 1-based |
| delimiter | Optional | delimiter(s), e.g. " " or {",",";"} (default space) |
| by_col | Optional | TRUE spreads parts across columns (default FALSE: down rows) |
| pad_with | Optional | fill value for short rows when by_col (default "") |
| ignore_empty | Optional | TRUE drops empty parts |

## Returns

Returns the exploded table as a 2-D spilled array. Returns #VALUE! when the array argument is not an array, and #NUM! when col_number is outside 1..number-of-columns.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=EXPLODE({"A","x,y";"B","z"},2,",")` | {"A","x";"A","y";"B","z"} | Split column 2 into rows |
| `=EXPLODE({"A","x,y";"B","z"},2,",",TRUE)` | {"A","x","y";"B","z",""} | Spread across columns, padded |

## Notes

- Multiple delimiters are allowed; longer delimiters are applied first (default: space).
- With by_col TRUE, short rows are padded with pad_with (default "").
- ignore_empty TRUE drops empty parts (one empty string is kept if all parts are empty).
- Supported: Excel 2010+. Always registered as `EXPLODE` on every Excel version.
