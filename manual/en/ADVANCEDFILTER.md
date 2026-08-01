# ADVANCEDFILTER

**Category**: EGTools-only function

Filters a table with an Advanced-Filter style criteria grid (rows=OR, columns=AND).

## Syntax

```
=ADVANCEDFILTER(data_array, filter_array, [unique])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| data_array | Required | data with header row |
| filter_array | Required | criteria grid with header row; supports =, <>, >, <, >=, <=, * and ? |
| unique | Optional | TRUE removes duplicate result rows |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ADVANCEDFILTER(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `ADVANCEDFILTER` on every Excel version (no `EG.`/`x` prefix).
