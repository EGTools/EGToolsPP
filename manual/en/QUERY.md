# QUERY

**Category**: Google Sheets compatibility function

Runs a SQL query over the data (table name: EGTOOLS).

## Syntax

```
=QUERY(data, query, [headers])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| data | Required | range or array to query |
| query | Required | SQL text, e.g. "SELECT F1, SUM(F2) FROM EGTOOLS GROUP BY F1" |
| headers | Optional | number of header rows: N>=1 top rows are field names, 0 none (fields F1, F2, …), -1/omitted auto-detect |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=QUERY(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `QUERY` on every Excel version.
