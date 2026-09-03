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

Returns the SQL result as a spilled 2-D array (data rows only, no header row). Missing data/query or an empty query gives #VALUE!; no result rows gives #N/A; ADO/ACE provider errors and temp-CSV write failures are returned as "ERROR: …" text values.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=QUERY(A1:C10,"SELECT F1, SUM(F3) FROM EGTOOLS GROUP BY F1",0)` |  | result depends on the range data |

## Notes

- Uses SQL via the ACE text driver + ADO, not Google Query Language; the table name is EGTOOLS (occurrences of EGTOOLS in the query are rewritten to EGTOOLS#csv).
- Column names come from the header row(s) selected by headers, otherwise F1, F2, … (headers: N>=1 top rows are field names, 0 none, omitted/-1 auto-detect).
- The data is written to %TEMP%\EGTOOLS.csv and queried there, so the Microsoft ACE OLEDB provider (12.0/16.0) is required; the shared temp file makes it non-thread-safe.
- Supported: Excel 2010+. Always registered as `QUERY` on every Excel version.
