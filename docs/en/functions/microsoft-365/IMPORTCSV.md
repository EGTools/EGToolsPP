# IMPORTCSV

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365 (new in 2026)

Imports a CSV file (local path or URL) as an array — comma delimiter, UTF-8.

## Syntax

```
=IMPORTCSV(path, [skip_rows], [take_rows], [locale])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| path | Required | local file path or URL |
| skip_rows | Optional | rows to skip; negative skips from the end |
| take_rows | Optional | rows to take; negative takes from the end |
| locale | Optional | locale for number parsing, e.g. "de-DE" |

## Returns

Returns the CSV content as a spilled 2-D array; numeric and date (yyyy-M-d style) text is auto-converted to values. A missing path or a file-read/download failure gives #VALUE!; if no rows remain the result is #N/A.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTCSV("C:\data\sales.csv",1)` |  | result depends on the file |

## Notes

- Delimiter is fixed to comma and encoding to UTF-8 (a BOM takes precedence); use IMPORTTEXT for other delimiters or encodings.
- Double-quoted fields (including "" escapes) are recognized.
- Negative skip_rows/take_rows count from the end.
- An array in the path and take_rows returns #VALUE! — for multiple items, copy the formula down row by row.
- Supported: Excel 2010+. Registered as `IMPORTCSV` (drop-in) on hosts without the native function, and as `EG.IMPORTCSV` on modern Excel that has it.
