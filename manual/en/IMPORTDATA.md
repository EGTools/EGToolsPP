# IMPORTDATA

**Category**: Google Sheets compatibility function

Imports CSV/TSV data from a URL or local file path.

## Syntax

```
=IMPORTDATA(url_or_path, [delimiter], [locale], [charset])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| url_or_path | Required | web URL or local file path providing the data |
| delimiter | Optional | field delimiter (default: auto-detect tab/comma) |
| locale | Optional | locale for number parsing, e.g. "de-DE" |
| charset | Optional | encoding, e.g. "utf-8" (default), "euc-kr", or a codepage number |

## Returns

Returns the CSV/TSV content of a URL or local file as a spilled 2-D array; numeric and date text is auto-converted to values. A missing path, download/read failure, or invalid delimiter spec gives #VALUE!; empty content gives #N/A.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTDATA("https://example.com/data.csv")` |  | result depends on the external data |

## Notes

- When delimiter is omitted it auto-detects: tab if the first line contains a tab, otherwise comma.
- Legacy calls passing a charset name as the second argument (IMPORTDATA(url,"euc-kr")) are still accepted for backward compatibility.
- Interior empty lines are kept; only trailing empty lines are dropped.
- Supported: Excel 2010+. Always registered as `IMPORTDATA` on every Excel version.
