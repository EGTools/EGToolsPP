# IMPORTRANGE

**Category**: Google Sheets compatibility function

Imports data from a shared Google Sheets spreadsheet.

## Syntax

```
=IMPORTRANGE(spreadsheet_url, [range_address])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| spreadsheet_url | Required | Google Sheets URL (sharing must be enabled) |
| range_address | Optional | range to import, e.g. "A1:C10" (default: whole sheet) |

## Returns

Returns the values of a shared Google Sheets spreadsheet as a spilled 2-D array; numeric and date text is auto-converted. An empty URL or one without the /d/<ID> pattern gives #VALUE!; a download failure, an unshared sheet (HTML login page), an empty sheet, or a range outside the data gives #N/A.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1AbCdEf.../edit#gid=0","A1:C10")` |  | result depends on the sheet |

## Notes

- Only link-shared Google Sheets can be imported — unshared documents return an HTML login page, which yields #N/A.
- The sheet is selected by the gid in the URL (default gid=0); any sheet-name part of range_address is ignored.
- Data comes from the TSV export endpoint, so only values are imported, without formatting.
- Supported: Excel 2010+. Always registered as `IMPORTRANGE` on every Excel version.
