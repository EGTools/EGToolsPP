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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTRANGE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `IMPORTRANGE` on every Excel version.
