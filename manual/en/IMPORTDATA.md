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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTDATA(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `IMPORTDATA` on every Excel version.
