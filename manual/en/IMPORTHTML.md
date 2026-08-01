# IMPORTHTML

**Category**: Google Sheets compatibility function

Extracts a table or list from a web page.

## Syntax

```
=IMPORTHTML(url, [search], [index], [locale])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| url | Required | web page URL |
| search | Optional | "table" (default) or "list" (ul/ol) |
| index | Optional | which occurrence to extract, 1-based (default 1) |
| locale | Optional | locale for number parsing, e.g. "de-DE" |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTHTML(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `IMPORTHTML` on every Excel version.
