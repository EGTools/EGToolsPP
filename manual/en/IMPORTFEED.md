# IMPORTFEED

**Category**: Google Sheets compatibility function

Imports an RSS or ATOM feed.

## Syntax

```
=IMPORTFEED(url, [query], [headers], [num_items])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| url | Required | feed URL |
| query | Optional | "items" (default), "feed", or a field: title/summary/description/author/url/created |
| headers | Optional | TRUE to include a header row |
| num_items | Optional | number of items to return (default: all) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTFEED(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `IMPORTFEED` on every Excel version (no `EG.`/`x` prefix).
