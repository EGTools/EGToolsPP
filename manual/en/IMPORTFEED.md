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

Returns feed items (or feed info / a single field) as a spilled 2-D array. A missing URL or unsupported query field gives #VALUE!; a download failure, XML parse failure, or no items gives #N/A.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTFEED("https://blog.example.com/rss","items",TRUE,5)` |  | result depends on the feed |

## Notes

- RSS 2.0 and Atom are auto-detected. query "items" (default) returns five columns: Title/Summary/URL/Author/Created.
- RFC822/ISO8601 dates in the created field are converted to Excel serial values (time-zone offsets ignored).
- Uses COM (MSXML), so it is excluded from multi-threaded recalculation; requires an internet connection.
- Supported: Excel 2010+. Always registered as `IMPORTFEED` on every Excel version.
