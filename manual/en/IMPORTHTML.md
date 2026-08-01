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

Returns a web-page table (2-D array) or list (N-by-1 array) as a spilled array; numeric text is auto-converted. A missing URL, a search value other than "table"/"list", or index below 1 gives #VALUE!; a download failure or no table/list at that index gives #N/A.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTHTML("https://en.wikipedia.org/wiki/Microsoft_Excel","table",1)` |  | result depends on the page |

## Notes

- The page is parsed with MSHTML after removing <script> blocks, so script-generated dynamic content cannot be imported.
- Table rowspan/colspan are normalized onto a grid; the value goes into the top-left cell only.
- Uses COM, so it is excluded from multi-threaded recalculation; requires an internet connection.
- Supported: Excel 2010+. Always registered as `IMPORTHTML` on every Excel version.
