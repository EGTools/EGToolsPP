# SHEETLIST

**Category**: EGTools-only function

Returns the visible sheet names of the calling workbook.

## Syntax

```
=SHEETLIST([start], [end], [strictly_between], [except_me])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| start | Optional | first sheet number or name (default: first) |
| end | Optional | last sheet number or name (default: last) |
| strictly_between | Optional | TRUE excludes the start/end sheets |
| except_me | Optional | TRUE excludes the calling sheet (default) |

## Returns

Returns the visible sheet names as a one-column spilled array. Returns #VALUE! when the caller/workbook cannot be resolved, COM is unavailable, or the start/end specification is invalid, and #N/A when no sheet remains after filtering.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SHEETLIST()` |  | Result depends on the workbook |

## Notes

- Macro-type function; reads sheet information through COM (not thread-safe).
- Hidden sheets are always excluded, and the calling sheet is excluded by default (except_me defaults to TRUE).
- start and end accept either a sheet number or a sheet name.
- Supported: Excel 2010+. Always registered as `SHEETLIST` on every Excel version.
