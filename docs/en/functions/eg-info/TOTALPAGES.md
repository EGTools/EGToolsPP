# TOTALPAGES

**Category**: EGTools-only function

Total number of pages with the current print settings (recalculate before printing).

## Syntax

```
=TOTALPAGES([ignore_start])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| ignore_start | Optional | TRUE numbers from 1, ignoring the first-page-number setting |

## Returns

Returns the total page count as a numeric scalar, reflecting the first-page-number print setting (counts from 1 when ignore_start is TRUE). Returns #VALUE! when the page count query fails.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TOTALPAGES()` |  | Result depends on the print settings |

## Notes

- Macro-type function; uses XLM GET.DOCUMENT (not thread-safe).
- Recalculate after changing print settings to refresh the value.
- Supported: Excel 2010+. Always registered as `TOTALPAGES` on every Excel version.
