# IMPORTTEXT

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365 (new in 2026)

Imports data from a text-based file (local path or URL) as an array.

## Syntax

```
=IMPORTTEXT(path, [delimiter], [skip_rows], [take_rows], [encoding], [locale])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| path | Required | local file path or URL |
| delimiter | Optional | delimiter text (default tab), or ascending numbers like {1,3} for fixed-width columns |
| skip_rows | Optional | rows to skip; negative skips from the end |
| take_rows | Optional | rows to take; negative takes from the end |
| encoding | Optional | file encoding, e.g. "utf-8" (default), "windows-1252", "euc-kr" |
| locale | Optional | locale for number parsing, e.g. "de-DE" |

## Returns

Returns a text file (local path or URL) split by delimiter or fixed widths as a spilled 2-D array; numeric and date text is auto-converted. A missing path, read/download failure, or invalid delimiter/fixed-width spec gives #VALUE!; if no rows remain after skip/take the result is #N/A.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTTEXT("C:\data\log.txt",";",1)` |  | result depends on the file |

## Notes

- The default delimiter is tab. String value(s) are delimiters; a numeric array ({1,11,21}) is interpreted as 1-based ascending fixed-width column start positions.
- encoding accepts charset names like "euc-kr" or a codepage number; a BOM takes precedence.
- Negative skip_rows/take_rows count from the end.
- Supported: Excel 2010+. Registered as `IMPORTTEXT` (drop-in) on hosts without the native function, and as `EG.IMPORTTEXT` on modern Excel that has it.
