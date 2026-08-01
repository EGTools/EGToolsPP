# IMPORTTEXT

**Category**: MS Excel compatibility function · **Native since**: Excel 365 (new in 2026)

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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTTEXT(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `IMPORTTEXT` (drop-in) on hosts without the native function, and as `EG.IMPORTTEXT` on modern Excel that has it.
