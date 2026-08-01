# IMPORTCSV

**Category**: MS Excel compatibility function · **Native since**: Excel 365 (new in 2026)

Imports a CSV file (local path or URL) as an array — comma delimiter, UTF-8.

## Syntax

```
=IMPORTCSV(path, [skip_rows], [take_rows], [locale])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| path | Required | local file path or URL |
| skip_rows | Optional | rows to skip; negative skips from the end |
| take_rows | Optional | rows to take; negative takes from the end |
| locale | Optional | locale for number parsing, e.g. "de-DE" |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMPORTCSV(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `IMPORTCSV` (drop-in) on hosts without the native function, and as `EG.IMPORTCSV` on modern Excel that has it.
