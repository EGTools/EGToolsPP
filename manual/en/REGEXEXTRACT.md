# REGEXEXTRACT

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Extracts matches of a regular expression from text.

## Syntax

```
=REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| pattern | Required | regex pattern |
| return_mode | Optional | 0 first match, 1 all matches, 2 capture groups |
| case_sensitivity | Optional | 0 case-sensitive (default), 1 case-insensitive |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=REGEXEXTRACT(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `REGEXEXTRACT` (drop-in) on hosts without the native function, and as `EG.REGEXEXTRACT` on modern Excel that has it.
