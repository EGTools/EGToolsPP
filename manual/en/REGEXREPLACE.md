# REGEXREPLACE

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Replaces text matching a regular expression ($1.. backrefs).

## Syntax

```
=REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| pattern | Required | regex pattern |
| replacement | Required | replacement text ($1, $2…) |
| occurrence | Optional | 0 all (default), N = Nth match |
| case_sensitivity | Optional | 0 case-sensitive (default), 1 case-insensitive |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=REGEXREPLACE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `REGEXREPLACE` (drop-in) on hosts without the native function, and as `EG.REGEXREPLACE` on modern Excel that has it.
