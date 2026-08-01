# REGEXTEST

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Tests whether text matches a regular expression.

## Syntax

```
=REGEXTEST(text, pattern, [case_sensitivity])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text to test |
| pattern | Required | regex pattern (ECMAScript) |
| case_sensitivity | Optional | 0 case-sensitive (default), 1 case-insensitive |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=REGEXTEST(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `REGEXTEST` (drop-in) on hosts without the native function, and as `EG.REGEXTEST` on modern Excel that has it.
