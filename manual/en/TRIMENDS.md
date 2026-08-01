# TRIMENDS

**Category**: EGTools-only function

Removes leading/trailing whitespace only (keeps inner spaces).

## Syntax

```
=TRIMENDS(text, [options])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text or range |
| options | Optional | extra chars to trim: +1 control chars, +2 nbsp, +4 invisible Unicode (default 0) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TRIMENDS(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `TRIMENDS` on every Excel version.
