# STREXT

**Category**: EGTools-only function

Extracts or removes characters by type (digits, Latin, Korean, Japanese, Hanja) or by regex.

## Syntax

```
=STREXT(text, option, [delimiter])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| option | Required | S=space A=Latin N=digits K=Korean J=Japanese H=Hanja; leading "-" removes instead; anything else is a regex |
| delimiter | Optional | joiner between extracted runs (default none) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=STREXT(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `STREXT` on every Excel version.
