# HANTONUMBER

**Category**: EGTools-only function

Converts Korean/Hanja money notation (e.g. 일금 오만원) to a number.

## Syntax

```
=HANTONUMBER(text, [number_only])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | Korean money text |
| number_only | Optional | TRUE returns a number (default); FALSE returns formatted text with 금/원 |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=HANTONUMBER(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `HANTONUMBER` on every Excel version.
