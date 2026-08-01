# TEXTBETWEEN

**Category**: EGTools-only function

Extracts the text between StartKey and EndKey.

## Syntax

```
=TEXTBETWEEN(text, left_text, right_text, [delimiter_or_instance], [include])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| left_text | Required | start marker |
| right_text | Required | end marker |
| delimiter_or_instance | Optional | joiner for all matches (default ","), or a number N for the Nth match only |
| include | Optional | TRUE includes the markers |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTBETWEEN(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `TEXTBETWEEN` on every Excel version (no `EG.`/`x` prefix).
