# TEXTREPLACE

**Category**: EGTools-only function

Replaces every StartKey…EndKey section of the text.

## Syntax

```
=TEXTREPLACE(text, start_key, end_key, [replace], [include])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| start_key | Required | section start marker |
| end_key | Required | section end marker |
| replace | Optional | replacement text (default empty) |
| include | Optional | TRUE keeps the markers around the replacement |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTREPLACE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `TEXTREPLACE` on every Excel version (no `EG.`/`x` prefix).
