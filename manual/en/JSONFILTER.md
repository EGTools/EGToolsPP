# JSONFILTER

**Category**: EGTools-only function

Parses JSON text into a two-column (path, value) table.

## Syntax

```
=JSONFILTER(json_text, [key_path])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| json_text | Required | JSON text to parse |
| key_path | Optional | path to filter, e.g. "data.items[0].name" (default: whole document) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=JSONFILTER(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `JSONFILTER` on every Excel version (no `EG.`/`x` prefix).
