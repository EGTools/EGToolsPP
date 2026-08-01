# SUMBYBOM

**Category**: EGTools-only function

Expands one BOM level and sums required quantities per item.

## Syntax

```
=SUMBYBOM(array1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array1 | Required | columns in order: parent, child, usage, item name, quantity column(s) — ranges may be split across arguments |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SUMBYBOM(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `SUMBYBOM` on every Excel version.
