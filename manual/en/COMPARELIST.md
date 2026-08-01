# COMPARELIST

**Category**: EGTools-only function

Compares lists and returns a chosen set: 0 union, N only-in-list-N, -1 in-exactly-one, -2 in-all, -3 in-some-but-not-all.

## Syntax

```
=COMPARELIST(unique_type, by_row, array1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| unique_type | Required | 0 union, N only in list N, -1 exactly one, -2 all, -3 some-not-all |
| by_row | Required | TRUE compares whole rows, FALSE each cell |
| array1 | Required | list range or array |
| ... | Optional | the argument(s) above may be repeated |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=COMPARELIST(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `COMPARELIST` on every Excel version.
