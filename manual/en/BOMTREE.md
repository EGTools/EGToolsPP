# BOMTREE

**Category**: EGTools-only function

Builds a BOM tree from parent/child/quantity lists.

## Syntax

```
=BOMTREE(parents, children, values, [header], [wbs])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| parents | Required | parent item column |
| children | Required | child item column (same length) |
| values | Required | usage quantity column (same length) |
| header | Optional | TRUE adds a header row |
| wbs | Optional | TRUE adds a WBS column |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BOMTREE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `BOMTREE` on every Excel version (no `EG.`/`x` prefix).
