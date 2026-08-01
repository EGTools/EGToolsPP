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

Returns the expanded BOM tree as a 2-D spilled array with level/parent/child/value columns (plus WBS when requested). Returns #VALUE! when the three lists are empty or have different lengths, and the text "Error: ..." when the same parent-child pair appears with different values.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BOMTREE({"X";"X";"Y"},{"Y";"Z";"W"},{1;2;3})` | {0,"Top Level","X","";1,"X","Y",1;2,"Y","W",3;1,"X","Z",2} | Expand the BOM tree |
| `=BOMTREE({"X";"X";"Y"},{"Y";"Z";"W"},{1;2;3},TRUE)` | {"level","Parent","Child","Value";0,"Top Level","X","";1,"X","Y",1;2,"Y","W",3;1,"X","Z",2} | With header row |

## Notes

- WBS numbers are padded to 5 digits with zero-width spaces (U+200B) so they sort correctly.
- Expansion depth is limited to 64 levels (guard against circular references).
- Reading stops at the first row whose parent cell is empty.
- Supported: Excel 2010+. Always registered as `BOMTREE` on every Excel version.
