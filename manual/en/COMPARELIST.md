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

Returns the items of the chosen set as a spilled array (multiple columns when by_row is TRUE, otherwise one column). Returns #VALUE! when no list is given or unique_type is outside -3..number-of-lists, and #N/A when the result set is empty.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=COMPARELIST(0,FALSE,{"a";"b"},{"b";"c"})` | {"a";"b";"c"} | Union without duplicates |
| `=COMPARELIST(-2,FALSE,{"a";"b"},{"b";"c"})` | b | Items in every list |
| `=COMPARELIST(1,FALSE,{"a";"b"},{"b";"c"})` | a | Items only in list 1 |

## Notes

- Comparison is case-insensitive; numbers are compared by value.
- The result keeps the order in which items first appear.
- With by_row TRUE, whole rows are compared as one key (trailing empty cells ignored).
- Supported: Excel 2010+. Always registered as `COMPARELIST` on every Excel version.
