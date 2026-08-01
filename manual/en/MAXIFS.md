# MAXIFS

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Returns the maximum of cells that meet all criteria.

## Syntax

```
=MAXIFS(max_range, criteria_range1, criteria1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| max_range | Required | cells to maximise |
| criteria_range1 | Required | range to test |
| criteria1 | Required | criterion, e.g. ">5" |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns the maximum numeric value among cells meeting all criteria, as a numeric scalar. Returns #VALUE! when a criteria range's size differs from the value range, and 0 when nothing matches.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=MAXIFS({1;5;3},{1;2;3},">1")` | 5 | Max meeting criteria |
| `=MAXIFS({10;20},{1;2},">5")` | 0 | Returns 0 when nothing matches |

## Notes

- Criteria support numbers, comparison operators (e.g. ">5") and wildcards (*, ?); text comparison is case-insensitive.
- Related functions: MINIFS.
- Supported: Excel 2010+. Registered as `MAXIFS` (drop-in) on hosts without the native function, and as `EG.MAXIFS` on modern Excel that has it.
