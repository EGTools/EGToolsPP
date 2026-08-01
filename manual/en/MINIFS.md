# MINIFS

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Returns the minimum of cells that meet all criteria.

## Syntax

```
=MINIFS(min_range, criteria_range1, criteria1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| min_range | Required | cells to minimise |
| criteria_range1 | Required | range to test |
| criteria1 | Required | criterion, e.g. "<5" |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns the minimum numeric value among cells meeting all criteria, as a numeric scalar. Returns #VALUE! when a criteria range's size differs from the value range, and 0 when nothing matches.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=MINIFS({10;20;30},{"a";"b";"a"},"a")` | 10 | Min meeting criteria |
| `=MINIFS({7;8;9},{"x";"y";"z"},"?")` | 7 | Wildcard criterion |

## Notes

- Criteria support numbers, comparison operators (e.g. "<5") and wildcards (*, ?); text comparison is case-insensitive.
- Related functions: MAXIFS.
- Supported: Excel 2010+. Registered as `MINIFS` (drop-in) on hosts without the native function, and as `EG.MINIFS` on modern Excel that has it.
