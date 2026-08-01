# UNIQUE

**Category**: MS Excel compatibility function · **Native since**: Excel 2021

Returns the unique rows (or columns) of an array.

## Syntax

```
=UNIQUE(array, [by_col], [exactly_once])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | data |
| by_col | Optional | TRUE for unique columns |
| exactly_once | Optional | TRUE to keep items appearing once |

## Returns

Returns an array of the unique rows (or columns); on dynamic-array hosts it spills. Returns #N/A when nothing remains and #VALUE! on internal errors.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=UNIQUE({1;2;2;3})` | {1;2;3} | Remove duplicates |
| `=UNIQUE({1;2;2;3},,TRUE)` | {1;3} | Items appearing exactly once |

## Notes

- With an empty result, #N/A is returned instead of #CALC!.
- Related functions: FILTER, SORT.
- Supported: Excel 2010+. Registered as `UNIQUE` (drop-in) on hosts without the native function, and as `EG.UNIQUE` on modern Excel that has it.
