# MVLOOKUP

**Category**: EGTools-only function

Runs VLOOKUP for many keys at once, returning one or more columns.

## Syntax

```
=MVLOOKUP(lookup_value, lookup_array, return_column, [match_mode], [if_not_found])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| lookup_value | Required | key column (or single value) |
| lookup_array | Required | data table; keys in first column |
| return_column | Required | column number(s) to return, e.g. 2 or {2,3} |
| match_mode | Optional | FALSE exact (default), TRUE approximate (largest ≤ key) |
| if_not_found | Optional | value when no match (default "") |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=MVLOOKUP(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `MVLOOKUP` on every Excel version.
