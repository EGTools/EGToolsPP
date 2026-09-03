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

Returns a spilled array with one item-name column plus one summed column per quantity column, sorted by item name ascending. Returns #VALUE! when fewer than 5 columns are given, column lengths differ, or usage/quantity is not numeric; #NUM! for negative quantities; #N/A when nothing is aggregated.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SUMBYBOM({"X","P",2,"X",1;"X","Q",3,"P",5})` | {"P",7;"Q",3} | Expand X one level and sum |

## Notes

- Only one BOM level is expanded (not multi-level).
- Item names are aggregated case-insensitively; the result is sorted by name ascending.
- Non-parent items with zero quantity are dropped from the result.
- Supported: Excel 2010+. Always registered as `SUMBYBOM` on every Excel version.
