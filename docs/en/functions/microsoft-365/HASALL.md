# HASALL

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365 (new in 2026)

Returns TRUE if all of the values appear in a list or array.

## Syntax

```
=HASALL(array, values)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | list text, range or array to search |
| values | Required | values to look for (range, array or list text) |

## Returns

Returns TRUE/FALSE (scalar). If array or values itself is an error, that error is returned as is; omitting values returns #VALUE!. Empty values gives TRUE.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=HASALL("a, b, c",{"a","c"})` | TRUE | TRUE when all values are present |
| `=HASALL("a, b",{"a","z"})` | FALSE | FALSE when any value is missing |

## Notes

- values accepts a range, an array constant or list text; all items including nested ones are flattened into one set, and each is looked up with the same rules as HAS (whole-item case-insensitive match, numeric equality). Text given as values is parsed as list text — split at commas, with numeric-looking parts read as numbers ("z, B" specifies several values; "1, 2" is the numbers 1 and 2). To compare a single value as text, pass it as the array constant {"1"}.
- A blank cell in a values range counts as an empty-text item, so the result is FALSE unless the list contains an empty item. array is interpreted exactly as in HAS (list text, range or array constant; a single text without commas stays text; unparsable text is one item). Convert Lists in Cells first with ribbon [EGTools] › [List/Array Cells to Text].
- Related: HAS, HASANY, FLATTEN
- Supported: Excel 2010+. Registered as `HASALL` (drop-in) on hosts without the native function, and as `EG.HASALL` on modern Excel that has it.
