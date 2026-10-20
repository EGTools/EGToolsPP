# HASANY

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365 (new in 2026)

Returns TRUE if any of the values appears anywhere in a list or array.

## Syntax

```
=HASANY(array, values)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | list text, range or array to search |
| values | Required | values to look for (range, array or list text) |

## Returns

Returns TRUE/FALSE (scalar). If array or values itself is an error, that error is returned as is; omitting values returns #VALUE!. Empty values gives FALSE.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=HASANY("a, b, c",{"x","c"})` | TRUE | TRUE if any value is present |
| `=HASANY("1, 2",{"1"})` | FALSE | text "1" in an array constant differs from number 1 |

## Notes

- values accepts a range, an array constant or list text; all items including nested ones are flattened into one set, and each is looked up with the same rules as HAS (whole-item case-insensitive match, numeric equality). Text given as values is parsed as list text — split at commas, with numeric-looking parts read as numbers ("z, B" specifies several values; "1, 2" is the numbers 1 and 2). To compare a single value as text, pass it as the array constant {"1"}.
- array is interpreted exactly as in HAS (list text, range or array constant; a single text without commas stays text; unparsable text is one item). Convert Lists in Cells first with ribbon [EGTools] › [List/Array Cells to Text].
- Related: HAS, HASALL, FLATTEN
- Supported: Excel 2010+. Registered as `HASANY` (drop-in) on hosts without the native function, and as `EG.HASANY` on modern Excel that has it.
