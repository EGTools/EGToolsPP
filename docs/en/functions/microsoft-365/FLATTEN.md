# FLATTEN

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365 (new in 2026)

Removes nesting levels from list/array text and spills the items into the grid.

## Syntax

```
=FLATTEN(array, [pad_value], [levels])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | list text, range or array to flatten |
| pad_value | Optional | value for ragged rows (default #N/A) |
| levels | Optional | nesting levels to remove (default all) |

## Returns

Returns the un-nested 2-D array, spilling on dynamic-array hosts. If array itself is an error, that error is returned as is; an array in pad_value, a non-numeric or negative levels, or an empty result returns #VALUE!.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=FLATTEN("a, b, c")` | {a;b;c} | list → one column |
| `=FLATTEN("{""a"",""b"";""c""}")` | {a,b;c,#N/A} | 2-D array text, shortfall padded with #N/A |
| `=FLATTEN({"a, b";"c"})` | {a;b;c} | a column of list cells stacks vertically |

## Notes

- One level = in each row of the outer array, the elements (scalar = 1×1, array = its own shape) are joined side by side and the rows are stacked, with widths equalised using pad_value (default #N/A). With levels omitted or 0 this repeats until no nesting remains; a positive N removes N levels and any remaining nested item is shown as list text ({…}). Example: =FLATTEN("{""a"",{""b"",""c""}}",,1) → {a,{"b","c"}} — only the first nesting of that cell is removed.
- A one-column list gives one vertical column, a column range of several list cells stacks vertically, and a single cell holding 2-D array text keeps that 2-D shape. Ragged rows inside the array text itself are always padded with #N/A regardless of pad_value; error items in the list (#N/A etc.) are preserved as data. A single text without commas or braces becomes a 1×1 text.
- Lists in Cells from the Microsoft 365 beta show only as #UNKNOWN! (#VALUE! in Excel 2016) in older Excel, so convert them first with ribbon [EGTools] › [List/Array Cells to Text]. Related: HAS, HASANY, HASALL
- Supported: Excel 2010+. Registered as `FLATTEN` (drop-in) on hosts without the native function, and as `EG.FLATTEN` on modern Excel that has it.
