# HAS

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365 (new in 2026)

Returns TRUE if the value appears anywhere in a list or array (list text, range or array constant).

## Syntax

```
=HAS(array, value)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | list text, range or array to search |
| value | Required | value to look for (whole item, case-insensitive) |

## Returns

Returns TRUE/FALSE (scalar); an array in value is evaluated element-wise and spills. If the array argument itself is an error, that error is returned as is; omitting value returns #VALUE!.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=HAS("Mon, Tue, Wed","tue")` | TRUE | whole-item match, case-insensitive |
| `=HAS("Ann, Bob","An")` | FALSE | no partial match |
| `=HAS("{""a"",{""b"",""c""}}","c")` | TRUE | nested array text |
| `=HAS("1, 2","1")` | FALSE | text "1" ≠ number 1 |

## Notes

- array accepts list text (Mon, Tue, Wed, or array-constant form such as {"a",{"b","c"}}), a range or an array constant, and is searched recursively including nested levels. Items of a comma-separated list or inside {…} are read as numbers when they look numeric, as logicals for TRUE/FALSE and as errors for error names such as #N/A, but a single text without commas or braces is not classified and stays one text item (a plain cell holding "1" is the text "1"). Text that cannot be parsed (e.g. an unclosed quote) is treated as a single item without raising an error.
- Comparison is whole-item and case-insensitive for text ("Ann"≠"Anna") and numeric for numbers; text "1" and number 1 differ. Error items inside the list never match. array is not lifted — a range is taken as one whole array, so for per-row results use a table's [@column] reference or BYROW.
- Lists in Cells from the Microsoft 365 beta show only as #UNKNOWN! (#VALUE! in Excel 2016) in older Excel, so convert them first with ribbon [EGTools] › [List/Array Cells to Text]. Related: HASANY, HASALL, FLATTEN
- Supported: Excel 2010+. Registered as `HAS` (drop-in) on hosts without the native function, and as `EG.HAS` on modern Excel that has it.
