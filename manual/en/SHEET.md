# SHEET

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the sheet number of a reference (or the current sheet).

## Syntax

```
=SHEET([value])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| value | Optional | reference (optional; default: current sheet) |

## Returns

Returns the 1-based position of the sheet containing the reference (or the calling cell when omitted) as a number. Returns #N/A when the sheet cannot be found and #VALUE! when the sheet list lookup fails.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SHEET()` |  | Depends on workbook and calling sheet |

## Notes

- Accepts only a cell reference (native also accepts a sheet name as text).
- Registered as a macro-sheet function.
- Supported: Excel 2010+. Registered as `SHEET` (drop-in) on hosts without the native function, and as `EG.SHEET` on modern Excel that has it.
