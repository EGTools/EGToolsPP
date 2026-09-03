# SHEETS

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the number of sheets in the workbook (or in a reference).

## Syntax

```
=SHEETS([reference])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| reference | Optional | reference (optional) |

## Returns

Returns the number of sheets in the workbook as a number, or always 1 when a reference is given. Returns #VALUE! when the sheet-count lookup fails.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=SHEETS()` |  | Depends on the workbook |

## Notes

- With a reference, always returns 1 (counting sheets of a 3-D reference is not supported).
- Registered as a macro-sheet function.
- Supported: Excel 2010+. Registered as `SHEETS` (drop-in) on hosts without the native function, and as `EG.SHEETS` on modern Excel that has it.
