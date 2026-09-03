# FORMULATEXT

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the formula of the referenced cell as text.

## Syntax

```
=FORMULATEXT(reference)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| reference | Required | cell whose formula is returned |

## Returns

Returns the referenced cell's formula as a text scalar. Returns #N/A when the reference is omitted, the cell has no formula, or the lookup fails.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=FORMULATEXT(A1)` |  | Depends on cell A1's content |

## Notes

- Registered as a macro-sheet function.
- Related functions: ISFORMULA.
- Supported: Excel 2010+. Registered as `FORMULATEXT` (drop-in) on hosts without the native function, and as `EG.FORMULATEXT` on modern Excel that has it.
