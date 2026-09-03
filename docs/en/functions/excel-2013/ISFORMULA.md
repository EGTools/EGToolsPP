# ISFORMULA

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns TRUE if the referenced cell contains a formula.

## Syntax

```
=ISFORMULA(reference)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| reference | Required | cell to test |

## Returns

Returns TRUE when the referenced cell contains a formula, otherwise FALSE. Returns #VALUE! when the reference is omitted or the lookup fails.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ISFORMULA(A1)` |  | Depends on cell A1's content |

## Notes

- Registered as a macro-sheet function.
- Related functions: FORMULATEXT.
- Supported: Excel 2010+. Registered as `ISFORMULA` (drop-in) on hosts without the native function, and as `EG.ISFORMULA` on modern Excel that has it.
