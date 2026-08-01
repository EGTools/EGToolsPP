# ARABIC

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Converts a Roman numeral to an Arabic number.

## Syntax

```
=ARABIC(text)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | Roman numeral text |

## Returns

Returns the converted number as a scalar. Any character other than the Roman symbols (I, V, X, L, C, D, M) and a leading '-' returns a #VALUE! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ARABIC("LVII")` | 57 | Roman numeral to number |
| `=ARABIC("MCMXII")` | 1912 | Convert a large Roman numeral |
| `=ARABIC("-X")` | -10 | Negative notation supported |

## Notes

- Whitespace is ignored and letters are case-insensitive.
- Empty text returns 0.
- Related functions: BASE, DECIMAL
- Supported: Excel 2010+. Registered as `ARABIC` (drop-in) on hosts without the native function, and as `EG.ARABIC` on modern Excel that has it.
