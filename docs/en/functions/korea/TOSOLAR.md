# TOSOLAR

**Category**: EGTools-only function

Converts a Korean lunar date to the solar (Gregorian) date (919–2050).

## Syntax

```
=TOSOLAR(lunar_date)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| lunar_date | Required | lunar date text like "2025-06-05" — add 윤/LM for a leap month — or a range |

## Returns

Returns the solar (Gregorian) date serial value as a scalar number, spilling an array of the same shape for range input. Unparseable text, a date outside the lunar table range (918-2050), or a nonexistent lunar date returns a #VALUE! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TOSOLAR("2025-01-01")` | 45686 | Lunar New Year = 2025-01-29 |
| `=TOSOLAR("2025-08-15")` | 45936 | Chuseok = 2025-10-06 |

## Notes

- Mark leap months with "윤" or "LM" (e.g. "2025-윤06-01").
- Apply a date number format to the result cell.
- Related function: TOLUNAR
- Supported: Excel 2010+. Always registered as `TOSOLAR` on every Excel version.
