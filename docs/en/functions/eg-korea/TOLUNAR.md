# TOLUNAR

**Category**: EGTools-only function

Converts a solar (Gregorian) date to the Korean lunar date (919–2050).

## Syntax

```
=TOLUNAR(solar_date)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| solar_date | Required | date, serial, "yyyy-mm-dd" text, yyyymmdd number, or a range |

## Returns

Returns the lunar date as "yyyy-mm-dd" text (with a suffix for leap months) as a scalar, spilling an array of the same shape for range input. An unparseable date or one outside the lunar table range (918-2050) returns a #VALUE! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TOLUNAR(DATE(2025,1,29))` | 2025-01-01 | Lunar New Year 2025 = lunar 1/1 |
| `=TOLUNAR(DATE(2025,10,6))` | 2025-08-15 | Chuseok 2025 = lunar 8/15 |

## Notes

- Leap months get a "(윤)" (Korean) or "(LM)" (English) suffix depending on the language setting.
- Uses a built-in table (918-2050) generated from .NET KoreanLunisolarCalendar.
- Related function: TOSOLAR
- Supported: Excel 2010+. Always registered as `TOLUNAR` on every Excel version.
