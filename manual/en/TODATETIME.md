# TODATETIME

**Category**: EGTools-only function

Parses date/time text containing Korean or Hanja units (년월일시분초) into a date-time value.

## Syntax

```
=TODATETIME(datetime_text)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| datetime_text | Required | text containing a date and/or time |

## Returns

Returns a date-time serial value (scalar number): an integer for date only, a fraction below 1 for time only. Returns empty text when no date or time is found, and #VALUE! on a parsing failure.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TODATETIME("2026년 1월 15일")` | 46037 | Korean date text to serial |
| `=TODATETIME("오후 2시 30분")` | 0.604166667 | Afternoon time to time value |

## Notes

- Hanja units (年月日時分秒) and AM/PM markers (오전/오후, AM/PM, 上午/下午) are recognized.
- Month-day only text uses the current year; two-digit years above 31 are mapped to the 2000s.
- Apply a date/time number format to the result cell.
- Supported: Excel 2010+. Always registered as `TODATETIME` on every Excel version.
