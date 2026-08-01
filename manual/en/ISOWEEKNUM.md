# ISOWEEKNUM

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns the ISO 8601 week number of the year for a date.

## Syntax

```
=ISOWEEKNUM(date)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| date | Required | the date |

## Returns

Returns the ISO 8601 week number (1-53) of the date as a number. An array argument is computed element-wise and spills as an array of the same shape on dynamic-array hosts. Returns #VALUE! if the date is non-numeric or negative, and #NUM! if the serial cannot be converted to a date.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ISOWEEKNUM(DATE(2026,1,1))` | 1 | 2026-01-01 is week 1 |
| `=ISOWEEKNUM(DATE(2023,1,1))` | 52 | Last week of prior ISO year |

## Notes

- Array input is computed element-wise and returned as an array of the same shape.
- Supported: Excel 2010+. Registered as `ISOWEEKNUM` (drop-in) on hosts without the native function, and as `EG.ISOWEEKNUM` on modern Excel that has it.
