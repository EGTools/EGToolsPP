# KOREANHOLIDAYS

**Category**: EGTools-only function

Returns the list of Korean public holidays for a year (substitute-holiday rules included).

## Syntax

```
=KOREANHOLIDAYS([calendar_year], [holiday_name], [may_day], [api_key])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| calendar_year | Optional | year (default: this year) |
| holiday_name | Optional | TRUE includes names (default), FALSE dates only |
| may_day | Optional | TRUE includes Labor Day (default) |
| api_key | Optional | data.go.kr key for temporary holidays; entered once, then stored in the registry |

## Returns

Spills a two-column array of [date serial, holiday name] rows (dates only when holiday_name=FALSE). A year outside 919-2050 returns a #VALUE! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ROWS(KOREANHOLIDAYS(2025))` | 18 | Holiday rows for 2025 (built-in calculation — with an api_key registered, temporary holidays may add rows) |
| `=INDEX(KOREANHOLIDAYS(2025,FALSE),1,1)` | 45658 | First holiday (Jan 1) serial |

## Notes

- Substitute-holiday rules are applied per their effective years (Children's Day 2014+, Mar 1/Liberation/Foundation Day 2022+, Buddha's Birthday and Christmas 2023+).
- Holiday names follow the add-in language setting (Korean/English); holidays falling on the same day are merged with commas.
- Lunar holidays (Lunar New Year, Buddha's Birthday, Chuseok) use the built-in lunar table (918-2050).
- With an api_key (Korean open-data portal, data.go.kr) the holiday API is queried and items missing from the built-in calculation — temporary holidays, election days — are merged in. A key passed once as the argument is stored in the registry (HKCU\Software\EGTools\ApiKeys) and can be omitted afterwards; it can also be registered via ribbon EGTools → Manage API Keys. Without a key, or when the lookup fails, only the built-in result is returned (cached per year).
- Supported: Excel 2010+. Always registered as `KOREANHOLIDAYS` on every Excel version.
