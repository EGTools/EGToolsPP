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
| api_key | Optional | data.go.kr key for temporary holidays. Entered once, it is stored and can be omitted later. Pass 0 to skip the API and use the built-in calculation only |

## Returns

Spills a two-column array of [date serial, holiday name] rows (dates only when holiday_name=FALSE). A year outside 919-2050 returns a #VALUE! error.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ROWS(KOREANHOLIDAYS(2025,,,0))` | 18 | Holiday rows for 2025 — the trailing 0 skips the API and uses the built-in calculation only |
| `=INDEX(KOREANHOLIDAYS(2025,FALSE,,0),1,1)` | 45658 | First holiday (Jan 1) serial, dates only (names excluded) |
| `=KOREANHOLIDAYS(2026)` |  | Default behaviour — the stored data.go.kr key merges temporary holidays (such as the nationwide local election on 2026-06-03). The row count depends on whether a key is registered and on network access |

## Notes

- Substitute-holiday rules are applied per their effective years (Children's Day 2014+, Mar 1/Liberation/Foundation Day 2022+, Buddha's Birthday and Christmas 2023+).
- Holiday names follow the add-in language setting (Korean/English); holidays falling on the same day are merged with commas.
- Lunar holidays (Lunar New Year, Buddha's Birthday, Chuseok) use the built-in lunar table (918-2050).
- The api_key argument has three uses. **Omit it** (the default) and the stored data.go.kr key is used to query the holiday API, merging in items the built-in calculation lacks — temporary holidays, election days; without a key, or when the lookup fails, only the built-in result is returned (the formula does not break). **Pass a key string** and it is stored and used; if the lookup then fails the cell shows `ERROR: 공휴일 API - <reason>` with the cause (no key / key rejected plus the server message / connection failure / no items in the response). **Pass 0** and the API is never called — the built-in calculation only, with no network access.
- The data.go.kr (Korean open-data portal) key is a single per-account key shared by KOREANHOLIDAYS, the ribbon Calendar commands and BRNSTATUS. Pass it once as the last argument, or register it under ribbon [EGTools] → [Manage API Keys]; it is stored in your user account and can be omitted afterwards. A key registered for business-registration lookups by an older version is moved to the shared slot automatically on first use.
- Lookup results are cached per year, so a recalculation queries a given year only once. Failures are not cached, so fixing the key takes effect on the next calculation.
- Supported: Excel 2010+. Always registered as `KOREANHOLIDAYS` on every Excel version.
