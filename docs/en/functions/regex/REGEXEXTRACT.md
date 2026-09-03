# REGEXEXTRACT

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365

Extracts matches of a regular expression from text.

## Syntax

```
=REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| pattern | Required | regex pattern |
| return_mode | Optional | 0 first match, 1 all matches, 2 capture groups |
| case_sensitivity | Optional | 0 case-sensitive (default), 1 case-insensitive |

## Returns

return_mode 0 returns the first match (scalar text), 1 returns all matches (one-column array, spills), 2 returns the capture groups of the first match (one-row array, spills). Returns #N/A when there is no match (or no groups in mode 2) and #VALUE! for an invalid pattern.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=REGEXEXTRACT("abc123def45","\d+")` | 123 | first match |
| `=REGEXEXTRACT("abc123def45","\d+",1)` | {123;45} | all matches (column) |
| `=REGEXEXTRACT("2026-08-01","(\d+)-(\d+)",2)` | {2026,08} | capture groups (row) |

## Notes

- The regex flavor is std::wregex ECMAScript (may differ from native 365's PCRE2 in places).
- text and pattern accept arrays and are computed element-wise, spilling an array of the same shape (an error element returns that error); with an array text and return_mode 1 or 2, each element yields only its first match (demotion).
- Related functions: REGEXTEST, REGEXREPLACE
- Supported: Excel 2010+. Registered as `REGEXEXTRACT` (drop-in) on hosts without the native function, and as `EG.REGEXEXTRACT` on modern Excel that has it.
