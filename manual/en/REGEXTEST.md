# REGEXTEST

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365

Tests whether text matches a regular expression.

## Syntax

```
=REGEXTEST(text, pattern, [case_sensitivity])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text to test |
| pattern | Required | regex pattern (ECMAScript) |
| case_sensitivity | Optional | 0 case-sensitive (default), 1 case-insensitive |

## Returns

Returns a logical TRUE/FALSE (scalar) indicating whether a match exists. Returns #VALUE! for an invalid pattern.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=REGEXTEST("abc123","\d+")` | TRUE | contains digits |
| `=REGEXTEST("ABC","[a-z]+",1)` | TRUE | case-insensitive |
| `=REGEXTEST("abc","^\d")` | FALSE | no match |

## Notes

- The regex flavor is std::wregex ECMAScript (may differ from native 365's PCRE2 in places).
- Related functions: REGEXEXTRACT, REGEXREPLACE
- Supported: Excel 2010+. Registered as `REGEXTEST` (drop-in) on hosts without the native function, and as `EG.REGEXTEST` on modern Excel that has it.
