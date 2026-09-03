# REGEXREPLACE

**Category**: MS Excel compatibility function · **Native since**: Microsoft 365

Replaces text matching a regular expression ($1.. backrefs).

## Syntax

```
=REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| pattern | Required | regex pattern |
| replacement | Required | replacement text ($1, $2…) |
| occurrence | Optional | 0 all (default), N = Nth match |
| case_sensitivity | Optional | 0 case-sensitive (default), 1 case-insensitive |

## Returns

Returns the replaced text as a single scalar. Returns #VALUE! for an invalid pattern.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=REGEXREPLACE("a1b2c3","\d","-")` | a-b-c- | replace all |
| `=REGEXREPLACE("a1b2c3","\d","-",2)` | a1b-c3 | replace 2nd only |
| `=REGEXREPLACE("john smith","(\w+) (\w+)","$2 $1")` | smith john | backreference swap |

## Notes

- $1, $2 backreferences are supported in the replacement.
- occurrence 0 (default) replaces all matches; N>0 replaces only the Nth.
- The regex flavor is std::wregex ECMAScript.
- text·pattern·replacement accept arrays and are computed element-wise, spilling an array of the same shape — scalars are broadcast, a column vector × a row vector expands to their outer product, size-mismatched elements yield #N/A, and an error element returns that error.
- Supported: Excel 2010+. Registered as `REGEXREPLACE` (drop-in) on hosts without the native function, and as `EG.REGEXREPLACE` on modern Excel that has it.
