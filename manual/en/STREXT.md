# STREXT

**Category**: EGTools-only function

Extracts or removes characters by type (digits, Latin, Korean, Japanese, Hanja) or by regex.

## Syntax

```
=STREXT(text, option, [delimiter])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| option | Required | S=space A=Latin N=digits K=Korean J=Japanese H=Hanja; leading "-" removes instead; anything else is a regex |
| delimiter | Optional | joiner between extracted runs (default none) |

## Returns

Returns the extracted (or removed) result as a single text (scalar). Returns #VALUE! when text/option is missing or option is empty, no valid character class results, or the regex is invalid.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=STREXT("abc123한글45","N")` | 12345 | extract digits |
| `=STREXT("abc123한글45","K")` | 한글 | extract Korean |
| `=STREXT("a1b2","N","-")` | 1-2 | join with delimiter |

## Notes

- The N class also includes the decimal point '.'.
- If the option is not composed of -SANKJH letters, the whole option is treated as an ECMAScript regex (std::wregex).
- A leading '-' removes the matched characters instead of extracting them.
- If the text is an array, it is processed element-wise and spills an array of the same shape (an error element returns that error); an array in the option and separator arguments returns #VALUE!.
- Supported: Excel 2010+. Always registered as `STREXT` on every Excel version.
