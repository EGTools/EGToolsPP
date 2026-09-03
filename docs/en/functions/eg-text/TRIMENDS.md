# TRIMENDS

**Category**: EGTools-only function

Removes leading/trailing whitespace only (keeps inner spaces).

## Syntax

```
=TRIMENDS(text, [options])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text or range |
| options | Optional | extra chars to trim: +1 control chars, +2 nbsp, +4 invisible Unicode (default 0) |

## Returns

Returns the text with both ends trimmed (scalar); an array/range input is processed element-wise and spills. Returns #VALUE! when options is outside 0..7; empty cells become an empty string.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TRIMENDS("  a  b  ")` | a  b | trim ends only |
| `=TRIMENDS("a",8)` | #VALUE! | options out of range |

## Notes

- Even with options 0, spaces, tabs, CR and LF are always trimmed from both ends.
- Unlike TRIM, inner spaces are kept as-is.
- Options are additive: +1 control chars, +2 NBSP, +4 invisible Unicode spaces (ZWSP, full-width space, etc.).
- Supported: Excel 2010+. Always registered as `TRIMENDS` on every Excel version.
