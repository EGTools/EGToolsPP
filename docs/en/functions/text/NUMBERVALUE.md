# NUMBERVALUE

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Converts text to a number in a locale-independent way.

## Syntax

```
=NUMBERVALUE(text, [decimal_separator], [group_separator])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text to convert |
| decimal_separator | Optional | decimal separator (default .) |
| group_separator | Optional | group separator (default ,) |

## Returns

Returns a number (scalar); if text is an array it is evaluated element-wise and spills on dynamic-array hosts. Returns #VALUE! when the text does not parse as a number after separator handling; empty (whitespace-only) text returns 0.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=NUMBERVALUE("2.500,27",",",".")` | 2500.27 | European-style notation |
| `=NUMBERVALUE("9%%")` | 0.0009 | each % divides by 100 |
| `=NUMBERVALUE("abc")` | #VALUE! | not a number |

## Notes

- All whitespace is ignored anywhere in the text, and group separators are removed wherever they appear.
- Trailing % signs divide the result by 100 once per sign.
- Supported: Excel 2010+. Registered as `NUMBERVALUE` (drop-in) on hosts without the native function, and as `EG.NUMBERVALUE` on modern Excel that has it.
