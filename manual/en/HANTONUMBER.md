# HANTONUMBER

**Category**: EGTools-only function

Converts Korean/Hanja money notation (e.g. 일금 오만원) to a number.

## Syntax

```
=HANTONUMBER(text, [number_only])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | Korean money text |
| number_only | Optional | TRUE returns a number (default); FALSE returns formatted text with 금/원 |

## Returns

By default (number_only=TRUE) returns a number (scalar); when the integer part exceeds 15 digits the result is returned as text to preserve precision. With number_only=FALSE returns formatted text like "금1,234원". Returns #VALUE! when the text is empty or contains two or more consecutive Arabic digits/dots.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=HANTONUMBER("일금오만원정")` | 50000 | Korean money to number |
| `=HANTONUMBER("삼억이천오백만")` | 325000000 | eok/man units |
| `=HANTONUMBER("일금오만원정",FALSE)` | 일금50,000원정 | formatted text output |

## Notes

- Hanja notation and variant forms (一二三, 壹貳參, 拾佰仟萬, etc.) are recognized.
- Spaces and commas in the input are ignored; the 금/일금 and 원/원정 prefix/suffix are restored when number_only=FALSE.
- If the text is an array, it is processed element-wise and spills an array of the same shape (an error element returns that error); an array in the other arguments returns #VALUE!.
- Supported: Excel 2010+. Always registered as `HANTONUMBER` on every Excel version.
