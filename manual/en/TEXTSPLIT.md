# TEXTSPLIT

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Splits text into an array using column and row delimiters.

## Syntax

```
=TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | text to split |
| col_delimiter | Required | delimiter between columns |
| row_delimiter | Optional | delimiter between rows (optional) |
| ignore_empty | Optional | TRUE to skip empty values |
| match_mode | Optional | 0 case-sensitive (default), 1 insensitive |
| pad_with | Optional | value for ragged rows (default #N/A) |

## Returns

Returns an array of the split text, spilling on dynamic-array hosts. Ragged rows are padded with pad_with (default #N/A); returns #VALUE! when the split result is empty.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTSPLIT("a,b,c",",")` | {a,b,c} | split into columns |
| `=TEXTSPLIT("a,b;c",",",";")` | {a,b;c,#N/A} | rows and columns, padded |
| `=TEXTSPLIT("1,,2",",",,TRUE)` | {1,2} | ignore empty values |

## Notes

- Each delimiter argument accepts a single text only (the native array-of-delimiters form is not supported).
- match_mode=1 ignores case.
- Related functions: TEXTBEFORE, TEXTAFTER, TEXTJOIN
- Supported: Excel 2010+. Registered as `TEXTSPLIT` (drop-in) on hosts without the native function, and as `EG.TEXTSPLIT` on modern Excel that has it.
