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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTSPLIT(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `TEXTSPLIT` (drop-in) on hosts without the native function, and as `EG.TEXTSPLIT` on modern Excel that has it.
