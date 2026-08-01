# TEXTAFTER

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Returns the text after a delimiter.

## Syntax

```
=TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| text | Required | source text |
| delimiter | Required | delimiter to search for |
| instance_num | Optional | which occurrence, negative from the end (default 1) |
| match_mode | Optional | 1 to ignore case (default 0) |
| match_end | Optional | 1 treats the end of text as a delimiter (default 0) |
| if_not_found | Optional | value when not found (default #N/A) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTAFTER(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `TEXTAFTER` (drop-in) on hosts without the native function, and as `EG.TEXTAFTER` on modern Excel that has it.
