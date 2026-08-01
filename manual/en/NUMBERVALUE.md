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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=NUMBERVALUE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `NUMBERVALUE` (drop-in) on hosts without the native function, and as `EG.NUMBERVALUE` on modern Excel that has it.
