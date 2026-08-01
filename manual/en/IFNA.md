# IFNA

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Returns value_if_na if the value is the #N/A error; otherwise the value.

## Syntax

```
=IFNA(value, value_if_na)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| value | Required | the value or expression |
| value_if_na | Required | value to return if #N/A |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IFNA(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `IFNA` (drop-in) on hosts without the native function, and as `EG.IFNA` on modern Excel that has it.
