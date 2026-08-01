# BASE

**Category**: MS Excel compatibility function · **Native since**: Excel 2013

Converts a number into text in the given radix (base).

## Syntax

```
=BASE(number, radix, [min_length])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| number | Required | the number to convert (>= 0) |
| radix | Required | base, 2 to 36 |
| min_length | Optional | minimum length of the result (optional) |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=BASE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `BASE` (drop-in) on hosts without the native function, and as `EG.BASE` on modern Excel that has it.
