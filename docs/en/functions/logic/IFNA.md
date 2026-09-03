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

Returns value_if_na when the value is the #N/A error, otherwise the value itself. Given an array, it is applied element-wise and returns an array that spills.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IFNA(NA(),"none")` | none | Replaces #N/A |
| `=IFNA(1/0,0)` | #DIV/0! | Other errors pass through |
| `=IFNA(5,0)` | 5 | Normal value unchanged |

## Notes

- Both value and value_if_na accept arrays and are broadcast against each other, replacing only #N/A element-wise (scalars repeat, size-mismatched elements yield #N/A; the array result spills).
- Supported: Excel 2010+. Registered as `IFNA` (drop-in) on hosts without the native function, and as `EG.IFNA` on modern Excel that has it.
