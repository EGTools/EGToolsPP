# ARRAYTOTEXT

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

Returns a text representation of an array.

## Syntax

```
=ARRAYTOTEXT(array, [format])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| array | Required | the array |
| format | Optional | 0 concise (default), 1 strict |

## Returns

Returns a single text value (scalar) that joins the whole array. The implementation itself returns no error values.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ARRAYTOTEXT({1,2;3,4})` | 1, 2, 3, 4 | concise format |
| `=ARRAYTOTEXT({"a",1;"b",2},1)` | {"a",1;"b",2} | strict format |

## Notes

- Concise (0) format joins every element with ", " regardless of rows/columns.
- Strict (1) format produces {a,b;c,d}; only text is quoted, inner quotes doubled ("").
- Related function: VALUETOTEXT
- Supported: Excel 2010+. Registered as `ARRAYTOTEXT` (drop-in) on hosts without the native function, and as `EG.ARRAYTOTEXT` on modern Excel that has it.
