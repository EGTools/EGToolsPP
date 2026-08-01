# TEXTJOIN

**Category**: MS Excel compatibility function · **Native since**: Excel 2016/2019

Joins text using a delimiter, optionally ignoring empty cells.

## Syntax

```
=TEXTJOIN(delimiter, ignore_empty, text1, ...)
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| delimiter | Required | separator between items |
| ignore_empty | Required | TRUE to skip empty cells |
| text1 | Required | text or range |
| ... | Optional | the argument(s) above may be repeated |

## Returns

Returns a single text value (scalar) joined with the delimiter. The implementation itself returns no error values.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=TEXTJOIN("-",TRUE,"a","","b")` | a-b | skip empty values |
| `=TEXTJOIN("-",FALSE,"a","","b")` | a--b | keep empty values |
| `=TEXTJOIN(",",TRUE,{1,2;3,4})` | 1,2,3,4 | join array elements |

## Notes

- Omitting ignore_empty is treated as TRUE.
- Up to 255 text arguments are supported.
- Related function: CONCAT
- Supported: Excel 2010+. Registered as `TEXTJOIN` (drop-in) on hosts without the native function, and as `EG.TEXTJOIN` on modern Excel that has it.
