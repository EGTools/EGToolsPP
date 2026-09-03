# JSONFILTER

**Category**: EGTools-only function

Parses JSON text into a two-column (path, value) table.

## Syntax

```
=JSONFILTER(json_text, [key_path])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| json_text | Required | JSON text to parse |
| key_path | Optional | path to filter, e.g. "data.items[0].name" (default: whole document) |

## Returns

Returns a two-column (path, value) spilled array — numbers stay numeric, true/false become booleans, null becomes an empty string. Returns #VALUE! when the text is empty, JSON parsing fails, or key_path is not found, and #N/A when the target is an empty object/array.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=JSONFILTER("{""a"":1,""b"":[true,null]}")` | {"a",1;"b[0]",TRUE;"b[1]",""} | Flatten the whole document |
| `=JSONFILTER("{""a"":{""b"":5,""c"":""x""}}","a")` | {"b",5;"c","x"} | Only below key_path |

## Notes

- Uses its own lightweight parser, no external library.
- The whole text must be valid JSON; comments and trailing commas are not supported.
- key_path uses the "a.b[0].c" form; array indexes are 0-based.
- Supported: Excel 2010+. Always registered as `JSONFILTER` on every Excel version.
