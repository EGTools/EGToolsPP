# DIRFOLDER

**Category**: EGTools-only function

Lists the files (and optionally folders) of a directory.

## Syntax

```
=DIRFOLDER(path, [filter], [directory], [sub_folder], [hidden])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| path | Required | folder path |
| filter | Optional | file filter(s), e.g. "*.xlsx;*.csv" (default *.*) |
| directory | Optional | TRUE includes folder names |
| sub_folder | Optional | TRUE recurses into subfolders |
| hidden | Optional | TRUE includes hidden items |

## Returns

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DIRFOLDER(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Always registered as `DIRFOLDER` on every Excel version.
