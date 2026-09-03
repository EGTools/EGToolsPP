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

Returns a spilled array of item names; with sub_folder TRUE the columns are staggered by folder depth. Returns #VALUE! when the path is missing or empty, and #N/A when nothing matches.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=DIRFOLDER("C:\Data","*.xlsx;*.csv",FALSE,TRUE)` |  | Result depends on the folder contents |

## Notes

- Non-macro function; reads the folder directly through the Win32 API.
- Several filters can be given separated by ";"; * and ? wildcards are supported.
- Hidden items are excluded by default; set hidden to TRUE to include them.
- Supported: Excel 2010+. Always registered as `DIRFOLDER` on every Excel version.
