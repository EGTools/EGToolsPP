# IMAGE

**Category**: MS Excel compatibility function · **Native since**: Excel 2024

Inserts a picture into a cell from a source URL or file path.

## Syntax

```
=IMAGE(source, [alt_text], [sizing], [height], [width])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| source | Required | image URL or file path |
| alt_text | Optional | alternative (accessibility) text (optional) |
| sizing | Optional | 0 fit & keep ratio, 1 fill, 2 original size, 3 custom (optional) |
| height | Optional | height in pixels when sizing is 3 (optional) |
| width | Optional | width in pixels when sizing is 3 (optional) |

## Returns

Always returns an empty string ("") as a scalar; after calculation it inserts a floating picture shape sized to the calling (merged) cell. An empty source gives #VALUE!; if the caller cell address cannot be resolved it returns #REF!.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMAGE("https://www.google.com/favicon.ico")` |  | result depends on the image source |

## Notes

- Unlike native 365 IMAGE it does not produce an in-cell value: it inserts a floating picture shape (EG.Image_Sheet_Cell) over the cell, replacing any same-named shape on recalculation.
- alt_text, height and width are accepted for compatibility but ignored; resize_mode 0 (default) keeps the aspect ratio.
- On modern Excel where native IMAGE exists, this function is registered as EG.IMAGE.
- Supported: Excel 2010+. Registered as `IMAGE` (drop-in) on hosts without the native function, and as `EG.IMAGE` on modern Excel that has it.
