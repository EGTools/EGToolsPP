# IMAGE

**Category**: MS Excel compatibility function · **Native since**: Excel 2024 / Microsoft 365

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

<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->

## Examples

| Formula | Result | Description |
|---|---|---|
| `=IMAGE(...)` | | <!-- TODO --> |

## Notes

<!-- TODO: differences from the Excel/Google original, related functions -->
- Supported: Excel 2010+. Registered as `IMAGE` (drop-in) on hosts without the native function, and as `EG.IMAGE` on modern Excel that has it.
