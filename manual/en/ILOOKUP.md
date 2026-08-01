# ILOOKUP

**Category**: EGTools-only function

Finds a match and copies the picture anchored at the matched cell into the calling cell.

## Syntax

```
=ILOOKUP(find_value, lookup_array, image_range, [if_not_found], [match_mode], [search_mode])
```

## Arguments

| Argument | Required | Description |
|---|---|---|
| find_value | Required | the value to search for |
| lookup_array | Required | the array or range to search |
| image_range | Required | range whose cells anchor the pictures |
| if_not_found | Optional | value when no match (default #N/A) |
| match_mode | Optional | 0 exact (default), -1 next smaller, 1 next larger, 2 wildcard, 3 regex |
| search_mode | Optional | 1 first-to-last (default), -1 last-to-first |

## Returns

Finds a match and, on success, returns an empty string ("") while the picture anchored at the matched cell is inserted separately, sized to the calling cell. Returns #REF! when lookup_array and image_range differ in size, the if_not_found value (default #N/A) when there is no match or no picture, and #VALUE! for an invalid regex or a COM/clipboard failure.

## Examples

| Formula | Result | Description |
|---|---|---|
| `=ILOOKUP("apple",A1:A5,B1:B5)` |  | Copy the picture of the matched row (image insertion, environment-dependent) |

## Notes

- Copies the picture (msoPicture/msoLinkedPicture) whose TopLeftCell is the matched cell in image_range, via the clipboard.
- match_mode: 0 exact (default), -1 next smaller, 1 next larger, 2 wildcard, 3 regex; search_mode: 1 first-to-last (default), -1 last-to-first.
- Moving pictures or changing formats alone does not trigger recalculation (press F9); macro-type, excluded from multithreaded recalculation.
- Supported: Excel 2010+. Always registered as `ILOOKUP` on every Excel version.
