# Forms

**Location**: ribbon `EGTools++` tab → `Forms` group

A group of features that fill a form (template area) drawn on a sheet with a
data list, repeatedly. It consists of a single split button whose default
action is **Make Labels**.

| Menu item | Description |
|---|---|
| [Make Labels](#make-labels) | Tile the form on one sheet, once per data row, substituting `{{fieldname}}` with values |
| [Mail Merge](mailmerge.md) | Clone the form sheet per data row into individual sheets/files (xlsx/PDF), print each and/or send email (file attached) |

## Make Labels

Clones the form once per data row while substituting `{{fieldname}}`
placeholders with each row's values, building a printable label sheet
(name tags, address labels, etc.).

### How it runs

1. **Designate the form area** — pick the form range in the range input box
   (default = the current selection's contiguous region; cancelling exits).
   Designating a single empty cell exits.
2. **Designate the data list** — pick the list range. The **first row holds
   the field names**; data starts on the second row.
3. **Enter the labels-across count** — how many labels to place side by side
   per row (1–10, default 1).
4. On a new sheet (form sheet name + `_1` pattern) the form is copied once per
   data row and each `{{fieldname}}` is replaced with that row's value.
   A completion notice is shown.

### Substitution rules

| Item | Rule |
|---|---|
| Placeholders | Only cells containing `{{fieldname}}` are substituted. A cell may hold several placeholders combined with other text (e.g. `Address: {{addr}}`) |
| No placeholders | If the form contains no `{{ }}` placeholders at all, empty forms are simply cloned once per data row |
| Empty field names | Columns whose header cell is empty can be referenced as `Empty1`, `Empty2`, … |
| Unknown field names | Referencing a field name that is not in the list shows a warning and aborts |
| Formatting | The form's formatting (borders, colors, …) is copied along, and **column widths / row heights are cloned** from the form |

### Notes

- Substituted cells are written as **text values** (not formulas).
- To build an individual sheet/file per row, use [Mail Merge](mailmerge.md)
  (labels = tiled repeatedly on one sheet, mail merge = one sheet/file per row).
