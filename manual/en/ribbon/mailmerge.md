# Mail Merge

**Location**: ribbon `EGTools++` tab → `Forms` group → `Labels` split-button menu → **Mail Merge**

Clones the **whole form sheet once per data-list row**, substituting `{{fieldname}}`
placeholders with that row's values, then — depending on the chosen option — saves
individual files (xlsx/PDF) and/or prints each copy. Sending mail is planned for a
future version; the SMTP settings will then be stored with the password excluded.

## How it runs

1. **Enter the option** — a number 0–3 (default 1).
   `0` = clone form sheets only (the clones are kept) / `1` = save files /
   `2` = print / `3` = save + print.
2. **Printer setup** — if printing is included (2 or 3), the printer setup dialog
   is shown first. Cancelling it aborts the whole run.
3. **Designate the form sheet** — pick any cell of the form sheet in the range
   input box. The **entire sheet** containing the picked range is used as the form.
4. **Designate the data range** — pick a cell in the data list. The selection is
   expanded to its contiguous region (CurrentRegion); the **first row holds the
   field names** and data starts on the second row.
5. For each data row the form sheet is cloned at the end of the workbook, every
   `{{fieldname}}` is substituted, and printing/saving runs per the option.
   "Items completed: N" is shown at the end.

## Rules

| Item | Rule |
|---|---|
| Placeholders | `{{fieldname}}` is matched against the list header row **ignoring case and surrounding spaces**. A cell may hold several placeholders combined with other text. If a field name is not in the list, an error naming that field is shown and the run aborts |
| Sheet/file names | The **first-column value** names the cloned sheet and the saved file. Rows with an empty first column are skipped; forbidden characters become `_`. If the sheet name already exists, an `_1`, `_2` … suffix is appended |
| Output location | Files are saved into an `Output\` subfolder of the workbook's folder. The **workbook must have been saved**; otherwise a notice is shown and the run aborts |
| PDF output | If the **first column's header is `PDF`** (case-insensitive), files are saved as PDF instead of xlsx |
| Open password | If the list has a `PassWord` (or Korean `암호`) column, that row's value becomes the xlsx **file-open password** (empty = no password; not applied to PDF) |
| Clone cleanup | With any option other than 0, each cloned sheet is deleted after printing/saving. With option 0 the sheets remain in the workbook |

## Notes

- Substituted cells are written as **text values** (not formulas).
- Saving copies the cloned sheet into a new single-sheet workbook, saves it and
  closes it, so each output file contains just that one sheet.
- Difference from [Make Labels](forms.md#make-labels): labels **tile the form
  repeatedly on one sheet**, while mail merge builds an **individual sheet/file
  per row**.
