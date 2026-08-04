# Ribbon Menu

Loading EGTools++ adds an **EGTools++** tab to the ribbon.
Feature buttons are organised into **groups**; new groups will be added as more features arrive.

## Group: Visible Cells

Copying and pasting under filters/hidden rows (two split buttons: Copy = Copy
Visible Cells / Paste = Paste Values). Details: [Visible Cells](visible-cells.md)

| Menu item | Description |
|---|---|
| [Copy Visible Cells](visible-cells.md#copy-visible-cells) | Remember only the visible cells of the selection as the copy source and copy them to the clipboard |
| [Copy Range](visible-cells.md#copy-range) | Remember the entire selection (hidden cells included) as the copy source |
| [Paste Values](visible-cells.md#paste-values--formulas--all) | Paste the remembered source into the target's visible cells only, as values |
| [Paste Formulas](visible-cells.md#paste-values--formulas--all) | Paste into the visible cells only, as formulas |
| [Paste All](visible-cells.md#paste-values--formulas--all) | Paste values, formulas and formats into the visible cells only |

## Group: Merge/Unmerge Cells

Features for working with merged cells (two split buttons: Merge = Merge
Contents Only / Unmerge = Split to Rows Only).
Details: [Merge/Unmerge Cells](merge-split.md)

| Menu item | Description |
|---|---|
| [Merge Same Values](merge-split.md#merge-same-values) | Merge vertical runs of equal values (hierarchical across columns) |
| [Merge Keeping Values](merge-split.md#merge-keeping-values) | Merge in appearance only, keeping every value — formulas and filters keep working |
| [Merge Same Values Keeping](merge-split.md#merge-same-values-keeping) | Visually merge each run of equal values while keeping the values |
| [Merge Contents Only/with Style](merge-split.md#merge-contents-only--with-style) | Merge the area into one cell, joining contents with separator/line breaks (with Style keeps character formatting) |
| [Merge Rows Only/with Style](merge-split.md#merge-rows-only--with-style) | Make each row one merged cell, joining its contents |
| [Merge Columns Only/with Style](merge-split.md#merge-columns-only--with-style) | Make each column one merged cell, joining contents with line breaks |
| [Unmerge and Fill](merge-split.md#unmerge-and-fill) | Unmerge cells and fill the merged value back in (formulas become values) |
| [Split to Rows Only/with Style](merge-split.md#split-to-rows-only--with-style) | Split merged-cell text at line breaks into rows (inserting rows as needed) |
| [Split to Columns Only/with Style](merge-split.md#split-to-columns-only--with-style) | Split merged-cell text at a separator into columns (inserting columns as needed) |

## Group: Image/Picture

Inserts, aligns and extracts images on cells (two split buttons: Image = Insert
Image / Fit = Fit Selected). Details: [Image/Picture](pictures.md)

| Menu item | Description |
|---|---|
| [Insert Image](pictures.md#insert-image) | Pick an image file and insert it fitted to the selection |
| [Insert by Name](pictures.md#insert-by-name) | Find the folder image matching each cell value and insert it into that cell |
| [Insert into Form](pictures.md#insert-into-form) | Clone the form and insert the folder's images at the `{{}}` placeholders in order |
| [Fit Selected](pictures.md#fit-selected) | Fit the selected pictures to their top-left cells |
| [Fit All](pictures.md#fit-all) | Fit every picture on the sheet to its cell in one pass |
| [Export All](pictures.md#export-all) | Save every picture on the sheet to a folder at original quality |

## Group: Calendar

Builds calendar/schedule sheets (two split buttons: Calendar = Monthly /
Schedule = Daily B). Details: [Calendar](calendar.md)

| Menu item | Description |
|---|---|
| [Monthly](calendar.md#monthly) | Monthly calendar sheet (week numbers, holidays, lunar dates) |
| [Annual A](calendar.md#annual-a--annual-b) | Annual calendar, portrait (3×4) |
| [Annual B](calendar.md#annual-a--annual-b) | Annual calendar, landscape (4×3) |
| [Weekly](calendar.md#weekly) | Weekly schedule (Sun–Sat, 08:00–21:00 time axis) |
| [Daily A](calendar.md#daily-a--daily-b) | Daily schedule, two panes (carries over unfinished tasks) |
| [Daily B](calendar.md#daily-a--daily-b) | Daily schedule, three panes (carries over unfinished tasks) |

## Group: Forms

Fills a form with a data list, repeatedly (split button, default = Make Labels).
Details: [Forms](forms.md)

| Menu item | Description |
|---|---|
| [Make Labels](forms.md#make-labels) | Clone the form per data row, substituting `{{fieldname}}` with values |
| [Mail Merge](mailmerge.md) | Clone the form sheet per data row into individual sheets/files (xlsx/PDF), print each and/or send email (file attached) |

## Group: Utilities

Frequently used helpers (split button, default = Recalculate All). Details: [Utilities](utilities.md)

| Menu item | Description |
|---|---|
| [Delete Styles](utilities.md#delete-styles) | Batch-delete custom cell styles (choose to keep built-in / in-use ones) |
| [Delete Names](utilities.md#delete-names) | Delete defined names with broken references or hidden names |
| [Clear Errors](utilities.md#clear-errors) | Wrap erroring formulas with `IFERROR` so errors disappear |
| [Clear Empty Text](utilities.md#clear-empty-text) | Turn zero-length-string cells into truly empty cells |
| [Arrange Notes](utilities.md#arrange-notes) | Move every note next to its cell and auto-size it |
| [Change Border Colors](utilities.md#change-border-colors) | Change all border colors in the selection (line styles kept) |
| [Draw Free Form](utilities.md#draw-free-form) | Create a freeform shape tracing the outline of the colored cells |
| [Freeze UDF Values](utilities.md#freeze-udf-values) | Replace EGTools formulas with values, keeping the formula in a note (for sharing) |
| [Apply CheckBox](utilities.md#apply-checkbox) | Cell format showing 0/1 input as ✅/⬜ |
| [Recalculate All](utilities.md#recalculate-all) | Force a full recalculation of all open workbooks |

## Group: EGTools

A split button collecting batch function-name conversion, SMTP settings, API key
management and add-in information (default = About EGTools++).

| Menu item | Description |
|---|---|
| [Apply Legacy Compatibility](apply-compat.md) | Batch-convert modern/native (`_xlfn.*`) function tokens to EGTools-compatible names |
| [Restore Native Functions](restore-native.md) | Batch-restore EGTools function names to native built-ins |
| [SMTP Settings](smtp-settings.md) | Register the SMTP server for mail-merge sending (password never stored) |
| [Manage API Keys](api-keys.md) | Register/remove public-API keys (juso · data.go.kr · vworld — the data.go.kr key is shared by holidays, calendars and business-registration lookups) |
| About EGTools++ | Shows version, distribution page and base libraries; notifies with a link when a newer GitHub Release is available |

## Cell Right-Click Menu

Besides the ribbon, items are also added to the cell context menu.
Details: [Cell Right-Click Menu](context-menu.md)

| Item | Description |
|---|---|
| [Date Picker](context-menu.md#date-picker) | Pop up a calendar next to the cell and enter the clicked date |
| [Open URL](context-menu.md#open-url) | Open the http/https URL found in the cell in the default browser |

## Insert Tab CheckBox

An [Apply CheckBox](utilities.md#apply-checkbox) button is also added to Excel's
`Insert` tab. It is not shown on Microsoft 365, which has native checkboxes.

<!-- Future features are added here as new group sections. -->
