# Image/Picture

**Location**: ribbon `EGTools` tab → `Image/Picture` group

A group of features for inserting, aligning and extracting images on cells.
It consists of two split buttons — **Image** (default action = Insert Image;
menu: Insert Image, Insert by Name, Insert into Form) and **Fit** (default
action = Fit Selected; menu: Fit Selected, Fit All, Export All).

## Insert Image

Picks one image file and inserts it fitted to the current selection.

1. Choose an image file in the file dialog
   (jpg, jpeg, png, gif, bmp, tif, tiff, svg, emf, wmf, etc.; cancelling exits).
2. The picture is inserted over the selection's position and size with a
   **0.2 pt margin** (embedded in the document, not linked to the file).
3. The inserted picture is **named after the file name**, and its
   **aspect-ratio lock is turned off** so later cell-fitting can resize it freely.

## Insert by Name

Picks a folder and inserts, into each cell, the image file whose **name matches
the cell's value**. Useful to attach photos in bulk to a roster or product list
that already holds the names.

1. Choose the folder holding the images (cancelling exits).
2. Every cell value in the active sheet's used range (from A1) is checked
   against the folder for an image file with the same name.
   - The extension is optional (cell value `photo1` matches the file
     `photo1.jpg` either way). Case is ignored.
   - Only image extensions are considered
     (jpg, jpeg, png, gif, bmp, tif, tiff, svg, emf, wmf, etc.).
3. A matching image is inserted into that cell (the whole merged area for a
   merged cell) with a **0.3 pt margin**, embedded in the document. The
   picture is named after the file name without its extension, and the
   aspect-ratio lock is turned off.
4. A result message is shown at the end (folder notice / no pictures added).

If the same value appears in several cells, the image is inserted into each.

## Insert into Form

Clones the form while inserting the folder's images, in order, at the `{{}}`
placeholder positions. Useful to fill photo-sheet (photo report) templates.

1. **Designate the form area** — pick the form range in the range input box
   (default = the current selection; cancelling exits). The form must contain
   at least one cell holding only `{{}}` (the photo position); otherwise a
   notice is shown and the command exits.
2. **Enter the clone method** — `1` = clone the whole form sheet /
   `2` = clone the form area downward / `3` = clone it to the right.
3. **Choose the folder** — the folder's image files are used **sorted by
   name** (with no images, a notice is shown and the command exits).
4. The form is cloned until the images run out, inserting an image into each
   `{{}}` cell (merged area for merged cells) with a **0.3 pt margin**.
   The **cell value is set to the file name**, and the picture is named after
   the file as well. Leftover `{{}}` markers after the images run out are
   cleared. Progress is shown in the status bar.

- Method 1 copies the sheet, so sheet names grow as `original_1`, `_2`, ….
- Methods 2/3 copy entire rows/columns, so row heights and column widths are
  preserved (existing content at the clone position is cleared).

## Fit Selected

Fits the selected pictures to the cell under each picture's top-left corner.

1. Select the picture(s) and press the button; a **margin prompt** appears
   (0–10, default 0.3). Entering a non-numeric value or cancelling exits.
2. Each picture is fitted inside its top-left cell, inset by the margin.
   If that cell belongs to a **merged area**, the picture is fitted to the whole
   merged area.
3. Pictures rotated by **90 or 270 degrees** have their width and height swapped
   so they still fit the cell in their rotated state.

If no picture is selected, a "no images selected" notice is shown.
Only picture objects are processed; shapes, charts, etc. are left alone.

## Fit All

Fits **every picture** on the active sheet the same way as Fit Selected.

1. After you confirm the prompt, a **margin prompt** appears (0–5, default 0.3).
2. All pictures on the sheet are fitted to their own top-left cell
   (merged area if merged). Progress is shown in the Excel **status bar**
   with each picture's name.
3. A result message (done / no images) is shown at the end.

## Export All

Saves **every picture** on the active sheet to a folder of your choice **at
original quality**. Useful to get the photos placed on a sheet back as files.

1. With no pictures on the sheet, a notice is shown and the command exits.
   Choose the destination folder (cancelling exits).
2. The file name is the **value of the cell under the picture's top-left
   corner**. If that cell is empty, the name becomes
   `picturename_R{row}C{col}`. Characters not allowed in file names become
   `_`, and duplicate names get `_1`, `_2`, … appended. The extension is the
   original format stored in the document (jpg, png, etc.).
3. Internally a **copy of the workbook is saved to a temporary folder** and
   the original image data is extracted from it directly. Since the pictures
   are not re-rendered at display quality, **the quality they were inserted
   with is preserved exactly**, and the original document is never modified
   (temporary files are removed automatically).
4. The exported count and destination folder are reported at the end.

- The document must be in **xlsx/xlsm format**. Other formats (.xls, etc.)
  can show the "Picture export failed" notice.
- Only picture objects are exported; shapes and charts are not.

## Notes

- A margin of 0 makes the picture fill the cell completely; larger values leave a
  wider gap from the cell borders.
- The fit target is the **cell under the picture's top-left corner**. If it snaps to
  the wrong cell, move the picture over the intended cell and run it again.
