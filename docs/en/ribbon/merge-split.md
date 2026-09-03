# Merge/Unmerge Cells

**Location**: ribbon `EGTools` tab → `Merge/Unmerge Cells` group

A group of features for working with merged cells. It consists of two split
buttons — **Merge** (default action = Merge Contents Only) and **Unmerge**
(default action = Split to Rows Only).

- **Merge menu, 9 items** (ribbon order): Merge Same Values · Merge Keeping
  Values · Merge Same Values Keeping · Merge Contents Only/with Style ·
  Merge Rows Only/with Style · Merge Columns Only/with Style
- **Unmerge menu, 5 items**: Unmerge and Fill · Split to Rows Only/with Style ·
  Split to Columns Only/with Style

## Merge Same Values

Turns vertical runs of equal values in the selection into real merged cells.
Useful to tidy up a sorted list where the same category repeats down a column.

### How it runs

1. Select the range to merge (a sorted list is recommended) and press the
   button.
2. In each column, **runs of the same value spanning two or more cells
   vertically** are merged. Runs of empty cells are not merged.
3. With multiple columns the merge is **hierarchical** — runs in a right
   column are broken at the group boundaries of the column(s) to its left.
   For example, with departments in column A and teams in column B, team
   merges also break wherever the department changes.

### Notes

- Values are compared by their displayed value (a formula's calculated result).
- Since only runs of equal values are merged, no value is lost even though it
  runs without a confirmation prompt (Excel keeps only the top-left value on
  merge, but every value in the run is the same).
- A multi-area selection (Ctrl+click) is processed area by area.

## Merge Keeping Values

Merges the selected area(s) **in appearance only**. Every cell keeps its own
value, so functions like `SUM`/`VLOOKUP`, filtering and sorting **keep
working**, while the screen shows only the top-left value like a merged cell.

1. Select the range(s) to look merged and press the menu item (multi-area
   selections are processed area by area).
2. Internally the current sheet is **copied into a temporary workbook**, the
   real merge is made in the copy, and only its **formatting** is pasted back
   onto the same area of the original. Values and formulas are untouched; the
   temporary workbook is closed without saving.

- Because the merge is appearance-only, values other than the top-left one are
  merely hidden from view, not deleted. Formulas referencing them still read
  them.

## Merge Same Values Keeping

Finds runs using the same rules as [Merge Same Values](#merge-same-values)
(vertical runs of two or more equal cells, hierarchical across columns), but
merges them **in appearance only** like
[Merge Keeping Values](#merge-keeping-values). Every cell keeps its value, so
filters and aggregation still work on the values after the merge.

- Existing merged cells inside the selection are first unmerged and filled (in
  the copy) before the runs are computed. Values in the original never change.

## Merge Contents Only / with Style

Merges the whole selection into **one merged cell** while joining the
contents.

1. Select the range and press the menu item. With two or more columns you are
   asked for the **column separator** (default `,`; cancelling exits).
2. Cells within a row are joined with the separator; **rows are joined with a
   line break**. Empty cells are skipped. Formula cells contribute their
   calculated text.
3. The area is merged and the joined text is written in. The result cell gets
   the **Text number format (@)**.
   - **Only**: the cell style is reset to Normal and Wrap Text is turned on.
   - **with Style**: eleven font attributes of **each character** of the
     source (font, size, color, bold, italic, underline, strikethrough,
     sub/superscript, …) are remembered and re-applied character by character
     to the joined text.

## Merge Rows Only / with Style

Processes **each row** of the selection with the content merge above, making
one merged cell per row. With two or more columns the separator is asked
**once** and reused for every row. Progress is shown in the status bar.

## Merge Columns Only / with Style

Makes **each column** of the selection into one merged cell. The vertical
values are joined with **line breaks**, so no separator prompt appears.

## Unmerge and Fill

Unmerges the merged cells inside the selection and fills the cells back with the value
each merge held. Useful to make filtering, sorting and formula references work on
plain values after unmerging.

### How it runs

1. Select the range to process and press the button. If the selection is not a cell
   range, the command exits silently.
2. Each merged cell in the selection is unmerged and the value it held is written back
   into the unmerged cells. Merged cells that were empty stay empty.
3. If a merge spanned multiple **columns**, the value is filled into the **leftmost
   column only** (a multi-row merge fills every row of that first column).

### Caution

- **Formulas become values.** If a merged cell contained a formula, its calculated
  result is filled in and the formula is lost. Back up first if you need to keep
  the formulas.

## Split to Rows Only / with Style

Splits the text of **one merged cell** at **line breaks** and distributes the
pieces downward, one per row. The reverse of the content merge.

1. Select the merged cell (or a single cell) and press the menu item.
2. The text is split at line breaks into the unmerged rows. If there are more
   pieces than the merge had rows, the missing **rows are inserted**.
3. Each target cell gets the Text number format (@) and its row height is
   auto-fitted. **with Style** preserves the source character formatting per
   piece.

- If there is no line break and only one piece results, a notice is shown and
  the command exits.

## Split to Columns Only / with Style

Splits the text of one merged cell at a **separator** and distributes the
pieces to the right, one per column.

1. Select the merged cell and press the menu item; the separator is asked
   (default `,`; cancelling or an empty value exits).
2. If there are more pieces than the merge had columns, **columns are
   inserted**; each target cell gets the Text number format (@) and its column
   width is auto-fitted. **with Style** preserves the source character
   formatting per piece.

## Common notes

### When it refuses to run

| Situation | Applies to | Notice |
|---|---|---|
| Selection is not a cell range | all | "Work on a cell range" |
| Selection overlaps a table (ListObject) | Merge Same Values · Merge Same Values Keeping · content merge · row/column merge | "Table ranges cannot be processed" — merged cells are not supported inside tables |
| Only one row selected | Merge Same Values · Merge Same Values Keeping | "A single row cannot be processed" |
| Multi-area (Ctrl+click) selection | content merge | "Only one area can be processed at a time" |
| More than 100 rows | content merge · column merge | "Ranges over 100 rows cannot be processed" |
| Already-merged cells included | content merge (whole selection is already one merge) · row/column merge | "The range contains already-merged cells. Unmerge first" |
| Not a single merged cell | split to rows/columns | "Select a single merged cell to run this" |
| Nothing to split | split to rows | "The cell has no line breaks (Alt+Enter) to split by" |
| Nothing to split | split to columns | "The cell has no delimiter to split by" |

### Caution

- **"with Style" reads and re-applies formatting character by character, so
  long texts take a while.** Progress is shown in the status bar. If character
  formatting is not needed, "Only" is much faster.
- Result cells of content merge/split carry the Text number format, so they
  are not auto-recognised as numbers/dates. Cells that held formulas end up
  as their calculated text.
