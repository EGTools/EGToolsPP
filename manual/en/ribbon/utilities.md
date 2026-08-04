# Utilities

**Location**: ribbon `EGTools++` tab → `Utilities` group

A split button collecting frequently used helpers. Pressing the button itself runs the
default action, **Recalculate All**; the arrow opens a menu of ten items.

## Delete Styles

Batch-deletes the **custom cell styles** piled up in the active workbook.
Use it when styles copied in from other documents have snowballed and bloated
the file. A prompt asks how much to delete.

| Choice | Action |
|---|---|
| Yes | Delete every custom style, keeping only built-in styles |
| No | Keep built-in styles plus styles **actually in use**, delete the rest |
| Cancel | Abort |

- The "in use" check scans every sheet's used range cell by cell, so it can be
  **slow on large documents** (progress is shown in the status bar).
- Locked styles are unlocked before deletion; the deleted/failed counts are
  reported at the end.

## Delete Names

Deletes defined names in the active workbook that match any of the following.

| Deleted | Condition |
|---|---|
| Broken references | RefersTo contains `#NAME?` / `#REF!` / `#N/A` |
| Macro leftovers | RefersTo contains `BlankMacro1` |
| Hidden names | Names not visible in the Name Manager (Visible=False) |

Excel-internal names (`_xl*.*` pattern, e.g. print areas and filters) are never
touched. With many names, progress is shown in the status bar; the deleted count is
reported at the end (or a "nothing to delete" notice).

## Clear Errors

Wraps every **formula currently producing an error value** in the active
sheet's used range with `=IFERROR(formula,"")` so the errors disappear from
view. CSE array formulas (`{=…}`) are wrapped as a whole array. Non-formula
cells (typed-in error constants) are left alone; the fixed count is reported
at the end.

- The error's cause is not fixed — it is merely **displayed as an empty
  string**. Best suited for tidying up right before printing or sharing.
- If a wrapped formula errors again and you re-run the command, the IFERROR
  wrappers can nest.

## Clear Empty Text

Turns cells holding only a **zero-length string ("")** into truly empty cells
(clears their contents). Such cells — left behind by formula results or
external data — look empty but are not treated as blank by ISBLANK, COUNTA or
the end-cell (Ctrl+End) logic.

1. Select a range and run the command. **With only a single cell selected, the
   sheet's entire used range** becomes the target.
2. A prompt asks "Exclude formulas?" — **Yes** = formula cells returning `""`
   are preserved / **No** = formula cells are cleared too (the formula is
   deleted) / **Cancel** = abort.
3. The cleared-cell count is reported at the end. CSE array formulas are
   cleared as a whole array.

## Arrange Notes

Moves **every note (comment)** on the active sheet right next to its cell and
auto-sizes it to its content. For merged cells the note moves next to the merged
area. The number of arranged notes is reported at the end.

## Change Border Colors

After you pick a color in the color dialog (cancelling exits), all borders in the
selection — top/bottom/left/right plus both diagonals — change to that color in one
pass. **Line styles (solid/dashed/weight) are preserved**; only the color changes,
and no new borders are added where none exist.

## Draw Free Form

Creates a freeform shape tracing the outline of the **background-colored
cells** in the selection. Useful to extract a map or zone drawn by coloring
cells into a single shape.

1. Select a range containing colored cells and run the command. With no
   colored cells, the "Select a range containing colored cells" notice is
   shown and the command exits.
2. **White backgrounds (including no fill)** count as background; cells filled
   with any other color form the shape area.
3. The outline of the first **connected block** of colored cells is traced
   along the actual cell widths/heights, and a style preset is applied to the
   resulting shape.

- With several separate blocks, only the first block becomes a shape. Run
  again on a narrowed range for the others.
- If the shape cannot be created, the "Shape conversion failed" notice is
  shown.

## Freeze UDF Values

Replaces formulas containing EGTools functions (including the `EG.`/`x` prefix
variants) with their **calculated values**. A distribution helper: when the
document is sent somewhere without the add-in, the values survive instead of
formulas breaking into `#NAME?`.

1. **Scan** — all worksheet formulas are scanned read-only and the target
   count is shown in a confirmation prompt (with no targets, a notice is shown
   and the command exits).
2. **Save gate + backup** — if the workbook is unsaved you are asked to save,
   and a backup file (SaveCopyAs) is saved automatically right before
   applying. If the backup fails, nothing is changed.
3. **Freeze** — target formulas are replaced with their values. CSE arrays are
   frozen as the whole array, spill formulas as the whole spill range. The
   original formula is preserved in a cell note as `EgUdfRemoved: <formula>`
   (an existing note is appended after it).
4. The frozen-cell count and the backup file path are reported at the end.

- The target function list is managed from the same single source as the
  add-in's function registration list, so they always match.
- Whereas [Apply Legacy Compatibility](apply-compat.md) **converts function
  names** and keeps the formulas alive, Freeze UDF Values **turns the formulas
  into values**.

## Apply CheckBox

Makes the selection display 0/1 input like checkboxes.

1. Applies **data validation allowing only 0 or 1** to the selection.
2. Sets the number format to `[=1]"✅";[=0]"⬜";""`
   — entering 1 shows ✅, 0 shows ⬜, anything else shows blank.
3. Empty cells are filled with 0 and center-aligned.

Cell values remain plain 0/1 numbers, so they aggregate directly with `SUM`,
`COUNTIF`, etc.

## Recalculate All

Forces a full recalculation of **all open workbooks** (`CalculateFull`).
Use it to recompute from scratch — cells missed by automatic calculation, volatile
functions and external references included.

> Registering/removing keys for the public-API functions has moved to
> [Manage API Keys in the EGTools group](api-keys.md).
