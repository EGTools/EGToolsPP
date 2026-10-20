# List/Array Cells to Text

**Location**: ribbon `EGTools` tab → `EGTools` group → split-button dropdown menu

**Lists in Cells** (in-cell lists/arrays) from the Microsoft 365 beta (version 2610,
build 20520+) are stored as rich values, so older Excel (including regular Microsoft 365)
shows them only as `#UNKNOWN!` (`#VALUE!` in Excel 2016) and cannot read the items. This
command reads the richData of the saved file directly and rewrites each list cell in the
**text notation** below. The converted text can then be evaluated with the compatibility
functions [HAS](../functions/microsoft-365/HAS.md) · [HASANY](../functions/microsoft-365/HASANY.md) ·
[HASALL](../functions/microsoft-365/HASALL.md) · [FLATTEN](../functions/microsoft-365/FLATTEN.md)
(on Microsoft 365 they are registered with the `EG.` prefix, e.g. `EG.HAS`, so they coexist with the native ones).

## Text notation

| List shape | Notation | Example |
|---|---|---|
| One-column list (no nesting) | Comma-separated — same as the `ARRAYTOTEXT` concise format. An item containing `,` `;` `{` `}` `"` or leading/trailing spaces, or text that would read as a number/TRUE/error, is quoted as `"…"` (inner `"` doubled to `""`) | `Mon, Tue, Wed` / `"1,000", "TRUE"` |
| 2-D or nested array | Array-constant form — the `ARRAYTOTEXT` strict format (all text quoted, no spaces, `;` rows, `,` columns) extended with nested braces | `{"a",1;"b",2}` / `{"a",{"b","c"},"d"}` |
| List with a single numeric/logical item | Written as the numeric/logical **value itself**, not as text | `3`, `TRUE` |

- In the beta Excel, prefixing array-constant text with `=` turns it straight back into an in-cell array formula.
- The HAS family reads both notations. Items of a comma-separated list or inside `{…}` are read as
  numbers when they look numeric, as logicals for TRUE/FALSE and as errors for error names such as
  `#N/A`; a single text without commas or braces is not classified and stays one text item. Text that
  cannot be parsed (e.g. an unclosed quote) is treated as a single item without raising an error.

## How it runs

1. **Read the file** — a copy of the current workbook (SaveCopyAs) is unpacked into a temporary
   folder and its richData (list/array rich values) is read. Unsaved edits are included in the copy.
2. **Target check** — if there are no list cells, a notice is shown and the command exits.
3. **Confirmation** — reports that N value list cells will be converted. If M list cells are formula
   results, it asks whether to freeze those to text as well: **Yes** = all / **No** = value cells only /
   **Cancel**. (Formula list cells are stored as `_xlfn._ARRAY(...)`, which is `#NAME?` on older Excel
   anyway. Choosing No with no value cells exits.)
4. **Save gate + backup** — if the workbook is unsaved you are asked to save (declining or cancelling
   aborts), and a backup file is saved automatically right before applying (same folder as the workbook,
   `name_egbak_<timestamp>.<ext>` — the same procedure as [Apply Legacy Compatibility](apply-compat.md)).
   If the backup fails, nothing is changed.
5. **Apply** — each list cell is rewritten as text (or a value) according to the rules below.
6. **Result** — shows the number of converted cells, the number of skipped cells (if any) and the backup path.

## Rules

| Target | Handling |
|---|---|
| Value list cell | Only cells whose current value is an error (`#UNKNOWN!`/`#VALUE!`) are rewritten. Cells edited after saving (no longer an error) are skipped |
| Formula list cell | Converted only when you chose **Yes**. The original formula is preserved in the cell note as `EgListRemoved: <formula>` (an existing note is appended after it) |
| CSE array area | The whole area is frozen to values first, then the text is written to that cell |
| Text that Excel would auto-interpret as a formula/number/TRUE | The cell format is set to Text (`@`) before writing |
| List with a single numeric/logical item | Written as the numeric/logical value itself, not as text |
| Rich values other than `_array` (in-cell pictures, linked data types, etc.) | Not targeted |

## Notes

- The command also works in the beta Excel itself, so it can be used to convert in advance before sharing a workbook.
- After conversion, HAS / HASANY / HASALL / FLATTEN give the same results. If the workbook was saved in the
  beta Excel with formulas such as `=HAS(...)`, [Apply Legacy Compatibility](apply-compat.md) converts the
  `_xlfn.HAS` and `_xlfn._xlws.FLATTEN` tokens to the EGTools-compatible names.
- Recommended order: (1) List/Array Cells to Text, (2) [Apply Legacy Compatibility](apply-compat.md) (or
  [Restore Native Functions](restore-native.md) on non-beta Microsoft 365). Keeping this order is what makes
  formulas such as `=HAS(...)` return correct results instead of `#VALUE!`. If you answered **No** and left
  the formula list cells alone, Apply Legacy Compatibility rewrites their `_xlfn._ARRAY(...)` formulas to
  `ARRAYTOTEXT(…,1)`.
- To undo, open the backup file. For formula cells the original formula can also be read from the `EgListRemoved:` note.
