# Visible Cells

**Location**: ribbon `EGTools++` tab → `Visible Cells` group

A group of features for copying and pasting ranges under filters or hidden
rows/columns. It consists of two split buttons — **Copy** (default action =
Copy Visible Cells) and **Paste** (default action = Paste Values).

## Copy Visible Cells

Picks only the **visible cells** of the selection, remembers them as the copy
source and copies them to the clipboard.

1. Select the range to copy and press the button. If the selection is not a
   cell range, the command exits silently.
2. The visible cells of the selection — excluding cells hidden by filters or
   row/column hiding — are remembered as the copy source
   (workbook and sheet name plus the area addresses).
3. The same area is also copied to the Excel clipboard (marching-ants border),
   so it can be pasted right away with a normal paste.

## Copy Range

Remembers the **entire selection (hidden cells included)** as the copy source
— **remember only**. Nothing is copied to the clipboard, so nothing changes
on screen.

## Paste Values / Formulas / All

Pastes the remembered copy source **into the visible cells only** of the
current selection, in order. The three menu items differ only in what they
paste — **Paste Values** (calculated values, the split button's default
action), **Paste Formulas** and **Paste All** (values, formulas and formats).

### How it runs

1. First run [Copy Visible Cells](#copy-visible-cells) or
   [Copy Range](#copy-range) to remember a copy source.
2. Select the target range and press the menu item. Cells hidden by filters or
   row/column hiding in the target are skipped, and only the **visible cells**
   are filled in source order. With a single cell selected as the target, the
   paste starts at that cell and extends down/right.
3. With a multi-cell source, a **temporary workbook** is created internally to
   reshape the data around the target's hidden row/column positions before
   pasting. The temporary workbook is closed without saving automatically.
4. During this process the source sheet's **AutoFilter is temporarily removed
   and then restored with its criteria intact**. A single-cell source skips
   these intermediate steps and is filled straight into all visible target
   cells.

### When it refuses to run

| Situation | Notice |
|---|---|
| No copy source remembered | "Run [Copy Visible Cells] first" |
| Source workbook/sheet was closed or renamed | "The copy-source workbook/sheet is gone. Copy again" |

## Notes

- The source is remembered by workbook/sheet name and addresses. Closing the
  source workbook or renaming the sheet means you must copy again.
- If the visible area copied under a filter is extremely fragmented (its area
  address exceeds 255 characters), re-resolving the source can fail. The
  "copy again" notice is shown in that case too.
