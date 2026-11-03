# Home Tab Alignment Group (Center Across Selection)

**Location**: Excel `Home` tab → `Alignment` group (replaced by EGTools++) → `Center Across Selection` below the merge button

When EGTools++ is loaded, the built-in **Alignment** group on Excel's `Home` tab is hidden and an
**Alignment** group with the same layout is rebuilt in its place. An add-in cannot insert controls
into a built-in Office group, so the whole group is recreated; the only difference is one extra
button, **Center Across Selection**, placed below the merge button. It changes only the horizontal
alignment to "Center Across Selection" without merging cells, so a title can be centered over several
columns while sorting, filtering and formula references keep working.

The replacement happens **only when** Excel's merge dropdown does not already have the native
"Center Across Selection" item (idMso `CenterAcrossSelection`). On an Excel that already has the native
item the default `Alignment` group is left untouched and the EGTools++ button does not appear — use
Excel's own item instead.

## Group layout

| Position | Control | Origin |
|---|---|---|
| Left, row 1 | Top Align · Middle Align · Bottom Align · Orientation | Excel built-in |
| Left, row 2 | Align Left · Center · Align Right · Decrease Indent · Increase Indent | Excel built-in |
| Right, row 1 | Wrap Text | Excel built-in |
| Right, row 2 | Merge & Center (dropdown) | Excel built-in |
| Right, row 3 | **Center Across Selection** | Added by EGTools++ |
| Dialog launcher | Format Cells: Alignment | Excel built-in |

The built-in controls are the hidden default group's own controls placed again, so their behaviour,
labels and keyboard shortcuts are Excel's.

## How it runs — Center Across Selection

1. Select the column range to spread across within a row, with the cell holding the content at the
   left edge (for example `A1:F1`). If several rows are selected, each row is handled on its own.
2. Click **Center Across Selection** in the `Home` tab → `Alignment` group.
3. The horizontal alignment of the selection is set to **Center Across Selection**
   (`xlHAlignCenterAcrossSelection`): in each row the content is displayed centered over the run of
   empty cells to its right. No cells are merged.
4. Clicking again on the same range sets the horizontal alignment back to **General**
   (`xlHAlignGeneral`) — the button is a toggle.

## Rules

| Situation | Behaviour |
|---|---|
| The whole selection is already "Center Across Selection" | Reverts to **General** alignment |
| Mixed (only part of the selection has it) or not applied | Applies **Center Across Selection** to the whole selection |
| A non-cell object is selected (shape, chart, picture, …) | Does nothing |
| Vertical direction | Only the horizontal alignment changes; content is not centered vertically (this is how Excel's "Center Across Selection" is defined). Use Top/Middle/Bottom Align for vertical placement |
| Undo | This is an add-in command, so it is not recorded in Excel's Undo (`Ctrl+Z`). To revert, click the button again or change the horizontal alignment in Format Cells |

## Notes

- Microsoft is gradually rolling out a "Center Across Selection" item inside the merge dropdown,
  starting with the Microsoft 365 Insider channel (September 2026 What's New). EGTools++ provides the
  feature on the `Home` tab so it is also available in Excel 2010–2024 and in Microsoft 365 builds
  that have not received it yet. On an Excel that already has the native item, EGTools++ detects it
  at load time and does not replace the `Alignment` group, so the two never appear together — use
  Excel's own item. The check runs once when EGTools++ loads, so if a channel or build change adds or
  removes the native item, the change takes effect after Excel is restarted.
- The same effect is available from the `Format Cells` dialog (`Ctrl+1`) → `Alignment` tab →
  Horizontal `Center Across Selection`. This button is a one-click shortcut to apply or clear it.
- To make several value-holding cells merely look merged, see
  [Merge Keeping Values](merge-split.md#merge-keeping-values); that command applies merge formatting,
  whereas this button only changes the horizontal alignment.
- When EGTools++ is unloaded, the replacement group disappears and Excel's default `Alignment` group
  reappears.
