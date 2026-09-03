# Cell Right-Click Menu

**Location**: right-click a cell range → context menu

Loading EGTools++ adds items to the cell right-click menu.
They appear in the same place for regular cells, table (list) ranges and Page Layout
view.

## Date Picker

Pops up a calendar next to the cell you right-clicked and enters the date you
pick — two clicks instead of typing a date.

1. Right-click a cell and choose **Date Picker**; a calendar appears to the
   right of the active cell (it shifts to the left when it would leave the
   screen). If the current cell value is a date, the calendar opens **on that
   date's month with it selected**; otherwise it starts on today.
2. Use `◀◀`/`◀`/`▶`/`▶▶` at the top to move by year/month. **Clicking a date
   enters it into the cell** and closes the calendar (the grayed previous/next
   month dates are clickable too).
3. **Clicking the "Today: yyyy-mm-dd" line at the bottom jumps to today.**
   Cancelling with `Esc` or the close button (X) leaves the cell unchanged.

- Sundays and **holidays are red**, Saturdays blue, and today is outlined.
  Hovering over a holiday shows its name at the bottom.
- Holidays are the built-in Korean calculation; with a data.go.kr key
  registered, **temporary holidays** are included too — see
  [Calendar — Holidays](calendar.md#holidays).

## Open URL

Finds an `http://` or `https://` URL in the selected cell's content and opens it in
the default browser.

- The first URL is found and opened even when it is mixed with other text in the cell.
- With multiple cells selected, the **first (top-left) cell** is used.
- If the cell contains no URL, nothing happens.

This opens plain-text URLs that have no hyperlink formatting in two clicks — no need
to enter edit mode or create a hyperlink.

---

More context-menu items may be added over time.
