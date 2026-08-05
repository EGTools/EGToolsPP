# Calendar

**Location**: ribbon `EGTools` tab → `Calendar` group

A group of features that build calendar and schedule sheets. It consists of two
split buttons — **Calendar** (default action = Monthly; menu: Monthly, Annual A,
Annual B) and **Schedule** (default action = Daily B; menu: Weekly, Daily A,
Daily B).

Common behaviour:

- The date/month prompt is pre-filled with today (this month); `-`, `/` and `.`
  separators are all accepted. Cancelling or entering an unparseable value
  exits silently.
- The new sheet is added **after** the active sheet, with print setup
  (orientation + fit to one page), gridlines hidden and panes frozen.

## Monthly

Enter a year-month (`YYYY-MM`) to build a monthly calendar sheet on a
six-week (42-day) grid.

- The first column of each week shows the **week number** as `W##`
  (Wednesday-based week of the year).
- **Holidays** show their name next to the date, and the whole day block is
  highlighted with a light-yellow background and red font.
- Days whose **lunar day is 1, 11 or 21** get a `(MM/dd)` lunar date mark.
- Days spilling over from the previous/next month are shown in grey `M/D`
  (January also reflects last year's Christmas, December next year's New
  Year's Day), and weeks entirely outside the month are removed.
- Sundays are red, Saturdays blue. Print is set to landscape, one page.

## Annual A / Annual B

Enter a year (2000–2050) to build an annual calendar sheet with all twelve
months on one page.

- **Annual A**: 3 columns × 4 rows, portrait print / **Annual B**: 4 columns ×
  3 rows, landscape print. Only the layout differs.
- **Holiday** dates are highlighted with a yellow background and red font, and
  the holiday name is attached as a **cell comment**.
- Month titles have a green background; Sundays are red, Saturdays blue.

## Weekly

Enter a date (`YYYY-MM-DD`) to build a weekly schedule sheet for the week
containing that date (Sunday–Saturday).

- The title row shows the year/month, the week's start–end dates and the week
  number.
- Each weekday column holds the date, lunar date (`MM/dd`), holiday name
  (highlighted) and a free schedule area, followed by a time-schedule area
  with a **08:00–21:00 axis in 30-minute steps**.
- Sunday red / Saturday blue; print is set to landscape, one page.

## Daily A / Daily B

Enter a date (`YYYY-MM-DD`) to build a one-day schedule sheet. The sheet is
named `MMdd` (`MMdd_1` … on duplicates).

- The layout consists of **Main Tasks** (a checkbox column plus task text),
  a **Time Schedule** (08–21 h) and **Task Notes**. Daily A is a two-pane
  layout (tasks/time schedule + notes column), Daily B a three-pane layout
  (tasks | time | notes).
- The checkbox column gets 0/1 data validation with the ✅/⬜ number format,
  so completion can be aggregated as numbers (same format as
  [Apply CheckBox](utilities.md#apply-checkbox)).
- **Unfinished tasks are carried over automatically**: tasks left unchecked
  (blank/0) on the most recent past daily sheet in the workbook are copied
  into the new sheet's Main Tasks. Daily sheets are identified by a hidden
  comment on cell A1 (`EGCalendar:date`).
- The title shows the date, weekday and lunar date; on a holiday its name is
  shown below the title.
- Print is set to landscape, one page.

## Holidays

- The basis for holidays is the **built-in calculation** — solar and lunar
  holidays, Labor Day and the **substitute-holiday rules** (per their effective
  years). The calculation is shared with the
  [KOREANHOLIDAYS](../KOREANHOLIDAYS.md) function.
- On top of that, the **holiday API is merged in by default**: the stored
  data.go.kr key ([EGTools → Manage API Keys](api-keys.md)) is used to query the
  holiday-information service so that **temporary holidays and election days**
  missing from the built-in calculation also appear. The same rule applies to
  **every sheet type** — Monthly, Annual A/B, Weekly and Daily A/B.
- When there is no internet connection, no key registered, or the key has
  expired or been rejected, the sheet is **built from the built-in calculation
  only** and the following notice is shown **once** after the command finishes
  (together with the reason for the failure).

  > The holiday API was unavailable, so the built-in calculation was used.
  > Temporary holidays and election days are not shown.

  Sheet creation itself is never aborted. If you see this notice, register or
  renew the key and run the command again to get temporary holidays.
- Lookup results are cached per year, so building several calendars for the same
  year queries the service only once (failures are not cached, so fixing the key
  takes effect immediately).
- Lunar dates use the built-in lunar table (same as the
  [TOLUNAR](../TOLUNAR.md) function); dates outside its range show no lunar
  mark.
