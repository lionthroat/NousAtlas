# Nous Atlas

A spreadsheet for worldbuilding: workbooks that are mostly text, spread over
many sheets, plus maps. A sibling of Nous PDF.

Files stay ordinary `.xlsx`. Excel, LibreOffice and scripts (openpyxl) can
still open everything Atlas saves.

## What it does

- **Sheets in a sidebar**, in collapsible groups you make yourself. Drag to
  reorder or regroup (Excel's tab order follows). `Ctrl+P` jumps to a sheet by
  typing part of its name.
- **Back and forward** (`Alt+←` / `Alt+→`, or the mouse's side buttons)
  between the places you've been, as in a browser.
- **Smooth scrolling** by the pixel, so tall rows don't jump.
- **Compact rows** (on by default) cap tall rows at three lines. The **Cell
  panel** on the right shows and edits the whole text (`Ctrl+Enter` saves),
  plus the cell's note and formula result.
- **Find across all sheets** (`Ctrl+F`), including notes and formulas.
- **Formatting kept to what matters**: font, size, bold, italic, underline,
  strikethrough, text and fill colour, wrap, horizontal and vertical
  alignment, merge and center, borders, cell styles.
- **Your colours, not Excel's.** The colour buttons offer the workbook's own
  palette first, then Nord, Dracula and Biomes. *Format → Palette* names your
  colours, and changing one there can change it everywhere it's used.
  *Format → Swap colours for your palette* remaps a whole imported workbook.
- **Styles**: *Format → Make styles from repeated formatting* turns the
  formatting a workbook already repeats into named styles. Change a style
  (*Update style to match this cell*) and every cell using it follows.
- **Formulas**: `+ - * / ^ & %`, comparisons, cell and range references
  (including other sheets), and SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ROUND,
  ROUNDUP, ROUNDDOWN, INT, ABS, SQRT, MOD, POWER, PI, IF, IFERROR, AND, OR, NOT,
  CONCAT, LEN, UPPER, LOWER, TRIM. Inserting or deleting rows and columns
  keeps every reference in step.
- **Margins and spacing per sheet** (*Format → Sheet margins and spacing*):
  a margin between the pane's edge and column A, cell padding (sides, and
  top and bottom), and cell spacing between cells, like an old HTML table's
  margin, cellpadding and cellspacing. Display only, so columns and formulas
  stay the same.
- **Themes**: Nord (default), Dracula and Snow. On dark themes, dark text that
  was written for a white page is shown lighter. That's display only; the
  file isn't changed.

## Range layers

Map sheets (a row lettered A, B, C… with 1, 2, 3… down its left side) are found
automatically. Open one and use the **Ranges** tab:

- One **layer** per creature or plant, each with its own colour.
- Times are the game's six 4-hour blocks: 00, 04, 08, 12, 16, 20. **Day** is
  08–20 and **Night** is 20–08. **Split day / night** shows day in each square's
  top-left half and night in its bottom-right half.
- **Paint** with Common, Uncommon or Rare. Shift-drag (or the Erase brush)
  removes squares. Painting applies to the times that are switched on.
- Double-click a layer to show only that one.
- **Here** lists everything in the selected square, block by block.
- In any other sheet, right-click a name and choose *Make a range layer for
  “…”*. After that, selecting that row shows a button that jumps straight to
  its range on the map.

## Square cards

Hover over a map square for a card listing what's there: **places** (any row
that links to the square, such as a Gazetteer row whose coordinates link to
G6), **life** (range layers painted there, labelled by the sheet each one
is written up on), and a **note** for the square. Click anything in the
card to go there; Alt+← comes back. Select a square to see the same card in
the Cell panel and edit its note there.

Links draw as buttons. Right-click a square like `G6` (or a whole column of
them) → *Link to a map*. Following a link pulses a ring around the square
it lands on. On maps, the selection is a two-tone ring that shows up on any
colour.

Ranges are saved in a plain sheet called **Ranges** (Creature | Map | Times |
Abundance | Squares | Colour), with squares written like `I6` or `G9:K12`.
Atlas rewrites that sheet on save and hides it from the sidebar (*View → Show
the Ranges data sheet*). Atlas's own settings (groups, palette, which sheets
are maps) live in a very hidden sheet called `_NousAtlas`.

## What doesn't come over from Excel

*File → What came over from Excel* lists anything a given file uses that Atlas
can't show. Conditional formatting and dropdown lists are kept in the file but
not shown or enforced. Charts, pictures and macros aren't kept, so a file with
any of them saves as a new file instead of over the original.

## Running from source

```
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe nousatlas.py [workbook.xlsx]
```

Tests (each prints OK): `tests\test_formula.py`, `tests\test_model.py`,
`tests\test_app.py`, `tests\test_palette.py`. They only ever work on copies,
in a temp folder.

## Building

`powershell -ExecutionPolicy Bypass -File build.ps1` makes
`dist\NousAtlas\NousAtlas.exe`. The version lives in `VERSION`.
