# Nous Atlas

A spreadsheet for worldbuilding: workbooks that are mostly text, spread over
many sheets, plus maps. A sibling of Nous PDF.

Files stay ordinary `.xlsx`. Excel, LibreOffice and scripts (openpyxl) can
still open everything Atlas saves.

**Download** for Windows, Mac or Linux: https://nousatlas.lionthroat.com

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

## Layers

Map sheets (a row lettered A, B, C… with 1, 2, 3… down its left side) are found
automatically. A **layer** is anything with a footprint on a map: creatures,
plants, characters, paths, quests.

- **Select squares, then add them**: select squares on the map (click or drag)
  and press **Ctrl+L** or click **Add selected squares to …**.
  **Ctrl+Shift+L** / **Remove** takes them out. Optionally, **Paint by
  dragging** lets you drag over squares instead (Shift-drag removes).
- **Times and rarity are the workbook's own** (Format → Layer times and
  rarity, or the link in the Layers section): no time at all (the default for
  new workbooks), Day / Night, Dawn / Day / Dusk / Night, hourly, six 4-hour
  ticks grouped Day / Night, or your own named periods with optional groups.
  Rarity is off by default, or up to three names of your own. With no time or
  no rarity, those controls don't appear. Changing the scheme keeps what it
  can; anything that doesn't fit becomes "all the time".
- **Show** picks what the map displays. With two to four periods (or groups)
  each square splits: two diagonally, four into quarters. Day is the layer's
  colour; night is a darker, bluer shade of it.
- The selected layer is drawn at full strength; other layers' areas fade
  (**Show every ticked area at full strength** turns that off). Paths and
  markers always show.
- **Square editor**: select one square to see every layer in it, by period;
  click to change.
- In any other sheet, right-click a name → *Make a layer for “…”* (or *Go to
  the … layer* once it exists).

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

Each layer draws as an **area** (fills squares: ranges, regions), a **path**
(a line joining neighbouring squares: roads, tracks, rivers) or a **marker**
(a dot: one-off things). Square cards also name the **terrain**, read from
the map sheet's own key (a filled swatch beside its label). **Alt+click**
any cell to pick up its fill; select cells and press **Ctrl+Shift+F** (or
the fill button) to paint it.

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
in a temp folder. Most need a private sample workbook in `tests\fixtures\`;
`tests\test_smoke.py` makes its own, so it's the one GitHub runs.

## Building

`powershell -ExecutionPolicy Bypass -File build.ps1` makes
`dist\NousAtlas\NousAtlas.exe`. The version lives in `VERSION`. With Inno
Setup 6 installed it also makes `installer-output\NousAtlas-Setup.exe`.

Releases are built by GitHub (`.github/workflows/build.yml`): bump `VERSION`,
commit, then `git tag v0.2.0 && git push && git push --tags`. That builds the
Windows installer, both Mac zips (`NousAtlas-mac.spec`) and the Linux .deb
(`build_linux.sh`), checks each one installs and opens a workbook, and
publishes a GitHub Release. The download page is `site/`, a static Cloudflare
Worker: `cd site && npx wrangler@4 deploy`.

## License

MIT (see `LICENSE`). The downloadable builds bundle Qt / PySide6 (LGPL-3.0) as
separate, replaceable library files.
