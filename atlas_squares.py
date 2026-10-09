"""What's on a map square: the square card.

Pulled together from data that's already in the workbook, plus one layer of
Atlas's own:

- places: rows anywhere that link to the square (a Gazetteer row whose Map
  Coordinates cell links to G6 makes G6 "RED-m6 · Scout's Transect")
- life: range layers painted on the square, labelled by the sheet each
  creature or plant is written up on (Fauna, Flora...)
- a note per square, kept in Atlas's settings (Excel has nowhere to put it)

Links in the card use "atlas:<sheet>|<row>|<col>" so the window can follow them.
"""

import html
from urllib.parse import quote, unquote

from atlas_formula import Ref, num_to_col
import atlas_ranges as R

MAX_LABEL = 60


def href(sheet, row, col):
    return f"atlas:{quote(sheet)}|{row}|{col}"


def parse_href(link):
    if not link.startswith("atlas:"):
        return None
    sheet, row, col = link[6:].rsplit("|", 2)
    return unquote(sheet), int(row), int(col)


class SquareIndex:
    """Built lazily; rebuilt whenever the book changes."""

    def __init__(self, book, skip_sheet=None):
        self.book = book
        self.skip_sheet = skip_sheet       # fn(ws) -> True for sheets not to search (the Ranges sheet)
        self.version = None
        self.places = {}                   # (map, (x, y)) -> [(sheet, row, col, label)]
        self.homes = {}                    # layer name (lower) -> (sheet, row, col)
        self.legends = {}                  # map title -> {fill hex: label}

    def refresh(self):
        if self.version == self.book.version:
            return
        self.version = self.book.version
        self.places, self.homes, self.legends = {}, {}, {}
        maps = self.book.meta["maps"]
        names = {l.name.lower() for l in self.book.ranges.layers} if self.book.ranges else set()
        for ws in self.book.user_sheets():
            if self.skip_sheet and self.skip_sheet(ws):
                continue
            is_map = ws.title in maps
            for (r, c), cell in sorted(ws._cells.items()):
                link = getattr(cell, "_hyperlink", None)
                if link is not None and link.location and not is_map:
                    target = self._square_for(link.location)
                    if target is not None:
                        self.places.setdefault(target, []).append((ws.title, r, c, self.row_label(ws, r, c)))
                v = cell.value
                if not is_map and isinstance(v, str) and v.strip().lower() in names:
                    self.homes.setdefault(v.strip().lower(), (ws.title, r, c))

    def legend(self, map_title):
        """{fill colour: name} read from the map sheet's own key: a filled
        swatch beside a text label, outside the grid. Under a heading with
        KEY or LEGEND in it, if there is one; colours used for more than one
        label are left out."""
        self.refresh()
        if map_title in self.legends:
            return self.legends[map_title]
        out = {}
        ws = self.book.sheet(map_title)
        info = self.book.meta["maps"].get(map_title)
        if ws is None or not info:
            return out
        r0, c0 = info["row"], info["col"]
        r1, c1 = r0 + info["rows"] - 1, c0 + info["cols"] - 1

        def fill_of(cell):
            if cell is None or not cell.has_style or not cell.fill or cell.fill.fill_type != "solid":
                return None
            from atlas_model import resolve_color
            return resolve_color(cell.fill.fgColor, self.book.theme)

        def label_right(r, c):
            for cc in range(c + 1, c + 4):
                cell = ws._cells.get((r, cc))
                if cell is not None and isinstance(cell.value, str) and cell.value.strip():
                    return cell.value.strip()
            return None

        heading = None
        for (r, c), cell in ws._cells.items():
            if isinstance(cell.value, str) and ("KEY" in cell.value.upper() or "LEGEND" in cell.value.upper()) \
                    and not (r0 <= r <= r1 and c0 <= c <= c1):
                heading = (r, c)
                break
        candidates = {}
        for (r, c), cell in ws._cells.items():
            if r0 <= r <= r1 and c0 <= c <= c1:
                continue
            if heading and not (r > heading[0] and abs(c - heading[1]) <= 1):
                continue
            hexv = fill_of(cell)
            label = label_right(r, c) if hexv else None
            if hexv and label:
                candidates.setdefault(hexv, set()).add(label)
        for hexv, labels in candidates.items():
            if len(labels) == 1:
                out[hexv] = labels.pop()
        self.legends[map_title] = out
        return out

    def terrain(self, map_title, sq):
        info = self.book.meta["maps"][map_title]
        ws = self.book.sheet(map_title)
        r, c = R.square_cell(info, sq)
        cell = ws._cells.get((r, c)) if ws is not None else None
        if cell is None or not cell.has_style or not cell.fill or cell.fill.fill_type != "solid":
            return None
        from atlas_model import resolve_color
        return self.legend(map_title).get(resolve_color(cell.fill.fgColor, self.book.theme))

    def _square_for(self, location):
        sheet, _, ref = location.rpartition("!")
        sheet = sheet.strip()
        if sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        info = self.book.meta["maps"].get(sheet)
        if not info:
            return None
        try:
            r = Ref.parse(ref)
        except Exception:
            return None
        sq = R.map_square(info, r.r1, r.c1)
        return (sheet, sq) if sq else None

    def row_label(self, ws, row, skip_col):
        """'RED-m6 · Scout's Transect': the first short texts in the row."""
        parts = []
        for c in range(1, min(ws.max_column, 40) + 1):
            if c == skip_col:
                continue
            cell = ws._cells.get((row, c))
            v = cell.value if cell is not None else None
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                continue
            if isinstance(v, str) and v.strip() and not v.startswith("=") and len(v) <= MAX_LABEL and "\n" not in v:
                parts.append(v.strip())
            if len(parts) == 2:
                break
        return " · ".join(parts) or f"{ws.title} row {row}"

    def card(self, map_title, sq):
        """Everything known about one square."""
        self.refresh()
        info = self.book.meta["maps"][map_title]
        ws = self.book.sheet(map_title)
        r, c = R.square_cell(info, sq)
        mark = ws._cells.get((r, c)) if ws is not None else None
        mark = mark.value if mark is not None and isinstance(mark.value, str) else None
        life = []
        if self.book.ranges is not None:
            for layer, levels in self.book.ranges.who(map_title, sq):
                blocks = {b for b, lv in levels.items() if lv}
                life.append((layer, max(levels.values()), blocks, self.homes.get(layer.name.lower())))
        note = self.book.meta.get("square_notes", {}).get(map_title, {}).get(R.square_label(sq), "")
        return {"square": R.square_label(sq), "mark": mark, "places": self.places.get((map_title, sq), []),
                "life": life, "note": note, "terrain": self.terrain(map_title, sq)}

    def has_content(self, map_title, sq):
        cd = self.card(map_title, sq)
        return bool(cd["places"] or cd["life"] or cd["note"] or cd["mark"] or cd["terrain"])


def card_html(cd, theme, map_title, with_note=True):
    """Rich text for a square card (hover card and Cell panel)."""
    t = theme
    esc = html.escape
    out = [f"<div style='color:{t['group']}; font-weight:bold'>{esc(cd['square'])}"
           + (f" <span style='color:{t['muted']}; font-weight:normal'>· {esc(cd['mark'])}</span>" if cd["mark"] else "")
           + "</div>"]
    if cd.get("terrain"):
        out.append(f"<div style='color:{t['text']}; margin-top:2px'><span style='color:{t['faint']}'>Terrain:</span> "
                   f"{esc(cd['terrain'])}</div>")
    if cd["places"]:
        out.append(f"<div style='color:{t['faint']}; margin-top:6px'>Places</div>")
        for sheet, r, c, label in cd["places"]:
            out.append(f"<div><a style='color:{t['link']}; text-decoration:none' href='{href(sheet, r, c)}'>"
                       f"{esc(label)}</a> <span style='color:{t['faint']}'>· {esc(sheet)}</span></div>")
    if cd["life"]:
        out.append(f"<div style='color:{t['faint']}; margin-top:6px'>Life here</div>")
        for layer, level, blocks, home in cd["life"]:
            when = R.times_label(blocks)
            what = f"{R.LEVELS[level]} · {when.lower() if when in ('Day', 'Night', 'All day') else when}"
            if layer.style != "area":
                what = layer.style
            kind = f"{esc(home[0])}: " if home else ""
            name = (f"<a style='color:#{layer.color}; text-decoration:none' href='{href(*home)}'>{esc(layer.name)}</a>"
                    if home else f"<span style='color:#{layer.color}'>{esc(layer.name)}</span>")
            out.append(f"<div><span style='color:{t['muted']}'>{kind}</span>{name} "
                       f"<span style='color:{t['faint']}'>· {esc(what)}</span></div>")
    if cd["note"] and with_note:
        out.append(f"<div style='color:{t['faint']}; margin-top:6px'>Note</div>")
        out.append(f"<div style='color:{t['text']}'>{esc(cd['note']).replace(chr(10), '<br>')}</div>")
    if not (cd["places"] or cd["life"] or cd["note"]):
        out.append(f"<div style='color:{t['faint']}; margin-top:4px'>Nothing here yet. Name it with a note "
                   f"(right-click the square), link a Gazetteer row's coordinates to it, or paint a layer.</div>")
    return "".join(out)
