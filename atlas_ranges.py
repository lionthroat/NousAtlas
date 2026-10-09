"""Range layers: where each plant or animal lives, painted over a map sheet.

A layer belongs to one creature (or plant). For each map it records, per
time block (the game's six 4-hour ticks: 00, 04, 08, 12, 16, 20), which
squares it's found in and how often: common, uncommon or rare.

On disk this is a plain sheet called "Ranges" (Creature | Map | Times |
Abundance | Squares | Colour), with squares written the way the Gazetteer
writes them ("I6", "G9:K12"), so Excel and scripts can read it too. Atlas
rewrites that sheet on save; in Atlas you paint instead of typing.
"""

import re

from PySide6.QtCore import QPoint, QPointF, QRect, QSize, Qt, Signal
from PySide6.QtGui import QBrush, QColor, QIcon, QPainter, QPen, QPixmap, QPolygon
from PySide6.QtWidgets import (QButtonGroup, QCheckBox, QColorDialog, QComboBox, QFrame,
                               QTableWidget, QTableWidgetItem,
                               QHBoxLayout, QInputDialog, QLabel, QListWidget,
                               QListWidgetItem, QMenu, QMessageBox, QPushButton,
                               QToolButton, QVBoxLayout, QWidget)
from openpyxl.styles import Alignment, Font, PatternFill

from atlas_formula import col_to_num, num_to_col
from atlas_model import Book, RANGES_SHEET
from atlas_theme import LAYER_COLORS, qcolor

BLOCKS = [0, 4, 8, 12, 16, 20]
DAY = {8, 12, 16}            # 08:00-20:00
NIGHT = {20, 0, 4}           # 20:00-08:00
LEVELS = {3: "common", 2: "uncommon", 1: "rare"}
LEVEL_NAMES = {v: k for k, v in LEVELS.items()}
ALPHA = {3: 215, 2: 135, 1: 70}
BRUSH_OFF = 9                # the brush row's Off button (Qt reserves -1)
HEADERS = ["Layer", "Map", "Times", "Abundance", "Squares", "Colour", "Style"]
STYLES = {"area": "Area", "path": "Path", "marker": "Marker"}


def block_label(b):
    return f"{b:02d}:00"


def times_label(blocks):
    blocks = set(blocks)
    if blocks == set(BLOCKS):
        return "All day"
    if blocks == DAY:
        return "Day"
    if blocks == NIGHT:
        return "Night"
    return ", ".join(block_label(b) for b in BLOCKS if b in blocks)


def parse_times(text):
    t = str(text or "").strip().lower()
    if t in ("", "all", "all day", "always"):
        return set(BLOCKS)
    if t == "day":
        return set(DAY)
    if t == "night":
        return set(NIGHT)
    out = set()
    for part in re.split(r"[,;\s]+", t):
        m = re.match(r"(\d{1,2})(?::00)?$", part)
        if m and int(m.group(1)) in BLOCKS:
            out.add(int(m.group(1)))
    return out or set(BLOCKS)


def square_label(sq):
    return f"{num_to_col(sq[0])}{sq[1]}"


_SQ = re.compile(r"([A-Za-z]{1,3})(\d+)")


def parse_squares(text):
    out = set()
    for part in re.split(r"[,;\s]+", str(text or "")):
        if not part:
            continue
        ends = part.split(":")
        a = _SQ.fullmatch(ends[0])
        b = _SQ.fullmatch(ends[-1])
        if not a or not b:
            continue
        x1, y1 = col_to_num(a.group(1)), int(a.group(2))
        x2, y2 = col_to_num(b.group(1)), int(b.group(2))
        for x in range(min(x1, x2), max(x1, x2) + 1):
            for y in range(min(y1, y2), max(y1, y2) + 1):
                out.add((x, y))
    return out


def squares_text(squares):
    """Compress a set of squares into rectangles: 'G9:K12, I13'."""
    left = set(squares)
    parts = []
    for sq in sorted(squares, key=lambda s: (s[1], s[0])):
        if sq not in left:
            continue
        x, y = sq
        x2 = x
        while (x2 + 1, y) in left:
            x2 += 1
        y2 = y
        while all((xx, y2 + 1) in left for xx in range(x, x2 + 1)):
            y2 += 1
        for xx in range(x, x2 + 1):
            for yy in range(y, y2 + 1):
                left.discard((xx, yy))
        parts.append(square_label((x, y)) if (x, y) == (x2, y2)
                     else f"{square_label((x, y))}:{square_label((x2, y2))}")
    return ", ".join(parts)


class Layer:
    def __init__(self, name, color):
        self.name = name
        self.color = color
        self.visible = True
        self.style = "area"      # area (fill), path (a line through squares) or marker (a dot)
        self.maps = {}           # map title -> {block: {(x, y): level}}

    def level(self, map_title, block, sq):
        return self.maps.get(map_title, {}).get(block, {}).get(sq, 0)

    def has_map(self, map_title):
        return any(self.maps.get(map_title, {}).values())


class Ranges:
    def __init__(self):
        self.layers = []

    def get(self, name):
        for layer in self.layers:
            if layer.name.lower() == name.lower():
                return layer
        return None

    def add(self, name, color=None):
        if color is None:
            color = LAYER_COLORS[len(self.layers) % len(LAYER_COLORS)]
        layer = Layer(name, color)
        self.layers.append(layer)
        return layer

    def paint(self, layer, map_title, blocks, squares, level):
        if layer.style != "area":
            blocks, level = BLOCKS, (3 if level else 0)     # paths and markers are just there
        m = layer.maps.setdefault(map_title, {})
        for b in blocks:
            cells = m.setdefault(b, {})
            for sq in squares:
                if level:
                    cells[sq] = level
                else:
                    cells.pop(sq, None)

    def who(self, map_title, sq):
        """[(layer, {block: level})] for one square."""
        out = []
        for layer in self.layers:
            levels = {b: layer.level(map_title, b, sq) for b in BLOCKS}
            if any(levels.values()):
                out.append((layer, levels))
        return out

    def rename_map(self, old, new):
        for layer in self.layers:
            if old in layer.maps:
                layer.maps[new] = layer.maps.pop(old)

    def drop_map(self, title):
        for layer in self.layers:
            layer.maps.pop(title, None)

    # -- the Ranges sheet
    @classmethod
    def load(cls, book):
        r = cls()
        title = book.meta.get("ranges_sheet", RANGES_SHEET)
        ws = book.sheet(title)
        heads = [str(ws.cell(1, i + 1).value or "") for i in range(6)] if ws is not None else []
        if not heads or heads[0] not in ("Layer", "Creature") or heads[1:] != HEADERS[1:6]:
            return r
        book.meta["ranges_sheet"] = title
        for row in ws.iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * 7
            name, map_title, times, level, squares, color, style = row[:7]
            if not name:
                continue
            layer = r.get(str(name)) or r.add(str(name), _hex(color))
            if _hex(color):
                layer.color = _hex(color)
            if str(style or "").strip().lower() in STYLES:
                layer.style = str(style).strip().lower()
            if map_title:
                lv = LEVEL_NAMES.get(str(level or "common").strip().lower(), 3)
                r.paint(layer, str(map_title), parse_times(times), parse_squares(squares), lv)
        return r

    def write(self, book):
        title = book.meta.get("ranges_sheet")
        if title is None:
            if not self.layers:
                return
            title = RANGES_SHEET if book.sheet(RANGES_SHEET) is None else book.unique_title("Ranges (Atlas)")
            book.meta["ranges_sheet"] = title
        ws = book.sheet(title)
        if ws is None:
            ws = book.wb.create_sheet(title)
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)
        head_font = Font(name="Arial", size=10, bold=True, color="FFECEFF4")
        head_fill = PatternFill("solid", fgColor="FF3B4252")
        for i, h in enumerate(HEADERS, 1):
            cell = ws.cell(1, i, h)
            cell.font, cell.fill = head_font, head_fill
        widths = [26, 22, 22, 12, 60, 10, 10]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[num_to_col(i)].width = w
        ws.freeze_panes = "A2"
        row = 2
        for layer in self.layers:
            wrote = False
            for map_title, blocks in layer.maps.items():
                # one row per (abundance, set of times that share exactly the same squares)
                rows_out = []
                for lv in (3, 2, 1):
                    by_squares = {}
                    for b in BLOCKS:
                        sqs = frozenset(sq for sq, v in blocks.get(b, {}).items() if v == lv)
                        if sqs:
                            by_squares.setdefault(sqs, set()).add(b)
                    for sqs, bs in by_squares.items():
                        rows_out.append((lv, bs, sqs))
                rows_out.sort(key=lambda x: (min(BLOCKS.index(b) for b in x[1]) if x[1] != NIGHT else -1, -x[0]))
                for lv, bs, sqs in rows_out:
                    values = [layer.name, map_title, times_label(bs), LEVELS[lv], squares_text(sqs), layer.color,
                              layer.style]
                    for i, v in enumerate(values, 1):
                        ws.cell(row, i, v).alignment = Alignment(vertical="top", wrap_text=(i == 5))
                    ws.cell(row, 6).fill = PatternFill("solid", fgColor="FF" + layer.color)
                    row += 1
                    wrote = True
            if not wrote:
                ws.cell(row, 1, layer.name)
                ws.cell(row, 6, layer.color).fill = PatternFill("solid", fgColor="FF" + layer.color)
                ws.cell(row, 7, layer.style)
                row += 1


def _hex(v):
    s = str(v or "").strip().lstrip("#").upper()
    return s if re.fullmatch(r"[0-9A-F]{6}", s) else None


Book.ranges_writer = staticmethod(lambda book: book.ranges.write(book) if book.ranges is not None else None)


# ---------------------------------------------------------------- maps


def detect_map(ws):
    """Find a lettered row (A, B, C ...) with a numbered column (1, 2, 3 ...)
    down its left side. Returns {"row", "col", "rows", "cols"} for square A1,
    or None."""
    cells = ws._cells
    for (r, c), cell in sorted(cells.items()):
        if r > 40:
            break
        if str(cell.value).strip() != "A":
            continue
        n = 0
        while str((cells.get((r, c + n)) or _Empty).value).strip() == num_to_col(n + 1):
            n += 1
        if n < 4:
            continue
        m = 0
        while True:
            v = (cells.get((r + 1 + m, c - 1)) or _Empty).value
            try:
                ok = float(v) == m + 1
            except (TypeError, ValueError):
                ok = False
            if not ok:
                break
            m += 1
        if m >= 3:
            return {"row": r + 1, "col": c, "rows": m, "cols": n}
    return None


class _Empty:
    value = None


def map_square(info, row, col):
    """Map square (x, y) for a sheet cell, or None if outside the grid."""
    x, y = col - info["col"] + 1, row - info["row"] + 1
    if 1 <= x <= info["cols"] and 1 <= y <= info["rows"]:
        return (x, y)
    return None


def square_cell(info, sq):
    return info["row"] + sq[1] - 1, info["col"] + sq[0] - 1


# ---------------------------------------------------------------- drawing


def night_shade(hex6):
    """The night version of a layer colour: darker and a little bluer."""
    import colorsys
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    r, g, b = colorsys.hsv_to_rgb(h, min(1.0, s * 1.1), v * 0.52)
    nr, ng, nb = 0.20, 0.24, 0.42                     # a night blue to lean toward
    mix = 0.28
    r, g, b = r * (1 - mix) + nr * mix, g * (1 - mix) + ng * mix, b * (1 - mix) + nb * mix
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def view_blocks(view):
    return {"day": DAY, "night": NIGHT, "all": set(BLOCKS)}.get(view, {view} if isinstance(view, int) else set(BLOCKS))


class ViewState:
    @property
    def split(self):
        return self.view == "split"

    @split.setter
    def split(self, on):
        self.view = "split" if on else "all"

    def __init__(self):
        self.blocks = set(BLOCKS)
        self.view = "split"
        self.level = 3
        self.painting = False
        self.layer = None            # name of the layer being painted
        self.solo = None
        self.all_areas = False       # False: only the selected layer's area draws (paths/markers always do)


def path_end(book, layer, map_title, sq, toward, info):
    """How a path ends in a square with one neighbour: 'run' (straight on to
    the far edge) or 'end' (stops in the square). Ends on the edge of the map
    run off it unless set otherwise."""
    label = square_label(sq)
    mode = book.meta.get("path_ends", {}).get(layer.name, {}).get(map_title, {}).get(label)
    if mode:
        return mode
    dx, dy = toward
    nx, ny = sq[0] - dx, sq[1] - dy
    off_map = not (1 <= nx <= info["cols"] and 1 <= ny <= info["rows"])
    return "run" if off_map else "end"


def path_ends_here(book, map_title, sq):
    """[(layer, toward)] for each visible path that ends in this square."""
    out = []
    for layer in book.ranges.layers:
        if layer.style != "path" or not layer.level(map_title, 0, sq):
            continue
        dirs = []
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                if (dx or dy) and layer.level(map_title, 0, (sq[0] + dx, sq[1] + dy)):
                    if dx and dy and (layer.level(map_title, 0, (sq[0] + dx, sq[1]))
                                      or layer.level(map_title, 0, (sq[0], sq[1] + dy))):
                        continue
                    dirs.append((dx, dy))
        if len(dirs) == 1:
            out.append((layer, dirs[0]))
    return out


class Overlay:
    """Paints visible layers onto map squares (called by the grid).
    Area layers fill (side by side when several share a square), paths draw
    a line joining neighbouring squares, markers draw a dot."""

    def __init__(self, window):
        self.window = window

    def line_through(self, ws, row, col):
        """True when a path or marker is drawn on this square (labels get a backing)."""
        book = self.window.book
        if book is None or book.ranges is None or not self.window.layers_showing():
            return False
        info = book.meta["maps"].get(ws.title)
        sq = map_square(info, row, col) if info else None
        if sq is None:
            return False
        st = self.window.range_state
        for layer in book.ranges.layers:
            if layer.style == "area" or (st.solo and layer.name != st.solo) or (not st.solo and not layer.visible):
                continue
            if any(self._view(layer, ws.title, sq)):
                return True
        return False

    def _view(self, layer, map_title, sq):
        """(day level, night level) in split view, else (level, level)."""
        st = self.window.range_state
        if st.view == "split":
            d = max((layer.level(map_title, b, sq) for b in DAY), default=0)
            n = max((layer.level(map_title, b, sq) for b in NIGHT), default=0)
            return d, n
        lv = max((layer.level(map_title, b, sq) for b in view_blocks(st.view)), default=0)
        return lv, lv

    def paint(self, p, ws, row, col, rect):
        book = self.window.book
        if book is None or book.ranges is None:
            return
        info = book.meta["maps"].get(ws.title)
        if not info:
            return
        sq = map_square(info, row, col)
        if sq is None:
            return
        if not self.window.layers_showing():
            return
        st = self.window.range_state
        areas, lines, marks = [], [], []
        for layer in book.ranges.layers:
            if st.solo and layer.name != st.solo:
                continue
            if not st.solo and not layer.visible:
                continue
            if not st.solo and not st.all_areas and layer.style == "area" and layer.name != st.layer:
                continue
            d, n = self._view(layer, ws.title, sq)
            if not (d or n):
                continue
            {"path": lines, "marker": marks}.get(layer.style, areas).append((layer, d, n))
        if not (areas or lines or marks):
            return
        inner = rect.adjusted(0, 0, -1, -1)
        p.save()
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setPen(Qt.NoPen)
        k = len(areas)
        for i, (layer, d, n) in enumerate(areas):
            x1 = inner.left() + inner.width() * i // k
            x2 = inner.left() + inner.width() * (i + 1) // k
            stripe = QRect(x1, inner.top(), x2 - x1, inner.height())
            if st.view != "split":
                night_view = st.view == "night" or (isinstance(st.view, int) and st.view in NIGHT)
                p.fillRect(stripe, qcolor(night_shade(layer.color) if night_view else layer.color, ALPHA[d]))
                continue
            tl, br = stripe.topLeft(), stripe.bottomRight()
            tr, bl = stripe.topRight(), stripe.bottomLeft()
            if d:
                p.setBrush(qcolor(layer.color, ALPHA[d]))
                p.drawPolygon(QPolygon([tl, tr, bl]))
            if n:
                p.setBrush(qcolor(night_shade(layer.color), ALPHA[n]))
                p.drawPolygon(QPolygon([tr, br, bl]))
            if d and n:
                p.setPen(QPen(QColor(20, 20, 26, 110), 1))
                p.drawLine(tr, bl)
                p.setPen(Qt.NoPen)
        c = QPointF(inner.center())
        size = min(inner.width(), inner.height())
        for layer, d, n in lines:
            # a line from the centre toward every neighbour on the same path
            ends, dirs = [], []
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    if (dx, dy) == (0, 0):
                        continue
                    nb = (sq[0] + dx, sq[1] + dy)
                    if not any(self._view(layer, ws.title, nb)):
                        continue
                    if dx and dy:
                        # skip a diagonal when the two squares beside it already join them
                        a = self._view(layer, ws.title, (sq[0] + dx, sq[1]))
                        b = self._view(layer, ws.title, (sq[0], sq[1] + dy))
                        if any(a) or any(b):
                            continue
                    dirs.append((dx, dy))
            if len(dirs) == 1 and path_end(book, layer, ws.title, sq, dirs[0], info) == "run":
                dx, dy = dirs[0]
                dirs.append((-dx, -dy))         # carry on straight through to the far edge
            for dx, dy in dirs:
                ends.append(QPointF(c.x() + dx * inner.width() / 2, c.y() + dy * inner.height() / 2))
            width = max(3.0, size * 0.17)
            for pen_color, extra in ((QColor(20, 20, 26, 110), 2.0), (qcolor(layer.color), 0.0)):
                w2 = width + extra
                p.setPen(QPen(pen_color, w2, Qt.SolidLine, Qt.FlatCap, Qt.RoundJoin))
                for e in ends:
                    p.drawLine(c, e)
                # a round joint (or a rounded stop) at the centre
                p.setPen(Qt.NoPen)
                p.setBrush(pen_color)
                p.drawEllipse(c, w2 / 2, w2 / 2)
                p.setBrush(Qt.NoBrush)
        for i, (layer, d, n) in enumerate(marks):
            level = max(d, n)
            r = size * 0.2
            off = (i - (len(marks) - 1) / 2) * r * 2.2
            centre = QPointF(c.x() + off, c.y())
            p.setPen(QPen(QColor(20, 20, 26, 170), 2))
            p.setBrush(qcolor(layer.color, 255 if level == 3 else ALPHA[level] + 30))
            p.drawEllipse(centre, r, r)
        p.restore()


class Brush:
    """The grid's paint tool when painting ranges."""

    def __init__(self, window):
        self.window = window
        self.stroke = None

    def active(self):
        w = self.window
        st = w.range_state
        return (st.painting and st.layer is not None and w.book is not None
                and w.current_ws() is not None and w.current_ws().title in w.book.meta["maps"])

    def _apply(self, row, col):
        w = self.window
        info = w.book.meta["maps"][w.current_ws().title]
        sq = map_square(info, row, col)
        if sq is None or sq in self.stroke["done"]:
            return
        self.stroke["done"].add(sq)
        layer = w.book.ranges.get(w.range_state.layer)
        if layer is None:
            return
        w.book.ranges.paint(layer, w.current_ws().title, w.range_state.blocks, [sq],
                            0 if self.stroke["erase"] else w.range_state.level)
        w.grid.viewport().update()

    def press(self, row, col, e):
        if not self.active() or e.button() not in (Qt.LeftButton,):
            return False
        info = self.window.book.meta["maps"][self.window.current_ws().title]
        if map_square(info, row, col) is None:
            return False
        if self.stroke is not None:
            self.window.book.done("ranges")
        self.window.book.begin([])
        erase = bool(e.modifiers() & (Qt.ShiftModifier | Qt.AltModifier)) or self.window.range_state.level == 0
        self.stroke = {"done": set(), "erase": erase}
        self._apply(row, col)
        return True

    def move(self, row, col, e):
        if self.stroke is not None:
            self._apply(row, col)

    def release(self, e):
        if self.stroke is not None:
            self.stroke = None
            self.window.book.done("ranges")

    def hover(self, row, col):
        pass

    def stop(self):
        self.window.stop_painting()


# ---------------------------------------------------------------- the panel


def swatch(hex6, size=14):
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setPen(Qt.NoPen)
    p.setBrush(qcolor(hex6))
    p.drawRoundedRect(0, 0, size, size, 3, 3)
    p.end()
    return QIcon(pm)


class RangesPanel(QWidget):
    changed = Signal()

    def __init__(self, window):
        super().__init__()
        self.window = window
        self.st = window.range_state
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(8)

        self.title = QLabel()
        self.title.setObjectName("heading")
        self.title.setWordWrap(True)
        lay.addWidget(self.title)

        self.no_map = QWidget()
        nm = QVBoxLayout(self.no_map)
        nm.setContentsMargins(0, 0, 0, 0)
        self.no_map_text = QLabel()
        self.no_map_text.setObjectName("muted")
        self.no_map_text.setWordWrap(True)
        nm.addWidget(self.no_map_text)
        self.btn_detect = QPushButton("Find the map grid on this sheet")
        self.btn_detect.clicked.connect(self.detect)
        self.btn_use_sel = QPushButton("Use the selected squares as the map")
        self.btn_use_sel.clicked.connect(self.use_selection)
        nm.addWidget(self.btn_detect)
        nm.addWidget(self.btn_use_sel)
        lay.addWidget(self.no_map)

        self.body = QWidget()
        b = QVBoxLayout(self.body)
        b.setContentsMargins(0, 0, 0, 0)
        b.setSpacing(8)

        # Show: what the map displays. Only changes the view.
        self.view_box = QWidget()
        vb = QVBoxLayout(self.view_box)
        vb.setContentsMargins(0, 0, 0, 0)
        vb.setSpacing(4)
        lbl = QLabel("Show")
        lbl.setObjectName("faint")
        vb.addWidget(lbl)
        self.view_group = QButtonGroup(self)
        self.view_btns = {}
        row = QHBoxLayout()
        row.setSpacing(3)
        for key, label, tip in (("split", "Day / night", "Each square: day in the bright top-left half, night in the "
                                                         "dark bottom-right half"),
                                ("day", "Day", "08:00–20:00"), ("night", "Night", "20:00–08:00"),
                                ("all", "All", "Anywhere at any time, one colour")):
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setToolTip(tip)
            self.view_group.addButton(btn)
            self.view_btns[key] = btn
            btn.clicked.connect(lambda _=False, k=key: self.set_view(k))
            row.addWidget(btn)
        row.addStretch()
        vb.addLayout(row)
        row = QHBoxLayout()
        row.setSpacing(3)
        for blk in BLOCKS:
            btn = QPushButton(f"{blk:02d}")
            btn.setCheckable(True)
            btn.setFixedWidth(34)
            btn.setToolTip(f"Show {block_label(blk)}–{block_label((blk + 4) % 24)} only")
            self.view_group.addButton(btn)
            self.view_btns[blk] = btn
            btn.clicked.connect(lambda _=False, k=blk: self.set_view(k))
            row.addWidget(btn)
        row.addStretch()
        vb.addLayout(row)
        b.addWidget(self.view_box)

        self.list = QListWidget()
        self.list.setMinimumHeight(90)
        self.list.setMaximumHeight(170)
        self.list.setToolTip("Anything with a footprint on the map: creatures, plants, characters, hazards, quests.\n"
                             "Tick to show; double-click to show only that one; right-click for more.")
        self.list.itemChanged.connect(self.item_changed)
        self.list.currentItemChanged.connect(self.current_layer_changed)
        self.list.itemDoubleClicked.connect(self.solo_toggle)
        self.list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list.customContextMenuRequested.connect(self.layer_menu)
        b.addWidget(self.list)
        self.empty_note = QLabel("Nothing here yet: click New, or select squares and press Ctrl+L.")
        self.empty_note.setObjectName("muted")
        self.empty_note.setWordWrap(True)
        b.addWidget(self.empty_note)
        row = QHBoxLayout()
        for label, fn in (("New", self.new_layer), ("Rename", self.rename_layer),
                          ("Colour", self.recolor_layer), ("Delete", self.delete_layer)):
            btn = QPushButton(label)
            btn.clicked.connect(fn)
            row.addWidget(btn)
        b.addLayout(row)
        row = QHBoxLayout()
        lbl = QLabel("Draw as")
        lbl.setObjectName("faint")
        row.addWidget(lbl)
        self.style_box = QComboBox()
        for key, label in STYLES.items():
            self.style_box.addItem({"area": "Area (fills squares)", "path": "Path (a line: roads, tracks, rivers)",
                                    "marker": "Marker (a dot: one-off things)"}[key], key)
        self.style_box.activated.connect(self.style_chosen)
        row.addWidget(self.style_box, 1)
        b.addLayout(row)
        self.all_box = QCheckBox("Show every ticked area at once")
        self.all_box.setToolTip("Off: only the selected layer's area is drawn, so the map stays readable.\n"
                                "Paths and markers always draw.")
        self.all_box.toggled.connect(self.all_toggled)
        b.addWidget(self.all_box)
        self.solo_note = QPushButton()
        self.solo_note.setFlat(True)
        self.solo_note.setStyleSheet("text-align: left;")
        self.solo_note.clicked.connect(lambda: self.show_all())
        b.addWidget(self.solo_note)

        # Brush: one row; choosing a brush starts painting, Off stops
        lbl = QLabel("Brush")
        lbl.setObjectName("faint")
        b.addWidget(lbl)
        row = QHBoxLayout()
        row.setSpacing(3)
        self.level_group = QButtonGroup(self)
        for lv, label in ((BRUSH_OFF, "Off"), (3, "Common"), (2, "Uncommon"), (1, "Rare"), (0, "Erase")):
            btn = QPushButton(label)
            btn.setCheckable(True)
            self.level_group.addButton(btn, lv)
            row.addWidget(btn)
        self.level_group.idClicked.connect(self.level_clicked)
        row.addStretch()
        b.addLayout(row)
        # Paint for: which times the brush paints. Only for areas.
        self.time_box = QWidget()
        tb = QVBoxLayout(self.time_box)
        tb.setContentsMargins(0, 0, 0, 0)
        tb.setSpacing(4)
        lbl = QLabel("Paint for")
        lbl.setObjectName("faint")
        tb.addWidget(lbl)
        self.paint_group = QButtonGroup(self)
        self.paint_btns = {}
        row = QHBoxLayout()
        row.setSpacing(3)
        for key, label in (("all", "All day"), ("day", "Day"), ("night", "Night")):
            btn = QPushButton(label)
            btn.setCheckable(True)
            self.paint_group.addButton(btn)
            self.paint_btns[key] = btn
            btn.clicked.connect(lambda _=False, k=key: self.set_blocks({"all": set(BLOCKS), "day": DAY, "night": NIGHT}[k]))
            row.addWidget(btn)
        row.addStretch()
        tb.addLayout(row)
        row = QHBoxLayout()
        row.setSpacing(3)
        self.block_btns = {}
        for blk in BLOCKS:
            btn = QPushButton(f"{blk:02d}")
            btn.setCheckable(True)
            btn.setFixedWidth(34)
            btn.setToolTip(f"Paint {block_label(blk)}–{block_label((blk + 4) % 24)} only")
            self.paint_group.addButton(btn)
            self.block_btns[blk] = btn
            btn.clicked.connect(lambda _=False, k=blk: self.set_blocks({k}))
            row.addWidget(btn)
        row.addStretch()
        tb.addLayout(row)
        b.addWidget(self.time_box)
        # kept for the window's own use: painting on / off
        self.paint_btn = QPushButton()
        self.paint_btn.setCheckable(True)
        self.paint_btn.toggled.connect(self.paint_toggled)
        self.paint_btn.hide()
        b.addWidget(self.paint_btn)

        self.here_title = QLabel()
        self.here_title.setWordWrap(True)
        self.here_title.setTextFormat(Qt.RichText)
        b.addWidget(self.here_title)
        self.here_table = QTableWidget(0, 8)
        self.here_table.setHorizontalHeaderLabels([""] + [f"{blk:02d}" for blk in BLOCKS] + ["All"])
        self.here_table.verticalHeader().hide()
        self.here_table.setShowGrid(True)
        self.here_table.setSelectionMode(QTableWidget.NoSelection)
        self.here_table.setFocusPolicy(Qt.NoFocus)
        self.here_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.here_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.here_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        hh = self.here_table.horizontalHeader()
        for i in range(1, 8):
            self.here_table.setColumnWidth(i, 30 if i < 7 else 34)
        hh.setStretchLastSection(False)
        self.here_table.verticalHeader().setDefaultSectionSize(24)
        self.here_table.cellClicked.connect(self.here_clicked)
        b.addWidget(self.here_table)
        self.here_legend = QLabel()
        self.here_legend.setTextFormat(Qt.RichText)
        self.here_legend.setObjectName("faint")
        b.addWidget(self.here_legend)
        self.here_add = QComboBox()
        self.here_add.activated.connect(self.here_add_chosen)
        b.addWidget(self.here_add)
        self.here_sq = None
        self.here_layers = []
        b.addStretch(1)                  # spare height goes to the bottom, not between sections
        lay.addWidget(self.body, 1)
        lay.addStretch()
        self.refresh()

    # -- state <-> controls
    def refresh(self):
        w = self.window
        ws = w.current_ws()
        book = w.book
        is_map = bool(book is not None and ws is not None and ws.title in book.meta["maps"])
        self.no_map.setVisible(not is_map)
        self.body.setVisible(is_map)
        if ws is None:
            self.title.setText("Layers")
            self.no_map_text.setText("Open a workbook to paint layers.")
            self.btn_detect.hide()
            self.btn_use_sel.hide()
            return
        if not is_map:
            self.title.setText("")
            maps = [t for t in book.meta["maps"] if book.sheet(t) is not None]
            self.no_map_text.setText(f"{ws.title} isn't a map. Layers live on map sheets; "
                                     "if this sheet has a grid, Atlas can use it.")
            self.btn_detect.show()
            self.btn_use_sel.show()
            return
        self.title.setText(ws.title)
        self.sync_times()
        self.all_box.blockSignals(True)
        self.all_box.setChecked(self.st.all_areas)
        self.all_box.blockSignals(False)
        self.sync_brush()
        self.paint_btn.blockSignals(True)
        self.paint_btn.setChecked(self.st.painting)
        self.paint_btn.blockSignals(False)
        self.fill_list()
        self.update_notes()
        self.update_here()
        self.window.update_paint_bar()

    def fill_list(self):
        book = self.window.book
        ws = self.window.current_ws()
        self.list.blockSignals(True)
        self.list.clear()
        current = None
        here = [l for l in book.ranges.layers if l.has_map(ws.title)]
        rest = [l for l in book.ranges.layers if not l.has_map(ws.title)]
        maps = [t for t in book.meta["maps"] if book.sheet(t) is not None]
        for group, layers in (("", here), ("Not on this map", rest)):
            if not layers:
                continue
            if group:
                head = QListWidgetItem(group)
                head.setFlags(Qt.NoItemFlags)
                f = head.font()
                f.setPointSizeF(max(7.0, f.pointSizeF() - 1))
                head.setFont(f)
                self.list.addItem(head)
            for layer in layers:
                text = layer.name
                if group:
                    where = [t for t in maps if layer.has_map(t)]
                    text += f"  · on {', '.join(where)}" if where else "  · not painted yet"
                it = QListWidgetItem(swatch(layer.color), text)
                it.setData(Qt.UserRole, layer.name)
                it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
                it.setCheckState(Qt.Checked if layer.visible else Qt.Unchecked)
                if group:
                    it.setForeground(QColor(self.window.theme["faint"]))
                self.list.addItem(it)
                if layer.name == self.st.layer:
                    current = it
        if current is None:
            for i in range(self.list.count()):
                if self.list.item(i).data(Qt.UserRole):
                    current = self.list.item(i)
                    self.st.layer = current.data(Qt.UserRole)
                    break
        if current is not None:
            self.list.setCurrentItem(current)
        self.list.blockSignals(False)
        self.sync_style_box()

    def timeless_selected(self):
        layer = self.window.book.ranges.get(self.st.layer) if (self.window.book and self.st.layer) else None
        return layer is not None and layer.style != "area"

    def sync_brush(self):
        """The brush row: Off / Common / Uncommon / Rare / Erase, or Off / Draw / Erase
        for paths and markers."""
        st = self.st
        timeless = self.timeless_selected()
        if timeless and st.level in (1, 2):
            st.level = 3
        self.level_group.button(3).setText("Draw" if timeless else "Common")
        for lv in (2, 1):
            self.level_group.button(lv).setVisible(not timeless)
        self.level_group.button(BRUSH_OFF if not st.painting else st.level).setChecked(True)

    def update_notes(self):
        st = self.st
        timeless = self.timeless_selected()
        self.time_box.setVisible(not timeless)
        self.sync_brush()
        self.solo_note.setText(f"Only {st.solo} is showing · show all")
        self.solo_note.setVisible(bool(st.solo))
        has = bool(self.window.book is not None and self.window.book.ranges.layers)
        self.empty_note.setVisible(not has)

    def show_all(self):
        self.st.solo = None
        self.update_notes()
        self.window.grid.viewport().update()

    def update_here(self):
        """The square editor: every layer on the selected square, by time.
        Click a slot to cycle common > uncommon > rare > none; the All column
        sets the whole row."""
        w = self.window
        ws = w.current_ws()
        if ws is None or w.book is None or ws.title not in w.book.meta["maps"]:
            return
        info = w.book.meta["maps"][ws.title]
        sq = map_square(info, *w.grid.cur)
        self.here_sq = (ws.title, sq) if sq else None
        t = w.theme
        tbl = self.here_table
        tbl.blockSignals(True)
        if sq is None:
            self.here_title.setText("<span style='color:gray'>Select a square on the map to edit what's in it.</span>")
            tbl.setRowCount(0)
            tbl.hide()
            self.here_add.hide()
            self.here_legend.hide()
            tbl.blockSignals(False)
            return
        rows = w.book.ranges.who(ws.title, sq)
        self.here_title.setText(f"<span style='color:{t['group']}'><b>Square {square_label(sq)}</b></span>"
                                + (f" <span style='color:{t['muted']}'>· click a slot to change it</span>" if rows else
                                   f" <span style='color:{t['muted']}'>· nothing here yet</span>"))
        tbl.setRowCount(len(rows))
        self.here_layers = [layer.name for layer, _ in rows]
        for i, (layer, levels) in enumerate(rows):
            name = QTableWidgetItem(swatch(layer.color), layer.name + ("" if layer.style == "area" else f" ({layer.style})"))
            name.setFlags(Qt.ItemIsEnabled)
            name.setToolTip(layer.name)
            tbl.setItem(i, 0, name)
            for j, b in enumerate(BLOCKS):
                lv = levels[b]
                # filled like the map: solid common, medium uncommon, faint rare, empty none
                it = QTableWidgetItem("" if lv else "–")
                it.setTextAlignment(Qt.AlignCenter)
                it.setForeground(QColor(t["faint"]))
                if lv:
                    tone = night_shade(layer.color) if (b in NIGHT and layer.style == "area") else layer.color
                    it.setBackground(qcolor(tone, ALPHA[lv]))
                it.setFlags(Qt.ItemIsEnabled)
                it.setToolTip(f"{layer.name} at {block_label(b)}: {LEVELS.get(lv, 'not here')}. Click to change.")
                tbl.setItem(i, 1 + j, it)
            if layer.style != "area":
                for j in range(6):
                    it = tbl.item(i, 1 + j)
                    it.setText("")
                    it.setToolTip(f"{layer.name} is a {layer.style}: it's simply here. Click to remove it.")
            top = max(levels.values())
            all_it = QTableWidgetItem("all" if layer.style == "area" else "on")
            all_it.setTextAlignment(Qt.AlignCenter)
            all_it.setForeground(QColor(t["muted"]))
            all_it.setFlags(Qt.ItemIsEnabled)
            all_it.setToolTip("Set every time at once (cycles common > uncommon > rare > none)")
            tbl.setItem(i, 7, all_it)
        tbl.resizeColumnToContents(0)
        tbl.setColumnWidth(0, min(120, max(80, tbl.columnWidth(0))))
        tbl.setFixedHeight(tbl.horizontalHeader().height() + sum(tbl.rowHeight(i) for i in range(len(rows))) + 4)
        tbl.setVisible(bool(rows))
        tbl.setStyleSheet(f"QTableWidget {{ gridline-color: {t['panel']}; background: {t['window']}; }}")
        def chip(alpha):
            c = qcolor(t["accent"].lstrip("#"), alpha)
            return (f"<span style='background-color: rgba({c.red()},{c.green()},{c.blue()},{alpha / 255:.2f})'>"
                    "&nbsp;&nbsp;&nbsp;&nbsp;</span>")
        self.here_legend.setText(f"{chip(ALPHA[3])} common &nbsp; {chip(ALPHA[2])} uncommon &nbsp; "
                                 f"{chip(ALPHA[1])} rare &nbsp; – not here")
        self.here_legend.setVisible(any(layer.style == "area" for layer, _ in rows))
        # add another layer to this square
        others = [l.name for l in w.book.ranges.layers if l.name not in self.here_layers]
        self.here_add.blockSignals(True)
        self.here_add.clear()
        self.here_add.addItem("Add a layer to this square…", None)
        for n in others:
            layer = w.book.ranges.get(n)
            self.here_add.addItem(swatch(layer.color), n, n)
        self.here_add.addItem("New layer…", "__new__")
        self.here_add.blockSignals(False)
        self.here_add.show()
        tbl.blockSignals(False)

    def here_clicked(self, row, col):
        if self.here_sq is None or col == 0 or row >= len(self.here_layers):
            return
        w = self.window
        map_title, sq = self.here_sq
        layer = w.book.ranges.get(self.here_layers[row])
        if layer is None:
            return
        nxt = {3: 2, 2: 1, 1: 0, 0: 3}
        if layer.style != "area":
            on = any(layer.level(map_title, b, sq) for b in BLOCKS)
            blocks, level = set(BLOCKS), 0 if on else 3
        elif col == 7:
            top = max(layer.level(map_title, b, sq) for b in BLOCKS)
            blocks, level = set(BLOCKS), nxt[top] if top else 3
        else:
            b = BLOCKS[col - 1]
            blocks, level = {b}, nxt[layer.level(map_title, b, sq)]
        w.book.begin([])
        w.book.ranges.paint(layer, map_title, blocks, [sq], level)
        w.book.done("ranges")

    def here_add_chosen(self, idx):
        data = self.here_add.itemData(idx)
        if not data or self.here_sq is None:
            return
        w = self.window
        map_title, sq = self.here_sq
        if data == "__new__":
            layer = self.new_layer()
            if layer is None:
                self.update_here()
                return
        else:
            layer = w.book.ranges.get(data)
        w.book.begin([])
        w.book.ranges.paint(layer, map_title, self.st.blocks, [sq], self.st.level or 3)
        w.book.done("ranges")

    # -- handlers
    def sync_times(self):
        st = self.st
        btn = self.view_btns.get(st.view)
        if btn is not None:
            btn.setChecked(True)
        key = {frozenset(BLOCKS): "all", frozenset(DAY): "day", frozenset(NIGHT): "night"}.get(frozenset(st.blocks))
        if key:
            self.paint_btns[key].setChecked(True)
        elif len(st.blocks) == 1:
            self.block_btns[next(iter(st.blocks))].setChecked(True)
        else:
            self.paint_group.setExclusive(False)
            for b, bb in self.block_btns.items():
                bb.setChecked(b in st.blocks)
            for bb in self.paint_btns.values():
                bb.setChecked(False)
            self.paint_group.setExclusive(True)

    def set_view(self, view):
        self.st.view = view
        self.sync_times()
        self.window.grid.viewport().update()

    def blocks_clicked(self):
        chosen = {b for b, btn in self.block_btns.items() if btn.isChecked()}
        if not chosen:
            chosen = {next(b for b, btn in self.block_btns.items() if btn is self.sender())}
        self.set_blocks(chosen, keep_split=True)

    def set_blocks(self, blocks, keep_split=False):
        """Which times the brush paints. Doesn't touch what's shown."""
        self.st.blocks = set(blocks)
        self.sync_times()
        self.update_notes()
        self.window.update_paint_bar()
        self.window.grid.viewport().update()

    def split_toggled(self, on):
        self.st.split = on
        self.window.grid.viewport().update()

    def item_changed(self, it):
        layer = self.window.book.ranges.get(it.data(Qt.UserRole))
        if layer is not None:
            layer.visible = it.checkState() == Qt.Checked
            self.window.grid.viewport().update()

    def all_toggled(self, on):
        self.st.all_areas = on
        self.window.grid.viewport().update()

    def style_chosen(self, idx):
        layer = self._selected_layer()
        style = self.style_box.itemData(idx)
        if layer is None or style == layer.style:
            return
        self.window.book.begin([])
        live = self.window.book.ranges.get(layer.name)
        live.style = style
        if style != "area":
            # it's simply there now: wherever it was, at any time, it is all day
            for m in live.maps.values():
                where = set()
                for cells in m.values():
                    where |= {sq for sq, lv in cells.items() if lv}
                for b in BLOCKS:
                    m[b] = {sq: 3 for sq in where}
        self.window.book.done("ranges")

    def sync_style_box(self):
        layer = self._selected_layer() if self.list.currentItem() is not None else None
        self.style_box.setEnabled(layer is not None)
        if layer is not None:
            self.style_box.setCurrentIndex(self.style_box.findData(layer.style))

    def current_layer_changed(self, cur, prev):
        if cur is not None:
            self.sync_style_box()
            self.st.layer = cur.data(Qt.UserRole)
            self.update_notes()
            self.window.grid.viewport().update()
            self.window.update_paint_bar()

    def solo_toggle(self, it):
        name = it.data(Qt.UserRole)
        self.st.solo = None if self.st.solo == name else name
        self.update_notes()
        self.window.grid.viewport().update()

    def paint_toggled(self, on):
        self.st.painting = on
        if on and self.st.layer is None:
            self.new_layer()
        self.update_notes()
        self.window.update_paint_bar()

    def level_clicked(self, lv):
        if lv == BRUSH_OFF:
            self.paint_btn.setChecked(False)
        else:
            self.st.level = lv
            if not self.st.painting:
                self.paint_btn.setChecked(True)
        self.update_notes()
        self.window.update_paint_bar()

    def layer_menu(self, pos):
        it = self.list.itemAt(pos)
        menu = QMenu(self)
        if it is not None:
            name = it.data(Qt.UserRole)
            menu.addAction(f"Show only {name}", lambda: self.solo_toggle(it))
            menu.addAction(f"Find “{name}” in the sheets", lambda: self.window.find_text(name))
            menu.addSeparator()
            menu.addAction("Rename…", self.rename_layer)
            menu.addAction("Colour…", self.recolor_layer)
            menu.addAction("Delete…", self.delete_layer)
        menu.addAction("New layer…", self.new_layer)
        menu.exec(self.list.mapToGlobal(pos))

    def _selected_layer(self):
        it = self.list.currentItem()
        return self.window.book.ranges.get(it.data(Qt.UserRole)) if it is not None else None

    def new_layer(self, name=None):
        book = self.window.book
        if name is None:
            name, ok = QInputDialog.getText(self, "New layer", "What's on the map? A creature, plant, character, "
                                                          "hazard, quest… (a name that matches a row elsewhere links to it)")
            if not ok:
                if self.st.painting and self.st.layer is None:
                    self.paint_btn.setChecked(False)
                return None
        name = name.strip()
        if not name:
            return None
        existing = book.ranges.get(name)
        if existing is not None:
            self.st.layer = existing.name
            self.refresh()
            return existing
        book.begin([])
        layer = book.ranges.add(name)
        book.done("ranges")
        self.st.layer = layer.name
        self.refresh()
        return layer

    def rename_layer(self):
        layer = self._selected_layer()
        if layer is None:
            return
        name, ok = QInputDialog.getText(self, "Rename layer", "New name:", text=layer.name)
        name = name.strip()
        if not ok or not name or name == layer.name:
            return
        if self.window.book.ranges.get(name) is not None and name.lower() != layer.name.lower():
            QMessageBox.information(self, "Rename layer", f"There's already a layer called {name}.")
            return
        self.window.book.begin([])
        old = layer.name
        self.window.book.ranges.get(old).name = name
        ends = self.window.book.meta.get("path_ends", {})
        if old in ends:
            ends[name] = ends.pop(old)
        if self.st.layer == old:
            self.st.layer = name
        if self.st.solo == old:
            self.st.solo = name
        self.window.book.done("ranges")

    def recolor_layer(self):
        layer = self._selected_layer()
        if layer is None:
            return
        hex6 = self.window.pick_color(layer.color, f"Colour for {layer.name}")
        if hex6 is None:
            return
        self.window.book.begin([])
        self.window.book.ranges.get(layer.name).color = hex6
        self.window.book.done("ranges")

    def delete_layer(self):
        layer = self._selected_layer()
        if layer is None:
            return
        book = self.window.book
        here = self.window.current_ws().title
        maps = [t for t in book.meta["maps"] if book.sheet(t) is not None]
        counts = {t: len({sq for cells in layer.maps.get(t, {}).values() for sq, lv in cells.items() if lv})
                  for t in maps}
        painted = [t for t in maps if counts[t]]
        if painted:
            box = QMessageBox(self)
            box.setWindowTitle(layer.name)
            where = "; ".join(f"{t} ({counts[t]} square{'s' if counts[t] != 1 else ''})" for t in painted)
            box.setText(f"{layer.name} is painted on: {where}.")
            box.setInformativeText("Ctrl+Z undoes either choice.")
            only_here = box.addButton(f"Remove from {here}", QMessageBox.AcceptRole) if here in painted and len(painted) > 1 else None
            everywhere = box.addButton("Delete the layer" if painted == [here] or len(painted) == 1 else
                                       "Delete it everywhere", QMessageBox.DestructiveRole)
            box.addButton(QMessageBox.Cancel)
            box.exec()
            choice = box.clickedButton()
            if choice is only_here and only_here is not None:
                book.begin([])
                book.ranges.get(layer.name).maps.pop(here, None)
                book.done("ranges")
                return
            if choice is not everywhere:
                return
        book.begin([])
        book.ranges.layers = [l for l in book.ranges.layers if l.name != layer.name]
        self.window.book.meta.get("path_ends", {}).pop(layer.name, None)
        if self.st.layer == layer.name:
            self.st.layer = None
        if self.st.solo == layer.name:
            self.st.solo = None
        self.window.book.done("ranges")

    def detect(self):
        w = self.window
        ws = w.current_ws()
        info = detect_map(ws)
        if info is None:
            QMessageBox.information(self, "Find the map grid",
                                    "I couldn't find a lettered row (A, B, C…) with numbers (1, 2, 3…) down "
                                    "its left side. Select the squares of the map and use the other button.")
            return
        w.book.begin([])
        w.book.meta["maps"][ws.title] = info
        w.book.done("maps")

    def use_selection(self):
        w = self.window
        r1, c1, r2, c2 = w.grid.selection()
        if (r1, c1) == (r2, c2):
            QMessageBox.information(self, "Use selection as map",
                                    "Select all the squares of the map first (the top-left one becomes A1).")
            return
        w.book.begin([])
        w.book.meta["maps"][w.current_ws().title] = {"row": r1, "col": c1, "rows": r2 - r1 + 1, "cols": c2 - c1 + 1}
        w.book.done("maps")


def _esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
