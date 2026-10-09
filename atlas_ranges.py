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

from PySide6.QtCore import QPoint, QRect, QSize, Qt, Signal
from PySide6.QtGui import QBrush, QColor, QIcon, QPainter, QPen, QPixmap, QPolygon
from PySide6.QtWidgets import (QButtonGroup, QCheckBox, QColorDialog, QFrame,
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
HEADERS = ["Creature", "Map", "Times", "Abundance", "Squares", "Colour"]


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
        if ws is None or [str(ws.cell(1, i + 1).value or "") for i in range(len(HEADERS))] != HEADERS:
            return r
        book.meta["ranges_sheet"] = title
        for row in ws.iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * 6
            name, map_title, times, level, squares, color = row[:6]
            if not name:
                continue
            layer = r.get(str(name)) or r.add(str(name), _hex(color))
            if _hex(color):
                layer.color = _hex(color)
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
        widths = [26, 22, 22, 12, 60, 10]
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
                    values = [layer.name, map_title, times_label(bs), LEVELS[lv], squares_text(sqs), layer.color]
                    for i, v in enumerate(values, 1):
                        ws.cell(row, i, v).alignment = Alignment(vertical="top", wrap_text=(i == 5))
                    ws.cell(row, 6).fill = PatternFill("solid", fgColor="FF" + layer.color)
                    row += 1
                    wrote = True
            if not wrote:
                ws.cell(row, 1, layer.name)
                ws.cell(row, 6, layer.color).fill = PatternFill("solid", fgColor="FF" + layer.color)
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


class ViewState:
    def __init__(self):
        self.blocks = set(BLOCKS)
        self.split = True
        self.level = 3
        self.painting = False
        self.layer = None            # name of the layer being painted
        self.solo = None


class Overlay:
    """Paints visible layers onto map squares (called by the grid)."""

    def __init__(self, window):
        self.window = window

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
        entries = []
        for layer in book.ranges.layers:
            if st.solo and layer.name != st.solo:
                continue
            if not st.solo and not layer.visible:
                continue
            if st.split:
                d = max((layer.level(ws.title, b, sq) for b in DAY), default=0)
                n = max((layer.level(ws.title, b, sq) for b in NIGHT), default=0)
                if d or n:
                    entries.append((layer, d, n))
            else:
                lv = max((layer.level(ws.title, b, sq) for b in st.blocks), default=0)
                if lv:
                    entries.append((layer, lv, lv))
        if not entries:
            return
        k = len(entries)
        inner = rect.adjusted(0, 0, -1, -1)
        p.save()
        p.setPen(Qt.NoPen)
        for i, (layer, d, n) in enumerate(entries):
            x1 = inner.left() + inner.width() * i // k
            x2 = inner.left() + inner.width() * (i + 1) // k
            stripe = QRect(x1, inner.top(), x2 - x1, inner.height())
            if d == n:
                p.fillRect(stripe, qcolor(layer.color, ALPHA[d]))
                continue
            tl, br = stripe.topLeft(), stripe.bottomRight()
            tr, bl = stripe.topRight(), stripe.bottomLeft()
            if d:
                p.setBrush(qcolor(layer.color, ALPHA[d]))
                p.drawPolygon(QPolygon([tl, tr, bl]))
            if n:
                p.setBrush(qcolor(layer.color, ALPHA[n]))
                p.drawPolygon(QPolygon([tr, br, bl]))
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
        about = QLabel("Anything with a footprint on the map: creatures, plants, characters, hazards, quests. "
                       "Layers only show while this tab is open.")
        about.setObjectName("muted")
        about.setWordWrap(True)
        lay.addWidget(about)

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

        lbl = QLabel("Time")
        lbl.setObjectName("faint")
        b.addWidget(lbl)
        row = QHBoxLayout()
        row.setSpacing(3)
        self.block_btns = {}
        for blk in BLOCKS:
            btn = QPushButton(f"{blk:02d}")
            btn.setCheckable(True)
            btn.setFixedWidth(34)
            btn.setToolTip(f"{block_label(blk)}–{block_label((blk + 4) % 24)}")
            btn.clicked.connect(self.blocks_clicked)
            self.block_btns[blk] = btn
            row.addWidget(btn)
        row.addStretch()
        b.addLayout(row)
        row = QHBoxLayout()
        row.setSpacing(3)
        for label, blocks in (("Day", DAY), ("Night", NIGHT), ("All", set(BLOCKS))):
            btn = QPushButton(label)
            btn.clicked.connect(lambda _=False, bl=blocks: self.set_blocks(bl))
            row.addWidget(btn)
        self.split_box = QCheckBox("Split day / night")
        self.split_box.setToolTip("Each square shows day in its top-left half and night in its bottom-right half.\n"
                                  "Day = 08:00–20:00, night = 20:00–08:00.")
        self.split_box.toggled.connect(self.split_toggled)
        row.addWidget(self.split_box)
        row.addStretch()
        b.addLayout(row)

        lbl = QLabel("Layers")
        lbl.setObjectName("faint")
        b.addWidget(lbl)
        self.list = QListWidget()
        self.list.setMinimumHeight(120)
        self.list.itemChanged.connect(self.item_changed)
        self.list.currentItemChanged.connect(self.current_layer_changed)
        self.list.itemDoubleClicked.connect(self.solo_toggle)
        self.list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list.customContextMenuRequested.connect(self.layer_menu)
        b.addWidget(self.list, 1)
        row = QHBoxLayout()
        for label, fn in (("New", self.new_layer), ("Rename", self.rename_layer),
                          ("Colour", self.recolor_layer), ("Delete", self.delete_layer)):
            btn = QPushButton(label)
            btn.clicked.connect(fn)
            row.addWidget(btn)
        b.addLayout(row)
        self.solo_note = QLabel()
        self.solo_note.setObjectName("muted")
        b.addWidget(self.solo_note)

        lbl = QLabel("Paint")
        lbl.setObjectName("faint")
        b.addWidget(lbl)
        row = QHBoxLayout()
        row.setSpacing(3)
        self.paint_btn = QPushButton("Paint")
        self.paint_btn.setCheckable(True)
        self.paint_btn.setToolTip("Click or drag over squares. Shift-drag (or the Erase brush) removes.")
        self.paint_btn.toggled.connect(self.paint_toggled)
        row.addWidget(self.paint_btn)
        self.level_group = QButtonGroup(self)
        for lv, label in ((3, "Common"), (2, "Uncommon"), (1, "Rare"), (0, "Erase")):
            btn = QPushButton(label)
            btn.setCheckable(True)
            self.level_group.addButton(btn, lv)
            row.addWidget(btn)
        self.level_group.idClicked.connect(self.level_clicked)
        row.addStretch()
        b.addLayout(row)
        self.paint_note = QLabel()
        self.paint_note.setObjectName("muted")
        self.paint_note.setWordWrap(True)
        b.addWidget(self.paint_note)

        lbl = QLabel("Here")
        lbl.setObjectName("faint")
        b.addWidget(lbl)
        self.here = QLabel()
        self.here.setWordWrap(True)
        self.here.setTextFormat(Qt.RichText)
        self.here.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        b.addWidget(self.here)
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
            self.title.setText("Layers")
            maps = [t for t in book.meta["maps"] if book.sheet(t) is not None]
            text = f"“{ws.title}” isn't a map yet."
            if maps:
                text += " Maps in this workbook: " + ", ".join(maps) + "."
            self.no_map_text.setText(text)
            self.btn_detect.show()
            self.btn_use_sel.show()
            return
        self.title.setText(f"Layers · {ws.title}")
        for blk, btn in self.block_btns.items():
            btn.setChecked(blk in self.st.blocks)
            btn.setEnabled(True)
        self.split_box.blockSignals(True)
        self.split_box.setChecked(self.st.split)
        self.split_box.blockSignals(False)
        self.level_group.button(self.st.level).setChecked(True)
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
        for layer in book.ranges.layers:
            text = layer.name if layer.has_map(ws.title) else f"{layer.name}  (not painted here)"
            it = QListWidgetItem(swatch(layer.color), text)
            it.setData(Qt.UserRole, layer.name)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if layer.visible else Qt.Unchecked)
            self.list.addItem(it)
            if layer.name == self.st.layer:
                current = it
        if current is None and self.list.count():
            current = self.list.item(0)
            self.st.layer = current.data(Qt.UserRole)
        if current is not None:
            self.list.setCurrentItem(current)
        self.list.blockSignals(False)

    def update_notes(self):
        st = self.st
        self.solo_note.setText(f"Showing only {st.solo}. Double-click it again to show all."
                               if st.solo else "Double-click a layer to show only that one.")
        if st.layer is None:
            self.paint_note.setText("Make a layer first.")
        elif st.painting:
            what = "Erasing" if st.level == 0 else f"Painting {LEVELS[st.level]}"
            self.paint_note.setText(f"{what} {st.layer} at {times_label(st.blocks)}. "
                                    "Click or drag over squares; Shift-drag erases.")
        else:
            self.paint_note.setText("Turn on Paint, then click or drag over squares.")

    def update_here(self):
        w = self.window
        ws = w.current_ws()
        if ws is None or w.book is None or ws.title not in w.book.meta["maps"]:
            return
        info = w.book.meta["maps"][ws.title]
        sq = map_square(info, *w.grid.cur)
        if sq is None:
            self.here.setText("<span style='color:gray'>Select a map square to see who's there.</span>")
            return
        rows = w.book.ranges.who(ws.title, sq)
        t = w.theme
        head = "".join(f"<td align='center' style='color:{t['faint']}; padding:0 3px'>{b:02d}</td>" for b in BLOCKS)
        html = [f"<b>{square_label(sq)}</b>"]
        if not rows:
            html.append(f"<br><span style='color:{t['muted']}'>Nothing's painted here.</span>")
        else:
            html.append(f"<table cellspacing='0' cellpadding='1'><tr><td></td>{head}</tr>")
            for layer, levels in rows:
                cells = ""
                for b in BLOCKS:
                    lv = levels[b]
                    mark = {3: "●", 2: "◐", 1: "·"}.get(lv, "")
                    cells += f"<td align='center' style='color:#{layer.color}; padding:0 3px'>{mark}</td>"
                html.append(f"<tr><td style='color:#{layer.color}; padding-right:6px'>{_esc(layer.name)}</td>{cells}</tr>")
            html.append(f"</table><span style='color:{t['faint']}'>● common  ◐ uncommon  · rare</span>")
        self.here.setText("".join(html))

    # -- handlers
    def blocks_clicked(self):
        chosen = {b for b, btn in self.block_btns.items() if btn.isChecked()}
        if not chosen:
            chosen = {next(b for b, btn in self.block_btns.items() if btn is self.sender())}
        self.set_blocks(chosen, keep_split=True)

    def set_blocks(self, blocks, keep_split=False):
        self.st.blocks = set(blocks)
        if set(blocks) != set(BLOCKS):
            self.st.split = False      # splitting only makes sense when looking at the whole day
        self.refresh()
        self.window.grid.viewport().update()

    def split_toggled(self, on):
        self.st.split = on
        self.window.grid.viewport().update()

    def item_changed(self, it):
        layer = self.window.book.ranges.get(it.data(Qt.UserRole))
        if layer is not None:
            layer.visible = it.checkState() == Qt.Checked
            self.window.grid.viewport().update()

    def current_layer_changed(self, cur, prev):
        if cur is not None:
            self.st.layer = cur.data(Qt.UserRole)
            self.update_notes()
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
        if QMessageBox.question(self, "Delete layer",
                                f"Delete the {layer.name} layer and everything painted for it, on every map?\n"
                                "You can undo this.") != QMessageBox.Yes:
            return
        self.window.book.begin([])
        self.window.book.ranges.layers = [l for l in self.window.book.ranges.layers if l.name != layer.name]
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
