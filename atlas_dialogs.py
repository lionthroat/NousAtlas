"""Dialogs: the colour picker (your palettes only, never Excel's), the
palette editor, colour remapping, and style suggestions."""

import copy

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtGui import QColor, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (QColorDialog, QComboBox, QDialog,
                               QDialogButtonBox, QGridLayout, QHBoxLayout,
                               QInputDialog, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMessageBox, QPushButton,
                               QScrollArea, QTableWidget, QTableWidgetItem,
                               QToolButton, QVBoxLayout, QWidget, QCheckBox,
                               QAbstractItemView, QHeaderView)

from atlas_theme import PALETTES, contrast, qcolor


def swatch_icon(hex6, w=16, h=16, radius=3):
    pm = QPixmap(w, h)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    if hex6 is None:
        p.setPen(QColor("#888"))
        p.drawRoundedRect(0, 0, w - 1, h - 1, radius, radius)
        p.setPen(QColor("#bf616a"))
        p.drawLine(2, h - 3, w - 3, 2)
    else:
        p.setPen(Qt.NoPen)
        p.setBrush(qcolor(hex6))
        p.drawRoundedRect(0, 0, w, h, radius, radius)
    p.end()
    return QIcon(pm)


def nearest(hex6, choices):
    """Closest colour by a weighted RGB distance."""
    def rgb(h):
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    r1, g1, b1 = rgb(hex6)
    best, best_d = None, None
    for c in choices:
        r2, g2, b2 = rgb(c)
        rm = (r1 + r2) / 2
        d = (2 + rm / 256) * (r1 - r2) ** 2 + 4 * (g1 - g2) ** 2 + (2 + (255 - rm) / 256) * (b1 - b2) ** 2
        if best_d is None or d < best_d:
            best, best_d = c, d
    return best


class ColorPicker(QDialog):
    """Pick a colour from the workbook palette, the built-in palettes, or the
    colours already in the workbook. Result: 'RRGGBB', '' for none, or None."""

    def __init__(self, parent, book, current=None, title="Colour", allow_none=True):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.result_hex = None
        self.book = book
        lay = QVBoxLayout(self)
        lay.setSpacing(6)
        palette = (book.meta.get("palette") or {}).get("colors") or []
        if palette:
            self._section(lay, (book.meta["palette"].get("name") or "Workbook palette"),
                          [(c["name"], c["hex"]) for c in palette], current)
        for name, colors in PALETTES.items():
            if palette and name == (book.meta.get("palette") or {}).get("name"):
                continue
            self._section(lay, name, colors, current)
        used = sorted(book.colors_in_use().items(), key=lambda kv: -kv[1])[:32]
        if used:
            self._section(lay, "In this workbook", [(f"{h} · used {n}×", h) for h, n in used], current)
        row = QHBoxLayout()
        if allow_none:
            none = QPushButton(swatch_icon(None), "No colour")
            none.clicked.connect(lambda: self._done(""))
            row.addWidget(none)
        self.hex = QLineEdit(current or "")
        self.hex.setPlaceholderText("RRGGBB")
        self.hex.setMaxLength(7)
        self.hex.setFixedWidth(90)
        self.hex.returnPressed.connect(self._hex_entered)
        row.addWidget(QLabel("Hex"))
        row.addWidget(self.hex)
        more = QPushButton("Mix a colour…")
        more.clicked.connect(self._more)
        row.addWidget(more)
        row.addStretch()
        lay.addLayout(row)

    def _section(self, lay, title, colors, current):
        lbl = QLabel(title)
        lbl.setObjectName("faint")
        lay.addWidget(lbl)
        grid = QGridLayout()
        grid.setSpacing(3)
        for i, (name, hex6) in enumerate(colors):
            b = QToolButton()
            b.setIcon(swatch_icon(hex6, 22, 22, 4))
            b.setIconSize(QSize(22, 22))
            b.setToolTip(f"{name}\n#{hex6}")
            b.setAutoRaise(True)
            if current and current.upper() == hex6.upper():
                b.setStyleSheet("QToolButton { border: 2px solid palette(highlight); }")
            b.clicked.connect(lambda _=False, h=hex6: self._done(h))
            grid.addWidget(b, i // 16, i % 16)
        lay.addLayout(grid)

    def _hex_entered(self):
        s = self.hex.text().strip().lstrip("#").upper()
        if len(s) == 6 and all(ch in "0123456789ABCDEF" for ch in s):
            self._done(s)

    def _more(self):
        dlg = QColorDialog(self)
        palette = (self.book.meta.get("palette") or {}).get("colors") or []
        source = [c["hex"] for c in palette] or [h for _, h in PALETTES["Nord"]]
        for i in range(min(16, len(source))):
            QColorDialog.setCustomColor(i, qcolor(source[i]))
        for i in range(48):
            src = [h for p in PALETTES.values() for _, h in p]
            QColorDialog.setStandardColor(i, qcolor(src[i % len(src)]))
        s = self.hex.text().strip().lstrip("#")
        if len(s) == 6:
            dlg.setCurrentColor(qcolor(s))
        if dlg.exec():
            self._done(dlg.currentColor().name()[1:].upper())

    def _done(self, hex6):
        self.result_hex = hex6
        self.accept()

    @classmethod
    def get(cls, parent, book, current=None, title="Colour", allow_none=True):
        dlg = cls(parent, book, current, title, allow_none)
        dlg.exec()
        return dlg.result_hex


class PaletteDialog(QDialog):
    """Edit the workbook's named palette. Changing a colour can also change
    it everywhere it's used."""

    def __init__(self, window):
        super().__init__(window)
        self.window = window
        self.book = window.book
        self.setWindowTitle("Workbook palette")
        self.resize(520, 520)
        pal = copy.deepcopy(self.book.meta.get("palette") or {"name": "My palette", "colors": []})
        self.pal = pal
        self.replacements = {}
        lay = QVBoxLayout(self)
        row = QHBoxLayout()
        row.addWidget(QLabel("Name"))
        self.name = QLineEdit(pal.get("name") or "My palette")
        row.addWidget(self.name)
        lay.addLayout(row)
        hint = QLabel("These are the colours the colour buttons offer first. Styles and range layers pick from them too.")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        lay.addWidget(hint)
        self.list = QListWidget()
        self.list.setDragDropMode(QAbstractItemView.InternalMove)
        self.list.itemDoubleClicked.connect(lambda it: self.edit_color())
        lay.addWidget(self.list, 1)
        row = QHBoxLayout()
        for label, fn in (("Add", self.add), ("Change colour", self.edit_color), ("Rename", self.rename),
                          ("Remove", self.remove)):
            b = QPushButton(label)
            b.clicked.connect(fn)
            row.addWidget(b)
        lay.addLayout(row)
        row = QHBoxLayout()
        row.addWidget(QLabel("Add a whole palette:"))
        for name in PALETTES:
            b = QPushButton(name)
            b.clicked.connect(lambda _=False, n=name: self.add_builtin(n))
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        lay.addWidget(box)
        self.fill()

    def fill(self):
        self.list.clear()
        for c in self.pal["colors"]:
            it = QListWidgetItem(swatch_icon(c["hex"], 20, 20), f"{c['name']}   #{c['hex']}")
            it.setData(Qt.UserRole, c)
            self.list.addItem(it)

    def _sync_order(self):
        self.pal["colors"] = [self.list.item(i).data(Qt.UserRole) for i in range(self.list.count())]

    def add(self):
        self._sync_order()
        hex6 = ColorPicker.get(self, self.book, None, "Add a colour", allow_none=False)
        if not hex6:
            return
        name, ok = QInputDialog.getText(self, "Add a colour", "Name it:")
        if not ok:
            return
        self.pal["colors"].append({"name": name.strip() or f"#{hex6}", "hex": hex6})
        self.fill()

    def add_builtin(self, name):
        self._sync_order()
        have = {c["hex"] for c in self.pal["colors"]}
        for cname, hex6 in PALETTES[name]:
            if hex6 not in have:
                self.pal["colors"].append({"name": cname, "hex": hex6})
        if not self.name.text().strip() or self.name.text() == "My palette":
            self.name.setText(name)
        self.fill()

    def edit_color(self):
        self._sync_order()
        it = self.list.currentItem()
        if it is None:
            return
        c = it.data(Qt.UserRole)
        new = ColorPicker.get(self, self.book, c["hex"], f"Change {c['name']}", allow_none=False)
        if not new or new == c["hex"]:
            return
        used = self.book.colors_in_use().get(c["hex"], 0)
        if used and QMessageBox.question(
                self, "Change colour",
                f"#{c['hex']} is used in {used} places. Change it there too?") == QMessageBox.Yes:
            self.replacements[c["hex"]] = new
        c["hex"] = new
        self.fill()

    def rename(self):
        self._sync_order()
        it = self.list.currentItem()
        if it is None:
            return
        c = it.data(Qt.UserRole)
        name, ok = QInputDialog.getText(self, "Rename colour", "Name:", text=c["name"])
        if ok and name.strip():
            c["name"] = name.strip()
            self.fill()

    def remove(self):
        self._sync_order()
        row = self.list.currentRow()
        if row >= 0:
            del self.pal["colors"][row]
            self.fill()

    def accept(self):
        self._sync_order()
        self.pal["name"] = self.name.text().strip() or "My palette"
        book = self.book
        book.begin()
        book.meta["palette"] = self.pal
        if self.replacements:
            book.replace_colors(self.replacements)
        book.done("palette")
        super().accept()


class RemapDialog(QDialog):
    """Swap each colour in the workbook for one from your palette."""

    def __init__(self, window):
        super().__init__(window)
        self.window = window
        self.book = window.book
        self.setWindowTitle("Swap colours for your palette")
        self.resize(560, 600)
        lay = QVBoxLayout(self)
        pal = (self.book.meta.get("palette") or {}).get("colors") or []
        if not pal:
            pal = [{"name": n, "hex": h} for n, h in PALETTES["Nord"]]
            note = "Your workbook has no palette yet, so this offers Nord. Set one up under Format → Palette."
        else:
            note = "Every cell, text colour and border using a colour on the left switches to the one you pick."
        self.pal = pal
        lbl = QLabel(note)
        lbl.setObjectName("muted")
        lbl.setWordWrap(True)
        lay.addWidget(lbl)
        used = sorted(self.book.colors_in_use().items(), key=lambda kv: -kv[1])
        self.table = QTableWidget(len(used), 3)
        self.table.setHorizontalHeaderLabels(["In the workbook", "Used", "Becomes"])
        self.table.verticalHeader().hide()
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.combos = []
        for i, (hex6, n) in enumerate(used):
            it = QTableWidgetItem(swatch_icon(hex6, 18, 18), f"#{hex6}")
            it.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(i, 0, it)
            cnt = QTableWidgetItem(str(n))
            cnt.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(i, 1, cnt)
            combo = QComboBox()
            combo.addItem("Keep as is", None)
            for c in pal:
                combo.addItem(swatch_icon(c["hex"]), f"{c['name']}  #{c['hex']}", c["hex"])
            self.table.setCellWidget(i, 2, combo)
            self.combos.append((hex6, combo))
        lay.addWidget(self.table, 1)
        row = QHBoxLayout()
        b = QPushButton("Suggest the nearest for each")
        b.clicked.connect(self.suggest)
        row.addWidget(b)
        b = QPushButton("Keep all")
        b.clicked.connect(lambda: [c.setCurrentIndex(0) for _, c in self.combos])
        row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        box.button(QDialogButtonBox.Ok).setText("Swap colours")
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        lay.addWidget(box)

    def suggest(self):
        choices = [c["hex"] for c in self.pal]
        for hex6, combo in self.combos:
            best = nearest(hex6, choices)
            combo.setCurrentIndex(1 + choices.index(best))

    def accept(self):
        mapping = {h: c.currentData() for h, c in self.combos if c.currentData()}
        if mapping:
            self.book.begin()
            n = self.book.replace_colors(mapping)
            self.book.done("palette")
            self.window.statusBar().showMessage(f"Swapped colours in {n} places.", 5000)
        super().accept()


class SuggestStylesDialog(QDialog):
    """Turn the formatting that repeats most into named styles."""

    def __init__(self, window, combos):
        super().__init__(window)
        self.window = window
        self.setWindowTitle("Make styles from repeated formatting")
        self.resize(640, 560)
        lay = QVBoxLayout(self)
        lbl = QLabel("These combinations of font, colour, fill and alignment come up again and again. "
                     "Name the ones you want as styles; their cells are linked to the style, so changing "
                     "the style later changes them all. Leave a name empty to skip it.")
        lbl.setObjectName("muted")
        lbl.setWordWrap(True)
        lay.addWidget(lbl)
        self.rows = []
        area = QScrollArea()
        area.setWidgetResizable(True)
        inner = QWidget()
        g = QGridLayout(inner)
        g.setColumnStretch(1, 1)
        for i, (key, count, sample) in enumerate(combos):
            prev = QLabel(sample[:40] or "Sample")
            name, size, bold, italic, color, fill, halign, valign, wrap = key
            css = [f"font-family: '{name}'", f"font-size: {max(8, int(size or 11))}pt", "padding: 4px 8px"]
            if bold:
                css.append("font-weight: bold")
            if italic:
                css.append("font-style: italic")
            if color:
                css.append(f"color: #{color}")
            css.append(f"background: #{fill}" if fill else "background: transparent")
            prev.setStyleSheet("; ".join(css))
            prev.setMinimumWidth(220)
            desc = QLabel(f"{count} cells · {name} {size:g}{' bold' if bold else ''}{' italic' if italic else ''}"
                          f"{' · wrap' if wrap else ''}")
            desc.setObjectName("faint")
            edit = QLineEdit()
            edit.setPlaceholderText("Style name")
            g.addWidget(prev, i * 2, 0)
            g.addWidget(edit, i * 2, 1)
            g.addWidget(desc, i * 2 + 1, 0, 1, 2)
            self.rows.append((key, edit))
        area.setWidget(inner)
        lay.addWidget(area, 1)
        box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        box.button(QDialogButtonBox.Ok).setText("Make styles")
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        lay.addWidget(box)

    def chosen(self):
        return [(key, edit.text().strip()) for key, edit in self.rows if edit.text().strip()]


class SheetLayoutDialog(QDialog):
    """Margin, cell padding and cell spacing for a sheet, previewed live.
    Display only: column widths and the data in the file don't change."""

    FIELDS = [("margin", "Margin", "Space between the edge of the pane and the grid (left and top)", 200),
              ("pad_x", "Cell padding, sides", "Space between a cell's left and right edges and its text", 60),
              ("pad_y", "Cell padding, top and bottom", "Space between a cell's top and bottom edges and its text", 60),
              ("spacing", "Cell spacing", "A gap between cells, where the sheet shows through (gridlines go away)", 40)]

    def __init__(self, window):
        from PySide6.QtWidgets import QSpinBox
        super().__init__(window)
        self.window = window
        self.book = window.book
        self.ws = window.current_ws()
        self.setWindowTitle(f"Margins and spacing · {self.ws.title}")
        self.original = copy.deepcopy(self.book.meta.get("sheet_layout", {}))
        lay = QVBoxLayout(self)
        note = QLabel("Like an HTML table's margin, cellpadding and cellspacing. Only changes how the sheet looks "
                      "in Atlas; columns, formulas and the file's data stay the same.")
        note.setObjectName("muted")
        note.setWordWrap(True)
        lay.addWidget(note)
        grid = QGridLayout()
        current = window.grid.sheet_layout()
        self.spins = {}
        for i, (key, label, tip, top) in enumerate(self.FIELDS):
            lbl = QLabel(label)
            lbl.setToolTip(tip)
            spin = QSpinBox()
            spin.setRange(0, top)
            spin.setSuffix(" px")
            spin.setValue(int(current[key]))
            spin.setToolTip(tip)
            spin.valueChanged.connect(self.preview)
            grid.addWidget(lbl, i, 0)
            grid.addWidget(spin, i, 1)
            self.spins[key] = spin
        lay.addLayout(grid)
        row = QHBoxLayout()
        for label, values in (("Excel-like", (0, 4, 2, 0)), ("Roomy", (16, 8, 4, 0)), ("Table", (24, 10, 6, 3))):
            b = QPushButton(label)
            b.clicked.connect(lambda _=False, v=values: self.set_values(v))
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        self.all = QCheckBox("Use these on every sheet")
        lay.addWidget(self.all)
        box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        lay.addWidget(box)

    def set_values(self, values):
        for (key, *_), v in zip(self.FIELDS, values):
            self.spins[key].setValue(v)

    def values(self):
        return {key: spin.value() for key, spin in self.spins.items()}

    def _apply(self, meta_layout):
        self.book.meta["sheet_layout"] = meta_layout
        self.window.grid.invalidate_layout(refit=True)

    def preview(self):
        trial = copy.deepcopy(self.original)
        trial[self.ws.title] = self.values()
        self._apply(trial)

    def reject(self):
        self._apply(copy.deepcopy(self.original))
        super().reject()

    def accept(self):
        final = copy.deepcopy(self.original)
        titles = [s.title for s in self.book.user_sheets()] if self.all.isChecked() else [self.ws.title]
        for t in titles:
            final[t] = self.values()
        self._apply(copy.deepcopy(self.original))
        self.book.begin([])
        self.book.meta["sheet_layout"] = final
        self.book.done("layout")
        super().accept()
