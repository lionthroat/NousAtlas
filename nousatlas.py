"""Nous Atlas: a spreadsheet for worldbuilding, made for workbooks that are
mostly text across many sheets.

Sheets live in a grouped sidebar, the grid scrolls by the pixel, a side
panel shows the whole cell, and map sheets can carry layers (where
each creature or plant lives, by time of day). Files stay ordinary .xlsx.
"""

import copy
import os
import re
import sys

from PySide6.QtCore import QPointF, QSettings, QSize, Qt, QTimer, QUrl, Signal
from PySide6.QtGui import (QAction, QActionGroup, QColor, QDesktopServices,
                           QFont, QFontDatabase, QFontMetrics, QIcon, QKeySequence, QPainter, QPen, QPixmap)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox,
                               QComboBox, QFileDialog, QFontComboBox, QFrame,
                               QHBoxLayout, QInputDialog, QLabel, QLineEdit,
                               QMainWindow, QMenu, QMessageBox, QPlainTextEdit,
                               QPushButton, QSplitter, QTabWidget, QToolBar,
                               QToolButton, QTreeWidget, QTreeWidgetItem,
                               QVBoxLayout, QWidget)
from openpyxl.cell.cell import MergedCell
from openpyxl.comments import Comment
from openpyxl.styles import NamedStyle
from openpyxl.worksheet.hyperlink import Hyperlink

import atlas_formula as fx
import atlas_model as model
import atlas_ranges as ranges
import atlas_squares as squares
from atlas_dialogs import (ColorPicker, PaletteDialog, RemapDialog, SheetLayoutDialog,
                           SuggestStylesDialog, swatch_icon)
from atlas_grid import SheetView, link_target
from atlas_model import Book
from atlas_theme import THEMES, qcolor, stylesheet

APP_NAME = "Nous Atlas"
XLSX_FILTER = "Excel workbooks (*.xlsx *.xlsm);;All files (*)"
BUILTIN_STYLE_NAMES = {"Normal"}


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def read_version():
    try:
        with open(resource_path("VERSION"), encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return "dev"


# ---------------------------------------------------------------- icons


def draw_icon(kind, color="#d8dee9", accent=None):
    """Small toolbar icons drawn in code, so they follow the theme."""
    pm = QPixmap(20, 20)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor(color), 1.6)
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)
    if kind in ("left", "center", "right"):
        lines = [(3, 17), (3, 13), (3, 17), (3, 11)]
        for i, (a, b) in enumerate(lines):
            y = 4 + i * 4
            w = b - a
            if kind == "left":
                x1 = 3
            elif kind == "right":
                x1 = 17 - w
            else:
                x1 = 10 - w / 2
            p.drawLine(int(x1), y, int(x1 + w), y)
    elif kind in ("top", "middle", "bottom"):
        p.drawLine(3, 2 if kind == "top" else 18 if kind == "bottom" else 10,
                   17, 2 if kind == "top" else 18 if kind == "bottom" else 10)
        y = {"top": 5, "middle": 6, "bottom": 9}[kind]
        p.setPen(QPen(QColor(color), 1.2))
        for i in range(2):
            p.drawLine(6, y + i * 4, 14, y + i * 4)
    elif kind == "wrap":
        p.drawLine(3, 5, 17, 5)
        p.drawLine(3, 10, 14, 10)
        p.drawArc(10, 10, 7, 6, 90 * 16, -180 * 16)
        p.drawLine(13, 16, 3, 16)
    elif kind == "merge":
        p.drawRect(2, 5, 16, 10)
        p.drawLine(5, 10, 8, 10)
        p.drawLine(12, 10, 15, 10)
        p.drawLine(5, 10, 7, 8)
        p.drawLine(5, 10, 7, 12)
        p.drawLine(15, 10, 13, 8)
        p.drawLine(15, 10, 13, 12)
    elif kind == "borders":
        p.drawRect(3, 3, 14, 14)
        p.setPen(QPen(QColor(color), 1, Qt.DotLine))
        p.drawLine(10, 3, 10, 17)
        p.drawLine(3, 10, 17, 10)
    elif kind == "text":
        f = QFont("Arial")
        f.setPixelSize(13)
        f.setBold(True)
        p.setFont(f)
        p.drawText(0, -2, 20, 18, Qt.AlignCenter, "A")
        p.fillRect(3, 15, 14, 4, qcolor(accent) if accent else QColor(color))
    elif kind == "fill":
        p.drawRoundedRect(4, 3, 12, 9, 2, 2)
        p.fillRect(3, 15, 14, 4, qcolor(accent) if accent else Qt.transparent)
        if not accent:
            p.drawRect(3, 15, 14, 3)
    elif kind == "bordercolor":
        p.setPen(QPen(QColor(color), 1.2, Qt.DotLine))
        p.drawRect(3, 2, 14, 11)
        p.setPen(Qt.NoPen)
        p.fillRect(3, 15, 14, 4, qcolor(accent) if accent else QColor(color))
    elif kind in ("pane_left", "pane_right"):
        p.setPen(QPen(QColor(color), 1.5))
        p.drawRoundedRect(2, 3, 16, 14, 2, 2)
        x = 7 if kind == "pane_left" else 13
        p.drawLine(x, 3, x, 17)
        p.fillRect(3 if kind == "pane_left" else 13, 4, 4, 12, QColor(color))
    elif kind == "spacing":
        p.setPen(QPen(QColor(color), 1.2))
        for x, y in ((3, 3), (11, 3), (3, 11), (11, 11)):
            p.drawRect(x, y, 6, 6)
        p.setPen(QPen(QColor(color), 1, Qt.DotLine))
        p.drawRect(1, 1, 18, 18)
    elif kind == "clear":
        p.drawLine(4, 16, 16, 4)
        f = QFont("Arial")
        f.setPixelSize(11)
        p.setFont(f)
        p.drawText(0, 0, 14, 14, Qt.AlignCenter, "A")
    p.end()
    return QIcon(pm)


def app_icon():
    path = resource_path("icon.ico")
    return QIcon(path) if os.path.exists(path) else QIcon()


# ---------------------------------------------------------------- panel sections


class Section(QWidget):
    """A titled part of the side panel that folds open and shut."""

    toggled = Signal(bool)

    def __init__(self, title, content, key, settings):
        super().__init__()
        self.key, self.settings = key, settings
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        self.header = QToolButton()
        self.header.setText(title)
        self.header.setCheckable(True)
        self.header.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.header.setSizePolicy(self.header.sizePolicy().horizontalPolicy().Expanding,
                                  self.header.sizePolicy().verticalPolicy())
        self.header.setObjectName("sectionHeader")
        self.header.toggled.connect(self._toggled)
        lay.addWidget(self.header)
        self.content = content
        lay.addWidget(content)
        self.header.setChecked(settings.value(key, "true") == "true")
        self._toggled(self.header.isChecked())

    def _toggled(self, on):
        self.header.setArrowType(Qt.DownArrow if on else Qt.RightArrow)
        self.content.setVisible(on)
        self.settings.setValue(self.key, "true" if on else "false")
        self.toggled.emit(on)

    def expanded(self):
        return self.header.isChecked()

    def expand(self, on=True):
        self.header.setChecked(on)


# ---------------------------------------------------------------- font picker


class FontBox(QComboBox):
    """Fonts with the ones this workbook uses first, each shown in itself.
    Type a name and press Enter to pick it."""

    fontChosen = Signal(str)

    def __init__(self, used_fn):
        super().__init__()
        self.used_fn = used_fn
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.setMaxVisibleItems(24)
        self.view().setMinimumWidth(300)
        self.view().setTextElideMode(Qt.ElideNone)
        # every row the same height: Qt then doesn't measure 300+ fonts one by one
        self.view().setUniformItemSizes(True)
        self.families = sorted(set(QFontDatabase.families()), key=str.lower)
        self._filled_for = None
        self.lineEdit().returnPressed.connect(self._typed)
        self.activated.connect(self._picked)
        comp = self.completer()
        comp.setCaseSensitivity(Qt.CaseInsensitive)
        comp.setFilterMode(Qt.MatchStartsWith)

    def _header(self, text):
        self.addItem(text)
        it = self.model().item(self.count() - 1)
        it.setEnabled(False)
        it.setSelectable(False)
        f = QFont(self.font())
        f.setBold(True)
        f.setPointSizeF(max(7.0, f.pointSizeF() - 1))
        it.setFont(f)

    def _font_item(self, name):
        self.addItem(name, name)
        it = self.model().item(self.count() - 1)
        f = QFont(name)
        f.setPointSizeF(self.font().pointSizeF() + 1)
        it.setFont(f)

    def fill(self):
        used = [n for n in self.used_fn()]
        if used == self._filled_for:
            return
        text = self.currentText()
        self.blockSignals(True)
        self.clear()
        if used:
            self._header("In this workbook")
            for name in used:
                self._font_item(name)
            self.insertSeparator(self.count())
            self._header("All fonts")
        for name in self.families:
            self._font_item(name)
        self.setEditText(text)
        self.blockSignals(False)
        self._filled_for = used

    def prewarm(self):
        """Fill the list and load the fonts a few at a time while the app is
        idle, so the first open is instant."""
        self.fill()
        names = list(self.families)
        def step():
            for _ in range(12):
                if not names:
                    return
                QFontMetrics(QFont(names.pop())).horizontalAdvance("Ag")
            QTimer.singleShot(15, step)
        QTimer.singleShot(15, step)

    def showPopup(self):
        self.fill()
        name = self.currentText()
        idx = self.findData(name)
        if idx >= 0:
            self.setCurrentIndex(idx)
        super().showPopup()

    def show_name(self, name):
        self.blockSignals(True)
        self.setEditText(name)
        self.blockSignals(False)

    def _picked(self, idx):
        name = self.itemData(idx)
        if name:
            self.show_name(name)
            self.fontChosen.emit(name)

    def _typed(self):
        text = self.currentText().strip()
        match = next((f for f in self.families if f.lower() == text.lower()), None)
        if match:
            self.show_name(match)
            self.fontChosen.emit(match)


# ---------------------------------------------------------------- square card


class HoverCard(QFrame):
    """The card that appears over a map square. You can move onto it and
    click what's in it."""

    def __init__(self, window):
        super().__init__(window, Qt.Tool | Qt.FramelessWindowHint | Qt.WindowDoesNotAcceptFocus)
        self.window = window
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setObjectName("card")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)
        self.label = QLabel()
        self.label.setTextFormat(Qt.RichText)
        self.label.setWordWrap(True)
        self.label.setMinimumWidth(240)
        self.label.setMaximumWidth(380)
        self.label.setOpenExternalLinks(False)
        self.label.linkActivated.connect(self.clicked)
        lay.addWidget(self.label)
        self.key = None

    def show_for(self, key, html_text, gpos):
        self.key = key
        self.label.setText(html_text)
        self.adjustSize()
        screen = QApplication.screenAt(gpos) or QApplication.primaryScreen()
        area = screen.availableGeometry()
        x, y = gpos.x() + 16, gpos.y() + 18
        if x + self.width() > area.right():
            x = gpos.x() - self.width() - 12
        if y + self.height() > area.bottom():
            y = gpos.y() - self.height() - 12
        self.move(x, y)
        self.show()
        self.raise_()

    def clicked(self, link):
        self.hide()
        self.key = None
        self.window.go_to_href(link)

    def leaveEvent(self, e):
        super().leaveEvent(e)
        self.window.card_timer_hide.start()

    def enterEvent(self, e):
        super().enterEvent(e)
        self.window.card_timer_hide.stop()


# ---------------------------------------------------------------- sidebar


class SheetTree(QTreeWidget):
    """Sheets in collapsible groups. Drag to reorder or to move a sheet into
    a group; groups can't nest."""

    def __init__(self, window):
        super().__init__()
        self.window = window
        self.setHeaderHidden(True)
        self.setIndentation(14)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.setExpandsOnDoubleClick(False)

    def dropEvent(self, e):
        dragged = self.currentItem()
        super().dropEvent(e)
        # undo anything that nested a group or put a sheet under a sheet
        for i in range(self.topLevelItemCount()):
            top = self.topLevelItem(i)
            for j in reversed(range(top.childCount())):
                child = top.child(j)
                bad = child.data(0, Qt.UserRole)[0] == "group" or top.data(0, Qt.UserRole)[0] == "sheet"
                if bad:
                    top.takeChild(j)
                    self.insertTopLevelItem(i + 1, child)
                for k in reversed(range(child.childCount())):
                    grand = child.takeChild(k)
                    top.insertChild(j + 1, grand)
        self.window.sidebar_reordered(dragged)


class Sidebar(QWidget):
    def __init__(self, window):
        super().__init__()
        self.window = window
        lay = QVBoxLayout(self)
        lay.setContentsMargins(6, 6, 0, 6)
        lay.setSpacing(6)
        self.filter = QLineEdit()
        self.filter.setPlaceholderText("Go to sheet  (Ctrl+P)")
        self.filter.setClearButtonEnabled(True)
        self.filter.textChanged.connect(self.apply_filter)
        self.filter.returnPressed.connect(self.open_first)
        self.filter.installEventFilter(self)
        lay.addWidget(self.filter)
        self.tree = SheetTree(window)
        self.tree.itemClicked.connect(self.clicked)
        self.tree.itemActivated.connect(self.clicked)
        self.tree.itemExpanded.connect(lambda it: self.window.group_collapsed(it, False))
        self.tree.itemCollapsed.connect(lambda it: self.window.group_collapsed(it, True))
        self.tree.customContextMenuRequested.connect(window.sidebar_menu)
        lay.addWidget(self.tree, 1)

    def eventFilter(self, obj, e):
        if obj is self.filter and e.type() == e.Type.KeyPress:
            if e.key() == Qt.Key_Escape:
                self.filter.clear()
                self.window.grid.setFocus()
                return True
            if e.key() == Qt.Key_Down:
                self.tree.setFocus()
                first = self._visible_sheets()
                if first:
                    self.tree.setCurrentItem(first[0])
                return True
        return False

    def clicked(self, it):
        kind, name = it.data(0, Qt.UserRole)
        if kind == "sheet":
            self.window.open_sheet(name)
        else:
            it.setExpanded(not it.isExpanded())

    def _visible_sheets(self):
        out = []
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            items = [top] if top.data(0, Qt.UserRole)[0] == "sheet" else [top.child(j) for j in range(top.childCount())]
            out += [it for it in items if not it.isHidden()]
        return out

    def apply_filter(self, text):
        text = text.strip().lower()
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            if top.data(0, Qt.UserRole)[0] == "sheet":
                top.setHidden(bool(text) and text not in top.data(0, Qt.UserRole)[1].lower())
                continue
            any_shown = False
            group_hit = bool(text) and text in top.data(0, Qt.UserRole)[1].lower()
            for j in range(top.childCount()):
                ch = top.child(j)
                shown = not text or group_hit or text in ch.data(0, Qt.UserRole)[1].lower()
                ch.setHidden(not shown)
                any_shown |= shown
            top.setHidden(bool(text) and not any_shown)
            if text and any_shown:
                top.setExpanded(True)
            elif not text:
                top.setExpanded(not self.window.is_group_collapsed(top.data(0, Qt.UserRole)[1]))

    def open_first(self):
        items = self._visible_sheets()
        if items:
            self.window.open_sheet(items[0].data(0, Qt.UserRole)[1])
            self.filter.clear()
            self.window.grid.setFocus()

    def select_sheet(self, title):
        for it in self._all_items():
            if it.data(0, Qt.UserRole) == ("sheet", title):
                self.tree.blockSignals(True)
                self.tree.setCurrentItem(it)
                self.tree.blockSignals(False)
                return

    def _all_items(self):
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            yield top
            for j in range(top.childCount()):
                yield top.child(j)


# ---------------------------------------------------------------- inspector


class BigEdit(QPlainTextEdit):
    """Commits on Ctrl+Enter or when focus leaves; Esc puts it back."""

    def __init__(self, on_commit, on_revert):
        super().__init__()
        self.on_commit, self.on_revert = on_commit, on_revert

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Return, Qt.Key_Enter) and e.modifiers() & Qt.ControlModifier:
            self.on_commit()
            return
        if e.key() == Qt.Key_Escape:
            self.on_revert()
            return
        super().keyPressEvent(e)

    def focusOutEvent(self, e):
        super().focusOutEvent(e)
        self.on_commit()


class Inspector(QWidget):
    def __init__(self, window):
        super().__init__()
        self.window = window
        self.loading = False
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(6)
        self.where = QLabel()
        self.where.setObjectName("muted")
        self.where.setWordWrap(True)
        lay.addWidget(self.where)
        self.text = BigEdit(self.commit_text, self.load)
        self.text.setPlaceholderText("Empty cell. Type here, then Ctrl+Enter.")
        self.text.setToolTip("Ctrl+Enter saves; Esc puts it back. Alt+← / Alt+→ go back and forward.")
        lay.addWidget(self.text, 3)
        self.result = QLabel()
        self.result.setWordWrap(True)
        self.result.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lay.addWidget(self.result)
        self.note_label = lbl = QLabel("Note")
        lbl.setObjectName("faint")
        lay.addWidget(lbl)
        self.note = BigEdit(self.commit_note, self.load)
        self.note.setPlaceholderText("No note")
        self.note.setMaximumHeight(110)
        lay.addWidget(self.note, 1)
        self.link = QLabel()
        self.link.setObjectName("muted")
        self.link.setWordWrap(True)
        lay.addWidget(self.link)
        self.square_box = QWidget()
        sb = QVBoxLayout(self.square_box)
        sb.setContentsMargins(0, 6, 0, 0)
        sb.setSpacing(4)
        self.square_card = QLabel()
        self.square_card.setTextFormat(Qt.RichText)
        self.square_card.setWordWrap(True)
        self.square_card.linkActivated.connect(lambda link: self.window.go_to_href(link))
        sb.addWidget(self.square_card)
        lbl = QLabel("Note for this square")
        lbl.setObjectName("faint")
        sb.addWidget(lbl)
        self.square_note = BigEdit(self.commit_square_note, self.load)
        self.square_note.setPlaceholderText("Anything about this square. Shows on its hover card.")
        self.square_note.setMaximumHeight(110)
        sb.addWidget(self.square_note)
        lay.addWidget(self.square_box)
        self.square_key = None
        self.ranges_box = QWidget()
        rb = QVBoxLayout(self.ranges_box)
        rb.setContentsMargins(0, 6, 0, 0)
        rb.setSpacing(4)
        self.ranges_label = QLabel("Range")
        self.ranges_label.setObjectName("faint")
        rb.addWidget(self.ranges_label)
        self.ranges_row = QVBoxLayout()
        self.ranges_row.setSpacing(4)
        rb.addLayout(self.ranges_row)
        lay.addWidget(self.ranges_box)

    def load(self):
        w = self.window
        ws = w.current_ws()
        self.loading = True
        try:
            if ws is None:
                self.square_box.hide()
                self.where.setText("")
                self.text.setPlainText("")
                self.note.setPlainText("")
                self.result.setText("")
                self.link.setText("")
                self.ranges_box.hide()
                return
            r, c = w.grid.cur
            cell = ws._cells.get((r, c))
            ref = f"{fx.num_to_col(c)}{r}"
            r1, c1, r2, c2 = w.grid.selection()
            where = f"{ws.title} · {ref}"
            if (r1, c1, r2, c2) != (r, c, r, c) and not w.grid._is_single_merge(r1, c1, r2, c2):
                where += f"   ({fx.num_to_col(c1)}{r1}:{fx.num_to_col(c2)}{r2})"
            info = w.book.meta["maps"].get(ws.title)
            sq = ranges.map_square(info, r, c) if info else None
            if sq:
                where += f" · map square {ranges.square_label(sq)}"
            self.where.setText(where)
            self.text.setReadOnly(w.read_only_sheet(ws))
            self.text.setPlainText(w.book.edit_text(ws, r, c))
            raw = cell.value if cell is not None else None
            if fx.is_formula(raw):
                v = w.book.calc.value(ws.title, r, c)
                shown = model.format_value(v, cell.number_format)
                color = w.theme["danger"] if isinstance(v, fx.XlError) else w.theme["accent"]
                self.result.setText(f"<span style='color:{color}'>= {shown}</span>")
                self.result.show()
            else:
                self.result.hide()
            self.note.setPlainText(cell.comment.text if cell is not None and cell.comment else "")
            link = cell.hyperlink if cell is not None else None
            if link is not None and (link.target or link.location):
                self.link.setText(f"Links to {link.target or link.location}. Click the button in the cell to go there; "
                                  "click beside it (or use the arrow keys) to select the cell. Right-click → "
                                  "Remove link turns it back into plain text.")
                self.link.show()
            else:
                self.link.hide()
            self.load_ranges(ws, r)
            self.load_square(ws, r, c)
        finally:
            self.loading = False

    def load_square(self, ws, r, c):
        w = self.window
        info = w.book.meta["maps"].get(ws.title)
        sq = ranges.map_square(info, r, c) if info else None
        self.square_key = (ws.title, ranges.square_label(sq)) if sq else None
        if sq is None:
            self.square_box.hide()
            self.text.setMaximumHeight(16777215)
            self.note_label.show()
            self.note.show()
            return
        # a map square: its own note (below) replaces the cell note, unless the cell has one
        has_note = bool(self.note.toPlainText())
        self.note_label.setVisible(has_note)
        self.note.setVisible(has_note)
        cd = w.squares.card(ws.title, sq)
        self.square_card.setText(squares.card_html(cd, w.theme, ws.title, with_note=False))
        self.text.setMaximumHeight(70)          # map squares hold a short mark; make room for the card
        self.square_note.setPlainText(cd["note"])
        self.square_box.show()

    def commit_square_note(self):
        w = self.window
        if self.loading or self.square_key is None or w.book is None:
            return
        map_title, label = self.square_key
        old = w.book.meta.get("square_notes", {}).get(map_title, {}).get(label, "")
        new = self.square_note.toPlainText().strip()
        if new == old:
            return
        w.book.begin([])
        notes = w.book.meta.setdefault("square_notes", {}).setdefault(map_title, {})
        if new:
            notes[label] = new
        else:
            notes.pop(label, None)
        w.book.done("notes")

    def load_ranges(self, ws, row):
        w = self.window
        while self.ranges_row.count():
            item = self.ranges_row.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        if w.book is None or w.book.ranges is None:
            self.ranges_box.hide()
            return
        names = {}
        for (r, c), cell in ws._cells.items():
            if r == row and isinstance(cell.value, str) and not fx.is_formula(cell.value):
                names[cell.value.strip().lower()] = cell.value.strip()
        layers = [l for l in w.book.ranges.layers if l.name.lower() in names]
        shown = False
        maps = [t for t in w.book.meta["maps"] if w.book.sheet(t) is not None]
        for layer in layers:
            for title in maps:
                if layer.has_map(title) or len(maps) == 1 or not any(layer.has_map(t) for t in maps):
                    b = QPushButton(swatch_icon(layer.color), f"{layer.name} on {title}")
                    b.setStyleSheet("text-align: left;")
                    b.clicked.connect(lambda _=False, n=layer.name, t=title: w.show_range(n, t))
                    self.ranges_row.addWidget(b)
                    shown = True
        self.ranges_box.setVisible(shown)

    def commit_text(self):
        w = self.window
        if self.loading or w.current_ws() is None or self.text.isReadOnly():
            return
        r, c = w.grid.cur
        text = self.text.toPlainText()
        if text != w.book.edit_text(w.current_ws(), r, c):
            w.grid.set_cell_text(r, c, text)

    def commit_note(self):
        w = self.window
        ws = w.current_ws()
        if self.loading or ws is None or w.read_only_sheet(ws):
            return
        r, c = w.grid.cur
        cell = ws._cells.get((r, c))
        old = cell.comment.text if cell is not None and cell.comment else ""
        new = self.note.toPlainText()
        if new == old:
            return
        w.book.begin([ws])
        cell = ws.cell(r, c)
        cell.comment = Comment(new, "") if new.strip() else None
        if cell.comment is not None:
            cell.comment.width, cell.comment.height = 300, 150
        w.book.done()


# ---------------------------------------------------------------- find bar


class FindBar(QFrame):
    def __init__(self, window):
        super().__init__()
        self.window = window
        self.setObjectName("panel")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(8, 4, 8, 4)
        self.edit = QLineEdit()
        self.edit.setPlaceholderText("Find")
        self.edit.setClearButtonEnabled(True)
        self.edit.textChanged.connect(lambda: self.timer.start())
        self.edit.installEventFilter(self)
        lay.addWidget(self.edit, 1)
        self.all = QCheckBox("All sheets")
        self.all.setChecked(True)
        self.all.toggled.connect(self.search)
        lay.addWidget(self.all)
        self.case = QCheckBox("Match case")
        self.case.toggled.connect(self.search)
        lay.addWidget(self.case)
        prev = QPushButton("↑")
        prev.setFixedWidth(30)
        prev.clicked.connect(lambda: self.step(-1))
        nxt = QPushButton("↓")
        nxt.setFixedWidth(30)
        nxt.clicked.connect(lambda: self.step(1))
        lay.addWidget(prev)
        lay.addWidget(nxt)
        self.count = QLabel()
        self.count.setObjectName("muted")
        self.count.setMinimumWidth(110)
        lay.addWidget(self.count)
        close = QPushButton("✕")
        close.setFixedWidth(30)
        close.clicked.connect(self.close_bar)
        lay.addWidget(close)
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(180)
        self.timer.timeout.connect(self.search)
        self.hits = []
        self.index = -1

    def eventFilter(self, obj, e):
        if e.type() == e.Type.KeyPress:
            if e.key() in (Qt.Key_Return, Qt.Key_Enter):
                if self.timer.isActive():
                    self.timer.stop()
                    self.search()
                self.step(-1 if e.modifiers() & Qt.ShiftModifier else 1)
                return True
            if e.key() == Qt.Key_Escape:
                self.close_bar()
                return True
        return False

    def open_bar(self, text=None):
        self.show()
        if text is not None:
            self.edit.setText(text)
            self.search()
        self.edit.setFocus()
        self.edit.selectAll()

    def close_bar(self):
        self.hide()
        self.hits = []
        self.window.grid.find_hits = set()
        self.window.grid.find_current = None
        self.window.grid.viewport().update()
        self.window.grid.setFocus()

    def search(self):
        w = self.window
        needle = self.edit.text()
        self.hits = []
        self.index = -1
        if needle and w.book is not None:
            case = self.case.isChecked()
            n = needle if case else needle.lower()
            sheets = w.sidebar_order() if self.all.isChecked() else [w.current_ws()]
            for ws in sheets:
                if ws is None:
                    continue
                for (r, c) in sorted(ws._cells):
                    cell = ws._cells[(r, c)]
                    if cell.value is None:
                        continue
                    texts = [w.book.display_text(ws, r, c)]
                    if fx.is_formula(cell.value):
                        texts.append(cell.value)
                    if cell.comment is not None:
                        texts.append(cell.comment.text)
                    if any(n in (t if case else t.lower()) for t in texts):
                        self.hits.append((ws.title, r, c))
        self.update_marks()
        if self.hits:
            cur_title = w.current_ws().title if w.current_ws() else None
            order = [i for i, h in enumerate(self.hits) if h[0] == cur_title]
            self.index = (order[0] if order else 0) - 1
            self.count.setText(f"{len(self.hits)} found")
        else:
            self.count.setText("Nothing found" if needle else "")

    def update_marks(self):
        w = self.window
        ws = w.current_ws()
        title = ws.title if ws else None
        w.grid.find_hits = {(r, c) for t, r, c in self.hits if t == title}
        cur = self.hits[self.index] if 0 <= self.index < len(self.hits) else None
        w.grid.find_current = (cur[1], cur[2]) if cur and cur[0] == title else None
        w.grid.viewport().update()

    def step(self, d):
        if not self.hits:
            return
        self.index = (self.index + d) % len(self.hits)
        title, r, c = self.hits[self.index]
        self.window.open_sheet(title, (r, c))
        self.update_marks()
        sheets = len({h[0] for h in self.hits})
        self.count.setText(f"{self.index + 1} of {len(self.hits)}" + (f" · {sheets} sheets" if sheets > 1 else ""))


# ---------------------------------------------------------------- main window


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("Nous", os.environ.get("NOUSATLAS_SETTINGS", "NousAtlas"))
        self.theme_name = self.settings.value("theme", "Nord")
        if self.theme_name not in THEMES:
            self.theme_name = "Nord"
        self.theme = THEMES[self.theme_name]
        self.book = None
        self.ws_title = None
        self.history, self.future = [], []
        self.range_state = ranges.ViewState()
        self.clip = {}
        self.show_ranges_sheet = False
        self.last_text_color = "BF616A"
        self.last_fill_color = "EBCB8B"
        self._syncing = False

        self.setWindowIcon(app_icon())
        self.grid = SheetView(self.theme)
        self.grid.clip = self.clip
        self.grid.overlay = ranges.Overlay(self)
        self.grid.tool = ranges.Brush(self)
        self.grid.row_mode = self.settings.value("row_mode", "compact")
        self.grid.adapt_colors = self.settings.value("adapt_colors", "true") == "true"
        self.grid.currentChanged.connect(self.on_current_changed)
        self.grid.linkActivated.connect(self.follow_link)
        self.grid.zoomChanged.connect(lambda z: self.zoom_label.setText(f"{round(z * 100)}%"))
        self.grid.contextRequested.connect(self.grid_menu)
        self.grid.cellHovered.connect(self.on_hover)
        self.grid.colorPicked.connect(self.pick_fill_from)
        self.grid.deleteRequested.connect(lambda kind: self.delete_rows() if kind == "rows" else self.delete_cols())
        self.squares = None
        self.card = HoverCard(self)
        self.card_timer_show = QTimer(self)
        self.card_timer_show.setSingleShot(True)
        self.card_timer_show.setInterval(350)
        self.card_timer_show.timeout.connect(self.show_card)
        self.card_timer_hide = QTimer(self)
        self.card_timer_hide.setSingleShot(True)
        self.card_timer_hide.setInterval(300)
        self.card_timer_hide.timeout.connect(self.hide_card)
        self._hover = None

        self.sidebar = Sidebar(self)
        self.inspector = Inspector(self)
        self.find_bar = FindBar(self)
        self.find_bar.hide()
        center = QWidget()
        cl = QVBoxLayout(center)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        cl.addWidget(self.find_bar)
        self.paint_bar = QFrame()
        self.paint_bar.setObjectName("panel")
        pb = QHBoxLayout(self.paint_bar)
        pb.setContentsMargins(10, 4, 8, 4)
        self.paint_label = QLabel()
        pb.addWidget(self.paint_label, 1)
        done = QPushButton("Done painting")
        done.setToolTip("Stop painting (Esc)")
        done.clicked.connect(lambda: self.stop_painting())
        pb.addWidget(done)
        self.paint_bar.hide()
        cl.addWidget(self.paint_bar)
        cl.addWidget(self.grid, 1)

        self.ranges_panel = ranges.RangesPanel(self)
        from PySide6.QtWidgets import QScrollArea
        self.right = QScrollArea()
        self.right.setWidgetResizable(True)
        self.right.setFrameShape(QFrame.NoFrame)
        self.right.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        holder = QWidget()
        hl = QVBoxLayout(holder)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(0)
        self.cell_section = Section("Cell", self.inspector, "sec_cell", self.settings)
        self.layers_section = Section("Layers", self.ranges_panel, "sec_layers", self.settings)
        self.layers_section.toggled.connect(lambda on: self.ranges_panel.refresh() if on else None)
        hl.addWidget(self.cell_section)
        hl.addWidget(self.layers_section)
        hl.addStretch(1)
        self.right.setWidget(holder)
        self.layers_on = self.settings.value("layers_on", "false") == "true"

        self.splitter = QSplitter()
        self.splitter.addWidget(self.sidebar)
        self.splitter.addWidget(center)
        self.splitter.setCollapsible(0, True)
        self.splitter.setCollapsible(1, False)
        self.splitter.setStretchFactor(1, 1)
        self.splitter.setSizes([230, 1200])
        self.splitter.setHandleWidth(1)
        self.splitter.splitterMoved.connect(lambda *a: self.sync_pane_buttons())
        self.setCentralWidget(self.splitter)

        # Cell and Layers: a panel you can close, float and drag anywhere, or dock left or right
        from PySide6.QtWidgets import QDockWidget
        self.dock = QDockWidget("", self)
        self.dock.setObjectName("panel_dock")
        self.dock.setWidget(self.right)
        self.dock.setFeatures(QDockWidget.DockWidgetClosable | QDockWidget.DockWidgetMovable |
                              QDockWidget.DockWidgetFloatable)
        self.dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock)
        self.dock.visibilityChanged.connect(self._dock_visibility)
        self._panel_auto = False                    # opened by Layers; close it again when Layers goes off

        self.build_actions()
        self.build_toolbar()
        self.build_menus()
        self.row_mode_btn.setChecked(self.grid.row_mode == "compact")
        self.a_compact.setChecked(self.grid.row_mode == "compact")
        self.zoom_label = QLabel("100%")
        self.zoom_label.setObjectName("muted")
        self.stats_label = QLabel()
        self.stats_label.setObjectName("muted")
        self.statusBar().addPermanentWidget(self.stats_label)
        self.statusBar().addPermanentWidget(self.zoom_label)
        self.apply_theme()
        self.update_enabled()
        self.update_title()
        geo = self.settings.value("geometry")
        if geo is not None:
            self.restoreGeometry(geo)
        else:
            screen = QApplication.primaryScreen().availableGeometry()
            self.resize(int(screen.width() * 0.85), int(screen.height() * 0.85))
        sizes = self.settings.value("splitter2")
        if sizes:
            try:
                self.splitter.setSizes([int(x) for x in sizes][:2])
            except (TypeError, ValueError):
                pass
        state = self.settings.value("window_state")
        if state is not None:
            self.restoreState(state)
        self.fit_panel_width()
        self.sync_pane_buttons()

    # ------------------------------------------------------------ actions
    def act(self, text, fn, shortcut=None, checkable=False, tip=None):
        a = QAction(text, self)
        if shortcut:
            a.setShortcut(QKeySequence(shortcut))
        a.setCheckable(checkable)
        if tip:
            a.setToolTip(tip)
        a.triggered.connect(fn)
        return a

    def build_actions(self):
        self.a_new = self.act("New workbook", self.new_book, QKeySequence.New)
        self.a_open = self.act("Open…", self.open_dialog, QKeySequence.Open)
        self.a_save = self.act("Save", self.save, QKeySequence.Save)
        self.a_save_as = self.act("Save as…", self.save_as, "Ctrl+Shift+S")
        self.a_close = self.act("Close workbook", self.close_book, "Ctrl+W")
        self.a_quit = self.act("Quit", self.close, "Ctrl+Q")
        self.a_undo = self.act("Undo", self.undo, QKeySequence.Undo)
        self.a_redo = self.act("Redo", self.redo, "Ctrl+Y")
        self.a_redo.setShortcuts([QKeySequence("Ctrl+Y"), QKeySequence("Ctrl+Shift+Z")])
        self.a_find = self.act("Find…", lambda: self.find_bar.open_bar(), QKeySequence.Find)
        self.a_goto = self.act("Go to sheet…", self.focus_filter, "Ctrl+P")
        self.a_back = self.act("Back", self.go_back, "Alt+Left")
        self.a_fwd = self.act("Forward", self.go_forward, "Alt+Right")
        self.a_bold = self.act("Bold", lambda: self.toggle_font("b"), QKeySequence.Bold, True)
        self.a_italic = self.act("Italic", lambda: self.toggle_font("i"), QKeySequence.Italic, True)
        self.a_under = self.act("Underline", lambda: self.toggle_font("u"), QKeySequence.Underline, True)
        self.a_strike = self.act("Strikethrough", lambda: self.toggle_font("strike"), "Ctrl+5", True)
        self.a_bigger = self.act("Bigger text", lambda: self.step_font_size(1), "Ctrl+Shift+>")
        self.a_smaller = self.act("Smaller text", lambda: self.step_font_size(-1), "Ctrl+Shift+<")
        self.addAction(self.a_bigger)
        self.addAction(self.a_smaller)
        self.a_wrap = self.act("Wrap text", self.toggle_wrap, None, True)
        self.a_merge = self.act("Merge and center", self.toggle_merge, None, True)
        self.a_clear_fmt = self.act("Clear formatting", self.clear_formatting)
        self.a_note = self.act("Edit note…", self.edit_cell_note, "Shift+F2")
        self.a_layer_here = self.act("New layer on the selected squares", self.new_layer_here, "Ctrl+L",
                                     tip="On a map: put the selected squares in a layer (new or existing). "
                                         "On another sheet: make a layer for the selected name. (Ctrl+L)")
        self.halign = {}
        for k in ("left", "center", "right"):
            self.halign[k] = self.act(f"Align {k}", lambda _=False, k=k: self.set_halign(k), None, True)
        self.valign = {}
        for k, v in (("top", "top"), ("middle", "center"), ("bottom", "bottom")):
            self.valign[v] = self.act(f"Align {k}", lambda _=False, v=v: self.set_valign(v), None, True)
        for a in self.halign.values():
            a.setToolTip(a.text())
        for a in self.valign.values():
            a.setToolTip(a.text())

    def build_toolbar(self):
        tb = QToolBar("Format")
        tb.setMovable(False)
        tb.setIconSize(QSize(18, 18))
        self.addToolBar(tb)
        self.toolbar = tb
        self.sidebar_btn = QToolButton()
        self.sidebar_btn.setCheckable(True)
        self.sidebar_btn.setChecked(True)
        self.sidebar_btn.setToolTip("Show or hide the sheet list (Ctrl+Shift+\\)")
        self.sidebar_btn.clicked.connect(self.set_sidebar)
        tb.addWidget(self.sidebar_btn)
        tb.addSeparator()
        self.font_box = FontBox(self.fonts_in_use)
        self.font_box.setMinimumWidth(210)
        self.font_box.setMaximumWidth(260)
        self.font_box.fontChosen.connect(lambda name: self.set_font_attr(name=name))
        tb.addWidget(self.font_box)
        self.size_box = QComboBox()
        self.size_box.setEditable(True)
        self.size_box.addItems([str(s) for s in (8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36)])
        self.size_box.setFixedWidth(62)
        self.size_box.lineEdit().returnPressed.connect(self.size_entered)
        self.size_box.activated.connect(lambda i: self.size_entered())
        tb.addWidget(self.size_box)
        for label, step, tip in (("▲", 1, "Bigger text (Ctrl+Shift+>)"), ("▼", -1, "Smaller text (Ctrl+Shift+<)")):
            b = QToolButton()
            b.setText(label)
            b.setToolTip(tip)
            b.setAutoRepeat(True)
            b.setFixedWidth(22)
            f = b.font()
            f.setPointSizeF(max(7.0, f.pointSizeF() - 1))
            b.setFont(f)
            b.clicked.connect(lambda _=False, s=step: self.step_font_size(s))
            tb.addWidget(b)
        tb.addSeparator()
        for a, label in ((self.a_bold, "B"), (self.a_italic, "I"), (self.a_under, "U"), (self.a_strike, "S")):
            a.setIconText(label)
            b = QToolButton()
            b.setDefaultAction(a)
            b.setToolButtonStyle(Qt.ToolButtonTextOnly)
            f = b.font()
            f.setBold(label == "B")
            f.setItalic(label == "I")
            f.setUnderline(label == "U")
            f.setStrikeOut(label == "S")
            b.setFont(f)
            b.setFixedWidth(28)
            tb.addWidget(b)
        self.text_color_btn = QToolButton()
        self.text_color_btn.setToolTip("Text colour (click the arrow to choose)")
        self.text_color_btn.setPopupMode(QToolButton.MenuButtonPopup)
        self.text_color_btn.clicked.connect(lambda: self.apply_text_color(self.last_text_color))
        m = QMenu(self)
        m.aboutToShow.connect(lambda: (m.close(), QTimer.singleShot(0, self.choose_text_color)))
        self.text_color_btn.setMenu(m)
        tb.addWidget(self.text_color_btn)
        self.fill_color_btn = QToolButton()
        self.fill_color_btn.setToolTip("Fill colour (click the arrow to choose)")
        self.fill_color_btn.setPopupMode(QToolButton.MenuButtonPopup)
        self.fill_color_btn.clicked.connect(lambda: self.apply_fill(self.last_fill_color or None))
        m2 = QMenu(self)
        m2.aboutToShow.connect(lambda: (m2.close(), QTimer.singleShot(0, self.choose_fill_color)))
        self.fill_color_btn.setMenu(m2)
        tb.addWidget(self.fill_color_btn)
        tb.addSeparator()
        self.align_buttons = []
        for a in list(self.halign.values()) + [None] + list(self.valign.values()):
            if a is None:
                tb.addSeparator()
                continue
            tb.addAction(a)
        tb.addSeparator()
        tb.addAction(self.a_wrap)
        tb.addAction(self.a_merge)
        self.border_btn = QToolButton()
        self.border_btn.setToolTip("Borders")
        self.border_btn.setPopupMode(QToolButton.InstantPopup)
        bm = QMenu(self)
        bm.addAction("All borders", lambda: self.apply_borders("all"))
        bm.addAction("Outside border", lambda: self.apply_borders("outside"))
        bm.addSeparator()
        bm.addAction("Top border", lambda: self.apply_borders("top"))
        bm.addAction("Bottom border", lambda: self.apply_borders("bottom"))
        bm.addAction("Left border", lambda: self.apply_borders("left"))
        bm.addAction("Right border", lambda: self.apply_borders("right"))
        bm.addSeparator()
        bm.addAction("No borders", lambda: self.apply_borders("none"))
        bm.addSeparator()
        bm.addAction("Border colour…", self.choose_border_color)
        self.border_btn.setMenu(bm)
        tb.addWidget(self.border_btn)
        self.border_color_btn = QToolButton()
        self.border_color_btn.setToolTip("Border colour: new borders use it, and choosing one recolours the borders "
                                         "already on the selected cells")
        self.border_color_btn.clicked.connect(self.choose_border_color)
        tb.addWidget(self.border_color_btn)
        tb.addSeparator()
        self.style_box = QComboBox()
        self.style_box.setMinimumWidth(150)
        self.style_box.setToolTip("Cell styles")
        self.style_box.activated.connect(self.style_chosen)
        tb.addWidget(self.style_box)
        self.spacing_btn = QToolButton()
        self.spacing_btn.setToolTip("Margins and spacing for this sheet")
        self.spacing_btn.clicked.connect(self.edit_sheet_layout)
        tb.addWidget(self.spacing_btn)
        spacer = QWidget()
        spacer.setStyleSheet("background: transparent;")
        spacer.setSizePolicy(spacer.sizePolicy().horizontalPolicy().Expanding, spacer.sizePolicy().verticalPolicy())
        tb.addWidget(spacer)
        self.layers_btn = QToolButton()
        self.layers_btn.setText("Layers")
        self.layers_btn.setCheckable(True)
        self.layers_btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.layers_btn.setPopupMode(QToolButton.MenuButtonPopup)
        self.layers_btn.setToolTip("Show layers on maps (the Layers tab). The arrow has New layer (Ctrl+L).")
        self.layers_btn.clicked.connect(lambda on: self.show_layers_tab() if on else self.hide_layers())
        self.layers_btn.setToolTip("Layers on or off: creatures, plants, characters, paths… on the map. "
                                   "The arrow has New layer (Ctrl+L) and Draw as.")
        lm = QMenu(self)
        lm.addAction(self.a_layer_here)
        lm.addAction("Done painting", lambda: self.stop_painting())
        lm.addSeparator()
        self.style_menu = lm.addMenu("Draw the selected layer as")
        for key, label in (("area", "Area (fills squares)"), ("path", "Path (a line: roads, tracks, rivers)"),
                           ("marker", "Marker (a dot)")):
            a = self.style_menu.addAction(label, lambda k=key: self.set_layer_style(k))
            a.setCheckable(True)
            a.setData(key)
        lm.aboutToShow.connect(self.sync_style_menu)
        self.layers_btn.setMenu(lm)
        tb.addWidget(self.layers_btn)
        self.row_mode_btn = QToolButton()
        self.row_mode_btn.setCheckable(True)
        self.row_mode_btn.setText("Compact rows")
        self.row_mode_btn.setToolTip("Cap tall rows at three lines; the full text is in the Cell panel")
        self.row_mode_btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.row_mode_btn.toggled.connect(self.set_row_mode)
        tb.addWidget(self.row_mode_btn)
        self.panel_btn = QToolButton()
        self.panel_btn.setCheckable(True)
        self.panel_btn.setToolTip("Show or hide the Cell and Layers panel (Ctrl+\\). Drag its title bar to float it.")
        self.panel_btn.clicked.connect(lambda on: (self.set_panel(on), setattr(self, "_panel_auto", False)))
        tb.addWidget(self.panel_btn)

    def build_menus(self):
        mb = self.menuBar()
        f = mb.addMenu("&File")
        for a in (self.a_new, self.a_open):
            f.addAction(a)
        self.recent_menu = f.addMenu("Open recent")
        self.recent_menu.aboutToShow.connect(self.fill_recent)
        f.addSeparator()
        for a in (self.a_save, self.a_save_as, self.a_close):
            f.addAction(a)
        f.addAction("What came over from Excel…", self.show_import_report)
        f.addSeparator()
        f.addAction(self.a_quit)

        e = mb.addMenu("&Edit")
        e.addAction(self.a_undo)
        e.addAction(self.a_redo)
        e.addSeparator()
        e.addAction(self.act("Cut", lambda: self.grid.copy(cut=True), QKeySequence.Cut))
        e.addAction(self.act("Copy", self.grid.copy, QKeySequence.Copy))
        e.addAction(self.act("Paste", self.grid.paste, QKeySequence.Paste))
        e.addAction(self.act("Fill down", lambda: self.grid.fill("down"), "Ctrl+D"))
        e.addAction(self.act("Fill right", lambda: self.grid.fill("right"), "Ctrl+R"))
        e.addAction(self.act("Clear contents", self.grid.clear_contents))
        e.addAction(self.a_note)
        e.addSeparator()
        e.addAction(self.a_find)
        for a in e.actions():
            if a.shortcut() in (QKeySequence(QKeySequence.Cut), QKeySequence(QKeySequence.Copy),
                                QKeySequence(QKeySequence.Paste), QKeySequence("Ctrl+D"), QKeySequence("Ctrl+R")):
                a.setShortcutContext(Qt.WidgetShortcut)    # the grid handles these keys itself

        v = mb.addMenu("&View")
        v.addAction(self.a_goto)
        self.a_panel = self.act("Cell and Layers panel", lambda on: (self.set_panel(on),
                                setattr(self, "_panel_auto", False)), "Ctrl+\\", True)
        self.a_panel.setChecked(True)
        v.addAction(self.a_panel)
        self.a_sidebar = self.act("Sheet list", self.set_sidebar, "Ctrl+Shift+\\", True)
        self.a_sidebar.setChecked(True)
        v.addAction(self.a_sidebar)
        v.addAction(self.act("Put the panel back on the right", self.redock_panel))
        v.addAction(self.a_back)
        v.addAction(self.a_fwd)
        v.addSeparator()
        self.a_compact = self.act("Compact rows", lambda on: self.row_mode_btn.setChecked(on), None, True)
        v.addAction(self.a_compact)
        self.a_headers = self.act("Row and column headings", self.toggle_headers, None, True)
        self.a_headers.setChecked(True)
        v.addAction(self.a_headers)
        self.a_adapt = self.act("Lighten dark text on the dark theme", self.toggle_adapt, None, True)
        self.a_adapt.setChecked(self.grid.adapt_colors)
        self.a_adapt.setToolTip("Display only: dark text written for a white page is shown lighter. The file isn't changed.")
        v.addAction(self.a_adapt)
        fz = v.addMenu("Freeze panes")
        fz.addAction(self.act("Freeze above and left of the selected cell", lambda: self.set_freeze("here")))
        fz.addAction(self.act("Freeze top row", lambda: self.set_freeze("row")))
        fz.addAction(self.act("Freeze first column", lambda: self.set_freeze("col")))
        fz.addSeparator()
        fz.addAction(self.act("Unfreeze", lambda: self.set_freeze(None)))
        self.a_show_ranges_sheet = self.act("Show the Ranges data sheet", self.toggle_ranges_sheet, None, True)
        v.addAction(self.a_show_ranges_sheet)
        v.addSeparator()
        v.addAction(self.act("Zoom in", lambda: self.grid.set_zoom(self.grid.zoom * 1.1), "Ctrl+="))
        v.addAction(self.act("Zoom out", lambda: self.grid.set_zoom(self.grid.zoom / 1.1), "Ctrl+-"))
        v.addAction(self.act("Actual size", lambda: self.grid.set_zoom(1.0), "Ctrl+0"))
        for a in v.actions()[-3:]:
            a.setShortcutContext(Qt.WidgetShortcut)
        v.addSeparator()
        tm = v.addMenu("Theme")
        group = QActionGroup(self)
        for name in THEMES:
            a = self.act(name, lambda _=False, n=name: self.set_theme(n), None, True)
            a.setChecked(name == self.theme_name)
            group.addAction(a)
            tm.addAction(a)

        ins = mb.addMenu("&Insert")
        ins.addAction(self.act("Rows above", lambda: self.insert_rows(above=True), "Ctrl+Shift+="))
        ins.addAction(self.act("Rows below", lambda: self.insert_rows(above=False)))
        ins.addAction(self.act("Columns left", lambda: self.insert_cols(left=True)))
        ins.addAction(self.act("Columns right", lambda: self.insert_cols(left=False)))
        ins.addSeparator()
        ins.addAction(self.act("Sheet", self.new_sheet, "Shift+F11"))
        ins.addAction(self.act("Group", self.new_group))
        ins.addSeparator()
        ins.addAction(self.a_layer_here)
        dl = mb.addMenu("&Delete")
        dl.addAction(self.act("Rows", self.delete_rows, "Ctrl+-"))
        dl.actions()[-1].setShortcutContext(Qt.WidgetShortcut)
        dl.addAction(self.act("Columns", self.delete_cols))
        dl.addAction(self.act("Sheet…", lambda: self.delete_sheet(self.current_ws())))

        fm = mb.addMenu("F&ormat")
        for a in (self.a_bold, self.a_italic, self.a_under, self.a_strike, self.a_wrap, self.a_merge):
            fm.addAction(a)
        fm.addAction(self.act("Text colour…", self.choose_text_color))
        fm.addAction(self.act("Fill colour…", self.choose_fill_color))
        fm.addAction(self.act("Apply the picked fill", lambda: self.apply_fill(self.last_fill_color or None), "Ctrl+Shift+F",
                              tip="Alt+click a cell to pick up its fill, then select cells and press Ctrl+Shift+F"))
        fm.addAction(self.a_clear_fmt)
        fm.addSeparator()
        fm.addAction(self.act("Sheet margins and spacing…", self.edit_sheet_layout))
        fm.addAction(self.act("Link squares to a map…", self.link_selection_dialog))
        fm.addAction(self.act("Remove links (keep the text)", self.remove_links))
        fm.addSeparator()
        fm.addAction(self.act("Palette…", self.edit_palette))
        fm.addAction(self.act("Swap colours for your palette…", self.remap_colors))
        fm.addSeparator()
        fm.addAction(self.act("New style from this cell…", self.new_style))
        self.a_update_style = self.act("Update style to match this cell", self.update_style)
        fm.addAction(self.a_update_style)
        fm.addAction(self.act("Rename a style…", self.rename_style))
        fm.addAction(self.act("Make styles from repeated formatting…", self.suggest_styles))

        h = mb.addMenu("&Help")
        h.addAction(self.act("Keyboard shortcuts", self.show_shortcuts, "F1"))
        h.addAction(self.act(f"About {APP_NAME}", self.show_about))

    # ------------------------------------------------------------ theme
    def apply_theme(self):
        t = self.theme
        QApplication.instance().setStyleSheet(stylesheet(t))
        self.grid.set_theme(t)
        for k, a in self.halign.items():
            a.setIcon(draw_icon(k, t["text"]))
        for k, a in self.valign.items():
            a.setIcon(draw_icon({"top": "top", "center": "middle", "bottom": "bottom"}[k], t["text"]))
        self.a_wrap.setIcon(draw_icon("wrap", t["text"]))
        self.a_merge.setIcon(draw_icon("merge", t["text"]))
        self.a_clear_fmt.setIcon(draw_icon("clear", t["text"]))
        self.border_btn.setIcon(draw_icon("borders", t["text"]))
        self.spacing_btn.setIcon(draw_icon("spacing", t["text"]))
        self.sidebar_btn.setIcon(draw_icon("pane_left", t["text"]))
        self.panel_btn.setIcon(draw_icon("pane_right", t["text"]))
        self.update_color_buttons()
        if self.book is not None:
            self.rebuild_sidebar()
            self.inspector.load()
            self.ranges_panel.refresh()

    def update_color_buttons(self):
        t = self.theme
        self.text_color_btn.setIcon(draw_icon("text", t["text"], self.last_text_color))
        self.fill_color_btn.setIcon(draw_icon("fill", t["text"], self.last_fill_color or None))
        self.border_color_btn.setIcon(draw_icon("bordercolor", t["text"], getattr(self, "border_color", None)))

    def set_theme(self, name):
        self.theme_name = name
        self.theme = THEMES[name]
        self.settings.setValue("theme", name)
        self.apply_theme()

    # ------------------------------------------------------------ files
    def new_book(self):
        if not self.maybe_save():
            return
        book = Book()
        book.wb.active.title = "Sheet1"
        self.set_book(book)

    def open_dialog(self):
        start = os.path.dirname(self.book.path) if self.book and self.book.path else \
            self.settings.value("last_dir", os.path.expanduser("~"))
        path, _ = QFileDialog.getOpenFileName(self, "Open workbook", start, XLSX_FILTER)
        if path:
            self.open_path(path)

    def open_path(self, path):
        if not self.maybe_save():
            return False
        try:
            book = Book.open(path)
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Couldn't open {os.path.basename(path)}:\n{e}")
            return False
        self.settings.setValue("last_dir", os.path.dirname(path))
        recent = [p for p in (self.settings.value("recent") or []) if p != path]
        self.settings.setValue("recent", [path] + recent[:9])
        self.set_book(book)
        if book.warnings:
            self.statusBar().showMessage("Some things in this file can't be shown. File → What came over from Excel.", 8000)
        return True

    def set_book(self, book):
        self.book = book
        book.listeners.append(self.on_book_changed)
        book.ranges = ranges.Ranges.load(book)
        self.squares = squares.SquareIndex(book, skip_sheet=self.read_only_sheet)
        self.prepare_meta(book)
        self.history, self.future = [], []
        self.range_state = ranges.ViewState()
        self.ranges_panel.st = self.range_state
        self.clip.clear()
        self.rebuild_sidebar()
        sheets = self.sidebar_order()
        first = book.meta.get("last_sheet")
        if not first or book.sheet(first) is None:
            active = book.wb.active
            first = active.title if active is not None and active in sheets else (sheets[0].title if sheets else None)
        self.ws_title = None
        if first:
            self.open_sheet(first, record=False)
        self.update_title()
        self.update_enabled()
        self.update_style_box()
        QTimer.singleShot(600, self.font_box.prewarm)
        self.layers_btn.setChecked(self.layers_on)

    def prepare_meta(self, book):
        """First time a workbook comes into Atlas: find map grids and put the
        maps in their own group. Nothing is saved until you save."""
        if book.meta.get("scanned"):
            return
        found = []
        for ws in book.user_sheets():
            if ws.title in book.meta["maps"]:
                continue
            info = ranges.detect_map(ws)
            if info:
                book.meta["maps"][ws.title] = info
                found.append(ws.title)
        if found and not book.meta["groups"] and len(book.user_sheets()) > 6:
            book.meta["groups"].append({"name": "Maps", "sheets": found, "collapsed": False})
        book.meta["scanned"] = True

    def close_book(self):
        if not self.maybe_save():
            return
        self.book = None
        self.ws_title = None
        self.grid.ws = None
        self.grid.book = None
        self.grid.viewport().update()
        self.sidebar.tree.clear()
        self.inspector.load()
        self.ranges_panel.refresh()
        self.update_title()
        self.update_enabled()

    def save(self):
        if self.book is None:
            return False
        if self.book.path is None or self.book.must_save_as:
            if self.book.must_save_as:
                QMessageBox.information(self, APP_NAME, "This file has things Atlas can't keep (see File → What came "
                                        "over from Excel), so it'll save as a new file and leave the original alone.")
            return self.save_as()
        return self._save_to(self.book.path)

    def save_as(self):
        if self.book is None:
            return False
        start = self.book.path or os.path.join(self.settings.value("last_dir", os.path.expanduser("~")), "Untitled.xlsx")
        if self.book.must_save_as and self.book.path:
            root, ext = os.path.splitext(self.book.path)
            start = root + " (Atlas)" + ext
        path, _ = QFileDialog.getSaveFileName(self, "Save workbook", start, "Excel workbook (*.xlsx)")
        if not path:
            return False
        if not path.lower().endswith((".xlsx", ".xlsm")):
            path += ".xlsx"
        return self._save_to(path)

    def _save_to(self, path):
        self.grid.commit_if_editing()
        self.inspector.commit_text()
        self.inspector.commit_note()
        if self.current_ws() is not None:
            self.book.meta["last_sheet"] = self.current_ws().title
        try:
            self.book.save(path)
        except PermissionError:
            QMessageBox.warning(self, APP_NAME, f"Couldn't save {os.path.basename(path)}: it's open somewhere else "
                                "(Excel, maybe). Close it there and save again.")
            return False
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Couldn't save:\n{e}")
            return False
        recent = [p for p in (self.settings.value("recent") or []) if p != path]
        self.settings.setValue("recent", [path] + recent[:9])
        self.statusBar().showMessage(f"Saved {os.path.basename(path)}", 4000)
        self.update_title()
        self.rebuild_sidebar()
        return True

    def maybe_save(self):
        if self.book is None or not self.book.dirty:
            return True
        name = os.path.basename(self.book.path) if self.book.path else "this workbook"
        r = QMessageBox.question(self, APP_NAME, f"Save changes to {name}?",
                                 QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
        if r == QMessageBox.Save:
            return self.save()
        return r == QMessageBox.Discard

    def fill_recent(self):
        self.recent_menu.clear()
        for p in self.settings.value("recent") or []:
            self.recent_menu.addAction(p, lambda _=False, p=p: self.open_path(p))
        if self.recent_menu.isEmpty():
            a = self.recent_menu.addAction("Nothing yet")
            a.setEnabled(False)

    def show_import_report(self):
        if self.book is None:
            return
        lines = self.book.warnings or ["Everything in this file came over: values, formulas, fonts, colours, "
                                       "borders, wrapping, alignment, merged cells, sizes, frozen panes and notes."]
        lines.append("Print settings (orientation, margins) are kept in the file but Atlas doesn't use them.")
        QMessageBox.information(self, "What came over from Excel", "\n\n".join(lines))

    def closeEvent(self, e):
        if not self.maybe_save():
            e.ignore()
            return
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("splitter2", self.splitter.sizes())
        self.settings.setValue("window_state", self.saveState())
        e.accept()

    def update_title(self):
        if self.book is None:
            self.setWindowTitle(APP_NAME)
            return
        name = os.path.basename(self.book.path) if self.book.path else "Untitled"
        self.setWindowTitle(f"{'• ' if self.book.dirty else ''}{name} — {APP_NAME}")

    def update_enabled(self):
        has = self.book is not None
        for a in (self.a_save, self.a_save_as, self.a_close, self.a_find, self.a_goto):
            a.setEnabled(has)
        self.a_undo.setEnabled(has and bool(self.book.undo_stack))
        self.a_redo.setEnabled(has and bool(self.book.redo_stack))
        self.a_back.setEnabled(bool(self.history))
        self.a_fwd.setEnabled(bool(self.future))
        self.toolbar.setEnabled(has)

    # ------------------------------------------------------------ sheets & sidebar
    def current_ws(self):
        if self.book is None or self.ws_title is None:
            return None
        return self.book.sheet(self.ws_title)

    def read_only_sheet(self, ws):
        return ws is not None and self.book is not None and ws.title == self.book.meta.get("ranges_sheet")

    def hidden_from_sidebar(self, ws):
        if self.read_only_sheet(ws) and not self.show_ranges_sheet:
            return True
        return False

    def sidebar_order(self):
        """User sheets in the order the sidebar shows them."""
        book = self.book
        if book is None:
            return []
        sheets = [ws for ws in book.user_sheets() if not self.hidden_from_sidebar(ws)]
        by_title = {ws.title: ws for ws in sheets}
        out, seen = [], set()
        for g in book.meta["groups"]:
            for t in g["sheets"]:
                if t in by_title and t not in seen:
                    out.append(by_title[t])
                    seen.add(t)
        out += [ws for ws in sheets if ws.title not in seen]
        return out

    def is_group_collapsed(self, name):
        for g in self.book.meta["groups"]:
            if g["name"] == name:
                return g.get("collapsed", False)
        return False

    def rebuild_sidebar(self):
        tree = self.sidebar.tree
        tree.blockSignals(True)
        tree.clear()
        if self.book is None:
            tree.blockSignals(False)
            return
        t = self.theme
        book = self.book
        sheets = [ws for ws in book.user_sheets() if not self.hidden_from_sidebar(ws)]
        titles = {ws.title for ws in sheets}
        grouped = set()
        gfont = QFont(tree.font())
        gfont.setBold(True)

        def sheet_item(ws):
            label = ws.title
            it = QTreeWidgetItem([label])
            it.setData(0, Qt.UserRole, ("sheet", ws.title))
            it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled)
            if ws.title in book.meta["maps"]:
                it.setIcon(0, map_icon(t["accent"]))
                it.setToolTip(0, f"{ws.title} — map")
            if ws.sheet_state == "hidden":
                it.setForeground(0, QColor(t["faint"]))
                it.setText(0, label + "  (hidden in Excel)")
            if self.read_only_sheet(ws):
                it.setForeground(0, QColor(t["faint"]))
                it.setText(0, label + "  (written by Atlas)")
            return it

        for g in book.meta["groups"]:
            gi = QTreeWidgetItem([g["name"]])
            gi.setData(0, Qt.UserRole, ("group", g["name"]))
            gi.setFlags(Qt.ItemIsEnabled | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled)
            gi.setFont(0, gfont)
            gi.setForeground(0, QColor(t["group"]))
            tree.addTopLevelItem(gi)
            for title in g["sheets"]:
                if title in titles and title not in grouped:
                    gi.addChild(sheet_item(book.sheet(title)))
                    grouped.add(title)
            gi.setExpanded(not g.get("collapsed", False))
        for ws in sheets:
            if ws.title not in grouped:
                tree.addTopLevelItem(sheet_item(ws))
        tree.blockSignals(False)
        if self.ws_title:
            self.sidebar.select_sheet(self.ws_title)
        self.sidebar.apply_filter(self.sidebar.filter.text())

    def sidebar_reordered(self, dragged):
        """Read groups and order back from the tree after a drag."""
        tree = self.sidebar.tree
        groups, order = [], []
        old = {g["name"]: g for g in self.book.meta["groups"]}
        for i in range(tree.topLevelItemCount()):
            top = tree.topLevelItem(i)
            kind, name = top.data(0, Qt.UserRole)
            if kind == "group":
                members = [top.child(j).data(0, Qt.UserRole)[1] for j in range(top.childCount())]
                groups.append({"name": name, "sheets": members, "collapsed": old.get(name, {}).get("collapsed", False)})
                order += members
            else:
                order.append(name)
        self.book.begin([])
        self.book.meta["groups"] = groups
        # Excel's tab order follows the sidebar
        wb = self.book.wb
        rest = [ws for ws in wb._sheets if ws.title not in order]
        wb._sheets = [wb[t] for t in order if t in wb.sheetnames] + rest
        self.book.done("groups")

    def group_collapsed(self, item, collapsed):
        if self.book is None or self.sidebar.filter.text():
            return
        kind, name = item.data(0, Qt.UserRole)
        for g in self.book.meta["groups"]:
            if g["name"] == name:
                g["collapsed"] = collapsed

    def focus_filter(self):
        self.sidebar.filter.setFocus()
        self.sidebar.filter.selectAll()

    def open_sheet(self, title, cur=None, record=True):
        if self.book is None or self.book.sheet(title) is None:
            return
        self.grid.commit_if_editing()
        if record and self.ws_title is not None and (self.ws_title != title or cur is not None):
            self.history.append((self.ws_title, self.grid.cur))
            del self.history[:-100]
            self.future.clear()
        same = self.ws_title == title
        if not same:
            self.stop_painting(quiet=True)
        self.ws_title = title
        if same and cur is not None:
            self.grid.set_current(*cur)
        elif not same:
            self.grid.set_sheet(self.book, self.book.sheet(title), cur)
        self.sidebar.select_sheet(title)
        self.grid.ring_selection = title in self.book.meta["maps"]
        self.card.hide()
        if self.find_bar.isVisible():
            self.find_bar.update_marks()
        self.ranges_panel.refresh()
        self.update_enabled()

    # ------------------------------------------------------------ notes
    def _note_dialog(self, title, label, text):
        from PySide6.QtWidgets import QDialog, QDialogButtonBox
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        dlg.resize(460, 280)
        lay = QVBoxLayout(dlg)
        lbl = QLabel(label)
        lbl.setObjectName("muted")
        lbl.setWordWrap(True)
        lay.addWidget(lbl)
        edit = QPlainTextEdit(text)
        lay.addWidget(edit, 1)
        box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        box.accepted.connect(dlg.accept)
        box.rejected.connect(dlg.reject)
        lay.addWidget(box)
        edit.setFocus()
        return edit.toPlainText().strip() if dlg.exec() else None

    def edit_cell_note(self):
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        r, c = self.grid.cur
        cell = ws._cells.get((r, c))
        old = cell.comment.text if cell is not None and cell.comment else ""
        new = self._note_dialog(f"Note · {fx.num_to_col(c)}{r}",
                                "A note on this cell. Cells with a note get a small corner mark; the note shows in "
                                "the Cell panel. (Excel keeps these too.)", old)
        if new is None or new == old.strip():
            return
        self.book.begin([ws])
        cell = ws.cell(r, c)
        cell.comment = Comment(new, "") if new else None
        if cell.comment is not None:
            cell.comment.width, cell.comment.height = 300, 150
        self.book.done()

    def edit_square_note(self, map_title, sq):
        label = ranges.square_label(sq)
        old = self.book.meta.get("square_notes", {}).get(map_title, {}).get(label, "")
        new = self._note_dialog(f"Square {label} · {map_title}",
                                f"What's at {label}? Shows on the square's hover card and in the Cell panel. "
                                "Kept by Atlas (Excel has nowhere to put it).", old)
        if new is None or new == old:
            return
        self.book.begin([])
        notes = self.book.meta.setdefault("square_notes", {}).setdefault(map_title, {})
        if new:
            notes[label] = new
        else:
            notes.pop(label, None)
        self.book.done("notes")

    # ------------------------------------------------------------ painting
    def layers_showing(self):
        return bool(self.layers_on)

    def set_layers_on(self, on):
        self.layers_on = bool(on)
        self.settings.setValue("layers_on", "true" if on else "false")
        if not on:
            self.stop_painting(quiet=True)
        self.layers_btn.setChecked(self.layers_on)
        self.grid.viewport().update()

    def panel_open(self):
        return self.dock.isVisible()

    def fit_panel_width(self):
        """The panel is never narrower than its contents need."""
        need = max(self.ranges_panel.body.minimumSizeHint().width(),
                   self.ranges_panel.minimumSizeHint().width(),
                   self.inspector.minimumSizeHint().width()) + 30      # + a scrollbar
        self.right.setMinimumWidth(need)
        return need

    def set_panel(self, show):
        if show and not self.dock.isVisible():
            self.dock.show()
            if not self.dock.isFloating():
                self.resizeDocks([self.dock], [self.fit_panel_width() + 20], Qt.Horizontal)
        elif not show and self.dock.isVisible():
            self.dock.hide()
        self.sync_pane_buttons()
        self.grid.viewport().update()

    def set_sidebar(self, show):
        sizes = self.splitter.sizes()
        if show and sizes[0] == 0:
            self.splitter.setSizes([230, max(300, sizes[1] - 230)])
        elif not show and sizes[0] > 0:
            self.splitter.setSizes([0, sizes[0] + sizes[1]])
        self.sync_pane_buttons()

    def sync_pane_buttons(self):
        if not hasattr(self, "panel_btn"):
            return
        for btn, on in ((self.sidebar_btn, self.splitter.sizes()[0] > 0), (self.panel_btn, self.panel_open())):
            btn.blockSignals(True)
            btn.setChecked(on)
            btn.blockSignals(False)
        self.a_panel.setChecked(self.panel_open())
        self.a_sidebar.setChecked(self.splitter.sizes()[0] > 0)
        self.layers_btn.setChecked(self.layers_showing())

    def _dock_visibility(self, *_):
        try:
            if self.isVisible():
                self.right_panel_moved()
        except RuntimeError:        # the window is being torn down
            pass

    def right_panel_moved(self):
        if not self.panel_open():
            self._panel_auto = False
            self.stop_painting(quiet=True)
        self.sync_pane_buttons()
        self.grid.viewport().update()

    def redock_panel(self):
        self.dock.setFloating(False)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock)
        self.set_panel(True)

    def set_path_end(self, map_title, sq, layer_name, mode):
        self.book.begin([])
        ends = self.book.meta.setdefault("path_ends", {}).setdefault(layer_name, {}).setdefault(map_title, {})
        ends[ranges.square_label(sq)] = mode
        self.book.done("ranges")

    def show_layers_tab(self):
        """Layers on: draw them, open the Layers section, and slide the panel
        open if it's shut."""
        if not self.panel_open():
            self._panel_auto = True
            self.set_panel(True)
        self.layers_section.expand(True)
        self.set_layers_on(True)
        self.ranges_panel.refresh()
        self.fit_panel_width()
        QTimer.singleShot(0, lambda: self.right.ensureWidgetVisible(self.layers_section.header))

    def hide_layers(self):
        self.set_layers_on(False)
        if self._panel_auto:
            self._panel_auto = False
            self.set_panel(False)

    def sync_style_menu(self):
        layer = self.book.ranges.get(self.range_state.layer) if (self.book and self.range_state.layer) else None
        self.style_menu.setEnabled(layer is not None)
        self.style_menu.setTitle(f"Draw {layer.name} as" if layer else "Draw the selected layer as")
        for a in self.style_menu.actions():
            a.setChecked(layer is not None and a.data() == layer.style)

    def set_layer_style(self, key):
        layer = self.book.ranges.get(self.range_state.layer) if (self.book and self.range_state.layer) else None
        if layer is None or layer.style == key:
            return
        self.book.begin([])
        self.book.ranges.get(layer.name).style = key
        self.book.done("ranges")
        self.show_layers_tab()

    def new_layer_here(self):
        """Ctrl+L. On a map: put the selected squares in a layer. Elsewhere:
        make a layer for the selected name."""
        ws = self.current_ws()
        if ws is None or self.book is None:
            return
        info = self.book.meta["maps"].get(ws.title)
        if not info:
            cell = self.current_cell()
            name = cell.value.strip() if cell is not None and isinstance(cell.value, str) else ""
            if name and "\n" not in name and len(name) <= 60 and not fx.is_formula(name):
                existing = self.book.ranges.get(name)
                if existing is not None:
                    maps = self.map_titles()
                    if maps:
                        self.show_range(existing.name, maps[0])
                    return
                self.make_layer_for(name)
            else:
                self.statusBar().showMessage("Ctrl+L: select squares on a map, or a name (e.g. on Fauna) "
                                             "to make a layer for it.", 6000)
            return
        r1, c1, r2, c2 = self.grid.selection()
        squares_sel = sorted({sq for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)
                              for sq in [ranges.map_square(info, r, c)] if sq})
        if not squares_sel:
            self.statusBar().showMessage("Ctrl+L: select one or more squares on the map first.", 6000)
            return
        names = [l.name for l in self.book.ranges.layers]
        where = ranges.square_label(squares_sel[0]) + (f" and {len(squares_sel) - 1} more" if len(squares_sel) > 1 else "")
        st = self.range_state
        name, ok = QInputDialog.getItem(
            self, "Layer on these squares",
            f"Put {where} in a layer as {ranges.LEVELS.get(st.level or 3).upper()}, "
            f"{ranges.times_label(st.blocks).upper()}.\n(Those come from Time and Paint on the Layers tab; "
            "change any square afterwards in the Here table there.)\n\nType a new name or pick an existing layer:",
            names, -1 if not names else 0, True)
        name = (name or "").strip()
        if not ok or not name:
            return
        self.book.begin([])
        layer = self.book.ranges.get(name) or self.book.ranges.add(name)
        self.book.ranges.paint(layer, ws.title, st.blocks, squares_sel, st.level or 3)
        self.book.done("ranges")
        st.layer = layer.name
        self.show_layers_tab()
        self.ranges_panel.refresh()
        self.statusBar().showMessage(f"{len(squares_sel)} square{'s' if len(squares_sel) != 1 else ''} in {layer.name}. "
                                     "Ctrl+Z undoes it; turn on Paint on the Layers tab to keep going.", 6000)

    def stop_painting(self, quiet=False):
        st = self.range_state
        if not st.painting:
            return
        st.painting = False
        try:
            self.ranges_panel.paint_btn.blockSignals(True)
            self.ranges_panel.paint_btn.setChecked(False)
            self.ranges_panel.paint_btn.blockSignals(False)
            self.ranges_panel.update_notes()
        except RuntimeError:            # the window is closing
            return
        self.update_paint_bar()
        if not quiet:
            self.statusBar().showMessage("Stopped painting. Ctrl+Z undoes strokes one at a time.", 5000)

    def update_paint_bar(self):
        st = self.range_state
        ws = self.current_ws()
        on = bool(st.painting and st.layer and ws is not None and self.book is not None
                  and ws.title in self.book.meta["maps"] and self.layers_showing())
        if on and self.ranges_panel.timeless_selected():
            what = "Erasing" if st.level == 0 else "Painting"
            self.paint_label.setText(f"<b>{what} · {st.layer}</b> (a path or marker: no times). "
                                     "Click or drag squares; Shift-drag erases; Ctrl+Z undoes a stroke; Esc stops.")
        elif on:
            what = "Erasing" if st.level == 0 else f"Painting {ranges.LEVELS[st.level]}"
            self.paint_label.setText(f"<b>{what} · {st.layer}</b> · {ranges.times_label(st.blocks)}. "
                                     "Click or drag squares; Shift-drag erases; Ctrl+Z undoes a stroke; Esc stops.")
        self.paint_bar.setVisible(on)

    # ------------------------------------------------------------ square cards
    def on_hover(self, row, col, gpos):
        ws = self.current_ws()
        info = self.book.meta["maps"].get(ws.title) if (ws is not None and self.book is not None) else None
        sq = ranges.map_square(info, row, col) if (info and row and col) else None
        key = (ws.title, sq) if sq else None
        if key is None:
            self.card_timer_show.stop()
            if self.card.isVisible():
                self.card_timer_hide.start()
            self._hover = None
            return
        if self.card.isVisible() and self.card.key == key:
            self.card_timer_hide.stop()
            return
        self._hover = (key, gpos)
        if self.card.isVisible():
            self.show_card()                 # already showing: follow the mouse right away
        else:
            self.card_timer_show.start()

    def show_card(self):
        if self._hover is None or self.book is None or self.grid.tool.active() and self.grid._drag:
            return
        (map_title, sq), gpos = self._hover
        if not self.squares.has_content(map_title, sq):
            self.card.hide()
            return
        cd = self.squares.card(map_title, sq)
        self.card.setStyleSheet(f"QFrame#card {{ background: {self.theme['panel']}; border: 1px solid {self.theme['border']}; "
                                f"border-radius: 8px; }} QLabel {{ background: transparent; }}")
        self.card.show_for((map_title, sq), squares.card_html(cd, self.theme, map_title), gpos)

    def hide_card(self):
        from PySide6.QtGui import QCursor
        if self.card.isVisible() and self.card.geometry().contains(QCursor.pos()):
            return
        self.card.hide()
        self.card.key = None

    def go_to_href(self, link):
        target = squares.parse_href(link)
        if target is None:
            self.follow_link(link)
            return
        sheet, r, c = target
        self.go_to(sheet, r, c)

    def go_to(self, sheet, r, c):
        self.card.hide()
        self.open_sheet(sheet, (r, c))
        self.grid.flash(r, c)
        self.statusBar().showMessage("Alt+← goes back.", 4000)

    def go_back(self):
        if not self.history:
            return
        self.future.append((self.ws_title, self.grid.cur))
        title, cur = self.history.pop()
        self.open_sheet(title, cur, record=False)

    def go_forward(self):
        if not self.future:
            return
        self.history.append((self.ws_title, self.grid.cur))
        title, cur = self.future.pop()
        self.open_sheet(title, cur, record=False)

    def mousePressEvent(self, e):
        if e.button() == Qt.BackButton:
            self.go_back()
        elif e.button() == Qt.ForwardButton:
            self.go_forward()
        else:
            super().mousePressEvent(e)

    def follow_link(self, target):
        if target.startswith("#"):
            loc = target[1:]
            sheet, _, cell = loc.rpartition("!")
            sheet = sheet.strip("'").replace("''", "'") or self.ws_title
            try:
                ref = fx.Ref.parse(cell)
                self.go_to(sheet, ref.r1, ref.c1)
            except Exception:
                self.open_sheet(sheet)
        else:
            QDesktopServices.openUrl(QUrl(target))

    def sidebar_menu(self, pos):
        if self.book is None:
            return
        it = self.sidebar.tree.itemAt(pos)
        menu = QMenu(self)
        if it is not None:
            kind, name = it.data(0, Qt.UserRole)
            if kind == "sheet":
                ws = self.book.sheet(name)
                menu.addAction("Rename…", lambda: self.rename_sheet(ws))
                menu.addAction("Duplicate", lambda: self.duplicate_sheet(ws))
                menu.addAction("Margins and spacing…", lambda: (self.open_sheet(name), self.edit_sheet_layout()))
                move = menu.addMenu("Move to group")
                for g in self.book.meta["groups"]:
                    move.addAction(g["name"], lambda _=False, g=g["name"]: self.move_to_group(name, g))
                move.addSeparator()
                move.addAction("New group…", lambda: self.new_group(name))
                move.addAction("No group", lambda: self.move_to_group(name, None))
                if name in self.book.meta["maps"]:
                    menu.addAction("Stop treating as a map", lambda: self.unmap(name))
                else:
                    menu.addAction("Make this a map…", lambda: (self.open_sheet(name), self.show_layers_tab()))
                menu.addSeparator()
                menu.addAction("Delete sheet…", lambda: self.delete_sheet(ws))
            else:
                menu.addAction("Rename group…", lambda: self.rename_group(name))
                menu.addAction("Remove group (keeps its sheets)", lambda: self.remove_group(name))
            menu.addSeparator()
        menu.addAction("New sheet", self.new_sheet)
        menu.addAction("New group…", self.new_group)
        menu.exec(self.sidebar.tree.viewport().mapToGlobal(pos))

    def rename_sheet(self, ws):
        name, ok = QInputDialog.getText(self, "Rename sheet", "Name:", text=ws.title)
        name = name.strip()
        if not ok or not name or name == ws.title:
            return
        bad = set('[]:*?/\\')
        if any(ch in bad for ch in name) or len(name) > 31:
            QMessageBox.information(self, "Rename sheet", "Sheet names can't be longer than 31 characters "
                                    "or use [ ] : * ? / \\")
            return
        if name.lower() in (t.lower() for t in self.book.wb.sheetnames if t != ws.title):
            QMessageBox.information(self, "Rename sheet", f"There's already a sheet called {name}.")
            return
        was_current = ws.title == self.ws_title
        self.book.begin()
        self.book.rename_sheet(ws, name)
        if was_current:
            self.ws_title = name
        self.history = [(name if t == ws.title else t, c) for t, c in self.history]
        self.book.done("sheets")

    def duplicate_sheet(self, ws):
        self.book.begin()
        new = self.book.duplicate_sheet(ws)
        self.book.done("sheets")
        self.open_sheet(new.title)

    def delete_sheet(self, ws):
        if ws is None:
            return
        if len(self.book.user_sheets()) <= 1:
            QMessageBox.information(self, "Delete sheet", "A workbook needs at least one sheet.")
            return
        if QMessageBox.question(self, "Delete sheet", f"Delete “{ws.title}”? You can undo this.") != QMessageBox.Yes:
            return
        order = self.sidebar_order()
        idx = order.index(ws) if ws in order else 0
        self.book.begin()
        self.book.delete_sheet(ws)
        self.book.done("sheets")
        order = self.sidebar_order()
        if order:
            self.ws_title = None
            self.open_sheet(order[min(idx, len(order) - 1)].title, record=False)

    def new_sheet(self):
        if self.book is None:
            return
        self.book.begin()
        cur = self.current_ws()
        idx = self.book.wb.worksheets.index(cur) + 1 if cur is not None else None
        ws = self.book.add_sheet("Sheet", idx)
        for g in self.book.meta["groups"]:
            if cur is not None and cur.title in g["sheets"]:
                g["sheets"].insert(g["sheets"].index(cur.title) + 1, ws.title)
                break
        self.book.done("sheets")
        self.open_sheet(ws.title)
        self.rename_sheet(ws)

    def new_group(self, with_sheet=None):
        if self.book is None:
            return
        name, ok = QInputDialog.getText(self, "New group", "Group name:")
        name = name.strip()
        if not ok or not name:
            return
        if any(g["name"] == name for g in self.book.meta["groups"]):
            QMessageBox.information(self, "New group", f"There's already a group called {name}.")
            return
        self.book.begin([])
        self.book.meta["groups"].append({"name": name, "sheets": [], "collapsed": False})
        self.book.done("groups")
        if isinstance(with_sheet, str):
            self.move_to_group(with_sheet, name)

    def move_to_group(self, title, group):
        self.book.begin([])
        for g in self.book.meta["groups"]:
            g["sheets"] = [s for s in g["sheets"] if s != title]
            if g["name"] == group:
                g["sheets"].append(title)
        self.book.done("groups")

    def rename_group(self, name):
        new, ok = QInputDialog.getText(self, "Rename group", "Name:", text=name)
        new = new.strip()
        if not ok or not new or new == name:
            return
        self.book.begin([])
        for g in self.book.meta["groups"]:
            if g["name"] == name:
                g["name"] = new
        self.book.done("groups")

    def remove_group(self, name):
        self.book.begin([])
        self.book.meta["groups"] = [g for g in self.book.meta["groups"] if g["name"] != name]
        self.book.done("groups")

    def unmap(self, title):
        self.book.begin([])
        self.book.meta["maps"].pop(title, None)
        self.book.done("maps")

    def toggle_ranges_sheet(self, on):
        self.show_ranges_sheet = on
        self.rebuild_sidebar()

    # ------------------------------------------------------------ book changes
    def on_book_changed(self, kind):
        if self.book is not None and self.ws_title:
            self.grid.ring_selection = self.ws_title in self.book.meta["maps"]
        if self.current_ws() is None and self.book is not None:
            order = self.sidebar_order()
            self.ws_title = None
            if order:
                self.open_sheet(order[0].title, record=False)
        if self.current_ws() is not None and self.grid.ws is not self.current_ws():
            self.grid.set_sheet(self.book, self.current_ws(), self.grid.cur)
        elif kind in ("ranges",):
            self.grid.viewport().update()
        else:
            self.grid.content_changed()
        if kind in ("sheets", "groups", "maps", "undo", "saved", "palette"):
            self.rebuild_sidebar()
        if kind in ("undo", "maps") and self.ws_title and self.book.sheet(self.ws_title) is None:
            order = self.sidebar_order()
            if order:
                self.open_sheet(order[0].title, record=False)
        self.inspector.load()
        self.ranges_panel.refresh()
        self.update_title()
        self.update_enabled()
        self.sync_toolbar()
        self.update_style_box()
        if self.find_bar.isVisible() and kind != "ranges":
            self.find_bar.timer.start()

    def undo(self):
        if self.grid.is_editing():
            self.grid.finish_edit(False)
            return
        if self.book is not None:
            self.book.undo()

    def redo(self):
        if self.book is not None:
            self.book.redo()

    def on_current_changed(self):
        self.inspector.load()
        self.sync_toolbar()
        self.update_stats()
        if self.layers_section.expanded():
            self.ranges_panel.update_here()

    def update_stats(self):
        ws = self.current_ws()
        if ws is None:
            self.stats_label.setText("")
            return
        r1, c1, r2, c2 = self.grid.selection()
        if (r1, c1) == (r2, c2) or (r2 - r1 + 1) * (c2 - c1 + 1) > 50000:
            self.stats_label.setText("")
            return
        nums, filled = [], 0
        for (r, c), cell in ws._cells.items():
            if r1 <= r <= r2 and c1 <= c <= c2 and cell.value is not None:
                filled += 1
                v = self.book.calc.value(ws.title, r, c) if fx.is_formula(cell.value) else cell.value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    nums.append(v)
        parts = [f"{filled} filled"]
        if nums:
            parts.append(f"Sum {fx.format_general(float(sum(nums)))}")
            parts.append(f"Average {fx.format_general(sum(nums) / len(nums))}")
        self.stats_label.setText("   ".join(parts) + "   ")

    # ------------------------------------------------------------ formatting
    def selected_cells(self, create=True):
        ws = self.current_ws()
        r1, c1, r2, c2 = self.grid.selection()
        out = []
        for cell in self.book.cells_in(ws, r1, c1, r2, c2, create=create):
            if not isinstance(cell, MergedCell):
                out.append(cell)
        return out

    def format_cells(self, fn):
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        self.grid.commit_if_editing()
        self.book.begin([ws])
        for cell in self.selected_cells():
            fn(cell)
        self.book.done("format")

    def current_cell(self):
        ws = self.current_ws()
        if ws is None:
            return None
        return ws._cells.get(self.grid.cur)

    def sync_toolbar(self):
        cell = self.current_cell()
        self._syncing = True
        try:
            font = cell.font if cell is not None else None
            al = cell.alignment if cell is not None else None
            self.a_bold.setChecked(bool(font and font.b))
            self.a_italic.setChecked(bool(font and font.i))
            self.a_under.setChecked(bool(font and font.u))
            self.a_strike.setChecked(bool(font and font.strike))
            self.a_wrap.setChecked(bool(al and al.wrap_text))
            h = al.horizontal if al else None
            for k, a in self.halign.items():
                a.setChecked(h == k)
            v = (al.vertical if al else None) or "bottom"
            for k, a in self.valign.items():
                a.setChecked(v == k)
            self.font_box.show_name((font.name if font and font.name else None) or "Calibri")
            size = (font.sz if font and font.sz else None) or 11
            self.size_box.setEditText(f"{size:g}")
            ws = self.current_ws()
            self.a_merge.setChecked(bool(ws is not None and self.book.merged_at(ws, *self.grid.cur)))
            name = None
            if cell is not None and cell.has_style:
                try:
                    name = cell.style
                except Exception:
                    name = None
            idx = self.style_box.findData(name) if name and name not in BUILTIN_STYLE_NAMES else 0
            self.style_box.setCurrentIndex(max(0, idx))
            self.a_update_style.setEnabled(bool(name and name not in BUILTIN_STYLE_NAMES))
        finally:
            self._syncing = False

    def toggle_font(self, attr):
        if self._syncing or self.book is None:
            return
        cell = self.current_cell()
        cur = getattr(cell.font, attr) if cell is not None and cell.font is not None else None
        on = not bool(cur)
        value = ("single" if on else None) if attr == "u" else on
        self.format_cells(lambda c: model.set_font(c, **{attr: value}))

    def set_font_attr(self, **kw):
        if self._syncing or self.book is None:
            return
        self.format_cells(lambda c: model.set_font(c, **kw))
        self.grid.setFocus()

    def step_font_size(self, step):
        """Each selected cell's text one point bigger or smaller."""
        if self.book is None:
            return
        def fn(cell):
            size = (cell.font.sz if cell.font is not None and cell.font.sz else None) or 11
            model.set_font(cell, sz=max(1.0, min(409.0, float(round(size)) + step)))
        self.format_cells(fn)
        cell = self.current_cell()
        if cell is not None and cell.font is not None and cell.font.sz:
            self.size_box.setEditText(f"{cell.font.sz:g}")

    def size_entered(self):
        if self._syncing:
            return
        try:
            size = float(self.size_box.currentText())
        except ValueError:
            return
        if 1 <= size <= 409:
            self.set_font_attr(sz=size)

    def toggle_wrap(self):
        if self._syncing or self.book is None:
            return
        cell = self.current_cell()
        on = not bool(cell is not None and cell.alignment and cell.alignment.wrap_text)
        self.format_cells(lambda c: model.set_alignment(c, wrap_text=on))

    def set_halign(self, k):
        if self._syncing or self.book is None:
            return
        cell = self.current_cell()
        cur = cell.alignment.horizontal if cell is not None and cell.alignment else None
        val = None if cur == k else k
        self.format_cells(lambda c: model.set_alignment(c, horizontal=val))

    def set_valign(self, v):
        if self._syncing or self.book is None:
            return
        self.format_cells(lambda c: model.set_alignment(c, vertical=v))

    def toggle_merge(self):
        if self._syncing or self.book is None:
            return
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        r1, c1, r2, c2 = self.grid.selection()
        self.grid.commit_if_editing()
        self.book.begin([ws])
        if not self.book.unmerge(ws, r1, c1, r2, c2):
            if (r1, c1) == (r2, c2):
                self.book.undo_stack.pop()
                self.sync_toolbar()
                return
            lost = sum(1 for cell in self.book.cells_in(ws, r1, c1, r2, c2, create=False)
                       if cell.value is not None and (cell.row, cell.column) != (r1, c1))
            self.book.merge(ws, r1, c1, r2, c2)
            model.set_alignment(ws.cell(r1, c1), horizontal="center")
            if lost:
                self.statusBar().showMessage(f"Merged. Only the top-left cell's text is kept ({lost} other cells "
                                             "cleared); Ctrl+Z brings them back.", 8000)
        self.book.done("format")
        self.grid.select_range(r1, c1, r2, c2)

    def clear_formatting(self):
        self.format_cells(lambda c: setattr(c, "style", "Normal"))

    def choose_text_color(self):
        if self.book is None:
            return
        cell = self.current_cell()
        cur = model.resolve_color(cell.font.color, self.book.theme) if cell is not None and cell.font else None
        hex6 = ColorPicker.get(self, self.book, cur, "Text colour")
        if hex6 is None:
            return
        self.apply_text_color(hex6 or None)

    def apply_text_color(self, hex6):
        if hex6:
            self.last_text_color = hex6
            self.update_color_buttons()
        self.format_cells(lambda c: model.set_font(c, color=("FF" + hex6) if hex6 else None))

    def choose_fill_color(self):
        if self.book is None:
            return
        cell = self.current_cell()
        cur = model.resolve_color(cell.fill.fgColor, self.book.theme) if cell is not None and cell.fill and cell.fill.fill_type else None
        hex6 = ColorPicker.get(self, self.book, cur, "Fill colour")
        if hex6 is None:
            return
        self.apply_fill(hex6 or None)

    def pick_fill_from(self, r, c):
        """Alt+click: the dropper. The fill button now paints this colour."""
        ws = self.current_ws()
        cell = ws._cells.get((r, c)) if ws is not None else None
        hexv = (model.resolve_color(cell.fill.fgColor, self.book.theme)
                if cell is not None and cell.has_style and cell.fill and cell.fill.fill_type else None)
        if not hexv:
            self.last_fill_color = ""        # "no fill": the sheet's own background
            self.update_color_buttons()
            self.statusBar().showMessage("Picked no fill (the sheet's background). Select cells, then click the fill "
                                         "button or press Ctrl+Shift+F to clear their fill.", 8000)
            return
        self.last_fill_color = hexv
        self.update_color_buttons()
        name = None
        if ws.title in self.book.meta["maps"] and self.squares is not None:
            name = self.squares.legend(ws.title).get(hexv)
        self.statusBar().showMessage(f"Picked #{hexv}{f' ({name})' if name else ''}. Select cells, then click the fill "
                                     "button or press Ctrl+Shift+F.", 8000)

    def apply_fill(self, hex6):
        if hex6:
            self.last_fill_color = hex6
            self.update_color_buttons()
        hex6 = hex6 or None
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        r1, c1, r2, c2 = self.grid.selection()
        self.book.begin([ws])
        for cell in self.book.cells_in(ws, r1, c1, r2, c2):
            model.set_fill(cell, hex6)
        self.book.done("format")

    def apply_borders(self, kind):
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        r1, c1, r2, c2 = self.grid.selection()
        color = getattr(self, "border_color", None)
        color = ("FF" + color) if color else None
        self.book.begin([ws])
        for cell in self.book.cells_in(ws, r1, c1, r2, c2):
            r, c = cell.row, cell.column
            if kind == "none":
                model.set_borders(cell, False, False, False, False)
            elif kind == "all":
                model.set_borders(cell, color or True, color or True, color or True, color or True)
            elif kind == "outside":
                model.set_borders(cell, (color or True) if c == c1 else None, (color or True) if c == c2 else None,
                                  (color or True) if r == r1 else None, (color or True) if r == r2 else None)
            elif kind == "bottom" and r == r2:
                model.set_borders(cell, bottom=color or True)
            elif kind == "top" and r == r1:
                model.set_borders(cell, top=color or True)
            elif kind == "left" and c == c1:
                model.set_borders(cell, left=color or True)
            elif kind == "right" and c == c2:
                model.set_borders(cell, right=color or True)
        self.book.done("format")

    def choose_border_color(self):
        if self.book is None:
            return
        hex6 = ColorPicker.get(self, self.book, getattr(self, "border_color", None), "Border colour")
        if hex6 is None:
            return
        self.border_color = hex6 or None
        self.update_color_buttons()
        # recolour the borders already on the selection
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        r1, c1, r2, c2 = self.grid.selection()
        cells = [cell for cell in self.book.cells_in(ws, r1, c1, r2, c2, create=False)
                 if cell.border is not None and any(getattr(cell.border, s) is not None and getattr(cell.border, s).style
                                                    for s in ("left", "right", "top", "bottom"))]
        if not cells:
            return
        from openpyxl.styles import Border, Side
        color = ("FF" + self.border_color) if self.border_color else None
        self.book.begin([ws])
        for cell in cells:
            sides = {}
            for s in ("left", "right", "top", "bottom"):
                cur = getattr(cell.border, s)
                sides[s] = Side(style=cur.style, color=color) if cur is not None and cur.style else Side()
            cell.border = Border(**sides)
        self.book.done("format")
        self.statusBar().showMessage(f"Recoloured borders on {len(cells)} cells.", 4000)

    def pick_color(self, current, title):
        hex6 = ColorPicker.get(self, self.book, current, title, allow_none=False)
        return hex6 or None

    def fonts_in_use(self):
        """Font names used in the workbook, most used first."""
        if self.book is None:
            return []
        counts = {}
        for ws in self.book.user_sheets():
            for cell in ws._cells.values():
                if cell.value is None and not cell.has_style:
                    continue
                f = cell.font
                name = (f.name if f is not None and f.name else None) or "Calibri"
                counts[name] = counts.get(name, 0) + 1
        return sorted(counts, key=lambda n: -counts[n])

    def edit_sheet_layout(self):
        if self.current_ws() is not None:
            SheetLayoutDialog(self).exec()

    def edit_palette(self):
        if self.book is not None:
            PaletteDialog(self).exec()

    def remap_colors(self):
        if self.book is not None:
            RemapDialog(self).exec()

    # ------------------------------------------------------------ styles
    def user_styles(self):
        return [ns for ns in self.book.wb._named_styles if ns.name not in BUILTIN_STYLE_NAMES
                and not getattr(ns, "builtinId", None)]

    def update_style_box(self):
        self.style_box.blockSignals(True)
        self.style_box.clear()
        self.style_box.addItem("Styles…", None)
        if self.book is not None:
            for ns in self.user_styles():
                hex6 = model.resolve_color(ns.fill.fgColor, self.book.theme) if ns.fill and ns.fill.fill_type else None
                txt = model.resolve_color(ns.font.color, self.book.theme) if ns.font else None
                self.style_box.addItem(swatch_icon(hex6 or txt or self.theme["canvas"].lstrip("#")), ns.name, ns.name)
        self.style_box.blockSignals(False)
        self.sync_toolbar() if self.book is not None else None

    def style_chosen(self, idx):
        name = self.style_box.itemData(idx)
        if not name or self.book is None:
            return
        self.apply_style(name)
        self.grid.setFocus()

    def apply_style(self, name):
        def fn(cell):
            border, numfmt = copy.copy(cell.border), cell.number_format
            cell.style = name
            cell.border, cell.number_format = border, numfmt
        self.format_cells(fn)

    def new_style(self):
        cell = self.current_cell()
        if self.book is None:
            return
        name, ok = QInputDialog.getText(self, "New style", "Name for a style that looks like this cell:")
        name = name.strip()
        if not ok or not name:
            return
        if name in self.book.wb.named_styles:
            QMessageBox.information(self, "New style", f"There's already a style called {name}.")
            return
        ws = self.current_ws()
        cell = ws.cell(*self.grid.cur)
        self.book.begin()
        ns = NamedStyle(name=name)
        ns.font, ns.fill, ns.alignment = copy.copy(cell.font), copy.copy(cell.fill), copy.copy(cell.alignment)
        self.book.wb.add_named_style(ns)
        border = copy.copy(cell.border)
        cell.style = name
        cell.border = border
        self.book.done("format")
        self.update_style_box()

    def update_style(self):
        cell = self.current_cell()
        if cell is None or self.book is None:
            return
        try:
            name = cell.style
        except Exception:
            return
        ns = next((s for s in self.book.wb._named_styles if s.name == name), None)
        if ns is None:
            return
        self.book.begin()
        ns.font, ns.fill, ns.alignment = copy.copy(cell.font), copy.copy(cell.fill), copy.copy(cell.alignment)
        n = self.reapply_style(name)
        self.book.done("format")
        self.statusBar().showMessage(f"Updated “{name}” in {n} cells.", 5000)

    def reapply_style(self, name):
        n = 0
        for ws in self.book.user_sheets():
            for cell in ws._cells.values():
                if not cell.has_style or isinstance(cell, MergedCell):
                    continue
                try:
                    if cell.style != name:
                        continue
                except Exception:
                    continue
                border, numfmt = copy.copy(cell.border), cell.number_format
                cell.style = name
                cell.border, cell.number_format = border, numfmt
                n += 1
        return n

    def rename_style(self):
        styles = [ns.name for ns in self.user_styles()]
        if not styles:
            QMessageBox.information(self, "Rename a style", "This workbook has no styles of its own yet.")
            return
        old, ok = QInputDialog.getItem(self, "Rename a style", "Style:", styles, 0, False)
        if not ok:
            return
        new, ok = QInputDialog.getText(self, "Rename a style", "New name:", text=old)
        new = new.strip()
        if not ok or not new or new == old or new in self.book.wb.named_styles:
            return
        self.book.begin()
        for ns in self.book.wb._named_styles:
            if ns.name == old:
                ns.name = new
        self.book.done("format")
        self.update_style_box()

    def style_key(self, cell):
        f, al = cell.font, cell.alignment
        theme = self.book.theme
        fill = model.resolve_color(cell.fill.fgColor, theme) if cell.fill and cell.fill.fill_type else None
        return ((f.name if f else None) or "Calibri", float((f.sz if f else None) or 11), bool(f and f.b),
                bool(f and f.i), model.resolve_color(f.color, theme) if f else None, fill,
                (al.horizontal if al else None), (al.vertical if al else None), bool(al and al.wrap_text))

    def suggest_styles(self):
        if self.book is None:
            return
        counts, samples = {}, {}
        for ws in self.book.user_sheets():
            if self.read_only_sheet(ws):
                continue
            for cell in ws._cells.values():
                if cell.value is None or isinstance(cell, MergedCell):
                    continue
                key = self.style_key(cell)
                counts[key] = counts.get(key, 0) + 1
                if key not in samples and isinstance(cell.value, str) and not fx.is_formula(cell.value):
                    samples[key] = cell.value.split("\n")[0]
        combos = [(k, n, samples.get(k, "")) for k, n in sorted(counts.items(), key=lambda kv: -kv[1]) if n >= 5][:14]
        if not combos:
            QMessageBox.information(self, APP_NAME, "No formatting repeats often enough to suggest styles.")
            return
        dlg = SuggestStylesDialog(self, combos)
        if not dlg.exec():
            return
        chosen = dlg.chosen()
        if not chosen:
            return
        self.book.begin()
        made = 0
        for key, name in chosen:
            if name in self.book.wb.named_styles:
                continue
            sample = None
            for ws in self.book.user_sheets():
                for cell in ws._cells.values():
                    if cell.value is not None and not isinstance(cell, MergedCell) and self.style_key(cell) == key:
                        sample = cell
                        break
                if sample:
                    break
            ns = NamedStyle(name=name)
            ns.font, ns.fill, ns.alignment = copy.copy(sample.font), copy.copy(sample.fill), copy.copy(sample.alignment)
            self.book.wb.add_named_style(ns)
            for ws in self.book.user_sheets():
                for cell in ws._cells.values():
                    if cell.value is not None and not isinstance(cell, MergedCell) and self.style_key(cell) == key:
                        border, numfmt = copy.copy(cell.border), cell.number_format
                        cell.style = name
                        cell.border, cell.number_format = border, numfmt
            made += 1
        self.book.done("format")
        self.update_style_box()
        self.statusBar().showMessage(f"Made {made} style{'s' if made != 1 else ''}.", 5000)

    # ------------------------------------------------------------ rows & columns
    def insert_rows(self, above=True):
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        r1, c1, r2, c2 = self.grid.selection()
        n = r2 - r1 + 1
        self.book.begin()
        self.book.insert(ws, "row", r1 if above else r2 + 1, n)
        self.book.done("structure")
        if not above:
            self.grid.select_range(r2 + 1, c1, r2 + n, c2)
        else:
            self.grid.select_range(r1, c1, r2, c2)

    def insert_cols(self, left=True):
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        r1, c1, r2, c2 = self.grid.selection()
        n = c2 - c1 + 1
        self.book.begin()
        self.book.insert(ws, "col", c1 if left else c2 + 1, n)
        self.book.done("structure")
        if not left:
            self.grid.select_range(r1, c2 + 1, r2, c2 + n)

    def delete_rows(self):
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        r1, c1, r2, c2 = self.grid.selection()
        self.book.begin()
        self.book.delete(ws, "row", r1, r2 - r1 + 1)
        self.book.done("structure")
        self.grid.set_current(r1, c1)
        what = f"row {r1}" if r1 == r2 else f"rows {r1}–{r2}"
        self.statusBar().showMessage(f"Deleted {what}. Ctrl+Z brings it back.", 6000)

    def delete_cols(self):
        ws = self.current_ws()
        if ws is None or self.read_only_sheet(ws):
            return
        r1, c1, r2, c2 = self.grid.selection()
        self.book.begin()
        self.book.delete(ws, "col", c1, c2 - c1 + 1)
        self.book.done("structure")
        self.grid.set_current(r1, c1)
        what = f"column {fx.num_to_col(c1)}" if c1 == c2 else f"columns {fx.num_to_col(c1)}–{fx.num_to_col(c2)}"
        self.statusBar().showMessage(f"Deleted {what}. Ctrl+Z brings it back.", 6000)

    # ------------------------------------------------------------ links
    SQUARE_RE = re.compile(r"\s*([A-Za-z]{1,2})(\d{1,3})(?![\d])")

    def map_titles(self):
        return [t for t in self.book.meta["maps"] if self.book.sheet(t) is not None]

    def square_of(self, value):
        """'I1' -> (9, 1); also takes the leading square of 'H10 + block'."""
        if not isinstance(value, str):
            return None
        m = self.SQUARE_RE.match(value)
        if not m:
            return None
        return fx.col_to_num(m.group(1)), int(m.group(2))

    def link_cells(self, cells, location_for):
        """Give cells internal links; location_for(cell) -> "'Sheet'!A1" or None."""
        ws = self.current_ws()
        self.book.begin([ws])
        n = 0
        for cell in cells:
            loc = location_for(cell)
            if loc:
                cell.hyperlink = Hyperlink(ref=cell.coordinate, location=loc)
                n += 1
        if n:
            self.book.done("format")
        else:
            self.book.undo_stack.pop()
        return n

    def link_to_square(self, map_title, cells=None):
        info = self.book.meta["maps"][map_title]
        cells = cells or self.selected_cells(create=False)
        def loc(cell):
            sq = self.square_of(cell.value)
            if sq is None or not ranges.map_square(info, *ranges.square_cell(info, sq)):
                return None
            r, c = ranges.square_cell(info, sq)
            return f"{fx.quote_sheet(map_title)}!{fx.num_to_col(c)}{r}"
        n = self.link_cells(cells, loc)
        self.statusBar().showMessage(f"Linked {n} cell{'s' if n != 1 else ''} to {map_title}." if n else
                                     "Nothing in the selection looks like a map square (I6, G13…).", 6000)

    def link_to_sheet(self, title):
        self.link_cells([self.current_ws().cell(*self.grid.cur)], lambda cell: f"{fx.quote_sheet(title)}!A1")

    def remove_links(self):
        ws = self.current_ws()
        if ws is None:
            return
        cells = [c for c in self.selected_cells(create=False) if c.hyperlink is not None]
        if not cells:
            return
        self.book.begin([ws])
        for cell in cells:
            cell.hyperlink = None
        self.book.done("format")

    def add_link_actions(self, menu):
        cell = self.current_cell()
        menu.addSeparator()
        target = link_target(cell)
        if target:
            menu.addAction("Follow link", lambda: self.follow_link(target))
        if any(c.hyperlink is not None for c in self.selected_cells(create=False)):
            menu.addAction("Remove link (keep the text)", self.remove_links)
        maps = self.map_titles()
        sq_cells = [c for c in self.selected_cells(create=False) if self.square_of(c.value)]
        if sq_cells and maps:
            label = (f"Link {cell.value.strip()} to a map" if len(sq_cells) == 1 and cell is not None and isinstance(cell.value, str)
                     else f"Link these {len(sq_cells)} squares to a map")
            sub = menu.addMenu(label)
            for t in maps:
                sub.addAction(t, lambda t=t: self.link_to_square(t, sq_cells))
        sheets = menu.addMenu("Link to a sheet")
        for ws in self.sidebar_order():
            if ws.title != self.ws_title:
                sheets.addAction(ws.title, lambda t=ws.title: self.link_to_sheet(t))

    def link_selection_dialog(self):
        if self.current_ws() is None:
            return
        maps = self.map_titles()
        if not maps:
            QMessageBox.information(self, "Link squares to a map", "This workbook has no map sheets yet.")
            return
        title, ok = QInputDialog.getItem(self, "Link squares to a map",
                                         "Turn the squares in the selection (I6, G13…) into buttons that go to:",
                                         maps, 0, False)
        if ok:
            self.link_to_square(title)

    def set_freeze(self, how):
        """Frozen rows/columns stay put while the rest scrolls (undoable)."""
        ws = self.current_ws()
        if ws is None:
            return
        r, c = self.grid.cur
        target = {None: None, "row": "A2", "col": "B1",
                  "here": None if (r, c) == (1, 1) else f"{fx.num_to_col(c)}{r}"}[how]
        if target == ws.freeze_panes:
            return
        self.book.begin([ws])
        ws.freeze_panes = target
        self.book.done("layout")
        self.statusBar().showMessage("Unfrozen." if target is None else
                                     f"Frozen above and left of {target}. View → Freeze panes → Unfreeze undoes it.", 6000)

    def set_row_mode(self, on):
        self.grid.row_mode = "compact" if on else "fit"
        self.a_compact.setChecked(on)
        self.settings.setValue("row_mode", self.grid.row_mode)
        self.grid.invalidate_layout(refit=True)
        self.grid.ensure_visible(*self.grid.cur)

    def toggle_headers(self, on):
        self.grid.show_headers = on
        self.grid.invalidate_layout()

    def toggle_adapt(self, on):
        self.grid.adapt_colors = on
        self.settings.setValue("adapt_colors", "true" if on else "false")
        self.grid.viewport().update()

    def grid_menu(self, menu):
        ws = self.current_ws()
        ro = self.read_only_sheet(ws)
        menu.addAction("Cut", lambda: self.grid.copy(cut=True)).setEnabled(not ro)
        menu.addAction("Copy", self.grid.copy)
        menu.addAction("Paste", self.grid.paste).setEnabled(not ro)
        menu.addSeparator()
        if not ro:
            menu.addAction("Insert rows above", lambda: self.insert_rows(True))
            menu.addAction("Insert rows below", lambda: self.insert_rows(False))
            menu.addAction("Insert columns left", lambda: self.insert_cols(True))
            menu.addAction("Insert columns right", lambda: self.insert_cols(False))
            menu.addAction("Delete rows", self.delete_rows)
            menu.addAction("Delete columns", self.delete_cols)
            menu.addSeparator()
            menu.addAction("Clear contents", self.grid.clear_contents)
            menu.addAction("Clear formatting", self.clear_formatting)
            menu.addAction("Edit note", self.a_note.trigger)
            menu.addSeparator()
            menu.addAction("Fit row height to text", lambda: self.grid.autofit("row", self.grid.cur[0]))
            menu.addAction("Fit column width to text", lambda: self.grid.autofit("col", self.grid.cur[1]))
            menu.addSeparator()
            r, c = self.grid.cur
            if ws.freeze_panes:
                menu.addAction(f"Unfreeze panes (frozen at {ws.freeze_panes})", lambda: self.set_freeze(None))

        info = self.book.meta["maps"].get(ws.title) if ws is not None else None
        sq = ranges.map_square(info, *self.grid.cur) if info else None
        if sq and not ro:
            menu.addSeparator()
            menu.addAction(f"Note for square {ranges.square_label(sq)}…", lambda: self.edit_square_note(ws.title, sq))
            for layer, toward in ranges.path_ends_here(self.book, ws.title, sq):
                mode = ranges.path_end(self.book, layer, ws.title, sq, toward, info)
                other = "end" if mode == "run" else "run"
                label = (f"{layer.name}: stop in this square" if other == "end"
                         else f"{layer.name}: run on to the edge")
                menu.addAction(label, lambda l=layer.name, o=other: self.set_path_end(ws.title, sq, l, o))
        if not ro:
            self.add_link_actions(menu)
        cell = self.current_cell()
        if cell is not None and isinstance(cell.value, str) and cell.value.strip() and not fx.is_formula(cell.value):
            menu.addSeparator()
            text = cell.value.strip().split("\n")[0][:40]
            menu.addAction(f"Find “{text}” everywhere", lambda: self.find_text(text))
            name = cell.value.strip()
            existing = self.book.ranges.get(name)
            if existing is not None and ws.title not in self.book.meta["maps"] and self.map_titles():
                painted = [t for t in self.map_titles() if existing.has_map(t)]
                target = painted[0] if painted else self.map_titles()[0]
                menu.addAction(f"Go to the {existing.name} layer", lambda: self.show_range(existing.name, target))
            elif (not ro and "\n" not in name and len(name) <= 60
                    and ws.title not in self.book.meta["maps"] and self.map_titles()):
                menu.addAction(f"Make a layer for “{name}”", lambda: self.make_layer_for(name))

    # ------------------------------------------------------------ ranges
    def find_text(self, text):
        self.find_bar.all.setChecked(True)
        self.find_bar.open_bar(text)

    def show_range(self, name, map_title):
        st = self.range_state
        st.solo = name
        st.layer = name
        st.split = True
        st.blocks = set(ranges.BLOCKS)
        self.open_sheet(map_title)
        self.show_layers_tab()
        self.ranges_panel.refresh()
        self.grid.viewport().update()

    def make_layer_for(self, name):
        maps = [t for t in self.book.meta["maps"] if self.book.sheet(t) is not None]
        if not maps:
            QMessageBox.information(self, "Range layer", "This workbook has no map sheets yet. Open a map sheet, "
                                    "then use the Ranges tab to mark its grid.")
            return
        map_title = maps[0]
        if len(maps) > 1:
            map_title, ok = QInputDialog.getItem(self, "New layer", f"Paint {name} on which map?", maps, 0, False)
            if not ok:
                return                      # cancelled: nothing made, nowhere to go
        self.book.begin([])
        layer = self.book.ranges.add(name)
        self.book.done("ranges")
        self.show_range(layer.name, map_title)
        self.range_state.painting = True
        self.ranges_panel.refresh()
        self.update_paint_bar()

    # ------------------------------------------------------------ help
    def show_shortcuts(self):
        QMessageBox.information(self, "Keyboard shortcuts", "\n".join([
            "Ctrl+P  go to a sheet (type part of its name, Enter)",
            "Alt+← / Alt+→  back / forward between places you've been",
            "Ctrl+F  find (all sheets; Enter for next, Shift+Enter for previous)",
            "Type or F2  edit a cell · Enter keeps it · Esc cancels",
            "Shift+Enter or Alt+Enter  new line inside a cell",
            "Ctrl+Enter  save the text in the Cell panel",
            "Ctrl+arrows  jump to the edge of the text",
            "Ctrl+D / Ctrl+R  fill down / right",
            "Ctrl+wheel or Ctrl+= / Ctrl+-  zoom",
            "Ctrl+Z / Ctrl+Y  undo / redo",
            "Double-click a header edge  fit the column or row",
            "Ctrl+click a link  follow it",
            "Ranges tab · Paint on: drag over squares · Shift-drag erases",
        ]))

    def show_about(self):
        QMessageBox.about(self, f"About {APP_NAME}",
                          f"<b>{APP_NAME}</b> {read_version()}<br><br>A spreadsheet for worldbuilding: many sheets of "
                          "mostly text, grouped in a sidebar, with range layers for maps.<br><br>"
                          "Files stay ordinary .xlsx. Built with Qt (PySide6) and openpyxl.")


def map_icon(color):
    pm = QPixmap(14, 14)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setPen(QPen(QColor(color), 1.3))
    p.drawPolyline([QPointF(x, y) for x, y in
                    [(1, 3), (5, 1), (9, 3), (13, 1), (13, 11), (9, 13), (5, 11), (1, 13), (1, 3)]])
    p.drawLine(5, 1, 5, 11)
    p.drawLine(9, 3, 9, 13)
    p.end()
    return QIcon(pm)


def main():
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName("Nous")
    app.setWindowIcon(app_icon())
    app.setStyle("Fusion")
    if sys.platform == "win32":
        app.setFont(QFont("Segoe UI", 10))
    w = MainWindow()
    w.show()
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args and os.path.exists(args[0]):
        w.open_path(os.path.abspath(args[0]))
    else:
        recent = w.settings.value("recent") or []
        if recent and os.path.exists(recent[0]) and "--fresh" not in sys.argv:
            w.open_path(recent[0])
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
