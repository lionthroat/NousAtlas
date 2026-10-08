"""The sheet view: a hand-painted grid that scrolls by the pixel (not a row
at a time), with frozen panes, merged cells, wrapped text, borders, text
that spills into empty neighbours, and an in-place editor.

Range layers and the paint brush plug in through `overlay` and `tool`.
"""

import bisect
import copy
import csv
import io

from PySide6.QtCore import QPoint, QRect, QSize, Qt, Signal
from PySide6.QtGui import (QBrush, QColor, QFont, QFontMetrics, QKeySequence,
                           QPainter, QPen, QPolygon, QRegion, QTextOption)
from PySide6.QtWidgets import (QAbstractScrollArea, QApplication, QMenu,
                               QPlainTextEdit)
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

import atlas_formula as fx
from atlas_model import (DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, resolve_color)
from atlas_theme import contrast, qcolor, readable_on

PAD_X, PAD_Y = 4, 2
HEADER_H = 22
MIN_HEADER_W = 36
COMPACT_LINES = 3
EXTRA_ROWS, EXTRA_COLS = 40, 8
MDW = 7                               # Excel's "maximum digit width" at 100%


class Style:
    """A cell style worked out once per distinct style id."""
    __slots__ = ("font", "fm", "color", "fill", "halign", "valign", "wrap",
                 "indent", "borders", "underline", "font_key")


class CellEditor(QPlainTextEdit):
    def __init__(self, view):
        super().__init__(view.viewport())
        self.view = view
        self.setFrameShape(QPlainTextEdit.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setTabChangesFocus(False)
        self.document().setDocumentMargin(PAD_X - 1)
        self.textChanged.connect(self.grow)

    def keyPressEvent(self, e):
        k, mods = e.key(), e.modifiers()
        if k in (Qt.Key_Return, Qt.Key_Enter):
            if mods & (Qt.ShiftModifier | Qt.AltModifier):
                self.insertPlainText("\n")
                return
            self.view.finish_edit(True, (-1 if mods & Qt.ShiftModifier else 1, 0))
            return
        if k == Qt.Key_Tab:
            self.view.finish_edit(True, (0, 1))
            return
        if k == Qt.Key_Backtab:
            self.view.finish_edit(True, (0, -1))
            return
        if k == Qt.Key_Escape:
            self.view.finish_edit(False)
            return
        super().keyPressEvent(e)

    def focusOutEvent(self, e):
        super().focusOutEvent(e)
        if e.reason() not in (Qt.PopupFocusReason, Qt.ActiveWindowFocusReason) and self.isVisible():
            self.view.finish_edit(True, None)

    def grow(self):
        doc_h = int(self.document().size().height()) + 6
        rect = self.view.editor_base
        if rect is None:
            return
        avail = self.view.viewport().height() - rect.top() - 4
        self.resize(self.width(), max(rect.height(), min(doc_h, max(avail, rect.height()))))


class SheetView(QAbstractScrollArea):
    currentChanged = Signal()          # selection or current cell moved
    linkActivated = Signal(str)        # an internal "#Sheet!A1" or external URL
    zoomChanged = Signal(float)
    contextRequested = Signal(object)  # a QMenu to add to before it pops up

    def __init__(self, theme):
        super().__init__()
        self.theme = theme
        self.book = None
        self.ws = None
        self.zoom = 1.0
        self.row_mode = "fit"           # or "compact"
        self.show_headers = True
        self.adapt_colors = True
        self.overlay = None             # range layers: .paint(p, ws, row, col, rect)
        self.tool = None                # paint brush: .press/.move/.release(row, col, event)
        self.find_hits = set()
        self.find_current = None
        self.cur = (1, 1)
        self.anchor = (1, 1)
        self.editor = None
        self.editor_base = None
        self.edit_cell = None
        self._drag = None
        self._styles = {}
        self._fonts = {}
        self._fit = {}
        self._layout_dirty = True
        self._merges = {}
        self._covered = {}
        self._region = QRect()
        self.pad_x, self.pad_y, self.gap = PAD_X, PAD_Y, 0
        self.ox = self.oy = 0
        self._spill_to = 0
        self.clip = None                # set by the window: shared clipboard
        self.setFocusPolicy(Qt.StrongFocus)
        self.viewport().setMouseTracking(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.horizontalScrollBar().valueChanged.connect(self._scrolled)
        self.verticalScrollBar().valueChanged.connect(self._scrolled)
        self.horizontalScrollBar().setSingleStep(20)
        self.verticalScrollBar().setSingleStep(20)

    # ------------------------------------------------------------ setup
    def set_sheet(self, book, ws, cur=None):
        self.cancel_edit()
        self.book, self.ws = book, ws
        self._styles.clear()
        self.invalidate_layout(refit=True)
        self.cur = self.anchor = cur or (1, 1)
        self.horizontalScrollBar().setValue(0)
        self.verticalScrollBar().setValue(0)
        if cur:
            self.ensure_visible(*cur)
        self.viewport().update()
        self.currentChanged.emit()

    def set_theme(self, theme):
        self.theme = theme
        self._styles.clear()
        self.viewport().update()

    def set_zoom(self, z):
        z = max(0.25, min(4.0, round(z, 2)))
        if z == self.zoom:
            return
        self.cancel_edit()
        self.zoom = z
        self._styles.clear()
        self.invalidate_layout()
        self.ensure_visible(*self.cur)
        self.zoomChanged.emit(z)

    def invalidate_layout(self, refit=False):
        if refit:
            self._fit.clear()
            self._fit_done = False
        self._layout_dirty = True
        self.viewport().update()

    def content_changed(self):
        """Called by the window after the book changes."""
        self._styles.clear()
        self.invalidate_layout(refit=True)
        if self.ws is not None and self.ws not in self.book.wb.worksheets:
            return
        self.viewport().update()

    # ------------------------------------------------------------ geometry
    def _layout(self):
        if not self._layout_dirty or self.ws is None:
            return
        ws = self.ws
        self._merges, self._covered = {}, {}
        for m in ws.merged_cells.ranges:
            self._merges[(m.min_row, m.min_col)] = (m.max_row, m.max_col)
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1):
                    if (r, c) != (m.min_row, m.min_col):
                        self._covered[(r, c)] = (m.min_row, m.min_col)
        if not getattr(self, "_fit_done", False):
            self._compute_fit()
        self.nrows = max(ws.max_row, self.cur[0], self.anchor[0]) + EXTRA_ROWS
        self.ncols = max(ws.max_column, self.cur[1], self.anchor[1]) + EXTRA_COLS
        z = self.zoom
        lay = self.sheet_layout()
        self.pad_x = int(round(lay["pad_x"] * z))
        self.pad_y = int(round(lay["pad_y"] * z))
        self.gap = int(round(lay["spacing"] * z))
        margin = int(round(lay["margin"] * z))
        default_w = ws.sheet_format.defaultColWidth or ws.sheet_format.baseColWidth and (ws.sheet_format.baseColWidth + 0.71) or DEFAULT_COL_WIDTH
        self.colx = [0]
        dims = ws.column_dimensions
        for c in range(1, self.ncols + 1):
            d = dims.get(get_column_letter(c))
            if d is not None and d.hidden:
                w = 0
            else:
                width = d.width if d is not None and d.width else default_w
                w = int((int(width * MDW + 5)) * z) + self.gap
            self.colx.append(self.colx[-1] + w)
        self.rowy = [0]
        default_h = ws.sheet_format.defaultRowHeight or DEFAULT_ROW_HEIGHT
        rdims = ws.row_dimensions
        cap = self._compact_cap()
        pinned = set(self.book.meta.get("pinned_rows", {}).get(ws.title, []))
        for r in range(1, self.nrows + 1):
            d = rdims.get(r)
            if d is not None and d.hidden:
                h = 0
            elif r in pinned:
                # a height you set by hand is kept exactly, in either row mode
                h = int(((d.height if d is not None and d.height else None) or default_h) * 96 / 72 * z) + self.gap
            else:
                stored = d.height if d is not None and d.height else None
                base = (stored or default_h) * 96 / 72
                fit = self._fit.get(r, 0)
                h = max(base, fit)
                if self.row_mode == "compact":
                    h = min(h, max(base if stored and base < cap else 0, cap))
                h = int(h * z) + self.gap
            self.rowy.append(self.rowy[-1] + h)
        # frozen panes
        self.frow, self.fcol = 1, 1
        if ws.freeze_panes:
            cell = ws[ws.freeze_panes]
            self.frow, self.fcol = cell.row, cell.column
        self.header_w = 0
        self.header_h = 0
        if self.show_headers:
            fm = QFontMetrics(self._ui_font())
            self.header_w = max(MIN_HEADER_W, fm.horizontalAdvance(str(self.nrows)) + 14)
            self.header_h = int(HEADER_H * max(0.8, min(1.3, z)))
        self.ox = self.header_w + margin       # where column A starts
        self.oy = self.header_h + margin       # where row 1 starts
        self._layout_dirty = False
        self._update_scrollbars()

    DEFAULT_LAYOUT = {"margin": 0, "pad_x": PAD_X, "pad_y": PAD_Y, "spacing": 0}

    def sheet_layout(self):
        """This sheet's margin, cell padding and cell spacing (unzoomed px)."""
        out = dict(self.DEFAULT_LAYOUT)
        if self.book is not None and self.ws is not None:
            out.update(self.book.meta.get("sheet_layout", {}).get(self.ws.title, {}))
        return out

    def _compact_cap(self):
        fm = self._metrics("Arial", 10, False, False, 1.0)
        return fm.lineSpacing() * COMPACT_LINES + 2 * self.sheet_layout()["pad_y"] + 4

    def _compute_fit(self):
        """Heights each row needs for its text, at 100% zoom."""
        ws = self.ws
        lay = self.sheet_layout()
        self._fit = {}
        self._fit_done = True
        if self.row_mode is None:
            return
        save_zoom, self.zoom = self.zoom, 1.0
        try:
            dims = ws.column_dimensions
            default_w = ws.sheet_format.defaultColWidth or DEFAULT_COL_WIDTH
            def colw(c):
                d = dims.get(get_column_letter(c))
                if d is not None and d.hidden:
                    return 0
                return int((d.width if d is not None and d.width else default_w) * MDW + 5)
            for (r, c), cell in ws._cells.items():
                v = cell.value
                if v is None or isinstance(cell, MergedCell) or (r, c) in self._covered:
                    continue
                if not isinstance(v, str) or fx.is_formula(v):
                    continue
                st = self.style_of(cell)
                span = self._merges.get((r, c))
                if span and span[0] > r:
                    continue                    # tall merges size themselves
                w = sum(colw(cc) for cc in range(c, (span[1] if span else c) + 1)) - 2 * lay["pad_x"]
                if st.wrap and w > 4:
                    rect = st.fm.boundingRect(QRect(0, 0, w, 100000), Qt.TextWordWrap, v)
                    h = rect.height()
                elif "\n" in v:
                    h = st.fm.lineSpacing() * (v.count("\n") + 1)
                else:
                    continue
                h += 2 * lay["pad_y"] + 2
                if h > self._fit.get(r, 0):
                    self._fit[r] = h
        finally:
            self.zoom = save_zoom
            self._styles.clear()

    def _update_scrollbars(self):
        vp = self.viewport().size()
        fw = self.colx[self.fcol - 1]
        fh = self.rowy[self.frow - 1]
        total_w = self.colx[-1] - fw
        total_h = self.rowy[-1] - fh
        avail_w = vp.width() - self.ox - fw
        avail_h = vp.height() - self.oy - fh
        self.horizontalScrollBar().setRange(0, max(0, total_w - avail_w))
        self.verticalScrollBar().setRange(0, max(0, total_h - avail_h))
        self.horizontalScrollBar().setPageStep(max(20, avail_w))
        self.verticalScrollBar().setPageStep(max(20, avail_h))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if self.ws is not None:
            self._layout()
            self._update_scrollbars()

    def _scrolled(self):
        if self.editor is not None and self.editor.isVisible():
            self._place_editor()
        self.viewport().update()

    def _ensure_extent(self, row, col):
        if self.ws is None:
            return
        self._layout()
        if row > self.nrows - 5 or col > self.ncols - 3:
            self._layout_dirty = True
            self._layout()

    # screen position of a cell's left / top edge
    def x_of(self, c):
        if c < self.fcol:
            return self.ox + self.colx[c - 1]
        return (self.ox + self.colx[self.fcol - 1] + (self.colx[c - 1] - self.colx[self.fcol - 1])
                - self.horizontalScrollBar().value())

    def y_of(self, r):
        if r < self.frow:
            return self.oy + self.rowy[r - 1]
        return (self.oy + self.rowy[self.frow - 1] + (self.rowy[r - 1] - self.rowy[self.frow - 1])
                - self.verticalScrollBar().value())

    def cell_rect(self, r, c, merged=True):
        r2, c2 = r, c
        if merged:
            if (r, c) in self._covered:
                r, c = self._covered[(r, c)]
            span = self._merges.get((r, c))
            if span:
                r2, c2 = span
        self._ensure_extent(r2, c2)
        x, y = self.x_of(c), self.y_of(r)
        return QRect(x, y, self.colx[c2] - self.colx[c - 1] - self.gap, self.rowy[r2] - self.rowy[r - 1] - self.gap)

    def hit(self, pos):
        """(row, col) under a viewport position, or None. Headers give 0."""
        self._layout()
        x, y = pos.x(), pos.y()
        fx_w = self.colx[self.fcol - 1]
        fy_h = self.rowy[self.frow - 1]
        if x < self.header_w:
            col = 0
        else:
            ax = max(0, x - self.ox)
            if ax >= fx_w:
                ax += self.horizontalScrollBar().value()
            col = min(bisect.bisect_right(self.colx, ax), self.ncols)
        if y < self.header_h:
            row = 0
        else:
            ay = max(0, y - self.oy)
            if ay >= fy_h:
                ay += self.verticalScrollBar().value()
            row = min(bisect.bisect_right(self.rowy, ay), self.nrows)
        return row, col

    def ensure_visible(self, r, c):
        if self.ws is None:
            return
        self._ensure_extent(r, c)
        self._layout()
        vp = self.viewport().size()
        if c >= self.fcol:
            left = self.colx[c - 1] - self.colx[self.fcol - 1]
            right = self.colx[c] - self.colx[self.fcol - 1]
            avail = vp.width() - self.ox - self.colx[self.fcol - 1]
            sb = self.horizontalScrollBar()
            if left < sb.value():
                sb.setValue(left)
            elif right > sb.value() + avail:
                sb.setValue(min(left, right - avail))
        if r >= self.frow:
            top = self.rowy[r - 1] - self.rowy[self.frow - 1]
            bottom = self.rowy[r] - self.rowy[self.frow - 1]
            avail = vp.height() - self.oy - self.rowy[self.frow - 1]
            sb = self.verticalScrollBar()
            if top < sb.value():
                sb.setValue(top)
            elif bottom > sb.value() + avail:
                sb.setValue(min(top, bottom - avail))

    # ------------------------------------------------------------ styles
    def _ui_font(self):
        f = QFont(self.font())
        f.setPixelSize(max(9, int(11 * max(0.8, min(1.3, self.zoom)))))
        return f

    def _metrics(self, name, size, bold, italic, zoom=None):
        key = (name, size, bold, italic, zoom or self.zoom)
        hit = self._fonts.get(key)
        if hit is None:
            f = QFont(name)
            f.setPixelSize(max(1, round(size * 96 / 72 * (zoom or self.zoom))))
            f.setBold(bold)
            f.setItalic(italic)
            hit = self._fonts[key] = (f, QFontMetrics(f))
        return hit[1]

    def style_of(self, cell):
        key = tuple(cell._style) if cell._style is not None else ()
        st = self._styles.get(key)
        if st is not None:
            return st
        st = Style()
        theme = self.book.theme
        font = cell.font
        name = (font.name if font is not None and font.name else None) or "Calibri"
        size = (font.sz if font is not None and font.sz else None) or 11
        bold = bool(font.b) if font is not None else False
        italic = bool(font.i) if font is not None else False
        st.font_key = (name, size, bold, italic, self.zoom)
        self._metrics(*st.font_key)
        f, st.fm = self._fonts[st.font_key]
        f = QFont(f)
        st.underline = bool(font.u) if font is not None else False
        f.setUnderline(st.underline)
        f.setStrikeOut(bool(font.strike) if font is not None else False)
        st.font = f
        st.fill = None
        fill = cell.fill
        if fill is not None and fill.fill_type:
            st.fill = resolve_color(fill.fgColor, theme) if fill.fill_type == "solid" else \
                resolve_color(fill.fgColor, theme) or resolve_color(fill.bgColor, theme)
        st.color = resolve_color(font.color, theme) if font is not None else None
        al = cell.alignment
        st.halign = al.horizontal if al is not None else None
        st.valign = (al.vertical if al is not None else None) or "bottom"
        st.wrap = bool(al.wrap_text) if al is not None else False
        st.indent = int(al.indent or 0) if al is not None else 0
        st.borders = []
        b = cell.border
        if b is not None:
            for side in ("left", "top", "right", "bottom"):
                s = getattr(b, side)
                if s is not None and s.style:
                    st.borders.append((side, s.style, resolve_color(s.color, theme)))
        self._styles[key] = st
        return st

    def text_color(self, st):
        t = self.theme
        canvas = t["canvas"].lstrip("#").upper()
        if st.fill:
            if st.color:
                return qcolor(st.color)
            return QColor("#000000") if _lum(st.fill) > 0.35 else QColor("#ffffff")
        if not st.color:
            return QColor(t["text"])
        return qcolor(readable_on(st.color, canvas) if self.adapt_colors else st.color)

    def border_color(self, hex6, fill):
        if hex6 is None:
            hex6 = "000000"
        if fill is None and self.adapt_colors:
            hex6 = readable_on(hex6, self.theme["canvas"].lstrip("#").upper(), 2.0)
        return qcolor(hex6)

    # ------------------------------------------------------------ painting
    def paintEvent(self, e):
        p = QPainter(self.viewport())
        t = self.theme
        p.fillRect(self.viewport().rect(), QColor(t["canvas"]))
        if self.ws is None:
            return
        self._layout()
        vp = self.viewport().rect()
        hw, hh = self.ox, self.oy
        fw = self.colx[self.fcol - 1]
        fh = self.rowy[self.frow - 1]
        regions = [
            (QRect(hw + fw, hh + fh, vp.width(), vp.height()), False, False),
            (QRect(hw + fw, hh, vp.width(), fh), True, False),
            (QRect(hw, hh + fh, fw, vp.height()), False, True),
            (QRect(hw, hh, fw, fh), True, True),
        ]
        for rect, frozen_r, frozen_c in regions:
            rect = rect.intersected(vp)
            if rect.isEmpty():
                continue
            rows = self._visible(self.rowy, rect.top(), rect.bottom(), frozen_r, True)
            cols = self._visible(self.colx, rect.left(), rect.right(), frozen_c, False)
            if rows and cols:
                p.save()
                p.setClipRect(rect)
                # text in frozen columns may spill right across the freeze line
                self._region = rect
                self._spill_to = vp.right() if (frozen_c and self.horizontalScrollBar().value() == 0) else rect.right()
                self._paint_cells(p, rows, cols)
                p.restore()
        # freeze lines
        if self.frow > 1 or self.fcol > 1:
            p.setPen(QPen(QColor(t["border"]), 1))
            if self.frow > 1:
                p.drawLine(hw, hh + fh, vp.width(), hh + fh)
            if self.fcol > 1:
                p.drawLine(hw + fw, hh, hw + fw, vp.height())
        if self.show_headers:
            self._paint_headers(p)

    def _visible(self, edges, lo, hi, frozen, is_row):
        """1-based indices whose span overlaps [lo, hi] on screen."""
        f = self.frow if is_row else self.fcol
        n = self.nrows if is_row else self.ncols
        if frozen:
            rng = range(1, f)
        else:
            rng = None
        pos = self.y_of if is_row else self.x_of
        out = []
        if rng is None:
            # binary search the first visible one
            off = self.verticalScrollBar().value() if is_row else self.horizontalScrollBar().value()
            base = edges[f - 1]
            start = (self.oy if is_row else self.ox) + base
            target = lo - start + off + base
            i = max(f, bisect.bisect_left(edges, target))
            rng = range(i, n + 1)
        for i in rng:
            a = pos(i)
            if a > hi:
                break
            if a + (edges[i] - edges[i - 1]) >= lo:
                out.append(i)
        return out

    def _paint_cells(self, p, rows, cols):
        t = self.theme
        ws = self.ws
        cells = ws._cells
        grid_pen = QPen(QColor(t["grid"]), 1)
        show_grid = ws.sheet_view.showGridLines is not False and self.gap == 0
        # 1. gridlines
        if show_grid:
            p.setPen(grid_pen)
            top, bottom = self.y_of(rows[0]), self.y_of(rows[-1]) + self.rowy[rows[-1]] - self.rowy[rows[-1] - 1]
            left, right = self.x_of(cols[0]), self.x_of(cols[-1]) + self.colx[cols[-1]] - self.colx[cols[-1] - 1]
            for c in cols:
                x = self.x_of(c) + self.colx[c] - self.colx[c - 1] - 1
                p.drawLine(x, top, x, bottom)
            for r in rows:
                y = self.y_of(r) + self.rowy[r] - self.rowy[r - 1] - 1
                p.drawLine(left, y, right, y)
        # which cells to draw (merges drawn once, from their top-left)
        todo = []
        seen = set()
        for r in rows:
            for c in cols:
                key = (r, c)
                if key in self._covered:
                    key = self._covered[key]
                    if key in seen:
                        continue
                if key in seen:
                    continue
                seen.add(key)
                todo.append(key)
        # 2. fills and overlays
        for (r, c) in todo:
            cell = cells.get((r, c))
            rect = self.cell_rect(r, c)
            if rect.width() <= 0 or rect.height() <= 0:
                continue
            span = self._merges.get((r, c))
            st = self.style_of(cell) if cell is not None else None
            if st is not None and st.fill:
                p.fillRect(rect, qcolor(st.fill))
            elif span:
                p.fillRect(rect.adjusted(0, 0, -1, -1), QColor(t["canvas"]))
            if self.overlay is not None:
                self.overlay.paint(p, ws, r, c, rect)
        # 3. borders
        for (r, c) in todo:
            cell = cells.get((r, c))
            if cell is None:
                continue
            st = self.style_of(cell)
            if not st.borders:
                continue
            rect = self.cell_rect(r, c)
            self._paint_borders(p, rect, st)
        # merged cells also take borders from their right/bottom edge cells
        # 4. text
        for (r, c) in todo:
            cell = cells.get((r, c))
            if cell is None or cell.value is None:
                continue
            rect = self.cell_rect(r, c)
            if rect.width() <= 2 or rect.height() <= 2:
                continue
            self._paint_text(p, r, c, cell, rect)
        # 5. notes, find hits, selection
        accent = QColor(t["accent"])
        for (r, c) in todo:
            cell = cells.get((r, c))
            if cell is not None and cell.comment is not None:
                rect = self.cell_rect(r, c)
                tri = QPolygon([QPoint(rect.right() - 7, rect.top()), QPoint(rect.right(), rect.top()),
                                QPoint(rect.right(), rect.top() + 7)])
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(t["link"]))
                p.drawPolygon(tri)
        if self.find_hits:
            for key in todo:
                if key in self.find_hits:
                    rect = self.cell_rect(*key)
                    color = QColor(t["find"])
                    p.setBrush(Qt.NoBrush)
                    p.setPen(QPen(color, 3 if key == self.find_current else 1.5))
                    p.drawRect(rect.adjusted(1, 1, -2, -2))
        self._paint_selection(p, rows, cols)

    def _paint_borders(self, p, rect, st):
        for side, style, color in st.borders:
            width = {"medium": 2, "thick": 3, "double": 3, "mediumDashed": 2,
                     "mediumDashDot": 2, "mediumDashDotDot": 2, "slantDashDot": 2}.get(style, 1)
            pen = QPen(self.border_color(color, st.fill), width)
            if "Dash" in style or style == "dashed":
                pen.setStyle(Qt.DashLine)
            elif style in ("dotted", "hair"):
                pen.setStyle(Qt.DotLine)
            p.setPen(pen)
            x1, y1, x2, y2 = rect.left(), rect.top(), rect.left() + rect.width() - 1, rect.top() + rect.height() - 1
            if side == "left":
                p.drawLine(x1, y1, x1, y2)
            elif side == "right":
                p.drawLine(x2, y1, x2, y2)
            elif side == "top":
                p.drawLine(x1, y1, x2, y1)
            else:
                p.drawLine(x1, y2, x2, y2)

    def _paint_text(self, p, r, c, cell, rect):
        st = self.style_of(cell)
        value = cell.value
        if fx.is_formula(value):
            shown = self.book.calc.value(self.ws.title, r, c)
        else:
            shown = value
        text = self.book.display_text(self.ws, r, c)
        if not text:
            return
        is_number = isinstance(shown, (int, float)) and not isinstance(shown, bool)
        h = st.halign
        if h in (None, "general"):
            h = "right" if is_number else ("center" if isinstance(shown, (bool, fx.XlError)) else "left")
        hflag = {"left": Qt.AlignLeft, "right": Qt.AlignRight, "center": Qt.AlignHCenter,
                 "centerContinuous": Qt.AlignHCenter, "distributed": Qt.AlignHCenter,
                 "justify": Qt.AlignJustify, "fill": Qt.AlignLeft}.get(h, Qt.AlignLeft)
        vflag = {"top": Qt.AlignTop, "center": Qt.AlignVCenter, "justify": Qt.AlignTop,
                 "distributed": Qt.AlignVCenter}.get(st.valign, Qt.AlignBottom)
        indent = int(st.indent * 9 * self.zoom)
        px, py = self.pad_x, self.pad_y
        inner = rect.adjusted(px + (indent if h == "left" else 0), py,
                              -px - (indent if h == "right" else 0), -py - 1)
        clip = QRect(rect)
        wrap = 0
        if st.wrap:
            wrap = int(Qt.TextWordWrap)
            # when a row is too short for its text, show the beginning, not the middle
            need = st.fm.boundingRect(QRect(0, 0, max(1, inner.width()), 100000), wrap | int(hflag), text).height()
            if need > inner.height():
                vflag = Qt.AlignTop
                ls = max(1, st.fm.lineSpacing())
                inner.setHeight(max(ls, inner.height() // ls * ls))   # whole lines only
        elif "\n" not in text and h == "left" and (r, c) not in self._merges:
            # spill into empty cells to the right, the way Excel does
            need = st.fm.horizontalAdvance(text) + 2 * px + indent
            cc = c
            right = rect.right()
            while need > right - rect.left() and cc < self.ncols:
                nxt = self.ws._cells.get((r, cc + 1))
                if (nxt is not None and nxt.value not in (None, "")) or (r, cc + 1) in self._covered \
                        or (r, cc + 1) in self._merges:
                    break
                cc += 1
                right += self.colx[cc] - self.colx[cc - 1]
            if cc > c:
                clip.setRight(right)
                inner.setRight(right - px)
        elif "\n" not in text and h == "right" and is_number and st.fm.horizontalAdvance(text) > inner.width():
            text = "#" * max(1, inner.width() // max(1, st.fm.horizontalAdvance("#")))
        p.save()
        region = self._region
        limit = QRect(region.left(), region.top(), max(region.width(), self._spill_to - region.left() + 1), region.height())
        p.setClipRect(clip.intersected(limit), Qt.ReplaceClip)
        p.setFont(st.font)
        color = self.text_color(st)
        if isinstance(shown, fx.XlError):
            color = QColor(self.theme["danger"])
        p.setPen(color)
        flags = int(hflag) | int(vflag) | wrap
        if clip.right() > rect.right():
            # text spilling past its cell: over each neighbour it takes the
            # colour that reads on that neighbour (its fill, or the bare sheet)
            p.setClipRect(rect.intersected(limit), Qt.ReplaceClip)
            p.drawText(inner, flags, text)
            base = st.color or ("000000" if st.fill and _lum(st.fill) > 0.35 else None)
            canvas = self.theme["canvas"].lstrip("#").upper()
            cc = c + 1
            while cc <= self.ncols and self.x_of(cc) <= clip.right():
                x = self.x_of(cc)
                seg = QRect(x, clip.top(), self.colx[cc] - self.colx[cc - 1], clip.height())
                nb = self.ws._cells.get((r, cc))
                nfill = self.style_of(nb).fill if nb is not None else None
                if nfill:
                    col = base or ("000000" if _lum(nfill) > 0.35 else "FFFFFF")
                    if contrast(col, nfill) < 2.0:
                        col = "000000" if _lum(nfill) > 0.35 else "FFFFFF"
                    pen = qcolor(col)
                elif base is None:
                    pen = QColor(self.theme["text"])
                else:
                    pen = qcolor(readable_on(base, canvas) if self.adapt_colors else base)
                p.setClipRect(seg.intersected(clip).intersected(limit), Qt.ReplaceClip)
                p.setPen(pen)
                p.drawText(inner, flags, text)
                cc += 1
            p.restore()
            return
        p.drawText(inner, flags, text)
        # a small mark when a compact row hides part of the text
        if st.wrap and self.row_mode == "compact":
            need = st.fm.boundingRect(QRect(0, 0, max(1, inner.width()), 100000), Qt.TextWordWrap, text).height()
            if need > inner.height() + 2:
                m = QRect(rect.right() - 16, rect.bottom() - 9, 14, 8)
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(self.theme["accent"]))
                for i in range(3):
                    p.drawEllipse(QPoint(m.left() + 3 + i * 4, m.center().y()), 1, 1)
        p.restore()

    def selection(self):
        (r1, c1), (r2, c2) = self.anchor, self.cur
        r1, r2 = sorted((r1, r2))
        c1, c2 = sorted((c1, c2))
        # grow to cover merged cells that are partly inside
        changed = True
        while changed:
            changed = False
            for (mr, mc), (mr2, mc2) in self._merges.items():
                if mr <= r2 and mr2 >= r1 and mc <= c2 and mc2 >= c1:
                    if mr < r1 or mr2 > r2 or mc < c1 or mc2 > c2:
                        r1, r2, c1, c2 = min(r1, mr), max(r2, mr2), min(c1, mc), max(c2, mc2)
                        changed = True
        return r1, c1, r2, c2

    def _paint_selection(self, p, rows, cols):
        if self.ws is None:
            return
        r1, c1, r2, c2 = self.selection()
        t = self.theme
        acc = QColor(t["accent"])
        x1, y1 = self.x_of(c1), self.y_of(r1)
        x2 = self.x_of(c2) + self.colx[c2] - self.colx[c2 - 1] - self.gap
        y2 = self.y_of(r2) + self.rowy[r2] - self.rowy[r2 - 1] - self.gap
        sel = QRect(x1, y1, x2 - x1, y2 - y1)
        if (r1, c1) != (r2, c2) and not self._is_single_merge(r1, c1, r2, c2):
            # tint the selection, leaving the active cell (where typing goes) clear
            fill = QColor(acc)
            fill.setAlpha(45)
            area = QRegion(sel).subtracted(QRegion(self.cell_rect(*self.cur)))
            for part in area:
                p.fillRect(part, fill)
        p.setBrush(Qt.NoBrush)
        p.setPen(QPen(acc, 2))
        p.drawRect(sel.adjusted(0, 0, -1, -1))

    def _is_single_merge(self, r1, c1, r2, c2):
        return self._merges.get((r1, c1)) == (r2, c2)

    def _paint_headers(self, p):
        t = self.theme
        vp = self.viewport().rect()
        hw, hh = self.header_w, self.header_h
        p.setFont(self._ui_font())
        r1, c1, r2, c2 = self.selection()
        hdr, sel_bg = QColor(t["header"]), QColor(t["panel2"])
        p.fillRect(QRect(0, 0, vp.width(), hh), hdr)
        p.fillRect(QRect(0, 0, hw, vp.height()), hdr)
        line = QPen(QColor(t["grid"]), 1)
        txt = QColor(t["header_text"])
        acc = QColor(t["accent"])
        for frozen in (False, True):
            area = QRect(self.ox + (0 if frozen else self.colx[self.fcol - 1]), 0, vp.width(), hh)
            cols = self._visible(self.colx, area.left(), vp.right(), frozen, False)
            p.save()
            p.setClipRect(area)
            for c in cols:
                x = self.x_of(c)
                w = self.colx[c] - self.colx[c - 1]
                if w <= 0:
                    continue
                rect = QRect(x, 0, w, hh)
                if c1 <= c <= c2:
                    p.fillRect(rect, sel_bg)
                p.setPen(line)
                p.drawLine(rect.right(), 2, rect.right(), hh - 3)
                p.setPen(acc if c1 <= c <= c2 else txt)
                p.drawText(rect, Qt.AlignCenter, get_column_letter(c))
            p.restore()
            area = QRect(0, self.oy + (0 if frozen else self.rowy[self.frow - 1]), hw, vp.height())
            rows = self._visible(self.rowy, area.top(), vp.bottom(), frozen, True)
            p.save()
            p.setClipRect(area)
            for r in rows:
                y = self.y_of(r)
                h = self.rowy[r] - self.rowy[r - 1]
                if h <= 0:
                    continue
                rect = QRect(0, y, hw, h)
                if r1 <= r <= r2:
                    p.fillRect(rect, sel_bg)
                p.setPen(line)
                p.drawLine(3, rect.bottom(), hw - 3, rect.bottom())
                p.setPen(acc if r1 <= r <= r2 else txt)
                p.drawText(rect.adjusted(0, 0, -6, 0), Qt.AlignRight | Qt.AlignVCenter, str(r))
            p.restore()
        p.fillRect(QRect(0, 0, hw, hh), hdr)

    # ------------------------------------------------------------ selection & keys
    def set_current(self, r, c, extend=False):
        r, c = max(1, r), max(1, c)
        if (r, c) in self._covered:
            r, c = self._covered[(r, c)]
        self.cur = (r, c)
        if not extend:
            self.anchor = (r, c)
        self._ensure_extent(r, c)
        self.ensure_visible(r, c)
        self.viewport().update()
        self.currentChanged.emit()

    def select_range(self, r1, c1, r2, c2):
        self.anchor, self.cur = (r1, c1), (r2, c2)
        self.ensure_visible(r2, c2)
        self.viewport().update()
        self.currentChanged.emit()

    def move(self, dr, dc, extend=False):
        r, c = self.cur
        if not extend:
            span = self._merges.get((r, c))
            if span:
                if dr > 0:
                    r = span[0]
                if dc > 0:
                    c = span[1]
        self.set_current(r + dr, c + dc, extend)

    def _has(self, r, c):
        cell = self.ws._cells.get((r, c))
        return cell is not None and cell.value not in (None, "")

    def jump(self, dr, dc, extend=False):
        """Ctrl+arrow: to the edge of the block of filled cells."""
        r, c = self.cur
        maxr, maxc = max(self.ws.max_row, 1), max(self.ws.max_column, 1)
        def inside(rr, cc):
            return 1 <= rr <= maxr and 1 <= cc <= maxc
        nr, nc = r + dr, c + dc
        if self._has(r, c) and inside(nr, nc) and self._has(nr, nc):
            while inside(nr + dr, nc + dc) and self._has(nr + dr, nc + dc):
                nr, nc = nr + dr, nc + dc
        else:
            while inside(nr, nc) and not self._has(nr, nc):
                nr, nc = nr + dr, nc + dc
            if not inside(nr, nc):
                nr = max(1, min(nr, maxr)) if dr else r
                nc = max(1, min(nc, maxc)) if dc else c
        self.set_current(nr, nc, extend)

    def keyPressEvent(self, e):
        if self.ws is None:
            return
        k, mods = e.key(), e.modifiers()
        shift = bool(mods & Qt.ShiftModifier)
        ctrl = bool(mods & Qt.ControlModifier)
        arrows = {Qt.Key_Up: (-1, 0), Qt.Key_Down: (1, 0), Qt.Key_Left: (0, -1), Qt.Key_Right: (0, 1)}
        if k in arrows:
            dr, dc = arrows[k]
            (self.jump if ctrl else self.move)(dr, dc, shift)
            return
        if k in (Qt.Key_PageDown, Qt.Key_PageUp):
            rows_per_page = max(1, self.viewport().height() // max(1, int(20 * self.zoom)))
            self.move(rows_per_page if k == Qt.Key_PageDown else -rows_per_page, 0, shift)
            return
        if k == Qt.Key_Home:
            self.set_current(1 if ctrl else self.cur[0], 1, shift)
            return
        if k == Qt.Key_End and ctrl:
            self.set_current(self.ws.max_row, self.ws.max_column, shift)
            return
        if k in (Qt.Key_Return, Qt.Key_Enter):
            self.move(-1 if shift else 1, 0)
            return
        if k == Qt.Key_Tab:
            self.move(0, 1)
            return
        if k == Qt.Key_Backtab:
            self.move(0, -1)
            return
        if k == Qt.Key_F2:
            self.start_edit()
            return
        if k in (Qt.Key_Delete, Qt.Key_Backspace) and not ctrl:
            if k == Qt.Key_Backspace:
                self.start_edit("")
            else:
                self.clear_contents()
            return
        if e.matches(QKeySequence.Copy):
            self.copy()
            return
        if e.matches(QKeySequence.Cut):
            self.copy(cut=True)
            return
        if e.matches(QKeySequence.Paste):
            self.paste()
            return
        if e.matches(QKeySequence.SelectAll):
            self.select_range(1, 1, max(1, self.ws.max_row), max(1, self.ws.max_column))
            return
        if ctrl and k == Qt.Key_D:
            self.fill("down")
            return
        if ctrl and k == Qt.Key_R:
            self.fill("right")
            return
        if ctrl and k in (Qt.Key_Equal, Qt.Key_Plus):
            self.set_zoom(self.zoom * 1.1)
            return
        if ctrl and k == Qt.Key_Minus:
            self.set_zoom(self.zoom / 1.1)
            return
        if ctrl and k == Qt.Key_0:
            self.set_zoom(1.0)
            return
        text = e.text()
        if text and text.isprintable() and not ctrl and not (mods & Qt.AltModifier):
            self.start_edit(text)
            return
        super().keyPressEvent(e)

    # ------------------------------------------------------------ mouse
    def _border_hit(self, pos):
        """('col', index) / ('row', index) when the mouse is on a header edge."""
        if not self.show_headers:
            return None
        row, col = self.hit(pos)
        if row == 0 and col > 0:
            for c in (col - 1, col):
                if c >= 1 and abs(pos.x() - (self.x_of(c) + self.colx[c] - self.colx[c - 1])) <= 3:
                    return ("col", c)
        if col == 0 and row > 0:
            for r in (row - 1, row):
                if r >= 1 and abs(pos.y() - (self.y_of(r) + self.rowy[r] - self.rowy[r - 1])) <= 3:
                    return ("row", r)
        return None

    def mousePressEvent(self, e):
        if self.ws is None:
            return
        self.setFocus()
        self.commit_if_editing()
        pos = e.position().toPoint()
        edge = self._border_hit(pos)
        if edge and e.button() == Qt.LeftButton:
            kind, idx = edge
            start = self.colx[idx] - self.colx[idx - 1] if kind == "col" else self.rowy[idx] - self.rowy[idx - 1]
            self._drag = ("resize", kind, idx, pos, start, start)
            return
        row, col = self.hit(pos)
        if e.button() == Qt.RightButton:
            r1, c1, r2, c2 = self.selection()
            if row and col and not (r1 <= row <= r2 and c1 <= col <= c2):
                self.set_current(row, col)
            return
        if e.button() != Qt.LeftButton:
            return
        if row == 0 and col == 0:
            self.select_range(1, 1, max(1, self.ws.max_row), max(1, self.ws.max_column))
            return
        if row == 0:
            if e.modifiers() & Qt.ShiftModifier:
                self.anchor = (1, self.anchor[1])
                self.cur = (self.nrows - EXTRA_ROWS, col)
            else:
                self.anchor, self.cur = (1, col), (max(1, self.ws.max_row), col)
            self._drag = ("cols", col)
            self.viewport().update()
            self.currentChanged.emit()
            return
        if col == 0:
            if e.modifiers() & Qt.ShiftModifier:
                self.anchor = (self.anchor[0], 1)
                self.cur = (row, max(1, self.ws.max_column))
            else:
                self.anchor, self.cur = (row, 1), (row, max(1, self.ws.max_column))
            self._drag = ("rows", row)
            self.viewport().update()
            self.currentChanged.emit()
            return
        if self.tool is not None and self.tool.press(row, col, e):
            self._drag = ("tool",)
            return
        if e.modifiers() & Qt.ControlModifier:
            cell = self.ws._cells.get((row, col))
            if cell is not None and cell.hyperlink is not None:
                link = cell.hyperlink
                target = link.target or ("#" + link.location if link.location else None)
                if target:
                    self.linkActivated.emit(target)
                    return
        self.set_current(row, col, bool(e.modifiers() & Qt.ShiftModifier))
        self._drag = ("cells",)

    def mouseMoveEvent(self, e):
        pos = e.position().toPoint()
        if self._drag is None:
            edge = self._border_hit(pos)
            if edge:
                self.viewport().setCursor(Qt.SplitHCursor if edge[0] == "col" else Qt.SplitVCursor)
            else:
                self.viewport().unsetCursor()
            # say what the frozen-pane divider is when the mouse is on it
            tip = ""
            if self.ws is not None and self.ws.freeze_panes:
                fx_line = self.ox + self.colx[self.fcol - 1]
                fy_line = self.oy + self.rowy[self.frow - 1]
                if (self.fcol > 1 and abs(pos.x() - fx_line) <= 3) or (self.frow > 1 and abs(pos.y() - fy_line) <= 3):
                    tip = (f"Frozen panes: everything above and left of {self.ws.freeze_panes} stays put while "
                           "you scroll.\nView → Freeze panes, or right-click a cell, to change or unfreeze.")
            if tip != self.viewport().toolTip():
                self.viewport().setToolTip(tip)
            if self.tool is not None:
                row, col = self.hit(pos)
                self.tool.hover(row, col)
            return
        kind = self._drag[0]
        if kind == "resize":
            _, axis, idx, start_pos, start, _ = self._drag
            delta = (pos.x() - start_pos.x()) if axis == "col" else (pos.y() - start_pos.y())
            size = max(4, start + delta)
            self._drag = ("resize", axis, idx, start_pos, start, size)
            self._preview_size(axis, idx, size)
            return
        row, col = self.hit(pos)
        if kind == "tool":
            if row and col:
                self.tool.move(row, col, e)
            return
        if kind == "cells" and row and col:
            if (row, col) != self.cur:
                self.cur = (row, col)
                self.ensure_visible(row, col)
                self.viewport().update()
                self.currentChanged.emit()
        elif kind == "cols" and col:
            self.cur = (self.cur[0], col)
            self.viewport().update()
            self.currentChanged.emit()
        elif kind == "rows" and row:
            self.cur = (row, self.cur[1])
            self.viewport().update()
            self.currentChanged.emit()

    def mouseReleaseEvent(self, e):
        drag, self._drag = self._drag, None
        if not drag:
            return
        if drag[0] == "resize":
            _, axis, idx, _, start, size = drag
            if size != start:
                self.set_size(axis, idx, size)
        elif drag[0] == "tool":
            self.tool.release(e)

    def mouseDoubleClickEvent(self, e):
        pos = e.position().toPoint()
        edge = self._border_hit(pos)
        if edge:
            self.autofit(*edge)
            return
        row, col = self.hit(pos)
        if row and col and (self.tool is None or not self.tool.active()):
            self.set_current(row, col)
            self.start_edit()

    def wheelEvent(self, e):
        if e.modifiers() & Qt.ControlModifier:
            steps = e.angleDelta().y() / 120
            if steps:
                self.set_zoom(self.zoom * (1.1 ** steps))
            return
        dy = e.pixelDelta().y() or int(e.angleDelta().y() / 120 * 60)
        dx = e.pixelDelta().x() or int(e.angleDelta().x() / 120 * 60)
        if e.modifiers() & Qt.ShiftModifier and not dx:
            dx, dy = dy, 0
        self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - dx)
        self.verticalScrollBar().setValue(self.verticalScrollBar().value() - dy)

    def contextMenuEvent(self, e):
        if self.ws is None:
            return
        menu = QMenu(self)
        self.contextRequested.emit(menu)
        if not menu.isEmpty():
            menu.exec(e.globalPos())

    # ------------------------------------------------------------ sizes
    def _preview_size(self, axis, idx, px):
        edges = self.colx if axis == "col" else self.rowy
        delta = px - (edges[idx] - edges[idx - 1])
        for i in range(idx, len(edges)):
            edges[i] += delta
        self._update_scrollbars()
        self.viewport().update()

    def set_size(self, axis, idx, px, idxs=None):
        """Set a column width / row height in screen pixels (undoable)."""
        r1, c1, r2, c2 = self.selection()
        if idxs is None:
            if axis == "col" and r1 == 1 and r2 >= self.ws.max_row and c1 <= idx <= c2:
                idxs = range(c1, c2 + 1)
            elif axis == "row" and c1 == 1 and c2 >= self.ws.max_column and r1 <= idx <= r2:
                idxs = range(r1, r2 + 1)
            else:
                idxs = [idx]
        px = max(4, px - self.gap)
        self.book.begin([self.ws])
        for i in idxs:
            if axis == "col":
                d = self.ws.column_dimensions[get_column_letter(i)]
                d.width = round(max(0.5, (px / self.zoom - 5) / MDW), 2)
            else:
                self.ws.row_dimensions[i].height = round(max(3, px / self.zoom * 72 / 96), 2)
                self._pin(i)
        self.book.done("layout")

    def _pin(self, row):
        pinned = self.book.meta.setdefault("pinned_rows", {}).setdefault(self.ws.title, [])
        if row not in pinned:
            pinned.append(row)
            pinned.sort()

    def autofit(self, axis, idx):
        if axis == "row":
            # fit the row to all of its text, even in compact mode
            self._layout()
            default_h = (self.ws.sheet_format.defaultRowHeight or DEFAULT_ROW_HEIGHT) * 96 / 72
            r1, c1, r2, c2 = self.selection()
            rows = range(r1, r2 + 1) if (c1 == 1 and c2 >= self.ws.max_column and r1 <= idx <= r2) else [idx]
            self.book.begin([self.ws])
            for r in rows:
                px = max(self._fit.get(r, 0), default_h)
                self.ws.row_dimensions[r].height = round(px * 72 / 96, 2)
                self._pin(r)
            self.book.done("layout")
            return
        best = 0
        for (r, c), cell in self.ws._cells.items():
            if c != idx or cell.value is None or (r, c) in self._merges or (r, c) in self._covered:
                continue
            st = self.style_of(cell)
            if st.wrap:
                continue
            text = self.book.display_text(self.ws, r, c)
            w = max(st.fm.horizontalAdvance(line) for line in text.split("\n")) if text else 0
            best = max(best, w + 2 * self.pad_x + 4 + self.gap)
        if best:
            self.set_size("col", idx, min(best, int(900 * self.zoom)), [idx])

    # ------------------------------------------------------------ editing
    def start_edit(self, initial=None):
        if self.ws is None or self.book is None:
            return
        r, c = self.cur
        if (r, c) in self._covered:
            r, c = self._covered[(r, c)]
        self.cancel_edit()
        self.edit_cell = (r, c)
        ed = self.editor = CellEditor(self)
        cell = self.ws._cells.get((r, c))
        st = self.style_of(cell) if cell is not None else None
        if st is not None:
            ed.setFont(st.font)
        else:
            f = QFont("Calibri")
            f.setPixelSize(round(11 * 96 / 72 * self.zoom))
            ed.setFont(f)
        t = self.theme
        bg = qcolor(st.fill) if st is not None and st.fill else QColor(t["canvas"])
        fg = self.text_color(st) if st is not None else QColor(t["text"])
        ed.setStyleSheet(f"QPlainTextEdit {{ background: {bg.name()}; color: {fg.name()}; "
                         f"border: 2px solid {t['accent']}; border-radius: 0; padding: 0; }}")
        ed.setPlainText(self.book.edit_text(self.ws, r, c) if initial is None else initial)
        self._place_editor()
        ed.show()
        ed.setFocus()
        ed.moveCursor(ed.textCursor().MoveOperation.End)

    def _place_editor(self):
        if self.editor is None:
            return
        rect = self.cell_rect(*self.edit_cell)
        rect = QRect(rect.left() - 1, rect.top() - 1, max(rect.width() + 2, 120), max(rect.height() + 2, 26))
        self.editor_base = rect
        self.editor.setGeometry(rect)
        self.editor.grow()

    def is_editing(self):
        return self.editor is not None and self.editor.isVisible()

    def commit_if_editing(self):
        if self.is_editing():
            self.finish_edit(True, None)

    def cancel_edit(self):
        if self.editor is not None:
            ed, self.editor = self.editor, None
            ed.hide()
            ed.deleteLater()
            self.edit_cell = None

    def finish_edit(self, commit, step=(1, 0)):
        if self.editor is None:
            return
        text = self.editor.toPlainText()
        r, c = self.edit_cell
        self.cancel_edit()
        self.setFocus()
        if commit and text != self.book.edit_text(self.ws, r, c):
            self.book.begin([self.ws])
            self.book.set_value(self.ws, r, c, text)
            self.book.done()
        if step:
            self.set_current(r, c)
            self.move(*step)

    def set_cell_text(self, r, c, text):
        """Used by the inspector."""
        if text == self.book.edit_text(self.ws, r, c):
            return
        self.book.begin([self.ws])
        self.book.set_value(self.ws, r, c, text)
        self.book.done()

    def clear_contents(self):
        r1, c1, r2, c2 = self.selection()
        cells = [cell for cell in self.book.cells_in(self.ws, r1, c1, r2, c2, create=False)
                 if cell.value is not None and not isinstance(cell, MergedCell)]
        if not cells:
            return
        self.book.begin([self.ws])
        for cell in cells:
            cell.value = None
        self.book.done()

    # ------------------------------------------------------------ clipboard
    def copy(self, cut=False):
        r1, c1, r2, c2 = self.selection()
        rows = []
        grid = []
        for r in range(r1, r2 + 1):
            line, items = [], []
            for c in range(c1, c2 + 1):
                cell = self.ws._cells.get((r, c))
                line.append(self.book.display_text(self.ws, r, c))
                if cell is None:
                    items.append(None)
                else:
                    items.append((cell.value, copy.copy(cell.font), copy.copy(cell.fill),
                                  copy.copy(cell.border), copy.copy(cell.alignment), cell.number_format))
            rows.append(line)
            grid.append(items)
        buf = io.StringIO()
        csv.writer(buf, delimiter="\t", lineterminator="\r\n").writerows(rows)
        text = buf.getvalue()
        QApplication.clipboard().setText(text)
        merges = [(m.min_row - r1, m.min_col - c1, m.max_row - r1, m.max_col - c1)
                  for m in self.ws.merged_cells.ranges
                  if m.min_row >= r1 and m.max_row <= r2 and m.min_col >= c1 and m.max_col <= c2]
        self.clip.update(text=text, grid=grid, origin=(r1, c1), cut=cut, ws=self.ws, merges=merges,
                         size=(r2 - r1 + 1, c2 - c1 + 1))

    def paste(self):
        text = QApplication.clipboard().text()
        r0, c0 = self.selection()[:2]
        clip = self.clip
        self.book.begin()
        if clip.get("text") == text and clip.get("grid") is not None:
            src_r, src_c = clip["origin"]
            if clip.get("cut"):
                src = clip["ws"]
                if src in self.book.wb.worksheets:
                    nr, nc = clip["size"]
                    self.book.unmerge(src, src_r, src_c, src_r + nr - 1, src_c + nc - 1)
                    for cell in self.book.cells_in(src, src_r, src_c, src_r + nr - 1, src_c + nc - 1, create=False):
                        cell.value = None
                        cell.style = "Normal"
            for i, items in enumerate(clip["grid"]):
                for j, item in enumerate(items):
                    r, c = r0 + i, c0 + j
                    cell = self.ws.cell(r, c)
                    if isinstance(cell, MergedCell):
                        self.book.unmerge(self.ws, r, c, r, c)
                        cell = self.ws.cell(r, c)
                    if item is None:
                        cell.value = None
                        continue
                    value, font, fill, border, align, numfmt = item
                    if fx.is_formula(value) and not clip.get("cut"):
                        value = fx.translate(value, r - (src_r + i), c - (src_c + j))
                    cell.value = value
                    cell.font, cell.fill, cell.border, cell.alignment = font, fill, border, align
                    cell.number_format = numfmt
            for a, b, c, d in clip.get("merges", []):
                self.book.merge(self.ws, r0 + a, c0 + b, r0 + c, c0 + d)
            if clip.get("cut"):
                clip.clear()
            rows_n, cols_n = len(clip.get("grid") or [[]]), len((clip.get("grid") or [[]])[0])
        else:
            rows = list(csv.reader(io.StringIO(text), delimiter="\t"))
            if text.endswith("\n") and rows and rows[-1] == []:
                rows.pop()
            for i, line in enumerate(rows):
                for j, value in enumerate(line):
                    cell = self.ws.cell(r0 + i, c0 + j)
                    if not isinstance(cell, MergedCell):
                        self.book.set_value(self.ws, r0 + i, c0 + j, value)
            rows_n, cols_n = len(rows), max((len(x) for x in rows), default=1)
        self.book.done()
        self.select_range(r0, c0, r0 + max(1, rows_n) - 1, c0 + max(1, cols_n) - 1)

    def fill(self, direction):
        r1, c1, r2, c2 = self.selection()
        if (direction == "down" and r1 == r2) or (direction == "right" and c1 == c2):
            return
        self.book.begin([self.ws])
        if direction == "down":
            for c in range(c1, c2 + 1):
                src = self.ws.cell(r1, c)
                for r in range(r1 + 1, r2 + 1):
                    self._copy_into(src, self.ws.cell(r, c), r - r1, 0)
        else:
            for r in range(r1, r2 + 1):
                src = self.ws.cell(r, c1)
                for c in range(c1 + 1, c2 + 1):
                    self._copy_into(src, self.ws.cell(r, c), 0, c - c1)
        self.book.done()

    @staticmethod
    def _copy_into(src, dst, dr, dc):
        if isinstance(dst, MergedCell):
            return
        v = src.value
        dst.value = fx.translate(v, dr, dc) if fx.is_formula(v) else v
        if src.has_style:
            dst._style = copy.copy(src._style)


def _lum(hex6):
    from atlas_theme import luminance
    return luminance(hex6)
