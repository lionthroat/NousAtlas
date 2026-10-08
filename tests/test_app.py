r"""The whole window on a copy of the sample workbook: sidebar, editing, undo,
rows, find, range layers, save and reopen.

Run from the project folder:  .venv\Scripts\python.exe tests\test_app.py
Opens a real (offscreen) window; prints OK at the end."""
import os, sys
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", "C:/Windows/Fonts")
os.environ["NOUSATLAS_SETTINGS"] = "NousAtlasTest"
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import sample_copy
from PySide6.QtCore import Qt
import faulthandler
faulthandler.dump_traceback_later(120, exit=True)      # never hang on a stray dialog
from PySide6.QtWidgets import QApplication, QInputDialog, QMessageBox
import openpyxl
import nousatlas, atlas_ranges as R

QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.Yes)
QMessageBox.information = staticmethod(lambda *a, **k: None)
QMessageBox.warning = staticmethod(lambda *a, **k: None)
QInputDialog.getItem = staticmethod(lambda parent, title, label, items, *a, **k: (items[0], True))

path = sample_copy()
app = QApplication(sys.argv)
w = nousatlas.MainWindow(); w.resize(1400, 900); w.show()
assert w.open_path(path)
book = w.book

# sidebar: maps found and grouped, every sheet listed once
assert "MAP — Day One" in book.meta["maps"], book.meta["maps"]
assert book.meta["maps"]["MAP — Day One"] == {"row": 5, "col": 2, "rows": 15, "cols": 17}, book.meta["maps"]["MAP — Day One"]
assert [g["name"] for g in book.meta["groups"]] == ["Maps"]
assert len(w.sidebar_order()) == 21

# go to a sheet by typing
w.sidebar.filter.setText("fau"); w.sidebar.open_first()
assert w.ws_title == "Fauna"
w.open_sheet("Gazetteer")
assert w.history[-1][0] == "Fauna"
w.go_back(); assert w.ws_title == "Fauna"
w.go_forward(); assert w.ws_title == "Gazetteer"

# edit through the grid, then undo
g = w.grid
g.set_current(5, 5); g.start_edit("10"); g.finish_edit(True)
ws = book.sheet("Gazetteer")
assert ws["E5"].value == 10 and book.display_text(ws, 5, 7) == "5"
assert g.cur == (6, 5)
w.undo(); assert ws["E5"].value == 9
w.redo(); assert ws["E5"].value == 10

# inspector edits
g.set_current(5, 2)
assert w.inspector.text.toPlainText() == "The Springs"
w.inspector.text.setPlainText("The Springs (cold)"); w.inspector.commit_text()
assert ws["B5"].value == "The Springs (cold)"

# formatting and merge
g.select_range(5, 1, 5, 2); w.toggle_font("i")
assert ws["A5"].font.i and ws["B5"].font.i
g.select_range(2, 1, 2, 4); w.toggle_merge()
assert "A2:D2" in [str(m) for m in ws.merged_cells.ranges]
w.toggle_merge(); assert "A2:D2" not in [str(m) for m in ws.merged_cells.ranges]

# rows
g.select_range(6, 1, 6, 1); w.insert_rows(above=True)
assert ws["B7"].value == "The Horns" and ws["G7"].value.startswith("=ROUND(SQRT((E7")
w.undo(); assert ws["B6"].value == "The Horns"

# find across sheets
w.find_text("Saint's Hair")
assert len({h[0] for h in w.find_bar.hits}) >= 2, w.find_bar.hits
w.find_bar.step(1)
assert w.grid.find_current is not None

# range layers: make one from a Fauna name, paint on the Day One map
fauna = book.sheet("Fauna")
w.open_sheet("Fauna"); g.set_current(3, 1)
name = fauna["A3"].value
w.make_layer_for(name)
assert w.ws_title == "MAP — Day One" and w.range_state.painting and w.range_state.solo == name

class Ev:
    def __init__(self, mods=Qt.NoModifier): self.m = mods
    def button(self): return Qt.LeftButton
    def modifiers(self): return self.m

info = book.meta["maps"]["MAP — Day One"]
brush = w.grid.tool
w.ranges_panel.set_blocks(R.DAY)
r, c = R.square_cell(info, (9, 6))                 # I6
assert brush.press(r, c, Ev())
brush.move(r, c + 1, Ev()); brush.move(r, c + 2, Ev())
brush.release(Ev())
layer = book.ranges.get(name)
assert layer.level("MAP — Day One", 8, (9, 6)) == 3 and layer.level("MAP — Day One", 8, (11, 6)) == 3
assert layer.level("MAP — Day One", 20, (9, 6)) == 0
w.ranges_panel.set_blocks({20}); w.range_state.level = 2
assert brush.press(*R.square_cell(info, (1, 1)), Ev()); brush.release(Ev())
assert layer.level("MAP — Day One", 20, (1, 1)) == 2
who = book.ranges.who("MAP — Day One", (9, 6))
assert who and who[0][0].name == name
w.undo(); assert layer.level("MAP — Day One", 20, (1, 1)) == 0 or book.ranges.get(name).level("MAP — Day One", 20, (1, 1)) == 0
w.redo()

# a chimney-bat style layer: common at home from 16:00, uncommon around it, everywhere from 20:00
w.range_state.painting = False
bats = book.ranges.add("Chimney Bats")
book.ranges.paint(bats, "MAP — Day One", {16}, [(5, 3)], 3)
book.ranges.paint(bats, "MAP — Day One", {16}, [(4, 2), (5, 2), (6, 2), (4, 3), (6, 3), (4, 4), (5, 4), (6, 4)], 2)
book.ranges.paint(bats, "MAP — Day One", {20, 0, 4}, [(x, y) for x in range(1, 18) for y in range(1, 16)], 3)
book.dirty = True

# paint into a screenshot to make sure drawing doesn't fail
w.range_state.split = True; w.range_state.solo = None
w.right.setCurrentIndex(1); app.processEvents()
shot_dir = os.path.join(os.path.dirname(path), "shots"); os.makedirs(shot_dir, exist_ok=True)
w.grab().save(os.path.join(shot_dir, "ranges.png"))

# save, reopen: everything back
assert w.save()
wb = openpyxl.load_workbook(path)
assert "Ranges" in wb.sheetnames and wb["_NousAtlas"].sheet_state == "veryHidden"
rows = list(wb["Ranges"].iter_rows(values_only=True))
assert rows[0] == tuple(R.HEADERS), rows[0]
bat_rows = [r for r in rows if r[0] == "Chimney Bats"]
assert ("Chimney Bats", "MAP — Day One", "Night", "common", "A1:Q15", bats.color) in bat_rows, bat_rows
assert any(r[2] == "16:00" and r[3] == "common" and r[4] == "E3" for r in bat_rows), bat_rows
assert any(r[2] == "16:00" and r[3] == "uncommon" and R.parse_squares(r[4]) == {(4, 2), (5, 2), (6, 2), (4, 3), (6, 3), (4, 4), (5, 4), (6, 4)} for r in bat_rows), bat_rows
assert len(bat_rows) == 3, bat_rows
assert wb["Gazetteer"]["E5"].value == 10

w.book.dirty = False
w.close_book()
assert w.open_path(path)
b2 = w.book
assert b2.ranges.get("Chimney Bats").level("MAP — Day One", 0, (17, 15)) == 3
assert b2.ranges.get("Chimney Bats").level("MAP — Day One", 16, (5, 2)) == 2
assert b2.ranges.get(name).level("MAP — Day One", 12, (10, 6)) == 3
assert [g["name"] for g in b2.meta["groups"]] == ["Maps"]
assert "Ranges" not in [ws.title for ws in w.sidebar_order()]      # Atlas writes it; hidden from the sidebar
print("OK")
