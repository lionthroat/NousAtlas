r"""Palette, colour swapping and styles, on a copy of the sample workbook.

Run from the project folder:  .venv\Scripts\python.exe tests\test_palette.py"""
import os, sys
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", "C:/Windows/Fonts")
os.environ["NOUSATLAS_SETTINGS"] = "NousAtlasTest"
import faulthandler
faulthandler.dump_traceback_later(120, exit=True)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import sample_copy
from PySide6.QtWidgets import QApplication, QInputDialog, QMessageBox
import openpyxl
import nousatlas, atlas_dialogs as D
from atlas_model import resolve_color

QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.Yes)
QMessageBox.information = staticmethod(lambda *a, **k: None)
path = sample_copy()
app = QApplication(sys.argv)
w = nousatlas.MainWindow(); w.show()
w.open_path(path)
book = w.book
gz = book.sheet("Gazetteer")
assert resolve_color(gz["A4"].fill.fgColor, book.theme) == "3F6F53"

# a palette from Nord
dlg = D.PaletteDialog(w); dlg.add_builtin("Nord"); dlg.accept()
assert book.meta["palette"]["name"] == "Nord" and len(book.meta["palette"]["colors"]) == 16

# swap every colour for its nearest Nord colour
dlg = D.RemapDialog(w); dlg.suggest(); dlg.accept()
nord = {c["hex"] for c in book.meta["palette"]["colors"]}
used = set(book.colors_in_use())
assert used <= nord, used - nord
w.undo()
assert resolve_color(gz["A4"].fill.fgColor, book.theme) == "3F6F53"

# styles from repeated formatting
combos = []
orig_exec = D.SuggestStylesDialog.exec
def fake_exec(self):
    for i, (key, edit) in enumerate(self.rows[:2]):
        edit.setText(["Body", "Header"][i])
    return True
D.SuggestStylesDialog.exec = fake_exec
w.suggest_styles()
assert "Body" in book.wb.named_styles and "Header" in book.wb.named_styles
body_cells = [c for ws in book.user_sheets() for c in ws._cells.values() if c.has_style and c.style == "Body"]
assert len(body_cells) > 100, len(body_cells)
# change the style through a cell, everything follows
w.open_sheet(body_cells[0].parent.title)
w.grid.set_current(body_cells[0].row, body_cells[0].column)
w.format_cells(lambda c: None)
import atlas_model as M
book.begin(); M.set_font(body_cells[0], color="FF88C0D0"); book.done()
w.update_style()
assert all(resolve_color(c.font.color, book.theme) == "88C0D0" for c in body_cells[:50])
assert w.save()
wb = openpyxl.load_workbook(path)
assert "Body" in wb.named_styles
print("OK")
