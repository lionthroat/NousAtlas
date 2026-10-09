r"""Layer time schemes and rarity: none, Day/Night, four, hourly, custom;
rarity on and off; switching schemes keeps what it can; the Ranges sheet
round-trips. Run:  .venv\Scripts\python.exe tests\test_layer_schemes.py"""
import os, sys
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", "C:/Windows/Fonts")
os.environ["NOUSATLAS_SETTINGS"] = "NousAtlasTestSchemes"
import faulthandler
faulthandler.dump_traceback_later(120, exit=True)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import sample_copy
from PySide6.QtWidgets import QApplication, QInputDialog, QMessageBox
import nousatlas, atlas_ranges as R
import atlas_squares as SQ

QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.Yes)
QInputDialog.getText = staticmethod(lambda *a, **k: ("Scorpions", True))
path = sample_copy()
app = QApplication(sys.argv)
w = nousatlas.MainWindow(); w.resize(1400, 900); w.show()
w.open_path(path)
b = w.book
mp = "MAP — Day One"
info = b.meta["maps"][mp]
p = w.ranges_panel

def use(time, rarity):
    b.begin([]); b.meta["time"] = time; b.meta["rarity"] = rarity
    b.ranges.convert(R.TimeScheme.from_meta(b.meta)); w.range_state.blocks = None; w.range_state.view = "split"
    b.done("ranges")

def select(*sqs):
    (r1, c1), (r2, c2) = R.square_cell(info, sqs[0]), R.square_cell(info, sqs[-1])
    w.grid.select_range(r1, c1, r2, c2)

# --- a new workbook: no time, no rarity -> none of those controls show
w.open_sheet(mp); w.show_layers_tab()
assert not b.ranges.scheme.has_time() and R.rarity_of(b) is None
w.range_state.layer = None
select((7, 9), (11, 12)); w.new_layer_here()                       # asks for a name, then adds
sc = b.ranges.get("Scorpions")
assert sc.present(mp, (9, 10)) and not sc.present(mp, (1, 1))
p.refresh()
assert not p.view_row.isVisibleTo(p) and not p.opts.isVisibleTo(p), "no time / rarity controls"
w.grid.set_current(*R.square_cell(info, (9, 10))); p.update_here()
assert [p.here_table.horizontalHeaderItem(i).text() for i in range(1, p.here_table.columnCount())] == ["Here"]
html = SQ.card_html(w.squares.card(mp, (9, 10)), w.theme, mp)
assert "Scorpions" in html and "common" not in html.lower() and "all day" not in html.lower(), html
p.here_clicked(0, 1); assert not b.ranges.get("Scorpions").present(mp, (9, 10))   # toggles off
w.undo()

# --- Day / Night: data carried over to every period; split view in two
use({"preset": "daynight"}, None)
sc = b.ranges.get("Scorpions")
assert sc.level(mp, "day", (9, 10)) and sc.level(mp, "night", (9, 10))
p.refresh(); assert p.view_row.isVisibleTo(p) and p.for_box.isVisibleTo(p) and not p.as_box.isVisibleTo(p)
assert [p.for_box.itemText(i) for i in range(p.for_box.count())] == ["All day", "Day", "Night"]
w.range_state.layer = "Scorpions"
p.set_blocks({"night"}); select((1, 1)); w.add_selection_to_layer()
sc = b.ranges.get("Scorpions")
assert sc.level(mp, "night", (1, 1)) and not sc.level(mp, "day", (1, 1))
assert len(R.view_parts(b.ranges.scheme, "split")) == 2

# --- four periods: squares split in quarters
use({"preset": "four"}, ["Common", "Rare"])
assert len(R.view_parts(b.ranges.scheme, "split")) == 4 and len(R._part_shapes(w.grid.cell_rect(5, 5), 4)) == 4
assert b.ranges.get("Scorpions").present(mp, (9, 10))
p.refresh(); assert p.as_box.isVisibleTo(p) and [p.as_box.itemText(i) for i in range(p.as_box.count())] == ["Common", "Rare"]

# --- hourly: the square table uses the Day / Night groups, not 24 columns
use({"preset": "hourly"}, None)
w.grid.set_current(*R.square_cell(info, (9, 10))); p.update_here()
heads = [p.here_table.horizontalHeaderItem(i).text() for i in range(1, p.here_table.columnCount())]
assert heads == ["Day", "Night", "All"], heads

# --- custom periods with a group, and the Ranges sheet round trip
use({"preset": "custom", "periods": ["Dawn", "Noon", "Dusk", "Midnight"],
     "groups": {"Daylight": ["Dawn", "Noon", "Dusk"]}}, ["Plenty", "Some", "Few"])
w.range_state.layer = "Scorpions"; p.set_blocks({"Midnight"}); w.range_state.level = 1
select((2, 2)); w.add_selection_to_layer()
assert b.ranges.get("Scorpions").level(mp, "Midnight", (2, 2)) == 1
assert w.save()
w.close_book(); w.open_path(path); b = w.book; p = w.ranges_panel
assert b.ranges.scheme.ids == ["Dawn", "Noon", "Dusk", "Midnight"]
assert b.ranges.get("Scorpions").level(mp, "Midnight", (2, 2)) == 1
import openpyxl
rows = list(openpyxl.load_workbook(path)["Ranges"].iter_rows(values_only=True))
assert any(r[2] == "Midnight" and r[3] == "Few" and r[4] == "B2" for r in rows), rows
print("OK schemes")
