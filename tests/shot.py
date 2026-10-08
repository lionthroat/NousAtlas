"""Open a copy of the sample workbook and save screenshots (for eyeballing)."""
import os, sys
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", "C:/Windows/Fonts")
os.environ["NOUSATLAS_SETTINGS"] = "NousAtlasTest"
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import sample_copy
from PySide6.QtWidgets import QApplication
import nousatlas
out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(sample_copy()), "shots")
os.makedirs(out, exist_ok=True)
app = QApplication(sys.argv)
w = nousatlas.MainWindow(); w.resize(1500, 900); w.show()
w.open_path(sample_copy())
for name in sys.argv[2:] or ["Gazetteer", "MAP — Day One", "Fauna", "The Shrine"]:
    w.open_sheet(name)
    app.processEvents()
    w.grab().save(os.path.join(out, name.replace("—", "-").replace(" ", "_") + ".png"))
print(out)
