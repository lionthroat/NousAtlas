# PyInstaller recipe for the macOS app bundle, "Nous Atlas.app".
# Built by .github/workflows/build.yml on GitHub's Mac machines:
#   python -m PyInstaller --noconfirm --clean NousAtlas-mac.spec
# (needs Pillow, which turns icon.png into the .icns macOS wants)
import os

root = SPECPATH
with open(os.path.join(root, "VERSION")) as f:
    VERSION = f.read().strip()

a = Analysis(
    [os.path.join(root, "nousatlas.py")],
    datas=[(os.path.join(root, "icon.ico"), "."), (os.path.join(root, "VERSION"), ".")],
    excludes=["PySide6.QtNetwork", "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtPdf",
              "tkinter"],
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="NousAtlas", console=False)
coll = COLLECT(exe, a.binaries, a.datas, name="NousAtlas")
app = BUNDLE(
    coll,
    name="Nous Atlas.app",
    icon=os.path.join(root, "icon.png"),
    bundle_identifier="com.lionthroat.nousatlas",
    version=VERSION,
    info_plist={
        "CFBundleDisplayName": "Nous Atlas",
        "CFBundleShortVersionString": VERSION,
        "NSHighResolutionCapable": True,
        "LSMinimumSystemVersion": "12.0",
        # Lets Finder offer Nous Atlas in "Open With" for .xlsx (never the default unless chosen)
        "CFBundleDocumentTypes": [{
            "CFBundleTypeName": "Excel Workbook",
            "CFBundleTypeRole": "Editor",
            "LSItemContentTypes": ["org.openxmlformats.spreadsheetml.sheet"],
            "CFBundleTypeExtensions": ["xlsx"],
            "LSHandlerRank": "Alternate",
        }],
    },
)
