#!/usr/bin/env bash
# Builds nous-atlas_<version>_amd64.deb for Linux Mint / Ubuntu / Debian.
# Run on Linux from the project folder, with the venv set up
# (python3 -m venv .venv && .venv/bin/pip install -r requirements.txt):
#   bash build_linux.sh
# GitHub builds it on version tags (see .github/workflows/build.yml).
#
# The package installs the app to /opt/nous-atlas, a `nous-atlas` command,
# a menu entry, icons, and offers it for .xlsx workbooks in Open With.
set -euo pipefail
cd "$(dirname "$0")"
PY=.venv/bin/python
VERSION=$(tr -d '[:space:]' < VERSION)
PKG=nous-atlas

# --strip drops debug symbols (libpython alone goes from 32 MB to a few)
"$PY" -m PyInstaller --noconfirm --clean --windowed --name NousAtlas --strip \
    --add-data "$PWD/icon.ico:." --add-data "$PWD/VERSION:." \
    --distpath dist --workpath build --specpath build \
    --exclude-module PySide6.QtNetwork --exclude-module PySide6.QtQml \
    --exclude-module PySide6.QtQuick --exclude-module PySide6.QtPdf \
    "$PWD/nousatlas.py"

# Qt drags in pieces this app never uses (Wayland, EGLFS/framebuffer/VNC
# platforms, networking, SVG, OpenGL helpers, translations, image formats
# other than the .ico window icon). Drop them, as build.ps1 does on Windows.
# Kept: the X11 (xcb) platform, the GTK and portal themes (native file
# dialogs), input methods (compose key, IBus) and ICU, which Qt Core needs.
# Under a Wayland session Qt falls back to X11 (XWayland).
INT=dist/NousAtlas/_internal
QT=$INT/PySide6/Qt
rm -rf "$QT/translations" \
    "$QT"/lib/libQt6{Network,OpenGL,Svg,WaylandClient,WlShellIntegration}.so* \
    "$QT"/lib/libQt6{EglFS,EglFs,Quick,Qml,Pdf,VirtualKeyboard}*.so* \
    "$QT"/plugins/{egldeviceintegrations,generic,iconengines,tls,networkinformation,wayland-*}
find "$QT/plugins/platforms" -type f ! -name libqxcb.so -delete
find "$QT/plugins/imageformats" -type f ! -name libqico.so -delete

# PyInstaller also copies the build machine's own system libraries (GTK,
# GLib, X11, fontconfig, OpenSSL, the C++ runtime...). Use the user's instead
# (they're in Depends below): smaller, and an old bundled GLib or C++ runtime
# can clash with a newer system's GTK modules or graphics drivers.
find "$INT" -maxdepth 1 -name 'lib*.so*' ! -name 'libpython3*' -delete

# Nothing left behind may need something that was just deleted
MISSING=$(find dist/NousAtlas -type f \( -name '*.so*' -o -name NousAtlas \) -print0 |
    xargs -0 ldd 2>/dev/null | awk '/:$/ { f = $0 } /not found/ { print f, $0 }' || true)
if [ -n "$MISSING" ]; then echo "$MISSING"; echo "Pruned too much"; exit 1; fi

STAGE=build/deb
rm -rf "$STAGE"
mkdir -p "$STAGE/DEBIAN" "$STAGE/opt" "$STAGE/usr/bin" "$STAGE/usr/share/applications"
cp -a dist/NousAtlas "$STAGE/opt/$PKG"
ln -s "/opt/$PKG/NousAtlas" "$STAGE/usr/bin/$PKG"

# Icons at the sizes desktops look for
QT_QPA_PLATFORM=offscreen "$PY" - "$STAGE" "$PKG" <<'EOF'
import os, sys
from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication, QImage
app = QGuiApplication(sys.argv[:1])
stage, pkg = sys.argv[1], sys.argv[2]
src = QImage("icon.png")
for size in (48, 64, 128, 256, 512):
    d = os.path.join(stage, f"usr/share/icons/hicolor/{size}x{size}/apps")
    os.makedirs(d, exist_ok=True)
    src.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation).save(os.path.join(d, pkg + ".png"))
EOF

cat > "$STAGE/usr/share/applications/$PKG.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Nous Atlas
GenericName=Workbook
Comment=A text-first workbook for worldbuilding: sheets, notes and painted map layers
Exec=$PKG %f
Icon=$PKG
Terminal=false
Categories=Office;Spreadsheet;
MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;
Keywords=xlsx;spreadsheet;workbook;map;worldbuilding;
StartupWMClass=NousAtlas
EOF

SIZE_KB=$(du -sk "$STAGE" | cut -f1)
cat > "$STAGE/DEBIAN/control" <<EOF
Package: $PKG
Version: $VERSION
Section: editors
Priority: optional
Architecture: amd64
Maintainer: lionthroat <lionthroat@users.noreply.github.com>
Homepage: https://nousatlas.lionthroat.com
Installed-Size: $SIZE_KB
Depends: libxcb-cursor0, libxcb-icccm4, libxcb-image0, libxcb-keysyms1, libxcb-randr0, libxcb-render-util0, libxcb-shape0, libxcb-xinerama0, libxcb-xkb1, libxcb-util1, libxcb-shm0, libxcb-sync1, libxcb-xfixes0, libxcb-render0, libxcb1, libx11-6, libx11-xcb1, libxkbcommon0, libxkbcommon-x11-0, libegl1, libgl1, libfontconfig1, libfreetype6, libdbus-1-3, libglib2.0-0, libgtk-3-0, libstdc++6, libssl3, libffi8, zlib1g, libzstd1, libbz2-1.0, liblzma5
Description: Text-first workbook app for worldbuilding
 Nous Atlas opens .xlsx workbooks with many tabs of text: a grouped sheet
 sidebar, links between sheets, notes, and map sheets you can paint range
 layers onto. Free and open source (MIT).
EOF

# Files root-owned and world-readable, whatever the build machine's umask was
chmod -R u+rwX,go+rX,go-w "$STAGE"
OUT="${PKG}_${VERSION}_amd64.deb"
dpkg-deb --build --root-owner-group "$STAGE" "$OUT"
echo "Built: $OUT ($(du -h "$OUT" | cut -f1))"
