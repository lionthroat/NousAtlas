"""Draws the app icon (a folded map with a painted range) with Qt: icon.ico
for Windows and a 1024px icon.png for other platforms.
Run before building (build.ps1 does it); commit both files."""

import os
import sys

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QGuiApplication, QImage, QPainter, QPen, QPolygonF


def draw(size):
    """The artwork is laid out on a 256px grid and scaled to `size`."""
    img = QImage(size, size, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    p.scale(size / 256, size / 256)
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#2e3440"))
    p.drawRoundedRect(QRectF(8, 8, 240, 240), 44, 44)
    # three folded panels
    xs = [36, 98, 158, 220]
    tops = [58, 44, 58, 44]
    bottoms = [212, 198, 212, 198]
    shades = ["#ebcb8b", "#d9b878", "#ebcb8b"]
    for i in range(3):
        poly = QPolygonF([QPointF(xs[i], tops[i]), QPointF(xs[i + 1], tops[i + 1]),
                          QPointF(xs[i + 1], bottoms[i + 1]), QPointF(xs[i], bottoms[i])])
        p.setBrush(QColor(shades[i]))
        p.setPen(QPen(QColor("#4c566a"), 4, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.drawPolygon(poly)
    # a painted range: violet blob with a pink night half
    p.setPen(Qt.NoPen)
    violet = QColor("#a58cf0")
    violet.setAlpha(225)
    p.setBrush(violet)
    p.drawEllipse(QRectF(70, 92, 88, 70))
    pink = QColor("#f28ad6")
    pink.setAlpha(230)
    p.setBrush(pink)
    p.drawPolygon(QPolygonF([QPointF(146, 112), QPointF(196, 104), QPointF(190, 164), QPointF(134, 160)]))
    # grid lines
    p.setPen(QPen(QColor(46, 52, 64, 90), 3))
    for y in (100, 140, 180):
        p.drawLine(QPointF(40, y + 4), QPointF(216, y - 4))
    p.end()
    return img


app = QGuiApplication(sys.argv)
here = os.path.dirname(os.path.abspath(__file__))
for name, size in (("icon.ico", 256), ("icon.png", 1024)):
    out = os.path.join(here, name)
    print("saved" if draw(size).save(out) else "FAILED", out)
