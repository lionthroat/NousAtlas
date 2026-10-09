"""Looks: the app's own themes (Nord, Dracula, Snow) and the built-in colour
palettes a workbook can use for its cells (Nord, Dracula, Biomes)."""

import colorsys

from PySide6.QtGui import QColor

# ---------------------------------------------------------------- UI themes

THEMES = {
    "Nord": {
        "window": "#2e3440", "panel": "#3b4252", "panel2": "#434c5e", "border": "#4c566a",
        "text": "#eceff4", "muted": "#a3abbd", "faint": "#7b88a1",
        "accent": "#88c0d0", "accent_text": "#2e3440", "select": "#5e81ac",
        "canvas": "#2e3440", "grid": "#3b4252", "header": "#353c4a", "header_text": "#a3abbd",
        "link": "#b48ead", "find": "#ebcb8b", "danger": "#bf616a", "group": "#ebcb8b",
        "dark": True,
    },
    "Dracula": {
        "window": "#282a36", "panel": "#21222c", "panel2": "#343746", "border": "#44475a",
        "text": "#f8f8f2", "muted": "#b6b9cc", "faint": "#6272a4",
        "accent": "#bd93f9", "accent_text": "#282a36", "select": "#6272a4",
        "canvas": "#282a36", "grid": "#343746", "header": "#21222c", "header_text": "#a7abc4",
        "link": "#ff79c6", "find": "#f1fa8c", "danger": "#ff5555", "group": "#8be9fd",
        "dark": True,
    },
    "Snow": {
        "window": "#eceff4", "panel": "#e5e9f0", "panel2": "#d8dee9", "border": "#c5ccd8",
        "text": "#2e3440", "muted": "#4c566a", "faint": "#7b88a1",
        "accent": "#5e81ac", "accent_text": "#ffffff", "select": "#81a1c1",
        "canvas": "#ffffff", "grid": "#e1e5ec", "header": "#e5e9f0", "header_text": "#4c566a",
        "link": "#b48ead", "find": "#d08770", "danger": "#bf616a", "group": "#5e81ac",
        "dark": False,
    },
}


def stylesheet(t):
    return f"""
    QWidget {{ background: {t['window']}; color: {t['text']}; font-size: 10pt; }}
    QMainWindow::separator {{ background: transparent; width: 4px; height: 4px; }}
    QMainWindow::separator:hover {{ background: {t['accent']}; }}
    QScrollArea#panelScroll {{ border: none; border-left: 1px solid {t['border']}; }}
    QToolBar {{ background: {t['panel']}; border: none; border-bottom: 1px solid {t['border']}; spacing: 3px; padding: 3px 6px; }}
    QToolBar::separator {{ background: {t['border']}; width: 1px; margin: 4px 5px; }}
    QToolButton {{ background: transparent; border: 1px solid transparent; border-radius: 5px; padding: 3px 6px; color: {t['text']}; }}
    QToolButton:hover {{ background: {t['panel2']}; }}
    QToolButton:checked {{ background: {t['select']}; color: {t['text']}; }}
    QToolButton::menu-indicator {{ image: none; }}
    QPushButton {{ background: {t['panel2']}; border: 1px solid {t['border']}; border-radius: 5px; padding: 4px 10px; }}
    QPushButton:hover {{ border-color: {t['accent']}; }}
    QPushButton:checked {{ background: {t['accent']}; color: {t['accent_text']}; border-color: {t['accent']}; }}
    QPushButton:default {{ border-color: {t['accent']}; }}
    QLineEdit, QPlainTextEdit, QTextEdit, QSpinBox, QDoubleSpinBox {{ background: {t['canvas']}; border: 1px solid {t['border']}; border-radius: 5px; padding: 3px 5px; selection-background-color: {t['select']}; }}
    QLineEdit:focus, QPlainTextEdit:focus {{ border-color: {t['accent']}; }}
    QComboBox, QFontComboBox {{ background: {t['panel2']}; border: 1px solid {t['border']}; border-radius: 5px; padding: 2px 6px; }}
    QComboBox QAbstractItemView {{ background: {t['panel']}; selection-background-color: {t['select']}; border: 1px solid {t['border']}; }}
    QTreeWidget, QListWidget, QTableWidget {{ background: {t['panel']}; border: none; outline: none; }}
    QTreeView {{ show-decoration-selected: 1; }}
    QListWidget::item {{ padding: 3px 2px; border-radius: 4px; }}
    QTreeWidget::item {{ padding: 3px 2px; border-radius: 0; }}
    QTreeWidget::item:selected, QListWidget::item:selected {{ background: {t['select']}; color: {t['text']}; }}
    QTreeWidget::item:hover, QListWidget::item:hover {{ background: {t['panel2']}; }}
    QHeaderView::section {{ background: {t['panel']}; color: {t['muted']}; border: none; padding: 3px; }}
    QMenu {{ background: {t['panel']}; border: 1px solid {t['border']}; padding: 4px; }}
    QMenu::item {{ padding: 4px 22px 4px 18px; border-radius: 4px; }}
    QMenu::item:selected {{ background: {t['select']}; }}
    QMenu::separator {{ height: 1px; background: {t['border']}; margin: 4px 6px; }}
    QMenuBar {{ background: {t['panel']}; }}
    QMenuBar::item:selected {{ background: {t['panel2']}; }}
    QTabWidget::pane {{ border: none; }}
    QTabBar::tab {{ background: transparent; color: {t['muted']}; padding: 5px 12px; border: none; border-bottom: 2px solid transparent; }}
    QTabBar::tab:selected {{ color: {t['text']}; border-bottom: 2px solid {t['accent']}; }}
    QScrollBar:vertical {{ background: transparent; width: 11px; margin: 0; }}
    QScrollBar:horizontal {{ background: transparent; height: 11px; margin: 0; }}
    QScrollBar::handle {{ background: {t['border']}; border-radius: 4px; min-height: 30px; min-width: 30px; margin: 2px; }}
    QScrollBar::handle:hover {{ background: {t['faint']}; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
    QScrollBar::add-page, QScrollBar::sub-page {{ background: none; }}
    QSplitter::handle {{ background: {t['border']}; }}
    QDockWidget {{ color: {t['muted']}; }}
    QDockWidget::title {{ background: {t['panel']}; padding: 5px 8px; border-bottom: 1px solid {t['border']}; }}
    QStatusBar {{ background: {t['panel']}; color: {t['muted']}; }}
    QLabel#muted {{ color: {t['muted']}; background: transparent; }}
    QLabel#faint {{ color: {t['faint']}; background: transparent; }}
    QLabel#heading {{ color: {t['group']}; font-weight: bold; background: transparent; }}
    QFrame#panel {{ background: {t['panel']}; }}
    QToolButton#sectionHeader {{ background: {t['panel']}; color: {t['group']}; font-weight: bold; text-align: left;
        border: none; border-bottom: 1px solid {t['border']}; border-radius: 0; padding: 6px 8px; }}
    QToolButton#sectionHeader:hover {{ background: {t['panel2']}; }}
    QFrame#workbox {{ background: {t['panel']}; border: 1px solid {t['accent']}; border-radius: 6px; }}
    QFrame#workbox QWidget {{ background: transparent; }}
    QFrame#workbox QComboBox {{ background: {t['panel2']}; }}
    QFrame#workbox QPushButton {{ background: {t['panel2']}; }}
    QLabel#boxTitle {{ color: {t['accent']}; font-weight: bold; background: transparent; }}
    QLabel#subhead {{ color: {t['group']}; font-weight: bold; background: transparent; padding-top: 4px; }}
    QPushButton#bigToggle {{ padding: 7px; font-weight: bold; }}
    QPushButton#bigToggle:checked {{ background: {t['accent']}; color: {t['accent_text']}; }}
    QCheckBox {{ background: transparent; spacing: 6px; }}
    QCheckBox::indicator {{ width: 13px; height: 13px; border: 1px solid {t['faint']}; border-radius: 3px; background: {t['canvas']}; }}
    QCheckBox::indicator:checked {{ background: {t['accent']}; border-color: {t['accent']}; }}
    QToolTip {{ background: {t['panel']}; color: {t['text']}; border: 1px solid {t['border']}; }}
    """


# ---------------------------------------------------------------- cell palettes
# Each entry: (name, RRGGBB). The Biomes colours come from the Godot client.

PALETTES = {
    "Nord": [
        ("Polar night", "2E3440"), ("Night 2", "3B4252"), ("Night 3", "434C5E"), ("Night 4", "4C566A"),
        ("Snow", "D8DEE9"), ("Snow 2", "E5E9F0"), ("Snow 3", "ECEFF4"),
        ("Frost green", "8FBCBB"), ("Frost", "88C0D0"), ("Frost blue", "81A1C1"), ("Deep frost", "5E81AC"),
        ("Aurora red", "BF616A"), ("Aurora orange", "D08770"), ("Aurora yellow", "EBCB8B"),
        ("Aurora green", "A3BE8C"), ("Aurora purple", "B48EAD"),
    ],
    "Dracula": [
        ("Background", "282A36"), ("Current line", "44475A"), ("Foreground", "F8F8F2"),
        ("Comment", "6272A4"), ("Cyan", "8BE9FD"), ("Green", "50FA7B"), ("Orange", "FFB86C"),
        ("Pink", "FF79C6"), ("Purple", "BD93F9"), ("Red", "FF5555"), ("Yellow", "F1FA8C"),
    ],
    "Biomes": [
        ("Land violet", "A58CF0"), ("Built cyan", "5FD0E6"), ("Strange pink", "F28AD6"),
        ("Lamp gold", "F6C86A"), ("Sand", "F2DCB8"), ("Tan", "C9A46B"), ("Umber", "5A4D39"),
        ("Bark", "3A3226"), ("Deep teal", "2F7F73"), ("Sky", "A8D8FF"), ("Lilac", "C9A6FF"),
        ("Iris", "8F7FE0"), ("Ice", "8FE9F7"), ("Leaf", "9BE37A"), ("Coral", "E38A7A"),
        ("Ink", "101014"),
    ],
}

# Colours handed out to new range layers, in order
LAYER_COLORS = ["A58CF0", "F28AD6", "5FD0E6", "F6C86A", "9BE37A", "E38A7A",
                "8FBCBB", "D08770", "BD93F9", "8BE9FD", "EBCB8B", "FF79C6"]


# ---------------------------------------------------------------- contrast


def luminance(hex6):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hex6[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(a, b):
    la, lb = luminance(a), luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def readable_on(text_hex, bg_hex, minimum=4.5):
    """Keep the hue but move lightness until the text reads on the background.
    Used only for display: dark text written for a white page, shown on a
    dark canvas."""
    if contrast(text_hex, bg_hex) >= 3.0:
        return text_hex
    r, g, b = (int(text_hex[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    dark_bg = luminance(bg_hex) < 0.2
    l = 1 - l
    best = text_hex
    for _ in range(20):
        rr, gg, bb = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
        best = "%02X%02X%02X" % (round(rr * 255), round(gg * 255), round(bb * 255))
        if contrast(best, bg_hex) >= minimum:
            break
        l += 0.04 if dark_bg else -0.04
    return best


def qcolor(hex6, alpha=255):
    c = QColor("#" + hex6.lstrip("#"))
    c.setAlpha(alpha)
    return c
