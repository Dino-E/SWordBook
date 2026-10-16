import os
import re
import bisect
import sys
import zipfile
import datetime
import getpass
import hashlib
import time
import json
import posixpath
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QToolBar, QToolButton, QFontComboBox,
    QComboBox, QMenu, QWidget, QWidgetAction, QGridLayout, QColorDialog, QLabel,
    QVBoxLayout, QStyleFactory, QSpinBox, QGraphicsView, QGraphicsScene,
    QFrame, QStatusBar, QSlider, QDialog, QFormLayout, QDoubleSpinBox,
    QDialogButtonBox, QGroupBox, QAbstractScrollArea, QMessageBox, QFileDialog,
    QRadioButton, QLineEdit, QCheckBox, QTabBar, QPushButton, QHBoxLayout,
    QDockWidget, QTreeWidget, QTreeWidgetItem, QScrollArea, QSizePolicy,
    QScrollBar, QSplitter,
)
from PySide6.QtGui import (
    QFont, QColor, QTextBlockFormat, QTextCharFormat, QTextCursor, QIcon, QPixmap,
    QPainter, QPen, QBrush, QAction, QKeySequence, QPalette, QIntValidator,
    QActionGroup, QPolygonF, QDoubleValidator, QTextFormat, QTransform,
    QCursor, QMouseEvent, QTextDocument, QTextDocumentFragment,
    QAbstractTextDocumentLayout, QDropEvent, QDragMoveEvent, QPainterPath,
    QImage, QTextImageFormat, QLinearGradient, QTextListFormat,
)
from PySide6.QtCore import (
    Qt, QSize, QSizeF, QRectF, QPointF, QPoint, Signal, QLocale, QTimer, QEvent,
    QObject, QBuffer, QByteArray, QIODevice, QUrl, QMimeData,
    QStandardPaths, QLockFile, QTranslator, QLibraryInfo, QSettings,
    QProcess, QDateTime,
)

# ---------------------------------------------------------------------------
# Interface language. The texts are written in English; _t() translates them
# with the _TRANSLATIONS table (at the end of the file). The language is chosen
# in Settings and saved in settings.ini ([Interface] language=…)
# ---------------------------------------------------------------------------
LANGUAGES = [
    ("en", "English"), ("es", "Español"), ("pt", "Português"),
    ("fr", "Français"), ("de", "Deutsch"), ("it", "Italiano"),
    ("zh", "简体中文"), ("ja", "日本語"),
]
LANGUAGE = "en"
# Qt locale (numbers, dates and Qt's own dialogs) and the language tag written
# to the .docx (Word spell checking, etc.)
QT_LOCALES = {"en": "en_US", "es": "es_MX", "pt": "pt_BR", "fr": "fr_FR",
              "de": "de_DE", "it": "it_IT", "zh": "zh_CN", "ja": "ja_JP"}
# Letters on the Bold, Italic and Underline buttons in each language
GLYPHS = {"en": "BIU", "es": "NKS", "pt": "NIS", "fr": "GIS", "de": "FKU",
          "it": "GCS", "zh": "BIU", "ja": "BIU"}


def set_language(code):
    global LANGUAGE
    LANGUAGE = code if code in dict(LANGUAGES) else "en"


def _t(source, **data):
    """Translate an interface text into the chosen language. With `data`, fill
    in its fields ({name}, {n}…). A "context|text" key is shown in English
    as "text"."""
    t = _TRANSLATIONS.get(LANGUAGE, {}).get(source)
    if t is None:
        t = source.split("|", 1)[-1]
    return t.format(**data) if data else t


def glyph(kind):
    letters = GLYPHS.get(LANGUAGE, GLYPHS["en"])
    return letters["bold italic underline".split().index(kind)]


def language_tag():
    return QT_LOCALES.get(LANGUAGE, "en_US").replace("_", "-")


def local_date(seconds):
    """Date and time in the language's format."""
    loc = QLocale()
    format_spec = loc.dateTimeFormat(QLocale.ShortFormat)
    if "yyyy" not in format_spec:                     # year always with 4 digits
        format_spec = format_spec.replace("yy", "yyyy")
    return loc.toString(QDateTime.fromSecsSinceEpoch(int(seconds)), format_spec)


# ---------------------------------------------------------------------------
# "Chapter body" style for e-books
# ---------------------------------------------------------------------------
CHAPTER_BODY = {
    "font_family": "Georgia",           # Serif, readable on screen
    "font_size": 12,                    # pt (≈1em in readers)
    "font_color": "#1A1A1A",            # Almost black, less eye strain
    "alignment": Qt.AlignJustify,       # Justified, typical of books
    "line_spacing": 150,                # % (line spacing 1.5)
    "space_before": 0,                  # px
    "space_after": 6,                   # px
    "indent_left": 0,                   # px
    "indent_right": 0,                  # px
    "first_line_indent": 0,             # px (first-line indent)
    "bold": False,
    "italic": False,
    "heading_level": 0,                 # 0 = normal paragraph (<p>)
}

# ---------------------------------------------------------------------------
# "Chapter title" style
# ---------------------------------------------------------------------------
CHAPTER_TITLE = {
    "font_family": "Georgia",           # Same family as the body
    "font_size": 24,                    # pt (≈2em, clear hierarchy)
    "font_color": "#1A1A1A",
    "alignment": Qt.AlignHCenter,       # Centered, typical of chapters
    "line_spacing": 120,                # % (more compact at large sizes)
    "space_before": 48,                 # px (space before the chapter)
    "space_after": 12,                  # px
    "indent_left": 0,
    "indent_right": 0,
    "first_line_indent": 0,
    "bold": True,
    "italic": False,
    "heading_level": 1,                 # <h1>: readers use it for the table of contents
}

# ---------------------------------------------------------------------------
# "Chapter subtitle" style
# ---------------------------------------------------------------------------
CHAPTER_SUBTITLE = {
    "font_family": "Georgia",
    "font_size": 16,                    # pt (≈1.33em, between title and body)
    "font_color": "#4A4A4A",            # Dark gray, secondary to the title
    "alignment": Qt.AlignHCenter,
    "line_spacing": 130,                # %
    "space_before": 0,                  # px (close to the title)
    "space_after": 36,                  # px (separates it from the chapter text)
    "indent_left": 0,
    "indent_right": 0,
    "first_line_indent": 0,
    "bold": False,
    "italic": True,                     # Italic, the convention for subtitles
    "heading_level": 2,                 # <h2>
}

# "Chapter thoughts" style: same as the body, in italics
CHAPTER_THOUGHTS = {**CHAPTER_BODY, "italic": True}

# ---------------------------------------------------------------------------
# "Chapter section" style (section heading inside a chapter)
# ---------------------------------------------------------------------------
CHAPTER_SECTION = {
    "font_family": "Georgia",
    "font_size": 14,                    # pt (≈1.17em, smaller than the subtitle)
    "font_color": "#1A1A1A",
    "alignment": Qt.AlignLeft,          # Left: it divides, it doesn't open a chapter
    "line_spacing": 130,                # %
    "space_before": 24,                 # px (separates it from the previous section)
    "space_after": 8,                   # px (close to its own text)
    "indent_left": 0,
    "indent_right": 0,
    "first_line_indent": 0,
    "bold": True,
    "italic": False,
    "heading_level": 3,                 # <h3>, below title and subtitle
}

# ---------------------------------------------------------------------------
# "Chapter quote" style (block quote: letters, epigraphs, quoted texts)
# ---------------------------------------------------------------------------
CHAPTER_QUOTE = {
    "font_family": "Georgia",
    "font_size": 11,                    # pt (≈0.92em, a bit smaller than the body)
    "font_color": "#333333",            # A softer tone than the body
    "alignment": Qt.AlignJustify,
    "line_spacing": 140,                # %
    "space_before": 12,                 # px
    "space_after": 12,                  # px
    "indent_left": 30,                  # px (≈2em, the quote stands apart from the text)
    "indent_right": 30,                 # px
    "first_line_indent": 0,
    "bold": False,
    "italic": False,
    "heading_level": 0,                 # <p>
}

# ---------------------------------------------------------------------------
# "Scene break" style
# ---------------------------------------------------------------------------
SCENE_BREAK = {
    "font_family": "Georgia",
    "font_size": 12,
    "font_color": "#1A1A1A",
    "alignment": Qt.AlignHCenter,       # Centered
    "line_spacing": 150,
    "space_before": 18,                 # px (space that marks the scene change)
    "space_after": 18,                  # px
    "indent_left": 0,
    "indent_right": 0,
    "first_line_indent": 0,
    "bold": False,
    "italic": False,
    "heading_level": 0,
}

# Visible name of each style (ribbon, tooltips, navigator): the one for the
# language, unless settings.ini changes it ([Styles.Name] → name=…)
DEFAULT_STYLE_NAMES = {
    "body": "Body", "title": "Title", "subtitle": "Subtitle",
    "section": "Section", "thoughts": "Thoughts", "dialogue": "Dialogue",
    "quote": "Quote", "scene_break": "Scene break",
}
STYLE_NAMES = {}                    # only the ones changed in settings.ini
STYLE_NAME_MAX = 40                  # characters


def style_name(key):
    return STYLE_NAMES.get(key) or _t(DEFAULT_STYLE_NAMES[key])


# Version and license (shown, discreetly, in the Settings window)
APP_NAME = "SWordBook"
APP_VERSION = "0.1.0-alpha"
APP_LICENSE = "License: MIT | © DinoSoft.1998"

DASH = "\u2014"                         # "—" dialogue dash
SEPARATOR = "* * *"                     # Scene separator (shows in every reader)

# Style ribbon swatches (px)
SWATCH_WIDTH, SWATCH_HEIGHT = 112, 32
# Order of the swatches in the ribbon (2 rows of 4)
STYLE_ORDER = ["body", "title", "subtitle", "section",
                 "thoughts", "dialogue", "quote", "scene_break"]

SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]

# Font color palette (word-processor style, 10 columns)
FONT_PALETTE = [
    "#000000", "#434343", "#666666", "#999999", "#B7B7B7",
    "#CCCCCC", "#D9D9D9", "#EFEFEF", "#F3F3F3", "#FFFFFF",
    "#980000", "#FF0000", "#FF9900", "#FFFF00", "#00FF00",
    "#00FFFF", "#4A86E8", "#0000FF", "#9900FF", "#FF00FF",
    "#A61C00", "#CC0000", "#E69138", "#F1C232", "#6AA84F",
    "#45818E", "#3C78D8", "#3D85C6", "#674EA7", "#A64D79",
]

# Highlight palette: soft tones that don't hide the text
HIGHLIGHT_PALETTE = [
    "#FFF59D", "#C5E1A5", "#80DEEA", "#90CAF9", "#CE93D8",
    "#F48FB1", "#FFCC80", "#BCAAA4", "#E0E0E0", "#FFFF00",
]

ICON_COLOR = "#3A3A3A"

# ---------------------------------------------------------------------------
# Page (measurements in cm, as in a word processor)
# ---------------------------------------------------------------------------
PAGE_SIZES = [            # name, width, height (portrait)
    ("Letter", 21.59, 27.94),
    ("Oficio", 21.59, 34.01),
    ("Legal", 21.59, 35.56),
    ("A4", 21.0, 29.7),
    ("A5", 14.8, 21.0),
    ("Book 6 × 9 in", 15.24, 22.86),
    ("Book 5.5 × 8.5 in", 13.97, 21.59),
    ("Book 5 × 8 in", 12.7, 20.32),
]
PAGE = {                     # Initial setup
    "width": 21.59, "height": 27.94,          # Letter, portrait
    "top": 2.5, "bottom": 2.5, "left": 3.0, "right": 3.0,   # margins
}
# Block state (QTextBlock.userState) of a collapsed heading. It only affects
# the view: it is not saved to the file and is not part of undo
STATE_COLLAPSED = 1
# Internal property that marks a page break inserted with Ctrl+Enter
PROP_MANUAL_BREAK = QTextFormat.UserProperty + 1
# Root frame property with the page setup ("width;height;…"). This way, if
# Ctrl+Z undoes a margin change, the page size that went with it is restored
# too, and the view adapts to both
PROP_PAGE = QTextFormat.UserProperty + 2
PAGE_KEYS = ("width", "height", "top", "bottom", "left", "right")
# Character property with the comments that span the text ("3 7"). It lives in
# the text format: it is undone, redone and saved with it
PROP_COMMENT = QTextFormat.UserProperty + 3
# Comment highlight colors (each comment can have its own). New ones take the
# next color, to tell neighboring comments apart
# ---------------------------------------------------------------------------
# Images ("Top and bottom"): each image takes its own paragraph, which the text
# respects as a gap; it is drawn centered in the column
# ---------------------------------------------------------------------------
# The text reserves the gap with a transparent Qt image (the size at which the
# image is shown); the real image is drawn on top by the editor
URL_GAP = QUrl("swordbook://gap")
PROP_IMAGE_BLOCK = QTextFormat.UserProperty + 4   # paragraph of an image
PROP_IMG_NAME = QTextFormat.UserProperty + 5
PROP_IMG_WIDTH = QTextFormat.UserProperty + 6       # editor px (requested
PROP_IMG_HEIGHT = QTextFormat.UserProperty + 7        # size; reduced if it doesn't fit)
# When opening a .docx: list the paragraph belongs to
# ("bullet|decimal;numId;level")
PROP_LIST = QTextFormat.UserProperty + 8
# Lists: bullets and numbers (indent per level, in px)
LIST_INDENT = 30
OBJECT_CHAR = "\ufffc"
IMG_MARGIN = 8               # px of space above and below the image
IMG_MIN = 16              # px, minimum size when resizing
HANDLE = 8                      # on-screen px of each drag handle
IMAGE_SELECTION_COLOR = "#4A86E8"
MIME_BOOK = "application/x-swordbook"   # copy and paste with images
DOCX_IMAGE_TYPES = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                     "gif": "image/gif", "bmp": "image/bmp", "tif": "image/tiff",
                     "tiff": "image/tiff", "emf": "image/x-emf", "wmf": "image/x-wmf"}


def is_image(block):
    return block is not None and block.isValid() and \
        block.blockFormat().boolProperty(PROP_IMAGE_BLOCK)


COMMENT_COLORS = [
    "#E6DCF7", "#FFF59D", "#C5E1A5", "#B3E5FC", "#F8BBD0",
    "#FFCC80", "#B2DFDB", "#D7CCC8", "#E0E0E0", "#FFAB91",
]


def active_color(color):
    """More intense tone for the comment where the cursor is."""
    return QColor(color).darker(118)

PAGE_GAP = 20        # px between sheets (and around them)
# View modes
VIEW_PAGES, VIEW_NO_PAGES, VIEW_MIXED = "pages", "no_pages", "mixed"
# Names used by earlier versions (in settings.ini and in .docx files)
OLD_VIEWS = {"paginas": VIEW_PAGES, "sin_paginas": VIEW_NO_PAGES,
                   "mixto": VIEW_MIXED}
# Zoom: only fixed levels between 50 % and 200 %. The text is laid out at 100 %
# and the view scales it like an image: zoomed far out, letters crowd and
# overlap; at odd scales (113 %, 187 %…) the spacing between letters and words
# looks uneven. With these levels the text looks even
ZOOM_LEVELS = [50, 60, 67, 75, 80, 90, 100, 110, 125, 150, 175, 200]
ZOOM_MIN, ZOOM_MAX = ZOOM_LEVELS[0] / 100, ZOOM_LEVELS[-1] / 100
ZOOM_MENU = [50, 75, 100, 125, 150, 200]           # zoom button menu


def zoom_level(z, toward="nearest"):
    """The allowed level closest to z (or the largest one not above it, with
    toward="down": so what you want to see whole does fit)."""
    pct = z * 100
    if toward == "down":
        candidates = [n for n in ZOOM_LEVELS if n <= pct + 1e-6]
        return (candidates[-1] if candidates else ZOOM_LEVELS[0]) / 100
    return min(ZOOM_LEVELS, key=lambda n: (abs(n - pct), n)) / 100

# Auto-scroll when selecting by dragging toward the edge: the speed grows with
# the distance between the cursor and the edge of the view
AUTO_ZONE = 20             # px before the edge where scrolling starts
AUTO_INTERVAL = 16        # ms between steps (~60 per second)
AUTO_MAX_SPEED = 60          # px per step at most

# Indent: each step moves the paragraph; the maximum avoids unreadable columns
INDENT_STEP = 15          # px per increase/decrease click
INDENT_MAX = 240          # px, limit for left and right indent
SPACE_STEP = 6           # px per increase/decrease click for before/after
FIRST_LINE_MAX = 120    # px, limit for first-line indent
SPACE_MAX = 200          # px, limit for space before/after
LETTERS_MIN, LETTERS_MAX = -5.0, 20.0   # pt, letter spacing
LETTERS_STEP = 0.5                     # pt per click
LINE_SPACINGS = ["1.0", "1.15", "1.5", "2.0", "2.5", "3.0"]

# ---------------------------------------------------------------------------
# Interface style sheet (light, neutral theme)
# ---------------------------------------------------------------------------
STYLESHEET = """
QMainWindow { background: #F3F3F3; }

QToolBar {
    background: #F7F7F7;
    border: none;
    border-bottom: 1px solid #DADADA;
    padding: 4px 8px;
    spacing: 2px;
}
QToolBar::separator {
    background: #D6D6D6;
    width: 1px;
    margin: 5px 6px;
}

QToolButton {
    background: transparent;
    border: 1px solid transparent;
    border-radius: 4px;
    padding: 3px;
}
QToolButton:hover   { background: #E9E9E9; border-color: #D2D2D2; }
QToolButton:pressed { background: #DDDDDD; }
QToolButton:checked { background: #DCE8F7; border-color: #9DBCE3; }
QToolButton[popupMode="1"] { padding-right: 15px; }
QToolButton::menu-button {
    border: none;
    border-left: 1px solid transparent;
    width: 13px;
}
QToolButton::menu-button:hover { border-left-color: #D2D2D2; }

QToolBar QLabel#groupTitle {
    color: #6A6A6A;
    font-size: 11px;
    font-weight: bold;
    padding: 0 6px 0 2px;
}
QToolBar QLabel#fieldLabel {
    color: #3A3A3A;
    padding: 0 3px 0 8px;
}

QComboBox QAbstractItemView {
    background: #FFFFFF;
    color: #1A1A1A;
    border: 1px solid #C8C8C8;
    selection-background-color: #DCE8F7;
    selection-color: #1A1A1A;
    outline: none;
}

QMenu {
    background: #FFFFFF;
    border: 1px solid #C8C8C8;
    border-radius: 6px;
    padding: 6px;
}
QMenu::item {
    color: #1A1A1A;
    padding: 5px 16px;
    border-radius: 4px;
}
QMenu::item:selected { background: #DCE8F7; }
QMenu::separator { height: 1px; background: #E2E2E2; margin: 5px 4px; }

QToolTip {
    background: #FFFFFF;
    color: #1A1A1A;
    border: 1px solid #C8C8C8;
    padding: 4px 6px;
}

QGraphicsView#canvas { background: #E4E4E4; border: none; }

QStatusBar {
    background: #F7F7F7;
    border-top: 1px solid #DADADA;
    color: #3A3A3A;
}
QStatusBar::item { border: none; }
QStatusBar QToolButton { padding: 1px 6px; }
QStatusBar QLabel { color: #3A3A3A; padding: 0 8px; }

QTextEdit {
    background: #FFFFFF;
    border: none;
    selection-background-color: #B3D4FC;
    selection-color: #1A1A1A;
}
"""


def install_qt_translations(app):
    """Qt's own texts in the chosen language: right-click menu (Undo, Copy,
    Paste…), color picker, file dialogs and standard buttons. Also the date
    and number format."""
    loc = QLocale(QT_LOCALES.get(LANGUAGE, "en_US"))
    QLocale.setDefault(loc)
    app._translators = []                 # kept so they aren't freed
    if LANGUAGE == "en":
        return
    folder = QLibraryInfo.path(QLibraryInfo.TranslationsPath)
    for base in ("qtbase", "qt"):
        t = QTranslator(app)
        if t.load(loc, base, "_", folder):
            app.installTranslator(t)
            app._translators.append(t)


def restart_program():
    """Start the program again (to apply another language)."""
    if getattr(sys, "frozen", False):          # packaged executable
        QProcess.startDetached(sys.executable, sys.argv[1:])
    else:
        QProcess.startDetached(sys.executable,
                               [os.path.abspath(sys.argv[0])] + sys.argv[1:])


def light_palette():
    """Fixed light palette so the interface looks the same in dark mode."""
    p = QPalette()
    p.setColor(QPalette.Window, QColor("#F3F3F3"))
    p.setColor(QPalette.WindowText, QColor("#1A1A1A"))
    p.setColor(QPalette.Base, QColor("#FFFFFF"))
    p.setColor(QPalette.AlternateBase, QColor("#F7F7F7"))
    p.setColor(QPalette.Text, QColor("#1A1A1A"))
    p.setColor(QPalette.Button, QColor("#F7F7F7"))
    p.setColor(QPalette.ButtonText, QColor("#1A1A1A"))
    p.setColor(QPalette.Highlight, QColor("#4A86E8"))
    p.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    p.setColor(QPalette.ToolTipBase, QColor("#FFFFFF"))
    p.setColor(QPalette.ToolTipText, QColor("#1A1A1A"))
    p.setColor(QPalette.PlaceholderText, QColor("#8A8A8A"))
    return p


# ---------------------------------------------------------------------------
# Drawn icons (sharp on HiDPI screens, no external files)
# ---------------------------------------------------------------------------
def make_icon(text, *, bold=False, italic=False, underline=False,
                strike=False, small=None, small_pos="up",
                bar=None, background=None):
    scale_factor = 3
    size = 24
    pm = QPixmap(size * scale_factor, size * scale_factor)
    pm.setDevicePixelRatio(scale_factor)
    pm.fill(Qt.transparent)

    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)

    text_height = 19 if bar is not None else size

    if background is not None:  # highlight block behind the text
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(background))
        p.drawRoundedRect(QRectF(3, 4, 18, 13), 2, 2)

    f = QFont("Arial")
    f.setPixelSize(15 if small is None else 14)
    f.setBold(bold)
    f.setItalic(italic)
    f.setUnderline(underline)
    f.setStrikeOut(strike)
    p.setFont(f)
    p.setPen(QColor(ICON_COLOR))

    if small is None:
        p.drawText(QRectF(0, 0, size, text_height), Qt.AlignCenter, text)
    else:
        p.drawText(QRectF(1, 0, 15, size), Qt.AlignCenter, text)
        fp = QFont("Arial")
        fp.setPixelSize(10)
        fp.setBold(True)
        p.setFont(fp)
        y = 2 if small_pos == "up" else 12
        p.drawText(QRectF(14, y, 10, 10), Qt.AlignCenter, small)

    if bar is not None:  # strip with the current color
        c = QColor(bar)
        p.setPen(QPen(QColor("#9A9A9A"), 0.6) if not c.isValid() or c.lightness() > 190
                 else Qt.NoPen)
        p.setBrush(QBrush(c) if c.isValid() else Qt.NoBrush)
        p.drawRoundedRect(QRectF(3, 19.5, 18, 3.5), 1, 1)

    p.end()
    return QIcon(pm)


def clear_format_icon():
    """An A with an eraser on top, in the bottom-right corner."""
    scale_factor = 3
    size = 24
    pm = QPixmap(size * scale_factor, size * scale_factor)
    pm.setDevicePixelRatio(scale_factor)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)
    f = QFont("Arial")
    f.setPixelSize(16)
    f.setBold(True)
    p.setFont(f)
    p.setPen(QColor(ICON_COLOR))
    p.drawText(QRectF(0, 0, 17, 20), Qt.AlignCenter, "A")
    # Tilted eraser: light body and pink tip, with an outline
    p.translate(16.5, 16.5)
    p.rotate(-45)
    outline = QPen(QColor(ICON_COLOR), 1.2)
    outline.setJoinStyle(Qt.RoundJoin)
    p.setPen(outline)
    p.setBrush(QColor("#FFFFFF"))
    p.drawRoundedRect(QRectF(-6.5, -3, 13, 6), 1.2, 1.2)
    p.setBrush(QColor("#F48FB1"))
    p.drawRoundedRect(QRectF(-6.5, -3, 5, 6), 1.2, 1.2)
    p.end()
    return QIcon(pm)


def paragraph_icon(kind):
    """Line icons for alignment, indent and lists."""
    scale_factor = 3
    size = 24
    pm = QPixmap(size * scale_factor, size * scale_factor)
    pm.setDevicePixelRatio(scale_factor)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor(ICON_COLOR), 1.8)
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)

    ys = [5, 9.5, 14, 18.5]
    lengths = [16, 11, 16, 11]
    x0, x1 = 4, 20
    if kind in ("left", "center", "right", "justify"):
        for y, ln in zip(ys, lengths):
            if kind == "justify":
                a, b = x0, x1
            elif kind == "left":
                a, b = x0, x0 + ln
            elif kind == "right":
                a, b = x1 - ln, x1
            else:
                a = (x0 + x1 - ln) / 2
                b = a + ln
            p.drawLine(QPointF(a, y), QPointF(b, y))
    elif kind in ("bullets", "numbers"):
        # Three lines with their bullet (dot) or number on the left
        for k, y in enumerate((5.5, 12, 18.5)):
            p.drawLine(QPointF(10, y), QPointF(x1, y))
            if kind == "bullets":
                p.save()
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(ICON_COLOR))
                p.drawEllipse(QPointF(5, y), 1.9, 1.9)
                p.restore()
            else:
                p.save()
                f = QFont("Arial")
                f.setPixelSize(7)
                f.setBold(True)
                p.setFont(f)
                p.setPen(QColor(ICON_COLOR))
                p.drawText(QRectF(1, y - 4.5, 7.5, 9), Qt.AlignCenter, str(k + 1))
                p.restore()
    elif kind.startswith(("before", "after")):
        # Paragraph lines, the edge of the neighboring paragraph (dotted) and
        # an arrow: down increases the space, up decreases it
        before = kind.startswith("before")
        lines = (13.5, 17, 20.5) if before else (3.5, 7, 10.5)
        for y, ln in zip(lines, (16, 16, 10)):
            p.drawLine(QPointF(x0, y), QPointF(x0 + ln, y))
        border = 3 if before else 21
        guide = QPen(QColor(ICON_COLOR), 1.1, Qt.DotLine)
        p.setPen(guide)
        p.drawLine(QPointF(x0, border), QPointF(x1, border))
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(ICON_COLOR))
        cy = 7.75 if before else 15.75        # middle of the space
        if kind.endswith("increase"):        # ▼
            tri = QPolygonF([QPointF(15.5, cy - 2.5), QPointF(21, cy - 2.5),
                             QPointF(18.25, cy + 2.5)])
        else:                                # ▲
            tri = QPolygonF([QPointF(15.5, cy + 2.5), QPointF(21, cy + 2.5),
                             QPointF(18.25, cy - 2.5)])
        p.drawPolygon(tri)
    else:
        # First and last lines full; the middle ones shifted
        p.drawLine(QPointF(x0, ys[0]), QPointF(x1, ys[0]))
        p.drawLine(QPointF(x0, ys[3]), QPointF(x1, ys[3]))
        p.drawLine(QPointF(11, ys[1]), QPointF(x1, ys[1]))
        p.drawLine(QPointF(11, ys[2]), QPointF(x1, ys[2]))
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(ICON_COLOR))
        if kind == "increase":   # arrow to the right
            tri = QPolygonF([QPointF(4, 8.5), QPointF(8.5, 11.75), QPointF(4, 15)])
        else:                    # arrow to the left
            tri = QPolygonF([QPointF(8.5, 8.5), QPointF(4, 11.75), QPointF(8.5, 15)])
        p.drawPolygon(tri)
    p.end()
    return QIcon(pm)


# ---------------------------------------------------------------------------
# Color button with a palette menu
# ---------------------------------------------------------------------------
class ColorButton(QToolButton):
    colorChosen = Signal(QColor)  # invalid QColor = remove color

    def __init__(self, glyph, initial_color, palette, none_text,
                 none_color, tooltip, highlight=False, parent=None):
        super().__init__(parent)
        self.glyph = glyph
        self.highlight = highlight
        self.none_color = QColor(none_color) if none_color else QColor()
        self.current = QColor(initial_color)

        self.setToolTip(tooltip)
        self.setPopupMode(QToolButton.MenuButtonPopup)
        self.clicked.connect(lambda: self.colorChosen.emit(self.current))

        menu = QMenu(self)
        none_action = menu.addAction(none_text)
        none_action.triggered.connect(lambda: self._choose(self.none_color))
        menu.addSeparator()

        box_widget = QWidget()
        lay = QVBoxLayout(box_widget)
        lay.setContentsMargins(8, 4, 8, 4)
        lay.setSpacing(6)
        title = QLabel(_t("Colors"))
        title.setStyleSheet("color:#6A6A6A; font-size:11px;")
        lay.addWidget(title)

        grid = QGridLayout()
        grid.setSpacing(4)
        for i, hex_color in enumerate(palette):
            b = QToolButton()
            b.setFixedSize(20, 20)
            b.setToolTip(hex_color.upper())
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet(
                f"QToolButton {{ background:{hex_color}; border:1px solid #C4C4C4;"
                f" border-radius:3px; padding:0; }}"
                f"QToolButton:hover {{ border:2px solid #3A3A3A; }}"
            )
            b.clicked.connect(lambda _=False, c=hex_color: self._choose(QColor(c)))
            grid.addWidget(b, i // 10 if len(palette) > 10 else i // 5,
                           i % 10 if len(palette) > 10 else i % 5)
        lay.addLayout(grid)

        wa = QWidgetAction(menu)
        wa.setDefaultWidget(box_widget)
        menu.addAction(wa)
        menu.addSeparator()
        more_action = menu.addAction(_t("More colors…"))
        more_action.triggered.connect(self._dialog)

        self.palette_menu = menu
        self.setMenu(menu)
        self._update_icon()

    def _update_icon(self):
        if self.highlight:
            self.setIcon(make_icon(self.glyph, bar=self.current))
        else:
            self.setIcon(make_icon(self.glyph, bold=True, bar=self.current))

    def set_color_value(self, color):
        """Change the color shown without applying it to the text."""
        self.current = QColor(color)
        self._update_icon()

    def _choose(self, color):
        self.palette_menu.hide()
        self.current = QColor(color)
        self._update_icon()
        self.colorChosen.emit(self.current)

    def _dialog(self):
        initial = self.current if self.current.isValid() else QColor("#FFFFFF")
        c = QColorDialog.getColor(initial, self, _t("Choose color"))
        if c.isValid():
            self._choose(c)


# ---------------------------------------------------------------------------
# Comments
# ---------------------------------------------------------------------------
def comment_ids(cf):
    """Comments that span a text fragment."""
    if not cf.hasProperty(PROP_COMMENT):
        return set()
    ids = set()
    for x in cf.stringProperty(PROP_COMMENT).split():
        try:
            ids.add(int(x))
        except ValueError:
            pass
    return ids


def set_comment_ids(cf, ids):
    if ids:
        cf.setProperty(PROP_COMMENT, " ".join(str(i) for i in sorted(ids)))
    else:
        cf.clearProperty(PROP_COMMENT)


def now_iso():
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def readable_date(iso):
    try:
        d = datetime.datetime.fromisoformat(iso.replace("Z", "+00:00"))
        if d.tzinfo is None:
            d = d.replace(tzinfo=datetime.timezone.utc)
        return local_date(d.timestamp())
    except (ValueError, AttributeError):
        return iso or ""


class CommentManager(QObject):
    """The document's comments.
    #
    The commented text carries the PROP_COMMENT property in its format with
    the comment number; everything else (author, date and comment text) is
    kept here. This way undo and redo just work: if the commented text is
    deleted, the comment disappears from the panel, and if the deletion is
    undone, it comes back. Only the ones that have text are saved to the
    file."""
    listChanged = Signal()
    activeChanged = Signal(int)
    changed = Signal()           # text or color of a comment

    def __init__(self, editor):
        super().__init__(editor)
        self.editor = editor
        self.data = {}          # id → {"author", "date", "text", "color"}
        self.ranges = {}         # id → QTextCursor with the commented text
        self.order = []          # ids in the order they appear
        self.active_id = -1
        self._next_id = 1
        try:                      # author saved in the .docx
            self.author = getpass.getuser()
        except Exception:
            self.author = _t("Author")
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(200)
        self._timer.timeout.connect(self.scan)
        editor.document().contentsChange.connect(self._schedule)
        editor.cursorPositionChanged.connect(self._cursor_moved)

    # -- reading the document ---------------------------------------------
    def _schedule(self, *_):
        if self.data:            # without comments there is nothing to check
            self._timer.start()

    def _fragments(self, start=0, end=None):
        """(start, end, format) of each text fragment in the range."""
        doc = self.editor.document()
        if end is None:
            end = doc.characterCount()
        b = doc.findBlock(start)
        while b.isValid() and b.position() < end:
            it = b.begin()
            while not it.atEnd():
                f = it.fragment()
                if f.isValid():
                    a = max(start, f.position())
                    z = min(end, f.position() + f.length())
                    if a < z:
                        yield a, z, f.charFormat()
                it += 1
            b = b.next()

    def scan(self):
        """Find where each comment is in the text."""
        self._timer.stop()
        limits = {}
        if self.data:
            for a, z, cf in self._fragments():
                for i in comment_ids(cf):
                    if i in self.data:
                        r = limits.get(i)
                        limits[i] = (a, z) if r is None else (min(r[0], a), max(r[1], z))
        doc = self.editor.document()
        self.ranges = {}
        for i, (a, z) in limits.items():
            c = QTextCursor(doc)
            c.setPosition(a)
            c.setPosition(z, QTextCursor.KeepAnchor)
            self.ranges[i] = c
        self.order = sorted(limits, key=lambda i: (limits[i][0], i))
        self.active_id = -2              # forces a recalculation and notification
        self._cursor_moved()
        self.listChanged.emit()

    def excerpt(self, i, length=90):
        r = self.ranges.get(i)
        if r is None:
            return ""
        t = " ".join(r.selectedText().replace("\u2029", " ").replace("\u2028", " ")
                     .replace(OBJECT_CHAR, " ").split())
        return t if len(t) <= length else t[:length - 1].rstrip() + "…"

    # -- cursor -------------------------------------------------------------
    def _cursor_moved(self):
        c = self.editor.textCursor()
        pos = c.position()
        inside = [i for i, r in self.ranges.items()
                  if r.hasSelection() and r.selectionStart() <= pos <= r.selectionEnd()]
        # If comments are nested, the shortest one
        new = min(inside, key=lambda i: (self.ranges[i].selectionEnd()
                                           - self.ranges[i].selectionStart(), i)) \
            if inside else -1
        if new != self.active_id:
            self.active_id = new
            self.editor.refresh_selections()
            self.activeChanged.emit(new)
        self._dont_extend(c)

    def _dont_extend(self, c):
        """What is typed right at the end of a commented text does not go into
        the comment (inside it, it does)."""
        if c.hasSelection() or not self.data:
            return
        cf = self.editor.currentCharFormat()
        before = comment_ids(cf)
        if not before:
            return
        after = set()
        b = c.block()
        if c.position() < b.position() + b.length() - 1:
            d = QTextCursor(c)
            d.movePosition(QTextCursor.NextCharacter)
            after = comment_ids(d.charFormat())
        remaining = before & after
        if remaining != before:
            set_comment_ids(cf, remaining)
            self.editor.setCurrentCharFormat(cf)   # doesn't touch the document

    def selections(self):
        """Highlight of the commented text."""
        res = []
        for i in sorted(self.order, key=lambda i: i == self.active_id):
            r = self.ranges.get(i)
            if r is None or not r.hasSelection():
                continue
            sel = QTextEdit.ExtraSelection()
            sel.cursor = QTextCursor(r)
            # Always the same tone, whether or not the cursor is inside
            sel.format.setBackground(QColor(self.color(i)))
            res.append(sel)
        return res

    # -- changes --------------------------------------------------------------
    def _change_marks(self, start, end, func):
        """Change the comment marks in the range (a single undo step)."""
        changes = list(self._fragments(start, end))
        c = QTextCursor(self.editor.document())
        c.beginEditBlock()
        for a, z, cf in changes:
            before = comment_ids(cf)
            ids = func(before)
            if ids == before:
                continue
            set_comment_ids(cf, ids)
            c.setPosition(a)
            c.setPosition(z, QTextCursor.KeepAnchor)
            c.setCharFormat(cf)       # the same format, only the mark changes
        c.endEditBlock()

    def color(self, i):
        c = self.data.get(i, {}).get("color")
        return c if c and QColor(c).isValid() else \
            COMMENT_COLORS[(i - 1) % len(COMMENT_COLORS)]

    def set_color_value(self, i, color):
        if i in self.data and QColor(color).isValid():
            self.data[i]["color"] = QColor(color).name().upper()
            self.changed.emit()
            self.editor.document().setModified(True)
            self.editor.refresh_selections()
            self.listChanged.emit()

    def add(self, author=None):
        """Comment on the selection (or the word under the cursor). Returns the
        id, or None if there is no text to comment on."""
        c = self.editor.textCursor()
        if not c.hasSelection():
            c.select(QTextCursor.WordUnderCursor)
        if not c.hasSelection() or not c.selectedText().strip():
            return None
        i = self._next_id
        self._next_id += 1
        self.data[i] = {"author": author or self.author, "date": now_iso(),
                         "text": "",
                         "color": COMMENT_COLORS[(i - 1) % len(COMMENT_COLORS)]}
        self._change_marks(c.selectionStart(), c.selectionEnd(), lambda s: s | {i})
        self.scan()
        return i

    def delete(self, i):
        """Remove the mark from the text. The data is kept in case it is undone
        (it is not saved while no text is marked)."""
        r = self.ranges.get(i)
        if r is not None:
            self._change_marks(r.selectionStart(), r.selectionEnd(),
                                 lambda s: s - {i})
        self.scan()

    def set_text(self, i, text):
        if i in self.data and self.data[i]["text"] != text:
            self.data[i]["text"] = text
            self.changed.emit()
            self.editor.document().setModified(True)

    def go_to(self, i):
        r = self.ranges.get(i)
        if r is None:
            return
        c = QTextCursor(r)
        c.setPosition(r.selectionStart())
        self.editor.setTextCursor(c)       # expands it if it was collapsed
        view = self.editor.view
        if view is None:
            return
        # Scroll here rather than when following the cursor: that is skipped
        # while the mouse button is held (the click on the comment)
        QApplication.processEvents()
        rect = QRectF(self.editor.cursorRect())
        point = view.proxy.mapToScene(rect.center())
        # The commented text ends up a third of the way down the view
        height = view.viewport().height() / view.zoom
        view.centerOn(point.x(), point.y() + height / 6)

    def marks(self, start, end):
        """Marks in the range, to restore them after replacing the format."""
        return [(a, z, cf.stringProperty(PROP_COMMENT))
                for a, z, cf in self._fragments(start, end)
                if cf.hasProperty(PROP_COMMENT)]

    def restore(self, marks):
        c = QTextCursor(self.editor.document())
        for a, z, value in marks:
            cf = QTextCharFormat()
            cf.setProperty(PROP_COMMENT, value)
            c.setPosition(a)
            c.setPosition(z, QTextCursor.KeepAnchor)
            c.mergeCharFormat(cf)

    def load(self, data):
        self.data = {i: dict(d) for i, d in data.items()}
        self._next_id = max(self.data, default=0) + 1
        self.scan()

    def clear_all(self):
        self.load({})


# ---------------------------------------------------------------------------
# Editor
# ---------------------------------------------------------------------------
class Editor(QTextEdit):
    # Emitted when a tool changes the paragraph format without moving the
    # cursor, so that all the bars update.
    paragraphFormatChanged = Signal()
    # Emitted when an image is selected or released (for Cut and Copy)
    imageSelectionChanged = Signal()

    @property
    def _img_sel(self):
        return getattr(self, "_img_sel_value", None)

    @_img_sel.setter
    def _img_sel(self, value):
        before = getattr(self, "_img_sel_value", None)
        self._img_sel_value = value
        if (before is None) != (value is None) or value is not None:
            self.imageSelectionChanged.emit()

    def __init__(self):
        super().__init__()
        self._external_selections = []    # Find highlight
        self.view = None          # PageView that contains it
        self.page_size = None     # QSizeF used to paginate the document
        self._initial_style()
        # Collapsible headings
        self.collapse_enabled = True   # can be turned off in Settings
        self._collapsed = 0       # collapsed headings (0 = nothing to check)
        self._navigating = False    # moving the cursor with the keyboard
        self._folding_pending = False
        self.viewport().setMouseTracking(True)   # hand cursor over the triangle
        self.document().contentsChanged.connect(self._schedule_folding)
        self.cursorPositionChanged.connect(self._cursor_in_hidden)
        # Mixed view: cuts (in text coordinates) and height of each gap
        self._cuts = []
        self._gap = 0.0
        for signal in (self.cursorPositionChanged, self.selectionChanged,
                      self.document().contentsChanged):
            signal.connect(self._repaint_mixed)
        self.comments = CommentManager(self)
        # Images
        self.images = {}         # name → {"data", "ext", "img"}
        self._img_counter = 0
        self._img_sel = None       # QTextCursor in the paragraph of the selected image
        self._img_op = None        # move or resize in progress
        self._img_cache = {}
        self._gap_image = QImage(1, 1, QImage.Format_ARGB32)
        self._gap_image.fill(Qt.transparent)
        self.document().addResource(QTextDocument.ImageResource, URL_GAP, self._gap_image)
        self.document().contentsChanged.connect(self._validate_img_sel)

    # -- pagination ---------------------------------------------------------
    def set_page_size(self, size):
        self.page_size = QSizeF(size)
        self.ensure_pagination()

    def ensure_pagination(self):
        """QTextEdit resets the page height when resized; set it again so the
        document stays split into pages."""
        if self.page_size is not None and \
                self.document().pageSize() != self.page_size:
            self.document().setPageSize(self.page_size)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.ensure_pagination()

    def showEvent(self, e):
        super().showEvent(e)
        self.ensure_pagination()

    def setFocus(self, *args):
        # Inside the page view, the real focus belongs to the view
        if self.view is not None:
            self.view.setFocus()
        super().setFocus(*args)

    def keyPressEvent(self, e):
        key_code = e.key()
        mods = e.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier
                                | Qt.AltModifier | Qt.MetaModifier)
        cursor = self.textCursor()

        # Page Up / Page Down move one screen, not the whole document
        if self.view is not None and key_code in (Qt.Key_PageUp, Qt.Key_PageDown) \
                and not mods & Qt.ControlModifier:
            self.view.scroll_screen(key_code == Qt.Key_PageDown,
                                        bool(mods & Qt.ShiftModifier))
            return

        # With an image selected: Delete/Backspace remove it, Esc releases it
        # and Enter opens a paragraph below; any other key releases it
        sel = self.selected_image()
        if sel is not None:
            if key_code in (Qt.Key_Delete, Qt.Key_Backspace) and mods == Qt.NoModifier:
                self.delete_image(sel)
                return
            if e.matches(QKeySequence.Copy):
                self.copy_selection()
                return
            if e.matches(QKeySequence.Cut):
                self.cut_selection()
                return
            if key_code == Qt.Key_Escape:
                self.deselect_image()
                return
            if key_code in (Qt.Key_Return, Qt.Key_Enter) and mods == Qt.NoModifier:
                self._paragraph_after_image(sel)
                return
            self.deselect_image()

        if key_code in (Qt.Key_Return, Qt.Key_Enter):
            if mods == Qt.NoModifier:              # Enter: new paragraph
                self.new_paragraph()
                return
            if mods == Qt.ControlModifier:         # Ctrl+Enter: page break
                self.insert_page_break()
                return

        # Backspace at the start of a new page removes the manual break
        # (without joining the paragraphs), as in a word processor
        if key_code == Qt.Key_Backspace and mods == Qt.NoModifier \
                and not cursor.hasSelection() and cursor.atBlockStart() \
                and self.is_manual_break(cursor.block()):
            self.remove_break(cursor.block())
            return

        # Delete at the end of the previous page removes the manual break
        if key_code == Qt.Key_Delete and mods == Qt.NoModifier \
                and not cursor.hasSelection() and cursor.atBlockEnd() \
                and self.is_manual_break(cursor.block().next()):
            self.remove_break(cursor.block().next())
            return

        # Backspace right after an image: selects it (does not join the
        # paragraph with it); if the paragraph is empty, removes it
        if key_code == Qt.Key_Backspace and mods == Qt.NoModifier \
                and not cursor.hasSelection() and cursor.atBlockStart():
            previous = cursor.block().previous()
            if is_image(previous) and previous.isVisible():
                block = cursor.block()
                if block.length() == 1 and block.next().isValid():
                    c = QTextCursor(self.document())
                    c.setPosition(previous.position() + previous.length() - 1)
                    c.setPosition(block.position(), QTextCursor.KeepAnchor)
                    c.removeSelectedText()
                    self._leave_image()
                else:
                    self.select_image(previous)
                return

        # Delete right before an image: same, forward
        if key_code == Qt.Key_Delete and mods == Qt.NoModifier \
                and not cursor.hasSelection() and cursor.atBlockEnd():
            following = cursor.block().next()
            if is_image(following) and following.isVisible():
                block = cursor.block()
                previous = block.previous()
                if block.length() == 1 and previous.isValid() and not is_image(previous):
                    c = QTextCursor(self.document())
                    c.setPosition(block.position() - 1)
                    c.setPosition(block.position(), QTextCursor.KeepAnchor)
                    c.removeSelectedText()
                else:
                    self.select_image(following)
                return

        # Arrow and other movement keys skip collapsed text and images (they
        # don't expand it)
        before = cursor.position()
        self._navigating = True
        try:
            super().keyPressEvent(e)
            self._skip_hidden(before)
        finally:
            self._navigating = False

    def mousePressEvent(self, e):
        e = self._text_event(e)
        if e.button() == Qt.LeftButton:
            block = self._heading_at(e.position())
            if block is not None:          # click on the triangle
                self.toggle_collapsed(block)
                e.accept()
                return
            if self._press_image(e.position()):
                e.accept()
                return
        if self._img_sel is not None:
            self.deselect_image()
        super().mousePressEvent(e)

    def mouseDoubleClickEvent(self, e):
        e = self._text_event(e)
        if self._heading_at(e.position()) is not None:
            self.toggle_collapsed(self._heading_at(e.position()))
            e.accept()
            return
        if self._image_at(e.position()) is not None:
            e.accept()
            return
        super().mouseDoubleClickEvent(e)

    def mouseMoveEvent(self, e):
        e = self._text_event(e)
        if self._img_op is not None and e.buttons() & Qt.LeftButton:
            self._move_image_op(e.position(), e.modifiers())
            if self.view is not None:
                self.view.start_auto_scroll()
            return
        if not e.buttons():
            self.viewport().setCursor(self._cursor_for(e.position()))
        super().mouseMoveEvent(e)
        # When selecting by dragging toward the edge, the view scrolls by
        # itself with a timer (never directly here: every scroll makes
        # QGraphicsView replay the last mouse move, and scrolling from here
        # caused an infinite loop)
        if self.view is not None and e.buttons() & Qt.LeftButton:
            self.view.start_auto_scroll()

    def mouseReleaseEvent(self, e):
        if self.view is not None:
            self.view.stop_auto_scroll()
        if self._img_op is not None:
            op, self._img_op = self._img_op, None
            self._finish_image_op(op)
            self.viewport().update()
            return
        super().mouseReleaseEvent(self._text_event(e))

    # -- mixed view: the text is drawn in segments ----------------------------
    def _is_mixed(self):
        return self.view is not None and self.view.mixed and bool(self._cuts)

    def set_cuts(self, cuts, gap):
        if cuts != self._cuts or gap != self._gap:
            self._cuts, self._gap = list(cuts), gap
            self.viewport().update()

    def _offset(self, y):
        """How far down on screen the text at height y is shifted."""
        return bisect.bisect_right(self._cuts, y) * self._gap

    def _to_text(self, y):
        """Height on screen → height in the text (a click in the gap between
        sheets falls at the end of the sheet above)."""
        for k in range(len(self._cuts), 0, -1):
            cut = self._cuts[k - 1]
            if y >= cut + k * self._gap:
                return y - k * self._gap
            if y >= cut + (k - 1) * self._gap:
                return cut - 0.5
        return y

    def _text_point(self, p):
        return QPointF(p.x(), self._to_text(p.y())) if self._is_mixed() else QPointF(p)

    def _text_event(self, e):
        if not self._is_mixed():
            return e
        return QMouseEvent(e.type(), self._text_point(e.position()),
                           e.globalPosition(), e.button(), e.buttons(), e.modifiers())

    def _repaint_mixed(self, *_):
        # Qt only repaints the areas that change according to its coordinates,
        # which in the mixed view don't match the screen ones: repaint
        # everything
        if self._is_mixed():
            self.viewport().update()

    def cursorRect(self, cursor=None):
        r = super().cursorRect() if cursor is None else super().cursorRect(cursor)
        if self._is_mixed():
            r.translate(0, round(self._offset(r.top())))
        return r

    def cursorForPosition(self, pos):
        if self._is_mixed():
            pos = self._text_point(QPointF(pos)).toPoint()
        return super().cursorForPosition(pos)

    def setExtraSelections(self, selections):
        # Find uses the extra selections; comments are added underneath
        self._external_selections = list(selections)
        self.refresh_selections()

    def refresh_selections(self):
        own = self.comments.selections() \
            if getattr(self, "comments", None) is not None else []
        super().setExtraSelections(own + self._external_selections)
        self._repaint_mixed()

    def focusInEvent(self, e):
        super().focusInEvent(e)
        self._repaint_mixed()

    def focusOutEvent(self, e):
        super().focusOutEvent(e)
        self._repaint_mixed()

    def dragMoveEvent(self, e):
        if self._is_mixed():
            e = QDragMoveEvent(self._text_point(e.position()).toPoint(),
                               e.possibleActions(), e.mimeData(), e.buttons(),
                               e.modifiers())
        super().dragMoveEvent(e)

    def dropEvent(self, e):
        if self._is_mixed():
            e = QDropEvent(self._text_point(e.position()), e.possibleActions(),
                           e.mimeData(), e.buttons(), e.modifiers())
        super().dropEvent(e)

    def _paint_mixed(self, zone):
        doc = self.document()
        layout = doc.documentLayout()
        ctx = QAbstractTextDocumentLayout.PaintContext()
        ctx.palette = self.palette()
        focused = self.hasFocus() or (self.view is not None and self.view.hasFocus())
        cursor = self.textCursor()
        ctx.cursorPosition = cursor.position() if focused else -1
        selections = []
        for extra in self.extraSelections():
            sel = QAbstractTextDocumentLayout.Selection()
            sel.cursor, sel.format = extra.cursor, extra.format
            selections.append(sel)
        if cursor.hasSelection():
            sel = QAbstractTextDocumentLayout.Selection()
            sel.cursor = cursor
            fmt = QTextCharFormat()
            fmt.setBackground(QColor("#B3D4FC"))
            fmt.setForeground(QColor("#1A1A1A"))
            sel.format = fmt
            selections.append(sel)
        ctx.selections = selections
        height = layout.documentSize().height()
        limits = [0.0] + self._cuts + [height]
        p = QPainter(self.viewport())
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.TextAntialiasing)
        for k in range(len(limits) - 1):
            shift = k * self._gap
            # The clip is rounded outward when zoomed out: it goes 3 px into
            # each new segment (there is only margin there) so that the last
            # line of the previous sheet doesn't leak in
            top = limits[k] + (3 if k else 0)
            segment = QRectF(0, top, self.viewport().width(),
                           limits[k + 1] - top)
            visible = segment.translated(0, shift).intersected(QRectF(zone))
            if visible.isEmpty():
                continue
            clip_rect = visible.translated(0, -shift)
            p.save()
            p.translate(0, shift)
            p.setClipRect(clip_rect)
            ctx.clip = clip_rect
            layout.draw(p, ctx)
            self._draw_extras(clip_rect.toAlignedRect(), p)
            p.restore()
        p.end()

    def paintEvent(self, e):
        if self._is_mixed():
            self._paint_mixed(e.rect())
            return
        super().paintEvent(e)
        self._draw_extras(e.rect())

    def _draw_extras(self, zone, painter=None):
        """Draw, only in the visible paragraphs of the area: the triangle to
        collapse each heading and, in the No pages view, the line of each
        page break."""
        doc = self.document()
        layout = doc.documentLayout()
        ff = doc.rootFrame().frameFormat()
        x0 = ff.leftMargin()
        x1 = self.viewport().width() - ff.rightMargin()
        no_pages = self.view is not None and self.view.no_pages
        b = doc.findBlock(layout.hitTest(QPointF(x0, zone.top()), Qt.FuzzyHit))
        if b.previous().isValid():
            b = b.previous()
        p = painter if painter is not None else QPainter(self.viewport())
        p.setRenderHint(QPainter.Antialiasing)
        font = QFont(self.font())
        font.setPointSizeF(7.5)
        p.setFont(font)
        previous_block = None                     # last visible paragraph
        while b.isValid():
            if not b.isVisible():
                b = b.next()
                continue
            r = layout.blockBoundingRect(b)
            if r.top() > zone.bottom() + 20:
                break
            if is_image(b):
                self._paint_image(p, b)
            if self.collapse_enabled and b.blockFormat().headingLevel() > 0:
                self._draw_triangle(p, b)
            if no_pages and previous_block is not None and \
                    b.blockFormat().pageBreakPolicy() & QTextFormat.PageBreak_AlwaysBefore:
                # Halfway down the gap between this paragraph and the previous
                # one
                top = layout.blockBoundingRect(previous_block).bottom()
                y = round((top + r.top()) / 2) + 0.5
                p.setPen(QPen(QColor("#A8A8A8"), 1, Qt.DashLine))
                p.drawLine(QPointF(x0, y), QPointF(x1, y))
                # The label goes in the right margin (if it fits), so as not to
                # cover the text
                text = _t("Page break")
                label_width = p.fontMetrics().horizontalAdvance(text)
                if x1 + 8 + label_width <= self.viewport().width() - 4:
                    p.setPen(QColor("#909090"))
                    p.drawText(QRectF(x1 + 8, y - 8, label_width + 2, 16),
                               Qt.AlignLeft | Qt.AlignVCenter, text)
            previous_block = b
            b = b.next()
        self._draw_image_selection(p)
        if painter is None:
            p.end()

    # -- collapsible headings --------------------------------------------------
    def _triangle_rect(self, block):
        """Area of a heading's triangle: in the left margin, at the height of
        its first line."""
        layout = block.layout()
        if layout is None or layout.lineCount() == 0:
            return None
        line = layout.lineAt(0)
        cy = layout.position().y() + line.y() + line.height() / 2
        x = self.document().rootFrame().frameFormat().leftMargin() \
            + block.blockFormat().leftMargin() - 14
        return QRectF(x - 8, cy - 8, 16, 16)

    def _draw_triangle(self, p, block):
        r = self._triangle_rect(block)
        if r is None:
            return
        cx, cy = r.center().x(), r.center().y()
        collapsed = block.userState() == STATE_COLLAPSED
        p.setPen(Qt.NoPen)
        if collapsed:        # ▸ dark: there is hidden text
            p.setBrush(QColor("#555555"))
            p.drawPolygon(QPolygonF([QPointF(cx - 2.5, cy - 4.5),
                                     QPointF(cx + 3.5, cy), QPointF(cx - 2.5, cy + 4.5)]))
        else:                # ▾ faint
            p.setBrush(QColor("#BDBDBD"))
            p.drawPolygon(QPolygonF([QPointF(cx - 4.5, cy - 2.5),
                                     QPointF(cx + 4.5, cy - 2.5), QPointF(cx, cy + 3.5)]))

    def _heading_at(self, point):
        """The heading whose triangle is under the point, or None."""
        if not self.collapse_enabled:
            return None
        doc = self.document()
        x = doc.rootFrame().frameFormat().leftMargin()
        if point.x() >= x:                  # the triangle is in the margin
            return None
        pos = doc.documentLayout().hitTest(QPointF(x + 1, point.y()), Qt.FuzzyHit)
        block = doc.findBlock(max(0, pos))
        if not block.isValid() or block.blockFormat().headingLevel() <= 0:
            return None
        r = self._triangle_rect(block)
        return block if r is not None and r.contains(point) else None

    def toggle_collapsed(self, block):
        """Collapse or expand a heading: hides everything that follows up to
        the next heading of the same or a higher level (a Title hides its
        Subtitles and Sections; a Section, only its text)."""
        collapse = block.userState() != STATE_COLLAPSED
        block.setUserState(STATE_COLLAPSED if collapse else -1)
        self.apply_folding()
        if not self.textCursor().block().isVisible():
            # The cursor ended up inside the hidden part: move it to the end of
            # the heading
            c = QTextCursor(block)
            c.movePosition(QTextCursor.EndOfBlock)
            self._navigating = True
            self.setTextCursor(c)
            self._navigating = False
        self.viewport().update()

    def apply_folding(self):
        """Show or hide paragraphs according to the collapsed headings."""
        self._folding_pending = False
        doc = self.document()
        b = doc.begin()
        hide_below = 0       # level of the active collapsed heading (0 = none)
        collapsed_count = 0
        while b.isValid():
            level = b.blockFormat().headingLevel()
            is_collapsed = level > 0 and b.userState() == STATE_COLLAPSED
            collapsed_count += is_collapsed
            if hide_below and (level == 0 or level > hide_below):
                visible = False
            else:
                visible = True
                hide_below = level if is_collapsed else 0
            if b.isVisible() != visible:
                b.setVisible(visible)
                doc.markContentsDirty(b.position(), b.length())
            b = b.next()
        self._collapsed = collapsed_count

    def set_collapse(self, active_id):
        """Turn collapsible headings on or off. Turning them off expands
        everything and removes the triangles."""
        self.collapse_enabled = bool(active_id)
        if not self.collapse_enabled:
            self.expand_all()
            self.viewport().setCursor(Qt.IBeamCursor)
        self.viewport().update()

    def expand_all(self):
        b = self.document().begin()
        while b.isValid():
            if b.userState() == STATE_COLLAPSED:
                b.setUserState(-1)
            b = b.next()
        self.apply_folding()

    def _schedule_folding(self):
        # When editing (e.g., removing the style of a collapsed heading) it is
        # recalculated; only if something is collapsed, so as not to slow down
        # long books
        if self._collapsed and not self._folding_pending:
            self._folding_pending = True
            QTimer.singleShot(0, self.apply_folding)

    def _expand_to(self, block):
        """Expand the headings that hide the paragraph."""
        level = block.blockFormat().headingLevel()
        limit = level if level > 0 else 7
        b = block.previous()
        while b.isValid() and limit > 1:
            n = b.blockFormat().headingLevel()
            if 0 < n < limit:
                if b.userState() == STATE_COLLAPSED:
                    b.setUserState(-1)
                limit = n
            b = b.previous()
        self.apply_folding()
        self.viewport().update()

    def _cursor_in_hidden(self):
        """If the cursor reaches hidden text some other way (Find, undo, etc.),
        expand what is needed to show it. If it reaches an image paragraph,
        leave it (images are not text)."""
        if self._navigating:
            return
        if is_image(self.textCursor().block()):
            self._leave_image()
            return
        if not self._collapsed:
            return
        block = self.textCursor().block()
        if not block.isVisible():
            self._expand_to(block)

    def _skip_hidden(self, before):
        """After a movement key, if the cursor landed in hidden text, move it
        to the next (or previous) visible paragraph."""
        c = self.textCursor()
        if not self._not_navigable(c.block()):
            return
        forward = c.position() >= before
        b = c.block()
        while b.isValid() and self._not_navigable(b):
            b = b.next() if forward else b.previous()
        mode = QTextCursor.KeepAnchor if c.hasSelection() else QTextCursor.MoveAnchor
        if not b.isValid():                 # nothing visible after it
            b = c.block()
            while b.isValid() and self._not_navigable(b):
                b = b.previous()
            forward = False
        if not b.isValid():
            return
        c.setPosition(b.position() if forward else b.position() + b.length() - 1, mode)
        self.setTextCursor(c)

    # -- images -----------------------------------------------------------------
    @staticmethod
    def _not_navigable(b):
        """Paragraphs where the cursor does not stop."""
        return not b.isVisible() or is_image(b)

    def _leave_image(self):
        """Move the cursor to the next text paragraph (or the previous one)."""
        c = self.textCursor()
        mode = QTextCursor.KeepAnchor if c.hasSelection() else QTextCursor.MoveAnchor
        b = c.block().next()
        while b.isValid() and self._not_navigable(b):
            b = b.next()
        if b.isValid():
            c.setPosition(b.position(), mode)
        else:
            b = c.block().previous()
            while b.isValid() and self._not_navigable(b):
                b = b.previous()
            if not b.isValid():
                return
            c.setPosition(b.position() + b.length() - 1, mode)
        self._navigating = True
        self.setTextCursor(c)
        self._navigating = False

    def column_width(self):
        """Width of the text between the page margins."""
        if self.view is not None:
            c = self.view.config
            return max(20.0, self.view.cm_to_px(c["width"] - c["left"] - c["right"]))
        doc = self.document()
        ff = doc.rootFrame().frameFormat()
        width = doc.pageSize().width()
        if width <= 0:
            width = self.viewport().width()
        return max(20.0, width - ff.leftMargin() - ff.rightMargin())

    def max_image_height(self):
        """The tallest an image can be: the usable height of the page."""
        if self.view is None:
            return 1e6
        c = self.view.config
        return max(20.0, self.view.cm_to_px(c["height"] - c["top"] - c["bottom"])
                   - 2 * IMG_MARGIN - 4)

    def image_size(self, fmt):
        """Size at which it is shown: the saved one, reduced if it doesn't fit."""
        width = fmt.doubleProperty(PROP_IMG_WIDTH) or 100.0
        height = fmt.doubleProperty(PROP_IMG_HEIGHT) or 100.0
        k = min(1.0, (self.column_width() - 2) / width,
                self.max_image_height() / height)
        return width * k, height * k

    def image_block_format(self, page_break=False):
        bf = QTextBlockFormat()
        bf.setAlignment(Qt.AlignHCenter)
        bf.setTopMargin(IMG_MARGIN)
        bf.setBottomMargin(IMG_MARGIN)
        bf.setLeftMargin(0)
        bf.setRightMargin(0)
        bf.setTextIndent(0)
        bf.setLineHeight(100, QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
        bf.setHeadingLevel(0)
        bf.setProperty(PROP_IMAGE_BLOCK, True)
        if page_break:
            bf.setPageBreakPolicy(QTextFormat.PageBreak_AlwaysBefore)
            bf.setProperty(PROP_MANUAL_BREAK, True)
        return bf

    def image_format(self, name, width, height):
        cf = QTextImageFormat()
        cf.merge(self.char_format())
        cf.setName(URL_GAP.toString())
        cf.setProperty(PROP_IMG_NAME, name)
        cf.setProperty(PROP_IMG_WIDTH, float(width))
        cf.setProperty(PROP_IMG_HEIGHT, float(height))
        self._adjust_gap(cf)
        return cf

    def _adjust_gap(self, cf):
        """The gap takes the size at which the image is shown."""
        width, height = self.image_size(cf)
        cf.setProperty(QTextFormat.ImageWidth, float(width))
        cf.setProperty(QTextFormat.ImageHeight, float(height))

    def fit_images(self):
        """After changing the page: images that no longer fit are reduced (and
        those that had been reduced get their size back if they now fit)."""
        b = self.document().begin()
        while b.isValid():
            if is_image(b):
                cf, pos = self.object_format(b)
                if cf is not None:
                    width, height = self.image_size(cf)
                    if abs(cf.doubleProperty(QTextFormat.ImageWidth) - width) > 0.5 or \
                            abs(cf.doubleProperty(QTextFormat.ImageHeight) - height) > 0.5:
                        self._adjust_gap(cf)
                        c = QTextCursor(self.document())
                        c.setPosition(pos)
                        c.setPosition(pos + 1, QTextCursor.KeepAnchor)
                        c.setCharFormat(cf)
            b = b.next()

    @staticmethod
    def shown_size(cf):
        return cf.doubleProperty(QTextFormat.ImageWidth), \
            cf.doubleProperty(QTextFormat.ImageHeight)

    @staticmethod
    def object_format(block):
        """(format, position) of the paragraph's image object, or (None, -1)."""
        it = block.begin()
        while not it.atEnd():
            f = it.fragment()
            if f.isValid() and f.charFormat().isImageFormat() and \
                    f.charFormat().hasProperty(PROP_IMG_NAME):
                return f.charFormat().toImageFormat(), f.position()
            it += 1
        return None, -1

    def _has_image_object(self, block):
        return self.object_format(block)[0] is not None

    def register_image(self, data, ext, target=None):
        """Store an image's data and return its name."""
        target = self.images if target is None else target
        self._img_counter += 1
        while f"img{self._img_counter}" in target:
            self._img_counter += 1
        name = f"img{self._img_counter}"
        target[name] = {"data": bytes(data), "ext": ext.lower(),
                           "img": QImage.fromData(QByteArray(bytes(data)))}
        return name

    def clear_images(self):
        self.images = {}
        self._img_cache = {}
        self._img_sel = None
        self._img_op = None
        self.document().addResource(QTextDocument.ImageResource, URL_GAP, self._gap_image)

    def insert_image_file(self, path):
        with open(path, "rb") as f:
            data = f.read()
        img = QImage.fromData(QByteArray(data))
        if img.isNull():
            raise ValueError(_t("The file is not an image that can be read."))
        ext = os.path.splitext(path)[1].lower().lstrip(".")
        if ext not in ("png", "jpg", "jpeg", "gif", "bmp"):
            # Formats Word does not support (webp, etc.) are saved as PNG
            buf = QBuffer()
            buf.open(QIODevice.WriteOnly)
            img.save(buf, "PNG")
            data, ext = bytes(buf.data()), "png"
        self.insert_image(data, ext)

    def insert_image(self, data, ext):
        """Insert the image in its own paragraph: before the cursor's paragraph
        if the cursor is at its start; otherwise, after it."""
        name = self.register_image(data, ext)
        img = self.images[name]["img"]
        width, height = (img.width(), img.height()) if not img.isNull() else (300, 200)
        k = min(1.0, (self.column_width() - 2) / width, self.max_image_height() / height)
        cf = self.image_format(name, width * k, height * k)
        step = QTextCursor(self.document())
        step.beginEditBlock()
        new = self._place_image(cf)
        step.endEditBlock()
        self.select_image(new)
        return new

    def _place_image(self, cf):
        """Place an image next to the cursor's paragraph: before it if the
        cursor is at the start of the paragraph; otherwise, after it."""
        cursor = self.textCursor()
        block = self.document().findBlock(cursor.selectionEnd())
        after = cursor.hasSelection() or not cursor.atBlockStart()
        return self._insert_image_block(block, after,
                                            self.image_block_format(), cf)

    def _insert_image_block(self, target_block, after, bf, cf):
        """Create the image paragraph before or after `target` and return that
        paragraph. There is always a text paragraph after it."""
        doc = self.document()
        c = QTextCursor(doc)
        if not after and target_block.previous().isValid():
            target_block, after = target_block.previous(), True
        if after:
            c.setPosition(target_block.position() + target_block.length() - 1)
            c.insertBlock(bf, self.char_format())
            c.insertText(OBJECT_CHAR, cf)
            new = c.block()
            if not new.next().isValid():
                c.insertBlock(self.paragraph_format(), self.char_format())
            return doc.findBlock(new.position())
        # Before the first paragraph of the document
        bf0, cf0 = target_block.blockFormat(), target_block.charFormat()
        c.setPosition(0)
        c.insertText(OBJECT_CHAR, cf)
        c.insertBlock(bf0, cf0)          # the rest keeps its format
        first = doc.begin()
        k = QTextCursor(first)
        k.setBlockFormat(bf)
        k.setBlockCharFormat(self.char_format())
        return doc.begin()

    def _remove_image_block(self, block):
        """Remove the image paragraph without altering the neighboring
        paragraphs."""
        doc = self.document()
        c = QTextCursor(doc)
        if block.previous().isValid():
            c.setPosition(block.position() - 1)
            c.setPosition(block.position() + block.length() - 1, QTextCursor.KeepAnchor)
            c.removeSelectedText()
            return
        following = block.next()
        if not following.isValid():
            c.setPosition(0)
            c.setPosition(block.length() - 1, QTextCursor.KeepAnchor)
            c.removeSelectedText()
            c.setBlockFormat(self.paragraph_format())
            return
        bf, cf = following.blockFormat(), following.charFormat()
        c.setPosition(0)
        c.setPosition(following.position(), QTextCursor.KeepAnchor)
        c.removeSelectedText()
        c.setBlockFormat(bf)
        c.setBlockCharFormat(cf)

    def delete_image(self, block):
        following = QTextCursor(block.next()) if block.next().isValid() else None
        step = QTextCursor(self.document())
        step.beginEditBlock()
        self._remove_image_block(block)
        step.endEditBlock()
        self._img_sel = None
        if following is not None:
            c = QTextCursor(self.document())
            c.setPosition(following.position())
            self.setTextCursor(c)
        self.viewport().update()

    def move_image(self, block, target):
        """Move the image to another place between paragraphs (a single undo
        step)."""
        cf, _ = self.object_format(block)
        if cf is None:
            return
        bf = block.blockFormat()
        ref, after = target[0], target[1]
        step = QTextCursor(self.document())
        step.beginEditBlock()
        self._remove_image_block(block)
        new = self._insert_image_block(ref.block(), after, bf, cf)
        step.endEditBlock()
        self.select_image(new)

    def _resize_image(self, block, width, height):
        cf, pos = self.object_format(block)
        if cf is None:
            return
        cf.setProperty(PROP_IMG_WIDTH, float(width))
        cf.setProperty(PROP_IMG_HEIGHT, float(height))
        self._adjust_gap(cf)
        c = QTextCursor(self.document())
        c.setPosition(pos)
        c.setPosition(pos + 1, QTextCursor.KeepAnchor)
        c.setCharFormat(cf)

    def _paragraph_after_image(self, block):
        c = QTextCursor(block)
        c.movePosition(QTextCursor.EndOfBlock)
        c.insertBlock(self.paragraph_format(), self.char_format())
        self._img_sel = None
        self._navigating = True
        self.setTextCursor(c)
        self._navigating = False
        self.setCurrentCharFormat(self.char_format())
        self.viewport().update()

    # -- image selection ---------------------------------------------------
    def selected_image(self):
        if self._img_sel is None:
            return None
        b = self._img_sel.block()
        if is_image(b) and b.isVisible():
            return b
        self._img_sel = None
        return None

    # -- clipboard --------------------------------------------------------------
    def _image_mime(self, block):
        """The selected image for the clipboard: in our own format (to paste it
        here with its size) and as an image (for other programs)."""
        cf, _ = self.object_format(block)
        md = QMimeData()
        if cf is None:
            return md
        name = cf.stringProperty(PROP_IMG_NAME)
        md.setData(MIME_BOOK, QByteArray(json.dumps([{
            "img": name, "w": cf.doubleProperty(PROP_IMG_WIDTH),
            "h": cf.doubleProperty(PROP_IMG_HEIGHT)}]).encode("utf-8")))
        info = self.images.get(name)
        if info and not info["img"].isNull():
            md.setImageData(info["img"])
        return md

    def can_copy(self):
        return self.textCursor().hasSelection() or self.selected_image() is not None

    def copy_selection(self):
        sel = self.selected_image()
        if sel is not None:
            QApplication.clipboard().setMimeData(self._image_mime(sel))
        else:
            self.copy()

    def cut_selection(self):
        sel = self.selected_image()
        if sel is not None:
            QApplication.clipboard().setMimeData(self._image_mime(sel))
            self.delete_image(sel)
        else:
            self.cut()

    def paste_clipboard(self):
        self.deselect_image()
        self.paste()

    def select_image(self, block):
        self._img_sel = QTextCursor(block)
        # The text cursor moves to the next paragraph: what is typed goes below
        # the image
        c = QTextCursor(block)
        self._navigating = True
        self.setTextCursor(c)
        self._navigating = False
        self._leave_image()
        self.viewport().update()

    def deselect_image(self):
        if self._img_sel is not None:
            self._img_sel = None
            self.viewport().update()

    def _validate_img_sel(self):
        if self._img_sel is not None:
            if not is_image(self._img_sel.block()):
                self._img_sel = None
            self.viewport().update()

    def image_rect(self, block):
        """Rectangle of the image (text coordinates), or None."""
        if not is_image(block) or not block.isVisible():
            return None
        layout = block.layout()
        if layout is None or layout.lineCount() == 0:
            return None
        cf, _ = self.object_format(block)
        if cf is None:
            return None
        width, height = self.shown_size(cf)
        line = layout.lineAt(0)
        top = layout.position().y() + line.y() + line.ascent() - height
        center = self.document().rootFrame().frameFormat().leftMargin() \
            + self.column_width() / 2
        return QRectF(center - width / 2, top, width, height)

    def _image_at(self, point):
        doc = self.document()
        x = doc.rootFrame().frameFormat().leftMargin() + 1
        pos = doc.documentLayout().hitTest(QPointF(x, point.y()), Qt.FuzzyHit)
        b = doc.findBlock(max(0, pos))
        for cand in (b, b.previous(), b.next()):
            r = self.image_rect(cand)
            if r is not None and r.contains(point):
                return cand
        return None

    def _zoom_factor(self):
        return self.view.zoom if self.view is not None else 1.0

    def _handles(self, r):
        """The 8 drag handles: {(sx, sy): point}."""
        cx, cy = r.center().x(), r.center().y()
        xs = {-1: r.left(), 0: cx, 1: r.right()}
        ys = {-1: r.top(), 0: cy, 1: r.bottom()}
        return {(sx, sy): QPointF(xs[sx], ys[sy])
                for sx in (-1, 0, 1) for sy in (-1, 0, 1) if sx or sy}

    def _handle_at(self, block, point):
        r = self.image_rect(block)
        if r is None:
            return None
        t = (HANDLE + 4) / self._zoom_factor()
        for key, q in self._handles(r).items():
            if abs(q.x() - point.x()) <= t / 2 and abs(q.y() - point.y()) <= t / 2:
                return key
        return None

    def _cursor_for(self, point):
        if self._heading_at(point) is not None:
            return Qt.PointingHandCursor
        sel = self.selected_image()
        if sel is not None:
            handle = self._handle_at(sel, point)
            if handle is not None:
                if handle[0] and handle[1]:
                    return Qt.SizeFDiagCursor if handle[0] == handle[1] else Qt.SizeBDiagCursor
                return Qt.SizeHorCursor if handle[0] else Qt.SizeVerCursor
        if self._image_at(point) is not None:
            return Qt.SizeAllCursor
        return Qt.IBeamCursor

    # -- moving and resizing with the mouse ----------------------------------
    def _press_image(self, point):
        sel = self.selected_image()
        if sel is not None:
            handle = self._handle_at(sel, point)
            if handle is not None:
                r = self.image_rect(sel)
                self._img_op = {"kind": "resize", "handle": handle, "start": QPointF(point),
                                "rect": r, "new": QRectF(r)}
                return True
        b = self._image_at(point)
        if b is None:
            return False
        self.select_image(b)
        self._img_op = {"kind": "move", "start": QPointF(point), "active": False,
                        "target": None}
        return True

    def _move_image_op(self, point, mods):
        op = self._img_op
        if op["kind"] == "move":
            if not op["active"]:
                if (point - op["start"]).manhattanLength() < 5 / self._zoom_factor():
                    return
                op["active"] = True
            op["target"] = self._image_target(point)
        else:
            op["new"] = self._resized_rect(
                op, point, bool(mods & Qt.ShiftModifier))
        self.viewport().update()

    def _resized_rect(self, op, point, free):
        """New size according to the dragged point. The image stays centered:
        when a side edge moves, the opposite one moves the same. Corners
        keep the aspect ratio (not with Shift)."""
        r, (sx, sy) = op["rect"], op["handle"]
        dx = point.x() - op["start"].x()
        dy = point.y() - op["start"].y()
        w0, h0 = r.width(), r.height()
        nw = w0 + 2 * sx * dx if sx else w0
        nh = h0 + sy * dy if sy else h0
        maxw, maxh = self.column_width() - 2, self.max_image_height()
        if sx and sy and not free:
            kx, ky = nw / w0, nh / h0
            k = kx if abs(kx - 1) >= abs(ky - 1) else ky
            k = max(k, IMG_MIN / min(w0, h0))
            k = min(k, maxw / w0, maxh / h0)
            nw, nh = w0 * k, h0 * k
        else:
            nw = max(IMG_MIN, min(maxw, nw))
            nh = max(IMG_MIN, min(maxh, nh))
        return QRectF(r.center().x() - nw / 2, r.top(), nw, nh)

    def _image_target(self, point):
        """Where the image would land when dropped: between two paragraphs.
        Returns (cursor of the paragraph, after it?, y of the guide line)."""
        doc = self.document()
        layout = doc.documentLayout()
        x = doc.rootFrame().frameFormat().leftMargin() + 1
        b = doc.findBlock(max(0, layout.hitTest(QPointF(x, point.y()), Qt.FuzzyHit)))
        while b.isValid() and not b.isVisible():
            b = b.previous()
        if not b.isValid():
            return None
        r = layout.blockBoundingRect(b)
        after = point.y() > r.center().y()
        sel = self.selected_image()
        if sel is not None:
            def neighbor(blk, forward):
                blk = blk.next() if forward else blk.previous()
                while blk.isValid() and not blk.isVisible():
                    blk = blk.next() if forward else blk.previous()
                return blk
            if b == sel or (after and b == neighbor(sel, False)) or \
                    (not after and b == neighbor(sel, True)):
                return None                 # would stay where it already is
        y = r.bottom() if after else r.top()
        return QTextCursor(b), after, y

    def _finish_image_op(self, op):
        sel = self.selected_image()
        if sel is None:
            return
        if op["kind"] == "resize":
            n, r = op["new"], op["rect"]
            if abs(n.width() - r.width()) > 0.5 or abs(n.height() - r.height()) > 0.5:
                self._resize_image(sel, n.width(), n.height())
        elif op["active"] and op["target"] is not None:
            self.move_image(sel, op["target"])

    def _paint_image(self, p, block):
        r = self.image_rect(block)
        if r is None:
            return
        cf, _ = self.object_format(block)
        name = cf.stringProperty(PROP_IMG_NAME)
        info = self.images.get(name)
        img = info["img"] if info else None
        p.save()
        if img is None or img.isNull():          # no data: a placeholder box
            p.setPen(QPen(QColor("#B0B0B0"), 1, Qt.DashLine))
            p.setBrush(QColor("#F2F2F2"))
            p.drawRect(r)
            p.setPen(QColor("#909090"))
            p.drawText(r, Qt.AlignCenter, _t("Image"))
            p.restore()
            return
        # Scaled once per on-screen size (with the zoom) and cached
        scale_factor = max(abs(p.deviceTransform().m11()), 0.1)
        device = p.device()
        if device is not None:
            scale_factor *= device.devicePixelRatioF()
        key = (name, round(r.width() * scale_factor), round(r.height() * scale_factor))
        scaled = self._img_cache.get(key)
        if scaled is None:
            if len(self._img_cache) > 40:
                self._img_cache.clear()
            scaled = img.scaled(max(1, key[1]), max(1, key[2]),
                                  Qt.IgnoreAspectRatio, Qt.SmoothTransformation)
            self._img_cache[key] = scaled
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        p.drawImage(r, scaled)
        p.restore()

    def _draw_image_selection(self, p):
        """Frame and 8 handles of the selected image; guide while moving it;
        outline of the new size while resizing it."""
        op = self._img_op
        z = self._zoom_factor()
        blue = QColor(IMAGE_SELECTION_COLOR)
        p.save()
        p.setRenderHint(QPainter.Antialiasing)
        if op is not None and op["kind"] == "move" and op["active"] and op["target"]:
            y = op["target"][2]
            x0 = self.document().rootFrame().frameFormat().leftMargin()
            x1 = x0 + self.column_width()
            p.setPen(QPen(blue, 2.5 / z))
            p.drawLine(QPointF(x0, y), QPointF(x1, y))
            p.setPen(Qt.NoPen)
            p.setBrush(blue)
            t = 6 / z
            for x, d in ((x0, 1), (x1, -1)):
                p.drawPolygon(QPolygonF([QPointF(x, y - t), QPointF(x + d * t, y),
                                         QPointF(x, y + t)]))
        sel = self.selected_image()
        r = self.image_rect(sel) if sel is not None else None
        if r is not None:
            if op is not None and op["kind"] == "resize":
                p.setPen(QPen(blue, 1 / z, Qt.DashLine))
                p.setBrush(QColor(74, 134, 232, 35))
                r = op["new"]
                p.drawRect(r)
            p.setPen(QPen(blue, 1.5 / z))
            p.setBrush(Qt.NoBrush)
            p.drawRect(r)
            t = HANDLE / z
            p.setPen(QPen(blue, 1.2 / z))
            p.setBrush(QColor("#FFFFFF"))
            for q in self._handles(r).values():
                p.drawRect(QRectF(q.x() - t / 2, q.y() - t / 2, t, t))
        p.restore()

    # -- the tools don't touch images ----------------------------------------
    def _images_in_selection(self, cursor):
        if not cursor.hasSelection():
            return []
        doc = self.document()
        b = doc.findBlock(cursor.selectionStart())
        end = doc.findBlock(cursor.selectionEnd())
        saved = []
        while b.isValid():
            if is_image(b):
                cf, _ = self.object_format(b)
                saved.append((QTextCursor(b), b.blockFormat(), b.charFormat(), cf))
            if b == end:
                break
            b = b.next()
        return saved

    def _restore_images(self, saved):
        for c, bf, mark, cf in saved:
            b = c.block()
            k = QTextCursor(b)
            k.setBlockFormat(bf)
            k.setBlockCharFormat(mark)
            if cf is not None:
                k.setPosition(b.position())
                k.setPosition(b.position() + 1, QTextCursor.KeepAnchor)
                k.setCharFormat(cf)

    # -- page breaks and paragraphs ------------------------------------------
    @staticmethod
    def is_manual_break(block):
        return block.isValid() and \
            block.blockFormat().boolProperty(PROP_MANUAL_BREAK)

    @staticmethod
    def _put_break(block):
        fmt = QTextBlockFormat()
        fmt.setPageBreakPolicy(QTextFormat.PageBreak_AlwaysBefore)
        fmt.setProperty(PROP_MANUAL_BREAK, True)
        QTextCursor(block).mergeBlockFormat(fmt)

    def remove_break(self, block):
        fmt = block.blockFormat()
        fmt.setPageBreakPolicy(QTextFormat.PageBreak_Auto)
        fmt.clearProperty(PROP_MANUAL_BREAK)
        QTextCursor(block).setBlockFormat(fmt)
        self.paragraphFormatChanged.emit()

    def new_paragraph(self, page_break_before=False):
        """Enter. The new paragraph keeps the format of the previous one, but
        not its page break (otherwise every Enter would open a sheet)."""
        cursor = self.textCursor()
        cursor.beginEditBlock()                 # a single undo step
        if cursor.hasSelection():
            cursor.removeSelectedText()
        cursor.insertBlock()

        fmt = cursor.blockFormat()
        fmt.setPageBreakPolicy(QTextFormat.PageBreak_Auto)
        fmt.clearProperty(PROP_MANUAL_BREAK)
        cursor.setBlockFormat(fmt)

        # Enter at the end of a Title, Subtitle or Section: what follows is
        # normal text, so the new paragraph becomes Body (if the heading is
        # split in the middle, both parts remain headings)
        to_body = cursor.block().previous().blockFormat().headingLevel() > 0 \
            and cursor.block().length() == 1
        if to_body:
            cursor.setBlockFormat(self.paragraph_format(CHAPTER_BODY))
            cursor.setBlockCharFormat(self.char_format(CHAPTER_BODY))
            cursor.setCharFormat(self.char_format(CHAPTER_BODY))

        if page_break_before:
            self._put_break(cursor.block())
        cursor.endEditBlock()

        self.setTextCursor(cursor)
        if to_body:                     # whatever is typed, already in Body
            self.setCurrentCharFormat(self.char_format(CHAPTER_BODY))
        self.paragraphFormatChanged.emit()

    def insert_page_break(self):
        """Ctrl+Enter: what follows the cursor moves to a new page. Not allowed
        in the No pages view (there are no pages to cut)."""
        if self.view is not None and self.view.no_pages:
            window = self.view.window()    # the editor lives in the scene
            if isinstance(window, QMainWindow):
                window.statusBar().showMessage(
                    _t("Page breaks are not available in the No pages view "
                       "(Settings → View → Pages or Mixed)"), 4000)
            QApplication.beep()
            self.setFocus()
            return
        self.new_paragraph(page_break_before=True)
        self.setFocus()

    def _set_block_format(self, cursor, fmt):
        """Like cursor.setBlockFormat, but keeps manual page breaks (changing a
        paragraph's style must not remove them)."""
        manual = []
        doc = self.document()
        b = doc.findBlock(cursor.selectionStart())
        end = doc.findBlock(cursor.selectionEnd())
        while b.isValid():
            if self.is_manual_break(b):
                manual.append(b.position())
            if b == end:
                break
            b = b.next()
        cursor.setBlockFormat(fmt)
        for pos in manual:
            self._put_break(doc.findBlock(pos))

    def char_format(self, style=CHAPTER_BODY):
        fmt = QTextCharFormat()
        fmt.setFontFamilies([style["font_family"]])
        fmt.setFontPointSize(style["font_size"])
        fmt.setForeground(QColor(style["font_color"]))
        fmt.setFontWeight(QFont.Bold if style["bold"] else QFont.Normal)
        fmt.setFontItalic(style["italic"])
        return fmt

    def paragraph_format(self, style=CHAPTER_BODY):
        fmt = QTextBlockFormat()
        fmt.setAlignment(style["alignment"])
        fmt.setLineHeight(style["line_spacing"],
                          QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
        fmt.setTopMargin(style["space_before"])
        fmt.setBottomMargin(style["space_after"])
        fmt.setLeftMargin(style["indent_left"])
        fmt.setRightMargin(style["indent_right"])
        fmt.setTextIndent(style["first_line_indent"])
        fmt.setHeadingLevel(style["heading_level"])
        return fmt

    def _initial_style(self):
        """The whole document starts with the Chapter body style."""
        font = QFont(CHAPTER_BODY["font_family"], CHAPTER_BODY["font_size"])
        self.document().setDefaultFont(font)
        self.setFont(font)
        self.setTextColor(QColor(CHAPTER_BODY["font_color"]))

        cursor = self.textCursor()
        cursor.select(QTextCursor.Document)
        cursor.setBlockFormat(self.paragraph_format())
        cursor.setCharFormat(self.char_format())
        cursor.clearSelection()
        self.setTextCursor(cursor)
        self.setCurrentCharFormat(self.char_format())

    def loadResource(self, kind, url):
        if url == URL_GAP:                 # also after the document is cleared
            return self._gap_image
        return super().loadResource(kind, url)

    def canInsertFromMimeData(self, source):
        return source.hasImage() or source.hasFormat(MIME_BOOK) or \
            super().canInsertFromMimeData(source)

    def createMimeDataFromSelection(self):
        """If the selection includes images, it is also copied in our own
        format so they are kept when pasting (or dragging)."""
        md = super().createMimeDataFromSelection()
        c = self.textCursor()
        if not c.hasSelection():
            return md
        start, end = c.selectionStart(), c.selectionEnd()
        doc = self.document()
        elements, has_image = [], False
        b = doc.findBlock(start)
        while b.isValid() and b.position() <= end:
            if is_image(b):
                cf, pos = self.object_format(b)
                if cf is not None and start <= pos < end:
                    elements.append({"img": cf.stringProperty(PROP_IMG_NAME),
                                      "w": cf.doubleProperty(PROP_IMG_WIDTH),
                                      "h": cf.doubleProperty(PROP_IMG_HEIGHT)})
                    has_image = True
            else:
                chunks = []
                it = b.begin()
                while not it.atEnd():
                    f = it.fragment()
                    if f.isValid():
                        a, z = max(start, f.position()), min(end, f.position() + f.length())
                        if a < z:
                            text = f.text()[a - f.position():z - f.position()]
                            chunks.append([text, self._attributes(f.charFormat())])
                    it += 1
                elements.append({"p": chunks})
            b = b.next()
        if has_image:
            # QTextEdit's QMimeData does not accept added formats: copy it
            own_mime = QMimeData()
            for format_spec in md.formats():
                own_mime.setData(format_spec, md.data(format_spec))
            own_mime.setData(MIME_BOOK, QByteArray(json.dumps(elements).encode("utf-8")))
            return own_mime
        return md

    @staticmethod
    def _attributes(cf):
        """The attributes kept when pasting, as a serializable dict."""
        fmt = Editor._pasted_format(cf, QTextCharFormat())
        va = fmt.verticalAlignment()
        return {
            "b": fmt.fontWeight() >= QFont.DemiBold, "i": fmt.fontItalic(),
            "u": fmt.fontUnderline(), "s": fmt.fontStrikeOut(),
            "va": "sub" if va == QTextCharFormat.AlignSubScript else
                  "top" if va == QTextCharFormat.AlignSuperScript else "",
            "fg": fmt.foreground().color().name()
                  if fmt.hasProperty(QTextFormat.ForegroundBrush) else "",
            "bg": fmt.background().color().name()
                  if fmt.hasProperty(QTextFormat.BackgroundBrush) else "",
        }

    def _format_from_attributes(self, d, base):
        fmt = QTextCharFormat(base)
        if d.get("b"):
            fmt.setFontWeight(QFont.Bold)
        if d.get("i"):
            fmt.setFontItalic(True)
        if d.get("u"):
            fmt.setFontUnderline(True)
        if d.get("s"):
            fmt.setFontStrikeOut(True)
        if d.get("va") == "sub":
            fmt.setVerticalAlignment(QTextCharFormat.AlignSubScript)
        elif d.get("va") == "top":
            fmt.setVerticalAlignment(QTextCharFormat.AlignSuperScript)
        if d.get("fg") and QColor(d["fg"]).isValid():
            fmt.setForeground(QColor(d["fg"]))
        if d.get("bg") and QColor(d["bg"]).isValid():
            fmt.setBackground(QColor(d["bg"]))
        return fmt

    def insertFromMimeData(self, source):
        """When pasting, the text takes the Chapter body style (font, size and
        paragraph format), but keeps bold, italic, underline, strikethrough,
        subscript, superscript, text color and highlight. Pasted paragraphs
        don't inherit the page break of the paragraph where they are pasted."""
        base = self.char_format()
        if source.hasFormat(MIME_BOOK):          # copied here, with images
            try:
                elements = json.loads(bytes(source.data(MIME_BOOK)).decode("utf-8"))
            except ValueError:
                elements = []
            paragraphs = []
            for el in elements:
                if "img" in el:
                    if el["img"] in self.images:
                        paragraphs.append(("img", el["img"], el["w"], el["h"]))
                else:
                    paragraphs.append([(t, self._format_from_attributes(d, base))
                                     for t, d in el.get("p", []) if t])
        elif source.hasImage() and not source.hasText():   # copied image
            img = QImage(source.imageData())
            if not img.isNull():
                buf = QBuffer()
                buf.open(QIODevice.WriteOnly)
                img.save(buf, "PNG")
                self.insert_image(bytes(buf.data()), "png")
            return
        elif source.hasHtml():
            source_doc = QTextDocument()
            source_doc.setHtml(source.html())
            paragraphs = self._pasted_paragraphs(source_doc)
            # Browsers often add an empty paragraph at the end
            if not source.text().endswith(("\n", "\u2029")):
                while len(paragraphs) > 1 and not paragraphs[-1]:
                    paragraphs.pop()
        elif source.hasText():
            base = self.char_format()
            text = source.text().replace("\r\n", "\n").replace("\r", "\n")
            paragraphs = [[(line, base)] if line else []
                        for line in text.split("\n")]
        else:
            return
        if paragraphs and all(isinstance(el, tuple) for el in paragraphs):
            # Images only: they go between paragraphs, just like when inserting
            # them
            cursor = self.textCursor()
            cursor.beginEditBlock()
            cursor.removeSelectedText()
            new = None
            for _, name, width, height in paragraphs:
                cf = self.image_format(name, width, height)
                if new is None:
                    new = self._place_image(cf)
                else:
                    new = self._insert_image_block(
                        new, True, self.image_block_format(), cf)
            cursor.endEditBlock()
            self.select_image(new)
            return
        cursor = self.textCursor()
        cursor.beginEditBlock()
        cursor.removeSelectedText()
        start_pos = cursor.position()
        had_break = self.is_manual_break(cursor.block())
        pending = False          # the paragraph still has to be split before the next one
        for el in paragraphs:
            if isinstance(el, tuple):             # image: its own paragraph
                cf = self.image_format(el[1], el[2], el[3])
                if cursor.atBlockStart():
                    bf0, cf0 = cursor.blockFormat(), cursor.block().charFormat()
                    cursor.insertText(OBJECT_CHAR, cf)
                    cursor.insertBlock(bf0, cf0)
                else:
                    cursor.insertBlock(self.image_block_format(), base)
                    cursor.insertText(OBJECT_CHAR, cf)
                    cursor.insertBlock(QTextBlockFormat(), base)
                pending = False
                continue
            if pending:
                cursor.insertBlock(QTextBlockFormat(), base)
            for text, fmt in el:
                cursor.insertText(text, fmt)
            pending = True
        end = cursor.position()
        cursor.setPosition(start_pos, QTextCursor.KeepAnchor)
        cursor.setBlockFormat(self.paragraph_format())   # removes the copied breaks
        # Image paragraphs get their format back
        doc = self.document()
        b = doc.findBlock(start_pos)
        while b.isValid() and b.position() <= end:
            if self._has_image_object(b):
                QTextCursor(b).setBlockFormat(self.image_block_format())
            b = b.next()
        if had_break:
            self._put_break(doc.findBlock(start_pos))
        cursor.endEditBlock()

    def _pasted_paragraphs(self, doc):
        """[[(text, format)] per paragraph] of the pasted document (tables
        become paragraphs and images are skipped)."""
        base = self.char_format()
        paragraphs = []
        b = doc.begin()
        while b.isValid():
            chunks = []
            it = b.begin()
            while not it.atEnd():
                f = it.fragment()
                if f.isValid():
                    text = f.text().replace("\ufffc", "")
                    if text:
                        chunks.append((text, self._pasted_format(f.charFormat(), base)))
                it += 1
            paragraphs.append(chunks)
            b = b.next()
        return paragraphs

    @staticmethod
    def _pasted_format(cf, base):
        """The Chapter body format with only the attributes that are pasted."""
        fmt = QTextCharFormat(base)
        if cf.fontWeight() >= QFont.DemiBold:
            fmt.setFontWeight(QFont.Bold)
        if cf.fontItalic():
            fmt.setFontItalic(True)
        if cf.fontUnderline() or \
                cf.underlineStyle() != QTextCharFormat.NoUnderline:
            fmt.setFontUnderline(True)
        if cf.fontStrikeOut():
            fmt.setFontStrikeOut(True)
        va = cf.verticalAlignment()
        if va in (QTextCharFormat.AlignSubScript, QTextCharFormat.AlignSuperScript):
            fmt.setVerticalAlignment(va)
        if cf.hasProperty(QTextFormat.ForegroundBrush) and \
                cf.foreground().style() != Qt.NoBrush:
            color = cf.foreground().color()
            # The black of web pages and word processors = automatic color
            if color.isValid() and color.alpha() > 0 and color.name() != "#000000":
                fmt.setForeground(color)
        if cf.hasProperty(QTextFormat.BackgroundBrush) and \
                cf.background().style() != Qt.NoBrush:
            color = cf.background().color()
            # A white or transparent background is not a highlight
            if color.isValid() and color.alpha() > 0 and color.name() != "#ffffff":
                fmt.setBackground(color)
        return fmt

    def merge_format(self, fmt):
        """Apply the format to the selection or to the text typed afterward
        (images in the selection don't change)."""
        cursor = self.textCursor()
        saved = self._images_in_selection(cursor)
        step = QTextCursor(self.document())
        step.beginEditBlock()
        self.mergeCurrentCharFormat(fmt)   # with a selection, it changes it too
        self._restore_images(saved)
        step.endEditBlock()

    def merge_paragraph(self, fmt):
        """Apply a paragraph format to all selected paragraphs (except image
        ones)."""
        cursor = self.textCursor()
        saved = self._images_in_selection(cursor)
        cursor.beginEditBlock()
        cursor.mergeBlockFormat(fmt)
        self._restore_images(saved)
        cursor.endEditBlock()
        self.paragraphFormatChanged.emit()

    def apply_style(self, style=CHAPTER_BODY):
        """Apply a style like the other tools: the character format to the
        selection (or to what is typed afterward) and the paragraph format
        to all selected paragraphs (or the current one)."""
        cursor = self.textCursor()
        cursor.beginEditBlock()  # a single undo step
        saved = self._images_in_selection(cursor)
        if cursor.hasSelection():
            # Replace (not merge) to remove bold, highlight, etc., but keep the
            # comments
            marks = self.comments.marks(cursor.selectionStart(),
                                             cursor.selectionEnd())
            cursor.setCharFormat(self.char_format(style))
            self.comments.restore(marks)
        self._set_block_format(cursor, self.paragraph_format(style))
        self._restore_images(saved)
        cursor.endEditBlock()
        if not cursor.hasSelection():
            # With a selection the format has already been applied; repeating
            # it here would add a second undo step.
            self.setCurrentCharFormat(self.char_format(style))
        self.paragraphFormatChanged.emit()

    def selected_blocks(self):
        """All paragraphs the selection touches (or the current paragraph)."""
        cursor = self.textCursor()
        doc = self.document()
        end = doc.findBlock(cursor.selectionEnd())
        b = doc.findBlock(cursor.selectionStart())
        blocks = []
        while b.isValid():
            blocks.append(b)
            if b == end:
                break
            b = b.next()
        return blocks


# ---------------------------------------------------------------------------
# Tool container: creates the actions and controls of a group; the ribbon
# (Ribbon) lays them out. It is never shown itself
# ---------------------------------------------------------------------------
class ToolContainer(QWidget):
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.hide()
        self.controls = []

    def addWidget(self, w):
        self.controls.append(w)
        return w

    def addSeparator(self):
        return None

    def setMovable(self, *_):
        pass

    def setIconSize(self, *_):
        pass

    def setToolButtonStyle(self, *_):
        pass


# ---------------------------------------------------------------------------
# Font tools
# ---------------------------------------------------------------------------
class FontBar(ToolContainer):
    def __init__(self, editor, parent=None):
        super().__init__("Font", parent)
        self.editor = editor
        self.setMovable(False)
        self.setIconSize(QSize(20, 20))
        self.setToolButtonStyle(Qt.ToolButtonIconOnly)

        # Font family
        self.font_combo = QFontComboBox()
        self.font_combo.setToolTip(_t("Font"))
        self.font_combo.setMinimumWidth(170)
        self.font_combo.setCurrentFont(QFont(CHAPTER_BODY["font_family"]))
        self.font_combo.currentFontChanged.connect(self.set_font)
        self.addWidget(self.font_combo)

        # Size
        self.size_combo = QComboBox()
        self.size_combo.setToolTip(_t("Font size"))
        self.size_combo.setEditable(True)
        self.size_combo.setFixedWidth(64)
        self.size_combo.setInsertPolicy(QComboBox.NoInsert)
        self.size_combo.lineEdit().setValidator(QIntValidator(1, 1638, self))
        self.size_combo.addItems([str(t) for t in SIZES])
        self.size_combo.setCurrentText(str(CHAPTER_BODY["font_size"]))
        self.size_combo.textActivated.connect(self.set_size)
        self.addWidget(self.size_combo)

        # Increase / decrease
        self.act_bigger = self._action(
            make_icon("A", small="+"), _t("Increase size") + " (Ctrl+Shift+>)",
            lambda: self.step_size(+1), "Ctrl+Shift+>")
        self.act_smaller = self._action(
            make_icon("A", small="−"), _t("Decrease size") + " (Ctrl+Shift+<)",
            lambda: self.step_size(-1), "Ctrl+Shift+<")
        self.addAction(self.act_bigger)
        self.addAction(self.act_smaller)

        # Clear formatting: leaves the text in the Body style (same as the Body
        # button in the style ribbon)
        self.act_clear = self._action(
            clear_format_icon(),
            _t("Clear formatting: back to the {style} style",
               style=style_name("body")) + " (Ctrl+1)",
            lambda: self.set_style(CHAPTER_BODY))
        self.addAction(self.act_clear)

        self.addSeparator()

        # Styles
        self.act_bold = self._action(
            make_icon(glyph("bold"), bold=True), _t("Bold") + " (Ctrl+B)",
            self.set_bold, QKeySequence.Bold, checkable=True)
        self.act_italic = self._action(
            make_icon(glyph("italic"), italic=True), _t("Italic") + " (Ctrl+I)",
            self.set_italic, QKeySequence.Italic, checkable=True)
        self.act_underline = self._action(
            make_icon(glyph("underline"), underline=True), _t("Underline") + " (Ctrl+U)",
            self.set_underline, QKeySequence.Underline, checkable=True)
        self.act_strike = self._action(
            make_icon("abc", strike=True), _t("Strikethrough") + " (Ctrl+Shift+X)",
            self.set_strikethrough, "Ctrl+Shift+X", checkable=True)
        for a in (self.act_bold, self.act_italic,
                  self.act_underline, self.act_strike):
            self.addAction(a)

        self.addSeparator()

        # Subscript / superscript (mutually exclusive)
        self.act_sub = self._action(
            make_icon("X", small="2", small_pos="down"), _t("Subscript") + " (Ctrl+=)",
            self.set_subscript, "Ctrl+=", checkable=True)
        self.act_sup = self._action(
            make_icon("X", small="2", small_pos="up"), _t("Superscript") + " (Ctrl+Shift+=)",
            self.set_superscript, "Ctrl+Shift+=", checkable=True)
        self.addAction(self.act_sub)
        self.addAction(self.act_sup)

        self.addSeparator()

        # Colors
        self.btn_color = ColorButton(
            "A", "#C00000", FONT_PALETTE, _t("Automatic"), CHAPTER_BODY["font_color"],
            _t("Font color"))
        self.btn_color.colorChosen.connect(self.set_color)
        self.addWidget(self.btn_color)

        self.btn_highlight = ColorButton(
            "ab", "#FFF59D", HIGHLIGHT_PALETTE, _t("No color"), None,
            _t("Highlight color"), highlight=True)
        self.btn_highlight.colorChosen.connect(self.set_highlight)
        self.addWidget(self.btn_highlight)

        self.addSeparator()

        # Alignment (exclusive: only one active)
        self.alignment_group = QActionGroup(self)
        self.alignment_group.setExclusive(True)
        self.alignment_actions = {}
        self.alignment_list = []
        for kind, align_value, tip, shortcut in (
            ("left", Qt.AlignLeft, _t("Align left") + " (Ctrl+L)", "Ctrl+L"),
            ("center", Qt.AlignHCenter, _t("Center") + " (Ctrl+E)", "Ctrl+E"),
            ("right", Qt.AlignRight, _t("Align right") + " (Ctrl+R)", "Ctrl+R"),
            ("justify", Qt.AlignJustify, _t("Justify") + " (Ctrl+J)", "Ctrl+J"),
        ):
            a = self._action(paragraph_icon(kind), tip,
                             lambda _=False, algn=align_value: self.set_alignment(algn),
                             shortcut, checkable=True)
            self.alignment_group.addAction(a)
            self.alignment_list.append(a)
            self.addAction(a)
            self.alignment_actions[align_value] = a

        self.addSeparator()

        # Left indent (limited to 0 … INDENT_MAX px)
        self.act_decrease = self._action(
            paragraph_icon("decrease"), _t("Decrease indent") + " (Ctrl+Shift+M)",
            lambda: self.step_indent(-1), "Ctrl+Shift+M")
        self.act_increase = self._action(
            paragraph_icon("increase"), _t("Increase indent") + " (Ctrl+M)",
            lambda: self.step_indent(+1), "Ctrl+M")
        self.addAction(self.act_decrease)
        self.addAction(self.act_increase)

        self.addSeparator()

        # Lists (turn the list on or off in the selected paragraphs)
        self.act_numbers = self._action(
            paragraph_icon("numbers"), _t("Numbered list") + " (Ctrl+Shift+L)",
            lambda: self.toggle_list(QTextListFormat.ListDecimal),
            "Ctrl+Shift+L", checkable=True)
        self.act_bullets = self._action(
            paragraph_icon("bullets"), _t("Bulleted list") + " (Ctrl+Shift+8)",
            lambda: self.toggle_list(QTextListFormat.ListDisc),
            "Ctrl+Shift+8", checkable=True)
        for a in (self.act_numbers, self.act_bullets):
            self.addAction(a)

        # Style buttons, in their own bar: they set the tools to the style's
        # values and apply it to the text. They are shown as a ribbon of
        # swatches: each button looks like its style, and the one of the
        # cursor's paragraph is marked
        self.style_bar = ToolContainer("Styles", parent)
        self.style_bar.setMovable(False)
        self.style_bar.setObjectName("styleRibbon")
        self.style_bar.setStyleSheet(
            "QToolBar#styleRibbon { spacing: 0px; padding: 6px 8px; }")
        self.swatches = {}             # key → (button, style)
        self._marked_style = None
        n = {c: style_name(c) for c in DEFAULT_STYLE_NAMES}
        self.act_body = self._style_button(
            "body", n["body"], CHAPTER_BODY, "Ctrl+1", "Ctrl+1")
        self.act_title = self._style_button(
            "title", n["title"], CHAPTER_TITLE, "Ctrl+2", "Ctrl+2",
            name=_t("Chapter title"))
        self.act_subtitle = self._style_button(
            "subtitle", n["subtitle"], CHAPTER_SUBTITLE, "Ctrl+3", "Ctrl+3",
            name=_t("Chapter subtitle"))
        self.act_dialogue = self._style_button(
            "dialogue", n["dialogue"], CHAPTER_BODY, "Ctrl+6", "Ctrl+6",
            name=n["body"], symbol=DASH, sample=f"{DASH}{n['dialogue']}")
        self.act_thoughts = self._style_button(
            "thoughts", n["thoughts"], CHAPTER_THOUGHTS, "Ctrl+5", "Ctrl+5")
        self.act_section = self._style_button(
            "section", n["section"], CHAPTER_SECTION, "Ctrl+4", "Ctrl+4",
            name=_t("Section heading"))
        self.act_quote = self._style_button(
            "quote", n["quote"], CHAPTER_QUOTE, "Ctrl+7", "Ctrl+7")
        self.act_separator = self._style_button(
            "scene_break", n["scene_break"], SCENE_BREAK, "Ctrl+8", "Ctrl+8",
            name=_t("Scene separator"), symbol=SEPARATOR)

        # Keep the bar in sync with the format under the cursor
        editor.currentCharFormatChanged.connect(self.sync)
        editor.cursorPositionChanged.connect(self.sync_paragraph)
        editor.selectionChanged.connect(self.sync_paragraph)
        editor.paragraphFormatChanged.connect(self.sync_paragraph)
        self.sync(editor.currentCharFormat())
        self.sync_paragraph()

    # -- utilities ----------------------------------------------------------
    def _action(self, icon, tooltip, slot, shortcut=None, checkable=False):
        a = QAction(icon, tooltip.split(" (")[0], self)
        a.setToolTip(tooltip)
        a.setCheckable(checkable)
        if shortcut:
            a.setShortcut(QKeySequence(shortcut))
            a.setShortcutContext(Qt.WindowShortcut)
        a.triggered.connect(slot)
        return a

    def _style_button(self, key, text, style, shortcut, shortcut_text,
                      name=None, symbol=None, sample=None):
        """Style swatch in the ribbon. With `symbol`, it also types it at the
        cursor."""
        # The descriptive name only if the style keeps its usual name (if it
        # was changed in settings.ini, the new one is used)
        if key in STYLE_NAMES:
            name = None if key != "dialogue" else name
        tooltip = _t('Sets the tools to "{name}"', name=name or text)
        if symbol:
            tooltip += _t(' and types "{symbol}"', symbol=symbol)
        tooltip += f" ({shortcut_text})"
        a = self._action(
            QIcon(), tooltip,
            lambda: self.set_style(style, symbol), shortcut)
        a.setText(text)
        # The shortcut works in the whole window; the button only triggers it
        window = self.parentWidget()
        (window if window is not None else self).addAction(a)
        b = QPushButton(sample or text)
        b.setToolTip(tooltip)
        b.setFocusPolicy(Qt.NoFocus)          # doesn't take the focus away from the text
        b.setCursor(Qt.PointingHandCursor)
        # Minimum width; in the ribbon they stretch to fill the space
        b.setMinimumWidth(SWATCH_WIDTH)
        b.setFixedHeight(SWATCH_HEIGHT)
        b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        b.clicked.connect(a.trigger)
        self.swatches[key] = (b, style)
        self._paint_swatch(key)
        self.style_bar.addWidget(b)
        return a

    def _paint_swatch(self, key):
        """Each swatch shows the font, size, weight, italics and color of its
        style (the size is reduced so it fits in the ribbon)."""
        b, style = self.swatches[key]
        marked = key == self._marked_style
        k = STYLE_ORDER.index(key) if key in STYLE_ORDER else 0
        row, col = divmod(k, 4)
        size = max(12, min(20, round(style["font_size"] * 1.05)))
        border = "2px solid #7A7A7A" if marked else "1px solid #CFCFCF"
        # Shared borders: each swatch draws its left and top border only if it
        # is in the first column or row, or is marked (otherwise its neighbor
        # draws them)
        left_border = border if (col == 0 or marked) else "none"
        top = border if (row == 0 or marked) else "none"
        family = style["font_family"].replace('"', "")
        b.setStyleSheet(
            "QPushButton {"
            f" font-family: \"{family}\"; font-size: {size}px;"
            f" font-weight: {'bold' if style['bold'] else 'normal'};"
            f" font-style: {'italic' if style['italic'] else 'normal'};"
            f" color: {style['font_color']};"
            " background: #FFFFFF; text-align: left; padding: 0 10px;"
            f" border: {border}; border-left: {left_border}; border-top: {top};"
            " border-radius: 0; }"
            "QPushButton:hover { background: #EEF4FC; }"
            "QPushButton:pressed { background: #DCE8F7; }")

    def mark_style(self, key):
        if key == self._marked_style:
            return
        self._marked_style = key
        for c in self.swatches:              # the borders depend on the neighbors
            self._paint_swatch(c)

    def paragraph_style(self):
        """Which style the cursor's paragraph has (None if none matches)."""
        cursor = self.editor.textCursor()
        block = cursor.block()
        bf = block.blockFormat()
        level = bf.headingLevel()
        if level:
            return {1: "title", 2: "subtitle"}.get(level, "section")
        if block.length() > 1:               # format of the first character
            c = QTextCursor(block)
            c.movePosition(QTextCursor.NextCharacter)
            cf = c.charFormat()
        else:
            cf = self.editor.currentCharFormat()
        text = block.text().strip()
        if text == SEPARATOR or (bf.alignment() & Qt.AlignHCenter
                                  and round(bf.topMargin()) == SCENE_BREAK["space_before"]):
            return "scene_break"
        if round(bf.leftMargin()) == CHAPTER_QUOTE["indent_left"] and \
                round(bf.rightMargin()) == CHAPTER_QUOTE["indent_right"] and \
                bf.leftMargin() > 0:
            return "quote"
        if text.startswith(DASH):
            return "dialogue"
        if cf.fontItalic():
            return "thoughts"
        return "body"

    def _apply(self, fmt):
        self.editor.merge_format(fmt)
        self.editor.setFocus()

    def _current_size(self):
        fmt = self.editor.currentCharFormat()
        t = fmt.fontPointSize() or fmt.font().pointSizeF()
        return t if t > 0 else CHAPTER_BODY["font_size"]

    # -- actions ------------------------------------------------------------
    def set_font(self, font):
        fmt = QTextCharFormat()
        fmt.setFontFamilies([font.family()])
        self._apply(fmt)

    def set_size(self, text):
        try:
            t = float(text)
        except ValueError:
            return
        if t > 0:
            fmt = QTextCharFormat()
            fmt.setFontPointSize(t)
            self._apply(fmt)

    def step_size(self, direction):
        current = self._current_size()
        if direction > 0:
            new = next((t for t in SIZES if t > current), current + 10)
        else:
            smaller = [t for t in SIZES if t < current]
            new = smaller[-1] if smaller else max(1, current - 1)
        self.set_size(str(new))

    def set_bold(self, checked):
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Bold if checked else QFont.Normal)
        self._apply(fmt)

    def set_italic(self, checked):
        fmt = QTextCharFormat()
        fmt.setFontItalic(checked)
        self._apply(fmt)

    def set_underline(self, checked):
        fmt = QTextCharFormat()
        fmt.setFontUnderline(checked)
        self._apply(fmt)

    def set_strikethrough(self, checked):
        fmt = QTextCharFormat()
        fmt.setFontStrikeOut(checked)
        self._apply(fmt)

    def _vertical(self, alignment, checked):
        fmt = QTextCharFormat()
        fmt.setVerticalAlignment(
            alignment if checked else QTextCharFormat.AlignNormal)
        self._apply(fmt)

    def set_subscript(self, checked):
        self._vertical(QTextCharFormat.AlignSubScript, checked)

    def set_superscript(self, checked):
        self._vertical(QTextCharFormat.AlignSuperScript, checked)

    def set_color(self, color):
        fmt = QTextCharFormat()
        fmt.setForeground(color if color.isValid() else QColor(CHAPTER_BODY["font_color"]))
        self._apply(fmt)

    def set_highlight(self, color):
        fmt = QTextCharFormat()
        fmt.setBackground(QBrush(color) if color.isValid() else QBrush(Qt.NoBrush))
        self._apply(fmt)

    def set_style(self, style, symbol=None):
        """Apply the style to the text and, if given, type the symbol at the
        cursor (all in a single undo step). The tools take their values when
        they sync with the text; the color buttons are set here because they
        show the last chosen color."""
        self.btn_color.set_color_value(style["font_color"])
        self.btn_highlight.set_color_value(QColor())  # no highlight
        doc = QTextCursor(self.editor.document())
        doc.beginEditBlock()
        self.editor.apply_style(style)
        if symbol:
            cursor = self.editor.textCursor()
            cursor.setPosition(cursor.selectionStart())  # doesn't delete the selection
            cursor.insertText(symbol, self.editor.char_format(style))
            self.editor.setTextCursor(cursor)
        doc.endEditBlock()
        self.editor.setFocus()

    # -- paragraph ----------------------------------------------------------
    def set_alignment(self, alignment):
        fmt = QTextBlockFormat()
        fmt.setAlignment(alignment)
        self.editor.merge_paragraph(fmt)  # all selected paragraphs
        self.editor.setFocus()

    def step_indent(self, step):
        """Move the left indent to the next or previous multiple of
        INDENT_STEP, staying within 0 … INDENT_MAX."""
        cursor = QTextCursor(self.editor.document())
        cursor.beginEditBlock()  # a single undo step
        for b in self.editor.selected_blocks():
            if is_image(b):              # images always centered
                continue
            current = b.blockFormat().leftMargin()
            if step > 0:
                new = (int(current // INDENT_STEP) + 1) * INDENT_STEP
            else:
                new = (int(-(-current // INDENT_STEP)) - 1) * INDENT_STEP
            new = max(0, min(INDENT_MAX, new))
            cursor.setPosition(b.position())
            fmt = QTextBlockFormat()
            fmt.setLeftMargin(new)
            cursor.mergeBlockFormat(fmt)
        cursor.endEditBlock()
        self.editor.paragraphFormatChanged.emit()
        self.editor.setFocus()

    def step_space(self, before, step):
        """Move the space before (or after) the selected paragraphs to the next
        or previous multiple of SPACE_STEP, staying within 0 … SPACE_MAX.
        Images don't change."""
        cursor = QTextCursor(self.editor.document())
        cursor.beginEditBlock()  # a single undo step
        for b in self.editor.selected_blocks():
            if is_image(b):
                continue
            bf = b.blockFormat()
            current = bf.topMargin() if before else bf.bottomMargin()
            if step > 0:
                new = (int(current // SPACE_STEP) + 1) * SPACE_STEP
            else:
                new = (int(-(-current // SPACE_STEP)) - 1) * SPACE_STEP
            new = max(0, min(SPACE_MAX, new))
            fmt = QTextBlockFormat()
            if before:
                fmt.setTopMargin(new)
            else:
                fmt.setBottomMargin(new)
            cursor.setPosition(b.position())
            cursor.mergeBlockFormat(fmt)
        cursor.endEditBlock()
        self.editor.paragraphFormatChanged.emit()
        self.editor.setFocus()

    def toggle_list(self, style):
        """Put the selected paragraphs in a list of the requested style. If
        they are already in a list of that style, remove it (as in Word).
        Images and headings don't go into the list."""
        editor = self.editor
        cursor = editor.textCursor()
        blocks = [b for b in editor.selected_blocks()
                   if not is_image(b) and b.blockFormat().headingLevel() == 0]
        if not blocks:
            QApplication.beep()
            return
        current = cursor.currentList()
        numbered = style == QTextListFormat.ListDecimal
        same = current is not None and \
            (current.format().style() == QTextListFormat.ListDecimal) == numbered
        cursor.beginEditBlock()
        if same:                             # remove the list
            for b in blocks:
                lst = b.textList()
                if lst is not None:
                    lst.remove(b)
                c = QTextCursor(b)
                bf = c.blockFormat()
                bf.setIndent(0)
                c.setBlockFormat(bf)
        else:
            fmt = QTextListFormat()
            fmt.setStyle(style)
            fmt.setIndent(1)
            if numbered:
                fmt.setNumberSuffix(".")
            # Lists that already exist change type as a whole (all their items,
            # at the same level), as in Word
            touched = []
            for b in blocks:
                lst = b.textList()
                if lst is not None and lst not in touched:
                    touched.append(lst)
            for lst in touched:
                new = QTextListFormat(fmt)
                new.setIndent(lst.format().indent() or 1)
                lst.setFormat(new)
            # Loose paragraphs join the list above (if it is the same type),
            # the first one touched, or a new one
            target = touched[0] if touched else None
            if target is None:
                previous = blocks[0].previous()
                top = previous.textList() if previous.isValid() else None
                if top is not None and \
                        (top.format().style() == QTextListFormat.ListDecimal) == numbered:
                    target = top
            for b in blocks:
                if b.textList() is not None:
                    continue
                if target is None:
                    target = QTextCursor(b).createList(fmt)
                else:
                    target.add(b)
        cursor.endEditBlock()
        editor.paragraphFormatChanged.emit()
        editor.setFocus()

    def sync_paragraph(self):
        self.mark_style(self.paragraph_style())
        blocks = self.editor.selected_blocks()
        align_value = self.editor.textCursor().blockFormat().alignment() & Qt.AlignHorizontal_Mask
        if align_value == Qt.AlignLeading:
            align_value = Qt.AlignLeft
        elif align_value == Qt.AlignTrailing:
            align_value = Qt.AlignRight
        action = self.alignment_actions.get(align_value)
        if action:
            action.setChecked(True)

        margins = [b.blockFormat().leftMargin() for b in blocks]
        # The button is disabled when no paragraph can move any further
        self.act_decrease.setEnabled(any(m > 0 for m in margins))
        self.act_increase.setEnabled(any(m < INDENT_MAX for m in margins))
        text = [b.blockFormat() for b in blocks if not is_image(b)]
        lst = self.editor.textCursor().currentList()
        style = lst.format().style() if lst is not None else None
        self.act_numbers.setChecked(style == QTextListFormat.ListDecimal)
        self.act_bullets.setChecked(style in (QTextListFormat.ListDisc,
                                               QTextListFormat.ListCircle,
                                               QTextListFormat.ListSquare))

    # -- state --------------------------------------------------------------
    def sync(self, fmt):
        f = fmt.font()

        self.font_combo.blockSignals(True)
        families = fmt.fontFamilies()
        family = families[0] if families else f.family()
        self.font_combo.setCurrentFont(QFont(family))
        self.font_combo.blockSignals(False)

        t = fmt.fontPointSize() or f.pointSizeF()
        self.size_combo.blockSignals(True)
        self.size_combo.setCurrentText(f"{t:g}" if t > 0 else "")
        self.size_combo.blockSignals(False)

        self.act_bold.setChecked(f.bold())
        self.act_italic.setChecked(f.italic())
        self.act_underline.setChecked(fmt.fontUnderline())
        self.act_strike.setChecked(fmt.fontStrikeOut())

        va = fmt.verticalAlignment()
        self.act_sub.setChecked(va == QTextCharFormat.AlignSubScript)
        self.act_sup.setChecked(va == QTextCharFormat.AlignSuperScript)


# ---------------------------------------------------------------------------
# Paragraph toolbar (spacing and indents)
# ---------------------------------------------------------------------------
def pt_to_px(editor, pt):
    """Points → editor px (Qt's absolute spacing is in px)."""
    return pt * editor.logicalDpiY() / 72


def letter_spacing_pt(editor, cf):
    """Letter spacing of a format, in points (0 = normal)."""
    if not cf.hasProperty(QTextFormat.FontLetterSpacing):
        return 0.0
    value = cf.fontLetterSpacing()
    if cf.fontLetterSpacingType() == QFont.AbsoluteSpacing:
        return round(value * 72 / editor.logicalDpiY(), 1)
    # Percentage (100 = normal): converted according to the font size
    size = cf.fontPointSize() or CHAPTER_BODY["font_size"]
    return round((value - 100) / 100 * size / 2, 1) if value else 0.0


class ParagraphBar(ToolContainer):
    def __init__(self, editor, parent=None):
        super().__init__("Paragraph", parent)
        self.editor = editor
        self.setMovable(False)

        # -- Spacing --------------------------------------------------------
        self._heading_label(_t("Spacing"))

        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.setEditable(True)
        self.line_spacing_combo.setInsertPolicy(QComboBox.NoInsert)
        self.line_spacing_combo.setFixedWidth(68)
        self.line_spacing_combo.setToolTip(
            _t("Line spacing (multiple of the line height, 0.5 to 5)"))
        validator = QDoubleValidator(0.5, 5.0, 2, self)
        validator.setNotation(QDoubleValidator.StandardNotation)
        validator.setLocale(QLocale.c())  # always a decimal point
        self.line_spacing_combo.lineEdit().setValidator(validator)
        self.line_spacing_combo.addItems(LINE_SPACINGS)
        self.line_spacing_combo.textActivated.connect(self.set_line_spacing)
        self._label(_t("Line spacing"))
        self.addWidget(self.line_spacing_combo)

        self.spin_before = self._spin(
            _t("Before"), _t("Space before paragraph"), 0, SPACE_MAX, 6,
            self.set_space_before)
        self.spin_after = self._spin(
            _t("After"), _t("Space after paragraph"), 0, SPACE_MAX, 6,
            self.set_space_after)

        # Letter spacing: it is a character format (applied to the selection or
        # to what is typed), in points as in Word/Writer
        self._label(_t("Letters"))
        self.spin_letters = QDoubleSpinBox()
        self.spin_letters.setRange(LETTERS_MIN, LETTERS_MAX)
        self.spin_letters.setSingleStep(LETTERS_STEP)
        self.spin_letters.setDecimals(1)
        self.spin_letters.setSuffix(" pt")
        self.spin_letters.setLocale(QLocale.c())     # always a decimal point
        self.spin_letters.setFixedWidth(76)
        self.spin_letters.setAccelerated(True)
        self.spin_letters.setKeyboardTracking(False)
        self.spin_letters.setToolTip(
            _t("Letter spacing: positive spreads, negative tightens "
               "({min} to {max} pt; 0 = normal)",
               min=f"{LETTERS_MIN:g}", max=f"{LETTERS_MAX:g}"))
        self.spin_letters.valueChanged.connect(self.set_letter_spacing)
        self.addWidget(self.spin_letters)

        self.addSeparator()

        # -- Indents --------------------------------------------------------
        self._heading_label(_t("Indent"))
        self.spin_left_indent = self._spin(
            _t("indent|Left"), _t("Left indent"), 0, INDENT_MAX, INDENT_STEP,
            self.set_left_indent)
        self.spin_right_indent = self._spin(
            _t("indent|Right"), _t("Right indent"), 0, INDENT_MAX, INDENT_STEP,
            self.set_right_indent)
        self.spin_first_line = self._spin(
            _t("First line"), _t("First-line indent"), 0, FIRST_LINE_MAX, 6,
            self.set_first_line)

        editor.cursorPositionChanged.connect(self.sync)
        editor.selectionChanged.connect(self.sync)
        editor.paragraphFormatChanged.connect(self.sync)
        editor.currentCharFormatChanged.connect(self.sync_letters)
        self.sync()

    # -- widgets ------------------------------------------------------------
    def _heading_label(self, text):
        lbl = QLabel(text)
        lbl.setObjectName("groupTitle")
        self.addWidget(lbl)

    def _label(self, text):
        lbl = QLabel(text)
        lbl.setObjectName("fieldLabel")
        self.addWidget(lbl)

    def _spin(self, tag, tooltip, minimum, maximum, step, slot):
        self._label(tag)
        s = QSpinBox()
        s.setRange(minimum, maximum)
        s.setSingleStep(step)
        s.setSuffix(" px")
        s.setFixedWidth(76)
        s.setAccelerated(True)
        s.setKeyboardTracking(False)  # applied on Enter or when leaving the field
        s.setToolTip(f"{tooltip} ({minimum}–{maximum} px)")
        s.valueChanged.connect(slot)
        self.addWidget(s)
        return s

    # -- actions ------------------------------------------------------------
    def _apply(self, fmt):
        self.editor.merge_paragraph(fmt)

    def set_line_spacing(self, text):
        try:
            value = float(text.replace(",", "."))
        except ValueError:
            return
        value = max(0.5, min(5.0, value))
        fmt = QTextBlockFormat()
        fmt.setLineHeight(round(value * 100),
                          QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
        self._apply(fmt)
        self.editor.setFocus()

    def set_space_before(self, v):
        fmt = QTextBlockFormat()
        fmt.setTopMargin(v)
        self._apply(fmt)

    def set_space_after(self, v):
        fmt = QTextBlockFormat()
        fmt.setBottomMargin(v)
        self._apply(fmt)

    def set_letter_spacing(self, pt):
        fmt = QTextCharFormat()
        fmt.setFontLetterSpacingType(QFont.AbsoluteSpacing)
        fmt.setFontLetterSpacing(pt_to_px(self.editor, pt))
        self.editor.merge_format(fmt)

    def set_left_indent(self, v):
        fmt = QTextBlockFormat()
        fmt.setLeftMargin(v)
        self._apply(fmt)

    def set_right_indent(self, v):
        fmt = QTextBlockFormat()
        fmt.setRightMargin(v)
        self._apply(fmt)

    def set_first_line(self, v):
        fmt = QTextBlockFormat()
        fmt.setTextIndent(v)
        self._apply(fmt)

    # -- state --------------------------------------------------------------
    def sync(self):
        fmt = self.editor.textCursor().blockFormat()

        if fmt.lineHeightType() == \
                QTextBlockFormat.LineHeightTypes.ProportionalHeight.value:
            spacing = fmt.lineHeight() / 100
        else:
            spacing = 1.0
        self.line_spacing_combo.blockSignals(True)
        text = f"{spacing:.2f}".rstrip("0")
        self.line_spacing_combo.setCurrentText(text + "0" if text.endswith(".") else text)
        self.line_spacing_combo.blockSignals(False)

        for spin, value in (
            (self.spin_before, fmt.topMargin()),
            (self.spin_after, fmt.bottomMargin()),
            (self.spin_left_indent, fmt.leftMargin()),
            (self.spin_right_indent, fmt.rightMargin()),
            (self.spin_first_line, fmt.textIndent()),
        ):
            spin.blockSignals(True)
            spin.setValue(round(value))
            spin.blockSignals(False)
        self.sync_letters()

    def sync_letters(self, *_):
        pt = letter_spacing_pt(self.editor, self.editor.currentCharFormat())
        self.spin_letters.blockSignals(True)
        self.spin_letters.setValue(pt)
        self.spin_letters.blockSignals(False)


# ---------------------------------------------------------------------------
# Page view
# ---------------------------------------------------------------------------
class PageView(QGraphicsView):
    """Show the editor as sheets on a gray background, with zoom.
    #
    The document is paginated with the sheet height plus the gap; the top
    and bottom margins include half the gap, so each visible sheet keeps its
    exact size and margins.
    #
    In the "No pages" view the document is a single continuous sheet of the
    same width (lines break the same as when printing), with no cuts between
    pages; manual page breaks are shown as a line.
    #
    The "Mixed" view is also continuous, but each page break splits the
    sheet: the result is sheets of different lengths, each as long as its
    content. The text is laid out as in No pages and the editor draws each
    segment shifted onto its sheet (see Editor.paintEvent)."""
    zoomChanged = Signal(float)
    pageChanged = Signal(int, int)   # current page, total (0, 0 = no pages)
    viewChanged = Signal(str)         # VIEW_PAGES / VIEW_NO_PAGES / VIEW_MIXED

    def __init__(self, editor, parent=None):
        super().__init__(parent)
        self.setObjectName("canvas")
        self.setScene(QGraphicsScene(self))
        self.setBackgroundBrush(QColor("#E4E4E4"))
        self.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.setFrameShape(QFrame.NoFrame)
        self.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing
                            | QPainter.SmoothPixmapTransform)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)

        self.editor = editor
        editor.view = self
        editor.setFrameShape(QFrame.NoFrame)
        editor.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        editor.setStyleSheet("QTextEdit { background: transparent; }")
        editor.viewport().setAutoFillBackground(False)

        self.proxy = self.scene().addWidget(editor)
        self.proxy.setZValue(1)            # the text above the sheets
        self.sheets = []
        self.zoom = 1.0
        self.config = dict(PAGE)
        self.mode = VIEW_PAGES
        self._page = (0, 0)
        self._mixed_pending = False

        # Auto-scroll when selecting with the mouse
        self._auto = QTimer(self)
        self._auto.setInterval(AUTO_INTERVAL)
        self._auto.timeout.connect(self._auto_scroll_step)
        self._auto_remainder = [0.0, 0.0]   # accumulated pixel fractions

        editor.document().documentLayout().documentSizeChanged.connect(
            self._update_pages)
        editor.cursorPositionChanged.connect(
            lambda: QTimer.singleShot(0, self._follow_cursor))
        editor.textChanged.connect(
            lambda: QTimer.singleShot(0, self._follow_cursor))
        # In the mixed view, adding or removing a break doesn't always change
        # the document height: the cuts are checked after every change
        editor.document().contentsChanged.connect(self._schedule_mixed)
        # Undo/redo can change the root frame margins: the view catches up with
        # whatever the document says
        self._sync_pending = False
        editor.document().contentsChange.connect(self._schedule_sync)
        editor.document().contentsChanged.connect(self._schedule_sync)
        self.apply_config(self.config)

    # -- view mode ------------------------------------------------------------
    @property
    def no_pages(self):
        return self.mode == VIEW_NO_PAGES

    @property
    def mixed(self):
        return self.mode == VIEW_MIXED

    @property
    def continuous(self):
        """No pages and Mixed lay out the text without automatic cuts."""
        return self.mode in (VIEW_NO_PAGES, VIEW_MIXED)

    # -- measurements -------------------------------------------------------
    def cm_to_px(self, cm):
        # Same DPI the editor uses to draw fonts (pt → px)
        return cm / 2.54 * self.editor.logicalDpiY()

    def _measurements(self):
        c = self.config
        return (self.cm_to_px(c["width"]), self.cm_to_px(c["height"]))

    def apply_config(self, config):
        self.config = dict(config)
        width, height = self._measurements()
        g = PAGE_GAP
        doc = self.editor.document()
        root = doc.rootFrame()
        ff = root.frameFormat()
        ff.setTopMargin(self.cm_to_px(config["top"]) + g / 2)
        ff.setBottomMargin(self.cm_to_px(config["bottom"]) + g / 2)
        ff.setLeftMargin(self.cm_to_px(config["left"]))
        ff.setRightMargin(self.cm_to_px(config["right"]))
        ff.setProperty(PROP_PAGE, ";".join(
            repr(float(config[k])) for k in PAGE_KEYS))
        # Changing the frame format goes into the undo history: only do it if
        # something really changes. Images that no longer fit are adjusted in
        # the same step
        step = QTextCursor(doc)
        step.beginEditBlock()
        if ff != root.frameFormat():
            root.setFrameFormat(ff)
        self.editor.fit_images()
        step.endEditBlock()
        self._set_document_size()

    def _document_config(self):
        """The page setup stored in the root frame, or None."""
        value = self.editor.document().rootFrame().frameFormat().property(PROP_PAGE)
        try:
            numbers = [float(x) for x in str(value).split(";")]
        except (TypeError, ValueError):
            return None
        if len(numbers) != len(PAGE_KEYS):
            return None
        return dict(zip(PAGE_KEYS, numbers))

    def _schedule_sync(self, *_):
        if not self._sync_pending:
            self._sync_pending = True
            QTimer.singleShot(0, self._sync_config)

    def _sync_config(self):
        """After undoing or redoing a page change, adopt the document's setup
        (without touching the document: that would erase what is left to
        redo)."""
        self._sync_pending = False
        config = self._document_config()
        if config is None or config == self.config:
            return
        self.config = config
        self._set_document_size()

    def _set_document_size(self):
        width, height = self._measurements()
        # Height -1 = no limit: Qt neither cuts pages nor applies breaks
        self.editor.set_page_size(
            QSizeF(width, -1 if self.continuous else height + PAGE_GAP))
        self._update_pages(force=True)
        self._page = (-1, -1)
        self._follow_cursor()

    def set_mode(self, mode):
        """Change the view (Pages, No pages or Mixed). It only changes how it
        is shown: the text and its format are not touched (so it leaves
        nothing to undo)."""
        if mode not in (VIEW_PAGES, VIEW_NO_PAGES, VIEW_MIXED):
            mode = VIEW_PAGES
        if mode == self.mode:
            return
        self.mode = mode
        self.editor.set_cuts([], 0)
        self._set_document_size()
        self.editor.viewport().update()
        self.viewChanged.emit(mode)

    # -- sheets -------------------------------------------------------------
    def _update_pages(self, *_, force=False):
        self.editor.ensure_pagination()
        if self.mixed:
            self._update_mixed_sheets()
            return
        if self.no_pages:
            self._update_continuous_sheet(force)
            return
        width, height = self._measurements()
        g = PAGE_GAP
        step = height + g
        n = max(1, self.editor.document().pageCount())

        # Each sheet is two items (shadow and sheet). Only the sheets that
        # change are added or removed: rebuilding them all every time is very
        # slow in long books
        if force:
            for h in self.sheets:
                self.scene().removeItem(h)
            self.sheets = []
        existing = len(self.sheets) // 2
        if existing != n or force:
            while len(self.sheets) // 2 > n:
                for h in self.sheets[-2:]:
                    self.scene().removeItem(h)
                del self.sheets[-2:]
            for i in range(existing, n):
                y = i * step + g / 2
                shadow = self.scene().addRect(
                    QRectF(2, y + 3, width, height), Qt.NoPen, QColor(0, 0, 0, 40))
                sheet = self.scene().addRect(
                    QRectF(0, y, width, height),
                    QPen(QColor("#C8C8C8"), 0), QColor("#FFFFFF"))
                self.sheets += [shadow, sheet]
            self.pageChanged.emit(*self._compute_page())

        total = round(n * step)
        if self.editor.width() != round(width) or self.editor.height() != total:
            self.editor.setFixedSize(round(width), total)
        self.scene().setSceneRect(QRectF(-g, -g / 2, width + 2 * g, total + g))

    def _update_continuous_sheet(self, force=False):
        """No pages view: a single sheet that grows with the text (at least one
        page tall) and always reaches the bottom of the view: its bottom
        edge is never seen, not even at the end of the text."""
        width, height = self._measurements()
        g = PAGE_GAP
        content = self.editor.document().documentLayout().documentSize().height()
        total = round(max(content, height + g))
        # Visible height, in scene units (1 px less so the scroll bar doesn't
        # flicker on and off)
        visible = (self.viewport().height() - 1) / self.zoom
        background = max(total - g / 2, visible - g / 2)   # where the scene ends
        if force or len(self.sheets) != 2:
            for h in self.sheets:
                self.scene().removeItem(h)
            self.sheets = [
                self.scene().addRect(QRectF(), Qt.NoPen, QColor(0, 0, 0, 40)),
                self.scene().addRect(QRectF(), QPen(QColor("#C8C8C8"), 0),
                                     QColor("#FFFFFF")),
            ]
        shadow, sheet = self.sheets
        # The sheet goes a little past the end of the scene: its bottom edge
        # stays out of view
        sheet_height = background - g / 2 + 4 * g
        shadow.setRect(QRectF(2, g / 2 + 3, width, sheet_height))
        sheet.setRect(QRectF(0, g / 2, width, sheet_height))
        if self.editor.width() != round(width) or self.editor.height() != total:
            self.editor.setFixedSize(round(width), total)
        self.scene().setSceneRect(QRectF(-g, -g / 2, width + 2 * g, background + g / 2))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if self.no_pages:          # the sheet reaches the bottom of the view
            self._update_continuous_sheet()

    def _schedule_mixed(self):
        if self.mixed and not self._mixed_pending:
            self._mixed_pending = True
            QTimer.singleShot(0, self._update_mixed_sheets)

    def _set_sheets(self, rects):
        """Put a sheet (with a shadow) in each rectangle, reusing the existing
        ones."""
        while len(self.sheets) // 2 > len(rects):
            for h in self.sheets[-2:]:
                self.scene().removeItem(h)
            del self.sheets[-2:]
        while len(self.sheets) // 2 < len(rects):
            self.sheets += [
                self.scene().addRect(QRectF(), Qt.NoPen, QColor(0, 0, 0, 40)),
                self.scene().addRect(QRectF(), QPen(QColor("#C8C8C8"), 0),
                                     QColor("#FFFFFF")),
            ]
        for i, r in enumerate(rects):
            self.sheets[2 * i].setRect(r.translated(2, 3))
            self.sheets[2 * i + 1].setRect(r)

    def _update_mixed_sheets(self):
        """Mixed view: one sheet per segment between page breaks. Each break
        opens a gap (bottom margin + gap + top margin) that the editor
        respects when drawing."""
        self._mixed_pending = False
        if not self.mixed:
            return
        width, height = self._measurements()
        g = PAGE_GAP
        top_margin = self.cm_to_px(self.config["top"])
        bottom_margin = self.cm_to_px(self.config["bottom"])
        doc = self.editor.document()
        layout = doc.documentLayout()
        cuts, previous_block = [], None
        b = doc.begin()
        while b.isValid():
            if b.isVisible():
                if previous_block is not None and \
                        b.blockFormat().pageBreakPolicy() & QTextFormat.PageBreak_AlwaysBefore:
                    cuts.append(layout.blockBoundingRect(previous_block).bottom())
                previous_block = b
            b = b.next()
        gap = bottom_margin + g + top_margin
        content = layout.documentSize().height()
        rects = []
        starts = [None] + cuts
        ends = cuts + [None]
        for k, (start, end) in enumerate(zip(starts, ends)):
            shift = k * gap
            top = g / 2 if start is None else start + shift - top_margin
            down = content + shift - g / 2 if end is None else end + shift + bottom_margin
            if end is None:                     # the last one, at least one page
                down = max(down, top + height)
            rects.append(QRectF(0, top, width, down - top))
        self._set_sheets(rects)
        total = round(rects[-1].bottom() + g / 2)
        self.editor.set_cuts(cuts, gap)
        if self.editor.width() != round(width) or self.editor.height() != total:
            self.editor.setFixedSize(round(width), total)
        self.scene().setSceneRect(QRectF(-g, -g / 2, width + 2 * g, total + g))
        page = self._compute_page()
        if page != self._page:
            self._page = page
            self.pageChanged.emit(*page)

    def _compute_page(self):
        if self.no_pages:
            return 0, 0
        if self.mixed:                  # the cursor's sheet, out of however many there are
            y = QTextEdit.cursorRect(self.editor).center().y()
            cuts = self.editor._cuts
            return bisect.bisect_right(cuts, y) + 1, len(cuts) + 1
        _, height = self._measurements()
        step = height + PAGE_GAP
        y = self.editor.cursorRect().center().y()
        total = max(1, self.editor.document().pageCount())
        return min(total, int(y // step) + 1), total

    def _follow_cursor(self):
        """Scroll the view so the cursor is visible, as when typing in a word
        processor."""
        # While selecting with the mouse, scrolling is controlled by the
        # auto-scroll; following the cursor here creates a loop (scroll → the
        # selection grows → the cursor moves down → scroll…)
        if not QApplication.mouseButtons() & Qt.LeftButton:
            r = QRectF(self.editor.cursorRect())
            self.ensureVisible(self.proxy.mapRectToScene(r), 40, 60)
        page = self._compute_page()
        if page != self._page:
            self._page = page
            self.pageChanged.emit(*page)

    def scroll_screen(self, down, extend):
        cursor = self.editor.textCursor()
        r = self.editor.cursorRect(cursor)
        step = self.viewport().height() / self.zoom * 0.9
        y = r.center().y() + (step if down else -step)
        target = self.editor.cursorForPosition(QPoint(r.center().x(), round(y)))
        cursor.setPosition(target.position(),
                           QTextCursor.KeepAnchor if extend else QTextCursor.MoveAnchor)
        bar = self.verticalScrollBar()
        bar.setValue(bar.value() + round(step * self.zoom) * (1 if down else -1))
        self.editor.setTextCursor(cursor)

    # -- auto-scroll ----------------------------------------------------------
    def start_auto_scroll(self):
        if not self._auto.isActive():
            self._auto_remainder = [0.0, 0.0]
            self._auto.start()

    def stop_auto_scroll(self):
        self._auto.stop()

    @staticmethod
    def _speed(pos, start_pos, end):
        """Pixels per step depending on how far the cursor is from the edge
        zone: slow right on arrival, faster and faster as it moves away."""
        if pos < start_pos + AUTO_ZONE:
            d = start_pos + AUTO_ZONE - pos
            sign = -1
        elif pos > end - AUTO_ZONE:
            d = pos - (end - AUTO_ZONE)
            sign = 1
        else:
            return 0.0
        # Smooth curve (at ~60 steps/s): entering the zone ≈ 20 px/s, at the
        # edge ≈ 110 px/s, 20 px outside ≈ 300 px/s, 80 px outside ≈ 1 300 px/s
        return sign * min(AUTO_MAX_SPEED, 0.3 + (d / 15) ** 1.6)

    def _auto_scroll_step(self):
        if not QApplication.mouseButtons() & Qt.LeftButton:
            self._auto.stop()       # the button was released (in case the event didn't arrive)
            return
        vp = self.viewport()
        p = vp.mapFromGlobal(QCursor.pos())
        vx = self._speed(p.x(), 0, vp.width())
        vy = self._speed(p.y(), 0, vp.height())
        if not vx and not vy:
            self._auto.stop()       # resumes with the next movement
            return
        moved = False
        for i, (bar, v) in enumerate(((self.horizontalScrollBar(), vx),
                                        (self.verticalScrollBar(), vy))):
            self._auto_remainder[i] += v
            step = int(self._auto_remainder[i])
            if step:
                self._auto_remainder[i] -= step
                before = bar.value()
                bar.setValue(before + step)
                moved |= bar.value() != before
        if moved:
            # Extend the selection to where the cursor ended up
            pos = QPointF(p)
            QApplication.sendEvent(vp, QMouseEvent(
                QEvent.MouseMove, pos, vp.mapToGlobal(pos),
                Qt.NoButton, Qt.LeftButton, QApplication.keyboardModifiers()))

    # -- zoom ---------------------------------------------------------------
    def set_zoom(self, z, toward="nearest"):
        z = zoom_level(z, toward)            # only levels that don't distort
        if abs(z - self.zoom) < 1e-4:
            return
        self.zoom = z
        self.setTransform(QTransform.fromScale(z, z))
        if self.no_pages:
            self._update_continuous_sheet()
        self.zoomChanged.emit(z)

    def step_zoom(self, steps):
        """Raise or lower the zoom by as many levels as `steps`."""
        current = ZOOM_LEVELS.index(round(zoom_level(self.zoom) * 100))
        k = max(0, min(len(ZOOM_LEVELS) - 1, current + steps))
        self.set_zoom(ZOOM_LEVELS[k] / 100)

    def fit_width(self):
        width, _ = self._measurements()
        available = self.viewport().width() - 2 * PAGE_GAP
        self.set_zoom(available / width, "down")

    def whole_page(self):
        if self.continuous:             # sheets of variable height: fit to width
            self.fit_width()
            return
        width, height = self._measurements()
        g = PAGE_GAP
        z = min((self.viewport().width() - 2 * g) / width,
                (self.viewport().height() - g) / height)
        self.set_zoom(z, "down")
        current, _ = self._compute_page()
        y = (current - 1) * (height + g) + g / 2
        self.centerOn(self.proxy.mapToScene(QPointF(width / 2, y + height / 2)))

    def wheelEvent(self, e):
        if e.modifiers() & Qt.ControlModifier:   # Ctrl + wheel = zoom
            self._wheel_remainder = getattr(self, "_wheel_remainder", 0) + e.angleDelta().y()
            steps = int(self._wheel_remainder / 120)   # one notch = one level
            if steps:
                self._wheel_remainder -= steps * 120
                self.step_zoom(steps)
            e.accept()
            return
        # Normal scrolling of the view (not of the inner editor)
        QAbstractScrollArea.wheelEvent(self, e)


# ---------------------------------------------------------------------------
# "Settings" dialog (page, view and headings)
# ---------------------------------------------------------------------------
class SettingsDialog(QDialog):
    def __init__(self, config, mode=VIEW_PAGES, collapse=True, language=None,
                 parent=None):
        super().__init__(parent)
        self.setWindowTitle(_t("Settings"))
        self.setWindowIcon(file_icon("settings"))
        lay = QVBoxLayout(self)

        language_group = QGroupBox(_t("Language"))
        li = QVBoxLayout(language_group)
        self.language_combo = QComboBox()
        for code, name in LANGUAGES:       # each language with its own name
            self.language_combo.addItem(name, code)
        self.language_combo.setCurrentIndex(
            max(0, self.language_combo.findData(language or LANGUAGE)))
        li.addWidget(self.language_combo)
        lay.addWidget(language_group)

        view_group = QGroupBox(_t("View"))
        lv = QVBoxLayout(view_group)
        self.radio_pages = QRadioButton(_t("Pages"))
        self.radio_pages.setToolTip(
            _t("Shows the document on sheets, as it will look when printed"))
        self.radio_no_pages = QRadioButton(_t("No pages"))
        self.radio_no_pages.setToolTip(
            _t("A single continuous sheet, with no breaks between pages, to write "
               "and read without interruptions (page breaks are not allowed)"))
        self.radio_mixed = QRadioButton(_t("Mixed"))
        self.radio_mixed.setToolTip(
            _t("Like No pages, but each page break (Ctrl+Enter) splits the sheet: "
               "sheets of different lengths, each as long as its content"))
        lv.addWidget(self.radio_pages)
        lv.addWidget(self.radio_no_pages)
        lv.addWidget(self.radio_mixed)
        {VIEW_NO_PAGES: self.radio_no_pages,
         VIEW_MIXED: self.radio_mixed}.get(mode, self.radio_pages).setChecked(True)
        lay.addWidget(view_group)

        headings_group = QGroupBox(_t("Headings"))
        lt = QVBoxLayout(headings_group)
        self.chk_collapse = QCheckBox(_t(
            "Allow collapsing {title}, {subtitle} and {section}",
            title=style_name("title"), subtitle=style_name("subtitle"),
            section=style_name("section")))
        self.chk_collapse.setToolTip(
            _t("Shows a triangle next to each heading to hide or show its "
               "content. Turning it off expands everything."))
        self.chk_collapse.setChecked(collapse)
        lt.addWidget(self.chk_collapse)
        lay.addWidget(headings_group)

        paper_group = QGroupBox(_t("Paper"))
        form = QFormLayout(paper_group)
        self.size_combo = QComboBox()
        for name, wd, ht in PAGE_SIZES:
            self.size_combo.addItem(f"{_t(name)}  ({wd:g} × {ht:g} cm)", (wd, ht))
        self.size_combo.addItem(_t("Custom"), None)
        form.addRow(_t("Size"), self.size_combo)

        self.spin_width = self._spin_cm(5, 100)
        self.spin_height = self._spin_cm(5, 100)
        form.addRow(_t("Width"), self.spin_width)
        form.addRow(_t("Height"), self.spin_height)

        self.orientation_combo = QComboBox()
        self.orientation_combo.addItems([_t("Portrait"), _t("Landscape")])
        form.addRow(_t("Orientation"), self.orientation_combo)
        lay.addWidget(paper_group)

        margins_group = QGroupBox(_t("Margins"))
        fm = QFormLayout(margins_group)
        self.spin_top = self._spin_cm(0, 20)
        self.spin_bottom = self._spin_cm(0, 20)
        self.spin_left = self._spin_cm(0, 20)
        self.spin_right = self._spin_cm(0, 20)
        fm.addRow(_t("Top"), self.spin_top)
        fm.addRow(_t("Bottom"), self.spin_bottom)
        fm.addRow(_t("margin|Left"), self.spin_left)
        fm.addRow(_t("margin|Right"), self.spin_right)
        lay.addWidget(margins_group)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText(_t("OK"))
        buttons.button(QDialogButtonBox.Cancel).setText(_t("Cancel"))
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)
        # Version and license, discreetly, to the left of the buttons
        about = QLabel(f"{APP_NAME} v{APP_VERSION}<br>{APP_LICENSE}")
        about.setObjectName("about")
        about.setStyleSheet("QLabel#about { color: #9A9A9A; font-size: 10px; }")
        about.setTextInteractionFlags(Qt.TextSelectableByMouse)
        footer = QHBoxLayout()
        footer.addWidget(about, 0, Qt.AlignBottom)
        footer.addStretch()
        footer.addWidget(buttons)
        lay.addLayout(footer)

        # Current values
        self._loading = True
        self.spin_width.setValue(config["width"])
        self.spin_height.setValue(config["height"])
        self.spin_top.setValue(config["top"])
        self.spin_bottom.setValue(config["bottom"])
        self.spin_left.setValue(config["left"])
        self.spin_right.setValue(config["right"])
        self._loading = False
        self._sync_size()

        self.size_combo.currentIndexChanged.connect(self._choose_size)
        self.orientation_combo.currentIndexChanged.connect(self._rotate)
        self.spin_width.valueChanged.connect(self._sync_size)
        self.spin_height.valueChanged.connect(self._sync_size)

    def _spin_cm(self, minimum, maximum):
        s = QDoubleSpinBox()
        s.setRange(minimum, maximum)
        s.setDecimals(2)
        s.setSingleStep(0.1)
        s.setSuffix(" cm")
        s.setLocale(QLocale.c())  # always a decimal point
        return s

    def _choose_size(self, i):
        sizes = self.size_combo.itemData(i)
        if sizes is None or self._loading:
            return
        wd, ht = sizes
        if self.orientation_combo.currentIndex() == 1:
            wd, ht = ht, wd
        self._loading = True
        self.spin_width.setValue(wd)
        self.spin_height.setValue(ht)
        self._loading = False

    def _rotate(self, i):
        wd, ht = self.spin_width.value(), self.spin_height.value()
        if (i == 1) != (wd > ht):          # swap width and height
            self._loading = True
            self.spin_width.setValue(ht)
            self.spin_height.setValue(wd)
            self._loading = False

    def _sync_size(self, *_):
        """Show the predefined size and orientation that match the typed width
        and height."""
        if self._loading:
            return
        wd, ht = self.spin_width.value(), self.spin_height.value()
        short_side, length = sorted((wd, ht))
        index = self.size_combo.count() - 1   # Custom
        for i, (_, pa, pl) in enumerate(PAGE_SIZES):
            if abs(pa - short_side) < 0.01 and abs(pl - length) < 0.01:
                index = i
                break
        self._loading = True
        self.size_combo.setCurrentIndex(index)
        self.orientation_combo.setCurrentIndex(1 if wd > ht else 0)
        self._loading = False

    def _accept(self):
        c = self.config()
        if c["left"] + c["right"] > c["width"] - 2 or c["top"] + c["bottom"] > c["height"] - 2:
            QMessageBox.warning(
                self, _t("Settings"),
                _t("The margins are too large for the page size. They must leave "
                   "at least 2 cm of space for the text."))
            return
        self.accept()

    def mode(self):
        if self.radio_no_pages.isChecked():
            return VIEW_NO_PAGES
        if self.radio_mixed.isChecked():
            return VIEW_MIXED
        return VIEW_PAGES

    def collapse(self):
        return self.chk_collapse.isChecked()

    def language(self):
        return self.language_combo.currentData()

    def config(self):
        return {
            "width": self.spin_width.value(), "height": self.spin_height.value(),
            "top": self.spin_top.value(), "bottom": self.spin_bottom.value(),
            "left": self.spin_left.value(), "right": self.spin_right.value(),
        }


# ---------------------------------------------------------------------------
# Status bar with page indicator and zoom
# ---------------------------------------------------------------------------
# A word: letters or digits, with inner apostrophes or hyphens ("well-known",
# "l'amour"); the dialogue dash and punctuation don't count
WORD = re.compile(r"\w+(?:[’'\-]\w+)*")


def count_words(text):
    return sum(1 for _ in WORD.finditer(text))


class StatusBar(QStatusBar):
    def __init__(self, view, parent=None):
        super().__init__(parent)
        self.view = view
        self.setSizeGripEnabled(True)

        self.page_label = QLabel()
        self.addWidget(self.page_label)

        # Word counter: the total is recalculated shortly after you stop typing
        # (in long books counting on every keystroke would slow things down);
        # with selected text it shows "n of total"
        self.words_separator = _vertical_line()
        self.addWidget(self.words_separator)
        self.words_label = QLabel()
        self.addWidget(self.words_label)
        self._total_words = 0
        self._counter = QTimer(self)
        self._counter.setSingleShot(True)
        self._counter.setInterval(300)
        self._counter.timeout.connect(self._count_total)
        editor = view.editor
        editor.document().contentsChanged.connect(self._counter.start)
        editor.selectionChanged.connect(self._show_words)

        self.btn_minus = QToolButton()
        self.btn_minus.setText("−")
        self.btn_minus.setToolTip(_t("Zoom out"))
        self.btn_minus.setAutoRepeat(True)
        self.btn_minus.clicked.connect(lambda: self._step(-1))

        self.slider = QSlider(Qt.Horizontal)
        # One position per zoom level (no values in between)
        self.slider.setRange(0, len(ZOOM_LEVELS) - 1)
        self.slider.setFixedWidth(140)
        self.slider.setToolTip(_t("Zoom ({a} % to {b} %)", a=ZOOM_LEVELS[0],
                                     b=ZOOM_LEVELS[-1]))
        self.slider.setSingleStep(1)
        self.slider.setPageStep(1)
        self.slider.setTickPosition(QSlider.TicksBelow)
        self.slider.setTickInterval(1)
        self.slider.valueChanged.connect(
            lambda v: self.view.set_zoom(ZOOM_LEVELS[v] / 100))

        self.btn_plus = QToolButton()
        self.btn_plus.setText("+")
        self.btn_plus.setToolTip(_t("Zoom in"))
        self.btn_plus.setAutoRepeat(True)
        self.btn_plus.clicked.connect(lambda: self._step(+1))

        self.btn_zoom = QToolButton()
        self.btn_zoom.setToolTip(_t("Zoom level"))
        self.btn_zoom.setPopupMode(QToolButton.InstantPopup)
        self.btn_zoom.setMinimumWidth(56)
        menu = QMenu(self.btn_zoom)
        for z in ZOOM_MENU:
            menu.addAction(f"{z} %").triggered.connect(
                lambda _=False, z=z: self.view.set_zoom(z / 100))
        menu.addSeparator()
        menu.addAction(_t("Fit to width")).triggered.connect(self.view.fit_width)
        menu.addAction(_t("Whole page")).triggered.connect(self.view.whole_page)
        self.btn_zoom.setMenu(menu)

        for w in (self.btn_minus, self.slider, self.btn_plus, self.btn_zoom):
            w.setFocusPolicy(Qt.NoFocus)   # don't take the focus away from the text
            self.addPermanentWidget(w)

        view.zoomChanged.connect(self._show_zoom)
        view.pageChanged.connect(self._show_page)
        self._show_zoom(view.zoom)
        self._show_page(*view._compute_page())
        self._count_total()

    # -- words ---------------------------------------------------------------
    @staticmethod
    def _num(n):
        return QLocale().toString(int(n))

    def _count_total(self):
        self._counter.stop()
        text = self.view.editor.document().toPlainText()
        self._total_words = count_words(text)
        without_spaces = len("".join(text.replace("\ufffc", "").split()))
        with_spaces = len(text.replace("\ufffc", ""))
        paragraphs = sum(1 for line in text.split("\n") if line.strip(" \ufffc\t"))
        self.words_label.setToolTip("\n".join((
            _t("Words: {n}", n=self._num(self._total_words)),
            _t("Characters (with spaces): {n}", n=self._num(with_spaces)),
            _t("Characters (no spaces): {n}", n=self._num(without_spaces)),
            _t("Paragraphs: {n}", n=self._num(paragraphs)))))
        self._show_words()

    def _show_words(self):
        total = self._total_words
        words = _t("word") if total == 1 else _t("words")
        cursor = self.view.editor.textCursor()
        if cursor.hasSelection():
            n = count_words(cursor.selectedText())
            self.words_label.setText(_t("{n} of {total} {words}", n=self._num(n),
                                         total=self._num(total), words=words))
        else:
            self.words_label.setText(f"{self._num(total)} {words}")

    def _show_page(self, p, t):
        # In the No pages view only the word counter remains
        self.page_label.setText(_t("Page {p} of {t}", p=p, t=t) if t else "")
        self.page_label.setVisible(bool(t))
        self.words_separator.setVisible(bool(t))

    @staticmethod
    def _to_slider(z):
        return ZOOM_LEVELS.index(round(zoom_level(z) * 100))

    def _step(self, direction):
        """Go to the next or previous zoom level."""
        self.view.step_zoom(direction)

    def _show_zoom(self, z):
        self.btn_zoom.setText(f"{round(z * 100)} %")
        self.slider.blockSignals(True)
        self.slider.setValue(self._to_slider(z))
        self.slider.blockSignals(False)
        self.btn_minus.setEnabled(z > ZOOM_MIN + 1e-4)
        self.btn_plus.setEnabled(z < ZOOM_MAX - 1e-4)


# ---------------------------------------------------------------------------
# .docx files (Office Open XML)
#
# It is the format that Google Docs, Microsoft Word and LibreOffice Writer open
# and save without loss, and the one platforms such as Amazon KDP ask for. It
# is read and written with the standard library (zipfile + XML), with no
# dependencies.
#
# Preserved: font, size, bold, italic, underline, strikethrough,
# subscript/superscript, color, highlight, letter spacing, alignment, line
# spacing, space before/after, indents, page breaks, heading levels (as
# "Heading 1/2/3", which programs use for the table of contents and when
# exporting to EPUB), lists, and the page size and margins.
# ---------------------------------------------------------------------------
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
TWIPS_PER_CM = 1440 / 2.54
DOCX_FILTER = "Word document (*.docx)"
OPEN_FILTER = "Documents (*.docx *.txt);;Word document (*.docx);;Text (*.txt)"

# Named colors of <w:highlight>
DOCX_HIGHLIGHTS = {
    "yellow": "FFFF00", "green": "00FF00", "cyan": "00FFFF",
    "magenta": "FF00FF", "blue": "0000FF", "red": "FF0000",
    "darkBlue": "000080", "darkCyan": "008080", "darkGreen": "008000",
    "darkMagenta": "800080", "darkRed": "800000", "darkYellow": "808000",
    "darkGray": "808080", "lightGray": "C0C0C0", "black": "000000",
    "white": "FFFFFF",
}
_HIGHLIGHT_BY_HEX = {v: k for k, v in DOCX_HIGHLIGHTS.items()}
_INVALID_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffc\ufffe\uffff]")


def _w(tag):
    return f"{{{NS_W}}}{tag}"


def _attr(el, name, default=None):
    if el is None:
        return default
    return el.get(_w(name), default)


def _num(el, name, default=None):
    v = _attr(el, name)
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _enabled(el):
    """Yes/no properties: <w:b/> = yes, <w:b w:val="0"/> = no."""
    return _attr(el, "val", "true").lower() not in ("0", "false", "off", "none")


class DocxWriter:
    def __init__(self, editor, config, mode=VIEW_PAGES):
        self.editor = editor
        self.config = config
        self.mode = mode
        self.dpi = editor.logicalDpiY()
        self.comments = getattr(editor, "comments", None)
        self.used = set()        # comments that remain in the document
        self.media = {}           # name → (rId, file, ext, data)
        self.lists = {}           # id(QTextList) → (numId, numbered?, list)
        self._drawings = 0
        font = editor.document().defaultFont()
        self.default_family = font.family() or CHAPTER_BODY["font_family"]
        self.default_size = font.pointSizeF() if font.pointSizeF() > 0 \
            else CHAPTER_BODY["font_size"]

    def tw(self, px):
        """Editor px → twips (1/1440 of an inch)."""
        return round(px * 1440 / self.dpi)

    # -- XML for paragraphs and runs ----------------------------------------
    def rpr(self, cf):
        """Character properties, in the order the schema requires."""
        # (don't use cf.fontFamily(): it is deprecated and fails in PySide6
        # with empty formats; without its own font, the document's is used)
        families = cf.fontFamilies() if cf.hasProperty(QTextFormat.FontFamilies) else None
        family = (families[0] if families else None) or self.default_family
        f = xml_escape(family, {'"': "&quot;"})
        p = [f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>']
        if cf.fontWeight() >= 600:
            p.append("<w:b/><w:bCs/>")
        if cf.fontItalic():
            p.append("<w:i/><w:iCs/>")
        if cf.fontStrikeOut():
            p.append("<w:strike/>")
        if cf.hasProperty(QTextFormat.ForegroundBrush) and \
                cf.foreground().style() != Qt.NoBrush:
            p.append(f'<w:color w:val="{cf.foreground().color().name()[1:].upper()}"/>')
        letters = letter_spacing_pt(self.editor, cf)
        if letters:                                    # in 1/20 of a point
            p.append(f'<w:spacing w:val="{round(letters * 20)}"/>')
        size = cf.fontPointSize() or self.default_size
        media = max(2, round(size * 2))
        p.append(f'<w:sz w:val="{media}"/><w:szCs w:val="{media}"/>')
        background = None
        if cf.hasProperty(QTextFormat.BackgroundBrush) and \
                cf.background().style() != Qt.NoBrush:
            background = cf.background().color().name()[1:].upper()
            if background in _HIGHLIGHT_BY_HEX:          # standard highlight
                p.append(f'<w:highlight w:val="{_HIGHLIGHT_BY_HEX[background]}"/>')
                background = None
        if cf.fontUnderline():
            p.append('<w:u w:val="single"/>')
        if background:                                     # any other color
            p.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{background}"/>')
        va = cf.verticalAlignment()
        if va == QTextCharFormat.AlignSuperScript:
            p.append('<w:vertAlign w:val="superscript"/>')
        elif va == QTextCharFormat.AlignSubScript:
            p.append('<w:vertAlign w:val="subscript"/>')
        return "<w:rPr>" + "".join(p) + "</w:rPr>"

    @staticmethod
    def text(t):
        """Text → <w:t>, <w:tab/> and <w:br/> (line break) elements."""
        parts = []
        for i, line in enumerate(_INVALID_XML.sub("", t).split("\u2028")):
            if i:
                parts.append("<w:br/>")
            for j, chunk in enumerate(line.split("\t")):
                if j:
                    parts.append("<w:tab/>")
                if chunk:
                    parts.append(f'<w:t xml:space="preserve">{xml_escape(chunk)}</w:t>')
        return "".join(parts)

    def ppr(self, block):
        bf = block.blockFormat()
        p = []
        level = min(6, bf.headingLevel())
        if level > 0:
            p.append(f'<w:pStyle w:val="Heading{level}"/>')
        if bf.pageBreakPolicy() & QTextFormat.PageBreak_AlwaysBefore:
            p.append("<w:pageBreakBefore/>")
        lst = block.textList()
        if lst is not None:
            numbered = lst.format().style() in (
                QTextListFormat.ListDecimal, QTextListFormat.ListLowerAlpha,
                QTextListFormat.ListUpperAlpha, QTextListFormat.ListLowerRoman,
                QTextListFormat.ListUpperRoman)
            key = id(lst)
            if key not in self.lists:
                self.lists[key] = (len(self.lists) + 1, numbered, lst)
            list_level = max(0, min(8, lst.format().indent() - 1))
            p.append(f'<w:numPr><w:ilvl w:val="{list_level}"/>'
                     f'<w:numId w:val="{self.lists[key][0]}"/></w:numPr>')
        types = QTextBlockFormat.LineHeightTypes
        kind = bf.lineHeightType()
        if kind == types.ProportionalHeight.value:
            line, rule = round(240 * bf.lineHeight() / 100), "auto"
        elif kind == types.FixedHeight.value:
            line, rule = self.tw(bf.lineHeight()), "exact"
        elif kind == types.MinimumHeight.value:
            line, rule = self.tw(bf.lineHeight()), "atLeast"
        else:
            line, rule = 240, "auto"
        p.append(f'<w:spacing w:before="{self.tw(bf.topMargin())}" '
                 f'w:after="{self.tw(bf.bottomMargin())}" '
                 f'w:line="{max(1, line)}" w:lineRule="{rule}"/>')
        left = bf.leftMargin()
        if lst is not None:                 # the list's indent
            left += LIST_INDENT * max(1, lst.format().indent())
            if bf.textIndent() == 0:
                bf.setTextIndent(-LIST_INDENT * 0.6)
        indent = f'<w:ind w:left="{self.tw(left)}" ' \
                  f'w:right="{self.tw(bf.rightMargin())}" '
        first_line = self.tw(bf.textIndent())
        indent += f'w:hanging="{-first_line}"/>' if first_line < 0 \
            else f'w:firstLine="{first_line}"/>'
        p.append(indent)
        a = bf.alignment() & Qt.AlignHorizontal_Mask
        if a & Qt.AlignJustify:
            jc = "both"
        elif a & Qt.AlignHCenter:
            jc = "center"
        elif a & (Qt.AlignRight | Qt.AlignTrailing):
            jc = "right"
        else:
            jc = "left"
        p.append(f'<w:jc w:val="{jc}"/>')
        p.append(self.rpr(block.charFormat()))      # paragraph mark
        return "<w:pPr>" + "".join(p) + "</w:pPr>"

    def runs_of(self, block, valid_ids):
        """[(rPr, [texts], comment ids)], merging identical runs."""
        segs = []
        it = block.begin()
        while not it.atEnd():
            frag = it.fragment()
            if frag.isValid():
                cf = frag.charFormat()
                rpr = self.rpr(cf)
                ids = frozenset(comment_ids(cf) & valid_ids)
                if segs and segs[-1][0] == rpr and segs[-1][2] == ids:
                    segs[-1][1].append(frag.text())
                else:
                    segs.append((rpr, [frag.text()], ids))
            it += 1
        return segs

    # -- package parts -------------------------------------------------------
    def document_xml(self):
        c = self.config
        tw = lambda cm: round(cm * TWIPS_PER_CM)
        orient = ' w:orient="landscape"' if c["width"] > c["height"] else ""
        valid_ids = set(self.comments.data) if self.comments else set()
        blocks = []
        b = self.editor.document().begin()
        while b.isValid():
            blocks.append(b)
            b = b.next()
        segs = [[] if is_image(b) else self.runs_of(b, valid_ids) for b in blocks]
        # A comment closes when the next run (even in another paragraph) no
        # longer carries it
        flat = [ids for ss in segs for _, _, ids in ss]
        self.used = set()
        open_ids, k, body = set(), 0, []
        for b, ss in zip(blocks, segs):
            if is_image(b):
                body.append(self.image_paragraph(b))
                continue
            parts = []
            for rpr, texts, ids in ss:
                for n in sorted(ids - open_ids):
                    parts.append(f'<w:commentRangeStart w:id="{n}"/>')
                    open_ids.add(n)
                    self.used.add(n)
                parts.append(f"<w:r>{rpr}{self.text(''.join(texts))}</w:r>")
                following = flat[k + 1] if k + 1 < len(flat) else frozenset()
                for n in sorted(ids - following):
                    parts.append(f'<w:commentRangeEnd w:id="{n}"/>'
                                  f'<w:r><w:commentReference w:id="{n}"/></w:r>')
                    open_ids.discard(n)
                k += 1
            body.append(f"<w:p>{self.ppr(b)}{''.join(parts)}</w:p>")
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document xmlns:w="{NS_W}" xmlns:r="{NS_R}" xmlns:wp="{NS_WP}" '
            f'xmlns:a="{NS_A}" xmlns:pic="{NS_PIC}"><w:body>'
            + "".join(body) +
            f'<w:sectPr><w:pgSz w:w="{tw(c["width"])}" w:h="{tw(c["height"])}"{orient}/>'
            f'<w:pgMar w:top="{tw(c["top"])}" w:right="{tw(c["right"])}" '
            f'w:bottom="{tw(c["bottom"])}" w:left="{tw(c["left"])}" '
            'w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>'
            "</w:body></w:document>")

    def image_paragraph(self, b):
        """Centered paragraph with the inline image: in Word, Writer, Google
        Docs and when converting to EPUB it looks the same as "Top and
        bottom"."""
        cf, _ = self.editor.object_format(b)
        info = self.editor.images.get(cf.stringProperty(PROP_IMG_NAME)) \
            if cf is not None else None
        if info is None:
            return f"<w:p>{self.ppr(b)}</w:p>"
        name = cf.stringProperty(PROP_IMG_NAME)
        if name not in self.media:
            k = len(self.media) + 1
            self.media[name] = (f"rIdImg{k}", f"image{k}.{info['ext']}",
                                   info["ext"], info["data"])
        rid, file = self.media[name][:2]
        width, height = self.editor.shown_size(cf)
        emu = lambda px: max(1, round(px / self.dpi * 914400))
        cx, cy = emu(width), emu(height)
        self._drawings += 1
        n = self._drawings
        drawing_xml = (
            '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{n}" name="Imagen {n}"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/>'
            '</wp:cNvGraphicFramePr><a:graphic>'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="{file}"/><pic:cNvPicPr/>'
            f'</pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/>'
            f'</a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/>'
            f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>'
            '</a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
            '</wp:inline></w:drawing>')
        return f"<w:p>{self.ppr(b)}<w:r>{drawing_xml}</w:r></w:p>"

    def numbering_xml(self):
        """List definitions: abstractNum 1 = bullets, 2 = numbers. Each list in
        the document has its own numId (so each numbered list starts at 1)."""
        def levels(numbered):
            xml = []
            bullets = ["\u2022", "\u25e6", "\u25aa"]
            for n in range(9):
                left = self.tw(LIST_INDENT * (n + 1))
                hanging = self.tw(LIST_INDENT * 0.6)
                if numbered:
                    format_spec, text = "decimal", f"%{n + 1}."
                else:
                    format_spec, text = "bullet", bullets[n % 3]
                xml.append(f'<w:lvl w:ilvl="{n}"><w:start w:val="1"/>'
                           f'<w:numFmt w:val="{format_spec}"/><w:lvlText w:val="{text}"/>'
                           '<w:lvlJc w:val="left"/>'
                           f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr>'
                           '</w:lvl>')
            return "".join(xml)
        parts = [f'<w:abstractNum w:abstractNumId="1">{levels(False)}</w:abstractNum>',
                  f'<w:abstractNum w:abstractNumId="2">{levels(True)}</w:abstractNum>']
        for num, numbered, _ in sorted(self.lists.values(), key=lambda x: x[0]):
            restart_xml = ('<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/>'
                        '</w:lvlOverride>') if numbered else ""
            parts.append(f'<w:num w:numId="{num}"><w:abstractNumId '
                          f'w:val="{2 if numbered else 1}"/>{restart_xml}</w:num>')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:numbering xmlns:w="{NS_W}">' + "".join(parts) + "</w:numbering>")

    def comments_xml(self):
        q = {'"': "&quot;"}
        parts = []
        for n in sorted(self.used):
            d = self.comments.data[n]
            author = d.get("author") or _t("Author")
            initials = "".join(p[0] for p in author.split()[:3]).upper()
            ps = []
            for j, line in enumerate((d.get("text") or "").split("\n")):
                ref = "<w:r><w:annotationRef/></w:r>" if j == 0 else ""
                ps.append(f"<w:p>{ref}<w:r>{self.text(line)}</w:r></w:p>")
            parts.append(
                f'<w:comment w:id="{n}" w:author="{xml_escape(author, q)}" '
                f'w:date="{xml_escape(d.get("date") or now_iso(), q)}" '
                f'w:initials="{xml_escape(initials, q)}">{"".join(ps)}</w:comment>')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:comments xmlns:w="{NS_W}">' + "".join(parts) + "</w:comments>")

    def _style_xml(self, ident, name, style, level=None):
        f = xml_escape(style["font_family"], {'"': "&quot;"})
        jc = {Qt.AlignHCenter: "center", Qt.AlignRight: "right",
              Qt.AlignJustify: "both"}.get(style["alignment"], "left")
        ppr = (f'<w:pPr><w:keepNext/><w:keepLines/>' if level else "<w:pPr>")
        ppr += (f'<w:spacing w:before="{self.tw(style["space_before"])}" '
                f'w:after="{self.tw(style["space_after"])}" '
                f'w:line="{round(240 * style["line_spacing"] / 100)}" w:lineRule="auto"/>'
                f'<w:ind w:left="{self.tw(style["indent_left"])}" '
                f'w:right="{self.tw(style["indent_right"])}" '
                f'w:firstLine="{self.tw(style["first_line_indent"])}"/>'
                f'<w:jc w:val="{jc}"/>')
        if level:
            ppr += f'<w:outlineLvl w:val="{level - 1}"/>'
        ppr += "</w:pPr>"
        media = round(style["font_size"] * 2)
        rpr = (f'<w:rPr><w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>'
               + ("<w:b/><w:bCs/>" if style["bold"] else "")
               + ("<w:i/><w:iCs/>" if style["italic"] else "")
               + f'<w:color w:val="{style["font_color"][1:].upper()}"/>'
               f'<w:sz w:val="{media}"/><w:szCs w:val="{media}"/></w:rPr>')
        default = ' w:default="1"' if ident == "Normal" else ""
        base = "" if ident == "Normal" else '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
        return (f'<w:style w:type="paragraph"{default} w:styleId="{ident}">'
                f'<w:name w:val="{name}"/>{base}<w:qFormat/>{ppr}{rpr}</w:style>')

    def styles(self):
        b = CHAPTER_BODY
        f = xml_escape(b["font_family"], {'"': "&quot;"})
        headings = {1: CHAPTER_TITLE, 2: CHAPTER_SUBTITLE, 3: CHAPTER_SECTION}
        xml = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:styles xmlns:w="{NS_W}"><w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>'
            f'<w:sz w:val="{round(b["font_size"] * 2)}"/>'
            f'<w:szCs w:val="{round(b["font_size"] * 2)}"/>'
            f'<w:lang w:val="{language_tag()}" w:eastAsia="{language_tag()}" '
            'w:bidi="ar-SA"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
            '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr></w:pPrDefault></w:docDefaults>',
            self._style_xml("Normal", "Normal", CHAPTER_BODY),
        ]
        for n in range(1, 7):
            xml.append(self._style_xml(
                f"Heading{n}", f"heading {n}", headings.get(n, CHAPTER_SECTION), n))
        xml.append('<w:style w:type="character" w:default="1" '
                   'w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>'
                   '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>')
        xml.append("</w:styles>")
        return "".join(xml)

    def save(self, path, title=""):
        now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        ct = "application/vnd.openxmlformats-officedocument"
        parts = {
            "[Content_Types].xml":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                f'<Override PartName="/word/document.xml" ContentType="{ct}.wordprocessingml.document.main+xml"/>'
                f'<Override PartName="/word/styles.xml" ContentType="{ct}.wordprocessingml.styles+xml"/>'
                f'<Override PartName="/word/settings.xml" ContentType="{ct}.wordprocessingml.settings+xml"/>'
                '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                f'<Override PartName="/docProps/app.xml" ContentType="{ct}.extended-properties+xml"/>'
                '</Types>',
            "_rels/.rels":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Relationships xmlns="{NS_REL}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                '</Relationships>',
            "word/_rels/document.xml.rels":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Relationships xmlns="{NS_REL}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                '</Relationships>',
            "word/document.xml": self.document_xml(),
            "word/styles.xml": self.styles(),
            "word/settings.xml":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:settings xmlns:w="{NS_W}">'
                # Continuous views → "Web layout" in Word and Writer
                + ('<w:view w:val="web"/>' if self.mode != VIEW_PAGES else "") +
                '<w:defaultTabStop w:val="709"/>'
                '<w:characterSpacingControl w:val="doNotCompress"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                # The exact view of this editor (Word keeps it without using
                # it)
                f'<w:docVars><w:docVar w:name="SWordBookView" w:val="{self.mode}"/>'
                # Color of each comment (Word has nowhere to store it)
                + "".join(f'<w:docVar w:name="SWordBookComment{n}" '
                          f'w:val="{self.comments.color(n)}"/>'
                          for n in sorted(self.used)) +
                '</w:docVars>'
                '</w:settings>',
            "docProps/core.xml":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                f'<dc:title>{xml_escape(title)}</dc:title><dc:language>{language_tag()}</dc:language>'
                f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
                f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
                '</cp:coreProperties>',
            "docProps/app.xml":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                '<Application>SWordBook</Application></Properties>',
        }
        if self.media:          # images: word/media
            extensions = sorted({m[2] for m in self.media.values()})
            defaults = "".join(
                f'<Default Extension="{e}" ContentType="'
                f'{DOCX_IMAGE_TYPES.get(e, "image/" + e)}"/>' for e in extensions)
            parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
                '<Default Extension="xml"', defaults + '<Default Extension="xml"')
            rels = "".join(
                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                f'officeDocument/2006/relationships/image" Target="media/{file}"/>'
                for rid, file, _, _ in self.media.values())
            parts["word/_rels/document.xml.rels"] = \
                parts["word/_rels/document.xml.rels"].replace(
                    "</Relationships>", rels + "</Relationships>")
            for _, file, _, data in self.media.values():
                parts[f"word/media/{file}"] = data
        if self.used:          # comments: their own package part
            parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
                "</Types>",
                f'<Override PartName="/word/comments.xml" '
                f'ContentType="{ct}.wordprocessingml.comments+xml"/></Types>')
            parts["word/_rels/document.xml.rels"] = \
                parts["word/_rels/document.xml.rels"].replace(
                    "</Relationships>",
                    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/comments" Target="comments.xml"/>'
                    "</Relationships>")
            parts["word/comments.xml"] = self.comments_xml()
        if self.lists:          # lists: word/numbering.xml
            parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
                "</Types>",
                f'<Override PartName="/word/numbering.xml" '
                f'ContentType="{ct}.wordprocessingml.numbering+xml"/></Types>')
            parts["word/_rels/document.xml.rels"] = \
                parts["word/_rels/document.xml.rels"].replace(
                    "</Relationships>",
                    '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
                    "</Relationships>")
            parts["word/numbering.xml"] = self.numbering_xml()
        # Write to a temporary file and then replace: if something fails
        # halfway, the previous file stays intact
        temp = path + ".tmp"
        try:
            with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as z:
                for name, content in parts.items():
                    z.writestr(name, content if isinstance(content, bytes)
                               else content.encode("utf-8"))
            # Make sure the data reaches the disk before replacing (in case the
            # power goes out right after)
            with open(temp, "rb+") as f:
                os.fsync(f.fileno())
            os.replace(temp, path)
        finally:
            if os.path.exists(temp):
                os.remove(temp)


class DocxReader:
    """Read a .docx (from this editor, Word, Google Docs or Writer). Resolves
    style inheritance so the text looks as in the source program. Tables,
    notes and headers are not imported (table text is, as paragraphs)."""

    def __init__(self, path, dpi):
        self.dpi = dpi
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            main_part = "word/document.xml"
            if "_rels/.rels" in names:
                for rel in ET.fromstring(z.read("_rels/.rels")):
                    if rel.get("Type", "").endswith("/officeDocument"):
                        main_part = rel.get("Target", main_part).lstrip("/")
            if main_part not in names:
                raise ValueError(_t("The file does not contain a Word document."))
            self.doc = ET.fromstring(z.read(main_part))
            folder = main_part.rsplit("/", 1)[0]
            read = lambda n: ET.fromstring(z.read(n)) if n in names else None
            self._read_theme(read(f"{folder}/theme/theme1.xml"))
            self._read_styles(read(f"{folder}/styles.xml"))
            self._read_numbering(read(f"{folder}/numbering.xml"))
            comments_path = None
            self.media = {}       # rId → (data, ext) of each image
            rels = f"{folder}/_rels/{main_part.rsplit('/', 1)[-1]}.rels"
            if rels in names:
                for rel in ET.fromstring(z.read(rels)):
                    if rel.get("Type", "").endswith("/image") and \
                            rel.get("TargetMode") != "External":
                        target = rel.get("Target", "")
                        image_path = target.lstrip("/") if target.startswith("/") \
                            else posixpath.normpath(f"{folder}/{target}")
                        if image_path in names:
                            ext = posixpath.splitext(image_path)[1].lstrip(".").lower()
                            self.media[rel.get("Id")] = (z.read(image_path), ext or "png")
                    if rel.get("Type", "").endswith("/comments"):
                        target = rel.get("Target", "")
                        comments_path = target.lstrip("/") if target.startswith("/") \
                            else f"{folder}/{target}"
            self._read_comments(read(comments_path) if comments_path else None)
            settings_xml = read(f"{folder}/settings.xml")
            view = settings_xml.find(_w("view")) if settings_xml is not None else None
            self.mode = VIEW_NO_PAGES if _attr(view, "val") == "web" else VIEW_PAGES
            if settings_xml is not None:
                for var in settings_xml.iter(_w("docVar")):
                    name, value = _attr(var, "name") or "", _attr(var, "val")
                    # (the EditorLibro… names come from earlier versions)
                    if name in ("SWordBookView", "EditorLibroVista"):
                        value = OLD_VIEWS.get(value, value)
                        if value in (VIEW_PAGES, VIEW_NO_PAGES, VIEW_MIXED):
                            self.mode = value
                    elif re.fullmatch(r"(?:SWordBookComment|EditorLibroComentario)\d+",
                                      name):
                        n = int(re.search(r"\d+$", name).group())
                        if n in self.comments and QColor(value or "").isValid():
                            self.comments[n]["color"] = value

    def px(self, twips):
        return (twips or 0) * self.dpi / 1440

    def _read_numbering(self, root):
        """numId → {level: numbered?} (to know whether it is a bullet or a
        number)."""
        self.numbering = {}
        if root is None:
            return
        abstracts = {}
        for a in root.findall(_w("abstractNum")):
            levels = {}
            for lvl in a.findall(_w("lvl")):
                try:
                    n = int(_attr(lvl, "ilvl", "0"))
                except ValueError:
                    continue
                fmt = _attr(lvl.find(_w("numFmt")), "val", "bullet")
                levels[n] = fmt not in ("bullet", "none")
            abstracts[_attr(a, "abstractNumId")] = levels
        for num in root.findall(_w("num")):
            self.numbering[_attr(num, "numId")] = abstracts.get(
                _attr(num.find(_w("abstractNumId")), "val"), {})

    def _read_comments(self, root):
        self.comments = {}
        if root is None:
            return
        for c in root.findall(_w("comment")):
            try:
                i = int(_attr(c, "id"))
            except (TypeError, ValueError):
                continue
            paragraphs = []
            for p in c.iter(_w("p")):
                chunks = []
                for el in p.iter():
                    t = el.tag.split("}")[-1]
                    if t == "t":
                        chunks.append(el.text or "")
                    elif t == "tab":
                        chunks.append("\t")
                    elif t in ("br", "cr"):
                        chunks.append("\n")
                paragraphs.append("".join(chunks))
            self.comments[i] = {"author": _attr(c, "author") or "",
                                   "date": _attr(c, "date") or "",
                                   "text": "\n".join(paragraphs).strip()}

    def _comment_mark(self, tag, el, open_ids):
        try:
            i = int(_attr(el, "id"))
        except (TypeError, ValueError):
            return
        if tag == "commentRangeStart" and i in self.comments:
            open_ids.add(i)
        elif tag == "commentRangeEnd":
            open_ids.discard(i)

    # -- theme and styles ----------------------------------------------------
    def _read_theme(self, theme):
        self.theme_fonts = {}
        if theme is None:
            return
        for key, name in (("minor", "minorFont"), ("major", "majorFont")):
            el = theme.find(f".//{{{NS_A}}}{name}/{{{NS_A}}}latin")
            if el is not None and el.get("typeface"):
                self.theme_fonts[key] = el.get("typeface")

    def _read_styles(self, root):
        self.styles = {}
        self.ppr_default, self.rpr_default = {}, {}
        self.default_paragraph_style = None
        if root is None:
            return
        dd = root.find(_w("docDefaults"))
        if dd is not None:
            self.rpr_default = self.read_rpr(dd.find(f"{_w('rPrDefault')}/{_w('rPr')}"))
            self.ppr_default = self.read_ppr(dd.find(f"{_w('pPrDefault')}/{_w('pPr')}"))
        for st in root.findall(_w("style")):
            ident = _attr(st, "styleId")
            kind = _attr(st, "type")
            self.styles[ident] = {
                "kind": kind,
                "name": (_attr(st.find(_w("name")), "val") or "").lower(),
                "base": _attr(st.find(_w("basedOn")), "val"),
                "ppr": self.read_ppr(st.find(_w("pPr"))),
                "rpr": self.read_rpr(st.find(_w("rPr"))),
            }
            if kind == "paragraph" and _is_default(st):
                self.default_paragraph_style = ident

    def chain(self, ident):
        """The style and all its base styles, from the most general to its own."""
        res, seen = [], set()
        while ident and ident in self.styles and ident not in seen:
            seen.add(ident)
            res.insert(0, self.styles[ident])
            ident = self.styles[ident]["base"]
        return res

    # -- properties ------------------------------------------------------------
    def read_rpr(self, el):
        d = {}
        if el is None:
            return d
        for child in el:
            t = child.tag.split("}")[-1]
            if t == "rFonts":
                font = _attr(child, "ascii") or _attr(child, "hAnsi")
                theme = _attr(child, "asciiTheme") or _attr(child, "hAnsiTheme")
                if not font and theme:
                    font = self.theme_fonts.get("major" if theme.startswith("major") else "minor")
                if font:
                    d["font"] = font
            elif t in ("b", "i", "strike"):
                d[t] = _enabled(child)
            elif t == "u":
                d["u"] = _attr(child, "val", "single") != "none"
            elif t == "color":
                v = _attr(child, "val", "auto")
                d["color"] = None if v == "auto" else v
            elif t == "sz":
                v = _num(child, "val")
                if v:
                    d["size"] = v / 2
            elif t == "highlight":
                d["background"] = DOCX_HIGHLIGHTS.get(_attr(child, "val"))
            elif t == "shd":
                fill = _attr(child, "fill", "auto")
                if fill != "auto" and "background" not in d:
                    d["background"] = fill
            elif t == "spacing":                      # letter spacing
                v = _num(child, "val")
                if v is not None:
                    d["letters"] = v / 20                # 1/20 pt → pt
            elif t == "vertAlign":
                d["va"] = _attr(child, "val", "baseline")
            elif t == "rStyle":
                d["style"] = _attr(child, "val")
        return d

    def read_ppr(self, el):
        d = {}
        if el is None:
            return d
        for child in el:
            t = child.tag.split("}")[-1]
            if t == "pStyle":
                d["style"] = _attr(child, "val")
            elif t == "jc":
                d["jc"] = _attr(child, "val")
            elif t == "pageBreakBefore":
                d["break"] = _enabled(child)
            elif t == "numPr":
                num = _attr(child.find(_w("numId")), "val")
                list_level = _num(child.find(_w("ilvl")), "val", 0) or 0
                if num is not None:
                    d["list"] = (num, int(list_level))
            elif t == "outlineLvl":
                d["outline"] = _num(child, "val")
            elif t == "spacing":
                for key, name in (("before", "before"), ("after", "after"),
                                      ("line", "line")):
                    v = _num(child, name)
                    if v is not None:
                        d[key] = v
                if _attr(child, "lineRule"):
                    d["rule"] = _attr(child, "lineRule")
            elif t == "ind":
                for key, names in (("left", ("left", "start")),
                                       ("right", ("right", "end"))):
                    for name in names:
                        v = _num(child, name)
                        if v is not None:
                            d[key] = v
                            break
                if _num(child, "hanging") is not None:
                    d["first"] = -_num(child, "hanging")
                elif _num(child, "firstLine") is not None:
                    d["first"] = _num(child, "firstLine")
        return d

    def heading_level(self, chain, ppr):
        for st in reversed(chain):
            m = re.fullmatch(r"heading\s*(\d)", st["name"])
            if m:
                return min(6, max(1, int(m.group(1))))
            if st["name"] == "title":
                return 1
            if st["name"] == "subtitle":
                return 2
        outline_level = ppr.get("outline")
        if outline_level is not None and outline_level < 6:
            return int(outline_level) + 1
        return 0

    # -- conversion to Qt formats ----------------------------------------------
    def block_format(self, ppr, level):
        bf = QTextBlockFormat()
        jc = ppr.get("jc", "left")
        bf.setAlignment({"center": Qt.AlignHCenter, "right": Qt.AlignRight,
                         "end": Qt.AlignRight, "both": Qt.AlignJustify,
                         "distribute": Qt.AlignJustify,
                         "justify": Qt.AlignJustify}.get(jc, Qt.AlignLeft))
        bf.setTopMargin(self.px(ppr.get("before")))
        bf.setBottomMargin(self.px(ppr.get("after")))
        bf.setLeftMargin(self.px(ppr.get("left")))
        bf.setRightMargin(self.px(ppr.get("right")))
        bf.setTextIndent(self.px(ppr.get("first")))
        types = QTextBlockFormat.LineHeightTypes
        line, rule = ppr.get("line"), ppr.get("rule", "auto")
        if line and rule == "exact":
            bf.setLineHeight(self.px(line), types.FixedHeight.value)
        elif line and rule == "atLeast":
            bf.setLineHeight(self.px(line), types.MinimumHeight.value)
        else:
            bf.setLineHeight(round((line or 240) / 240 * 100),
                             types.ProportionalHeight.value)
        bf.setHeadingLevel(level)
        return bf

    def char_format(self, rpr):
        cf = QTextCharFormat()
        cf.setFontFamilies([rpr.get("font") or "Times New Roman"])
        cf.setFontPointSize(rpr.get("size", 10))
        cf.setFontWeight(QFont.Bold if rpr.get("b") else QFont.Normal)
        cf.setFontItalic(bool(rpr.get("i")))
        cf.setFontUnderline(bool(rpr.get("u")))
        cf.setFontStrikeOut(bool(rpr.get("strike")))
        color = rpr.get("color")
        cf.setForeground(QColor("#" + color) if color and QColor("#" + color).isValid()
                         else QColor("#000000"))
        background = rpr.get("background")
        if background and QColor("#" + background).isValid():
            cf.setBackground(QColor("#" + background))
        if rpr.get("letters"):
            cf.setFontLetterSpacingType(QFont.AbsoluteSpacing)
            cf.setFontLetterSpacing(rpr["letters"] * self.dpi / 72)
        va = rpr.get("va")
        if va == "superscript":
            cf.setVerticalAlignment(QTextCharFormat.AlignSuperScript)
        elif va == "subscript":
            cf.setVerticalAlignment(QTextCharFormat.AlignSubScript)
        return cf

    # -- walking the document --------------------------------------------------
    def paragraphs(self):
        """Yield (block format, mark format, [(text, format)]). Page breaks
        inside a paragraph split it in two."""
        body = self.doc.find(_w("body"))
        if body is None:
            return
        pending_break = False
        open_ids = set()           # open comments (they can span paragraphs)
        for p in self._iter_paragraphs(body):
            tag = p.tag.split("}")[-1]
            if tag != "p":
                self._comment_mark(tag, p, open_ids)
                continue
            ppr_el = p.find(_w("pPr"))
            direct = self.read_ppr(ppr_el)
            chain = self.chain(direct.get("style") or self.default_paragraph_style)
            ppr = dict(self.ppr_default)
            rpr_paragraph = dict(self.rpr_default)
            for st in chain:
                ppr.update(st["ppr"])
                rpr_paragraph.update(st["rpr"])
            ppr.update(direct)
            level = self.heading_level(chain, ppr)
            mark = dict(rpr_paragraph)
            if ppr_el is not None:
                mark.update(self.read_rpr(ppr_el.find(_w("rPr"))))

            def new_block(page_break):
                bf = self.block_format(ppr, level)
                lst = ppr.get("list")
                if lst and lst[0] != "0" and lst[0] in self.numbering and not level:
                    numbered = self.numbering[lst[0]].get(lst[1], False)
                    bf.setProperty(PROP_LIST, f"{'decimal' if numbered else 'bullet'};"
                                               f"{lst[0]};{lst[1]}")
                    # The list provides the indent
                    bf.setLeftMargin(0)
                    bf.setTextIndent(0)
                if page_break:
                    bf.setPageBreakPolicy(QTextFormat.PageBreak_AlwaysBefore)
                    bf.setProperty(PROP_MANUAL_BREAK, True)
                return [bf, self.char_format(mark), []]

            current = new_block(pending_break or ppr.get("break"))
            images = []      # (rId, cx, cy) in this paragraph
            pending_break = False
            cut = False      # there was a page break after some text
            for r in self._iter_runs(p):
                tag = r.tag.split("}")[-1]
                if tag != "r":
                    self._comment_mark(tag, r, open_ids)
                    continue
                rpr = dict(rpr_paragraph)
                direct_r = self.read_rpr(r.find(_w("rPr")))
                for st in self.chain(direct_r.get("style")):
                    rpr.update(st["rpr"])
                rpr.update(direct_r)
                cf = self.char_format(rpr)
                if open_ids:
                    set_comment_ids(cf, open_ids)
                seen = set()
                for drawing in r.iter(_w("drawing")):      # images (inline or floating)
                    blip = drawing.find(f".//{{{NS_A}}}blip")
                    rid = blip.get(f"{{{NS_R}}}embed") if blip is not None else None
                    ext = drawing.find(f".//{{{NS_WP}}}extent")
                    if rid in self.media and rid not in seen and ext is not None:
                        seen.add(rid)
                        images.append((rid, _plain_num(ext.get("cx")),
                                         _plain_num(ext.get("cy"))))
                for child in r:
                    t = child.tag.split("}")[-1]
                    text = None
                    if t == "t":
                        text = child.text or ""
                    elif t == "tab":
                        text = "\t"
                    elif t == "noBreakHyphen":
                        text = "\u2011"
                    elif t == "softHyphen":
                        text = "\u00ad"
                    elif t == "cr":
                        text = "\u2028"
                    elif t == "br":
                        if _attr(child, "type") == "page":
                            if current[2]:
                                cut = True         # resolved below
                            else:
                                current[0].setPageBreakPolicy(QTextFormat.PageBreak_AlwaysBefore)
                                current[0].setProperty(PROP_MANUAL_BREAK, True)
                        else:
                            text = "\u2028"
                    if text:
                        if cut:      # text after the break: split the paragraph
                            yield current
                            current = new_block(True)
                            cut = False
                        current[2].append((text, cf))
            # Images go in their own paragraph, after the text (a paragraph
            # that only had images doesn't leave an empty one)
            image_break = False
            if images and not current[2] and not cut:
                image_break = bool(current[0].pageBreakPolicy()
                                 & QTextFormat.PageBreak_AlwaysBefore)
            else:
                yield current
            for k, (rid, cx, cy) in enumerate(images):
                yield ("image", rid, cx, cy, image_break and k == 0)
            # Word and Writer put the break at the end of the previous
            # paragraph: here it moves to the start of the next one
            pending_break = cut
            # A section break (not continuous) starts a new page
            sect = ppr_el.find(_w("sectPr")) if ppr_el is not None else None
            if sect is not None and \
                    _attr(sect.find(_w("type")), "val", "nextPage") != "continuous":
                pending_break = True

    def _iter_paragraphs(self, container):
        for el in container:
            t = el.tag.split("}")[-1]
            if t in ("p", "commentRangeStart", "commentRangeEnd"):
                yield el
            elif t == "tbl":
                yield from el.iter(_w("p"))
            elif t in ("sdt", "sdtContent", "customXml", "ins"):
                yield from self._iter_paragraphs(el)

    def _iter_runs(self, el):
        for child in el:
            t = child.tag.split("}")[-1]
            if t in ("r", "commentRangeStart", "commentRangeEnd"):
                yield child
            elif t in ("pPr", "del", "moveFrom", "rPr"):
                continue
            else:                 # hyperlinks, insertions, fields, etc.
                yield from self._iter_runs(child)

    def page_config(self, current):
        """Page size and margins (from the last sectPr)."""
        body = self.doc.find(_w("body"))
        sect = body.find(_w("sectPr")) if body is not None else None
        config = dict(current)
        if sect is None:
            return config
        cm = lambda tw: round(abs(tw) / TWIPS_PER_CM, 2)
        size, margins_el = sect.find(_w("pgSz")), sect.find(_w("pgMar"))
        if _num(size, "w") and _num(size, "h"):
            config["width"], config["height"] = cm(_num(size, "w")), cm(_num(size, "h"))
        for key, name in (("top", "top"), ("bottom", "bottom"),
                              ("left", "left"), ("right", "right")):
            v = _num(margins_el, name)
            if v is not None:
                config[key] = cm(v)
        return config


def _plain_num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _is_default(st):
    return _attr(st, "default", "0") in ("1", "true", "on")


def open_docx(editor, view, path):
    """Replace the editor's content with that of the .docx."""
    reader = DocxReader(path, editor.logicalDpiY())
    paragraphs = list(reader.paragraphs())
    config = reader.page_config(view.config)
    # Built in a separate document and moved into the editor in one go
    # (inserting paragraph by paragraph into the editor repaginates every time)
    temp = QTextDocument()
    temp.setUndoRedoEnabled(False)
    cursor = QTextCursor(temp)
    images = {}
    lists = {}             # (numId, level) → QTextList
    dpi = editor.logicalDpiY()
    first = None
    for i, item in enumerate(paragraphs):
        if item[0] == "image":
            _, rid, cx, cy, page_break = item
            data, ext = reader.media[rid]
            name = editor.register_image(data, ext, images)
            width, height = cx / 914400 * dpi, cy / 914400 * dpi
            if width <= 0 or height <= 0:
                img = images[name]["img"]
                width, height = (img.width(), img.height()) if not img.isNull() else (300, 200)
            bf, mark = editor.image_block_format(page_break), editor.char_format()
            runs = [(OBJECT_CHAR, editor.image_format(name, width, height))]
        else:
            bf, mark, runs = item
        if i == 0:
            cursor.setBlockFormat(bf)
            cursor.setBlockCharFormat(mark)
            first = (bf, mark)
        else:
            cursor.insertBlock(bf, mark)
        for text, cf in runs:
            cursor.insertText(text, cf)
        _add_to_list(cursor, bf, lists)
    if paragraphs and paragraphs[-1][0] == "image":    # text after the image
        cursor.insertBlock(editor.paragraph_format(), editor.char_format())
    doc = editor.document()
    editor.blockSignals(True)
    doc.setUndoRedoEnabled(False)
    try:
        doc.clear()
        editor.clear_images()
        editor.images = images
        if paragraphs:
            target = QTextCursor(doc)
            target.setBlockFormat(first[0])
            target.setBlockCharFormat(first[1])
            target.insertFragment(QTextDocumentFragment(temp))
        else:
            editor._initial_style()
    finally:
        doc.setUndoRedoEnabled(True)     # also clears the history
        editor.blockSignals(False)
    editor.comments.load(reader.comments)
    view.mode = reader.mode                    # apply_config uses it
    editor.set_cuts([], 0)
    view.apply_config(config)
    view.editor.viewport().update()
    view.viewChanged.emit(view.mode)
    _cursor_to_start(editor)


def _add_to_list(cursor, bf, lists):
    """Put the paragraph just written into its list (read from the .docx)."""
    value = bf.stringProperty(PROP_LIST) if bf.hasProperty(PROP_LIST) else ""
    if not value:
        return
    kind, num, level = value.split(";")
    block = cursor.block()
    c = QTextCursor(block)
    clean = c.blockFormat()
    clean.clearProperty(PROP_LIST)
    c.setBlockFormat(clean)
    key = (num, level)
    lst = lists.get(key)
    if lst is None:
        fmt = QTextListFormat()
        fmt.setStyle(QTextListFormat.ListDecimal if kind == "decimal"
                     else [QTextListFormat.ListDisc, QTextListFormat.ListCircle,
                           QTextListFormat.ListSquare][int(level) % 3])
        fmt.setIndent(int(level) + 1)
        if kind == "decimal":
            fmt.setNumberSuffix(".")
        lists[key] = c.createList(fmt)
    else:
        lst.add(block)


def open_text(editor, view, path):
    with open(path, "rb") as f:
        data = f.read()
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    doc = editor.document()
    editor.blockSignals(True)
    doc.setUndoRedoEnabled(False)
    try:
        editor.setPlainText(text.replace("\r\n", "\n").replace("\r", "\n"))
        editor._initial_style()          # everything in the Body style
    finally:
        doc.setUndoRedoEnabled(True)
        editor.blockSignals(False)
    editor.comments.clear_all()
    editor.clear_images()
    view.apply_config(view.config)
    _cursor_to_start(editor)


def _cursor_to_start(editor):
    cursor = QTextCursor(editor.document())
    editor.setTextCursor(cursor)
    editor.setCurrentCharFormat(cursor.charFormat())
    editor.paragraphFormatChanged.emit()
    editor.document().clearUndoRedoStacks()   # nothing to undo after opening
    editor.document().setModified(False)


# ---------------------------------------------------------------------------
# Automatic backup: while there are unsaved changes, a copy is written ("name
# (backup).docx", next to the original). It is deleted when saving and when
# closing normally; if the program closes abruptly, it stays so it can be
# recovered. Untitled documents are backed up in the program folder. A registry
# makes it possible to offer recovery on the next launch.
# ---------------------------------------------------------------------------
BACKUP_WAIT = 5000         # ms without typing before backing up
BACKUP_MAX = 30           # s: even if typing continues


def app_folder():
    base = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation) \
        or os.path.join(os.path.expanduser("~"), ".editor_libros")
    os.makedirs(base, exist_ok=True)
    return base


def backup_path_of(path):
    folder, name = os.path.split(path)
    return os.path.join(folder, f"{os.path.splitext(name)[0]} (backup).docx")


def old_backup_path_of(path):
    """Backup name used by earlier versions of the program."""
    folder, name = os.path.split(path)
    return os.path.join(folder, f"{os.path.splitext(name)[0]} (respaldo).docx")


def _registry_path():
    return os.path.join(app_folder(), "respaldos.json")


def read_registry():
    try:
        with open(_registry_path(), encoding="utf-8") as f:
            data = json.load(f)
        entries = []
        for e in data:
            if isinstance(e, dict):
                if "respaldo" in e and "backup" not in e:   # earlier versions
                    e["backup"] = e.pop("respaldo")
                if e.get("backup"):
                    entries.append(e)
        return entries
    except (OSError, ValueError):
        return []


def _save_registry(lst):
    path = _registry_path()
    try:
        with open(path + ".tmp", "w", encoding="utf-8") as f:
            json.dump(lst, f, ensure_ascii=False, indent=1)
        os.replace(path + ".tmp", path)
    except OSError:
        pass


def register_backup(backup, original):
    lst = [e for e in read_registry() if e["backup"] != backup]
    lst.append({"backup": backup, "original": original})
    _save_registry(lst)


def remove_from_registry(backup):
    lst = read_registry()
    new_list = [e for e in lst if e["backup"] != backup]
    if len(new_list) != len(lst):
        _save_registry(new_list)


def backup_lock(backup):
    """Backup lock: held by the window that uses it. If the program closed
    abruptly, the lock is free (Qt detects it)."""
    folder = os.path.join(app_folder(), "bloqueos")
    os.makedirs(folder, exist_ok=True)
    key = hashlib.sha1(os.path.abspath(backup).encode("utf-8")).hexdigest()[:20]
    lock = QLockFile(os.path.join(folder, key + ".lock"))
    lock.setStaleLockTime(0)      # only stale if its program no longer exists
    return lock


def backup_in_use(backup):
    lock = backup_lock(backup)
    if lock.tryLock(0):
        lock.unlock()
        return False
    return True


# ---------------------------------------------------------------------------
# settings.ini: the program configuration. It is created next to the program
# (if that folder is not writable, in the program's data folder) with all the
# values, and can be edited by hand while the program is closed
# ---------------------------------------------------------------------------
DEFAULT_CONFIG = {
    # Interface language: en, es, pt, fr, de, it, zh, ja
    "Interface/language": "en",
    # Page for new documents (cm)
    "Page/width": PAGE["width"],
    "Page/height": PAGE["height"],
    "Page/margin_top": PAGE["top"],
    "Page/margin_bottom": PAGE["bottom"],
    "Page/margin_left": PAGE["left"],
    "Page/margin_right": PAGE["right"],
    # View: pages, no_pages or mixed; zoom in %
    "View/mode": VIEW_PAGES,
    "View/zoom": 100,
    "View/collapse_headings": True,
    # Window (0 = automatic size)
    "Window/x": -1,
    "Window/y": -1,
    "Window/width": 0,
    "Window/height": 0,
    "Window/maximized": False,
    "Panels/navigator": False,
    "Panels/comments": False,
    # Last used folders
    "Files/last_folder": "",
    "Files/images_folder": "",
    # Automatic backup (seconds)
    "Backup/enabled": True,
    "Backup/wait_seconds": BACKUP_WAIT // 1000,
    "Backup/max_seconds": BACKUP_MAX,
}
# Styles: each one in its own section [Styles.Name]. Dialogue uses the Body
# style; only its name and the symbol it types are configured
INI_STYLES = [            # section, internal key, style
    ("Body", "body", CHAPTER_BODY), ("Title", "title", CHAPTER_TITLE),
    ("Subtitle", "subtitle", CHAPTER_SUBTITLE), ("Section", "section", CHAPTER_SECTION),
    ("Thoughts", "thoughts", CHAPTER_THOUGHTS), ("Quote", "quote", CHAPTER_QUOTE),
    ("SceneBreak", "scene_break", SCENE_BREAK),
]
DEFAULT_STYLES = {section: dict(style) for section, _, style in INI_STYLES}
ALIGNMENTS = {"left": Qt.AlignLeft, "center": Qt.AlignHCenter,
                "right": Qt.AlignRight, "justify": Qt.AlignJustify}
_ALIGNMENT_NAME = {v: k for k, v in ALIGNMENTS.items()}


def _ini_number(v):
    """12.0 → 12; 1.5 → 1.5 (so the .ini reads well)."""
    v = round(float(v), 2)
    return int(v) if v == int(v) else v


for _section, _key, _e in INI_STYLES:
    _g = f"Styles.{_section}/"
    DEFAULT_CONFIG.update({
        # Empty = the style name in the interface language
        _g + "name": "",
        _g + "font": _e["font_family"],
        _g + "size": _ini_number(_e["font_size"]),
        _g + "color": _e["font_color"],
        _g + "bold": _e["bold"],
        _g + "italic": _e["italic"],
        _g + "alignment": _ALIGNMENT_NAME.get(_e["alignment"], "left"),
        _g + "line_spacing": _ini_number(_e["line_spacing"] / 100),
        _g + "space_before": _ini_number(_e["space_before"]),
        _g + "space_after": _ini_number(_e["space_after"]),
        _g + "indent_left": _ini_number(_e["indent_left"]),
        _g + "indent_right": _ini_number(_e["indent_right"]),
        _g + "indent_first_line": _ini_number(_e["first_line_indent"]),
    })
DEFAULT_CONFIG["Styles.Dialogue/name"] = ""
DEFAULT_CONFIG["Styles.Dialogue/symbol"] = DASH
DEFAULT_CONFIG["Styles.SceneBreak/symbol"] = SEPARATOR

# settings.ini from earlier versions (in Spanish): when opened, each value
# moves to its English key
_OLD_KEYS = {
    "Pagina/ancho": "Page/width", "Pagina/alto": "Page/height",
    "Pagina/margen_superior": "Page/margin_top",
    "Pagina/margen_inferior": "Page/margin_bottom",
    "Pagina/margen_izquierdo": "Page/margin_left",
    "Pagina/margen_derecho": "Page/margin_right",
    "Vista/modo": "View/mode", "Vista/zoom": "View/zoom",
    "Vista/contraer_titulos": "View/collapse_headings",
    "Ventana/x": "Window/x", "Ventana/y": "Window/y",
    "Ventana/ancho": "Window/width", "Ventana/alto": "Window/height",
    "Ventana/maximizada": "Window/maximized",
    "Paneles/navegador": "Panels/navigator", "Paneles/comentarios": "Panels/comments",
    "Archivos/ultima_carpeta": "Files/last_folder",
    "Archivos/carpeta_imagenes": "Files/images_folder",
    "Respaldo/activo": "Backup/enabled",
    "Respaldo/espera_segundos": "Backup/wait_seconds",
    "Respaldo/maximo_segundos": "Backup/max_seconds",
}
_OLD_SECTIONS = {
    "Cuerpo": "Body", "Titulo": "Title", "Subtitulo": "Subtitle", "Seccion": "Section",
    "Pensamientos": "Thoughts", "Cita": "Quote", "Separador": "SceneBreak",
    "Dialogo": "Dialogue",
}
_OLD_FIELDS = {
    "nombre": "name", "fuente": "font", "tamano": "size", "color": "color",
    "negrita": "bold", "cursiva": "italic", "alineacion": "alignment",
    "interlineado": "line_spacing", "espacio_antes": "space_before",
    "espacio_despues": "space_after", "sangria_izquierda": "indent_left",
    "sangria_derecha": "indent_right", "sangria_primera_linea": "indent_first_line",
    "simbolo": "symbol",
}
_OLD_VALUES = {"izquierda": "left", "centro": "center", "derecha": "right",
                     "justificado": "justify", **OLD_VIEWS}
# Style names the earlier version wrote by default: they are not a user change,
# so they are left empty (= the name for the language)
_OLD_NAMES = {"Cuerpo", "Título", "Subtítulo", "Sección", "Pensamientos",
                     "Diálogo", "Cita", "Separador"}


def load_styles(config):
    """Take the styles from settings.ini. The dictionaries are changed in place
    (the buttons and the editor use them directly), so it must be called
    before creating the editor. A value that can't be understood or is out
    of range keeps the usual one."""
    global DASH, SEPARATOR

    def in_range(v, minimum, maximum, default):
        return v if minimum <= v <= maximum else default

    for section, _, style in INI_STYLES:
        d = DEFAULT_STYLES[section]
        g = f"Styles.{section}/"
        font = config.value(g + "font").strip()
        style["font_family"] = font or d["font_family"]
        style["font_size"] = in_range(config.value(g + "size"), 1, 400, d["font_size"])
        color = config.value(g + "color").strip()
        style["font_color"] = QColor(color).name().upper() \
            if QColor(color).isValid() else d["font_color"]
        style["bold"] = config.value(g + "bold")
        style["italic"] = config.value(g + "italic")
        style["alignment"] = ALIGNMENTS.get(
            config.value(g + "alignment").strip().lower(), d["alignment"])
        style["line_spacing"] = round(
            in_range(config.value(g + "line_spacing"), 0.5, 5, d["line_spacing"] / 100) * 100)
        style["space_before"] = in_range(config.value(g + "space_before"),
                                       0, SPACE_MAX, d["space_before"])
        style["space_after"] = in_range(config.value(g + "space_after"),
                                      0, SPACE_MAX, d["space_after"])
        style["indent_left"] = in_range(config.value(g + "indent_left"),
                                      0, INDENT_MAX, d["indent_left"])
        style["indent_right"] = in_range(config.value(g + "indent_right"),
                                       0, INDENT_MAX, d["indent_right"])
        style["first_line_indent"] = in_range(config.value(g + "indent_first_line"),
                                            0, FIRST_LINE_MAX, d["first_line_indent"])
    # Visible name of each style (empty = the one for the language)
    STYLE_NAMES.clear()
    for section, key in [(s, c) for s, c, _ in INI_STYLES] + [("Dialogue", "dialogue")]:
        text = " ".join(config.value(f"Styles.{section}/name").split())
        if text:
            STYLE_NAMES[key] = text[:STYLE_NAME_MAX]
    DASH = config.value("Styles.Dialogue/symbol") or "\u2014"
    SEPARATOR = config.value("Styles.SceneBreak/symbol").strip() or "* * *"


_PAGE_KEYS_INI = {"width": "Page/width", "height": "Page/height",
                  "top": "Page/margin_top", "bottom": "Page/margin_bottom",
                  "left": "Page/margin_left", "right": "Page/margin_right"}


def settings_path():
    folder = os.path.dirname(os.path.abspath(sys.argv[0] or __file__))
    if os.path.isdir(folder) and os.access(folder, os.W_OK):
        return os.path.join(folder, "settings.ini")
    return os.path.join(app_folder(), "settings.ini")


class Settings:
    """Read and write settings.ini. A missing or unreadable value takes the
    default."""

    def __init__(self, path=None):
        self.path = path or settings_path()
        self.settings_xml = QSettings(self.path, QSettings.IniFormat)
        self._migrate()
        for key, value in DEFAULT_CONFIG.items():   # fill in the file
            if not self.settings_xml.contains(key):
                self.settings_xml.setValue(key, value)
        self.settings_xml.sync()

    def _migrate(self):
        """Move the Spanish keys from earlier versions to the new (English)
        ones and remove the old ones."""
        for old in self.settings_xml.allKeys():
            new_list = _OLD_KEYS.get(old)
            m = re.fullmatch(r"Estilos\.(\w+)/(\w+)", old)
            if m and m.group(1) in _OLD_SECTIONS and m.group(2) in _OLD_FIELDS:
                new_list = f"Styles.{_OLD_SECTIONS[m.group(1)]}/" \
                        f"{_OLD_FIELDS[m.group(2)]}"
            if new_list is None:
                continue
            value = self.settings_xml.value(old)
            if isinstance(value, str):
                if new_list.endswith("/name") and value.strip() in _OLD_NAMES:
                    value = ""
                else:
                    value = _OLD_VALUES.get(value.strip().lower(), value)
            if not self.settings_xml.contains(new_list):
                self.settings_xml.setValue(new_list, value)
            self.settings_xml.remove(old)

    def value(self, key):
        default = DEFAULT_CONFIG[key]
        v = self.settings_xml.value(key, default)
        if isinstance(default, bool):
            return str(v).strip().lower() in ("true", "1", "yes", "si", "sí")
        if isinstance(default, (int, float)):
            try:
                n = float(str(v).replace(",", "."))
            except ValueError:
                return default
            if isinstance(default, int) and n != int(n) and key.startswith("Styles."):
                return n                   # e.g., size=11.5
            return int(n) if isinstance(default, int) else n
        return str(v) if v is not None else default

    def set_value(self, key, value):
        self.settings_xml.setValue(key, value)

    def save(self):
        self.settings_xml.sync()

    # -- groups ------------------------------------------------------------
    def language(self):
        code = self.value("Interface/language").strip().lower()
        return code if code in dict(LANGUAGES) else "en"

    def page(self):
        c = {k: self.value(key) for k, key in _PAGE_KEYS_INI.items()}
        # Impossible measurements (file edited by hand): the usual ones
        valid = 5 <= c["width"] <= 100 and 5 <= c["height"] <= 100 and \
            all(0 <= c[k] <= 20 for k in ("top", "bottom", "left", "right")) and \
            c["left"] + c["right"] <= c["width"] - 2 and c["top"] + c["bottom"] <= c["height"] - 2
        return c if valid else dict(PAGE)

    def set_page(self, config):
        for k, key in _PAGE_KEYS_INI.items():
            self.set_value(key, round(float(config[k]), 2))

    def mode(self):
        m = self.value("View/mode").strip().lower()
        m = OLD_VIEWS.get(m, m)
        return m if m in (VIEW_PAGES, VIEW_NO_PAGES, VIEW_MIXED) else VIEW_PAGES


def _draw_app_icon(p, detail):
    """Draw the program icon on a 64 × 64 grid: a book with a blue cover and
    its pages, a bookmark ribbon and a pen nib on the cover (writing a
    book)."""
    # Pages (peeking out to the right of and below the cover)
    p.setPen(QPen(QColor("#B9AE98"), 1.2))
    p.setBrush(QColor("#FBF7EE"))
    p.drawRoundedRect(QRectF(13, 7.5, 39, 50), 3, 3)
    if detail:                       # page edges
        p.setPen(QPen(QColor("#D8CFBD"), 0.8))
        for k in range(1, 4):
            p.drawLine(QPointF(49 + k * 0.8, 10 + k), QPointF(49 + k * 0.8, 52 + k))
            p.drawLine(QPointF(16 + k, 54 + k * 0.8), QPointF(49, 54 + k * 0.8))
    # Bookmark ribbon (coming out at the bottom, between the pages)
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#E69138"))
    p.drawPolygon(QPolygonF([QPointF(37, 44), QPointF(43, 44), QPointF(43, 62),
                             QPointF(40, 58.5), QPointF(37, 62)]))
    # Cover with a blue gradient
    gradient = QLinearGradient(QPointF(10, 4), QPointF(48, 54))
    gradient.setColorAt(0, QColor("#5B9BEF"))
    gradient.setColorAt(1, QColor("#2D68B0"))
    p.setPen(QPen(QColor("#22528C"), 1.2))
    p.setBrush(gradient)
    p.drawRoundedRect(QRectF(9, 4, 39, 50), 4, 4)
    # Spine
    spine = QPainterPath()
    spine.addRoundedRect(QRectF(9, 4, 39, 50), 4, 4)
    strip = QPainterPath()
    strip.addRect(QRectF(9, 4, 7, 50))
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(20, 60, 110, 90))
    p.drawPath(spine.intersected(strip))
    p.setPen(QPen(QColor(255, 255, 255, 70), 1))
    p.drawLine(QPointF(16.5, 5), QPointF(16.5, 53))
    # White pen nib in the center of the cover
    cx = 32.5
    nib = QPainterPath()
    nib.moveTo(cx - 7, 16)
    nib.lineTo(cx + 7, 16)
    nib.lineTo(cx + 7, 24)
    nib.cubicTo(cx + 7, 31, cx + 2, 37, cx, 43)
    nib.cubicTo(cx - 2, 37, cx - 7, 31, cx - 7, 24)
    nib.closeSubpath()
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#FFFFFF"))
    p.drawPath(nib)
    if detail:                       # nib slit and breather hole
        p.setBrush(QColor("#2D68B0"))
        p.drawEllipse(QPointF(cx, 27), 2.1, 2.1)
        slit = QPen(QColor("#2D68B0"), 1.4)
        slit.setCapStyle(Qt.RoundCap)
        p.setPen(slit)
        p.drawLine(QPointF(cx, 29), QPointF(cx, 40))
        p.setPen(QPen(QColor("#2D68B0"), 1.2))
        p.drawLine(QPointF(cx - 7, 19.5), QPointF(cx + 7, 19.5))


def app_icon():
    """Program icon, drawn in several sizes so it looks sharp in the title bar,
    the taskbar and the desktop."""
    icon = QIcon()
    for size in (16, 20, 24, 32, 40, 48, 64, 96, 128, 256):
        pm = QPixmap(size, size)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.scale(size / 64, size / 64)
        _draw_app_icon(p, detail=size >= 32)
        p.end()
        icon.addPixmap(pm)
    return icon


def change_color_icon(color):
    """Tipped paint can with a drop and a strip of the current color (the
    "change color" icon of word processors)."""
    scale_factor = 3
    size = 24
    pm = QPixmap(size * scale_factor, size * scale_factor)
    pm.setDevicePixelRatio(scale_factor)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    ink = QColor(color)
    stroke = QPen(QColor(ICON_COLOR), 1.5)
    stroke.setJoinStyle(Qt.RoundJoin)
    stroke.setCapStyle(Qt.RoundCap)
    # Tilted can (a rhombus), with the paint in its lower half
    can = QPolygonF([QPointF(4, 9.5), QPointF(10.5, 3), QPointF(17, 9.5),
                      QPointF(10.5, 16)])
    paint = QPolygonF([QPointF(4, 9.5), QPointF(17, 9.5), QPointF(10.5, 16)])
    p.setPen(Qt.NoPen)
    p.setBrush(ink)
    p.drawPolygon(paint)
    p.setPen(stroke)
    p.setBrush(Qt.NoBrush)
    p.drawPolygon(can)
    # Handle
    p.drawLine(QPointF(7.2, 6.3), QPointF(4.5, 3.6))
    # Drop falling from the can
    drop = QPainterPath()
    drop.moveTo(19.5, 10)
    drop.cubicTo(20.5, 12, 21.5, 13.2, 21.5, 14.3)
    drop.arcTo(QRectF(17.5, 12.3, 4, 4), 0, -180)
    drop.cubicTo(17.5, 13.2, 18.5, 12, 19.5, 10)
    p.setPen(QPen(QColor(ICON_COLOR), 1))
    p.setBrush(ink)
    p.drawPath(drop)
    # Strip with the current color
    p.setPen(QPen(QColor("#9A9A9A"), 0.6) if ink.lightness() > 190 else Qt.NoPen)
    p.setBrush(ink)
    p.drawRoundedRect(QRectF(3, 19, 18, 3.5), 1, 1)
    p.end()
    return QIcon(pm)


def file_icon(kind):
    """Line icons for New, Open, Save, Save as and Settings (same stroke and
    color as the paragraph ones)."""
    scale_factor = 3
    size = 24
    pm = QPixmap(size * scale_factor, size * scale_factor)
    pm.setDevicePixelRatio(scale_factor)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    color = QColor(ICON_COLOR)
    pen = QPen(color, 1.6)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen)
    p.setBrush(Qt.NoBrush)

    def sheet(x0, y0, x1, y1, fold=4.5):
        """Sheet with the top-right corner folded."""
        p.drawPolygon(QPolygonF([
            QPointF(x0, y0), QPointF(x1 - fold, y0), QPointF(x1, y0 + fold),
            QPointF(x1, y1), QPointF(x0, y1)]))
        p.drawPolyline(QPolygonF([
            QPointF(x1 - fold, y0), QPointF(x1 - fold, y0 + fold),
            QPointF(x1, y0 + fold)]))

    def floppy(x0, y0, x1, y1):
        width = x1 - x0
        p.drawPolygon(QPolygonF([
            QPointF(x0, y0), QPointF(x1 - 3, y0), QPointF(x1, y0 + 3),
            QPointF(x1, y1), QPointF(x0, y1)]))
        # Label (bottom) and shutter (top)
        p.drawRect(QRectF(x0 + width * 0.2, y1 - (y1 - y0) * 0.42,
                          width * 0.6, (y1 - y0) * 0.42))
        p.drawLine(QPointF(x0 + width * 0.28, y0), QPointF(x0 + width * 0.28, y0 + 4.5))
        p.drawLine(QPointF(x0 + width * 0.28, y0 + 4.5), QPointF(x1 - width * 0.3, y0 + 4.5))
        p.drawLine(QPointF(x1 - width * 0.3, y0 + 4.5), QPointF(x1 - width * 0.3, y0))

    if kind == "new":
        sheet(5, 3, 19, 21)
        p.drawLine(QPointF(12, 11), QPointF(12, 17))
        p.drawLine(QPointF(9, 14), QPointF(15, 14))
    elif kind == "open":
        # Folder: back part with a tab and a slanted front cover
        p.drawPolyline(QPolygonF([
            QPointF(4, 18), QPointF(4, 6), QPointF(9, 6), QPointF(11, 8),
            QPointF(18, 8), QPointF(18, 10.5)]))
        p.drawPolygon(QPolygonF([
            QPointF(4, 18), QPointF(7, 10.5), QPointF(21, 10.5), QPointF(18, 18)]))
    elif kind == "save":
        floppy(4, 4, 20, 20)
    elif kind == "save_as":
        floppy(3, 3, 16, 16)
        # Diagonal pencil
        pencil = QPen(color, 1.6)
        pencil.setCapStyle(Qt.RoundCap)
        p.setPen(pencil)
        p.setBrush(QColor("#FFFFFF"))
        p.drawPolygon(QPolygonF([
            QPointF(13, 19.5), QPointF(13.6, 16.6), QPointF(19.3, 10.9),
            QPointF(21.6, 13.2), QPointF(15.9, 18.9)]))
        p.drawLine(QPointF(18.1, 12.1), QPointF(20.4, 14.4))
    elif kind == "cut":
        # Scissors: two crossed blades and two handles
        p.drawLine(QPointF(7.5, 3.5), QPointF(15.2, 14.8))
        p.drawLine(QPointF(16.5, 3.5), QPointF(8.8, 14.8))
        p.drawEllipse(QPointF(7, 17.5), 3, 3)
        p.drawEllipse(QPointF(17, 17.5), 3, 3)
    elif kind == "copy":
        # Two sheets: the back one peeks out at the top left
        p.drawPolyline(QPolygonF([QPointF(5, 16), QPointF(5, 3.5), QPointF(14, 3.5)]))
        sheet(8.5, 7, 19.5, 21, 3.5)
    elif kind == "paste":
        # Clipboard with its clip and a sheet on top
        p.drawPolyline(QPolygonF([QPointF(8, 5), QPointF(4.5, 5), QPointF(4.5, 21),
                                  QPointF(10, 21)]))
        p.drawPolyline(QPolygonF([QPointF(16, 5), QPointF(19.5, 5), QPointF(19.5, 9)]))
        p.drawRoundedRect(QRectF(8, 3, 8, 4), 1, 1)
        sheet(11, 11, 21, 22, 3)
    elif kind == "find":
        p.drawEllipse(QRectF(4, 4, 11, 11))
        thick = QPen(color, 2.4)
        thick.setCapStyle(Qt.RoundCap)
        p.setPen(thick)
        p.drawLine(QPointF(13.6, 13.6), QPointF(19.5, 19.5))
    elif kind == "navigator":
        # Outline: bullets with lines indented by level
        for y, x, length in ((5.5, 4, 15), (11, 8, 11), (16.5, 8, 11)):
            p.setPen(Qt.NoPen)
            p.setBrush(color)
            p.drawEllipse(QPointF(x, y), 1.6, 1.6)
            p.setPen(pen)
            p.drawLine(QPointF(x + 3.5, y), QPointF(x + 3.5 + length - 4, y))
        p.setPen(QPen(color, 1.1))
        p.drawLine(QPointF(4, 8), QPointF(4, 16.5))
        p.drawLine(QPointF(4, 11), QPointF(6, 11))
        p.drawLine(QPointF(4, 16.5), QPointF(6, 16.5))
    elif kind == "image":
        p.drawRoundedRect(QRectF(3.5, 5, 17, 14), 2, 2)
        p.drawPolyline(QPolygonF([QPointF(4, 17.5), QPointF(9.5, 11.5), QPointF(13, 15),
                                  QPointF(15.5, 12.5), QPointF(20, 17.5)]))
        p.setPen(Qt.NoPen)
        p.setBrush(color)
        p.drawEllipse(QPointF(15.5, 9), 1.6, 1.6)
    elif kind in ("comment", "comments"):
        balloon = QPainterPath()
        balloon.addRoundedRect(QRectF(3.5, 4, 17, 12), 3, 3)
        tail = QPainterPath()
        tail.moveTo(7.5, 15)
        tail.lineTo(7, 20.5)
        tail.lineTo(12.5, 15)
        tail.closeSubpath()
        p.drawPath(balloon.united(tail))
        if kind == "comment":       # balloon with "+": new comment
            p.drawLine(QPointF(12, 7), QPointF(12, 13))
            p.drawLine(QPointF(9, 10), QPointF(15, 10))
        else:                          # balloon with lines: the panel
            p.drawLine(QPointF(7.5, 8.2), QPointF(16.5, 8.2))
            p.drawLine(QPointF(7.5, 11.8), QPointF(13.5, 11.8))
    elif kind == "delete":
        # Trash can: lid with a handle, tapering body and three lines
        p.drawLine(QPointF(4.5, 6.5), QPointF(19.5, 6.5))
        p.drawPolyline(QPolygonF([QPointF(9.5, 6.5), QPointF(9.5, 4),
                                  QPointF(14.5, 4), QPointF(14.5, 6.5)]))
        p.drawPolygon(QPolygonF([QPointF(6.5, 6.5), QPointF(17.5, 6.5),
                                 QPointF(16.5, 20.5), QPointF(7.5, 20.5)]))
        for x in (10, 12, 14):
            p.drawLine(QPointF(x, 10), QPointF(x, 17.5))
    elif kind == "settings":
        # Gear: wheel with 8 teeth and a hole in the center
        center = QPointF(12, 12)
        wheel = QPainterPath()
        wheel.addEllipse(center, 6.2, 6.2)
        for k in range(8):
            tooth = QPainterPath()
            tooth.addRoundedRect(QRectF(-1.7, -9.3, 3.4, 4.5), 0.8, 0.8)
            rotation = QTransform()
            rotation.translate(center.x(), center.y())
            rotation.rotate(k * 45)
            wheel = wheel.united(rotation.map(tooth))
        gap = QPainterPath()
        gap.addEllipse(center, 2.6, 2.6)
        p.drawPath(wheel.subtracted(gap))
    p.end()
    return QIcon(pm)


class FileBar(ToolContainer):
    """New, Open, Save, Save as, editing, Find, Navigator, images, comments and
    Settings."""
    def __init__(self, window):
        super().__init__("File", window)
        self.actions = {}
        self.setMovable(False)
        self.setIconSize(QSize(24, 24))
        for icon, text, shortcut, shortcut_label, slot in (
                ("new", _t("New"), QKeySequence.New, "Ctrl+N", window.new),
                ("open", _t("Open"), QKeySequence.Open, "Ctrl+O", window.open_file),
                ("save", _t("Save"), QKeySequence.Save, "Ctrl+S", window.save),
                ("save_as", _t("Save as"), QKeySequence("Ctrl+Shift+S"),
                 "Ctrl+Shift+S", window.save_as),
                # The editor handles the Cut, Copy and Paste shortcuts
                ("cut", _t("Cut"), None, "Ctrl+X", window.cut_selection),
                ("copy", _t("Copy"), None, "Ctrl+C", window.copy_selection),
                ("paste", _t("Paste"), None, "Ctrl+V", window.paste_clipboard),
                ("find", _t("Find and replace"), QKeySequence.Find, "Ctrl+F",
                 window.find_text),
                ("navigator", _t("Navigator: titles, subtitles and sections"),
                 QKeySequence("F5"), "F5", window.toggle_navigator),
                ("image", _t("Insert image (top and bottom, centered)"), None, None,
                 window.insert_image),
                ("comment", _t("New comment on the selected text"),
                 QKeySequence("Ctrl+Alt+M"), "Ctrl+Alt+M", window.new_comment),
                ("comments", _t("Comments panel"), QKeySequence("F6"), "F6",
                 window.toggle_comments),
                ("settings", _t("Settings (page, margins, view, headings and language)"), None,
                 None, window.open_settings)):
            if icon in ("cut", "find", "settings"):
                self.addSeparator()
            a = QAction(file_icon(icon), text, self)
            self.addAction(a)
            self.actions[icon] = a
            a.setToolTip(f"{text} ({shortcut_label})" if shortcut_label else text)
            if shortcut is not None:
                a.setShortcut(shortcut)
                a.setShortcutContext(Qt.WindowShortcut)
            a.triggered.connect(slot)
            if icon in ("cut", "copy", "paste"):
                setattr(self, f"action_{icon}", a)
            if icon == "navigator":
                a.setCheckable(True)
                self.action_navigator = a
            elif icon == "comments":
                a.setCheckable(True)
                self.action_comments = a
        # Ctrl+H opens "Find and replace" directly
        replace = QAction(_t("Replace"), self)
        replace.setShortcut(QKeySequence("Ctrl+H"))
        replace.setShortcutContext(Qt.WindowShortcut)
        replace.triggered.connect(lambda: window.find_text(replace=True))
        window.addAction(replace)


# ---------------------------------------------------------------------------
# Ribbon: all the tools in a single strip, in groups
# ---------------------------------------------------------------------------
def _vertical_line():
    line = QFrame()
    line.setFrameShape(QFrame.VLine)
    line.setStyleSheet("color: #D6D6D6;")
    return line


class Ribbon(QWidget):
    """File and editing (3 × 4) | font, paragraph and spacing | styles (2 × 4)
    | Insert image."""
    ROW_HEIGHT = 30

    def __init__(self, file, font, paragraph, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(6, 4, 6, 4)
        lay.setSpacing(10)

        # -- File, search, comments and editing -------------------------------
        g1 = QGridLayout()
        g1.setSpacing(2)
        for f, row in enumerate((("new", "open", "save", "save_as"),
                                  ("find", "navigator", "comment", "comments"),
                                  ("paste", "copy", "cut", "settings"))):
            for c, key in enumerate(row):
                g1.addWidget(self._button(file.actions[key], 22), f, c)
        lay.addLayout(g1)
        lay.addWidget(_vertical_line())

        # -- Center: font, paragraph and styles in rows 1 and 2; spacing and indent in row 3, from side to side ---
        center = QVBoxLayout()
        center.setSpacing(2)
        top = QHBoxLayout()
        top.setSpacing(8)
        left = QVBoxLayout()
        left.setSpacing(2)
        f1 = self._row()
        f1.addWidget(font.font_combo)
        f1.addWidget(font.size_combo)
        f1.addSpacing(4)
        for a in (font.act_bigger, font.act_smaller):
            f1.addWidget(self._button(a))
        f1.addStretch()
        f2 = self._row()
        for a in (font.act_bold, font.act_italic, font.act_underline,
                  font.act_strike):
            f2.addWidget(self._button(a))
        f2.addSpacing(6)
        for a in (font.act_sub, font.act_sup):
            f2.addWidget(self._button(a))
        f2.addSpacing(6)
        for b in (font.btn_color, font.btn_highlight):
            b.setFocusPolicy(Qt.NoFocus)
            b.setAutoRaise(True)
            b.setFixedHeight(self.ROW_HEIGHT)
            b.setIconSize(QSize(20, 20))
            f2.addWidget(b)
        f2.addWidget(self._button(font.act_clear))
        f2.addStretch()
        left.addLayout(f1)
        left.addLayout(f2)
        top.addLayout(left)
        top.addWidget(_vertical_line())

        right = QGridLayout()
        right.setSpacing(2)
        for c, a in enumerate(font.alignment_list):
            right.addWidget(self._button(a), 0, c)
        for c, a in enumerate((font.act_decrease, font.act_increase,
                               font.act_numbers, font.act_bullets)):
            right.addWidget(self._button(a), 1, c)
        top.addLayout(right)
        top.addWidget(_vertical_line())

        styles = QGridLayout()                    # 2 rows of 4
        styles.setSpacing(0)
        for k, key in enumerate(STYLE_ORDER):
            styles.addWidget(font.swatches[key][0], k // 4, k % 4)
        # The swatches stretch to the end of the spacing row (leaving no gap on
        # the right)
        top.addLayout(styles, 1)
        center.addLayout(top)

        f3 = self._row()
        for k, (text, w) in enumerate(((_t("Line spacing"), paragraph.line_spacing_combo),
                                        (_t("Before"), paragraph.spin_before),
                                        (_t("After"), paragraph.spin_after),
                                        (_t("Letters"), paragraph.spin_letters))):
            if k:
                f3.addSpacing(10)
            f3.addWidget(self._label(text))
            f3.addWidget(w)
        f3.addSpacing(12)
        f3.addWidget(_vertical_line())
        f3.addSpacing(4)
        title = QLabel(_t("Indent"))
        title.setObjectName("groupTitle")
        f3.addWidget(title)
        f3.addSpacing(4)
        for k, (text, w) in enumerate(((_t("indent|Left"), paragraph.spin_left_indent),
                                        (_t("indent|Right"), paragraph.spin_right_indent),
                                        (_t("First line"), paragraph.spin_first_line))):
            if k:
                f3.addSpacing(10)
            f3.addWidget(self._label(text))
            f3.addWidget(w)
        f3.addStretch()
        center.addLayout(f3)
        lay.addLayout(center)
        lay.addWidget(_vertical_line())

        # -- Insert image ----------------------------------------------------
        action = file.actions["image"]
        action.setText(_t("Insert\nimage"))
        big = QToolButton()
        big.setDefaultAction(action)
        big.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        big.setIconSize(QSize(46, 46))
        big.setAutoRaise(True)
        big.setFocusPolicy(Qt.NoFocus)
        big.setStyleSheet("QToolButton { font-weight: bold; color: #3A3A3A;"
                             " padding: 4px 8px; }")
        lay.addWidget(big, 0, Qt.AlignVCenter)
        lay.addStretch()

    def _row(self):
        f = QHBoxLayout()
        f.setSpacing(2)
        return f

    def _button(self, action, icon=20):
        b = QToolButton()
        b.setDefaultAction(action)
        b.setAutoRaise(True)
        b.setIconSize(QSize(icon, icon))
        b.setFixedSize(self.ROW_HEIGHT, self.ROW_HEIGHT)
        b.setFocusPolicy(Qt.NoFocus)        # the focus stays in the text
        return b

    @staticmethod
    def _label(text):
        lbl = QLabel(text)
        lbl.setObjectName("fieldLabel")
        return lbl


# ---------------------------------------------------------------------------
# Find / Find and replace window
# ---------------------------------------------------------------------------
MATCH_COLOR = "#FFE98A"        # all matches
CURRENT_MATCH_COLOR = "#FFA94D"  # the selected one
MAX_HIGHLIGHTED = 5000                 # limit so huge books don't slow down


class FindDialog(QDialog):
    """Modeless window: you can keep typing with it open. Highlights all
    matches and shows "n of N"."""

    def __init__(self, editor, parent=None):
        super().__init__(parent)
        self.editor = editor
        self.setWindowTitle(_t("Find"))
        self.setModal(False)
        lay = QVBoxLayout(self)

        self.tabs = QTabBar()
        self.tabs.addTab(_t("Find"))
        self.tabs.addTab(_t("Find and replace"))
        self.tabs.setExpanding(False)
        self.tabs.setDrawBase(False)
        self.tabs.currentChanged.connect(self._change_mode)
        lay.addWidget(self.tabs)

        form = QGridLayout()
        form.setColumnStretch(1, 1)
        self.find_field = QLineEdit()
        self.find_field.setPlaceholderText(_t("Text to find"))
        self.find_field.setClearButtonEnabled(True)
        self.find_field.textChanged.connect(self._schedule)
        self.find_field.returnPressed.connect(self._enter_find)
        self.replace_field = QLineEdit()
        self.replace_field.setPlaceholderText(_t("Replace with"))
        self.replace_field.setClearButtonEnabled(True)
        self.replace_field.returnPressed.connect(self.replace)
        self.replace_label = QLabel(_t("Replace"))
        form.addWidget(QLabel(_t("Find")), 0, 0)
        form.addWidget(self.find_field, 0, 1)
        form.addWidget(self.replace_label, 1, 0)
        form.addWidget(self.replace_field, 1, 1)
        # The label column doesn't change width when "Replace" is hidden
        form.setColumnMinimumWidth(0, self.replace_label.sizeHint().width())
        lay.addLayout(form)

        options = QHBoxLayout()
        self.chk_case = QCheckBox(_t("Match case"))
        self.chk_word = QCheckBox(_t("Whole words"))
        for chk in (self.chk_case, self.chk_word):
            chk.toggled.connect(self._schedule)
            options.addWidget(chk)
        options.addStretch()
        lay.addLayout(options)

        self.status_label = QLabel(" ")
        self.status_label.setStyleSheet("color: #666666;")
        lay.addWidget(self.status_label)

        buttons = QHBoxLayout()
        self.btn_previous = QPushButton(_t("Previous"))
        self.btn_previous.setToolTip(_t("Previous match (Shift+Enter)"))
        self.btn_previous.clicked.connect(lambda: self.find_text(backward=True))
        self.btn_next = QPushButton(_t("Next"))
        self.btn_next.setToolTip(_t("Next match (Enter)"))
        self.btn_next.clicked.connect(lambda: self.find_text())
        self.btn_replace = QPushButton(_t("Replace"))
        self.btn_replace.setToolTip(
            _t("Replaces the selected match and moves to the next one"))
        self.btn_replace.clicked.connect(self.replace)
        self.btn_all = QPushButton(_t("Replace all"))
        self.btn_all.setToolTip(_t("Replaces all matches (a single undo step)"))
        self.btn_all.clicked.connect(self.replace_all)
        btn_close = QPushButton(_t("Close"))
        btn_close.clicked.connect(self.close)
        for b in (self.btn_previous, self.btn_next, self.btn_replace, self.btn_all):
            b.setAutoDefault(False)
            buttons.addWidget(b)
        buttons.addStretch()
        btn_close.setAutoDefault(False)
        buttons.addWidget(btn_close)
        lay.addStretch()      # in "Find", the extra space goes above the buttons
        lay.addLayout(buttons)

        # Recalculate the highlight shortly after typing stops
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(150)
        self._timer.timeout.connect(self.refresh_highlight)
        editor.document().contentsChanged.connect(self._doc_changed)
        editor.cursorPositionChanged.connect(self._show_status)

        self._matches = []           # [(start, end)]
        self._extra_status = ""            # e.g. "Continued from…"
        # Both tabs use the same size: that of "Find and replace"
        self._change_mode(1)
        size = self.sizeHint().expandedTo(QSize(460, 0))
        self.setMinimumSize(size)
        self.resize(size)
        self._change_mode(0)

    # -- opening ----------------------------------------------------------
    def open_file(self, replace=False):
        self.tabs.setCurrentIndex(1 if replace else 0)
        selection = self.editor.textCursor().selectedText()
        if selection and "\u2029" not in selection and len(selection) <= 200:
            self.find_field.setText(selection)
        self.show()
        self.raise_()
        self.activateWindow()
        self.find_field.setFocus()
        self.find_field.selectAll()
        self.refresh_highlight()

    def _change_mode(self, i):
        replacing = i == 1
        for w in (self.replace_label, self.replace_field,
                  self.btn_replace, self.btn_all):
            w.setVisible(replacing)
        self.setWindowTitle(_t("Find and replace") if replacing else _t("Find"))

    # -- searching --------------------------------------------------------
    def _flags(self, backward=False):
        f = QTextDocument.FindFlag(0)
        if self.chk_case.isChecked():
            f |= QTextDocument.FindCaseSensitively
        if self.chk_word.isChecked():
            f |= QTextDocument.FindWholeWords
        if backward:
            f |= QTextDocument.FindBackward
        return f

    def _enter_find(self):
        backward = bool(QApplication.keyboardModifiers() & Qt.ShiftModifier)
        self.find_text(backward=backward)

    def find_text(self, backward=False):
        text = self.find_field.text()
        if not text:
            return False
        self._extra_status = ""
        doc = self.editor.document()
        cursor = self.editor.textCursor()
        found = doc.find(text, cursor, self._flags(backward))
        if found.isNull():               # wrap around the document
            start_pos = QTextCursor(doc)
            if backward:
                start_pos.movePosition(QTextCursor.End)
            found = doc.find(text, start_pos, self._flags(backward))
            if not found.isNull():
                self._extra_status = _t("Continued from the end") if backward \
                    else _t("Continued from the beginning")
        if found.isNull():
            self.status_label.setText(_t("No matches"))
            QApplication.beep()
            return False
        self.editor.setTextCursor(found)
        self._paint()
        return True

    def _selection_matches(self):
        """Is the current selection exactly a match?"""
        c = self.editor.textCursor()
        return c.hasSelection() and \
            (c.selectionStart(), c.selectionEnd()) in set(self._matches)

    def replace(self):
        text = self.find_field.text()
        if not text:
            return
        self._recalculate()
        if self._selection_matches():
            c = self.editor.textCursor()
            c.insertText(self.replace_field.text())   # keeps the format
            self.editor.setTextCursor(c)
            self._recalculate()
        self.find_text()

    def replace_all(self):
        text = self.find_field.text()
        if not text:
            return
        self._recalculate()
        if not self._matches:
            self.status_label.setText(_t("No matches"))
            return
        new = self.replace_field.text()
        doc = self.editor.document()
        c = QTextCursor(doc)
        c.beginEditBlock()                     # a single undo step
        for start_pos, end in reversed(self._matches):   # from back to front
            c.setPosition(start_pos)
            c.setPosition(end, QTextCursor.KeepAnchor)
            c.insertText(new)
        c.endEditBlock()
        n = len(self._matches)
        self._recalculate()
        self._paint()
        self.status_label.setText(
            _t("Replaced 1 match") if n == 1 else _t("Replaced {n} matches", n=n))

    # -- highlight and status -----------------------------------------------
    def _schedule(self, *_):
        self._extra_status = ""
        self._timer.start()

    def _doc_changed(self):
        if self.isVisible():
            self._timer.start()

    def _recalculate(self):
        self._matches = []
        text = self.find_field.text()
        if not text:
            return
        doc = self.editor.document()
        c = QTextCursor(doc)
        flags = self._flags()
        while True:
            c = doc.find(text, c, flags)
            if c.isNull():
                break
            self._matches.append((c.selectionStart(), c.selectionEnd()))
            if c.selectionStart() == c.selectionEnd():
                break

    def refresh_highlight(self):
        if not self.isVisible():
            return
        self._recalculate()
        self._paint()

    def _paint(self):
        doc = self.editor.document()
        current = self.editor.textCursor()
        selections = []
        for start_pos, end in self._matches[:MAX_HIGHLIGHTED]:
            sel = QTextEdit.ExtraSelection()
            sel.cursor = QTextCursor(doc)
            sel.cursor.setPosition(start_pos)
            sel.cursor.setPosition(end, QTextCursor.KeepAnchor)
            is_current = (start_pos, end) == (current.selectionStart(), current.selectionEnd())
            sel.format.setBackground(QColor(
                CURRENT_MATCH_COLOR if is_current else MATCH_COLOR))
            selections.append(sel)
        self.editor.setExtraSelections(selections)
        self._show_status()

    def _show_status(self):
        if not self.isVisible():
            return
        if not self.find_field.text():
            self.status_label.setText(" ")
            return
        n = len(self._matches)
        if n == 0:
            self.status_label.setText(_t("No matches"))
            return
        c = self.editor.textCursor()
        try:
            i = self._matches.index((c.selectionStart(), c.selectionEnd())) + 1
            text = _t("{i} of {n} matches", i=i, n=n)
        except ValueError:
            text = _t("1 match") if n == 1 else _t("{n} matches", n=n)
        extra = self._extra_status
        self.status_label.setText(f"{text} · {extra}" if extra else text)

    def hideEvent(self, e):
        self.editor.setExtraSelections([])     # remove the highlight when closing
        super().hideEvent(e)

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape:
            self.close()
            return
        super().keyPressEvent(e)


# ---------------------------------------------------------------------------
# Navigator: outline of headings to jump through the book
# ---------------------------------------------------------------------------
class NavigatorPanel(QDockWidget):
    """Lists Titles, Subtitles and Sections as a tree. A click goes to the
    heading; the one that contains the cursor is marked. It updates by
    itself while typing."""
    @property
    def NAMES(self):
        return {1: style_name("title"), 2: style_name("subtitle"),
                3: style_name("section")}

    def __init__(self, editor, view, parent=None):
        super().__init__(_t("Navigator"), parent)
        self.editor = editor
        self.view = view
        self.setObjectName("navigator")
        self.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.setFeatures(QDockWidget.DockWidgetClosable)

        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.setSpacing(6)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText(_t("Filter headings"))
        self.filter_edit.setClearButtonEnabled(True)
        self.filter_edit.textChanged.connect(self._filter)
        lay.addWidget(self.filter_edit)
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setIndentation(14)
        self.tree.setUniformRowHeights(True)
        self.tree.setExpandsOnDoubleClick(False)
        self.tree.itemClicked.connect(self._go)
        self.tree.itemActivated.connect(self._go)      # Enter
        lay.addWidget(self.tree, 1)
        self.empty_label = QLabel(_t(
            "No headings yet.\nUse the {title}, {subtitle} or {section} styles "
            "to see them here.", title=style_name("title"),
            subtitle=style_name("subtitle"), section=style_name("section")))
        self.empty_label.setWordWrap(True)
        self.empty_label.setStyleSheet("color: #808080; padding: 8px;")
        lay.addWidget(self.empty_label)
        self.setWidget(content)
        self.setMinimumWidth(200)

        self._headings = []        # [(position, level, text)]
        self._items = []          # QTreeWidgetItem per heading, in order
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(300)
        self._timer.timeout.connect(self.refresh)
        editor.document().contentsChanged.connect(self._schedule)
        editor.cursorPositionChanged.connect(self._mark_current)
        self.visibilityChanged.connect(lambda v: v and self.refresh())

    # -- content ----------------------------------------------------------
    def _schedule(self):
        if self.isVisible():
            self._timer.start()

    def _read_headings(self):
        headings = []
        b = self.editor.document().begin()
        while b.isValid():
            level = b.blockFormat().headingLevel()
            if level > 0:
                text = " ".join(b.text().replace("\u2028", " ").split())
                headings.append((b.position(), level, text))
            b = b.next()
        return headings

    def refresh(self):
        """Re-read the headings. If only their positions changed (typing in
        normal text), don't rebuild the tree: that way whatever was folded
        in it is not lost."""
        self._timer.stop()
        headings = self._read_headings()
        if [t[1:] for t in headings] == [t[1:] for t in self._headings]:
            self._headings = headings
            self._mark_current()
            return
        # Folded branches in the tree (by level and text), to keep them folded
        folded = {(self._headings[i][1], it.text(0)) for i, it in enumerate(self._items)
                    if i < len(self._headings) and it.childCount() and not it.isExpanded()}
        self._headings = headings
        self.tree.clear()
        self._items = []
        stack = []                 # [(level, item)] open ancestors
        bold = QFont(self.tree.font())
        bold.setBold(True)
        for i, (_, level, text) in enumerate(headings):
            while stack and stack[-1][0] >= level:
                stack.pop()
            item = QTreeWidgetItem([text or _t("(no text)")])
            item.setData(0, Qt.UserRole, i)
            item.setToolTip(0, f"{self.NAMES.get(level, _t('Heading {n}', n=level))}: {text}")
            if level == 1:
                item.setFont(0, bold)
            if not text:
                item.setForeground(0, QColor("#9A9A9A"))
            if stack:
                stack[-1][1].addChild(item)
            else:
                self.tree.addTopLevelItem(item)
            stack.append((level, item))
            self._items.append(item)
        self.tree.expandAll()
        for i, it in enumerate(self._items):
            if (headings[i][1], it.text(0)) in folded:
                it.setExpanded(False)
        self.tree.setVisible(bool(headings))
        self.empty_label.setVisible(not headings)
        self._filter(self.filter_edit.text())
        self._mark_current()

    # -- navigation -------------------------------------------------------
    def _go(self, item, _column=0):
        if self._timer.isActive():         # positions up to date
            self.refresh()
        i = item.data(0, Qt.UserRole)
        if i is None or i >= len(self._headings):
            return
        block = self.editor.document().findBlock(self._headings[i][0])
        if not block.isValid():
            return
        self.editor.setTextCursor(QTextCursor(block))   # expands it if it was hidden
        QApplication.processEvents()
        # Leave the heading near the top edge of the view
        r = self.editor.cursorRect()
        top = self.view.mapFromScene(self.view.proxy.mapToScene(QPointF(r.topLeft()))).y()
        bar = self.view.verticalScrollBar()
        bar.setValue(bar.value() + top - 48)
        self.editor.setFocus()

    def _mark_current(self):
        """Mark the heading the cursor is under."""
        if not self.isVisible() or not self._headings:
            return
        pos = self.editor.textCursor().position()
        positions = [t[0] for t in self._headings]
        i = bisect.bisect_right(positions, pos) - 1
        self.tree.blockSignals(True)
        if i >= 0 and i < len(self._items):
            # If it is inside a folded branch, mark the branch (marking the
            # heading would unfold it)
            target_block = self._items[i]
            parent_item = target_block.parent()
            while parent_item is not None:
                if not parent_item.isExpanded():
                    target_block = parent_item
                parent_item = parent_item.parent()
            self.tree.setCurrentItem(target_block)
            self.tree.scrollToItem(target_block)
        else:
            self.tree.setCurrentItem(None)
            self.tree.clearSelection()
        self.tree.blockSignals(False)

    def _filter(self, text):
        text = text.strip().lower()

        def show_item(item):
            child_visible = False
            for j in range(item.childCount()):
                child_visible |= show_item(item.child(j))
            matches = not text or text in item.text(0).lower()
            item.setHidden(not (matches or child_visible))
            return matches or child_visible

        for j in range(self.tree.topLevelItemCount()):
            show_item(self.tree.topLevelItem(j))
        if text:
            self.tree.expandAll()


# ---------------------------------------------------------------------------
# Comments panel
# ---------------------------------------------------------------------------
class CommentCard(QFrame):
    """A comment: "Comment n", its color, the commented text and what was
    written. A click goes to its location in the text."""
    chosen = Signal(int, bool)      # id, move the focus to the book text?
    commentTextChanged = Signal(int, str)
    colorChanged = Signal(int, QColor)
    deleteRequested = Signal(int)

    def __init__(self, ident, parent=None):
        super().__init__(parent)
        self.ident = ident
        self.color = QColor(COMMENT_COLORS[0])
        self.active = False
        self.setObjectName("card")
        self.setCursor(Qt.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 8, 8, 8)
        lay.setSpacing(4)

        header = QHBoxLayout()
        header.setSpacing(4)
        self.title_label = QLabel()
        self.title_label.setStyleSheet("font-weight: bold; color: #1A1A1A;")
        header.addWidget(self.title_label)
        header.addStretch()

        # Highlight color: swatch with the color and a menu with the palette
        self.btn_color = QToolButton()
        self.btn_color.setToolTip(_t("Change the highlight color"))
        self.btn_color.setCursor(Qt.ArrowCursor)
        self.btn_color.setPopupMode(QToolButton.InstantPopup)
        self.btn_color.setFixedSize(28, 24)
        self.btn_color.setIconSize(QSize(18, 18))
        menu = QMenu(self.btn_color)
        box_widget = QWidget()
        grid = QGridLayout(box_widget)
        grid.setContentsMargins(8, 6, 8, 6)
        grid.setSpacing(4)
        for k, hex_color in enumerate(COMMENT_COLORS):
            b = QToolButton()
            b.setFixedSize(20, 20)
            b.setToolTip(hex_color)
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet(
                f"QToolButton {{ background:{hex_color}; border:1px solid #C4C4C4;"
                f" border-radius:3px; padding:0; }}"
                f"QToolButton:hover {{ border:2px solid #3A3A3A; }}")
            b.clicked.connect(lambda _=False, c=hex_color: self._choose_color(QColor(c)))
            grid.addWidget(b, k // 5, k % 5)
        wa = QWidgetAction(menu)
        wa.setDefaultWidget(box_widget)
        menu.addAction(wa)
        menu.addSeparator()
        menu.addAction(_t("More colors…")).triggered.connect(self._color_dialog)
        self.color_menu = menu
        self.btn_color.setMenu(menu)
        header.addWidget(self.btn_color)

        btn = QToolButton()
        btn.setIcon(file_icon("delete"))
        btn.setIconSize(QSize(16, 16))
        btn.setFixedSize(24, 22)
        btn.setAutoRaise(True)
        btn.setToolTip(_t("Delete the comment (can be undone with Ctrl+Z)"))
        btn.setCursor(Qt.ArrowCursor)
        btn.setStyleSheet(
            "QToolButton { border: 1px solid transparent; border-radius: 4px; padding: 0; }"
            "QToolButton:hover { background: #FDE7E7; border-color: #E8B4B4; }"
            "QToolButton:pressed { background: #F8D0D0; }")
        btn.clicked.connect(lambda: self.deleteRequested.emit(self.ident))
        header.addWidget(btn)
        lay.addLayout(header)

        # A single line: if the commented text doesn't fit, it is shortened
        # with "…" (the full text is in the tooltip)
        self.quote = ""
        self.quote_label = QLabel()
        self.quote_label.setWordWrap(False)
        self.quote_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.quote_label.setMinimumWidth(0)
        self.quote_label.installEventFilter(self)
        lay.addWidget(self.quote_label)

        self.text = QTextEdit()
        self.text.setAcceptRichText(False)
        self.text.setTabChangesFocus(True)
        self.text.setPlaceholderText(_t("Write a comment…"))
        self.text.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.text.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.text.setCursor(Qt.IBeamCursor)
        self.text.setStyleSheet(
            "QTextEdit { border: 1px solid #DADADA; border-radius: 4px;"
            " background: #FFFFFF; }")
        self.text.textChanged.connect(
            lambda: self.commentTextChanged.emit(self.ident, self.text.toPlainText()))
        self.text.document().documentLayout().documentSizeChanged.connect(
            self._adjust_height)
        self.text.installEventFilter(self)
        lay.addWidget(self.text)
        self._restyle()
        self._adjust_height()

    # -- interaction --------------------------------------------------------
    def eventFilter(self, obj, e):
        # When starting to type the comment, also go to its text
        if obj is getattr(self, "text", None) and e.type() == QEvent.FocusIn:
            self.chosen.emit(self.ident, False)
        elif obj is getattr(self, "quote_label", None) and e.type() == QEvent.Resize:
            self._show_quote()
        return False

    def _show_quote(self):
        """Shorten the quote to the available width, ending in "…"."""
        lbl = self.quote_label
        lbl.ensurePolished()
        fm = lbl.fontMetrics()
        width = lbl.contentsRect().width() - 12 - fm.horizontalAdvance(_t("“{text}”", text=""))
        short = fm.elidedText(self.quote, Qt.ElideRight, max(20, width))
        lbl.setText(_t("“{text}”", text=short))
        lbl.setToolTip(self.quote if short != self.quote else "")

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.chosen.emit(self.ident, True)
        super().mousePressEvent(e)

    def _choose_color(self, color):
        self.color_menu.hide()
        self.colorChanged.emit(self.ident, color)

    def _color_dialog(self):
        c = QColorDialog.getColor(self.color, self, _t("Comment color"))
        if c.isValid():
            self._choose_color(c)

    # -- appearance ---------------------------------------------------------
    def _adjust_height(self, *_):
        line = self.text.fontMetrics().lineSpacing()
        height = max(self.text.document().size().height(), 2 * line + 8)
        self.text.setFixedHeight(round(height) + 4)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._adjust_height()

    def refresh(self, number, data, quote, color):
        self.title_label.setText(_t("Comment {n}", n=number))
        author, date = data.get("author", ""), readable_date(data.get("date", ""))
        self.setToolTip(" · ".join(x for x in (author, date) if x))
        if quote != self.quote:
            self.quote = quote
            self._show_quote()
        if QColor(color) != self.color:
            self.color = QColor(color)
            self._restyle()
        # While typing, the data is updated instantly: it only differs if the
        # comment changed another way (e.g., when opening a file)
        if self.text.toPlainText() != data.get("text", ""):
            self.text.blockSignals(True)
            self.text.setPlainText(data.get("text", ""))
            self.text.blockSignals(False)

    def set_active(self, active):
        if active != self.active:
            self.active = active
            self._restyle()

    def _restyle(self):
        strong = active_color(self.color).name()
        border = f"2px solid {strong}" if self.active else "1px solid #DADADA"
        # Left strip with the comment color
        self.setStyleSheet(
            f"QFrame#card {{ background: #FFFFFF; border: {border};"
            f" border-left: 5px solid {self.color.name()}; border-radius: 6px; }}")
        self.btn_color.setIcon(change_color_icon(self.color))
        self.btn_color.setStyleSheet(
            "QToolButton { background: transparent; border: 1px solid transparent;"
            " border-radius: 4px; padding: 0; }"
            "QToolButton:hover { background: #E9E9E9; border-color: #D2D2D2; }"
            "QToolButton:pressed { background: #DDDDDD; }"
            "QToolButton::menu-indicator { image: none; width: 0; }")
        self.quote_label.setStyleSheet(
            f"color: #5A5A5A; font-style: italic; background: {self.color.name()};"
            " border-radius: 3px; padding: 2px 5px;")


class CommentsPanel(QDockWidget):
    """Lists the comments in text order ("Comment 1, 2…"). A click goes to its
    location; the comment where the cursor is gets marked."""

    def __init__(self, editor, parent=None):
        super().__init__(_t("Comments"), parent)
        self.editor = editor
        self.manager = editor.comments
        self.setObjectName("comments")
        self.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.setFeatures(QDockWidget.DockWidgetClosable)

        content = QWidget()
        lay = QVBoxLayout(content)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.setSpacing(6)
        self.area = QScrollArea()
        self.area.setWidgetResizable(True)
        self.area.setFrameShape(QFrame.NoFrame)
        self.area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.lst = QWidget()
        self.list_layout = QVBoxLayout(self.lst)
        self.list_layout.setContentsMargins(0, 0, 2, 0)
        self.list_layout.setSpacing(8)
        self.list_layout.addStretch()
        self.area.setWidget(self.lst)
        lay.addWidget(self.area, 1)
        self.empty_label = QLabel(
            _t("No comments yet.\nSelect text and click New comment (Ctrl+Alt+M)."))
        self.empty_label.setWordWrap(True)
        self.empty_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.empty_label.setStyleSheet("color: #808080; padding: 8px;")
        lay.addWidget(self.empty_label, 1)
        self.setWidget(content)
        self.setMinimumWidth(240)

        self.cards = {}
        self._order = []
        self.manager.listChanged.connect(self.refresh)
        self.manager.activeChanged.connect(self._mark)
        self.refresh()

    def refresh(self):
        g = self.manager
        for i in list(self.cards):
            if i not in g.ranges:
                t = self.cards.pop(i)
                self.list_layout.removeWidget(t)
                t.deleteLater()
        for i in g.order:
            if i not in self.cards:
                t = CommentCard(i)
                t.chosen.connect(self._go)
                t.commentTextChanged.connect(g.set_text)
                t.colorChanged.connect(g.set_color_value)
                t.deleteRequested.connect(g.delete)
                self.cards[i] = t
        if self._order != g.order:            # reorder according to the text
            for t in self.cards.values():
                self.list_layout.removeWidget(t)
            for pos, i in enumerate(g.order):
                self.list_layout.insertWidget(pos, self.cards[i])
            self._order = list(g.order)
        for n, i in enumerate(g.order, 1):
            self.cards[i].refresh(n, g.data[i], g.excerpt(i, 500), g.color(i))
        empty = not g.order
        self.area.setVisible(not empty)
        self.empty_label.setVisible(empty)
        self._mark(g.active_id)

    def _go(self, i, focus_editor):
        """Go to the comment's location in the book."""
        self.manager.go_to(i)
        if focus_editor:                    # click on the card
            self.editor.setFocus()

    def _mark(self, active_id):
        for i, t in self.cards.items():
            t.set_active(i == active_id)
        if active_id in self.cards and self.isVisible():
            self.area.ensureWidgetVisible(self.cards[active_id], 0, 20)

    def focus_comment(self, i):
        """Show the panel with the cursor in the comment, to write it."""
        self.show()
        self.refresh()
        t = self.cards.get(i)
        if t is not None:
            QTimer.singleShot(0, lambda: (self.area.ensureWidgetVisible(t),
                                          t.text.setFocus()))


# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self, config=None):
        super().__init__()
        self.path = None               # open file (None = unsaved)
        self._restart = False        # on close, start again (language)
        # settings.ini first: the styles must be loaded before the editor
        self.config = config or Settings()
        load_styles(self.config)
        self.editor = Editor()
        self.view = PageView(self.editor)
        # Center of the window: the document, the comments panel and, on the
        # right edge, the document scroll bar
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(4)
        self.splitter.addWidget(self.view)
        center = QWidget()
        center_layout = QHBoxLayout(center)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(0)
        center_layout.addWidget(self.splitter, 1)
        center_layout.addWidget(self._scroll_bar())
        self.setCentralWidget(center)
        file_bar_tools = FileBar(self)
        font_bar = FontBar(self.editor, self)
        paragraph_bar = ParagraphBar(self.editor, self)
        # All the tools in a ribbon
        self.ribbon = Ribbon(file_bar_tools, font_bar, paragraph_bar)
        bar = QToolBar(_t("Ribbon"), self)
        bar.setObjectName("ribbon")
        bar.setMovable(False)
        bar.setFloatable(False)
        bar.addWidget(self.ribbon)
        self.addToolBar(Qt.TopToolBarArea, bar)

        # Wide enough for the whole ribbon to show, without exceeding the
        # screen size
        screen = QApplication.primaryScreen().availableGeometry()
        width = self.ribbon.sizeHint().width() + 30
        self.resize(min(max(width, 1000), screen.width()),
                    min(820, screen.height()))

        # No menu bar: the file actions are in the File bar. Quit only has a
        # shortcut; page breaks are inserted with Ctrl+Enter (handled by the
        # editor)
        quit_action = QAction(_t("Quit"), self)
        quit_action.setShortcut(QKeySequence("Ctrl+Q"))
        quit_action.triggered.connect(self.close)
        self.addAction(quit_action)

        self.navigator = NavigatorPanel(self.editor, self.view, self)
        self.splitter.insertWidget(0, self.navigator)    # to the left of the document
        self.navigator.hide()
        self._navigator_action = file_bar_tools.action_navigator
        self.navigator.visibilityChanged.connect(self._sync_navigator_button)

        # Cut and Copy only with something selected; Paste only with something
        # pasteable
        self._clipboard_actions = (file_bar_tools.action_cut,
                                       file_bar_tools.action_copy,
                                       file_bar_tools.action_paste)
        for signal in (self.editor.selectionChanged, self.editor.imageSelectionChanged,
                      QApplication.clipboard().dataChanged):
            signal.connect(self._update_clipboard)
        self._update_clipboard()

        self.comments_panel = CommentsPanel(self.editor, self)
        self.splitter.addWidget(self.comments_panel)
        # navigator | document | comments: only the document stretches
        for k, stretch in enumerate((0, 1, 0)):
            self.splitter.setStretchFactor(k, stretch)
        self.splitter.setSizes([240, 10000, 280])
        self.comments_panel.hide()
        self._comments_action = file_bar_tools.action_comments
        self.comments_panel.visibilityChanged.connect(
            self._sync_comments_button)

        self.setStatusBar(StatusBar(self.view, self))
        self._start_backup()
        self._apply_settings()
        self.editor.document().modificationChanged.connect(self.setWindowModified)
        # The initial style and the page margins are not user changes: if they
        # stayed in the history, Ctrl+Z would remove them and the text would
        # spill off the sheet
        _cursor_to_start(self.editor)
        self._update_title()
        self.editor.setFocus()

    def _scroll_bar(self):
        """Vertical scroll bar of the document on the right edge of the window
        (to the right of the comments panel). The view's own bar is hidden
        and this one follows it: same range, same position."""
        internal = self.view.verticalScrollBar()
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        external = QScrollBar(Qt.Vertical)
        external.setFocusPolicy(Qt.NoFocus)

        def sync_range(minimum, maximum):
            external.setRange(minimum, maximum)
            external.setPageStep(internal.pageStep())
            external.setSingleStep(internal.singleStep())
            external.setVisible(maximum > minimum)

        internal.rangeChanged.connect(sync_range)
        internal.valueChanged.connect(external.setValue)
        external.valueChanged.connect(internal.setValue)
        sync_range(internal.minimum(), internal.maximum())
        external.setValue(internal.value())
        self.vertical_bar = external
        return external

    # -- file ---------------------------------------------------------------
    def _update_title(self):
        name = os.path.basename(self.path) if self.path else _t("Untitled")
        self.setWindowTitle(f"{name}[*] — SWordBook")
        self.setWindowModified(self.editor.document().isModified())

    def _folder(self):
        if self.path:
            return os.path.dirname(self.path)
        last = self.config.value("Files/last_folder") \
            if hasattr(self, "config") else ""
        return last if last and os.path.isdir(last) else os.path.expanduser("~")

    def _confirm_discard(self):
        """If there are unsaved changes, ask what to do. Returns True if it can
        continue (it was saved or the changes were discarded)."""
        if not self.editor.document().isModified():
            return True
        name = os.path.basename(self.path) if self.path else _t("Untitled")
        box = QMessageBox(QMessageBox.Warning, _t("Unsaved changes"),
                           _t("Save the changes to “{name}” before continuing?",
                              name=name),
                           parent=self)
        save = box.addButton(_t("Save"), QMessageBox.AcceptRole)
        dont_save = box.addButton(_t("Don't save"), QMessageBox.DestructiveRole)
        cancel = box.addButton(_t("Cancel"), QMessageBox.RejectRole)
        box.setDefaultButton(save)
        box.setEscapeButton(cancel)
        box.exec()
        if box.clickedButton() is save:
            return self.save()
        return box.clickedButton() is dont_save

    # -- settings.ini -------------------------------------------------------
    def _apply_settings(self):
        """Page, view, window and panels as they were last time."""
        c = self.config
        self.view.mode = c.mode()               # apply_config uses it
        self.view.apply_config(c.page())
        self.view.viewChanged.emit(self.view.mode)
        self.editor.set_collapse(c.value("View/collapse_headings"))
        self.view.set_zoom(max(ZOOM_MIN, min(ZOOM_MAX, c.value("View/zoom") / 100)))
        width, height = c.value("Window/width"), c.value("Window/height")
        screen = QApplication.primaryScreen().availableGeometry()
        if width > 200 and height > 200:
            self.resize(min(width, screen.width()), min(height, screen.height()))
        x, y = c.value("Window/x"), c.value("Window/y")
        if x >= 0 and y >= 0 and screen.contains(x + 50, y + 20):
            self.move(x, y)
        self.navigator.setVisible(c.value("Panels/navigator"))
        self.comments_panel.setVisible(c.value("Panels/comments"))
        self._backup_timer.setInterval(max(1, c.value("Backup/wait_seconds")) * 1000)

    def show_window(self):
        if self.config.value("Window/maximized"):
            self.showMaximized()
        else:
            self.show()

    def _save_settings(self):
        """On close: zoom, window and panels."""
        c = self.config
        c.set_value("View/zoom", round(self.view.zoom * 100))
        c.set_value("Window/maximized", self.isMaximized())
        g = self.normalGeometry() if self.isMaximized() else self.geometry()
        c.set_value("Window/x", g.x())
        c.set_value("Window/y", g.y())
        c.set_value("Window/width", g.width())
        c.set_value("Window/height", g.height())
        c.set_value("Panels/navigator", self.navigator.isVisible())
        c.set_value("Panels/comments", self.comments_panel.isVisible())
        c.save()

    def _remember_folder(self, path, key="Files/last_folder"):
        self.config.set_value(key, os.path.dirname(path))
        self.config.save()

    # -- automatic backup ---------------------------------------------------
    def _start_backup(self):
        self._backup_timer = QTimer(self)
        self._backup_timer.setSingleShot(True)
        self._backup_timer.setInterval(BACKUP_WAIT)
        self._backup_timer.timeout.connect(self._backup)
        self._current_backup = None       # backup file that exists
        self._lock = None
        self._pending = False
        self._first_change = 0.0
        self._untitled_backup = None   # backup of an untitled document
        self._session_id = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        self.editor.document().contentsChanged.connect(self._schedule_backup)
        self.editor.comments.changed.connect(self._schedule_backup)

    def _schedule_backup(self, *_):
        """Back up 5 s after typing stops, or every 30 s if typing doesn't
        stop."""
        if not self.editor.document().isModified() or \
                not self.config.value("Backup/enabled"):
            return
        now = time.monotonic()
        if not self._pending:
            self._pending = True
            self._first_change = now
        if now - self._first_change >= max(5, self.config.value("Backup/max_seconds")):
            QTimer.singleShot(0, self._backup)
        else:
            self._backup_timer.start()

    def _backup_path(self, in_app_folder=False):
        if self.path and not in_app_folder:
            return backup_path_of(self.path)
        if not self.path and self._untitled_backup:
            return self._untitled_backup
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path \
            else _t("Untitled")
        return os.path.join(app_folder(), "respaldos",
                            f"{base} (backup {self._session_id}).docx")

    def _backup(self):
        self._backup_timer.stop()
        if not self._pending:
            return
        self._pending = False
        if not self.editor.document().isModified():
            return
        written = None
        # Next to the original; if that folder is not writable, in the
        # program's folder
        for target in dict.fromkeys((self._backup_path(), self._backup_path(True))):
            try:
                os.makedirs(os.path.dirname(target), exist_ok=True)
                DocxWriter(self.editor, self.view.config, self.view.mode).save(
                    target, os.path.splitext(os.path.basename(target))[0])
                written = target
                break
            except Exception:
                continue
        if written is None:
            self.statusBar().showMessage(_t("Could not create the backup file"), 5000)
            return
        if self._current_backup and self._current_backup != written:
            self._delete_backup()          # e.g., after Save as
        if self._current_backup != written:
            self._current_backup = written
            self._lock = backup_lock(written)
            self._lock.tryLock(0)
        register_backup(written, self.path)

    def _delete_backup(self):
        """No longer needed (it was saved or closed normally)."""
        self._backup_timer.stop()
        self._pending = False
        if self._current_backup:
            try:
                if os.path.exists(self._current_backup):
                    os.remove(self._current_backup)
            except OSError:
                pass
            remove_from_registry(self._current_backup)
        if self._lock is not None:
            self._lock.unlock()
        self._current_backup = None
        self._lock = None
        if not self.path:
            self._untitled_backup = None

    def _recover(self, backup, original):
        """Open the backup as the original document, with the changes pending
        to be saved."""
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            open_docx(self.editor, self.view, backup)
        except Exception as error:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, _t("Could not recover"),
                                 _t("Could not open the backup.") + f"\n\n{error}")
            return False
        QApplication.restoreOverrideCursor()
        valid = original and os.path.isdir(os.path.dirname(original))
        self.path = original if valid else None
        self._untitled_backup = None if self.path else backup
        self._current_backup = backup
        self._lock = backup_lock(backup)
        self._lock.tryLock(0)
        register_backup(backup, self.path)
        self.editor.document().setModified(True)
        if self.editor.comments.order:
            self.comments_panel.show()
        self._update_title()
        self.statusBar().showMessage(
            _t("The backup was recovered. Save the document to keep the changes."), 8000)
        self.editor.setFocus()
        return True

    def offer_recovery(self):
        """At startup: backups left behind because the program didn't close
        normally (the ones used by another open window are not touched)."""
        def date(e):
            try:
                return os.path.getmtime(e["backup"])
            except OSError:
                return 0
        for e in sorted(read_registry(), key=date, reverse=True):
            backup = e["backup"]
            if not os.path.exists(backup):
                remove_from_registry(backup)
                continue
            if backup_in_use(backup):
                continue
            original = e.get("original")
            name = os.path.basename(original) if original else _t("Untitled")
            box = QMessageBox(QMessageBox.Warning, _t("Recover document"),
                               _t("SWordBook did not close properly.\n\nThere is a "
                                  "backup of “{name}” from {when}.\nDo you want to "
                                  "recover it?", name=name, when=local_date(date(e))),
                               parent=self)
            recover = box.addButton(_t("Recover"), QMessageBox.AcceptRole)
            discard = box.addButton(_t("Discard"), QMessageBox.DestructiveRole)
            box.addButton(_t("Later"), QMessageBox.RejectRole)
            box.setDefaultButton(recover)
            box.exec()
            if box.clickedButton() is recover:
                if self._recover(backup, original):
                    return
            elif box.clickedButton() is discard:
                try:
                    os.remove(backup)
                except OSError:
                    pass
                remove_from_registry(backup)

    def _check_backup_of(self, path):
        """Before opening a file: if it has a backup newer than itself, ask
        what to do. Returns "recovered", "open" or "cancel"."""
        backup = backup_path_of(path)
        if not os.path.exists(backup):            # name used by earlier versions
            backup = old_backup_path_of(path)
        if backup == path or not os.path.exists(backup) or backup_in_use(backup):
            return "open"
        try:
            newer = os.path.getmtime(backup) > os.path.getmtime(path)
        except OSError:
            newer = False
        if not newer:                  # old backup: no longer useful
            try:
                os.remove(backup)
            except OSError:
                pass
            remove_from_registry(backup)
            return "open"
        box = QMessageBox(QMessageBox.Warning, _t("Backup found"),
                           _t("“{name}” has a more recent backup (from {when}) with "
                              "changes that were not saved.\n\nDo you want to "
                              "recover it?", name=os.path.basename(path),
                              when=local_date(os.path.getmtime(backup))),
                           parent=self)
        recover = box.addButton(_t("Recover changes"), QMessageBox.AcceptRole)
        original = box.addButton(_t("Open the original"), QMessageBox.DestructiveRole)
        box.addButton(_t("Cancel"), QMessageBox.RejectRole)
        box.setDefaultButton(recover)
        box.exec()
        if box.clickedButton() is recover:
            self._delete_backup()          # the one of the document being closed
            return "recovered" if self._recover(backup, path) else "cancel"
        if box.clickedButton() is original:
            try:
                os.remove(backup)
            except OSError:
                pass
            remove_from_registry(backup)
            return "open"
        return "cancel"

    def new(self):
        if not self._confirm_discard():
            return
        self._delete_backup()
        self._session_id = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        doc = self.editor.document()
        doc.setUndoRedoEnabled(False)
        doc.clear()
        self.editor._initial_style()
        doc.setUndoRedoEnabled(True)
        self.editor.comments.clear_all()
        self.editor.clear_images()
        self.view.set_mode(self.config.mode())
        self.view.apply_config(self.config.page())
        _cursor_to_start(self.editor)
        self.path = None
        self._update_title()
        self.editor.setFocus()

    def open_file(self):
        if not self._confirm_discard():
            return
        path, _ = QFileDialog.getOpenFileName(
            self, _t("Open"), self._folder(), _t(OPEN_FILTER))
        if not path:
            return
        if path.lower().endswith(".docx"):
            action = self._check_backup_of(path)
            if action != "open":
                return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            if path.lower().endswith(".txt"):
                open_text(self.editor, self.view, path)
            else:
                open_docx(self.editor, self.view, path)
        except Exception as error:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, _t("Could not open"),
                                 _t("Could not open “{name}”.", name=os.path.basename(path))
                                 + f"\n\n{error}")
            return
        QApplication.restoreOverrideCursor()
        self._delete_backup()      # the one of the previous document (already discarded)
        if self.editor.comments.order:      # the file has comments
            self.comments_panel.show()
        # A .txt is later saved as .docx (Save will ask for the name)
        self.path = path if path.lower().endswith(".docx") else None
        self._remember_folder(path)
        self._update_title()
        self.editor.setFocus()

    def save(self):
        if not self.path:
            return self.save_as()
        return self._write(self.path)

    def save_as(self):
        suggested = self.path or os.path.join(self._folder(), _t("Untitled") + ".docx")
        path, _ = QFileDialog.getSaveFileName(
            self, _t("Save as"), suggested, _t(DOCX_FILTER))
        if not path:
            return False
        if not path.lower().endswith(".docx"):
            path += ".docx"
        if self._write(path):
            self.path = path
            self._remember_folder(path)
            self._update_title()
            return True
        return False

    def _write(self, path):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            title = os.path.splitext(os.path.basename(path))[0]
            DocxWriter(self.editor, self.view.config,
                         self.view.mode).save(path, title)
        except Exception as error:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, _t("Could not save"),
                                 _t("Could not save “{name}”.", name=os.path.basename(path))
                                 + f"\n\n{error}")
            return False
        QApplication.restoreOverrideCursor()
        self.editor.document().setModified(False)
        self._delete_backup()      # what was saved is now safe
        self.statusBar().showMessage(_t("Saved: {path}", path=path), 4000)
        return True

    def closeEvent(self, e):
        if self._confirm_discard():
            self._delete_backup()  # normal close: the backup is no longer needed
            self._save_settings()
            e.accept()
            if self._restart:      # the language was changed
                restart_program()
        else:
            e.ignore()
            self._restart = False

    def _sync_navigator_button(self, *_):
        try:
            self._navigator_action.setChecked(self.navigator.isVisible())
        except RuntimeError:      # when closing the window the button no longer exists
            pass

    def _sync_comments_button(self, *_):
        try:
            self._comments_action.setChecked(self.comments_panel.isVisible())
        except RuntimeError:
            pass

    def toggle_comments(self, *_):
        self.comments_panel.setVisible(not self.comments_panel.isVisible())

    def insert_image(self, *_):
        path, _ = QFileDialog.getOpenFileName(
            self, _t("Insert image"),
            self.config.value("Files/images_folder") or self._folder(),
            _t("Images") + " (*.png *.jpg *.jpeg *.gif *.bmp *.webp *.tif *.tiff);;"
            + _t("All files") + " (*)")
        if not path:
            return
        self._remember_folder(path, "Files/images_folder")
        try:
            self.editor.insert_image_file(path)
        except Exception as error:
            QMessageBox.critical(self, _t("Could not insert the image"),
                                 _t("“{name}”: {error}", name=os.path.basename(path),
                                    error=error))
            return
        self.statusBar().showMessage(
            _t("Drag the image to move it, or its handles to resize it "
               "(Delete removes it)"), 5000)
        self.editor.setFocus()

    def new_comment(self, *_):
        i = self.editor.comments.add()
        if i is None:
            self.statusBar().showMessage(
                _t("Select the text you want to comment on "
                   "(or place the cursor on a word)"), 4000)
            QApplication.beep()
            return
        self.comments_panel.focus_comment(i)

    def _update_clipboard(self, *_):
        try:
            cut_selection, copy_selection, paste_clipboard = self._clipboard_actions
            has = self.editor.can_copy()
            cut_selection.setEnabled(has)
            copy_selection.setEnabled(has)
            md = QApplication.clipboard().mimeData()
            paste_clipboard.setEnabled(md is not None and self.editor.canInsertFromMimeData(md))
        except RuntimeError:      # when closing the window
            pass

    def cut_selection(self, *_):
        self.editor.cut_selection()
        self.editor.setFocus()

    def copy_selection(self, *_):
        self.editor.copy_selection()
        self._update_clipboard()
        self.editor.setFocus()

    def paste_clipboard(self, *_):
        self.editor.paste_clipboard()
        self.editor.setFocus()

    def toggle_navigator(self, *_):
        self.navigator.setVisible(not self.navigator.isVisible())
        if self.navigator.isVisible():
            self.navigator.refresh()

    def find_text(self, replace=False):
        if getattr(self, "_find_dialog", None) is None:
            self._find_dialog = FindDialog(self.editor, self)
        self._find_dialog.open_file(replace)

    def open_settings(self):
        dialog = SettingsDialog(self.view.config, self.view.mode,
                                self.editor.collapse_enabled, language=self.config.language(),
                                parent=self)
        if dialog.exec() == QDialog.Accepted:
            self.view.set_mode(dialog.mode())
            self.editor.set_collapse(dialog.collapse())
            self.view.apply_config(dialog.config())
            # Becomes the setting for new documents
            self.config.set_page(dialog.config())
            self.config.set_value("View/mode", dialog.mode())
            self.config.set_value("View/collapse_headings", dialog.collapse())
            self.config.set_value("Interface/language", dialog.language())
            self.config.save()
            if dialog.language() != LANGUAGE:
                self._offer_restart(dialog.language())
        self.editor.setFocus()

    def _offer_restart(self, language):
        """The language is applied when the program starts: offer to restart
        it. The notice already appears in the chosen language."""
        current = LANGUAGE
        set_language(language)
        box = QMessageBox(QMessageBox.Question, _t("Language"),
                           _t("The language will change when SWordBook restarts."
                              "\n\nRestart now?"), parent=self)
        now = box.addButton(_t("Restart now"), QMessageBox.AcceptRole)
        box.addButton(_t("Later"), QMessageBox.RejectRole)
        box.setDefaultButton(now)
        set_language(current)          # the rest stays the same until restarting
        box.exec()
        if box.clickedButton() is now:
            self._restart = True
            self.close()              # asks if there are unsaved changes


# ---------------------------------------------------------------------------
# Interface translations. The original text is in English; each row: (English,
# Spanish, Portuguese, French, German, Italian, Chinese, Japanese). A key with
# "context|text" distinguishes identical English words that translate
# differently (e.g., "Left" for indent and for margin)
# ---------------------------------------------------------------------------
_ROWS = [
    # -- styles ----------------------------------------------------------------
    ("Body", "Cuerpo", "Corpo", "Corps", "Fließtext", "Corpo", "正文", "本文"),
    ("Title", "Título", "Título", "Titre", "Titel", "Titolo", "标题", "タイトル"),
    ("Subtitle", "Subtítulo", "Subtítulo", "Sous-titre", "Untertitel", "Sottotitolo",
     "副标题", "サブタイトル"),
    ("Section", "Sección", "Seção", "Section", "Abschnitt", "Sezione", "小节", "セクション"),
    ("Thoughts", "Pensamientos", "Pensamentos", "Pensées", "Gedanken", "Pensieri",
     "内心独白", "心の声"),
    ("Dialogue", "Diálogo", "Diálogo", "Dialogue", "Dialog", "Dialogo", "对话", "会話"),
    ("Quote", "Cita", "Citação", "Citation", "Zitat", "Citazione", "引文", "引用"),
    ("Scene break", "Separador", "Quebra de cena", "Saut de scène", "Szenenwechsel",
     "Cambio scena", "场景分隔", "場面転換"),
    ("Chapter title", "Título de capítulo", "Título do capítulo", "Titre de chapitre",
     "Kapiteltitel", "Titolo del capitolo", "章节标题", "章タイトル"),
    ("Chapter subtitle", "Subtítulo de capítulo", "Subtítulo do capítulo",
     "Sous-titre de chapitre", "Kapiteluntertitel", "Sottotitolo del capitolo",
     "章节副标题", "章サブタイトル"),
    ("Section heading", "Encabezado de sección", "Cabeçalho de seção", "En-tête de section",
     "Abschnittsüberschrift", "Intestazione di sezione", "小节标题", "セクション見出し"),
    ("Scene separator", "Separador de escena", "Separador de cena", "Séparateur de scène",
     "Szenentrenner", "Separatore di scena", "场景分隔符", "場面区切り"),
    ('Sets the tools to "{name}"', 'Establece las herramientas a "{name}"',
     'Ajusta as ferramentas para "{name}"', 'Règle les outils sur « {name} »',
     'Stellt die Werkzeuge auf „{name}“', 'Imposta gli strumenti su "{name}"',
     '将工具设为“{name}”', 'ツールを「{name}」に設定'),
    (' and types "{symbol}"', ' y escribe "{symbol}"', ' e escreve "{symbol}"',
     ' et écrit « {symbol} »', ' und schreibt „{symbol}“', ' e scrive "{symbol}"',
     '并输入“{symbol}”', '、「{symbol}」を入力'),
    # -- page sizes ------------------------------------------------------------
    ("Letter", "Carta", "Carta", "Lettre US", "US-Letter", "Lettera", "信纸", "レター"),
    ("Oficio", "Oficio", "Ofício", "Oficio", "Oficio", "Oficio", "Oficio", "オフィシオ"),
    ("Legal", "Legal", "Legal", "Légal US", "US-Legal", "Legale", "法律文书纸", "リーガル"),
    ("Book 6 × 9 in", "Libro 6 × 9 in", "Livro 6 × 9 pol", "Livre 6 × 9 po",
     "Buch 6 × 9 Zoll", "Libro 6 × 9 pollici", "图书 6 × 9 英寸", "書籍 6 × 9 インチ"),
    ("Book 5.5 × 8.5 in", "Libro 5.5 × 8.5 in", "Livro 5,5 × 8,5 pol", "Livre 5,5 × 8,5 po",
     "Buch 5,5 × 8,5 Zoll", "Libro 5,5 × 8,5 pollici", "图书 5.5 × 8.5 英寸",
     "書籍 5.5 × 8.5 インチ"),
    ("Book 5 × 8 in", "Libro 5 × 8 in", "Livro 5 × 8 pol", "Livre 5 × 8 po",
     "Buch 5 × 8 Zoll", "Libro 5 × 8 pollici", "图书 5 × 8 英寸", "書籍 5 × 8 インチ"),
    # -- colors ------------------------------------------------------------------
    ("Colors", "Colores", "Cores", "Couleurs", "Farben", "Colori", "颜色", "色"),
    ("More colors…", "Más colores…", "Mais cores…", "Autres couleurs…", "Weitere Farben…",
     "Altri colori…", "更多颜色…", "その他の色…"),
    ("Choose color", "Elegir color", "Escolher cor", "Choisir une couleur", "Farbe wählen",
     "Scegli colore", "选择颜色", "色の選択"),
    ("Automatic", "Automático", "Automático", "Automatique", "Automatisch", "Automatico",
     "自动", "自動"),
    ("Font color", "Color de fuente", "Cor da fonte", "Couleur de police", "Schriftfarbe",
     "Colore carattere", "字体颜色", "フォントの色"),
    ("No color", "Sin color", "Sem cor", "Aucune couleur", "Keine Farbe", "Nessun colore",
     "无颜色", "色なし"),
    ("Highlight color", "Color de resaltado", "Cor de realce", "Couleur de surlignage",
     "Hervorhebungsfarbe", "Colore evidenziatore", "突出显示颜色", "蛍光ペンの色"),
    # -- editor ------------------------------------------------------------------
    ("Author", "Autor", "Autor", "Auteur", "Autor", "Autore", "作者", "作成者"),
    ("Page break", "Salto de página", "Quebra de página", "Saut de page", "Seitenumbruch",
     "Interruzione di pagina", "分页符", "改ページ"),
    ("Page breaks are not available in the No pages view (Settings → View → Pages or Mixed)",
     "Los saltos de página no están disponibles en la vista Sin páginas "
     "(Ajustes → Vista → Páginas o Mixto)",
     "As quebras de página não estão disponíveis no modo Sem páginas "
     "(Configurações → Exibição → Páginas ou Misto)",
     "Les sauts de page ne sont pas disponibles en mode Sans pages "
     "(Paramètres → Affichage → Pages ou Mixte)",
     "Seitenumbrüche sind in der Ansicht „Ohne Seiten“ nicht verfügbar "
     "(Einstellungen → Ansicht → Seiten oder Gemischt)",
     "Le interruzioni di pagina non sono disponibili nella vista Senza pagine "
     "(Impostazioni → Vista → Pagine o Mista)",
     "“无分页”视图中不能插入分页符（设置 → 视图 → 分页或混合）",
     "「ページなし」表示では改ページを使えません（設定 → 表示 → ページまたは混合）"),
    ("Image", "Imagen", "Imagem", "Image", "Bild", "Immagine", "图片", "画像"),
    ("The file is not an image that can be read.",
     "El archivo no es una imagen que se pueda leer.",
     "O arquivo não é uma imagem que possa ser lida.",
     "Le fichier n’est pas une image lisible.",
     "Die Datei ist kein lesbares Bild.",
     "Il file non è un’immagine leggibile.",
     "该文件不是可读取的图片。", "このファイルは読み込める画像ではありません。"),
    # -- font ----------------------------------------------------------------------
    ("Font", "Fuente", "Fonte", "Police", "Schriftart", "Carattere", "字体", "フォント"),
    ("Font size", "Tamaño de fuente", "Tamanho da fonte", "Taille de police",
     "Schriftgröße", "Dimensione carattere", "字号", "フォントサイズ"),
    ("Increase size", "Aumentar tamaño", "Aumentar tamanho", "Agrandir la police",
     "Schrift vergrößern", "Aumenta dimensione", "增大字号", "フォントの拡大"),
    ("Decrease size", "Disminuir tamaño", "Diminuir tamanho", "Réduire la police",
     "Schrift verkleinern", "Riduci dimensione", "减小字号", "フォントの縮小"),
    ("Clear formatting: back to the {style} style",
     "Borrar formato: regresa al estilo {style}",
     "Limpar formatação: volta ao estilo {style}",
     "Effacer la mise en forme : revient au style {style}",
     "Formatierung löschen: zurück zur Formatvorlage {style}",
     "Cancella formattazione: torna allo stile {style}",
     "清除格式：恢复为“{style}”样式", "書式のクリア：「{style}」スタイルに戻す"),
    ("Bold", "Negrita", "Negrito", "Gras", "Fett", "Grassetto", "加粗", "太字"),
    ("Italic", "Cursiva", "Itálico", "Italique", "Kursiv", "Corsivo", "倾斜", "斜体"),
    ("Underline", "Subrayado", "Sublinhado", "Souligné", "Unterstrichen", "Sottolineato",
     "下划线", "下線"),
    ("Strikethrough", "Tachado", "Tachado", "Barré", "Durchgestrichen", "Barrato",
     "删除线", "取り消し線"),
    ("Subscript", "Subíndice", "Subscrito", "Indice", "Tiefgestellt", "Pedice",
     "下标", "下付き"),
    ("Superscript", "Superíndice", "Sobrescrito", "Exposant", "Hochgestellt", "Apice",
     "上标", "上付き"),
    # -- paragraph -----------------------------------------------------------------
    ("Align left", "Alinear a la izquierda", "Alinhar à esquerda", "Aligner à gauche",
     "Linksbündig", "Allinea a sinistra", "左对齐", "左揃え"),
    ("Center", "Centrar", "Centralizar", "Centrer", "Zentriert", "Centra", "居中", "中央揃え"),
    ("Align right", "Alinear a la derecha", "Alinhar à direita", "Aligner à droite",
     "Rechtsbündig", "Allinea a destra", "右对齐", "右揃え"),
    ("Justify", "Justificar", "Justificar", "Justifier", "Blocksatz", "Giustifica",
     "两端对齐", "両端揃え"),
    ("Decrease indent", "Reducir sangría", "Diminuir recuo", "Diminuer le retrait",
     "Einzug verkleinern", "Riduci rientro", "减少缩进", "インデントを減らす"),
    ("Increase indent", "Aumentar sangría", "Aumentar recuo", "Augmenter le retrait",
     "Einzug vergrößern", "Aumenta rientro", "增加缩进", "インデントを増やす"),
    ("Decrease space after paragraph", "Reducir espacio después del párrafo",
     "Diminuir espaço depois do parágrafo", "Réduire l’espace après le paragraphe",
     "Abstand nach Absatz verkleinern", "Riduci spazio dopo il paragrafo",
     "减少段后间距", "段落後の間隔を狭くする"),
    ("Increase space after paragraph", "Aumentar espacio después del párrafo",
     "Aumentar espaço depois do parágrafo", "Augmenter l’espace après le paragraphe",
     "Abstand nach Absatz vergrößern", "Aumenta spazio dopo il paragrafo",
     "增加段后间距", "段落後の間隔を広くする"),
    ("Numbered list", "Lista numérica", "Lista numerada", "Liste numérotée",
     "Nummerierte Liste", "Elenco numerato", "编号列表", "番号付きリスト"),
    ("Bulleted list", "Lista de viñetas", "Lista com marcadores", "Liste à puces",
     "Aufzählung", "Elenco puntato", "项目符号列表", "箇条書き"),
    ("Spacing", "Espaciado", "Espaçamento", "Espacement", "Abstand", "Spaziatura",
     "间距", "間隔"),
    ("Line spacing", "Interlineado", "Entrelinhas", "Interligne", "Zeilenabstand",
     "Interlinea", "行距", "行間"),
    ("Line spacing (multiple of the line height, 0.5 to 5)",
     "Interlineado (múltiplo del alto de línea, de 0.5 a 5)",
     "Entrelinhas (múltiplo da altura da linha, de 0.5 a 5)",
     "Interligne (multiple de la hauteur de ligne, de 0.5 à 5)",
     "Zeilenabstand (Vielfaches der Zeilenhöhe, 0.5 bis 5)",
     "Interlinea (multiplo dell’altezza della riga, da 0.5 a 5)",
     "行距（行高的倍数，0.5 到 5）", "行間（行の高さの倍数、0.5～5）"),
    ("Before", "Antes", "Antes", "Avant", "Vor", "Prima", "段前", "前"),
    ("Space before paragraph", "Espacio antes del párrafo", "Espaço antes do parágrafo",
     "Espace avant le paragraphe", "Abstand vor dem Absatz", "Spazio prima del paragrafo",
     "段前间距", "段落前の間隔"),
    ("After", "Después", "Depois", "Après", "Nach", "Dopo", "段后", "後"),
    ("Space after paragraph", "Espacio después del párrafo", "Espaço depois do parágrafo",
     "Espace après le paragraphe", "Abstand nach dem Absatz", "Spazio dopo il paragrafo",
     "段后间距", "段落後の間隔"),
    ("Letters", "Letras", "Letras", "Lettres", "Zeichen", "Lettere", "字距", "文字間"),
    ("Letter spacing: positive spreads, negative tightens ({min} to {max} pt; 0 = normal)",
     "Espaciado entre letras: positivo separa, negativo junta ({min} a {max} pt; 0 = normal)",
     "Espaçamento entre letras: positivo separa, negativo junta ({min} a {max} pt; 0 = normal)",
     "Espacement des lettres : positif écarte, négatif resserre ({min} à {max} pt ; 0 = normal)",
     "Zeichenabstand: positiv sperrt, negativ verdichtet ({min} bis {max} pt; 0 = normal)",
     "Spaziatura tra lettere: positiva allarga, negativa stringe ({min} a {max} pt; 0 = normale)",
     "字距：正值加宽，负值紧缩（{min} 到 {max} 磅；0 = 标准）",
     "文字間隔：正の値で広げ、負の値で詰める（{min}～{max} pt、0 = 標準）"),
    ("Indent", "Sangría", "Recuo", "Retrait", "Einzug", "Rientro", "缩进", "インデント"),
    ("indent|Left", "Izquierda", "Esquerda", "Gauche", "Links", "Sinistra", "左", "左"),
    ("Left indent", "Sangría izquierda", "Recuo à esquerda", "Retrait gauche",
     "Einzug links", "Rientro sinistro", "左缩进", "左インデント"),
    ("indent|Right", "Derecha", "Direita", "Droite", "Rechts", "Destra", "右", "右"),
    ("Right indent", "Sangría derecha", "Recuo à direita", "Retrait droit",
     "Einzug rechts", "Rientro destro", "右缩进", "右インデント"),
    ("First line", "1.ª línea", "1.ª linha", "1re ligne", "1. Zeile", "1ª riga",
     "首行", "1 行目"),
    ("First-line indent", "Sangría de primera línea", "Recuo da primeira linha",
     "Retrait de première ligne", "Erstzeileneinzug", "Rientro prima riga",
     "首行缩进", "1 行目のインデント"),
    # -- settings ------------------------------------------------------------------
    ("Settings", "Ajustes", "Configurações", "Paramètres", "Einstellungen", "Impostazioni",
     "设置", "設定"),
    ("Language", "Idioma", "Idioma", "Langue", "Sprache", "Lingua", "语言", "言語"),
    ("View", "Vista", "Exibição", "Affichage", "Ansicht", "Vista", "视图", "表示"),
    ("Pages", "Páginas", "Páginas", "Pages", "Seiten", "Pagine", "分页", "ページ"),
    ("Shows the document on sheets, as it will look when printed",
     "Muestra el documento en hojas, como quedará al imprimir",
     "Mostra o documento em folhas, como ficará ao imprimir",
     "Affiche le document en feuilles, tel qu’il sera imprimé",
     "Zeigt das Dokument auf Blättern, wie es gedruckt aussieht",
     "Mostra il documento in fogli, come apparirà in stampa",
     "按纸张显示文档，与打印效果一致", "印刷時と同じように用紙単位で文書を表示します"),
    ("No pages", "Sin páginas", "Sem páginas", "Sans pages", "Ohne Seiten", "Senza pagine",
     "无分页", "ページなし"),
    ("A single continuous sheet, with no breaks between pages, to write and read "
     "without interruptions (page breaks are not allowed)",
     "Una sola hoja continua, sin cortes entre páginas, para escribir y leer sin "
     "interrupciones (no permite saltos de página)",
     "Uma única folha contínua, sem cortes entre páginas, para escrever e ler sem "
     "interrupções (não permite quebras de página)",
     "Une seule feuille continue, sans coupure entre les pages, pour écrire et lire "
     "sans interruption (pas de sauts de page)",
     "Ein einziges fortlaufendes Blatt ohne Seitenwechsel, um ohne Unterbrechung zu "
     "schreiben und zu lesen (keine Seitenumbrüche möglich)",
     "Un unico foglio continuo, senza tagli tra le pagine, per scrivere e leggere "
     "senza interruzioni (non consente interruzioni di pagina)",
     "一张连续的纸，页面之间没有分隔，便于不间断地写作和阅读（不能插入分页符）",
     "ページの区切りがない 1 枚の連続した用紙で、途切れずに書いたり読んだりできます"
     "（改ページは使えません）"),
    ("Mixed", "Mixto", "Misto", "Mixte", "Gemischt", "Mista", "混合", "混合"),
    ("Like No pages, but each page break (Ctrl+Enter) splits the sheet: sheets of "
     "different lengths, each as long as its content",
     "Como Sin páginas, pero cada salto de página (Ctrl+Enter) parte la hoja: hojas de "
     "distintas longitudes, cada una del largo de su contenido",
     "Como Sem páginas, mas cada quebra de página (Ctrl+Enter) divide a folha: folhas de "
     "comprimentos diferentes, cada uma do tamanho do seu conteúdo",
     "Comme Sans pages, mais chaque saut de page (Ctrl+Entrée) coupe la feuille : des "
     "feuilles de longueurs différentes, chacune aussi longue que son contenu",
     "Wie „Ohne Seiten“, aber jeder Seitenumbruch (Strg+Eingabe) teilt das Blatt: "
     "unterschiedlich lange Blätter, jedes so lang wie sein Inhalt",
     "Come Senza pagine, ma ogni interruzione di pagina (Ctrl+Invio) divide il foglio: "
     "fogli di lunghezze diverse, ognuno lungo quanto il suo contenuto",
     "与“无分页”相同，但每个分页符（Ctrl+Enter）会把纸张分开：纸张长短不一，"
     "各自与其内容等长",
     "「ページなし」と同じですが、改ページ（Ctrl+Enter）ごとに用紙が分かれます。"
     "用紙の長さはそれぞれの内容に合わせて変わります"),
    ("Headings", "Títulos", "Títulos", "Titres", "Überschriften", "Titoli", "标题", "見出し"),
    ("Allow collapsing {title}, {subtitle} and {section}",
     "Permitir contraer {title}, {subtitle} y {section}",
     "Permitir recolher {title}, {subtitle} e {section}",
     "Permettre de réduire {title}, {subtitle} et {section}",
     "{title}, {subtitle} und {section} einklappbar machen",
     "Consenti di comprimere {title}, {subtitle} e {section}",
     "允许折叠{title}、{subtitle}和{section}",
     "{title}、{subtitle}、{section}の折りたたみを許可"),
    ("Shows a triangle next to each heading to hide or show its content. Turning it "
     "off expands everything.",
     "Muestra un triángulo junto a cada título para ocultar o mostrar su contenido. "
     "Al desactivarlo se despliega todo.",
     "Mostra um triângulo ao lado de cada título para ocultar ou mostrar seu conteúdo. "
     "Ao desativar, tudo é expandido.",
     "Affiche un triangle à côté de chaque titre pour masquer ou afficher son contenu. "
     "Le désactiver développe tout.",
     "Zeigt neben jeder Überschrift ein Dreieck, um ihren Inhalt aus- oder einzublenden. "
     "Beim Ausschalten wird alles aufgeklappt.",
     "Mostra un triangolo accanto a ogni titolo per nascondere o mostrare il contenuto. "
     "Disattivandolo si espande tutto.",
     "在每个标题旁显示三角形，用于隐藏或显示其内容。关闭后全部展开。",
     "各見出しの横に三角形を表示し、内容の表示と非表示を切り替えます。"
     "オフにするとすべて展開されます。"),
    ("Paper", "Papel", "Papel", "Papier", "Papier", "Carta", "纸张", "用紙"),
    ("Size", "Tamaño", "Tamanho", "Taille", "Größe", "Dimensione", "大小", "サイズ"),
    ("Custom", "Personalizado", "Personalizado", "Personnalisé", "Benutzerdefiniert",
     "Personalizzato", "自定义", "ユーザー設定"),
    ("Width", "Ancho", "Largura", "Largeur", "Breite", "Larghezza", "宽度", "幅"),
    ("Height", "Alto", "Altura", "Hauteur", "Höhe", "Altezza", "高度", "高さ"),
    ("Orientation", "Orientación", "Orientação", "Orientation", "Ausrichtung",
     "Orientamento", "方向", "向き"),
    ("Portrait", "Vertical", "Retrato", "Portrait", "Hochformat", "Verticale", "纵向", "縦"),
    ("Landscape", "Horizontal", "Paisagem", "Paysage", "Querformat", "Orizzontale",
     "横向", "横"),
    ("Margins", "Márgenes", "Margens", "Marges", "Ränder", "Margini", "页边距", "余白"),
    ("Top", "Superior", "Superior", "Haut", "Oben", "Superiore", "上", "上"),
    ("Bottom", "Inferior", "Inferior", "Bas", "Unten", "Inferiore", "下", "下"),
    ("margin|Left", "Izquierdo", "Esquerda", "Gauche", "Links", "Sinistro", "左", "左"),
    ("margin|Right", "Derecho", "Direita", "Droite", "Rechts", "Destro", "右", "右"),
    ("OK", "Aceptar", "OK", "OK", "OK", "OK", "确定", "OK"),
    ("Cancel", "Cancelar", "Cancelar", "Annuler", "Abbrechen", "Annulla", "取消",
     "キャンセル"),
    ("The margins are too large for the page size. They must leave at least 2 cm of "
     "space for the text.",
     "Los márgenes son demasiado grandes para el tamaño de página. Deben dejar al menos "
     "2 cm de espacio para el texto.",
     "As margens são grandes demais para o tamanho da página. Devem deixar pelo menos "
     "2 cm de espaço para o texto.",
     "Les marges sont trop grandes pour la taille de page. Elles doivent laisser au "
     "moins 2 cm d’espace pour le texte.",
     "Die Ränder sind für die Seitengröße zu groß. Sie müssen mindestens 2 cm Platz "
     "für den Text lassen.",
     "I margini sono troppo grandi per le dimensioni della pagina. Devono lasciare "
     "almeno 2 cm di spazio per il testo.",
     "页边距相对于页面大小过大，必须为正文留出至少 2 厘米的空间。",
     "用紙サイズに対して余白が大きすぎます。本文用に 2 cm 以上の幅を残してください。"),
    ("The language will change when SWordBook restarts.\n\nRestart now?",
     "El idioma cambiará al reiniciar SWordBook.\n\n¿Reiniciar ahora?",
     "O idioma mudará quando o SWordBook for reiniciado.\n\nReiniciar agora?",
     "La langue changera au redémarrage de SWordBook.\n\nRedémarrer maintenant ?",
     "Die Sprache ändert sich beim Neustart von SWordBook.\n\nJetzt neu starten?",
     "La lingua cambierà al riavvio di SWordBook.\n\nRiavviare ora?",
     "重新启动 SWordBook 后将更改语言。\n\n现在重新启动吗？",
     "SWordBook を再起動すると言語が変わります。\n\n今すぐ再起動しますか？"),
    ("Restart now", "Reiniciar ahora", "Reiniciar agora", "Redémarrer maintenant",
     "Jetzt neu starten", "Riavvia ora", "立即重新启动", "今すぐ再起動"),
    ("Later", "Más tarde", "Mais tarde", "Plus tard", "Später", "Più tardi", "稍后",
     "後で"),
    # -- status bar ----------------------------------------------------------------
    ("Words: {n}", "Palabras: {n}", "Palavras: {n}", "Mots : {n}", "Wörter: {n}",
     "Parole: {n}", "字数：{n}", "単語数: {n}"),
    ("Characters (with spaces): {n}", "Caracteres (con espacios): {n}",
     "Caracteres (com espaços): {n}", "Caractères (espaces compris) : {n}",
     "Zeichen (mit Leerzeichen): {n}", "Caratteri (spazi inclusi): {n}",
     "字符数（含空格）：{n}", "文字数 (スペースを含む): {n}"),
    ("Characters (no spaces): {n}", "Caracteres (sin espacios): {n}",
     "Caracteres (sem espaços): {n}", "Caractères (sans espaces) : {n}",
     "Zeichen (ohne Leerzeichen): {n}", "Caratteri (spazi esclusi): {n}",
     "字符数（不含空格）：{n}", "文字数 (スペースを含まない): {n}"),
    ("Paragraphs: {n}", "Párrafos: {n}", "Parágrafos: {n}", "Paragraphes : {n}",
     "Absätze: {n}", "Paragrafi: {n}", "段落数：{n}", "段落数: {n}"),
    ("word", "palabra", "palavra", "mot", "Wort", "parola", "词", "語"),
    ("words", "palabras", "palavras", "mots", "Wörter", "parole", "词", "語"),
    ("{n} of {total} {words}", "{n} de {total} {words}", "{n} de {total} {words}",
     "{n} sur {total} {words}", "{n} von {total} {words}", "{n} di {total} {words}",
     "{total} {words}中的 {n} 个", "{total} {words}中 {n} {words}"),
    ("Page {p} of {t}", "Página {p} de {t}", "Página {p} de {t}", "Page {p} sur {t}",
     "Seite {p} von {t}", "Pagina {p} di {t}", "第 {p} 页，共 {t} 页", "{p} / {t} ページ"),
    ("Zoom out", "Alejar", "Reduzir", "Zoom arrière", "Verkleinern", "Riduci",
     "缩小", "縮小"),
    ("Zoom in", "Acercar", "Ampliar", "Zoom avant", "Vergrößern", "Ingrandisci",
     "放大", "拡大"),
    ("Zoom ({a} % to {b} %)", "Zoom ({a} % a {b} %)", "Zoom ({a} % a {b} %)",
     "Zoom ({a} % à {b} %)", "Zoom ({a} % bis {b} %)", "Zoom (da {a} % a {b} %)",
     "缩放（{a}% 到 {b}%）", "ズーム（{a}%～{b}%）"),
    ("Zoom level", "Nivel de zoom", "Nível de zoom", "Niveau de zoom", "Zoomstufe",
     "Livello di zoom", "缩放级别", "ズームレベル"),
    ("Fit to width", "Ajustar al ancho", "Ajustar à largura", "Ajuster à la largeur",
     "An Breite anpassen", "Adatta alla larghezza", "适应宽度", "幅に合わせる"),
    ("Whole page", "Página completa", "Página inteira", "Page entière", "Ganze Seite",
     "Pagina intera", "整页", "ページ全体"),
    # -- files ---------------------------------------------------------------------
    ("Word document (*.docx)", "Documento de Word (*.docx)", "Documento do Word (*.docx)",
     "Document Word (*.docx)", "Word-Dokument (*.docx)", "Documento di Word (*.docx)",
     "Word 文档 (*.docx)", "Word 文書 (*.docx)"),
    ("Documents (*.docx *.txt);;Word document (*.docx);;Text (*.txt)",
     "Documentos (*.docx *.txt);;Documento de Word (*.docx);;Texto (*.txt)",
     "Documentos (*.docx *.txt);;Documento do Word (*.docx);;Texto (*.txt)",
     "Documents (*.docx *.txt);;Document Word (*.docx);;Texte (*.txt)",
     "Dokumente (*.docx *.txt);;Word-Dokument (*.docx);;Text (*.txt)",
     "Documenti (*.docx *.txt);;Documento di Word (*.docx);;Testo (*.txt)",
     "文档 (*.docx *.txt);;Word 文档 (*.docx);;文本 (*.txt)",
     "文書 (*.docx *.txt);;Word 文書 (*.docx);;テキスト (*.txt)"),
    ("The file does not contain a Word document.",
     "El archivo no contiene un documento de Word.",
     "O arquivo não contém um documento do Word.",
     "Le fichier ne contient pas de document Word.",
     "Die Datei enthält kein Word-Dokument.",
     "Il file non contiene un documento di Word.",
     "该文件不包含 Word 文档。", "このファイルには Word 文書が含まれていません。"),
    ("Untitled", "Sin título", "Sem título", "Sans titre", "Unbenannt", "Senza titolo",
     "未命名", "無題"),
    ("Images", "Imágenes", "Imagens", "Images", "Bilder", "Immagini", "图片", "画像"),
    ("All files", "Todos los archivos", "Todos os arquivos", "Tous les fichiers",
     "Alle Dateien", "Tutti i file", "所有文件", "すべてのファイル"),
    # -- backup --------------------------------------------------------------------
    ("Could not create the backup file", "No se pudo crear el archivo de respaldo",
     "Não foi possível criar o arquivo de backup",
     "Impossible de créer le fichier de sauvegarde",
     "Die Sicherungsdatei konnte nicht erstellt werden",
     "Impossibile creare il file di backup",
     "无法创建备份文件", "バックアップファイルを作成できませんでした"),
    ("Could not recover", "No se pudo recuperar", "Não foi possível recuperar",
     "Récupération impossible", "Wiederherstellung fehlgeschlagen",
     "Impossibile recuperare", "无法恢复", "回復できませんでした"),
    ("Could not open the backup.", "No se pudo abrir el respaldo.",
     "Não foi possível abrir o backup.", "Impossible d’ouvrir la sauvegarde.",
     "Die Sicherung konnte nicht geöffnet werden.", "Impossibile aprire il backup.",
     "无法打开备份。", "バックアップを開けませんでした。"),
    ("The backup was recovered. Save the document to keep the changes.",
     "Se recuperó el respaldo. Guarda el documento para conservar los cambios.",
     "O backup foi recuperado. Salve o documento para manter as alterações.",
     "La sauvegarde a été récupérée. Enregistrez le document pour conserver les "
     "modifications.",
     "Die Sicherung wurde wiederhergestellt. Speichern Sie das Dokument, um die "
     "Änderungen zu behalten.",
     "Il backup è stato recuperato. Salva il documento per conservare le modifiche.",
     "已恢复备份。请保存文档以保留更改。",
     "バックアップを回復しました。変更を残すには文書を保存してください。"),
    ("Recover document", "Recuperar documento", "Recuperar documento",
     "Récupérer le document", "Dokument wiederherstellen", "Recupera documento",
     "恢复文档", "文書の回復"),
    ("SWordBook did not close properly.\n\nThere is a backup of “{name}” from {when}."
     "\nDo you want to recover it?",
     "SWordBook no se cerró correctamente.\n\nHay un respaldo de «{name}» del {when}."
     "\n¿Quieres recuperarlo?",
     "O SWordBook não foi fechado corretamente.\n\nHá um backup de “{name}” de {when}."
     "\nDeseja recuperá-lo?",
     "SWordBook ne s’est pas fermé correctement.\n\nIl existe une sauvegarde de "
     "« {name} » du {when}.\nVoulez-vous la récupérer ?",
     "SWordBook wurde nicht ordnungsgemäß beendet.\n\nEs gibt eine Sicherung von "
     "„{name}“ vom {when}.\nMöchten Sie sie wiederherstellen?",
     "SWordBook non è stato chiuso correttamente.\n\nC’è un backup di «{name}» del "
     "{when}.\nVuoi recuperarlo?",
     "SWordBook 未正常关闭。\n\n有一份“{name}”的备份（{when}）。\n要恢复它吗？",
     "SWordBook は正常に終了しませんでした。\n\n「{name}」のバックアップ（{when}）が"
     "あります。\n回復しますか？"),
    ("Recover", "Recuperar", "Recuperar", "Récupérer", "Wiederherstellen", "Recupera",
     "恢复", "回復"),
    ("Discard", "Descartar", "Descartar", "Ignorer", "Verwerfen", "Scarta", "丢弃", "破棄"),
    ("Backup found", "Respaldo encontrado", "Backup encontrado", "Sauvegarde trouvée",
     "Sicherung gefunden", "Backup trovato", "发现备份", "バックアップが見つかりました"),
    ("“{name}” has a more recent backup (from {when}) with changes that were not "
     "saved.\n\nDo you want to recover it?",
     "«{name}» tiene un respaldo más reciente (del {when}) con cambios que no se "
     "guardaron.\n\n¿Quieres recuperarlo?",
     "“{name}” tem um backup mais recente (de {when}) com alterações que não foram "
     "salvas.\n\nDeseja recuperá-lo?",
     "« {name} » a une sauvegarde plus récente (du {when}) avec des modifications non "
     "enregistrées.\n\nVoulez-vous la récupérer ?",
     "„{name}“ hat eine neuere Sicherung (vom {when}) mit nicht gespeicherten "
     "Änderungen.\n\nMöchten Sie sie wiederherstellen?",
     "«{name}» ha un backup più recente (del {when}) con modifiche non salvate.\n\n"
     "Vuoi recuperarlo?",
     "“{name}”有一份较新的备份（{when}），其中包含未保存的更改。\n\n要恢复它吗？",
     "「{name}」には保存されていない変更を含む新しいバックアップ（{when}）があります。"
     "\n\n回復しますか？"),
    ("Recover changes", "Recuperar cambios", "Recuperar alterações",
     "Récupérer les modifications", "Änderungen wiederherstellen", "Recupera modifiche",
     "恢复更改", "変更を回復"),
    ("Open the original", "Abrir el original", "Abrir o original", "Ouvrir l’original",
     "Original öffnen", "Apri l’originale", "打开原文件", "元のファイルを開く"),
    # -- window --------------------------------------------------------------------
    ("Unsaved changes", "Cambios sin guardar", "Alterações não salvas",
     "Modifications non enregistrées", "Nicht gespeicherte Änderungen",
     "Modifiche non salvate", "未保存的更改", "未保存の変更"),
    ("Save the changes to “{name}” before continuing?",
     "¿Guardar los cambios de «{name}» antes de continuar?",
     "Salvar as alterações de “{name}” antes de continuar?",
     "Enregistrer les modifications de « {name} » avant de continuer ?",
     "Änderungen an „{name}“ vor dem Fortfahren speichern?",
     "Salvare le modifiche a «{name}» prima di continuare?",
     "继续之前是否保存对“{name}”的更改？", "続行する前に「{name}」の変更を保存しますか？"),
    ("Save", "Guardar", "Salvar", "Enregistrer", "Speichern", "Salva", "保存", "保存"),
    ("Don't save", "No guardar", "Não salvar", "Ne pas enregistrer", "Nicht speichern",
     "Non salvare", "不保存", "保存しない"),
    ("Open", "Abrir", "Abrir", "Ouvrir", "Öffnen", "Apri", "打开", "開く"),
    ("Could not open", "No se pudo abrir", "Não foi possível abrir", "Ouverture impossible",
     "Öffnen fehlgeschlagen", "Impossibile aprire", "无法打开", "開けませんでした"),
    ("Could not open “{name}”.", "No se pudo abrir «{name}».",
     "Não foi possível abrir “{name}”.", "Impossible d’ouvrir « {name} ».",
     "„{name}“ konnte nicht geöffnet werden.", "Impossibile aprire «{name}».",
     "无法打开“{name}”。", "「{name}」を開けませんでした。"),
    ("Save as", "Guardar como", "Salvar como", "Enregistrer sous", "Speichern unter",
     "Salva con nome", "另存为", "名前を付けて保存"),
    ("Could not save", "No se pudo guardar", "Não foi possível salvar",
     "Enregistrement impossible", "Speichern fehlgeschlagen", "Impossibile salvare",
     "无法保存", "保存できませんでした"),
    ("Could not save “{name}”.", "No se pudo guardar «{name}».",
     "Não foi possível salvar “{name}”.", "Impossible d’enregistrer « {name} ».",
     "„{name}“ konnte nicht gespeichert werden.", "Impossibile salvare «{name}».",
     "无法保存“{name}”。", "「{name}」を保存できませんでした。"),
    ("Saved: {path}", "Guardado: {path}", "Salvo: {path}", "Enregistré : {path}",
     "Gespeichert: {path}", "Salvato: {path}", "已保存：{path}", "保存しました: {path}"),
    ("Insert image", "Insertar imagen", "Inserir imagem", "Insérer une image",
     "Bild einfügen", "Inserisci immagine", "插入图片", "画像の挿入"),
    ("Could not insert the image", "No se pudo insertar la imagen",
     "Não foi possível inserir a imagem", "Impossible d’insérer l’image",
     "Das Bild konnte nicht eingefügt werden", "Impossibile inserire l’immagine",
     "无法插入图片", "画像を挿入できませんでした"),
    ("“{name}”: {error}", "«{name}»: {error}", "“{name}”: {error}", "« {name} » : {error}",
     "„{name}“: {error}", "«{name}»: {error}", "“{name}”：{error}", "「{name}」: {error}"),
    ("Drag the image to move it, or its handles to resize it (Delete removes it)",
     "Arrastra la imagen para moverla o sus puntos para cambiar el tamaño (Supr la borra)",
     "Arraste a imagem para movê-la ou seus pontos para mudar o tamanho (Delete a apaga)",
     "Faites glisser l’image pour la déplacer, ou ses poignées pour la redimensionner "
     "(Suppr l’efface)",
     "Ziehen Sie das Bild, um es zu verschieben, oder seine Griffpunkte, um die Größe "
     "zu ändern (Entf löscht es)",
     "Trascina l’immagine per spostarla o i suoi punti per ridimensionarla "
     "(Canc la elimina)",
     "拖动图片可移动它，拖动控制点可调整大小（Delete 键删除）",
     "画像をドラッグすると移動、ハンドルをドラッグするとサイズ変更できます（Delete で削除）"),
    ("Select the text you want to comment on (or place the cursor on a word)",
     "Selecciona el texto que quieres comentar (o pon el cursor sobre una palabra)",
     "Selecione o texto que deseja comentar (ou coloque o cursor sobre uma palavra)",
     "Sélectionnez le texte à commenter (ou placez le curseur sur un mot)",
     "Markieren Sie den Text, den Sie kommentieren möchten (oder setzen Sie den Cursor "
     "auf ein Wort)",
     "Seleziona il testo da commentare (o posiziona il cursore su una parola)",
     "请选择要批注的文本（或将光标放在某个词上）",
     "コメントするテキストを選択してください（または単語の上にカーソルを置きます）"),
    ("Quit", "Salir", "Sair", "Quitter", "Beenden", "Esci", "退出", "終了"),
    # -- file bar --------------------------------------------------------------------
    ("New", "Nuevo", "Novo", "Nouveau", "Neu", "Nuovo", "新建", "新規"),
    ("Cut", "Cortar", "Recortar", "Couper", "Ausschneiden", "Taglia", "剪切", "切り取り"),
    ("Copy", "Copiar", "Copiar", "Copier", "Kopieren", "Copia", "复制", "コピー"),
    ("Paste", "Pegar", "Colar", "Coller", "Einfügen", "Incolla", "粘贴", "貼り付け"),
    ("Find and replace", "Buscar y reemplazar", "Localizar e substituir",
     "Rechercher et remplacer", "Suchen und ersetzen", "Trova e sostituisci",
     "查找和替换", "検索と置換"),
    ("Navigator: titles, subtitles and sections",
     "Navegador: títulos, subtítulos y secciones",
     "Navegador: títulos, subtítulos e seções",
     "Navigateur : titres, sous-titres et sections",
     "Navigator: Titel, Untertitel und Abschnitte",
     "Navigatore: titoli, sottotitoli e sezioni",
     "导航：标题、副标题和小节", "ナビゲーター：タイトル、サブタイトル、セクション"),
    ("Insert image (top and bottom, centered)",
     "Insertar imagen (Arriba y abajo, centrada)",
     "Inserir imagem (acima e abaixo, centralizada)",
     "Insérer une image (haut et bas, centrée)",
     "Bild einfügen (oben und unten, zentriert)",
     "Inserisci immagine (sopra e sotto, centrata)",
     "插入图片（上下型环绕，居中）", "画像の挿入（上下、中央揃え）"),
    ("New comment on the selected text", "Nuevo comentario en el texto seleccionado",
     "Novo comentário no texto selecionado", "Nouveau commentaire sur le texte sélectionné",
     "Neuer Kommentar zum markierten Text", "Nuovo commento sul testo selezionato",
     "对所选文本新建批注", "選択したテキストに新しいコメント"),
    ("Comments panel", "Panel de comentarios", "Painel de comentários",
     "Volet des commentaires", "Kommentarbereich", "Pannello commenti",
     "批注窗格", "コメント ウィンドウ"),
    ("Settings (page, margins, view, headings and language)",
     "Ajustes (página, márgenes, vista, títulos e idioma)",
     "Configurações (página, margens, exibição, títulos e idioma)",
     "Paramètres (page, marges, affichage, titres et langue)",
     "Einstellungen (Seite, Ränder, Ansicht, Überschriften und Sprache)",
     "Impostazioni (pagina, margini, vista, titoli e lingua)",
     "设置（页面、页边距、视图、标题和语言）", "設定（ページ、余白、表示、見出し、言語）"),
    ("Replace", "Reemplazar", "Substituir", "Remplacer", "Ersetzen", "Sostituisci",
     "替换", "置換"),
    ("Insert\nimage", "Insertar\nimagen", "Inserir\nimagem", "Insérer\nune image",
     "Bild\neinfügen", "Inserisci\nimmagine", "插入\n图片", "画像の\n挿入"),
    ("Ribbon", "Cinta", "Faixa de opções", "Ruban", "Menüband", "Barra multifunzione",
     "功能区", "リボン"),
    # -- find ------------------------------------------------------------------------
    ("Find", "Buscar", "Localizar", "Rechercher", "Suchen", "Trova", "查找", "検索"),
    ("Text to find", "Texto a buscar", "Texto a localizar", "Texte à rechercher",
     "Suchtext", "Testo da trovare", "要查找的文本", "検索する文字列"),
    ("Replace with", "Reemplazar con", "Substituir por", "Remplacer par", "Ersetzen durch",
     "Sostituisci con", "替换为", "置換後の文字列"),
    ("Match case", "Coincidir mayúsculas", "Diferenciar maiúsculas",
     "Respecter la casse", "Groß-/Kleinschreibung beachten", "Maiuscole/minuscole",
     "区分大小写", "大文字と小文字を区別"),
    ("Whole words", "Palabras completas", "Palavras inteiras", "Mots entiers",
     "Ganze Wörter", "Parole intere", "全字匹配", "単語単位"),
    ("Previous", "Anterior", "Anterior", "Précédent", "Zurück", "Precedente", "上一个", "前へ"),
    ("Previous match (Shift+Enter)", "Coincidencia anterior (Mayús+Enter)",
     "Ocorrência anterior (Shift+Enter)", "Occurrence précédente (Maj+Entrée)",
     "Vorheriger Treffer (Umschalt+Eingabe)", "Corrispondenza precedente (Maiusc+Invio)",
     "上一个匹配项（Shift+Enter）", "前の一致（Shift+Enter）"),
    ("Next", "Siguiente", "Próximo", "Suivant", "Weiter", "Successivo", "下一个", "次へ"),
    ("Next match (Enter)", "Coincidencia siguiente (Enter)", "Próxima ocorrência (Enter)",
     "Occurrence suivante (Entrée)", "Nächster Treffer (Eingabe)",
     "Corrispondenza successiva (Invio)", "下一个匹配项（Enter）", "次の一致（Enter）"),
    ("Replaces the selected match and moves to the next one",
     "Reemplaza la coincidencia seleccionada y pasa a la siguiente",
     "Substitui a ocorrência selecionada e passa para a próxima",
     "Remplace l’occurrence sélectionnée et passe à la suivante",
     "Ersetzt den markierten Treffer und springt zum nächsten",
     "Sostituisce la corrispondenza selezionata e passa alla successiva",
     "替换所选匹配项并转到下一个", "選択した一致を置換して次へ移動します"),
    ("Replace all", "Reemplazar todo", "Substituir tudo", "Tout remplacer", "Alle ersetzen",
     "Sostituisci tutto", "全部替换", "すべて置換"),
    ("Replaces all matches (a single undo step)",
     "Reemplaza todas las coincidencias (un solo paso de deshacer)",
     "Substitui todas as ocorrências (uma única etapa de desfazer)",
     "Remplace toutes les occurrences (une seule étape d’annulation)",
     "Ersetzt alle Treffer (ein einziger Rückgängig-Schritt)",
     "Sostituisce tutte le corrispondenze (un solo passo di annullamento)",
     "替换所有匹配项（一步即可撤销）", "すべての一致を置換します（1 回で元に戻せます）"),
    ("Close", "Cerrar", "Fechar", "Fermer", "Schließen", "Chiudi", "关闭", "閉じる"),
    ("Continued from the end", "Se continuó desde el final", "Continuou a partir do fim",
     "Recherche reprise depuis la fin", "Am Ende fortgesetzt", "Ripreso dalla fine",
     "已从末尾继续", "末尾から続行しました"),
    ("Continued from the beginning", "Se continuó desde el principio",
     "Continuou a partir do início", "Recherche reprise depuis le début",
     "Am Anfang fortgesetzt", "Ripreso dall’inizio", "已从开头继续", "先頭から続行しました"),
    ("No matches", "Sin coincidencias", "Nenhuma ocorrência", "Aucune occurrence",
     "Keine Treffer", "Nessuna corrispondenza", "无匹配项", "一致なし"),
    ("Replaced 1 match", "Se reemplazó 1 coincidencia", "1 ocorrência substituída",
     "1 occurrence remplacée", "1 Treffer ersetzt", "1 corrispondenza sostituita",
     "已替换 1 处", "1 件置換しました"),
    ("Replaced {n} matches", "Se reemplazaron {n} coincidencias",
     "{n} ocorrências substituídas", "{n} occurrences remplacées", "{n} Treffer ersetzt",
     "{n} corrispondenze sostituite", "已替换 {n} 处", "{n} 件置換しました"),
    ("{i} of {n} matches", "{i} de {n} coincidencias", "{i} de {n} ocorrências",
     "{i} sur {n} occurrences", "{i} von {n} Treffern", "{i} di {n} corrispondenze",
     "第 {i} 个，共 {n} 个匹配项", "{i} / {n} 件"),
    ("1 match", "1 coincidencia", "1 ocorrência", "1 occurrence", "1 Treffer",
     "1 corrispondenza", "1 个匹配项", "1 件"),
    ("{n} matches", "{n} coincidencias", "{n} ocorrências", "{n} occurrences",
     "{n} Treffer", "{n} corrispondenze", "{n} 个匹配项", "{n} 件"),
    # -- navigator ---------------------------------------------------------------------
    ("Navigator", "Navegador", "Navegador", "Navigateur", "Navigator", "Navigatore",
     "导航", "ナビゲーター"),
    ("Filter headings", "Filtrar títulos", "Filtrar títulos", "Filtrer les titres",
     "Überschriften filtern", "Filtra titoli", "筛选标题", "見出しを絞り込む"),
    ("No headings yet.\nUse the {title}, {subtitle} or {section} styles to see them here.",
     "Aún no hay títulos.\nUsa los estilos {title}, {subtitle} o {section} para verlos aquí.",
     "Ainda não há títulos.\nUse os estilos {title}, {subtitle} ou {section} para vê-los aqui.",
     "Aucun titre pour l’instant.\nUtilisez les styles {title}, {subtitle} ou {section} "
     "pour les voir ici.",
     "Noch keine Überschriften.\nVerwenden Sie die Formatvorlagen {title}, {subtitle} "
     "oder {section}, um sie hier zu sehen.",
     "Ancora nessun titolo.\nUsa gli stili {title}, {subtitle} o {section} per vederli qui.",
     "还没有标题。\n使用“{title}”“{subtitle}”或“{section}”样式即可在此处看到它们。",
     "見出しはまだありません。\n「{title}」「{subtitle}」「{section}」スタイルを使うと"
     "ここに表示されます。"),
    ("Heading {n}", "Título {n}", "Título {n}", "Titre {n}", "Überschrift {n}", "Titolo {n}",
     "标题 {n}", "見出し {n}"),
    ("(no text)", "(sin texto)", "(sem texto)", "(sans texte)", "(kein Text)",
     "(nessun testo)", "（无文本）", "（テキストなし）"),
    # -- comments ----------------------------------------------------------------------
    ("Comment {n}", "Comentario {n}", "Comentário {n}", "Commentaire {n}", "Kommentar {n}",
     "Commento {n}", "批注 {n}", "コメント {n}"),
    ("Change the highlight color", "Cambiar el color del resaltado",
     "Alterar a cor do realce", "Changer la couleur du surlignage",
     "Hervorhebungsfarbe ändern", "Cambia il colore dell’evidenziazione",
     "更改突出显示颜色", "強調表示の色を変更"),
    ("Delete the comment (can be undone with Ctrl+Z)",
     "Eliminar el comentario (se puede deshacer con Ctrl+Z)",
     "Excluir o comentário (pode ser desfeito com Ctrl+Z)",
     "Supprimer le commentaire (annulable avec Ctrl+Z)",
     "Kommentar löschen (mit Strg+Z rückgängig zu machen)",
     "Elimina il commento (si può annullare con Ctrl+Z)",
     "删除批注（可用 Ctrl+Z 撤销）", "コメントを削除（Ctrl+Z で元に戻せます）"),
    ("Write a comment…", "Escribe un comentario…", "Escreva um comentário…",
     "Écrivez un commentaire…", "Kommentar schreiben…", "Scrivi un commento…",
     "输入批注…", "コメントを入力…"),
    ("Comment color", "Color del comentario", "Cor do comentário",
     "Couleur du commentaire", "Kommentarfarbe", "Colore del commento",
     "批注颜色", "コメントの色"),
    ("“{text}”", "«{text}»", "“{text}”", "« {text} »", "„{text}“", "«{text}»",
     "“{text}”", "「{text}」"),
    ("Comments", "Comentarios", "Comentários", "Commentaires", "Kommentare", "Commenti",
     "批注", "コメント"),
    ("No comments yet.\nSelect text and click New comment (Ctrl+Alt+M).",
     "Aún no hay comentarios.\nSelecciona texto y pulsa Nuevo comentario (Ctrl+Alt+M).",
     "Ainda não há comentários.\nSelecione um texto e clique em Novo comentário "
     "(Ctrl+Alt+M).",
     "Aucun commentaire pour l’instant.\nSélectionnez du texte et cliquez sur Nouveau "
     "commentaire (Ctrl+Alt+M).",
     "Noch keine Kommentare.\nMarkieren Sie Text und klicken Sie auf „Neuer Kommentar“ "
     "(Strg+Alt+M).",
     "Ancora nessun commento.\nSeleziona del testo e premi Nuovo commento (Ctrl+Alt+M).",
     "还没有批注。\n选择文本后点击“新建批注”（Ctrl+Alt+M）。",
     "コメントはまだありません。\nテキストを選択して［新しいコメント］"
     "（Ctrl+Alt+M）をクリックします。"),
]

_COLUMNS = ("es", "pt", "fr", "de", "it", "zh", "ja")
_TRANSLATIONS = {c: {} for c in _COLUMNS}
for _row in _ROWS:
    assert len(_row) == 8, _row[0]
    for _c, _v in zip(_COLUMNS, _row[1:]):
        _TRANSLATIONS[_c][_row[0]] = _v
del _row, _c, _v


if __name__ == "__main__":
    if sys.platform == "win32":
        # Without this, Windows shows the Python icon in the taskbar
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "SWordBook.SWordBook")
        except Exception:
            pass
    app = QApplication(sys.argv)
    # The internal name doesn't change: the folder where backups are saved
    # depends on it (so existing ones aren't lost)
    app.setApplicationName("Editor de libros")
    app.setApplicationDisplayName("SWordBook")
    app.setOrganizationName("EditorLibros")
    app.setWindowIcon(app_icon())
    # The language, before creating any window
    config = Settings()
    set_language(config.language())
    install_qt_translations(app)
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setPalette(light_palette())
    app.setStyleSheet(STYLESHEET)
    window = MainWindow(config)
    window.show_window()
    # If the program closed abruptly last time, offer to recover
    QTimer.singleShot(0, window.offer_recovery)
    sys.exit(app.exec())
