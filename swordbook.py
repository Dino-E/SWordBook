# =============================================================================
# DESCRIPCIÓN
# -----------------------------------------------------------------------------
# SWordBook es un editor de documentos de texto inspirado en WordPad, diseñado
# para ofrecer una experiencia de escritura sencilla, intuitiva y personalizable.
# Permite crear, editar y guardar documentos en formato DOCX, proporcionando
# herramientas esenciales para la edición y organización del contenido.
#
# FUNCIONALIDADES PRINCIPALES
# -----------------------------------------------------------------------------
# - Creación, edición y guardado de documentos DOCX.
# - Sistema de seguridad y recuperación de archivos para ayudar a prevenir
#   la pérdida de información ante cierres inesperados o errores.
# - Sistema de comentarios para realizar anotaciones en los documentos.
# - Panel de navegación para explorar y organizar el contenido.
# - Herramientas de búsqueda de texto dentro del documento.
# - Opciones de formato de texto: fuente, tamaño, color y otros atributos.
# - Herramientas de formato de párrafo y alineación.
# - Personalización de colores y apariencia de la interfaz.
#
# AUTORÍA Y CRÉDITOS
# -----------------------------------------------------------------------------
# Desarrollado por DinoSoft.1998.
# Claude.ai se utilizó como herramienta de asistencia durante el desarrollo.
#
# LICENCIA
# -----------------------------------------------------------------------------
# Este software se distribuye bajo los términos de la licencia MIT.
# Consulta el archivo LICENSE para conocer los términos completos de uso,
# modificación, distribución y redistribución del software.
#
# AVISO
# -----------------------------------------------------------------------------
# Microsoft Word y DOCX son marcas y denominaciones de sus respectivos
# propietarios. Este proyecto es un editor independiente inspirado en
# WordPad y no está afiliado ni respaldado oficialmente por Microsoft.
# =============================================================================

import sys
import os
import re
import copy
import math
import string
import json
import bisect
import zipfile
import datetime
import getpass
import posixpath
import configparser
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape
from PySide6.QtWidgets import (
    QFileDialog, QTabBar, QDockWidget, QTreeWidget, QTreeWidgetItem,
    QScrollArea, QSizePolicy, QSplitter, QScrollBar,
)
from PySide6.QtGui import (
    QTextDocument, QTextDocumentFragment, QPainterPath,
)
from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QUrl, QMimeData
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QToolBar, QToolButton, QFontComboBox,
    QComboBox, QMenu, QWidget, QWidgetAction, QGridLayout, QColorDialog, QLabel,
    QVBoxLayout, QStyleFactory, QSpinBox, QDialog, QFormLayout, QGroupBox,
    QPushButton, QHBoxLayout, QGraphicsView, QGraphicsScene, QFrame,
    QAbstractScrollArea, QStatusBar, QSlider, QDoubleSpinBox,
    QDialogButtonBox, QMessageBox, QListWidget, QListWidgetItem,
    QAbstractItemView, QLineEdit, QCheckBox,
)
from PySide6.QtGui import (
    QFont, QColor, QLinearGradient, QTextBlockFormat, QTextCharFormat, QTextCursor, QIcon, QPixmap,
    QPainter, QPen, QBrush, QAction, QKeySequence, QPalette, QIntValidator,
    QActionGroup, QShortcut, QFontMetrics, QTextFormat, QTextListFormat, QPolygonF, QDoubleValidator, QTransform,
    QImage, QImageReader, QImageIOHandler, QTextImageFormat, QFontMetricsF,
)
from PySide6.QtCore import (
    Qt, QSize, QSizeF, QRectF, QPointF, Signal, QLocale, QObject, QEvent, QTimer,
    QKeyCombination,
    QDateTime, QTranslator, QLibraryInfo, QStandardPaths,
)

# ---------------------------------------------------------------------------
# Estilo "Body" recomendado para libros electrónicos
# ---------------------------------------------------------------------------
BODY = {
    "font_family": "Georgia",           # Serif legible en pantalla
    "font_size": 12,                    # pt (≈1em en lectores)
    "font_color": "#1A1A1A",            # Casi negro, menos fatiga visual
    "alignment": Qt.AlignJustify,       # Justificado, típico de libros
    "line_spacing": 150,                # % (interlineado 1.5)
    "space_before": 0,                  # px
    "space_after": 6,                   # px
    "indent_left": 0,                   # px
    "indent_right": 0,                  # px
    "first_line_indent": 0,             # px (sangría de primera línea)
    "letter_spacing": 0,                # px (espaciado entre letras)
    "bold": False,
    "italic": False,
}


# ---------------------------------------------------------------------------
# Galería de estilos: cada estilo guarda todos sus valores, su nivel de
# esquema (jerarquía de títulos) y su estilo siguiente (el que toma el
# párrafo nuevo al pulsar Enter). Los predeterminados no se pueden borrar.
# ---------------------------------------------------------------------------
CAMPOS_ESTILO = list(BODY)
PROP_ESTILO = QTextFormat.UserProperty + 1      # id del estilo del párrafo


# Nombres que se muestran (cinta, Ajustes de estilo, navegador…). Los ids
# internos siguen en inglés: con ellos se guardan los .docx y Word reconoce
# el Normal, los títulos y los encabezados
NOMBRES_ESTILO = {
    "Body": "Cuerpo", "Title": "Título", "Subtitle": "Subtítulo",
    "Heading 1": "Encabezado 1", "Heading 2": "Encabezado 2",
    "Heading 3": "Encabezado 3",
}


def _predeterminados():
    """Estilos para libros electrónicos: misma familia que el cuerpo,
    tamaños en escala clara, títulos sin justificar y con más espacio antes
    que después (el título queda unido a su texto)."""
    def e(id_, base, siguiente, nivel, **propio):
        return {"id": id_, "nombre": NOMBRES_ESTILO.get(id_, id_),
                "visible": True, "base": base,
                "siguiente": siguiente, "nivel": nivel, "propio": propio,
                "fijo": True}
    lista = [
        e("Body", None, "Body", 0, **BODY),
        e("Title", "Body", "Subtitle", 0, font_size=28, bold=True,
          alignment=Qt.AlignHCenter, line_spacing=120, space_before=48,
          space_after=12),
        e("Subtitle", "Body", "Body", 0, font_size=16, italic=True,
          font_color="#4A4A4A", alignment=Qt.AlignHCenter, line_spacing=130,
          space_after=36),
        e("Heading 1", "Body", "Body", 1, font_size=22, bold=True,
          alignment=Qt.AlignLeft, line_spacing=120, space_before=36,
          space_after=12),
        e("Heading 2", "Heading 1", "Body", 2, font_size=18, space_before=24,
          space_after=8),
        e("Heading 3", "Heading 2", "Body", 3, font_size=14, line_spacing=130,
          space_before=18, space_after=6),
    ]
    # Cada estilo guarda todos sus valores (sin herencia entre estilos)
    datos = {x["id"]: x for x in lista}
    for x in lista:
        completo = dict(BODY)
        cadena, b = [], x
        while b is not None:
            cadena.append(b)
            b = datos.get(b["base"]) if b["base"] else None
        for c in reversed(cadena):
            completo.update(c["propio"])
        x["completo"] = completo
    for x in lista:
        x["propio"], x["base"] = x.pop("completo"), None
    return datos, [x["id"] for x in lista]


class GestorEstilos(QObject):
    cambiado = Signal()

    def __init__(self, datos=None, orden=None):
        super().__init__()
        if datos is None:
            datos, orden = _predeterminados()
        self.datos = datos
        self.orden = orden

    def copia(self):
        return GestorEstilos(copy.deepcopy(self.datos), list(self.orden))

    def adoptar(self, otro):
        self.datos = copy.deepcopy(otro.datos)
        self.orden = list(otro.orden)
        self.cambiado.emit()

    def resuelto(self, id_):
        """Todos los valores del estilo."""
        e = self.datos.get(id_) or self.datos["Body"]
        cadena, vistos = [], set()
        while e is not None and e["id"] not in vistos:
            vistos.add(e["id"])
            cadena.append(e)
            e = self.datos.get(e["base"]) if e["base"] else None
        r = dict(BODY)
        for x in reversed(cadena):
            r.update(x["propio"])
        return r

    def nombre(self, id_):
        e = self.datos.get(id_)
        return (e["nombre"] or e["id"]) if e else id_

    def nuevo(self, origen="Body"):
        n = 1
        while f"estilo_{n}" in self.datos:
            n += 1
        id_ = f"estilo_{n}"
        self.datos[id_] = {"id": id_, "nombre": f"Estilo {n}", "visible": True,
                           "base": None, "siguiente": id_, "nivel": 0,
                           "propio": self.resuelto(origen), "fijo": False}
        self.orden.append(id_)
        return id_

    def eliminar(self, id_):
        e = self.datos.get(id_)
        if not e or e["fijo"]:
            return
        for h in self.datos.values():
            if h["siguiente"] == id_:
                h["siguiente"] = h["id"]
        del self.datos[id_]
        self.orden.remove(id_)

    def restablecer(self, id_):
        datos, _ = _predeterminados()
        if id_ in datos:
            self.datos[id_] = datos[id_]


GESTOR = GestorEstilos()


def formato_caracter(id_, gestor=None):
    e = (gestor or GESTOR).resuelto(id_)
    fmt = QTextCharFormat()
    fmt.setFontFamilies([e["font_family"]])
    fmt.setFontPointSize(e["font_size"])
    fmt.setForeground(QColor(e["font_color"]))
    fmt.setFontWeight(QFont.Bold if e["bold"] else QFont.Normal)
    fmt.setFontItalic(e["italic"])
    aplicar_letras(fmt, e["letter_spacing"])
    return fmt


def formato_parrafo(id_, gestor=None):
    g = gestor or GESTOR
    e = g.resuelto(id_)
    fmt = QTextBlockFormat()
    fmt.setAlignment(e["alignment"])
    fmt.setLineHeight(e["line_spacing"],
                      QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
    fmt.setTopMargin(e["space_before"])
    fmt.setBottomMargin(e["space_after"])
    fmt.setLeftMargin(e["indent_left"])
    fmt.setRightMargin(e["indent_right"])
    fmt.setTextIndent(e["first_line_indent"])
    fmt.setProperty(PROP_ESTILO, id_ if id_ in g.datos else "Body")
    fmt.setHeadingLevel(g.datos.get(id_, {}).get("nivel", 0))   # jerarquía
    return fmt


TAMANOS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]

# Paleta de color de fuente (estilo procesador de textos, 10 columnas)
PALETA_FUENTE = [
    "#000000", "#434343", "#666666", "#999999", "#B7B7B7",
    "#CCCCCC", "#D9D9D9", "#EFEFEF", "#F3F3F3", "#FFFFFF",
    "#980000", "#FF0000", "#FF9900", "#FFFF00", "#00FF00",
    "#00FFFF", "#4A86E8", "#0000FF", "#9900FF", "#FF00FF",
    "#A61C00", "#CC0000", "#E69138", "#F1C232", "#6AA84F",
    "#45818E", "#3C78D8", "#3D85C6", "#674EA7", "#A64D79",
]

# Paleta de resaltado: tonos suaves que no tapan el texto
PALETA_RESALTADO = [
    "#FFF59D", "#C5E1A5", "#80DEEA", "#90CAF9", "#CE93D8",
    "#F48FB1", "#FFCC80", "#BCAAA4", "#E0E0E0", "#FFFF00",
]

COLOR_ICONO = "#3A3A3A"

# Sangría: cada paso desplaza el párrafo; el máximo evita columnas ilegibles
SANGRIA_PASO = 30          # px por cada clic de aumentar/reducir
SANGRIA_MAX = 240          # px, límite para sangría izquierda y derecha
PRIMERA_LINEA_MAX = 120    # px, límite de sangría de primera línea
ESPACIO_MAX = 200          # px, límite de espacio antes/después
LETRAS_MIN = -5           # px, espaciado entre letras mínimo (condensado)
LETRAS_MAX = 20           # px, espaciado entre letras máximo (expandido)
INTERLINEADOS = ["1.0", "1.15", "1.5", "2.0", "2.5", "3.0"]

# ---------------------------------------------------------------------------
# Hoja (medidas en cm). Vista "Sin páginas": una sola hoja continua del ancho
# de la página, sin cortes entre páginas
# ---------------------------------------------------------------------------
# Hueco gris entre la barra y el borde superior de la hoja, en px de
# pantalla: el mismo con cualquier zoom (no crece al acercar). Es un ajuste
# de la vista: se cambia en Configurar página y no se guarda en el .docx
ESPACIO_SUPERIOR = 20
PAGINA = {
    "ancho": 21.59, "alto": 27.94,                           # Carta
    "superior": 2.5, "inferior": 2.5, "izquierdo": 3.0, "derecho": 3.0,
    "espacio_superior": ESPACIO_SUPERIOR,
    # Vista (no se guardan en el .docx): tema de la interfaz y color de la hoja
    "tema": "claro", "tema_interfaz": "#F7F7F7", "tema_fondo": "#E4E4E4",
    "color_hoja": "#FFFFFF",
}
SEPARACION = 20        # px alrededor de la hoja
TAMANOS_PAGINA = [         # nombre, ancho, alto (vertical)
    ("Carta", 21.59, 27.94),
    ("Oficio", 21.59, 34.01),
    ("Legal", 21.59, 35.56),
    ("A4", 21.0, 29.7),
    ("A5", 14.8, 21.0),
    ("Libro 6 × 9 in", 15.24, 22.86),
    ("Libro 5.5 × 8.5 in", 13.97, 21.59),
    ("Libro 5 × 8 in", 12.7, 20.32),
]

# Zoom: entre ZOOM_MIN y el zoom en que la hoja llena el ancho de la vista
# (ZOOM_MAX es solo un tope de seguridad para ventanas enormes)
ZOOM_MIN = 0.10
ZOOM_MAX = 5.00
ZOOM_PASO = 10             # puntos de % por clic de −/+ o muesca de Ctrl + rueda
SCROLL_MUESCA = 100        # px que avanza cada muesca de la rueda
MENU_ZOOM = [50, 75, 100, 125, 150, 200, 300, 400]



# ---------------------------------------------------------------------------
# Hoja de estilo de la interfaz (tema claro y neutro)
# ---------------------------------------------------------------------------
PLANTILLA_ESTILO = string.Template("""
QMainWindow { background: $ventana; }

QToolBar {
    background: $barra;
    border: none;
    border-bottom: 1px solid $borde;
    padding: 4px 8px;
    spacing: 2px;
}
QToolBar::separator {
    background: $separador;
    width: 1px;
    margin: 5px 6px;
}
/* Secciones dentro de una fila: sin borde ni relleno propios */
QToolBar#seccion {
    background: transparent;
    border: none;
    padding: 0px;
    margin: 0px;
}
/* Botón » de lo que no cabe */
QToolBar#seccion > QToolButton#qt_toolbar_ext_button {
    padding: 0px;
    border-radius: 3px;
}
/* La fila de arriba no lleva línea: las dos filas forman una sola barra */
QToolBar#filaSuperior { border-bottom: none; padding-bottom: 0px; }
QToolBar#filaEstilos { padding-top: 2px; }
QWidget#filaHerramientas { background: transparent; }

QToolButton {
    background: transparent;
    border: 1px solid transparent;
    border-radius: 4px;
    padding: 3px;
}
QToolButton:hover   { background: $hover; border-color: $hover_borde; }
QToolButton:pressed { background: $pulsado; }
QToolButton:checked { background: $marcado; border-color: $marcado_borde; }
QToolButton[popupMode="1"] { padding-right: 15px; }
QToolButton::menu-button {
    border: none;
    border-left: 1px solid transparent;
    width: 13px;
}
QToolButton::menu-button:hover { border-left-color: $hover_borde; }

QToolBar QLabel#tituloGrupo {
    color: $texto_sec;
    font-size: 11px;
    font-weight: bold;
    padding: 0 6px 0 2px;
}
QToolBar QLabel#etiquetaCampo {
    color: $texto_barra;
    padding: 0 3px 0 8px;
}

QComboBox QAbstractItemView {
    background: $campo;
    color: $texto;
    border: 1px solid $campo_borde;
    selection-background-color: $marcado;
    selection-color: $texto;
    outline: none;
}

QMenu {
    background: $campo;
    border: 1px solid $campo_borde;
    border-radius: 6px;
    padding: 6px;
}
QMenu::item {
    color: $texto;
    padding: 5px 16px;
    border-radius: 4px;
}
QMenu::item:selected { background: $marcado; }
QMenu::item:disabled { color: $texto_suave; }
QMenu::separator { height: 1px; background: $menu_sep; margin: 5px 4px; }


QGraphicsView#lienzo { background: $lienzo; border: none; }

QStatusBar {
    background: $barra;
    border-top: 1px solid $borde;
    color: $texto_barra;
}
QStatusBar::item { border: none; }
QStatusBar QToolButton { padding: 1px 6px; }

QTextEdit {
    background: $campo;
    color: $texto;
    border: none;
    selection-background-color: $seleccion_suave;
    selection-color: $texto;
}
""")


# ---------------------------------------------------------------------------
# Temas de la interfaz. Solo cambian barras, paneles, menús, diálogos y el
# fondo alrededor de la hoja: el contenido (texto, colores del documento,
# imágenes) y la hoja (que tiene su propio color) no cambian
# ---------------------------------------------------------------------------
COLOR_CONTENIDO = "#1A1A1A"        # texto sin color propio, en la hoja
SELECCION_CONTENIDO = "#B3D4FC"    # selección dentro de la hoja
COLOR_HOJA = "#FFFFFF"
TEMAS = {
    "claro": {
        "ventana": "#F3F3F3", "barra": "#F7F7F7", "borde": "#DADADA",
        "separador": "#D6D6D6", "hover": "#E9E9E9", "hover_borde": "#D2D2D2",
        "pulsado": "#DDDDDD", "marcado": "#DCE8F7", "marcado_borde": "#9DBCE3",
        "texto": "#1A1A1A", "texto_sec": "#6A6A6A", "texto_suave": "#808080",
        "texto_barra": "#3A3A3A", "campo": "#FFFFFF", "campo_borde": "#C8C8C8",
        "menu_sep": "#E2E2E2", "lienzo": "#E4E4E4", "borde_hoja": "#C8C8C8",
        "icono": "#3A3A3A", "seleccion": "#4A86E8", "seleccion_suave": "#B3D4FC",
    },
    "oscuro": {
        "ventana": "#1E1E1E", "barra": "#2A2A2A", "borde": "#3B3B3B",
        "separador": "#474747", "hover": "#383838", "hover_borde": "#4C4C4C",
        "pulsado": "#444444", "marcado": "#2E4A6E", "marcado_borde": "#4F7CB3",
        "texto": "#E6E6E6", "texto_sec": "#A6A6A6", "texto_suave": "#8A8A8A",
        "texto_barra": "#D2D2D2", "campo": "#232323", "campo_borde": "#4A4A4A",
        "menu_sep": "#3A3A3A", "lienzo": "#141414", "borde_hoja": "#000000",
        "icono": "#D6D6D6", "seleccion": "#3D7BD9", "seleccion_suave": "#2E4A6E",
    },
}
NOMBRES_TEMA = [("claro", "Claro"), ("oscuro", "Oscuro"),
                ("personalizado", "Personalizado")]


def _mezcla(a, b, t):
    """Color entre a (t=0) y b (t=1)."""
    a, b = QColor(a), QColor(b)
    return QColor(round(a.red() + (b.red() - a.red()) * t),
                  round(a.green() + (b.green() - a.green()) * t),
                  round(a.blue() + (b.blue() - a.blue()) * t)).name().upper()


def contraste(a, b):
    """Relación de contraste WCAG entre dos colores (1 a 21)."""
    def lum(c):
        c = QColor(c)
        v = []
        for x in (c.redF(), c.greenF(), c.blueF()):
            v.append(x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4)
        return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2]
    l1, l2 = sorted((lum(a), lum(b)), reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)


def tema_personalizado(interfaz, fondo):
    """Tema completo a partir de dos colores: el de la interfaz (barras y
    paneles) y el del fondo alrededor de la hoja. Bordes, estados y textos
    se calculan para que todo siga siendo legible."""
    base = QColor(interfaz)
    oscuro = base.lightness() < 128
    fg = "#E6E6E6" if oscuro else "#1A1A1A"
    acento = "#3D7BD9" if oscuro else "#4A86E8"
    b = base.name().upper()
    campo = _mezcla(b, "#000000", 0.18) if oscuro else _mezcla(b, "#FFFFFF", 0.75)
    return {
        "ventana": _mezcla(b, "#000000", 0.28 if oscuro else 0.016),
        "barra": b, "borde": _mezcla(b, fg, 0.13), "separador": _mezcla(b, fg, 0.10),
        "hover": _mezcla(b, fg, 0.07), "hover_borde": _mezcla(b, fg, 0.15),
        "pulsado": _mezcla(b, fg, 0.12),
        "marcado": _mezcla(b, acento, 0.30 if oscuro else 0.16),
        "marcado_borde": _mezcla(b, acento, 0.55),
        "texto": fg, "texto_sec": _mezcla(fg, b, 0.40), "texto_suave": _mezcla(fg, b, 0.50),
        "texto_barra": _mezcla(fg, b, 0.12), "campo": campo,
        "campo_borde": _mezcla(campo, fg, 0.22), "menu_sep": _mezcla(campo, fg, 0.10),
        "lienzo": QColor(fondo).name().upper(),
        "borde_hoja": _mezcla(QColor(fondo).name(), "#000000", 0.25),
        "icono": _mezcla(fg, b, 0.12), "seleccion": acento,
        "seleccion_suave": _mezcla(b, acento, 0.35),
    }


def resolver_tema(config):
    """Tokens del tema elegido en la configuración de la vista."""
    nombre = config.get("tema", "claro")
    if nombre == "personalizado":
        return tema_personalizado(config.get("tema_interfaz", TEMAS["claro"]["barra"]),
                                  config.get("tema_fondo", TEMAS["claro"]["lienzo"]))
    return dict(TEMAS.get(nombre, TEMAS["claro"]))


def estilo_app(t):
    return PLANTILLA_ESTILO.substitute(t)


def paleta_tema(t):
    p = QPalette()
    for rol, k in ((QPalette.Window, "ventana"), (QPalette.WindowText, "texto"),
                   (QPalette.Base, "campo"), (QPalette.AlternateBase, "barra"),
                   (QPalette.Text, "texto"), (QPalette.Button, "barra"),
                   (QPalette.ButtonText, "texto"), (QPalette.Highlight, "seleccion"),
                   (QPalette.PlaceholderText, "texto_suave"),
                   (QPalette.ToolTipBase, "campo"), (QPalette.ToolTipText, "texto")):
        p.setColor(rol, QColor(t[k]))
    p.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    # Relieves de Fusion (marcos, flechas) a partir del color de los botones
    boton = QColor(t["barra"])
    p.setColor(QPalette.Light, boton.lighter(150))
    p.setColor(QPalette.Midlight, boton.lighter(125))
    p.setColor(QPalette.Mid, QColor(t["campo_borde"]))
    p.setColor(QPalette.Dark, QColor(t["borde"]).darker(130))
    p.setColor(QPalette.Shadow, QColor("#000000"))
    for rol in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        p.setColor(QPalette.Disabled, rol, QColor(t["texto_suave"]))
    return p


def recolorear_icono(icono, de, a):
    """Pasa un icono dibujado con el tema `de` al tema `a`. Cambian los
    píxeles que están en la gama del tema anterior —entre su color de icono
    y su color de barra, sean grises o teñidos (temas personalizados)— y se
    conservan los de color propio (la barra roja del color de letra, el
    amarillo del resaltado…)."""
    if icono.isNull():
        return icono
    tamanos = icono.availableSizes()
    pm = icono.pixmap(tamanos[0]) if tamanos else icono.pixmap(QSize(72, 72))
    if pm.isNull():
        return icono
    img = pm.toImage().convertToFormat(QImage.Format_ARGB32)
    ci, cb = QColor(de["icono"]), QColor(de["barra"])
    ni, nb = QColor(a["icono"]), QColor(a["barra"])
    i0 = (ci.red(), ci.green(), ci.blue())
    eje = (cb.red() - i0[0], cb.green() - i0[1], cb.blue() - i0[2])
    largo2 = sum(x * x for x in eje) or 1
    cache = {}

    def nuevo_color(r, g, bl):
        clave = (r, g, bl)
        if clave in cache:
            return cache[clave]
        d = (r - i0[0], g - i0[1], bl - i0[2])
        t = (d[0] * eje[0] + d[1] * eje[1] + d[2] * eje[2]) / largo2
        res = None
        if -0.35 <= t <= 1.6:              # más allá de la barra: rellenos claros
            cerca = (i0[0] + eje[0] * t, i0[1] + eje[1] * t, i0[2] + eje[2] * t)
            dist2 = (r - cerca[0]) ** 2 + (g - cerca[1]) ** 2 + (bl - cerca[2]) ** 2
            if dist2 <= 30 * 30:           # está en la gama del tema anterior
                t = max(0.0, min(1.0, t))
                res = (round(ni.blue() + (nb.blue() - ni.blue()) * t),
                       round(ni.green() + (nb.green() - ni.green()) * t),
                       round(ni.red() + (nb.red() - ni.red()) * t))
        cache[clave] = res
        return res

    mv = memoryview(img.bits()).cast("B")
    for k in range(0, len(mv), 4):         # BGRA
        if mv[k + 3] == 0:
            continue
        res = nuevo_color(mv[k + 2], mv[k + 1], mv[k])
        if res is not None:
            mv[k], mv[k + 1], mv[k + 2] = res
    nuevo = QPixmap.fromImage(img)
    nuevo.setDevicePixelRatio(pm.devicePixelRatio())
    return QIcon(nuevo)


class GestorTema(QObject):
    """Tema de la interfaz y color de la hoja. Los widgets con colores
    propios se registran con `estilo()` (o se conectan a `cambiado`) y se
    actualizan solos al cambiar de tema."""
    cambiado = Signal()

    def __init__(self):
        super().__init__()
        self.t = dict(TEMAS["claro"])
        self.color_hoja = COLOR_HOJA
        self._estilos = []

    def estilo(self, widget, fn):
        """Aplica fn(tokens) como hoja de estilo del widget, ahora y en
        cada cambio de tema."""
        widget.setStyleSheet(fn(self.t))
        self._estilos.append((widget, fn))

    def aplicar(self, t, color_hoja=None):
        global COLOR_ICONO
        anterior = self.t
        hoja_cambia = color_hoja is not None and color_hoja != self.color_hoja
        if color_hoja is not None:
            self.color_hoja = color_hoja
        if t == anterior and not hoja_cambia:
            return
        self.t = dict(t)
        app = QApplication.instance()
        if t != anterior and app is not None:
            COLOR_ICONO = t["icono"]       # los iconos nuevos ya salen así
            app.setPalette(paleta_tema(t))
            app.setStyleSheet(estilo_app(t))
            self._recolorear_iconos(anterior, t)
        vivos = []
        for w, fn in self._estilos:
            try:
                w.setStyleSheet(fn(self.t))
                vivos.append((w, fn))
            except RuntimeError:           # el widget ya no existe
                pass
        self._estilos = vivos
        self.cambiado.emit()

    @staticmethod
    def _recolorear_iconos(de, a):
        app = QApplication.instance()
        acciones, botones = {}, []
        for w in app.allWidgets():
            for act in w.actions():
                acciones[id(act)] = act
            if isinstance(w, QToolButton) and w.defaultAction() is not None:
                acciones[id(w.defaultAction())] = w.defaultAction()
            elif isinstance(w, (QToolButton, QPushButton)) and not w.icon().isNull():
                botones.append(w)
            if w.isWindow() and w.testAttribute(Qt.WA_SetWindowIcon):
                w.setWindowIcon(recolorear_icono(w.windowIcon(), de, a))
        for top in app.topLevelWidgets():
            for act in top.findChildren(QAction):
                acciones[id(act)] = act
        for obj in list(acciones.values()) + botones:
            try:                           # puede haberse borrado ya
                if not obj.icon().isNull():
                    obj.setIcon(recolorear_icono(obj.icon(), de, a))
            except RuntimeError:
                pass


TEMA = GestorTema()
ESTILO = estilo_app(TEMAS["claro"])      # hoja de estilo inicial (tema Claro)


def paleta_clara():
    """Paleta clara fija para que la interfaz se vea igual en modo oscuro."""
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
    p.setColor(QPalette.PlaceholderText, QColor("#8A8A8A"))
    return p


# ---------------------------------------------------------------------------
# Iconos dibujados (nítidos en pantallas HiDPI, sin archivos externos)
# ---------------------------------------------------------------------------
def crear_icono(texto, *, negrita=False, cursiva=False, subrayado=False,
                tachado=False, pequeno=None, pos_pequeno="arriba",
                barra=None, fondo=None):
    escala = 3
    tam = 24
    pm = QPixmap(tam * escala, tam * escala)
    pm.setDevicePixelRatio(escala)
    pm.fill(Qt.transparent)

    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)

    alto_texto = 19 if barra is not None else tam

    if fondo is not None:  # bloque de resaltado detrás del texto
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(fondo))
        p.drawRoundedRect(QRectF(3, 4, 18, 13), 2, 2)

    f = QFont("Arial")
    f.setPixelSize(15 if pequeno is None else 14)
    f.setBold(negrita)
    f.setItalic(cursiva)
    f.setUnderline(subrayado)
    f.setStrikeOut(tachado)
    p.setFont(f)
    p.setPen(QColor(COLOR_ICONO))

    if pequeno is None:
        p.drawText(QRectF(0, 0, tam, alto_texto), Qt.AlignCenter, texto)
    else:
        p.drawText(QRectF(1, 0, 15, tam), Qt.AlignCenter, texto)
        fp = QFont("Arial")
        fp.setPixelSize(10)
        fp.setBold(True)
        p.setFont(fp)
        y = 2 if pos_pequeno == "arriba" else 12
        p.drawText(QRectF(14, y, 10, 10), Qt.AlignCenter, pequeno)

    if barra is not None:  # franja con el color actual
        c = QColor(barra)
        p.setPen(QPen(QColor("#9A9A9A"), 0.6) if not c.isValid() or c.lightness() > 190
                 else Qt.NoPen)
        p.setBrush(QBrush(c) if c.isValid() else Qt.NoBrush)
        p.drawRoundedRect(QRectF(3, 19.5, 18, 3.5), 1, 1)

    p.end()
    return QIcon(pm)


def icono_parrafo(tipo):
    """Iconos de líneas para alineación y sangría."""
    escala = 3
    tam = 24
    pm = QPixmap(tam * escala, tam * escala)
    pm.setDevicePixelRatio(escala)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor(COLOR_ICONO), 1.8)
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)

    ys = [5, 9.5, 14, 18.5]
    largos = [16, 11, 16, 11]
    x0, x1 = 4, 20
    if tipo in ("izquierda", "centro", "derecha", "justificado"):
        for y, l in zip(ys, largos):
            if tipo == "justificado":
                a, b = x0, x1
            elif tipo == "izquierda":
                a, b = x0, x0 + l
            elif tipo == "derecha":
                a, b = x1 - l, x1
            else:
                a = (x0 + x1 - l) / 2
                b = a + l
            p.drawLine(QPointF(a, y), QPointF(b, y))
    else:
        # Primera y última línea completas; las intermedias desplazadas
        p.drawLine(QPointF(x0, ys[0]), QPointF(x1, ys[0]))
        p.drawLine(QPointF(x0, ys[3]), QPointF(x1, ys[3]))
        p.drawLine(QPointF(11, ys[1]), QPointF(x1, ys[1]))
        p.drawLine(QPointF(11, ys[2]), QPointF(x1, ys[2]))
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(COLOR_ICONO))
        if tipo == "aumentar":   # flecha hacia la derecha
            tri = QPolygonF([QPointF(4, 8.5), QPointF(8.5, 11.75), QPointF(4, 15)])
        else:                    # flecha hacia la izquierda
            tri = QPolygonF([QPointF(8.5, 8.5), QPointF(4, 11.75), QPointF(8.5, 15)])
        p.drawPolygon(tri)
    p.end()
    return QIcon(pm)


def icono_pagina():
    """Hoja con la esquina doblada y tres líneas de texto, en un solo color."""
    escala = 3
    tam = 24
    pm = QPixmap(tam * escala, tam * escala)
    pm.setDevicePixelRatio(escala)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor(COLOR_ICONO), 1.6)
    pen.setJoinStyle(Qt.RoundJoin)
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)
    p.setBrush(Qt.NoBrush)

    x0, y0, x1, y1, d = 5, 3, 19, 21, 5
    p.drawPolygon(QPolygonF([QPointF(x0, y0), QPointF(x1 - d, y0),
                             QPointF(x1, y0 + d), QPointF(x1, y1),
                             QPointF(x0, y1)]))
    p.drawPolyline(QPolygonF([QPointF(x1 - d, y0), QPointF(x1 - d, y0 + d),
                              QPointF(x1, y0 + d)]))
    for y, fin in ((11.5, 15.5), (14.5, 15.5), (17.5, 12.5)):
        p.drawLine(QPointF(8.5, y), QPointF(fin, y))
    p.end()
    return QIcon(pm)


def icono_ajustes_estilo():
    """Letra «A» con tres renglones al lado: la galería de estilos."""
    escala = 3
    tam = 24
    pm = QPixmap(tam * escala, tam * escala)
    pm.setDevicePixelRatio(escala)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)
    f = QFont("Georgia")
    f.setPixelSize(16)
    f.setBold(True)
    p.setFont(f)
    p.setPen(QColor(COLOR_ICONO))
    p.drawText(QRectF(1, 2, 13, 18), Qt.AlignCenter, "A")
    pen = QPen(QColor(COLOR_ICONO), 1.6)
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)
    for y, x1 in ((7, 21), (12, 21), (17, 18)):
        p.drawLine(QPointF(15, y), QPointF(x1, y))
    p.end()
    return QIcon(pm)


def icono_lista(tipo):
    """Tres renglones con viñetas o con números a la izquierda."""
    escala = 3
    tam = 24
    pm = QPixmap(tam * escala, tam * escala)
    pm.setDevicePixelRatio(escala)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)
    color = QColor(COLOR_ICONO)
    ys = [6, 12, 18]
    pen = QPen(color, 1.8)
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)
    for y in ys:
        p.drawLine(QPointF(10, y), QPointF(20, y))
    if tipo == "vinetas":
        p.setPen(Qt.NoPen)
        p.setBrush(color)
        for y in ys:
            p.drawEllipse(QPointF(5, y), 1.7, 1.7)
    else:
        f = QFont("Arial")
        f.setPixelSize(7)
        f.setBold(True)
        p.setFont(f)
        p.setPen(color)
        for n, y in enumerate(ys, 1):
            p.drawText(QRectF(1, y - 4.5, 7, 9), Qt.AlignCenter, str(n))
    p.end()
    return QIcon(pm)


# ---------------------------------------------------------------------------
# Botón de color con menú de paleta
# ---------------------------------------------------------------------------
class BotonColor(QToolButton):
    colorElegido = Signal(QColor)  # QColor inválido = quitar color

    def __init__(self, glifo, color_inicial, paleta, texto_ninguno,
                 color_ninguno, resaltado=False, parent=None):
        super().__init__(parent)
        self.glifo = glifo
        self.resaltado = resaltado
        self.color_ninguno = QColor(color_ninguno) if color_ninguno else QColor()
        self.actual = QColor(color_inicial)

        self.setPopupMode(QToolButton.MenuButtonPopup)
        self.clicked.connect(lambda: self.colorElegido.emit(self.actual))

        menu = QMenu(self)
        accion_ninguno = menu.addAction(texto_ninguno)
        accion_ninguno.triggered.connect(lambda: self._elegir(self.color_ninguno))
        menu.addSeparator()

        cont = QWidget()
        lay = QVBoxLayout(cont)
        lay.setContentsMargins(8, 4, 8, 4)
        lay.setSpacing(6)
        titulo = QLabel("Colores")
        TEMA.estilo(titulo, lambda t: f"color:{t['texto_sec']}; font-size:11px;")
        lay.addWidget(titulo)

        grid = QGridLayout()
        grid.setSpacing(4)
        for i, hexa in enumerate(paleta):
            b = QToolButton()
            b.setFixedSize(20, 20)
            b.setCursor(Qt.PointingHandCursor)
            TEMA.estilo(b, lambda t, hexa=hexa: (
                f"QToolButton {{ background:{hexa}; border:1px solid {t['campo_borde']};"
                f" border-radius:3px; padding:0; }}"
                f"QToolButton:hover {{ border:2px solid {t['texto']}; }}"))
            b.clicked.connect(lambda _=False, c=hexa: self._elegir(QColor(c)))
            grid.addWidget(b, i // 10 if len(paleta) > 10 else i // 5,
                           i % 10 if len(paleta) > 10 else i % 5)
        lay.addLayout(grid)

        wa = QWidgetAction(menu)
        wa.setDefaultWidget(cont)
        menu.addAction(wa)
        menu.addSeparator()
        accion_mas = menu.addAction("Más colores…")
        accion_mas.triggered.connect(self._dialogo)

        self.menu_paleta = menu
        self.setMenu(menu)
        self._actualizar_icono()

    def _actualizar_icono(self):
        if self.resaltado:
            self.setIcon(crear_icono(self.glifo, barra=self.actual))
        else:
            self.setIcon(crear_icono(self.glifo, negrita=True, barra=self.actual))

    def _elegir(self, color):
        self.menu_paleta.hide()
        self.actual = QColor(color)
        self._actualizar_icono()
        self.colorElegido.emit(self.actual)

    def _dialogo(self):
        inicial = self.actual if self.actual.isValid() else QColor("#FFFFFF")
        c = QColorDialog.getColor(inicial, self, "Elegir color")
        if c.isValid():
            self._elegir(c)


def aplicar_letras(fmt, px):
    """Espaciado entre letras. Con 0 se usa el espaciado normal de la fuente
    (100 %), sin ningún ajuste: así se conservan el interletraje y las
    ligaduras originales. Con otro valor, px absolutos."""
    if px == 0:
        fmt.setFontLetterSpacingType(QFont.PercentageSpacing)
        fmt.setFontLetterSpacing(100)
    else:
        fmt.setFontLetterSpacingType(QFont.AbsoluteSpacing)
        fmt.setFontLetterSpacing(px)


# ---------------------------------------------------------------------------
# Listas
# ---------------------------------------------------------------------------
VINETAS = [QTextListFormat.ListDisc, QTextListFormat.ListCircle,
           QTextListFormat.ListSquare]
TIPOS_LISTA = [
    ("Viñeta  ●", QTextListFormat.ListDisc),
    ("Viñeta  ○", QTextListFormat.ListCircle),
    ("Viñeta  ■", QTextListFormat.ListSquare),
    None,
    ("Numérica  1. 2. 3.", QTextListFormat.ListDecimal),
    ("Letras  a. b. c.", QTextListFormat.ListLowerAlpha),
    ("Letras  A. B. C.", QTextListFormat.ListUpperAlpha),
    ("Romanos  i. ii. iii.", QTextListFormat.ListLowerRoman),
    ("Romanos  I. II. III.", QTextListFormat.ListUpperRoman),
]


def estilo_base(estilo):
    """Las viñetas cuentan como un mismo tipo (cambian con el nivel)."""
    return "vineta" if estilo in VINETAS else estilo


def vineta_nivel(estilo, desde, hasta):
    """En las listas con viñetas, cada nivel usa la viñeta siguiente
    (● → ○ → ■ …). Las numéricas conservan su tipo."""
    if estilo in VINETAS:
        return VINETAS[(VINETAS.index(estilo) + hasta - desde) % len(VINETAS)]
    return estilo


# ---------------------------------------------------------------------------
# Imágenes: cada imagen va sola en su propio párrafo (centrado por defecto).
# Se dibujan con QTextImageFormat + un recurso del documento; se mueven solo
# entre párrafos y se redimensionan con 8 puntos de arrastre.
# ---------------------------------------------------------------------------
MIME_IMAGEN = "application/x-editorlibro-imagen"
IMAGEN_MIN = 16                # px, lado mínimo al redimensionar
IMAGEN_MAX_RECURSO = 2400      # px, ancho máximo de la copia que se dibuja
ASA = 8                        # px de pantalla, lado de cada punto de arrastre
COLOR_MARCO = "#2B79D7"
ESPACIO_IMAGEN = 12            # px iniciales (luego se calculan, ver abajo)
# Separación visible entre la imagen y el texto, en renglones del estilo
# Body: igual arriba y abajo, medida de la línea base del texto anterior y
# hasta la altura de las mayúsculas del texto siguiente (como en imprenta)
ESPACIO_IMAGEN_RENGLONES = 0.0      # separación predeterminada (pegada)
PROP_SEPARACION = QTextFormat.UserProperty + 10   # separación de cada imagen
SEPARACIONES = [("Pegada", 0.0), ("Pequeña", 0.75), ("Normal", 1.0),
                ("Amplia", 1.5), ("Muy amplia", 2.0)]
# Formatos que Word acepta tal cual; los demás se guardan como PNG
TIPOS_IMAGEN_DOCX = {"png": "image/png", "jpeg": "image/jpeg",
                     "gif": "image/gif", "bmp": "image/bmp"}
# Puntos de arrastre: posición relativa (x, y) dentro del marco
ASAS = [(0, 0), (0.5, 0), (1, 0), (1, 0.5), (1, 1), (0.5, 1), (0, 1), (0, 0.5)]
CURSOR_ASA = {
    (0, 0): Qt.SizeFDiagCursor, (1, 1): Qt.SizeFDiagCursor,
    (1, 0): Qt.SizeBDiagCursor, (0, 1): Qt.SizeBDiagCursor,
    (0.5, 0): Qt.SizeVerCursor, (0.5, 1): Qt.SizeVerCursor,
    (0, 0.5): Qt.SizeHorCursor, (1, 0.5): Qt.SizeHorCursor,
}


def formato_bloque_imagen(alineacion=Qt.AlignHCenter):
    """Párrafo de una imagen: centrado, interlineado sencillo y sin sangrías."""
    bf = QTextBlockFormat()
    bf.setAlignment(alineacion)
    bf.setLineHeight(100, QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
    bf.setTopMargin(ESPACIO_IMAGEN)
    bf.setBottomMargin(ESPACIO_IMAGEN)
    return bf


def formato_imagen_en(doc, pos):
    """QTextImageFormat del carácter en `pos`, o None si no es una imagen."""
    if pos < 0 or pos >= doc.characterCount() - 1 or \
            doc.characterAt(pos) != CARACTER_OBJETO:
        return None
    c = QTextCursor(doc)
    c.setPosition(pos + 1)          # charFormat() = el del carácter anterior
    cf = c.charFormat()
    return cf.toImageFormat() if cf.isImageFormat() else None


def es_bloque_imagen(b):
    """Párrafo que solo contiene una imagen."""
    return b.isValid() and b.length() == 2 and \
        formato_imagen_en(b.document(), b.position()) is not None


def imagen_a_bytes(img, formato="PNG", calidad=-1):
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QIODevice.WriteOnly)
    img.save(buf, formato, calidad)
    buf.close()
    return bytes(ba)


# ---------------------------------------------------------------------------
# Editor
# ---------------------------------------------------------------------------
def tramos_de_documento(doc, inicio, fin, interno):
    """Párrafos entre `inicio` y `fin` como [{"tramos": [...]}]. Cada tramo
    es {"t": texto, "b", "i", "u", "s", "va"} con solo el formato puesto a
    mano: la negrita o cursiva que da el estilo del párrafo (un Title, un
    Subtitle) o un título HTML no cuentan, ni el subrayado de un enlace.
    Las imágenes del propio documento van como {"img", "w", "h"}."""
    parrafos = []
    b = doc.findBlock(inicio)
    while b.isValid() and b.position() <= fin:
        bf = b.blockFormat()
        estilo_b = estilo_i = False
        if interno and bf.hasProperty(PROP_ESTILO):
            e = GESTOR.resuelto(bf.stringProperty(PROP_ESTILO)
                                if bf.stringProperty(PROP_ESTILO) in GESTOR.datos else "Body")
            estilo_b, estilo_i = e["bold"], e["italic"]
        titulo = not interno and bf.headingLevel() > 0
        tramos = []
        it = b.begin()
        while not it.atEnd():
            fr = it.fragment()
            it += 1
            if not fr.isValid():
                continue
            a = max(inicio, fr.position())
            z = min(fin, fr.position() + fr.length())
            if a >= z:
                continue
            cf = fr.charFormat()
            texto = fr.text()[a - fr.position():z - fr.position()]
            if cf.isImageFormat():
                if interno:
                    f = cf.toImageFormat()
                    for _ in texto:
                        tramos.append({"img": f.name(), "w": f.width(), "h": f.height()})
                continue
            texto = texto.replace(CARACTER_OBJETO, "").replace("\u2028", "\n")
            if not texto:
                continue
            t = {"t": texto}
            if cf.fontWeight() >= QFont.DemiBold and not estilo_b and not titulo:
                t["b"] = 1
            if cf.fontItalic() and not estilo_i:
                t["i"] = 1
            if cf.fontUnderline() and not cf.isAnchor():
                t["u"] = 1
            if cf.fontStrikeOut():
                t["s"] = 1
            va = cf.verticalAlignment()
            if va == QTextCharFormat.AlignSubScript:
                t["va"] = "sub"
            elif va == QTextCharFormat.AlignSuperScript:
                t["va"] = "sup"
            if tramos and "img" not in tramos[-1] and \
                    {k: v for k, v in tramos[-1].items() if k != "t"} == \
                    {k: v for k, v in t.items() if k != "t"}:
                tramos[-1]["t"] += texto   # mismo formato: un solo tramo
            else:
                tramos.append(t)
        parrafos.append({"tramos": tramos})
        b = b.next()
    return parrafos


def formato_pegado(base, t):
    """Formato de Body con lo que el tramo trae puesto a mano."""
    f = QTextCharFormat(base)
    if t.get("b"):
        f.setFontWeight(QFont.Bold)
    if t.get("i"):
        f.setFontItalic(True)
    if t.get("u"):
        f.setFontUnderline(True)
    if t.get("s"):
        f.setFontStrikeOut(True)
    if t.get("va") == "sub":
        f.setVerticalAlignment(QTextCharFormat.AlignSubScript)
    elif t.get("va") == "sup":
        f.setVerticalAlignment(QTextCharFormat.AlignSuperScript)
    return f


class Editor(QTextEdit):
    # Se emite cuando una herramienta cambia el formato del párrafo sin mover
    # el cursor, para que todas las barras se actualicen.
    formatoParrafoCambiado = Signal()
    imagenCambiada = Signal()      # se seleccionó o deseleccionó una imagen

    def __init__(self):
        super().__init__()
        self.vista = None        # VistaSinPaginas que lo contiene
        self.document().setIndentWidth(28)    # px por nivel de lista
        self.aplicar_body()
        self._selecciones_externas = []     # resaltado de Buscar
        self.comentarios = GestorComentarios(self)
        # Imágenes: nombre del recurso → (bytes originales, formato, copia
        # que se dibuja). Los bytes se guardan tal cual en el .docx
        self.imagenes = {}
        self._contador_imagen = 0
        self._img = None             # QTextCursor sobre la imagen elegida
        self._accion = None          # arrastre en curso (mover / redimensionar)
        self._eligiendo = False
        self._forma_cursor = None
        self.viewport().setMouseTracking(True)
        self.document().contentsChanged.connect(self._validar_imagen)
        self.cursorPositionChanged.connect(self._cursor_movido_imagen)

    # -- selecciones extra (Buscar + comentarios) ---------------------------
    def setExtraSelections(self, selecciones):
        self._selecciones_externas = list(selecciones)
        self.refrescar_selecciones()

    def refrescar_selecciones(self):
        propias = self.comentarios.selecciones() \
            if getattr(self, "comentarios", None) is not None else []
        super().setExtraSelections(propias + self._selecciones_externas)

    # -- portapapeles -------------------------------------------------------
    def createMimeDataFromSelection(self):
        """Al copiar dentro del editor se guarda también el fragmento completo
        (formato e imágenes)."""
        md = super().createMimeDataFromSelection()
        c = self.textCursor()
        if not c.hasSelection():
            return md
        propio = QMimeData()
        for f in md.formats():
            propio.setData(f, md.data(f))
        propio.setData(MIME_LIBRO, QByteArray(
            c.selection().toHtml().encode("utf-8")))
        tramos = tramos_de_documento(self.document(), c.selectionStart(),
                                     c.selectionEnd(), interno=True)
        propio.setData(MIME_TRAMOS, QByteArray(json.dumps(tramos).encode("utf-8")))
        return propio

    def canInsertFromMimeData(self, source):
        return source.hasFormat(MIME_LIBRO) or source.hasFormat(MIME_IMAGEN) \
            or source.hasImage() or super().canInsertFromMimeData(source)

    def setFocus(self, *args):
        # Dentro de la vista, el foco real es de la vista
        if self.vista is not None:
            self.vista.setFocus()
        super().setFocus(*args)

    def char_format(self, nombre="Body"):
        return formato_caracter(nombre)

    def block_format(self, nombre="Body"):
        return formato_parrafo(nombre)

    def body_char_format(self):
        return self.char_format("Body")

    def body_block_format(self):
        return self.block_format("Body")

    def aplicar_body(self):
        font = QFont(BODY["font_family"], BODY["font_size"])
        self.document().setDefaultFont(font)
        self.setFont(font)
        self.setTextColor(QColor(BODY["font_color"]))

        cursor = self.textCursor()
        cursor.select(QTextCursor.Document)
        cursor.setBlockFormat(self.body_block_format())
        cursor.setCharFormat(self.body_char_format())
        cursor.clearSelection()
        self.setTextCursor(cursor)
        self.setCurrentCharFormat(self.body_char_format())

    def aplicar_estilo_body(self):
        self.aplicar_estilo("Body")

    def aplicar_estilo(self, nombre):
        """Aplica un estilo completo (carácter y párrafo).
        Con selección: al texto seleccionado y a los párrafos que toca.
        Sin selección: solo al puntero de texto, es decir, a lo nuevo que se
        escriba; el texto existente no cambia."""
        cursor = self.textCursor()
        self._pendiente = None
        if cursor.hasSelection():
            cursor.beginEditBlock()  # un solo paso de deshacer
            marcas = self.comentarios.marcas(cursor.selectionStart(),
                                             cursor.selectionEnd())
            imagenes = self._imagenes_en(cursor.selectionStart(),
                                         cursor.selectionEnd())
            cursor.setCharFormat(self.char_format(nombre))
            self._restaurar_imagenes(imagenes)     # no pierden su formato
            self.comentarios.restaurar(marcas)
            cursor.mergeBlockFormat(self.block_format(nombre))
            self._interlineado_imagenes(cursor.selectionStart(),
                                        cursor.selectionEnd())
            cursor.endEditBlock()
        elif cursor.block().length() <= 1:
            # Párrafo vacío: no hay texto que cambiar, así que el párrafo
            # entero toma el estilo (espacios, alineación, estilo siguiente)
            cursor.mergeBlockFormat(self.block_format(nombre))
        else:
            # Párrafo con texto: solo lo que se escriba; la cinta lo recuerda
            # mientras sigas en este párrafo
            self._pendiente = (nombre, cursor.blockNumber())
        # Con selección, Qt aplicaría esto otra vez a todo lo seleccionado
        # (y borraría las imágenes); el texto ya tiene el estilo
        if not cursor.hasSelection():
            self.setCurrentCharFormat(self.char_format(nombre))
        self.currentCharFormatChanged.emit(self.currentCharFormat())
        self.formatoParrafoCambiado.emit()
        self.setFocus()

    # -- listas -------------------------------------------------------------
    def aplicar_lista(self, estilo):
        """Convierte en lista los párrafos seleccionados (o el actual).
        Si ya son una lista del mismo tipo, la quita; si son de otro tipo,
        la cambia conservando los niveles."""
        cursor = self.textCursor()
        bloques = self.bloques_seleccionados()
        actual = cursor.currentList()
        cursor.beginEditBlock()
        if actual is not None and all(b.textList() is not None for b in bloques):
            listas = {b.textList() for b in bloques}
            mismo = all(estilo_base(l.format().style()) == estilo_base(estilo)
                        for l in listas)
            if mismo:
                for b in bloques:
                    self._quitar_de_lista(b)
            else:
                for l in listas:
                    f = l.format()
                    f.setStyle(vineta_nivel(estilo, 1, f.indent()))
                    l.setFormat(f)
        else:
            f = QTextListFormat()
            f.setStyle(estilo)
            f.setIndent(1)
            # Continúa la lista anterior si es del mismo tipo y nivel
            previo = bloques[0].previous()
            lista = previo.textList() if previo.isValid() else None
            if lista is not None and lista.format().indent() == 1 and \
                    estilo_base(lista.format().style()) == estilo_base(estilo):
                for b in bloques:
                    if b.textList() is not None:
                        b.textList().remove(b)
                    lista.add(b)
            else:
                c = QTextCursor(self.document())
                c.setPosition(bloques[0].position())
                c.setPosition(bloques[-1].position(), QTextCursor.KeepAnchor)
                c.createList(f)
        cursor.endEditBlock()
        self.formatoParrafoCambiado.emit()
        self.setFocus()

    def _quitar_de_lista(self, b):
        l = b.textList()
        if l is None:
            return
        l.remove(b)
        c = QTextCursor(b)
        fmt = QTextBlockFormat()
        fmt.setIndent(0)
        fmt.setObjectIndex(-1)
        c.mergeBlockFormat(fmt)

    def cambiar_nivel_lista(self, paso):
        """Sube o baja de nivel los elementos seleccionados. Por debajo del
        nivel 1 el elemento sale de la lista."""
        cursor = self.textCursor()
        cursor.beginEditBlock()
        for b in self.bloques_seleccionados():
            l = b.textList()
            if l is None:
                continue
            f = l.format()
            nivel = f.indent() + paso
            if nivel < 1:
                self._quitar_de_lista(b)
                continue
            nivel = min(nivel, 6)
            # Busca hacia atrás una lista de ese nivel y tipo para unirse
            destino, p = None, b.previous()
            while p.isValid() and p.textList() is not None:
                pf = p.textList().format()
                # Al subir de nivel se vuelve a la lista de arriba sea cual
                # sea su tipo; al bajar, solo se une a una del mismo tipo
                if pf.indent() == nivel and (paso < 0 or
                        estilo_base(pf.style()) == estilo_base(f.style())):
                    destino = p.textList()
                    break
                if pf.indent() < nivel:
                    break
                p = p.previous()
            l.remove(b)
            sin_sangria = QTextBlockFormat()
            sin_sangria.setIndent(0)       # la sangría la da la lista
            QTextCursor(b).mergeBlockFormat(sin_sangria)
            if destino is not None:
                destino.add(b)
            else:
                nf = QTextListFormat(f)
                nf.setIndent(nivel)
                nf.setStyle(vineta_nivel(f.style(), f.indent(), nivel))
                QTextCursor(b).createList(nf)
        cursor.endEditBlock()
        self.formatoParrafoCambiado.emit()
        self.setFocus()

    def contextMenuEvent(self, e):
        """Menú de clic derecho solo con Cortar, Copiar y Pegar (las mismas
        acciones de Qt, con sus atajos y su estado activado/desactivado).
        Sobre una imagen elegida: acciones de la imagen."""
        if self.imagen_seleccionada() is not None:
            menu = QMenu(self._padre_menu())
            menu.addAction("Cortar").triggered.connect(self.cut)
            menu.addAction("Copiar").triggered.connect(self.copy)
            menu.addSeparator()
            self.menu_separacion(menu.addMenu("Separación arriba y abajo"))
            menu.addSeparator()
            menu.addAction("Tamaño original").triggered.connect(
                lambda: self.redimensionar_imagen(original=True))
            menu.addAction("Ajustar al ancho del texto").triggered.connect(
                lambda: self.redimensionar_imagen(ancho_texto=True))
            menu.addSeparator()
            menu.addAction("Eliminar imagen").triggered.connect(self.eliminar_imagen)
            menu.exec(self._a_global(QPointF(e.pos())))
            menu.deleteLater()
            return
        menu = self.createStandardContextMenu(e.pos())
        dejar = {"edit-cut", "edit-copy", "edit-paste"}
        for accion in menu.actions():
            if accion.objectName() not in dejar:
                menu.removeAction(accion)
        menu.exec(e.globalPos())
        menu.deleteLater()

    def estilo_actual(self):
        """Id del estilo del párrafo bajo el cursor (o None)."""
        cursor = self.textCursor()
        pend = getattr(self, "_pendiente", None)
        if pend and not cursor.hasSelection() and pend[0] in GESTOR.datos:
            if cursor.blockNumber() == pend[1]:
                return pend[0]
            self._pendiente = None
        v = cursor.blockFormat().property(PROP_ESTILO)
        return v if isinstance(v, str) and v in GESTOR.datos else None

    def keyPressEvent(self, e):
        if self._teclas_imagen(e):
            return
        cursor = self.textCursor()
        lista = cursor.currentList()
        if lista is not None:
            enter = e.key() in (Qt.Key_Return, Qt.Key_Enter) and \
                not e.modifiers() & Qt.ShiftModifier
            # Enter en un elemento vacío: sube un nivel o termina la lista
            if enter and cursor.block().length() <= 1:
                self.cambiar_nivel_lista(-1)
                return
            if enter:
                super().keyPressEvent(e)     # nuevo elemento de la lista
                return
            # Tab / Mayús+Tab al inicio del elemento (o con selección): nivel
            if e.key() == Qt.Key_Tab and (cursor.atBlockStart() or cursor.hasSelection()):
                self.cambiar_nivel_lista(+1)
                return
            if e.key() == Qt.Key_Backtab:
                self.cambiar_nivel_lista(-1)
                return
            # Retroceso al inicio del elemento: lo saca de la lista
            if e.key() == Qt.Key_Backspace and cursor.atBlockStart() \
                    and not cursor.hasSelection():
                self.cambiar_nivel_lista(-1)
                return
        # Enter al final de un párrafo: el nuevo toma el "estilo siguiente"
        if e.key() in (Qt.Key_Return, Qt.Key_Enter) and \
                not e.modifiers() & Qt.ShiftModifier:
            actual = self.estilo_actual()
            al_final = self.textCursor().atBlockEnd()
            super().keyPressEvent(e)
            if actual and al_final:
                sig = GESTOR.datos[actual]["siguiente"]
                if sig in GESTOR.datos and sig != actual:
                    cursor = self.textCursor()
                    cursor.setBlockFormat(self.block_format(sig))
                    self.setCurrentCharFormat(self.char_format(sig))
                    self.formatoParrafoCambiado.emit()
            return
        super().keyPressEvent(e)

    def actualizar_estilos(self):
        """Tras cambiar la galería: cada párrafo vuelve a tomar los valores
        de su estilo (un solo paso de deshacer)."""
        doc = self.document()
        cursor = QTextCursor(doc)
        cursor.beginEditBlock()
        b = doc.begin()
        while b.isValid():
            v = b.blockFormat().property(PROP_ESTILO)
            if isinstance(v, str):
                id_ = v if v in GESTOR.datos else "Body"
                cursor.setPosition(b.position())
                cursor.mergeBlockFormat(self.block_format(id_))
                cursor.setPosition(b.position() + b.length() - 1,
                                   QTextCursor.KeepAnchor)
                cursor.mergeCharFormat(self.char_format(id_))
            b = b.next()
        self._interlineado_imagenes()
        cursor.endEditBlock()
        actual = self.estilo_actual()
        if actual and not self.textCursor().hasSelection():
            self.setCurrentCharFormat(self.char_format(actual))
        self.formatoParrafoCambiado.emit()

    def insertFromMimeData(self, source):
        # Imagen copiada en este editor: va en su propio párrafo
        if source.hasFormat(MIME_IMAGEN):
            try:
                d = json.loads(bytes(source.data(MIME_IMAGEN)).decode("utf-8"))
                nombre = d["nombre"]
            except (ValueError, KeyError, TypeError):
                nombre = None
            if nombre and self.document().resource(
                    QTextDocument.ImageResource, QUrl(nombre)) is not None:
                self._insertar_imagen_nueva(self._formato_imagen(
                    nombre, d.get("ancho", 0), d.get("alto", 0)))
                return
        # Imagen de otra aplicación (captura, navegador…)
        if source.hasImage() and not source.hasFormat(MIME_LIBRO) \
                and not source.hasText():
            img = QImage(source.imageData())
            if not img.isNull():
                self.insertar_imagen_datos(imagen_a_bytes(img))
                return
        # Nunca se pega texto dentro del párrafo de una imagen
        c = self.textCursor()
        if not c.hasSelection() and es_bloque_imagen(c.block()):
            self._parrafo_junto_a_imagen(c.block(),
                                         antes=c.position() == c.block().position())
        # Arrastrar texto dentro del documento lo mueve tal cual (un título
        # sigue siendo título)
        if getattr(self, "_soltando", False) and source.hasFormat(MIME_LIBRO):
            html = bytes(source.data(MIME_LIBRO)).decode("utf-8", "replace")
            self.textCursor().insertFragment(QTextDocumentFragment.fromHtml(html))
            return
        # Al pegar, el texto adopta el estilo Body y conserva solo negrita,
        # cursiva, subrayado, tachado, subíndice y superíndice
        tramos = None
        if source.hasFormat(MIME_TRAMOS):
            try:
                tramos = json.loads(bytes(source.data(MIME_TRAMOS)).decode("utf-8"))
            except ValueError:
                tramos = None
        if tramos is None and source.hasHtml():
            tmp = QTextDocument()
            QTextCursor(tmp).insertFragment(QTextDocumentFragment.fromHtml(source.html()))
            tramos = tramos_de_documento(tmp, 0, tmp.characterCount() - 1, interno=False)
            if not any(t for par in tramos for t in par["tramos"]) and source.hasText():
                tramos = None              # HTML vacío: mejor el texto
        if tramos is None and source.hasText():
            tramos = [{"tramos": [{"t": linea}] if linea else []}
                      for linea in source.text().replace("\r\n", "\n")
                      .replace("\r", "\n").split("\n")]
        if tramos:
            self.pegar_tramos(tramos)

    def combinar_formato(self, fmt):
        """Aplica el formato a la selección o al texto que se escriba después."""
        cursor = self.textCursor()
        if cursor.hasSelection():
            cursor.mergeCharFormat(fmt)
        self.mergeCurrentCharFormat(fmt)

    def combinar_parrafo(self, fmt):
        """Aplica un formato de párrafo a todos los párrafos seleccionados."""
        c = self.textCursor()
        c.beginEditBlock()               # un solo paso de deshacer
        c.mergeBlockFormat(fmt)
        self._interlineado_imagenes(c.selectionStart(), c.selectionEnd())
        c.endEditBlock()
        self.formatoParrafoCambiado.emit()

    def _interlineado_imagenes(self, *_):
        """Ajusta todos los párrafos de imagen del documento:
        - interlineado sencillo (con uno proporcional, Qt multiplica la
          altura de la imagen y deja un hueco debajo);
        - el mismo espacio óptico arriba y abajo (ver _margenes_imagen).
        Se llama dentro del paso de deshacer de cada operación."""
        doc = self.document()
        sencillo = QTextBlockFormat.LineHeightTypes.ProportionalHeight.value
        c = QTextCursor(doc)
        b = doc.begin()
        while b.isValid():
            if es_bloque_imagen(b):
                bf = b.blockFormat()
                arriba, abajo = self._margenes_imagen(b)
                if bf.lineHeightType() != sencillo or bf.lineHeight() != 100 \
                        or bf.textIndent() != 0 \
                        or abs(bf.topMargin() - arriba) > 0.5 \
                        or abs(bf.bottomMargin() - abajo) > 0.5:
                    f = QTextBlockFormat()
                    f.setLineHeight(100, sencillo)
                    f.setTextIndent(0)   # la 1.ª línea no descentra la imagen
                    f.setTopMargin(arriba)
                    f.setBottomMargin(abajo)
                    c.setPosition(b.position())
                    c.mergeBlockFormat(f)
            b = b.next()

    # -- espacio alrededor de las imágenes ----------------------------------
    def _metricas_linea(self, b, ultima):
        """(alto natural, ascendente) de la primera o última línea del
        párrafo, según su fuente (la maquetación puede estar sin
        actualizar mientras dura un paso de deshacer)."""
        c = QTextCursor(b)
        if ultima:
            c.movePosition(QTextCursor.EndOfBlock)
        elif b.length() > 1:
            c.movePosition(QTextCursor.NextCharacter)
        fm = QFontMetricsF(c.charFormat().font(), self.viewport())
        return fm.lineSpacing(), fm.ascent()

    def _altura_mayusculas(self, b):
        """Altura de las mayúsculas del primer carácter del párrafo."""
        c = QTextCursor(b)
        if b.length() > 1:
            c.movePosition(QTextCursor.NextCharacter)
        return QFontMetricsF(c.charFormat().font(), self.viewport()).capHeight()

    @staticmethod
    def _alto_renglon(bf, natural):
        """Altura que ocupa un renglón según el interlineado del párrafo."""
        tipos = QTextBlockFormat.LineHeightTypes
        t, v = bf.lineHeightType(), bf.lineHeight()
        if t == tipos.ProportionalHeight.value:
            return natural * v / 100
        if t == tipos.FixedHeight.value:
            return v
        if t == tipos.MinimumHeight.value:
            return max(natural, v)
        if t == tipos.LineDistanceHeight.value:
            return natural + v
        return natural

    def renglon_body(self):
        """Alto de un renglón del estilo Body (px), unidad de la separación."""
        e = GESTOR.resuelto("Body")
        f = QFont(e["font_family"])
        f.setPointSizeF(e["font_size"])
        fm = QFontMetricsF(f, self.viewport())
        return fm.height() * e["line_spacing"] / 100

    def separacion_imagen(self, b):
        """Separación elegida para esta imagen, en renglones de Body."""
        v = b.blockFormat().property(PROP_SEPARACION) if b.isValid() else None
        return float(v) if isinstance(v, (int, float)) else ESPACIO_IMAGEN_RENGLONES

    def espacio_imagen(self, b=None):
        """Separación deseada (px) entre la imagen y el texto."""
        renglones = self.separacion_imagen(b) if b is not None \
            else ESPACIO_IMAGEN_RENGLONES
        return renglones * self.renglon_body()

    def _vecinos_imagen(self, b):
        """(espacio fijo bajo la línea base del texto de arriba, margen que
        ya impone ese párrafo, espacio fijo sobre las mayúsculas del texto
        de abajo, margen que impone ese párrafo). None si no hay vecino."""
        previo, siguiente = b.previous(), b.next()
        arriba = abajo = None
        if previo.isValid():
            if es_bloque_imagen(previo):
                bajo_base = 0.0          # imagen sobre imagen: borde a borde
            else:
                alto, asc = self._metricas_linea(previo, ultima=True)
                bajo_base = self._alto_renglon(previo.blockFormat(), alto) - asc
            arriba = (bajo_base, previo.blockFormat().bottomMargin())
        if siguiente.isValid():
            if es_bloque_imagen(siguiente):
                sobre_mayus = 0.0
            else:
                _, asc = self._metricas_linea(siguiente, ultima=False)
                sobre_mayus = asc - self._altura_mayusculas(siguiente)
            abajo = (sobre_mayus, siguiente.blockFormat().topMargin())
        return arriba, abajo

    def espacio_visible(self, b):
        """Separación (px) que de verdad queda arriba y abajo: la elegida,
        o la mínima que permiten los párrafos vecinos si es mayor. Qt no
        admite márgenes negativos y entre párrafos usa el mayor margen (no
        la suma), así que el espacio propio del texto vecino (su «Después»
        o «Antes» y el aire de su interlineado) es el mínimo posible."""
        g = self.espacio_imagen(b)
        arriba, abajo = self._vecinos_imagen(b)
        minimos = [sum(x) for x in (arriba, abajo) if x is not None]
        return max([g] + minimos)

    def _margenes_imagen(self, b):
        """Márgenes del párrafo de imagen para que el espacio *visible* sea
        igual arriba y abajo: de la línea base del texto anterior al borde
        superior de la imagen, y del borde inferior a la altura de las
        mayúsculas del texto siguiente. Qt pone el interlineado extra debajo
        de cada renglón y, entre dos párrafos, usa el mayor de los márgenes
        (no los suma); ambas cosas se compensan aquí."""
        g = self.espacio_visible(b)      # simétrico aunque se pida menos
        arriba, abajo = self._vecinos_imagen(b)
        m_arriba = m_abajo = 0.0
        if arriba is not None and g - arriba[0] > arriba[1]:
            m_arriba = g - arriba[0]
        if abajo is not None and g - abajo[0] > abajo[1]:
            m_abajo = g - abajo[0]
        return round(max(0.0, m_arriba), 1), round(max(0.0, m_abajo), 1)

    def poner_separacion(self, renglones):
        """Separación de la imagen elegida (un solo paso de deshacer)."""
        pos = self.imagen_seleccionada()
        if pos is None:
            return
        b = self.document().findBlock(pos)
        if abs(self.separacion_imagen(b) - renglones) < 1e-6 and \
                b.blockFormat().hasProperty(PROP_SEPARACION):
            return
        c = QTextCursor(self.document())
        c.beginEditBlock()
        f = QTextBlockFormat()
        f.setProperty(PROP_SEPARACION, float(renglones))
        c.setPosition(b.position())
        c.mergeBlockFormat(f)
        self._interlineado_imagenes()
        c.endEditBlock()
        self.viewport().update()

    def _a_global(self, punto):
        """Punto del viewport a coordenadas de pantalla (con el zoom)."""
        v = self.vista
        if v is None:
            return self.viewport().mapToGlobal(punto.toPoint())
        return v.viewport().mapToGlobal(v.mapFromScene(v.proxy.mapToScene(punto)))

    def _padre_menu(self):
        # Un menú hijo del editor se dibujaría dentro de la hoja (con zoom)
        return self.vista.window() if self.vista is not None else self

    def menu_separacion(self, menu):
        """Opciones de separación arriba y abajo (simétrica)."""
        pos = self.imagen_seleccionada()
        b = self.document().findBlock(pos)
        actual = self.separacion_imagen(b)
        titulo = QLabel("Separación arriba y abajo")
        titulo.setStyleSheet(f"color:{TEMA.t['texto_sec']}; font-size:11px; font-weight:bold;"
                             " padding:4px 12px 2px 12px;")
        wa = QWidgetAction(menu)
        wa.setDefaultWidget(titulo)
        menu.addAction(wa)
        grupo = QActionGroup(menu)
        for nombre, valor in SEPARACIONES:
            a = menu.addAction(nombre)
            a.setCheckable(True)
            a.setChecked(abs(actual - valor) < 1e-6)
            grupo.addAction(a)
            a.triggered.connect(lambda _=False, v=valor: self.poner_separacion(v))
        menu.addSeparator()
        cont = QWidget()
        fila = QHBoxLayout(cont)
        fila.setContentsMargins(12, 2, 12, 2)
        fila.addWidget(QLabel("Personalizada:"))
        spin = QDoubleSpinBox()
        spin.setLocale(QLocale.c())
        spin.setRange(0, 4)
        spin.setSingleStep(0.25)
        spin.setDecimals(2)
        spin.setSuffix(" renglones")
        spin.setKeyboardTracking(False)
        spin.setValue(actual)
        fila.addWidget(spin)
        nota = QLabel()
        nota.setStyleSheet(f"color:{TEMA.t['texto_suave']}; font-size:11px;"
                           " padding:0 12px 4px 12px;")

        def mostrar_px():
            bb = self.document().findBlock(self.imagen_seleccionada() or 0)
            pedido = self.espacio_imagen(bb)
            real = self.espacio_visible(bb)
            texto = f"{real:.0f} px arriba y abajo"
            if real > pedido + 0.5:
                texto += " (lo más cerca que permite\nel espacio del párrafo vecino)"
            nota.setText(texto)

        def cambiar(v):
            self.poner_separacion(v)
            for a, (_, valor) in zip(grupo.actions(), SEPARACIONES):
                a.setChecked(abs(v - valor) < 1e-6)
            mostrar_px()

        spin.valueChanged.connect(cambiar)
        # Exclusivas, pero un valor personalizado puede no marcar ninguna
        grupo.setExclusionPolicy(QActionGroup.ExclusionPolicy.ExclusiveOptional)
        wa = QWidgetAction(menu)
        wa.setDefaultWidget(cont)
        menu.addAction(wa)
        wa2 = QWidgetAction(menu)
        wa2.setDefaultWidget(nota)
        menu.addAction(wa2)
        mostrar_px()
        return spin

    def _imagenes_en(self, inicio, fin):
        """[(posición, formato)] de las imágenes entre inicio y fin."""
        doc = self.document()
        res, pos = [], inicio
        while True:
            hallado = doc.find(CARACTER_OBJETO, pos)
            if hallado.isNull() or hallado.selectionStart() >= fin:
                return res
            pos = hallado.selectionStart()
            fmt = formato_imagen_en(doc, pos)
            if fmt is not None:
                res.append((pos, fmt))
            pos += 1

    def _restaurar_imagenes(self, imagenes):
        """Devuelve su formato a imágenes que un setCharFormat dejó sin él."""
        c = QTextCursor(self.document())
        for pos, fmt in imagenes:
            c.setPosition(pos)
            c.setPosition(pos + 1, QTextCursor.KeepAnchor)
            c.setCharFormat(fmt)

    def bloques_seleccionados(self):
        """Todos los párrafos que toca la selección (o el párrafo actual)."""
        cursor = self.textCursor()
        doc = self.document()
        fin = doc.findBlock(cursor.selectionEnd())
        b = doc.findBlock(cursor.selectionStart())
        bloques = []
        while b.isValid():
            bloques.append(b)
            if b == fin:
                break
            b = b.next()
        return bloques

    # =======================================================================
    # Imágenes
    # =======================================================================
    def _zoom(self):
        return self.vista.zoom if self.vista is not None else 1.0

    def _desplazamiento(self):
        return QPointF(self.horizontalScrollBar().value(),
                       self.verticalScrollBar().value())

    def ancho_texto(self, bloque=None):
        """Ancho disponible para el texto (y para una imagen) en px."""
        doc = self.document()
        ff = doc.rootFrame().frameFormat()
        ancho = doc.pageSize().width() - ff.leftMargin() - ff.rightMargin()
        if bloque is not None and bloque.isValid():
            bf = bloque.blockFormat()
            ancho -= bf.leftMargin() + bf.rightMargin()
        return max(IMAGEN_MIN, ancho)

    def _alto_texto(self):
        if self.vista is None:
            return 10000
        c = self.vista.config
        return max(IMAGEN_MIN, self.vista.cm_a_px(c["alto"] - c["superior"] - c["inferior"]))

    # -- registro de imágenes -----------------------------------------------
    def registrar_imagen(self, datos, recurso=True):
        """Guarda los bytes de una imagen y crea el recurso que dibuja el
        documento. Devuelve (nombre, tamaño natural) o None si no es imagen."""
        buf = QBuffer()
        buf.setData(QByteArray(datos))
        buf.open(QIODevice.ReadOnly)
        lector = QImageReader(buf)
        lector.setAutoTransform(True)        # respeta la rotación EXIF
        formato = bytes(lector.format()).decode("ascii", "ignore").lower()
        girar = lector.transformation() != QImageIOHandler.TransformationNone
        img = lector.read()
        buf.close()
        if img.isNull():
            return None
        formato = "jpeg" if formato == "jpg" else formato
        if girar or formato not in TIPOS_IMAGEN_DOCX:
            # Se guarda ya girada y en un formato que Word abre
            if formato == "jpeg" and not img.hasAlphaChannel():
                datos = imagen_a_bytes(img, "JPEG", 92)
            else:
                formato, datos = "png", imagen_a_bytes(img)
        vista = img
        if img.width() > IMAGEN_MAX_RECURSO:     # dibujar es más ligero
            vista = img.scaledToWidth(IMAGEN_MAX_RECURSO, Qt.SmoothTransformation)
        self._contador_imagen += 1
        nombre = f"imagen://{self._contador_imagen}"
        self.imagenes[nombre] = (datos, formato, vista)
        if recurso:
            self.document().addResource(QTextDocument.ImageResource,
                                        QUrl(nombre), vista)
        return nombre, img.size()

    def recargar_recursos(self):
        """Vuelve a crear los recursos (QTextDocument.clear() los borra)."""
        for nombre, (_, _, vista) in self.imagenes.items():
            self.document().addResource(QTextDocument.ImageResource,
                                        QUrl(nombre), vista)

    def _formato_imagen(self, nombre, ancho, alto):
        f = QTextImageFormat()
        f.setName(nombre)
        if ancho <= 0 or alto <= 0:
            img = self.document().resource(QTextDocument.ImageResource, QUrl(nombre))
            tam = img.size() if img is not None else QSize(100, 100)
            ancho, alto = tam.width(), tam.height()
        f.setWidth(ancho)
        f.setHeight(alto)
        return f

    def _tamano_imagen(self, fmt):
        w, h = fmt.width(), fmt.height()
        if w <= 0 or h <= 0:
            img = self.document().resource(QTextDocument.ImageResource, QUrl(fmt.name()))
            if img is not None:
                tam = img.size()
                if w <= 0 and h > 0 and tam.height():
                    w = h * tam.width() / tam.height()
                elif h <= 0 and w > 0 and tam.width():
                    h = w * tam.height() / tam.width()
                elif w <= 0 and h <= 0:
                    w, h = tam.width(), tam.height()
        return max(1.0, w), max(1.0, h)

    # -- insertar -----------------------------------------------------------
    def insertar_imagen_archivo(self, ruta):
        with open(ruta, "rb") as f:
            return self.insertar_imagen_datos(f.read())

    def insertar_imagen_datos(self, datos):
        """Inserta la imagen centrada en su propio párrafo, a su tamaño
        natural salvo que no quepa en el ancho del texto."""
        r = self.registrar_imagen(datos)
        if r is None:
            return False
        nombre, tam = r
        w, h = tam.width(), tam.height()
        maximo = self.ancho_texto()
        if w > maximo:
            w, h = maximo, h * maximo / w
        self._insertar_imagen_nueva(self._formato_imagen(nombre, round(w), round(h)))
        return True

    def _insertar_antes(self, c, destino, fmt, bf):
        """Crea un párrafo de imagen antes del bloque `destino` (o al final
        si es None) sin partir ningún texto. Devuelve la posición de la
        imagen."""
        doc = self.document()
        cf = self.body_char_format()
        if destino is None or not destino.isValid():
            ultimo = doc.lastBlock()
            c.setPosition(ultimo.position() + ultimo.length() - 1)
            c.insertBlock(bf, cf)
        elif destino.previous().isValid():
            previo = destino.previous()
            c.setPosition(previo.position() + previo.length() - 1)
            c.insertBlock(bf, cf)
        else:
            # Antes del primer párrafo: el texto pasa a un párrafo nuevo
            # con su mismo formato y el primero queda para la imagen
            bfd, cfd = destino.blockFormat(), destino.charFormat()
            c.setPosition(destino.position())
            c.insertBlock(bfd, cfd)
            c.setPosition(0)
            c.setBlockFormat(bf)
            c.setBlockCharFormat(cf)
        pos = c.position()
        c.insertImage(fmt)
        return pos

    def _insertar_imagen_nueva(self, fmt):
        """Párrafo vacío: la imagen lo ocupa. Cursor al inicio de un
        párrafo: va antes. Si no: va después del párrafo (nunca en medio)."""
        doc = self.document()
        actual = self.textCursor()
        self.deseleccionar()
        b = doc.findBlock(actual.selectionEnd())
        c = QTextCursor(doc)
        c.beginEditBlock()
        if b.length() <= 1 and b.textList() is None:
            c.setPosition(b.position())
            c.setBlockFormat(formato_bloque_imagen())
            c.setBlockCharFormat(self.body_char_format())
            pos = c.position()
            c.insertImage(fmt)
        elif not actual.hasSelection() and actual.atBlockStart():
            pos = self._insertar_antes(c, b, fmt, formato_bloque_imagen())
        else:
            pos = self._insertar_antes(c, b.next(), fmt, formato_bloque_imagen())
        self._interlineado_imagenes()
        c.endEditBlock()
        self.seleccionar_imagen(pos)
        self.setFocus()

    def _quitar_imagen(self, c, pos):
        """Borra la imagen; si estaba sola en su párrafo, borra el párrafo
        sin tocar el formato de los de alrededor."""
        doc = self.document()
        b = doc.findBlock(pos)
        if es_bloque_imagen(b):
            if b.next().isValid():
                c.setPosition(b.position())
                c.setPosition(b.next().position(), QTextCursor.KeepAnchor)
            elif b.previous().isValid():
                previo = b.previous()
                c.setPosition(previo.position() + previo.length() - 1)
                c.setPosition(b.position() + b.length() - 1, QTextCursor.KeepAnchor)
            else:                        # era lo único del documento
                c.setPosition(pos)
                c.setPosition(pos + 1, QTextCursor.KeepAnchor)
                c.removeSelectedText()
                c.setBlockFormat(self.body_block_format())
                c.setBlockCharFormat(self.body_char_format())
                return
        else:
            c.setPosition(pos)
            c.setPosition(pos + 1, QTextCursor.KeepAnchor)
        c.removeSelectedText()

    def eliminar_imagen(self):
        pos = self.imagen_seleccionada()
        if pos is None:
            return
        self.deseleccionar()
        c = QTextCursor(self.document())
        c.beginEditBlock()
        self._quitar_imagen(c, pos)
        self._interlineado_imagenes()
        c.endEditBlock()
        c.clearSelection()
        self.setTextCursor(c)
        self.formatoParrafoCambiado.emit()

    def mover_imagen(self, pos, destino):
        """Lleva la imagen a un párrafo propio antes del bloque número
        `destino` (== blockCount: al final). Un solo paso de deshacer."""
        doc = self.document()
        fmt = formato_imagen_en(doc, pos)
        if fmt is None:
            return
        origen = doc.findBlock(pos)
        bf = QTextBlockFormat(origen.blockFormat()) if es_bloque_imagen(origen) \
            else formato_bloque_imagen()
        bloque = doc.findBlockByNumber(destino) if destino < doc.blockCount() else None
        viejo = QTextCursor(doc)
        viejo.setPosition(pos)          # sigue a la imagen original
        c = QTextCursor(doc)
        c.beginEditBlock()
        nueva = QTextCursor(doc)
        nueva.setPosition(self._insertar_antes(c, bloque, fmt, bf))
        self._quitar_imagen(c, viejo.position())
        self._interlineado_imagenes()
        c.endEditBlock()
        self.seleccionar_imagen(nueva.position())
        self.formatoParrafoCambiado.emit()

    def redimensionar_imagen(self, ancho=None, alto=None, original=False,
                             ancho_texto=False):
        pos = self.imagen_seleccionada()
        if pos is None:
            return
        doc = self.document()
        fmt = formato_imagen_en(doc, pos)
        w, h = self._tamano_imagen(fmt)
        if original or ancho_texto:
            datos = self.imagenes.get(fmt.name())
            if datos is not None:
                img = QImage.fromData(datos[0])
                nat = img.size() if not img.isNull() else datos[2].size()
            else:
                res = doc.resource(QTextDocument.ImageResource, QUrl(fmt.name()))
                nat = res.size() if res is not None else QSize(round(w), round(h))
            ratio = nat.height() / nat.width() if nat.width() else h / w
            if original:
                ancho = min(nat.width(), self.ancho_texto(doc.findBlock(pos)))
            else:
                ancho = self.ancho_texto(doc.findBlock(pos))
            alto = ancho * ratio
        if ancho is None or alto is None:
            return
        ancho, alto = max(IMAGEN_MIN, round(ancho)), max(IMAGEN_MIN, round(alto))
        if (ancho, alto) == (round(w), round(h)):
            return
        nuevo = QTextImageFormat(fmt)
        nuevo.setWidth(ancho)
        nuevo.setHeight(alto)
        c = QTextCursor(doc)
        c.setPosition(pos)
        c.setPosition(pos + 1, QTextCursor.KeepAnchor)
        c.setCharFormat(nuevo)
        self.seleccionar_imagen(pos)

    # -- selección ----------------------------------------------------------
    def imagen_seleccionada(self):
        """Posición de la imagen elegida (o None)."""
        if self._img is None:
            return None
        pos = self._img.position()
        if formato_imagen_en(self.document(), pos) is None:
            self.deseleccionar()
            return None
        return pos

    def seleccionar_imagen(self, pos):
        doc = self.document()
        self._img = QTextCursor(doc)
        self._img.setPosition(pos)
        c = QTextCursor(doc)
        c.setPosition(pos)
        self._eligiendo = True
        self.setTextCursor(c)            # la cinta y los botones la siguen
        self._eligiendo = False
        self.setCursorWidth(0)           # sin cursor parpadeando
        self.viewport().update()
        self.imagenCambiada.emit()

    def deseleccionar(self):
        if self._img is None:
            return
        self._img = None
        self._accion = None
        self.setCursorWidth(1)
        self.viewport().update()
        self.imagenCambiada.emit()

    def _validar_imagen(self):
        if self._img is not None:
            self.imagen_seleccionada()
            self.viewport().update()

    def _cursor_movido_imagen(self):
        if self._eligiendo or self._img is None:
            return
        c = self.textCursor()
        if c.hasSelection() or c.position() != self._img.position():
            self.deseleccionar()

    def copy(self):
        if self.imagen_seleccionada() is not None:
            self._copiar_imagen()
        else:
            super().copy()

    def cut(self):
        if self.imagen_seleccionada() is not None:
            self._copiar_imagen()
            self.eliminar_imagen()
        else:
            super().cut()

    def _copiar_imagen(self):
        pos = self.imagen_seleccionada()
        fmt = formato_imagen_en(self.document(), pos)
        w, h = self._tamano_imagen(fmt)
        md = QMimeData()
        md.setData(MIME_IMAGEN, QByteArray(json.dumps(
            {"nombre": fmt.name(), "ancho": round(w), "alto": round(h)}).encode("utf-8")))
        datos = self.imagenes.get(fmt.name())
        img = QImage.fromData(datos[0]) if datos else self.document().resource(
            QTextDocument.ImageResource, QUrl(fmt.name()))
        if img is not None and not (isinstance(img, QImage) and img.isNull()):
            md.setImageData(img)         # para pegar en otras aplicaciones
        QApplication.clipboard().setMimeData(md)

    # -- geometría ----------------------------------------------------------
    def rect_imagen(self, pos):
        """Rectángulo de la imagen en coordenadas del viewport."""
        doc = self.document()
        fmt = formato_imagen_en(doc, pos)
        if fmt is None:
            return None
        b = doc.findBlock(pos)
        lay = b.layout()
        rel = pos - b.position()
        linea = lay.lineForTextPosition(rel)
        if not linea.isValid():
            return None
        x = linea.cursorToX(rel)
        x = x[0] if isinstance(x, tuple) else x
        w, h = self._tamano_imagen(fmt)
        origen = lay.position()
        r = QRectF(origen.x() + x, origen.y() + linea.y() + linea.ascent() - h, w, h)
        return r.translated(-self._desplazamiento())

    def imagen_en(self, punto):
        """Posición de la imagen bajo el punto (viewport) o None."""
        doc = self.document()
        p = doc.documentLayout().hitTest(punto + self._desplazamiento(), Qt.FuzzyHit)
        if p < 0:
            return None
        for pos in (p, p - 1):
            r = self.rect_imagen(pos)
            if r is not None and r.contains(punto):
                return pos
        return None

    def _centros_asas(self, r):
        return [(a, QPointF(r.x() + a[0] * r.width(), r.y() + a[1] * r.height()))
                for a in ASAS]

    def asa_en(self, punto):
        pos = self.imagen_seleccionada()
        r = self.rect_imagen(pos) if pos is not None else None
        if r is None:
            return None
        tol = (ASA / 2 + 3) / self._zoom()
        for asa, centro in self._centros_asas(r):
            if abs(punto.x() - centro.x()) <= tol and abs(punto.y() - centro.y()) <= tol:
                return asa
        return None

    def _destino_mover(self, punto, pos):
        """Número del bloque antes del cual caería la imagen (entre
        párrafos), o None si quedaría en el mismo sitio."""
        doc = self.document()
        lay = doc.documentLayout()
        p = punto + self._desplazamiento()
        b = doc.findBlock(max(0, lay.hitTest(p, Qt.FuzzyHit)))
        if not b.isValid():
            b = doc.lastBlock()
        r = lay.blockBoundingRect(b)
        destino = b.blockNumber() + (0 if p.y() < r.center().y() else 1)
        origen = doc.findBlock(pos)
        if es_bloque_imagen(origen) and destino in (origen.blockNumber(),
                                                   origen.blockNumber() + 1):
            return None
        return destino

    def _y_destino(self, destino):
        doc = self.document()
        lay = doc.documentLayout()
        if destino < doc.blockCount():
            y = lay.blockBoundingRect(doc.findBlockByNumber(destino)).top()
        else:
            y = lay.blockBoundingRect(doc.lastBlock()).bottom()
        return y - self._desplazamiento().y()

    def _tamano_redimensionado(self, punto, shift):
        """Tamaño nuevo según el asa arrastrada. Como la imagen está
        alineada (centrada por defecto), un asa lateral mueve los dos lados
        a la vez: el ancho cambia el doble del desplazamiento y el borde
        sigue al ratón. Arriba la imagen no se mueve (va en su renglón):
        crece hacia abajo y el texto de abajo se recorre."""
        a = self._accion
        w0, h0 = a["tam"]
        (fx, fy), d = a["asa"], punto - a["inicio"]
        alin = a["alineacion"]
        if fx == 0.5:
            dw = 0.0
        else:
            signo = 1 if fx == 1 else -1
            factor = 2.0 if alin & Qt.AlignHCenter else 1.0   # centrada: ambos lados
            dw = signo * d.x() * factor
        dh = 0.0 if fy == 0.5 else (d.y() if fy == 1 else -d.y())
        w, h = w0 + dw, h0 + dh
        max_w, max_h = a["max_w"], self._alto_texto()
        proporcional = fx != 0.5 and fy != 0.5 and not shift   # esquinas
        if proporcional:
            # Un solo factor de escala, continuo: el cambio (dw, dh) se
            # proyecta sobre la diagonal (w0, h0). Así no hay saltos al
            # pasar por el tamaño original (antes se elegía el eje que más
            # cambiaba y la elección brincaba de uno a otro).
            k = 1 + (dw * w0 + dh * h0) / (w0 * w0 + h0 * h0)
            # Límites sobre el factor (nunca llega a 0 ni a negativo)
            k_min = max(IMAGEN_MIN / w0, IMAGEN_MIN / h0)
            k_max = max(k_min, min(max_w / w0, max_h / h0))
            k = max(k_min, min(k_max, k))
            w, h = w0 * k, h0 * k
        else:
            w = max(IMAGEN_MIN, min(max_w, w))
            h = max(IMAGEN_MIN, min(max_h, h))
        return round(w), round(h)

    def _aplicar_tamano(self, pos, w, h):
        """Cambia el tamaño de la imagen (un paso de deshacer)."""
        doc = self.document()
        fmt = formato_imagen_en(doc, pos)
        if fmt is None:
            return
        nuevo = QTextImageFormat(fmt)
        nuevo.setWidth(w)
        nuevo.setHeight(h)
        c = QTextCursor(doc)
        c.setPosition(pos)
        c.setPosition(pos + 1, QTextCursor.KeepAnchor)
        c.setCharFormat(nuevo)

    def _terminar_redimensionado(self, a):
        """Deshace los pasos intermedios del arrastre y deja uno solo."""
        doc = self.document()
        pos = a["pos"]
        final = a.get("final")
        self._eligiendo = True           # deshacer no debe deseleccionar
        while doc.availableUndoSteps() > a["pasos"]:
            doc.undo()
        if final is not None and final != tuple(round(x) for x in a["tam"]):
            self._aplicar_tamano(pos, *final)
        self._eligiendo = False
        self.seleccionar_imagen(pos)

    # -- ratón --------------------------------------------------------------
    def mousePressEvent(self, e):
        punto = e.position()
        if e.button() in (Qt.LeftButton, Qt.RightButton):
            sel = self.imagen_seleccionada()
            asa = self.asa_en(punto) if sel is not None else None
            if asa is not None and e.button() == Qt.LeftButton:
                b = self.document().findBlock(sel)
                self._accion = {"tipo": "redimensionar", "asa": asa, "inicio": punto,
                                "pos": sel,
                                "tam": self._tamano_imagen(
                                    formato_imagen_en(self.document(), sel)),
                                "alineacion": b.blockFormat().alignment()
                                & Qt.AlignHorizontal_Mask,
                                "pasos": self.document().availableUndoSteps(),
                                "max_w": self.ancho_texto(self.document().findBlock(sel))}
                e.accept()
                return
            pos = self.imagen_en(punto)
            if pos is not None:
                self.seleccionar_imagen(pos)
                if e.button() == Qt.LeftButton:
                    self._accion = {"tipo": "presionado", "inicio": punto, "pos": pos}
                self.setFocus()
                e.accept()
                return
            self.deseleccionar()
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        a = self._accion
        punto = e.position()
        if a is not None and a["tipo"] == "redimensionar":
            tam = self._tamano_redimensionado(
                punto, bool(e.modifiers() & Qt.ShiftModifier))
            if tam != a.get("final"):
                # En vivo: la imagen real cambia y el texto se recoloca
                a["final"] = tam
                self._eligiendo = True
                self._aplicar_tamano(a["pos"], *tam)
                self._eligiendo = False
                self.viewport().update()
            e.accept()
            return
        if a is not None and a["tipo"] in ("presionado", "mover"):
            if a["tipo"] == "presionado":
                umbral = QApplication.startDragDistance() / self._zoom()
                if (punto - a["inicio"]).manhattanLength() < umbral:
                    return
                a["tipo"] = "mover"
                self._poner_cursor(Qt.ClosedHandCursor)
            self._autodesplazar(punto)
            a["punto"] = punto
            a["destino"] = self._destino_mover(punto, a["pos"])
            self.viewport().update()
            e.accept()
            return
        super().mouseMoveEvent(e)
        if not e.buttons():
            self._actualizar_forma(punto)

    def mouseReleaseEvent(self, e):
        a = self._accion
        if a is not None:
            self._accion = None
            if a["tipo"] == "redimensionar":
                self._terminar_redimensionado(a)
            elif a["tipo"] == "mover" and a.get("destino") is not None:
                self.mover_imagen(a["pos"], a["destino"])
            self._actualizar_forma(e.position())
            self.viewport().update()
            e.accept()
            return
        super().mouseReleaseEvent(e)

    def mouseDoubleClickEvent(self, e):
        if self.imagen_en(e.position()) is not None:
            e.accept()                   # no selecciona palabras
            return
        super().mouseDoubleClickEvent(e)

    def viewportEvent(self, e):
        if e.type() == QEvent.HoverMove and self._accion is None:
            self._actualizar_forma(e.position())
        return super().viewportEvent(e)

    def _poner_cursor(self, forma):
        if forma != self._forma_cursor:
            self._forma_cursor = forma
            if forma is None:
                self.viewport().setCursor(Qt.IBeamCursor)
            else:
                self.viewport().setCursor(forma)

    def _actualizar_forma(self, punto):
        asa = self.asa_en(punto) if self._img is not None else None
        if asa is not None:
            self._poner_cursor(CURSOR_ASA[asa])
        elif self.imagen_en(punto) is not None:
            self._poner_cursor(Qt.SizeAllCursor)
        else:
            self._poner_cursor(None)

    def _autodesplazar(self, punto):
        """Al arrastrar cerca del borde de la ventana, la hoja se desplaza."""
        v = self.vista
        if v is None:
            return
        en_vista = v.mapFromScene(v.proxy.mapToScene(punto))
        alto = v.viewport().height()
        barra = v.verticalScrollBar()
        if en_vista.y() < 40:
            barra.setValue(barra.value() - 24)
        elif en_vista.y() > alto - 40:
            barra.setValue(barra.value() + 24)

    # -- dibujo del marco y los puntos de arrastre --------------------------
    def paintEvent(self, e):
        super().paintEvent(e)
        if self._img is None and self._accion is None:
            return
        p = QPainter(self.viewport())
        p.setRenderHint(QPainter.Antialiasing)
        self._pintar_imagen(p)
        p.end()

    def _pintar_imagen(self, p):
        z = self._zoom()
        color = QColor(COLOR_MARCO)
        fino = QPen(color, 0)
        a = self._accion
        pos = self.imagen_seleccionada()
        r = self.rect_imagen(pos) if pos is not None else None
        if a is not None and a["tipo"] == "mover":
            if a.get("destino") is not None:
                # Línea donde caerá la imagen (entre párrafos)
                ff = self.document().rootFrame().frameFormat()
                x0 = ff.leftMargin() - self._desplazamiento().x()
                x1 = x0 + self.ancho_texto()
                y = self._y_destino(a["destino"])
                grueso = QPen(color, 2.5)
                grueso.setCosmetic(True)
                p.setPen(grueso)
                p.drawLine(QPointF(x0, y), QPointF(x1, y))
                p.setPen(Qt.NoPen)
                p.setBrush(color)
                t = 5 / z
                for x, s in ((x0, 1), (x1, -1)):
                    p.drawPolygon(QPolygonF([QPointF(x, y - t), QPointF(x + s * t, y),
                                             QPointF(x, y + t)]))
            if r is not None and "punto" in a:
                # Silueta de la imagen siguiendo al ratón
                fantasma = QRectF(r).translated(a["punto"] - a["inicio"])
                p.setBrush(QColor(43, 121, 215, 40))
                guion = QPen(color, 0, Qt.DashLine)
                p.setPen(guion)
                p.drawRect(fantasma)
        if a is not None and a["tipo"] == "redimensionar" and r is not None:
            # La imagen real ya tiene el tamaño nuevo: solo la medida
            texto = f"{round(r.width())} × {round(r.height())} px"
            f = QFont(self.font())
            f.setPointSizeF(max(1.0, 9 / z))
            p.setFont(f)
            fm = QFontMetrics(f)
            caja = QRectF(0, 0, fm.horizontalAdvance(texto) + 10 / z, fm.height() + 4 / z)
            caja.moveCenter(QPointF(r.center().x(), r.bottom() - 10 / z - caja.height() / 2))
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(40, 40, 40, 200))
            p.drawRoundedRect(caja, 3 / z, 3 / z)
            p.setPen(QColor("#FFFFFF"))
            p.drawText(caja, Qt.AlignCenter, texto)
        if r is None:
            return
        p.setBrush(Qt.NoBrush)
        p.setPen(fino)
        p.drawRect(r)
        s = ASA / z
        p.setBrush(QColor("#FFFFFF"))
        p.setPen(fino)
        for _, c in self._centros_asas(r):
            p.drawRect(QRectF(c.x() - s / 2, c.y() - s / 2, s, s))

    # -- teclado ------------------------------------------------------------
    def _parrafo_junto_a_imagen(self, b, antes):
        """Crea un párrafo Body vacío antes o después del párrafo de una
        imagen y lleva ahí el cursor (el texto nunca entra en él)."""
        c = QTextCursor(self.document())
        c.beginEditBlock()
        if antes and b.previous().isValid():
            previo = b.previous()
            c.setPosition(previo.position() + previo.length() - 1)
            c.insertBlock(self.body_block_format(), self.body_char_format())
        elif antes:
            bf, cf = b.blockFormat(), b.charFormat()
            c.setPosition(b.position())
            c.insertBlock(bf, cf)
            c.setPosition(0)
            c.setBlockFormat(self.body_block_format())
            c.setBlockCharFormat(self.body_char_format())
        else:
            c.setPosition(b.position() + b.length() - 1)
            c.insertBlock(self.body_block_format(), self.body_char_format())
        self._interlineado_imagenes()
        c.endEditBlock()
        self.setTextCursor(c)
        self.setCurrentCharFormat(self.body_char_format())
        self.formatoParrafoCambiado.emit()

    def dropEvent(self, e):
        self._soltando = True              # mover dentro del documento
        try:
            super().dropEvent(e)
        finally:
            self._soltando = False

    def pegar_tramos(self, parrafos):
        """Pega los párrafos en estilo Body (un solo paso de deshacer)."""
        doc = self.document()
        c = self.textCursor()
        c.beginEditBlock()
        if c.hasSelection():
            c.removeSelectedText()
        primer_bloque = c.blockNumber()
        base = self.body_char_format()
        for k, par in enumerate(parrafos):
            if k > 0:
                c.insertBlock(self.body_block_format(), base)
            for t in par["tramos"]:
                if "img" in t:
                    nombre = t["img"]
                    if doc.resource(QTextDocument.ImageResource, QUrl(nombre)) is None:
                        continue
                    c.insertImage(self._formato_imagen(nombre, t.get("w", 0), t.get("h", 0)))
                    continue
                c.insertText(t["t"], formato_pegado(base, t))
        ultimo = c.blockNumber()
        # Todos los párrafos que tocó el pegado quedan en Body (los de una
        # imagen sola, con su formato de imagen)
        b = doc.findBlockByNumber(primer_bloque)
        q = QTextCursor(doc)
        while b.isValid() and b.blockNumber() <= ultimo:
            q.setPosition(b.position())
            if es_bloque_imagen(b):
                q.setBlockFormat(formato_bloque_imagen())
            else:
                q.setBlockFormat(self.body_block_format())
                if b.length() <= 1:
                    q.setBlockCharFormat(base)
            b = b.next()
        self._interlineado_imagenes()
        c.endEditBlock()
        self.setTextCursor(c)
        self.ensureCursorVisible()
        self.formatoParrafoCambiado.emit()

    def insertar_texto(self, texto):
        """Escribe `texto` en el cursor con el formato actual, como si se
        tecleara (reemplaza la selección; un solo paso de deshacer)."""
        doc = self.document()
        # Un bloque de edición: Qt no lo fusiona con lo tecleado antes
        grupo = QTextCursor(doc)
        grupo.beginEditBlock()
        sel = self.imagen_seleccionada()
        if sel is not None:
            # Imagen elegida: el texto va en un párrafo nuevo debajo
            self.deseleccionar()
            self._parrafo_junto_a_imagen(doc.findBlock(sel), antes=False)
        else:
            c = self.textCursor()
            if not c.hasSelection() and es_bloque_imagen(c.block()):
                self._parrafo_junto_a_imagen(
                    c.block(), antes=c.position() == c.block().position())
        c = self.textCursor()
        c.insertText(texto)
        grupo.endEditBlock()
        self.setTextCursor(c)
        self.ensureCursorVisible()

    def _borrar_vecino(self, q):
        """Borra un párrafo vacío junto a una imagen y recalcula su espacio
        (un solo paso de deshacer)."""
        q.beginEditBlock()
        q.removeSelectedText()
        self._interlineado_imagenes()
        q.endEditBlock()

    def _teclas_imagen(self, e):
        """Teclas especiales con imágenes. Devuelve True si ya se atendió."""
        k = e.key()
        escribe = bool(e.text()) and e.text().isprintable() and \
            not e.modifiers() & (Qt.ControlModifier | Qt.AltModifier | Qt.MetaModifier)
        enter = k in (Qt.Key_Return, Qt.Key_Enter)
        doc = self.document()
        sel = self.imagen_seleccionada()
        if sel is not None:
            b = doc.findBlock(sel)
            if k == Qt.Key_Escape:
                self.deseleccionar()
                return True
            if k in (Qt.Key_Delete, Qt.Key_Backspace):
                self.eliminar_imagen()
                return True
            if e.matches(QKeySequence.Copy):
                self.copy()
                return True
            if e.matches(QKeySequence.Cut):
                self.cut()
                return True
            if enter and es_bloque_imagen(b):
                self.deseleccionar()
                self._parrafo_junto_a_imagen(b, antes=False)
                return True
            if k in (Qt.Key_Left, Qt.Key_Up, Qt.Key_Right, Qt.Key_Down):
                self.deseleccionar()
                c = QTextCursor(doc)
                c.setPosition(sel if k in (Qt.Key_Left, Qt.Key_Up) else sel + 1)
                self.setTextCursor(c)
                return True
            if escribe:
                return True              # no se escribe sobre la imagen
            self.deseleccionar()
            return False
        c = self.textCursor()
        if c.hasSelection():
            return False
        b = c.block()
        if es_bloque_imagen(b):
            antes = c.position() == b.position()
            if enter or escribe:
                self._parrafo_junto_a_imagen(b, antes)
                return enter             # la letra se escribe en el nuevo
            if k == Qt.Key_Backspace:
                previo = b.previous()
                if antes and previo.isValid() and previo.length() <= 1:
                    # Quita el párrafo vacío de arriba (la imagen no cambia)
                    q = QTextCursor(doc)
                    q.setPosition(previo.position())
                    q.setPosition(b.position(), QTextCursor.KeepAnchor)
                    self._borrar_vecino(q)
                elif antes and previo.isValid():
                    q = QTextCursor(previo)
                    q.movePosition(QTextCursor.EndOfBlock)
                    self.setTextCursor(q)
                elif not antes:
                    self.seleccionar_imagen(b.position())
                return True
            if k == Qt.Key_Delete:
                siguiente = b.next()
                if not antes and siguiente.isValid() and siguiente.length() <= 1:
                    q = QTextCursor(doc)
                    q.setPosition(b.position() + b.length() - 1)
                    q.setPosition(siguiente.position(), QTextCursor.KeepAnchor)
                    self._borrar_vecino(q)
                elif not antes and siguiente.isValid():
                    self.setTextCursor(QTextCursor(siguiente))
                elif antes:
                    self.seleccionar_imagen(b.position())
                return True
            return False
        # Borrar junto a una imagen no la mete en el texto: la selecciona
        if k == Qt.Key_Backspace and c.atBlockStart() and c.currentList() is None \
                and es_bloque_imagen(b.previous()):
            if b.length() <= 1:
                q = QTextCursor(doc)
                q.setPosition(b.position() - 1)
                q.setPosition(b.position(), QTextCursor.KeepAnchor)
                self._borrar_vecino(q)
            else:
                self.seleccionar_imagen(b.previous().position())
            return True
        if k == Qt.Key_Delete and c.atBlockEnd() and es_bloque_imagen(b.next()):
            if b.length() <= 1:
                q = QTextCursor(doc)
                q.setPosition(b.position())
                q.setPosition(b.next().position(), QTextCursor.KeepAnchor)
                self._borrar_vecino(q)
            else:
                self.seleccionar_imagen(b.next().position())
            return True
        return False


# ---------------------------------------------------------------------------
# Cinta de estilos (galería de casillas, estilo procesador de textos)
# ---------------------------------------------------------------------------
def fuente_muestra(e):
    """Fuente de vista previa: misma familia y estilo, tamaño a escala."""
    f = QFont(e["font_family"])
    f.setPointSizeF(max(9.0, min(14.0, 9 + (e["font_size"] - 12) * 0.32)))
    f.setBold(e["bold"])
    f.setItalic(e["italic"])
    return f


class CintaEstilos(QWidget):
    def __init__(self, editor, parent=None):
        super().__init__(parent)
        self.editor = editor
        self.setObjectName("cinta")
        self.lay = QHBoxLayout(self)
        self.lay.setContentsMargins(0, 0, 0, 0)
        self.lay.setSpacing(0)
        self.botones = {}
        self.ids = []
        # Ctrl+1 … Ctrl+9: los nueve primeros estilos de la cinta, en su
        # orden (también con el teclado numérico)
        self.atajos = []
        for k in range(9):
            tecla = Qt.Key(Qt.Key_1.value + k)
            for mod in (Qt.ControlModifier, Qt.ControlModifier | Qt.KeypadModifier):
                a = QShortcut(QKeySequence(QKeyCombination(mod, tecla)), self)
                a.setContext(Qt.WindowShortcut)
                a.activated.connect(lambda k=k: self._atajo(k))
                self.atajos.append(a)
        self.reconstruir()
        GESTOR.cambiado.connect(self.reconstruir)
        TEMA.cambiado.connect(self.reconstruir)
        editor.cursorPositionChanged.connect(self.sincronizar)
        editor.formatoParrafoCambiado.connect(self.sincronizar)

    def reconstruir(self):
        """Casillas unidas con los estilos visibles, en el orden elegido."""
        while self.lay.count():
            w = self.lay.takeAt(0).widget()
            if w is not None:
                w.deleteLater()
        self.botones = {}
        ids = [i for i in GESTOR.orden if GESTOR.datos[i]["visible"]]
        self.ids = ids
        for k, id_ in enumerate(ids):
            e = GESTOR.resuelto(id_)
            b = QToolButton()
            b.setText(GESTOR.nombre(id_))
            b.setCheckable(True)
            b.setToolButtonStyle(Qt.ToolButtonTextOnly)
            b.setFocusPolicy(Qt.NoFocus)
            b.setFixedHeight(30)
            b.setFont(fuente_muestra(e))
            # Ancho según el texto: nunca se recorta el nombre
            ancho = QFontMetrics(b.font()).horizontalAdvance(b.text()) + 30
            b.setFixedWidth(max(96, ancho))
            # Colores de las herramientas (los del tema), no los de la hoja.
            # En Claro: fondo blanco, borde #CFCFCF, como siempre
            t = TEMA.t
            fondo = t["campo"]
            borde = _mezcla(fondo, t["texto"], 0.21)
            letra = e["font_color"]
            if contraste(letra, fondo) < 3:      # no se leería: color del tema
                letra = t["texto"]
            borde_izq = f"1px solid {borde}" if k == 0 else "none"
            radio_i = "3px" if k == 0 else "0"
            radio_d = "3px" if k == len(ids) - 1 else "0"
            b.setStyleSheet(
                "QToolButton {"
                f" color:{letra}; background:{fondo};"
                f" border:1px solid {borde};"
                f" border-left:{borde_izq};"
                f" border-top-left-radius:{radio_i}; border-bottom-left-radius:{radio_i};"
                f" border-top-right-radius:{radio_d}; border-bottom-right-radius:{radio_d};"
                " padding:0 12px; }"
                f"QToolButton:hover {{ background:{_mezcla(fondo, t['texto'], 0.057)}; }}"
                f"QToolButton:pressed {{ background:{_mezcla(fondo, t['texto'], 0.11)}; }}"
                f"QToolButton:checked {{ background:{fondo};"
                f" border:2px solid {_mezcla(fondo, t['texto'], 0.58)}; padding:0 11px; }}")
            b.clicked.connect(lambda _=False, n=id_: self._elegir(n))
            self.lay.addWidget(b)
            self.botones[id_] = b
        self.sincronizar()

    def _elegir(self, id_):
        self.editor.aplicar_estilo(id_)
        self.sincronizar()

    def _atajo(self, k):
        """Ctrl+número: el estilo k de la cinta (si hay tantos)."""
        if k < len(self.ids):
            self._elegir(self.ids[k])

    def sincronizar(self, *_):
        """Marca el estilo del párrafo bajo el cursor."""
        actual = self.editor.estilo_actual()
        for id_, b in self.botones.items():
            b.setChecked(id_ == actual)


# ---------------------------------------------------------------------------
# Icono de la aplicación: libro azul con cinta y plumilla (dibujado, sin
# archivos externos; cada tamaño con el detalle que le cabe)
# ---------------------------------------------------------------------------
NOMBRE_APP = "SWordBook"
APP_VERSION = "0.1.1"


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


def icono_aplicacion():
    """Icono de SWordBook (libro azul con cinta y plumilla), dibujado en
    varios tamaños para que se vea nítido en la barra de título, la barra de
    tareas y el escritorio."""
    icono = QIcon()
    for lado in (16, 20, 24, 32, 40, 48, 64, 96, 128, 256):
        pm = QPixmap(lado, lado)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.scale(lado / 64, lado / 64)
        _draw_app_icon(p, detail=lado >= 32)
        p.end()
        icono.addPixmap(pm)
    return icono


app_icon = icono_aplicacion  # nombre que usa build.py


def guardar_ico(ruta):
    """Guarda el icono como .ico con todos sus tamaños (para un acceso
    directo o un ejecutable). Cada tamaño va como PNG dentro del .ico."""
    import struct
    icono = icono_aplicacion()
    imagenes = []
    for lado in (16, 20, 24, 32, 40, 48, 64, 128, 256):
        imagenes.append((lado, imagen_a_bytes(icono.pixmap(lado, lado).toImage())))
    cabecera = struct.pack("<HHH", 0, 1, len(imagenes))
    directorio, datos = b"", b""
    desplazamiento = 6 + 16 * len(imagenes)
    for lado, png in imagenes:
        directorio += struct.pack("<BBBBHHII", lado % 256, lado % 256, 0, 0, 1, 32,
                                  len(png), desplazamiento + len(datos))
        datos += png
    with open(ruta, "wb") as f:
        f.write(cabecera + directorio + datos)
    return True


# ---------------------------------------------------------------------------
# Botón ↔ ventana
# ---------------------------------------------------------------------------
class EnlaceVentana(QObject):
    """Une un botón de la barra con la ventana que abre:
    - el botón queda marcado mientras la ventana está a la vista;
    - pulsarlo otra vez cierra la ventana;
    - la ventana lleva el mismo icono que el botón.
    `crear()` construye la ventana. Con `reutilizar` se conserva entre usos
    (Párrafo, Buscar); si no, se crea una nueva cada vez y se destruye al
    cerrarla (las que muestran valores del momento). `mostrar(ventana,
    *args)` sustituye al show() normal. `al_aceptar(ventana)` se llama si
    la ventana es un QDialog y se acepta."""

    def __init__(self, accion, crear, *, reutilizar=False, mostrar=None,
                 al_aceptar=None, al_cerrar=None):
        super().__init__(accion)
        self.accion = accion
        self._crear = crear
        self._reutilizar = reutilizar
        self._mostrar = mostrar
        self._al_aceptar = al_aceptar
        self._al_cerrar = al_cerrar
        self.ventana = None
        accion.setCheckable(True)
        accion.triggered.connect(self._pulsado)

    def _pulsado(self, marcado):
        if marcado:
            self.abrir()
        else:
            self.cerrar()

    def abierta(self):
        return self.ventana is not None and self.ventana.isVisible()

    def abrir(self, *args):
        if self.ventana is None:
            v = self._crear()
            v.setWindowIcon(self.accion.icon())
            if isinstance(v, QDialog):
                v.setModal(False)            # la barra sigue respondiendo
                v.finished.connect(lambda r, v=v: self._terminada(v, r))
            v.installEventFilter(self)
            self.ventana = v
        if self._mostrar is not None:
            self._mostrar(self.ventana, *args)
        else:
            self.ventana.show()
            self.ventana.raise_()
            self.ventana.activateWindow()
        self.accion.setChecked(True)

    def cerrar(self):
        if self.ventana is not None:
            self.ventana.close()
        self.accion.setChecked(False)

    def _terminada(self, v, resultado):
        if resultado == QDialog.Accepted and self._al_aceptar is not None:
            self._al_aceptar(v)
        if self._al_cerrar is not None:
            self._al_cerrar()
        if not self._reutilizar and v is self.ventana:
            self.ventana = None
            v.removeEventFilter(self)
            v.deleteLater()

    def eventFilter(self, obj, e):
        if obj is self.ventana:
            t = e.type()
            if t == QEvent.Show:
                self.accion.setChecked(True)
            elif t == QEvent.Hide:
                self.accion.setChecked(False)
        return False


# ---------------------------------------------------------------------------
# Fila de herramientas: izquierda, centro y derecha
# ---------------------------------------------------------------------------
def seccion_barra(nombre):
    """Barra sin borde ni relleno para una sección de una fila."""
    b = QToolBar(nombre)
    b.setObjectName("seccion")
    b.setMovable(False)
    b.setFloatable(False)
    b.setIconSize(QSize(20, 20))
    b.setToolButtonStyle(Qt.ToolButtonIconOnly)
    return b


class FilaHerramientas(QToolBar):
    """Una fila de la barra de herramientas con tres secciones: la de la
    izquierda pegada al borde izquierdo, la de la derecha al derecho y la
    del centro centrada en la ventana (no en el hueco que dejan las otras).
    Si la ventana es estrecha, la sección del centro se encoge y muestra
    el botón » con lo que no cabe."""

    def __init__(self, titulo, izquierda, centro, derecha, parent=None):
        super().__init__(titulo, parent)
        self.setMovable(False)
        self.setFloatable(False)
        self.toggleViewAction().setVisible(False)   # no se puede ocultar
        self.secciones = (izquierda, centro, derecha)
        for barra in self.secciones:
            barra.installEventFilter(self)
            self._sin_margen(barra)
        cont = QWidget()
        cont.setObjectName("filaHerramientas")
        self.grid = QGridLayout(cont)
        self.grid.setContentsMargins(0, 0, 0, 0)
        self.grid.setHorizontalSpacing(12)
        for col, (barra, alin) in enumerate(((izquierda, Qt.AlignLeft),
                                             (centro, Qt.AlignHCenter),
                                             (derecha, Qt.AlignRight))):
            self.grid.addWidget(barra, 0, col, alin | Qt.AlignVCenter)
        # Laterales con el mismo ancho: así el centro queda centrado
        self.grid.setColumnStretch(0, 1)
        self.grid.setColumnStretch(2, 1)
        cont.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.addWidget(cont)
        self._igualar_laterales()

    @staticmethod
    def _sin_margen(barra):
        """Qt devuelve a la barra su margen interno (8 px) cada vez que le
        aplica el estilo; dentro de una fila sobra."""
        lay = barra.layout()
        if lay is not None and not lay.contentsMargins().isNull():
            lay.setContentsMargins(0, 0, 0, 0)

    def _igualar_laterales(self):
        izq, _, der = self.secciones
        ancho = max(izq.sizeHint().width(), der.sizeHint().width())
        self.grid.setColumnMinimumWidth(0, ancho)
        self.grid.setColumnMinimumWidth(2, ancho)

    def eventFilter(self, obj, e):
        t = e.type()
        if t in (QEvent.Polish, QEvent.StyleChange, QEvent.Show,
                 QEvent.LayoutRequest):
            self._sin_margen(obj)
            if obj is not self.secciones[1]:
                self._igualar_laterales()   # cambió una sección lateral
        return False


# ---------------------------------------------------------------------------
# Barra de herramientas de fuente
# ---------------------------------------------------------------------------
class BarraFuente(QToolBar):
    def __init__(self, editor, parent=None):
        super().__init__("Fuente", parent)
        self.editor = editor
        self.setMovable(False)
        self.setIconSize(QSize(20, 20))
        self.setToolButtonStyle(Qt.ToolButtonIconOnly)

        # Estilos: cinta de casillas unidas; cada una muestra su nombre con
        # el aspecto del estilo y lo aplica completo (como Body). La ventana
        # la coloca en su propia fila, con el botón de Ajustes de estilo
        self.cinta = CintaEstilos(editor)
        self.botones_estilo = self.cinta.botones
        self.act_ajustes_estilo = self._accion(
            icono_ajustes_estilo(), "Ajustes de estilo", None)
        self.enlace_estilos = EnlaceVentana(
            self.act_ajustes_estilo, lambda: VentanaEstilos(self.window()),
            al_aceptar=self._estilos_aceptados, al_cerrar=self.editor.setFocus)
        # Atajo sin botón propio (una QAction en la barra crearía un botón vacío)
        atajo_body = QShortcut(QKeySequence("Ctrl+Shift+N"), self)
        atajo_body.setContext(Qt.WindowShortcut)
        atajo_body.activated.connect(editor.aplicar_estilo_body)

        # Tipo de fuente
        self.combo_fuente = QFontComboBox()
        self.combo_fuente.setMinimumWidth(170)
        self.combo_fuente.setCurrentFont(QFont(BODY["font_family"]))
        self.combo_fuente.currentFontChanged.connect(self.cambiar_fuente)
        self.addWidget(self.combo_fuente)

        # Tamaño
        self.combo_tamano = QComboBox()
        self.combo_tamano.setEditable(True)
        self.combo_tamano.setFixedWidth(64)
        self.combo_tamano.setInsertPolicy(QComboBox.NoInsert)
        self.combo_tamano.lineEdit().setValidator(QIntValidator(1, 1638, self))
        self.combo_tamano.addItems([str(t) for t in TAMANOS])
        self.combo_tamano.setCurrentText(str(BODY["font_size"]))
        self.combo_tamano.textActivated.connect(self.cambiar_tamano)
        self.addWidget(self.combo_tamano)

        # Aumentar / disminuir
        self.addAction(self._accion(
            crear_icono("A", pequeno="+"), "Aumentar tamaño (Ctrl+Shift+>)",
            lambda: self.paso_tamano(+1), "Ctrl+Shift+>"))
        self.addAction(self._accion(
            crear_icono("A", pequeno="−"), "Disminuir tamaño (Ctrl+Shift+<)",
            lambda: self.paso_tamano(-1), "Ctrl+Shift+<"))

        self.addSeparator()

        # Estilos
        self.act_negrita = self._accion(
            crear_icono("N", negrita=True), "Negrita (Ctrl+B)",
            self.cambiar_negrita, QKeySequence.Bold, marcable=True)
        self.act_cursiva = self._accion(
            crear_icono("K", cursiva=True), "Cursiva (Ctrl+I)",
            self.cambiar_cursiva, QKeySequence.Italic, marcable=True)
        self.act_subrayado = self._accion(
            crear_icono("S", subrayado=True), "Subrayado (Ctrl+U)",
            self.cambiar_subrayado, QKeySequence.Underline, marcable=True)
        self.act_tachado = self._accion(
            crear_icono("abc", tachado=True), "Tachado (Ctrl+Shift+X)",
            self.cambiar_tachado, "Ctrl+Shift+X", marcable=True)
        for a in (self.act_negrita, self.act_cursiva,
                  self.act_subrayado, self.act_tachado):
            self.addAction(a)

        self.addSeparator()

        # Subíndice / superíndice (mutuamente excluyentes)
        self.act_sub = self._accion(
            crear_icono("X", pequeno="2", pos_pequeno="abajo"), "Subíndice (Ctrl+=)",
            self.cambiar_subindice, "Ctrl+=", marcable=True)
        self.act_sup = self._accion(
            crear_icono("X", pequeno="2", pos_pequeno="arriba"), "Superíndice (Ctrl+Shift+=)",
            self.cambiar_superindice, "Ctrl+Shift+=", marcable=True)
        self.addAction(self.act_sub)
        self.addAction(self.act_sup)

        self.addSeparator()

        # Colores
        self.btn_color = BotonColor(
            "A", "#C00000", PALETA_FUENTE, "Automático", BODY["font_color"])
        self.btn_color.colorElegido.connect(self.cambiar_color)
        self.addWidget(self.btn_color)

        self.btn_resaltado = BotonColor(
            "ab", "#FFF59D", PALETA_RESALTADO, "Sin color", None,
            resaltado=True)
        self.btn_resaltado.colorElegido.connect(self.cambiar_resaltado)
        self.addWidget(self.btn_resaltado)

        self.addSeparator()

        # Alineación (exclusiva: solo una activa)
        self.grupo_alineacion = QActionGroup(self)
        self.grupo_alineacion.setExclusive(True)
        self.acciones_alineacion = {}
        for tipo, alin, tip, atajo in (
            ("izquierda", Qt.AlignLeft, "Alinear a la izquierda (Ctrl+L)", "Ctrl+L"),
            ("centro", Qt.AlignHCenter, "Centrar (Ctrl+E)", "Ctrl+E"),
            ("derecha", Qt.AlignRight, "Alinear a la derecha (Ctrl+R)", "Ctrl+R"),
            ("justificado", Qt.AlignJustify, "Justificar (Ctrl+J)", "Ctrl+J"),
        ):
            a = self._accion(icono_parrafo(tipo), tip,
                             lambda _=False, al=alin: self.cambiar_alineacion(al),
                             atajo, marcable=True)
            self.grupo_alineacion.addAction(a)
            self.addAction(a)
            self.acciones_alineacion[alin] = a

        self.addSeparator()

        # Sangría izquierda (con límites 0 … SANGRIA_MAX px)
        self.act_reducir = self._accion(
            icono_parrafo("reducir"), "Reducir sangría (Ctrl+Shift+M)",
            lambda: self.cambiar_sangria(-1), "Ctrl+Shift+M")
        self.act_aumentar = self._accion(
            icono_parrafo("aumentar"), "Aumentar sangría (Ctrl+M)",
            lambda: self.cambiar_sangria(+1), "Ctrl+M")
        self.addAction(self.act_reducir)
        self.addAction(self.act_aumentar)

        self.addSeparator()

        # Lista: clic = viñetas (o el último tipo usado); flecha = elegir tipo
        self.tipo_lista = QTextListFormat.ListDisc
        self.btn_lista = QToolButton()
        self.btn_lista.setPopupMode(QToolButton.MenuButtonPopup)
        self.btn_lista.setCheckable(True)
        self.btn_lista.setIcon(icono_lista("vinetas"))
        self.btn_lista.clicked.connect(
            lambda: editor.aplicar_lista(self.tipo_lista))
        menu = QMenu(self.btn_lista)
        for opcion in TIPOS_LISTA:
            if opcion is None:
                menu.addSeparator()
                continue
            texto, estilo = opcion
            a = menu.addAction(texto)
            a.setIcon(icono_lista("vinetas" if estilo in VINETAS else "numeros"))
            a.triggered.connect(lambda _=False, st=estilo: self.elegir_lista(st))
        menu.addSeparator()
        menu.addAction("Quitar lista").triggered.connect(self.quitar_lista)
        self.btn_lista.setMenu(menu)
        self.addWidget(self.btn_lista)

        self.addSeparator()

        # Párrafo: abre la ventana con espaciado y sangrías
        self.act_parrafo = self._accion(
            crear_icono("¶", negrita=True), "Párrafo: espaciado y sangría", None)
        self.enlace_parrafo = EnlaceVentana(
            self.act_parrafo, lambda: VentanaParrafo(self.editor, self.window()),
            reutilizar=True)
        self.addAction(self.act_parrafo)

        # Sincronizar la barra con el formato bajo el cursor
        editor.currentCharFormatChanged.connect(self.sincronizar)
        editor.cursorPositionChanged.connect(self.sincronizar_parrafo)
        editor.selectionChanged.connect(self.sincronizar_parrafo)
        editor.formatoParrafoCambiado.connect(self.sincronizar_parrafo)
        self.sincronizar(editor.currentCharFormat())
        self.sincronizar_parrafo()

    # -- utilidades ---------------------------------------------------------
    def _accion(self, icono, texto, slot, atajo=None, marcable=False):
        a = QAction(icono, texto.split(" (")[0], self)
        a.setCheckable(marcable)
        if atajo:
            a.setShortcut(QKeySequence(atajo))
            a.setShortcutContext(Qt.WindowShortcut)
        if slot is not None:
            a.triggered.connect(slot)
        return a

    def _aplicar(self, fmt):
        self.editor.combinar_formato(fmt)
        self.editor.setFocus()

    def _tamano_actual(self):
        fmt = self.editor.currentCharFormat()
        t = fmt.fontPointSize() or fmt.font().pointSizeF()
        return t if t > 0 else BODY["font_size"]

    # -- acciones -----------------------------------------------------------
    def cambiar_fuente(self, fuente):
        fmt = QTextCharFormat()
        fmt.setFontFamilies([fuente.family()])
        self._aplicar(fmt)

    def cambiar_tamano(self, texto):
        try:
            t = float(texto)
        except ValueError:
            return
        if t > 0:
            fmt = QTextCharFormat()
            fmt.setFontPointSize(t)
            self._aplicar(fmt)

    def paso_tamano(self, direccion):
        actual = self._tamano_actual()
        if direccion > 0:
            nuevo = next((t for t in TAMANOS if t > actual), actual + 10)
        else:
            menores = [t for t in TAMANOS if t < actual]
            nuevo = menores[-1] if menores else max(1, actual - 1)
        self.cambiar_tamano(str(nuevo))

    def cambiar_negrita(self, marcado):
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Bold if marcado else QFont.Normal)
        self._aplicar(fmt)

    def cambiar_cursiva(self, marcado):
        fmt = QTextCharFormat()
        fmt.setFontItalic(marcado)
        self._aplicar(fmt)

    def cambiar_subrayado(self, marcado):
        fmt = QTextCharFormat()
        fmt.setFontUnderline(marcado)
        self._aplicar(fmt)

    def cambiar_tachado(self, marcado):
        fmt = QTextCharFormat()
        fmt.setFontStrikeOut(marcado)
        self._aplicar(fmt)

    def _vertical(self, alineacion, marcado):
        fmt = QTextCharFormat()
        fmt.setVerticalAlignment(
            alineacion if marcado else QTextCharFormat.AlignNormal)
        self._aplicar(fmt)

    def cambiar_subindice(self, marcado):
        self._vertical(QTextCharFormat.AlignSubScript, marcado)

    def cambiar_superindice(self, marcado):
        self._vertical(QTextCharFormat.AlignSuperScript, marcado)

    def cambiar_color(self, color):
        fmt = QTextCharFormat()
        fmt.setForeground(color if color.isValid() else QColor(BODY["font_color"]))
        self._aplicar(fmt)

    def cambiar_resaltado(self, color):
        fmt = QTextCharFormat()
        fmt.setBackground(QBrush(color) if color.isValid() else QBrush(Qt.NoBrush))
        self._aplicar(fmt)

    def abrir_ajustes_estilo(self):
        self.enlace_estilos.abrir()

    def _estilos_aceptados(self, dialogo):
        GESTOR.adoptar(dialogo.trabajo)
        self.editor.actualizar_estilos()
        guardar_estilos(GESTOR)

    # -- párrafo ------------------------------------------------------------
    @property
    def ventana_parrafo(self):
        return self.enlace_parrafo.ventana

    def abrir_parrafo(self):
        self.enlace_parrafo.abrir()

    def cambiar_alineacion(self, alineacion):
        fmt = QTextBlockFormat()
        fmt.setAlignment(alineacion)
        self.editor.combinar_parrafo(fmt)  # todos los párrafos seleccionados
        self.editor.setFocus()

    def elegir_lista(self, estilo):
        self.tipo_lista = estilo
        self.btn_lista.setIcon(icono_lista("vinetas" if estilo in VINETAS else "numeros"))
        cursor = self.editor.textCursor()
        l = cursor.currentList()
        if l is not None and estilo_base(l.format().style()) == estilo_base(estilo) \
                and l.format().style() != estilo:
            # Mismo tipo con otra viñeta: solo cambia el símbolo
            f = l.format()
            f.setStyle(estilo)
            l.setFormat(f)
            self.editor.formatoParrafoCambiado.emit()
            self.editor.setFocus()
            return
        if l is not None and l.format().style() == estilo:
            self.editor.setFocus()      # ya es ese tipo: no la quita
            return
        self.editor.aplicar_lista(estilo)

    def quitar_lista(self):
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        for b in self.editor.bloques_seleccionados():
            self.editor._quitar_de_lista(b)
        cursor.endEditBlock()
        self.editor.formatoParrafoCambiado.emit()
        self.editor.setFocus()

    def cambiar_sangria(self, paso):
        """Mueve la sangría izquierda al múltiplo de SANGRIA_PASO siguiente
        o anterior, sin salirse de 0 … SANGRIA_MAX."""
        cursor = QTextCursor(self.editor.document())
        cursor.beginEditBlock()  # un solo paso de deshacer
        for b in self.editor.bloques_seleccionados():
            actual = b.blockFormat().leftMargin()
            if paso > 0:
                nuevo = (int(actual // SANGRIA_PASO) + 1) * SANGRIA_PASO
            else:
                nuevo = (int(-(-actual // SANGRIA_PASO)) - 1) * SANGRIA_PASO
            nuevo = max(0, min(SANGRIA_MAX, nuevo))
            cursor.setPosition(b.position())
            fmt = QTextBlockFormat()
            fmt.setLeftMargin(nuevo)
            cursor.mergeBlockFormat(fmt)
        cursor.endEditBlock()
        self.editor.formatoParrafoCambiado.emit()
        self.editor.setFocus()

    def sincronizar_parrafo(self):
        self.btn_lista.setChecked(self.editor.textCursor().currentList() is not None)
        bloques = self.editor.bloques_seleccionados()
        alin = self.editor.textCursor().blockFormat().alignment() & Qt.AlignHorizontal_Mask
        if alin == Qt.AlignLeading:
            alin = Qt.AlignLeft
        elif alin == Qt.AlignTrailing:
            alin = Qt.AlignRight
        accion = self.acciones_alineacion.get(alin)
        if accion:
            accion.setChecked(True)

        margenes = [b.blockFormat().leftMargin() for b in bloques]
        # Se desactiva el botón cuando ningún párrafo puede moverse más
        self.act_reducir.setEnabled(any(m > 0 for m in margenes))
        self.act_aumentar.setEnabled(any(m < SANGRIA_MAX for m in margenes))

    # -- estado -------------------------------------------------------------
    def sincronizar(self, fmt):
        f = fmt.font()

        self.combo_fuente.blockSignals(True)
        familias = fmt.fontFamilies()
        familia = familias[0] if familias else f.family()
        self.combo_fuente.setCurrentFont(QFont(familia))
        self.combo_fuente.blockSignals(False)

        t = fmt.fontPointSize() or f.pointSizeF()
        self.combo_tamano.blockSignals(True)
        self.combo_tamano.setCurrentText(f"{t:g}" if t > 0 else "")
        self.combo_tamano.blockSignals(False)

        self.act_negrita.setChecked(f.bold())
        self.act_cursiva.setChecked(f.italic())
        self.act_subrayado.setChecked(fmt.fontUnderline())
        self.act_tachado.setChecked(fmt.fontStrikeOut())

        va = fmt.verticalAlignment()
        self.act_sub.setChecked(va == QTextCharFormat.AlignSubScript)
        self.act_sup.setChecked(va == QTextCharFormat.AlignSuperScript)


# ---------------------------------------------------------------------------
# Ventana "Párrafo" (todas las herramientas de espaciado)
# ---------------------------------------------------------------------------
class VentanaParrafo(QDialog):
    """Ventana no modal: se puede dejar abierta mientras se escribe y siempre
    muestra los valores del párrafo bajo el cursor."""

    def __init__(self, editor, parent=None):
        super().__init__(parent)
        self.editor = editor
        self.setWindowTitle("Párrafo")
        self.setModal(False)

        grupo = QGroupBox("Espaciado")
        form = QFormLayout(grupo)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)

        self.combo_interlineado = QComboBox()
        self.combo_interlineado.setEditable(True)
        self.combo_interlineado.setInsertPolicy(QComboBox.NoInsert)
        self.combo_interlineado.setFixedWidth(90)
        validador = QDoubleValidator(0.5, 5.0, 2, self)
        validador.setNotation(QDoubleValidator.StandardNotation)
        validador.setLocale(QLocale.c())  # siempre punto decimal
        self.combo_interlineado.lineEdit().setValidator(validador)
        self.combo_interlineado.addItems(INTERLINEADOS)
        self.combo_interlineado.textActivated.connect(self.cambiar_interlineado)
        form.addRow("Interlineado:", self.combo_interlineado)

        self.spin_antes = self._spin(0, ESPACIO_MAX, 6, self.cambiar_antes)
        form.addRow("Antes:", self.spin_antes)
        self.spin_despues = self._spin(0, ESPACIO_MAX, 6, self.cambiar_despues)
        form.addRow("Después:", self.spin_despues)
        self.spin_letras = self._spin(LETRAS_MIN, LETRAS_MAX, 1, self.cambiar_letras)
        form.addRow("Entre letras:", self.spin_letras)

        grupo_sangria = QGroupBox("Sangría")
        form_s = QFormLayout(grupo_sangria)
        form_s.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form_s.setHorizontalSpacing(12)
        form_s.setVerticalSpacing(8)
        self.spin_izquierda = self._spin(0, SANGRIA_MAX, SANGRIA_PASO, self.cambiar_izquierda)
        form_s.addRow("Izquierda:", self.spin_izquierda)
        self.spin_derecha = self._spin(0, SANGRIA_MAX, SANGRIA_PASO, self.cambiar_derecha)
        form_s.addRow("Derecha:", self.spin_derecha)
        self.spin_primera = self._spin(0, PRIMERA_LINEA_MAX, 6, self.cambiar_primera)
        form_s.addRow("1.ª línea:", self.spin_primera)

        cerrar = QPushButton("Cerrar")
        cerrar.clicked.connect(self.close)
        botones = QHBoxLayout()
        botones.addStretch()
        botones.addWidget(cerrar)

        lay = QVBoxLayout(self)
        lay.addWidget(grupo)
        lay.addWidget(grupo_sangria)
        lay.addLayout(botones)

        editor.cursorPositionChanged.connect(self.sincronizar)
        editor.selectionChanged.connect(self.sincronizar)
        editor.formatoParrafoCambiado.connect(self.sincronizar)
        editor.currentCharFormatChanged.connect(self.sincronizar_letras)
        self.sincronizar()
        self.sincronizar_letras(editor.currentCharFormat())

    def _spin(self, minimo, maximo, paso, slot):
        s = QSpinBox()
        s.setRange(minimo, maximo)
        s.setSingleStep(paso)
        s.setSuffix(" px")
        s.setFixedWidth(90)
        s.setAccelerated(True)
        s.setKeyboardTracking(False)  # aplica al pulsar Enter o al salir
        s.valueChanged.connect(slot)
        return s

    # -- acciones -----------------------------------------------------------
    def _aplicar(self, fmt):
        self.editor.combinar_parrafo(fmt)

    def cambiar_interlineado(self, texto):
        try:
            valor = float(texto.replace(",", "."))
        except ValueError:
            return
        valor = max(0.5, min(5.0, valor))
        fmt = QTextBlockFormat()
        fmt.setLineHeight(round(valor * 100),
                          QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
        self._aplicar(fmt)

    def cambiar_antes(self, v):
        fmt = QTextBlockFormat()
        fmt.setTopMargin(v)
        self._aplicar(fmt)

    def cambiar_despues(self, v):
        fmt = QTextBlockFormat()
        fmt.setBottomMargin(v)
        self._aplicar(fmt)

    def cambiar_letras(self, v):
        """Espaciado entre letras (formato de carácter, en px absolutos)."""
        fmt = QTextCharFormat()
        aplicar_letras(fmt, v)
        self.editor.combinar_formato(fmt)

    def cambiar_izquierda(self, v):
        fmt = QTextBlockFormat()
        fmt.setLeftMargin(v)
        self._aplicar(fmt)

    def cambiar_derecha(self, v):
        fmt = QTextBlockFormat()
        fmt.setRightMargin(v)
        self._aplicar(fmt)

    def cambiar_primera(self, v):
        fmt = QTextBlockFormat()
        fmt.setTextIndent(v)
        self._aplicar(fmt)

    # -- estado -------------------------------------------------------------
    def sincronizar(self):
        fmt = self.editor.textCursor().blockFormat()
        if fmt.lineHeightType() == \
                QTextBlockFormat.LineHeightTypes.ProportionalHeight.value:
            inter = fmt.lineHeight() / 100
        else:
            inter = 1.0
        self.combo_interlineado.blockSignals(True)
        texto = f"{inter:.2f}".rstrip("0")
        self.combo_interlineado.setCurrentText(texto + "0" if texto.endswith(".") else texto)
        self.combo_interlineado.blockSignals(False)

        for spin, valor in ((self.spin_antes, fmt.topMargin()),
                            (self.spin_despues, fmt.bottomMargin()),
                            (self.spin_izquierda, fmt.leftMargin()),
                            (self.spin_derecha, fmt.rightMargin()),
                            (self.spin_primera, fmt.textIndent())):
            spin.blockSignals(True)
            spin.setValue(round(valor))
            spin.blockSignals(False)

    def sincronizar_letras(self, fmt):
        if fmt.fontLetterSpacingType() == QFont.AbsoluteSpacing:
            valor = fmt.fontLetterSpacing()
        else:
            valor = 0  # espaciado porcentual por defecto (100 %) = sin cambio
        self.spin_letras.blockSignals(True)
        self.spin_letras.setValue(round(valor))
        self.spin_letras.blockSignals(False)


# ---------------------------------------------------------------------------
# Ventana
# ---------------------------------------------------------------------------
class VistaSinPaginas(QGraphicsView):
    """Muestra el editor como una sola hoja continua sobre fondo gris. La hoja
    tiene el ancho de la página (las líneas cortan igual que al imprimir),
    crece con el texto (al menos una página de alto) y siempre llega al borde
    inferior de la vista."""
    zoomCambiado = Signal(float)

    def __init__(self, editor, parent=None):
        super().__init__(parent)
        self.setObjectName("lienzo")
        self.setScene(QGraphicsScene(self))
        self.setBackgroundBrush(QColor(TEMA.t["lienzo"]))
        self.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.setFrameShape(QFrame.NoFrame)
        # Barras fijas para que la vista no cambie de tamaño al hacer zoom
        # La barra visible es externa (en el borde derecho de la ventana)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        # Solo suavizado de texto: la hoja es un rectángulo recto
        # y no lo necesitan (antialiasing general es caro a zoom alto)
        self.setRenderHints(QPainter.TextAntialiasing)
        # Repintado inteligente: al hacer scroll solo se pinta la franja nueva
        self.setViewportUpdateMode(QGraphicsView.SmartViewportUpdate)
        self.setOptimizationFlags(QGraphicsView.DontSavePainterState
                                  | QGraphicsView.DontAdjustForAntialiasing)
        self.setCacheMode(QGraphicsView.CacheBackground)

        self.editor = editor
        editor.vista = self
        editor.setFrameShape(QFrame.NoFrame)
        editor.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        # El contenido conserva sus colores con cualquier tema
        editor.setStyleSheet(
            f"QTextEdit {{ background: transparent; color: {COLOR_CONTENIDO};"
            f" selection-background-color: {SELECCION_CONTENIDO};"
            f" selection-color: {COLOR_CONTENIDO}; }}")
        editor.viewport().setAutoFillBackground(False)

        self.hoja = self.scene().addRect(QRectF(), QPen(QColor(TEMA.t["borde_hoja"]), 0),
                                         QColor(TEMA.color_hoja))
        TEMA.cambiado.connect(self._aplicar_tema)
        self.proxy = self.scene().addWidget(editor)
        self.proxy.setZValue(1)            # el texto sobre la hoja
        self.config = dict(PAGINA)
        self.zoom = 1.0
        self.setTransformationAnchor(QGraphicsView.NoAnchor)
        self._frac_x = 0.5          # empieza centrada a lo ancho
        self._frac_y = 0.0          # y al principio del documento
        self._aplicando = False
        self.verticalScrollBar().valueChanged.connect(self._guardar_ubicacion)
        self.horizontalScrollBar().valueChanged.connect(self._guardar_ubicacion)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)

        editor.document().documentLayout().documentSizeChanged.connect(
            self._actualizar_hoja)
        editor.cursorPositionChanged.connect(
            lambda: QTimer.singleShot(0, self._seguir_cursor))
        editor.textChanged.connect(
            lambda: QTimer.singleShot(0, self._seguir_cursor))
        self.aplicar_config(self.config)

    def cm_a_px(self, cm):
        return cm / 2.54 * self.editor.logicalDpiY()

    def _aplicar_tema(self):
        self.setBackgroundBrush(QColor(TEMA.t["lienzo"]))
        self.hoja.setPen(QPen(QColor(TEMA.t["borde_hoja"]), 0))
        self.hoja.setBrush(QColor(TEMA.color_hoja))
        self.resetCachedContent()          # el fondo está en caché
        self.viewport().update()

    def aplicar_config(self, config):
        """Márgenes de la página en el marco raíz; sin límite de alto."""
        self.config = dict(PAGINA)
        self.config.update(config)
        TEMA.aplicar(resolver_tema(self.config),
                     self.config.get("color_hoja", COLOR_HOJA))
        doc = self.editor.document()
        doc.setUndoRedoEnabled(False)       # no es un cambio del usuario
        ff = doc.rootFrame().frameFormat()
        ff.setTopMargin(self.cm_a_px(config["superior"]) + SEPARACION / 2)
        ff.setBottomMargin(self.cm_a_px(config["inferior"]) + SEPARACION / 2)
        ff.setLeftMargin(self.cm_a_px(config["izquierdo"]))
        ff.setRightMargin(self.cm_a_px(config["derecho"]))
        doc.rootFrame().setFrameFormat(ff)
        doc.setUndoRedoEnabled(True)
        self._fijar_ancho()
        self._actualizar_hoja()

    def _fijar_ancho(self):
        ancho = self.cm_a_px(self.config["ancho"])
        # Alto -1 = sin límite: Qt no corta páginas
        self.editor.document().setPageSize(QSizeF(ancho, -1))

    def _actualizar_hoja(self, *_):
        if self.editor.document().pageSize().height() != -1:
            self._fijar_ancho()       # QTextEdit lo restablece al redimensionarse
        ancho = self.cm_a_px(self.config["ancho"])
        alto = self.cm_a_px(self.config["alto"])
        g = SEPARACION
        contenido = self.editor.document().documentLayout().documentSize().height()
        total = round(max(contenido, alto + g))
        visible = (self.viewport().height() - 1) / self.zoom
        # Inicio de la escena: ESPACIO_SUPERIOR px de pantalla sobre la hoja
        espacio = self.config.get("espacio_superior", ESPACIO_SUPERIOR)
        arriba = g / 2 - espacio / self.zoom
        fondo = max(total - g / 2, visible + arriba)       # donde acaba la escena
        # La hoja pasa un poco del final de la escena: su borde inferior no se ve
        alto_hoja = fondo - g / 2 + 4 * g
        self.hoja.setRect(QRectF(0, g / 2, ancho, alto_hoja))
        if self.editor.width() != round(ancho) or self.editor.height() != total:
            self.editor.setFixedSize(round(ancho), total)
        self.scene().setSceneRect(QRectF(-g, arriba, ancho + 2 * g, fondo - arriba))

    def _seguir_cursor(self):
        """Desplaza la vista para que el cursor esté visible."""
        if not QApplication.mouseButtons() & Qt.LeftButton:
            r = QRectF(self.editor.cursorRect())
            self.ensureVisible(self.proxy.mapRectToScene(r), 40, 60)

    def resizeEvent(self, e):
        self._aplicando = True     # al cambiar de tamaño se conserva la ubicación
        super().resizeEvent(e)
        self._actualizar_hoja()    # la hoja llega al borde inferior de la vista
        self._restaurar_ubicacion()
        self._aplicando = False
        # Si la ventana se achica, el zoom no puede pasar del nuevo ancho
        if self.zoom > self.zoom_max() + 1e-4:
            self.poner_zoom(self.zoom_max())
        else:
            self.zoomCambiado.emit(self.zoom)   # actualiza el límite en la barra

    # -- zoom ---------------------------------------------------------------
    def poner_zoom(self, z, ancla=None):
        """Aplica el zoom de inmediato, en un solo repintado. No toma como
        referencia el puntero ni el centro: la vista conserva la misma
        posición relativa del scroll (si estabas al 40 % del documento,
        sigues al 40 %). `ancla` se acepta solo por compatibilidad."""
        z = max(ZOOM_MIN, min(self.zoom_max(), z))
        if abs(z - self.zoom) < 1e-5:
            return
        vp = self.viewport()
        vp.setUpdatesEnabled(False)          # evita cuadros intermedios
        self._aplicando = True
        self.zoom = z
        self.setTransform(QTransform.fromScale(z, z))
        self._actualizar_hoja()
        self._restaurar_ubicacion()
        self._aplicando = False
        vp.setUpdatesEnabled(True)
        vp.update()
        self.zoomCambiado.emit(z)

    # -- ubicación guardada -------------------------------------------------
    def _guardar_ubicacion(self, *_):
        """Recuerda la posición relativa (0–1) de cada barra, aunque esté
        oculta. Si en este momento no hay recorrido (la hoja cabe), se
        conserva el último valor conocido."""
        if getattr(self, "_aplicando", False):
            return
        for barra, clave in ((self.verticalScrollBar(), "_frac_y"),
                             (self.horizontalScrollBar(), "_frac_x")):
            rango = barra.maximum() - barra.minimum()
            if rango > 0:
                setattr(self, clave, (barra.value() - barra.minimum()) / rango)

    def _restaurar_ubicacion(self):
        for barra, frac in ((self.verticalScrollBar(), self._frac_y),
                            (self.horizontalScrollBar(), self._frac_x)):
            rango = barra.maximum() - barra.minimum()
            barra.setValue(barra.minimum() + round(frac * rango))

    def _centrar_horizontal(self):
        """Compatibilidad: ahora se restaura la ubicación guardada."""
        self._aplicando = True
        self._restaurar_ubicacion()
        self._aplicando = False

    def paso_zoom(self, pasos):
        """Sube o baja de 10 en 10 (100 → 110 → 120…). Si el zoom no es
        múltiplo de 10 (p. ej. tras "Ajustar al ancho"), primero se ajusta
        al múltiplo siguiente en esa dirección."""
        pct = self.zoom * 100
        if pasos > 0:
            base = math.floor(pct / ZOOM_PASO + 1e-6) * ZOOM_PASO
        else:
            base = math.ceil(pct / ZOOM_PASO - 1e-6) * ZOOM_PASO
        self.poner_zoom((base + pasos * ZOOM_PASO) / 100)

    def zoom_max(self):
        """Zoom máximo: el múltiplo de 10 % más grande en que la hoja cabe a
        lo ancho de la vista."""
        ancho = self.cm_a_px(self.config["ancho"]) + 2 * SEPARACION
        z = (self.viewport().width() - 1) / ancho
        # Redondeado hacia abajo a múltiplo de 10 % (así la hoja siempre cabe)
        z = math.floor(z * 100 / ZOOM_PASO + 1e-6) * ZOOM_PASO / 100
        return max(ZOOM_MIN, min(ZOOM_MAX, z))

    def ajustar_ancho(self):
        """La hoja ocupa todo el ancho de la vista, con el mismo margen gris
        a cada lado y centrada. La escena mide ancho + 2·SEPARACION (el
        margen también se escala), así que se ajusta esa medida completa."""
        self._frac_x = 0.5
        self.poner_zoom(self.zoom_max())
        self._centrar_horizontal()      # por si el zoom no cambió

    def wheelEvent(self, e):
        delta = e.angleDelta().y()
        if e.modifiers() & Qt.ControlModifier:     # Ctrl + rueda = zoom
            # Una muesca (120) = un paso de 10 %; el touchpad acumula
            self._resto_zoom = getattr(self, "_resto_zoom", 0) + delta
            pasos = int(self._resto_zoom / 120)
            if pasos:
                self._resto_zoom -= pasos * 120
                self.paso_zoom(pasos)
            e.accept()
            return
        if not e.pixelDelta().isNull() or not delta:
            QAbstractScrollArea.wheelEvent(self, e)   # touchpad: nativo
            return
        # Rueda del ratón: avance moderado por muesca
        barra = self.horizontalScrollBar() if e.modifiers() & Qt.ShiftModifier \
            else self.verticalScrollBar()
        barra.setValue(barra.value() - round(delta / 120 * SCROLL_MUESCA))
        e.accept()


# ---------------------------------------------------------------------------
# Barra de estado con el zoom
# ---------------------------------------------------------------------------
# Palabras: letras y números, con guiones o apóstrofos internos («e-mail»,
# «l'estel»); un número con decimales o miles («1,5», «12.000») es una
# palabra; las imágenes y los signos no cuentan
RE_PALABRA = re.compile(r"\d+(?:[.,]\d+)+|[^\W_]+(?:[-'’][^\W_]+)*")
LOCALE_MX = QLocale(QLocale.Spanish, QLocale.Mexico)


def contar_palabras(texto):
    return sum(1 for _ in RE_PALABRA.finditer(texto))


def texto_palabras(n):
    return f"{LOCALE_MX.toString(n)} {'palabra' if n == 1 else 'palabras'}"


class BarraEstado(QStatusBar):
    def __init__(self, vista, parent=None):
        super().__init__(parent)
        self.vista = vista
        # Izquierda: contador de palabras y, junto a él, los avisos
        # (Guardado…, etc.), que así no lo tapan
        self.contador = QLabel()
        self.contador.setContentsMargins(6, 0, 6, 0)
        self.mensaje = QLabel()
        self.mensaje.setContentsMargins(10, 0, 0, 0)
        self.mensaje.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        TEMA.estilo(self.mensaje, lambda t: f"color: {t['texto_sec']};")
        self.addWidget(self.contador)
        self.addWidget(self.mensaje, 1)
        self._t_mensaje = QTimer(self)
        self._t_mensaje.setSingleShot(True)
        self._t_mensaje.timeout.connect(self.clearMessage)
        # Recuento al dejar de escribir (en un libro largo no se nota)
        self._t_contar = QTimer(self)
        self._t_contar.setSingleShot(True)
        self._t_contar.setInterval(250)
        self._t_contar.timeout.connect(self.actualizar_contador)
        editor = vista.editor
        editor.document().contentsChanged.connect(self._t_contar.start)
        editor.selectionChanged.connect(self._t_contar.start)
        self.actualizar_contador()
        self.btn_menos = QToolButton()
        self.btn_menos.setText("−")
        self.btn_menos.setAutoRepeat(True)
        self.btn_menos.clicked.connect(lambda: vista.paso_zoom(-1))
        self.deslizador = QSlider(Qt.Horizontal)
        # Escala logarítmica: 100 % queda en el centro
        self.deslizador.setRange(0, 1000)
        self.deslizador.setFixedWidth(160)
        self.deslizador.setPageStep(50)
        self.deslizador.valueChanged.connect(
            lambda v: vista.poner_zoom(self._a_zoom(v)))
        self.btn_mas = QToolButton()
        self.btn_mas.setText("+")
        self.btn_mas.setAutoRepeat(True)
        self.btn_mas.clicked.connect(lambda: vista.paso_zoom(+1))
        self.btn_zoom = QToolButton()
        self.btn_zoom.setPopupMode(QToolButton.InstantPopup)
        self.btn_zoom.setMinimumWidth(56)
        menu = QMenu(self.btn_zoom)
        for z in MENU_ZOOM:
            menu.addAction(f"{z} %").triggered.connect(
                lambda _=False, z=z: vista.poner_zoom(z / 100))
        menu.addSeparator()
        menu.addAction("Ajustar al ancho").triggered.connect(vista.ajustar_ancho)
        self.btn_zoom.setMenu(menu)
        for w in (self.btn_menos, self.deslizador, self.btn_mas, self.btn_zoom):
            w.setFocusPolicy(Qt.NoFocus)    # el foco se queda en el texto
            self.addPermanentWidget(w)
        vista.zoomCambiado.connect(self._mostrar)
        self._mostrar(vista.zoom)

    def actualizar_contador(self):
        """«1,234 palabras»; con texto seleccionado, «56 de 1,234 palabras»."""
        editor = self.vista.editor
        total = contar_palabras(editor.document().toPlainText())
        sel = editor.textCursor()
        if sel.hasSelection():
            n = contar_palabras(sel.selection().toPlainText())
            self.contador.setText(f"{LOCALE_MX.toString(n)} de {texto_palabras(total)}")
        else:
            self.contador.setText(texto_palabras(total))
        # La barra de estado no ensancha sola la etiqueta al crecer el texto
        self.contador.setMinimumWidth(self.contador.sizeHint().width())

    # Avisos en su propia etiqueta: el contador sigue visible
    def showMessage(self, texto, tiempo=0):
        self.mensaje.setText(texto)
        if tiempo > 0:
            self._t_mensaje.start(tiempo)
        else:
            self._t_mensaje.stop()

    def clearMessage(self):
        self.mensaje.clear()
        self._t_mensaje.stop()

    def currentMessage(self):
        return self.mensaje.text()

    def _limites(self):
        a = math.log(ZOOM_MIN)
        return a, max(a + 1e-6, math.log(self.vista.zoom_max()))

    def _a_zoom(self, v):
        a, b = self._limites()
        return math.exp(a + (b - a) * v / 1000)

    def _a_valor(self, z):
        a, b = self._limites()
        return round(max(0.0, min(1.0, (math.log(z) - a) / (b - a))) * 1000)

    def _mostrar(self, z):
        self.btn_zoom.setText(f"{round(z * 100)} %")
        if not self.deslizador.isSliderDown():
            self.deslizador.blockSignals(True)
            self.deslizador.setValue(self._a_valor(z))
            self.deslizador.blockSignals(False)
        self.btn_menos.setEnabled(z > ZOOM_MIN + 1e-4)
        self.btn_mas.setEnabled(z < self.vista.zoom_max() - 1e-4)


# ---------------------------------------------------------------------------
# Ventana "Ajustes de estilo" (galería: personalizar, ordenar, ocultar,
# crear estilos, jerarquía y estilo siguiente, con vista previa)
# ---------------------------------------------------------------------------
ALINEACIONES = [("Izquierda", Qt.AlignLeft), ("Centro", Qt.AlignHCenter),
                ("Derecha", Qt.AlignRight), ("Justificado", Qt.AlignJustify)]
NIVELES = ["Texto (sin nivel)"] + [f"Nivel {n}" for n in range(1, 7)]
LOREM = ("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do "
         "eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim "
         "ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut "
         "aliquip ex ea commodo consequat.")
LOREM_CORTO = "Lorem ipsum dolor sit amet"
COLOR_CONTEXTO = "#A0A0A0"     # párrafos de alrededor en la vista previa


class VentanaEstilos(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ajustes de estilo")
        self.setWindowIcon(icono_ajustes_estilo())
        self.resize(900, 640)
        self.trabajo = GESTOR.copia()      # se aplica solo al aceptar
        self.actual = None
        self._cargando = False

        # -- izquierda: lista de la galería ----------------------------------
        self.lista = QListWidget()
        self.lista.setDragDropMode(QAbstractItemView.InternalMove)
        self.lista.setDefaultDropAction(Qt.MoveAction)
        self.lista.setMinimumWidth(220)
        self.lista.currentItemChanged.connect(self._seleccionar)
        self.lista.itemChanged.connect(self._marcar_visible)
        self.lista.model().rowsMoved.connect(self._reordenar)

        self.btn_nuevo = QPushButton("Nuevo")
        self.btn_duplicar = QPushButton("Duplicar")
        self.btn_eliminar = QPushButton("Eliminar")
        self.btn_subir = QPushButton("▲")
        self.btn_bajar = QPushButton("▼")
        self.btn_restablecer = QPushButton("Restablecer")
        for b in (self.btn_subir, self.btn_bajar):
            b.setFixedWidth(36)
        self.btn_nuevo.clicked.connect(self._nuevo)
        self.btn_duplicar.clicked.connect(self._duplicar)
        self.btn_eliminar.clicked.connect(self._eliminar)
        self.btn_subir.clicked.connect(lambda: self._mover(-1))
        self.btn_bajar.clicked.connect(lambda: self._mover(+1))
        self.btn_restablecer.clicked.connect(self._restablecer)

        fila1 = QHBoxLayout()
        for b in (self.btn_nuevo, self.btn_duplicar, self.btn_eliminar):
            fila1.addWidget(b)
        fila2 = QHBoxLayout()
        fila2.addWidget(self.btn_subir)
        fila2.addWidget(self.btn_bajar)
        fila2.addStretch()
        fila2.addWidget(self.btn_restablecer)
        nota = QLabel("Marca la casilla para mostrar el estilo en la cinta. "
                      "Arrastra o usa ▲ ▼ para cambiar el orden.")
        nota.setWordWrap(True)
        nota.setStyleSheet(f"color:{TEMA.t['texto_sec']};")
        izq = QVBoxLayout()
        izq.addWidget(QLabel("Estilos"))
        izq.addWidget(self.lista, 1)
        izq.addLayout(fila1)
        izq.addLayout(fila2)
        izq.addWidget(nota)

        # -- derecha: propiedades ---------------------------------------------
        g_general = QGroupBox("General")
        fg = QFormLayout(g_general)
        self.ed_nombre = QLineEdit()
        self.ed_nombre.textEdited.connect(self._cambiar_nombre)
        fg.addRow("Nombre en la cinta:", self.ed_nombre)
        self.combo_siguiente = QComboBox()
        self.combo_siguiente.activated.connect(self._cambiar_siguiente)
        fg.addRow("Estilo siguiente:", self.combo_siguiente)
        self.combo_nivel = QComboBox()
        self.combo_nivel.addItems(NIVELES)
        self.combo_nivel.activated.connect(self._cambiar_nivel)
        fg.addRow("Nivel de esquema:", self.combo_nivel)

        g_fuente = QGroupBox("Fuente")
        ff = QFormLayout(g_fuente)
        self.combo_fuente = QFontComboBox()
        self.combo_fuente.currentFontChanged.connect(
            lambda f: self._valor("font_family", f.family()))
        ff.addRow("Tipo:", self.combo_fuente)
        self.spin_tamano = QDoubleSpinBox()
        self.spin_tamano.setRange(4, 200)
        self.spin_tamano.setDecimals(1)
        self.spin_tamano.setSuffix(" pt")
        self.spin_tamano.setLocale(QLocale.c())
        self.spin_tamano.valueChanged.connect(lambda v: self._valor("font_size", v))
        ff.addRow("Tamaño:", self.spin_tamano)
        self.btn_color = QPushButton()
        self.btn_color.setFixedWidth(90)
        self.btn_color.clicked.connect(self._elegir_color)
        ff.addRow("Color:", self.btn_color)
        estilos = QHBoxLayout()
        self.chk_negrita = QCheckBox("Negrita")
        self.chk_cursiva = QCheckBox("Cursiva")
        self.chk_negrita.toggled.connect(lambda v: self._valor("bold", v))
        self.chk_cursiva.toggled.connect(lambda v: self._valor("italic", v))
        estilos.addWidget(self.chk_negrita)
        estilos.addWidget(self.chk_cursiva)
        estilos.addStretch()
        ff.addRow("", estilos)

        g_parrafo = QGroupBox("Párrafo")
        fp = QFormLayout(g_parrafo)
        self.combo_alineacion = QComboBox()
        for nombre, _ in ALINEACIONES:
            self.combo_alineacion.addItem(nombre)
        self.combo_alineacion.activated.connect(
            lambda i: self._valor("alignment", ALINEACIONES[i][1]))
        fp.addRow("Alineación:", self.combo_alineacion)
        self.combo_interlineado = QComboBox()
        self.combo_interlineado.setEditable(True)
        self.combo_interlineado.setInsertPolicy(QComboBox.NoInsert)
        validador = QDoubleValidator(0.5, 5.0, 2, self)
        validador.setNotation(QDoubleValidator.StandardNotation)
        validador.setLocale(QLocale.c())
        self.combo_interlineado.lineEdit().setValidator(validador)
        self.combo_interlineado.addItems(INTERLINEADOS)
        self.combo_interlineado.textActivated.connect(self._cambiar_interlineado)
        fp.addRow("Interlineado:", self.combo_interlineado)
        self.spin_antes = self._spin(0, ESPACIO_MAX, "space_before")
        fp.addRow("Espacio antes:", self.spin_antes)
        self.spin_despues = self._spin(0, ESPACIO_MAX, "space_after")
        fp.addRow("Espacio después:", self.spin_despues)
        self.spin_izq = self._spin(0, SANGRIA_MAX, "indent_left")
        fp.addRow("Sangría izquierda:", self.spin_izq)
        self.spin_der = self._spin(0, SANGRIA_MAX, "indent_right")
        fp.addRow("Sangría derecha:", self.spin_der)
        self.spin_primera = self._spin(0, PRIMERA_LINEA_MAX, "first_line_indent")
        fp.addRow("Primera línea:", self.spin_primera)
        self.spin_letras = self._spin(LETRAS_MIN, LETRAS_MAX, "letter_spacing")
        fp.addRow("Entre letras:", self.spin_letras)


        self.vista_previa = QTextEdit()
        # Contenido: se ve como en la hoja, con cualquier tema
        self.vista_previa.setStyleSheet(
            f"QTextEdit {{ background: {TEMA.color_hoja}; color: {COLOR_CONTENIDO};"
            f" border: 1px solid {TEMA.t['campo_borde']}; }}")
        self.vista_previa.setReadOnly(True)
        self.vista_previa.setMinimumHeight(170)
        self.vista_previa.setFocusPolicy(Qt.NoFocus)

        columnas = QHBoxLayout()
        c1 = QVBoxLayout()
        c1.addWidget(g_general)
        c1.addWidget(g_fuente)
        c1.addStretch()
        c2 = QVBoxLayout()
        c2.addWidget(g_parrafo)
        c2.addStretch()
        columnas.addLayout(c1)
        columnas.addLayout(c2)
        der = QVBoxLayout()
        der.addLayout(columnas)
        der.addWidget(QLabel("Vista previa"))
        der.addWidget(self.vista_previa, 1)

        cuerpo = QHBoxLayout()
        cuerpo.addLayout(izq)
        cuerpo.addLayout(der, 1)

        botones = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        botones.button(QDialogButtonBox.Ok).setText("Aceptar")
        botones.button(QDialogButtonBox.Cancel).setText("Cancelar")
        botones.accepted.connect(self.accept)
        botones.rejected.connect(self.reject)

        lay = QVBoxLayout(self)
        lay.addLayout(cuerpo, 1)
        lay.addWidget(botones)

        self._llenar_lista()
        self.lista.setCurrentRow(0)

    def _spin(self, minimo, maximo, campo):
        s = QSpinBox()
        s.setRange(minimo, maximo)
        s.setSuffix(" px")
        s.valueChanged.connect(lambda v: self._valor(campo, v))
        return s

    # -- lista --------------------------------------------------------------
    def _llenar_lista(self, seleccionar=None):
        self._cargando = True
        self.lista.clear()
        for id_ in self.trabajo.orden:
            d = self.trabajo.datos[id_]
            it = QListWidgetItem(self.trabajo.nombre(id_))
            it.setData(Qt.UserRole, id_)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsDragEnabled)
            it.setCheckState(Qt.Checked if d["visible"] else Qt.Unchecked)
            it.setFont(fuente_muestra(self.trabajo.resuelto(id_)))
            if not d["fijo"]:
                it.setToolTip("Estilo personalizado")
            self.lista.addItem(it)
        self._cargando = False
        if seleccionar is not None:
            self.lista.setCurrentRow(self.trabajo.orden.index(seleccionar))

    def _item(self, id_):
        for i in range(self.lista.count()):
            if self.lista.item(i).data(Qt.UserRole) == id_:
                return self.lista.item(i)
        return None

    def _refrescar_item(self, id_):
        it = self._item(id_)
        if it is not None:
            self._cargando = True
            it.setText(self.trabajo.nombre(id_))
            it.setFont(fuente_muestra(self.trabajo.resuelto(id_)))
            self._cargando = False

    def _marcar_visible(self, it):
        if self._cargando:
            return
        id_ = it.data(Qt.UserRole)
        self.trabajo.datos[id_]["visible"] = it.checkState() == Qt.Checked

    def _reordenar(self, *_):
        self.trabajo.orden = [self.lista.item(i).data(Qt.UserRole)
                              for i in range(self.lista.count())]

    def _mover(self, paso):
        fila = self.lista.currentRow()
        nueva = fila + paso
        if fila < 0 or not 0 <= nueva < self.lista.count():
            return
        orden = self.trabajo.orden
        orden[fila], orden[nueva] = orden[nueva], orden[fila]
        self._llenar_lista(seleccionar=orden[nueva])

    def _nuevo(self):
        base = self.actual or "Body"
        id_ = self.trabajo.nuevo(base)
        self._llenar_lista(seleccionar=id_)
        self.ed_nombre.setFocus()
        self.ed_nombre.selectAll()

    def _duplicar(self):
        if not self.actual:
            return
        origen = self.trabajo.datos[self.actual]
        id_ = self.trabajo.nuevo(self.actual)
        d = self.trabajo.datos[id_]
        d["nombre"] = f"{self.trabajo.nombre(self.actual)} (copia)"
        d["nivel"] = origen["nivel"]
        d["siguiente"] = origen["siguiente"] if origen["siguiente"] != self.actual else id_
        self._llenar_lista(seleccionar=id_)

    def _eliminar(self):
        if not self.actual or self.trabajo.datos[self.actual]["fijo"]:
            return
        fila = self.lista.currentRow()
        self.trabajo.eliminar(self.actual)
        self.actual = None
        self._llenar_lista()
        self.lista.setCurrentRow(min(fila, self.lista.count() - 1))

    def _restablecer(self):
        if self.actual and self.trabajo.datos[self.actual]["fijo"]:
            self.trabajo.restablecer(self.actual)
            id_ = self.actual
            self._llenar_lista(seleccionar=id_)
            self._cargar(id_)

    # -- propiedades --------------------------------------------------------
    def _seleccionar(self, it, _anterior=None):
        if it is None:
            return
        self._cargar(it.data(Qt.UserRole))

    def _cargar(self, id_):
        self.actual = id_
        d = self.trabajo.datos[id_]
        e = self.trabajo.resuelto(id_)
        self._cargando = True
        widgets = (self.ed_nombre, self.combo_siguiente,
                   self.combo_nivel, self.combo_fuente, self.spin_tamano,
                   self.chk_negrita, self.chk_cursiva, self.spin_letras,
                   self.combo_alineacion, self.combo_interlineado,
                   self.spin_antes, self.spin_despues, self.spin_izq,
                   self.spin_der, self.spin_primera)
        for w in widgets:
            w.blockSignals(True)
        self.ed_nombre.setText(d["nombre"])
        self.ed_nombre.setPlaceholderText(id_)

        self.combo_siguiente.clear()
        for o in self.trabajo.orden:
            self.combo_siguiente.addItem(self.trabajo.nombre(o), o)
        self.combo_siguiente.setCurrentIndex(
            max(0, self.combo_siguiente.findData(d["siguiente"])))
        self.combo_nivel.setCurrentIndex(d["nivel"])

        self.combo_fuente.setCurrentFont(QFont(e["font_family"]))
        self.spin_tamano.setValue(e["font_size"])
        self.chk_negrita.setChecked(e["bold"])
        self.chk_cursiva.setChecked(e["italic"])
        self.spin_letras.setValue(round(e["letter_spacing"]))
        self.combo_alineacion.setCurrentIndex(next(
            (i for i, (_, a) in enumerate(ALINEACIONES) if a == e["alignment"]), 0))
        texto = f"{e['line_spacing'] / 100:.2f}".rstrip("0")
        self.combo_interlineado.setCurrentText(texto + "0" if texto.endswith(".") else texto)
        self.spin_antes.setValue(round(e["space_before"]))
        self.spin_despues.setValue(round(e["space_after"]))
        self.spin_izq.setValue(round(e["indent_left"]))
        self.spin_der.setValue(round(e["indent_right"]))
        self.spin_primera.setValue(round(e["first_line_indent"]))
        for w in widgets:
            w.blockSignals(False)
        self._pintar_color(e["font_color"])
        self._cargando = False

        self.btn_eliminar.setEnabled(not d["fijo"])
        self.btn_restablecer.setEnabled(d["fijo"])
        self._vista_previa()

    def _valor(self, campo, valor):
        if self._cargando or not self.actual:
            return
        self.trabajo.datos[self.actual]["propio"][campo] = valor
        self._refrescar_item(self.actual)
        self._cargar(self.actual)

    def _cambiar_nombre(self, texto):
        if self.actual:
            self.trabajo.datos[self.actual]["nombre"] = texto.strip()
            self._refrescar_item(self.actual)
            self._vista_previa()

    def _cambiar_siguiente(self, i):
        if self.actual:
            self.trabajo.datos[self.actual]["siguiente"] = self.combo_siguiente.itemData(i)
            self._vista_previa()

    def _cambiar_nivel(self, i):
        if self.actual:
            self.trabajo.datos[self.actual]["nivel"] = i

    def _cambiar_interlineado(self, texto):
        try:
            v = float(texto.replace(",", "."))
        except ValueError:
            return
        self._valor("line_spacing", round(max(0.5, min(5.0, v)) * 100))

    def _elegir_color(self):
        if not self.actual:
            return
        actual = QColor(self.trabajo.resuelto(self.actual)["font_color"])
        c = QColorDialog.getColor(actual, self, "Color del texto")
        if c.isValid():
            self._valor("font_color", c.name().upper())

    def _pintar_color(self, hexa):
        self.btn_color.setText(hexa)
        claro = QColor(hexa).lightness() > 140
        self.btn_color.setStyleSheet(
            f"QPushButton {{ background:{hexa}; color:{'#1A1A1A' if claro else '#FFFFFF'};"
            " border:1px solid #9A9A9A; border-radius:3px; padding:3px; }")

    def _vista_previa(self):
        """Párrafo anterior (Body, en gris), el estilo elegido y el párrafo
        siguiente (su estilo siguiente, en gris): así se ven los espacios
        antes y después, sangrías y alineación en contexto."""
        if not self.actual:
            return
        g = self.trabajo
        doc = self.vista_previa.document()
        doc.clear()
        ff = doc.rootFrame().frameFormat()
        ff.setMargin(18)
        doc.rootFrame().setFrameFormat(ff)
        c = QTextCursor(doc)

        def contexto(id_):
            fmt = formato_caracter(id_, g)
            fmt.setForeground(QColor(COLOR_CONTEXTO))
            return fmt

        # Un título se ve con texto corto; el texto corrido con un párrafo
        e = g.resuelto(self.actual)
        corto = e["font_size"] >= 14 or g.datos[self.actual]["nivel"] > 0
        c.setBlockFormat(formato_parrafo("Body", g))
        c.insertText(LOREM, contexto("Body"))
        c.insertBlock(formato_parrafo(self.actual, g))
        c.insertText(LOREM_CORTO if corto else LOREM,
                     formato_caracter(self.actual, g))
        sig = g.datos[self.actual]["siguiente"]
        c.insertBlock(formato_parrafo(sig, g))
        c.insertText(LOREM, contexto(sig))
        self.vista_previa.verticalScrollBar().setValue(0)


# ---------------------------------------------------------------------------
# Ventana "Configurar página" (tamaño, orientación y márgenes)
# ---------------------------------------------------------------------------
class BotonElegirColor(QPushButton):
    """Botón con una muestra del color y su código; al pulsarlo abre el
    selector de color."""
    colorCambiado = Signal(str)

    def __init__(self, color, titulo, parent=None):
        super().__init__(parent)
        self._titulo = titulo
        self.setMinimumWidth(130)
        self.clicked.connect(self._elegir)
        self.poner(color)

    def color(self):
        return self._color

    def poner(self, color):
        self._color = QColor(color).name().upper()
        pm = QPixmap(30, 16)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setPen(QPen(QColor("#808080"), 1))
        p.setBrush(QColor(self._color))
        p.drawRoundedRect(QRectF(0.5, 0.5, 29, 15), 3, 3)
        p.end()
        self.setIcon(QIcon(pm))
        self.setIconSize(QSize(30, 16))
        self.setText(self._color)

    def _elegir(self):
        c = QColorDialog.getColor(QColor(self._color), self.window(), self._titulo)
        if c.isValid():
            self.poner(c.name())
            self.colorCambiado.emit(self._color)


class VentanaPagina(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configurar página")
        self.setWindowIcon(icono_pagina())
        lay = QVBoxLayout(self)

        grupo_papel = QGroupBox("Papel")
        form = QFormLayout(grupo_papel)
        self.combo_tamano = QComboBox()
        for nombre, an, al in TAMANOS_PAGINA:
            self.combo_tamano.addItem(f"{nombre}  ({an:g} × {al:g} cm)", (an, al))
        self.combo_tamano.addItem("Personalizado", None)
        form.addRow("Tamaño:", self.combo_tamano)
        self.spin_ancho = self._spin_cm(5, 100)
        self.spin_alto = self._spin_cm(5, 100)
        form.addRow("Ancho:", self.spin_ancho)
        form.addRow("Alto:", self.spin_alto)
        self.combo_orientacion = QComboBox()
        self.combo_orientacion.addItems(["Vertical", "Horizontal"])
        form.addRow("Orientación:", self.combo_orientacion)
        lay.addWidget(grupo_papel)

        grupo_margenes = QGroupBox("Márgenes")
        fm = QFormLayout(grupo_margenes)
        self.spin_superior = self._spin_cm(0, 20)
        self.spin_inferior = self._spin_cm(0, 20)
        self.spin_izquierdo = self._spin_cm(0, 20)
        self.spin_derecho = self._spin_cm(0, 20)
        fm.addRow("Superior:", self.spin_superior)
        fm.addRow("Inferior:", self.spin_inferior)
        fm.addRow("Izquierdo:", self.spin_izquierdo)
        fm.addRow("Derecho:", self.spin_derecho)
        lay.addWidget(grupo_margenes)

        grupo_vista = QGroupBox("Vista")
        fv = QFormLayout(grupo_vista)
        self.combo_tema = QComboBox()
        for clave, nombre in NOMBRES_TEMA:
            self.combo_tema.addItem(nombre, clave)
        fv.addRow("Tema:", self.combo_tema)
        self.btn_interfaz = BotonElegirColor(
            config.get("tema_interfaz", TEMAS["claro"]["barra"]), "Color de la interfaz")
        self.btn_fondo = BotonElegirColor(
            config.get("tema_fondo", TEMAS["claro"]["lienzo"]), "Color del fondo")
        self.btn_hoja = BotonElegirColor(
            config.get("color_hoja", COLOR_HOJA), "Color de la hoja")
        fv.addRow("Color de la interfaz:", self.btn_interfaz)
        fv.addRow("Color del fondo:", self.btn_fondo)
        fv.addRow("Color de la hoja:", self.btn_hoja)
        self.spin_espacio_superior = QSpinBox()
        self.spin_espacio_superior.setRange(0, 200)
        self.spin_espacio_superior.setSingleStep(2)
        self.spin_espacio_superior.setSuffix(" px")
        fv.addRow("Separación superior:", self.spin_espacio_superior)
        lay.addWidget(grupo_vista)

        botones = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        botones.button(QDialogButtonBox.Ok).setText("Aceptar")
        botones.button(QDialogButtonBox.Cancel).setText("Cancelar")
        botones.accepted.connect(self._aceptar)
        botones.rejected.connect(self.reject)
        lay.addWidget(botones)

        self._cargando = True
        for spin, k in ((self.spin_ancho, "ancho"), (self.spin_alto, "alto"),
                        (self.spin_superior, "superior"), (self.spin_inferior, "inferior"),
                        (self.spin_izquierdo, "izquierdo"), (self.spin_derecho, "derecho")):
            spin.setValue(config[k])
        self.spin_espacio_superior.setValue(
            round(config.get("espacio_superior", ESPACIO_SUPERIOR)))
        i = self.combo_tema.findData(config.get("tema", "claro"))
        self.combo_tema.setCurrentIndex(max(0, i))
        self._cargando = False
        self._tema_elegido()
        self.combo_tema.currentIndexChanged.connect(self._tema_elegido)
        self._sincronizar_tamano()
        self.combo_tamano.currentIndexChanged.connect(self._elegir_tamano)
        self.combo_orientacion.currentIndexChanged.connect(self._girar)
        self.spin_ancho.valueChanged.connect(self._sincronizar_tamano)
        self.spin_alto.valueChanged.connect(self._sincronizar_tamano)

    def _tema_elegido(self, *_):
        """Claro y Oscuro muestran sus colores (fijos); Personalizado deja
        elegirlos, empezando por los del tema que se veía."""
        clave = self.combo_tema.currentData()
        personal = clave == "personalizado"
        self.btn_interfaz.setEnabled(personal)
        self.btn_fondo.setEnabled(personal)
        if not personal:
            t = TEMAS[clave]
            self.btn_interfaz.poner(t["barra"])
            self.btn_fondo.poner(t["lienzo"])

    def _spin_cm(self, minimo, maximo):
        s = QDoubleSpinBox()
        s.setRange(minimo, maximo)
        s.setDecimals(2)
        s.setSingleStep(0.1)
        s.setSuffix(" cm")
        s.setLocale(QLocale.c())  # siempre punto decimal
        return s

    def _elegir_tamano(self, i):
        medidas = self.combo_tamano.itemData(i)
        if medidas is None or self._cargando:
            return
        an, al = medidas
        if self.combo_orientacion.currentIndex() == 1:
            an, al = al, an
        self._cargando = True
        self.spin_ancho.setValue(an)
        self.spin_alto.setValue(al)
        self._cargando = False

    def _girar(self, i):
        an, al = self.spin_ancho.value(), self.spin_alto.value()
        if (i == 1) != (an > al):          # intercambia ancho y alto
            self._cargando = True
            self.spin_ancho.setValue(al)
            self.spin_alto.setValue(an)
            self._cargando = False

    def _sincronizar_tamano(self, *_):
        """Muestra el tamaño y la orientación que coinciden con las medidas."""
        if self._cargando:
            return
        an, al = self.spin_ancho.value(), self.spin_alto.value()
        corto, largo = sorted((an, al))
        indice = self.combo_tamano.count() - 1   # Personalizado
        for i, (_, pa, pl) in enumerate(TAMANOS_PAGINA):
            if abs(pa - corto) < 0.01 and abs(pl - largo) < 0.01:
                indice = i
                break
        self._cargando = True
        self.combo_tamano.setCurrentIndex(indice)
        self.combo_orientacion.setCurrentIndex(1 if an > al else 0)
        self._cargando = False

    def _aceptar(self):
        c = self.config()
        if c["izquierdo"] + c["derecho"] > c["ancho"] - 2 or \
                c["superior"] + c["inferior"] > c["alto"] - 2:
            QMessageBox.warning(
                self, "Configurar página",
                "Los márgenes son demasiado grandes para el tamaño de página. "
                "Deben dejar al menos 2 cm de espacio para el texto.")
            return
        self.accept()

    def config(self):
        return {
            "ancho": self.spin_ancho.value(), "alto": self.spin_alto.value(),
            "superior": self.spin_superior.value(), "inferior": self.spin_inferior.value(),
            "izquierdo": self.spin_izquierdo.value(), "derecho": self.spin_derecho.value(),
            "espacio_superior": self.spin_espacio_superior.value(),
            "tema": self.combo_tema.currentData(),
            "tema_interfaz": self.btn_interfaz.color(),
            "tema_fondo": self.btn_fondo.color(),
            "color_hoja": self.btn_hoja.color(),
        }


# ===========================================================================
# Herramientas tomadas de SWordBook: Nuevo, Abrir, Guardar, Guardar como,
# Buscar, Navegador, Comentarios y Cortar/Copiar/Pegar
# ===========================================================================
PROP_COMENTARIO = QTextFormat.UserProperty + 3   # "3 7": comentarios del texto
PROP_LISTA = QTextFormat.UserProperty + 8        # al abrir: "tipo;numId;nivel"
MIME_LIBRO = "application/x-editorlibro"
# Al copiar: párrafos y tramos con el formato puesto a mano (sin el que
# viene del estilo), para pegar en Body conservando solo ese formato
MIME_TRAMOS = "application/x-swordbook-tramos"
CARACTER_OBJETO = "\ufffc"
COLORES_COMENTARIO = [
    "#E6DCF7", "#FFF59D", "#C5E1A5", "#B3E5FC", "#F8BBD0",
    "#FFCC80", "#B2DFDB", "#D7CCC8", "#E0E0E0", "#FFAB91",
]


def ids_comentario(cf):
    if not cf.hasProperty(PROP_COMENTARIO):
        return set()
    ids = set()
    for x in cf.stringProperty(PROP_COMENTARIO).split():
        try:
            ids.add(int(x))
        except ValueError:
            pass
    return ids


def poner_ids_comentario(cf, ids):
    if ids:
        cf.setProperty(PROP_COMENTARIO, " ".join(str(i) for i in sorted(ids)))
    else:
        cf.clearProperty(PROP_COMENTARIO)


def ahora_iso():
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def fecha_legible(iso):
    try:
        d = datetime.datetime.fromisoformat(iso.replace("Z", "+00:00"))
        if d.tzinfo is None:
            d = d.replace(tzinfo=datetime.timezone.utc)
        return d.astimezone().strftime("%d/%m/%Y %H:%M")
    except (ValueError, AttributeError):
        return iso or ""


# ---------------------------------------------------------------------------
# Comentarios: el texto comentado lleva PROP_COMENTARIO en su formato; autor,
# fecha y texto del comentario se guardan aquí (deshacer funciona solo)
# ---------------------------------------------------------------------------
class GestorComentarios(QObject):
    listaCambiada = Signal()
    activoCambiado = Signal(int)
    cambiado = Signal()

    def __init__(self, editor):
        super().__init__(editor)
        self.editor = editor
        self.datos = {}          # id → {"autor", "fecha", "texto", "color"}
        self.rangos = {}         # id → QTextCursor con el texto comentado
        self.orden = []
        self.activo = -1
        self._siguiente = 1
        try:
            self.autor = getpass.getuser()
        except Exception:
            self.autor = "Autor"
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(200)
        self._timer.timeout.connect(self.escanear)
        editor.document().contentsChange.connect(self._programar)
        editor.cursorPositionChanged.connect(self._cursor_movido)

    def _programar(self, *_):
        if self.datos:
            self._timer.start()

    def _fragmentos(self, inicio=0, fin=None):
        doc = self.editor.document()
        if fin is None:
            fin = doc.characterCount()
        b = doc.findBlock(inicio)
        while b.isValid() and b.position() < fin:
            it = b.begin()
            while not it.atEnd():
                f = it.fragment()
                if f.isValid():
                    a = max(inicio, f.position())
                    z = min(fin, f.position() + f.length())
                    if a < z:
                        yield a, z, f.charFormat()
                it += 1
            b = b.next()

    def escanear(self):
        self._timer.stop()
        limites = {}
        if self.datos:
            for a, z, cf in self._fragmentos():
                for i in ids_comentario(cf):
                    if i in self.datos:
                        r = limites.get(i)
                        limites[i] = (a, z) if r is None else (min(r[0], a), max(r[1], z))
        doc = self.editor.document()
        self.rangos = {}
        for i, (a, z) in limites.items():
            c = QTextCursor(doc)
            c.setPosition(a)
            c.setPosition(z, QTextCursor.KeepAnchor)
            self.rangos[i] = c
        self.orden = sorted(limites, key=lambda i: (limites[i][0], i))
        self.activo = -2
        self._cursor_movido()
        self.listaCambiada.emit()

    def extracto(self, i, largo=500):
        r = self.rangos.get(i)
        if r is None:
            return ""
        t = " ".join(r.selectedText().replace("\u2029", " ").replace("\u2028", " ")
                     .replace(CARACTER_OBJETO, " ").split())
        return t if len(t) <= largo else t[:largo - 1].rstrip() + "…"

    def _cursor_movido(self):
        c = self.editor.textCursor()
        pos = c.position()
        dentro = [i for i, r in self.rangos.items()
                  if r.hasSelection() and r.selectionStart() <= pos <= r.selectionEnd()]
        nuevo = min(dentro, key=lambda i: (self.rangos[i].selectionEnd()
                                           - self.rangos[i].selectionStart(), i)) \
            if dentro else -1
        if nuevo != self.activo:
            self.activo = nuevo
            self.editor.refrescar_selecciones()
            self.activoCambiado.emit(nuevo)
        self._no_extender(c)

    def _no_extender(self, c):
        """Lo que se escribe justo al final de un texto comentado no entra
        en el comentario."""
        if c.hasSelection() or not self.datos:
            return
        cf = self.editor.currentCharFormat()
        antes = ids_comentario(cf)
        if not antes:
            return
        despues = set()
        b = c.block()
        if c.position() < b.position() + b.length() - 1:
            d = QTextCursor(c)
            d.movePosition(QTextCursor.NextCharacter)
            despues = ids_comentario(d.charFormat())
        quedan = antes & despues
        if quedan != antes:
            poner_ids_comentario(cf, quedan)
            self.editor.setCurrentCharFormat(cf)

    def selecciones(self):
        res = []
        for i in self.orden:
            r = self.rangos.get(i)
            if r is None or not r.hasSelection():
                continue
            sel = QTextEdit.ExtraSelection()
            sel.cursor = QTextCursor(r)
            # Mismo color siempre, esté o no el cursor dentro
            sel.format.setBackground(QColor(self.color(i)))
            res.append(sel)
        return res

    def _cambiar_marcas(self, inicio, fin, func):
        cambios = list(self._fragmentos(inicio, fin))
        c = QTextCursor(self.editor.document())
        c.beginEditBlock()
        for a, z, cf in cambios:
            antes = ids_comentario(cf)
            ids = func(antes)
            if ids == antes:
                continue
            poner_ids_comentario(cf, ids)
            c.setPosition(a)
            c.setPosition(z, QTextCursor.KeepAnchor)
            c.setCharFormat(cf)
        c.endEditBlock()

    def color(self, i):
        c = self.datos.get(i, {}).get("color")
        return c if c and QColor(c).isValid() else \
            COLORES_COMENTARIO[(i - 1) % len(COLORES_COMENTARIO)]

    def poner_color(self, i, color):
        if i in self.datos and QColor(color).isValid():
            self.datos[i]["color"] = QColor(color).name().upper()
            self.cambiado.emit()
            self.editor.document().setModified(True)
            self.editor.refrescar_selecciones()
            self.listaCambiada.emit()

    def agregar(self):
        c = self.editor.textCursor()
        if not c.hasSelection():
            c.select(QTextCursor.WordUnderCursor)
        if not c.hasSelection() or not c.selectedText().strip():
            return None
        i = self._siguiente
        self._siguiente += 1
        self.datos[i] = {"autor": self.autor, "fecha": ahora_iso(), "texto": "",
                         "color": COLORES_COMENTARIO[(i - 1) % len(COLORES_COMENTARIO)]}
        self._cambiar_marcas(c.selectionStart(), c.selectionEnd(), lambda s: s | {i})
        self.escanear()
        return i

    def eliminar(self, i):
        r = self.rangos.get(i)
        if r is not None:
            self._cambiar_marcas(r.selectionStart(), r.selectionEnd(), lambda s: s - {i})
        self.escanear()

    def poner_texto(self, i, texto):
        if i in self.datos and self.datos[i]["texto"] != texto:
            self.datos[i]["texto"] = texto
            self.cambiado.emit()
            self.editor.document().setModified(True)

    def ir_a(self, i):
        r = self.rangos.get(i)
        if r is None:
            return
        c = QTextCursor(r)
        c.setPosition(r.selectionStart())
        self.editor.setTextCursor(c)
        vista = self.editor.vista
        if vista is None:
            return
        QApplication.processEvents()
        punto = vista.proxy.mapToScene(QRectF(self.editor.cursorRect()).center())
        alto = vista.viewport().height() / vista.zoom
        vista.centerOn(punto.x(), punto.y() + alto / 6)

    def marcas(self, inicio, fin):
        return [(a, z, cf.stringProperty(PROP_COMENTARIO))
                for a, z, cf in self._fragmentos(inicio, fin)
                if cf.hasProperty(PROP_COMENTARIO)]

    def restaurar(self, marcas):
        c = QTextCursor(self.editor.document())
        for a, z, valor in marcas:
            cf = QTextCharFormat()
            cf.setProperty(PROP_COMENTARIO, valor)
            c.setPosition(a)
            c.setPosition(z, QTextCursor.KeepAnchor)
            c.mergeCharFormat(cf)

    def cargar(self, datos):
        self.datos = {i: dict(d) for i, d in datos.items()}
        self._siguiente = max(self.datos, default=0) + 1
        self.escanear()

    def limpiar(self):
        self.cargar({})


# ---------------------------------------------------------------------------
# Archivos .docx (Word, Google Docs, LibreOffice, Amazon KDP)
# ---------------------------------------------------------------------------
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
TWIPS_POR_CM = 1440 / 2.54
FILTRO_DOCX = "Documento de Word (*.docx)"
# Copia de seguridad: mientras hay cambios sin guardar se escribe junto al
# documento; se borra al guardar o al cerrar con normalidad. Si el programa
# se cierra de forma inesperada (corte de luz…), al abrir el documento se
# ofrece recuperarla
SUFIJO_RESPALDO = " (copia de seguridad)"
RESPALDO_PAUSA_MS = 4000        # tras dejar de escribir
RESPALDO_MAX_MS = 30000         # como mucho, mientras se sigue escribiendo


def ruta_respaldo(ruta):
    """Copia de seguridad de un documento: en su misma carpeta, «nombre
    (copia de seguridad).docx». Sin documento (Sin título): en Documentos."""
    if ruta:
        carpeta = os.path.dirname(os.path.abspath(ruta))
        base = os.path.splitext(os.path.basename(ruta))[0]
    else:
        carpeta = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        if not carpeta or not os.path.isdir(carpeta):
            carpeta = os.path.expanduser("~")
        base = "Sin título"
    return os.path.join(carpeta, f"{base}{SUFIJO_RESPALDO}.docx")
FILTRO_ABRIR = "Documentos (*.docx *.txt);;Documento de Word (*.docx);;Texto (*.txt)"
RESALTADOS_DOCX = {
    "yellow": "FFFF00", "green": "00FF00", "cyan": "00FFFF",
    "magenta": "FF00FF", "blue": "0000FF", "red": "FF0000",
    "darkBlue": "000080", "darkCyan": "008080", "darkGreen": "008000",
    "darkMagenta": "800080", "darkRed": "800000", "darkYellow": "808000",
    "darkGray": "808080", "lightGray": "C0C0C0", "black": "000000",
    "white": "FFFFFF",
}
_RESALTADO_POR_HEX = {v: k for k, v in RESALTADOS_DOCX.items()}
_XML_INVALIDO = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffc\ufffe\uffff]")
FORMATOS_NUMERO = {
    QTextListFormat.ListDecimal: "decimal",
    QTextListFormat.ListLowerAlpha: "lowerLetter",
    QTextListFormat.ListUpperAlpha: "upperLetter",
    QTextListFormat.ListLowerRoman: "lowerRoman",
    QTextListFormat.ListUpperRoman: "upperRoman",
}
_ESTILO_POR_FORMATO = {v: k for k, v in FORMATOS_NUMERO.items()}


def _w(tag):
    return f"{{{NS_W}}}{tag}"


def _attr(el, nombre, defecto=None):
    if el is None:
        return defecto
    return el.get(_w(nombre), defecto)


def _num(el, nombre, defecto=None):
    try:
        return float(_attr(el, nombre))
    except (TypeError, ValueError):
        return defecto


def _activo(el):
    return _attr(el, "val", "true").lower() not in ("0", "false", "off", "none")


def _num_simple(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def id_estilo_docx(id_):
    """Id de estilo para Word: Body = Normal; "Heading 1" = Heading1."""
    if id_ == "Body":
        return "Normal"
    return re.sub(r"[^A-Za-z0-9]", "", id_) or "Estilo"


class EscritorDocx:
    def __init__(self, editor, config):
        self.editor = editor
        self.config = config
        self.dpi = editor.logicalDpiY()
        self.com = editor.comentarios
        self.usados = set()
        self.listas = {}           # índice de la lista → (numId, formato)
        self.medios = []           # (rId, ruta en el paquete, bytes, formato)
        self._medio_por_nombre = {}
        f = editor.document().defaultFont()
        self.familia = f.family() or BODY["font_family"]
        self.tamano = f.pointSizeF() if f.pointSizeF() > 0 else BODY["font_size"]

    def tw(self, px):
        return round(px * 1440 / self.dpi)

    def rpr(self, cf):
        familias = cf.fontFamilies() if cf.hasProperty(QTextFormat.FontFamilies) else None
        familia = (familias[0] if familias else None) or self.familia
        f = xml_escape(familia, {'"': "&quot;"})
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
        if cf.hasProperty(QTextFormat.FontLetterSpacing) and \
                cf.fontLetterSpacingType() == QFont.AbsoluteSpacing and \
                cf.fontLetterSpacing():
            pt = cf.fontLetterSpacing() * 72 / self.dpi
            p.append(f'<w:spacing w:val="{round(pt * 20)}"/>')
        medio = max(2, round((cf.fontPointSize() or self.tamano) * 2))
        p.append(f'<w:sz w:val="{medio}"/><w:szCs w:val="{medio}"/>')
        fondo = None
        if cf.hasProperty(QTextFormat.BackgroundBrush) and \
                cf.background().style() != Qt.NoBrush:
            fondo = cf.background().color().name()[1:].upper()
            if fondo in _RESALTADO_POR_HEX:
                p.append(f'<w:highlight w:val="{_RESALTADO_POR_HEX[fondo]}"/>')
                fondo = None
        if cf.fontUnderline():
            p.append('<w:u w:val="single"/>')
        if fondo:
            p.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{fondo}"/>')
        va = cf.verticalAlignment()
        if va == QTextCharFormat.AlignSuperScript:
            p.append('<w:vertAlign w:val="superscript"/>')
        elif va == QTextCharFormat.AlignSubScript:
            p.append('<w:vertAlign w:val="subscript"/>')
        return "<w:rPr>" + "".join(p) + "</w:rPr>"

    @staticmethod
    def texto(t):
        partes = []
        for i, linea in enumerate(_XML_INVALIDO.sub("", t).split("\u2028")):
            if i:
                partes.append("<w:br/>")
            for j, trozo in enumerate(linea.split("\t")):
                if j:
                    partes.append("<w:tab/>")
                if trozo:
                    partes.append(f'<w:t xml:space="preserve">{xml_escape(trozo)}</w:t>')
        return "".join(partes)

    def ppr(self, block):
        bf = block.blockFormat()
        p = []
        v = bf.property(PROP_ESTILO)
        if isinstance(v, str) and v in GESTOR.datos and v != "Body":
            p.append(f'<w:pStyle w:val="{id_estilo_docx(v)}"/>')
        elif bf.headingLevel() > 0:
            p.append(f'<w:pStyle w:val="Heading{min(6, bf.headingLevel())}"/>')
        lst = block.textList()
        izquierda, primera = bf.leftMargin(), bf.textIndent()
        if lst is not None:
            clave = lst.objectIndex()
            if clave not in self.listas:
                self.listas[clave] = (len(self.listas) + 1, lst.format().style())
            nivel = max(0, min(8, lst.format().indent() - 1))
            p.append(f'<w:numPr><w:ilvl w:val="{nivel}"/>'
                     f'<w:numId w:val="{self.listas[clave][0]}"/></w:numPr>')
            ancho = self.editor.document().indentWidth()
            izquierda += ancho * max(1, lst.format().indent())
            if primera == 0:
                primera = -ancho * 0.6
        tipos = QTextBlockFormat.LineHeightTypes
        tipo = bf.lineHeightType()
        if tipo == tipos.ProportionalHeight.value:
            linea, regla = round(240 * bf.lineHeight() / 100), "auto"
        elif tipo == tipos.FixedHeight.value:
            linea, regla = self.tw(bf.lineHeight()), "exact"
        elif tipo == tipos.MinimumHeight.value:
            linea, regla = self.tw(bf.lineHeight()), "atLeast"
        else:
            linea, regla = 240, "auto"
        p.append(f'<w:spacing w:before="{self.tw(bf.topMargin())}" '
                 f'w:after="{self.tw(bf.bottomMargin())}" '
                 f'w:line="{max(1, linea)}" w:lineRule="{regla}"/>')
        pl = self.tw(primera)
        p.append(f'<w:ind w:left="{self.tw(izquierda)}" w:right="{self.tw(bf.rightMargin())}" '
                 + (f'w:hanging="{-pl}"/>' if pl < 0 else f'w:firstLine="{pl}"/>'))
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
        p.append(self.rpr(block.charFormat()))
        return "<w:pPr>" + "".join(p) + "</w:pPr>"

    def segmentos(self, block, validos):
        """[(rPr | "img", textos | formato, ids, n)], uniendo corridas iguales."""
        segs = []
        it = block.begin()
        while not it.atEnd():
            fr = it.fragment()
            if fr.isValid() and fr.charFormat().isImageFormat():
                cf = fr.charFormat()
                ids = frozenset(ids_comentario(cf) & validos)
                for _ in fr.text():      # una corrida por imagen
                    segs.append(("img", [cf.toImageFormat()], ids, 0))
            elif fr.isValid():
                cf = fr.charFormat()
                ids = frozenset(ids_comentario(cf) & validos)
                rpr = self.rpr(cf)
                if segs and segs[-1][0] == rpr and segs[-1][2] == ids:
                    segs[-1][1].append(fr.text())
                else:
                    segs.append((rpr, [fr.text()], ids, 0))
            it += 1
        return segs

    def documento_xml(self):
        c = self.config
        tw = lambda cm: round(cm * TWIPS_POR_CM)
        orient = ' w:orient="landscape"' if c["ancho"] > c["alto"] else ""
        validos = set(self.com.datos)
        bloques = []
        b = self.editor.document().begin()
        while b.isValid():
            bloques.append(b)
            b = b.next()
        segs = [self.segmentos(b, validos) for b in bloques]
        plano = [s[2] for ss in segs for s in ss]
        self.usados = set()
        abiertos, k, cuerpo = set(), 0, []
        for b, ss in zip(bloques, segs):
            partes = []
            for clave, contenido, ids, n in ss:
                for i in sorted(ids - abiertos):
                    partes.append(f'<w:commentRangeStart w:id="{i}"/>')
                    abiertos.add(i)
                    self.usados.add(i)
                if clave == "img":
                    partes.append(self.dibujo(contenido[0], b))
                else:
                    partes.append(f"<w:r>{clave}{self.texto(''.join(contenido))}</w:r>")
                siguiente = plano[k + 1] if k + 1 < len(plano) else frozenset()
                for i in sorted(ids - siguiente):
                    partes.append(f'<w:commentRangeEnd w:id="{i}"/>'
                                  f'<w:r><w:commentReference w:id="{i}"/></w:r>')
                    abiertos.discard(i)
                k += 1
            cuerpo.append(f"<w:p>{self.ppr(b)}{''.join(partes)}</w:p>")
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document xmlns:w="{NS_W}" xmlns:r="{NS_R}" xmlns:wp="{NS_WP}" '
            f'xmlns:a="{NS_A}" xmlns:pic="{NS_PIC}"><w:body>' + "".join(cuerpo) +
            f'<w:sectPr><w:pgSz w:w="{tw(c["ancho"])}" w:h="{tw(c["alto"])}"{orient}/>'
            f'<w:pgMar w:top="{tw(c["superior"])}" w:right="{tw(c["derecho"])}" '
            f'w:bottom="{tw(c["inferior"])}" w:left="{tw(c["izquierdo"])}" '
            'w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>')

    def _medio(self, nombre):
        """rId de la imagen (los bytes originales, sin volver a comprimir)."""
        if nombre in self._medio_por_nombre:
            return self._medio_por_nombre[nombre]
        datos = self.editor.imagenes.get(nombre)
        if datos is not None:
            contenido, formato = datos[0], datos[1]
        else:
            res = self.editor.document().resource(QTextDocument.ImageResource, QUrl(nombre))
            if isinstance(res, QPixmap):
                res = res.toImage()
            if not isinstance(res, QImage) or res.isNull():
                self._medio_por_nombre[nombre] = None
                return None
            contenido, formato = imagen_a_bytes(res), "png"
        k = len(self.medios) + 1
        rid = f"rIdImg{k}"
        self.medios.append((rid, f"media/imagen{k}.{formato}", contenido, formato))
        self._medio_por_nombre[nombre] = rid
        return rid

    def dibujo(self, fmt, bloque=None):
        """Imagen en línea (<w:drawing>) con su tamaño en EMU."""
        rid = self._medio(fmt.name())
        if rid is None:
            return ""
        if not hasattr(self, "separaciones"):
            self.separaciones = {}
        w, h = self.editor._tamano_imagen(fmt)
        cx = round(w / self.dpi * 914400)
        cy = round(h / self.dpi * 914400)
        self._dibujos = getattr(self, "_dibujos", 0) + 1
        n = self._dibujos
        if bloque is not None and es_bloque_imagen(bloque) and \
                bloque.blockFormat().hasProperty(PROP_SEPARACION):
            # Word no la conoce (usa el espacio antes/después ya calculado);
            # se guarda para recuperarla al abrir en este editor
            self.separaciones[n] = self.editor.separacion_imagen(bloque)
        return (
            f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{n}" name="Imagen {n}"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="Imagen {n}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

    def numeracion_xml(self):
        ancho = self.editor.document().indentWidth()
        vinetas = ["\u2022", "\u25e6", "\u25aa"]
        partes_abs, partes_num = [], []
        for num, estilo in sorted(self.listas.values(), key=lambda x: x[0]):
            formato = FORMATOS_NUMERO.get(estilo)
            niveles = []
            for n in range(9):
                if formato:
                    fmt, texto = formato, f"%{n + 1}."
                else:
                    fmt, texto = "bullet", vinetas[n % 3]
                niveles.append(
                    f'<w:lvl w:ilvl="{n}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                    f'<w:lvlText w:val="{texto}"/><w:lvlJc w:val="left"/>'
                    f'<w:pPr><w:ind w:left="{self.tw(ancho * (n + 1))}" '
                    f'w:hanging="{self.tw(ancho * 0.6)}"/></w:pPr></w:lvl>')
            partes_abs.append(f'<w:abstractNum w:abstractNumId="{num}">'
                              + "".join(niveles) + "</w:abstractNum>")
            partes_num.append(f'<w:num w:numId="{num}"><w:abstractNumId w:val="{num}"/></w:num>')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:numbering xmlns:w="{NS_W}">' + "".join(partes_abs + partes_num)
                + "</w:numbering>")

    def comentarios_xml(self):
        q = {'"': "&quot;"}
        partes = []
        for n in sorted(self.usados):
            d = self.com.datos[n]
            autor = d.get("autor") or "Autor"
            iniciales = "".join(p[0] for p in autor.split()[:3]).upper()
            ps = []
            for j, linea in enumerate((d.get("texto") or "").split("\n")):
                ref = "<w:r><w:annotationRef/></w:r>" if j == 0 else ""
                ps.append(f"<w:p>{ref}<w:r>{self.texto(linea)}</w:r></w:p>")
            partes.append(
                f'<w:comment w:id="{n}" w:author="{xml_escape(autor, q)}" '
                f'w:date="{xml_escape(d.get("fecha") or ahora_iso(), q)}" '
                f'w:initials="{xml_escape(iniciales, q)}">{"".join(ps)}</w:comment>')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:comments xmlns:w="{NS_W}">' + "".join(partes) + "</w:comments>")

    def estilos_xml(self):
        b = GESTOR.resuelto("Body")
        f = xml_escape(b["font_family"], {'"': "&quot;"})
        xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               f'<w:styles xmlns:w="{NS_W}"><w:docDefaults><w:rPrDefault><w:rPr>'
               f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{f}"/>'
               f'<w:sz w:val="{round(b["font_size"] * 2)}"/>'
               f'<w:szCs w:val="{round(b["font_size"] * 2)}"/>'
               '<w:lang w:val="es-MX" w:eastAsia="es-MX" w:bidi="ar-SA"/>'
               '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
               '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
               '</w:pPr></w:pPrDefault></w:docDefaults>']
        for id_ in GESTOR.orden:
            e = GESTOR.resuelto(id_)
            nivel = GESTOR.datos[id_]["nivel"]
            sid = id_estilo_docx(id_)
            if id_ == "Body":
                nombre = "Normal"
            elif re.fullmatch(r"Heading \d", id_):
                nombre = id_.lower()
            elif id_ in ("Title", "Subtitle"):
                nombre = id_
            else:
                nombre = GESTOR.nombre(id_)
            jc = {Qt.AlignHCenter: "center", Qt.AlignRight: "right",
                  Qt.AlignJustify: "both"}.get(e["alignment"], "left")
            ppr = "<w:pPr>" + ("<w:keepNext/><w:keepLines/>" if nivel else "")
            ppr += (f'<w:spacing w:before="{self.tw(e["space_before"])}" '
                    f'w:after="{self.tw(e["space_after"])}" '
                    f'w:line="{round(240 * e["line_spacing"] / 100)}" w:lineRule="auto"/>'
                    f'<w:ind w:left="{self.tw(e["indent_left"])}" '
                    f'w:right="{self.tw(e["indent_right"])}" '
                    f'w:firstLine="{self.tw(e["first_line_indent"])}"/>'
                    f'<w:jc w:val="{jc}"/>')
            if nivel:
                ppr += f'<w:outlineLvl w:val="{nivel - 1}"/>'
            ppr += "</w:pPr>"
            ff = xml_escape(e["font_family"], {'"': "&quot;"})
            medio = round(e["font_size"] * 2)
            rpr = (f'<w:rPr><w:rFonts w:ascii="{ff}" w:hAnsi="{ff}" w:cs="{ff}" w:eastAsia="{ff}"/>'
                   + ("<w:b/><w:bCs/>" if e["bold"] else "")
                   + ("<w:i/><w:iCs/>" if e["italic"] else "")
                   + f'<w:color w:val="{e["font_color"][1:].upper()}"/>'
                   f'<w:sz w:val="{medio}"/><w:szCs w:val="{medio}"/></w:rPr>')
            defecto = ' w:default="1"' if id_ == "Body" else ""
            base = "" if id_ == "Body" else '<w:basedOn w:val="Normal"/>'
            sig = GESTOR.datos[id_]["siguiente"]
            sig_xml = f'<w:next w:val="{id_estilo_docx(sig)}"/>' if sig in GESTOR.datos else ""
            xml.append(f'<w:style w:type="paragraph"{defecto} w:styleId="{sid}">'
                       f'<w:name w:val="{xml_escape(nombre, {chr(34): "&quot;"})}"/>'
                       f'{base}{sig_xml}<w:qFormat/>{ppr}{rpr}</w:style>')
        xml.append('<w:style w:type="character" w:default="1" '
                   'w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>'
                   '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
                   "</w:styles>")
        return "".join(xml)

    def guardar(self, ruta, titulo=""):
        ahora = ahora_iso()
        ct = "application/vnd.openxmlformats-officedocument"
        documento = self.documento_xml()        # antes: llena listas y usados
        tipos = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                 '<Default Extension="xml" ContentType="application/xml"/>']
        tipos += [f'<Override PartName="/word/document.xml" ContentType="{ct}.wordprocessingml.document.main+xml"/>',
                  f'<Override PartName="/word/styles.xml" ContentType="{ct}.wordprocessingml.styles+xml"/>',
                  f'<Override PartName="/word/settings.xml" ContentType="{ct}.wordprocessingml.settings+xml"/>',
                  '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>']
        rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        rels = [f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>',
                f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/>']
        partes = {
            "word/document.xml": documento,
            "word/styles.xml": self.estilos_xml(),
            "word/settings.xml":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:settings xmlns:w="{NS_W}"><w:view w:val="web"/>'
                '<w:defaultTabStop w:val="709"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                '<w:docVars>' + "".join(
                    f'<w:docVar w:name="EditorLibroComentario{n}" w:val="{self.com.color(n)}"/>'
                    for n in sorted(self.usados)) + "".join(
                    f'<w:docVar w:name="EditorLibroSepImagen{n}" w:val="{v:g}"/>'
                    for n, v in sorted(getattr(self, "separaciones", {}).items()))
                + '</w:docVars></w:settings>',
            "docProps/core.xml":
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                f'<dc:title>{xml_escape(titulo)}</dc:title><dc:language>es-MX</dc:language>'
                f'<dcterms:created xsi:type="dcterms:W3CDTF">{ahora}</dcterms:created>'
                f'<dcterms:modified xsi:type="dcterms:W3CDTF">{ahora}</dcterms:modified>'
                '</cp:coreProperties>',
        }
        if self.usados:
            tipos.append(f'<Override PartName="/word/comments.xml" ContentType="{ct}.wordprocessingml.comments+xml"/>')
            rels.append(f'<Relationship Id="rId3" Type="{rel}/comments" Target="comments.xml"/>')
            partes["word/comments.xml"] = self.comentarios_xml()
        for formato in sorted({m[3] for m in self.medios}):
            tipos.append(f'<Default Extension="{formato}" '
                         f'ContentType="{TIPOS_IMAGEN_DOCX.get(formato, "image/" + formato)}"/>')
        for rid, archivo, contenido, _ in self.medios:
            rels.append(f'<Relationship Id="{rid}" Type="{rel}/image" Target="{archivo}"/>')
            partes[f"word/{archivo}"] = contenido
        if self.listas:
            tipos.append(f'<Override PartName="/word/numbering.xml" ContentType="{ct}.wordprocessingml.numbering+xml"/>')
            rels.append(f'<Relationship Id="rIdNum" Type="{rel}/numbering" Target="numbering.xml"/>')
            partes["word/numbering.xml"] = self.numeracion_xml()
        partes["[Content_Types].xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            + "".join(tipos) + "</Types>")
        partes["_rels/.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<Relationships xmlns="{NS_REL}">'
            f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/'
            'relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            "</Relationships>")
        partes["word/_rels/document.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<Relationships xmlns="{NS_REL}">' + "".join(rels) + "</Relationships>")
        # Se escribe en un temporal y luego se reemplaza: si algo falla,
        # el archivo anterior queda intacto
        temporal = ruta + ".tmp"
        try:
            with zipfile.ZipFile(temporal, "w", zipfile.ZIP_DEFLATED) as z:
                for nombre, contenido in partes.items():
                    z.writestr(nombre, contenido if isinstance(contenido, bytes)
                               else contenido.encode("utf-8"))
            with open(temporal, "rb+") as f:
                os.fsync(f.fileno())
            os.replace(temporal, ruta)
        finally:
            if os.path.exists(temporal):
                os.remove(temporal)


class LectorDocx:
    """Lee un .docx (de este editor, Word, Google Docs o Writer) resolviendo
    la herencia de estilos. Las tablas se importan como párrafos."""

    def __init__(self, ruta, dpi):
        self.dpi = dpi
        self.mapa = {id_estilo_docx(i): i for i in GESTOR.datos}
        with zipfile.ZipFile(ruta) as z:
            nombres = set(z.namelist())
            principal = "word/document.xml"
            if "_rels/.rels" in nombres:
                for r in ET.fromstring(z.read("_rels/.rels")):
                    if r.get("Type", "").endswith("/officeDocument"):
                        principal = r.get("Target", principal).lstrip("/")
            if principal not in nombres:
                raise ValueError("El archivo no contiene un documento de Word.")
            self.doc = ET.fromstring(z.read(principal))
            carpeta = principal.rsplit("/", 1)[0]
            leer = lambda n: ET.fromstring(z.read(n)) if n and n in nombres else None
            self._leer_tema(leer(f"{carpeta}/theme/theme1.xml"))
            self._leer_estilos(leer(f"{carpeta}/styles.xml"))
            self._leer_numeracion(leer(f"{carpeta}/numbering.xml"))
            ruta_com = None
            self.medios = {}          # rId → bytes de la imagen
            rels = f"{carpeta}/_rels/{principal.rsplit('/', 1)[-1]}.rels"
            if rels in nombres:
                for r in ET.fromstring(z.read(rels)):
                    destino = r.get("Target", "")
                    ruta_parte = destino.lstrip("/") if destino.startswith("/") \
                        else posixpath.normpath(f"{carpeta}/{destino}")
                    tipo = r.get("Type", "")
                    if tipo.endswith("/comments"):
                        ruta_com = ruta_parte
                    elif tipo.endswith("/image") and ruta_parte in nombres \
                            and r.get("TargetMode") != "External":
                        self.medios[r.get("Id")] = z.read(ruta_parte)
            self._leer_comentarios(leer(ruta_com))
            self.sep_imagenes = {}    # n.º de imagen → separación (renglones)
            self._n_imagen = 0
            ajustes = leer(f"{carpeta}/settings.xml")
            if ajustes is not None:
                for var in ajustes.iter(_w("docVar")):
                    nombre, valor = _attr(var, "name") or "", _attr(var, "val")
                    m_sep = re.fullmatch(r"EditorLibroSepImagen(\d+)", nombre)
                    if m_sep:
                        try:
                            self.sep_imagenes[int(m_sep.group(1))] = float(valor)
                        except (TypeError, ValueError):
                            pass
                        continue
                    m = re.fullmatch(r"(?:EditorLibroComentario|SWordBookComment)(\d+)", nombre)
                    if m and int(m.group(1)) in self.comentarios and QColor(valor or "").isValid():
                        self.comentarios[int(m.group(1))]["color"] = valor

    def px(self, twips):
        return (twips or 0) * self.dpi / 1440

    def _leer_tema(self, tema):
        self.fuentes_tema = {}
        if tema is None:
            return
        for clave, nombre in (("minor", "minorFont"), ("major", "majorFont")):
            el = tema.find(f".//{{{NS_A}}}{nombre}/{{{NS_A}}}latin")
            if el is not None and el.get("typeface"):
                self.fuentes_tema[clave] = el.get("typeface")

    def _leer_estilos(self, raiz):
        self.estilos = {}
        self.ppr_defecto, self.rpr_defecto = {}, {}
        self.estilo_defecto = None
        if raiz is None:
            return
        dd = raiz.find(_w("docDefaults"))
        if dd is not None:
            self.rpr_defecto = self.leer_rpr(dd.find(f"{_w('rPrDefault')}/{_w('rPr')}"))
            self.ppr_defecto = self.leer_ppr(dd.find(f"{_w('pPrDefault')}/{_w('pPr')}"))
        for st in raiz.findall(_w("style")):
            id_ = _attr(st, "styleId")
            tipo = _attr(st, "type")
            self.estilos[id_] = {
                "nombre": (_attr(st.find(_w("name")), "val") or "").lower(),
                "base": _attr(st.find(_w("basedOn")), "val"),
                "ppr": self.leer_ppr(st.find(_w("pPr"))),
                "rpr": self.leer_rpr(st.find(_w("rPr"))),
            }
            if tipo == "paragraph" and _attr(st, "default", "0") in ("1", "true", "on"):
                self.estilo_defecto = id_

    def _leer_numeracion(self, raiz):
        self.numeracion = {}          # numId → {nivel: formato}
        if raiz is None:
            return
        abstractos = {}
        for a in raiz.findall(_w("abstractNum")):
            niveles = {}
            for lvl in a.findall(_w("lvl")):
                try:
                    n = int(_attr(lvl, "ilvl", "0"))
                except ValueError:
                    continue
                niveles[n] = _attr(lvl.find(_w("numFmt")), "val", "bullet")
            abstractos[_attr(a, "abstractNumId")] = niveles
        for num in raiz.findall(_w("num")):
            self.numeracion[_attr(num, "numId")] = abstractos.get(
                _attr(num.find(_w("abstractNumId")), "val"), {})

    def _leer_comentarios(self, raiz):
        self.comentarios = {}
        if raiz is None:
            return
        for c in raiz.findall(_w("comment")):
            try:
                i = int(_attr(c, "id"))
            except (TypeError, ValueError):
                continue
            parrafos = []
            for p in c.iter(_w("p")):
                trozos = []
                for el in p.iter():
                    t = el.tag.split("}")[-1]
                    if t == "t":
                        trozos.append(el.text or "")
                    elif t == "tab":
                        trozos.append("\t")
                    elif t in ("br", "cr"):
                        trozos.append("\n")
                parrafos.append("".join(trozos))
            self.comentarios[i] = {"autor": _attr(c, "author") or "",
                                   "fecha": _attr(c, "date") or "",
                                   "texto": "\n".join(parrafos).strip()}

    def _marca_comentario(self, tag, el, abiertos):
        try:
            i = int(_attr(el, "id"))
        except (TypeError, ValueError):
            return
        if tag == "commentRangeStart" and i in self.comentarios:
            abiertos.add(i)
        elif tag == "commentRangeEnd":
            abiertos.discard(i)

    def cadena(self, id_):
        res, vistos = [], set()
        while id_ and id_ in self.estilos and id_ not in vistos:
            vistos.add(id_)
            res.insert(0, self.estilos[id_])
            id_ = self.estilos[id_]["base"]
        return res

    def leer_rpr(self, el):
        d = {}
        if el is None:
            return d
        for h in el:
            t = h.tag.split("}")[-1]
            if t == "rFonts":
                fuente = _attr(h, "ascii") or _attr(h, "hAnsi")
                tema = _attr(h, "asciiTheme") or _attr(h, "hAnsiTheme")
                if not fuente and tema:
                    fuente = self.fuentes_tema.get("major" if tema.startswith("major") else "minor")
                if fuente:
                    d["fuente"] = fuente
            elif t in ("b", "i", "strike"):
                d[t] = _activo(h)
            elif t == "u":
                d["u"] = _attr(h, "val", "single") != "none"
            elif t == "color":
                v = _attr(h, "val", "auto")
                d["color"] = None if v == "auto" else v
            elif t == "sz":
                v = _num(h, "val")
                if v:
                    d["tamano"] = v / 2
            elif t == "highlight":
                d["fondo"] = RESALTADOS_DOCX.get(_attr(h, "val"))
            elif t == "shd":
                relleno = _attr(h, "fill", "auto")
                if relleno != "auto" and "fondo" not in d:
                    d["fondo"] = relleno
            elif t == "spacing":
                v = _num(h, "val")
                if v is not None:
                    d["letras"] = v / 20
            elif t == "vertAlign":
                d["va"] = _attr(h, "val", "baseline")
            elif t == "rStyle":
                d["estilo"] = _attr(h, "val")
        return d

    def leer_ppr(self, el):
        d = {}
        if el is None:
            return d
        for h in el:
            t = h.tag.split("}")[-1]
            if t == "pStyle":
                d["estilo"] = _attr(h, "val")
            elif t == "jc":
                d["jc"] = _attr(h, "val")
            elif t == "numPr":
                num = _attr(h.find(_w("numId")), "val")
                nivel = _num(h.find(_w("ilvl")), "val", 0) or 0
                if num is not None:
                    d["lista"] = (num, int(nivel))
            elif t == "outlineLvl":
                d["esquema"] = _num(h, "val")
            elif t == "spacing":
                for clave in ("before", "after", "line"):
                    v = _num(h, clave)
                    if v is not None:
                        d[clave] = v
                if _attr(h, "lineRule"):
                    d["regla"] = _attr(h, "lineRule")
            elif t == "ind":
                for clave, nombres in (("left", ("left", "start")), ("right", ("right", "end"))):
                    for n in nombres:
                        v = _num(h, n)
                        if v is not None:
                            d[clave] = v
                            break
                if _num(h, "hanging") is not None:
                    d["primera"] = -_num(h, "hanging")
                elif _num(h, "firstLine") is not None:
                    d["primera"] = _num(h, "firstLine")
        return d

    def estilo_editor(self, sid, cadena):
        """Id del estilo del editor que corresponde al estilo de Word."""
        if sid in self.mapa:
            return self.mapa[sid]
        for st in reversed(cadena):
            n = st["nombre"]
            m = re.fullmatch(r"heading\s*(\d)", n)
            if m and f"Heading {min(3, int(m.group(1)))}" in GESTOR.datos:
                return f"Heading {min(3, int(m.group(1)))}"
            if n in ("title", "subtitle"):
                return n.capitalize()
        return "Body"

    def formato_bloque(self, ppr, id_):
        bf = QTextBlockFormat()
        bf.setAlignment({"center": Qt.AlignHCenter, "right": Qt.AlignRight,
                         "end": Qt.AlignRight, "both": Qt.AlignJustify,
                         "distribute": Qt.AlignJustify}.get(ppr.get("jc", "left"), Qt.AlignLeft))
        bf.setTopMargin(self.px(ppr.get("before")))
        bf.setBottomMargin(self.px(ppr.get("after")))
        bf.setLeftMargin(self.px(ppr.get("left")))
        bf.setRightMargin(self.px(ppr.get("right")))
        bf.setTextIndent(self.px(ppr.get("primera")))
        tipos = QTextBlockFormat.LineHeightTypes
        linea, regla = ppr.get("line"), ppr.get("regla", "auto")
        if linea and regla == "exact":
            bf.setLineHeight(self.px(linea), tipos.FixedHeight.value)
        elif linea and regla == "atLeast":
            bf.setLineHeight(self.px(linea), tipos.MinimumHeight.value)
        else:
            bf.setLineHeight(round((linea or 240) / 240 * 100), tipos.ProportionalHeight.value)
        bf.setProperty(PROP_ESTILO, id_)
        bf.setHeadingLevel(GESTOR.datos.get(id_, {}).get("nivel", 0))
        return bf

    def formato_caracter(self, rpr):
        cf = QTextCharFormat()
        cf.setFontFamilies([rpr.get("fuente") or "Times New Roman"])
        cf.setFontPointSize(rpr.get("tamano", 10))
        cf.setFontWeight(QFont.Bold if rpr.get("b") else QFont.Normal)
        cf.setFontItalic(bool(rpr.get("i")))
        cf.setFontUnderline(bool(rpr.get("u")))
        cf.setFontStrikeOut(bool(rpr.get("strike")))
        color = rpr.get("color")
        cf.setForeground(QColor("#" + color) if color and QColor("#" + color).isValid()
                         else QColor("#000000"))
        fondo = rpr.get("fondo")
        if fondo and QColor("#" + fondo).isValid():
            cf.setBackground(QColor("#" + fondo))
        aplicar_letras(cf, rpr.get("letras", 0) * self.dpi / 72)
        va = rpr.get("va")
        if va == "superscript":
            cf.setVerticalAlignment(QTextCharFormat.AlignSuperScript)
        elif va == "subscript":
            cf.setVerticalAlignment(QTextCharFormat.AlignSubScript)
        return cf

    def parrafos(self):
        """Produce (formato de bloque, formato de marca, [(texto, formato)]).
        Las imágenes llegan como (CARACTER_OBJETO, {datos, ancho, alto, cf})."""
        cuerpo = self.doc.find(_w("body"))
        if cuerpo is None:
            return
        abiertos = set()
        for p in self._iter_parrafos(cuerpo):
            tag = p.tag.split("}")[-1]
            if tag != "p":
                self._marca_comentario(tag, p, abiertos)
                continue
            ppr_el = p.find(_w("pPr"))
            directo = self.leer_ppr(ppr_el)
            sid = directo.get("estilo") or self.estilo_defecto
            cadena = self.cadena(sid)
            ppr = dict(self.ppr_defecto)
            rpr_parrafo = dict(self.rpr_defecto)
            for st in cadena:
                ppr.update(st["ppr"])
                rpr_parrafo.update(st["rpr"])
            ppr.update(directo)
            id_ = self.estilo_editor(sid, cadena)
            bf = self.formato_bloque(ppr, id_)
            lst = ppr.get("lista")
            if lst and lst[0] != "0" and lst[0] in self.numeracion and \
                    not GESTOR.datos.get(id_, {}).get("nivel"):
                fmt = self.numeracion[lst[0]].get(lst[1], "bullet")
                bf.setProperty(PROP_LISTA, f"{fmt};{lst[0]};{lst[1]}")
                bf.setLeftMargin(0)
                bf.setTextIndent(0)
            marca = dict(rpr_parrafo)
            if ppr_el is not None:
                marca.update(self.leer_rpr(ppr_el.find(_w("rPr"))))
            corridas = []
            for r in self._iter_corridas(p):
                tag = r.tag.split("}")[-1]
                if tag != "r":
                    self._marca_comentario(tag, r, abiertos)
                    continue
                rpr = dict(rpr_parrafo)
                directo_r = self.leer_rpr(r.find(_w("rPr")))
                for st in self.cadena(directo_r.get("estilo")):
                    rpr.update(st["rpr"])
                rpr.update(directo_r)
                cf = self.formato_caracter(rpr)
                if abiertos:
                    poner_ids_comentario(cf, abiertos)
                for h in r:
                    t = h.tag.split("}")[-1]
                    if t in ("drawing", "AlternateContent"):
                        dibujo = h if t == "drawing" else h.find(f".//{_w('drawing')}")
                        img = self._imagen(dibujo) if dibujo is not None else None
                        if img is not None:
                            img["cf"] = cf
                            corridas.append((CARACTER_OBJETO, img))
                        continue
                    texto = {"t": h.text or "", "tab": "\t", "noBreakHyphen": "\u2011",
                             "softHyphen": "\u00ad", "cr": "\u2028"}.get(t)
                    if t == "br" and _attr(h, "type") != "page":
                        texto = "\u2028"
                    if texto:
                        corridas.append((texto, cf))
            yield bf, self.formato_caracter(marca), corridas

    def _imagen(self, dibujo):
        """Datos de una imagen (<w:drawing>): bytes y tamaño en px."""
        blip = dibujo.find(f".//{{{NS_A}}}blip")
        if blip is None:
            return None
        datos = self.medios.get(blip.get(f"{{{NS_R}}}embed"))
        if not datos:
            return None
        ext = dibujo.find(f".//{{{NS_WP}}}extent")
        ancho = alto = 0
        if ext is not None:
            ancho = _num_simple(ext.get("cx")) / 914400 * self.dpi
            alto = _num_simple(ext.get("cy")) / 914400 * self.dpi
        self._n_imagen += 1
        return {"datos": datos, "ancho": round(ancho), "alto": round(alto),
                "sep": self.sep_imagenes.get(self._n_imagen)}

    def _iter_parrafos(self, contenedor):
        for el in contenedor:
            t = el.tag.split("}")[-1]
            if t in ("p", "commentRangeStart", "commentRangeEnd"):
                yield el
            elif t == "tbl":
                yield from el.iter(_w("p"))
            elif t in ("sdt", "sdtContent", "customXml", "ins"):
                yield from self._iter_parrafos(el)

    def _iter_corridas(self, el):
        for h in el:
            t = h.tag.split("}")[-1]
            if t in ("r", "commentRangeStart", "commentRangeEnd"):
                yield h
            elif t in ("pPr", "del", "moveFrom", "rPr"):
                continue
            else:
                yield from self._iter_corridas(h)

    def config_pagina(self, actual):
        cuerpo = self.doc.find(_w("body"))
        sect = cuerpo.find(_w("sectPr")) if cuerpo is not None else None
        config = dict(actual)
        if sect is None:
            return config
        cm = lambda tw: round(abs(tw) / TWIPS_POR_CM, 2)
        tam, mar = sect.find(_w("pgSz")), sect.find(_w("pgMar"))
        if _num(tam, "w") and _num(tam, "h"):
            config["ancho"], config["alto"] = cm(_num(tam, "w")), cm(_num(tam, "h"))
        for clave, nombre in (("superior", "top"), ("inferior", "bottom"),
                              ("izquierdo", "left"), ("derecho", "right")):
            v = _num(mar, nombre)
            if v is not None:
                config[clave] = cm(v)
        return config


def _agregar_a_lista(cursor, bf, listas):
    valor = bf.stringProperty(PROP_LISTA) if bf.hasProperty(PROP_LISTA) else ""
    if not valor:
        return
    fmt, num, nivel = valor.split(";")
    bloque = cursor.block()
    c = QTextCursor(bloque)
    limpio = c.blockFormat()
    limpio.clearProperty(PROP_LISTA)
    c.setBlockFormat(limpio)
    clave = (num, nivel)
    lst = listas.get(clave)
    if lst is None:
        f = QTextListFormat()
        if fmt in _ESTILO_POR_FORMATO:
            f.setStyle(_ESTILO_POR_FORMATO[fmt])
        else:
            f.setStyle(VINETAS[int(nivel) % 3])
        f.setIndent(int(nivel) + 1)
        listas[clave] = c.createList(f)
    else:
        lst.add(bloque)


def _cursor_al_inicio(editor):
    cursor = QTextCursor(editor.document())
    editor.setTextCursor(cursor)
    editor.setCurrentCharFormat(cursor.charFormat())
    editor.formatoParrafoCambiado.emit()
    editor.document().clearUndoRedoStacks()   # nada que deshacer al abrir
    editor.document().setModified(False)


def abrir_docx(editor, vista, ruta):
    """Reemplaza el contenido del editor con el del .docx."""
    lector = LectorDocx(ruta, editor.logicalDpiY())
    parrafos = list(lector.parrafos())
    config = lector.config_pagina(vista.config)
    temp = QTextDocument()
    temp.setUndoRedoEnabled(False)
    cur = QTextCursor(temp)
    listas = {}
    primero = None
    nuevas = set()            # imágenes de este archivo
    for i, (bf, marca, corridas) in enumerate(parrafos):
        if i == 0:
            cur.setBlockFormat(bf)
            cur.setBlockCharFormat(marca)
            primero = (bf, marca)
        else:
            cur.insertBlock(bf, marca)
        solo_imagen = len(corridas) == 1 and isinstance(corridas[0][1], dict)
        for texto, cf in corridas:
            if isinstance(cf, dict):
                r = editor.registrar_imagen(cf["datos"], recurso=False)
                if r is None:
                    continue
                nombre, tam = r
                nuevas.add(nombre)
                fmt = editor._formato_imagen(nombre, cf["ancho"] or tam.width(),
                                             cf["alto"] or tam.height())
                ids = ids_comentario(cf["cf"])
                if ids:
                    poner_ids_comentario(fmt, ids)
                cur.insertImage(fmt)
            else:
                cur.insertText(texto, cf)
        if solo_imagen:
            # Párrafo de imagen: interlineado sencillo para que no se corte
            bfi = cur.blockFormat()
            bfi.setLineHeight(100, QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
            bfi.setTextIndent(0)
            bfi.clearProperty(PROP_ESTILO)    # los estilos no lo reformatean
            if corridas[0][1].get("sep") is not None:
                bfi.setProperty(PROP_SEPARACION, corridas[0][1]["sep"])
            cur.setBlockFormat(bfi)
        _agregar_a_lista(cur, bf, listas)
    doc = editor.document()
    editor.deseleccionar()
    editor.blockSignals(True)
    doc.setUndoRedoEnabled(False)
    try:
        doc.clear()
        editor.imagenes = {k: v for k, v in editor.imagenes.items() if k in nuevas}
        editor.recargar_recursos()
        if parrafos:
            t = QTextCursor(doc)
            t.setBlockFormat(primero[0])
            t.setBlockCharFormat(primero[1])
            t.insertFragment(QTextDocumentFragment(temp))
        else:
            editor.aplicar_body()
    finally:
        doc.setUndoRedoEnabled(True)
        editor.blockSignals(False)
    vista.aplicar_config(config)
    editor._interlineado_imagenes()
    editor.comentarios.cargar(lector.comentarios)
    _cursor_al_inicio(editor)


def abrir_texto(editor, vista, ruta):
    with open(ruta, "rb") as f:
        data = f.read()
    for codificacion in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = data.decode(codificacion)
            break
        except UnicodeDecodeError:
            continue
    doc = editor.document()
    editor.deseleccionar()
    editor.imagenes.clear()
    editor.blockSignals(True)
    doc.setUndoRedoEnabled(False)
    try:
        editor.setPlainText(texto.replace("\r\n", "\n").replace("\r", "\n"))
        editor.aplicar_body()
    finally:
        doc.setUndoRedoEnabled(True)
        editor.blockSignals(False)
    vista.aplicar_config(vista.config)
    editor.comentarios.limpiar()
    _cursor_al_inicio(editor)


# ---------------------------------------------------------------------------
# Iconos de las herramientas de archivo
# ---------------------------------------------------------------------------
def icono_archivo(tipo):
    escala, tam = 3, 24
    pm = QPixmap(tam * escala, tam * escala)
    pm.setDevicePixelRatio(escala)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    color = QColor(COLOR_ICONO)
    pen = QPen(color, 1.6)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen)
    p.setBrush(Qt.NoBrush)

    def hoja(x0, y0, x1, y1, d=4.5):
        p.drawPolygon(QPolygonF([QPointF(x0, y0), QPointF(x1 - d, y0), QPointF(x1, y0 + d),
                                 QPointF(x1, y1), QPointF(x0, y1)]))
        p.drawPolyline(QPolygonF([QPointF(x1 - d, y0), QPointF(x1 - d, y0 + d),
                                  QPointF(x1, y0 + d)]))

    def disquete(x0, y0, x1, y1):
        an = x1 - x0
        p.drawPolygon(QPolygonF([QPointF(x0, y0), QPointF(x1 - 3, y0), QPointF(x1, y0 + 3),
                                 QPointF(x1, y1), QPointF(x0, y1)]))
        p.drawRect(QRectF(x0 + an * 0.2, y1 - (y1 - y0) * 0.42, an * 0.6, (y1 - y0) * 0.42))
        p.drawPolyline(QPolygonF([QPointF(x0 + an * 0.28, y0), QPointF(x0 + an * 0.28, y0 + 4.5),
                                  QPointF(x1 - an * 0.3, y0 + 4.5), QPointF(x1 - an * 0.3, y0)]))

    if tipo == "nuevo":
        hoja(5, 3, 19, 21)
        p.drawLine(QPointF(12, 11), QPointF(12, 17))
        p.drawLine(QPointF(9, 14), QPointF(15, 14))
    elif tipo == "abrir":
        p.drawPolyline(QPolygonF([QPointF(4, 18), QPointF(4, 6), QPointF(9, 6), QPointF(11, 8),
                                  QPointF(18, 8), QPointF(18, 10.5)]))
        p.drawPolygon(QPolygonF([QPointF(4, 18), QPointF(7, 10.5), QPointF(21, 10.5),
                                 QPointF(18, 18)]))
    elif tipo == "guardar":
        disquete(4, 4, 20, 20)
    elif tipo == "guardar_como":
        disquete(3, 3, 16, 16)
        p.setBrush(QColor("#FFFFFF"))
        p.drawPolygon(QPolygonF([QPointF(13, 19.5), QPointF(13.6, 16.6), QPointF(19.3, 10.9),
                                 QPointF(21.6, 13.2), QPointF(15.9, 18.9)]))
    elif tipo == "cortar":
        p.drawLine(QPointF(7.5, 3.5), QPointF(15.2, 14.8))
        p.drawLine(QPointF(16.5, 3.5), QPointF(8.8, 14.8))
        p.drawEllipse(QPointF(7, 17.5), 3, 3)
        p.drawEllipse(QPointF(17, 17.5), 3, 3)
    elif tipo == "copiar":
        p.drawPolyline(QPolygonF([QPointF(5, 16), QPointF(5, 3.5), QPointF(14, 3.5)]))
        hoja(8.5, 7, 19.5, 21, 3.5)
    elif tipo == "pegar":
        p.drawPolyline(QPolygonF([QPointF(8, 5), QPointF(4.5, 5), QPointF(4.5, 21),
                                  QPointF(10, 21)]))
        p.drawPolyline(QPolygonF([QPointF(16, 5), QPointF(19.5, 5), QPointF(19.5, 9)]))
        p.drawRoundedRect(QRectF(8, 3, 8, 4), 1, 1)
        hoja(11, 11, 21, 22, 3)
    elif tipo == "buscar":
        p.drawEllipse(QRectF(4, 4, 11, 11))
        grueso = QPen(color, 2.4)
        grueso.setCapStyle(Qt.RoundCap)
        p.setPen(grueso)
        p.drawLine(QPointF(13.6, 13.6), QPointF(19.5, 19.5))
    elif tipo == "navegador":
        for y, x, largo in ((5.5, 4, 15), (11, 8, 11), (16.5, 8, 11)):
            p.setPen(Qt.NoPen)
            p.setBrush(color)
            p.drawEllipse(QPointF(x, y), 1.6, 1.6)
            p.setPen(pen)
            p.drawLine(QPointF(x + 3.5, y), QPointF(x + largo - 0.5, y))
    elif tipo in ("comentario", "comentarios"):
        globo = QPainterPath()
        globo.addRoundedRect(QRectF(3.5, 4, 17, 12), 3, 3)
        cola = QPainterPath()
        cola.moveTo(7.5, 15)
        cola.lineTo(7, 20.5)
        cola.lineTo(12.5, 15)
        cola.closeSubpath()
        p.drawPath(globo.united(cola))
        if tipo == "comentario":
            p.drawLine(QPointF(12, 7), QPointF(12, 13))
            p.drawLine(QPointF(9, 10), QPointF(15, 10))
        else:
            p.drawLine(QPointF(7.5, 8.2), QPointF(16.5, 8.2))
            p.drawLine(QPointF(7.5, 11.8), QPointF(13.5, 11.8))
    elif tipo == "imagen":
        p.drawRoundedRect(QRectF(3.5, 5, 17, 14), 2, 2)
        p.drawEllipse(QPointF(15.5, 9.3), 1.7, 1.7)
        p.drawPolyline(QPolygonF([QPointF(3.8, 17), QPointF(9, 11.5),
                                  QPointF(13, 15.5), QPointF(15.2, 13.3),
                                  QPointF(20.2, 18)]))
    elif tipo == "fecha":
        # Calendario con un reloj pequeño en la esquina
        p.drawRoundedRect(QRectF(3.5, 5, 14, 13.5), 1.5, 1.5)
        p.drawLine(QPointF(3.5, 8.8), QPointF(17.5, 8.8))
        p.drawLine(QPointF(7, 3.5), QPointF(7, 6.5))
        p.drawLine(QPointF(14, 3.5), QPointF(14, 6.5))
        for x, y in ((7, 12), (10.5, 12), (7, 15.2)):
            p.drawPoint(QPointF(x, y))
        p.setBrush(QColor("#FFFFFF"))
        p.drawEllipse(QPointF(16.5, 16.5), 4.5, 4.5)
        p.setBrush(Qt.NoBrush)
        p.drawPolyline(QPolygonF([QPointF(16.5, 14), QPointF(16.5, 16.7),
                                  QPointF(18.3, 17.8)]))
    elif tipo == "eliminar":
        p.drawLine(QPointF(4.5, 6.5), QPointF(19.5, 6.5))
        p.drawPolyline(QPolygonF([QPointF(9.5, 6.5), QPointF(9.5, 4),
                                  QPointF(14.5, 4), QPointF(14.5, 6.5)]))
        p.drawPolygon(QPolygonF([QPointF(6.5, 6.5), QPointF(17.5, 6.5),
                                 QPointF(16.5, 20.5), QPointF(7.5, 20.5)]))
        for x in (10, 12, 14):
            p.drawLine(QPointF(x, 10), QPointF(x, 17.5))
    p.end()
    return QIcon(pm)


# ---------------------------------------------------------------------------
# Insertar fecha y hora
# ---------------------------------------------------------------------------
# Formatos de QLocale (español de México: «oct», «sábado», «A.M.»…)
FORMATOS_FECHA_HORA = [
    "dd/MM/yyyy",
    "dd/MM/yy",
    "yyyy-MM-dd",
    "dd-MMM-yy",
    "dddd, d 'de' MMMM 'de' yyyy",
    "d 'de' MMMM 'de' yyyy",
    "MMMM 'de' yyyy",
    "dd/MM/yyyy HH:mm",
    "HH:mm",
    "HH:mm:ss",
    "h:mm AP",
    "h:mm:ss AP",
]


class VentanaFechaHora(QDialog):
    """Lista de formatos con la fecha y hora actuales (se actualiza cada
    segundo). Aceptar, Intro o doble clic insertan el formato elegido."""
    ultimo = 0                       # formato elegido la última vez

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Fecha y hora")
        self.setWindowIcon(icono_archivo("fecha"))
        self.locale = QLocale(QLocale.Spanish, QLocale.Mexico)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Formatos disponibles:"))
        self.lista = QListWidget()
        self.lista.setSelectionMode(QAbstractItemView.SingleSelection)
        for _ in FORMATOS_FECHA_HORA:
            self.lista.addItem(QListWidgetItem())
        self.lista.itemDoubleClicked.connect(lambda _it: self.accept())
        lay.addWidget(self.lista)
        botones = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        botones.button(QDialogButtonBox.Ok).setText("Aceptar")
        botones.button(QDialogButtonBox.Cancel).setText("Cancelar")
        botones.accepted.connect(self.accept)
        botones.rejected.connect(self.reject)
        lay.addWidget(botones)
        self._actualizar()
        # Ancho para que quepa el formato más largo sin barra horizontal
        fm = self.lista.fontMetrics()
        ancho = max(fm.horizontalAdvance(self.lista.item(i).text())
                    for i in range(self.lista.count()))
        self.lista.setMinimumWidth(ancho + 48)
        self.lista.setMinimumHeight(
            self.lista.sizeHintForRow(0) * self.lista.count() + 2 * self.lista.frameWidth() + 8)
        self.lista.setCurrentRow(min(VentanaFechaHora.ultimo, self.lista.count() - 1))
        self.lista.setFocus()
        self._reloj = QTimer(self)
        self._reloj.timeout.connect(self._actualizar)
        self._reloj.start(1000)

    def _actualizar(self):
        self._ahora = QDateTime.currentDateTime()
        for i, f in enumerate(FORMATOS_FECHA_HORA):
            self.lista.item(i).setText(self.locale.toString(self._ahora, f))

    def texto(self):
        """Texto del formato elegido, con la hora del momento de aceptar."""
        fila = max(0, self.lista.currentRow())
        VentanaFechaHora.ultimo = fila
        return self.locale.toString(QDateTime.currentDateTime(),
                                    FORMATOS_FECHA_HORA[fila])


# ---------------------------------------------------------------------------
# Buscar / Buscar y reemplazar
# ---------------------------------------------------------------------------
COLOR_COINCIDENCIA = "#FFE98A"
COLOR_COINCIDENCIA_ACTUAL = "#FFA94D"
MAX_RESALTADAS = 5000


class VentanaBuscar(QDialog):
    """Ventana no modal que resalta todas las coincidencias y muestra "n de N"."""

    def __init__(self, editor, parent=None):
        super().__init__(parent)
        self.editor = editor
        self.setWindowTitle("Buscar")
        self.setModal(False)
        lay = QVBoxLayout(self)
        self.pestanas = QTabBar()
        self.pestanas.addTab("Buscar")
        self.pestanas.addTab("Buscar y reemplazar")
        self.pestanas.setExpanding(False)
        self.pestanas.setDrawBase(False)
        self.pestanas.currentChanged.connect(self._cambiar_modo)
        lay.addWidget(self.pestanas)

        form = QGridLayout()
        form.setColumnStretch(1, 1)
        self.campo_buscar = QLineEdit()
        self.campo_buscar.setPlaceholderText("Texto a buscar")
        self.campo_buscar.setClearButtonEnabled(True)
        self.campo_buscar.textChanged.connect(self._programar)
        self.campo_buscar.returnPressed.connect(self._enter)
        self.campo_reemplazar = QLineEdit()
        self.campo_reemplazar.setPlaceholderText("Reemplazar con")
        self.campo_reemplazar.setClearButtonEnabled(True)
        self.campo_reemplazar.returnPressed.connect(self.reemplazar)
        self.etiqueta_reemplazar = QLabel("Reemplazar")
        form.addWidget(QLabel("Buscar"), 0, 0)
        form.addWidget(self.campo_buscar, 0, 1)
        form.addWidget(self.etiqueta_reemplazar, 1, 0)
        form.addWidget(self.campo_reemplazar, 1, 1)
        form.setColumnMinimumWidth(0, self.etiqueta_reemplazar.sizeHint().width())
        lay.addLayout(form)

        opciones = QHBoxLayout()
        self.chk_mayus = QCheckBox("Coincidir mayúsculas")
        self.chk_palabra = QCheckBox("Palabras completas")
        for chk in (self.chk_mayus, self.chk_palabra):
            chk.toggled.connect(self._programar)
            opciones.addWidget(chk)
        opciones.addStretch()
        lay.addLayout(opciones)

        self.estado = QLabel(" ")
        TEMA.estilo(self.estado, lambda t: f"color: {t['texto_sec']};")
        lay.addWidget(self.estado)

        botones = QHBoxLayout()
        self.btn_anterior = QPushButton("Anterior")
        self.btn_anterior.clicked.connect(lambda: self.buscar(atras=True))
        self.btn_siguiente = QPushButton("Siguiente")
        self.btn_siguiente.clicked.connect(lambda: self.buscar())
        self.btn_reemplazar = QPushButton("Reemplazar")
        self.btn_reemplazar.clicked.connect(self.reemplazar)
        self.btn_todo = QPushButton("Reemplazar todo")
        self.btn_todo.clicked.connect(self.reemplazar_todo)
        cerrar = QPushButton("Cerrar")
        cerrar.clicked.connect(self.close)
        for b in (self.btn_anterior, self.btn_siguiente, self.btn_reemplazar, self.btn_todo):
            b.setAutoDefault(False)
            botones.addWidget(b)
        botones.addStretch()
        cerrar.setAutoDefault(False)
        botones.addWidget(cerrar)
        lay.addStretch()
        lay.addLayout(botones)

        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(150)
        self._timer.timeout.connect(self.refrescar)
        editor.document().contentsChanged.connect(self._doc_cambiado)
        editor.cursorPositionChanged.connect(self._mostrar_estado)
        self._coincidencias = []
        self._extra = ""
        self._cambiar_modo(1)
        tam = self.sizeHint().expandedTo(QSize(460, 0))
        self.setMinimumSize(tam)
        self.resize(tam)
        self._cambiar_modo(0)

    def abrir(self, reemplazar=False):
        self.pestanas.setCurrentIndex(1 if reemplazar else 0)
        sel = self.editor.textCursor().selectedText()
        if sel and "\u2029" not in sel and len(sel) <= 200:
            self.campo_buscar.setText(sel)
        self.show()
        self.raise_()
        self.activateWindow()
        self.campo_buscar.setFocus()
        self.campo_buscar.selectAll()
        self.refrescar()

    def _cambiar_modo(self, i):
        for w in (self.etiqueta_reemplazar, self.campo_reemplazar,
                  self.btn_reemplazar, self.btn_todo):
            w.setVisible(i == 1)
        self.setWindowTitle("Buscar y reemplazar" if i == 1 else "Buscar")

    def _banderas(self, atras=False):
        f = QTextDocument.FindFlag(0)
        if self.chk_mayus.isChecked():
            f |= QTextDocument.FindCaseSensitively
        if self.chk_palabra.isChecked():
            f |= QTextDocument.FindWholeWords
        if atras:
            f |= QTextDocument.FindBackward
        return f

    def _enter(self):
        self.buscar(atras=bool(QApplication.keyboardModifiers() & Qt.ShiftModifier))

    def buscar(self, atras=False):
        texto = self.campo_buscar.text()
        if not texto:
            return False
        self._extra = ""
        doc = self.editor.document()
        hallado = doc.find(texto, self.editor.textCursor(), self._banderas(atras))
        if hallado.isNull():                 # da la vuelta al documento
            inicio = QTextCursor(doc)
            if atras:
                inicio.movePosition(QTextCursor.End)
            hallado = doc.find(texto, inicio, self._banderas(atras))
            if not hallado.isNull():
                self._extra = "Se continuó desde el final" if atras \
                    else "Se continuó desde el principio"
        if hallado.isNull():
            self.estado.setText("Sin coincidencias")
            QApplication.beep()
            return False
        self.editor.setTextCursor(hallado)
        self._pintar()
        return True

    def reemplazar(self):
        if not self.campo_buscar.text():
            return
        self._recalcular()
        c = self.editor.textCursor()
        if c.hasSelection() and (c.selectionStart(), c.selectionEnd()) in set(self._coincidencias):
            c.insertText(self.campo_reemplazar.text())
            self.editor.setTextCursor(c)
            self._recalcular()
        self.buscar()

    def reemplazar_todo(self):
        if not self.campo_buscar.text():
            return
        self._recalcular()
        if not self._coincidencias:
            self.estado.setText("Sin coincidencias")
            return
        nuevo = self.campo_reemplazar.text()
        c = QTextCursor(self.editor.document())
        c.beginEditBlock()                     # un solo paso de deshacer
        for inicio, fin in reversed(self._coincidencias):
            c.setPosition(inicio)
            c.setPosition(fin, QTextCursor.KeepAnchor)
            c.insertText(nuevo)
        c.endEditBlock()
        n = len(self._coincidencias)
        self._recalcular()
        self._pintar()
        self.estado.setText("Se reemplazó 1 coincidencia" if n == 1
                            else f"Se reemplazaron {n} coincidencias")

    def _programar(self, *_):
        self._extra = ""
        self._timer.start()

    def _doc_cambiado(self):
        if self.isVisible():
            self._timer.start()

    def _recalcular(self):
        self._coincidencias = []
        texto = self.campo_buscar.text()
        if not texto:
            return
        doc = self.editor.document()
        c = QTextCursor(doc)
        while True:
            c = doc.find(texto, c, self._banderas())
            if c.isNull() or c.selectionStart() == c.selectionEnd():
                break
            self._coincidencias.append((c.selectionStart(), c.selectionEnd()))

    def refrescar(self):
        if self.isVisible():
            self._recalcular()
            self._pintar()

    def _pintar(self):
        doc = self.editor.document()
        actual = self.editor.textCursor()
        sels = []
        for inicio, fin in self._coincidencias[:MAX_RESALTADAS]:
            sel = QTextEdit.ExtraSelection()
            sel.cursor = QTextCursor(doc)
            sel.cursor.setPosition(inicio)
            sel.cursor.setPosition(fin, QTextCursor.KeepAnchor)
            es_actual = (inicio, fin) == (actual.selectionStart(), actual.selectionEnd())
            sel.format.setBackground(QColor(COLOR_COINCIDENCIA_ACTUAL if es_actual
                                            else COLOR_COINCIDENCIA))
            sels.append(sel)
        self.editor.setExtraSelections(sels)
        self._mostrar_estado()

    def _mostrar_estado(self):
        if not self.isVisible():
            return
        if not self.campo_buscar.text():
            self.estado.setText(" ")
            return
        n = len(self._coincidencias)
        if n == 0:
            self.estado.setText("Sin coincidencias")
            return
        c = self.editor.textCursor()
        try:
            i = self._coincidencias.index((c.selectionStart(), c.selectionEnd())) + 1
            texto = f"{i} de {n} coincidencias"
        except ValueError:
            texto = "1 coincidencia" if n == 1 else f"{n} coincidencias"
        self.estado.setText(f"{texto} · {self._extra}" if self._extra else texto)

    def hideEvent(self, e):
        self.editor.setExtraSelections([])
        super().hideEvent(e)

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape:
            self.close()
            return
        super().keyPressEvent(e)


# ---------------------------------------------------------------------------
# Navegador: esquema de títulos (párrafos con nivel de esquema)
# ---------------------------------------------------------------------------
class PanelNavegador(QDockWidget):
    def __init__(self, editor, vista, parent=None):
        super().__init__("Navegador", parent)
        self.editor = editor
        self.vista = vista
        self.setObjectName("navegador")
        self.setFeatures(QDockWidget.DockWidgetClosable)
        contenido = QWidget()
        lay = QVBoxLayout(contenido)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.setSpacing(6)
        self.filtro = QLineEdit()
        self.filtro.setPlaceholderText("Filtrar títulos")
        self.filtro.setClearButtonEnabled(True)
        self.filtro.textChanged.connect(self._filtrar)
        lay.addWidget(self.filtro)
        self.arbol = QTreeWidget()
        self.arbol.setHeaderHidden(True)
        self.arbol.setIndentation(14)
        self.arbol.setUniformRowHeights(True)
        self.arbol.setExpandsOnDoubleClick(False)
        self.arbol.itemClicked.connect(self._ir)
        self.arbol.itemActivated.connect(self._ir)
        lay.addWidget(self.arbol, 1)
        self.vacio = QLabel("Aún no hay títulos.\nUsa estilos con nivel de esquema "
                            "(Encabezado 1, 2, 3…) para verlos aquí.")
        self.vacio.setWordWrap(True)
        TEMA.estilo(self.vacio, lambda t: f"color: {t['texto_suave']}; padding: 8px;")
        lay.addWidget(self.vacio)
        self.setWidget(contenido)
        self.setMinimumWidth(200)
        self._titulos = []
        self._items = []
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(300)
        self._timer.timeout.connect(self.refrescar)
        editor.document().contentsChanged.connect(self._programar)
        editor.cursorPositionChanged.connect(self._marcar_actual)
        self.visibilityChanged.connect(lambda v: v and self.refrescar())

    def _programar(self):
        if self.isVisible():
            self._timer.start()

    def _leer_titulos(self):
        titulos = []
        b = self.editor.document().begin()
        while b.isValid():
            nivel = b.blockFormat().headingLevel()
            if nivel > 0:
                texto = " ".join(b.text().replace("\u2028", " ")
                                 .replace(CARACTER_OBJETO, " ").split())
                titulos.append((b.position(), nivel, texto))
            b = b.next()
        return titulos

    def refrescar(self):
        self._timer.stop()
        titulos = self._leer_titulos()
        if [t[1:] for t in titulos] == [t[1:] for t in self._titulos]:
            self._titulos = titulos
            self._marcar_actual()
            return
        plegados = {(self._titulos[i][1], it.text(0)) for i, it in enumerate(self._items)
                    if i < len(self._titulos) and it.childCount() and not it.isExpanded()}
        self._titulos = titulos
        self.arbol.clear()
        self._items = []
        pila = []
        negrita = QFont(self.arbol.font())
        negrita.setBold(True)
        doc = self.editor.document()
        for i, (pos, nivel, texto) in enumerate(titulos):
            while pila and pila[-1][0] >= nivel:
                pila.pop()
            it = QTreeWidgetItem([texto or "(sin texto)"])
            it.setData(0, Qt.UserRole, i)
            v = doc.findBlock(pos).blockFormat().property(PROP_ESTILO)
            estilo = GESTOR.nombre(v) if isinstance(v, str) else f"Nivel {nivel}"
            it.setToolTip(0, f"{estilo}: {texto}")
            if nivel == 1:
                it.setFont(0, negrita)
            if not texto:
                it.setForeground(0, QColor("#9A9A9A"))
            if pila:
                pila[-1][1].addChild(it)
            else:
                self.arbol.addTopLevelItem(it)
            pila.append((nivel, it))
            self._items.append(it)
        self.arbol.expandAll()
        for i, it in enumerate(self._items):
            if (titulos[i][1], it.text(0)) in plegados:
                it.setExpanded(False)
        self.arbol.setVisible(bool(titulos))
        self.vacio.setVisible(not titulos)
        self._filtrar(self.filtro.text())
        self._marcar_actual()

    def _ir(self, it, _col=0):
        if self._timer.isActive():
            self.refrescar()
        i = it.data(0, Qt.UserRole)
        if i is None or i >= len(self._titulos):
            return
        bloque = self.editor.document().findBlock(self._titulos[i][0])
        if not bloque.isValid():
            return
        self.editor.setTextCursor(QTextCursor(bloque))
        QApplication.processEvents()
        r = self.editor.cursorRect()
        arriba = self.vista.mapFromScene(
            self.vista.proxy.mapToScene(QPointF(r.topLeft()))).y()
        barra = self.vista.verticalScrollBar()
        barra.setValue(barra.value() + arriba - 48)
        self.editor.setFocus()

    def _marcar_actual(self):
        if not self.isVisible() or not self._titulos:
            return
        pos = self.editor.textCursor().position()
        i = bisect.bisect_right([t[0] for t in self._titulos], pos) - 1
        self.arbol.blockSignals(True)
        if 0 <= i < len(self._items):
            destino = self._items[i]
            padre = destino.parent()
            while padre is not None:
                if not padre.isExpanded():
                    destino = padre
                padre = padre.parent()
            self.arbol.setCurrentItem(destino)
            self.arbol.scrollToItem(destino)
        else:
            self.arbol.setCurrentItem(None)
            self.arbol.clearSelection()
        self.arbol.blockSignals(False)

    def _filtrar(self, texto):
        texto = texto.strip().lower()

        def mostrar(it):
            hijo = False
            for j in range(it.childCount()):
                hijo |= mostrar(it.child(j))
            coincide = not texto or texto in it.text(0).lower()
            it.setHidden(not (coincide or hijo))
            return coincide or hijo

        for j in range(self.arbol.topLevelItemCount()):
            mostrar(self.arbol.topLevelItem(j))
        if texto:
            self.arbol.expandAll()


# ---------------------------------------------------------------------------
# Panel de comentarios
# ---------------------------------------------------------------------------
class TarjetaComentario(QFrame):
    elegida = Signal(int, bool)          # id, ¿llevar el foco al texto?
    textoCambiado = Signal(int, str)
    colorCambiado = Signal(int, QColor)
    eliminarPedido = Signal(int)

    def __init__(self, id_, parent=None):
        super().__init__(parent)
        self.id_ = id_
        self.color = QColor(COLORES_COMENTARIO[0])
        self.activa = False
        self.cita = ""
        self.setObjectName("tarjeta")
        self.setCursor(Qt.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 8, 8, 8)
        lay.setSpacing(4)

        cabecera = QHBoxLayout()
        self.titulo = QLabel()
        cabecera.addWidget(self.titulo)
        cabecera.addStretch()
        self.btn_color = QToolButton()
        self.btn_color.setCursor(Qt.ArrowCursor)
        self.btn_color.setPopupMode(QToolButton.InstantPopup)
        self.btn_color.setFixedSize(24, 22)
        menu = QMenu(self.btn_color)
        cont = QWidget()
        grid = QGridLayout(cont)
        grid.setContentsMargins(8, 6, 8, 6)
        grid.setSpacing(4)
        for k, hexa in enumerate(COLORES_COMENTARIO):
            b = QToolButton()
            b.setFixedSize(20, 20)
            b.setCursor(Qt.PointingHandCursor)
            TEMA.estilo(b, lambda t, hexa=hexa: (
                f"QToolButton {{ background:{hexa}; border:1px solid {t['campo_borde']};"
                f" border-radius:3px; padding:0; }}"
                f"QToolButton:hover {{ border:2px solid {t['texto']}; }}"))
            b.clicked.connect(lambda _=False, c=hexa: self._elegir_color(QColor(c)))
            grid.addWidget(b, k // 5, k % 5)
        wa = QWidgetAction(menu)
        wa.setDefaultWidget(cont)
        menu.addAction(wa)
        menu.addSeparator()
        menu.addAction("Más colores…").triggered.connect(self._dialogo_color)
        self.menu_color = menu
        self.btn_color.setMenu(menu)
        cabecera.addWidget(self.btn_color)
        btn = QToolButton()
        btn.setIcon(icono_archivo("eliminar"))
        btn.setIconSize(QSize(16, 16))
        btn.setFixedSize(24, 22)
        btn.setAutoRaise(True)
        btn.setCursor(Qt.ArrowCursor)
        btn.clicked.connect(lambda: self.eliminarPedido.emit(self.id_))
        cabecera.addWidget(btn)
        lay.addLayout(cabecera)

        self.etiqueta_cita = QLabel()
        self.etiqueta_cita.setWordWrap(False)
        self.etiqueta_cita.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.etiqueta_cita.installEventFilter(self)
        lay.addWidget(self.etiqueta_cita)

        self.texto = QTextEdit()
        self.texto.setAcceptRichText(False)
        self.texto.setTabChangesFocus(True)
        self.texto.setPlaceholderText("Escribe un comentario…")
        self.texto.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.texto.setCursor(Qt.IBeamCursor)
        self.texto.textChanged.connect(
            lambda: self.textoCambiado.emit(self.id_, self.texto.toPlainText()))
        self.texto.document().documentLayout().documentSizeChanged.connect(self._ajustar_alto)
        self.texto.installEventFilter(self)
        lay.addWidget(self.texto)
        TEMA.cambiado.connect(self._reestilizar)
        self._reestilizar()
        self._ajustar_alto()

    def eventFilter(self, obj, e):
        if obj is getattr(self, "texto", None) and e.type() == QEvent.FocusIn:
            self.elegida.emit(self.id_, False)
        elif obj is getattr(self, "etiqueta_cita", None) and e.type() == QEvent.Resize:
            self._mostrar_cita()
        return False

    def _mostrar_cita(self):
        fm = self.etiqueta_cita.fontMetrics()
        ancho = self.etiqueta_cita.contentsRect().width() - 24
        corto = fm.elidedText(self.cita, Qt.ElideRight, max(20, ancho))
        self.etiqueta_cita.setText(f"«{corto}»")

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.elegida.emit(self.id_, True)
        super().mousePressEvent(e)

    def _elegir_color(self, color):
        self.menu_color.hide()
        self.colorCambiado.emit(self.id_, color)

    def _dialogo_color(self):
        c = QColorDialog.getColor(self.color, self, "Color del comentario")
        if c.isValid():
            self._elegir_color(c)

    def _ajustar_alto(self, *_):
        linea = self.texto.fontMetrics().lineSpacing()
        alto = max(self.texto.document().size().height(), 2 * linea + 8)
        self.texto.setFixedHeight(round(alto) + 4)

    def refrescar(self, numero, datos, cita, color):
        self.titulo.setText(f"Comentario {numero}")
        autor, fecha = datos.get("autor", ""), fecha_legible(datos.get("fecha", ""))
        self.titulo.setToolTip(" · ".join(x for x in (autor, fecha) if x))
        if cita != self.cita:
            self.cita = cita
            self._mostrar_cita()
        if QColor(color) != self.color:
            self.color = QColor(color)
            self._reestilizar()
        if self.texto.toPlainText() != datos.get("texto", ""):
            self.texto.blockSignals(True)
            self.texto.setPlainText(datos.get("texto", ""))
            self.texto.blockSignals(False)

    def poner_activa(self, activa):
        if activa != self.activa:
            self.activa = activa
            self._reestilizar()

    def _reestilizar(self):
        t = TEMA.t
        fuerte = self.color.darker(118).name()
        borde = f"2px solid {fuerte}" if self.activa else f"1px solid {t['borde']}"
        self.setStyleSheet(
            f"QFrame#tarjeta {{ background: {t['campo']}; border: {borde};"
            f" border-left: 5px solid {self.color.name()}; border-radius: 6px; }}")
        self.titulo.setStyleSheet(f"font-weight: bold; color: {t['texto']};")
        self.texto.setStyleSheet(
            f"QTextEdit {{ border: 1px solid {t['borde']}; border-radius: 4px;"
            f" background: {t['campo']}; color: {t['texto']}; }}")
        self.btn_color.setStyleSheet(
            f"QToolButton {{ background: {self.color.name()}; border: 1px solid #9A9A9A;"
            " border-radius: 4px; } QToolButton::menu-indicator { image: none; width: 0; }")
        self.etiqueta_cita.setStyleSheet(
            f"color: #5A5A5A; font-style: italic; background: {self.color.name()};"
            " border-radius: 3px; padding: 2px 5px;")


class PanelComentarios(QDockWidget):
    def __init__(self, editor, parent=None):
        super().__init__("Comentarios", parent)
        self.editor = editor
        self.gestor = editor.comentarios
        self.setObjectName("comentarios")
        self.setFeatures(QDockWidget.DockWidgetClosable)
        contenido = QWidget()
        lay = QVBoxLayout(contenido)
        lay.setContentsMargins(6, 6, 6, 6)
        self.area = QScrollArea()
        self.area.setWidgetResizable(True)
        self.area.setFrameShape(QFrame.NoFrame)
        self.area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.lista = QWidget()
        self.lay_lista = QVBoxLayout(self.lista)
        self.lay_lista.setContentsMargins(0, 0, 2, 0)
        self.lay_lista.setSpacing(8)
        self.lay_lista.addStretch()
        self.area.setWidget(self.lista)
        lay.addWidget(self.area, 1)
        self.vacio = QLabel("Aún no hay comentarios.\nSelecciona texto y pulsa "
                            "Nuevo comentario (Ctrl+Alt+M).")
        self.vacio.setWordWrap(True)
        self.vacio.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        TEMA.estilo(self.vacio, lambda t: f"color: {t['texto_suave']}; padding: 8px;")
        lay.addWidget(self.vacio, 1)
        self.setWidget(contenido)
        self.setMinimumWidth(240)
        self.tarjetas = {}
        self._orden = []
        self.gestor.listaCambiada.connect(self.refrescar)
        self.gestor.activoCambiado.connect(self._marcar)
        self.refrescar()

    def refrescar(self):
        g = self.gestor
        for i in list(self.tarjetas):
            if i not in g.rangos:
                t = self.tarjetas.pop(i)
                self.lay_lista.removeWidget(t)
                t.deleteLater()
        for i in g.orden:
            if i not in self.tarjetas:
                t = TarjetaComentario(i)
                t.elegida.connect(self._ir)
                t.textoCambiado.connect(g.poner_texto)
                t.colorCambiado.connect(g.poner_color)
                t.eliminarPedido.connect(g.eliminar)
                self.tarjetas[i] = t
        if self._orden != g.orden:
            for t in self.tarjetas.values():
                self.lay_lista.removeWidget(t)
            for pos, i in enumerate(g.orden):
                self.lay_lista.insertWidget(pos, self.tarjetas[i])
            self._orden = list(g.orden)
        for n, i in enumerate(g.orden, 1):
            self.tarjetas[i].refrescar(n, g.datos[i], g.extracto(i), g.color(i))
        vacio = not g.orden
        self.area.setVisible(not vacio)
        self.vacio.setVisible(vacio)
        self._marcar(g.activo)

    def _ir(self, i, foco_editor):
        self.gestor.ir_a(i)
        if foco_editor:
            self.editor.setFocus()

    def _marcar(self, activo):
        for i, t in self.tarjetas.items():
            t.poner_activa(i == activo)
        if activo in self.tarjetas and self.isVisible():
            self.area.ensureWidgetVisible(self.tarjetas[activo], 0, 20)

    def enfocar(self, i):
        self.show()
        self.refrescar()
        t = self.tarjetas.get(i)
        if t is not None:
            QTimer.singleShot(0, lambda: (self.area.ensureWidgetVisible(t),
                                          t.texto.setFocus()))


class CentroSuperpuesto(QWidget):
    """La hoja con su barra a la derecha; los paneles laterales se dibujan
    encima de la hoja (no cambian su tamaño ni su posición)."""
    ANCHO_IZQ = 240
    ANCHO_DER = 280

    def __init__(self, vista, barra, izquierdo, derecho, parent=None):
        super().__init__(parent)
        self.vista, self.izquierdo, self.derecho = vista, izquierdo, derecho
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(vista, 1)
        lay.addWidget(barra)
        for panel in (izquierdo, derecho):
            panel.setParent(self)
            panel.setAutoFillBackground(True)
            TEMA.estilo(panel, lambda t: (
                f"QDockWidget > QWidget {{ background: {t['barra']};"
                f" border: 1px solid {t['borde']}; }}"))
            panel.installEventFilter(self)
            panel.raise_()

    def eventFilter(self, obj, e):
        if e.type() == QEvent.Show:
            QTimer.singleShot(0, self._colocar)
        return False

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._colocar()

    def _colocar(self):
        alto = self.height()
        ancho_vista = self.vista.width()
        self.izquierdo.setGeometry(0, 0, min(self.ANCHO_IZQ, ancho_vista), alto)
        an = min(self.ANCHO_DER, ancho_vista)
        self.derecho.setGeometry(ancho_vista - an, 0, an, alto)
        self.izquierdo.raise_()
        self.derecho.raise_()


# ---------------------------------------------------------------------------
# Ventana principal
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Configuración del editor en config.ini: estilos, tema, colores, página y
# tamaño de la ventana. Se guarda sola al aceptar los diálogos y al cerrar, y
# se lee al abrir el programa. Un valor mal escrito se ignora (queda el de
# fábrica); si falta el archivo se crea al guardar el primer cambio.
# ---------------------------------------------------------------------------
NOMBRE_CONFIG = "config.ini"
CLAVES_VISTA = ("tema", "tema_interfaz", "tema_fondo", "color_hoja",
                "espacio_superior")
CLAVES_PAGINA = ("ancho", "alto", "superior", "inferior", "izquierdo", "derecho")
_ALINEACION_POR_NOMBRE = {"izquierda": Qt.AlignLeft, "centro": Qt.AlignHCenter,
                          "derecha": Qt.AlignRight, "justificado": Qt.AlignJustify}
_NOMBRE_POR_ALINEACION = {v: k for k, v in _ALINEACION_POR_NOMBRE.items()}
_RANGOS_ESTILO = {                       # campo numérico: (mínimo, máximo)
    "font_size": (4, 200), "line_spacing": (50, 500),
    "space_before": (0, 1000), "space_after": (0, 1000),
    "indent_left": (-1000, 1000), "indent_right": (-1000, 1000),
    "first_line_indent": (-1000, 1000), "letter_spacing": (-50, 100),
}
_ruta_config_cache = None


def _carpeta_escribible(carpeta):
    try:
        prueba = os.path.join(carpeta, f".swordbook_{os.getpid()}.tmp")
        with open(prueba, "w"):
            pass
        os.remove(prueba)
        return True
    except OSError:
        return False


def ruta_config():
    """config.ini junto al programa (versión portable). Si esa carpeta no se
    puede escribir (Program Files, el paquete .app de macOS) va a la carpeta
    de configuración del usuario."""
    global _ruta_config_cache
    if _ruta_config_cache:
        return _ruta_config_cache
    if getattr(sys, "frozen", False):
        carpeta = os.path.dirname(sys.executable)
        en_paquete = sys.platform == "darwin" and ".app/Contents/" in sys.executable
    else:
        carpeta = os.path.dirname(os.path.abspath(__file__))
        en_paquete = False
    junto = os.path.join(carpeta, NOMBRE_CONFIG)
    if os.path.isfile(junto) or (not en_paquete and _carpeta_escribible(carpeta)):
        ruta = junto
    else:
        datos = QStandardPaths.writableLocation(QStandardPaths.AppConfigLocation)
        try:
            os.makedirs(datos, exist_ok=True)
            ruta = os.path.join(datos, NOMBRE_CONFIG)
        except (OSError, TypeError):
            ruta = junto
    _ruta_config_cache = ruta
    return ruta


def _leer_ini():
    cp = configparser.ConfigParser(interpolation=None)
    try:
        cp.read(ruta_config(), encoding="utf-8")
    except (OSError, UnicodeDecodeError, configparser.Error):
        cp = configparser.ConfigParser(interpolation=None)    # archivo dañado
    return cp


def _escribir_ini(cp):
    ruta = ruta_config()
    try:
        with open(ruta + ".tmp", "w", encoding="utf-8") as f:
            f.write("# Configuración de SWordBook. Puedes editarla con el "
                    "programa cerrado;\n# borra este archivo para volver a "
                    "los valores de fábrica.\n\n")
            cp.write(f)
        os.replace(ruta + ".tmp", ruta)
    except OSError:
        pass                      # sin permisos o disco lleno: no es grave


def _texto_valor(v):
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, float):
        v = round(v, 3)
        return str(int(v)) if v == int(v) else str(v)
    return str(v)


def _ini_numero(texto, minimo, maximo):
    v = float(str(texto).replace(",", "."))
    if not math.isfinite(v) or not minimo <= v <= maximo:
        raise ValueError(texto)
    return int(v) if v == int(v) else v


def _ini_color(texto):
    c = QColor(str(texto).strip())
    if not c.isValid():
        raise ValueError(texto)
    return c.name().upper()


# -- guardar ---------------------------------------------------------------
def guardar_vista(config, con_pagina=False):
    """Tema, colores y espacio superior; con `con_pagina`, también el tamaño
    y los márgenes de la página (valores para documentos nuevos)."""
    cp = _leer_ini()
    cp["vista"] = {k: _texto_valor(config[k]) for k in CLAVES_VISTA if k in config}
    if con_pagina:
        cp["pagina"] = {k: _texto_valor(config[k]) for k in CLAVES_PAGINA}
    _escribir_ini(cp)


def guardar_estilos(gestor):
    cp = _leer_ini()
    for sec in [s for s in cp.sections() if s.startswith("estilo:")]:
        cp.remove_section(sec)
    cp["estilos"] = {"orden": ", ".join(gestor.orden)}
    for id_ in gestor.orden:
        e = gestor.datos[id_]
        sec = {"nombre": e["nombre"], "visible": _texto_valor(bool(e["visible"])),
               "siguiente": e["siguiente"], "nivel": str(int(e["nivel"]))}
        for campo in CAMPOS_ESTILO:
            v = e["propio"][campo]
            sec[campo] = (_NOMBRE_POR_ALINEACION.get(v, "izquierda")
                          if campo == "alignment" else _texto_valor(v))
        cp[f"estilo:{id_}"] = sec
    _escribir_ini(cp)


def guardar_ventana(ventana):
    cp = _leer_ini()
    maximizada = ventana.isMaximized() or ventana.isFullScreen()
    g = ventana.normalGeometry() if maximizada else ventana.geometry()
    cp["ventana"] = {"ancho": str(g.width()), "alto": str(g.height()),
                     "maximizada": _texto_valor(maximizada)}
    _escribir_ini(cp)


# -- cargar ----------------------------------------------------------------
def _cargar_estilos(cp):
    datos, orden_defecto = _predeterminados()
    ids = [i.strip() for i in cp.get("estilos", "orden", fallback="").split(",")
           if i.strip()]
    orden = []
    for id_ in ids:
        sec = f"estilo:{id_}"
        if id_ in orden or not cp.has_section(sec):
            continue
        s = cp[sec]
        e = datos.get(id_) or {
            "id": id_, "nombre": id_, "visible": True, "base": None,
            "siguiente": id_, "nivel": 0, "propio": dict(BODY), "fijo": False}
        if "nombre" in s:
            e["nombre"] = s["nombre"].strip()
        try:
            e["visible"] = s.getboolean("visible", fallback=e["visible"])
        except ValueError:
            pass
        if s.get("siguiente", "").strip():
            e["siguiente"] = s["siguiente"].strip()
        try:
            e["nivel"] = int(_ini_numero(s.get("nivel", e["nivel"]), 0, 6))
        except ValueError:
            pass
        for campo in CAMPOS_ESTILO:
            if campo not in s:
                continue
            try:
                if campo == "alignment":
                    e["propio"][campo] = _ALINEACION_POR_NOMBRE[s[campo].strip().lower()]
                elif campo in ("bold", "italic"):
                    e["propio"][campo] = s.getboolean(campo)
                elif campo == "font_color":
                    e["propio"][campo] = _ini_color(s[campo])
                elif campo == "font_family":
                    if s[campo].strip():
                        e["propio"][campo] = s[campo].strip()
                else:
                    e["propio"][campo] = _ini_numero(s[campo], *_RANGOS_ESTILO[campo])
            except (ValueError, KeyError):
                pass
        datos[id_] = e
        orden.append(id_)
    for id_ in orden_defecto:              # los de fábrica nunca faltan
        if id_ not in orden:
            orden.append(id_)
    for e in datos.values():               # «siguiente» debe existir
        if e["siguiente"] not in datos:
            e["siguiente"] = e["id"]
    GESTOR.datos, GESTOR.orden = datos, orden


def cargar_config():
    """Lee config.ini y lo deja puesto (estilos y valores iniciales de la
    vista y la página). Se llama antes de crear la ventana. Devuelve el
    tamaño guardado de la ventana ({} si no hay)."""
    cp = _leer_ini()
    if cp.has_section("estilos"):
        _cargar_estilos(cp)
    v = cp["vista"] if cp.has_section("vista") else {}
    if v.get("tema", "").strip() in dict(NOMBRES_TEMA):
        PAGINA["tema"] = v["tema"].strip()
    for k in ("tema_interfaz", "tema_fondo", "color_hoja"):
        try:
            PAGINA[k] = _ini_color(v[k])
        except (KeyError, ValueError):
            pass
    try:
        PAGINA["espacio_superior"] = _ini_numero(v["espacio_superior"], 0, 200)
    except (KeyError, ValueError):
        pass
    if cp.has_section("pagina"):
        try:
            p = {k: _ini_numero(cp["pagina"][k], 0, 300) for k in CLAVES_PAGINA}
            if (p["ancho"] >= 5 and p["alto"] >= 5          # como Configurar página
                    and p["izquierdo"] + p["derecho"] <= p["ancho"] - 2
                    and p["superior"] + p["inferior"] <= p["alto"] - 2):
                PAGINA.update(p)
        except (KeyError, ValueError):
            pass
    ventana = {}
    if cp.has_section("ventana"):
        try:
            ventana["ancho"] = int(_ini_numero(cp["ventana"]["ancho"], 400, 10000))
            ventana["alto"] = int(_ini_numero(cp["ventana"]["alto"], 300, 10000))
            ventana["maximizada"] = cp["ventana"].getboolean("maximizada",
                                                             fallback=False)
        except (KeyError, ValueError):
            ventana = {}
    return ventana


class Ventana(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ruta = None               # archivo abierto (None = sin guardar)
        self._carpeta_imagen = None
        self.resize(1100, 700)
        self.editor = Editor()
        self.vista = VistaSinPaginas(self.editor)

        # Centro: la hoja ocupa todo el espacio; el navegador y los
        # comentarios flotan encima (a izquierda y derecha) sin moverla
        self.navegador = PanelNavegador(self.editor, self.vista)
        self.panel_comentarios = PanelComentarios(self.editor)

        # Barra de desplazamiento de la hoja: siempre en el borde derecho
        self.barra_v = QScrollBar(Qt.Vertical)
        interna = self.vista.verticalScrollBar()

        def copiar_rango(minimo, maximo):
            self.barra_v.setRange(minimo, maximo)
            self.barra_v.setPageStep(interna.pageStep())
            self.barra_v.setSingleStep(interna.singleStep())

        interna.rangeChanged.connect(copiar_rango)
        interna.valueChanged.connect(self.barra_v.setValue)
        self.barra_v.valueChanged.connect(interna.setValue)
        copiar_rango(interna.minimum(), interna.maximum())
        self.barra_v.setValue(interna.value())

        self.setCentralWidget(CentroSuperpuesto(
            self.vista, self.barra_v, self.navegador, self.panel_comentarios))
        self.navegador.hide()
        self.panel_comentarios.hide()

        # Acciones de la barra (se reparten en las filas de abajo)
        self.acciones = {}
        for icono, texto, atajo, slot in (
                ("nuevo", "Nuevo", QKeySequence.New, self.nuevo),
                ("abrir", "Abrir", QKeySequence.Open, self.abrir),
                ("guardar", "Guardar", QKeySequence.Save, self.guardar),
                ("guardar_como", "Guardar como", QKeySequence("Ctrl+Shift+S"),
                 self.guardar_como),
                # Los atajos de Cortar, Copiar y Pegar los maneja el editor
                ("cortar", "Cortar", None, self.cortar),
                ("copiar", "Copiar", None, self.copiar),
                ("pegar", "Pegar", None, self.pegar),
                # Abren una ventana: se enlazan abajo (EnlaceVentana)
                ("imagen", "Insertar imagen", None, self.insertar_imagen),
                ("fecha", "Insertar fecha y hora", None, None),
                ("buscar", "Buscar y reemplazar", None, None),
                ("navegador", "Navegador", QKeySequence("F5"), self.alternar_navegador),
                ("comentario", "Nuevo comentario", QKeySequence("Ctrl+Alt+M"),
                 self.nuevo_comentario),
                ("comentarios", "Panel de comentarios", QKeySequence("F6"),
                 self.alternar_comentarios)):
            a = QAction(icono_archivo(icono), texto, self)
            if atajo is not None:
                a.setShortcut(atajo)
                a.setShortcutContext(Qt.WindowShortcut)
            if slot is not None:
                a.triggered.connect(slot)
            self.acciones[icono] = a
        self.acciones["pagina"] = QAction(icono_pagina(), "Configurar página", self)

        # Botones que abren una ventana: marcados mientras está abierta,
        # otra pulsación la cierra y la ventana lleva el icono del botón
        enfocar = self.editor.setFocus
        self.enlace_pagina = EnlaceVentana(
            self.acciones["pagina"], lambda: VentanaPagina(self.vista.config, self),
            al_aceptar=self._pagina_aceptada,
            al_cerrar=enfocar)
        self.enlace_fecha = EnlaceVentana(
            self.acciones["fecha"], lambda: VentanaFechaHora(self),
            al_aceptar=lambda d: self.editor.insertar_texto(d.texto()),
            al_cerrar=enfocar)
        self.enlace_buscar = EnlaceVentana(
            self.acciones["buscar"], lambda: VentanaBuscar(self.editor, self),
            reutilizar=True, mostrar=lambda v, reemplazar=False: v.abrir(reemplazar))
        # Ctrl+F abre (o enfoca) Buscar; no la cierra como el botón
        atajo_buscar = QShortcut(QKeySequence.Find, self)
        atajo_buscar.activated.connect(lambda: self.buscar())
        for clave, panel in (("navegador", self.navegador),
                             ("comentarios", self.panel_comentarios)):
            self.acciones[clave].setCheckable(True)
            panel.visibilityChanged.connect(
                lambda _v, a=self.acciones[clave], p=panel: self._sincronizar_boton(a, p))
        reemplazar = QAction("Reemplazar", self)
        reemplazar.setShortcut(QKeySequence("Ctrl+H"))
        reemplazar.triggered.connect(lambda: self.buscar(reemplazar=True))
        self.addAction(reemplazar)

        def seccion(nombre, claves):
            b = seccion_barra(nombre)
            for clave in claves:
                if clave is None:
                    b.addSeparator()
                else:
                    b.addAction(self.acciones[clave])
            return b

        # Fila 1 — izquierda: archivo · centro: formato e insertar ·
        # derecha: página y portapapeles
        barra = BarraFuente(self.editor, self)
        barra.setObjectName("seccion")
        barra.setFloatable(False)
        barra.addSeparator()
        for clave in ("imagen", "fecha", "comentario", "buscar"):
            barra.addAction(self.acciones[clave])
        self.barra_fuente = barra
        self.fila_superior = FilaHerramientas(
            "Herramientas",
            seccion("Archivo", ("abrir", "guardar", "guardar_como", None, "nuevo")),
            barra,
            seccion("Página y portapapeles",
                    ("pagina", None, "copiar", "cortar", "pegar")),
            self)
        self.fila_superior.setObjectName("filaSuperior")
        self.addToolBar(Qt.TopToolBarArea, self.fila_superior)
        self.addToolBarBreak(Qt.TopToolBarArea)

        # Fila 2 — izquierda: navegador · centro: estilos · derecha: comentarios
        estilos = seccion_barra("Estilos")
        estilos.addWidget(barra.cinta)
        estilos.addAction(barra.act_ajustes_estilo)
        self.fila_estilos = FilaHerramientas(
            "Estilos", seccion("Navegador", ("navegador",)), estilos,
            seccion("Comentarios", ("comentarios",)), self)
        self.fila_estilos.setObjectName("filaEstilos")
        self.addToolBar(Qt.TopToolBarArea, self.fila_estilos)
        self.setStatusBar(BarraEstado(self.vista, self))

        # Cortar y Copiar solo con selección; Pegar solo con algo pegable
        for senal in (self.editor.selectionChanged, self.editor.imagenCambiada,
                      QApplication.clipboard().dataChanged):
            senal.connect(self._actualizar_portapapeles)
        self._actualizar_portapapeles()

        self.editor.document().modificationChanged.connect(self.setWindowModified)

        # Copia de seguridad automática de los cambios sin guardar
        self._origen = None            # archivo abierto (también un .txt)
        self._respaldo = None          # copia de seguridad escrita ahora
        self._aviso_respaldo = False
        self._t_pausa = QTimer(self)
        self._t_pausa.setSingleShot(True)
        self._t_pausa.setInterval(RESPALDO_PAUSA_MS)
        self._t_pausa.timeout.connect(self._hacer_respaldo)
        self._t_max = QTimer(self)
        self._t_max.setSingleShot(True)
        self._t_max.setInterval(RESPALDO_MAX_MS)
        self._t_max.timeout.connect(self._hacer_respaldo)
        doc = self.editor.document()
        doc.contentsChanged.connect(self._programar_respaldo)
        doc.modificationChanged.connect(lambda m: m and self._programar_respaldo())
        # Los comentarios marcan el documento justo después de avisar
        self.editor.comentarios.cambiado.connect(
            lambda: QTimer.singleShot(0, self._programar_respaldo))
        _cursor_al_inicio(self.editor)
        self._actualizar_titulo()
        self.editor.setFocus()

    # -- copia de seguridad -------------------------------------------------
    def _programar_respaldo(self, *_):
        if not self.editor.document().isModified():
            return
        self._t_pausa.start()              # se reinicia con cada cambio
        if not self._t_max.isActive():
            self._t_max.start()            # pero no se espera más de esto

    def _hacer_respaldo(self):
        self._t_pausa.stop()
        self._t_max.stop()
        if not self.editor.document().isModified():
            return
        destino = ruta_respaldo(self._origen)
        temporal = os.path.join(os.path.dirname(destino),
                                "~" + os.path.basename(destino) + ".tmp")
        try:
            EscritorDocx(self.editor, self.vista.config).guardar(
                temporal, os.path.splitext(os.path.basename(destino))[0])
            # Se escribe aparte y se renombra: un corte de luz a mitad de la
            # escritura no deja la copia anterior dañada
            os.replace(temporal, destino)
        except Exception as error:
            try:
                os.remove(temporal)
            except OSError:
                pass
            if not self._aviso_respaldo:   # se avisa una sola vez
                self._aviso_respaldo = True
                self.statusBar().showMessage(
                    f"No se pudo crear la copia de seguridad: {error}", 8000)
            return
        if self._respaldo and self._respaldo != destino:
            self._borrar_archivo(self._respaldo)  # p. ej. tras «Guardar como»
        self._respaldo = destino
        self._aviso_respaldo = False

    @staticmethod
    def _borrar_archivo(ruta):
        try:
            os.remove(ruta)
        except OSError:
            pass

    def _borrar_respaldo(self):
        """Ya no hace falta: se guardó, se descartó o se cerró con normalidad."""
        self._t_pausa.stop()
        self._t_max.stop()
        if self._respaldo:
            self._borrar_archivo(self._respaldo)
        self._respaldo = None

    def _ofrecer_recuperar(self, origen):
        """Si quedó una copia de seguridad de `origen` (cierre inesperado),
        pregunta si se recupera. Devuelve True si se cargó la copia."""
        respaldo = ruta_respaldo(origen)
        if not os.path.isfile(respaldo):
            return False
        nombre = os.path.basename(origen) if origen else "Sin título"
        fecha = datetime.datetime.fromtimestamp(os.path.getmtime(respaldo))
        texto = (f"Hay una copia de seguridad de «{nombre}» del "
                 f"{fecha.strftime('%d/%m/%Y a las %H:%M')}, con cambios que no "
                 "se guardaron (el programa se cerró de forma inesperada).")
        if origen and os.path.isfile(origen) and \
                os.path.getmtime(origen) > os.path.getmtime(respaldo):
            texto += "\n\nEl documento se guardó después de esa copia."
        caja = QMessageBox(QMessageBox.Question, "Copia de seguridad", texto, parent=self)
        recuperar = caja.addButton("Recuperar la copia", QMessageBox.AcceptRole)
        original = caja.addButton("Abrir el original" if origen else "Descartar la copia",
                                  QMessageBox.DestructiveRole)
        caja.setDefaultButton(recuperar)
        caja.exec()
        if caja.clickedButton() is not recuperar:
            self._borrar_archivo(respaldo)
            return False
        try:
            abrir_docx(self.editor, self.vista, respaldo)
        except Exception as error:
            QMessageBox.critical(self, "Copia de seguridad",
                                 f"No se pudo leer la copia de seguridad.\n\n{error}")
            return False
        self._origen = origen
        self.ruta = origen if origen and origen.lower().endswith(".docx") else None
        self._respaldo = respaldo          # se mantiene hasta guardar o cerrar
        self.editor.document().setModified(True)
        self._actualizar_titulo()
        if self.editor.comentarios.orden:
            self.panel_comentarios.show()
        self.statusBar().showMessage("Copia de seguridad recuperada. Guarda el "
                                     "documento para conservar los cambios.", 8000)
        return True

    def comprobar_respaldo_sin_titulo(self):
        """Al iniciar: copia de un documento sin título que no se guardó."""
        self._ofrecer_recuperar(None)

    # -- archivo ------------------------------------------------------------
    def _actualizar_titulo(self):
        nombre = os.path.basename(self.ruta) if self.ruta else "Sin título"
        self.setWindowTitle(f"{nombre}[*] — {NOMBRE_APP}")
        self.setWindowModified(self.editor.document().isModified())

    def _carpeta(self):
        return os.path.dirname(self.ruta) if self.ruta else os.path.expanduser("~")

    def _confirmar_descartar(self):
        if not self.editor.document().isModified():
            return True
        nombre = os.path.basename(self.ruta) if self.ruta else "Sin título"
        caja = QMessageBox(QMessageBox.Warning, "Cambios sin guardar",
                           f"¿Guardar los cambios de «{nombre}» antes de continuar?",
                           parent=self)
        guardar = caja.addButton("Guardar", QMessageBox.AcceptRole)
        no_guardar = caja.addButton("No guardar", QMessageBox.DestructiveRole)
        cancelar = caja.addButton("Cancelar", QMessageBox.RejectRole)
        caja.setDefaultButton(guardar)
        caja.setEscapeButton(cancelar)
        caja.exec()
        if caja.clickedButton() is guardar:
            return self.guardar()
        return caja.clickedButton() is no_guardar

    def nuevo(self):
        if not self._confirmar_descartar():
            return
        doc = self.editor.document()
        self.editor.deseleccionar()
        self.editor.imagenes.clear()
        doc.setUndoRedoEnabled(False)
        doc.clear()
        self.editor.aplicar_body()
        doc.setUndoRedoEnabled(True)
        self.editor.comentarios.limpiar()
        self.vista.aplicar_config(self.vista.config)
        _cursor_al_inicio(self.editor)
        self._borrar_respaldo()            # los cambios se descartaron
        self.ruta = None
        self._origen = None
        self._actualizar_titulo()
        self.editor.setFocus()

    def abrir(self):
        if not self._confirmar_descartar():
            return
        ruta, _ = QFileDialog.getOpenFileName(self, "Abrir", self._carpeta(), FILTRO_ABRIR)
        if not ruta:
            return
        self._borrar_respaldo()            # el documento anterior se descartó
        if self._ofrecer_recuperar(ruta):
            self.editor.setFocus()
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            if ruta.lower().endswith(".txt"):
                abrir_texto(self.editor, self.vista, ruta)
            else:
                abrir_docx(self.editor, self.vista, ruta)
        except Exception as error:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "No se pudo abrir",
                                 f"No se pudo abrir «{os.path.basename(ruta)}».\n\n{error}")
            self._hacer_respaldo()         # el documento actual sigue abierto
            return
        QApplication.restoreOverrideCursor()
        if self.editor.comentarios.orden:
            self.panel_comentarios.show()
        # Un .txt se guarda después como .docx (Guardar pedirá el nombre)
        self.ruta = ruta if ruta.lower().endswith(".docx") else None
        self._origen = ruta
        self._actualizar_titulo()
        self.editor.setFocus()

    def guardar(self):
        if not self.ruta:
            return self.guardar_como()
        return self._escribir(self.ruta)

    def guardar_como(self):
        sugerido = self.ruta or os.path.join(self._carpeta(), "Sin título.docx")
        ruta, _ = QFileDialog.getSaveFileName(self, "Guardar como", sugerido, FILTRO_DOCX)
        if not ruta:
            return False
        if not ruta.lower().endswith(".docx"):
            ruta += ".docx"
        if self._escribir(ruta):
            self.ruta = ruta
            self._origen = ruta
            self._actualizar_titulo()
            return True
        return False

    def _escribir(self, ruta):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            EscritorDocx(self.editor, self.vista.config).guardar(
                ruta, os.path.splitext(os.path.basename(ruta))[0])
        except Exception as error:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "No se pudo guardar",
                                 f"No se pudo guardar «{os.path.basename(ruta)}».\n\n{error}")
            return False
        QApplication.restoreOverrideCursor()
        self.editor.document().setModified(False)
        self._borrar_respaldo()            # ya está guardado
        self.statusBar().showMessage(f"Guardado: {ruta}", 4000)
        return True

    def closeEvent(self, e):
        if self._confirmar_descartar():
            self._borrar_respaldo()        # cierre normal: ya no hace falta
            guardar_ventana(self)
            e.accept()
        else:
            e.ignore()

    # -- edición ------------------------------------------------------------
    def _actualizar_portapapeles(self, *_):
        try:
            hay = self.editor.textCursor().hasSelection() or \
                self.editor.imagen_seleccionada() is not None
            self.acciones["cortar"].setEnabled(hay)
            self.acciones["copiar"].setEnabled(hay)
            md = QApplication.clipboard().mimeData()
            self.acciones["pegar"].setEnabled(
                md is not None and self.editor.canInsertFromMimeData(md))
        except RuntimeError:          # al cerrar la ventana
            pass

    def cortar(self):
        self.editor.cut()
        self.editor.setFocus()

    def copiar(self):
        self.editor.copy()
        self._actualizar_portapapeles()
        self.editor.setFocus()

    def pegar(self):
        self.editor.paste()
        self.editor.setFocus()

    def insertar_fecha_hora(self):
        self.enlace_fecha.abrir()

    def insertar_imagen(self):
        formatos = sorted({bytes(f).decode("ascii", "ignore").lower()
                           for f in QImageReader.supportedImageFormats()})
        filtro = "Imágenes (" + " ".join(f"*.{f}" for f in formatos) + ");;Todos (*)"
        ruta, _ = QFileDialog.getOpenFileName(self, "Insertar imagen",
                                              self._carpeta_imagen or self._carpeta(),
                                              filtro)
        if not ruta:
            self.editor.setFocus()
            return
        self._carpeta_imagen = os.path.dirname(ruta)
        try:
            ok = self.editor.insertar_imagen_archivo(ruta)
        except OSError as error:
            ok, detalle = False, str(error)
        else:
            detalle = "El archivo no es una imagen o su formato no es compatible."
        if not ok:
            QMessageBox.warning(self, "Insertar imagen",
                                f"No se pudo insertar «{os.path.basename(ruta)}».\n\n{detalle}")
        self.editor.setFocus()

    @property
    def _ventana_buscar(self):
        return self.enlace_buscar.ventana

    def buscar(self, reemplazar=False):
        self.enlace_buscar.abrir(reemplazar)

    # -- paneles ------------------------------------------------------------
    @staticmethod
    def _sincronizar_boton(accion, panel):
        try:
            accion.setChecked(panel.isVisible())
        except RuntimeError:
            pass

    def alternar_navegador(self, *_):
        self.navegador.setVisible(not self.navegador.isVisible())
        if self.navegador.isVisible():
            self.navegador.refrescar()

    def alternar_comentarios(self, *_):
        self.panel_comentarios.setVisible(not self.panel_comentarios.isVisible())

    def nuevo_comentario(self, *_):
        i = self.editor.comentarios.agregar()
        if i is None:
            self.statusBar().showMessage(
                "Selecciona el texto que quieres comentar "
                "(o pon el cursor sobre una palabra)", 4000)
            QApplication.beep()
            return
        self.panel_comentarios.enfocar(i)

    def _pagina_aceptada(self, dialogo):
        antes = dict(self.vista.config)
        nueva = dialogo.config()
        self.vista.aplicar_config(nueva)
        # El tamaño y los márgenes solo se guardan si se cambiaron aquí (no
        # los que traía el .docx abierto)
        guardar_vista(self.vista.config,
                      con_pagina=any(antes.get(k) != nueva[k] for k in CLAVES_PAGINA))

    def configurar_pagina(self):
        self.enlace_pagina.abrir()


class SinTooltips(QObject):
    """Bloquea todos los tooltips de la aplicación (incluidos los automáticos
    que la barra de herramientas genera con el nombre de cada acción)."""
    def eventFilter(self, obj, event):
        return event.type() == QEvent.ToolTip


def portapapeles_seguro_al_salir(app):
    """Si al cerrar el programa el portapapeles tiene datos copiados desde
    este editor (un QMimeData creado en Python), Qt los borra después de
    que Python liberó su parte y el programa termina con un fallo. Antes de
    salir se cambian por una copia que crea Qt, conservando el contenido
    para pegarlo en otras aplicaciones."""
    def al_salir():
        cb = QApplication.clipboard()
        md = cb.mimeData()
        # Solo los datos copiados aquí (llevan nuestros formatos propios);
        # ownsClipboard() no es fiable en todas las plataformas
        if md is None or not (md.hasFormat(MIME_LIBRO) or md.hasFormat(MIME_IMAGEN)):
            return
        if md.hasImage() and not md.hasText():
            img = QImage(md.imageData())
            if not img.isNull():
                cb.setImage(img)
                return
        if md.hasHtml():
            # Un QTextEdit normal crea sus datos en C++ (HTML + texto)
            tmp = QTextEdit()
            tmp.setHtml(md.html())
            tmp.selectAll()
            tmp.copy()
            tmp.deleteLater()
        elif md.hasText():
            cb.setText(md.text())
        else:
            cb.clear()
    app.aboutToQuit.connect(al_salir)


if __name__ == "__main__":
    if sys.platform == "win32":
        # Windows agrupa la ventana en la barra de tareas con su propio
        # icono (si no, mostraría el de Python)
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "SWordBook.Editor")
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setApplicationName(NOMBRE_APP)
    app.setWindowIcon(icono_aplicacion())
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setPalette(paleta_clara())
    app.setStyleSheet(ESTILO)
    # Textos estándar de Qt en español (selector de archivos, menús, etc.)
    traductor = QTranslator(app)
    if traductor.load(QLocale(QLocale.Spanish, QLocale.Mexico), "qtbase", "_",
                      QLibraryInfo.path(QLibraryInfo.TranslationsPath)):
        app.installTranslator(traductor)
    sin_tooltips = SinTooltips(app)
    app.installEventFilter(sin_tooltips)
    portapapeles_seguro_al_salir(app)
    tamano = cargar_config()           # estilos, tema y página de config.ini
    ventana = Ventana()
    if tamano:
        pantalla = app.primaryScreen().availableGeometry()
        ventana.resize(min(tamano["ancho"], pantalla.width()),
                       min(tamano["alto"], pantalla.height()))
    if tamano.get("maximizada"):
        ventana.showMaximized()
    else:
        ventana.show()
    # Copia de un documento sin título que quedó de un cierre inesperado
    QTimer.singleShot(0, ventana.comprobar_respaldo_sin_titulo)
    sys.exit(app.exec())