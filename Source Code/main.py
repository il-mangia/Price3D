# -*- coding: utf-8 -*-
"""
Price3D — Calcolo Prezzo Stampa 3D
----------------------------------
- Profili stampante in JSON nella cartella dati dell'app:
    Windows : %APPDATA%/Price3D/profiles
    macOS   : ~/Documents/Price3D/profiles
    Linux   : ~/Documents/Price3D/profiles
- config.json salvato nella stessa cartella dati.
- Modalita' singolo / multi-oggetto:
    singolo -> 1 stampa, 1 oggetto
    multi   -> N stampe, M oggetti totali; tempo e materiale indicati
               sono PER STAMPA, il prezzo al pezzo e' totale / oggetti.
- Checkbox con tick (non riempimento verde).
- Logo app: Logo.png nella cartella dello script.
"""

import json
import math
import os
import re
import shutil
import sys
import time
import unicodedata
import urllib.request
from datetime import date
from pathlib import Path

try:
    from PyQt6.QtWidgets import (
        QApplication, QWidget, QLabel, QLineEdit, QTextEdit, QPushButton,
        QGridLayout, QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox,
        QSpinBox, QComboBox, QDialog, QDoubleSpinBox, QListWidget,
        QDialogButtonBox, QGroupBox, QCheckBox, QScrollArea, QFrame,
        QSizePolicy, QStackedWidget, QSplashScreen, QMenuBar, QMenu,
    )
    from PyQt6.QtGui import QFont, QPixmap, QIcon, QDesktopServices, QAction, QPainter, QColor
    from PyQt6.QtCore import Qt, QUrl, QThread
except Exception as e:
    raise RuntimeError(
        "PyQt6 is required to run this application. Install with: pip install PyQt6"
    ) from e

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.pdfbase.pdfmetrics import stringWidth
    _REPORTLAB_OK = True
except Exception:
    _REPORTLAB_OK = False

try:
    from docx import Document
    from docx.shared import Mm, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _PYDOCX_OK = True
except Exception:
    _PYDOCX_OK = False

try:
    from PIL import Image, ImageDraw
    _PIL_OK = True
except Exception:
    _PIL_OK = False

# ----------------------------------------------------------------------
# Percorsi
# ----------------------------------------------------------------------
APP_NAME = "Price3D"
__version__ = "1.0"
GITHUB_URL = "https://github.com/il-mangia/Price3D"
GITHUB_API_LATEST = "https://api.github.com/repos/il-mangia/Price3D/releases/latest"
BASE_DIR = Path(__file__).resolve().parent          # file bundlati (Logo.png, languages.json)
LANG_FILE = BASE_DIR / "languages.json"
LOGO_FILE = BASE_DIR / "Logo.png"


def app_dir_path() -> Path:
    """Cartella dati fissa per OS (senza crearla).

    Windows : %APPDATA%/Price3D
    macOS   : ~/Documents/Price3D
    Linux   : ~/Documents/Price3D
    """
    if os.name == "nt":
        base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        return Path(base) / APP_NAME
    else:
        # macOS e Linux: cartella Documenti dell'utente
        return Path.home() / "Documents" / APP_NAME


def get_app_dir() -> Path:
    """Cartella dati fissa per OS. Creata se manca."""
    d = app_dir_path()
    d.mkdir(parents=True, exist_ok=True)
    (d / "profiles").mkdir(parents=True, exist_ok=True)
    return d


SETTINGS_FILE = app_dir_path() / "settings.json"

# Primo avvio: la cartella dati o settings.json non esistono ancora.
# Rilevato QUI, prima che get_app_dir() crei la cartella qui sotto.
FIRST_RUN = (not app_dir_path().is_dir()) or (not SETTINGS_FILE.is_file())

APP_DIR = get_app_dir()
PROFILES_DIR = APP_DIR / "profiles"
CONFIG_FILE = APP_DIR / "config.json"


def _migrate_legacy():
    """Copia config/profili vecchi (accanto allo script) nella cartella dati."""
    try:
        legacy_cfg = BASE_DIR / "config.json"
        if legacy_cfg.is_file() and not CONFIG_FILE.is_file():
            shutil.copy2(legacy_cfg, CONFIG_FILE)
        legacy_prof = BASE_DIR / "profiles"
        if legacy_prof.is_dir():
            for f in legacy_prof.glob("*.json"):
                dest = PROFILES_DIR / f.name
                if not dest.exists():
                    shutil.copy2(f, dest)
    except Exception:
        pass


_migrate_legacy()

DEFAULT_PROFILE = {
    "nome": "Default",
    "costo_corrente": 0.25,      # €/kWh
    "consumo_w": 1300.0,         # W stampante
    "costo_materiale": 13.0,     # €/kg
    "costo_progettazione": 5.0,  # €/h
    "consumo_pc_w": 150.0,       # W del PC durante la progettazione
}

FALLBACK_LANG = {"en": {"_name": "English", "header": APP_NAME}}


# ----------------------------------------------------------------------
# Icona tick per i checkbox (niente riempimento verde)
# ----------------------------------------------------------------------
def get_tick_icon_path() -> str:
    """Ritorna il percorso di un tick bianco su trasparente per il QSS.

    Cerca tick.png / check.png accanto allo script, altrimenti lo genera
    con Pillow in APP_DIR/tick.png. Ritorna "" se non disponibile.
    """
    for cand in (BASE_DIR / "tick.png", BASE_DIR / "check.png"):
        if cand.is_file():
            return cand.as_posix()
    dest = APP_DIR / "tick.png"
    if dest.is_file():
        return dest.as_posix()
    if _PIL_OK:
        try:
            size = 48
            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            dr = ImageDraw.Draw(img)
            dr.line([(13, 25), (22, 34), (36, 15)], fill=(255, 255, 255, 255),
                    width=7, joint="curve")
            img.save(dest)
            return dest.as_posix()
        except Exception:
            pass
    return ""


TICK_PATH = get_tick_icon_path()
_TICK_QSS = f'\nQCheckBox::indicator:checked {{ image: url("{TICK_PATH}"); }}' if TICK_PATH else ""


# ----------------------------------------------------------------------
# Stylesheet (compatto per schermi 720p, look moderno)
# ----------------------------------------------------------------------
STYLE = """
QWidget {
    background-color: #0f1420;
    color: #e9edf4;
    font-family: 'Segoe UI', 'Helvetica Neue', sans-serif;
    font-size: 10pt;
}
QLabel#brandTitle {
    color: #ffffff;
    font-size: 17pt;
    font-weight: 800;
    letter-spacing: 0.4px;
}
QLabel#subheader {
    color: #7fd0ff;
    font-size: 9pt;
    font-weight: 500;
}
QLabel#footer { color: #5f6a80; font-size: 8pt; }
QLabel#hint { color: #8b96ad; font-size: 8.5pt; font-style: italic; }
QLabel#totalBadge {
    color: #ffffff;
    font-size: 11pt;
    font-weight: 700;
}
QGroupBox {
    background-color: #161c2b;
    border: 1px solid #2a3348;
    border-radius: 12px;
    margin-top: 16px;
    padding: 13px 12px 11px 12px;
    font-weight: 700;
    font-size: 10pt;
    color: #8fc3f7;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 2px 9px;
    background-color: #22304a;
    border: 1px solid #354664;
    border-radius: 6px;
    color: #a8d4ff;
}
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {
    background-color: #1b2233;
    border: 1px solid #323c52;
    border-radius: 7px;
    padding: 6px 9px;
    selection-background-color: #2f6fd6;
    color: #f1f4f9;
    min-height: 20px;
}
QLineEdit:hover, QSpinBox:hover, QDoubleSpinBox:hover, QComboBox:hover {
    border: 1px solid #4a5872;
    background-color: #202844;
}
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {
    border: 1.5px solid #4fa8f7;
    background-color: #212a3d;
}
QLineEdit:disabled { color: #6b7590; background-color: #151b2a; }
QComboBox::drop-down { border: none; width: 26px; }
QComboBox QAbstractItemView {
    background-color: #1b2233;
    border: 1px solid #3a4560;
    selection-background-color: #2f6fd6;
    outline: none;
    padding: 4px;
}
QCheckBox { spacing: 9px; font-size: 10pt; }
QCheckBox::indicator {
    width: 18px; height: 18px;
    border: 1.5px solid #46536e;
    border-radius: 5px;
    background-color: #1b2233;
}
QCheckBox::indicator:hover { border: 1.5px solid #4fa8f7; }
QCheckBox::indicator:checked {
    background-color: #2f6fd6;
    border-color: #2f6fd6;
}
QCheckBox::indicator:checked:hover { background-color: #3f81ea; border-color: #3f81ea; }
""" + _TICK_QSS + """
QPushButton {
    background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #26314a, stop:1 #1f2740);
    border: 1px solid #3a4666;
    border-radius: 9px;
    padding: 8px 14px;
    font-weight: 600;
    font-size: 10pt;
    color: #eef2f8;
    min-height: 22px;
}
QPushButton:hover { background-color: #2e3a58; border-color: #4c5d85; }
QPushButton:pressed { background-color: #1a2236; }
QPushButton#primary {
    background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #27b56d, stop:1 #1d9457);
    border: 1px solid #35cf82;
    color: #ffffff;
}
QPushButton#primary:hover { background-color: #27bd74; }
QPushButton#primary:pressed { background-color: #167a44; }
QPushButton#accent {
    background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #3a7be0, stop:1 #2c63bd);
    border: 1px solid #4a8ef2;
    color: #ffffff;
}
QPushButton#accent:hover { background-color: #4a8aea; }
QPushButton#accent:pressed { background-color: #2659ac; }
QPushButton#danger { background-color: #8f3a38; border: 1px solid #b5504e; color: #ffffff; }
QPushButton#danger:hover { background-color: #ac4643; }
QPushButton#ghost {
    background-color: transparent;
    border: 1px dashed #3d4a66;
    color: #9fb0cc;
    font-weight: 600;
}
QPushButton#ghost:hover {
    background-color: #1b2233;
    border: 1px dashed #4fa8f7;
    color: #cfe6ff;
}
QTextEdit {
    background-color: #0a0e17;
    border: 1px solid #2a3348;
    border-radius: 12px;
    padding: 12px;
    color: #cdf3dd;
    selection-background-color: #1d9457;
}
QListWidget {
    background-color: #1b2233;
    border: 1px solid #323c52;
    border-radius: 8px;
    padding: 4px;
}
QListWidget::item {
    padding: 6px 8px;
    border-radius: 6px;
}
QListWidget::item:selected {
    background-color: #2f6fd6;
    color: #ffffff;
}
QListWidget::item:hover:!selected {
    background-color: #22304a;
}
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical {
    background: transparent;
    width: 10px;
    margin: 2px;
}
QScrollBar::handle:vertical {
    background: #2e3a55;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover { background: #3d4f78; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal {
    background: transparent;
    height: 10px;
    margin: 2px;
}
QScrollBar::handle:horizontal {
    background: #2e3a55;
    border-radius: 5px;
    min-width: 30px;
}
QScrollBar::handle:horizontal:hover { background: #3d4f78; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QDialog { background-color: #0f1420; }
QMessageBox { background-color: #161c2b; }
QMessageBox QLabel { color: #e9edf4; font-size: 10pt; }
QMessageBox QPushButton { min-width: 80px; }
"""


# ----------------------------------------------------------------------
# Utility
# ----------------------------------------------------------------------

def euro(v: float) -> str:
    return f"{v:.2f} €"


def slugify(nome: str) -> str:
    """'Artillery Genius' -> 'artillery_genius'"""
    nome = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    nome = nome.strip().lower()
    nome = re.sub(r"[^a-z0-9]+", "_", nome)
    return nome.strip("_") or "profile"


def load_languages() -> dict:
    try:
        with open(LANG_FILE, "r", encoding="utf-8") as fp:
            data = json.load(fp)
        if isinstance(data, dict) and data:
            return data
    except Exception:
        pass
    return dict(FALLBACK_LANG)


def load_config() -> dict:
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as fp:
            return json.load(fp)
    except Exception:
        return {}


def save_config(cfg: dict):
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as fp:
            json.dump(cfg, fp, indent=2, ensure_ascii=False)
    except Exception:
        pass


def startup_language() -> tuple:
    """(languages, lang_code) per le finestre mostrate prima di quella principale."""
    langs = load_languages()
    code = load_config().get("language", "en")
    if code not in langs:
        code = "en" if "en" in langs else next(iter(langs))
    return langs, code


def load_company() -> dict:
    """Dati azienda da settings.json (nome + logo)."""
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as fp:
            d = json.load(fp)
        if isinstance(d, dict):
            return d
    except Exception:
        pass
    return {}


def save_company(company_name: str, logo_src: str | None) -> dict:
    """Salva settings.json; copia il logo scelto nella cartella dati.

    logo_src None o vuoto = tieni il logo precedente (se esiste).
    Ritorna il dict salvato.
    """
    APP_DIR.mkdir(parents=True, exist_ok=True)
    logo_file = ""
    if logo_src:
        try:
            src = Path(logo_src)
            if src.is_file():
                dest = APP_DIR / f"company_logo{src.suffix.lower() or '.png'}"
                for old in APP_DIR.glob("company_logo.*"):
                    try:
                        if old != dest:
                            old.unlink()
                    except Exception:
                        pass
                shutil.copy2(src, dest)
                logo_file = dest.name
        except Exception:
            pass
    if not logo_file:
        prev = load_company().get("company_logo", "")
        if prev and (APP_DIR / prev).is_file():
            logo_file = prev
    data = {"company_name": (company_name or "").strip(), "company_logo": logo_file}
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as fp:
            json.dump(data, fp, indent=2, ensure_ascii=False)
    except Exception:
        pass
    return data


def company_name() -> str:
    return (load_company().get("company_name") or "").strip()


def company_logo_path() -> str | None:
    fname = (load_company().get("company_logo") or "").strip()
    if fname and (APP_DIR / fname).is_file():
        return str(APP_DIR / fname)
    return None


# ----------------------------------------------------------------------
# Versione e controllo aggiornamenti (release GitHub)
# ----------------------------------------------------------------------
def is_beta_version(v: str) -> bool:
    """True se la versione contiene beta/alpha/rc/dev/pre."""
    return bool(re.search(r"(beta|alpha|rc|dev|pre|snapshot)", v or "", re.IGNORECASE))


def parse_version_tuple(v: str) -> tuple:
    """'v1.2-beta' -> (1, 2, 0). Solo parti numeriche, max 3, pad a 3."""
    nums = [int(x) for x in re.findall(r"\d+", v or "")]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums[:3])


def fetch_latest_release(timeout: float = 6.0) -> tuple:
    """(tag, html_url, error) dall'ultima release GitHub. Solo stdlib."""
    try:
        req = urllib.request.Request(
            GITHUB_API_LATEST,
            headers={"User-Agent": APP_NAME,
                     "Accept": "application/vnd.github+json"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.load(resp)
        return (data.get("tag_name") or "", data.get("html_url") or "", "")
    except Exception as e:
        return ("", "", str(e))


class UpdateCheckThread(QThread):
    """Controllo update in background: non blocca lo splash all'avvio."""

    def __init__(self, timeout: float = 6.0):
        super().__init__()
        self.timeout = timeout
        self.tag = ""
        self.url = ""
        self.error = ""

    def run(self):
        self.tag, self.url, self.error = fetch_latest_release(self.timeout)


# Ultimo esito del controllo update (riempito all'avvio, letto dall'About)
UPDATE_INFO = {"tag": "", "url": "", "error": "", "available": False,
               "current": __version__}


def update_status_text(tr, info: dict) -> str:
    """Messaggio localizzato sullo stato degli aggiornamenti."""
    if info.get("available") and info.get("tag"):
        return tr("upd_available").format(v=info["tag"])
    if info.get("tag"):
        return tr("upd_ok")
    return tr("upd_offline")


def build_splash_pixmap(tr, w: int = 420, h: int = 470) -> QPixmap:
    """Immagine dello splash/about: logo, nome, versione, avviso Beta."""
    pm = QPixmap(w, h)
    pm.fill(QColor("#0f1420"))
    p = QPainter(pm)
    try:
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(QColor("#2a3348"))
        p.setBrush(QColor("#161c2b"))
        p.drawRoundedRect(4, 4, w - 8, h - 8, 18, 18)

        # logo
        logo = app_logo_pixmap(120)
        if logo is not None:
            p.drawPixmap((w - 120) // 2, 28, logo)
            y0 = 160
        else:
            p.setPen(QColor("#ffffff"))
            p.setFont(QFont("Segoe UI", 64))
            p.drawText(0, 30, w, 120, Qt.AlignmentFlag.AlignCenter, "🖨️")
            y0 = 160

        p.setPen(QColor("#ffffff"))
        p.setFont(QFont("Segoe UI", 26, QFont.Weight.ExtraBold))
        p.drawText(0, y0, w, 44, Qt.AlignmentFlag.AlignCenter, APP_NAME)

        p.setPen(QColor("#7fd0ff"))
        p.setFont(QFont("Segoe UI", 12))
        p.drawText(0, y0 + 44, w, 26, Qt.AlignmentFlag.AlignCenter,
                   f"{tr('about_version')} {__version__}")

        y_msg = y0 + 76
        if is_beta_version(__version__):
            p.setPen(QColor("#0f1420"))
            p.setBrush(QColor("#e8a13c"))
            p.drawRoundedRect(30, y_msg, w - 60, 52, 10, 10)
            p.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            p.drawText(38, y_msg, w - 76, 52,
                       Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap,
                       tr("upd_beta"))
            y_msg += 60
        p.setPen(QColor("#5f6a80"))
        p.setFont(QFont("Segoe UI", 9))
        p.drawText(0, h - 44, w, 24, Qt.AlignmentFlag.AlignCenter, GITHUB_URL)
    finally:
        p.end()
    return pm


def load_profiles() -> dict:
    """{nome_profilo: dati} da tutti i .json nella cartella dati."""
    PROFILES_DIR.mkdir(parents=True, exist_ok=True)
    profiles = {}
    for f in sorted(PROFILES_DIR.glob("*.json")):
        try:
            with open(f, "r", encoding="utf-8") as fp:
                d = json.load(fp)
            nome = d.get("nome") or f.stem
            profiles[nome] = {
                "nome": nome,
                "costo_corrente": float(d.get("costo_corrente", DEFAULT_PROFILE["costo_corrente"])),
                "consumo_w": float(d.get("consumo_w", DEFAULT_PROFILE["consumo_w"])),
                "costo_materiale": float(d.get("costo_materiale", DEFAULT_PROFILE["costo_materiale"])),
                "costo_progettazione": float(d.get("costo_progettazione", DEFAULT_PROFILE["costo_progettazione"])),
                "consumo_pc_w": float(d.get("consumo_pc_w", DEFAULT_PROFILE["consumo_pc_w"])),
                "_file": str(f),
            }
        except Exception:
            continue
    return profiles


def save_profile(d: dict) -> Path:
    PROFILES_DIR.mkdir(parents=True, exist_ok=True)
    path = PROFILES_DIR / f"{slugify(d['nome'])}.json"
    to_save = {k: d[k] for k in (
        "nome", "costo_corrente", "consumo_w", "costo_materiale",
        "costo_progettazione", "consumo_pc_w"
    )}
    with open(path, "w", encoding="utf-8") as fp:
        json.dump(to_save, fp, indent=2, ensure_ascii=False)
    return path


def app_logo_pixmap(size: int = 44):
    """QPixmap del logo (Logo.png) scalato, o None se manca."""
    try:
        if LOGO_FILE.is_file():
            pm = QPixmap(str(LOGO_FILE))
            if not pm.isNull():
                return pm.scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio,
                                 Qt.TransformationMode.SmoothTransformation)
    except Exception:
        pass
    return None


# ----------------------------------------------------------------------
# Wizard di primo avvio (solo se mancano cartella dati o settings.json)
# ----------------------------------------------------------------------
class FirstRunWizard(QDialog):
    """Setup guidato: 1) benvenuto 2) dati azienda 3) primo profilo."""

    def __init__(self, tr, parent=None):
        super().__init__(parent)
        self.tr_ = tr
        self.setWindowTitle(APP_NAME)
        self.company_logo_src = None  # scelto dall'utente, copiato solo alla fine
        self._build_ui()
        self._go(0)
        self.setMinimumSize(540, 470)
        try:
            avail = self.screen().availableGeometry() if self.screen() else None
            if avail is not None:
                self.resize(min(620, avail.width() - 60),
                            min(560, avail.height() - 60))
            else:
                self.resize(600, 540)
        except Exception:
            self.resize(600, 540)

    # ------------------------------------------------------------------
    def _spin(self, decimals, vmin, vmax, step, value):
        s = QDoubleSpinBox()
        s.setDecimals(decimals)
        s.setRange(vmin, vmax)
        s.setSingleStep(step)
        s.setValue(value)
        return s

    def _build_ui(self):
        tr = self.tr_
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 14, 16, 12)
        outer.setSpacing(8)

        self.step_label = QLabel()
        self.step_label.setObjectName("hint")
        self.step_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(self.step_label)

        self.stack = QStackedWidget()
        outer.addWidget(self.stack, 1)

        # ---- pagina 1: benvenuto
        p1 = QWidget()
        l1 = QVBoxLayout(p1)
        l1.setSpacing(8)
        l1.addStretch()
        logo_lbl = QLabel()
        logo_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pm = app_logo_pixmap(96)
        if pm is not None:
            logo_lbl.setPixmap(pm)
        else:
            logo_lbl.setText("🖨️")
            logo_lbl.setStyleSheet("font-size: 48pt;")
        l1.addWidget(logo_lbl)
        t1 = QLabel(APP_NAME)
        t1.setObjectName("brandTitle")
        t1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        l1.addWidget(t1)
        w1 = QLabel(tr("wiz_welcome"))
        w1.setObjectName("subheader")
        w1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        l1.addWidget(w1)
        i1 = QLabel(tr("wiz_intro"))
        i1.setWordWrap(True)
        i1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        l1.addWidget(i1)
        self.btn_github = QPushButton("🌐  " + tr("wiz_github"))
        self.btn_github.setObjectName("ghost")
        self.btn_github.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl(GITHUB_URL))
        )
        l1.addWidget(self.btn_github, alignment=Qt.AlignmentFlag.AlignCenter)
        l1.addStretch()
        self.stack.addWidget(p1)

        # ---- pagina 2: dati azienda
        p2 = QWidget()
        l2 = QVBoxLayout(p2)
        l2.setSpacing(8)
        box = QGroupBox(tr("wiz_company_title"))
        form = QGridLayout(box)
        form.setVerticalSpacing(10)
        form.addWidget(QLabel(tr("wiz_company_name")), 0, 0)
        self.company_edit = QLineEdit()
        self.company_edit.setPlaceholderText(tr("wiz_company_name_ph"))
        form.addWidget(self.company_edit, 0, 1)
        form.addWidget(QLabel(tr("wiz_company_logo")), 1, 0)
        logo_row = QHBoxLayout()
        logo_row.setSpacing(8)
        self.logo_preview = QLabel()
        self.logo_preview.setFixedSize(64, 64)
        self.logo_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.logo_preview.setStyleSheet(
            "border: 1px dashed #3d4a66; border-radius: 8px; color: #5f6a80; font-size: 8pt;")
        self.logo_preview.setText(tr("wiz_no_logo"))
        self.logo_preview.setWordWrap(True)
        logo_row.addWidget(self.logo_preview)
        logo_col = QVBoxLayout()
        self.btn_logo = QPushButton(tr("wiz_choose_logo"))
        self.btn_logo.clicked.connect(self._choose_logo)
        logo_col.addWidget(self.btn_logo)
        self.logo_file_label = QLabel("")
        self.logo_file_label.setObjectName("hint")
        self.logo_file_label.setWordWrap(True)
        logo_col.addWidget(self.logo_file_label)
        logo_row.addLayout(logo_col, 1)
        form.addLayout(logo_row, 1, 1)
        l2.addWidget(box)
        l2.addStretch()
        self.stack.addWidget(p2)

        # ---- pagina 3: primo profilo stampante
        p3 = QWidget()
        l3 = QVBoxLayout(p3)
        l3.setSpacing(8)
        desc = QLabel(tr("wiz_profile_desc"))
        desc.setWordWrap(True)
        desc.setObjectName("hint")
        l3.addWidget(desc)
        pbox = QGroupBox(tr("wiz_profile_title"))
        pform = QGridLayout(pbox)
        pform.setVerticalSpacing(8)
        pform.addWidget(QLabel(tr("profile_name")), 0, 0)
        self.p_name = QLineEdit()
        self.p_name.setPlaceholderText(tr("name_placeholder"))
        pform.addWidget(self.p_name, 0, 1)
        pform.addWidget(QLabel(tr("energy_cost")), 1, 0)
        self.p_energy = self._spin(3, 0.0, 10.0, 0.01, DEFAULT_PROFILE["costo_corrente"])
        pform.addWidget(self.p_energy, 1, 1)
        pform.addWidget(QLabel(tr("printer_power")), 2, 0)
        self.p_power = self._spin(0, 1.0, 10000.0, 50.0, DEFAULT_PROFILE["consumo_w"])
        pform.addWidget(self.p_power, 2, 1)
        pform.addWidget(QLabel(tr("material_cost")), 3, 0)
        self.p_material = self._spin(2, 0.0, 1000.0, 0.5, DEFAULT_PROFILE["costo_materiale"])
        pform.addWidget(self.p_material, 3, 1)
        pform.addWidget(QLabel(tr("design_rate")), 4, 0)
        self.p_design = self._spin(2, 0.0, 500.0, 0.5, DEFAULT_PROFILE["costo_progettazione"])
        pform.addWidget(self.p_design, 4, 1)
        pform.addWidget(QLabel(tr("pc_power")), 5, 0)
        self.p_pc = self._spin(0, 0.0, 5000.0, 10.0, DEFAULT_PROFILE["consumo_pc_w"])
        pform.addWidget(self.p_pc, 5, 1)
        l3.addWidget(pbox)
        l3.addStretch()
        self.stack.addWidget(p3)

        # ---- navigazione
        nav = QHBoxLayout()
        nav.setSpacing(8)
        self.btn_back = QPushButton(tr("wiz_back"))
        self.btn_back.clicked.connect(lambda: self._go(self.stack.currentIndex() - 1))
        nav.addWidget(self.btn_back)
        nav.addStretch()
        self.btn_next = QPushButton(tr("wiz_next"))
        self.btn_next.setObjectName("primary")
        self.btn_next.clicked.connect(self._next)
        nav.addWidget(self.btn_next)
        outer.addLayout(nav)

    # ------------------------------------------------------------------
    def _go(self, idx: int):
        idx = max(0, min(2, idx))
        self.stack.setCurrentIndex(idx)
        self.step_label.setText(self.tr_("wiz_step").format(n=idx + 1))
        self.btn_back.setVisible(idx > 0)
        self.btn_next.setText(self.tr_("wiz_finish") if idx == 2 else self.tr_("wiz_next"))
        if idx == 2:
            self.p_name.setFocus()

    def _next(self):
        if self.stack.currentIndex() < 2:
            self._go(self.stack.currentIndex() + 1)
        else:
            self._finish()

    def _choose_logo(self):
        tr = self.tr_
        path, _ = QFileDialog.getOpenFileName(
            self, tr("wiz_company_logo"), "",
            "Images (*.png *.jpg *.jpeg);;All files (*)"
        )
        if not path:
            return
        try:
            pm = QPixmap(path)
            if pm.isNull():
                raise ValueError("bad image")
            self.company_logo_src = path
            self.logo_preview.setPixmap(
                pm.scaled(62, 62, Qt.AspectRatioMode.KeepAspectRatio,
                          Qt.TransformationMode.SmoothTransformation))
            self.logo_file_label.setText(Path(path).name)
        except Exception:
            QMessageBox.warning(self, tr("msg_invalid_title"), tr("msg_invalid_text"))

    def _finish(self):
        tr = self.tr_
        nome = self.p_name.text().strip()
        if not nome:
            QMessageBox.warning(self, tr("msg_noname_title"), tr("msg_noname_text"))
            self.p_name.setFocus()
            return
        save_company(self.company_edit.text(), self.company_logo_src)
        save_profile({
            "nome": nome,
            "costo_corrente": self.p_energy.value(),
            "consumo_w": self.p_power.value(),
            "costo_materiale": self.p_material.value(),
            "costo_progettazione": self.p_design.value(),
            "consumo_pc_w": self.pc_spin.value(),
        })
        self.accept()


# ----------------------------------------------------------------------
# About / Informazioni (stessa grafica dello splash, aperto con F1)
# ----------------------------------------------------------------------
class AboutDialog(QDialog):
    def __init__(self, tr, parent=None):
        super().__init__(parent)
        self.tr_ = tr
        self.setWindowTitle(tr("about_title"))
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 14, 16, 12)
        outer.setSpacing(8)

        img = QLabel()
        img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        img.setPixmap(build_splash_pixmap(tr))
        outer.addWidget(img)

        status = QLabel(update_status_text(tr, UPDATE_INFO))
        status.setObjectName("hint")
        status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status.setWordWrap(True)
        outer.addWidget(status)

        row = QHBoxLayout()
        row.setSpacing(8)
        btn_github = QPushButton("🌐  " + tr("wiz_github"))
        btn_github.setObjectName("ghost")
        btn_github.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl(GITHUB_URL)))
        row.addWidget(btn_github)
        btn_close = QPushButton(tr("close"))
        btn_close.setObjectName("primary")
        btn_close.clicked.connect(self.accept)
        row.addWidget(btn_close)
        outer.addLayout(row)

        self.setMinimumSize(440, 560)
        try:
            avail = self.screen().availableGeometry() if self.screen() else None
            if avail is not None:
                self.resize(min(460, avail.width() - 60),
                            min(640, avail.height() - 60))
            else:
                self.resize(452, 620)
        except Exception:
            self.resize(452, 620)


# ----------------------------------------------------------------------
# Dialog impostazioni / profili (compatto per 720p)
# ----------------------------------------------------------------------
class SettingsDialog(QDialog):
    def __init__(self, tr, parent=None):
        super().__init__(parent)
        self.tr_ = tr
        self.setWindowTitle(tr("dlg_title"))
        self._build_ui()
        self._reload_list()
        self.setMinimumSize(600, 420)
        try:
            avail = self.screen().availableGeometry() if self.screen() else None
            if avail is not None:
                self.resize(min(700, avail.width() - 60),
                            min(520, avail.height() - 60))
            else:
                self.resize(680, 500)
        except Exception:
            self.resize(680, 500)

    def _build_ui(self):
        tr = self.tr_
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(10)

        # ---- colonna sinistra: lista profili
        sx = QVBoxLayout()
        sx.addWidget(QLabel(tr("saved_profiles")))
        self.lista = QListWidget()
        self.lista.currentTextChanged.connect(self._load_in_form)
        sx.addWidget(self.lista)

        self.btn_delete = QPushButton(tr("delete_profile"))
        self.btn_delete.setObjectName("danger")
        self.btn_delete.clicked.connect(self._delete)
        sx.addWidget(self.btn_delete)
        layout.addLayout(sx, 1)

        # ---- colonna destra: form scrollabile (entra anche in 720p)
        right_col = QVBoxLayout()
        right_col.setSpacing(8)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll_content = QWidget()
        dx = QVBoxLayout(scroll_content)
        dx.setContentsMargins(2, 2, 6, 2)
        dx.setSpacing(8)
        box = QGroupBox(tr("profile_data"))
        form = QGridLayout(box)
        form.setVerticalSpacing(8)

        r = 0
        form.addWidget(QLabel(tr("profile_name")), r, 0)
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText(tr("name_placeholder"))
        form.addWidget(self.name_edit, r, 1); r += 1

        form.addWidget(QLabel(tr("energy_cost")), r, 0)
        self.energy_spin = QDoubleSpinBox()
        self.energy_spin.setDecimals(3)
        self.energy_spin.setRange(0.0, 10.0)
        self.energy_spin.setSingleStep(0.01)
        self.energy_spin.setValue(DEFAULT_PROFILE["costo_corrente"])
        form.addWidget(self.energy_spin, r, 1); r += 1

        form.addWidget(QLabel(tr("printer_power")), r, 0)
        self.power_spin = QDoubleSpinBox()
        self.power_spin.setDecimals(0)
        self.power_spin.setRange(1.0, 10000.0)
        self.power_spin.setSingleStep(50.0)
        self.power_spin.setValue(DEFAULT_PROFILE["consumo_w"])
        form.addWidget(self.power_spin, r, 1); r += 1

        form.addWidget(QLabel(tr("material_cost")), r, 0)
        self.material_spin = QDoubleSpinBox()
        self.material_spin.setDecimals(2)
        self.material_spin.setRange(0.0, 1000.0)
        self.material_spin.setSingleStep(0.5)
        self.material_spin.setValue(DEFAULT_PROFILE["costo_materiale"])
        form.addWidget(self.material_spin, r, 1); r += 1

        form.addWidget(QLabel(tr("design_rate")), r, 0)
        self.design_spin = QDoubleSpinBox()
        self.design_spin.setDecimals(2)
        self.design_spin.setRange(0.0, 500.0)
        self.design_spin.setSingleStep(0.5)
        self.design_spin.setValue(DEFAULT_PROFILE["costo_progettazione"])
        form.addWidget(self.design_spin, r, 1); r += 1

        form.addWidget(QLabel(tr("pc_power")), r, 0)
        self.pc_spin = QDoubleSpinBox()
        self.pc_spin.setDecimals(0)
        self.pc_spin.setRange(0.0, 5000.0)
        self.pc_spin.setSingleStep(10.0)
        self.pc_spin.setValue(DEFAULT_PROFILE["consumo_pc_w"])
        form.addWidget(self.pc_spin, r, 1); r += 1

        dx.addWidget(box)

        # ---- dati azienda (usati negli export al posto del logo Price3D)
        self.company_logo_src = None  # nuovo logo scelto, copiato al salvataggio
        cbox = QGroupBox(tr("wiz_company_title"))
        cform = QGridLayout(cbox)
        cform.setVerticalSpacing(8)
        cform.addWidget(QLabel(tr("wiz_company_name")), 0, 0)
        self.company_edit = QLineEdit()
        self.company_edit.setPlaceholderText(tr("wiz_company_name_ph"))
        cform.addWidget(self.company_edit, 0, 1)
        cform.addWidget(QLabel(tr("wiz_company_logo")), 1, 0)
        clogo_row = QHBoxLayout()
        clogo_row.setSpacing(8)
        self.company_preview = QLabel()
        self.company_preview.setFixedSize(48, 48)
        self.company_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.company_preview.setStyleSheet(
            "border: 1px dashed #3d4a66; border-radius: 8px;")
        clogo_row.addWidget(self.company_preview)
        clogo_col = QVBoxLayout()
        clogo_col.setSpacing(4)
        self.btn_company_logo = QPushButton(tr("wiz_choose_logo"))
        self.btn_company_logo.clicked.connect(self._choose_company_logo)
        clogo_col.addWidget(self.btn_company_logo)
        self.company_logo_label = QLabel("")
        self.company_logo_label.setObjectName("hint")
        self.company_logo_label.setWordWrap(True)
        clogo_col.addWidget(self.company_logo_label)
        clogo_row.addLayout(clogo_col, 1)
        cform.addLayout(clogo_row, 1, 1)
        dx.addWidget(cbox)
        comp = load_company()
        self.company_edit.setText(comp.get("company_name", ""))
        if comp.get("company_logo"):
            self.company_logo_label.setText(comp.get("company_logo"))
        self._refresh_company_preview()

        self.btn_save = QPushButton("💾 " + tr("save_profile"))
        self.btn_save.setObjectName("primary")
        self.btn_save.clicked.connect(self._save)
        dx.addWidget(self.btn_save)

        self.btn_new = QPushButton("🆕 " + tr("new_profile"))
        self.btn_new.clicked.connect(self._new)
        dx.addWidget(self.btn_new)

        dx.addStretch()
        scroll.setWidget(scroll_content)
        right_col.addWidget(scroll, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.button(QDialogButtonBox.StandardButton.Close).setText(tr("close"))
        buttons.rejected.connect(self.reject)
        right_col.addWidget(buttons)
        layout.addLayout(right_col, 2)

    def _refresh_company_preview(self):
        """Miniatura del logo aziendale salvato (o scelto, se in attesa)."""
        src = self.company_logo_src or company_logo_path()
        try:
            if src and Path(src).is_file():
                pm = QPixmap(src)
                if not pm.isNull():
                    self.company_preview.setPixmap(
                        pm.scaled(46, 46, Qt.AspectRatioMode.KeepAspectRatio,
                                  Qt.TransformationMode.SmoothTransformation))
                    return
        except Exception:
            pass
        self.company_preview.setPixmap(QPixmap())

    # ------------------------------------------------------------------
    def _reload_list(self):
        self.lista.blockSignals(True)
        self.lista.clear()
        for nome in load_profiles():
            self.lista.addItem(nome)
        self.lista.blockSignals(False)

    def _load_in_form(self, nome):
        p = load_profiles().get(nome)
        if not p:
            return
        self.name_edit.setText(p["nome"])
        self.energy_spin.setValue(p["costo_corrente"])
        self.power_spin.setValue(p["consumo_w"])
        self.material_spin.setValue(p["costo_materiale"])
        self.design_spin.setValue(p["costo_progettazione"])
        self.pc_spin.setValue(p["consumo_pc_w"])

    def _new(self):
        self.lista.clearSelection()
        self.name_edit.clear()
        self.energy_spin.setValue(DEFAULT_PROFILE["costo_corrente"])
        self.power_spin.setValue(DEFAULT_PROFILE["consumo_w"])
        self.material_spin.setValue(DEFAULT_PROFILE["costo_materiale"])
        self.design_spin.setValue(DEFAULT_PROFILE["costo_progettazione"])
        self.pc_spin.setValue(DEFAULT_PROFILE["consumo_pc_w"])
        self.name_edit.setFocus()

    def _choose_company_logo(self):
        tr = self.tr_
        path, _ = QFileDialog.getOpenFileName(
            self, tr("wiz_company_logo"), "",
            "Images (*.png *.jpg *.jpeg);;All files (*)"
        )
        if not path:
            return
        try:
            if QPixmap(path).isNull():
                raise ValueError("bad image")
            self.company_logo_src = path
            self.company_logo_label.setText(Path(path).name)
            self._refresh_company_preview()
        except Exception:
            QMessageBox.warning(self, tr("msg_invalid_title"), tr("msg_invalid_text"))

    def _save(self):
        tr = self.tr_
        nome = self.name_edit.text().strip()
        if not nome:
            QMessageBox.warning(self, tr("msg_noname_title"), tr("msg_noname_text"))
            return
        d = {
            "nome": nome,
            "costo_corrente": self.energy_spin.value(),
            "consumo_w": self.power_spin.value(),
            "costo_materiale": self.material_spin.value(),
            "costo_progettazione": self.design_spin.value(),
            "consumo_pc_w": self.pc_spin.value(),
        }
        path = save_profile(d)
        save_company(self.company_edit.text(), self.company_logo_src)
        self.company_logo_src = None
        comp = load_company()
        if comp.get("company_logo"):
            self.company_logo_label.setText(comp.get("company_logo"))
        self._reload_list()
        QMessageBox.information(self, tr("msg_saved_title"), f"{tr('msg_saved_text')}\n{path}")

    def _delete(self):
        tr = self.tr_
        item = self.lista.currentItem()
        if not item:
            QMessageBox.warning(self, tr("msg_nosel_title"), tr("msg_nosel_text"))
            return
        nome = item.text()
        conferma = QMessageBox.question(
            self, tr("confirm_del_title"),
            tr("confirm_del_text").format(name=nome),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if conferma != QMessageBox.StandardButton.Yes:
            return
        p = load_profiles().get(nome)
        if p and p.get("_file"):
            try:
                Path(p["_file"]).unlink(missing_ok=True)
            except Exception as e:
                QMessageBox.critical(self, tr("msg_export_err_title"), f"{tr('msg_delerr')} {e}")
                return
        self._reload_list()
        self._new()


# ----------------------------------------------------------------------
# Finestra principale
# ----------------------------------------------------------------------
class PriceCalculator(QWidget):
    def __init__(self):
        super().__init__()
        self.languages = load_languages()
        cfg = load_config()
        self.lang_code = cfg.get("language", "en")
        if self.lang_code not in self.languages:
            self.lang_code = "en" if "en" in self.languages else next(iter(self.languages))
        self.last_receipt = None
        self.last_calc = None
        self._build_ui()
        self._reload_profiles()
        self._retranslate()
        self._on_mode_changed(0)

        # 720p-friendly: minimo piccolo + dimensione iniziale adattata
        # allo schermo disponibile (il contenuto scorre dentro QScrollArea).
        self.setMinimumSize(620, 460)
        try:
            avail = self.screen().availableGeometry() if self.screen() else None
            if avail is not None:
                self.resize(min(880, avail.width() - 40),
                            min(860, avail.height() - 60))
            else:
                self.resize(860, 820)
        except Exception:
            self.resize(860, 820)

    # ------------------------------------------------------------------
    def tr_(self, key: str) -> str:
        lang = self.languages.get(self.lang_code, {})
        if key in lang:
            return lang[key]
        return self.languages.get("en", {}).get(key, key)

    # ------------------------------------------------------------------
    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 10, 14, 10)
        outer.setSpacing(8)

        # ---- intestazione brand: logo + Price3D
        brand_row = QHBoxLayout()
        brand_row.setSpacing(10)
        self.logo_label = QLabel()
        pm = app_logo_pixmap(44)
        if pm is not None:
            self.logo_label.setPixmap(pm)
            self.logo_label.setFixedSize(46, 46)
            self.logo_label.setScaledContents(False)
            self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        else:
            self.logo_label.setText("🖨️")
            self.logo_label.setStyleSheet("font-size: 26pt;")
        brand_row.addWidget(self.logo_label)

        title_col = QVBoxLayout()
        title_col.setSpacing(0)
        self.header = QLabel(APP_NAME)
        self.header.setObjectName("brandTitle")
        title_col.addWidget(self.header)
        self.subheader = QLabel()
        self.subheader.setObjectName("subheader")
        title_col.addWidget(self.subheader)
        brand_row.addLayout(title_col, 1)
        outer.addLayout(brand_row)

        # ---- area scrollabile con tutto il contenuto
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        outer.addWidget(scroll, 1)

        content = QWidget()
        scroll.setWidget(content)
        main_layout = QVBoxLayout(content)
        main_layout.setContentsMargins(2, 2, 2, 2)
        main_layout.setSpacing(8)
        self.main_layout = main_layout

        # ---- riga profilo + lingua + impostazioni
        self.box_printer = QGroupBox()
        riga = QHBoxLayout(self.box_printer)
        riga.setSpacing(8)
        self.lbl_profile = QLabel()
        riga.addWidget(self.lbl_profile)
        self.profile_combo = QComboBox()
        self.profile_combo.setMinimumWidth(140)
        self.profile_combo.setSizePolicy(QSizePolicy.Policy.Expanding,
                                         QSizePolicy.Policy.Fixed)
        riga.addWidget(self.profile_combo, 1)

        self.lbl_language = QLabel()
        riga.addWidget(self.lbl_language)
        self.lang_combo = QComboBox()
        for code, d in self.languages.items():
            self.lang_combo.addItem(d.get("_name", code), code)
        idx = self.lang_combo.findData(self.lang_code)
        if idx >= 0:
            self.lang_combo.setCurrentIndex(idx)
        self.lang_combo.currentIndexChanged.connect(self._change_language)
        riga.addWidget(self.lang_combo)

        self.btn_settings = QPushButton()
        self.btn_settings.clicked.connect(self._open_settings)
        riga.addWidget(self.btn_settings)
        main_layout.addWidget(self.box_printer)

        # ---- dati di stampa
        self.box_data = QGroupBox()
        grid = QGridLayout(self.box_data)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)

        self.lbl_mode = QLabel()
        grid.addWidget(self.lbl_mode, 0, 0)
        self.mode_combo = QComboBox()
        self.mode_combo.setSizePolicy(QSizePolicy.Policy.Expanding,
                                      QSizePolicy.Policy.Fixed)
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        grid.addWidget(self.mode_combo, 0, 1, 1, 3)

        self.lbl_hours = QLabel()
        grid.addWidget(self.lbl_hours, 1, 0)
        self.ore_edit = QLineEdit("0")
        self.ore_edit.setFixedWidth(70)
        grid.addWidget(self.ore_edit, 1, 1)

        self.lbl_minutes = QLabel()
        grid.addWidget(self.lbl_minutes, 1, 2)
        self.minuti_edit = QLineEdit("0")
        self.minuti_edit.setFixedWidth(70)
        grid.addWidget(self.minuti_edit, 1, 3)

        self.lbl_grams = QLabel()
        grid.addWidget(self.lbl_grams, 2, 0)
        self.grammi_edit = QLineEdit("0")
        self.grammi_edit.setFixedWidth(100)
        grid.addWidget(self.grammi_edit, 2, 1)

        self.lbl_design = QLabel()
        grid.addWidget(self.lbl_design, 2, 2)
        self.prog_edit = QLineEdit("0")
        self.prog_edit.setFixedWidth(100)
        grid.addWidget(self.prog_edit, 2, 3)

        # Riga multi-oggetto (visibile solo in modalita' multi)
        self.multi_widget = QWidget()
        multi_grid = QGridLayout(self.multi_widget)
        multi_grid.setContentsMargins(0, 0, 0, 0)
        multi_grid.setHorizontalSpacing(10)
        multi_grid.setVerticalSpacing(8)
        self.lbl_prints = QLabel()
        multi_grid.addWidget(self.lbl_prints, 0, 0)
        self.prints_spin = QSpinBox()
        self.prints_spin.setMinimum(1)
        self.prints_spin.setMaximum(9999)
        self.prints_spin.setValue(1)
        self.prints_spin.setFixedWidth(100)
        multi_grid.addWidget(self.prints_spin, 0, 1)
        self.lbl_objects = QLabel()
        multi_grid.addWidget(self.lbl_objects, 0, 2)
        self.objects_spin = QSpinBox()
        self.objects_spin.setMinimum(1)
        self.objects_spin.setMaximum(99999)
        self.objects_spin.setValue(1)
        self.objects_spin.setFixedWidth(100)
        multi_grid.addWidget(self.objects_spin, 0, 3)
        grid.addWidget(self.multi_widget, 3, 0, 1, 4)

        self.multi_hint = QLabel()
        self.multi_hint.setObjectName("hint")
        self.multi_hint.setWordWrap(True)
        grid.addWidget(self.multi_hint, 4, 0, 1, 4)
        grid.setColumnStretch(4, 1)
        main_layout.addWidget(self.box_data)

        # ---- oggetti extra
        self.box_extras = QGroupBox()
        extras_layout = QVBoxLayout(self.box_extras)
        extras_layout.setSpacing(6)

        self.extras_list = QListWidget()
        self.extras_list.setMaximumHeight(84)
        extras_layout.addWidget(self.extras_list)

        extras_form = QHBoxLayout()
        extras_form.setSpacing(6)
        self.extra_name_edit = QLineEdit()
        extras_form.addWidget(self.extra_name_edit, 2)
        self.extra_price_edit = QLineEdit()
        self.extra_price_edit.setFixedWidth(90)
        extras_form.addWidget(self.extra_price_edit)
        self.btn_add_extra = QPushButton()
        self.btn_add_extra.setObjectName("accent")
        self.btn_add_extra.clicked.connect(self._add_extra)
        extras_form.addWidget(self.btn_add_extra)
        self.btn_remove_extra = QPushButton()
        self.btn_remove_extra.setObjectName("ghost")
        self.btn_remove_extra.clicked.connect(self._remove_extra)
        extras_form.addWidget(self.btn_remove_extra)
        extras_layout.addLayout(extras_form)

        main_layout.addWidget(self.box_extras)
        self.extra_items = []

        # ---- checkbox opzioni
        chk_layout = QHBoxLayout()
        self.chk_design = QCheckBox()
        self.chk_design.setChecked(True)
        self.chk_design.toggled.connect(
            lambda on: self.prog_edit.setEnabled(on)
        )
        chk_layout.addWidget(self.chk_design)

        self.chk_round = QCheckBox()
        self.chk_round.setChecked(True)
        chk_layout.addWidget(self.chk_round)
        chk_layout.addStretch()
        main_layout.addLayout(chk_layout)

        # ---- pulsanti azione
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        self.calc_btn = QPushButton()
        self.calc_btn.setObjectName("primary")
        self.calc_btn.clicked.connect(self.calculate)
        self.calc_btn.setSizePolicy(QSizePolicy.Policy.Expanding,
                                    QSizePolicy.Policy.Fixed)
        btn_layout.addWidget(self.calc_btn)

        self.export_btn = QPushButton()
        self.export_btn.clicked.connect(self.export_pdf)
        self.export_btn.setSizePolicy(QSizePolicy.Policy.Expanding,
                                      QSizePolicy.Policy.Fixed)
        btn_layout.addWidget(self.export_btn)

        self.export_word_btn = QPushButton()
        self.export_word_btn.setObjectName("accent")
        self.export_word_btn.clicked.connect(self.export_docx)
        self.export_word_btn.setSizePolicy(QSizePolicy.Policy.Expanding,
                                           QSizePolicy.Policy.Fixed)
        btn_layout.addWidget(self.export_word_btn)
        main_layout.addLayout(btn_layout)

        # ---- scontrino
        self.receipt = QTextEdit()
        self.receipt.setReadOnly(True)
        self.receipt.setFont(QFont("Consolas", 10))
        self.receipt.setMinimumHeight(80)
        main_layout.addWidget(self.receipt, 1)

        self.footer = QLabel()
        self.footer.setObjectName("footer")
        self.footer.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.footer.setWordWrap(True)
        footer_row = QHBoxLayout()
        footer_row.setSpacing(8)
        self.btn_about = QPushButton("ⓘ")
        self.btn_about.setFixedSize(30, 26)
        self.btn_about.setToolTip(self.tr_("about_title"))
        self.btn_about.clicked.connect(self._open_about)
        footer_row.addWidget(self.btn_about)
        footer_row.addWidget(self.footer, 1)
        outer.addLayout(footer_row)

        # F1 (Windows/Linux) apre le informazioni sul programma
        self.about_action = QAction(self.tr_("about_title"), self)
        self.about_action.setShortcut("F1")
        self.about_action.triggered.connect(self._open_about)
        self.addAction(self.about_action)

        self._setup_macos_menu()

    # ------------------------------------------------------------------
    def _setup_macos_menu(self):
        """Su macOS la voce About finisce nel menu dell'app (dopo la mela)."""
        if sys.platform != "darwin":
            return
        try:
            self._mac_menu_bar = QMenuBar()
            app_menu = QMenu(APP_NAME, self._mac_menu_bar)
            about = QAction(self.tr_("about_title"), self)
            about.setMenuRole(QAction.MenuRole.AboutRole)
            about.triggered.connect(self._open_about)
            app_menu.addAction(about)
            self._mac_menu_bar.addMenu(app_menu)
        except Exception:
            pass

    def _open_about(self):
        AboutDialog(self.tr_, self).exec()

    # ------------------------------------------------------------------
    def _on_mode_changed(self, index: int):
        multi = (self.mode_combo.currentData() == "multi")
        self.multi_widget.setVisible(multi)
        self.multi_hint.setVisible(multi)
        # In multi-stampa tempo e materiale si intendono PER STAMPA
        tr = self.tr_
        if multi:
            self.lbl_hours.setText(tr("time_hours_multi"))
            self.lbl_grams.setText(tr("grams_multi"))
        else:
            self.lbl_hours.setText(tr("time_hours"))
            self.lbl_grams.setText(tr("grams"))

    # ------------------------------------------------------------------
    def _retranslate(self):
        tr = self.tr_
        self.setWindowTitle(APP_NAME)
        self.header.setText(APP_NAME)
        self.subheader.setText(tr("subheader"))
        self.box_printer.setTitle(tr("group_printer"))
        self.lbl_profile.setText(tr("profile"))
        self.lbl_language.setText(tr("language"))
        self.btn_settings.setText("⚙️  " + tr("settings"))
        self.box_data.setTitle(tr("group_data"))
        self.lbl_mode.setText(tr("mode_label"))
        # mode combo (ricostruita per tradurre le voci)
        self.mode_combo.blockSignals(True)
        prev = self.mode_combo.currentData()
        self.mode_combo.clear()
        self.mode_combo.addItem(tr("mode_single"), "single")
        self.mode_combo.addItem(tr("mode_multi"), "multi")
        if prev == "multi":
            self.mode_combo.setCurrentIndex(1)
        self.mode_combo.blockSignals(False)
        self.lbl_hours.setText(tr("time_hours"))
        self.lbl_minutes.setText(tr("minutes"))
        self.lbl_grams.setText(tr("grams"))
        self.lbl_design.setText(tr("design_hours"))
        self.lbl_prints.setText(tr("num_prints"))
        self.lbl_objects.setText(tr("total_objects"))
        self.multi_hint.setText(tr("multi_hint"))
        self.box_extras.setTitle(tr("group_extras"))
        self.extra_name_edit.setPlaceholderText(tr("extra_name_placeholder"))
        self.extra_price_edit.setPlaceholderText(tr("extra_price_placeholder"))
        self.btn_add_extra.setText("➕  " + tr("add_extra"))
        self.btn_remove_extra.setText("🗑️  " + tr("remove_extra"))
        self.chk_design.setText(tr("chk_design"))
        self.chk_round.setText(tr("chk_round"))
        self.calc_btn.setText("🧮  " + tr("calc"))
        self.export_btn.setText("📄  " + tr("export_pdf"))
        self.export_word_btn.setText("📝  " + tr("export_word"))
        self.footer.setText(tr("footer"))
        self.about_action.setText(tr("about_title"))
        self.btn_about.setToolTip(tr("about_title"))
        self._on_mode_changed(self.mode_combo.currentIndex())

    def _change_language(self):
        code = self.lang_combo.currentData()
        if not code:
            return
        self.lang_code = code
        cfg = load_config()
        cfg["language"] = code
        save_config(cfg)
        self._retranslate()

    # ------------------------------------------------------------------
    def _reload_profiles(self, select: str | None = None):
        current = select or self.profile_combo.currentText()
        self.profile_combo.blockSignals(True)
        self.profile_combo.clear()
        self.profiles = load_profiles()
        if not self.profiles:
            self.profiles = {DEFAULT_PROFILE["nome"]: dict(DEFAULT_PROFILE)}
        for nome in self.profiles:
            self.profile_combo.addItem(nome)
        idx = self.profile_combo.findText(current)
        if idx >= 0:
            self.profile_combo.setCurrentIndex(idx)
        self.profile_combo.blockSignals(False)

    def _open_settings(self):
        dlg = SettingsDialog(self.tr_, self)
        dlg.exec()
        self._reload_profiles()

    # ------------------------------------------------------------------
    def _add_extra(self):
        tr = self.tr_
        nome = self.extra_name_edit.text().strip()
        prezzo_txt = (self.extra_price_edit.text() or "").strip().replace(",", ".")
        try:
            prezzo = float(prezzo_txt)
        except ValueError:
            prezzo = None
        if not nome or prezzo is None or prezzo < 0:
            QMessageBox.warning(self, tr("msg_extra_invalid_title"), tr("msg_extra_invalid_text"))
            return
        self.extra_items.append({"nome": nome, "prezzo": prezzo})
        self.extras_list.addItem(f"{nome}  —  {euro(prezzo)}")
        self.extra_name_edit.clear()
        self.extra_price_edit.clear()
        self.extra_name_edit.setFocus()

    def _remove_extra(self):
        tr = self.tr_
        row = self.extras_list.currentRow()
        if row < 0:
            QMessageBox.warning(self, tr("msg_extra_nosel_title"), tr("msg_extra_nosel_text"))
            return
        self.extras_list.takeItem(row)
        del self.extra_items[row]

    # ------------------------------------------------------------------
    def _is_multi(self) -> bool:
        return self.mode_combo.currentData() == "multi"

    def calculate(self):
        tr = self.tr_
        try:
            ore = float((self.ore_edit.text() or "0").replace(",", "."))
            minuti = float((self.minuti_edit.text() or "0").replace(",", "."))
            grammi = float((self.grammi_edit.text() or "0").replace(",", "."))
            proj_ore = float((self.prog_edit.text() or "0").replace(",", "."))
            if self._is_multi():
                n_prints = int(self.prints_spin.value() or 1)
                n_objects = int(self.objects_spin.value() or 1)
            else:
                n_prints = 1
                n_objects = 1
        except ValueError:
            QMessageBox.warning(self, tr("msg_invalid_title"), tr("msg_invalid_text"))
            return

        p = self.profiles.get(self.profile_combo.currentText(), DEFAULT_PROFILE)

        # Tempo / costi PER STAMPA
        tempo = ore + (minuti / 60.0)
        material_pp = p["costo_materiale"] * (grammi / 1000.0)
        printer_kwh = (p["consumo_w"] / 1000.0) * tempo
        energy_kwh = printer_kwh
        energy_pp = p["costo_corrente"] * energy_kwh
        print_pp = material_pp + energy_pp

        # Costo di progettazione: una tantum per ordine
        if self.chk_design.isChecked():
            design_cost = p["costo_progettazione"] * proj_ore
        else:
            proj_ore = 0.0
            design_cost = 0.0

        # Oggetti extra: una tantum per ordine
        extras_total = sum(item["prezzo"] for item in self.extra_items)

        prints_total = print_pp * n_prints
        grand_total = prints_total + design_cost + extras_total

        if self.chk_round.isChecked():
            rounded_total = math.ceil(grand_total)
            diff = rounded_total - grand_total
            # Differenza assorbita nell'energia (ripartita sulle stampe)
            energy_pp += diff / n_prints
            print_pp += diff / n_prints
            prints_total += diff
            total_shown = float(rounded_total)
            total_str = f"{rounded_total:.0f} €"
        else:
            total_shown = grand_total
            total_str = euro(grand_total)

        per_piece = total_shown / n_objects

        self.last_receipt = self._format_receipt(
            p, n_prints, n_objects, grammi, energy_kwh, proj_ore, tempo,
            material_pp, energy_pp, print_pp, prints_total, design_cost,
            per_piece, total_shown, total_str,
            show_design=self.chk_design.isChecked(),
            extra_items=list(self.extra_items), extras_total=extras_total,
        )
        self.receipt.setPlainText(self.last_receipt)
        self.last_calc = {
            "profile": p, "n_prints": n_prints, "n_objects": n_objects,
            "multi": self._is_multi(), "grammi": grammi,
            "energy_kwh": energy_kwh, "proj_ore": proj_ore, "tempo": tempo,
            "material_pp": material_pp, "energy_pp": energy_pp,
            "print_pp": print_pp, "prints_total": prints_total,
            "design_cost": design_cost, "per_piece": per_piece,
            "total": total_shown,
            "total_str": total_str, "show_design": self.chk_design.isChecked(),
            "extra_items": list(self.extra_items), "extras_total": extras_total,
        }

    # ------------------------------------------------------------------
    @staticmethod
    def _row(left: str, right: str, width: int) -> str:
        dots = max(width - len(left) - len(right), 2)
        return f"{left} {'.' * dots} {right}"

    def _format_receipt(self, p, n_prints, n_objects, grammi, energy_kwh, proj_ore, tempo,
                        material_pp, energy_pp, print_pp, prints_total, design_cost,
                        per_piece, total, total_str,
                        show_design: bool, extra_items=None, extras_total: float = 0.0):
        tr = self.tr_
        L = 34
        top = "━" * L
        sep = "─" * (L - 2)
        r = self._row
        row_w = L - 1
        extra_items = extra_items or []

        lines = []
        lines.append(f"*{tr('r_title')}*")
        lines.append(f"{tr('r_time')}: {tempo:.2f} h    {tr('r_prints')}: {n_prints}    {tr('r_objects')}: {n_objects}")
        lines.append(sep)
        lines.append(f"*{tr('r_costs_per_piece')}*")
        lines.append(r(f"{tr('r_material')} ({grammi:.0f} g)", euro(material_pp), row_w))
        lines.append(r(f"{tr('r_energy')} ({energy_kwh:.2f} kWh)", euro(energy_pp), row_w))
        lines.append(r(tr("r_cost_per_print"), euro(print_pp), row_w))
        lines.append(sep)
        if n_prints > 1:
            lines.append(r(tr("r_prints_total") + f" (x{n_prints})", euro(prints_total), row_w))
        else:
            lines.append(r(tr("r_subtotal_total"), euro(prints_total), row_w))
        if show_design:
            lines.append(r(f"{tr('r_design')} ({proj_ore:.1f} h)", euro(design_cost), row_w))
        if extra_items:
            lines.append(sep)
            lines.append(f"*{tr('r_extras')}*")
            for item in extra_items:
                lines.append(r(item["nome"], euro(item["prezzo"]), row_w))
        lines.append(top)
        lines.append(f"*{tr('r_total')}: {total_str}*")
        if n_objects > 1:
            lines.append(r(tr("r_per_piece") + f" ({total_str} / {n_objects})", euro(per_piece), row_w))
        return "\n".join(lines)

    # ------------------------------------------------------------------
    def _default_logo(self) -> str | None:
        corp = company_logo_path()
        if corp:
            return corp
        if LOGO_FILE.is_file():
            return str(LOGO_FILE)
        return None

    def export_pdf(self):
        tr = self.tr_
        if not self.last_calc:
            QMessageBox.warning(self, tr("msg_nocalc_title"), tr("msg_nocalc_text"))
            return
        if not _REPORTLAB_OK:
            QMessageBox.critical(
                self, tr("msg_export_err_title"),
                "reportlab is required for PDF export. Install with: pip install reportlab"
            )
            return

        # 1) Logo: usa Logo.png dell'app, altrimenti chiedi
        logo_path = self._default_logo()
        if not logo_path:
            QMessageBox.information(
                self, tr("pdf_choose_logo_title"), tr("pdf_choose_logo_text")
            )
            picked, _ = QFileDialog.getOpenFileName(
                self, tr("pdf_choose_logo_title"), "",
                "PNG images (*.png);;All files (*)"
            )
            logo_path = picked or None

        # 2) Foto del prodotto (facoltativa)
        QMessageBox.information(
            self, tr("pdf_choose_photo_title"), tr("pdf_choose_photo_text")
        )
        photo_path, _ = QFileDialog.getOpenFileName(
            self, tr("pdf_choose_photo_title"), "",
            "Images (*.png *.jpg *.jpeg);;All files (*)"
        )

        # 3) Percorso di salvataggio del PDF
        default_name = "riepilogo_ordine.pdf"
        path, _ = QFileDialog.getSaveFileName(
            self, tr("save_receipt"), default_name,
            "PDF files (*.pdf);;All files (*)"
        )
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"

        try:
            self._build_pdf(path, logo_path, photo_path or None)
            QMessageBox.information(self, tr("msg_export_title"), f"{tr('msg_export_text')}\n{path}")
        except Exception as e:
            QMessageBox.critical(self, tr("msg_export_err_title"), f"{tr('msg_export_err_text')} {e}")

    def _build_pdf(self, path: str, logo_path: str | None, photo_path: str | None):
        """Genera il PDF 'Riepilogo Ordine e Accettazione'."""
        tr = self.tr_
        c = self.last_calc
        page_w, page_h = A4
        margin = 18 * mm
        content_w = page_w - 2 * margin

        cv = pdf_canvas.Canvas(path, pagesize=A4)

        ACCENT = colors.HexColor("#2f6f8f")
        DARK = colors.HexColor("#1e2430")
        GREY = colors.HexColor("#6b7280")
        LINE = colors.HexColor("#c8cdd6")

        y = page_h - margin

        # ---- intestazione: logo (quadrato, in alto a sinistra) + titolo
        header_h = 22 * mm
        if logo_path:
            try:
                img = ImageReader(logo_path)
                cv.drawImage(
                    img, margin, y - header_h,
                    width=header_h, height=header_h,
                    preserveAspectRatio=True, mask="auto"
                )
            except Exception:
                pass

        title_x = margin + header_h + 6 * mm if logo_path else margin
        corp_name = company_name()
        if corp_name:
            cv.setFillColor(DARK)
            cv.setFont("Helvetica-Bold", 15)
            cv.drawString(title_x, y - 8 * mm, corp_name)
            cv.setFillColor(ACCENT)
            cv.setFont("Helvetica-Bold", 11)
            cv.drawString(title_x, y - 13.5 * mm, tr("pdf_doc_title"))
            cv.setFillColor(GREY)
            cv.setFont("Helvetica", 9)
            cv.drawString(title_x, y - 18 * mm, f"{tr('pdf_date')}: {date.today().strftime('%d/%m/%Y')}")
        else:
            cv.setFillColor(DARK)
            cv.setFont("Helvetica-Bold", 15)
            cv.drawString(title_x, y - 8 * mm, tr("pdf_doc_title"))
            cv.setFillColor(GREY)
            cv.setFont("Helvetica", 9)
            cv.drawString(title_x, y - 14 * mm, f"{tr('pdf_date')}: {date.today().strftime('%d/%m/%Y')}")

        y -= header_h + 6 * mm
        cv.setStrokeColor(LINE)
        cv.setLineWidth(1)
        cv.line(margin, y, page_w - margin, y)
        y -= 8 * mm

        # ---- riepilogo lavorazione
        cv.setFillColor(DARK)
        cv.setFont("Helvetica-Bold", 11)
        cv.drawString(margin, y, f"{tr('r_time')}: {c['tempo']:.2f} h    {tr('r_prints')}: {c['n_prints']}    {tr('r_objects')}: {c['n_objects']}")
        y -= 10 * mm

        # ---- tabella costi
        cv.setFont("Helvetica-Bold", 11)
        cv.setFillColor(ACCENT)
        cv.drawString(margin, y, tr("r_costs_per_piece"))
        y -= 7 * mm

        def row(label: str, value: str, bold=False, indent=0):
            nonlocal y
            cv.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
            cv.setFillColor(DARK)
            cv.drawString(margin + indent, y, label)
            cv.drawRightString(page_w - margin, y, value)
            y -= 6.5 * mm

        row(f"{tr('r_material')} ({c['grammi']:.0f} g)", euro(c["material_pp"]))
        row(f"{tr('r_energy')} ({c['energy_kwh']:.2f} kWh)", euro(c["energy_pp"]))
        row(tr("r_cost_per_print"), euro(c["print_pp"]), bold=True)

        y -= 2 * mm
        cv.setStrokeColor(LINE)
        cv.line(margin, y, page_w - margin, y)
        y -= 7 * mm

        if c["n_prints"] > 1:
            row(tr("r_prints_total") + f" (x{c['n_prints']})", euro(c["prints_total"]))
        else:
            row(tr("r_subtotal_total"), euro(c["prints_total"]))
        if c["show_design"]:
            row(f"{tr('r_design')} ({c['proj_ore']:.1f} h)", euro(c["design_cost"]))

        extra_items = c.get("extra_items") or []
        if extra_items:
            y -= 2 * mm
            cv.setStrokeColor(LINE)
            cv.setLineWidth(1)
            cv.line(margin, y, page_w - margin, y)
            y -= 7 * mm
            cv.setFont("Helvetica-Bold", 11)
            cv.setFillColor(ACCENT)
            cv.drawString(margin, y, tr("r_extras"))
            y -= 7 * mm
            for item in extra_items:
                row(item["nome"], euro(item["prezzo"]))

        y -= 2 * mm
        cv.setStrokeColor(ACCENT)
        cv.setLineWidth(1.2)
        cv.line(margin, y, page_w - margin, y)
        y -= 10 * mm

        # ---- totale + prezzo al pezzo
        cv.setFont("Helvetica-Bold", 14)
        cv.setFillColor(ACCENT)
        cv.drawString(margin, y, tr("r_total"))
        cv.drawRightString(page_w - margin, y, c["total_str"])
        y -= 8 * mm
        if c["n_objects"] > 1:
            cv.setFont("Helvetica-Bold", 10)
            cv.drawString(margin, y, f"{tr('r_per_piece')} ({c['total_str']} / {c['n_objects']})")
            cv.drawRightString(page_w - margin, y, euro(c["per_piece"]))
            y -= 8 * mm
        else:
            y -= 4 * mm

        # ---- nota: non è una fattura
        cv.setFont("Helvetica-Oblique", 8.5)
        cv.setFillColor(GREY)
        y = self._draw_wrapped(cv, tr("pdf_disclaimer"), margin, y, content_w, 8.5)
        y -= 4 * mm
        cv.setFont("Helvetica", 9)
        cv.setFillColor(DARK)
        y = self._draw_wrapped(cv, tr("pdf_client_note"), margin, y, content_w, 9)
        y -= 8 * mm

        # ---- foto del progetto
        if photo_path:
            try:
                img = ImageReader(photo_path)
                iw, ih = img.getSize()
                max_w = content_w
                max_h = 70 * mm
                scale = min(max_w / iw, max_h / ih)
                draw_w, draw_h = iw * scale, ih * scale
                if y - draw_h < margin + 10 * mm:
                    cv.showPage()
                    y = page_h - margin
                cv.setFont("Helvetica-Bold", 9)
                cv.setFillColor(GREY)
                cv.drawString(margin, y, tr("pdf_project_photo"))
                y -= 4 * mm
                img_x = margin + (content_w - draw_w) / 2
                cv.drawImage(
                    img, img_x, y - draw_h, width=draw_w, height=draw_h,
                    preserveAspectRatio=True, mask="auto"
                )
                y -= draw_h + 6 * mm
            except Exception:
                pass

        cv.setFont("Helvetica", 7.5)
        cv.setFillColor(GREY)
        cv.drawCentredString(page_w / 2, margin / 2, tr("footer"))

        cv.save()

    @staticmethod
    def _draw_wrapped(cv, text: str, x: float, y: float, max_width: float, font_size: float,
                       font_name: str = "Helvetica", leading_mm: float = 4.6) -> float:
        """Disegna testo andando a capo automaticamente entro max_width. Ritorna la nuova y."""
        words = text.split()
        line = ""
        for word in words:
            trial = f"{line} {word}".strip()
            if stringWidth(trial, font_name, font_size) <= max_width:
                line = trial
            else:
                cv.drawString(x, y, line)
                y -= leading_mm * mm
                line = word
        if line:
            cv.drawString(x, y, line)
            y -= leading_mm * mm
        return y

    # ------------------------------------------------------------------
    def export_docx(self):
        tr = self.tr_
        if not self.last_calc:
            QMessageBox.warning(self, tr("msg_nocalc_title"), tr("msg_nocalc_text"))
            return
        if not _PYDOCX_OK:
            QMessageBox.critical(
                self, tr("msg_export_err_title"),
                "python-docx is required for Word export. Install with: pip install python-docx"
            )
            return

        # 1) Logo: usa Logo.png dell'app, altrimenti chiedi
        logo_path = self._default_logo()
        if not logo_path:
            QMessageBox.information(
                self, tr("pdf_choose_logo_title"), tr("pdf_choose_logo_text")
            )
            picked, _ = QFileDialog.getOpenFileName(
                self, tr("pdf_choose_logo_title"), "",
                "PNG images (*.png);;All files (*)"
            )
            logo_path = picked or None

        # 2) Foto del prodotto (facoltativa)
        QMessageBox.information(
            self, tr("pdf_choose_photo_title"), tr("pdf_choose_photo_text")
        )
        photo_path, _ = QFileDialog.getOpenFileName(
            self, tr("pdf_choose_photo_title"), "",
            "Images (*.png *.jpg *.jpeg);;All files (*)"
        )

        # 3) Percorso di salvataggio del documento Word
        default_name = "riepilogo_ordine.docx"
        path, _ = QFileDialog.getSaveFileName(
            self, tr("save_receipt"), default_name,
            "Word files (*.docx);;All files (*)"
        )
        if not path:
            return
        if not path.lower().endswith(".docx"):
            path += ".docx"

        try:
            self._build_docx(path, logo_path, photo_path or None)
            QMessageBox.information(self, tr("msg_export_title"), f"{tr('msg_export_text')}\n{path}")
        except Exception as e:
            QMessageBox.critical(self, tr("msg_export_err_title"), f"{tr('msg_export_err_text')} {e}")

    def _build_docx(self, path: str, logo_path: str | None, photo_path: str | None):
        """Genera il documento Word 'Riepilogo Ordine e Accettazione', stesso stile del PDF."""
        tr = self.tr_
        c = self.last_calc

        ACCENT = RGBColor(0x2F, 0x6F, 0x8F)
        DARK = RGBColor(0x1E, 0x24, 0x30)
        GREY = RGBColor(0x6B, 0x72, 0x80)
        LINE_GREY = "C8CDD6"
        LINE_ACCENT = "2F6F8F"

        doc = Document()

        section = doc.sections[0]
        section.page_height = Mm(297)
        section.page_width = Mm(210)
        section.top_margin = Mm(18)
        section.bottom_margin = Mm(18)
        section.left_margin = Mm(18)
        section.right_margin = Mm(18)

        content_w_mm = 210 - 36

        def no_table_borders(table):
            tbl = table._tbl
            tblPr = tbl.tblPr
            borders = OxmlElement('w:tblBorders')
            for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
                el = OxmlElement(f'w:{edge}')
                el.set(qn('w:val'), 'none')
                el.set(qn('w:sz'), '0')
                el.set(qn('w:space'), '0')
                borders.append(el)
            tblPr.append(borders)

        # ---- intestazione: logo + titolo/data
        header_table = doc.add_table(rows=1, cols=2)
        header_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        no_table_borders(header_table)
        header_table.columns[0].width = Mm(24)
        header_table.columns[1].width = Mm(content_w_mm - 24)
        logo_cell, title_cell = header_table.rows[0].cells
        logo_cell.width = Mm(24)
        title_cell.width = Mm(content_w_mm - 24)

        logo_p = logo_cell.paragraphs[0]
        if logo_path:
            try:
                run = logo_p.add_run()
                run.add_picture(logo_path, width=Mm(22), height=Mm(22))
            except Exception:
                pass

        title_p = title_cell.paragraphs[0]
        corp_name = company_name()
        if corp_name:
            title_run = title_p.add_run(corp_name)
            title_run.bold = True
            title_run.font.size = Pt(16)
            title_run.font.color.rgb = DARK
            sub_p = title_cell.add_paragraph()
            sub_run = sub_p.add_run(tr("pdf_doc_title"))
            sub_run.bold = True
            sub_run.font.size = Pt(11)
            sub_run.font.color.rgb = ACCENT
        else:
            title_run = title_p.add_run(tr("pdf_doc_title"))
            title_run.bold = True
            title_run.font.size = Pt(16)
            title_run.font.color.rgb = DARK

        date_p = title_cell.add_paragraph()
        date_run = date_p.add_run(f"{tr('pdf_date')}: {date.today().strftime('%d/%m/%Y')}")
        date_run.font.size = Pt(9)
        date_run.font.color.rgb = GREY

        sep_p = doc.add_paragraph()
        sep_p.paragraph_format.space_before = Pt(6)
        sep_p.paragraph_format.space_after = Pt(10)
        pPr = sep_p._p.get_or_add_pPr()
        pbdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:color'), LINE_GREY)
        pbdr.append(bottom)
        pPr.append(pbdr)

        # ---- tempo / stampe / oggetti
        info_p = doc.add_paragraph()
        info_run = info_p.add_run(
            f"{tr('r_time')}: {c['tempo']:.2f} h    {tr('r_prints')}: {c['n_prints']}    {tr('r_objects')}: {c['n_objects']}")
        info_run.bold = True
        info_run.font.size = Pt(11)
        info_run.font.color.rgb = DARK
        info_p.paragraph_format.space_after = Pt(10)

        section_p = doc.add_paragraph()
        section_run = section_p.add_run(tr("r_costs_per_piece"))
        section_run.bold = True
        section_run.font.size = Pt(11)
        section_run.font.color.rgb = ACCENT
        section_p.paragraph_format.space_after = Pt(4)

        def add_row_table(rows_data, top_border_after=None, accent_border_after=None):
            t = doc.add_table(rows=len(rows_data), cols=2)
            no_table_borders(t)
            t.columns[0].width = Mm(content_w_mm * 0.65)
            t.columns[1].width = Mm(content_w_mm * 0.35)
            for i, (label, value, bold) in enumerate(rows_data):
                lc, vc = t.rows[i].cells
                lc.width = Mm(content_w_mm * 0.65)
                vc.width = Mm(content_w_mm * 0.35)
                lp = lc.paragraphs[0]
                lr = lp.add_run(label)
                lr.bold = bold
                lr.font.size = Pt(10)
                lr.font.color.rgb = DARK
                vp = vc.paragraphs[0]
                vp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                vr = vp.add_run(value)
                vr.bold = bold
                vr.font.size = Pt(10)
                vr.font.color.rgb = DARK
                lp.paragraph_format.space_after = Pt(2)
                vp.paragraph_format.space_after = Pt(2)
            return t

        add_row_table([
            (f"{tr('r_material')} ({c['grammi']:.0f} g)", euro(c["material_pp"]), False),
            (f"{tr('r_energy')} ({c['energy_kwh']:.2f} kWh)", euro(c["energy_pp"]), False),
            (tr("r_cost_per_print"), euro(c["print_pp"]), True),
        ])

        sep2 = doc.add_paragraph()
        sep2.paragraph_format.space_after = Pt(8)
        pPr2 = sep2._p.get_or_add_pPr()
        pbdr2 = OxmlElement('w:pBdr')
        top2 = OxmlElement('w:top')
        top2.set(qn('w:val'), 'single')
        top2.set(qn('w:sz'), '4')
        top2.set(qn('w:color'), LINE_GREY)
        pbdr2.append(top2)
        pPr2.append(pbdr2)

        design_rows = []
        if c["n_prints"] > 1:
            design_rows.append((tr("r_prints_total") + f" (x{c['n_prints']})", euro(c["prints_total"]), False))
        else:
            design_rows.append((tr("r_subtotal_total"), euro(c["prints_total"]), False))
        if c["show_design"]:
            design_rows.append((f"{tr('r_design')} ({c['proj_ore']:.1f} h)", euro(c["design_cost"]), False))
        add_row_table(design_rows)

        extra_items = c.get("extra_items") or []
        if extra_items:
            sep_extra = doc.add_paragraph()
            sep_extra.paragraph_format.space_after = Pt(8)
            pPr_e = sep_extra._p.get_or_add_pPr()
            pbdr_e = OxmlElement('w:pBdr')
            top_e = OxmlElement('w:top')
            top_e.set(qn('w:val'), 'single')
            top_e.set(qn('w:sz'), '4')
            top_e.set(qn('w:color'), LINE_GREY)
            pbdr_e.append(top_e)
            pPr_e.append(pbdr_e)

            extras_title_p = doc.add_paragraph()
            extras_title_run = extras_title_p.add_run(tr("r_extras"))
            extras_title_run.bold = True
            extras_title_run.font.size = Pt(11)
            extras_title_run.font.color.rgb = ACCENT
            extras_title_p.paragraph_format.space_after = Pt(4)

            add_row_table([(item["nome"], euro(item["prezzo"]), False) for item in extra_items])

        sep3 = doc.add_paragraph()
        sep3.paragraph_format.space_before = Pt(6)
        sep3.paragraph_format.space_after = Pt(10)
        pPr3 = sep3._p.get_or_add_pPr()
        pbdr3 = OxmlElement('w:pBdr')
        top3 = OxmlElement('w:top')
        top3.set(qn('w:val'), 'single')
        top3.set(qn('w:sz'), '8')
        top3.set(qn('w:color'), LINE_ACCENT)
        pbdr3.append(top3)
        pPr3.append(pbdr3)

        # ---- totale + prezzo al pezzo
        total_t = doc.add_table(rows=1, cols=2)
        no_table_borders(total_t)
        total_t.columns[0].width = Mm(content_w_mm * 0.65)
        total_t.columns[1].width = Mm(content_w_mm * 0.35)
        lc, vc = total_t.rows[0].cells
        lc.width = Mm(content_w_mm * 0.65)
        vc.width = Mm(content_w_mm * 0.35)
        lp = lc.paragraphs[0]
        lr = lp.add_run(tr("r_total"))
        lr.bold = True
        lr.font.size = Pt(14)
        lr.font.color.rgb = ACCENT
        vp = vc.paragraphs[0]
        vp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        vr = vp.add_run(c["total_str"])
        vr.bold = True
        vr.font.size = Pt(14)
        vr.font.color.rgb = ACCENT

        if c["n_objects"] > 1:
            pp_t = doc.add_table(rows=1, cols=2)
            no_table_borders(pp_t)
            lc2, vc2 = pp_t.rows[0].cells
            lp2 = lc2.paragraphs[0]
            lr2 = lp2.add_run(f"{tr('r_per_piece')} ({c['total_str']} / {c['n_objects']})")
            lr2.bold = True
            lr2.font.size = Pt(10)
            lr2.font.color.rgb = DARK
            vp2 = vc2.paragraphs[0]
            vp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            vr2 = vp2.add_run(euro(c["per_piece"]))
            vr2.bold = True
            vr2.font.size = Pt(10)
            vr2.font.color.rgb = DARK

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

        disc_p = doc.add_paragraph()
        disc_run = disc_p.add_run(tr("pdf_disclaimer"))
        disc_run.italic = True
        disc_run.font.size = Pt(8.5)
        disc_run.font.color.rgb = GREY
        disc_p.paragraph_format.space_after = Pt(8)

        note_p = doc.add_paragraph()
        note_run = note_p.add_run(tr("pdf_client_note"))
        note_run.font.size = Pt(9)
        note_run.font.color.rgb = DARK
        note_p.paragraph_format.space_after = Pt(10)

        if photo_path:
            try:
                label_p = doc.add_paragraph()
                label_run = label_p.add_run(tr("pdf_project_photo"))
                label_run.bold = True
                label_run.font.size = Pt(9)
                label_run.font.color.rgb = GREY
                label_p.paragraph_format.space_after = Pt(4)

                photo_p = doc.add_paragraph()
                photo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                photo_run = photo_p.add_run()
                max_w_mm = content_w_mm
                max_h_mm = 70.0
                if _PIL_OK:
                    with Image.open(photo_path) as im:
                        iw, ih = im.size
                    scale = min(max_w_mm / iw, max_h_mm / ih)
                    draw_w_mm = iw * scale
                    photo_run.add_picture(photo_path, width=Mm(draw_w_mm))
                else:
                    photo_run.add_picture(photo_path, width=Mm(min(max_w_mm, 90)))
                photo_p.paragraph_format.space_after = Pt(10)
            except Exception:
                pass

        footer_p = doc.add_paragraph()
        footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_run = footer_p.add_run(tr("footer"))
        footer_run.font.size = Pt(7.5)
        footer_run.font.color.rgb = GREY

        doc.save(path)


# ----------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(__version__)
    app.setOrganizationName(APP_NAME)
    app.setStyleSheet(STYLE)
    if LOGO_FILE.is_file():
        try:
            app.setWindowIcon(QIcon(str(LOGO_FILE)))
        except Exception:
            pass

    _langs, _code = startup_language()

    def _wtr(key: str) -> str:
        return _langs.get(_code, {}).get(key, _langs.get("en", {}).get(key, key))

    # ---- splash + controllo aggiornamenti (release GitHub, non bloccante)
    splash = QSplashScreen(build_splash_pixmap(_wtr))
    splash.show()
    app.processEvents()
    splash.showMessage(_wtr("upd_check"),
                       Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter,
                       QColor("#8b96ad"))
    app.processEvents()
    checker = UpdateCheckThread()
    checker.start()
    t0 = time.monotonic()
    while checker.isRunning() and (time.monotonic() - t0) < 6.0:
        app.processEvents()
        QThread.msleep(120)
    if checker.tag:
        try:
            UPDATE_INFO["available"] = parse_version_tuple(checker.tag) > parse_version_tuple(__version__)
        except Exception:
            UPDATE_INFO["available"] = False
        UPDATE_INFO["tag"] = checker.tag
        UPDATE_INFO["url"] = checker.url
    else:
        UPDATE_INFO["error"] = checker.error
    splash.showMessage(update_status_text(_wtr, UPDATE_INFO),
                       Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter,
                       QColor("#8b96ad"))
    app.processEvents()
    while (time.monotonic() - t0) < 1.4:
        app.processEvents()
        QThread.msleep(100)

    if FIRST_RUN:
        wiz = FirstRunWizard(_wtr)
        splash.finish(wiz)
        wiz.exec()
    w = PriceCalculator()
    splash.finish(w)
    w.show()
    sys.exit(app.exec())
