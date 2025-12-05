# bootstrap_itc_v2.py - Version complète BASIC + FEED + Rapports + Livrables
"""
Bootstrap ITC Manager V2 - Application ITC Manager complète.

1) Place ce fichier dans un dossier (par ex. C:\ITC_Manager_App_V4) sous le nom bootstrap_itc_v2.py
2) Assure-toi d'avoir ton fichier Excel ITC_MASTER.xlsm (ou autre .xlsm) dans le même dossier
3) Dans l'environnement conda "itcmanager":
       cd /d C:\ITC_Manager_App_V4
       python bootstrap_itc_v2.py
4) Puis :
       python test_v2.py
       python main.py
"""

import os
import json
from pathlib import Path

# ---------------------------------------------------------------------
# Préparation des dossiers
# ---------------------------------------------------------------------
BASE = Path.cwd()
os.makedirs(BASE / "ui", exist_ok=True)
os.makedirs(BASE / "assets", exist_ok=True)

print("📦 Bootstrap ITC Manager V2...")
print("📍 Dossier:", BASE)

# Auto-détection du fichier .xlsm (le plus récent)
xlsm_candidates = sorted(
    BASE.glob("*.xlsm"),
    key=lambda p: p.stat().st_mtime,
    reverse=True
)
auto_excel = xlsm_candidates[0].name if xlsm_candidates else ""

# ---------------------------------------------------------------------
# 1) config.json
# ---------------------------------------------------------------------
config = {
    "excel_path": auto_excel,                     # Fichier Excel à ouvrir (auto-détecté)
    "sheet_name": "01.3-ITC MASTER WBS",         # Feuille principale pour BASIC/FEED
    "locale": "fr_FR",
    "theme": "dark",
    "window": {"width": 1100, "height": 750, "start_maximized": False},
    "excel": {"visible_on_first_run": False, "auto_save": True, "max_retries": 3},
    # Macros
    "macro_deliverables": "Module1.Generer_Livrables",
    "macro_reports": "Module1.Rapports_Numer_Graph_Export"
}
(Path("config.json")).write_text(
    json.dumps(config, indent=2, ensure_ascii=False),
    encoding="utf-8"
)
print("✅ config.json")

# ---------------------------------------------------------------------
# 2) schema_v2.json - Structure des onglets BASIC + FEED
# ---------------------------------------------------------------------
schema = {
    "sheet": "01.3-ITC MASTER WBS",
    "tabs": {
        "basic": {
            "label": "BASIC",
            "icon": "📊",
            "fields": [
                {
                    "id": "decret",
                    "label": "Décret",
                    "type": "select",
                    "options": ["Décret 92", "Décret 94"],
                    "cell": "C2",
                    "group": "Paramètres généraux"
                },
                {
                    "id": "phase",
                    "label": "Phase",
                    "type": "select",
                    "options": ["BASIC", "FEED", "REAL"],
                    "cell": "C3",
                    "group": "Paramètres généraux"
                },
                {
                    "id": "mh_directes",
                    "label": "Heures estimées directes Standard Construction",
                    "type": "int",
                    "cell": "J8",
                    "ui_format": "thousands_MHrs",
                    "placeholder": "10999999",
                    "group": "Estimation des heures"
                },
                {
                    "id": "facteur_prod",
                    "label": "Facteur de productivité",
                    "type": "float",
                    "cell": "J9",
                    "placeholder": "1.2",
                    "decimals": 1,
                    "group": "Estimation des heures"
                },
                {
                    "id": "date_debut",
                    "label": "Construction Starting date",
                    "type": "date",
                    "cell": "J14",
                    "locale": "fr_FR",
                    "group": "Planning"
                },
                {
                    "id": "date_fin",
                    "label": "End of construction work",
                    "type": "date",
                    "cell": "J15",
                    "locale": "fr_FR",
                    "group": "Planning"
                },
                {
                    "id": "taux_supervision",
                    "label": "Taux global de supervision (%)",
                    "type": "float",
                    "cell": "J19",
                    "ui_format": "percent_fr",
                    "placeholder": "12.5",
                    "decimals": 1,
                    "group": "Supervision"
                }
            ]
        },
        "feed_qty": {
            "label": "FEED - Quantités",
            "icon": "📦",
            "fields": [
                {
                    "id": "surface_construction",
                    "label": "Surface de construction au sol / Constructed plant area",
                    "type": "int",
                    "cell": "J26",
                    "ui_format": "thousands_m3",
                    "group": "Surfaces et volumes"
                },
                {
                    "id": "terrassement",
                    "label": "Terrassement / Site Preparation",
                    "type": "int",
                    "cell": "J27",
                    "ui_format": "thousands_m3",
                    "group": "Surfaces et volumes"
                },
                {
                    "id": "voirie_lourde",
                    "label": "Surface Générale des Voirie Lourdes / General Area of heavy road",
                    "type": "int",
                    "cell": "J28",
                    "ui_format": "thousands_m3",
                    "group": "Surfaces et volumes"
                },
                {
                    "id": "poids_equipment",
                    "label": "Poids des Equipment à installer / Equipment Installation",
                    "type": "int",
                    "cell": "J29",
                    "ui_format": "thousands_tons",
                    "group": "Équipements"
                },
                {
                    "id": "tuyauterie_48",
                    "label": "Tuyauterie Dia < 48\" / A/G Piping, Dia < 48\"",
                    "type": "int",
                    "cell": "J30",
                    "ui_format": "thousands_tons",
                    "group": "Tuyauterie"
                },
                {
                    "id": "hvac",
                    "label": "HVAC (Ml ou Tons)",
                    "type": "int",
                    "cell": "J31",
                    "ui_format": "thousands_tons",
                    "group": "Équipements"
                },
                {
                    "id": "tuyauterie_pression",
                    "label": "Tuyauterie sous pression U/G + tranchée",
                    "type": "int",
                    "cell": "J32",
                    "ui_format": "thousands_tons",
                    "group": "Tuyauterie"
                },
                {
                    "id": "tuyauterie_gravit",
                    "label": "Tuyauterie gravitationnelle U/G + tranchée",
                    "type": "int",
                    "cell": "J33",
                    "ui_format": "thousands_tons",
                    "group": "Tuyauterie"
                },
                {
                    "id": "isolation_piping",
                    "label": "Isolation / Insulation (piping)",
                    "type": "int",
                    "cell": "J34",
                    "ui_format": "thousands_m2",
                    "group": "Isolation"
                },
                {
                    "id": "instrumentation_cv",
                    "label": "Instrumentation (CV)",
                    "type": "int",
                    "cell": "J35",
                    "ui_format": "thousands_cv",
                    "group": "Instrumentation"
                },
                {
                    "id": "instrumentation_km",
                    "label": "Instrumentation (Km)",
                    "type": "int",
                    "cell": "J36",
                    "ui_format": "thousands_km",
                    "group": "Instrumentation"
                },
                {
                    "id": "electricite",
                    "label": "Électricité / Electricity",
                    "type": "int",
                    "cell": "J37",
                    "ui_format": "thousands_km",
                    "group": "Électricité"
                },
                {
                    "id": "genie_civil",
                    "label": "Travaux de Génie civil / Civil Works",
                    "type": "int",
                    "cell": "J38",
                    "ui_format": "thousands_m3",
                    "group": "Génie civil"
                },
                {
                    "id": "pieux",
                    "label": "Pieux / Piles",
                    "type": "int",
                    "cell": "J39",
                    "ui_format": "thousands_ml",
                    "group": "Génie civil"
                },
                {
                    "id": "ignifugation",
                    "label": "Ignifugation incendie / Fireproofing",
                    "type": "int",
                    "cell": "J40",
                    "ui_format": "thousands_m2",
                    "group": "Sécurité"
                },
                {
                    "id": "structure_metal",
                    "label": "Structure Métallique / Steel Structures",
                    "type": "int",
                    "cell": "J41",
                    "ui_format": "thousands_tons",
                    "group": "Structures"
                },
                {
                    "id": "equipment_incendie",
                    "label": "Equipment incendie / Fire Equipment",
                    "type": "int",
                    "cell": "J42",
                    "ui_format": "thousands_tons",
                    "group": "Sécurité"
                },
                {
                    "id": "reservoir_tons",
                    "label": "Réservoir / Tanks (Tons)",
                    "type": "int",
                    "cell": "J43",
                    "ui_format": "thousands_tons",
                    "group": "Réservoirs"
                },
                {
                    "id": "reservoir_m3",
                    "label": "Réservoir / Tanks (m³)",
                    "type": "int",
                    "cell": "J44",
                    "ui_format": "thousands_m3",
                    "group": "Réservoirs"
                },
                {
                    "id": "plus_grand_reservoir",
                    "label": "Le plus grand réservoir / Largest tank",
                    "type": "int",
                    "cell": "J45",
                    "ui_format": "thousands_m3",
                    "group": "Réservoirs"
                },
                {
                    "id": "batiments_surface",
                    "label": "Bâtiments / Buildings (surface à construire)",
                    "type": "int",
                    "cell": "J46",
                    "ui_format": "thousands_m2",
                    "group": "Bâtiments"
                },
                {
                    "id": "batiments_empreinte",
                    "label": "Bâtiments / Buildings (empreinte au sol)",
                    "type": "int",
                    "cell": "J47",
                    "ui_format": "thousands_m2",
                    "group": "Bâtiments"
                },
                {
                    "id": "isolation_bardage",
                    "label": "Isolation / Insulation (Bardage)",
                    "type": "int",
                    "cell": "J48",
                    "ui_format": "thousands_m2",
                    "group": "Isolation"
                },
                {
                    "id": "painting",
                    "label": "Painting",
                    "type": "int",
                    "cell": "J49",
                    "ui_format": "thousands_m2",
                    "group": "Finitions"
                },
                {
                    "id": "roads",
                    "label": "Routes / Roads (2 couches de 3 ou 4 cm)",
                    "type": "int",
                    "cell": "J50",
                    "ui_format": "thousands_m2",
                    "group": "Infrastructure"
                },
                {
                    "id": "salles_blanches",
                    "label": "Salles Blanches / White room",
                    "type": "int",
                    "cell": "J51",
                    "ui_format": "thousands_m2",
                    "group": "Spécialisé"
                }
            ]
        },
        "feed_surfaces": {
            "label": "FEED - Surfaces",
            "icon": "🗺️",
            "table_format": True,
            "wbs_options": [
                "Z.co / C.A.",
                "Z.C.E. / O.C.A",
                "Z.B.V. / C.V.",
                "Z.P.D.G. / G.A.",
                "Z.C.I. / I.C.A.",
                "Z.C. / C.F.A."
            ],
            "sections": {
                "sip": {
                    "label": "Surface Inside Plot (SIP)",
                    "rows": [
                        {"id": "sip_01", "label": "SIP 01", "wbs_cell": "I71", "surface_cell": "J71", "used_cell": "K71"},
                        {"id": "sip_02", "label": "SIP 02", "wbs_cell": "I72", "surface_cell": "J72", "used_cell": "K72"},
                        {"id": "sip_03", "label": "SIP 03", "wbs_cell": "I73", "surface_cell": "J73", "used_cell": "K73"},
                        {"id": "sip_04", "label": "SIP 04", "wbs_cell": "I74", "surface_cell": "J74", "used_cell": "K74"},
                        {"id": "sip_05", "label": "SIP 05", "wbs_cell": "I75", "surface_cell": "J75", "used_cell": "K75"},
                        {"id": "sip_06", "label": "SIP 06", "wbs_cell": "I76", "surface_cell": "J76", "used_cell": "K76"}
                    ]
                },
                "sop": {
                    "label": "Surface Outside Plot (SOP)",
                    "rows": [
                        {"id": "sop_01", "label": "SOP 01", "wbs_cell": "I77", "surface_cell": "J77", "used_cell": "K77"},
                        {"id": "sop_02", "label": "SOP 02", "wbs_cell": "I78", "surface_cell": "J78", "used_cell": "K78"},
                        {"id": "sop_03", "label": "SOP 03", "wbs_cell": "I79", "surface_cell": "J79", "used_cell": "K79"},
                        {"id": "sop_04", "label": "SOP 04", "wbs_cell": "I80", "surface_cell": "J80", "used_cell": "K80"}
                    ]
                },
                "wbs_recap": {
                    "label": "Récapitulatif WBS",
                    "readonly": True,
                    "rows": [
                        {"id": "wbs_ca", "label": "Z.co / C.A.", "used_cell": "K81", "display_cell": "J84"},
                        {"id": "wbs_oca", "label": "Z.C.E. / O.C.A", "used_cell": "K82", "display_cell": "J85"},
                        {"id": "wbs_cv", "label": "Z.B.V. / C.V.", "used_cell": "K83", "display_cell": "J86"},
                        {"id": "wbs_ga", "label": "Z.P.D.G. / G.A.", "used_cell": "K84", "display_cell": "J87"},
                        {"id": "wbs_ica", "label": "Z.C.I. / I.C.A.", "used_cell": "K85", "display_cell": "J88"},
                        {"id": "wbs_cfa", "label": "Z.C. / C.F.A.", "used_cell": "K86", "display_cell": "J89"}
                    ]
                }
            }
        },
        "feed_planning": {
            "label": "FEED - Planning",
            "icon": "📅",
            "table_format": True,
            "columns": [
                {
                    "id": "p1",
                    "label": "P1",
                    "subtitle": "Terrassement Fondations profondes",
                    "cell": "L14",
                    "readonly": True,
                    "source": "J14"
                },
                {
                    "id": "p2",
                    "label": "P2",
                    "subtitle": "Gros Œuvre",
                    "cell": "L15",
                    "type": "date"
                },
                {
                    "id": "p3",
                    "label": "P3",
                    "subtitle": "Second Œuvre ou Process",
                    "cell": "L16",
                    "type": "date"
                },
                {
                    "id": "p4",
                    "label": "P4",
                    "subtitle": "PCC et Hand-Over",
                    "cell": "L17",
                    "type": "date"
                },
                {
                    "id": "p5",
                    "label": "P5",
                    "subtitle": "À customiser si besoin",
                    "cell": "L18",
                    "type": "date"
                }
            ]
        }
    }
}
(Path("schema_v2.json")).write_text(
    json.dumps(schema, indent=2, ensure_ascii=False),
    encoding="utf-8"
)
print("✅ schema_v2.json (structure complète)")

# ---------------------------------------------------------------------
# 3) excel_client.py
# ---------------------------------------------------------------------
excel_client_code = r'''"""
Excel headless client (xlwings/COM) — self-healing & no sheets enumeration.
"""
import xlwings as xw
import time
from typing import Any
from pathlib import Path
from contextlib import contextmanager

class ExcelClient:
    def __init__(self, filepath: str, sheet_name: str = "01.3-ITC MASTER WBS",
                 visible: bool = False, max_retries: int = 3):
        if not Path(filepath).is_absolute():
            filepath = Path.cwd() / filepath
        self.filepath = Path(filepath)
        self.sheet_name = sheet_name
        self.visible = visible
        self.max_retries = max_retries
        self.app: xw.App | None = None
        self.book: xw.Book | None = None
        self._calc_mode = None

    def _spawn_app(self):
        self.app = xw.App(visible=self.visible, add_book=False)
        try:
            api = self.app.api
            self._calc_mode = api.Calculation
            api.ScreenUpdating = False
            api.DisplayAlerts = False
            api.EnableEvents = False
        except Exception:
            pass

    def _open_book(self):
        if self.app is None:
            self._spawn_app()
        try:
            self.book = self.app.books.open(str(self.filepath))
        except Exception:
            self.book = self.app.books.open(str(self.filepath), read_only=True)

    def _reopen_book(self):
        try:
            if self.book:
                self.book.close()
        except Exception:
            pass
        self._open_book()

    def _get_sheet(self):
        if self.app is None or self.book is None:
            self._open_book()
        try:
            return self.book.sheets[self.sheet_name]
        except Exception:
            try:
                ws = self.book.api.Worksheets
                names = [ws.Item(i).Name for i in range(1, ws.Count + 1)]
            except Exception:
                names = []
            self._reopen_book()
            try:
                return self.book.sheets[self.sheet_name]
            except Exception:
                raise ValueError(
                    f"Sheet '{self.sheet_name}' not found. "
                    f"Available: {', '.join(names) if names else '(unavailable)'}"
                )

    @contextmanager
    def _retry(self, op: str, attempts: int | None = None):
        tries = attempts or self.max_retries
        for i in range(tries):
            try:
                yield
                return
            except Exception as e:
                if i < tries - 1 and any(k in str(e) for k in ("RPC", "COM", "rejected", "disconnected", "busy")):
                    time.sleep(0.6 * (i + 1))
                    continue
                raise

    def connect(self):
        with self._retry("connect"):
            if not self.filepath.exists():
                raise FileNotFoundError(f"Excel file not found: {self.filepath}")
            self._open_book()
            sh = self._get_sheet()
            try:
                sh.activate()
            except Exception:
                pass

    def read_cell(self, address: str) -> Any:
        with self._retry(f"read {address}"):
            sh = self._get_sheet()
            try:
                return sh.range(address).value
            except AttributeError as e:
                if "Range" in str(e) or "range" in str(e):
                    self._reopen_book()
                    sh = self._get_sheet()
                    return sh.range(address).value
                raise

    def write_cell(self, address: str, value: Any):
        with self._retry(f"write {address}"):
            sh = self._get_sheet()
            try:
                sh.range(address).value = value
            except AttributeError as e:
                if "Range" in str(e) or "range" in str(e):
                    self._reopen_book()
                    sh = self._get_sheet()
                    sh.range(address).value = value
                else:
                    raise

    def write_cells_batch(self, cell_values: dict[str, Any]):
        """Écrire plusieurs cellules en batch"""
        with self._retry("batch write"):
            sh = self._get_sheet()
            if self.app:
                old_calc = self.app.api.Calculation
                self.app.api.Calculation = -4135  # xlCalculationManual
            try:
                for address, value in cell_values.items():
                    sh.range(address).value = value
            finally:
                if self.app:
                    self.app.api.Calculation = old_calc

    def run_macro(self, macro_name: str):
        with self._retry(f"macro {macro_name}"):
            if self.book is None:
                self._open_book()
            self.app.api.Run(macro_name)

    def save(self):
        with self._retry("save"):
            if self.book:
                self.book.save()

    def close(self):
        try:
            if self.app:
                try:
                    api = self.app.api
                    if self._calc_mode is not None:
                        api.Calculation = self._calc_mode
                    api.EnableEvents = True
                    api.ScreenUpdating = True
                    api.DisplayAlerts = True
                except Exception:
                    pass
            if self.book:
                self.book.close()
            if self.app:
                self.app.quit()
        except Exception:
            pass

    def set_visible(self, visible: bool):
        if self.app:
            self.app.visible = visible
'''
(Path("excel_client.py")).write_text(excel_client_code, encoding="utf-8")
print("✅ excel_client.py")

# ---------------------------------------------------------------------
# 4) main.py
# ---------------------------------------------------------------------
main_code = r'''"""
ITC Manager V2 — entrypoint
"""
import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPalette, QColor

def apply_theme(app):
    app.setStyle("Fusion")
    p = QPalette()
    p.setColor(QPalette.Window, QColor(45,45,48))
    p.setColor(QPalette.WindowText, Qt.white)
    p.setColor(QPalette.Base, QColor(60,60,64))
    p.setColor(QPalette.AlternateBase, QColor(70,70,74))
    p.setColor(QPalette.Text, Qt.white)
    p.setColor(QPalette.Button, QColor(70,70,74))
    p.setColor(QPalette.ButtonText, Qt.white)
    p.setColor(QPalette.Link, QColor(74,144,226))
    p.setColor(QPalette.Highlight, QColor(74,144,226))
    p.setColor(QPalette.HighlightedText, Qt.black)
    app.setPalette(p)

def main():
    app = QApplication(sys.argv)
    from ui.main_window_v2 import MainWindow
    apply_theme(app)
    w = MainWindow()
    w.show()
    QTimer.singleShot(500, w.connect_excel)
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
'''
(Path("main.py")).write_text(main_code, encoding="utf-8")
print("✅ main.py")

# ---------------------------------------------------------------------
# 5) ui/main_window_v2.py
# ---------------------------------------------------------------------
main_window_code = r'''"""
Main UI V2 - BASIC + FEED + Rapports + Livrables
"""
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QTabWidget, QScrollArea, QFormLayout,
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit,
    QGroupBox, QGridLayout, QProgressBar, QMessageBox,
    QGraphicsDropShadowEffect
)
from PySide6.QtCore import Qt, QTimer, QPropertyAnimation, QDate, QObject, Property
from PySide6.QtGui import QPalette, QColor, QIcon, QAction

import json
from pathlib import Path
from datetime import datetime, date
from typing import Dict, Any, List

# ---------------------------------------------------------------------
# Bouton animé réutilisé pour Rapports & Livrables
# ---------------------------------------------------------------------
class AnimatedButton(QPushButton):
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.animation = QPropertyAnimation(self, b"glow_intensity")
        self.animation.setDuration(1000)
        self.animation.setLoopCount(-1)
        self.animation.setStartValue(0)
        self.animation.setEndValue(100)
        self.shadow = QGraphicsDropShadowEffect()
        self.shadow.setBlurRadius(20)
        self.shadow.setOffset(0, 0)
        self.setGraphicsEffect(self.shadow)
        self.setStyleSheet("""
            QPushButton{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:1,
                    stop:0 #FF6B35, stop:1 #F72D7A);
                color:white;
                border-radius:25px;
                font-size:16px;
                font-weight:bold;
                padding:15px 30px;
                border:2px solid rgba(255,255,255,0.2)
            }
            QPushButton:hover{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:1,
                    stop:0 #FF7F50, stop:1 #FF1493)
            }
            QPushButton:pressed{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:1,
                    stop:0 #E55A2B, stop:1 #DD1B6A)
            }
        """)

    def enterEvent(self, e):
        self.animation.start()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self.animation.stop()
        self.shadow.setColor(QColor(0,0,0,50))
        super().leaveEvent(e)

    @Property(int)
    def glow_intensity(self):
        return getattr(self, "_glow_intensity", 0)

    @glow_intensity.setter
    def glow_intensity(self, v):
        self._glow_intensity = v
        self.shadow.setColor(QColor(255,107,53, int(v*2.55)))
        self.shadow.setBlurRadius(20 + v//5)

# ---------------------------------------------------------------------
# Onglets d'inputs BASIC / FEED
# ---------------------------------------------------------------------
class BaseInputTab(QWidget):
    """Classe de base pour les onglets d'inputs"""
    def __init__(self, tab_config: Dict, parent=None):
        super().__init__(parent)
        self.tab_config = tab_config
        self.widgets = {}
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        title = QLabel(f"{self.tab_config.get('icon', '📝')} {self.tab_config.get('label', 'Inputs')}")
        title.setStyleSheet("font-size:20px;font-weight:bold;padding:10px")
        layout.addWidget(title)

        scroll = QScrollArea()
        scroll_widget = QWidget()
        form_layout = QFormLayout()
        form_layout.setSpacing(12)

        groups = {}
        for field in self.tab_config.get('fields', []):
            group = field.get('group', 'Général')
            groups.setdefault(group, []).append(field)

        for group_name, fields in groups.items():
            group_label = QLabel(f"▸ {group_name}")
            group_label.setStyleSheet("font-weight:bold;color:#4A90E2;margin-top:10px")
            form_layout.addRow(group_label)

            for field in fields:
                widget = self.create_widget(field)
                if widget:
                    self.widgets[field['id']] = {
                        'widget': widget,
                        'cell': field['cell'],
                        'type': field['type'],
                        'ui_format': field.get('ui_format')
                    }
                    form_layout.addRow(field['label'], widget)

        scroll_widget.setLayout(form_layout)
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)
        self.setLayout(layout)

    def create_widget(self, field: Dict[str, Any]) -> QWidget:
        t = field['type']

        if t == 'select':
            w = QComboBox()
            w.addItems(field['options'])
            w.setStyleSheet("padding:5px")

        elif t == 'int':
            ui_format = field.get('ui_format', '')
            if 'thousands' in ui_format:
                w = QLineEdit()
                w.setPlaceholderText(field.get('placeholder', ''))
                if 'MHrs' in ui_format:
                    w.textChanged.connect(lambda: self.format_with_unit(w, 'MHrs'))
                elif 'tons' in ui_format:
                    w.textChanged.connect(lambda: self.format_with_unit(w, 'Tons'))
                elif 'm3' in ui_format:
                    w.textChanged.connect(lambda: self.format_with_unit(w, 'm³'))
                elif 'm2' in ui_format:
                    w.textChanged.connect(lambda: self.format_with_unit(w, 'm²'))
                elif 'cv' in ui_format:
                    w.textChanged.connect(lambda: self.format_with_unit(w, 'CV'))
                elif 'km' in ui_format:
                    w.textChanged.connect(lambda: self.format_with_unit(w, 'Km'))
                elif 'ml' in ui_format:
                    w.textChanged.connect(lambda: self.format_with_unit(w, 'ml'))
            else:
                w = QSpinBox()
                w.setMaximum(999_999_999)

        elif t == 'float':
            w = QDoubleSpinBox()
            w.setDecimals(field.get('decimals', 2))
            w.setMaximum(999_999.99)
            if field.get('ui_format') == 'percent_fr':
                w.setSuffix(' %')
                w.setMaximum(100.0)

        elif t == 'date':
            w = QDateEdit()
            w.setCalendarPopup(True)
            w.setDate(QDate.currentDate())
            w.setDisplayFormat('dd/MM/yyyy')

        else:
            w = QLineEdit()

        return w

    def format_with_unit(self, widget: QLineEdit, unit: str):
        text = widget.text()
        for u in ['MHrs', 'Tons', 'm³', 'm²', 'CV', 'Km', 'ml']:
            text = text.replace(u, '')
        text = text.replace(' ', '')

        if text.isdigit():
            formatted = f"{int(text):,}".replace(',', ' ') + f" {unit}"
            cursor_pos = widget.cursorPosition()
            widget.blockSignals(True)
            widget.setText(formatted)
            widget.setCursorPosition(min(cursor_pos, len(formatted)))
            widget.blockSignals(False)

    def get_values(self) -> Dict[str, Any]:
        vals = {}
        for field_id, info in self.widgets.items():
            w = info['widget']
            cell = info['cell']

            if isinstance(w, QComboBox):
                vals[cell] = w.currentText()
            elif isinstance(w, QLineEdit):
                text = w.text()
                for unit in ['MHrs', 'Tons', 'm³', 'm²', 'CV', 'Km', 'ml']:
                    text = text.replace(unit, '')
                text = text.replace(' ', '')
                vals[cell] = int(text) if text.isdigit() else 0
            elif isinstance(w, QSpinBox):
                vals[cell] = int(w.value())
            elif isinstance(w, QDoubleSpinBox):
                vals[cell] = float(w.value())
            elif isinstance(w, QDateEdit):
                try:
                    vals[cell] = w.date().toPython()
                except Exception:
                    d = w.date()
                    vals[cell] = (d.year(), d.month(), d.day())
        return vals

    def set_values(self, values: Dict[str, Any]):
        for field_id, info in self.widgets.items():
            cell = info['cell']
            if cell not in values:
                continue

            w = info['widget']
            val = values[cell]

            if isinstance(w, QComboBox):
                idx = w.findText(str(val))
                if idx >= 0:
                    w.setCurrentIndex(idx)

            elif isinstance(w, QLineEdit):
                if val not in (None, ""):
                    ui_format = info.get('ui_format', '')
                    try:
                        if 'thousands' in ui_format:
                            formatted = f"{int(float(val)):,}".replace(',', ' ')
                            if 'MHrs' in ui_format:
                                formatted += " MHrs"
                            elif 'tons' in ui_format:
                                formatted += " Tons"
                            elif 'm3' in ui_format:
                                formatted += " m³"
                            elif 'm2' in ui_format:
                                formatted += " m²"
                            elif 'cv' in ui_format:
                                formatted += " CV"
                            elif 'km' in ui_format:
                                formatted += " Km"
                            elif 'ml' in ui_format:
                                formatted += " ml"
                            w.setText(formatted)
                        else:
                            w.setText(str(val))
                    except Exception:
                        w.setText(str(val) if val else "")

            elif isinstance(w, QSpinBox):
                try:
                    w.setValue(int(float(val)))
                except Exception:
                    pass

            elif isinstance(w, QDoubleSpinBox):
                try:
                    w.setValue(float(val))
                except Exception:
                    pass

            elif isinstance(w, QDateEdit) and val:
                if isinstance(val, datetime):
                    w.setDate(QDate(val.year, val.month, val.day))
                elif isinstance(val, date):
                    w.setDate(QDate(val.year, val.month, val.day))
                elif isinstance(val, str):
                    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                        try:
                            d = datetime.strptime(val, fmt).date()
                            w.setDate(QDate(d.year, d.month, d.day))
                            break
                        except Exception:
                            pass

# ---------------------------------------------------------------------
# FEED - Surfaces
# ---------------------------------------------------------------------
class FeedSurfacesTab(QWidget):
    def __init__(self, tab_config: Dict, excel_client=None, parent=None):
        super().__init__(parent)
        self.tab_config = tab_config
        self.excel_client = excel_client
        self.widgets = {}
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        title = QLabel(f"{self.tab_config.get('icon', '🗺️')} {self.tab_config.get('label', 'FEED - Surfaces')}")
        title.setStyleSheet("font-size:20px;font-weight:bold;padding:10px")
        layout.addWidget(title)

        wbs_options = self.tab_config.get('wbs_options', [])

        scroll = QScrollArea()
        scroll_widget = QWidget()
        main_layout = QVBoxLayout(scroll_widget)

        # SIP
        sip_section = self.tab_config['sections']['sip']
        sip_group = QGroupBox(sip_section['label'])
        sip_layout = QGridLayout()
        sip_layout.addWidget(QLabel("Terrain"), 0, 0)
        sip_layout.addWidget(QLabel("WBS Associé"), 0, 1)
        sip_layout.addWidget(QLabel("Surface (m²)"), 0, 2)
        sip_layout.addWidget(QLabel("Utilisé?"), 0, 3)

        for i, row in enumerate(sip_section['rows'], 1):
            sip_layout.addWidget(QLabel(row['label']), i, 0)

            wbs_combo = QComboBox()
            wbs_combo.addItems(wbs_options)
            sip_layout.addWidget(wbs_combo, i, 1)
            self.widgets[row['wbs_cell']] = {
                'widget': wbs_combo,
                'cell': row['wbs_cell'],
                'type': 'select'
            }

            surface_edit = QLineEdit()
            surface_edit.setPlaceholderText("0")
            surface_edit.textChanged.connect(lambda t, w=surface_edit: self.format_surface(w))
            sip_layout.addWidget(surface_edit, i, 2)
            self.widgets[row['surface_cell']] = {
                'widget': surface_edit,
                'cell': row['surface_cell'],
                'type': 'int'
            }

            used_combo = QComboBox()
            used_combo.addItems(["", "X"])
            sip_layout.addWidget(used_combo, i, 3)
            self.widgets[row['used_cell']] = {
                'widget': used_combo,
                'cell': row['used_cell'],
                'type': 'select'
            }

        sip_group.setLayout(sip_layout)
        main_layout.addWidget(sip_group)

        # SOP
        sop_section = self.tab_config['sections']['sop']
        sop_group = QGroupBox(sop_section['label'])
        sop_layout = QGridLayout()
        sop_layout.addWidget(QLabel("Terrain"), 0, 0)
        sop_layout.addWidget(QLabel("WBS Associé"), 0, 1)
        sop_layout.addWidget(QLabel("Surface (m²)"), 0, 2)
        sop_layout.addWidget(QLabel("Utilisé?"), 0, 3)

        for i, row in enumerate(sop_section['rows'], 1):
            sop_layout.addWidget(QLabel(row['label']), i, 0)

            wbs_combo = QComboBox()
            wbs_combo.addItems(wbs_options)
            sop_layout.addWidget(wbs_combo, i, 1)
            self.widgets[row['wbs_cell']] = {
                'widget': wbs_combo,
                'cell': row['wbs_cell'],
                'type': 'select'
            }

            surface_edit = QLineEdit()
            surface_edit.setPlaceholderText("0")
            surface_edit.textChanged.connect(lambda t, w=surface_edit: self.format_surface(w))
            sop_layout.addWidget(surface_edit, i, 2)
            self.widgets[row['surface_cell']] = {
                'widget': surface_edit,
                'cell': row['surface_cell'],
                'type': 'int'
            }

            used_combo = QComboBox()
            used_combo.addItems(["", "X"])
            sop_layout.addWidget(used_combo, i, 3)
            self.widgets[row['used_cell']] = {
                'widget': used_combo,
                'cell': row['used_cell'],
                'type': 'select'
            }

        sop_group.setLayout(sop_layout)
        main_layout.addWidget(sop_group)

        # Récap WBS
        wbs_recap = self.tab_config['sections']['wbs_recap']
        recap_group = QGroupBox(wbs_recap['label'])
        recap_layout = QGridLayout()
        recap_layout.addWidget(QLabel("WBS"), 0, 0)
        recap_layout.addWidget(QLabel("Utilisé?"), 0, 1)
        recap_layout.addWidget(QLabel("Surface totale (m²)"), 0, 2)

        for i, row in enumerate(wbs_recap['rows'], 1):
            recap_layout.addWidget(QLabel(row['label']), i, 0)

            used_combo = QComboBox()
            used_combo.addItems(["", "X"])
            recap_layout.addWidget(used_combo, i, 1)
            self.widgets[row['used_cell']] = {
                'widget': used_combo,
                'cell': row['used_cell'],
                'type': 'select'
            }

            display_label = QLabel("0 m²")
            display_label.setStyleSheet(
                "background-color: #f0f0f0; padding: 5px; border-radius: 3px;"
            )
            recap_layout.addWidget(display_label, i, 2)
            self.widgets[row['display_cell']] = {
                'widget': display_label,
                'cell': row['display_cell'],
                'type': 'readonly',
                'wbs': row['label']
            }

        recap_group.setLayout(recap_layout)
        main_layout.addWidget(recap_group)

        calc_btn = QPushButton("🔄 Calculer les totaux")
        calc_btn.clicked.connect(self.calculate_totals)
        main_layout.addWidget(calc_btn)

        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)
        self.setLayout(layout)

    def format_surface(self, widget: QLineEdit):
        text = widget.text().replace(' ', '').replace('m²', '')
        if text.isdigit():
            formatted = f"{int(text):,}".replace(',', ' ') + " m²"
            cursor_pos = widget.cursorPosition()
            widget.blockSignals(True)
            widget.setText(formatted)
            widget.setCursorPosition(min(cursor_pos, len(formatted)))
            widget.blockSignals(False)

    def calculate_totals(self):
        wbs_totals = {}
        wbs_options = self.tab_config.get('wbs_options', [])
        for wbs in wbs_options:
            wbs_totals[wbs] = 0

        for section_key in ['sip', 'sop']:
            section = self.tab_config['sections'][section_key]
            for row in section['rows']:
                used_widget = self.widgets[row['used_cell']]['widget']
                if used_widget.currentText() == "X":
                    wbs_widget = self.widgets[row['wbs_cell']]['widget']
                    wbs = wbs_widget.currentText()

                    surface_widget = self.widgets[row['surface_cell']]['widget']
                    text = surface_widget.text().replace(' ', '').replace('m²', '')
                    if text.isdigit():
                        wbs_totals[wbs] += int(text)

        recap_section = self.tab_config['sections']['wbs_recap']
        for row in recap_section['rows']:
            if row['display_cell'] in self.widgets:
                widget_info = self.widgets[row['display_cell']]
                wbs = widget_info.get('wbs', '')
                total = wbs_totals.get(wbs, 0)
                widget_info['widget'].setText(
                    f"{total:,}".replace(',', ' ') + " m²"
                )

    def get_values(self) -> Dict[str, Any]:
        vals = {}
        for cell, info in self.widgets.items():
            w = info['widget']
            if info['type'] == 'readonly':
                continue
            elif isinstance(w, QComboBox):
                vals[cell] = w.currentText()
            elif isinstance(w, QLineEdit):
                text = w.text().replace(' ', '').replace('m²', '')
                vals[cell] = int(text) if text.isdigit() else 0
        return vals

    def set_values(self, values: Dict[str, Any]):
        for cell, info in self.widgets.items():
            if cell not in values:
                continue
            w = info['widget']
            val = values[cell]

            if info['type'] == 'readonly' and isinstance(w, QLabel):
                if val:
                    try:
                        w.setText(
                            f"{int(float(val)):,}".replace(',', ' ') + " m²"
                        )
                    except Exception:
                        w.setText("0 m²")
            elif isinstance(w, QComboBox):
                idx = w.findText(str(val) if val else "")
                if idx >= 0:
                    w.setCurrentIndex(idx)
            elif isinstance(w, QLineEdit) and val:
                try:
                    w.setText(
                        f"{int(float(val)):,}".replace(',', ' ') + " m²"
                    )
                except Exception:
                    w.setText(str(val))

# ---------------------------------------------------------------------
# FEED - Planning
# ---------------------------------------------------------------------
class FeedPlanningTab(QWidget):
    def __init__(self, tab_config: Dict, excel_client=None, parent=None):
        super().__init__(parent)
        self.tab_config = tab_config
        self.excel_client = excel_client
        self.widgets = {}
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        title = QLabel(f"{self.tab_config.get('icon', '📅')} {self.tab_config.get('label', 'FEED - Planning')}")
        title.setStyleSheet("font-size:20px;font-weight:bold;padding:10px")
        layout.addWidget(title)

        table_widget = QWidget()
        table_layout = QGridLayout()

        columns = self.tab_config.get('columns', [])
        for i, col in enumerate(columns):
            label = QLabel(col['label'])
            label.setStyleSheet("font-weight:bold;font-size:14px;text-align:center;padding:5px")
            label.setAlignment(Qt.AlignCenter)
            table_layout.addWidget(label, 0, i)

            subtitle = QLabel(col['subtitle'])
            subtitle.setStyleSheet("font-size:11px;color:#666;text-align:center")
            subtitle.setAlignment(Qt.AlignCenter)
            subtitle.setWordWrap(True)
            table_layout.addWidget(subtitle, 1, i)

            if col.get('readonly'):
                date_label = QLabel("--/--/----")
                date_label.setStyleSheet("background-color:#f0f0f0;padding:8px;border-radius:3px")
                date_label.setAlignment(Qt.AlignCenter)
                table_layout.addWidget(date_label, 2, i)
                self.widgets[col['cell']] = {
                    'widget': date_label,
                    'cell': col['cell'],
                    'type': 'readonly',
                    'source': col.get('source')
                }
            else:
                date_edit = QDateEdit()
                date_edit.setCalendarPopup(True)
                date_edit.setDate(QDate.currentDate())
                date_edit.setDisplayFormat('dd/MM/yyyy')
                date_edit.setMinimumWidth(120)
                table_layout.addWidget(date_edit, 2, i)
                self.widgets[col['cell']] = {
                    'widget': date_edit,
                    'cell': col['cell'],
                    'type': 'date'
                }

        table_widget.setLayout(table_layout)
        group = QGroupBox("Planning des phases")
        group_layout = QVBoxLayout()
        group_layout.addWidget(table_widget)
        group.setLayout(group_layout)
        layout.addWidget(group)
        layout.addStretch()
        self.setLayout(layout)

    def update_readonly_dates(self, source_values: Dict[str, Any]):
        for cell, info in self.widgets.items():
            if info['type'] == 'readonly' and info.get('source'):
                source_cell = info['source']
                if source_cell in source_values:
                    val = source_values[source_cell]
                    if val:
                        try:
                            if isinstance(val, (date, datetime)):
                                info['widget'].setText(val.strftime('%d/%m/%Y'))
                            else:
                                info['widget'].setText(str(val))
                        except Exception:
                            pass

    def get_values(self) -> Dict[str, Any]:
        vals = {}
        for cell, info in self.widgets.items():
            if info['type'] == 'date':
                w = info['widget']
                try:
                    vals[cell] = w.date().toPython()
                except Exception:
                    d = w.date()
                    vals[cell] = (d.year(), d.month(), d.day())
        return vals

    def set_values(self, values: Dict[str, Any]):
        for cell, info in self.widgets.items():
            if cell not in values:
                continue
            w = info['widget']
            val = values[cell]

            if info['type'] == 'readonly' and isinstance(w, QLabel):
                if val:
                    try:
                        if isinstance(val, (date, datetime)):
                            w.setText(val.strftime('%d/%m/%Y'))
                        else:
                            w.setText(str(val))
                    except Exception:
                        w.setText("--/--/----")
            elif info['type'] == 'date' and isinstance(w, QDateEdit):
                if isinstance(val, datetime):
                    w.setDate(QDate(val.year, val.month, val.day))
                elif isinstance(val, date):
                    w.setDate(QDate(val.year, val.month, val.day))

# ---------------------------------------------------------------------
# Container des sous-onglets d'inputs
# ---------------------------------------------------------------------
class InputsTabWidget(QTabWidget):
    def __init__(self, schema_path: str, excel_client=None, parent=None):
        super().__init__(parent)
        self.schema_path = schema_path
        self.excel_client = excel_client
        self.tabs = {}
        self.init_ui()

    def init_ui(self):
        with open(self.schema_path, 'r', encoding='utf-8') as f:
            schema = json.load(f)

        for tab_id, tab_config in schema.get('tabs', {}).items():
            if tab_config.get('table_format'):
                if 'surfaces' in tab_id:
                    tab_widget = FeedSurfacesTab(tab_config, self.excel_client)
                elif 'planning' in tab_id:
                    tab_widget = FeedPlanningTab(tab_config, self.excel_client)
                else:
                    continue
            else:
                tab_widget = BaseInputTab(tab_config)

            self.tabs[tab_id] = tab_widget
            self.addTab(tab_widget, tab_config.get('label', tab_id))

    def get_all_values(self) -> Dict[str, Any]:
        all_values = {}
        for tab_widget in self.tabs.values():
            all_values.update(tab_widget.get_values())
        return all_values

    def set_all_values(self, values: Dict[str, Any]):
        for tab_widget in self.tabs.values():
            tab_widget.set_values(values)

        if 'feed_planning' in self.tabs:
            self.tabs['feed_planning'].update_readonly_dates(values)

# ---------------------------------------------------------------------
# Onglet Livrables
# ---------------------------------------------------------------------
class DeliverablesTab(QWidget):
    """Onglet Livrables : inputs Master Guide + bouton macro"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop)

        title = QLabel("🚀 Génération des Livrables")
        title.setStyleSheet("font-size:28px;font-weight:bold;margin-bottom:20px")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        form_widget = QWidget()
        form_widget.setMaximumWidth(700)
        form_layout = QFormLayout(form_widget)
        form_layout.setSpacing(15)
        form_layout.setLabelAlignment(Qt.AlignRight)
        form_widget.setStyleSheet("QLabel{color:#333;font-size:14px;}")

        # Livrable (Master Guide!N7)
        self.livrable_combo = QComboBox()
        self.livrable_combo.addItems([
            "PS 8002 - TSF General (Eng) / PS 8002 - ITC Général (Fr)",
            "PS 8002 - TSF Preparatory Works for Civil Works (Eng) / PS 8002 - ITC Travaux préparatoires GC (Fr)",
            "PS 8002 - TSF Modular Buildings (Eng) / PS 8002 - ITC Bâtiment Modulaire (Fr)",
            "PS 8002 - TSF Preparatory E&I Works (Eng) / PS 8002 - ITC Travaux préparatoires E&I (Fr)",
            "PP 8002 - TSF General (Eng) / PP 8002 - ITC Général (Fr)",
            "SOW 8002 - TSF General (Eng) / SOW 8002 - ITC Général (Fr)",
            "SOW 8002 - TSF Preparatory Works for Civil Works (Eng) / SOW 8002 - ITC Travaux préparatoires GC (Fr)",
            "SOW 8002 - TSF Modular Buildings (Eng) / SOW 8002 - ITC Bâtiment Modulaire (Fr)",
            "SOW 8002 - TSF Preparatory E&I Works (Eng) / SOW 8002 - ITC Travaux préparatoires E&I (Fr)"
        ])
        self.livrable_combo.currentTextChanged.connect(self.on_livrable_changed)
        form_layout.addRow("Livrable :", self.livrable_combo)

        # Langue (Master Guide!N8)
        self.langue_combo = QComboBox()
        self.langue_combo.addItems(["FR", "ENG"])
        form_layout.addRow("Langue:", self.langue_combo)

        # Annexes (Master Guide!Q7-Q10)
        self.annexe1_combo = QComboBox()
        self.annexe1_combo.addItems(["FAUX", "VRAI"])
        form_layout.addRow("Annexe 1?", self.annexe1_combo)

        self.annexe2_combo = QComboBox()
        self.annexe2_combo.addItems(["FAUX", "VRAI"])
        form_layout.addRow("Annexe 2?", self.annexe2_combo)

        self.annexe3_combo = QComboBox()
        self.annexe3_combo.addItems(["FAUX", "VRAI"])
        form_layout.addRow("Annexe 3?", self.annexe3_combo)

        self.annexe4_combo = QComboBox()
        self.annexe4_combo.addItems(["FAUX", "VRAI"])
        form_layout.addRow("Annexe 4?", self.annexe4_combo)

        layout.addWidget(form_widget, alignment=Qt.AlignCenter)

        self.macro_btn = AnimatedButton("🔥 Générer les livrables SOW/PP/PS")
        self.macro_btn.setMinimumSize(350, 60)
        layout.addWidget(self.macro_btn, alignment=Qt.AlignCenter)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("margin-top:30px;font-size:14px")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar{
                border:2px solid #4A90E2;
                border-radius:5px;
                text-align:center;
                height:25px
            }
            QProgressBar::chunk{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 #4A90E2, stop:1 #357ABD);
                border-radius:3px
            }
        """)
        layout.addWidget(self.progress)

        layout.addStretch()
        self.setLayout(layout)

        # Init en fonction du livrable par défaut
        self.on_livrable_changed(self.livrable_combo.currentText())

    def on_livrable_changed(self, text: str):
        """
        Si livrable = PS 8002 ... -> Annexes bloquées & grisées.
        Sinon (PP ou SOW) -> Annexes actives & style normal.
        """
        is_ps = text.startswith("PS ")

        disabled_style = """
            QComboBox {
                background-color: #3a3a3a;
                color: #666;
            }
        """
        enabled_style = """
            QComboBox {
                background-color: #3f3f3f;
                color: white;
            }
        """

        combos = [
            self.annexe1_combo,
            self.annexe2_combo,
            self.annexe3_combo,
            self.annexe4_combo
        ]

        if is_ps:
            for c in combos:
                c.setEnabled(False)
                c.setStyleSheet(disabled_style)
                c.setCurrentText("Non")
        else:
            for c in combos:
                c.setEnabled(True)
                c.setStyleSheet(enabled_style)

    def get_deliverable_inputs(self) -> Dict[str, Any]:
        """
        Valeurs à écrire dans la feuille 'Master Guide':
          - N7 : Livrable
          - N8 : Langue
          - Q7-Q10 : Annexes
        """
        return {
            "N7": self.livrable_combo.currentText(),
            "N8": self.langue_combo.currentText(),
            "Q7": self.annexe1_combo.currentText(),
            "Q8": self.annexe2_combo.currentText(),
            "Q9": self.annexe3_combo.currentText(),
            "Q10": self.annexe4_combo.currentText(),
        }

# ---------------------------------------------------------------------
# Onglet Rapports
# ---------------------------------------------------------------------
class ReportsTab(QWidget):
    """Onglet Rapports : bouton pour Rapports_Numer_Graph_Export"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop)

        title = QLabel("📊 Génération des Rapports numériques")
        title.setStyleSheet("font-size:24px;font-weight:bold;margin-bottom:10px")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        desc = QLabel(
            "Ce module lance la macro Excel de génération des rapports numériques et graphiques.\n"
            "Assure-toi que les données sont à jour dans l’outil avant de lancer."
        )
        desc.setStyleSheet("font-size:13px;color:#666;margin-bottom:25px")
        desc.setAlignment(Qt.AlignCenter)
        desc.setWordWrap(True)
        layout.addWidget(desc)

        self.macro_btn = AnimatedButton("📈 Générer les rapports")
        self.macro_btn.setMinimumSize(320, 55)
        layout.addWidget(self.macro_btn, alignment=Qt.AlignCenter)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("margin-top:25px;font-size:14px")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar{
                border:2px solid #4A90E2;
                border-radius:5px;
                text-align:center;
                height:25px
            }
            QProgressBar::chunk{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 #4A90E2, stop:1 #357ABD);
                border-radius:3px
            }
        """)
        layout.addWidget(self.progress)

        layout.addStretch()
        self.setLayout(layout)

# ---------------------------------------------------------------------
# Fenêtre principale
# ---------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_client = None
        self.config = self._load_config()
        self._init_ui()

    def _load_config(self) -> Dict:
        p = Path("config.json")
        return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}

    def _init_ui(self):
        self.setWindowTitle("ITC Manager V2 - Gestionnaire Excel Complet")
        w = self.config.get('window', {})
        self.resize(w.get('width', 1100), w.get('height', 750))

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # Header
        header = QWidget()
        header.setStyleSheet(
            "background:qlineargradient(x1:0,y1:0,x2:1,y2:0, "
            "stop:0 #1e3c72, stop:1 #2a5298);padding:10px"
        )
        h = QHBoxLayout(header)
        title = QLabel("🏗️ ITC Manager V2")
        title.setStyleSheet("color:white;font-size:20px;font-weight:bold")
        h.addWidget(title)

        self.header_status = QLabel("✅ Prêt")
        self.header_status.setStyleSheet("color:#4CAF50;font-size:14px")
        h.addStretch()
        h.addWidget(self.header_status)
        root.addWidget(header)

        # Tabs principaux
        self.main_tabs = QTabWidget()
        self.main_tabs.setStyleSheet("""
            QTabWidget::pane{border:1px solid #ddd;background:white}
            QTabBar::tab{padding:10px 20px;margin-right:5px}
            QTabBar::tab:selected{background:white;border-bottom:3px solid #4A90E2}
        """)

        # Inputs
        self.inputs_widget = InputsTabWidget("schema_v2.json", self.excel_client)
        self.main_tabs.addTab(self.inputs_widget, "📋 Inputs")

        # Rapports
        self.reports_tab = ReportsTab()
        self.main_tabs.addTab(self.reports_tab, "📊 Rapports")

        # Livrables
        self.deliverables_tab = DeliverablesTab()
        self.main_tabs.addTab(self.deliverables_tab, "🚀 Livrables")

        root.addWidget(self.main_tabs)

        # Status bar
        self.statusBar().showMessage("Application prête")

        # Menus
        self._create_menus()

        # Connexions
        self.deliverables_tab.macro_btn.clicked.connect(self.run_deliverables_macro)
        self.reports_tab.macro_btn.clicked.connect(self.run_reports_macro)

        # Barre de boutons globaux
        button_bar = QWidget()
        button_layout = QHBoxLayout(button_bar)

        self.load_btn = QPushButton("🔄 Charger tout depuis Excel")
        self.load_btn.clicked.connect(self.load_all_from_excel)
        button_layout.addWidget(self.load_btn)

        self.save_btn = QPushButton("💾 Sauvegarder tout dans Excel")
        self.save_btn.setStyleSheet("""
            QPushButton{
                background:#4CAF50;
                color:white;
                padding:8px;
                font-weight:bold;
                border-radius:5px
            }
            QPushButton:hover{background:#45a049}
        """)
        self.save_btn.clicked.connect(self.save_all_to_excel)
        button_layout.addWidget(self.save_btn)

        button_layout.addStretch()
        root.addWidget(button_bar)

    def _create_menus(self):
        m = self.menuBar()
        f = m.addMenu("Fichier")

        a = QAction("Connecter Excel", self)
        a.triggered.connect(self.connect_excel)
        f.addAction(a)
        f.addSeparator()

        q = QAction("Quitter", self)
        q.triggered.connect(self.close)
        f.addAction(q)

        t = m.addMenu("Outils")
        vis = QAction("Rendre Excel visible", self)
        vis.triggered.connect(self.toggle_excel_visible)
        t.addAction(vis)

        hmenu = m.addMenu("Aide")
        about = QAction("À propos", self)
        about.triggered.connect(self.show_about)
        hmenu.addAction(about)

    # -----------------------------------------------------------------
    # Connexion Excel
    # -----------------------------------------------------------------
    def connect_excel(self):
        try:
            from excel_client import ExcelClient
            cfg_path = Path("config.json")
            excel_path = self.config.get('excel_path', "")
            p = Path(excel_path)

            if not p.is_absolute():
                p = Path.cwd() / p

            if not excel_path or not p.exists():
                candidates = sorted(
                    Path.cwd().glob("*.xlsm"),
                    key=lambda pp: pp.stat().st_mtime,
                    reverse=True
                )
                if not candidates:
                    QMessageBox.critical(self, "Erreur", "Aucun fichier .xlsm trouvé dans ce dossier.")
                    return
                p = candidates[0]
                self.config['excel_path'] = p.name
                cfg_path.write_text(
                    json.dumps(self.config, indent=2, ensure_ascii=False),
                    encoding="utf-8"
                )

            self.excel_client = ExcelClient(
                str(p),
                self.config.get('sheet_name', '01.3-ITC MASTER WBS'),
                visible=self.config.get('excel', {}).get('visible_on_first_run', False)
            )
            self.excel_client.connect()

            self.header_status.setText("✅ Connecté")
            self.header_status.setStyleSheet("color:#4CAF50;font-size:14px")
            self.statusBar().showMessage(f"Connexion Excel établie ({p.name})")

            self.inputs_widget.excel_client = self.excel_client
            self.load_all_from_excel()

        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Connexion impossible : {e}")
            self.header_status.setText("❌ Déconnecté")
            self.header_status.setStyleSheet("color:#F44336;font-size:14px")

    # -----------------------------------------------------------------
    # Chargement / sauvegarde des inputs BASIC/FEED
    # -----------------------------------------------------------------
    def load_all_from_excel(self):
        if not self.excel_client:
            self.connect_excel()
        if not self.excel_client:
            return

        try:
            all_cells = set()
            for tab_widget in self.inputs_widget.tabs.values():
                for info in tab_widget.widgets.values():
                    cell = info.get('cell')
                    if cell:
                        all_cells.add(cell)
                    source = info.get('source')
                    if source:
                        all_cells.add(source)

            values = {}
            for cell in all_cells:
                try:
                    values[cell] = self.excel_client.read_cell(cell)
                except Exception:
                    values[cell] = None

            self.inputs_widget.set_all_values(values)
            self.statusBar().showMessage("Toutes les valeurs chargées depuis Excel")

        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Chargement impossible : {e}")

    def save_all_to_excel(self):
        if not self.excel_client:
            self.connect_excel()
        if not self.excel_client:
            return

        try:
            values = self.inputs_widget.get_all_values()

            if hasattr(self.excel_client, 'write_cells_batch'):
                self.excel_client.write_cells_batch(values)
            else:
                for cell, value in values.items():
                    self.excel_client.write_cell(cell, value)

            if self.config.get('excel', {}).get('auto_save', True):
                self.excel_client.save()

            self.statusBar().showMessage("Toutes les valeurs écrites dans Excel")
            QMessageBox.information(self, "Succès", f"{len(values)} valeurs écrites dans Excel")

        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Écriture impossible : {e}")

    # -----------------------------------------------------------------
    # Macros Livrables & Rapports
    # -----------------------------------------------------------------
    def run_deliverables_macro(self):
        if not self.excel_client:
            self.connect_excel()
        if not self.excel_client:
            return

        try:
            # 1) Écrire les inputs dans 'Master Guide'
            deliverable_inputs = self.deliverables_tab.get_deliverable_inputs()
            original_sheet = self.excel_client.sheet_name
            try:
                self.excel_client.sheet_name = "Master Guide"
                if hasattr(self.excel_client, "write_cells_batch"):
                    self.excel_client.write_cells_batch(deliverable_inputs)
                else:
                    for cell, value in deliverable_inputs.items():
                        self.excel_client.write_cell(cell, value)
                if self.config.get('excel', {}).get('auto_save', True):
                    self.excel_client.save()
            finally:
                self.excel_client.sheet_name = original_sheet

            # 2) Lancer la macro de génération des livrables
            self.deliverables_tab.status_label.setText("⏳ Exécution de la macro Livrables...")
            self.deliverables_tab.progress.setVisible(True)
            self.deliverables_tab.progress.setRange(0, 0)
            QApplication.processEvents()

            macro = self.config.get('macro_deliverables', 'Module1.Generer_Livrables')
            self.excel_client.run_macro(macro)

            self.deliverables_tab.progress.setVisible(False)
            self.deliverables_tab.status_label.setText("✅ Livrables générés avec succès !")
            self.statusBar().showMessage("Livrables générés")
            self._toast_success("🎉 Livrables générés avec succès !")

        except Exception as e:
            self.deliverables_tab.progress.setVisible(False)
            self.deliverables_tab.status_label.setText("❌ Erreur lors de l'exécution")
            QMessageBox.critical(self, "Erreur", f"Macro livrables impossible : {e}")

    def run_reports_macro(self):
        if not self.excel_client:
            self.connect_excel()
        if not self.excel_client:
            return

        try:
            self.reports_tab.status_label.setText("⏳ Exécution de la macro Rapports...")
            self.reports_tab.progress.setVisible(True)
            self.reports_tab.progress.setRange(0, 0)
            QApplication.processEvents()

            macro = self.config.get('macro_reports', 'Module1.Rapports_Numer_Graph_Export')
            self.excel_client.run_macro(macro)

            self.reports_tab.progress.setVisible(False)
            self.reports_tab.status_label.setText("✅ Rapports générés avec succès !")
            self.statusBar().showMessage("Rapports générés")
            self._toast_success("📈 Rapports générés avec succès !")

        except Exception as e:
            self.reports_tab.progress.setVisible(False)
            self.reports_tab.status_label.setText("❌ Erreur lors de l'exécution")
            QMessageBox.critical(self, "Erreur", f"Macro rapports impossible : {e}")

    def _toast_success(self, message: str):
        toast = QLabel(message, self)
        toast.setStyleSheet("""
            QLabel{
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 #4CAF50, stop:1 #8BC34A);
                color:white;
                padding:15px 30px;
                border-radius:20px;
                font-size:16px;
                font-weight:bold
            }
        """)
        toast.setAlignment(Qt.AlignCenter)
        toast.adjustSize()
        toast.move(
            (self.width() - toast.width()) // 2,
            (self.height() - toast.height()) // 2
        )
        toast.show()

        self.fade = QPropertyAnimation(toast, b"windowOpacity")
        self.fade.setDuration(3000)
        self.fade.setStartValue(1.0)
        self.fade.setEndValue(0.0)
        self.fade.finished.connect(toast.deleteLater)
        QTimer.singleShot(1000, self.fade.start)

    # -----------------------------------------------------------------
    # Divers
    # -----------------------------------------------------------------
    def toggle_excel_visible(self):
        if self.excel_client and self.excel_client.app:
            cur = self.excel_client.app.visible
            self.excel_client.set_visible(not cur)
            self.statusBar().showMessage(f"Excel {'visible' if not cur else 'masqué'}")

    def show_about(self):
        QMessageBox.about(
            self,
            "À propos",
            "<h2>ITC Manager V2</h2>"
            "<p>Version complète BASIC + FEED + Rapports + Livrables</p>"
            "<p>PySide6 + xlwings</p>"
        )

    def closeEvent(self, e):
        if self.excel_client:
            reply = QMessageBox.question(
                self, "Confirmation",
                "Voulez-vous sauvegarder avant de quitter ?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )

            if reply == QMessageBox.Save:
                try:
                    self.excel_client.save()
                except Exception:
                    pass
                self.excel_client.close()
            elif reply == QMessageBox.Cancel:
                e.ignore()
                return
            else:
                self.excel_client.close()

        e.accept()
'''
(Path("ui/main_window_v2.py")).write_text(main_window_code, encoding="utf-8")
print("✅ ui/main_window_v2.py")

# ---------------------------------------------------------------------
# 6) test_v2.py
# ---------------------------------------------------------------------
test_code = r'''"""
Test de la V2
"""
from pathlib import Path

def test_imports():
    try:
        import PySide6
        print("✅ PySide6 OK")
        import xlwings
        print("✅ xlwings OK")
        import win32com.client
        print("✅ pywin32 OK")
        from excel_client import ExcelClient
        print("✅ excel_client.py OK")
        from ui.main_window_v2 import MainWindow
        print("✅ ui/main_window_v2.py OK")
        return True
    except Exception as e:
        print(f"❌ Import error: {e}")
        return False

def test_config():
    ok = True

    if not Path("config.json").exists():
        print("❌ Missing config.json")
        ok = False
    else:
        print("✅ config.json OK")

    if not Path("schema_v2.json").exists():
        print("❌ Missing schema_v2.json")
        ok = False
    else:
        print("✅ schema_v2.json OK")

        import json
        with open("schema_v2.json", "r", encoding="utf-8") as f:
            schema = json.load(f)
            tabs_count = len(schema.get('tabs', {}))
            print(f"  → {tabs_count} onglets configurés")

            for tab_id, tab in schema.get('tabs', {}).items():
                if 'fields' in tab:
                    print(f"  → {tab_id}: {len(tab['fields'])} champs")
                elif 'sections' in tab:
                    sections = tab['sections']
                    total_rows = sum(len(s.get('rows', [])) for s in sections.values())
                    print(f"  → {tab_id}: {total_rows} lignes totales")
                elif 'columns' in tab:
                    print(f"  → {tab_id}: {len(tab['columns'])} colonnes")

    return ok

if __name__ == "__main__":
    print("=" * 50)
    print("🔍 ITC Manager V2 Test")
    print("=" * 50)

    if test_imports() and test_config():
        print("\n✅ Application V2 prête!")
        print("\n▶️ Démarrage: python main.py")
    else:
        print("\n⚠️ Corrigez les erreurs ci-dessus")
'''
(Path("test_v2.py")).write_text(test_code, encoding="utf-8")
print("✅ test_v2.py")

# ---------------------------------------------------------------------
# Résumé
# ---------------------------------------------------------------------
print("\n" + "="*60)
print("🎉 BOOTSTRAP V2 TERMINÉ!")
print("="*60)
if auto_excel:
    print(f"📊 Excel détecté: {auto_excel}")
else:
    print("⚠️ Aucun .xlsm détecté - placez votre fichier dans ce dossier")

print("\n📋 STRUCTURE CRÉÉE:")
print("  • 4 onglets d'inputs:")
print("    - BASIC (7 champs)")
print("    - FEED Quantités (26 champs J26-J51)")
print("    - FEED Surfaces (tableaux SIP/SOP + récap WBS)")
print("    - FEED Planning (5 colonnes de dates)")
print("  • 1 onglet Rapports (macro Module1.Rapports_Numer_Graph_Export)")
print("  • 1 onglet Livrables (Master Guide N7/N8/Q7-Q10 + macro Module1.Generer_Livrables)")

print("\n🚀 DÉMARRAGE:")
print("  1. Test: python test_v2.py")
print("  2. Lancer: python main.py")

print("\n" + "="*60)
