import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sqlite3
import threading
import os
import sys
import shutil
import base64
import mimetypes
from datetime import datetime, date, timedelta
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import ctypes
import pathlib, sys

# ─────────────────────────────────────────────────────────────────────────────
# Force Windows to treat this script as a distinct application on the taskbar
# ─────────────────────────────────────────────────────────────────────────────
try:
    myappid = 'mycompany.myproduct.subproduct.version' # Custom arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except AttributeError:
    pass # Fallback for non-Windows operating systems

# ─────────────────────────────────────────────────────────────────────────────
# Resolve pathname for .ico referencing in same directory
# ─────────────────────────────────────────────────────────────────────────────
def _resource(filename: str) -> str:
    """Resolve a resource path for both dev and PyInstaller frozen builds."""
    base = pathlib.Path(getattr(sys, '_MEIPASS', pathlib.Path(__file__).parent))
    return str(base / filename)

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
# Constants & Paths
# ─────────────────────────────────────────────────────────────────────────────
APP_DIR        = Path(os.path.dirname(os.path.abspath(__file__)))
DB_PATH        = APP_DIR / "ems_data.db"
CFG_PATH       = APP_DIR / "ems_config.xlsx"
EXPORT_DIR     = APP_DIR / "exports"
ATTACHMENTS_DIR= APP_DIR / "attachments"

URGENCY_CRITICAL = "CRITICAL"
URGENCY_DUE      = "DUE TODAY"
URGENCY_SOON     = "DUE SOON"
URGENCY_OK       = "OK"

# Colour palette — industrial / precise
CLR_BG       = "#1A1F2E"
CLR_SURFACE  = "#242938"
CLR_SURFACE2 = "#2E3347"
CLR_ACCENT   = "#00C2A8"
CLR_ACCENT2  = "#0076CC"
CLR_WARN     = "#F5A623"
CLR_DANGER   = "#E84040"
CLR_OK       = "#27AE60"
CLR_TEXT     = "#E8ECF0"
CLR_MUTED    = "#8892A0"
CLR_BORDER   = "#3A4055"
CLR_ROW_ODD  = "#242938"
CLR_ROW_EVEN = "#2A2F44"
CLR_CAL_TODAY= "#00C2A8"
CLR_CAL_CRIT = "#E84040"
CLR_CAL_SOON = "#F5A623"
CLR_CAL_OK   = "#27AE60"
CLR_CAL_NONE = "#2E3347"

FONT_BODY  = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)
FONT_LABEL = ("Segoe UI", 9, "bold")
FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_HEAD  = ("Segoe UI", 12, "bold")
FONT_MONO  = ("Consolas", 9)

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}
DOC_EXTS   = {".pdf", ".docx", ".xlsx", ".txt", ".csv", ".doc", ".xls"}


# ─────────────────────────────────────────────────────────────────────────────
# Database Layer
# ─────────────────────────────────────────────────────────────────────────────
class Database:
    _local = threading.local()

    def __init__(self, path: Path):
        self.path = str(path)
        self._init_schema()

    def _conn(self):
        if not getattr(self._local, "conn", None):
            c = sqlite3.connect(self.path, check_same_thread=False)
            c.execute("PRAGMA journal_mode=WAL")
            c.execute("PRAGMA foreign_keys=ON")
            c.row_factory = sqlite3.Row
            self._local.conn = c
        return self._local.conn

    @property
    def conn(self):
        return self._conn()

    def _init_schema(self):
        c = self.conn
        c.executescript("""
        CREATE TABLE IF NOT EXISTS equipment (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            code        TEXT NOT NULL UNIQUE,
            name        TEXT NOT NULL,
            category    TEXT,
            location    TEXT,
            model       TEXT,
            serial_no   TEXT,
            status      TEXT DEFAULT 'Active',
            notes       TEXT,
            created_at  TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS maintenance_tasks (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            equipment_id    INTEGER NOT NULL REFERENCES equipment(id) ON DELETE CASCADE,
            task_name       TEXT NOT NULL,
            frequency_days  INTEGER NOT NULL,
            last_done       TEXT,
            next_due        TEXT,
            assigned_to     TEXT,
            procedure_notes TEXT,
            is_active       INTEGER DEFAULT 1,
            created_at      TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS maintenance_log (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id         INTEGER NOT NULL REFERENCES maintenance_tasks(id) ON DELETE CASCADE,
            equipment_id    INTEGER NOT NULL,
            done_date       TEXT NOT NULL,
            done_by         TEXT,
            duration_min    INTEGER,
            findings        TEXT,
            next_due        TEXT,
            logged_at       TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS maintenance_log_parts (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            log_id          INTEGER NOT NULL REFERENCES maintenance_log(id) ON DELETE CASCADE,
            part_id         INTEGER NOT NULL REFERENCES spare_parts(id),
            qty_used        REAL NOT NULL
        );

        CREATE TABLE IF NOT EXISTS spare_parts (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            part_number     TEXT NOT NULL UNIQUE,
            description     TEXT NOT NULL,
            category        TEXT,
            location        TEXT,
            unit            TEXT DEFAULT 'pcs',
            qty_on_hand     REAL DEFAULT 0,
            safety_stock    REAL DEFAULT 0,
            reorder_qty     REAL DEFAULT 0,
            unit_cost       REAL DEFAULT 0,
            supplier        TEXT,
            lead_time_days  INTEGER DEFAULT 0,
            compatible_equip TEXT,
            notes           TEXT,
            is_active       INTEGER DEFAULT 1,
            created_at      TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS parts_transactions (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            part_id         INTEGER NOT NULL REFERENCES spare_parts(id) ON DELETE CASCADE,
            txn_type        TEXT NOT NULL,
            qty             REAL NOT NULL,
            qty_before      REAL,
            qty_after       REAL,
            ref_equipment   TEXT,
            ref_task        TEXT,
            performed_by    TEXT,
            remarks         TEXT,
            txn_date        TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS attachments (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            entity_type     TEXT NOT NULL,
            entity_id       INTEGER NOT NULL,
            filename        TEXT NOT NULL,
            stored_name     TEXT NOT NULL,
            file_type       TEXT,
            file_size       INTEGER,
            description     TEXT,
            uploaded_at     TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS app_settings (
            key   TEXT PRIMARY KEY,
            value TEXT
        );
        """)
        c.commit()

    def fetchall(self, sql, params=()):
        return self.conn.execute(sql, params).fetchall()

    def fetchone(self, sql, params=()):
        return self.conn.execute(sql, params).fetchone()

    def execute(self, sql, params=()):
        cur = self.conn.execute(sql, params)
        self.conn.commit()
        return cur

    def lastrowid(self, sql, params=()):
        cur = self.conn.execute(sql, params)
        self.conn.commit()
        return cur.lastrowid

    def setting(self, key, default=None):
        row = self.fetchone("SELECT value FROM app_settings WHERE key=?", (key,))
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.execute(
            "INSERT INTO app_settings(key,value) VALUES(?,?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, str(value))
        )


# ─────────────────────────────────────────────────────────────────────────────
# Config Loader
# ─────────────────────────────────────────────────────────────────────────────
class ConfigLoader:
    def __init__(self, db: Database, path: Path):
        self.db   = db
        self.path = path

    def load(self):
        if not self.path.exists():
            self._create_template()
            return f"Template config created at {self.path}"
        wb   = openpyxl.load_workbook(self.path, data_only=True)
        msgs = []
        msgs += self._load_settings(wb)
        msgs += self._load_equipment(wb)
        msgs += self._load_tasks(wb)
        msgs += self._load_parts(wb)
        return "\n".join(msgs) if msgs else "Config loaded successfully."

    def _load_settings(self, wb):
        if "Settings" not in wb.sheetnames:
            return []
        for row in wb["Settings"].iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:
                self.db.set_setting(str(row[0]).strip(), str(row[1]).strip())
        return ["Settings loaded."]

    def _load_equipment(self, wb):
        if "Equipment" not in wb.sheetnames:
            return []
        count = 0
        for row in wb["Equipment"].iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            code = str(row[0]).strip()
            vals = {
                "code": code,
                "name": str(row[1]).strip() if row[1] else code,
                "category": str(row[2]).strip() if row[2] else "",
                "location": str(row[3]).strip() if row[3] else "",
                "model": str(row[4]).strip() if row[4] else "",
                "serial_no": str(row[5]).strip() if row[5] else "",
                "notes": str(row[6]).strip() if row[6] else "",
            }
            self.db.execute(
                """INSERT INTO equipment(code,name,category,location,model,serial_no,notes)
                   VALUES(:code,:name,:category,:location,:model,:serial_no,:notes)
                   ON CONFLICT(code) DO UPDATE SET
                   name=excluded.name,category=excluded.category,
                   location=excluded.location,model=excluded.model,
                   serial_no=excluded.serial_no,notes=excluded.notes""", vals)
            count += 1
        return [f"Equipment: {count} rows loaded."]

    def _load_tasks(self, wb):
        if "MaintenanceTasks" not in wb.sheetnames:
            return []
        count = 0
        for row in wb["MaintenanceTasks"].iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            eq_code = str(row[0]).strip()
            eq_row  = self.db.fetchone("SELECT id FROM equipment WHERE code=?", (eq_code,))
            if not eq_row: continue
            eq_id     = eq_row["id"]
            task_name = str(row[1]).strip() if row[1] else ""
            freq      = int(row[2]) if row[2] else 30
            last_done = None
            if row[3]:
                try:
                    last_done = row[3].strftime("%Y-%m-%d") if isinstance(row[3], (date, datetime)) else str(row[3]).strip()[:10]
                except Exception:
                    last_done = None
            next_due = None
            if last_done:
                try:
                    next_due = (datetime.strptime(last_done, "%Y-%m-%d") + timedelta(days=freq)).strftime("%Y-%m-%d")
                except Exception:
                    pass
            assigned = str(row[4]).strip() if row[4] else ""
            proc     = str(row[5]).strip() if row[5] else ""
            existing = self.db.fetchone(
                "SELECT id FROM maintenance_tasks WHERE equipment_id=? AND task_name=?", (eq_id, task_name))
            if existing:
                self.db.execute(
                    "UPDATE maintenance_tasks SET frequency_days=?,last_done=?,next_due=?,assigned_to=?,procedure_notes=? WHERE id=?",
                    (freq, last_done, next_due, assigned, proc, existing["id"]))
            else:
                self.db.execute(
                    "INSERT INTO maintenance_tasks(equipment_id,task_name,frequency_days,last_done,next_due,assigned_to,procedure_notes) VALUES(?,?,?,?,?,?,?)",
                    (eq_id, task_name, freq, last_done, next_due, assigned, proc))
            count += 1
        return [f"Maintenance tasks: {count} rows loaded."]

    def _load_parts(self, wb):
        if "SpareParts" not in wb.sheetnames:
            return []
        count = 0
        for row in wb["SpareParts"].iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            pn = str(row[0]).strip()
            vals = {
                "part_number": pn,
                "description": str(row[1]).strip() if row[1] else pn,
                "category": str(row[2]).strip() if row[2] else "",
                "location": str(row[3]).strip() if row[3] else "",
                "unit": str(row[4]).strip() if row[4] else "pcs",
                "qty_on_hand": float(row[5]) if row[5] else 0,
                "safety_stock": float(row[6]) if row[6] else 0,
                "reorder_qty": float(row[7]) if row[7] else 0,
                "unit_cost": float(row[8]) if row[8] else 0,
                "supplier": str(row[9]).strip() if row[9] else "",
                "lead_time_days": int(row[10]) if row[10] else 0,
                "compatible_equip": str(row[11]).strip() if row[11] else "",
                "notes": str(row[12]).strip() if row[12] else "",
            }
            self.db.execute(
                """INSERT INTO spare_parts(part_number,description,category,location,unit,
                   qty_on_hand,safety_stock,reorder_qty,unit_cost,supplier,
                   lead_time_days,compatible_equip,notes)
                   VALUES(:part_number,:description,:category,:location,:unit,
                   :qty_on_hand,:safety_stock,:reorder_qty,:unit_cost,:supplier,
                   :lead_time_days,:compatible_equip,:notes)
                   ON CONFLICT(part_number) DO UPDATE SET
                   description=excluded.description,category=excluded.category,
                   location=excluded.location,unit=excluded.unit,
                   safety_stock=excluded.safety_stock,reorder_qty=excluded.reorder_qty,
                   unit_cost=excluded.unit_cost,supplier=excluded.supplier,
                   lead_time_days=excluded.lead_time_days,
                   compatible_equip=excluded.compatible_equip,notes=excluded.notes""", vals)
            count += 1
        return [f"Spare parts: {count} rows loaded."]

    def _create_template(self):
        wb = openpyxl.Workbook()
        header_font = Font(bold=True, color="FFFFFF", size=10)
        header_fill = PatternFill("solid", fgColor="1A2744")
        border_side = Side(style="thin", color="CCCCCC")
        thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

        def make_sheet(wb, name, headers, examples):
            ws = wb.create_sheet(name)
            for ci, hdr in enumerate(headers, 1):
                cell = ws.cell(1, ci, hdr)
                cell.font = header_font; cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center")
                ws.column_dimensions[get_column_letter(ci)].width = 18
            for ri, row_data in enumerate(examples, 2):
                for ci, val in enumerate(row_data, 1):
                    ws.cell(ri, ci, val).border = thin_border
            return ws

        ws_s = make_sheet(wb, "Settings", ["key", "value"], [
            ["facility_name", "Acme Electronics - SMT Line 1"],
            ["alert_soon_days", "7"], ["default_user", "Technician"],
            ["currency", "USD"], ["app_theme", "dark"],
        ])
        ws_s.column_dimensions["A"].width = 22
        ws_s.column_dimensions["B"].width = 35

        make_sheet(wb, "Equipment",
            ["code","name","category","location","model","serial_no","notes"],
            [
                ["SMT-001","Fuji NXT III Mounter","SMT Placer","Line 1","NXT III","SN12345","High-speed placer"],
                ["SMT-002","DEK Horizon 03iX","Screen Printer","Line 1","Horizon 03iX","SN67890","Stencil printer"],
                ["AOI-001","Koh Young KY8030-3","AOI","Line 1","KY8030-3","SN11111","3D AOI system"],
                ["REFLOW-01","Heller 1964 MK7","Reflow Oven","Line 1","1964 MK7","SN22222","7-zone reflow"],
                ["WS-001","ERSA Wave Solder","Wave Solder","Line 2","VERSAFLOW","SN33333",""],
            ])

        make_sheet(wb, "MaintenanceTasks",
            ["equipment_code","task_name","frequency_days","last_done","assigned_to","procedure_notes"],
            [
                ["SMT-001","Nozzle Cleaning",7,"2024-12-01","TechA","Clean all nozzles with IPA"],
                ["SMT-001","Feeder Calibration",30,"2024-11-15","TechA","Calibrate all tape feeders"],
                ["SMT-001","Vision System Check",90,"2024-10-01","TechB","Verify camera calibration"],
                ["SMT-002","Squeegee Replacement",14,"2024-12-05","TechA","Replace squeegee blades"],
                ["SMT-002","Stencil Inspection",7,"2024-12-08","TechA","Check stencil apertures"],
                ["REFLOW-01","Flux Filter Replace",30,"2024-11-20","TechB","Replace exhaust filters"],
                ["REFLOW-01","Zone Temp Calibration",90,"2024-09-15","TechC","Calibrate all zones"],
                ["AOI-001","Camera Lens Cleaning",30,"2024-11-30","TechB","Clean all camera lenses"],
                ["WS-001","Flux Level Check",3,"2024-12-09","TechA","Check and top up flux"],
                ["WS-001","Solder Dross Removal",7,"2024-12-07","TechA","Remove dross from solder pot"],
            ])

        make_sheet(wb, "SpareParts",
            ["part_number","description","category","location","unit",
             "qty_on_hand","safety_stock","reorder_qty","unit_cost",
             "supplier","lead_time_days","compatible_equip","notes"],
            [
                ["NZ-0402-BLK","Nozzle 0402 Black","Nozzle","Rack A1","pcs",20,5,10,45.00,"Fuji Americas",14,"SMT-001",""],
                ["NZ-0603-BLK","Nozzle 0603 Black","Nozzle","Rack A1","pcs",15,5,10,48.00,"Fuji Americas",14,"SMT-001",""],
                ["SQ-12IN-MET","Squeegee Blade 12in","Squeegee","Rack B2","pcs",8,2,4,35.00,"DEK Corp",7,"SMT-002","Metal blade"],
                ["FL-ALPHA","Alpha EF-7000 Flux","Flux","Store","ltr",10,3,5,28.50,"Alpha Metals",5,"WS-001","No-clean flux"],
                ["FLT-EXH-01","Exhaust Filter Set","Filter","Store","set",6,2,4,22.00,"Heller",21,"REFLOW-01","Replace monthly"],
                ["SLD-SAC305","SAC305 Solder Bar","Solder","Store","kg",50,10,20,15.00,"Kester",7,"WS-001","Lead-free"],
                ["BELT-CNV-01","Conveyor Belt 1m","Belts","Rack C3","pcs",3,1,2,120.00,"Generic",30,"SMT-001,SMT-002",""],
                ["PASTE-SAC","SAC305 Solder Paste","Paste","Fridge","jar",12,4,8,65.00,"Alpha Metals",3,"SMT-002","Keep refrigerated"],
            ])

        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        wb.save(self.path)


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def days_until(date_str):
    if not date_str:
        return None
    try:
        return (datetime.strptime(date_str[:10], "%Y-%m-%d").date() - date.today()).days
    except Exception:
        return None

def urgency(days):
    if days is None: return URGENCY_OK
    if days < 0:     return URGENCY_CRITICAL
    if days == 0:    return URGENCY_DUE
    if days <= 7:    return URGENCY_SOON
    return URGENCY_OK

def urgency_color(u):
    return {
        URGENCY_CRITICAL: CLR_DANGER,
        URGENCY_DUE:      CLR_WARN,
        URGENCY_SOON:     "#E8A020",
        URGENCY_OK:       CLR_OK,
    }.get(u, CLR_MUTED)

def open_file_externally(path):
    path = str(path)
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        os.system(f'open "{path}"')
    else:
        os.system(f'xdg-open "{path}"')


# ─────────────────────────────────────────────────────────────────────────────
# Custom Widgets
# ─────────────────────────────────────────────────────────────────────────────
class StyledTreeview(ttk.Treeview):
    def __init__(self, master, columns, headings, col_widths=None,
                 col_anchors=None, show_scrollbar=True, **kwargs):
        frame = tk.Frame(master, bg=CLR_BG)
        self._frame = frame

        style = ttk.Style()
        style_name = f"Custom{id(self)}.Treeview"
        style.configure(style_name,
                        background=CLR_SURFACE, foreground=CLR_TEXT,
                        fieldbackground=CLR_SURFACE, rowheight=24, font=FONT_BODY)
        style.configure(f"{style_name}.Heading",
                        background=CLR_BG, foreground=CLR_ACCENT,
                        font=FONT_LABEL, relief="flat")
        style.map(style_name,
                  background=[("selected", CLR_ACCENT2)],
                  foreground=[("selected", "#FFFFFF")])

        super().__init__(frame, columns=columns, show="headings",
                         style=style_name, **kwargs)

        for col, hdr in zip(columns, headings):
            self.heading(col, text=hdr,
                         command=lambda c=col: self._sort_column(c, False))
            w = col_widths.get(col, 120) if col_widths else 120
            a = col_anchors.get(col, "w") if col_anchors else "w"
            self.column(col, width=w, minwidth=60, anchor=a, stretch=True)

        if show_scrollbar:
            vsb = ttk.Scrollbar(frame, orient="vertical", command=self.yview)
            self.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y")
            hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.xview)
            self.configure(xscrollcommand=hsb.set)
            hsb.pack(side="bottom", fill="x")

        ttk.Treeview.pack(self, side="left", fill="both", expand=True)

        self.tag_configure("odd",       background=CLR_ROW_ODD)
        self.tag_configure("even",      background=CLR_ROW_EVEN)
        self.tag_configure("critical",  background="#3A1A1A", foreground=CLR_DANGER)
        self.tag_configure("warning",   background="#3A2A10", foreground=CLR_WARN)
        self.tag_configure("due",       background="#3A2E10", foreground="#F0C040")
        self.tag_configure("ok",        background=CLR_ROW_ODD, foreground=CLR_OK)
        self.tag_configure("low_stock", background="#3A1A1A", foreground=CLR_DANGER)

    def grid(self, **kwargs):   self._frame.grid(**kwargs)
    def pack(self, **kwargs):   self._frame.pack(**kwargs)
    def place(self, **kwargs):  self._frame.place(**kwargs)

    def recolor_rows(self):
        for i, iid in enumerate(self.get_children()):
            existing_tags = list(self.item(iid, "tags"))
            sem = [t for t in existing_tags if t not in ("odd", "even")]
            parity = "even" if i % 2 == 0 else "odd"
            if not sem:
                self.item(iid, tags=(parity,))

    def _sort_column(self, col, reverse):
        rows = [(self.set(k, col), k) for k in self.get_children("")]
        try:
            rows.sort(key=lambda x: float(x[0].replace(",", "")), reverse=reverse)
        except (ValueError, AttributeError):
            rows.sort(key=lambda x: x[0].lower(), reverse=reverse)
        for idx, (_, k) in enumerate(rows):
            self.move(k, "", idx)
        self.heading(col, command=lambda: self._sort_column(col, not reverse))
        self.recolor_rows()

    def bind_context_menu(self, menu_builder):
        """Attach right-click context menu. menu_builder(event) -> tk.Menu."""
        def _show(event):
            # select row under cursor
            iid = self.identify_row(event.y)
            if iid:
                self.selection_set(iid)
            menu = menu_builder(event)
            if menu:
                try:
                    menu.tk_popup(event.x_root, event.y_root)
                finally:
                    menu.grab_release()
        self.bind("<Button-3>", _show)
        if sys.platform == "darwin":
            self.bind("<Button-2>", _show)

    def bind_double_click(self, callback):
        self.bind("<Double-Button-1>", callback)


class StatusBar(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=CLR_BG, height=24)
        self._lbl = tk.Label(self, text="Ready", bg=CLR_BG, fg=CLR_MUTED,
                             font=FONT_SMALL, anchor="w")
        self._lbl.pack(side="left", padx=10, fill="x", expand=True)
        self._ts  = tk.Label(self, text="", bg=CLR_BG, fg=CLR_MUTED,
                             font=FONT_MONO, anchor="e")
        self._ts.pack(side="right", padx=10)
        self._update_clock()

    def set(self, msg, color=None):
        self._lbl.config(text=msg, fg=color or CLR_MUTED)

    def _update_clock(self):
        self._ts.config(text=datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))
        self._ts.after(1000, self._update_clock)


# ─────────────────────────────────────────────────────────────────────────────
# Dialog helpers
# ─────────────────────────────────────────────────────────────────────────────
class FormDialog(tk.Toplevel):
    def __init__(self, parent, title, fields, initial=None):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.configure(bg=CLR_SURFACE)
        self.resizable(True, True)
        self.grab_set()
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx()+80}+{parent.winfo_rooty()+60}")
        self._vars   = {}
        self._fields = fields
        self._build(fields, initial or {})

    def _build(self, fields, initial):
        canvas = tk.Canvas(self, bg=CLR_SURFACE, highlightthickness=0)
        sb     = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner  = tk.Frame(canvas, bg=CLR_SURFACE, padx=20, pady=14)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def on_resize(e):
            canvas.itemconfig(win_id, width=e.width)
        canvas.bind("<Configure>", on_resize)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        for label, key, ftype, *opts in fields:
            row_f = tk.Frame(inner, bg=CLR_SURFACE)
            row_f.pack(fill="x", pady=3)
            tk.Label(row_f, text=label, bg=CLR_SURFACE, fg=CLR_MUTED,
                     font=FONT_LABEL, width=22, anchor="w").pack(side="left")

            val = initial.get(key, opts[0] if opts and ftype != "combo" else "")

            if ftype in ("text",):
                var = tk.StringVar(value=str(val) if val else "")
                tk.Entry(row_f, textvariable=var, bg=CLR_BG, fg=CLR_TEXT,
                         insertbackground=CLR_TEXT, font=FONT_BODY,
                         relief="flat", bd=4, width=32).pack(side="left", fill="x", expand=True)
                self._vars[key] = var

            elif ftype in ("int", "float"):
                var = tk.StringVar(value=str(val) if val is not None else "0")
                tk.Entry(row_f, textvariable=var, bg=CLR_BG, fg=CLR_TEXT,
                         insertbackground=CLR_TEXT, font=FONT_BODY,
                         relief="flat", bd=4, width=16).pack(side="left")
                self._vars[key] = var

            elif ftype == "date":
                var = tk.StringVar(value=str(val)[:10] if val else "")
                tk.Entry(row_f, textvariable=var, bg=CLR_BG, fg=CLR_TEXT,
                         insertbackground=CLR_TEXT, font=FONT_BODY,
                         relief="flat", bd=4, width=16).pack(side="left")
                tk.Label(row_f, text="YYYY-MM-DD", fg=CLR_MUTED, bg=CLR_SURFACE,
                         font=FONT_SMALL).pack(side="left", padx=4)
                self._vars[key] = var

            elif ftype == "combo":
                choices = opts[0] if opts else []
                var = tk.StringVar(value=str(val) if val else (choices[0] if choices else ""))
                cb_style = ttk.Style()
                cb_style.configure("Dark.TCombobox",
                                   fieldbackground=CLR_BG, background=CLR_SURFACE2,
                                   foreground=CLR_TEXT, selectbackground=CLR_ACCENT2,
                                   selectforeground="#FFFFFF", arrowcolor=CLR_ACCENT)
                cb_style.map("Dark.TCombobox",
                             fieldbackground=[("readonly", CLR_BG)],
                             foreground=[("readonly", CLR_TEXT)],
                             selectbackground=[("readonly", CLR_ACCENT2)])
                cb = ttk.Combobox(row_f, textvariable=var, values=choices,
                                  state="readonly", font=FONT_BODY, width=30,
                                  style="Dark.TCombobox")
                cb.pack(side="left")
                self._vars[key] = var

            elif ftype == "textarea":
                txt = tk.Text(row_f, bg=CLR_BG, fg=CLR_TEXT,
                              insertbackground=CLR_TEXT, font=FONT_BODY,
                              relief="flat", bd=4, width=32, height=3)
                txt.insert("1.0", str(val) if val else "")
                txt.pack(side="left", fill="x", expand=True)
                self._vars[key] = txt

            elif ftype == "bool":
                var = tk.BooleanVar(value=bool(val))
                tk.Checkbutton(row_f, variable=var, bg=CLR_SURFACE, fg=CLR_TEXT,
                               activebackground=CLR_SURFACE, selectcolor=CLR_BG).pack(side="left")
                self._vars[key] = var

            elif ftype == "equip_multi":
                eq_choices = opts[0] if opts else []
                var = tk.StringVar(value=str(val) if val else "")
                entry = tk.Entry(row_f, textvariable=var, bg=CLR_BG, fg=CLR_TEXT,
                                 insertbackground=CLR_TEXT, font=FONT_BODY,
                                 relief="flat", bd=4, width=24, state="readonly")
                entry.pack(side="left", fill="x", expand=True)
                def _open_picker(v=var, choices=eq_choices):
                    dlg = EquipMultiSelectDialog(self, choices, v.get())
                    self.wait_window(dlg)
                    if dlg.result is not None:
                        v.set(dlg.result)
                tk.Button(row_f, text="Select…", command=_open_picker,
                          bg=CLR_ACCENT2, fg="#FFFFFF", font=FONT_LABEL,
                          relief="flat", padx=10, pady=3, cursor="hand2"
                          ).pack(side="left", padx=(6, 0))
                self._vars[key] = var

        btn_f = tk.Frame(inner, bg=CLR_SURFACE, pady=10)
        btn_f.pack(fill="x")
        tk.Button(btn_f, text="Save", command=self._ok,
                  bg=CLR_ACCENT, fg="#FFFFFF", font=FONT_LABEL,
                  relief="flat", padx=16, pady=6, cursor="hand2"
                  ).pack(side="right", padx=4)
        tk.Button(btn_f, text="Cancel", command=self.destroy,
                  bg=CLR_SURFACE2, fg=CLR_MUTED, font=FONT_LABEL,
                  relief="flat", padx=16, pady=6, cursor="hand2"
                  ).pack(side="right", padx=4)
        self.minsize(500, 300)

    def _ok(self):
        result = {}
        for label, key, ftype, *_ in self._fields:
            var = self._vars.get(key)
            if var is None:
                result[key] = ""
                continue
            if isinstance(var, tk.Text):
                result[key] = var.get("1.0", "end-1c").strip()
            elif isinstance(var, tk.BooleanVar):
                result[key] = int(var.get())
            else:
                result[key] = var.get().strip()
        self.result = result
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# Parts Usage Dialog — shown during task completion
# ─────────────────────────────────────────────────────────────────────────────
class PartsUsageDialog(tk.Toplevel):
    """
    Let the user record which spare parts were consumed and how many.
    Shows all parts compatible with the equipment (or all if none tagged).
    """
    def __init__(self, parent, db: Database, equipment_code: str):
        super().__init__(parent)
        self.title("Parts Used in This Task")
        self.configure(bg=CLR_SURFACE)
        self.result = None   # None = dialog dismissed; [] = confirmed with no parts
        self._db  = db
        self._rows = []

        # Anchor position and grab relative to the true top-level window
        top = parent.winfo_toplevel()
        self.update_idletasks()
        self.geometry(f"560x480+{top.winfo_rootx()+100}+{top.winfo_rooty()+60}")
        self.transient(top)
        self.grab_set()

        tk.Label(self, text="Record Parts Consumed  (leave qty = 0 to skip)",
                 bg=CLR_SURFACE, fg=CLR_TEXT, font=FONT_LABEL,
                 padx=16, pady=10).pack(anchor="w")
        tk.Frame(self, bg=CLR_BORDER, height=1).pack(fill="x", padx=12)

        # scrollable area
        canvas = tk.Canvas(self, bg=CLR_SURFACE, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)
        inner = tk.Frame(canvas, bg=CLR_SURFACE, padx=16, pady=8)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # header row
        hdr = tk.Frame(inner, bg=CLR_SURFACE)
        hdr.pack(fill="x", pady=(0, 4))
        tk.Label(hdr, text="Part Number", bg=CLR_SURFACE, fg=CLR_MUTED,
                 font=FONT_LABEL, width=14, anchor="w").pack(side="left")
        tk.Label(hdr, text="Description", bg=CLR_SURFACE, fg=CLR_MUTED,
                 font=FONT_LABEL, width=22, anchor="w").pack(side="left")
        tk.Label(hdr, text="On Hand", bg=CLR_SURFACE, fg=CLR_MUTED,
                 font=FONT_LABEL, width=8, anchor="w").pack(side="left")
        tk.Label(hdr, text="Qty Used", bg=CLR_SURFACE, fg=CLR_MUTED,
                 font=FONT_LABEL, width=9, anchor="w").pack(side="left")

        tk.Frame(inner, bg=CLR_BORDER, height=1).pack(fill="x", pady=4)

        # load parts — prefer compatible ones
        parts = db.fetchall("""
            SELECT id, part_number, description, qty_on_hand, unit
            FROM spare_parts
            WHERE is_active=1
              AND (compatible_equip LIKE ? OR compatible_equip = '' OR compatible_equip IS NULL)
            ORDER BY part_number
        """, (f"%{equipment_code}%",))
        if not parts:
            parts = db.fetchall(
                "SELECT id,part_number,description,qty_on_hand,unit FROM spare_parts WHERE is_active=1 ORDER BY part_number")

        for p in parts:
            row_f = tk.Frame(inner, bg=CLR_SURFACE)
            row_f.pack(fill="x", pady=2)
            tk.Label(row_f, text=p["part_number"], bg=CLR_SURFACE, fg=CLR_TEXT,
                     font=FONT_MONO, width=14, anchor="w").pack(side="left")
            tk.Label(row_f, text=p["description"][:28], bg=CLR_SURFACE, fg=CLR_TEXT,
                     font=FONT_BODY, width=22, anchor="w").pack(side="left")
            tk.Label(row_f, text=f"{p['qty_on_hand']:g} {p['unit']}", bg=CLR_SURFACE,
                     fg=CLR_MUTED, font=FONT_SMALL, width=8, anchor="w").pack(side="left")
            qty_var = tk.StringVar(value="0")
            tk.Entry(row_f, textvariable=qty_var, bg=CLR_BG, fg=CLR_TEXT,
                     insertbackground=CLR_TEXT, font=FONT_BODY,
                     relief="flat", bd=3, width=8).pack(side="left", padx=4)
            self._rows.append((p["id"], qty_var))

        tk.Frame(self, bg=CLR_BORDER, height=1).pack(fill="x", padx=12)
        btn_f = tk.Frame(self, bg=CLR_SURFACE, pady=10, padx=16)
        btn_f.pack(fill="x")
        tk.Button(btn_f, text="Skip Parts", command=self.destroy,
                  bg=CLR_SURFACE2, fg=CLR_MUTED, font=FONT_LABEL,
                  relief="flat", padx=12, pady=6, cursor="hand2"
                  ).pack(side="right", padx=(4,0))
        tk.Button(btn_f, text="Confirm & Deduct", command=self._ok,
                  bg=CLR_ACCENT, fg="#FFFFFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=6, cursor="hand2"
                  ).pack(side="right")

    def _ok(self):
        used = []
        for part_id, qty_var in self._rows:
            try:
                q = float(qty_var.get())
                if q > 0:
                    used.append((part_id, q))
            except ValueError:
                pass
        self.result = used
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# Attachment Manager Dialog
# ─────────────────────────────────────────────────────────────────────────────
class AttachmentDialog(tk.Toplevel):
    def __init__(self, parent, db: Database, entity_type: str, entity_id: int, title_label: str):
        super().__init__(parent)
        self.title(f"Attachments — {title_label}")
        self.configure(bg=CLR_BG)
        self.grab_set()
        self.geometry(f"720x500+{parent.winfo_rootx()+60}+{parent.winfo_rooty()+40}")
        self._db          = db
        self._entity_type = entity_type
        self._entity_id   = entity_id
        ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)
        self._build()

    def _build(self):
        # Header bar
        hdr = tk.Frame(self, bg=CLR_BG)
        hdr.pack(fill="x", padx=16, pady=10)
        tk.Label(hdr, text="Attachments", bg=CLR_BG, fg=CLR_TEXT,
                 font=FONT_HEAD).pack(side="left")
        tk.Button(hdr, text="+ Upload File", command=self._upload,
                  bg=CLR_ACCENT, fg="#FFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2"
                  ).pack(side="right")

        tk.Frame(self, bg=CLR_BORDER, height=1).pack(fill="x", padx=16)

        # Column header row
        col_hdr = tk.Frame(self, bg=CLR_SURFACE2, padx=12, pady=6)
        col_hdr.pack(fill="x", padx=16, pady=(6, 0))
        tk.Label(col_hdr, text="Filename",    bg=CLR_SURFACE2, fg=CLR_MUTED,
                 font=FONT_LABEL, width=28, anchor="w").pack(side="left")
        tk.Label(col_hdr, text="Type",        bg=CLR_SURFACE2, fg=CLR_MUTED,
                 font=FONT_LABEL, width=8,  anchor="w").pack(side="left")
        tk.Label(col_hdr, text="Size",        bg=CLR_SURFACE2, fg=CLR_MUTED,
                 font=FONT_LABEL, width=8,  anchor="w").pack(side="left")
        tk.Label(col_hdr, text="Description", bg=CLR_SURFACE2, fg=CLR_MUTED,
                 font=FONT_LABEL, width=24, anchor="w").pack(side="left")
        tk.Label(col_hdr, text="Uploaded",    bg=CLR_SURFACE2, fg=CLR_MUTED,
                 font=FONT_LABEL, width=16, anchor="w").pack(side="left")

        # Scrollable list area
        outer = tk.Frame(self, bg=CLR_BG)
        outer.pack(fill="both", expand=True, padx=16, pady=4)

        canvas = tk.Canvas(outer, bg=CLR_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self._list_frame = tk.Frame(canvas, bg=CLR_BG)
        self._canvas_win = canvas.create_window((0, 0), window=self._list_frame, anchor="nw")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(self._canvas_win, width=e.width))
        self._list_frame.bind("<Configure>",
                              lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        self._canvas = canvas

        # Button bar
        tk.Frame(self, bg=CLR_BORDER, height=1).pack(fill="x", padx=16)
        btn_f = tk.Frame(self, bg=CLR_BG, pady=8)
        btn_f.pack(fill="x", padx=16)
        tk.Button(btn_f, text="Preview / Open", command=self._preview,
                  bg=CLR_ACCENT2, fg="#FFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2"
                  ).pack(side="left", padx=(0, 4))
        tk.Button(btn_f, text="Delete Selected", command=self._delete,
                  bg=CLR_DANGER, fg="#FFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2"
                  ).pack(side="left")
        tk.Button(btn_f, text="Close", command=self.destroy,
                  bg=CLR_SURFACE2, fg=CLR_MUTED, font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2"
                  ).pack(side="right")

        self._selected_id = None
        self._load()

    def _load(self):
        # Clear existing rows
        for w in self._list_frame.winfo_children():
            w.destroy()
        self._selected_id = None

        rows = self._db.fetchall(
            "SELECT * FROM attachments WHERE entity_type=? AND entity_id=? ORDER BY uploaded_at DESC",
            (self._entity_type, self._entity_id))

        if not rows:
            tk.Label(self._list_frame, text="No attachments yet. Click '+ Upload File' to add one.",
                     bg=CLR_BG, fg=CLR_MUTED, font=FONT_BODY,
                     padx=12, pady=20).pack(anchor="w")
            return

        self._row_frames = {}
        for i, r in enumerate(rows):
            bg = CLR_ROW_EVEN if i % 2 == 0 else CLR_ROW_ODD
            row_f = tk.Frame(self._list_frame, bg=bg, padx=12, pady=7,
                             cursor="hand2")
            row_f.pack(fill="x", pady=1)

            # File type icon
            icon = "🖼" if r["file_type"] == "image" else "📄"
            size_kb = f"{r['file_size']//1024} KB" if r["file_size"] else "—"

            tk.Label(row_f, text=f"{icon}  {r['filename']}", bg=bg, fg=CLR_TEXT,
                     font=FONT_BODY, width=28, anchor="w").pack(side="left")
            tk.Label(row_f, text=r["file_type"] or "—", bg=bg, fg=CLR_MUTED,
                     font=FONT_SMALL, width=8, anchor="w").pack(side="left")
            tk.Label(row_f, text=size_kb, bg=bg, fg=CLR_MUTED,
                     font=FONT_SMALL, width=8, anchor="w").pack(side="left")
            tk.Label(row_f, text=r["description"] or "—", bg=bg, fg=CLR_MUTED,
                     font=FONT_SMALL, width=24, anchor="w").pack(side="left")
            tk.Label(row_f, text=r["uploaded_at"][:16], bg=bg, fg=CLR_MUTED,
                     font=FONT_MONO, width=16, anchor="w").pack(side="left")

            rid = r["id"]
            self._row_frames[rid] = (row_f, bg)

            # Click to select
            def _select(e, row_id=rid):
                self._on_select(row_id)
            for widget in [row_f] + row_f.winfo_children():
                widget.bind("<Button-1>", _select)
                widget.bind("<Double-Button-1>", lambda e, row_id=rid: self._preview_by_id(row_id))

    def _on_select(self, row_id):
        # Deselect previous
        if self._selected_id and self._selected_id in self._row_frames:
            prev_f, prev_bg = self._row_frames[self._selected_id]
            prev_f.config(bg=prev_bg)
            for w in prev_f.winfo_children():
                w.config(bg=prev_bg)
        # Select new
        self._selected_id = row_id
        row_f, _ = self._row_frames[row_id]
        for w in [row_f] + row_f.winfo_children():
            w.config(bg=CLR_ACCENT2)

    def _upload(self):
        paths = filedialog.askopenfilenames(
            title="Select files to attach",
            filetypes=[
                ("Images", "*.jpg *.jpeg *.png *.gif *.bmp *.webp"),
                ("Documents", "*.pdf *.docx *.xlsx *.txt *.csv *.doc *.xls"),
                ("All files", "*.*"),
            ]
        )
        for src in paths:
            src = Path(src)
            suffix  = src.suffix.lower()
            ftype   = "image" if suffix in IMAGE_EXTS else "document"
            # build unique stored name
            ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
            stored  = f"{self._entity_type}_{self._entity_id}_{ts}_{src.name}"
            dest    = ATTACHMENTS_DIR / stored
            shutil.copy2(src, dest)
            fsize   = dest.stat().st_size
            desc    = simpledialog.askstring("Description",
                f"Optional description for:\n{src.name}",
                parent=self) or ""
            self._db.execute(
                """INSERT INTO attachments(entity_type,entity_id,filename,stored_name,file_type,file_size,description)
                   VALUES(?,?,?,?,?,?,?)""",
                (self._entity_type, self._entity_id,
                 src.name, stored, ftype, fsize, desc))
        self._load()

    def _selected_row(self):
        if not getattr(self, "_selected_id", None):
            messagebox.showwarning("No selection", "Click an attachment to select it first.")
            return None
        return self._db.fetchone("SELECT * FROM attachments WHERE id=?", (self._selected_id,))

    def _preview_by_id(self, row_id):
        self._on_select(row_id)
        self._preview()

    def _preview(self):
        row = self._selected_row()
        if not row: return
        dest = ATTACHMENTS_DIR / row["stored_name"]
        if not dest.exists():
            messagebox.showerror("File Missing", f"File not found:\n{dest}")
            return
        if row["file_type"] == "image" and PIL_AVAILABLE:
            self._show_image_preview(dest, row["filename"])
        else:
            open_file_externally(dest)

    def _show_image_preview(self, path, filename):
        win = tk.Toplevel(self)
        win.title(f"Preview — {filename}")
        win.configure(bg=CLR_BG)
        win.geometry("800x600")

        try:
            img = Image.open(path)
            img.thumbnail((760, 540), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(win, image=photo, bg=CLR_BG)
            lbl.image = photo
            lbl.pack(expand=True, pady=10)
        except Exception as e:
            tk.Label(win, text=f"Cannot preview image:\n{e}",
                     bg=CLR_BG, fg=CLR_DANGER, font=FONT_BODY).pack(expand=True)

        tk.Button(win, text="Open Externally", command=lambda: open_file_externally(path),
                  bg=CLR_ACCENT2, fg="#FFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2"
                  ).pack(pady=8)

    def _delete(self):
        row = self._selected_row()
        if not row: return
        if not messagebox.askyesno("Delete Attachment",
            f"Delete '{row['filename']}'?\nThis cannot be undone."):
            return
        # remove physical file
        dest = ATTACHMENTS_DIR / row["stored_name"]
        try:
            dest.unlink(missing_ok=True)
        except Exception:
            pass
        self._db.execute("DELETE FROM attachments WHERE id=?", (row["id"],))
        self._load()


# ─────────────────────────────────────────────────────────────────────────────
# Equipment Multi Select Dialog
# ─────────────────────────────────────────────────────────────────────────────
class EquipMultiSelectDialog(tk.Toplevel):
    def __init__(self, parent, eq_choices, current_value=""):
        super().__init__(parent)
        self.title("Select Compatible Equipment")
        self.configure(bg=CLR_SURFACE)
        self.resizable(False, False)
        self.grab_set()
        self.result = None
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx()+100}+{parent.winfo_rooty()+80}")

        tk.Label(self, text="Select all compatible equipment:",
                 bg=CLR_SURFACE, fg=CLR_TEXT, font=FONT_LABEL,
                 padx=16, pady=10).pack(anchor="w")
        tk.Frame(self, bg=CLR_BORDER, height=1).pack(fill="x", padx=12)

        list_frame = tk.Frame(self, bg=CLR_SURFACE, padx=16, pady=8)
        list_frame.pack(fill="both", expand=True)

        current_codes = {c.strip() for c in current_value.split(",") if c.strip()}
        self._vars = {}
        for choice in eq_choices:
            code = choice.split(" — ")[0]
            var  = tk.BooleanVar(value=code in current_codes)
            self._vars[code] = var
            tk.Checkbutton(list_frame, text=choice, variable=var,
                           bg=CLR_SURFACE, fg=CLR_TEXT,
                           activebackground=CLR_SURFACE2, activeforeground=CLR_TEXT,
                           selectcolor=CLR_BG, font=FONT_BODY, anchor="w"
                           ).pack(fill="x", pady=2)

        tk.Frame(self, bg=CLR_BORDER, height=1).pack(fill="x", padx=12)
        btn_f = tk.Frame(self, bg=CLR_SURFACE, pady=10, padx=16)
        btn_f.pack(fill="x")
        tk.Button(btn_f, text="Clear All", command=self._clear,
                  bg=CLR_SURFACE2, fg=CLR_MUTED, font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2"
                  ).pack(side="left")
        tk.Button(btn_f, text="Cancel", command=self.destroy,
                  bg=CLR_SURFACE2, fg=CLR_MUTED, font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2"
                  ).pack(side="right", padx=(4,0))
        tk.Button(btn_f, text="Confirm", command=self._ok,
                  bg=CLR_ACCENT, fg="#FFFFFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2"
                  ).pack(side="right")
        self.minsize(320, 100)

    def _clear(self):
        for var in self._vars.values(): var.set(False)

    def _ok(self):
        self.result = ",".join(code for code, var in self._vars.items() if var.get())
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# Panel base
# ─────────────────────────────────────────────────────────────────────────────
class Panel(tk.Frame):
    def __init__(self, master, db: Database, status_bar, **kwargs):
        super().__init__(master, bg=CLR_BG, **kwargs)
        self.db         = db
        self.status_bar = status_bar
        self.build()

    def build(self): pass
    def refresh(self): pass

    def _section_label(self, parent, text):
        tk.Label(parent, text=text, bg=CLR_BG, fg=CLR_ACCENT,
                 font=FONT_HEAD).pack(anchor="w", pady=(12, 4), padx=12)

    def _toolbar_btn(self, parent, text, cmd, color=CLR_SURFACE2, fg=CLR_TEXT, icon=""):
        lbl = f"{icon}  {text}" if icon else text
        return tk.Button(parent, text=lbl, command=cmd,
                         bg=color, fg=fg, font=FONT_LABEL,
                         relief="flat", padx=12, pady=6, cursor="hand2",
                         activebackground=CLR_BORDER, activeforeground=CLR_TEXT)

    def _search_bar(self, parent, var, callback):
        f = tk.Frame(parent, bg=CLR_BG)
        tk.Label(f, text="🔍", bg=CLR_BG, fg=CLR_MUTED,
                 font=("Segoe UI", 11)).pack(side="left", padx=(0,4))
        e = tk.Entry(f, textvariable=var, bg=CLR_SURFACE, fg=CLR_TEXT,
                     insertbackground=CLR_TEXT, font=FONT_BODY,
                     relief="flat", bd=4, width=26)
        e.pack(side="left")
        var.trace_add("write", lambda *_: callback())
        return f

    def _build_context_menu(self, entries):
        """
        entries = list of (label, command) or None for separator.
        Returns a tk.Menu (dark-themed).
        """
        menu = tk.Menu(self, tearoff=0,
                       bg=CLR_SURFACE2, fg=CLR_TEXT,
                       activebackground=CLR_ACCENT2,
                       activeforeground="#FFFFFF",
                       relief="flat", bd=0,
                       font=FONT_BODY)
        for item in entries:
            if item is None:
                menu.add_separator()
            else:
                label, cmd = item
                menu.add_command(label=label, command=cmd)
        return menu


# ─────────────────────────────────────────────────────────────────────────────
# PANEL 1 — Dashboard
# ─────────────────────────────────────────────────────────────────────────────
class DashboardPanel(Panel):
    def build(self):
        hdr = tk.Frame(self, bg=CLR_BG)
        hdr.pack(fill="x", padx=20, pady=(18, 0))
        facility = self.db.setting("facility_name", "Equipment Management System")
        tk.Label(hdr, text=facility, bg=CLR_BG, fg=CLR_TEXT, font=FONT_TITLE).pack(side="left")
        tk.Button(hdr, text="⟳ Refresh", command=self.refresh,
                  bg=CLR_SURFACE2, fg=CLR_MUTED, font=FONT_SMALL,
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="right")

        tk.Frame(self, bg=CLR_ACCENT, height=2).pack(fill="x", padx=20, pady=8)
        self._kpi_frame = tk.Frame(self, bg=CLR_BG)
        self._kpi_frame.pack(fill="x", padx=20, pady=4)
        self._sect_due  = tk.Frame(self, bg=CLR_BG)
        self._sect_due.pack(fill="both", expand=True, padx=20, pady=4)
        self.refresh()

    def refresh(self):
        self._build_kpis()
        self._build_due_list()

    def _kpi_card(self, parent, title, value, subtitle="", color=CLR_ACCENT):
        f = tk.Frame(parent, bg=CLR_SURFACE, padx=16, pady=14, relief="flat", bd=0)
        f.pack(side="left", fill="both", expand=True, padx=6)
        tk.Label(f, text=title,    bg=CLR_SURFACE, fg=CLR_MUTED, font=FONT_SMALL).pack(anchor="w")
        tk.Label(f, text=str(value),bg=CLR_SURFACE, fg=color, font=("Segoe UI", 26, "bold")).pack(anchor="w")
        if subtitle:
            tk.Label(f, text=subtitle, bg=CLR_SURFACE, fg=CLR_MUTED, font=FONT_SMALL).pack(anchor="w")
        return f

    def _build_kpis(self):
        for w in self._kpi_frame.winfo_children(): w.destroy()
        total_eq   = self.db.fetchone("SELECT COUNT(*) c FROM equipment WHERE status='Active'")["c"]
        today_str  = date.today().isoformat()
        alert_days = int(self.db.setting("alert_soon_days", 7))
        soon_date  = (date.today() + timedelta(days=alert_days)).isoformat()
        overdue    = self.db.fetchone("SELECT COUNT(*) c FROM maintenance_tasks WHERE is_active=1 AND next_due < ?", (today_str,))["c"]
        due_soon   = self.db.fetchone("SELECT COUNT(*) c FROM maintenance_tasks WHERE is_active=1 AND next_due >= ? AND next_due <= ?", (today_str, soon_date))["c"]
        parts_low  = self.db.fetchone("SELECT COUNT(*) c FROM spare_parts WHERE is_active=1 AND qty_on_hand < safety_stock")["c"]
        parts_total= self.db.fetchone("SELECT COUNT(*) c FROM spare_parts WHERE is_active=1")["c"]
        inv_value  = self.db.fetchone("SELECT SUM(qty_on_hand * unit_cost) v FROM spare_parts WHERE is_active=1")["v"] or 0
        curr       = self.db.setting("currency", "USD")

        self._kpi_card(self._kpi_frame, "Active Equipment",   total_eq,  "machines registered")
        self._kpi_card(self._kpi_frame, "Overdue Tasks",      overdue,   "need attention now",     CLR_DANGER)
        self._kpi_card(self._kpi_frame, "Due Soon",           due_soon,  f"within {alert_days} days", CLR_WARN)
        self._kpi_card(self._kpi_frame, "Parts Below Safety", parts_low, f"of {parts_total} parts",   CLR_DANGER if parts_low else CLR_OK)
        self._kpi_card(self._kpi_frame, "Inventory Value",    f"{curr} {inv_value:,.2f}", "total parts value", CLR_ACCENT2)

    def _build_due_list(self):
        for w in self._sect_due.winfo_children(): w.destroy()
        cols_frame = tk.Frame(self._sect_due, bg=CLR_BG)
        cols_frame.pack(fill="both", expand=True)

        left = tk.Frame(cols_frame, bg=CLR_BG)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))
        tk.Label(left, text="Maintenance Alerts", bg=CLR_BG, fg=CLR_ACCENT, font=FONT_HEAD).pack(anchor="w", pady=(6,4))

        alert_days = int(self.db.setting("alert_soon_days", 7))
        soon_date  = (date.today() + timedelta(days=alert_days)).isoformat()

        tasks = self.db.fetchall("""
            SELECT mt.id, mt.task_name, mt.next_due, mt.assigned_to, e.code, e.name
            FROM maintenance_tasks mt
            JOIN equipment e ON e.id = mt.equipment_id
            WHERE mt.is_active=1 AND e.status='Active'
              AND (mt.next_due IS NULL OR mt.next_due <= ?)
            ORDER BY mt.next_due ASC LIMIT 30
        """, (soon_date,))

        tv = StyledTreeview(left,
            columns=("equip","task","due","days","assignee"),
            headings=("Equipment","Task","Due Date","Days","Assigned To"),
            col_widths={"equip":100,"task":160,"due":90,"days":60,"assignee":90})
        tv.pack(fill="both", expand=True)
        for t in tasks:
            d = days_until(t["next_due"])
            u = urgency(d)
            tag = {"CRITICAL":"critical","DUE TODAY":"due","DUE SOON":"warning"}.get(u,"ok")
            days_str = f"{d}d" if d is not None else "—"
            if d is not None and d < 0: days_str = f"{abs(d)}d OVERDUE"
            tv.insert("", "end", values=(t["code"], t["task_name"], t["next_due"] or "Not set", days_str, t["assigned_to"] or "—"), tags=(tag,))

        right = tk.Frame(cols_frame, bg=CLR_BG)
        right.pack(side="left", fill="both", expand=True, padx=(8,0))
        tk.Label(right, text="Low Stock Alerts", bg=CLR_BG, fg=CLR_ACCENT, font=FONT_HEAD).pack(anchor="w", pady=(6,4))

        parts = self.db.fetchall("""
            SELECT part_number, description, qty_on_hand, safety_stock, unit, supplier
            FROM spare_parts WHERE is_active=1 AND qty_on_hand < safety_stock
            ORDER BY (qty_on_hand - safety_stock) ASC LIMIT 20""")
        pv = StyledTreeview(right,
            columns=("pn","desc","qty","safety","unit","supplier"),
            headings=("Part No.","Description","On Hand","Safety Stock","Unit","Supplier"),
            col_widths={"pn":100,"desc":150,"qty":70,"safety":90,"unit":50,"supplier":90})
        pv.pack(fill="both", expand=True)
        for p in parts:
            pv.insert("", "end", values=(p["part_number"], p["description"],
                f"{p['qty_on_hand']:g}", f"{p['safety_stock']:g}", p["unit"], p["supplier"] or "—"),
                tags=("critical",))


# ─────────────────────────────────────────────────────────────────────────────
# PANEL 1b — Calendar View (sub-window shown from Maintenance panel)
# ─────────────────────────────────────────────────────────────────────────────
class CalendarWindow(tk.Toplevel):
    """Monthly calendar with maintenance tasks colour-coded by urgency."""

    DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    def __init__(self, parent, db: Database):
        super().__init__(parent)
        self.title("Maintenance Calendar")
        self.configure(bg=CLR_BG)
        self.geometry("1000x680")
        self._db = db
        self._today = date.today()
        self._year  = self._today.year
        self._month = self._today.month
        self._build()
        self._render()

    def _build(self):
        # navigation bar
        nav = tk.Frame(self, bg=CLR_BG)
        nav.pack(fill="x", padx=16, pady=10)
        tk.Button(nav, text="◀", command=self._prev_month,
                  bg=CLR_SURFACE2, fg=CLR_TEXT, font=FONT_HEAD,
                  relief="flat", padx=12, pady=4, cursor="hand2").pack(side="left")
        self._month_lbl = tk.Label(nav, text="", bg=CLR_BG, fg=CLR_TEXT, font=FONT_TITLE)
        self._month_lbl.pack(side="left", padx=20)
        tk.Button(nav, text="▶", command=self._next_month,
                  bg=CLR_SURFACE2, fg=CLR_TEXT, font=FONT_HEAD,
                  relief="flat", padx=12, pady=4, cursor="hand2").pack(side="left")
        tk.Button(nav, text="Today", command=self._go_today,
                  bg=CLR_ACCENT2, fg="#FFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=4, cursor="hand2").pack(side="left", padx=12)

        # legend
        leg = tk.Frame(self, bg=CLR_BG)
        leg.pack(fill="x", padx=16, pady=(0,6))
        for clr, label in [(CLR_CAL_CRIT,"Overdue"),(CLR_CAL_SOON,"Due Soon"),(CLR_CAL_OK,"Scheduled"),(CLR_CAL_TODAY,"Today")]:
            tk.Frame(leg, bg=clr, width=14, height=14).pack(side="left", padx=(0,4))
            tk.Label(leg, text=label, bg=CLR_BG, fg=CLR_MUTED, font=FONT_SMALL).pack(side="left", padx=(0,16))

        tk.Frame(self, bg=CLR_BORDER, height=1).pack(fill="x", padx=16, pady=4)

        # day-of-week header
        dow = tk.Frame(self, bg=CLR_BG)
        dow.pack(fill="x", padx=16)
        for d in self.DAYS:
            tk.Label(dow, text=d, bg=CLR_BG, fg=CLR_MUTED, font=FONT_LABEL,
                     width=14, anchor="center").pack(side="left", expand=True, fill="x")

        # calendar grid container
        self._grid_frame = tk.Frame(self, bg=CLR_BG)
        self._grid_frame.pack(fill="both", expand=True, padx=16, pady=(4,12))

        # day detail pane at bottom
        self._detail_var = tk.StringVar(value="Click a day to see tasks.")
        self._detail_frame = tk.Frame(self, bg=CLR_SURFACE, height=130)
        self._detail_frame.pack(fill="x", padx=16, pady=(0,12))
        self._detail_frame.pack_propagate(False)
        self._detail_lbl = tk.Label(self._detail_frame, textvariable=self._detail_var,
                                    bg=CLR_SURFACE, fg=CLR_TEXT, font=FONT_BODY,
                                    justify="left", anchor="nw", padx=12, pady=8,
                                    wraplength=960)
        self._detail_lbl.pack(fill="both", expand=True)

    def _render(self):
        for w in self._grid_frame.winfo_children():
            w.destroy()

        self._month_lbl.config(text=f"{self._year} — {date(self._year, self._month, 1).strftime('%B')}")

        # build task map: {date_str: [(task_name, equip_code, urgency_tag)]}
        first = date(self._year, self._month, 1)
        if self._month == 12:
            last = date(self._year + 1, 1, 1) - timedelta(days=1)
        else:
            last = date(self._year, self._month + 1, 1) - timedelta(days=1)

        tasks = self._db.fetchall("""
            SELECT mt.task_name, mt.next_due, e.code
            FROM maintenance_tasks mt
            JOIN equipment e ON e.id = mt.equipment_id
            WHERE mt.is_active=1 AND mt.next_due IS NOT NULL
              AND mt.next_due BETWEEN ? AND ?
            ORDER BY mt.next_due
        """, (first.isoformat(), last.isoformat()))

        task_map = {}
        for t in tasks:
            d = days_until(t["next_due"])
            u = urgency(d)
            task_map.setdefault(t["next_due"][:10], []).append((t["task_name"], t["code"], u))

        # calendar grid — ISO week: Mon=0, Sun=6
        start_dow = first.weekday()  # 0=Mon
        cells_needed = start_dow + last.day
        rows_needed  = (cells_needed + 6) // 7

        for row in range(rows_needed):
            row_f = tk.Frame(self._grid_frame, bg=CLR_BG)
            row_f.pack(fill="both", expand=True)
            for col in range(7):
                cell_idx = row * 7 + col
                day_num  = cell_idx - start_dow + 1
                cell = tk.Frame(row_f, bg=CLR_SURFACE, relief="flat",
                                bd=0, padx=4, pady=4)
                cell.pack(side="left", fill="both", expand=True, padx=2, pady=2)

                if day_num < 1 or day_num > last.day:
                    cell.config(bg=CLR_BG)
                    continue

                this_date = date(self._year, self._month, day_num)
                dstr      = this_date.isoformat()
                is_today  = (this_date == self._today)

                # date number label
                num_color = CLR_CAL_TODAY if is_today else CLR_TEXT
                num_bg    = "#006655" if is_today else CLR_SURFACE
                cell.config(bg=num_bg)
                tk.Label(cell, text=str(day_num), bg=num_bg, fg=num_color,
                         font=FONT_LABEL, anchor="ne").pack(anchor="ne")

                # tasks for this day
                day_tasks = task_map.get(dstr, [])
                for task_name, eq_code, u in day_tasks[:4]:
                    dot_clr = {URGENCY_CRITICAL: CLR_CAL_CRIT,
                               URGENCY_DUE:      CLR_CAL_SOON,
                               URGENCY_SOON:     CLR_CAL_SOON,
                               URGENCY_OK:       CLR_CAL_OK}.get(u, CLR_CAL_OK)
                    tag_f = tk.Frame(cell, bg=dot_clr, padx=3, pady=1)
                    tag_f.pack(fill="x", pady=1)
                    short = f"{eq_code}: {task_name[:18]}"
                    tk.Label(tag_f, text=short, bg=dot_clr, fg="#FFF",
                             font=("Segoe UI", 7), anchor="w").pack(anchor="w")
                if len(day_tasks) > 4:
                    tk.Label(cell, text=f"+{len(day_tasks)-4} more",
                             bg=num_bg, fg=CLR_MUTED, font=("Segoe UI", 7)).pack(anchor="w")

                # click binding
                def _click(e, d=dstr, dt=day_tasks):
                    self._show_day(d, dt)
                cell.bind("<Button-1>", _click)
                for child in cell.winfo_children():
                    child.bind("<Button-1>", _click)

    def _show_day(self, dstr, day_tasks):
        if not day_tasks:
            self._detail_var.set(f"{dstr}  —  No maintenance tasks scheduled.")
            return
        lines = [f"Tasks on {dstr}:"]
        for task_name, eq_code, u in day_tasks:
            lines.append(f"  • [{u}]  {eq_code} — {task_name}")
        self._detail_var.set("\n".join(lines))

    def _prev_month(self):
        if self._month == 1:
            self._month = 12; self._year -= 1
        else:
            self._month -= 1
        self._render()

    def _next_month(self):
        if self._month == 12:
            self._month = 1; self._year += 1
        else:
            self._month += 1
        self._render()

    def _go_today(self):
        self._year  = self._today.year
        self._month = self._today.month
        self._render()


# ─────────────────────────────────────────────────────────────────────────────
# PANEL 2 — Maintenance
# ─────────────────────────────────────────────────────────────────────────────
class MaintenancePanel(Panel):
    def build(self):
        self._search_var = tk.StringVar()
        self._filter_var = tk.StringVar(value="All")

        tb = tk.Frame(self, bg=CLR_BG)
        tb.pack(fill="x", padx=12, pady=10)
        tk.Label(tb, text="Maintenance Tracker", bg=CLR_BG, fg=CLR_TEXT, font=FONT_HEAD).pack(side="left", padx=4)

        self._toolbar_btn(tb, "Mark Complete",  self._complete_task,  CLR_OK,     "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "Bulk Complete",  self._bulk_complete,  CLR_OK,     "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "+ Add Task",     self._add_task,       CLR_ACCENT, "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "Edit Task",      self._edit_task).pack(side="right", padx=4)
        self._toolbar_btn(tb, "View History",   self._view_history).pack(side="right", padx=4)
        self._toolbar_btn(tb, "📅 Calendar",   self._open_calendar,  CLR_ACCENT2,"#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "⟳",             self.refresh).pack(side="right", padx=4)

        fb = tk.Frame(self, bg=CLR_BG)
        fb.pack(fill="x", padx=12, pady=(0,6))
        sb = self._search_bar(fb, self._search_var, self.refresh)
        sb.pack(side="left", padx=(0,12))
        for label, val in [("All","All"),("⚠ Overdue","overdue"),("Due Soon","soon"),("OK","ok")]:
            tk.Radiobutton(fb, text=label, variable=self._filter_var, value=val,
                           command=self.refresh, bg=CLR_BG, fg=CLR_MUTED,
                           selectcolor=CLR_BG, activebackground=CLR_BG,
                           font=FONT_SMALL).pack(side="left", padx=6)

        hint = tk.Label(fb, text="Ctrl+click or Shift+click to select multiple rows for Bulk Complete",
                        bg=CLR_BG, fg=CLR_MUTED, font=("Segoe UI", 8))
        hint.pack(side="right", padx=8)

        self._tv = StyledTreeview(self,
            columns=("code","equip","task","freq","last","next","days","status","assigned"),
            headings=("Code","Equipment","Task","Freq(d)","Last Done","Next Due","Days Left","Status","Assigned To"),
            col_widths={"code":80,"equip":140,"task":180,"freq":60,"last":90,"next":90,"days":80,"status":90,"assigned":100},
            selectmode="extended")
        self._tv.pack(fill="both", expand=True, padx=12, pady=4)

        # Double-click to edit
        self._tv.bind_double_click(lambda e: self._edit_task())

        # Right-click context menu
        def _ctx(event):
            sel = self._tv.selection()
            n = len(sel)
            return self._build_context_menu([
                ("✏ Edit Task",          self._edit_task),
                ("✔ Mark Complete",       self._complete_task),
                (f"✔✔ Bulk Complete ({n} selected)", self._bulk_complete) if n > 1 else None,
                None,
                ("📋 View History",       self._view_history),
                None,
                ("🗑 Delete Task",        self._delete_task),
            ])
        self._tv.bind_context_menu(_ctx)
        self.refresh()

    def refresh(self):
        alert_days   = int(self.db.setting("alert_soon_days", 7))
        soon_date    = (date.today() + timedelta(days=alert_days)).isoformat()
        today_str    = date.today().isoformat()
        query_filter = self._filter_var.get()
        search       = self._search_var.get().lower()

        sql = """SELECT mt.id, e.code, e.name, mt.task_name, mt.frequency_days,
                        mt.last_done, mt.next_due, mt.assigned_to
                 FROM maintenance_tasks mt
                 JOIN equipment e ON e.id = mt.equipment_id
                 WHERE mt.is_active=1"""
        params = []
        if query_filter == "overdue":
            sql += " AND mt.next_due < ?"; params.append(today_str)
        elif query_filter == "soon":
            sql += " AND mt.next_due >= ? AND mt.next_due <= ?"; params += [today_str, soon_date]
        elif query_filter == "ok":
            sql += " AND mt.next_due > ?"; params.append(soon_date)
        sql += " ORDER BY mt.next_due ASC"

        rows = self.db.fetchall(sql, params)
        self._tv.delete(*self._tv.get_children())
        for r in rows:
            if search and search not in (r["code"]+r["name"]+r["task_name"]).lower():
                continue
            d = days_until(r["next_due"])
            u = urgency(d)
            days_str = "—" if d is None else (f"{d}d" if d >= 0 else f"{abs(d)}d OVERDUE")
            tag = {"CRITICAL":"critical","DUE TODAY":"due","DUE SOON":"warning","OK":"ok"}.get(u,"ok")
            self._tv.insert("", "end", iid=str(r["id"]),
                values=(r["code"], r["name"], r["task_name"], r["frequency_days"],
                        r["last_done"] or "Never", r["next_due"] or "Not set",
                        days_str, u, r["assigned_to"] or "—"),
                tags=(tag,))

    def _selected_task_id(self):
        sel = self._tv.selection()
        if not sel:
            messagebox.showwarning("No selection", "Please select a task first.")
            return None
        return int(sel[0])

    def _do_complete_one(self, tid, done_date, done_by, duration_min, findings, parts_used):
        """Core completion logic for a single task — called by both complete and bulk."""
        task = self.db.fetchone(
            "SELECT mt.*, e.code FROM maintenance_tasks mt "
            "JOIN equipment e ON e.id=mt.equipment_id WHERE mt.id=?", (tid,))
        if not task: return None
        try:
            next_due = (datetime.strptime(done_date, "%Y-%m-%d") +
                        timedelta(days=task["frequency_days"])).strftime("%Y-%m-%d")
        except Exception:
            next_due = None
        self.db.execute(
            "UPDATE maintenance_tasks SET last_done=?, next_due=? WHERE id=?",
            (done_date, next_due, tid))
        log_id = self.db.lastrowid(
            """INSERT INTO maintenance_log
               (task_id,equipment_id,done_date,done_by,duration_min,findings,next_due)
               VALUES(?,?,?,?,?,?,?)""",
            (tid, task["equipment_id"], done_date, done_by, duration_min, findings, next_due))
        # deduct parts
        for part_id, qty in parts_used:
            part = self.db.fetchone("SELECT * FROM spare_parts WHERE id=?", (part_id,))
            if not part: continue
            qty_before = part["qty_on_hand"]
            qty_after  = max(0, qty_before - qty)
            self.db.execute("UPDATE spare_parts SET qty_on_hand=? WHERE id=?", (qty_after, part_id))
            self.db.execute(
                """INSERT INTO maintenance_log_parts(log_id,part_id,qty_used) VALUES(?,?,?)""",
                (log_id, part_id, qty))
            self.db.execute(
                """INSERT INTO parts_transactions(part_id,txn_type,qty,qty_before,qty_after,ref_task,performed_by,remarks)
                   VALUES(?,?,?,?,?,?,?,?)""",
                (part_id, "Check Out", qty, qty_before, qty_after,
                 task["task_name"], done_by, f"Maintenance: {task['task_name']}"))
        return next_due

    def _complete_task(self):
        tid = self._selected_task_id()
        if tid is None: return
        task = self.db.fetchone(
            "SELECT mt.*, e.code FROM maintenance_tasks mt "
            "JOIN equipment e ON e.id=mt.equipment_id WHERE mt.id=?", (tid,))
        if not task: return

        default_user = self.db.setting("default_user", "")
        dlg = FormDialog(self, "Record Completion", [
            ("Done Date",      "done_date",    "date"),
            ("Done By",        "done_by",      "text"),
            ("Duration (min)", "duration_min", "int"),
            ("Findings/Notes", "findings",     "textarea"),
        ], initial={"done_date": date.today().isoformat(), "done_by": default_user})
        self.wait_window(dlg)
        if not dlg.result: return

        done_date = dlg.result["done_date"] or date.today().isoformat()
        try:
            dur = int(dlg.result["duration_min"]) if dlg.result["duration_min"] else None
        except ValueError:
            dur = None

        # Parts usage
        # Parts usage — None means dialog was dismissed without confirming
        parts_dlg = PartsUsageDialog(self, self.db, task["code"])
        self.wait_window(parts_dlg)
        parts_used = parts_dlg.result if parts_dlg.result is not None else []

        next_due = self._do_complete_one(
            tid, done_date, dlg.result["done_by"], dur,
            dlg.result["findings"], parts_used)
        self.status_bar.set(f"Task completed. Next due: {next_due}", CLR_OK)
        self.refresh()

    def _bulk_complete(self):
        sel = self._tv.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select multiple tasks first (Ctrl+click).")
            return
        if len(sel) == 1:
            self._complete_task()
            return

        default_user = self.db.setting("default_user", "")
        dlg = FormDialog(self, f"Bulk Complete — {len(sel)} Tasks", [
            ("Done Date",      "done_date",    "date"),
            ("Done By",        "done_by",      "text"),
            ("Duration (min)", "duration_min", "int"),
            ("Findings/Notes", "findings",     "textarea"),
        ], initial={"done_date": date.today().isoformat(), "done_by": default_user})
        self.wait_window(dlg)
        if not dlg.result: return

        done_date = dlg.result["done_date"] or date.today().isoformat()
        try:
            dur = int(dlg.result["duration_min"]) if dlg.result["duration_min"] else None
        except ValueError:
            dur = None

        completed = 0
        for iid in sel:
            self._do_complete_one(int(iid), done_date, dlg.result["done_by"],
                                  dur, dlg.result["findings"], [])
            completed += 1

        self.status_bar.set(f"Bulk complete: {completed} tasks marked done.", CLR_OK)
        self.refresh()

    def _add_task(self):
        equips = self.db.fetchall("SELECT id, code, name FROM equipment WHERE status='Active' ORDER BY code")
        eq_choices = [f"{e['code']} — {e['name']}" for e in equips]
        eq_map     = {f"{e['code']} — {e['name']}": e["id"] for e in equips}
        if not eq_choices:
            messagebox.showerror("No Equipment", "Add equipment first via Config.")
            return
        dlg = FormDialog(self, "Add Maintenance Task", [
            ("Equipment",       "equipment",       "combo",  eq_choices),
            ("Task Name",       "task_name",       "text"),
            ("Frequency (days)","frequency_days",  "int"),
            ("Last Done",       "last_done",        "date"),
            ("Assigned To",     "assigned_to",      "text"),
            ("Procedure Notes", "procedure_notes",  "textarea"),
        ])
        self.wait_window(dlg)
        if not dlg.result: return
        r     = dlg.result
        eq_id = eq_map.get(r["equipment"])
        if not eq_id: return
        freq  = int(r["frequency_days"]) if r["frequency_days"] else 30
        last  = r["last_done"] or None
        next_due = None
        if last:
            try:
                next_due = (datetime.strptime(last, "%Y-%m-%d") + timedelta(days=freq)).strftime("%Y-%m-%d")
            except Exception: pass
        self.db.execute(
            "INSERT INTO maintenance_tasks(equipment_id,task_name,frequency_days,last_done,next_due,assigned_to,procedure_notes) VALUES(?,?,?,?,?,?,?)",
            (eq_id, r["task_name"], freq, last, next_due, r["assigned_to"], r["procedure_notes"]))
        self.refresh()

    def _edit_task(self):
        tid = self._selected_task_id()
        if tid is None: return
        task   = self.db.fetchone("SELECT * FROM maintenance_tasks WHERE id=?", (tid,))
        equips = self.db.fetchall("SELECT id, code, name FROM equipment WHERE status='Active' ORDER BY code")
        eq_choices = [f"{e['code']} — {e['name']}" for e in equips]
        eq_map_r   = {e["id"]: f"{e['code']} — {e['name']}" for e in equips}
        eq_map     = {v: k for k, v in eq_map_r.items()}
        initial    = dict(task)
        initial["equipment"] = eq_map_r.get(task["equipment_id"], "")

        dlg = FormDialog(self, "Edit Maintenance Task", [
            ("Equipment",       "equipment",      "combo",   eq_choices),
            ("Task Name",       "task_name",      "text"),
            ("Frequency (days)","frequency_days", "int"),
            ("Last Done",       "last_done",      "date"),
            ("Assigned To",     "assigned_to",    "text"),
            ("Procedure Notes", "procedure_notes","textarea"),
        ], initial=initial)
        self.wait_window(dlg)
        if not dlg.result: return
        r    = dlg.result
        eq_id= eq_map.get(r["equipment"], task["equipment_id"])
        freq = int(r["frequency_days"]) if r["frequency_days"] else 30
        last = r["last_done"] or None
        next_due = None
        if last:
            try:
                next_due = (datetime.strptime(last, "%Y-%m-%d") + timedelta(days=freq)).strftime("%Y-%m-%d")
            except Exception: pass
        self.db.execute(
            "UPDATE maintenance_tasks SET equipment_id=?,task_name=?,frequency_days=?,last_done=?,next_due=?,assigned_to=?,procedure_notes=? WHERE id=?",
            (eq_id, r["task_name"], freq, last, next_due, r["assigned_to"], r["procedure_notes"], tid))
        self.refresh()

    def _delete_task(self):
        tid = self._selected_task_id()
        if tid is None: return
        if not messagebox.askyesno("Delete Task", "Permanently delete this maintenance task and all its history?"): return
        self.db.execute("DELETE FROM maintenance_tasks WHERE id=?", (tid,))
        self.refresh()

    def _view_history(self):
        tid = self._selected_task_id()
        if tid is None: return
        task = self.db.fetchone(
            "SELECT mt.task_name, e.code, e.name FROM maintenance_tasks mt "
            "JOIN equipment e ON e.id=mt.equipment_id WHERE mt.id=?", (tid,))
        logs = self.db.fetchall(
            "SELECT * FROM maintenance_log WHERE task_id=? ORDER BY done_date DESC", (tid,))

        win = tk.Toplevel(self)
        win.title(f"History — {task['code']} / {task['task_name']}")
        win.configure(bg=CLR_BG); win.geometry("780x460")
        tk.Label(win, text=f"{task['code']} · {task['name']} — {task['task_name']}",
                 bg=CLR_BG, fg=CLR_TEXT, font=FONT_HEAD).pack(padx=16, pady=10, anchor="w")

        tv = StyledTreeview(win,
            columns=("date","by","duration","findings","next","parts"),
            headings=("Done Date","Done By","Duration(min)","Findings","Next Due","Parts Used"),
            col_widths={"date":90,"by":100,"duration":90,"findings":200,"next":90,"parts":140})
        tv.pack(fill="both", expand=True, padx=12, pady=8)

        for lg in logs:
            parts_used = self.db.fetchall("""
                SELECT sp.part_number, mlp.qty_used, sp.unit
                FROM maintenance_log_parts mlp
                JOIN spare_parts sp ON sp.id = mlp.part_id
                WHERE mlp.log_id=?""", (lg["id"],))
            parts_str = "; ".join(f"{p['part_number']}×{p['qty_used']:g}" for p in parts_used) or "—"
            tv.insert("", "end",
                values=(lg["done_date"], lg["done_by"] or "—",
                        lg["duration_min"] or "—", lg["findings"] or "—",
                        lg["next_due"] or "—", parts_str))
        tv.recolor_rows()

        tk.Button(win, text="Export History",
                  command=lambda: self._export_history(tid, task, logs),
                  bg=CLR_ACCENT2, fg="#FFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=6).pack(padx=12, pady=8, anchor="e")

    def _export_history(self, tid, task, logs):
        EXPORT_DIR.mkdir(exist_ok=True)
        fname = EXPORT_DIR / f"maint_history_{task['code']}_{tid}_{date.today()}.xlsx"
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Maintenance History"
        ws.append(["Equipment Code", task["code"]])
        ws.append(["Equipment Name", task["name"]])
        ws.append(["Task",           task["task_name"]])
        ws.append([])
        ws.append(["Done Date","Done By","Duration(min)","Findings","Next Due","Parts Used","Logged At"])
        for lg in logs:
            parts_used = self.db.fetchall(
                "SELECT sp.part_number,mlp.qty_used FROM maintenance_log_parts mlp JOIN spare_parts sp ON sp.id=mlp.part_id WHERE mlp.log_id=?",
                (lg["id"],))
            pstr = "; ".join(f"{p['part_number']}×{p['qty_used']:g}" for p in parts_used) or ""
            ws.append([lg["done_date"], lg["done_by"], lg["duration_min"],
                       lg["findings"], lg["next_due"], pstr, lg["logged_at"]])
        wb.save(fname)
        messagebox.showinfo("Exported", f"Saved to:\n{fname}")

    def _open_calendar(self):
        CalendarWindow(self, self.db)


# ─────────────────────────────────────────────────────────────────────────────
# PANEL 3 — Equipment Registry
# ─────────────────────────────────────────────────────────────────────────────
class EquipmentPanel(Panel):
    def build(self):
        self._search_var = tk.StringVar()

        tb = tk.Frame(self, bg=CLR_BG)
        tb.pack(fill="x", padx=12, pady=10)
        tk.Label(tb, text="Equipment Registry", bg=CLR_BG, fg=CLR_TEXT, font=FONT_HEAD).pack(side="left", padx=4)
        self._toolbar_btn(tb, "+ Add Equipment",self._add,        CLR_ACCENT, "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "Edit",           self._edit).pack(side="right", padx=4)
        self._toolbar_btn(tb, "Attachments",    self._attachments, CLR_ACCENT2, "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "Deactivate",     self._deactivate, CLR_DANGER, "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "⟳",             self.refresh).pack(side="right", padx=4)

        fb = tk.Frame(self, bg=CLR_BG)
        fb.pack(fill="x", padx=12, pady=(0,6))
        self._search_bar(fb, self._search_var, self.refresh).pack(side="left")

        self._tv = StyledTreeview(self,
            columns=("code","name","category","location","model","serial","status","attachments","notes"),
            headings=("Code","Name","Category","Location","Model","Serial No.","Status","Files","Notes"),
            col_widths={"code":80,"name":160,"category":100,"location":100,"model":120,"serial":100,"status":70,"attachments":50,"notes":180})
        self._tv.pack(fill="both", expand=True, padx=12, pady=4)
        self._tv.bind_double_click(lambda e: self._edit())

        def _ctx(event):
            return self._build_context_menu([
                ("✏ Edit",            self._edit),
                ("📎 Attachments",    self._attachments),
                None,
                ("🔴 Deactivate",     self._deactivate),
            ])
        self._tv.bind_context_menu(_ctx)
        self.refresh()

    def refresh(self):
        search = self._search_var.get().lower()
        rows   = self.db.fetchall("SELECT * FROM equipment ORDER BY code")
        self._tv.delete(*self._tv.get_children())
        for r in rows:
            if search and search not in (r["code"]+r["name"]+r["category"]).lower():
                continue
            att_count = self.db.fetchone(
                "SELECT COUNT(*) c FROM attachments WHERE entity_type='equipment' AND entity_id=?",
                (r["id"],))["c"]
            tag = "ok" if r["status"] == "Active" else "critical"
            self._tv.insert("", "end", iid=str(r["id"]),
                values=(r["code"], r["name"], r["category"], r["location"],
                        r["model"], r["serial_no"], r["status"],
                        f"📎 {att_count}" if att_count else "—",
                        r["notes"]),
                tags=(tag,))
        self._tv.recolor_rows()

    def _fields(self):
        return [
            ("Code",       "code",      "text"),
            ("Name",       "name",      "text"),
            ("Category",   "category",  "text"),
            ("Location",   "location",  "text"),
            ("Model",      "model",     "text"),
            ("Serial No.", "serial_no", "text"),
            ("Status",     "status",    "combo", ["Active","Under Repair","Decommissioned","Standby"]),
            ("Notes",      "notes",     "textarea"),
        ]

    def _add(self):
        dlg = FormDialog(self, "Add Equipment", self._fields())
        self.wait_window(dlg)
        if not dlg.result: return
        try:
            self.db.execute(
                "INSERT INTO equipment(code,name,category,location,model,serial_no,status,notes) VALUES(:code,:name,:category,:location,:model,:serial_no,:status,:notes)",
                dlg.result)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        self.refresh()

    def _edit(self):
        sel = self._tv.selection()
        if not sel: return
        eq  = self.db.fetchone("SELECT * FROM equipment WHERE id=?", (int(sel[0]),))
        dlg = FormDialog(self, "Edit Equipment", self._fields(), initial=dict(eq))
        self.wait_window(dlg)
        if not dlg.result: return
        r    = dlg.result; r["id"] = int(sel[0])
        self.db.execute(
            "UPDATE equipment SET code=:code,name=:name,category=:category,location=:location,model=:model,serial_no=:serial_no,status=:status,notes=:notes WHERE id=:id",
            r)
        self.refresh()

    def _deactivate(self):
        sel = self._tv.selection()
        if not sel: return
        if not messagebox.askyesno("Confirm", "Deactivate selected equipment?"): return
        self.db.execute("UPDATE equipment SET status='Decommissioned' WHERE id=?", (int(sel[0]),))
        self.refresh()

    def _attachments(self):
        sel = self._tv.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select an equipment first.")
            return
        eq  = self.db.fetchone("SELECT code, name FROM equipment WHERE id=?", (int(sel[0]),))
        AttachmentDialog(self, self.db, "equipment", int(sel[0]),
                         f"{eq['code']} — {eq['name']}")


# ─────────────────────────────────────────────────────────────────────────────
# PANEL 4 — Spare Parts Inventory
# ─────────────────────────────────────────────────────────────────────────────
class PartsPanel(Panel):
    def build(self):
        self._search_var = tk.StringVar()
        self._show_low   = tk.BooleanVar(value=False)

        tb = tk.Frame(self, bg=CLR_BG)
        tb.pack(fill="x", padx=12, pady=10)
        tk.Label(tb, text="Spare Parts Inventory", bg=CLR_BG, fg=CLR_TEXT, font=FONT_HEAD).pack(side="left", padx=4)
        self._toolbar_btn(tb, "Check In  ↓",   self._check_in,     CLR_OK,     "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "Check Out ↑",   self._check_out,    CLR_ACCENT2,"#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "+ Add Part",     self._add_part,     CLR_ACCENT, "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "Edit Part",      self._edit_part).pack(side="right", padx=4)
        self._toolbar_btn(tb, "Attachments",    self._attachments,  CLR_ACCENT2,"#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "Txn History",    self._view_txn).pack(side="right", padx=4)
        self._toolbar_btn(tb, "📋 Reorder List",self._show_reorder, CLR_WARN,   "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "⟳",             self.refresh).pack(side="right", padx=4)

        fb = tk.Frame(self, bg=CLR_BG)
        fb.pack(fill="x", padx=12, pady=(0,6))
        self._search_bar(fb, self._search_var, self.refresh).pack(side="left", padx=(0,12))
        tk.Checkbutton(fb, text="Low stock only", variable=self._show_low,
                       command=self.refresh, bg=CLR_BG, fg=CLR_MUTED,
                       selectcolor=CLR_BG, activebackground=CLR_BG,
                       font=FONT_SMALL).pack(side="left")

        self._tv = StyledTreeview(self,
            columns=("pn","desc","cat","loc","unit","qty","safety","reorder","cost","value","supplier","compat","status","files"),
            headings=("Part Number","Description","Category","Location","Unit",
                      "On Hand","Safety Stock","Reorder Qty","Unit Cost","Value","Supplier","Compatible Equip","Stock Status","Files"),
            col_widths={"pn":110,"desc":150,"cat":90,"loc":80,"unit":50,"qty":70,
                        "safety":80,"reorder":80,"cost":80,"value":90,"supplier":100,"compat":130,"status":90,"files":50},
            col_anchors={"unit":"center","qty":"center","safety":"center","reorder":"center",
                         "cost":"center","value":"center","status":"center","files":"center"})
        self._tv.pack(fill="both", expand=True, padx=12, pady=4)
        self._tv.bind_double_click(lambda e: self._edit_part())

        def _ctx(event):
            return self._build_context_menu([
                ("✏ Edit Part",        self._edit_part),
                ("📎 Attachments",     self._attachments),
                None,
                ("↓ Check In",         self._check_in),
                ("↑ Check Out",        self._check_out),
                None,
                ("📋 Txn History",     self._view_txn),
            ])
        self._tv.bind_context_menu(_ctx)
        self.refresh()

    def refresh(self):
        search   = self._search_var.get().lower()
        low_only = self._show_low.get()
        sql = "SELECT * FROM spare_parts WHERE is_active=1"
        if low_only: sql += " AND qty_on_hand < safety_stock"
        sql += " ORDER BY part_number"
        rows = self.db.fetchall(sql)
        curr = self.db.setting("currency", "USD")
        self._tv.delete(*self._tv.get_children())
        for r in rows:
            if search and search not in (r["part_number"]+r["description"]+r["category"]).lower():
                continue
            low    = r["qty_on_hand"] < r["safety_stock"]
            tag    = "critical" if low else "ok"
            status = "⚠ LOW STOCK" if low else "OK"
            value  = r["qty_on_hand"] * r["unit_cost"]
            att_c  = self.db.fetchone(
                "SELECT COUNT(*) c FROM attachments WHERE entity_type='part' AND entity_id=?",
                (r["id"],))["c"]
            compat = r["compatible_equip"] or "—"
            compat = r["compatible_equip"] or "—"
            self._tv.insert("", "end", iid=str(r["id"]),
                values=(r["part_number"], r["description"], r["category"],
                        r["location"], r["unit"], f"{r['qty_on_hand']:g}",
                        f"{r['safety_stock']:g}", f"{r['reorder_qty']:g}",
                        f"{curr} {r['unit_cost']:.2f}", f"{curr} {value:.2f}",
                        r["supplier"] or "—", compat, status,
                        f"📎 {att_c}" if att_c else "—"),
                tags=(tag,))
        self._tv.recolor_rows()

    def _part_fields(self, eq_choices=None):
        return [
            ("Part Number",      "part_number",     "text"),
            ("Description",      "description",     "text"),
            ("Category",         "category",        "text"),
            ("Location",         "location",        "text"),
            ("Unit",             "unit",            "combo", ["pcs","set","kg","ltr","m","box","roll"]),
            ("Qty on Hand",      "qty_on_hand",     "float"),
            ("Safety Stock",     "safety_stock",    "float"),
            ("Reorder Qty",      "reorder_qty",     "float"),
            ("Unit Cost",        "unit_cost",       "float"),
            ("Supplier",         "supplier",        "text"),
            ("Lead Time (days)", "lead_time_days",  "int"),
            ("Compatible Equip", "compatible_equip","equip_multi", eq_choices or []),
            ("Notes",            "notes",           "textarea"),
        ]

    def _get_eq_choices(self):
        equips = self.db.fetchall("SELECT code, name FROM equipment WHERE status='Active' ORDER BY code")
        return [f"{e['code']} — {e['name']}" for e in equips]

    def _add_part(self):
        dlg = FormDialog(self, "Add Spare Part", self._part_fields(self._get_eq_choices()))
        self.wait_window(dlg)
        if not dlg.result: return
        r = dlg.result
        for k in ["qty_on_hand","safety_stock","reorder_qty","unit_cost"]:
            try: r[k] = float(r[k] or 0)
            except: r[k] = 0.0
        try: r["lead_time_days"] = int(r["lead_time_days"] or 0)
        except: r["lead_time_days"] = 0
        try:
            self.db.execute(
                """INSERT INTO spare_parts(part_number,description,category,location,unit,
                   qty_on_hand,safety_stock,reorder_qty,unit_cost,supplier,
                   lead_time_days,compatible_equip,notes)
                   VALUES(:part_number,:description,:category,:location,:unit,
                   :qty_on_hand,:safety_stock,:reorder_qty,:unit_cost,:supplier,
                   :lead_time_days,:compatible_equip,:notes)""", r)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        self.refresh()

    def _edit_part(self):
        sel = self._tv.selection()
        if not sel: return
        part    = self.db.fetchone("SELECT * FROM spare_parts WHERE id=?", (int(sel[0]),))
        initial = dict(part)
        dlg = FormDialog(self, "Edit Spare Part",
                         self._part_fields(self._get_eq_choices()), initial=initial)
        self.wait_window(dlg)
        if not dlg.result: return
        r = dlg.result
        for k in ["qty_on_hand","safety_stock","reorder_qty","unit_cost"]:
            try: r[k] = float(r[k] or 0)
            except: r[k] = 0.0
        try: r["lead_time_days"] = int(r["lead_time_days"] or 0)
        except: r["lead_time_days"] = 0
        r["id"] = int(sel[0])
        self.db.execute(
            """UPDATE spare_parts SET part_number=:part_number,description=:description,
               category=:category,location=:location,unit=:unit,
               qty_on_hand=:qty_on_hand,safety_stock=:safety_stock,
               reorder_qty=:reorder_qty,unit_cost=:unit_cost,supplier=:supplier,
               lead_time_days=:lead_time_days,compatible_equip=:compatible_equip,
               notes=:notes WHERE id=:id""", r)
        self.refresh()

    def _txn_dialog(self, txn_type):
        sel = self._tv.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select a part first."); return
        pid  = int(sel[0])
        part = self.db.fetchone("SELECT * FROM spare_parts WHERE id=?", (pid,))
        default_user = self.db.setting("default_user", "")
        dlg = FormDialog(self, f"{txn_type} — {part['part_number']}", [
            ("Quantity",        "qty",         "float"),
            ("Reference Equip", "ref_equip",   "text"),
            ("Reference Task",  "ref_task",    "text"),
            ("Performed By",    "performed_by","text"),
            ("Remarks",         "remarks",     "text"),
        ], initial={"performed_by": default_user})
        self.wait_window(dlg)
        if not dlg.result: return
        try:
            qty = float(dlg.result["qty"] or 0)
        except ValueError:
            messagebox.showerror("Error", "Invalid quantity"); return
        if qty <= 0:
            messagebox.showerror("Error", "Quantity must be positive."); return

        qty_before = part["qty_on_hand"]
        if txn_type == "Check Out" and qty > qty_before:
            if not messagebox.askyesno("Below Zero", "Quantity exceeds stock. Proceed anyway?"): return
        qty_after = qty_before + qty if txn_type == "Check In" else qty_before - qty

        self.db.execute("UPDATE spare_parts SET qty_on_hand=? WHERE id=?", (qty_after, pid))
        self.db.execute(
            """INSERT INTO parts_transactions(part_id,txn_type,qty,qty_before,qty_after,
               ref_equipment,ref_task,performed_by,remarks)
               VALUES(?,?,?,?,?,?,?,?,?)""",
            (pid, txn_type, qty, qty_before, qty_after,
             dlg.result["ref_equip"], dlg.result["ref_task"],
             dlg.result["performed_by"], dlg.result["remarks"]))
        self.status_bar.set(
            f"{txn_type}: {qty:g} × {part['part_number']} → New qty: {qty_after:g}", CLR_OK)
        self.refresh()

    def _check_in(self):  self._txn_dialog("Check In")
    def _check_out(self): self._txn_dialog("Check Out")

    def _view_txn(self):
        sel = self._tv.selection()
        if not sel: return
        pid  = int(sel[0])
        part = self.db.fetchone("SELECT part_number, description FROM spare_parts WHERE id=?", (pid,))
        logs = self.db.fetchall(
            """SELECT txn_type,qty,qty_before,qty_after,ref_equipment,ref_task,
                      performed_by,remarks,txn_date
               FROM parts_transactions WHERE part_id=? ORDER BY txn_date DESC""", (pid,))

        win = tk.Toplevel(self)
        win.title(f"Transactions — {part['part_number']}")
        win.configure(bg=CLR_BG); win.geometry("800x420")
        tk.Label(win, text=f"{part['part_number']} · {part['description']}",
                 bg=CLR_BG, fg=CLR_TEXT, font=FONT_HEAD).pack(padx=16, pady=10, anchor="w")
        tv = StyledTreeview(win,
            columns=("date","type","qty","before","after","equip","task","by","remarks"),
            headings=("Date","Type","Qty","Before","After","Equipment","Task","By","Remarks"),
            col_widths={"date":130,"type":80,"qty":60,"before":70,"after":70,"equip":90,"task":90,"by":80,"remarks":160})
        tv.pack(fill="both", expand=True, padx=12, pady=8)
        for lg in logs:
            tag = "ok" if lg["txn_type"] == "Check In" else "warning"
            tv.insert("", "end",
                values=(lg["txn_date"], lg["txn_type"],
                        f"{lg['qty']:g}", f"{lg['qty_before']:g}", f"{lg['qty_after']:g}",
                        lg["ref_equipment"] or "—", lg["ref_task"] or "—",
                        lg["performed_by"] or "—", lg["remarks"] or "—"),
                tags=(tag,))

    def _attachments(self):
        sel = self._tv.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select a part first.")
            return
        part = self.db.fetchone("SELECT part_number, description FROM spare_parts WHERE id=?", (int(sel[0]),))
        AttachmentDialog(self, self.db, "part", int(sel[0]),
                         f"{part['part_number']} — {part['description']}")

    def _show_reorder(self):
        """Show the Reorder List window."""
        ReorderWindow(self, self.db)


# ─────────────────────────────────────────────────────────────────────────────
# Reorder List Window
# ─────────────────────────────────────────────────────────────────────────────
class ReorderWindow(tk.Toplevel):
    def __init__(self, parent, db: Database):
        super().__init__(parent)
        self.title("Parts Reorder List")
        self.configure(bg=CLR_BG)
        self.geometry("900x520")
        self._db = db
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=CLR_BG)
        hdr.pack(fill="x", padx=16, pady=10)
        tk.Label(hdr, text="Parts Reorder List", bg=CLR_BG, fg=CLR_TEXT, font=FONT_HEAD).pack(side="left")
        tk.Button(hdr, text="Export to Excel", command=self._export,
                  bg=CLR_OK, fg="#FFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2").pack(side="right")
        tk.Button(hdr, text="⟳ Refresh", command=self._load,
                  bg=CLR_SURFACE2, fg=CLR_MUTED, font=FONT_LABEL,
                  relief="flat", padx=10, pady=5, cursor="hand2").pack(side="right", padx=6)

        # info label
        tk.Label(self, text="Parts where current stock ≤ safety stock.  Reorder Qty = suggested purchase amount.",
                 bg=CLR_BG, fg=CLR_MUTED, font=FONT_SMALL).pack(anchor="w", padx=16, pady=(0,4))
        tk.Frame(self, bg=CLR_ACCENT, height=1).pack(fill="x", padx=16, pady=4)

        curr = self._db.setting("currency", "USD")
        self._tv = StyledTreeview(self,
            columns=("pn","desc","cat","loc","unit","qty","safety","reorder","cost","reorder_cost","supplier","lead","equip"),
            headings=("Part Number","Description","Category","Location","Unit",
                      "On Hand","Safety Stock","Reorder Qty",f"Unit Cost ({curr})",
                      f"Reorder Cost ({curr})","Supplier","Lead (days)","Compatible Equip"),
            col_widths={"pn":110,"desc":180,"cat":90,"loc":80,"unit":50,"qty":70,
                        "safety":90,"reorder":90,"cost":90,"reorder_cost":110,
                        "supplier":110,"lead":80,"equip":120})
        self._tv.pack(fill="both", expand=True, padx=12, pady=4)

        # summary bar
        self._summary_var = tk.StringVar()
        tk.Label(self, textvariable=self._summary_var,
                 bg=CLR_SURFACE, fg=CLR_TEXT, font=FONT_LABEL,
                 padx=16, pady=8, anchor="w").pack(fill="x", padx=12, pady=4)

        self._load()

    def _load(self):
        rows = self._db.fetchall("""
            SELECT * FROM spare_parts
            WHERE is_active=1 AND qty_on_hand <= safety_stock
            ORDER BY (qty_on_hand - safety_stock) ASC, supplier, part_number
        """)
        curr = self._db.setting("currency", "USD")
        self._tv.delete(*self._tv.get_children())
        total_cost = 0
        for r in rows:
            reorder_cost = r["reorder_qty"] * r["unit_cost"]
            total_cost  += reorder_cost
            self._tv.insert("", "end",
                values=(r["part_number"], r["description"], r["category"],
                        r["location"], r["unit"],
                        f"{r['qty_on_hand']:g}", f"{r['safety_stock']:g}",
                        f"{r['reorder_qty']:g}",
                        f"{r['unit_cost']:.2f}", f"{reorder_cost:.2f}",
                        r["supplier"] or "—", r["lead_time_days"],
                        r["compatible_equip"] or "—"),
                tags=("critical",))
        self._summary_var.set(
            f"  {len(rows)} parts need reordering  |  "
            f"Total Estimated Cost: {curr} {total_cost:,.2f}")

    def _export(self):
        EXPORT_DIR.mkdir(exist_ok=True)
        curr  = self._db.setting("currency", "USD")
        rows  = self._db.fetchall("""
            SELECT * FROM spare_parts
            WHERE is_active=1 AND qty_on_hand <= safety_stock
            ORDER BY supplier, part_number""")
        wb = openpyxl.Workbook()
        ws = wb.active; ws.title = "Reorder List"

        # Header
        ws["A1"] = "PARTS REORDER LIST"
        ws["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws["A3"] = f"Facility: {self._db.setting('facility_name', '')}"
        ws.append([])

        header_fill = PatternFill("solid", fgColor="1A2744")
        header_font = Font(bold=True, color="FFFFFF", size=10)
        headers = ["Part Number","Description","Category","Location","Unit",
                   "Qty on Hand","Safety Stock","Reorder Qty",
                   f"Unit Cost ({curr})",f"Reorder Cost ({curr})",
                   "Supplier","Lead Time (days)","Compatible Equipment","Notes"]
        for ci, h in enumerate(headers, 1):
            cell       = ws.cell(5, ci, h)
            cell.font  = header_font
            cell.fill  = header_fill
            ws.column_dimensions[get_column_letter(ci)].width = max(14, len(h)+2)

        # Group by supplier
        supplier_groups = {}
        for r in rows:
            sup = r["supplier"] or "Unknown Supplier"
            supplier_groups.setdefault(sup, []).append(r)

        row_idx = 6
        total_cost = 0
        for supplier, parts in sorted(supplier_groups.items()):
            # Supplier sub-header
            ws.cell(row_idx, 1, f"── {supplier} ──")
            ws.cell(row_idx, 1).font = Font(bold=True, color="00C2A8")
            row_idx += 1
            sup_cost = 0
            for r in parts:
                rc = r["reorder_qty"] * r["unit_cost"]
                sup_cost   += rc
                total_cost += rc
                ws.append([  # uses next row implicitly — use explicit row
                ])
                for ci, val in enumerate([
                    r["part_number"], r["description"], r["category"],
                    r["location"], r["unit"], r["qty_on_hand"],
                    r["safety_stock"], r["reorder_qty"],
                    r["unit_cost"], rc,
                    r["supplier"], r["lead_time_days"],
                    r["compatible_equip"], r["notes"]
                ], 1):
                    ws.cell(row_idx, ci, val)
                row_idx += 1
            # Sub-total
            ws.cell(row_idx, 9, "Sub-total:")
            ws.cell(row_idx, 10, sup_cost)
            ws.cell(row_idx, 9).font = Font(bold=True)
            ws.cell(row_idx, 10).font = Font(bold=True)
            row_idx += 2

        # Grand total
        ws.cell(row_idx, 9, "GRAND TOTAL:")
        ws.cell(row_idx, 10, total_cost)
        ws.cell(row_idx, 9).font = Font(bold=True, color="00C2A8")
        ws.cell(row_idx, 10).font = Font(bold=True, color="00C2A8")

        path = EXPORT_DIR / f"reorder_list_{date.today().strftime('%Y%m%d')}.xlsx"
        wb.save(path)
        messagebox.showinfo("Exported", f"Reorder list saved to:\n{path}")


# ─────────────────────────────────────────────────────────────────────────────
# PANEL 5 — Export Centre
# ─────────────────────────────────────────────────────────────────────────────
class ExportPanel(Panel):
    def build(self):
        tk.Label(self, text="Export Centre", bg=CLR_BG, fg=CLR_TEXT, font=FONT_HEAD).pack(anchor="w", padx=20, pady=14)
        tk.Frame(self, bg=CLR_ACCENT, height=2).pack(fill="x", padx=20, pady=4)
        tk.Label(self, text="Generate Excel workbooks from your maintenance and inventory data.",
                 bg=CLR_BG, fg=CLR_MUTED, font=FONT_BODY, wraplength=600, justify="left"
                 ).pack(anchor="w", padx=20, pady=6)

        for title, subtitle, cmd, color in [
            ("All Maintenance History",  "Full log of every completed maintenance activity.",               self._export_all_maintenance, CLR_ACCENT),
            ("Overdue & Upcoming Tasks", "Snapshot of current maintenance status.",                        self._export_due_tasks,       CLR_WARN),
            ("Spare Parts Inventory",    "Current inventory levels, values, and supplier info.",           self._export_inventory,       CLR_ACCENT2),
            ("Parts Transaction Log",    "Full check-in / check-out history for all spare parts.",         self._export_txn_log,         CLR_ACCENT2),
            ("Parts Reorder List",       "Summary of all parts that need to be purchased soon.",           self._export_reorder,         CLR_WARN),
            ("Full System Export",       "All of the above in one workbook with separate sheets.",          self._export_full,            CLR_OK),
        ]:
            card = tk.Frame(self, bg=CLR_SURFACE, padx=20, pady=14)
            card.pack(fill="x", padx=20, pady=5)
            left = tk.Frame(card, bg=CLR_SURFACE)
            left.pack(side="left", fill="both", expand=True)
            tk.Label(left, text=title, bg=CLR_SURFACE, fg=CLR_TEXT, font=FONT_LABEL).pack(anchor="w")
            tk.Label(left, text=subtitle, bg=CLR_SURFACE, fg=CLR_MUTED, font=FONT_SMALL,
                     wraplength=500, justify="left").pack(anchor="w", pady=2)
            tk.Button(card, text="Export ↗", command=cmd,
                      bg=color, fg="#FFF", font=FONT_LABEL,
                      relief="flat", padx=14, pady=6, cursor="hand2").pack(side="right")

    def _stamp(self):
        return date.today().strftime("%Y%m%d")

    def _style_header(self, ws, headers, row=1):
        fill = PatternFill("solid", fgColor="1A2744")
        font = Font(bold=True, color="FFFFFF", size=10)
        for col, hdr in enumerate(headers, 1):
            c = ws.cell(row, col, hdr); c.fill = fill; c.font = font
            ws.column_dimensions[get_column_letter(col)].width = max(14, len(hdr)+4)

    def _export_all_maintenance(self):
        EXPORT_DIR.mkdir(exist_ok=True)
        rows = self.db.fetchall("""
            SELECT e.code,e.name,mt.task_name,mt.frequency_days,
                   ml.done_date,ml.done_by,ml.duration_min,ml.findings,ml.next_due,ml.logged_at
            FROM maintenance_log ml
            JOIN maintenance_tasks mt ON mt.id=ml.task_id
            JOIN equipment e ON e.id=ml.equipment_id
            ORDER BY ml.done_date DESC""")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Maintenance History"
        self._style_header(ws, ["Equipment Code","Equipment Name","Task","Freq(days)","Done Date","Done By","Duration(min)","Findings","Next Due","Logged At"])
        for r in rows: ws.append(list(r))
        path = EXPORT_DIR / f"maintenance_history_{self._stamp()}.xlsx"
        wb.save(path); messagebox.showinfo("Exported", f"Saved:\n{path}")

    def _export_due_tasks(self):
        EXPORT_DIR.mkdir(exist_ok=True)
        rows = self.db.fetchall("""
            SELECT e.code,e.name,e.location,mt.task_name,mt.frequency_days,
                   mt.last_done,mt.next_due,mt.assigned_to,
                   julianday(mt.next_due)-julianday('now') AS days_left
            FROM maintenance_tasks mt
            JOIN equipment e ON e.id=mt.equipment_id
            WHERE mt.is_active=1 ORDER BY mt.next_due ASC""")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Task Status"
        self._style_header(ws, ["Code","Equipment","Location","Task","Freq(d)","Last Done","Next Due","Assigned To","Days Left","Status"])
        for r in rows:
            d = r["days_left"]
            s = "Not Set" if d is None else ("OVERDUE" if d < 0 else ("DUE TODAY" if d == 0 else ("DUE SOON" if d <= 7 else "OK")))
            ws.append([r["code"],r["name"],r["location"],r["task_name"],r["frequency_days"],
                       r["last_done"],r["next_due"],r["assigned_to"],round(d,1) if d else None,s])
        path = EXPORT_DIR / f"task_status_{self._stamp()}.xlsx"
        wb.save(path); messagebox.showinfo("Exported", f"Saved:\n{path}")

    def _export_inventory(self):
        EXPORT_DIR.mkdir(exist_ok=True)
        curr = self.db.setting("currency","USD")
        rows = self.db.fetchall("SELECT * FROM spare_parts WHERE is_active=1 ORDER BY part_number")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Parts Inventory"
        self._style_header(ws, ["Part Number","Description","Category","Location","Unit","Qty on Hand","Safety Stock","Reorder Qty",f"Unit Cost ({curr})",f"Total Value ({curr})","Supplier","Lead Time(days)","Compatible Equip","Status"])
        for r in rows:
            s = "LOW STOCK" if r["qty_on_hand"] < r["safety_stock"] else "OK"
            ws.append([r["part_number"],r["description"],r["category"],r["location"],r["unit"],
                       r["qty_on_hand"],r["safety_stock"],r["reorder_qty"],r["unit_cost"],
                       r["qty_on_hand"]*r["unit_cost"],r["supplier"],r["lead_time_days"],r["compatible_equip"],s])
        path = EXPORT_DIR / f"inventory_{self._stamp()}.xlsx"
        wb.save(path); messagebox.showinfo("Exported", f"Saved:\n{path}")

    def _export_txn_log(self):
        EXPORT_DIR.mkdir(exist_ok=True)
        rows = self.db.fetchall("""
            SELECT sp.part_number,sp.description,pt.txn_type,pt.qty,pt.qty_before,pt.qty_after,
                   pt.ref_equipment,pt.ref_task,pt.performed_by,pt.remarks,pt.txn_date
            FROM parts_transactions pt JOIN spare_parts sp ON sp.id=pt.part_id
            ORDER BY pt.txn_date DESC""")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Transactions"
        self._style_header(ws, ["Part Number","Description","Type","Qty","Before","After","Equipment","Task","Performed By","Remarks","Date"])
        for r in rows: ws.append(list(r))
        path = EXPORT_DIR / f"parts_transactions_{self._stamp()}.xlsx"
        wb.save(path); messagebox.showinfo("Exported", f"Saved:\n{path}")

    def _export_reorder(self):
        """Reuse the ReorderWindow export logic."""
        rw = ReorderWindow.__new__(ReorderWindow)
        rw._db = self.db
        rw._export = lambda: None  # stub
        # Call the actual export directly
        EXPORT_DIR.mkdir(exist_ok=True)
        curr = self.db.setting("currency","USD")
        rows = self.db.fetchall("""
            SELECT * FROM spare_parts
            WHERE is_active=1 AND qty_on_hand <= safety_stock
            ORDER BY supplier, part_number""")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Reorder List"
        ws["A1"] = "PARTS REORDER LIST"
        ws["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws.append([])
        self._style_header(ws, ["Part Number","Description","Category","Location","Unit",
                                 "Qty on Hand","Safety Stock","Reorder Qty",f"Unit Cost ({curr})",
                                 f"Reorder Cost ({curr})","Supplier","Lead Time (days)","Compatible Equipment"], row=4)
        total = 0
        for r in rows:
            rc = r["reorder_qty"] * r["unit_cost"]; total += rc
            ws.append([r["part_number"],r["description"],r["category"],r["location"],r["unit"],
                       r["qty_on_hand"],r["safety_stock"],r["reorder_qty"],r["unit_cost"],rc,
                       r["supplier"],r["lead_time_days"],r["compatible_equip"]])
        ws.append([]); ws.append(["","","","","","","","","Grand Total:",total])
        path = EXPORT_DIR / f"reorder_list_{self._stamp()}.xlsx"
        wb.save(path); messagebox.showinfo("Exported", f"Saved:\n{path}")

    def _export_full(self):
        EXPORT_DIR.mkdir(exist_ok=True)
        wb   = openpyxl.Workbook()
        curr = self.db.setting("currency","USD")

        ws0 = wb.active; ws0.title = "Summary"
        ws0["A1"] = "Equipment Management System — Full Export"
        ws0["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws0["A4"] = "Metric"; ws0["B4"] = "Value"
        ws0["A5"] = "Total Equipment"
        ws0["B5"] = self.db.fetchone("SELECT COUNT(*) c FROM equipment WHERE status='Active'")["c"]
        ws0["A6"] = "Overdue Tasks"
        ws0["B6"] = self.db.fetchone("SELECT COUNT(*) c FROM maintenance_tasks WHERE is_active=1 AND next_due < date('now')")["c"]
        ws0["A7"] = "Parts Below Safety Stock"
        ws0["B7"] = self.db.fetchone("SELECT COUNT(*) c FROM spare_parts WHERE is_active=1 AND qty_on_hand < safety_stock")["c"]
        tv = self.db.fetchone("SELECT SUM(qty_on_hand*unit_cost) v FROM spare_parts WHERE is_active=1")["v"] or 0
        ws0["A8"] = f"Total Inventory Value ({curr})"; ws0["B8"] = round(tv, 2)

        ws1 = wb.create_sheet("Equipment")
        self._style_header(ws1, ["Code","Name","Category","Location","Model","Serial No.","Status","Notes"])
        for r in self.db.fetchall("SELECT * FROM equipment ORDER BY code"):
            ws1.append([r["code"],r["name"],r["category"],r["location"],r["model"],r["serial_no"],r["status"],r["notes"]])

        ws2 = wb.create_sheet("Task Status")
        self._style_header(ws2, ["Code","Equipment","Task","Freq(d)","Last Done","Next Due","Assigned To","Status"])
        for r in self.db.fetchall("SELECT e.code,e.name,mt.task_name,mt.frequency_days,mt.last_done,mt.next_due,mt.assigned_to FROM maintenance_tasks mt JOIN equipment e ON e.id=mt.equipment_id WHERE mt.is_active=1 ORDER BY mt.next_due ASC"):
            d = days_until(r["next_due"])
            s = "Not Set" if d is None else ("OVERDUE" if d < 0 else ("DUE TODAY" if d == 0 else ("DUE SOON" if d <= 7 else "OK")))
            ws2.append([r["code"],r["name"],r["task_name"],r["frequency_days"],r["last_done"],r["next_due"],r["assigned_to"],s])

        ws3 = wb.create_sheet("Maintenance Log")
        self._style_header(ws3, ["Code","Equipment","Task","Done Date","Done By","Duration(min)","Findings","Next Due","Parts Used"])
        for r in self.db.fetchall("SELECT ml.id,e.code,e.name,mt.task_name,ml.done_date,ml.done_by,ml.duration_min,ml.findings,ml.next_due FROM maintenance_log ml JOIN maintenance_tasks mt ON mt.id=ml.task_id JOIN equipment e ON e.id=ml.equipment_id ORDER BY ml.done_date DESC"):
            parts_used = self.db.fetchall("SELECT sp.part_number,mlp.qty_used FROM maintenance_log_parts mlp JOIN spare_parts sp ON sp.id=mlp.part_id WHERE mlp.log_id=?", (r["id"],))
            pstr = "; ".join(f"{p['part_number']}×{p['qty_used']:g}" for p in parts_used) or ""
            ws3.append([r["code"],r["name"],r["task_name"],r["done_date"],r["done_by"],r["duration_min"],r["findings"],r["next_due"],pstr])

        ws4 = wb.create_sheet("Parts Inventory")
        self._style_header(ws4, ["Part Number","Description","Category","Location","Unit","Qty","Safety Stock","Reorder Qty",f"Unit Cost({curr})",f"Value({curr})","Supplier","Status"])
        for r in self.db.fetchall("SELECT * FROM spare_parts WHERE is_active=1 ORDER BY part_number"):
            s = "LOW" if r["qty_on_hand"] < r["safety_stock"] else "OK"
            ws4.append([r["part_number"],r["description"],r["category"],r["location"],r["unit"],
                        r["qty_on_hand"],r["safety_stock"],r["reorder_qty"],r["unit_cost"],
                        r["qty_on_hand"]*r["unit_cost"],r["supplier"],s])

        ws5 = wb.create_sheet("Reorder List")
        self._style_header(ws5, ["Part Number","Description","Category","Unit","Qty on Hand","Safety Stock","Reorder Qty",f"Reorder Cost({curr})","Supplier","Lead Time(days)"])
        for r in self.db.fetchall("SELECT * FROM spare_parts WHERE is_active=1 AND qty_on_hand<=safety_stock ORDER BY supplier,part_number"):
            ws5.append([r["part_number"],r["description"],r["category"],r["unit"],
                        r["qty_on_hand"],r["safety_stock"],r["reorder_qty"],
                        r["reorder_qty"]*r["unit_cost"],r["supplier"],r["lead_time_days"]])

        ws6 = wb.create_sheet("Parts Transactions")
        self._style_header(ws6, ["Part Number","Description","Type","Qty","Before","After","Equipment","Task","By","Remarks","Date"])
        for r in self.db.fetchall("SELECT sp.part_number,sp.description,pt.txn_type,pt.qty,pt.qty_before,pt.qty_after,pt.ref_equipment,pt.ref_task,pt.performed_by,pt.remarks,pt.txn_date FROM parts_transactions pt JOIN spare_parts sp ON sp.id=pt.part_id ORDER BY pt.txn_date DESC"):
            ws6.append(list(r))

        path = EXPORT_DIR / f"full_export_{self._stamp()}.xlsx"
        wb.save(path)
        messagebox.showinfo("Full Export Complete", f"Saved:\n{path}")


# ─────────────────────────────────────────────────────────────────────────────
# PANEL 6 — Settings
# ─────────────────────────────────────────────────────────────────────────────
class SettingsPanel(Panel):
    def build(self):
        tk.Label(self, text="Settings", bg=CLR_BG, fg=CLR_TEXT, font=FONT_HEAD).pack(anchor="w", padx=20, pady=14)
        tk.Frame(self, bg=CLR_ACCENT, height=2).pack(fill="x", padx=20, pady=4)

        cf = tk.Frame(self, bg=CLR_SURFACE, padx=20, pady=16)
        cf.pack(fill="x", padx=20, pady=8)
        tk.Label(cf, text="Configuration File", bg=CLR_SURFACE, fg=CLR_TEXT, font=FONT_LABEL).pack(anchor="w")
        tk.Label(cf, text=str(CFG_PATH), bg=CLR_SURFACE, fg=CLR_MUTED, font=FONT_MONO).pack(anchor="w", pady=4)
        btn_row = tk.Frame(cf, bg=CLR_SURFACE)
        btn_row.pack(fill="x", pady=4)
        self._toolbar_btn(btn_row, "Reload Config from Excel", self._reload_config, CLR_ACCENT, "#FFF").pack(side="left", padx=(0,8))
        self._toolbar_btn(btn_row, "Create / Reset Template",  self._create_template, CLR_WARN, "#FFF").pack(side="left", padx=(0,8))
        self._toolbar_btn(btn_row, "Open Export Folder",       self._open_export_folder).pack(side="left", padx=(0,8))
        self._toolbar_btn(btn_row, "Open Attachments Folder",  self._open_attachments_folder).pack(side="left")

        sf = tk.Frame(self, bg=CLR_SURFACE, padx=20, pady=16)
        sf.pack(fill="x", padx=20, pady=8)
        tk.Label(sf, text="Application Settings", bg=CLR_SURFACE, fg=CLR_TEXT, font=FONT_LABEL).pack(anchor="w", pady=(0,8))
        self._setting_fields = [
            ("Facility Name",       "facility_name"),
            ("Alert Window (days)", "alert_soon_days"),
            ("Default User",        "default_user"),
            ("Currency",            "currency"),
        ]
        self._setting_vars = {}
        for label, key in self._setting_fields:
            row = tk.Frame(sf, bg=CLR_SURFACE)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label, bg=CLR_SURFACE, fg=CLR_MUTED,
                     font=FONT_LABEL, width=24, anchor="w").pack(side="left")
            var = tk.StringVar(value=self.db.setting(key, ""))
            tk.Entry(row, textvariable=var, bg=CLR_BG, fg=CLR_TEXT,
                     insertbackground=CLR_TEXT, font=FONT_BODY,
                     relief="flat", bd=4, width=36).pack(side="left")
            self._setting_vars[key] = var
        tk.Button(sf, text="Save Settings", command=self._save_settings,
                  bg=CLR_OK, fg="#FFF", font=FONT_LABEL,
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(anchor="e", pady=8)

        df = tk.Frame(self, bg=CLR_SURFACE, padx=20, pady=12)
        df.pack(fill="x", padx=20, pady=8)
        tk.Label(df, text="Database", bg=CLR_SURFACE, fg=CLR_TEXT, font=FONT_LABEL).pack(anchor="w")
        tk.Label(df, text=str(DB_PATH), bg=CLR_SURFACE, fg=CLR_MUTED, font=FONT_MONO).pack(anchor="w", pady=4)
        tk.Label(df, text="WAL mode enabled — safe for concurrent access.",
                 bg=CLR_SURFACE, fg=CLR_OK, font=FONT_SMALL).pack(anchor="w")

        if not PIL_AVAILABLE:
            tk.Label(self, text="⚠ Pillow not installed — image previews unavailable. Run: pip install pillow",
                     bg=CLR_BG, fg=CLR_WARN, font=FONT_SMALL).pack(anchor="w", padx=20, pady=4)

    def refresh(self):
        for key, var in self._setting_vars.items():
            var.set(self.db.setting(key, ""))

    def _reload_config(self):
        loader = ConfigLoader(self.db, CFG_PATH)
        msg = loader.load()
        messagebox.showinfo("Config Loaded", msg)
        self.status_bar.set("Config reloaded from Excel.", CLR_OK)

    def _create_template(self):
        if CFG_PATH.exists():
            if not messagebox.askyesno("Overwrite?", "This will OVERWRITE the existing config file. Continue?"): return
        ConfigLoader(self.db, CFG_PATH)._create_template()
        messagebox.showinfo("Template Created", f"Template saved to:\n{CFG_PATH}\n\nFill it in and reload config.")

    def _open_export_folder(self):
        EXPORT_DIR.mkdir(exist_ok=True)
        open_file_externally(EXPORT_DIR)

    def _open_attachments_folder(self):
        ATTACHMENTS_DIR.mkdir(exist_ok=True)
        open_file_externally(ATTACHMENTS_DIR)

    def _save_settings(self):
        for key, var in self._setting_vars.items():
            self.db.set_setting(key, var.get())
        self.status_bar.set("Settings saved.", CLR_OK)
        messagebox.showinfo("Saved", "Settings saved successfully.")

# ─────────────────────────────────────────────────────────────────────────────
# PANEL 7 — Documents Library
# ─────────────────────────────────────────────────────────────────────────────
class DocumentsPanel(Panel):
    """General-purpose attachment library — user manuals, SOPs, misc files.
    Uses entity_type='general', entity_id=0 so no equipment/part link is needed,
    but an optional 'tag' is stored in the description field as a prefix: [TAG] text."""

    def build(self):
        self._search_var = tk.StringVar()

        tb = tk.Frame(self, bg=CLR_BG)
        tb.pack(fill="x", padx=12, pady=10)
        tk.Label(tb, text="Documents Library", bg=CLR_BG, fg=CLR_TEXT,
                 font=FONT_HEAD).pack(side="left", padx=4)
        self._toolbar_btn(tb, "+ Upload File",   self._upload,  CLR_ACCENT,  "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "Preview / Open",  self._preview, CLR_ACCENT2, "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "Delete",          self._delete,  CLR_DANGER,  "#FFF").pack(side="right", padx=4)
        self._toolbar_btn(tb, "⟳",              self.refresh).pack(side="right", padx=4)

        fb = tk.Frame(self, bg=CLR_BG)
        fb.pack(fill="x", padx=12, pady=(0, 6))
        self._search_bar(fb, self._search_var, self.refresh).pack(side="left")
        tk.Label(fb, text="Search by filename, tag, or description",
                 bg=CLR_BG, fg=CLR_MUTED, font=FONT_SMALL).pack(side="left", padx=8)

        self._tv = StyledTreeview(self,
            columns=("icon", "filename", "tag", "desc", "size", "uploaded"),
            headings=("", "Filename", "Tag / Category", "Description", "Size", "Uploaded"),
            col_widths={"icon": 28, "filename": 220, "tag": 120,
                        "desc": 260, "size": 70, "uploaded": 130},
            col_anchors={"icon": "center", "size": "center"})
        self._tv.pack(fill="both", expand=True, padx=12, pady=4)
        self._tv.bind_double_click(lambda e: self._preview())

        def _ctx(event):
            return self._build_context_menu([
                ("🔍 Preview / Open", self._preview),
                None,
                ("✏ Edit Description", self._edit_desc),
                None,
                ("🗑 Delete",          self._delete),
            ])
        self._tv.bind_context_menu(_ctx)
        self.refresh()

    def refresh(self):
        search = self._search_var.get().lower()
        rows = self.db.fetchall(
            """SELECT * FROM attachments
               WHERE entity_type='general' AND entity_id=0
               ORDER BY uploaded_at DESC""")
        self._tv.delete(*self._tv.get_children())
        for r in rows:
            desc = r["description"] or ""
            # parse optional [TAG] prefix
            tag = ""
            display_desc = desc
            if desc.startswith("[") and "]" in desc:
                tag = desc[1:desc.index("]")]
                display_desc = desc[desc.index("]")+1:].strip()

            if search and search not in (r["filename"] + tag + display_desc).lower():
                continue

            icon = "🖼" if r["file_type"] == "image" else "📄"
            size_str = f"{r['file_size']//1024} KB" if r["file_size"] else "—"
            self._tv.insert("", "end", iid=str(r["id"]),
                values=(icon, r["filename"], tag, display_desc,
                        size_str, r["uploaded_at"][:16]),
                tags=("even",))
        self._tv.recolor_rows()

    def _upload(self):
        ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)
        paths = filedialog.askopenfilenames(
            title="Select files to upload",
            filetypes=[
                ("Images",    "*.jpg *.jpeg *.png *.gif *.bmp *.webp"),
                ("Documents", "*.pdf *.docx *.xlsx *.txt *.csv *.doc *.xls"),
                ("All files", "*.*"),
            ])
        for src in paths:
            src = Path(src)
            suffix  = src.suffix.lower()
            ftype   = "image" if suffix in IMAGE_EXTS else "document"
            ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
            stored  = f"general_0_{ts}_{src.name}"
            dest    = ATTACHMENTS_DIR / stored
            shutil.copy2(src, dest)
            fsize   = dest.stat().st_size

            tag  = simpledialog.askstring(
                "Tag / Category",
                f"Optional tag for:\n{src.name}\n\n"
                "Examples: User Manual, SOP, Drawing, Certificate\n"
                "(leave blank to skip)",
                parent=self) or ""
            desc_text = simpledialog.askstring(
                "Description",
                f"Optional description for:\n{src.name}",
                parent=self) or ""

            # store as "[TAG] description" so both survive in one field
            combined = f"[{tag}] {desc_text}".strip() if tag else desc_text
            self.db.execute(
                """INSERT INTO attachments
                   (entity_type, entity_id, filename, stored_name,
                    file_type, file_size, description)
                   VALUES('general', 0, ?, ?, ?, ?, ?)""",
                (src.name, stored, ftype, fsize, combined))
        self.refresh()

    def _selected_row(self):
        sel = self._tv.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select a file first.")
            return None
        return self.db.fetchone("SELECT * FROM attachments WHERE id=?", (int(sel[0]),))

    def _preview(self):
        row = self._selected_row()
        if not row: return
        dest = ATTACHMENTS_DIR / row["stored_name"]
        if not dest.exists():
            messagebox.showerror("File Missing", f"File not found:\n{dest}")
            return
        if row["file_type"] == "image" and PIL_AVAILABLE:
            self._show_image_preview(dest, row["filename"])
        else:
            open_file_externally(dest)

    def _show_image_preview(self, path, filename):
        win = tk.Toplevel(self)
        win.title(f"Preview — {filename}")
        win.configure(bg=CLR_BG)
        win.geometry("800x600")
        try:
            img = Image.open(path)
            img.thumbnail((760, 540), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(win, image=photo, bg=CLR_BG)
            lbl.image = photo
            lbl.pack(expand=True, pady=10)
        except Exception as e:
            tk.Label(win, text=f"Cannot preview:\n{e}",
                     bg=CLR_BG, fg=CLR_DANGER, font=FONT_BODY).pack(expand=True)
        tk.Button(win, text="Open Externally",
                  command=lambda: open_file_externally(path),
                  bg=CLR_ACCENT2, fg="#FFF", font=FONT_LABEL,
                  relief="flat", padx=12, pady=5, cursor="hand2").pack(pady=8)

    def _edit_desc(self):
        row = self._selected_row()
        if not row: return
        desc = row["description"] or ""
        tag, display_desc = "", desc
        if desc.startswith("[") and "]" in desc:
            tag = desc[1:desc.index("]")]
            display_desc = desc[desc.index("]")+1:].strip()

        new_tag = simpledialog.askstring("Tag / Category", "Tag:", initialvalue=tag, parent=self)
        if new_tag is None: return   # cancelled
        new_desc = simpledialog.askstring("Description", "Description:", initialvalue=display_desc, parent=self)
        if new_desc is None: return

        combined = f"[{new_tag}] {new_desc}".strip() if new_tag else new_desc
        self.db.execute("UPDATE attachments SET description=? WHERE id=?",
                        (combined, row["id"]))
        self.refresh()

    def _delete(self):
        row = self._selected_row()
        if not row: return
        if not messagebox.askyesno("Delete", f"Delete '{row['filename']}'?\nThis cannot be undone."):
            return
        dest = ATTACHMENTS_DIR / row["stored_name"]
        try:
            dest.unlink(missing_ok=True)
        except Exception:
            pass
        self.db.execute("DELETE FROM attachments WHERE id=?", (row["id"],))
        self.refresh()

# ─────────────────────────────────────────────────────────────────────────────
# Main Application
# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Equipment Management System")
        self.geometry("1400x820")
        self.minsize(960, 640)
        self.configure(bg=CLR_BG)
        self._setup_style()
        self.db = Database(DB_PATH)

        if not CFG_PATH.exists():
            ConfigLoader(self.db, CFG_PATH).load()
        elif self.db.fetchone("SELECT COUNT(*) c FROM equipment")["c"] == 0:
            ConfigLoader(self.db, CFG_PATH).load()

        ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)
        self._build_ui()
        self._schedule_alerts()

    def _setup_style(self):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TScrollbar",
                        background=CLR_SURFACE2, troughcolor=CLR_BG,
                        arrowcolor=CLR_MUTED, relief="flat")
        style.configure("TCombobox",
                        fieldbackground=CLR_BG, background=CLR_BG,
                        foreground=CLR_TEXT, selectbackground=CLR_ACCENT2,
                        arrowcolor=CLR_MUTED)

    def _build_ui(self):
        sidebar = tk.Frame(self, bg=CLR_SURFACE, width=190)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        logo = tk.Frame(sidebar, bg=CLR_SURFACE, pady=20)
        logo.pack(fill="x")
        tk.Label(logo, text="⚙", bg=CLR_SURFACE, fg=CLR_ACCENT, font=("Segoe UI", 28)).pack()
        tk.Label(logo, text="Equipment", bg=CLR_SURFACE, fg=CLR_TEXT, font=("Segoe UI", 14, "bold")).pack()
        tk.Label(logo, text="Manager", bg=CLR_SURFACE, fg=CLR_TEXT, font=("Segoe UI", 14, "bold")).pack()
        # tk.Label(logo, text="v6", bg=CLR_SURFACE, fg=CLR_MUTED, font=("Segoe UI", 8)).pack()
        tk.Frame(sidebar, bg=CLR_BORDER, height=1).pack(fill="x", padx=12, pady=8)

        self._status = StatusBar(self)
        self._status.pack(side="bottom", fill="x")

        content = tk.Frame(self, bg=CLR_BG)
        content.pack(side="left", fill="both", expand=True)

        self._panels = {}
        nav_items = [
            ("Dashboard",   "🏠", DashboardPanel),
            ("Maintenance", "🔧", MaintenancePanel),
            ("Equipment",   "🏭", EquipmentPanel),
            ("Spare Parts", "📦", PartsPanel),
            ("Export",      "📤", ExportPanel),
            ("Documents",   "📁", DocumentsPanel),
            ("Settings",    "⚙",  SettingsPanel),
        ]

        for name, icon, PanelClass in nav_items:
            panel = PanelClass(content, self.db, self._status)
            panel.place(relx=0, rely=0, relwidth=1, relheight=1)
            self._panels[name] = panel

            def cmd(n=name): self._switch(n)
            btn = tk.Button(sidebar, text=f"  {icon}  {name}",
                            command=cmd, bg=CLR_SURFACE, fg=CLR_MUTED,
                            font=FONT_BODY, relief="flat", padx=8, pady=10,
                            anchor="w", cursor="hand2",
                            activebackground=CLR_SURFACE2, activeforeground=CLR_TEXT)
            btn.pack(fill="x", padx=4)
            btn._name = name
            self._panels[name]._nav_btn = btn

        self._switch("Dashboard")

    def _switch(self, name):
        for pname, panel in self._panels.items():
            panel.lower()
            if hasattr(panel, "_nav_btn"):
                panel._nav_btn.config(bg=CLR_SURFACE, fg=CLR_MUTED)
        panel = self._panels[name]
        panel.lift()
        if hasattr(panel, "_nav_btn"):
            panel._nav_btn.config(bg=CLR_ACCENT2, fg="#FFFFFF")
        panel.refresh()
        self._status.set(f"Viewing: {name}")

    def _schedule_alerts(self):
        try:
            today   = date.today().isoformat()
            overdue = self.db.fetchone("SELECT COUNT(*) c FROM maintenance_tasks WHERE is_active=1 AND next_due < ?", (today,))["c"]
            low     = self.db.fetchone("SELECT COUNT(*) c FROM spare_parts WHERE is_active=1 AND qty_on_hand < safety_stock")["c"]
            alerts  = []
            if overdue: alerts.append(f"{overdue} overdue tasks")
            if low:     alerts.append(f"{low} low-stock parts")
            if alerts:
                self.title(f"⚠ ALERT — {', '.join(alerts)}")
                self._status.set(f"⚠ ALERT: {', '.join(alerts)}", CLR_DANGER)
            else:
                self.title("Equipment Management System")
        except Exception:
            pass
        self.after(60_000, self._schedule_alerts)


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    _ico = _resource("equipment.ico")
    if pathlib.Path(_ico).exists():
        app.iconbitmap(_ico)
    app.mainloop()