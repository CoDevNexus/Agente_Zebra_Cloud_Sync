from __future__ import annotations

import hashlib
import json
import logging
import os
import queue
import sqlite3
import sys
import threading
import time
from copy import deepcopy
from datetime import datetime, timedelta
from pathlib import Path
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Tuple

import customtkinter as ctk
import gspread
import mysql.connector
import pystray
import serial
import tkinter as tk
from mysql.connector import Error as MySQLError
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import Workbook
from PIL import Image, ImageDraw
from pystray import MenuItem as TrayItem
from serial.tools import list_ports
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from pynput import keyboard as pynput_keyboard
    _PYNPUT_AVAILABLE = True
except Exception:
    _PYNPUT_AVAILABLE = False


APP_NAME = "Agente Zebra Cloud Sync"
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config_agente.json"
LOG_PATH = BASE_DIR / "agente_zebra.log"
DEFAULT_SQLITE_PATH = BASE_DIR / "agente_buffer.db"
MYSQL_QUERY = """
SELECT TRIM(codigo) as codigo,
       TRIM(nombre_producto) as descripcion,
       SUM(stock) AS total_stock
FROM productos
WHERE TRIM(codigo) = %s
GROUP BY codigo, nombre_producto
LIMIT 1;
""".strip()
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(threadName)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(APP_NAME)

if TYPE_CHECKING:
    from pystray._base import Icon as PystrayIcon
else:
    PystrayIcon = Any


def hash_password(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


class ConfigManager:
    DEFAULTS: Dict[str, Any] = {
        "com_port": "",
        "mysql_host": "127.0.0.1",
        "mysql_user": "",
        "mysql_password": "",
        "mysql_database": "",
        "google_sheet_id": "",
        "google_credentials": "",
        "sqlite_db": str(DEFAULT_SQLITE_PATH),
        "sync_interval": 15,
        "scanner_baudrate": 9600,
        "correction_window_seconds": 60,
        "correction_repeat_count": 3,
        "history_page_size": 50,
        "config_password_hash": "",
        "start_with_windows": False,
        "configured_once": False,
        "scanner_mode": "hid",
        "hid_inter_char_ms": 150,
    }

    def __init__(self, path: Path):
        self.path = Path(path)
        self._lock = threading.RLock()
        self._data = deepcopy(self.DEFAULTS)
        self.load()

    def load(self) -> None:
        with self._lock:
            if not self.path.exists():
                self.save()
                return
            try:
                with self.path.open("r", encoding="utf-8") as fh:
                    data = json.load(fh)
                merged = deepcopy(self.DEFAULTS)
                if isinstance(data, dict):
                    merged.update(data)
                self._data = merged
            except Exception:
                logger.exception("No se pudo leer config_agente.json. Se usarán valores por defecto.")
                self._data = deepcopy(self.DEFAULTS)
                self.save()

    def save(self) -> None:
        with self._lock:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with self.path.open("w", encoding="utf-8") as fh:
                json.dump(self._data, fh, indent=2, ensure_ascii=False)

    def get(self) -> Dict[str, Any]:
        with self._lock:
            return deepcopy(self._data)

    def update(self, values: Dict[str, Any]) -> Dict[str, Any]:
        with self._lock:
            self._data.update(values)
            self.save()
            return deepcopy(self._data)


class StartupManager:
    """Gestiona el autoinicio usando el Registro de Windows (más confiable que .bat en Startup)."""

    _REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"

    def __init__(self, app_name: str):
        self.app_name = app_name
        # Mantener compatibilidad: si existía un .bat antiguo, saber dónde está para limpiarlo
        appdata = os.getenv("APPDATA", "")
        self.startup_dir = Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
        safe_name = "".join(ch if ch.isalnum() else "_" for ch in app_name)
        self.script_path = self.startup_dir / f"{safe_name}_startup.bat"

    def _reg_command(self) -> str:
        """Devuelve el comando que se registra en el Registro."""
        if getattr(sys, "frozen", False):
            exe_path = Path(sys.executable).resolve()
            return f'"{exe_path}" --hidden'
        python_exe = Path(sys.executable).resolve()
        pythonw_exe = python_exe.with_name("pythonw.exe")
        launcher = pythonw_exe if pythonw_exe.exists() else python_exe
        script_file = Path(__file__).resolve()
        return f'"{launcher}" "{script_file}" --hidden'

    def is_enabled(self) -> bool:
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self._REG_PATH, 0, winreg.KEY_READ)
            try:
                winreg.QueryValueEx(key, self.app_name)
                return True
            except FileNotFoundError:
                return False
            finally:
                winreg.CloseKey(key)
        except Exception:
            return self.script_path.exists()

    def sync(self, enabled: bool) -> None:
        if enabled:
            self.enable()
        else:
            self.disable()

    def enable(self) -> None:
        try:
            import winreg
            cmd = self._reg_command()
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, self._REG_PATH, 0, winreg.KEY_SET_VALUE
            )
            winreg.SetValueEx(key, self.app_name, 0, winreg.REG_SZ, cmd)
            winreg.CloseKey(key)
            logger.info("Autoinicio registrado en HKCU\\Run: %s -> %s", self.app_name, cmd)
            # Limpiar .bat antiguo si existe
            if self.script_path.exists():
                try:
                    self.script_path.unlink()
                except Exception:
                    pass
        except Exception:
            logger.exception("No se pudo registrar autoinicio en el Registro de Windows")
            raise

    def disable(self) -> None:
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, self._REG_PATH, 0, winreg.KEY_SET_VALUE
            )
            try:
                winreg.DeleteValue(key, self.app_name)
                logger.info("Autoinicio eliminado del Registro: %s", self.app_name)
            except FileNotFoundError:
                pass
            finally:
                winreg.CloseKey(key)
        except Exception:
            logger.exception("No se pudo eliminar autoinicio del Registro de Windows")
        # También limpiar .bat antiguo si existe
        try:
            if self.script_path.exists():
                self.script_path.unlink()
        except Exception:
            pass


class LocalStore:
    def __init__(self, config_manager: ConfigManager):
        self.config_manager = config_manager
        self.ensure_schema()

    def _resolve_db_path(self) -> Path:
        cfg = self.config_manager.get()
        db_path = Path(str(cfg.get("sqlite_db") or DEFAULT_SQLITE_PATH)).expanduser()
        if not db_path.is_absolute():
            db_path = (BASE_DIR / db_path).resolve()
        db_path.parent.mkdir(parents=True, exist_ok=True)
        return db_path

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self._resolve_db_path(), timeout=30)
        conn.row_factory = sqlite3.Row
        return conn

    @staticmethod
    def ensure_schema_at_path(db_path_text: str) -> Path:
        db_path = Path(str(db_path_text)).expanduser()
        if not db_path.is_absolute():
            db_path = (BASE_DIR / db_path).resolve()
        db_path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(db_path, timeout=15)
        conn.row_factory = sqlite3.Row
        try:
            LocalStore._ensure_schema_on_connection(conn)
        finally:
            conn.close()
        return db_path

    @staticmethod
    def _ensure_schema_on_connection(conn: sqlite3.Connection) -> None:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS scans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL,
                descripcion TEXT NOT NULL,
                stock REAL NOT NULL DEFAULT 0,
                fecha TEXT NOT NULL,
                sincronizado INTEGER NOT NULL DEFAULT 0,
                sync_error TEXT,
                anulado INTEGER NOT NULL DEFAULT 0,
                anulado_motivo TEXT,
                created_at TEXT NOT NULL
            )
            """
        )

        existing_columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(scans)").fetchall()
        }

        if "sync_error" not in existing_columns:
            conn.execute("ALTER TABLE scans ADD COLUMN sync_error TEXT")

        if "anulado" not in existing_columns:
            conn.execute("ALTER TABLE scans ADD COLUMN anulado INTEGER NOT NULL DEFAULT 0")

        if "anulado_motivo" not in existing_columns:
            conn.execute("ALTER TABLE scans ADD COLUMN anulado_motivo TEXT")

        if "created_at" not in existing_columns:
            conn.execute("ALTER TABLE scans ADD COLUMN created_at TEXT")
            conn.execute(
                """
                UPDATE scans
                SET created_at = fecha
                WHERE created_at IS NULL OR TRIM(created_at) = ''
                """
            )

        conn.execute(
            """
            UPDATE scans
            SET anulado = 0
            WHERE anulado IS NULL
            """
        )
        conn.commit()

    def ensure_schema(self) -> None:
        with self._connect() as conn:
            self._ensure_schema_on_connection(conn)

    def insert_scan(self, codigo: str, descripcion: str, stock: float) -> int:
        now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO scans (
                    codigo, descripcion, stock, fecha, sincronizado,
                    sync_error, anulado, anulado_motivo, created_at
                )
                VALUES (?, ?, ?, ?, 0, NULL, 0, NULL, ?)
                """,
                (codigo, descripcion, stock, now_text, now_text),
            )
            conn.commit()
            return int(cur.lastrowid)

    def count_history(self) -> int:
        with self._connect() as conn:
            cur = conn.execute("SELECT COUNT(*) AS total FROM scans")
            row = cur.fetchone()
            return int(row["total"] or 0)

    def get_history_page(self, limit: int, offset: int) -> List[sqlite3.Row]:
        with self._connect() as conn:
            cur = conn.execute(
                """
                SELECT id, codigo, descripcion, stock, fecha, sincronizado,
                       sync_error, anulado, anulado_motivo, created_at
                FROM scans
                ORDER BY id DESC
                LIMIT ? OFFSET ?
                """,
                (limit, offset),
            )
            return cur.fetchall()

    def get_history_by_range(self, start_date: Optional[str], end_date: Optional[str]) -> List[sqlite3.Row]:
        conditions = []
        params: List[Any] = []

        if start_date:
            conditions.append("fecha >= ?")
            params.append(start_date)

        if end_date:
            conditions.append("fecha <= ?")
            params.append(end_date)

        sql = """
            SELECT id, codigo, descripcion, stock, fecha, sincronizado,
                   sync_error, anulado, anulado_motivo, created_at
            FROM scans
        """
        if conditions:
            sql += " WHERE " + " AND ".join(conditions)
        sql += " ORDER BY id ASC"

        with self._connect() as conn:
            cur = conn.execute(sql, tuple(params))
            return cur.fetchall()

    def get_pending_ready(self, hold_seconds: int, limit: int = 100) -> List[sqlite3.Row]:
        cutoff = (datetime.now() - timedelta(seconds=max(0, hold_seconds))).strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            cur = conn.execute(
                """
                SELECT id, codigo, descripcion, stock, fecha, created_at,
                       sincronizado, anulado, anulado_motivo, sync_error
                FROM scans
                WHERE sincronizado = 0
                  AND created_at <= ?
                ORDER BY id ASC
                LIMIT ?
                """,
                (cutoff, limit),
            )
            return cur.fetchall()

    def get_recent_active_same_code(self, codigo: str, window_seconds: int) -> List[sqlite3.Row]:
        cutoff = (datetime.now() - timedelta(seconds=max(0, window_seconds))).strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            cur = conn.execute(
                """
                SELECT id, codigo, created_at
                FROM scans
                WHERE codigo = ?
                  AND sincronizado = 0
                  AND anulado = 0
                  AND created_at >= ?
                ORDER BY id ASC
                """,
                (codigo, cutoff),
            )
            return cur.fetchall()

    def maybe_cancel_scan_group(
        self,
        codigo: str,
        window_seconds: int,
        repeat_count: int,
    ) -> Tuple[bool, str, List[int]]:
        if window_seconds <= 0 or repeat_count <= 1:
            return False, "", []

        rows = self.get_recent_active_same_code(codigo, window_seconds)
        ids = [int(row["id"]) for row in rows]
        if len(ids) >= repeat_count:
            reason = (
                f"Anulado por repetición de {codigo}: "
                f"{len(ids)} lecturas dentro de {window_seconds}s"
            )
            self.mark_cancelled_many(ids, reason)
            return True, reason, ids
        return False, "", []

    def mark_synced(self, record_id: int) -> None:
        with self._connect() as conn:
            conn.execute(
                "UPDATE scans SET sincronizado = 1, sync_error = NULL WHERE id = ?",
                (record_id,),
            )
            conn.commit()

    def set_sync_error(self, record_id: int, message: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "UPDATE scans SET sync_error = ? WHERE id = ?",
                (message[:500], record_id),
            )
            conn.commit()

    def mark_cancelled(self, record_id: int, reason: str) -> None:
        self.mark_cancelled_many([record_id], reason)

    def mark_cancelled_many(self, record_ids: List[int], reason: str) -> None:
        if not record_ids:
            return
        placeholders = ",".join(["?"] * len(record_ids))
        params = [reason[:500], *record_ids]
        with self._connect() as conn:
            conn.execute(
                f"""
                UPDATE scans
                SET anulado = 1,
                    anulado_motivo = ?,
                    sync_error = NULL
                WHERE id IN ({placeholders})
                """,
                params,
            )
            conn.commit()


class MySQLService:
    def __init__(self, query: str):
        self.query = query

    def _connect(self, cfg: Dict[str, Any]):
        return mysql.connector.connect(
            host=str(cfg.get("mysql_host", "")).strip(),
            user=str(cfg.get("mysql_user", "")).strip(),
            password=str(cfg.get("mysql_password", "")),
            database=str(cfg.get("mysql_database", "")).strip(),
            connection_timeout=5,
            autocommit=True,
            use_pure=True,   # fuerza implementación Python pura; evita RuntimeError del C-extension en Python 3.13+
        )

    def fetch_product(self, cfg: Dict[str, Any], codigo: str) -> Tuple[str, float]:
        conn = None
        cursor = None
        try:
            conn = self._connect(cfg)
            cursor = conn.cursor(dictionary=True)
            cursor.execute(self.query, (codigo,))
            row = cursor.fetchone()
            if not row:
                return "No encontrado", 0.0
            descripcion = str(row.get("descripcion") or "Sin descripción").strip() or "Sin descripción"
            stock = float(row.get("total_stock") or 0)
            return descripcion, stock
        finally:
            try:
                if cursor is not None:
                    cursor.close()
            except Exception:
                pass
            try:
                if conn is not None:
                    conn.close()
            except Exception:
                pass

    def test_connection(self, cfg: Dict[str, Any]) -> Tuple[bool, str]:
        import traceback as _tb
        conn = None
        try:
            conn = self._connect(cfg)
            cursor = conn.cursor()
            cursor.execute("SELECT 1")
            cursor.fetchone()
            cursor.close()
            return True, "Conexión MySQL correcta."
        except Exception as exc:
            # Capturar causa raíz completa (mysql-connector a veces envuelve el error)
            cause = exc.__cause__ or exc.__context__ or exc
            detail = "".join(_tb.format_exception(type(cause), cause, cause.__traceback__))
            short = f"{type(cause).__name__}: {cause}"
            logger.error("Error en test_connection MySQL:\n%s", detail)
            return False, f"Error MySQL: {short}\n\nVer log para detalle completo."
        finally:
            try:
                if conn is not None:
                    conn.close()
            except Exception:
                pass


class GoogleSheetsService:
    HEADERS = [
        "ID",
        "Código",
        "Descripción",
        "Stock",
        "Fecha",
        "Estado",
        "Detalle",
        "Creado en",
    ]

    def _get_client(self, cfg: Dict[str, Any]):
        cred_file = Path(str(cfg.get("google_credentials", "")).strip()).expanduser()
        if not cred_file.exists():
            raise FileNotFoundError("No se encontró el archivo credentials.json")

        creds = ServiceAccountCredentials.from_json_keyfile_name(str(cred_file), GOOGLE_SCOPES)
        return gspread.authorize(creds)

    def _get_sheet(self, cfg: Dict[str, Any]):
        sheet_id = str(cfg.get("google_sheet_id", "")).strip()
        if not sheet_id:
            raise ValueError("Falta el ID de Google Sheets")
        client = self._get_client(cfg)
        return client.open_by_key(sheet_id).sheet1

    def _ensure_headers(self, ws) -> None:
        first_row = ws.row_values(1)
        if first_row[: len(self.HEADERS)] != self.HEADERS:
            ws.update("A1:H1", [self.HEADERS])

    def test_connection(self, cfg: Dict[str, Any]) -> Tuple[bool, str]:
        try:
            ws = self._get_sheet(cfg)
            self._ensure_headers(ws)
            _ = ws.title
            return True, "Conexión Google Sheets correcta."
        except Exception as exc:
            return False, f"Error Google Sheets: {exc}"

    def append_scan(self, cfg: Dict[str, Any], row: sqlite3.Row) -> None:
        ws = self._get_sheet(cfg)
        self._ensure_headers(ws)
        existing_ids = set(ws.col_values(1)[1:])
        if str(row["id"]) in existing_ids:
            return

        status = "ANULADO" if int(row["anulado"]) == 1 else "OK"
        detail = str(row["anulado_motivo"] or row["sync_error"] or "").strip()

        ws.append_row(
            [
                str(row["id"]),
                str(row["codigo"]),
                str(row["descripcion"]),
                str(row["stock"]),
                str(row["fecha"]),
                status,
                detail,
                str(row["created_at"]),
            ],
            value_input_option="USER_ENTERED",
        )


class ScannerWorker(threading.Thread):
    def __init__(
        self,
        config_manager: ConfigManager,
        local_store: LocalStore,
        mysql_service: MySQLService,
        ui_queue: queue.Queue,
        stop_event: threading.Event,
        reload_event: threading.Event,
    ):
        super().__init__(name="ScannerWorker", daemon=True)
        self.config_manager = config_manager
        self.local_store = local_store
        self.mysql_service = mysql_service
        self.ui_queue = ui_queue
        self.stop_event = stop_event
        self.reload_event = reload_event
        self.serial_conn: Optional[serial.Serial] = None
        self.current_port = ""

    def close_serial(self) -> None:
        try:
            if self.serial_conn is not None and self.serial_conn.is_open:
                self.serial_conn.close()
        except Exception:
            pass
        self.serial_conn = None
        self.current_port = ""

    def _open_serial(self, cfg: Dict[str, Any]) -> None:
        desired_port = str(cfg.get("com_port", "")).strip()
        baudrate = int(cfg.get("scanner_baudrate", 9600) or 9600)

        if not desired_port:
            self.close_serial()
            return

        if self.serial_conn is not None and self.serial_conn.is_open and self.current_port == desired_port:
            return

        self.close_serial()
        self.serial_conn = serial.Serial(
            port=desired_port,
            baudrate=baudrate,
            timeout=1,                # tiempo máx. de espera entre lecturas
            inter_byte_timeout=0.1,   # entrega los bytes tan pronto dejan de llegar
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            xonxoff=False,            # sin control de flujo XON/XOFF
            rtscts=False,             # sin control de flujo RTS/CTS (CDC)
            dsrdtr=False,             # sin control de flujo DSR/DTR (CDC)
        )
        self.current_port = desired_port
        self.ui_queue.put(("status", f"Escáner conectado en {desired_port}"))

    @staticmethod
    def _clean_code(raw: bytes) -> str:
        # El DS2278 termina con \r (CR) por defecto; strip() elimina \r, \n y espacios.
        return raw.decode("utf-8", errors="ignore").strip()

    def _process_code(self, codigo: str) -> None:
        cfg = self.config_manager.get()
        try:
            descripcion, stock = self.mysql_service.fetch_product(cfg, codigo)
        except MySQLError:
            logger.exception("MySQL fuera de línea. Se guarda el código con Error DB")
            descripcion, stock = "Error DB", 0.0
        except Exception:
            logger.exception("Error consultando producto. Se guarda el código con Error DB")
            descripcion, stock = "Error DB", 0.0

        self.local_store.insert_scan(codigo, descripcion, stock)
        cancelled, _reason, cancelled_ids = self.local_store.maybe_cancel_scan_group(
            codigo=codigo,
            window_seconds=int(cfg.get("correction_window_seconds", 60) or 60),
            repeat_count=int(cfg.get("correction_repeat_count", 3) or 3),
        )

        if cancelled:
            self.ui_queue.put(
                (
                    "status",
                    f"Grupo anulado para {codigo}: {len(cancelled_ids)} lecturas dentro de la ventana",
                )
            )
        else:
            self.ui_queue.put(("status", f"Escaneado: {codigo} | {descripcion}"))

        self.ui_queue.put(("refresh_history", None))

    def run(self) -> None:
        while not self.stop_event.is_set():
            cfg = self.config_manager.get()
            try:
                if self.reload_event.is_set():
                    self.reload_event.clear()
                    self.close_serial()

                self._open_serial(cfg)
                if self.serial_conn is None:
                    time.sleep(1)
                    continue

                # Lee bytes hasta que haya silencio de 100 ms (inter_byte_timeout).
                # Funciona con cualquier terminador: CR, LF, CR+LF o ninguno.
                raw = self.serial_conn.read(4096)
                if not raw:
                    continue

                codigo = self._clean_code(raw)
                if not codigo:
                    continue

                self._process_code(codigo)
            except serial.SerialException as exc:
                logger.exception("Error de puerto serial")
                self.ui_queue.put(("status", f"Error serial: {exc}"))
                self.close_serial()
                time.sleep(2)
            except Exception as exc:
                logger.exception("Error en ScannerWorker")
                self.ui_queue.put(("status", f"Scanner: {exc}"))
                time.sleep(1)

        self.close_serial()


class HIDScannerWorker(threading.Thread):
    """
    Captura códigos en modo HID Teclado (sin puerto COM).
    Compatible con cualquier escaner que emule teclado USB HID (Zebra, SAT, Honeywell, etc.).
    Los caracteres que llegan con <= hid_inter_char_ms entre sí se consideran
    del escáner; al llegar Enter/CR/LF se procesa el código.
    hid_inter_char_ms se configura en el panel (default 150 ms);
    aumenta si el escáner pierde dígitos; disminuye para ignorar mejor la escritura humana.
    """
    MIN_BARCODE_LEN = 3

    def __init__(
        self,
        config_manager: ConfigManager,
        local_store: LocalStore,
        mysql_service: MySQLService,
        ui_queue: queue.Queue,
        stop_event: threading.Event,
        reload_event: threading.Event,
    ):
        super().__init__(name="HIDScannerWorker", daemon=True)
        self.config_manager = config_manager
        self.local_store = local_store
        self.mysql_service = mysql_service
        self.ui_queue = ui_queue
        self.stop_event = stop_event
        self.reload_event = reload_event
        self._buffer: List[str] = []
        self._last_char_time: float = 0.0
        self._code_queue: "queue.Queue[str]" = queue.Queue()
        self._inter_char_ms: float = 150.0  # se actualiza desde config en run()

    def _on_press(self, key) -> None:
        try:
            if hasattr(key, "char") and key.char is not None:
                ch = key.char
            elif key in (pynput_keyboard.Key.enter, pynput_keyboard.Key.num_lock):
                ch = "\r"
            else:
                # Tecla especial (Shift, Ctrl, etc.) — no resetear buffer,
                # el escáner puede emitir shift antes de letras mayúsculas
                return

            now = time.monotonic() * 1000.0  # milisegundos

            if ch in ("\r", "\n"):
                if len(self._buffer) >= self.MIN_BARCODE_LEN:
                    codigo = "".join(self._buffer).strip()
                    if codigo:
                        self._code_queue.put(codigo)
                self._buffer.clear()
                self._last_char_time = 0.0
            else:
                elapsed = now - self._last_char_time
                if self._buffer and elapsed > self.INTER_CHAR_MAX_MS:
                    # Demasiado lento entre caracteres → escritura humana
                    self._buffer.clear()
                self._buffer.append(ch)
                self._last_char_time = now
        except Exception:
            pass

    def _process_code(self, codigo: str) -> None:
        cfg = self.config_manager.get()
        try:
            descripcion, stock = self.mysql_service.fetch_product(cfg, codigo)
        except MySQLError:
            logger.exception("MySQL fuera de línea. Se guarda código con Error DB")
            descripcion, stock = "Error DB", 0.0
        except Exception:
            logger.exception("Error consultando producto en modo HID")
            descripcion, stock = "Error DB", 0.0

        self.local_store.insert_scan(codigo, descripcion, stock)
        cancelled, _reason, cancelled_ids = self.local_store.maybe_cancel_scan_group(
            codigo=codigo,
            window_seconds=int(cfg.get("correction_window_seconds", 60) or 60),
            repeat_count=int(cfg.get("correction_repeat_count", 3) or 3),
        )

        if cancelled:
            self.ui_queue.put(
                ("status", f"Grupo anulado para {codigo}: {len(cancelled_ids)} lecturas dentro de la ventana")
            )
        else:
            self.ui_queue.put(("status", f"Escaneado (HID): {codigo} | {descripcion}"))

        self.ui_queue.put(("refresh_history", None))

    def run(self) -> None:
        if not _PYNPUT_AVAILABLE:
            self.ui_queue.put((
                "status",
                "ERROR: pynput no instalado. Ejecuta: pip install pynput",
            ))
            return

        self.ui_queue.put(("status", "Modo HID Teclado activo — apunta y escanea"))

        with pynput_keyboard.Listener(on_press=self._on_press, suppress=False) as listener:
            # Leer umbral inicial desde config
            _cfg0 = self.config_manager.get()
            self._inter_char_ms = float(_cfg0.get("hid_inter_char_ms", 150) or 150)
            while not self.stop_event.is_set():
                if self.reload_event.is_set():
                    self.reload_event.clear()
                    self._buffer.clear()
                    # Actualizar umbral desde config al recargar
                    cfg = self.config_manager.get()
                    self._inter_char_ms = float(cfg.get("hid_inter_char_ms", 150) or 150)
                try:
                    codigo = self._code_queue.get(timeout=0.3)
                except queue.Empty:
                    continue
                self._process_code(codigo)
            listener.stop()


class SyncWorker(threading.Thread):
    def __init__(
        self,
        config_manager: ConfigManager,
        local_store: LocalStore,
        sheets_service: GoogleSheetsService,
        ui_queue: queue.Queue,
        stop_event: threading.Event,
        wakeup_event: threading.Event,
    ):
        super().__init__(name="SyncWorker", daemon=True)
        self.config_manager = config_manager
        self.local_store = local_store
        self.sheets_service = sheets_service
        self.ui_queue = ui_queue
        self.stop_event = stop_event
        self.wakeup_event = wakeup_event

    def run(self) -> None:
        while not self.stop_event.is_set():
            cfg = self.config_manager.get()
            sync_interval = max(3, int(cfg.get("sync_interval", 15) or 15))
            hold_seconds = max(0, int(cfg.get("correction_window_seconds", 60) or 60))

            try:
                rows = self.local_store.get_pending_ready(hold_seconds=hold_seconds, limit=100)
                if rows:
                    synced_count = 0
                    for row in rows:
                        try:
                            self.sheets_service.append_scan(cfg, row)
                            self.local_store.mark_synced(int(row["id"]))
                            synced_count += 1
                        except Exception as exc:
                            logger.exception("No se pudo sincronizar el registro %s", row["id"])
                            self.local_store.set_sync_error(int(row["id"]), str(exc))
                    if synced_count:
                        self.ui_queue.put(("status", f"Sincronizados {synced_count} registros"))
                        self.ui_queue.put(("refresh_history", None))
            except Exception as exc:
                logger.exception("Error en SyncWorker")
                self.ui_queue.put(("status", f"Sincronización pendiente: {exc}"))

            self.wakeup_event.wait(timeout=sync_interval)
            self.wakeup_event.clear()


class TrayController:
    def __init__(self, ui_queue: queue.Queue):
        self.ui_queue = ui_queue
        self.icon: Optional[PystrayIcon] = None
        self.thread: Optional[threading.Thread] = None

    def _create_image(self) -> Image.Image:
        width, height = 64, 64
        image = Image.new("RGBA", (width, height), (33, 150, 243, 255))
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((8, 8, 56, 56), radius=10, fill=(13, 71, 161, 255))
        draw.text((22, 18), "Z", fill="white")
        return image

    def _run(self) -> None:
        menu = pystray.Menu(
            TrayItem("Abrir", lambda: self.ui_queue.put(("open_window", None))),
            TrayItem("Ver configuración", lambda: self.ui_queue.put(("open_config", None))),
            TrayItem("Salir", lambda: self.ui_queue.put(("exit_app", None))),
        )
        self.icon = pystray.Icon(APP_NAME, self._create_image(), APP_NAME, menu)
        self.icon.run()

    def start(self) -> None:
        if self.thread and self.thread.is_alive():
            return
        self.thread = threading.Thread(target=self._run, name="TrayThread", daemon=True)
        self.thread.start()

    def stop(self) -> None:
        try:
            if self.icon is not None:
                self.icon.stop()
        except Exception:
            pass


class ZebraCloudSyncApp(ctk.CTk):
    def __init__(self, start_hidden: bool = False):
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_NAME)
        self.geometry("1240x720")
        self.minsize(1080, 640)

        self.ui_queue: queue.Queue = queue.Queue()
        self.stop_event = threading.Event()
        self.scanner_reload_event = threading.Event()
        self.sync_wakeup_event = threading.Event()
        self.window_hidden = False
        self.exiting = False
        self.current_page = 1
        self.page_count = 1
        self.start_hidden = start_hidden
        self._tab_guard_active = False
        self._port_map: Dict[str, str] = {}

        self.config_manager = ConfigManager(CONFIG_PATH)
        self.startup_manager = StartupManager(APP_NAME)
        self.local_store = LocalStore(self.config_manager)
        self.mysql_service = MySQLService(MYSQL_QUERY)
        self.sheets_service = GoogleSheetsService()

        _init_cfg = self.config_manager.get()
        _scanner_mode = str(_init_cfg.get("scanner_mode", "hid"))
        if _scanner_mode == "serial":
            self.scanner_worker: threading.Thread = ScannerWorker(
                config_manager=self.config_manager,
                local_store=self.local_store,
                mysql_service=self.mysql_service,
                ui_queue=self.ui_queue,
                stop_event=self.stop_event,
                reload_event=self.scanner_reload_event,
            )
        else:
            self.scanner_worker = HIDScannerWorker(
                config_manager=self.config_manager,
                local_store=self.local_store,
                mysql_service=self.mysql_service,
                ui_queue=self.ui_queue,
                stop_event=self.stop_event,
                reload_event=self.scanner_reload_event,
            )
        self.sync_worker = SyncWorker(
            config_manager=self.config_manager,
            local_store=self.local_store,
            sheets_service=self.sheets_service,
            ui_queue=self.ui_queue,
            stop_event=self.stop_event,
            wakeup_event=self.sync_wakeup_event,
        )
        self.tray_controller = TrayController(self.ui_queue)

        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)

        self._build_ui()
        self._load_config_to_form()
        self.refresh_com_ports(update_value=False)
        self.refresh_history(reset_page=True)
        self._history_countdown_tick()
        self._process_ui_queue()
        self._periodic_refresh_ports()

        self.scanner_worker.start()
        self.sync_worker.start()
        self.tray_controller.start()

        if self.start_hidden:
            self.after(400, self.hide_to_tray)

        self._set_status("Aplicación iniciada")

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(self, command=self._on_tab_changed)
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=16, pady=(16, 8))
        self.tabview.add("Historial")
        self.tabview.add("Configuración")

        self._build_history_tab()
        self._build_config_tab()

        self.status_var = tk.StringVar(value="Listo")
        self.status_label = ctk.CTkLabel(self, textvariable=self.status_var, anchor="w")
        self.status_label.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))
        self.tabview.set("Historial")

    def _build_history_tab(self) -> None:
        tab = self.tabview.tab("Historial")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)

        title = ctk.CTkLabel(tab, text="Historial de escaneos", font=ctk.CTkFont(size=18, weight="bold"))
        title.grid(row=0, column=0, sticky="w", padx=24, pady=(12, 0))

        topbar = ctk.CTkFrame(tab)
        topbar.grid(row=1, column=0, sticky="ew", padx=12, pady=12)
        for col in range(8):
            topbar.grid_columnconfigure(col, weight=1)

        self.entry_export_start = ctk.CTkEntry(topbar, placeholder_text="Desde YYYY-MM-DD")
        self.entry_export_start.grid(row=0, column=0, sticky="ew", padx=6, pady=12)

        self.entry_export_end = ctk.CTkEntry(topbar, placeholder_text="Hasta YYYY-MM-DD")
        self.entry_export_end.grid(row=0, column=1, sticky="ew", padx=6, pady=12)

        ctk.CTkButton(topbar, text="Exportar Excel", command=self.export_excel_range, width=140).grid(
            row=0, column=2, padx=6, pady=12
        )
        ctk.CTkButton(topbar, text="Refrescar", command=lambda: self.refresh_history(reset_page=False), width=120).grid(
            row=0, column=3, padx=6, pady=12
        )
        ctk.CTkButton(topbar, text="Configuración", command=self.request_open_config, width=140).grid(
            row=0, column=4, padx=6, pady=12
        )

        tree_frame = ctk.CTkFrame(tab)
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        columns = ("id", "codigo", "descripcion", "stock", "fecha", "tiempo", "estado")
        self.history_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        self.history_tree.heading("id", text="ID")
        self.history_tree.heading("codigo", text="Código")
        self.history_tree.heading("descripcion", text="Descripción")
        self.history_tree.heading("stock", text="Stock")
        self.history_tree.heading("fecha", text="Fecha")
        self.history_tree.heading("tiempo", text="Tiempo restante")
        self.history_tree.heading("estado", text="Estado")

        self.history_tree.column("id", width=70, anchor="center")
        self.history_tree.column("codigo", width=150, anchor="center")
        self.history_tree.column("descripcion", width=300, anchor="w")
        self.history_tree.column("stock", width=90, anchor="center")
        self.history_tree.column("fecha", width=150, anchor="center")
        self.history_tree.column("tiempo", width=120, anchor="center")
        self.history_tree.column("estado", width=180, anchor="center")

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=yscroll.set)

        self.history_tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        yscroll.grid(row=0, column=1, sticky="ns", padx=(0, 12), pady=12)

        pager = ctk.CTkFrame(tab)
        pager.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 12))
        pager.grid_columnconfigure(1, weight=1)

        self.btn_prev = ctk.CTkButton(pager, text="Anterior", width=120, command=self.prev_page)
        self.btn_prev.grid(row=0, column=0, padx=12, pady=12, sticky="w")

        self.page_info_var = tk.StringVar(value="Página 1/1")
        self.page_info_label = ctk.CTkLabel(pager, textvariable=self.page_info_var, anchor="center")
        self.page_info_label.grid(row=0, column=1, padx=12, pady=12)

        self.btn_next = ctk.CTkButton(pager, text="Siguiente", width=120, command=self.next_page)
        self.btn_next.grid(row=0, column=2, padx=12, pady=12, sticky="e")

    def _build_config_tab(self) -> None:
        tab = self.tabview.tab("Configuración")
        for idx in range(4):
            tab.grid_columnconfigure(idx, weight=1)

        row = 0
        title = ctk.CTkLabel(tab, text="Configuración del agente", font=ctk.CTkFont(size=16, weight="bold"))
        title.grid(row=row, column=0, columnspan=4, sticky="w", padx=12, pady=(6, 2))
        row += 1

        # ── Modo escáner ──────────────────────────────────────────────
        self._label(tab, row, 0, "Modo escáner")
        self.combobox_scanner_mode = ctk.CTkComboBox(
            tab,
            values=["HID Teclado (USB, sin driver COM)", "Serial / USB CDC (con driver COM)"],
            width=340,
            state="readonly",
        )
        self.combobox_scanner_mode.set("HID Teclado (USB, sin driver COM)")
        self.combobox_scanner_mode.grid(row=row, column=1, columnspan=2, sticky="ew", padx=12, pady=4)
        ctk.CTkLabel(tab, text="Reinicia la app al cambiar modo", anchor="w", text_color="gray").grid(
            row=row, column=3, sticky="w", padx=12, pady=4
        )
        row += 1

        # ── Puerto COM ────────────────────────────────────────────────
        self._label(tab, row, 0, "Puerto COM")
        self.combobox_port = ctk.CTkComboBox(tab, values=[""], width=320)
        self.combobox_port.grid(row=row, column=1, sticky="ew", padx=12, pady=4)
        ctk.CTkButton(tab, text="Actualizar puertos", command=self.refresh_com_ports, height=30).grid(row=row, column=2, padx=12, pady=4, sticky="ew")
        ctk.CTkButton(tab, text="Test COM", command=self.test_com, height=30).grid(row=row, column=3, padx=12, pady=4, sticky="ew")
        row += 1

        # Baudrate + MySQL Host en la misma franja visual
        self._label(tab, row, 0, "Baudrate")
        self.combobox_baudrate = ctk.CTkComboBox(tab, values=["9600", "19200", "38400", "57600", "115200"], width=160)
        self.combobox_baudrate.set("9600")
        self.combobox_baudrate.grid(row=row, column=1, sticky="w", padx=12, pady=4)
        self._label(tab, row, 2, "MySQL Host")
        self.entry_mysql_host = self._entry(tab, row, 3)
        row += 1

        self._label(tab, row, 0, "MySQL Usuario")
        self.entry_mysql_user = self._entry(tab, row, 1)
        self._label(tab, row, 2, "MySQL Password")
        self.entry_mysql_password = ctk.CTkEntry(tab, show="*")
        self.entry_mysql_password.grid(row=row, column=3, sticky="ew", padx=12, pady=4)
        row += 1

        self._label(tab, row, 0, "MySQL Base de datos")
        self.entry_mysql_database = self._entry(tab, row, 1)
        ctk.CTkButton(tab, text="Test MySQL", command=self.test_mysql, height=30).grid(row=row, column=3, padx=12, pady=4, sticky="ew")
        row += 1

        self._label(tab, row, 0, "Google Sheet ID")
        self.entry_google_sheet_id = self._entry(tab, row, 1, columnspan=2)
        ctk.CTkButton(tab, text="Test Google", command=self.test_google, height=30).grid(row=row, column=3, padx=12, pady=4, sticky="ew")
        row += 1

        self._label(tab, row, 0, "credentials.json")
        self.entry_google_credentials = self._entry(tab, row, 1, columnspan=2)
        ctk.CTkButton(tab, text="Buscar archivo", command=self._browse_google_credentials, height=30).grid(row=row, column=3, padx=12, pady=4, sticky="ew")
        row += 1

        self._label(tab, row, 0, "SQLite .db")
        self.entry_sqlite_db = self._entry(tab, row, 1, columnspan=2)
        browse_sqlite = ctk.CTkFrame(tab, fg_color="transparent")
        browse_sqlite.grid(row=row, column=3, sticky="ew", padx=12, pady=4)
        browse_sqlite.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(browse_sqlite, text="Buscar", command=self._browse_sqlite_db, height=30).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ctk.CTkButton(browse_sqlite, text="Test SQLite", command=self.test_sqlite, height=30).grid(row=0, column=1, sticky="ew", padx=(4, 0))
        row += 1

        self._label(tab, row, 0, "Intervalo sync (seg)")
        self.entry_sync_interval = self._entry(tab, row, 1)
        self._label(tab, row, 2, "Ventana corrección (seg)")
        self.entry_correction_window = self._entry(tab, row, 3)
        row += 1

        self._label(tab, row, 0, "Repeticiones para anular")
        self.entry_repeat_count = self._entry(tab, row, 1)
        self._label(tab, row, 2, "Registros por página")
        self.entry_history_page_size = self._entry(tab, row, 3)
        row += 1

        self._label(tab, row, 0, "HID inter-carácter (ms)")
        self.entry_hid_inter_char_ms = self._entry(tab, row, 1)
        ctk.CTkLabel(
            tab,
            text="↑ sube si pierde dígitos (prueba 200-500 ms);  ↓ baja para ignorar escritura humana",
            anchor="w",
            text_color="gray",
            font=ctk.CTkFont(size=11),
        ).grid(row=row, column=2, columnspan=2, sticky="w", padx=12, pady=4)
        row += 1

        self.var_start_with_windows = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            tab,
            text="Iniciar con Windows en segundo plano",
            variable=self.var_start_with_windows,
            onvalue=True,
            offvalue=False,
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=12, pady=(4, 2))
        row += 1

        # ── Contraseña + footer en un frame compacto horizontal ───────
        bottom = ctk.CTkFrame(tab)
        bottom.grid(row=row, column=0, columnspan=4, sticky="ew", padx=12, pady=(4, 6))
        for idx in range(6):
            bottom.grid_columnconfigure(idx, weight=1)

        ctk.CTkLabel(bottom, text="Contraseña:", anchor="w", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=(12, 4), pady=6
        )
        self.entry_password_current = ctk.CTkEntry(bottom, show="*", placeholder_text="Actual", height=30)
        self.entry_password_current.grid(row=0, column=1, sticky="ew", padx=4, pady=6)
        self.entry_password_new = ctk.CTkEntry(bottom, show="*", placeholder_text="Nueva", height=30)
        self.entry_password_new.grid(row=0, column=2, sticky="ew", padx=4, pady=6)
        self.entry_password_confirm = ctk.CTkEntry(bottom, show="*", placeholder_text="Confirmar", height=30)
        self.entry_password_confirm.grid(row=0, column=3, sticky="ew", padx=4, pady=6)
        ctk.CTkButton(bottom, text="Cambiar", command=self.change_config_password, height=30, width=110).grid(
            row=0, column=4, padx=4, pady=6, sticky="ew"
        )
        ctk.CTkButton(bottom, text="Guardar y aplicar", command=self.save_configuration, height=32, width=160,
                      font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=5, padx=(4, 12), pady=6, sticky="ew"
        )

    def _label(self, parent, row: int, col: int, text: str) -> None:
        ctk.CTkLabel(parent, text=text, anchor="w").grid(row=row, column=col, sticky="w", padx=12, pady=4)

    def _entry(self, parent, row: int, col: int, columnspan: int = 1):
        entry = ctk.CTkEntry(parent)
        entry.grid(row=row, column=col, columnspan=columnspan, sticky="ew", padx=12, pady=4)
        return entry

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)
        logger.info(message)

    def _load_config_to_form(self) -> None:
        cfg = self.config_manager.get()
        self.refresh_com_ports(update_value=False)

        stored_port = str(cfg.get("com_port", ""))
        display_port = next(
            (disp for disp, dev in self._port_map.items() if dev == stored_port),
            stored_port,
        )
        self.combobox_port.set(display_port)

        baud = str(cfg.get("scanner_baudrate", 9600) or 9600)
        self.combobox_baudrate.set(baud if baud in ["9600", "19200", "38400", "57600", "115200"] else "9600")

        mode = str(cfg.get("scanner_mode", "hid"))
        if mode == "serial":
            self.combobox_scanner_mode.set("Serial / USB CDC (con driver COM)")
        else:
            self.combobox_scanner_mode.set("HID Teclado (USB, sin driver COM)")

        for entry, key in [
            (self.entry_mysql_host, "mysql_host"),
            (self.entry_mysql_user, "mysql_user"),
            (self.entry_mysql_password, "mysql_password"),
            (self.entry_mysql_database, "mysql_database"),
            (self.entry_google_sheet_id, "google_sheet_id"),
            (self.entry_google_credentials, "google_credentials"),
            (self.entry_sqlite_db, "sqlite_db"),
            (self.entry_sync_interval, "sync_interval"),
            (self.entry_correction_window, "correction_window_seconds"),
            (self.entry_repeat_count, "correction_repeat_count"),
            (self.entry_history_page_size, "history_page_size"),
            (self.entry_hid_inter_char_ms, "hid_inter_char_ms"),
        ]:
            entry.delete(0, tk.END)
            entry.insert(0, str(cfg.get(key, "")))

        self.var_start_with_windows.set(bool(cfg.get("start_with_windows", False)))

    def refresh_com_ports(self, update_value: bool = True) -> None:
        port_list = list_ports.comports()
        self._port_map = {}
        for p in port_list:
            desc = str(p.description or "").strip()
            if desc and desc.lower() != "n/a":
                display = f"{p.device} - {desc}"
            else:
                display = p.device
            self._port_map[display] = p.device

        display_values = list(self._port_map.keys()) if self._port_map else [""]
        current_display = self.combobox_port.get().strip() if hasattr(self, "combobox_port") else ""
        current_device = self._port_map.get(current_display, current_display.split(" - ")[0].strip())

        self.combobox_port.configure(values=display_values)

        if update_value:
            matching = next((d for d, dev in self._port_map.items() if dev == current_device), None)
            if matching:
                self.combobox_port.set(matching)
            elif display_values and display_values[0] != "":
                self.combobox_port.set(display_values[0])
            else:
                self.combobox_port.set("")
        else:
            if current_display in self._port_map:
                self.combobox_port.set(current_display)
            else:
                matching = next((d for d, dev in self._port_map.items() if dev == current_device), None)
                if matching:
                    self.combobox_port.set(matching)

    def _periodic_refresh_ports(self) -> None:
        if not self.exiting:
            try:
                self.refresh_com_ports(update_value=False)
            except Exception:
                logger.exception("No se pudieron refrescar los puertos COM")
            self.after(5000, self._periodic_refresh_ports)

    def _sanitize_path(self, value: str, fallback: Path) -> str:
        path = Path(value.strip() or fallback).expanduser()
        if not path.is_absolute():
            path = (BASE_DIR / path).resolve()
        return str(path)

    def _safe_int(self, value: str, default: int, minimum: Optional[int] = None) -> int:
        try:
            number = int(str(value).strip())
        except Exception:
            number = default
        if minimum is not None and number < minimum:
            number = minimum
        return number

    def get_form_config(self) -> Dict[str, Any]:
        current_cfg = self.config_manager.get()
        selected_display = self.combobox_port.get().strip()
        actual_port = self._port_map.get(selected_display, selected_display.split(" - ")[0].strip())
        return {
            "com_port": actual_port,
            "mysql_host": self.entry_mysql_host.get().strip(),
            "mysql_user": self.entry_mysql_user.get().strip(),
            "mysql_password": self.entry_mysql_password.get(),
            "mysql_database": self.entry_mysql_database.get().strip(),
            "google_sheet_id": self.entry_google_sheet_id.get().strip(),
            "google_credentials": self._sanitize_path(self.entry_google_credentials.get(), BASE_DIR / "credentials.json"),
            "sqlite_db": self._sanitize_path(self.entry_sqlite_db.get(), DEFAULT_SQLITE_PATH),
            "sync_interval": self._safe_int(self.entry_sync_interval.get(), default=15, minimum=3),
            "scanner_baudrate": int(self.combobox_baudrate.get() or "9600"),
            "scanner_mode": "serial" if "Serial" in self.combobox_scanner_mode.get() else "hid",
            "correction_window_seconds": self._safe_int(self.entry_correction_window.get(), default=60, minimum=0),
            "correction_repeat_count": self._safe_int(self.entry_repeat_count.get(), default=3, minimum=2),
            "history_page_size": self._safe_int(self.entry_history_page_size.get(), default=50, minimum=5),
            "hid_inter_char_ms": self._safe_int(self.entry_hid_inter_char_ms.get(), default=150, minimum=30),
            "start_with_windows": bool(self.var_start_with_windows.get()),
            "config_password_hash": str(current_cfg.get("config_password_hash", "")),
            "configured_once": True,
        }

    def save_configuration(self) -> None:
        prev_cfg = self.config_manager.get()
        prev_mode = str(prev_cfg.get("scanner_mode", "hid"))

        cfg = self.get_form_config()
        self.config_manager.update(cfg)
        startup_ok = True
        try:
            self.startup_manager.sync(bool(cfg.get("start_with_windows", False)))
        except Exception as exc:
            startup_ok = False
            messagebox.showerror(
                "Error — Autoinicio",
                f"No se pudo configurar el inicio con Windows:\n{exc}\n\n"
                "Revisa agente_zebra.log para más detalles.",
            )
        self.local_store.ensure_schema()
        self.current_page = 1
        self.scanner_reload_event.set()
        self.sync_wakeup_event.set()
        self.refresh_history(reset_page=True)

        new_mode = str(cfg.get("scanner_mode", "hid"))
        if new_mode != prev_mode:
            messagebox.showinfo(
                "Modo de escáner cambiado",
                f"El modo cambió a '{new_mode.upper()}'. "
                "Cierra y vuelve a abrir la aplicación para que tome efecto.",
            )
            self._set_status(f"Modo cambiado a {new_mode.upper()} — reinicia la app")
        elif not startup_ok:
            self._set_status("Configuración guardada — error al configurar autoinicio (ver log)")
        elif bool(cfg.get("start_with_windows", False)):
            self._set_status("Configuración guardada y auto inicio con Windows activado")
        else:
            self._set_status("Configuración guardada y auto inicio con Windows desactivado")

    def _browse_google_credentials(self) -> None:
        selected = filedialog.askopenfilename(
            title="Seleccionar credentials.json",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")],
        )
        if selected:
            self.entry_google_credentials.delete(0, tk.END)
            self.entry_google_credentials.insert(0, selected)

    def _browse_sqlite_db(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="Seleccionar o crear SQLite",
            defaultextension=".db",
            filetypes=[("SQLite", "*.db"), ("Todos", "*.*")],
            initialfile=Path(self.entry_sqlite_db.get() or "agente_buffer.db").name,
        )
        if selected:
            self.entry_sqlite_db.delete(0, tk.END)
            self.entry_sqlite_db.insert(0, selected)

    def _show_message_async(self, kind: str, title: str, message: str) -> None:
        self.ui_queue.put((kind, (title, message)))

    def _start_background_test(self, label: str, func) -> None:
        self._set_status(f"Probando {label}...")

        def worker() -> None:
            try:
                ok, message = func()
                self.ui_queue.put(("status", message))
                if ok:
                    self._show_message_async("show_info", f"Test {label}", message)
                else:
                    self._show_message_async("show_error", f"Test {label}", message)
            except Exception as exc:
                logger.exception("Error en test %s", label)
                message = f"Error probando {label}: {exc}"
                self.ui_queue.put(("status", message))
                self._show_message_async("show_error", f"Test {label}", message)

        threading.Thread(target=worker, name=f"Test{label}", daemon=True).start()

    def test_com(self) -> None:
        def task() -> Tuple[bool, str]:
            cfg = self.get_form_config()
            port = str(cfg.get("com_port", "")).strip()
            if not port:
                return False, "Selecciona un puerto COM primero."

            ports = [p.device for p in list_ports.comports()]
            if port not in ports:
                return False, f"El puerto {port} no está activo."

            # Solo ScannerWorker (modo serial) tiene current_port/serial_conn
            if isinstance(self.scanner_worker, ScannerWorker):
                if self.scanner_worker.current_port == port and self.scanner_worker.serial_conn is not None:
                    if self.scanner_worker.serial_conn.is_open:
                        return True, f"El puerto {port} ya está abierto por el agente."

            serial_test = None
            try:
                serial_test = serial.Serial(
                    port=port,
                    baudrate=int(cfg.get("scanner_baudrate", 9600) or 9600),
                    timeout=1,
                    bytesize=serial.EIGHTBITS,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_ONE,
                    xonxoff=False,
                    rtscts=False,
                    dsrdtr=False,
                )
                return True, f"El puerto {port} responde correctamente."
            except Exception as exc:
                return False, f"No se pudo probar el puerto {port}: {exc}"
            finally:
                try:
                    if serial_test is not None and serial_test.is_open:
                        serial_test.close()
                except Exception:
                    pass

        self._start_background_test("COM", task)

    def test_mysql(self) -> None:
        self._start_background_test("MySQL", lambda: self.mysql_service.test_connection(self.get_form_config()))

    def test_google(self) -> None:
        self._start_background_test("Google", lambda: self.sheets_service.test_connection(self.get_form_config()))

    def test_sqlite(self) -> None:
        def task() -> Tuple[bool, str]:
            cfg = self.get_form_config()
            db_path = LocalStore.ensure_schema_at_path(str(cfg.get("sqlite_db", DEFAULT_SQLITE_PATH)))
            return True, f"SQLite listo en: {db_path}"

        self._start_background_test("SQLite", task)

    def _seconds_since_created(self, created_at: Any) -> int:
        try:
            created = datetime.strptime(str(created_at), "%Y-%m-%d %H:%M:%S")
            return max(0, int((datetime.now() - created).total_seconds()))
        except Exception:
            return 0

    def _remaining_seconds(self, row: sqlite3.Row) -> int:
        cfg = self.config_manager.get()
        window = max(0, int(cfg.get("correction_window_seconds", 60) or 60))
        if int(row["sincronizado"]) == 1 or int(row["anulado"]) == 1:
            return 0
        elapsed = self._seconds_since_created(row["created_at"])
        return max(0, window - elapsed)

    def _build_time_label(self, row: sqlite3.Row) -> str:
        return f"{self._remaining_seconds(row)}s"

    def _row_state(self, row: sqlite3.Row) -> str:
        if int(row["anulado"]) == 1 and int(row["sincronizado"]) == 1:
            return "✖ Anulado y enviado"
        if int(row["anulado"]) == 1:
            return "✖ Anulado"
        if int(row["sincronizado"]) == 1:
            return "✔ Sincronizado"
        if self._remaining_seconds(row) > 0:
            return "🕒 En ventana"
        return "🕒 Pendiente de envío"

    def _history_countdown_tick(self) -> None:
        if self.exiting:
            return
        try:
            self.refresh_history(reset_page=False)
        except Exception:
            logger.exception("No se pudo refrescar el cronómetro del historial")
        self.after(1000, self._history_countdown_tick)

    def refresh_history(self, reset_page: bool = False) -> None:
        if reset_page:
            self.current_page = 1

        for item in self.history_tree.get_children():
            self.history_tree.delete(item)

        cfg = self.config_manager.get()
        page_size = max(5, int(cfg.get("history_page_size", 50) or 50))
        total_records = self.local_store.count_history()
        self.page_count = max(1, (total_records + page_size - 1) // page_size)

        if self.current_page > self.page_count:
            self.current_page = self.page_count

        offset = max(0, (self.current_page - 1) * page_size)
        rows = self.local_store.get_history_page(limit=page_size, offset=offset)

        for row in rows:
            self.history_tree.insert(
                "",
                "end",
                values=(
                    row["id"],
                    row["codigo"],
                    row["descripcion"],
                    self._format_stock(row["stock"]),
                    row["fecha"],
                    self._build_time_label(row),
                    self._row_state(row),
                ),
            )

        start_idx = 0 if total_records == 0 else offset + 1
        end_idx = min(offset + page_size, total_records)
        self.page_info_var.set(f"Página {self.current_page}/{self.page_count} | {start_idx}-{end_idx} de {total_records}")
        self.btn_prev.configure(state="normal" if self.current_page > 1 else "disabled")
        self.btn_next.configure(state="normal" if self.current_page < self.page_count else "disabled")

    def prev_page(self) -> None:
        if self.current_page > 1:
            self.current_page -= 1
            self.refresh_history(reset_page=False)

    def next_page(self) -> None:
        if self.current_page < self.page_count:
            self.current_page += 1
            self.refresh_history(reset_page=False)

    def _format_stock(self, value: Any) -> str:
        try:
            num = float(value)
            if num.is_integer():
                return str(int(num))
            return f"{num:.2f}"
        except Exception:
            return str(value)

    def _parse_export_date(self, value: str, end_of_day: bool = False) -> Optional[str]:
        value = value.strip()
        if not value:
            return None
        try:
            if end_of_day:
                parsed = datetime.strptime(value, "%Y-%m-%d") + timedelta(hours=23, minutes=59, seconds=59)
            else:
                parsed = datetime.strptime(value, "%Y-%m-%d")
            return parsed.strftime("%Y-%m-%d %H:%M:%S")
        except Exception as exc:
            raise ValueError(f"Fecha inválida: {value}. Usa formato YYYY-MM-DD") from exc

    def export_excel_range(self) -> None:
        try:
            start_text = self._parse_export_date(self.entry_export_start.get(), end_of_day=False)
            end_text = self._parse_export_date(self.entry_export_end.get(), end_of_day=True)
        except ValueError as exc:
            messagebox.showerror("Exportar Excel", str(exc))
            return

        rows = self.local_store.get_history_by_range(start_text, end_text)
        if not rows:
            messagebox.showinfo("Exportar Excel", "No hay registros en el rango indicado.")
            return

        suggested_name = f"historial_scans_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Guardar Excel",
            defaultextension=".xlsx",
            initialfile=suggested_name,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not save_path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Historial"
        ws.append([
            "ID",
            "Código",
            "Descripción",
            "Stock",
            "Fecha",
            "Tiempo restante (s)",
            "Estado",
            "Detalle",
        ])

        for row in rows:
            ws.append([
                int(row["id"]),
                str(row["codigo"]),
                str(row["descripcion"]),
                float(row["stock"]),
                str(row["fecha"]),
                self._remaining_seconds(row),
                self._row_state(row),
                str(row["anulado_motivo"] or row["sync_error"] or ""),
            ])

        for column in ws.columns:
            max_len = 0
            letter = column[0].column_letter
            for cell in column:
                max_len = max(max_len, len(str(cell.value or "")))
            ws.column_dimensions[letter].width = min(max_len + 2, 40)

        wb.save(save_path)
        self._set_status(f"Excel generado: {save_path}")
        messagebox.showinfo("Exportar Excel", f"Archivo generado correctamente:\n{save_path}")

    def _ask_password(self, title: str, prompt: str) -> Optional[str]:
        """Diálogo de contraseña con diseño mejorado (CTk nativo)."""
        result: List[Optional[str]] = [None]

        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.transient(self)
        dialog.configure(bg="#1a1a2e")
        dialog.withdraw()  # ocultar hasta posicionar

        # ── Centrar en la ventana principal ───────────────────────────
        w, h = 420, 320
        dialog.geometry(f"{w}x{h}")
        self.update_idletasks()
        dialog.update_idletasks()
        px = self.winfo_rootx() + (self.winfo_width() - w) // 2
        py = self.winfo_rooty() + (self.winfo_height() - h) // 2
        dialog.geometry(f"{w}x{h}+{max(0, px)}+{max(0, py)}")
        dialog.deiconify()

        # ── Fondo principal ────────────────────────────────────────────
        main_frame = ctk.CTkFrame(dialog, fg_color="#1a1a2e", corner_radius=0)
        main_frame.pack(fill="both", expand=True)

        # ── Ícono de candado (canvas) ──────────────────────────────────
        icon_canvas = tk.Canvas(main_frame, width=64, height=72,
                                bg="#1a1a2e", highlightthickness=0)
        icon_canvas.pack(pady=(24, 0))
        # arco del candado
        icon_canvas.create_arc(14, 2, 50, 36, start=0, extent=180,
                               outline="#4DA8DA", width=4, style="arc")
        # cuerpo del candado
        icon_canvas.create_rectangle(8, 30, 56, 64,
                                     fill="#0F3460", outline="#4DA8DA", width=2)
        # ojo
        icon_canvas.create_oval(26, 40, 38, 52, fill="#4DA8DA", outline="")
        # hendidura
        icon_canvas.create_rectangle(30, 46, 34, 58, fill="#1a1a2e", outline="")

        # ── Título y subtítulo ─────────────────────────────────────────
        ctk.CTkLabel(
            main_frame, text=title,
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#E0E0E0",
        ).pack(pady=(10, 2))

        ctk.CTkLabel(
            main_frame, text=prompt,
            font=ctk.CTkFont(size=12),
            text_color="#9E9E9E",
        ).pack(pady=(0, 14))

        # ── Campo contraseña + ojo ─────────────────────────────────────
        pw_frame = ctk.CTkFrame(main_frame, fg_color="#0F3460", corner_radius=10)
        pw_frame.pack(padx=36, fill="x")
        pw_frame.grid_columnconfigure(0, weight=1)

        entry_pw = ctk.CTkEntry(
            pw_frame, show="●", placeholder_text="Contraseña",
            font=ctk.CTkFont(size=14),
            fg_color="transparent", border_width=0,
            text_color="#E0E0E0",
            height=44,
        )
        entry_pw.grid(row=0, column=0, sticky="ew", padx=(12, 0))

        _show_pw = [False]

        def toggle_visibility() -> None:
            _show_pw[0] = not _show_pw[0]
            entry_pw.configure(show="" if _show_pw[0] else "●")
            btn_eye.configure(text="🙈" if _show_pw[0] else "👁")

        btn_eye = ctk.CTkButton(
            pw_frame, text="👁", width=38, height=38,
            fg_color="transparent", hover_color="#1a1a2e",
            text_color="#9E9E9E", font=ctk.CTkFont(size=16),
            command=toggle_visibility,
        )
        btn_eye.grid(row=0, column=1, padx=(0, 4))

        # ── Mensaje de error ───────────────────────────────────────────
        lbl_error = ctk.CTkLabel(
            main_frame, text="",
            font=ctk.CTkFont(size=11),
            text_color="#EF5350",
        )
        lbl_error.pack(pady=(6, 0))

        # ── Botones ────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_row.pack(pady=(10, 0), padx=36, fill="x")
        btn_row.grid_columnconfigure((0, 1), weight=1)

        def _shake() -> None:
            ox = dialog.winfo_x()
            oy = dialog.winfo_y()
            for dx in (8, -8, 6, -6, 4, -4, 0):
                dialog.geometry(f"+{ox + dx}+{oy}")
                dialog.update()
                time.sleep(0.03)

        def confirm() -> None:
            value = entry_pw.get()
            if not value:
                lbl_error.configure(text="⚠  Ingresa la contraseña")
                _shake()
                return
            result[0] = value
            dialog.destroy()

        def cancel() -> None:
            dialog.destroy()

        entry_pw.bind("<Return>", lambda _e: confirm())

        ctk.CTkButton(
            btn_row, text="Cancelar", command=cancel,
            fg_color="#263238", hover_color="#37474F",
            text_color="#B0BEC5", corner_radius=8, height=38,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))

        ctk.CTkButton(
            btn_row, text="Ingresar", command=confirm,
            fg_color="#0F3460", hover_color="#1565C0",
            text_color="white", corner_radius=8, height=38,
            font=ctk.CTkFont(weight="bold"),
        ).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        entry_pw.focus_set()
        dialog.wait_window()
        return result[0]

    def ensure_config_password_exists(self) -> None:
        """Si no hay hash guardado, inicializa con la clave por defecto '0000'."""
        cfg = self.config_manager.get()
        if not str(cfg.get("config_password_hash", "")).strip():
            self.config_manager.update({"config_password_hash": hash_password("0000")})

    def verify_config_password(self) -> bool:
        """Muestra un único diálogo con error inline y hasta 3 intentos."""
        self.ensure_config_password_exists()
        cfg = self.config_manager.get()
        expected_hash = str(cfg.get("config_password_hash", "")).strip()
        verified: List[bool] = [False]
        MAX_ATTEMPTS = 3
        attempts: List[int] = [0]

        dialog = tk.Toplevel(self)
        dialog.title("Acceso a Configuración")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.transient(self)
        dialog.configure(bg="#1a1a2e")
        dialog.withdraw()

        w, h = 420, 340
        dialog.geometry(f"{w}x{h}")
        self.update_idletasks()
        dialog.update_idletasks()
        px = self.winfo_rootx() + (self.winfo_width() - w) // 2
        py = self.winfo_rooty() + (self.winfo_height() - h) // 2
        dialog.geometry(f"{w}x{h}+{max(0, px)}+{max(0, py)}")
        dialog.deiconify()

        main_frame = ctk.CTkFrame(dialog, fg_color="#1a1a2e", corner_radius=0)
        main_frame.pack(fill="both", expand=True)

        # ── Ícono candado ──────────────────────────────────────────────
        icon_canvas = tk.Canvas(main_frame, width=64, height=72,
                                bg="#1a1a2e", highlightthickness=0)
        icon_canvas.pack(pady=(22, 0))
        icon_canvas.create_arc(14, 2, 50, 36, start=0, extent=180,
                               outline="#4DA8DA", width=4, style="arc")
        icon_canvas.create_rectangle(8, 30, 56, 64,
                                     fill="#0F3460", outline="#4DA8DA", width=2)
        icon_canvas.create_oval(26, 40, 38, 52, fill="#4DA8DA", outline="")
        icon_canvas.create_rectangle(30, 46, 34, 58, fill="#1a1a2e", outline="")

        ctk.CTkLabel(
            main_frame, text="Área protegida",
            font=ctk.CTkFont(size=18, weight="bold"), text_color="#E0E0E0",
        ).pack(pady=(8, 2))
        ctk.CTkLabel(
            main_frame, text="Ingresa la contraseña para acceder a Configuración",
            font=ctk.CTkFont(size=11), text_color="#9E9E9E",
        ).pack(pady=(0, 12))

        # ── Campo contraseña + ojo ─────────────────────────────────────
        pw_frame = ctk.CTkFrame(main_frame, fg_color="#0F3460", corner_radius=10)
        pw_frame.pack(padx=36, fill="x")
        pw_frame.grid_columnconfigure(0, weight=1)

        entry_pw = ctk.CTkEntry(
            pw_frame, show="●", placeholder_text="Contraseña",
            font=ctk.CTkFont(size=14), fg_color="transparent",
            border_width=0, text_color="#E0E0E0", height=44,
        )
        entry_pw.grid(row=0, column=0, sticky="ew", padx=(12, 0))

        _show = [False]

        def toggle_vis() -> None:
            _show[0] = not _show[0]
            entry_pw.configure(show="" if _show[0] else "●")
            btn_eye.configure(text="🙈" if _show[0] else "👁")

        btn_eye = ctk.CTkButton(
            pw_frame, text="👁", width=38, height=38,
            fg_color="transparent", hover_color="#1a1a2e",
            text_color="#9E9E9E", font=ctk.CTkFont(size=16),
            command=toggle_vis,
        )
        btn_eye.grid(row=0, column=1, padx=(0, 4))

        # ── Error / intentos ───────────────────────────────────────────
        lbl_error = ctk.CTkLabel(
            main_frame, text="",
            font=ctk.CTkFont(size=11), text_color="#EF5350",
        )
        lbl_error.pack(pady=(6, 0))

        # ── Botones ────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_row.pack(pady=(10, 0), padx=36, fill="x")
        btn_row.grid_columnconfigure((0, 1), weight=1)

        def _shake() -> None:
            ox, oy = dialog.winfo_x(), dialog.winfo_y()
            for dx in (10, -10, 7, -7, 4, -4, 0):
                dialog.geometry(f"+{ox + dx}+{oy}")
                dialog.update()
                time.sleep(0.03)

        def confirm() -> None:
            value = entry_pw.get()
            if not value:
                lbl_error.configure(text="⚠  Ingresa la contraseña")
                _shake()
                return
            attempts[0] += 1
            remaining = MAX_ATTEMPTS - attempts[0]
            if hash_password(value) == expected_hash:
                verified[0] = True
                dialog.destroy()
            else:
                entry_pw.delete(0, tk.END)
                if remaining > 0:
                    lbl_error.configure(
                        text=f"✖  Contraseña incorrecta  —  {remaining} intento{'s' if remaining != 1 else ''} restante{'s' if remaining != 1 else ''}"
                    )
                    _shake()
                else:
                    lbl_error.configure(text="✖  Demasiados intentos fallidos")
                    btn_ok.configure(state="disabled")
                    _shake()

        def cancel() -> None:
            dialog.destroy()

        entry_pw.bind("<Return>", lambda _e: confirm())

        ctk.CTkButton(
            btn_row, text="Cancelar", command=cancel,
            fg_color="#263238", hover_color="#37474F",
            text_color="#B0BEC5", corner_radius=8, height=38,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))

        btn_ok = ctk.CTkButton(
            btn_row, text="Ingresar", command=confirm,
            fg_color="#0F3460", hover_color="#1565C0",
            text_color="white", corner_radius=8, height=38,
            font=ctk.CTkFont(weight="bold"),
        )
        btn_ok.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        entry_pw.focus_set()
        dialog.wait_window()
        return verified[0]

    def change_config_password(self) -> None:
        cfg = self.config_manager.get()
        current_hash = str(cfg.get("config_password_hash", "")).strip()
        current = self.entry_password_current.get()
        new = self.entry_password_new.get()
        confirm = self.entry_password_confirm.get()

        if current_hash and hash_password(current) != current_hash:
            messagebox.showerror("Contraseña", "La contraseña actual no es correcta.")
            return
        if not new:
            messagebox.showerror("Contraseña", "Debes escribir la nueva contraseña.")
            return
        if new != confirm:
            messagebox.showerror("Contraseña", "La confirmación no coincide.")
            return

        self.config_manager.update({"config_password_hash": hash_password(new)})
        self.entry_password_current.delete(0, tk.END)
        self.entry_password_new.delete(0, tk.END)
        self.entry_password_confirm.delete(0, tk.END)
        self._set_status("Contraseña de configuración actualizada")
        messagebox.showinfo("Contraseña", "Contraseña actualizada correctamente.")

    def request_open_config(self) -> None:
        self.tabview.set("Configuración")

    def _on_tab_changed(self) -> None:
        if self._tab_guard_active:
            self._tab_guard_active = False
            return
        if self.tabview.get() == "Configuración":
            self.after(10, self._guard_config_access)

    def _guard_config_access(self) -> None:
        if self.tabview.get() != "Configuración":
            return
        if self.verify_config_password():
            return
        self._tab_guard_active = True
        self.tabview.set("Historial")

    def hide_to_tray(self) -> None:
        if self.exiting:
            return
        self.withdraw()
        self.window_hidden = True
        self._set_status("Aplicación minimizada a la bandeja")

    def restore_from_tray(self) -> None:
        self.deiconify()
        self.window_hidden = False
        self.state("normal")
        self.lift()
        self.focus_force()
        self._set_status("Ventana restaurada")

    def open_config_tab(self) -> None:
        self.restore_from_tray()
        self.request_open_config()

    def _process_ui_queue(self) -> None:
        while True:
            try:
                action, payload = self.ui_queue.get_nowait()
            except queue.Empty:
                break

            if action == "status":
                self._set_status(str(payload))
            elif action == "refresh_history":
                self.refresh_history(reset_page=False)
            elif action == "open_window":
                self.restore_from_tray()
            elif action == "open_config":
                self.open_config_tab()
            elif action == "exit_app":
                self.exit_app()
            elif action == "show_info":
                title, message = payload
                messagebox.showinfo(title, message)
            elif action == "show_error":
                title, message = payload
                messagebox.showerror(title, message)

        if not self.exiting:
            self.after(300, self._process_ui_queue)

    def exit_app(self) -> None:
        if self.exiting:
            return
        self.exiting = True
        self._set_status("Cerrando aplicación...")
        self.stop_event.set()
        self.scanner_reload_event.set()
        self.sync_wakeup_event.set()
        self.tray_controller.stop()
        self.after(200, self.destroy)


def main() -> None:
    start_hidden = "--hidden" in sys.argv
    app = ZebraCloudSyncApp(start_hidden=start_hidden)
    app.mainloop()


if __name__ == "__main__":
    main()