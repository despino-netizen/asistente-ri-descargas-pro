import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import queue
import time
import os
import sys
import math
import datetime
import json
import re
import uuid
import winsound
import subprocess
import urllib.request
import urllib.error

# InstalaciÃ³n automÃ¡tica del driver
from webdriver_manager.chrome import ChromeDriverManager

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ConfiguraciÃ³n Global de Estilo
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

def _get_app_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _get_app_data_dir():
    local_appdata = os.getenv("LOCALAPPDATA", "").strip()
    if local_appdata:
        return os.path.join(local_appdata, "AsistenteRIDescargasPro")
    return os.path.join(_get_app_base_dir(), "_user_data")


def _get_default_download_dir():
    profile_dir = os.path.expanduser("~")
    downloads_dir = os.path.join(profile_dir, "Downloads")
    if not os.path.isdir(downloads_dir):
        downloads_dir = os.path.join(profile_dir, "Documents")
    return os.path.join(downloads_dir, "Descargas_PDF_RI")


APP_BASE_DIR = _get_app_base_dir()
APP_DATA_DIR = _get_app_data_dir()
CONFIG_FILE = os.path.join(APP_DATA_DIR, "config_descargas.json")
LOG_FILE = os.path.join(APP_DATA_DIR, "historial_actividad.txt")
DEFAULT_DOWNLOAD_DIR = _get_default_download_dir()
APP_NAME = "Asistente RI Descargas Pro"
APP_VERSION = "3.3.5"
APP_VERSION_LABEL = f"V{APP_VERSION}"
DEFAULT_EXE_NAME = "AsistenteRIDescargasPro.exe"
DEFAULT_PORTABLE_DIR_NAME = "AsistenteRIDescargasPro"
DEFAULT_PORTABLE_ZIP_PREFIX = "AsistenteRIDescargasPro_portable_v"
DEFAULT_INSTALLER_PREFIX = "Instalar_AsistenteRIDescargasPro_v"
GITHUB_RELEASE_REPOSITORY = "despino-netizen/asistente-ri-descargas-pro"
GITHUB_RELEASE_ASSET_NAME = DEFAULT_EXE_NAME
GITHUB_API_VERSION = "2026-03-10"

# Theme visual
APP_BG = "#F4EFE8"
CARD_BG = "#FFFDF8"
CARD_BORDER = "#E8DDD0"
SURFACE_SOFT = "#F2E8DA"
SURFACE_SOFT_2 = "#E9DCC8"
PRIMARY = "#1D5A55"
PRIMARY_HOVER = "#164641"
SUCCESS = "#2E7B58"
SUCCESS_HOVER = "#265F46"
WARN = "#C88B31"
WARN_HOVER = "#A67224"
DANGER = "#B14C38"
DANGER_HOVER = "#903C2C"
TEXT_MAIN = "#1E2B2A"
TEXT_MUTED = "#6F6A63"
INPUT_BG = "#FBF7F0"
LOG_BG = "#13201F"
LOG_BORDER = "#2C4441"


class GobiernoPDFDownloader(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- ConfiguraciÃ³n de la Ventana ---
        self.title(f"{APP_NAME} | {APP_VERSION_LABEL}")
        self.geometry("1320x860")
        self.minsize(1120, 720)
        self.configure(fg_color=APP_BG)

        # --- Icono de la ventana ---
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        
        # Variables de control
        self.driver = None
        self.is_running = False
        self.is_paused = False
        self.log_queue = queue.Queue()
        self.stats = {"success": 0, "error": 0}
        self.status_text = "Esperando inicio"
        self.failed_rows = []
        self.run_download_dir = None
        self.browser_download_dir = None
        self.update_check_in_progress = False
        self.update_prompted_version = ""
        self.update_install_pending = False
        self.update_feedback_window = None
        self.update_feedback_title = None
        self.update_feedback_message = None
        self.update_feedback_detail = None
        self.update_feedback_progress = None
        self.update_feedback_percent = None
        
        # Cargar configuraciÃ³n guardada (Memoria)
        self.saved_config = self.load_config()
        default_dir = self._normalize_folder_path(self.saved_config.get("last_folder", DEFAULT_DOWNLOAD_DIR))
        self._initial_download_dir = default_dir
        
        self.download_dir = tk.StringVar(value=default_dir)

        # ConfiguraciÃ³n de Grid Principal
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Estado de animaciones de interfaz
        self._chip_mode = "idle"
        self._chip_pulse_tick = 0
        self._accent_tick = 0
        self._resize_after_id = None
        self._compact_layout = None

        self._setup_ui()
        self._run_intro_reveal()
        self._animate_status_chip()
        self._animate_background_accents()
        
        # Iniciar loop de logs
        self.after(100, self._process_log_queue)
        self.bind("<Configure>", self._handle_window_resize)
        self.after(180, self._apply_responsive_layout)
        self.after(1800, self.start_update_check_thread)
        
        # Cierre seguro
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def load_config(self):
        """Carga la configuraciÃ³n desde un archivo JSON local."""
        try:
            os.makedirs(APP_DATA_DIR, exist_ok=True)
        except Exception:
            pass
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    config = json.load(f)
                if not isinstance(config, dict):
                    return {}
                config["last_folder"] = self._normalize_folder_path(config.get("last_folder", DEFAULT_DOWNLOAD_DIR))
                return config
            except Exception:
                return {}
        return {}

    def _normalize_folder_path(self, folder_path):
        folder_path = str(folder_path or "").strip()
        if not folder_path:
            folder_path = DEFAULT_DOWNLOAD_DIR
        folder_path = os.path.expandvars(os.path.expanduser(folder_path))
        return os.path.abspath(folder_path)

    def _format_path_preview(self, folder_path, max_length=74):
        normalized = self._normalize_folder_path(folder_path)
        if len(normalized) <= max_length:
            return normalized

        drive, rest = os.path.splitdrive(normalized)
        tail_length = max(28, max_length - len(drive) - 7)
        tail = rest[-tail_length:].lstrip("\\/")
        return f"{drive}\\...\\{tail}" if drive else f"...\\{tail}"

    def _refresh_path_preview(self):
        preview_text = self._format_path_preview(self.download_dir.get())
        helper_text = "La carpeta se guarda automaticamente y se recuerda al volver a abrir."
        try:
            self.lbl_dir.configure(text=preview_text)
        except Exception:
            pass
        try:
            self.lbl_subtitle.configure(text=helper_text)
        except Exception:
            pass

    def save_config(self):
        """Guarda la configuraciÃ³n actual."""
        current_folder = self._normalize_folder_path(self.download_dir.get())
        self.download_dir.set(current_folder)
        config = dict(self.saved_config) if isinstance(self.saved_config, dict) else {}
        config["last_folder"] = current_folder
        try:
            os.makedirs(APP_DATA_DIR, exist_ok=True)
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.saved_config = dict(config)
        except Exception:
            pass

    def _set_download_base_dir(self, folder_path, persist=True, apply_to_browser=False):
        normalized_dir = self._normalize_folder_path(folder_path)
        os.makedirs(normalized_dir, exist_ok=True)
        self.run_download_dir = None
        self.download_dir.set(normalized_dir)
        self._refresh_path_preview()

        if persist:
            self.save_config()

        browser_updated = False
        if apply_to_browser and self.driver and not self.is_running:
            browser_updated = self._apply_runtime_download_dir(normalized_dir)

        return normalized_dir, browser_updated

    def _wait_if_paused(self):
        while self.is_paused:
            if not self.is_running:
                return False
            time.sleep(0.2)
        return self.is_running

    def _wait_if_paused_with_elapsed(self):
        paused_elapsed = 0.0
        while self.is_paused:
            if not self.is_running:
                return False, paused_elapsed
            tick_start = time.monotonic()
            time.sleep(0.2)
            paused_elapsed += time.monotonic() - tick_start
        return self.is_running, paused_elapsed

    def _sleep_with_pause(self, seconds, step=0.1):
        end_time = time.monotonic() + max(0.0, float(seconds))
        while time.monotonic() < end_time:
            if not self._wait_if_paused():
                return False
            remaining = end_time - time.monotonic()
            if remaining <= 0:
                break
            time.sleep(min(step, remaining))
        return self.is_running

    def _sleep_with_pause_and_elapsed(self, seconds, step=0.1):
        paused_elapsed = 0.0
        end_time = time.monotonic() + max(0.0, float(seconds))
        while time.monotonic() < end_time:
            can_continue, paused_delta = self._wait_if_paused_with_elapsed()
            paused_elapsed += paused_delta
            if not can_continue:
                return False, paused_elapsed
            remaining = end_time - time.monotonic()
            if remaining <= 0:
                break
            time.sleep(min(step, remaining))
        return self.is_running, paused_elapsed

    def _should_persist_config_on_close(self):
        current_dir = str(self.download_dir.get() or "").strip()
        if not current_dir:
            return False

        saved_dir = str(self.saved_config.get("last_folder", "") or "").strip()
        if saved_dir:
            return current_dir != saved_dir

        # Sin config previa: no crear archivo si no hubo cambio real.
        return current_dir != str(self._initial_download_dir or "").strip()

    def _setup_ui(self):
        # Layout general: sidebar izquierda + contenido principal.
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.configure(fg_color=APP_BG)

        # ================= SIDEBAR =================
        self.sidebar_frame = ctk.CTkFrame(
            self,
            width=170,
            corner_radius=0,
            fg_color="#F4F7FC",
            border_width=1,
            border_color="#D9E3F1",
        )
        self.sidebar_frame.grid(row=0, column=0, sticky="ns")
        self.sidebar_frame.grid_propagate(False)
        self.sidebar_frame.grid_rowconfigure(6, weight=1)

        self.sidebar_brand = ctk.CTkFrame(
            self.sidebar_frame,
            fg_color="#123B73",
            corner_radius=0,
            height=86,
        )
        self.sidebar_brand.grid(row=0, column=0, sticky="ew")
        self.sidebar_brand.grid_propagate(False)

        self.lbl_sidebar_title = ctk.CTkLabel(
            self.sidebar_brand,
            text="Asistente RI\nDescargas Pro",
            font=("Segoe UI", 21, "bold"),
            text_color="#F2F7FF",
            justify="left",
        )
        self.lbl_sidebar_title.pack(anchor="w", padx=14, pady=14)

        self.nav_dashboard = ctk.CTkButton(
            self.sidebar_frame,
            text="Dashboard",
            height=36,
            fg_color="#E8F0FB",
            hover_color="#D8E7FA",
            text_color="#2F67B9",
            corner_radius=8,
            font=("Segoe UI", 14, "bold"),
        )
        self.nav_dashboard.grid(row=1, column=0, sticky="ew", padx=10, pady=(12, 6))

        self.nav_history = ctk.CTkButton(
            self.sidebar_frame,
            text="History",
            height=34,
            fg_color="transparent",
            hover_color="#EAF1FC",
            text_color="#5F6F87",
            corner_radius=8,
            font=("Segoe UI", 13, "bold"),
            anchor="w",
        )
        self.nav_history.grid(row=2, column=0, sticky="ew", padx=10, pady=2)

        self.nav_settings = ctk.CTkButton(
            self.sidebar_frame,
            text="Settings",
            height=34,
            fg_color="transparent",
            hover_color="#EAF1FC",
            text_color="#5F6F87",
            corner_radius=8,
            font=("Segoe UI", 13, "bold"),
            anchor="w",
        )
        self.nav_settings.grid(row=3, column=0, sticky="ew", padx=10, pady=2)

        self.sidebar_hint = ctk.CTkLabel(
            self.sidebar_frame,
            text="Panel principal",
            font=("Segoe UI", 10, "bold"),
            text_color="#93A2B8",
        )
        self.sidebar_hint.grid(row=7, column=0, padx=10, pady=(6, 10), sticky="w")

        # ================= CONTENIDO PRINCIPAL =================
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)

        # Fondo con orbes suaves para evitar plano uniforme.
        self.bg_orb_left = ctk.CTkFrame(
            self.main_frame,
            width=280,
            height=280,
            corner_radius=140,
            fg_color="#DCE8FA",
        )
        self.bg_orb_left.place(x=-130, y=270)
        self.bg_orb_left.lower()

        self.bg_orb_right = ctk.CTkFrame(
            self.main_frame,
            width=340,
            height=340,
            corner_radius=170,
            fg_color="#E7EEFB",
        )
        self.bg_orb_right.place(relx=1.0, x=-80, y=520, anchor="ne")
        self.bg_orb_right.lower()

        # ================= CARD: PATH =================
        self.header_frame = ctk.CTkFrame(
            self.main_frame,
            corner_radius=10,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
        )
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=0)

        self.lbl_title = ctk.CTkLabel(
            self.header_frame,
            text="Path Configuration",
            font=("Segoe UI", 33, "bold"),
            text_color=TEXT_MAIN,
        )
        self.lbl_title.grid(row=0, column=0, columnspan=2, padx=16, pady=(12, 4), sticky="w")

        self.lbl_subtitle = ctk.CTkLabel(
            self.header_frame,
            text="Carpeta destino:",
            font=("Segoe UI", 13, "bold"),
            text_color=TEXT_MAIN,
        )
        self.lbl_subtitle.grid(row=1, column=0, columnspan=2, padx=16, pady=(0, 6), sticky="w")

        # Placeholder para conservar referencia heredada.
        self.lbl_dir = ctk.CTkLabel(
            self.header_frame,
            text="",
            font=("Segoe UI", 1),
            text_color=CARD_BG,
        )
        self.lbl_dir.grid(row=1, column=1, padx=0, pady=0, sticky="e")

        self.entry_dir = ctk.CTkEntry(
            self.header_frame,
            textvariable=self.download_dir,
            height=38,
            fg_color=INPUT_BG,
            border_color="#CDD9EA",
            border_width=1,
            text_color="#223249",
            font=("Consolas", 12),
            corner_radius=6,
        )
        self.entry_dir.grid(row=2, column=0, padx=(16, 8), pady=(0, 12), sticky="ew")

        self.btn_folder = ctk.CTkButton(
            self.header_frame,
            text="Cambiar",
            command=self._select_folder,
            width=92,
            height=36,
            fg_color="#3E7FD1",
            hover_color="#2F68AF",
            text_color="#FFFFFF",
            corner_radius=8,
            font=("Segoe UI", 13, "bold"),
        )
        self.btn_folder.grid(row=2, column=1, padx=(0, 14), pady=(0, 12), sticky="e")

        self.header_activity = ctk.CTkProgressBar(
            self.header_frame,
            height=4,
            corner_radius=999,
            progress_color="#3E7FD1",
            fg_color="#DCE6F5",
        )
        self.header_activity.grid(row=3, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 8))
        self.header_activity.configure(mode="indeterminate")
        self.header_activity.stop()

        # ================= CARD: CONTROLES =================
        self.controls_frame = ctk.CTkFrame(
            self.main_frame,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
            corner_radius=10,
        )
        self.controls_frame.grid(row=1, column=0, sticky="ew", padx=16, pady=(4, 8))
        self.controls_frame.grid_columnconfigure(0, weight=1)

        self.metrics_row = ctk.CTkFrame(self.controls_frame, fg_color="transparent")
        self.metrics_row.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))
        self.metrics_row.grid_columnconfigure((0, 1, 2), weight=1)

        self.console_frame = ctk.CTkFrame(self.controls_frame, fg_color="transparent")
        self.console_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))
        self.console_frame.grid_columnconfigure(0, weight=1)

        self.progress_shell = ctk.CTkFrame(self.console_frame, fg_color="transparent")
        self.progress_shell.grid(row=0, column=0, sticky="ew")
        self.progress_shell.grid_columnconfigure(0, weight=1)

        self.progress_ring_host = ctk.CTkFrame(
            self.progress_shell,
            fg_color="transparent",
            corner_radius=0,
        )
        self.progress_ring_host.grid(row=0, column=0, pady=(4, 8))

        self.progress_ring_canvas = tk.Canvas(
            self.progress_ring_host,
            width=220,
            height=220,
            bd=0,
            highlightthickness=0,
            bg=CARD_BG,
        )
        self.progress_ring_canvas.pack()
        self.progress_ring_canvas.create_oval(20, 20, 200, 200, outline="#D5E1F2", width=16)
        self.progress_ring_arc = self.progress_ring_canvas.create_arc(
            20,
            20,
            200,
            200,
            start=90,
            extent=0,
            style="arc",
            outline="#3E7FD1",
            width=16,
        )
        self.progress_ring_canvas.create_text(
            110,
            86,
            text="Descargando...",
            fill="#5A6B85",
            font=("Segoe UI", 11, "bold"),
        )
        self.progress_ring_text = self.progress_ring_canvas.create_text(
            110,
            124,
            text="0%",
            fill="#1D2B3E",
            font=("Segoe UI", 22, "bold"),
        )

        self.action_subframe = ctk.CTkFrame(self.progress_shell, fg_color="transparent")
        self.action_subframe.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 6))
        self.action_subframe.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        btn_font = ("Segoe UI", 13, "bold")
        btn_height = 36

        self.btn_browser = ctk.CTkButton(
            self.action_subframe,
            text="Abrir navegador",
            command=self.launch_browser,
            height=btn_height,
            corner_radius=7,
            font=btn_font,
            fg_color="#3E7FD1",
            hover_color="#2F68AF",
            text_color="#FFFFFF",
        )
        self.btn_browser.grid(row=0, column=0, padx=4, pady=6, sticky="ew")

        self.btn_start = ctk.CTkButton(
            self.action_subframe,
            text="Empezar descarga",
            command=self.start_scraping_thread,
            state="disabled",
            height=btn_height,
            corner_radius=7,
            font=btn_font,
            fg_color="#3E7FD1",
            hover_color="#2F68AF",
            text_color="#FFFFFF",
        )
        self.btn_start.grid(row=0, column=1, padx=4, pady=6, sticky="ew")

        self.btn_pause = ctk.CTkButton(
            self.action_subframe,
            text="Pausar",
            command=self.toggle_pause,
            state="disabled",
            height=btn_height,
            corner_radius=7,
            font=btn_font,
            fg_color="#3E7FD1",
            hover_color="#2F68AF",
            text_color="#FFFFFF",
        )
        self.btn_pause.grid(row=0, column=2, padx=4, pady=6, sticky="ew")

        self.btn_cancel = ctk.CTkButton(
            self.action_subframe,
            text="Cancelar",
            command=self.stop_process,
            state="disabled",
            height=btn_height,
            corner_radius=7,
            font=btn_font,
            fg_color="#3E7FD1",
            hover_color="#2F68AF",
            text_color="#FFFFFF",
        )
        self.btn_cancel.grid(row=0, column=3, padx=4, pady=6, sticky="ew")

        self.btn_retry_failed = ctk.CTkButton(
            self.action_subframe,
            text="Reintentar fallidos (0)",
            command=self.retry_failed_rows_thread,
            state="disabled",
            height=btn_height,
            corner_radius=7,
            font=btn_font,
            fg_color="#3E7FD1",
            hover_color="#2F68AF",
            text_color="#FFFFFF",
        )
        self.btn_retry_failed.grid(row=0, column=4, padx=4, pady=6, sticky="ew")

        # Barra lineal para mantener la logica interna actual.
        self.progress = ctk.CTkProgressBar(
            self.progress_shell,
            height=10,
            corner_radius=999,
            progress_color="#3E7FD1",
            fg_color="#DCE6F5",
        )
        self.progress.set(0)
        self.progress.grid(row=2, column=0, padx=10, pady=(2, 6), sticky="ew")

        self.lbl_runtime_status = ctk.CTkLabel(
            self.progress_shell,
            text="LISTO",
            font=("Segoe UI", 12, "bold"),
            fg_color="#2D67B2",
            corner_radius=999,
            padx=12,
            pady=4,
            text_color="#FFFFFF",
        )
        self.lbl_runtime_status.grid(row=3, column=0, pady=(0, 8))

        self.lbl_success = ctk.CTkLabel(
            self.metrics_row,
            text="Exitos: 0",
            font=("Segoe UI", 13, "bold"),
            text_color="#2A9B54",
            fg_color="#EDF8F1",
            corner_radius=8,
            padx=8,
            pady=6,
        )
        self.lbl_success.grid(row=0, column=0, padx=4, sticky="ew")

        self.lbl_error = ctk.CTkLabel(
            self.metrics_row,
            text="Errores: 0",
            font=("Segoe UI", 13, "bold"),
            text_color="#C0392B",
            fg_color="#FCEEEE",
            corner_radius=8,
            padx=8,
            pady=6,
        )
        self.lbl_error.grid(row=0, column=1, padx=4, sticky="ew")

        self.lbl_status = ctk.CTkLabel(
            self.metrics_row,
            text="Estado: Esperando inicio",
            font=("Segoe UI", 13, "bold"),
            text_color="#2F67B9",
            fg_color="#EAF1FC",
            corner_radius=8,
            padx=8,
            pady=6,
            justify="left",
        )
        self.lbl_status.grid(row=0, column=2, padx=4, sticky="ew")

        self.warning_frame = ctk.CTkFrame(
            self.controls_frame,
            fg_color="#F6FAFF",
            border_color="#D8E4F2",
            border_width=1,
            corner_radius=8,
        )
        self.warning_frame.grid(row=2, column=0, padx=12, pady=(0, 10), sticky="ew")
        instrucciones = (
            "INSTRUCCIONES RAPIDAS: 1) Inicie sesion y ubique la tabla. "
            "2) Coloque la tabla en 100 registros."
        )
        self.lbl_warning = ctk.CTkLabel(
            self.warning_frame,
            text=instrucciones,
            justify="left",
            font=("Segoe UI", 11, "bold"),
            text_color="#4F627C",
            wraplength=900,
        )
        self.lbl_warning.pack(pady=7, padx=10, anchor="w")

        # ================= CARD: LOG =================
        self.log_frame = ctk.CTkFrame(
            self.main_frame,
            corner_radius=10,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
        )
        self.log_frame.grid(row=2, column=0, sticky="nsew", padx=16, pady=(6, 8))
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)

        self.log_header = ctk.CTkFrame(
            self.log_frame,
            fg_color="#3E7FD1",
            corner_radius=8,
            border_width=0,
            height=34,
        )
        self.log_header.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 6))
        self.log_header.grid_propagate(False)
        ctk.CTkLabel(
            self.log_header,
            text="Live Log",
            font=("Segoe UI", 13, "bold"),
            text_color="#FFFFFF",
        ).pack(anchor="w", padx=10, pady=6)

        self.log_area = ctk.CTkTextbox(
            self.log_frame,
            font=("Consolas", 11),
            state="disabled",
            fg_color="#FFFFFF",
            text_color="#1E2E44",
            border_color="#D3DEEC",
            border_width=1,
        )
        self.log_area.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))

        # ================= FOOTER =================
        self.footer_frame = ctk.CTkFrame(self.main_frame, height=24, fg_color="transparent")
        self.footer_frame.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 8))
        self.footer_frame.grid_columnconfigure(0, weight=1)
        self.footer_frame.grid_columnconfigure(1, weight=0)
        self.footer_frame.grid_columnconfigure(2, weight=0)

        self.lbl_credits = ctk.CTkLabel(
            self.footer_frame,
            text="Desarrollado por Dariel Espino (809) 434-2700",
            font=("Segoe UI", 10, "bold"),
            text_color="#5A6B85",
        )
        self.lbl_credits.grid(row=0, column=0, sticky="w")

        self.btn_update = ctk.CTkButton(
            self.footer_frame,
            text="Buscar actualización",
            command=lambda: self.start_update_check_thread(manual=True),
            width=146,
            height=24,
            fg_color="#EAF1FC",
            hover_color="#D8E7FA",
            text_color="#2F67B9",
            corner_radius=999,
            font=("Segoe UI", 10, "bold"),
        )
        self.btn_update.grid(row=0, column=1, padx=(0, 10), sticky="e")

        self.lbl_version = ctk.CTkLabel(
            self.footer_frame,
            text=APP_VERSION_LABEL,
            font=("Segoe UI", 10, "bold"),
            text_color="#5A6B85",
        )
        self.lbl_version.grid(row=0, column=2, sticky="e")

        self._refresh_stats_labels("Esperando inicio")
        self._update_retry_button_state()
        self.log("Bienvenido. Configuracion cargada correctamente.", "INFO", persist=False)
        self._sync_progress_ring()

    def _sync_progress_ring(self):
        value = 0.0
        try:
            value = float(self.progress.get())
        except Exception:
            value = 0.0

        value = max(0.0, min(1.0, value))
        extent = -360 * value
        percent = int(round(value * 100))

        try:
            self.progress_ring_canvas.itemconfigure(self.progress_ring_arc, extent=extent)
            self.progress_ring_canvas.itemconfigure(self.progress_ring_text, text=f"{percent}%")
        except Exception:
            pass

        self.after(140, self._sync_progress_ring)

    # -------------------------------------------------------------------------
    # FUNCIONES LÃ“GICAS
    # -------------------------------------------------------------------------

    def _run_intro_reveal(self):
        reveal_sequence = [
            self.header_frame,
            self.controls_frame,
            self.log_frame,
            self.footer_frame,
        ]
        for widget in reveal_sequence:
            widget.grid_remove()

        for idx, widget in enumerate(reveal_sequence):
            self.after(80 + (idx * 95), lambda w=widget: w.grid())

    def _animate_background_accents(self):
        self._accent_tick += 1
        phase = self._accent_tick / 14.0

        left_x = -130 + int(8 * math.sin(phase))
        left_y = 270 + int(6 * math.cos(phase * 0.8))
        self.bg_orb_left.place(x=left_x, y=left_y)

        right_x = -80 + int(12 * math.cos(phase * 0.9))
        right_y = 520 + int(10 * math.sin(phase))
        self.bg_orb_right.place(relx=1.0, x=right_x, y=right_y, anchor="ne")

        self.after(85, self._animate_background_accents)

    def _animate_status_chip(self):
        self._chip_pulse_tick += 1

        if self._chip_mode == "running":
            palette = ["#2E70C4", "#3E84D8", "#337ACF"]
        elif self._chip_mode == "paused":
            palette = ["#376CAF", "#4682CC", "#3C75BC"]
        elif self._chip_mode == "error":
            palette = ["#B84036", "#CF4E43", "#BF463C"]
        elif self._chip_mode == "done":
            palette = ["#278D4E", "#2FA65C"]
        else:
            palette = ["#2D67B2", "#3A79C7"]

        color = palette[self._chip_pulse_tick % len(palette)]
        try:
            self.lbl_runtime_status.configure(fg_color=color)
        except Exception:
            pass

        self.after(450, self._animate_status_chip)

    def _refresh_stats_labels(self, status_text=None):
        if status_text is not None:
            self.status_text = status_text

        self.lbl_success.configure(text=str(self.stats['success']))
        self.lbl_error.configure(text=str(self.stats['error']))
        self.lbl_status.configure(text=self.status_text)

        chip_text = "LISTO"
        self._chip_mode = "idle"
        if self.is_running and self.is_paused:
            chip_text = "PAUSADO"
            self._chip_mode = "paused"
        elif self.is_running:
            chip_text = "EN PROCESO"
            self._chip_mode = "running"
        elif self.stats["error"] > 0 and self.stats["success"] == 0:
            chip_text = "CON ERRORES"
            self._chip_mode = "error"
        elif self.stats["success"] > 0:
            chip_text = "COMPLETADO"
            self._chip_mode = "done"

        self.lbl_runtime_status.configure(text=chip_text)

    def _show_info_threadsafe(self, title, message):
        self.after(0, lambda t=title, m=message: messagebox.showinfo(t, m))

    def _show_error_threadsafe(self, title, message):
        self.after(0, lambda t=title, m=message: messagebox.showerror(t, m))

    def _ask_yes_no_threadsafe(self, title, message):
        response = {"value": False}
        dialog_done = threading.Event()

        def _ask():
            try:
                response["value"] = messagebox.askyesno(title, message)
            finally:
                dialog_done.set()

        self.after(0, _ask)
        dialog_done.wait()
        return bool(response["value"])

    def _set_update_button_state(self, checking):
        def _apply():
            try:
                if not self.winfo_exists():
                    return
                self.btn_update.configure(
                    state="disabled" if checking else "normal",
                    text="Buscando..." if checking else "Buscar actualización",
                )
            except Exception:
                pass

        try:
            self.after(0, _apply)
        except Exception:
            pass

    def _format_progress_bytes(self, total_bytes):
        try:
            total = float(total_bytes)
        except Exception:
            total = 0.0

        units = ["B", "KB", "MB", "GB"]
        unit_index = 0
        while total >= 1024 and unit_index < len(units) - 1:
            total /= 1024.0
            unit_index += 1

        decimals = 0 if unit_index == 0 else 1
        return f"{total:.{decimals}f} {units[unit_index]}"

    def _ensure_update_feedback_window(self):
        if self.update_feedback_window and self.update_feedback_window.winfo_exists():
            return self.update_feedback_window

        update_window = ctk.CTkToplevel(self)
        update_window.title("Actualización en curso")
        update_window.geometry("430x220")
        update_window.resizable(False, False)
        update_window.transient(self)
        update_window.attributes("-topmost", True)
        update_window.protocol("WM_DELETE_WINDOW", lambda: None)
        update_window.configure(fg_color=APP_BG)
        update_window.grid_columnconfigure(0, weight=1)
        update_window.grid_rowconfigure(0, weight=1)

        self.update_idletasks()
        try:
            pos_x = self.winfo_x() + max(40, int((self.winfo_width() - 430) / 2))
            pos_y = self.winfo_y() + max(40, int((self.winfo_height() - 220) / 2))
            update_window.geometry(f"430x220+{pos_x}+{pos_y}")
        except Exception:
            pass

        shell = ctk.CTkFrame(
            update_window,
            corner_radius=22,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
        )
        shell.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)
        shell.grid_columnconfigure(0, weight=1)

        self.update_feedback_title = ctk.CTkLabel(
            shell,
            text="Actualización en curso",
            font=("Segoe UI", 20, "bold"),
            text_color=TEXT_MAIN,
        )
        self.update_feedback_title.grid(row=0, column=0, padx=18, pady=(18, 6), sticky="w")

        self.update_feedback_message = ctk.CTkLabel(
            shell,
            text="Preparando descarga...",
            font=("Segoe UI", 12),
            text_color=TEXT_MUTED,
            justify="left",
            wraplength=360,
        )
        self.update_feedback_message.grid(row=1, column=0, padx=18, sticky="w")

        self.update_feedback_detail = ctk.CTkLabel(
            shell,
            text="No cierres la aplicación mientras se actualiza.",
            font=("Segoe UI", 11, "bold"),
            text_color="#6C6A65",
            justify="left",
            wraplength=360,
        )
        self.update_feedback_detail.grid(row=2, column=0, padx=18, pady=(10, 10), sticky="w")

        self.update_feedback_progress = ctk.CTkProgressBar(
            shell,
            height=12,
            corner_radius=999,
            progress_color=PRIMARY,
            fg_color="#E4D8CA",
        )
        self.update_feedback_progress.grid(row=3, column=0, padx=18, pady=(0, 8), sticky="ew")
        self.update_feedback_progress.configure(mode="indeterminate")
        self.update_feedback_progress.start()

        self.update_feedback_percent = ctk.CTkLabel(
            shell,
            text="Procesando...",
            font=("Segoe UI", 11, "bold"),
            text_color=PRIMARY,
        )
        self.update_feedback_percent.grid(row=4, column=0, padx=18, pady=(0, 16), sticky="e")

        self.update_feedback_window = update_window
        return update_window

    def _set_update_feedback(
        self,
        title=None,
        message=None,
        detail=None,
        progress=None,
        indeterminate=None,
    ):
        def _apply():
            if not self.winfo_exists():
                return

            window = self._ensure_update_feedback_window()
            if not window or not window.winfo_exists():
                return

            if title is not None and self.update_feedback_title:
                self.update_feedback_title.configure(text=title)
                try:
                    window.title(title)
                except Exception:
                    pass

            if message is not None and self.update_feedback_message:
                self.update_feedback_message.configure(text=message)

            if detail is not None and self.update_feedback_detail:
                self.update_feedback_detail.configure(text=detail)

            if self.update_feedback_progress:
                use_indeterminate = bool(indeterminate) if indeterminate is not None else progress is None
                if use_indeterminate:
                    self.update_feedback_progress.stop()
                    self.update_feedback_progress.configure(mode="indeterminate")
                    self.update_feedback_progress.start()
                    if self.update_feedback_percent:
                        self.update_feedback_percent.configure(text="Procesando...")
                else:
                    normalized = max(0.0, min(1.0, float(progress or 0.0)))
                    self.update_feedback_progress.stop()
                    self.update_feedback_progress.configure(mode="determinate")
                    self.update_feedback_progress.set(normalized)
                    if self.update_feedback_percent:
                        self.update_feedback_percent.configure(text=f"{int(normalized * 100)}%")

            try:
                window.deiconify()
                window.lift()
                window.focus_force()
            except Exception:
                pass

            try:
                self.header_activity.start()
            except Exception:
                pass

        try:
            self.after(0, _apply)
        except Exception:
            pass

    def _close_update_feedback(self):
        def _apply():
            window = self.update_feedback_window
            if window and window.winfo_exists():
                try:
                    window.destroy()
                except Exception:
                    pass

            self.update_feedback_window = None
            self.update_feedback_title = None
            self.update_feedback_message = None
            self.update_feedback_detail = None
            self.update_feedback_progress = None
            self.update_feedback_percent = None

            if not self.is_running:
                try:
                    self.header_activity.stop()
                except Exception:
                    pass

        try:
            self.after(0, _apply)
        except Exception:
            pass

    def _normalize_version_tag(self, version_text):
        version = str(version_text or "").strip()
        return re.sub(r"^[Vv]\s*", "", version)

    def _version_key(self, version_text):
        normalized = self._normalize_version_tag(version_text)
        parts = [int(part) for part in re.findall(r"\d+", normalized)]
        return tuple(parts) if parts else (0,)

    def _current_exe_path(self):
        if not getattr(sys, "frozen", False):
            return ""
        return os.path.abspath(sys.executable)

    def _installed_app_dir(self):
        local_appdata = os.getenv("LOCALAPPDATA", "").strip()
        if not local_appdata:
            return ""
        return os.path.abspath(os.path.join(local_appdata, "Programs", "AsistenteRIDescargasPro"))

    def _is_running_from_installed_location(self):
        exe_path = self._current_exe_path()
        if not exe_path:
            return False

        current_dir = os.path.abspath(os.path.dirname(exe_path))
        installed_dir = self._installed_app_dir()
        if installed_dir and os.path.normcase(current_dir) == os.path.normcase(installed_dir):
            return True

        return os.path.exists(os.path.join(current_dir, "unins000.exe"))

    def _expected_direct_update_asset_name(self):
        current_name = os.path.basename(self._current_exe_path())
        candidates = [
            current_name,
            str(self.saved_config.get("github_asset_name", "") or "").strip(),
            GITHUB_RELEASE_ASSET_NAME,
            DEFAULT_EXE_NAME,
        ]
        for candidate in candidates:
            candidate = str(candidate or "").strip()
            if candidate:
                return candidate
        return DEFAULT_EXE_NAME

    def _expected_installer_asset_name(self, version_text=""):
        normalized = self._normalize_version_tag(version_text or APP_VERSION)
        return f"{DEFAULT_INSTALLER_PREFIX}{normalized}.exe"

    def _expected_portable_zip_asset_name(self, version_text=""):
        normalized = self._normalize_version_tag(version_text or APP_VERSION)
        return f"{DEFAULT_PORTABLE_ZIP_PREFIX}{normalized}.zip"

    def _expected_update_asset_name(self, version_text=""):
        if self._is_running_from_installed_location():
            return self._expected_portable_zip_asset_name(version_text)
        return self._expected_direct_update_asset_name()

    def _configured_github_repo(self):
        repo = str(self.saved_config.get("github_repo", "") or "").strip().strip("/")
        if repo:
            return repo
        return str(GITHUB_RELEASE_REPOSITORY or "").strip().strip("/")

    def _powershell_literal(self, value):
        return "'" + str(value).replace("'", "''") + "'"

    def _is_installer_asset_name(self, asset_name):
        normalized = str(asset_name or "").strip().lower()
        return normalized.startswith(DEFAULT_INSTALLER_PREFIX.lower()) and normalized.endswith(".exe")

    def _is_portable_zip_asset_name(self, asset_name):
        normalized = str(asset_name or "").strip().lower()
        return normalized.startswith(DEFAULT_PORTABLE_ZIP_PREFIX.lower()) and normalized.endswith(".zip")

    def _update_staging_dir(self):
        staging_dir = os.path.join(APP_DATA_DIR, "updates")
        os.makedirs(staging_dir, exist_ok=True)
        return staging_dir

    def _choose_release_asset(self, assets, latest_version=""):
        if not isinstance(assets, list):
            return None

        normalized_assets = [asset for asset in assets if isinstance(asset, dict)]

        def find_exact_name_match(preferred_names):
            for preferred_name in preferred_names:
                normalized_name = str(preferred_name or "").strip().lower()
                if not normalized_name:
                    continue
                for asset in normalized_assets:
                    asset_name = str(asset.get("name", "") or "").strip().lower()
                    if asset_name == normalized_name and asset.get("browser_download_url"):
                        return asset
            return None

        direct_preferred_names = [
            os.path.basename(self._current_exe_path()),
            str(self.saved_config.get("github_asset_name", "") or "").strip(),
            GITHUB_RELEASE_ASSET_NAME,
            DEFAULT_EXE_NAME,
        ]
        portable_preferred_names = [
            self._expected_portable_zip_asset_name(latest_version),
        ]
        installer_preferred_names = [
            self._expected_installer_asset_name(latest_version),
            str(self.saved_config.get("github_asset_name", "") or "").strip(),
        ]

        if self._is_running_from_installed_location():
            portable_asset = find_exact_name_match(portable_preferred_names)
            if portable_asset:
                return portable_asset

            for asset in normalized_assets:
                asset_name = str(asset.get("name", "") or "").strip()
                if self._is_portable_zip_asset_name(asset_name) and asset.get("browser_download_url"):
                    return asset

            installer_asset = find_exact_name_match(installer_preferred_names)
            if installer_asset:
                return installer_asset

            for asset in normalized_assets:
                asset_name = str(asset.get("name", "") or "").strip()
                if self._is_installer_asset_name(asset_name) and asset.get("browser_download_url"):
                    return asset

        direct_asset = find_exact_name_match(direct_preferred_names)
        if direct_asset:
            return direct_asset

        for asset in normalized_assets:
            asset_name = str(asset.get("name", "") or "").strip()
            if (
                asset.get("browser_download_url")
                and asset_name.lower().endswith(".exe")
                and not self._is_installer_asset_name(asset_name)
            ):
                return asset

        portable_asset = find_exact_name_match(portable_preferred_names)
        if portable_asset:
            return portable_asset

        for asset in normalized_assets:
            asset_name = str(asset.get("name", "") or "").strip()
            if self._is_portable_zip_asset_name(asset_name) and asset.get("browser_download_url"):
                return asset

        installer_asset = find_exact_name_match(installer_preferred_names)
        if installer_asset:
            return installer_asset

        for asset in normalized_assets:
            asset_name = str(asset.get("name", "") or "").strip()
            if self._is_installer_asset_name(asset_name) and asset.get("browser_download_url"):
                return asset

        return None

    def start_update_check_thread(self, manual=False):
        repo = self._configured_github_repo()
        if self.update_check_in_progress:
            if manual:
                messagebox.showinfo("Actualización", "Ya hay una verificación en curso.")
            return

        if self.is_running:
            if manual:
                messagebox.showinfo(
                    "Proceso en curso",
                    "Busque actualizaciones cuando termine la descarga actual.",
                )
            return

        if not repo:
            if manual:
                messagebox.showinfo(
                    "Configurar GitHub",
                    "Defina el repositorio en GITHUB_RELEASE_REPOSITORY o agregue "
                    "'github_repo' en config_descargas.json con formato 'usuario/repositorio' "
                    "antes de usar las actualizaciones automáticas.",
                )
            return

        if not getattr(sys, "frozen", False):
            if manual:
                messagebox.showinfo(
                    "Modo desarrollo",
                    "La autoactualización funciona desde el .exe compilado. "
                    "Pruebe esta función desde el ejecutable.",
                )
            return

        self.update_check_in_progress = True
        self.update_install_pending = False
        self._set_update_button_state(checking=True)
        threading.Thread(
            target=self._check_for_updates,
            args=(manual,),
            daemon=True,
        ).start()

    def _check_for_updates(self, manual=False):
        try:
            if manual:
                self._set_update_feedback(
                    title="Buscando actualización",
                    message="Consultando la versión más reciente en GitHub.",
                    detail="Esto suele tardar solo unos segundos.",
                    progress=None,
                    indeterminate=True,
                )

            repo = self._configured_github_repo()
            current_version = self._normalize_version_tag(APP_VERSION)

            api_url = f"https://api.github.com/repos/{repo}/releases/latest"
            request = urllib.request.Request(
                api_url,
                headers={
                    "Accept": "application/vnd.github+json",
                    "X-GitHub-Api-Version": GITHUB_API_VERSION,
                    "User-Agent": APP_NAME,
                },
            )

            with urllib.request.urlopen(request, timeout=30) as response:
                release_info = json.loads(response.read().decode("utf-8"))

            latest_version = self._normalize_version_tag(
                release_info.get("tag_name") or release_info.get("name") or ""
            )
            if not latest_version:
                raise RuntimeError("GitHub no devolvió una versión válida.")

            if self._version_key(latest_version) <= self._version_key(current_version):
                self._close_update_feedback()
                if manual:
                    self._show_info_threadsafe(
                        "Actualización",
                        f"Ya tienes la última versión instalada.\n\nActual: V{current_version}",
                    )
                return

            asset = self._choose_release_asset(release_info.get("assets"), latest_version)
            if not asset:
                self._close_update_feedback()
                expected_name = self._expected_update_asset_name(latest_version)
                message = (
                    "Hay una versión nueva en GitHub, pero no se encontró un archivo "
                    f"compatible para actualizar.\n\nSube a Releases un asset llamado '{expected_name}'."
                )
                if manual:
                    self._show_info_threadsafe("Actualización", message)
                else:
                    self.log(message, "WARN")
                return

            if not manual and self.update_prompted_version == latest_version:
                return

            self.update_prompted_version = latest_version
            asset_name = str(asset.get("name") or self._expected_update_asset_name()).strip()
            self._close_update_feedback()
            should_update = self._ask_yes_no_threadsafe(
                "Actualización disponible",
                "Hay una nueva versión disponible.\n\n"
                f"Actual: V{current_version}\n"
                f"Disponible: V{latest_version}\n"
                f"Archivo: {asset_name}\n\n"
                "La aplicación se cerrará para instalar la actualización.\n"
                "¿Desea actualizar ahora?",
            )
            if not should_update:
                self._close_update_feedback()
                self.log(f"Actualización pospuesta por el usuario: V{latest_version}", "INFO")
                return

            self._download_and_apply_update(release_info, asset)

        except urllib.error.HTTPError as exc:
            self.update_install_pending = False
            self._close_update_feedback()
            message = f"No se pudo consultar GitHub Releases (HTTP {exc.code})."
            if manual:
                self._show_error_threadsafe("Actualización", message)
            else:
                self.log(message, "WARN")
        except urllib.error.URLError as exc:
            self.update_install_pending = False
            self._close_update_feedback()
            message = f"No se pudo conectar con GitHub: {exc.reason}"
            if manual:
                self._show_error_threadsafe("Actualización", message)
            else:
                self.log(message, "WARN")
        except Exception as exc:
            self.update_install_pending = False
            self._close_update_feedback()
            message = f"Error buscando actualización: {exc}"
            if manual:
                self._show_error_threadsafe("Actualización", message)
            else:
                self.log(message, "WARN")
        finally:
            self.update_check_in_progress = False
            self._set_update_button_state(checking=False)
            if not self.update_install_pending:
                self._close_update_feedback()

    def _create_binary_swap_update_script(self, downloaded_exe_path, target_exe_path):
        script_path = os.path.join(
            self._update_staging_dir(),
            f"ri_auto_update_{uuid.uuid4().hex[:8]}.ps1",
        )
        script_content = f"""$ErrorActionPreference = 'SilentlyContinue'
$pidToWait = {os.getpid()}
$sourceExe = {self._powershell_literal(downloaded_exe_path)}
$targetExe = {self._powershell_literal(target_exe_path)}
$tempTarget = "$targetExe.new"
$backupTarget = "$targetExe.bak"
$deadline = (Get-Date).AddMinutes(5)

while ((Get-Date) -lt $deadline) {{
    $proc = Get-Process -Id $pidToWait -ErrorAction SilentlyContinue
    if (-not $proc) {{
        break
    }}
    Start-Sleep -Seconds 1
}}

Start-Sleep -Milliseconds 900
$sourceSize = 0
if (Test-Path -LiteralPath $sourceExe) {{
    try {{
        $sourceSize = (Get-Item -LiteralPath $sourceExe).Length
    }} catch {{
        $sourceSize = 0
    }}
}}
$updated = $false
for ($i = 0; $i -lt 12; $i++) {{
    try {{
        if (Test-Path -LiteralPath $tempTarget) {{
            Remove-Item -LiteralPath $tempTarget -Force
        }}
        Copy-Item -LiteralPath $sourceExe -Destination $tempTarget -Force
        $copiedSize = (Get-Item -LiteralPath $tempTarget).Length
        if ($sourceSize -le 0 -or $copiedSize -lt $sourceSize) {{
            throw "La copia temporal no quedó completa."
        }}
        if (Test-Path -LiteralPath $backupTarget) {{
            Remove-Item -LiteralPath $backupTarget -Force
        }}
        if (Test-Path -LiteralPath $targetExe) {{
            Move-Item -LiteralPath $targetExe -Destination $backupTarget -Force
        }}
        Move-Item -LiteralPath $tempTarget -Destination $targetExe -Force
        $updated = $true
        break
    }} catch {{
        if (Test-Path -LiteralPath $tempTarget) {{
            Remove-Item -LiteralPath $tempTarget -Force
        }}
        Start-Sleep -Seconds 1
    }}
}}

if ($updated) {{
    Start-Sleep -Seconds 2
    Start-Process -FilePath $targetExe
    Start-Sleep -Seconds 2
    if (Test-Path -LiteralPath $backupTarget) {{
        Remove-Item -LiteralPath $backupTarget -Force
    }}
}}

if (Test-Path -LiteralPath $sourceExe) {{
    Remove-Item -LiteralPath $sourceExe -Force
}}
Remove-Item -LiteralPath $PSCommandPath -Force
"""
        with open(script_path, "w", encoding="utf-8") as script_file:
            script_file.write(script_content)
        return script_path

    def _create_installer_update_script(self, installer_path, target_exe_path):
        script_path = os.path.join(
            self._update_staging_dir(),
            f"ri_installer_update_{uuid.uuid4().hex[:8]}.ps1",
        )
        script_content = f"""$ErrorActionPreference = 'SilentlyContinue'
$pidToWait = {os.getpid()}
$installerPath = {self._powershell_literal(installer_path)}
$targetExe = {self._powershell_literal(target_exe_path)}
$deadline = (Get-Date).AddMinutes(5)
$installArgs = @('/SILENT', '/SUPPRESSMSGBOXES', '/NOCANCEL', '/SP-')

while ((Get-Date) -lt $deadline) {{
    $proc = Get-Process -Id $pidToWait -ErrorAction SilentlyContinue
    if (-not $proc) {{
        break
    }}
    Start-Sleep -Seconds 1
}}

Start-Sleep -Milliseconds 900
$installed = $false
for ($i = 0; $i -lt 6; $i++) {{
    try {{
        $installerProc = Start-Process -FilePath $installerPath -ArgumentList $installArgs -PassThru
        $installerProc.WaitForExit()
        if ($installerProc.ExitCode -eq 0 -or $installerProc.ExitCode -eq 3010) {{
            $installed = $true
            break
        }}
    }} catch {{
        Start-Sleep -Seconds 2
    }}
}}

if ($installed -and (Test-Path -LiteralPath $targetExe)) {{
    Start-Sleep -Seconds 2
    Start-Process -FilePath $targetExe
}}

if (Test-Path -LiteralPath $installerPath) {{
    Remove-Item -LiteralPath $installerPath -Force
}}
Remove-Item -LiteralPath $PSCommandPath -Force
"""
        with open(script_path, "w", encoding="utf-8") as script_file:
            script_file.write(script_content)
        return script_path

    def _create_portable_update_script(self, zip_path, target_dir):
        script_path = os.path.join(
            self._update_staging_dir(),
            f"ri_portable_update_{uuid.uuid4().hex[:8]}.ps1",
        )
        script_content = f"""$ErrorActionPreference = 'Stop'
$pidToWait = {os.getpid()}
$zipPath = {self._powershell_literal(zip_path)}
$targetDir = {self._powershell_literal(target_dir)}
$appExe = Join-Path $targetDir {self._powershell_literal(DEFAULT_EXE_NAME)}
$parentDir = Split-Path -Parent $targetDir
$extractRoot = Join-Path $parentDir ('ri_update_' + [guid]::NewGuid().ToString('N'))
$backupDir = "$targetDir.old"
$logPath = Join-Path {self._powershell_literal(self._update_staging_dir())} 'portable_update_last.log'
$deadline = (Get-Date).AddMinutes(5)
$updateSucceeded = $false

function Write-UpdateLog($message) {{
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -LiteralPath $logPath -Value "[$stamp] $message"
}}

try {{
    Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue
}} catch {{}}

Write-UpdateLog "Inicio de actualización portable."

while ((Get-Date) -lt $deadline) {{
    $proc = Get-Process -Id $pidToWait -ErrorAction SilentlyContinue
    if (-not $proc) {{
        break
    }}
    Start-Sleep -Seconds 1
}}

Start-Sleep -Milliseconds 900

try {{
    Write-UpdateLog "Extrayendo paquete: $zipPath"
    if (Test-Path -LiteralPath $extractRoot) {{
        Remove-Item -LiteralPath $extractRoot -Recurse -Force
    }}
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force

    $expandedDir = Join-Path $extractRoot {self._powershell_literal(DEFAULT_PORTABLE_DIR_NAME)}
    if (-not (Test-Path -LiteralPath $expandedDir)) {{
        $expandedDir = $extractRoot
    }}
    Write-UpdateLog "Contenido extraído en: $expandedDir"

    if (Test-Path -LiteralPath $backupDir) {{
        Remove-Item -LiteralPath $backupDir -Recurse -Force
    }}

    if (Test-Path -LiteralPath $targetDir) {{
        Write-UpdateLog "Moviendo instalación actual a respaldo: $backupDir"
        Move-Item -LiteralPath $targetDir -Destination $backupDir -Force
    }}

    Write-UpdateLog "Moviendo nueva versión a: $targetDir"
    Move-Item -LiteralPath $expandedDir -Destination $targetDir -Force

    foreach ($name in @('unins000.exe', 'unins000.dat', 'unins000.msg')) {{
        $sourceFile = Join-Path $backupDir $name
        $targetFile = Join-Path $targetDir $name
        if ((Test-Path -LiteralPath $sourceFile) -and -not (Test-Path -LiteralPath $targetFile)) {{
            Copy-Item -LiteralPath $sourceFile -Destination $targetFile -Force
        }}
    }}

    if (Test-Path -LiteralPath $appExe) {{
        Write-UpdateLog "Lanzando nueva aplicación: $appExe"
        $updateSucceeded = $true
        Start-Sleep -Seconds 1
        Start-Process -FilePath $appExe
    }}
}} catch {{
    Write-UpdateLog ("Error: " + ($_ | Out-String))
    if (Test-Path -LiteralPath $backupDir -and -not (Test-Path -LiteralPath $targetDir)) {{
        Write-UpdateLog "Restaurando respaldo anterior."
        Move-Item -LiteralPath $backupDir -Destination $targetDir -Force
    }}
    if (Test-Path -LiteralPath $appExe) {{
        Write-UpdateLog "Reabriendo la versión anterior tras el fallo."
        Start-Process -FilePath $appExe
    }}
}} finally {{
    if (Test-Path -LiteralPath $extractRoot) {{
        Remove-Item -LiteralPath $extractRoot -Recurse -Force
    }}
    if ($updateSucceeded) {{
        Write-UpdateLog "Actualización completada correctamente."
        if (Test-Path -LiteralPath $backupDir) {{
            Remove-Item -LiteralPath $backupDir -Recurse -Force
        }}
        if (Test-Path -LiteralPath $zipPath) {{
            Remove-Item -LiteralPath $zipPath -Force
        }}
        Remove-Item -LiteralPath $PSCommandPath -Force
    }} else {{
        Write-UpdateLog "La actualización no terminó. Se conservaron archivos para diagnóstico."
    }}
}}
"""
        with open(script_path, "w", encoding="utf-8") as script_file:
            script_file.write(script_content)
        return script_path

    def _shutdown_for_update(self, update_script_path):
        self._set_update_feedback(
            title="Iniciando actualización",
            message="Cerrando la aplicación para continuar la instalación.",
            detail="En unos segundos se aplicará la nueva versión.",
            progress=None,
            indeterminate=True,
        )
        try:
            self.update_idletasks()
            self.update()
        except Exception:
            pass
        try:
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.Popen(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-WindowStyle",
                    "Hidden",
                    "-File",
                    update_script_path,
                ],
                creationflags=creationflags,
            )
        except Exception as exc:
            self._show_error_threadsafe(
                "Actualización",
                f"No se pudo iniciar el instalador de actualización: {exc}",
            )
            return

        if self._should_persist_config_on_close():
            self.save_config()

        self.is_running = False
        self.is_paused = False

        if self.driver:
            try:
                self._set_browser_close_warning_enabled(False)
                self.driver.quit()
            except Exception:
                pass
            finally:
                self.driver = None

        self.destroy()

    def _download_and_apply_update(self, release_info, asset):
        exe_path = self._current_exe_path()
        if not exe_path or not os.path.exists(exe_path):
            raise RuntimeError("No se encontró el ejecutable actual para reemplazarlo.")

        asset_name = str(asset.get("name") or self._expected_update_asset_name()).strip()
        asset_url = str(asset.get("browser_download_url") or "").strip()
        if not asset_url:
            raise RuntimeError("El release no incluye una URL de descarga válida.")

        staging_dir = self._update_staging_dir()
        temp_exe_path = os.path.join(
            staging_dir,
            f"{uuid.uuid4().hex[:8]}_{re.sub(r'[^A-Za-z0-9._-]', '_', asset_name)}",
        )

        self.log(f"Descargando actualización {asset_name}...", "INFO")
        self._set_update_feedback(
            title="Descargando actualización",
            message=f"Descargando {asset_name}",
            detail="No cierres la aplicación mientras se completa la descarga.",
            progress=0,
            indeterminate=False,
        )
        request = urllib.request.Request(
            asset_url,
            headers={
                "Accept": "application/octet-stream",
                "User-Agent": APP_NAME,
            },
        )

        with urllib.request.urlopen(request, timeout=120) as response, open(temp_exe_path, "wb") as output_file:
            total_size = int(response.headers.get("Content-Length", "0") or "0")
            downloaded_size = 0
            last_feedback_at = 0.0
            while True:
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                output_file.write(chunk)
                downloaded_size += len(chunk)

                now = time.time()
                if total_size > 0:
                    progress_ratio = downloaded_size / float(total_size)
                    if now - last_feedback_at >= 0.10 or downloaded_size >= total_size:
                        self._set_update_feedback(
                            title="Descargando actualización",
                            message=f"Descargando {asset_name}",
                            detail=(
                                f"{self._format_progress_bytes(downloaded_size)} de "
                                f"{self._format_progress_bytes(total_size)}"
                            ),
                            progress=progress_ratio,
                            indeterminate=False,
                        )
                        last_feedback_at = now
                elif now - last_feedback_at >= 0.20:
                    self._set_update_feedback(
                        title="Descargando actualización",
                        message=f"Descargando {asset_name}",
                        detail=f"{self._format_progress_bytes(downloaded_size)} descargados",
                        progress=None,
                        indeterminate=True,
                    )
                    last_feedback_at = now

        if not os.path.exists(temp_exe_path) or os.path.getsize(temp_exe_path) <= 0:
            raise RuntimeError("La descarga de la actualización llegó vacía o incompleta.")

        latest_version = self._normalize_version_tag(
            release_info.get("tag_name") or release_info.get("name") or ""
        )
        if self._is_portable_zip_asset_name(asset_name):
            install_message = "La aplicación se cerrará y se reemplazará por la nueva versión automáticamente."
            install_log_message = "Se aplicará la actualización portable."
        elif self._is_installer_asset_name(asset_name):
            install_message = "La aplicación se cerrará y se abrirá el instalador con progreso visible."
            install_log_message = "Se aplicará la actualización con instalador."
        else:
            install_message = "La aplicación se cerrará para aplicar el reemplazo del ejecutable."
            install_log_message = "Se aplicará la actualización reemplazando el ejecutable."

        self._set_update_feedback(
            title="Preparando instalación",
            message=f"La versión V{latest_version} ya se descargó correctamente.",
            detail=install_message,
            progress=None,
            indeterminate=True,
        )
        self.log(f"Actualización V{latest_version} descargada. Cerrando para instalar...", "WARN")
        self.log(install_log_message, "INFO")
        self.update_install_pending = True

        if self._is_portable_zip_asset_name(asset_name):
            target_dir = os.path.dirname(exe_path)
            update_script_path = self._create_portable_update_script(temp_exe_path, target_dir)
        elif self._is_installer_asset_name(asset_name):
            update_script_path = self._create_installer_update_script(temp_exe_path, exe_path)
        else:
            update_script_path = self._create_binary_swap_update_script(temp_exe_path, exe_path)
        self.after(0, lambda p=update_script_path: self._shutdown_for_update(p))

    def _confirm_logout_before_app_close(self):
        if not self.driver:
            return True
        return messagebox.askokcancel(
            "Aviso importante",
            "ANTES DE CERRAR SAL DE LA CUENTA.\n\n"
            "Cierre sesion en el navegador y luego pulse Aceptar para salir.",
        )

    def _active_download_dir(self):
        return self.run_download_dir or self.download_dir.get()

    def _prepare_run_download_dir(self):
        base_dir = self._normalize_folder_path(self.download_dir.get())
        os.makedirs(base_dir, exist_ok=True)

        fecha = datetime.datetime.now().strftime("%Y%m%d")
        unique_id = uuid.uuid4().hex[:8].upper()
        folder_name = f"Trabajo_{fecha}_{unique_id}"
        run_dir = os.path.join(base_dir, folder_name)
        while os.path.exists(run_dir):
            unique_id = uuid.uuid4().hex[:8].upper()
            folder_name = f"Trabajo_{fecha}_{unique_id}"
            run_dir = os.path.join(base_dir, folder_name)

        os.makedirs(run_dir, exist_ok=True)
        self.run_download_dir = run_dir
        self.log(f"Sesion de descarga creada: {run_dir}", "INFO")
        return run_dir

    def _apply_runtime_download_dir(self, target_dir):
        if not self.driver:
            return False
        target_dir = self._normalize_folder_path(target_dir)
        os.makedirs(target_dir, exist_ok=True)

        commands = [
            (
                "Browser.setDownloadBehavior",
                {
                    "behavior": "allow",
                    "downloadPath": target_dir,
                    "eventsEnabled": False,
                },
            ),
            (
                "Page.setDownloadBehavior",
                {
                    "behavior": "allow",
                    "downloadPath": target_dir,
                },
            ),
        ]

        last_error = None
        for command_name, payload in commands:
            try:
                self.driver.execute_cdp_cmd(command_name, payload)
                self.browser_download_dir = target_dir
                return True
            except Exception as e:
                last_error = e

        self.log(f"No se pudo aplicar carpeta de descarga en navegador: {last_error}", "WARN")
        return False

    def _fallback_browser_download_dir(self):
        fallback_dir = self.browser_download_dir or self.download_dir.get() or DEFAULT_DOWNLOAD_DIR
        fallback_dir = self._normalize_folder_path(fallback_dir)
        os.makedirs(fallback_dir, exist_ok=True)
        self.browser_download_dir = fallback_dir
        return fallback_dir

    def _sync_browser_download_dir(self, target_dir):
        target_dir = self._normalize_folder_path(target_dir)
        if self._apply_runtime_download_dir(target_dir):
            return target_dir, True

        fallback_dir = self._fallback_browser_download_dir()
        if os.path.normcase(fallback_dir) != os.path.normcase(target_dir):
            self.log(
                f"El navegador mantiene la carpeta actual: {fallback_dir}. "
                "Si necesita usar otra carpeta, cierre y abra de nuevo el navegador.",
                "WARN",
            )
        return fallback_dir, False

    def _capture_failure_evidence(self, row_index, attempt_number, error_message):
        evidence_root = os.path.join(self._active_download_dir(), "_evidencias_fallos")
        try:
            os.makedirs(evidence_root, exist_ok=True)
        except Exception:
            return

        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"fila_{row_index:04d}_intento_{attempt_number}_{stamp}"
        screenshot_path = os.path.join(evidence_root, f"{base_name}.png")
        html_path = os.path.join(evidence_root, f"{base_name}.html")
        info_path = os.path.join(evidence_root, f"{base_name}.txt")

        current_url = ""
        try:
            current_url = self.driver.current_url if self.driver else ""
        except Exception:
            current_url = ""

        try:
            if self.driver:
                self.driver.save_screenshot(screenshot_path)
        except Exception:
            pass

        try:
            if self.driver:
                with open(html_path, "w", encoding="utf-8", errors="ignore") as f:
                    f.write(self.driver.page_source or "")
        except Exception:
            pass

        try:
            with open(info_path, "w", encoding="utf-8") as f:
                f.write(f"timestamp={datetime.datetime.now().isoformat()}\n")
                f.write(f"fila={row_index}\n")
                f.write(f"intento={attempt_number}\n")
                f.write(f"url={current_url}\n")
                f.write(f"error={error_message}\n")
        except Exception:
            pass

        self.log(f"Evidencia guardada para fila {row_index}: {base_name}", "WARN")

    def _update_retry_button_state(self):
        total_failed = len(sorted(set(self.failed_rows)))
        state = "normal" if (total_failed > 0 and not self.is_running and self.driver) else "disabled"
        self.btn_retry_failed.configure(text=f"Reintentar fallidos ({total_failed})", state=state)

    def _get_table_rows(self):
        try:
            # Prioriza la tabla principal de resultados para evitar filas de modales u otras vistas.
            rows = self.driver.find_elements(By.XPATH, "//table[@id='DataTables_Table_0']//tbody/tr[td]")
            if rows:
                return rows

            rows = self.driver.find_elements(By.XPATH, "//table[contains(@class,'dataTable')]//tbody/tr[td]")
            if rows:
                return rows
        except Exception:
            pass
        return self.driver.find_elements(By.CSS_SELECTOR, "tr")

    def _extract_verdetalle_parts(self, ver_button):
        try:
            onclick = (ver_button.get_attribute("onclick") or "").strip()
            if "VerDetalle" not in onclick:
                return {}

            args = re.findall(r"'([^']*)'", onclick)
            if len(args) < 3:
                return {}

            tipo = args[0].strip() or "Documento"
            docid = args[1].strip()
            uddocid = args[2].strip()
            key = f"{tipo}|{docid}|{uddocid}"
            return {"tipo": tipo, "docid": docid, "uddocid": uddocid, "key": key}
        except Exception:
            return {}

    def _normalize_document_type(self, document_type):
        return re.sub(r"\s+", " ", str(document_type or "").strip()).lower()

    def _extract_row_document_type(self, row_element, row_parts=None):
        try:
            cells = row_element.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 2:
                tipo_visible = re.sub(r"\s+", " ", (cells[1].text or "").strip())
                if tipo_visible:
                    return tipo_visible
        except Exception:
            pass

        if row_parts:
            tipo = re.sub(r"\s+", " ", str(row_parts.get("tipo", "")).strip())
            if tipo:
                return tipo
        return ""

    def _find_modal_view_buttons(self, timeout=8):
        modal_link_xpaths = [
            "//*[@id='modalVerDetalle' and not(contains(@style,'display: none'))]//tr[td]//*[self::a or self::button][contains(normalize-space(.), 'Ver') or contains(normalize-space(.), 'Abrir') or contains(@title, 'Ver') or contains(@title, 'Abrir')]",
            "//*[@id='modalVerDetalle']//tr[td]//*[self::a or self::button][contains(normalize-space(.), 'Ver') or contains(normalize-space(.), 'Abrir') or contains(@title, 'Ver') or contains(@title, 'Abrir')]",
            "(//div[contains(@class,'modal') and (contains(@class,'show') or contains(@class,'in') or not(contains(@style,'display: none')))]//tr[td]//*[self::a or self::button][contains(normalize-space(.), 'Ver') or contains(normalize-space(.), 'Abrir') or contains(@title, 'Ver') or contains(@title, 'Abrir')])",
            "(//a[contains(normalize-space(.), 'Ver') or contains(normalize-space(.), 'Abrir') or @target='_blank' or contains(@href,'SecureImage')] | //button[contains(normalize-space(.), 'Ver') or contains(normalize-space(.), 'Abrir')])",
        ]

        last_error = None
        for candidate_xpath in modal_link_xpaths:
            try:
                WebDriverWait(self.driver, timeout).until(
                    lambda d, xp=candidate_xpath: len(d.find_elements(By.XPATH, xp)) > 0
                )
                buttons = self.driver.find_elements(By.XPATH, candidate_xpath)
                visible_buttons = []
                for button in buttons:
                    try:
                        if button.is_displayed():
                            visible_buttons.append(button)
                    except Exception:
                        pass
                if visible_buttons:
                    return visible_buttons, candidate_xpath
            except TimeoutException as e:
                last_error = e

        raise TimeoutException("No se encontro enlace 'Ver' en modal Detalle") from last_error

    def _get_modal_detail_entries(self):
        try:
            entries = self.driver.execute_script(
                """
                function isVisible(el) {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    return style.display !== "none" && style.visibility !== "hidden";
                }

                const modal = document.getElementById("modalVerDetalle");
                const scope = (modal && isVisible(modal)) ? modal : document;
                const rows = Array.from(scope.querySelectorAll("tr"));
                const data = [];

                for (const row of rows) {
                    const link = row.querySelector("a, button");
                    if (!link) continue;

                    const label = `${link.textContent || ""} ${link.getAttribute("title") || ""}`.trim();
                    const href = (link.getAttribute("href") || "").trim();
                    const onclick = (link.getAttribute("onclick") || "").trim();
                    if (!/ver|abrir/i.test(label) && !href && !onclick) continue;

                    const cells = row.querySelectorAll("td");
                    let info = "";
                    if (cells.length >= 2) {
                        info = (cells[1].textContent || "").trim();
                    } else if (cells.length >= 1) {
                        info = (cells[0].textContent || "").trim();
                    }

                    data.push({
                        info: (info || `Detalle ${data.length + 1}`).replace(/\\s+/g, " ").trim()
                    });
                }

                return data;
                """
            )
            return entries if isinstance(entries, list) else []
        except Exception:
            return []

    def _build_unique_filename(self, row_index, ver_button, detail_index=None):
        parts = self._extract_verdetalle_parts(ver_button)
        short_id = uuid.uuid4().hex[:6].upper()
        row_tag = f"{row_index + 1:03d}"
        detail_tag = f"_D{detail_index + 1:02d}" if detail_index is not None else ""

        docid = parts.get("docid", "")
        uddocid = parts.get("uddocid", "")
        if docid:
            return f"T{row_tag}{detail_tag}_{docid}_{short_id}.pdf"
        if uddocid:
            return f"T{row_tag}{detail_tag}_{uddocid}_{short_id}.pdf"
        return f"T{row_tag}{detail_tag}_{short_id}.pdf"

    def _set_browser_close_warning_enabled(self, enabled):
        if not self.driver:
            return
        try:
            self.driver.execute_script(
                "window.__closeAccountWarnEnabled = arguments[0] ? true : false;",
                bool(enabled),
            )
        except Exception:
            pass

    def _install_browser_close_warning(self, enabled=False):
        if not self.driver:
            return

        warning_text = "ANTES DE CERRAR SAL DE LA CUENTA"
        enabled_js = "true" if enabled else "false"
        cdp_source = """
            (function() {
                const warningText = "ANTES DE CERRAR SAL DE LA CUENTA";
                const initialEnabled = __INITIAL_ENABLED__;
                window.__closeAccountWarnEnabled = initialEnabled;
                if (window.__closeAccountWarnInstalled) return;
                window.__closeAccountWarnInstalled = true;
                window.addEventListener("beforeunload", function (event) {
                    if (!window.__closeAccountWarnEnabled) return;
                    event.preventDefault();
                    event.returnValue = warningText;
                    return warningText;
                });
            })();
        """.replace("__INITIAL_ENABLED__", enabled_js)

        try:
            # Inyecta el handler en futuras navegaciones del tab activo.
            self.driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {"source": cdp_source},
            )
        except Exception:
            pass

        try:
            # Inyecta tambien en el documento actual.
            self.driver.execute_script(
                """
                (function(warningText, enabled) {
                    window.__closeAccountWarnEnabled = enabled ? true : false;
                    if (window.__closeAccountWarnInstalled) return;
                    window.__closeAccountWarnInstalled = true;
                    window.addEventListener("beforeunload", function (event) {
                        if (!window.__closeAccountWarnEnabled) return;
                        event.preventDefault();
                        event.returnValue = warningText;
                        return warningText;
                    });
                })(arguments[0], arguments[1]);
                """,
                warning_text,
                bool(enabled),
            )
        except Exception:
            pass

    def _get_modal_detail_signature(self):
        try:
            signature = self.driver.execute_script(
                """
                try {
                    const modal = document.getElementById("modalVerDetalle");
                    if (!modal) return "";

                    const link = modal.querySelector("a[href], button[onclick], button");
                    const href = link ? (link.getAttribute("href") || "") : "";
                    const onclick = link ? (link.getAttribute("onclick") || "") : "";
                    const text = link ? ((link.textContent || "").trim()) : "";
                    const display = modal.style && modal.style.display ? modal.style.display : "";
                    return [display, href, onclick, text].join("|");
                } catch (e) {
                    return "";
                }
                """
            )
            return str(signature or "").strip()
        except Exception:
            return ""

    def _select_folder(self):
        if self.is_running:
            messagebox.showinfo("Proceso en curso", "Cambie la carpeta cuando termine la descarga actual.")
            return

        folder = filedialog.askdirectory(initialdir=self._normalize_folder_path(self.download_dir.get()))
        if folder:
            selected_dir, browser_updated = self._set_download_base_dir(
                folder,
                persist=True,
                apply_to_browser=bool(self.driver),
            )
            if self.driver:
                if browser_updated:
                    self.log(f"Carpeta cambiada, guardada y aplicada al navegador: {selected_dir}", "SUCCESS")
                else:
                    self.log(f"Carpeta cambiada y guardada: {selected_dir}", "WARN")
            else:
                self.log(f"Carpeta cambiada y guardada: {selected_dir}", "INFO")

    def log(self, message, level="INFO", persist=True):
        self.log_queue.put((message, level))
        if not persist:
            return
        try:
            os.makedirs(APP_DATA_DIR, exist_ok=True)
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                f.write(f"[{timestamp}] [{level}] {message}\n")
        except: pass

    def _process_log_queue(self):
        try:
            while True:
                msg, level = self.log_queue.get_nowait()
                timestamp = time.strftime('%H:%M:%S')
                prefix = "[INFO] "
                if level == "SUCCESS": prefix = "[OK] "
                elif level == "ERROR": prefix = "[ERR] "
                elif level == "WARN": prefix = "[WARN] "
                elif level == "WAIT": prefix = "[WAIT] "
                final_msg = f"[{timestamp}] {prefix}{msg}\n"
                
                self.log_area.configure(state="normal")
                self.log_area.insert("end", final_msg)
                self.log_area.see("end")
                self.log_area.configure(state="disabled")
        except queue.Empty: pass
        finally: self.after(100, self._process_log_queue)

    def on_close(self):
        if self.is_running:
            if not messagebox.askokcancel("Salir", "El programa esta trabajando. Seguro que quiere salir?"):
                return
            self.is_running = False

        if not self._confirm_logout_before_app_close():
            return

        if self._should_persist_config_on_close():
            self.save_config()

        if self.driver:
            try:
                self._set_browser_close_warning_enabled(False)
                self.driver.quit()
            except:
                pass
        self.destroy()

    def stop_process(self):
        if self.is_running:
            self.is_running = False
            self.log("Deteniendo... (TerminarÃ¡ el archivo actual)", "WARN")
            self.btn_cancel.configure(text="Deteniendo...", state="disabled")
            self.btn_pause.configure(state="disabled")
            self.after(0, lambda: self._refresh_stats_labels("Deteniendo..."))

    def toggle_pause(self):
        if not self.is_running: return
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.btn_pause.configure(text="Continuar", fg_color=SUCCESS)
            self.log("PAUSA ACTIVADA. Presione 'Continuar'.", "WAIT")
            self.header_activity.stop()
            self._refresh_stats_labels("Pausado")
        else:
            self.btn_pause.configure(text="Pausar", fg_color=WARN)
            self.log("Continuando...", "INFO")
            self.header_activity.start()
            self._refresh_stats_labels("Procesando")

    def toggle_ui_state(self, working):
        if working:
            self._set_browser_close_warning_enabled(False)
            self.entry_dir.configure(state="disabled")
            self.btn_folder.configure(state="disabled")
            self.btn_browser.configure(state="disabled")
            self.btn_start.configure(state="disabled")
            self.btn_cancel.configure(state="normal", text="Cancelar", fg_color=DANGER)
            self.btn_pause.configure(state="normal", text="Pausar", fg_color=WARN)
            self.btn_retry_failed.configure(state="disabled")
            self.progress.configure(mode="determinate")
            self.progress.set(0)
            self.header_activity.start()
            self._refresh_stats_labels("Procesando")
        else:
            self.entry_dir.configure(state="normal")
            self.btn_folder.configure(state="normal")
            self.btn_browser.configure(state="normal")
            self.btn_start.configure(state="normal")
            self.btn_cancel.configure(state="disabled", text="Cancelar", fg_color=DANGER)
            self.btn_pause.configure(state="disabled")
            self.progress.configure(mode="determinate")
            self.header_activity.stop()
            if self.stats["success"] == 0 and self.stats["error"] == 0:
                self.progress.set(0)
                self._refresh_stats_labels("Esperando inicio")
            self._update_retry_button_state()
            self._set_browser_close_warning_enabled(True)

    def launch_browser(self):
        try:
            self.log("Abriendo navegador...", "INFO")
            selected_dir, _ = self._set_download_base_dir(
                self.download_dir.get(),
                persist=True,
                apply_to_browser=False,
            )
            if not os.path.exists(selected_dir):
                os.makedirs(selected_dir)

            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None

            options = Options()
            prefs = {
                "download.default_directory": selected_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
                "pdfjs.disabled": True,
                "profile.default_content_settings.popups": 0
            }
            options.add_experimental_option("prefs", prefs)
            options.add_argument("--start-maximized")
            options.add_argument("--disable-infobars")
             
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.set_page_load_timeout(120)
            self.driver.set_script_timeout(90)
            self.browser_download_dir = selected_dir
            self._apply_runtime_download_dir(selected_dir)
            self._install_browser_close_warning(enabled=False)
              
            self.log("Navegador listo.", "SUCCESS")
            self.driver.get("https://oficinavirtual.ri.gob.do/PH")
            self._install_browser_close_warning(enabled=False)
            
            self.log(">>> ESPERANDO: Inicie sesiÃ³n y navegue a la tabla.", "WARN")
            self.btn_start.configure(state="normal")
            self._refresh_stats_labels("Navegador listo")
            self._update_retry_button_state()
            
        except Exception as e:
            self.log(f"Error navegador: {str(e)}", "ERROR")
            messagebox.showerror("Error", f"No se pudo iniciar Chrome: {e}")

    def start_scraping_thread(self):
        self._start_scraping_mode(target_rows=None, retry_mode=False)

    def retry_failed_rows_thread(self):
        if not self.driver:
            messagebox.showwarning("Error", "Primero presione el BotÃ³n 1.")
            return
        if self.is_running:
            return

        retry_rows = sorted(set(self.failed_rows))
        if not retry_rows:
            messagebox.showinfo("Sin fallidos", "No hay filas fallidas para reintentar.")
            return

        self._start_scraping_mode(target_rows=retry_rows, retry_mode=True)

    def _start_scraping_mode(self, target_rows=None, retry_mode=False):
        if not self.driver:
            messagebox.showwarning("Error", "Primero presione el BotÃ³n 1.")
            return

        if retry_mode:
            rows_to_process = sorted(set(target_rows or self.failed_rows))
            if not rows_to_process:
                messagebox.showinfo("Sin fallidos", "No hay filas fallidas para reintentar.")
                return
            self.failed_rows = []
        else:
            rows_to_process = None
            self.failed_rows = []

        selected_dir, _ = self._set_download_base_dir(
            self.download_dir.get(),
            persist=True,
            apply_to_browser=True,
        )
        run_dir = self._prepare_run_download_dir()
        active_run_dir, run_dir_synced = self._sync_browser_download_dir(run_dir)
        self.run_download_dir = active_run_dir
        if not run_dir_synced and os.path.normcase(active_run_dir) != os.path.normcase(run_dir):
            self.log(f"Esta ejecucion usara la carpeta activa del navegador: {active_run_dir}", "WARN")
        elif os.path.normcase(active_run_dir) == os.path.normcase(selected_dir):
            self.log(f"Esta ejecucion usara la carpeta base seleccionada: {active_run_dir}", "INFO")

        self.stats = {"success": 0, "error": 0}
        self.is_running = True
        self.is_paused = False
        self.toggle_ui_state(working=True)
        self._refresh_stats_labels("Reintentando fallidos..." if retry_mode else "Iniciando...")

        t = threading.Thread(target=self._scraping_logic, args=(rows_to_process,))
        t.daemon = True
        t.start()

    def _is_valid_pdf_file(self, file_path, min_size=1024):
        try:
            if not os.path.exists(file_path):
                return False

            size = os.path.getsize(file_path)
            if size < min_size:
                return False

            with open(file_path, "rb") as f:
                header = f.read(5)
            return header == b"%PDF-"
        except:
            return False

    def _wait_for_download(self, initial_files, timeout=45, expected_name=None):
        ignored_files = set()
        expected_base = None
        stable_sizes = {}
        stable_counts = {}
        if expected_name:
            expected_base = os.path.splitext(expected_name)[0].strip().lower()

        end_time = time.monotonic() + timeout
        while time.monotonic() < end_time:
            can_continue, paused_elapsed = self._wait_if_paused_with_elapsed()
            end_time += paused_elapsed
            if not can_continue:
                return None
            try:
                active_dir = self._active_download_dir()
                current_files = set(os.listdir(active_dir))
                new_files = current_files - initial_files - ignored_files
                if new_files:
                    valid_files = [f for f in new_files if not f.endswith(".crdownload") and not f.endswith(".tmp")]

                    if expected_base:
                        filtered_files = []
                        for file_name in valid_files:
                            base = os.path.splitext(file_name)[0].lower()
                            if base == expected_base or base.startswith(f"{expected_base} ("):
                                filtered_files.append(file_name)
                        valid_files = filtered_files

                    if valid_files:
                        valid_files.sort(
                            key=lambda f: os.path.getmtime(os.path.join(active_dir, f)),
                            reverse=True,
                        )

                        for file_name in valid_files:
                            file_path = os.path.join(active_dir, file_name)

                            try:
                                current_size = os.path.getsize(file_path)
                                previous_size = stable_sizes.get(file_name)
                                stable_sizes[file_name] = current_size

                                if previous_size is None or previous_size != current_size:
                                    stable_counts[file_name] = 0
                                    continue
                                stable_counts[file_name] = stable_counts.get(file_name, 0) + 1
                                if stable_counts[file_name] < 2:
                                    continue
                            except:
                                continue

                            if self._is_valid_pdf_file(file_path):
                                return file_name

                            self.log(f"Descarga invalida detectada y descartada: {file_name}", "WARN")
                            ignored_files.add(file_name)
                            stable_sizes.pop(file_name, None)
                            stable_counts.pop(file_name, None)
            except:
                pass
            can_continue, paused_elapsed = self._sleep_with_pause_and_elapsed(0.5)
            end_time += paused_elapsed
            if not can_continue:
                return None
        return None

    def _wait_for_pdf_ready(self, timeout=45):
        end_time = time.monotonic() + timeout
        last_uri = ""
        stable_uri_hits = 0
        while time.monotonic() < end_time:
            can_continue, paused_elapsed = self._wait_if_paused_with_elapsed()
            end_time += paused_elapsed
            if not can_continue:
                return None
            try:
                result = self.driver.execute_script(
                    """
                    const data = {
                        uri: "",
                        viewerDownloadUrl: "",
                        readyState: document.readyState || "",
                        buttonReady: false,
                        buttonEnabled: false
                    };

                    function toAbsolute(url) {
                        if (!url) return "";
                        try { return new URL(url, window.location.href).href; } catch (e) { return ""; }
                    }

                    function parseReaderControlDownloadUrl(readerSrc) {
                        try {
                            const full = new URL(readerSrc, window.location.href);
                            const hash = (full.hash || "").replace(/^#/, "");
                            if (!hash) return "";
                            const params = new URLSearchParams(hash);
                            const customRaw = params.get("custom");
                            if (!customRaw) return "";
                            const custom = JSON.parse(customRaw);
                            if (custom && custom.downloadAsPdfUrl) {
                                return toAbsolute(custom.downloadAsPdfUrl);
                            }
                        } catch (e) {}
                        return "";
                    }

                    // Caso especial del visor DocumentUltimate (SecureImage.aspx).
                    try {
                        const readerFrame = document.querySelector("iframe[src*='ReaderControl.html']");
                        if (readerFrame) {
                            const src = readerFrame.getAttribute("src") || "";
                            const duUrl = parseReaderControlDownloadUrl(src);
                            if (duUrl) data.viewerDownloadUrl = duUrl;
                        }
                    } catch (e) {}

                    // Fallback: extraer downloadAsPdfUrl desde el script inline del loader.
                    if (!data.viewerDownloadUrl) {
                        try {
                            const loader = document.getElementById("documentViewer-loader");
                            const txt = loader ? (loader.textContent || "") : "";
                            if (txt) {
                                const m = txt.match(/downloadAsPdfUrl\\\\?":\\\\?"([^"]+)/i);
                                if (m && m[1]) {
                                    const unescaped = m[1]
                                        .replace(/\\\\u0026/g, "&")
                                        .replace(/\\u0026/g, "&")
                                        .replace(/\\\\\\//g, "/")
                                        .replace(/\\\\/g, "");
                                    data.viewerDownloadUrl = toAbsolute(unescaped);
                                }
                            }
                        } catch (e) {}
                    }

                    const embed = document.querySelector("embed[src]");
                    const iframe = document.querySelector("iframe[src]");
                    const obj = document.querySelector("object[data]");

                    if (embed && embed.src) data.uri = embed.src;
                    if (!data.uri && iframe && iframe.src) data.uri = iframe.src;
                    if (!data.uri && obj && obj.data) data.uri = obj.data;

                    if (!data.uri && window.location && window.location.href) {
                        const href = window.location.href;
                        if (href.startsWith("blob:") || href.toLowerCase().includes(".pdf")) data.uri = href;
                    }

                    const blockedPrefixes = ["chrome-extension://", "edge://", "about:", "data:text/html"];
                    if (data.uri && blockedPrefixes.some(prefix => data.uri.startsWith(prefix))) {
                        data.uri = "";
                    }

                    // Si tenemos URL directa de DownloadAsPdf, esta es la fuente mas confiable.
                    if (data.viewerDownloadUrl) {
                        data.uri = data.viewerDownloadUrl;
                    }

                    const directBtn = document.querySelector("a[download], button[download], #download, cr-icon-button#download, [aria-label*='Download'], [aria-label*='Descargar'], [title*='Download'], [title*='Descargar']");
                    if (directBtn) {
                        data.buttonReady = true;
                        data.buttonEnabled = !directBtn.disabled && !directBtn.hasAttribute("disabled");
                    }

                    try {
                        const viewer = document.querySelector("pdf-viewer");
                        if (viewer && viewer.shadowRoot) {
                            const toolbar = viewer.shadowRoot.querySelector("viewer-toolbar");
                            if (toolbar && toolbar.shadowRoot) {
                                const shadowBtn = toolbar.shadowRoot.querySelector(
                                    "#download, cr-icon-button#download, button[aria-label*='Download'], button[aria-label*='Descargar']"
                                );
                                if (shadowBtn) {
                                    data.buttonReady = true;
                                    data.buttonEnabled = !shadowBtn.disabled && !shadowBtn.hasAttribute("disabled");
                                }
                            }
                        }
                    } catch (err) {}

                    return data;
                    """
                )

                result = result or {}
                uri = str(result.get("viewerDownloadUrl") or result.get("uri", "")).strip()
                page_ready = str(result.get("readyState", "")) == "complete"
                button_ready = bool(result.get("buttonReady"))
                button_enabled = bool(result.get("buttonEnabled"))

                # Si el boton del visor ya esta habilitado, preferir esa ruta
                if button_ready and button_enabled and page_ready:
                    return ""
                if uri and page_ready:
                    if uri == last_uri:
                        stable_uri_hits += 1
                    else:
                        last_uri = uri
                        stable_uri_hits = 1

                    if stable_uri_hits >= 2:
                        return uri
            except:
                pass

            can_continue, paused_elapsed = self._sleep_with_pause_and_elapsed(1)
            end_time += paused_elapsed
            if not can_continue:
                return None

        return None

    def _is_pdf_context_visible(self):
        try:
            return bool(
                self.driver.execute_script(
                    """
                    try {
                        const ct = (document.contentType || "").toLowerCase();
                        if (ct.includes("pdf")) return true;

                        const embed = document.querySelector("embed[src*='.pdf'], embed[src^='blob:']");
                        const frame = document.querySelector("iframe[src*='.pdf'], iframe[src^='blob:']");
                        const obj = document.querySelector("object[type*='pdf'], object[data*='.pdf'], object[data^='blob:']");
                        const viewer = document.querySelector("pdf-viewer");

                        return !!(embed || frame || obj || viewer);
                    } catch (e) {
                        return false;
                    }
                    """
                )
            )
        except Exception:
            return False

    def _download_via_source(self, source_uri, nombre_archivo):
        if not source_uri:
            return False
        try:
            result = self.driver.execute_async_script(
                """
                const uri = arguments[0];
                const filename = arguments[1];
                const done = arguments[arguments.length - 1];

                if (!uri) {
                    done({ ok: false, error: "sin_uri" });
                    return;
                }

                fetch(uri)
                    .then(resp => {
                        if (!resp.ok && !uri.startsWith("blob:")) {
                            throw new Error("HTTP " + resp.status);
                        }

                        const ct = (resp.headers.get("content-type") || "").toLowerCase();
                        if (
                            ct &&
                            !ct.includes("pdf") &&
                            !ct.includes("octet-stream") &&
                            !ct.includes("binary") &&
                            !uri.startsWith("blob:")
                        ) {
                            throw new Error("Respuesta no es PDF: " + ct);
                        }
                        return resp.blob();
                    })
                    .then(blob => {
                        if (blob.size < 1024) {
                            throw new Error("Blob demasiado pequeno");
                        }

                        return blob.slice(0, 5).text().then(head => {
                            if (head !== "%PDF-") {
                                throw new Error("Archivo no inicia con %PDF-");
                            }
                            return blob;
                        });
                    })
                    .then(blob => {
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement("a");
                        a.style.display = "none";
                        a.href = url;
                        a.download = filename;
                        document.body.appendChild(a);
                        a.click();
                        window.URL.revokeObjectURL(url);
                        done({ ok: true });
                    })
                    .catch(err => {
                        done({ ok: false, error: String(err) });
                    });
                """,
                source_uri,
                nombre_archivo,
            )
            if result and result.get("ok"):
                return True

            try:
                error_detail = str((result or {}).get("error", "")).strip()
            except Exception:
                error_detail = ""

            if error_detail:
                self.log(f"Descarga por URL fallo: {error_detail}", "WARN")
            return False
        except:
            return False

    def _click_pdf_download_button(self):
        try:
            clicked = self.driver.execute_script(
                """
                function clickFirst(root, selectors) {
                    if (!root) return false;
                    for (const selector of selectors) {
                        const el = root.querySelector(selector);
                        if (el) {
                            if (el.disabled || el.hasAttribute("disabled")) continue;
                            const style = window.getComputedStyle(el);
                            if (style.display === "none" || style.visibility === "hidden") continue;
                            el.click();
                            return true;
                        }
                    }
                    return false;
                }

                const selectors = [
                    "a[download]",
                    "button[download]",
                    "#download",
                    "cr-icon-button#download",
                    "button[aria-label*='Download']",
                    "button[aria-label*='Descargar']",
                    "[title*='Download']",
                    "[title*='Descargar']"
                ];

                if (clickFirst(document, selectors)) return true;

                // Fallback para visores que renderizan toolbar dentro de iframe (p. ej. DocumentUltimate).
                try {
                    const frameSelectors = [
                        "#downloadButton",
                        "[data-element='downloadButton']",
                        "[data-element='downloadAsPdfButton']",
                        "[data-element*='download']",
                        "button[aria-label*='Download']",
                        "button[title*='Download']",
                        "[title*='Download as PDF']",
                        "[title*='Download']",
                        "a[download]"
                    ];

                    const frames = Array.from(document.querySelectorAll("iframe"));
                    for (const frame of frames) {
                        let frameDoc = null;
                        try {
                            frameDoc = frame.contentWindow && frame.contentWindow.document ? frame.contentWindow.document : null;
                        } catch (e) {
                            frameDoc = null;
                        }
                        if (!frameDoc) continue;

                        if (clickFirst(frameDoc, frameSelectors)) return true;

                        const nestedFrames = Array.from(frameDoc.querySelectorAll("iframe"));
                        for (const nested of nestedFrames) {
                            let nestedDoc = null;
                            try {
                                nestedDoc = nested.contentWindow && nested.contentWindow.document ? nested.contentWindow.document : null;
                            } catch (e) {
                                nestedDoc = null;
                            }
                            if (!nestedDoc) continue;
                            if (clickFirst(nestedDoc, frameSelectors)) return true;
                        }
                    }
                } catch (err) {}

                try {
                    const viewer = document.querySelector("pdf-viewer");
                    if (viewer && viewer.shadowRoot) {
                        if (clickFirst(viewer.shadowRoot, selectors)) return true;
                        const toolbar = viewer.shadowRoot.querySelector("viewer-toolbar");
                        if (toolbar && toolbar.shadowRoot) {
                            if (clickFirst(toolbar.shadowRoot, selectors)) return true;
                        }
                    }
                } catch (err) {}

                return false;
                """
            )
            return bool(clicked)
        except:
            return False

    def _start_pdf_download(self, source_uri, nombre_archivo):
        source_text = (source_uri or "").lower()

        # En visores tipo DocumentUltimate, la URL DownloadAsPdf suele ser la via mas estable.
        if source_uri and ("downloadaspdf" in source_text or "documentviewer.ashx" in source_text):
            if self._download_via_source(source_uri, nombre_archivo):
                return True

        # Primero intentar boton real del visor (menos propenso a descargar contenido incorrecto)
        if self._click_pdf_download_button():
            return True

        if source_uri:
            if self._download_via_source(source_uri, nombre_archivo):
                return True
            self.log("No se pudo descargar por URL valida. Reintentando...", "WARN")

        return False

    def _close_extra_windows(self, main_window):
        try:
            handles = list(self.driver.window_handles)
        except Exception:
            return

        for handle in handles:
            if handle == main_window:
                continue
            try:
                self.driver.switch_to.window(handle)
                self.driver.close()
            except Exception:
                pass

        try:
            remaining = list(self.driver.window_handles)
            if main_window in remaining:
                self.driver.switch_to.window(main_window)
            elif remaining:
                self.driver.switch_to.window(remaining[0])
            self._install_browser_close_warning(enabled=False)
        except Exception:
            pass

    def _scraping_logic(self, target_rows=None):
        start_row_index = 0
        try:
            self.log("Iniciando secuencia de descargas...", "INFO")
            wait = WebDriverWait(self.driver, 20)
            main_window = self.driver.current_window_handle
            self._close_extra_windows(main_window)

            rows = self._get_table_rows()
            total_rows = len(rows)

            if total_rows <= 0:
                self.log("No se encontraron filas. Verifique que esta en la tabla correcta.", "ERROR")
                self.after(0, lambda: self._refresh_stats_labels("Sin filas para procesar"))
                return

            if target_rows:
                process_rows = [idx for idx in sorted(set(target_rows)) if start_row_index <= idx < total_rows]
                if not process_rows:
                    self.log("No hay filas fallidas validas para reintentar en la vista actual.", "WARN")
                    self.after(0, lambda: self._refresh_stats_labels("Sin filas fallidas visibles"))
                    return
                self.log(f"Reintentando {len(process_rows)} filas fallidas.", "WARN")
            else:
                process_rows = list(range(start_row_index, total_rows))
                self.log(f"Total filas detectadas: {total_rows}", "INFO")

            total_data_rows = max(1, len(process_rows))
            last_row_key = ""
            last_target_identity = ""

            for position, i in enumerate(process_rows):
                if not self.is_running:
                    self.log("Cancelado por usuario.", "WARN")
                    break

                if not self._wait_if_paused():
                    break

                current_progress = position / total_data_rows
                self.after(0, lambda v=current_progress: self.progress.set(v))
                self.after(0, lambda idx=i, total=process_rows[-1]: self._refresh_stats_labels(f"Procesando fila {idx}/{total}"))

                try:
                    fresh_rows = self._get_table_rows()
                    if i < len(fresh_rows):
                        row_element = fresh_rows[i]
                        self.driver.execute_script(
                            "arguments[0].style.backgroundColor = '#FFF2CC'; arguments[0].style.border = '2px solid #FFD966';",
                            row_element,
                        )
                except:
                    pass

                max_retries = 3
                success = False
                last_error_message = "Sin detalle"
                last_attempt_number = 0
                completed_detail_indexes = set()
                detail_plan_indexes = None
                detail_entries_by_index = {}
                row_document_type = ""

                for attempt in range(max_retries):
                    if not self._wait_if_paused():
                        break

                    url_before_open_pdf = ""
                    current_row_key = ""
                    target_identity = ""
                    try:
                        self._close_extra_windows(main_window)
                        prefix_msg = f"Intento {attempt + 1}" if attempt > 0 else "Procesando"
                        self.log(f"{prefix_msg} fila #{i}...", "INFO")

                        current_rows = self._get_table_rows()
                        if i >= len(current_rows):
                            break
                        current_row = current_rows[i]

                        btn_ver_1 = current_row.find_element(By.XPATH, ".//*[contains(text(), 'Ver') or contains(@class, 'view')]")
                        row_parts = self._extract_verdetalle_parts(btn_ver_1)
                        current_row_key = row_parts.get("key", "")
                        if current_row_key and last_row_key and current_row_key == last_row_key and i > 0:
                            raise Exception("Se detecto la misma fila que la anterior; reintentando para evitar duplicado")

                        row_document_type = self._extract_row_document_type(current_row, row_parts) or row_parts.get("tipo", "")
                        normalized_document_type = self._normalize_document_type(row_document_type)
                        modal_signature_before = self._get_modal_detail_signature()

                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_ver_1)
                        if not self._sleep_with_pause(0.5):
                            break
                        try:
                            btn_ver_1.click()
                        except:
                            self.driver.execute_script("arguments[0].click();", btn_ver_1)

                        try:
                            WebDriverWait(self.driver, 10).until(
                                lambda d: self._get_modal_detail_signature() != modal_signature_before
                            )
                        except TimeoutException:
                            pass

                        modal_buttons, btn_ver_2_xpath = self._find_modal_view_buttons(timeout=8)
                        modal_entries = self._get_modal_detail_entries()

                        if detail_plan_indexes is None:
                            if "certificado" in normalized_document_type:
                                detail_plan_indexes = [0]
                                self.log(
                                    f"Fila #{i}: tipo '{row_document_type or 'Documento'}'. Se descargara solo el primer detalle.",
                                    "INFO",
                                )
                            else:
                                detail_plan_indexes = list(range(len(modal_buttons)))
                                self.log(
                                    f"Fila #{i}: tipo '{row_document_type or 'Documento'}'. Se descargaran todos los detalles ({len(detail_plan_indexes)}).",
                                    "INFO",
                                )

                            if not detail_plan_indexes:
                                raise TimeoutException("No se encontraron detalles para descargar en el modal")

                            if max(detail_plan_indexes) >= len(modal_buttons):
                                raise TimeoutException("No se encontraron suficientes enlaces 'Ver' en el modal")

                            detail_entries_by_index = {}
                            for detail_index in detail_plan_indexes:
                                info = f"Detalle {detail_index + 1}"
                                if detail_index < len(modal_entries) and isinstance(modal_entries[detail_index], dict):
                                    info = str(modal_entries[detail_index].get("info") or info).strip() or info
                                detail_entries_by_index[detail_index] = info

                        use_current_modal = True

                        for detail_index in detail_plan_indexes:
                            if not self._wait_if_paused():
                                break
                            if detail_index in completed_detail_indexes:
                                continue

                            if not use_current_modal:
                                current_rows = self._get_table_rows()
                                if i >= len(current_rows):
                                    raise Exception("La fila ya no esta visible para reabrir el detalle")
                                current_row = current_rows[i]
                                btn_ver_1 = current_row.find_element(By.XPATH, ".//*[contains(text(), 'Ver') or contains(@class, 'view')]")

                                modal_signature_before = self._get_modal_detail_signature()
                                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_ver_1)
                                if not self._sleep_with_pause(0.5):
                                    break
                                try:
                                    btn_ver_1.click()
                                except Exception:
                                    self.driver.execute_script("arguments[0].click();", btn_ver_1)

                                try:
                                    WebDriverWait(self.driver, 10).until(
                                        lambda d: self._get_modal_detail_signature() != modal_signature_before
                                    )
                                except TimeoutException:
                                    pass

                                modal_buttons, btn_ver_2_xpath = self._find_modal_view_buttons(timeout=8)
                                modal_entries = self._get_modal_detail_entries()
                                if detail_index < len(modal_entries) and isinstance(modal_entries[detail_index], dict):
                                    detail_entries_by_index[detail_index] = (
                                        str(modal_entries[detail_index].get("info") or detail_entries_by_index.get(detail_index, f"Detalle {detail_index + 1}")).strip()
                                        or detail_entries_by_index.get(detail_index, f"Detalle {detail_index + 1}")
                                    )

                            use_current_modal = False

                            if detail_index >= len(modal_buttons):
                                raise Exception(f"No se encontro el detalle {detail_index + 1} en el modal")

                            btn_ver_2 = modal_buttons[detail_index]
                            detail_label = detail_entries_by_index.get(detail_index, f"Detalle {detail_index + 1}")
                            nombre_archivo = self._build_unique_filename(i, btn_ver_1, detail_index)

                            handles_before = set(self.driver.window_handles)
                            try:
                                url_before_open_pdf = self.driver.current_url
                            except Exception:
                                pass

                            link_href = ""
                            try:
                                link_href = (btn_ver_2.get_attribute("href") or "").strip()
                            except Exception:
                                link_href = ""

                            resolved_href = ""
                            if link_href:
                                try:
                                    resolved_href = (
                                        self.driver.execute_script(
                                            "try { return new URL(arguments[0], window.location.href).href; } catch (e) { return arguments[0] || ''; }",
                                            link_href,
                                        )
                                        or ""
                                    ).strip()
                                except Exception:
                                    resolved_href = link_href

                            if resolved_href:
                                target_identity = f"{current_row_key}|{resolved_href}".strip("|")
                            elif current_row_key:
                                target_identity = f"{current_row_key}|detalle_{detail_index + 1}"
                            if position > 0 and target_identity and target_identity == last_target_identity:
                                raise Exception("Se detecto el mismo documento que la fila previa; reintentando")

                            if resolved_href.lower().startswith(("http://", "https://")):
                                self.driver.execute_script("window.open(arguments[0], '_blank', 'noopener,noreferrer');", resolved_href)
                            else:
                                if btn_ver_2_xpath:
                                    try:
                                        refreshed_buttons = self.driver.find_elements(By.XPATH, btn_ver_2_xpath)
                                        visible_buttons = [button for button in refreshed_buttons if button.is_displayed()]
                                        if detail_index < len(visible_buttons):
                                            btn_ver_2 = visible_buttons[detail_index]
                                    except Exception:
                                        pass
                                try:
                                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_ver_2)
                                except Exception:
                                    pass
                                try:
                                    btn_ver_2.click()
                                except Exception:
                                    self.driver.execute_script("arguments[0].click();", btn_ver_2)

                            opened_new_window = False
                            try:
                                wait.until(lambda d: len(d.window_handles) > len(handles_before))
                                opened_new_window = True
                            except TimeoutException:
                                current_after_open = ""
                                try:
                                    current_after_open = self.driver.current_url
                                except Exception:
                                    current_after_open = ""

                                if (
                                    url_before_open_pdf
                                    and current_after_open
                                    and current_after_open != url_before_open_pdf
                                    and current_after_open.lower().startswith(("http://", "https://"))
                                ):
                                    self.driver.get(url_before_open_pdf)
                                    WebDriverWait(self.driver, 20).until(lambda d: len(self._get_table_rows()) > 0)
                                    handles_before = set(self.driver.window_handles)
                                    self.driver.execute_script(
                                        "window.open(arguments[0], '_blank', 'noopener,noreferrer');",
                                        current_after_open,
                                    )
                                    wait.until(lambda d: len(d.window_handles) > len(handles_before))
                                    opened_new_window = True
                                else:
                                    WebDriverWait(self.driver, 25).until(
                                        lambda d: len(d.window_handles) > len(handles_before)
                                        or self._is_pdf_context_visible()
                                        or (url_before_open_pdf and d.current_url != url_before_open_pdf)
                                    )
                                    opened_new_window = len(self.driver.window_handles) > len(handles_before)

                            if opened_new_window:
                                new_handles = [h for h in self.driver.window_handles if h not in handles_before]
                                if new_handles:
                                    self.driver.switch_to.window(new_handles[-1])
                                else:
                                    for handle in self.driver.window_handles:
                                        if handle != main_window:
                                            self.driver.switch_to.window(handle)
                                            break
                                self._install_browser_close_warning()

                            current_open_url = ""
                            try:
                                current_open_url = (self.driver.current_url or "").strip()
                            except Exception:
                                current_open_url = ""

                            if current_open_url:
                                if current_row_key:
                                    target_identity = f"{current_row_key}|{current_open_url}"
                                elif not target_identity:
                                    target_identity = current_open_url

                            if position > 0 and target_identity and target_identity == last_target_identity:
                                raise Exception("Se detecto el mismo documento abierto que la fila previa; reintentando")

                            self.log(f"Esperando que el PDF termine de cargar ({detail_label})...", "WAIT")
                            pdf_source = self._wait_for_pdf_ready(timeout=50)
                            if pdf_source is None:
                                raise Exception(f"El PDF no cargo a tiempo para {detail_label}")

                            initial_files = set(os.listdir(self._active_download_dir()))
                            started = self._start_pdf_download(pdf_source, nombre_archivo)
                            if not started:
                                raise Exception(f"No fue posible iniciar la descarga de {detail_label}")

                            filename = self._wait_for_download(
                                initial_files,
                                timeout=50,
                                expected_name=nombre_archivo,
                            )

                            if self.driver.current_window_handle == main_window:
                                if url_before_open_pdf and self.driver.current_url != url_before_open_pdf:
                                    self.driver.get(url_before_open_pdf)
                                    WebDriverWait(self.driver, 20).until(lambda d: len(self._get_table_rows()) > 0)
                            self._close_extra_windows(main_window)
                            if not self._sleep_with_pause(0.8):
                                break

                            if filename:
                                file_path = os.path.join(self._active_download_dir(), filename)
                                if not self._is_valid_pdf_file(file_path):
                                    raise Exception(f"Se descargo un archivo invalido para {detail_label}")
                                self.log(f"Guardado: {filename}", "SUCCESS")
                                self.after(0, lambda fn=filename: self._add_history_entry(fn))
                                if target_identity:
                                    last_target_identity = target_identity
                                self.stats["success"] += 1
                                completed_detail_indexes.add(detail_index)
                                completed_count = len(completed_detail_indexes)
                                total_count = len(detail_plan_indexes)
                                self.after(
                                    0,
                                    lambda idx=i, done=completed_count, total=total_count: self._refresh_stats_labels(
                                        f"Fila {idx} detalle {done}/{total}"
                                    ),
                                )
                                continue

                            raise Exception(f"Timeout de descarga en {detail_label}")

                        if detail_plan_indexes and len(completed_detail_indexes) == len(detail_plan_indexes):
                            try:
                                fresh_rows = self._get_table_rows()
                                if i < len(fresh_rows):
                                    self.driver.execute_script(
                                        "arguments[0].style.backgroundColor = '#E2EFDA'; arguments[0].style.border = 'none';",
                                        fresh_rows[i],
                                    )
                            except Exception:
                                pass

                            if current_row_key:
                                last_row_key = current_row_key
                            if i in self.failed_rows:
                                self.failed_rows = [row for row in self.failed_rows if row != i]
                            self.after(0, lambda idx=i: self._refresh_stats_labels(f"Fila {idx} completada"))
                            success = True
                            break

                        raise Exception(
                            f"Quedaron detalles pendientes por descargar ({len(completed_detail_indexes)}/{len(detail_plan_indexes or [])})"
                        )

                    except Exception as e:
                        last_error_message = str(e)
                        last_attempt_number = attempt + 1
                        self.log(f"Error temporal: {last_error_message}", "ERROR")
                        self._close_extra_windows(main_window)

                        # Si el fallo ocurrio en la vista PDF, regresar a la tabla para evitar cascada de errores.
                        if url_before_open_pdf:
                            try:
                                current_url = self.driver.current_url
                            except Exception:
                                current_url = ""

                            if current_url != url_before_open_pdf:
                                try:
                                    self.driver.get(url_before_open_pdf)
                                    WebDriverWait(self.driver, 20).until(lambda d: len(self._get_table_rows()) > 0)
                                except Exception:
                                    try:
                                        self.driver.back()
                                        WebDriverWait(self.driver, 20).until(lambda d: len(self._get_table_rows()) > 0)
                                    except Exception:
                                        pass
                        if not self._sleep_with_pause(1.5):
                            break

                if not success:
                    self.log(f"Fallo definitivo en fila {i}.", "ERROR")
                    self.stats["error"] += 1
                    if i not in self.failed_rows:
                        self.failed_rows.append(i)
                    self._capture_failure_evidence(i, max(last_attempt_number, 1), last_error_message)
                    self.after(0, lambda idx=i: self._refresh_stats_labels(f"Fila {idx} con error"))
                    try:
                        fresh_rows = self._get_table_rows()
                        if i < len(fresh_rows):
                            self.driver.execute_script("arguments[0].style.backgroundColor = '#FCE4D6';", fresh_rows[i])
                    except:
                        pass

                done_progress = (position + 1) / total_data_rows
                self.after(0, lambda v=done_progress: self.progress.set(min(1, v)))

            try:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            except:
                pass

            pending_failed = len(sorted(set(self.failed_rows)))
            msg_final = (
                f"Terminado.\n\nOK: {self.stats['success']}\nErrores: {self.stats['error']}\n"
                f"Fallidos pendientes: {pending_failed}"
            )
            self.log(msg_final, "SUCCESS")
            self.after(0, lambda: self._refresh_stats_labels("Proceso finalizado"))
            self._show_info_threadsafe("Reporte Final", msg_final)

        except Exception as e:
            self.log(f"Error critico: {str(e)}", "ERROR")
            self.after(0, lambda: self._refresh_stats_labels("Error critico"))
        finally:
            self.is_running = False
            self.is_paused = False
            self.after(0, lambda: self.toggle_ui_state(working=False))

    def _setup_ui(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.configure(fg_color=APP_BG)

        self._build_sidebar_modern()
        self._build_main_surface_modern()
        self._build_hero_modern()
        self._build_workspace_modern()
        self._build_log_modern()
        self._build_footer_modern()
        self._build_history_view()

        self._refresh_stats_labels("Esperando inicio")
        self._update_retry_button_state()
        self._refresh_path_preview()
        self.log("Bienvenido. Configuracion cargada correctamente.", "INFO", persist=False)
        self._sync_progress_ring()

    def _build_sidebar_modern(self):
        self.sidebar_frame = ctk.CTkFrame(
            self,
            width=220,
            corner_radius=0,
            fg_color="#173835",
            border_width=0,
        )
        self.sidebar_frame.grid(row=0, column=0, sticky="ns")
        self.sidebar_frame.grid_propagate(False)
        self.sidebar_frame.grid_columnconfigure(0, weight=1)
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        # --- Brand header ---
        self.sidebar_brand = ctk.CTkFrame(
            self.sidebar_frame,
            fg_color="#204A46",
            corner_radius=0,
            height=100,
        )
        self.sidebar_brand.grid(row=0, column=0, sticky="ew")
        self.sidebar_brand.grid_propagate(False)

        ctk.CTkLabel(
            self.sidebar_brand,
            text="Centro RI",
            font=("Segoe UI", 10, "bold"),
            text_color="#E4D7C1",
        ).pack(anchor="w", padx=16, pady=(14, 2))

        self.lbl_sidebar_title = ctk.CTkLabel(
            self.sidebar_brand,
            text="Descargas\ninteligentes",
            font=("Bahnschrift SemiBold", 22),
            text_color="#FFF8EC",
            justify="left",
        )
        self.lbl_sidebar_title.pack(anchor="w", padx=16)

        # --- Botones ---
        sidebar_buttons = [
            ("Abrir carpeta", self._open_download_folder),
            ("Ver historial", self._show_history_view),
            ("Actualizaciones", lambda: self._check_for_updates(manual=True)),
        ]
        for idx, (label, cmd) in enumerate(sidebar_buttons):
            ctk.CTkButton(
                self.sidebar_frame,
                text=label,
                command=cmd,
                height=36,
                fg_color="#2A5B56",
                hover_color="#346963",
                text_color="#FFF9EF",
                corner_radius=12,
                font=("Segoe UI", 12, "bold"),
                anchor="w",
            ).grid(row=idx + 1, column=0, sticky="ew", padx=12, pady=(12 if idx == 0 else 3, 3))

        # --- Footer ---
        self.sidebar_hint = ctk.CTkLabel(
            self.sidebar_frame,
            text=f"{APP_NAME}\n{APP_VERSION_LABEL}",
            font=("Segoe UI", 9),
            text_color="#6B8E89",
            wraplength=180,
            justify="left",
        )
        self.sidebar_hint.grid(row=5, column=0, padx=14, pady=(4, 10), sticky="sw")

    # ---- Historial en el area principal ----

    def _build_history_view(self):
        """Construye el frame de historial (oculto por defecto) en el area principal."""
        self.history_frame = ctk.CTkFrame(
            self.main_frame,
            fg_color="transparent",
        )
        self.history_frame.grid_columnconfigure(0, weight=1)
        self.history_frame.grid_rowconfigure(1, weight=1)

        # Cabecera con titulo y boton volver
        history_header = ctk.CTkFrame(
            self.history_frame,
            corner_radius=24,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
        )
        history_header.grid(row=0, column=0, sticky="ew", padx=22, pady=(20, 14))
        history_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            history_header,
            text="Historial de descargas",
            font=("Bahnschrift SemiBold", 28),
            text_color=TEXT_MAIN,
        ).grid(row=0, column=0, padx=24, pady=(22, 6), sticky="w")

        ctk.CTkLabel(
            history_header,
            text="Todas las descargas exitosas registradas por el sistema.",
            font=("Segoe UI", 13),
            text_color=TEXT_MUTED,
        ).grid(row=1, column=0, padx=24, pady=(0, 6), sticky="w")

        header_actions = ctk.CTkFrame(history_header, fg_color="transparent")
        header_actions.grid(row=0, column=1, rowspan=2, padx=22, pady=16, sticky="e")

        ctk.CTkButton(
            header_actions,
            text="Volver",
            command=self._hide_history_view,
            width=100,
            height=36,
            fg_color="#2A5B56",
            hover_color="#346963",
            text_color="#FFF9EF",
            corner_radius=12,
            font=("Segoe UI", 13, "bold"),
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            header_actions,
            text="Limpiar historial",
            command=self._clear_history,
            width=130,
            height=36,
            fg_color="transparent",
            hover_color="#F0E0CC",
            text_color=TEXT_MUTED,
            border_width=1,
            border_color=CARD_BORDER,
            corner_radius=12,
            font=("Segoe UI", 12),
        ).pack(side="right")

        # Contador de descargas
        self.history_count_label = ctk.CTkLabel(
            history_header,
            text="",
            font=("Segoe UI", 11),
            text_color=TEXT_MUTED,
        )
        self.history_count_label.grid(row=2, column=0, padx=24, pady=(0, 16), sticky="w")

        # Area scrollable con las entradas
        self.history_scroll = ctk.CTkScrollableFrame(
            self.history_frame,
            corner_radius=22,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
            scrollbar_button_color="#D4C9B8",
            scrollbar_button_hover_color="#B8A998",
        )
        self.history_scroll.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 12))
        self.history_scroll.grid_columnconfigure(0, weight=1)

        self.history_entries = []
        self._history_visible = False

    def _load_download_history(self):
        """Lee las ultimas descargas exitosas del archivo de log y las muestra."""
        # Limpiar entradas previas
        for widget in self.history_scroll.winfo_children():
            widget.destroy()
        self.history_entries.clear()

        entries = []
        if os.path.exists(LOG_FILE):
            try:
                with open(LOG_FILE, "r", encoding="utf-8") as f:
                    for line in f:
                        if "[SUCCESS] Guardado:" in line:
                            match = re.match(
                                r"\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] \[SUCCESS\] Guardado: (.+)",
                                line.strip(),
                            )
                            if match:
                                entries.append((match.group(1), match.group(2)))
            except Exception:
                pass

        if not entries:
            self._show_empty_history()
            self.history_count_label.configure(text="0 descargas registradas")
            return

        # Mostrar las ultimas 200 (mas recientes primero)
        recent = list(reversed(entries[-200:]))
        for timestamp, filename in recent:
            self._create_history_row(timestamp, filename)

        self.history_count_label.configure(text=f"{len(recent)} descargas registradas")

    def _show_empty_history(self):
        """Muestra mensaje cuando no hay historial."""
        empty_frame = ctk.CTkFrame(self.history_scroll, fg_color="transparent")
        empty_frame.pack(fill="x", padx=20, pady=40)

        ctk.CTkLabel(
            empty_frame,
            text="Sin descargas registradas",
            font=("Bahnschrift SemiBold", 18),
            text_color=TEXT_MUTED,
        ).pack(anchor="center", pady=(20, 6))

        ctk.CTkLabel(
            empty_frame,
            text="Cuando descargues archivos, apareceran aqui\ncon la fecha y nombre del documento.",
            font=("Segoe UI", 13),
            text_color="#A09888",
            justify="center",
        ).pack(anchor="center")

    def _create_history_row(self, timestamp, filename):
        """Crea una fila visual en el historial."""
        row = ctk.CTkFrame(
            self.history_scroll,
            fg_color=SURFACE_SOFT,
            corner_radius=14,
            height=52,
        )
        row.pack(fill="x", padx=12, pady=3)
        row.pack_propagate(False)
        row.grid_columnconfigure(0, weight=1)

        # Formato de fecha legible
        short_ts = timestamp
        try:
            dt = datetime.datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
            meses = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]
            short_ts = f"{dt.day} {meses[dt.month - 1]} {dt.year}  {dt.strftime('%H:%M')}"
        except Exception:
            pass

        ctk.CTkLabel(
            row,
            text=filename,
            font=("Segoe UI", 12),
            text_color=TEXT_MAIN,
            anchor="w",
        ).grid(row=0, column=0, padx=14, pady=(8, 0), sticky="w")

        ctk.CTkLabel(
            row,
            text=short_ts,
            font=("Segoe UI", 10),
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, padx=14, pady=(0, 6), sticky="w")

        self.history_entries.append(row)

    def _show_history_view(self):
        """Muestra la vista de historial ocultando el contenido principal."""
        if self._history_visible:
            return
        self._load_download_history()
        # Ocultar contenido principal
        self.hero_frame.grid_remove()
        self.content_grid.grid_remove()
        self.log_frame.grid_remove()
        self.footer_frame.grid_remove()
        # Mostrar historial
        self.history_frame.grid(row=0, column=0, rowspan=4, sticky="nsew", padx=0, pady=0)
        self._history_visible = True

    def _hide_history_view(self):
        """Oculta la vista de historial y restaura el contenido principal."""
        if not self._history_visible:
            return
        self.history_frame.grid_remove()
        # Restaurar contenido principal
        self.hero_frame.grid()
        self.content_grid.grid()
        self.log_frame.grid()
        self.footer_frame.grid()
        self._history_visible = False

    def _add_history_entry(self, filename):
        """Registra una descarga exitosa. Si la vista de historial esta abierta, la refresca."""
        if self._history_visible:
            self._load_download_history()

    def _open_download_folder(self):
        """Abre la carpeta de descargas en el explorador."""
        folder = self._normalize_folder_path(self.download_dir.get())
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            self.log(f"La carpeta no existe: {folder}", "WARN")

    def _clear_history(self):
        """Limpia el archivo de historial y la vista."""
        if not os.path.exists(LOG_FILE):
            self.log("No hay historial para limpiar.", "INFO")
            return
        confirm = messagebox.askyesno(
            "Limpiar historial",
            "Esto eliminara todas las entradas de descargas del historial.\n\nContinuar?",
        )
        if not confirm:
            return
        # Reescribir log sin las lineas de SUCCESS/Guardado
        try:
            with open(LOG_FILE, "r", encoding="utf-8") as f:
                lines = f.readlines()
            filtered = [l for l in lines if "[SUCCESS] Guardado:" not in l]
            with open(LOG_FILE, "w", encoding="utf-8") as f:
                f.writelines(filtered)
        except Exception:
            pass
        # Refrescar vista
        for widget in self.history_scroll.winfo_children():
            widget.destroy()
        self.history_entries.clear()
        self._show_empty_history()
        self.history_count_label.configure(text="0 descargas registradas")
        self.log("Historial de descargas limpiado.", "INFO")

    def _build_main_surface_modern(self):
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)

        self.bg_orb_left = ctk.CTkFrame(
            self.main_frame,
            width=260,
            height=260,
            corner_radius=130,
            fg_color="#E8DBC8",
        )
        self.bg_orb_left.place(x=-110, y=102)
        self.bg_orb_left.lower()

        self.bg_orb_center = ctk.CTkFrame(
            self.main_frame,
            width=180,
            height=180,
            corner_radius=90,
            fg_color="#E4EBDC",
        )
        self.bg_orb_center.place(relx=0.53, y=38, anchor="n")
        self.bg_orb_center.lower()

        self.bg_orb_right = ctk.CTkFrame(
            self.main_frame,
            width=320,
            height=320,
            corner_radius=160,
            fg_color="#E9E0D2",
        )
        self.bg_orb_right.place(relx=1.0, x=-64, y=456, anchor="ne")
        self.bg_orb_right.lower()

    def _build_hero_modern(self):
        self.hero_frame = ctk.CTkFrame(
            self.main_frame,
            corner_radius=18,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
        )
        self.hero_frame.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 8))
        self.hero_frame.grid_columnconfigure(0, weight=1)

        # --- Top row: title + badges ---
        hero_top = ctk.CTkFrame(self.hero_frame, fg_color="transparent")
        hero_top.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 0))
        hero_top.grid_columnconfigure(0, weight=1)

        self.lbl_title = ctk.CTkLabel(
            hero_top,
            text="Centro de descargas",
            font=("Bahnschrift SemiBold", 22),
            text_color=TEXT_MAIN,
        )
        self.lbl_title.grid(row=0, column=0, sticky="w")

        # Badges right
        badges = ctk.CTkFrame(hero_top, fg_color="transparent")
        badges.grid(row=0, column=1, sticky="e")

        self.lbl_runtime_status = ctk.CTkLabel(
            badges, text="LISTO",
            font=("Segoe UI", 10, "bold"), fg_color=PRIMARY,
            corner_radius=999, padx=10, pady=3, text_color="#FFF8EF",
        )
        self.lbl_runtime_status.pack(side="right", padx=(6, 0))

        self.lbl_version = ctk.CTkLabel(
            badges, text=APP_VERSION_LABEL,
            font=("Segoe UI", 10, "bold"), text_color="#4E4B45",
            fg_color=SURFACE_SOFT, corner_radius=999, padx=10, pady=3,
        )
        self.lbl_version.pack(side="right")

        # --- Metrics row ---
        metrics_bar = ctk.CTkFrame(self.hero_frame, fg_color="transparent")
        metrics_bar.grid(row=1, column=0, sticky="ew", padx=18, pady=(12, 0))
        metrics_bar.grid_columnconfigure((0, 1), weight=1)

        # Completados
        self.success_card = ctk.CTkFrame(
            metrics_bar, fg_color="#E8F5EC", corner_radius=12,
        )
        self.success_card.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.success_card.grid_columnconfigure((0, 1), weight=0)

        self.lbl_success = ctk.CTkLabel(
            self.success_card, text="0",
            font=("Bahnschrift SemiBold", 28), text_color=SUCCESS,
        )
        self.lbl_success.pack(side="left", padx=(14, 6), pady=8)
        ctk.CTkLabel(
            self.success_card, text="completados",
            font=("Segoe UI", 11), text_color="#5F7A68",
        ).pack(side="left", pady=8)

        # Errores
        self.error_card = ctk.CTkFrame(
            metrics_bar, fg_color="#FBEEEA", corner_radius=12,
        )
        self.error_card.grid(row=0, column=1, sticky="ew", padx=(4, 0))

        self.lbl_error = ctk.CTkLabel(
            self.error_card, text="0",
            font=("Bahnschrift SemiBold", 28), text_color=DANGER,
        )
        self.lbl_error.pack(side="left", padx=(14, 6), pady=8)
        ctk.CTkLabel(
            self.error_card, text="errores",
            font=("Segoe UI", 11), text_color="#84675D",
        ).pack(side="left", pady=8)

        # --- Progress bar ---
        self.header_activity = ctk.CTkProgressBar(
            self.hero_frame,
            height=4,
            corner_radius=999,
            progress_color=PRIMARY,
            fg_color="#E6DDCF",
        )
        self.header_activity.grid(row=2, column=0, sticky="ew", padx=18, pady=(12, 14))
        self.header_activity.configure(mode="indeterminate")
        self.header_activity.stop()

        # Hidden helpers for compatibility
        self.lbl_subtitle = ctk.CTkLabel(self.hero_frame, text="", height=0, fg_color="transparent")
        self.lbl_dir = ctk.CTkLabel(self.hero_frame, text="", height=0, fg_color="transparent")
        self.hero_actions = ctk.CTkFrame(self.hero_frame, fg_color="transparent", height=0)

    def _build_workspace_modern(self):
        self.content_grid = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.content_grid.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 8))
        self.content_grid.grid_columnconfigure(0, weight=1)
        self.content_grid.grid_rowconfigure(1, weight=1)

        # ---- Row 0: Carpeta ----
        self.header_frame = ctk.CTkFrame(
            self.content_grid,
            corner_radius=18,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
        )
        self.header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=0)

        self.header_description = ctk.CTkLabel(
            self.header_frame,
            text="Destino de guardado",
            font=("Segoe UI", 13, "bold"),
            text_color=TEXT_MAIN,
            wraplength=400,
        )
        self.header_description.grid(row=0, column=0, padx=16, pady=(12, 6), sticky="w")

        self.entry_dir = ctk.CTkEntry(
            self.header_frame,
            textvariable=self.download_dir,
            height=38,
            fg_color=INPUT_BG,
            border_color="#D8CBBC",
            border_width=1,
            text_color=TEXT_MAIN,
            font=("Consolas", 11),
            corner_radius=10,
        )
        self.entry_dir.grid(row=1, column=0, padx=(16, 8), pady=(0, 12), sticky="ew")

        self.btn_folder = ctk.CTkButton(
            self.header_frame,
            text="Cambiar",
            command=self._select_folder,
            width=90,
            height=38,
            fg_color="#314B48",
            hover_color="#243836",
            text_color="#FFF8EF",
            corner_radius=10,
            font=("Segoe UI", 12, "bold"),
        )
        self.btn_folder.grid(row=1, column=1, padx=(0, 14), pady=(0, 12), sticky="e")

        # ---- Row 1: Controles + Estado ----
        self.bottom_row = ctk.CTkFrame(self.content_grid, fg_color="transparent")
        self.bottom_row.grid(row=1, column=0, sticky="nsew")
        self.bottom_row.grid_columnconfigure(0, weight=3)
        self.bottom_row.grid_columnconfigure(1, weight=2)
        self.bottom_row.grid_rowconfigure(0, weight=1)

        # Panel de controles
        self.controls_frame = ctk.CTkFrame(
            self.bottom_row,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
            corner_radius=18,
        )
        self.controls_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.controls_frame.grid_columnconfigure(0, weight=1)
        self.controls_frame.grid_columnconfigure(1, weight=1)
        self.controls_frame.grid_rowconfigure(3, weight=0)

        self.controls_description = ctk.CTkLabel(
            self.controls_frame,
            text="Controles",
            font=("Segoe UI", 14, "bold"),
            text_color=TEXT_MAIN,
            wraplength=500,
        )
        self.controls_description.grid(row=0, column=0, columnspan=2, padx=16, pady=(14, 4), sticky="w")

        # Progress + Actions inside controls
        self.console_frame = ctk.CTkFrame(self.controls_frame, fg_color="transparent")
        self.console_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.console_frame.grid_columnconfigure(0, weight=2)
        self.console_frame.grid_columnconfigure(1, weight=3)

        self.progress_shell = ctk.CTkFrame(
            self.console_frame,
            fg_color=INPUT_BG,
            corner_radius=16,
            border_width=0,
        )
        self.progress_shell.grid(row=0, column=0, padx=(14, 6), pady=(4, 14), sticky="nsew")
        self.progress_shell.grid_columnconfigure(0, weight=1)

        self.progress_ring_host = ctk.CTkFrame(self.progress_shell, fg_color="transparent")
        self.progress_ring_host.grid(row=0, column=0, pady=(12, 6), padx=10)

        self.progress_ring_canvas = tk.Canvas(
            self.progress_ring_host,
            width=160,
            height=160,
            bd=0,
            highlightthickness=0,
            bg=INPUT_BG,
        )
        self.progress_ring_canvas.pack()
        self.progress_ring_canvas.create_oval(16, 16, 144, 144, outline="#DDD1C4", width=12)
        self.progress_ring_arc = self.progress_ring_canvas.create_arc(
            16, 16, 144, 144,
            start=90, extent=0, style="arc", outline=PRIMARY, width=12,
        )
        self.progress_ring_canvas.create_text(
            80, 56, text="PROGRESO", fill="#7D756D", font=("Segoe UI", 8, "bold"),
        )
        self.progress_ring_text = self.progress_ring_canvas.create_text(
            80, 82, text="0%", fill=TEXT_MAIN, font=("Bahnschrift SemiBold", 22),
        )
        self.progress_ring_canvas.create_text(
            80, 108, text="Tiempo real", fill="#7D756D", font=("Segoe UI", 9), justify="center",
        )

        self.progress = ctk.CTkProgressBar(
            self.progress_shell,
            height=8,
            corner_radius=999,
            progress_color=PRIMARY,
            fg_color="#DDD1C4",
        )
        self.progress.set(0)
        self.progress.grid(row=1, column=0, padx=14, pady=(0, 12), sticky="ew")

        # Action buttons
        self.action_subframe = ctk.CTkFrame(self.console_frame, fg_color="transparent")
        self.action_subframe.grid(row=0, column=1, padx=(6, 14), pady=(4, 14), sticky="nsew")
        self.action_subframe.grid_columnconfigure((0, 1), weight=1)

        self.action_description = ctk.CTkLabel(
            self.action_subframe,
            text="Acciones",
            font=("Segoe UI", 12, "bold"),
            text_color=TEXT_MAIN,
            wraplength=300,
        )
        self.action_description.grid(row=0, column=0, columnspan=2, sticky="w", pady=(4, 8))

        btn_font = ("Segoe UI", 12, "bold")
        btn_height = 40

        self.btn_browser = ctk.CTkButton(
            self.action_subframe,
            text="Abrir portal",
            command=self.launch_browser,
            height=btn_height,
            corner_radius=12,
            font=btn_font,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="#FFF8EF",
        )
        self.btn_browser.grid(row=1, column=0, padx=(0, 4), pady=(0, 6), sticky="ew")

        self.btn_start = ctk.CTkButton(
            self.action_subframe,
            text="Iniciar descarga",
            command=self.start_scraping_thread,
            state="disabled",
            height=btn_height,
            corner_radius=12,
            font=btn_font,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="#FFF8EF",
        )
        self.btn_start.grid(row=1, column=1, padx=(4, 0), pady=(0, 6), sticky="ew")

        self.btn_pause = ctk.CTkButton(
            self.action_subframe,
            text="Pausar",
            command=self.toggle_pause,
            state="disabled",
            height=btn_height,
            corner_radius=12,
            font=btn_font,
            fg_color=WARN,
            hover_color=WARN_HOVER,
            text_color="#FFF8EF",
        )
        self.btn_pause.grid(row=2, column=0, padx=(0, 4), pady=(0, 6), sticky="ew")

        self.btn_cancel = ctk.CTkButton(
            self.action_subframe,
            text="Cancelar",
            command=self.stop_process,
            state="disabled",
            height=btn_height,
            corner_radius=12,
            font=btn_font,
            fg_color=DANGER,
            hover_color=DANGER_HOVER,
            text_color="#FFF8EF",
        )
        self.btn_cancel.grid(row=2, column=1, padx=(4, 0), pady=(0, 6), sticky="ew")

        self.btn_retry_failed = ctk.CTkButton(
            self.action_subframe,
            text="Reintentar fallidos (0)",
            command=self.retry_failed_rows_thread,
            state="disabled",
            height=btn_height,
            corner_radius=12,
            font=btn_font,
            fg_color="#344846",
            hover_color="#253633",
            text_color="#FFF8EF",
        )
        self.btn_retry_failed.grid(row=3, column=0, columnspan=2, pady=(0, 4), sticky="ew")

        # Panel de estado (derecha)
        self.status_panel = ctk.CTkFrame(
            self.bottom_row,
            fg_color=CARD_BG,
            border_color=CARD_BORDER,
            border_width=1,
            corner_radius=18,
        )
        self.status_panel.grid(row=0, column=1, sticky="nsew")
        self.status_panel.grid_columnconfigure(0, weight=1)

        self.status_description = ctk.CTkLabel(
            self.status_panel,
            text="Estado",
            font=("Segoe UI", 14, "bold"),
            text_color=TEXT_MAIN,
            wraplength=260,
        )
        self.status_description.grid(row=0, column=0, padx=14, pady=(14, 8), sticky="w")

        self.status_card = ctk.CTkFrame(
            self.status_panel,
            fg_color=INPUT_BG,
            border_color="#E2D6C8",
            border_width=1,
            corner_radius=14,
        )
        self.status_card.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))
        ctk.CTkLabel(
            self.status_card,
            text="ESTADO ACTUAL",
            font=("Segoe UI", 9, "bold"),
            text_color="#7A7168",
        ).pack(anchor="w", padx=12, pady=(10, 2))
        self.lbl_status = ctk.CTkLabel(
            self.status_card,
            text="Esperando inicio",
            font=("Segoe UI", 13, "bold"),
            text_color=TEXT_MAIN,
            justify="left",
            wraplength=220,
        )
        self.lbl_status.pack(anchor="w", padx=12, pady=(0, 10))

        self.warning_frame = ctk.CTkFrame(
            self.status_panel,
            fg_color=SURFACE_SOFT,
            border_color="#E2D6C8",
            border_width=1,
            corner_radius=14,
        )
        self.warning_frame.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 14))
        ctk.CTkLabel(
            self.warning_frame,
            text="GUIA RAPIDA",
            font=("Segoe UI", 9, "bold"),
            text_color="#7A7168",
        ).pack(anchor="w", padx=12, pady=(10, 2))

        instrucciones = (
            "1. Abre el navegador.\n"
            "2. Inicia sesion y pon la tabla en 100.\n"
            "3. Ejecuta la descarga.\n"
            "4. Usa Pausar o Reintentar si falta."
        )
        self.lbl_warning = ctk.CTkLabel(
            self.warning_frame,
            text=instrucciones,
            justify="left",
            font=("Segoe UI", 11),
            text_color="#5E564D",
            wraplength=220,
        )
        self.lbl_warning.pack(padx=12, pady=(0, 10), anchor="w")

    def _build_log_modern(self):
        self.log_frame = ctk.CTkFrame(
            self.main_frame,
            corner_radius=18,
            fg_color="#162220",
            border_color="#2A3F3D",
            border_width=1,
        )
        self.log_frame.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 8))
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)

        self.log_header = ctk.CTkFrame(
            self.log_frame,
            fg_color="transparent",
            corner_radius=0,
            border_width=0,
        )
        self.log_header.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 4))
        self.log_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.log_header,
            text="Actividad en vivo",
            font=("Segoe UI", 13, "bold"),
            text_color="#FFF8EF",
        ).grid(row=0, column=0, sticky="w")

        self.log_area = ctk.CTkTextbox(
            self.log_frame,
            font=("Consolas", 11),
            state="disabled",
            fg_color=LOG_BG,
            text_color="#E8F0EE",
            border_color=LOG_BORDER,
            border_width=1,
            corner_radius=12,
        )
        self.log_area.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))

    def _build_footer_modern(self):
        self.footer_frame = ctk.CTkFrame(self.main_frame, height=24, fg_color="transparent")
        self.footer_frame.grid(row=3, column=0, sticky="ew", padx=22, pady=(0, 10))
        self.footer_frame.grid_columnconfigure(0, weight=1)
        self.footer_frame.grid_columnconfigure(1, weight=0)

        self.lbl_credits = ctk.CTkLabel(
            self.footer_frame,
            text="Desarrollado por Dariel Espino (809) 434-2700",
            font=("Segoe UI", 10, "bold"),
            text_color=TEXT_MUTED,
        )
        self.lbl_credits.grid(row=0, column=0, sticky="w")

        self.footer_note = ctk.CTkLabel(
            self.footer_frame,
            text="Diseno visual renovado para que todo se entienda desde el primer minuto.",
            font=("Segoe UI", 10),
            text_color=TEXT_MUTED,
        )
        self.footer_note.grid(row=0, column=1, sticky="e")

    def _handle_window_resize(self, event=None):
        if event is not None and event.widget is not self:
            return

        if self._resize_after_id:
            try:
                self.after_cancel(self._resize_after_id)
            except Exception:
                pass

        self._resize_after_id = self.after(90, self._apply_responsive_layout)

    def _apply_responsive_layout(self):
        self._resize_after_id = None
        try:
            main_width = max(self.main_frame.winfo_width(), self.winfo_width() - self.sidebar_frame.winfo_width())
            window_height = self.winfo_height()
        except Exception:
            return

        compact = main_width < 820
        compressed_height = window_height < 700
        if self._compact_layout != compact:
            if compact:
                # Stack vertically
                self.bottom_row.grid_columnconfigure(0, weight=1)
                self.bottom_row.grid_columnconfigure(1, weight=1)
                self.status_panel.grid_configure(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
                self.controls_frame.grid_configure(row=0, column=0, columnspan=2, padx=0)
                self.progress_shell.grid_configure(row=0, column=0, columnspan=2, padx=14, pady=(4, 8), sticky="ew")
                self.action_subframe.grid_configure(row=1, column=0, columnspan=2, padx=14, pady=(0, 14), sticky="ew")
                self.console_frame.grid_columnconfigure(0, weight=1)
                self.console_frame.grid_columnconfigure(1, weight=1)
            else:
                # Side-by-side
                self.bottom_row.grid_columnconfigure(0, weight=3)
                self.bottom_row.grid_columnconfigure(1, weight=2)
                self.status_panel.grid_configure(row=0, column=1, columnspan=1, sticky="nsew", pady=0)
                self.controls_frame.grid_configure(row=0, column=0, columnspan=1, padx=(0, 8))
                self.progress_shell.grid_configure(row=0, column=0, columnspan=1, padx=(14, 6), pady=(4, 14), sticky="nsew")
                self.action_subframe.grid_configure(row=0, column=1, columnspan=1, padx=(6, 14), pady=(4, 14), sticky="nsew")
                self.console_frame.grid_columnconfigure(0, weight=2)
                self.console_frame.grid_columnconfigure(1, weight=3)

            self._compact_layout = compact

        try:
            def wrap_for(widget, padding, minimum, fallback, maximum=None):
                width = widget.winfo_width()
                if width <= 1:
                    width = fallback
                wraplength = max(minimum, int(width - padding))
                if maximum is not None:
                    wraplength = min(maximum, wraplength)
                return wraplength

            header_wrap = wrap_for(self.header_frame, 40, 200, 360, 500)
            controls_wrap = wrap_for(self.controls_frame, 40, 200, 400, 600)
            action_wrap = wrap_for(self.action_subframe, 18, 180, 280, 400)
            status_wrap = wrap_for(self.status_panel, 32, 180, 240, 400)
            status_value_wrap = wrap_for(self.status_card, 28, 160, 210, 380)
            warning_wrap = wrap_for(self.warning_frame, 28, 160, 210, 380)

            self.sidebar_hint.configure(wraplength=180)
            self.header_description.configure(wraplength=header_wrap)
            self.controls_description.configure(wraplength=controls_wrap)
            self.action_description.configure(wraplength=action_wrap)
            self.status_description.configure(wraplength=status_wrap)
            self.lbl_status.configure(wraplength=status_value_wrap)
            self.lbl_warning.configure(wraplength=warning_wrap)
            self.footer_note.configure(wraplength=300 if compressed_height else 400, justify="right")
        except Exception:
            pass

    def _run_intro_reveal(self):
        reveal_sequence = [
            self.hero_frame,
            self.content_grid,
            self.log_frame,
            self.footer_frame,
        ]
        for widget in reveal_sequence:
            widget.grid_remove()

        for idx, widget in enumerate(reveal_sequence):
            self.after(90 + (idx * 110), lambda w=widget: w.grid())

    def _animate_background_accents(self):
        self._accent_tick += 1
        phase = self._accent_tick / 16.0

        try:
            self.bg_orb_left.place(
                x=-110 + int(10 * math.sin(phase)),
                y=102 + int(9 * math.cos(phase * 0.8)),
            )
            self.bg_orb_center.place(
                relx=0.53,
                y=38 + int(8 * math.sin(phase * 1.1)),
                anchor="n",
            )
            self.bg_orb_right.place(
                relx=1.0,
                x=-64 + int(12 * math.cos(phase * 0.9)),
                y=456 + int(10 * math.sin(phase * 0.7)),
                anchor="ne",
            )
        except Exception:
            pass

        self.after(95, self._animate_background_accents)

    def _animate_status_chip(self):
        self._chip_pulse_tick += 1

        if self._chip_mode == "running":
            palette = ["#1D5A55", "#256863", "#2D746F"]
        elif self._chip_mode == "paused":
            palette = ["#C88B31", "#D3963F", "#BE832D"]
        elif self._chip_mode == "error":
            palette = ["#B14C38", "#BD5944", "#A64533"]
        elif self._chip_mode == "done":
            palette = ["#2E7B58", "#378664", "#2A714F"]
        else:
            palette = ["#35514D", "#41625D", "#30504B"]

        try:
            self.lbl_runtime_status.configure(fg_color=palette[self._chip_pulse_tick % len(palette)])
        except Exception:
            pass

        self.after(420, self._animate_status_chip)

    def _refresh_stats_labels(self, status_text=None):
        if status_text is not None:
            self.status_text = status_text

        self.lbl_success.configure(text=str(self.stats['success']))
        self.lbl_error.configure(text=str(self.stats['error']))
        self.lbl_status.configure(text=self.status_text)

        chip_text = "LISTO"
        self._chip_mode = "idle"
        if self.is_running and self.is_paused:
            chip_text = "PAUSADO"
            self._chip_mode = "paused"
        elif self.is_running:
            chip_text = "EN PROCESO"
            self._chip_mode = "running"
        elif self.stats["error"] > 0 and self.stats["success"] == 0:
            chip_text = "CON ERRORES"
            self._chip_mode = "error"
        elif self.stats["success"] > 0:
            chip_text = "COMPLETADO"
            self._chip_mode = "done"

        self.lbl_runtime_status.configure(text=chip_text)

if __name__ == "__main__":
    app = GobiernoPDFDownloader()
    app.mainloop()
