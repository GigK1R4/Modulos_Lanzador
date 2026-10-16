"""
=============================================================================
LANZADOR DE AUTOMATIZACIÓN DE ARCHIVOS
=============================================================================
Lanzador modular con interfaz gráfica moderna (CustomTkinter), autenticación
contra Google Sheets, permisos dinámicos y actualización automática con Git/GitHub.
Los módulos corresponden a los scripts del repositorio:
https://github.com/GigK1R4/Modulos_Lanzador.git
Cada script en dicho repositorio es un módulo ejecutable.
Diseñado para ejecutarse de forma 100% portable en cualquier PC:
- Sin crear archivos ni carpetas en el escritorio o ruta del ejecutable.
- Descarga y actualiza automáticamente los módulos desde GitHub (con o sin Git instalado).
- Todo se ejecuta en la misma ventana unificada con tamaño adaptativo.
=============================================================================
"""
import os
import sys
import re
import json
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from PIL import Image

# Importar módulos locales
from icon_generator import generate_all_icons, get_icon_for_module
from sheet_auth import SheetAuth
from git_updater import GitUpdater, get_app_data_dir

# Importaciones requeridas por los módulos dinámicos (Gene_Documentos, PDF_Unir, etc.)
# para asegurar que PyInstaller empaquete todas sus dependencias en el .exe autónomo
try:
    import pandas
    import openpyxl
    import docxtpl
    import docx
    import pypdf
    import reportlab
except ImportError:
    pass

try:
    import pythoncom
    import win32com.client
except ImportError:
    pass

# Determinación de directorios
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
    STORAGE_DIR = get_app_data_dir()
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    STORAGE_DIR = BASE_DIR

# Archivos de configuración y sesión (Aislados para no ensuciar el PC del usuario)
EXE_CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
STORAGE_CONFIG_FILE = os.path.join(STORAGE_DIR, "config.json")
SESSION_FILE = os.path.join(STORAGE_DIR, "session.json")

DEFAULT_CONFIG = {
    "app_title": "AUTOMATIZACIÓN DE ARCHIVOS",
    "app_subtitle": "Seleccione una opción para continuar",
    "apps_script_url": "",
    "modules_repo_url": "https://github.com/GigK1R4/Modulos_Lanzador.git",
    "modules_dir": "Modulos",
    "auto_git_update": True,
    "auto_discover_modules": True,
    "network_timeout": 10,
    "modules": [
        {
            "id": "pdf_unir",
            "name": "PDF Unir",
            "title": "UNIR\nPDF",
            "script": "PDF_Unir.py",
            "description": "Unificación y generación de certificados PDF",
            "enabled_by_default": False
        },
        {
            "id": "gene_documentos",
            "name": "Gene Documentos",
            "title": "DOCUMENTOS\nMASIVOS",
            "script": "Gene_Documentos.py",
            "description": "Generación masiva de plantillas Word y conversión a PDF",
            "enabled_by_default": False
        },
        {
            "id": "formato_vinculacion",
            "name": "Formato Vinculación",
            "title": "FORMATO\nVINCULACIÓN",
            "script": "FORMATO VINCULACION.py",
            "description": "Generación masiva de formatos de vinculación FV y exportación a PDF",
            "enabled_by_default": False
        }
    ],
    "theme": {
        "window_bg": "#0b1626",
        "card_bg": "#f8fafc",
        "card_text": "#02293f",
        "card_border": "#ffffff",
        "card_active_border": "#00bcd4",
        "card_disabled_bg": "#152438",
        "card_disabled_text": "#506982",
        "badge_bg": "#0e223d",
        "badge_border": "#1b3a5f",
        "badge_dot": "#00b4d8",
        "badge_text": "#38bdf8",
        "font_family": "Segoe UI"
    }
}

def load_config():
    """
    Carga la configuración. Si existe un config.json manual junto al ejecutable lo utiliza.
    De lo contrario, lee de AppData o utiliza la configuración oficial embebida por defecto.
    NUNCA crea archivos junto al ejecutable en modo .exe.
    """
    config = dict(DEFAULT_CONFIG)
    target_file = None
    if os.path.exists(EXE_CONFIG_FILE):
        target_file = EXE_CONFIG_FILE
    elif os.path.exists(STORAGE_CONFIG_FILE):
        target_file = STORAGE_CONFIG_FILE

    if target_file and os.path.exists(target_file):
        try:
            with open(target_file, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                config.update(loaded)
        except Exception as e:
            print(f"Error al leer configuración: {e}")
    else:
        # En modo congelado (.exe), guardar copia interna en AppData (sin tocar el directorio del .exe)
        if getattr(sys, 'frozen', False):
            try:
                with open(STORAGE_CONFIG_FILE, "w", encoding="utf-8") as f:
                    json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
            except Exception:
                pass

    return config

def save_config(config):
    """Guarda la configuración respetando el almacenamiento seguro"""
    target = EXE_CONFIG_FILE if os.path.exists(EXE_CONFIG_FILE) else STORAGE_CONFIG_FILE
    try:
        with open(target, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Error al guardar config: {e}")

def save_session(username, remembered=True):
    """Guarda o borra la sesión recordada en AppData"""
    try:
        if remembered:
            with open(SESSION_FILE, "w", encoding="utf-8") as f:
                json.dump({"saved_user": username}, f)
        else:
            if os.path.exists(SESSION_FILE):
                os.remove(SESSION_FILE)
    except Exception:
        pass

def load_saved_session():
    """Carga el último usuario recordado desde AppData"""
    try:
        if os.path.exists(SESSION_FILE):
            with open(SESSION_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("saved_user", "")
    except Exception:
        pass
    return ""


class LauncherApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Iconos visuales 100% en memoria (Sin crear carpetas en el PC)
        self.icons = generate_all_icons(save_to_disk=False)

        # Cargar configuración y servicios
        self.config = load_config()
        self.auth_service = SheetAuth(self.config)
        
        modules_repo_url = self.config.get("modules_repo_url", "https://github.com/GigK1R4/Modulos_Lanzador.git")
        modules_dir_name = self.config.get("modules_dir", "Modulos_Lanzador")
        self.git_updater = GitUpdater(
            base_dir=BASE_DIR,
            repo_url=modules_repo_url,
            modules_dir_name=modules_dir_name
        )

        # Estado del usuario y vista
        self.current_user = None
        self.user_modules = []
        self.is_in_dashboard = False

        # Configuración de ventana
        self.title("Automatización de Archivos - Lanzador")
        self.geometry("1020x680")
        self.minsize(920, 600)
        
        # Color de fondo azul marino idéntico al diseño corporativo
        self.bg_color = self.config.get("theme", {}).get("window_bg", "#0b1626")
        self.configure(fg_color=self.bg_color)
        
        # Contenedor dinámico de pantallas
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True)

        # Comprobar actualizaciones de GitHub en segundo plano al abrir
        if self.config.get("auto_git_update", True):
            self.start_background_git_check()

        # Mostrar pantalla inicial de Login
        self.show_login_screen()

    # =========================================================================
    # DESCUBRIMIENTO Y GESTIÓN DE MÓDULOS DEL REPOSITORIO
    # =========================================================================
    def get_available_modules(self):
        """
        Devuelve la lista de módulos disponibles.
        Combina los metadatos de configuración con los scripts descubiertos en Modulos_Lanzador.
        """
        configured_modules = self.config.get("modules", [])
        modules_dir_path = self.git_updater.modules_dir

        seen_scripts = {}
        result_modules = []

        # 1. Procesar módulos explícitamente configurados
        for mod in configured_modules:
            script_raw = mod.get("script", "")
            base_script = os.path.basename(script_raw).lower()
            seen_scripts[base_script] = True
            
            candidates = [
                os.path.join(modules_dir_path, "Modulos", os.path.basename(script_raw)),
                os.path.join(modules_dir_path, os.path.basename(script_raw)),
                os.path.join(BASE_DIR, "Modulos", os.path.basename(script_raw)),
                os.path.join(BASE_DIR, script_raw),
                os.path.join(BASE_DIR, os.path.basename(script_raw)),
                os.path.join(modules_dir_path, script_raw)
            ]
            mod_copy = dict(mod)
            for cand in candidates:
                if os.path.exists(cand):
                    mod_copy["script"] = cand
                    break
            result_modules.append(mod_copy)

        # 2. Descubrimiento automático de cualquier nuevo script en el repositorio
        CORE_FILES = {"lanzador.py", "git_updater.py", "icon_generator.py", "sheet_auth.py", "__init__.py"}
        search_dirs = []
        modulos_in_repo = os.path.join(modules_dir_path, "Modulos")
        if os.path.isdir(modulos_in_repo):
            search_dirs.append(modulos_in_repo)
        if os.path.isdir(modules_dir_path) and modules_dir_path not in search_dirs:
            search_dirs.append(modules_dir_path)
        modulos_sub = os.path.join(BASE_DIR, "Modulos")
        if os.path.isdir(modulos_sub) and modulos_sub not in search_dirs:
            search_dirs.append(modulos_sub)
        if BASE_DIR not in search_dirs and os.path.isdir(BASE_DIR):
            search_dirs.append(BASE_DIR)

        if self.config.get("auto_discover_modules", True):
            for sdir in search_dirs:
                if not os.path.isdir(sdir):
                    continue
                try:
                    for fname in sorted(os.listdir(sdir)):
                        if fname.endswith(".py") and not fname.startswith((".", "_")) and fname.lower() not in CORE_FILES:
                            if fname.lower() not in seen_scripts:
                                seen_scripts[fname.lower()] = True
                                
                                # Formatear título amigable para la tarjeta
                                stem = os.path.splitext(fname)[0]
                                clean_words = stem.replace("-", " ").replace("_", " ").split()
                                mod_name = " ".join(clean_words).title()
                                
                                if len(clean_words) >= 2:
                                    card_title = f"{clean_words[0].upper()}\n{' '.join(clean_words[1:]).upper()}"
                                else:
                                    card_title = mod_name.upper()

                                new_mod = {
                                    "id": stem.lower(),
                                    "name": mod_name,
                                    "title": card_title,
                                    "script": os.path.join(sdir, fname),
                                    "description": f"Módulo {mod_name} ({fname})",
                                    "enabled_by_default": False
                                }
                                result_modules.append(new_mod)
                except Exception as e:
                    print(f"Error al escanear directorio {sdir}: {e}")

        return result_modules

    def check_module_permission(self, mod_data):
        """Verifica de forma flexible si el usuario activo tiene permiso para este módulo"""
        if "*" in self.user_modules:
            return True

        def _clean(val):
            return re.sub(r'[\s_\-\.\n]+', '', str(val)).lower()

        mod_id = _clean(mod_data.get("id", ""))
        mod_name = _clean(mod_data.get("name", ""))
        mod_title = _clean(mod_data.get("title", ""))
        script_file = os.path.basename(mod_data.get("script", ""))
        script_base = _clean(os.path.splitext(script_file)[0])
        script_full = _clean(script_file)

        for perm in self.user_modules:
            p = _clean(perm)
            if not p:
                continue
            if p == mod_id or p == mod_name or p == mod_title or p == script_base or p == script_full:
                return True
            if (p in mod_id or mod_id in p) and len(p) >= 3:
                return True
            if (p in script_base or script_base in p) and len(p) >= 3:
                return True

        return False

    # =========================================================================
    # CONTROL DE TAMAÑO ADAPTATIVO POR MÓDULO (MISMA VENTANA)
    # =========================================================================
    def set_window_size_for_view(self, view_type, custom_w=None, custom_h=None, min_w=None, min_h=None):
        """Ajusta suavemente las dimensiones y límites de la ventana para cada pantalla o módulo"""
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        sizes = {
            "login": (500, 560, 480, 520),
            "dashboard": (1020, 680, 920, 600),
            "pdf_unir": (1060, 720, 960, 640),
            "gene_documentos": (1080, 750, 980, 660),
            "default_module": (1020, 700, 920, 620)
        }

        target_w, target_h, default_mw, default_mh = sizes.get(view_type, sizes["default_module"])

        target_w = custom_w or target_w
        target_h = custom_h or target_h
        min_w = min_w or default_mw
        min_h = min_h or default_mh

        # Respetar límites máximos de pantalla
        max_w = max(600, screen_w - 60)
        max_h = max(520, screen_h - 90)

        final_w = min(target_w, max_w)
        final_h = min(target_h, max_h)
        final_mw = min(min_w, final_w)
        final_mh = min(min_h, final_h)

        self.minsize(final_mw, final_mh)

        cur_x = self.winfo_x()
        cur_y = self.winfo_y()

        if cur_x <= 0 or cur_y <= 0 or view_type == "login":
            pos_x = max(10, (screen_w - final_w) // 2)
            pos_y = max(10, (screen_h - final_h) // 2)
        else:
            pos_x = max(10, min(cur_x, screen_w - final_w - 20))
            pos_y = max(10, min(cur_y, screen_h - final_h - 60))

        self.geometry(f"{final_w}x{final_h}+{pos_x}+{pos_y}")

    # =========================================================================
    # GESTIÓN DE PANTALLAS
    # =========================================================================
    def clear_container(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()

    def show_login_screen(self):
        """Pantalla de inicio de sesión moderna y elegante"""
        self.is_in_dashboard = False
        self.title("Automatización de Archivos - Iniciar Sesión")
        self.set_window_size_for_view("login")
        self.clear_container()

        center_card = ctk.CTkFrame(
            self.main_container,
            fg_color="#0e1f38",
            corner_radius=24,
            border_width=2,
            border_color="#1a3557",
            width=460,
            height=500
        )
        center_card.place(relx=0.5, rely=0.5, anchor="center")
        center_card.pack_propagate(False)

        # Título del Login
        title_lbl = ctk.CTkLabel(
            center_card,
            text=self.config.get("app_title", "AUTOMATIZACIÓN DE ARCHIVOS"),
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#ffffff"
        )
        title_lbl.pack(pady=(35, 6))

        sub_lbl = ctk.CTkLabel(
            center_card,
            text="Ingrese sus credenciales para acceder a los módulos",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#8ea3bf"
        )
        sub_lbl.pack(pady=(0, 25))

        # Campo Usuario
        user_lbl = ctk.CTkLabel(
            center_card,
            text="Usuario:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        user_lbl.pack(anchor="w", padx=45, pady=(0, 4))

        saved_user = load_saved_session()
        self.user_entry = ctk.CTkEntry(
            center_card,
            width=370,
            height=46,
            corner_radius=12,
            fg_color="#0a1628",
            border_color="#1e385c",
            border_width=1.5,
            text_color="#ffffff",
            placeholder_text="Ingrese su usuario",
            font=ctk.CTkFont(size=14)
        )
        self.user_entry.pack(pady=(0, 16))
        if saved_user:
            self.user_entry.insert(0, saved_user)

        # Campo Contraseña
        pass_lbl = ctk.CTkLabel(
            center_card,
            text="Contraseña:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        pass_lbl.pack(anchor="w", padx=45, pady=(0, 4))

        self.pass_entry = ctk.CTkEntry(
            center_card,
            show="•",
            width=370,
            height=46,
            corner_radius=12,
            fg_color="#0a1628",
            border_color="#1e385c",
            border_width=1.5,
            text_color="#ffffff",
            placeholder_text="••••••••",
            font=ctk.CTkFont(size=15)
        )
        self.pass_entry.pack(pady=(0, 10))
        self.pass_entry.bind("<Return>", lambda e: self.attempt_login())

        # Checkbox Recordar usuario
        self.remember_var = ctk.BooleanVar(value=bool(saved_user))
        self.remember_cb = ctk.CTkCheckBox(
            center_card,
            text="Recordar usuario",
            variable=self.remember_var,
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8",
            fg_color="#00b4d8",
            hover_color="#0096c7",
            border_color="#244570"
        )
        self.remember_cb.pack(anchor="w", padx=45, pady=(0, 15))

        # Mensaje de Error
        self.login_msg = ctk.CTkLabel(
            center_card,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="#f87171",
            wraplength=370
        )
        self.login_msg.pack(pady=(0, 8))

        # Botón Ingresar
        self.btn_login = ctk.CTkButton(
            center_card,
            text="INGRESAR",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#00b4d8",
            hover_color="#0096c7",
            text_color="#031629",
            height=48,
            width=370,
            corner_radius=12,
            command=self.attempt_login
        )
        self.btn_login.pack(pady=(0, 15))

        # Pie de login
        footer_frame = ctk.CTkFrame(center_card, fg_color="transparent")
        footer_frame.pack(side="bottom", pady=15)

        footer_lbl = ctk.CTkLabel(
            footer_frame,
            text="⚙ Realizado Por DKSoftware",
            font=ctk.CTkFont(size=11),
            text_color="#577399"
        )
        footer_lbl.pack(side="left")

        # Botón sutil para configurar URL de Google Sheet
        btn_gear = ctk.CTkButton(
            footer_frame,
            text="⚙",
            width=26,
            height=24,
            corner_radius=6,
            fg_color="transparent",
            hover_color="#18365c",
            text_color="#577399",
            font=ctk.CTkFont(size=12),
            command=self.open_config_dialog
        )
        btn_gear.pack(side="left", padx=(6, 0))

        # Foco en contraseña si ya hay usuario guardado
        if saved_user:
            self.pass_entry.focus_set()
        else:
            self.user_entry.focus_set()

    def attempt_login(self):
        """Valida credenciales contra Google Apps Script en segundo plano"""
        user = self.user_entry.get().strip()
        pwd = self.pass_entry.get().strip()

        if not user or not pwd:
            self.login_msg.configure(text="Por favor ingrese usuario y contraseña.")
            return

        if not self.config.get("apps_script_url"):
            self.login_msg.configure(
                text="Falta configurar la URL de Apps Script en config.json o en el botón ⚙.",
                text_color="#f87171"
            )
            return

        self.btn_login.configure(state="disabled", text="VERIFICANDO...")
        self.login_msg.configure(text="", text_color="#38bdf8")

        def _bg_auth():
            self.auth_service = SheetAuth(self.config)
            success, modules_allowed, message = self.auth_service.autenticar(user, pwd)
            if success:
                save_session(user, remembered=self.remember_var.get())
                self.after(0, lambda: self._on_auth_success(user, modules_allowed))
            else:
                self.after(0, lambda: self._on_auth_failed(message))

        threading.Thread(target=_bg_auth, daemon=True).start()

    def _on_auth_success(self, user, modules):
        self.current_user = user
        self.user_modules = modules
        self.show_dashboard_screen()

    def _on_auth_failed(self, message):
        self.btn_login.configure(state="normal", text="INGRESAR")
        self.login_msg.configure(text=message, text_color="#f87171")

    # =========================================================================
    # DASHBOARD / MENÚ PRINCIPAL
    # =========================================================================
    def show_dashboard_screen(self):
        """Pantalla principal de tarjetas de módulos con diseño moderno"""
        self.is_in_dashboard = True
        self.title("Automatización de Archivos - Menú Principal")
        self.set_window_size_for_view("dashboard")
        self.clear_container()

        # Barra superior con Usuario y Cerrar Sesión
        top_bar = ctk.CTkFrame(self.main_container, fg_color="transparent")
        top_bar.pack(fill="x", padx=30, pady=(18, 0))

        user_badge = ctk.CTkLabel(
            top_bar,
            text=f"👤 Usuario: {self.current_user}",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#94a3b8"
        )
        user_badge.pack(side="left")

        btn_logout = ctk.CTkButton(
            top_bar,
            text="Cerrar Sesión",
            font=ctk.CTkFont(size=12),
            width=110,
            height=30,
            corner_radius=8,
            fg_color="#182c47",
            hover_color="#243f66",
            text_color="#94a3b8",
            command=self.show_login_screen
        )
        btn_logout.pack(side="right")

        # Encabezado Principal
        header_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header_frame.pack(fill="x", pady=(15, 10))

        title_lbl = ctk.CTkLabel(
            header_frame,
            text=self.config.get("app_title", "AUTOMATIZACIÓN DE ARCHIVOS"),
            font=ctk.CTkFont(family="Segoe UI", size=26, weight="bold"),
            text_color="#ffffff"
        )
        title_lbl.pack(pady=(0, 4))

        sub_lbl = ctk.CTkLabel(
            header_frame,
            text=self.config.get("app_subtitle", "Seleccione una opción para continuar"),
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color="#8ea3bf"
        )
        sub_lbl.pack(pady=(0, 14))

        # Obtener lista dinámica de módulos desde Modulos_Lanzador y config
        modulos = self.get_available_modules()
        total_mods = len(modulos)
        user_mods_count = sum(1 for m in modulos if self.check_module_permission(m))

        # Badge / Píldora de estado ("● Conectado...")
        badge_frame = ctk.CTkFrame(
            header_frame,
            fg_color="#0e223d",
            border_width=1,
            border_color="#1d3b61",
            corner_radius=18,
            height=34
        )
        badge_frame.pack(pady=(0, 15))

        badge_text = f"● Conectado ({user_mods_count} de {total_mods} módulos habilitados)"

        badge_lbl = ctk.CTkLabel(
            badge_frame,
            text=badge_text,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#00b4d8",
            padx=16,
            pady=4
        )
        badge_lbl.pack()

        # Contenedor de Tarjetas de Módulos (Adaptativo: centrado si <= 4, scrollable si > 4)
        if total_mods <= 4:
            cards_container = ctk.CTkFrame(self.main_container, fg_color="transparent")
            cards_container.pack(expand=True, pady=(10, 20))
        else:
            cards_container = ctk.CTkScrollableFrame(
                self.main_container,
                fg_color="transparent",
                orientation="horizontal",
                height=250
            )
            cards_container.pack(fill="x", padx=40, pady=(10, 20))

        # Renderizar cada módulo
        for idx, mod in enumerate(modulos):
            tiene_permiso = self.check_module_permission(mod)
            self.create_module_card(cards_container, mod, tiene_permiso, is_first=(idx == 0))

        # Barra inferior de estado (Git/GitHub y notificaciones)
        bottom_bar = ctk.CTkFrame(self.main_container, fg_color="transparent")
        bottom_bar.pack(fill="x", side="bottom", padx=30, pady=15)

        self.git_status_lbl = ctk.CTkLabel(
            bottom_bar,
            text="Git: Comprobando actualizaciones de módulos...",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#64748b"
        )
        self.git_status_lbl.pack(side="left")

        btn_sync = ctk.CTkButton(
            bottom_bar,
            text="⟳ Sincronizar Git",
            font=ctk.CTkFont(size=12),
            width=140,
            height=28,
            corner_radius=8,
            fg_color="#10253f",
            hover_color="#18365c",
            text_color="#38bdf8",
            command=self.manual_git_sync
        )
        btn_sync.pack(side="right")

    def create_module_card(self, parent, mod_data, habilitado, is_first=False):
        """
        Crea una tarjeta rectangular con bordes redondeados, icono superior en memoria
        y título en negrita abajo, exactamente como en el diseño de referencia.
        """
        mod_name = mod_data.get("name", "")
        mod_title = mod_data.get("title", mod_name.upper())
        mod_script = mod_data.get("script", "")

        # Estilo de tarjeta habilitada vs bloqueada
        if habilitado:
            card_bg = "#f8fafc"
            card_border = "#00bcd4" if is_first else "#f8fafc"
            border_width = 3 if is_first else 0
            text_color = "#02293f"
            hover_border = "#00b4d8"
            cursor = "hand2"
        else:
            card_bg = "#112239"
            card_border = "#1b3554"
            border_width = 1
            text_color = "#556e8f"
            hover_border = "#1b3554"
            cursor = "no"

        # Frame de la tarjeta (ancho ~200, alto ~220)
        card = ctk.CTkFrame(
            parent,
            fg_color=card_bg,
            border_color=card_border,
            border_width=border_width,
            corner_radius=22,
            width=200,
            height=220,
            cursor=cursor
        )
        card.pack(side="left", padx=14, pady=10)
        card.pack_propagate(False)

        # Cargar icono directamente desde memoria
        if not habilitado:
            pil_img = self.icons.get("lock")
        else:
            pil_img = get_icon_for_module(f"{mod_name} {mod_script}", self.icons)

        img_widget = None
        if pil_img:
            try:
                ctk_icon = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(64, 64))
                img_widget = ctk.CTkLabel(card, text="", image=ctk_icon, cursor=cursor)
                img_widget.pack(pady=(28, 12))
            except Exception as e:
                print(f"Error al renderizar icono: {e}")

        # Título en negrita centrado
        lbl_title = ctk.CTkLabel(
            card,
            text=mod_title,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=text_color,
            justify="center",
            cursor=cursor
        )
        lbl_title.pack(pady=(0, 10))

        if not habilitado:
            lbl_lock = ctk.CTkLabel(
                card,
                text="(No habilitado)",
                font=ctk.CTkFont(size=11),
                text_color="#f87171"
            )
            lbl_lock.pack()

        # Eventos de clic y hover
        def _on_enter(e):
            if habilitado:
                card.configure(border_color=hover_border, border_width=3)

        def _on_leave(e):
            if habilitado:
                if is_first:
                    card.configure(border_color="#00bcd4", border_width=3)
                else:
                    card.configure(border_color="#f8fafc", border_width=0)

        def _on_click(e):
            if habilitado:
                self.launch_module(mod_name, mod_script)
            else:
                messagebox.showwarning(
                    "Módulo no habilitado",
                    f"El usuario '{self.current_user}' no tiene permisos para acceder a '{mod_name}'.\n\nContacte al administrador para habilitar este permiso."
                )

        # Asociar eventos a los componentes de la tarjeta
        for w in [card, lbl_title]:
            w.bind("<Enter>", _on_enter)
            w.bind("<Leave>", _on_leave)
            w.bind("<Button-1>", _on_click)

        if img_widget:
            img_widget.bind("<Enter>", _on_enter)
            img_widget.bind("<Leave>", _on_leave)
            img_widget.bind("<Button-1>", _on_click)

    # =========================================================================
    # EJECUCIÓN DE MÓDULOS (INTEGRADOS EN LA MISMA VENTANA PRINCIPAL)
    # =========================================================================
    def load_embedded_module(self, script_path, name):
        """
        Importa dinámicamente el script del módulo y obtiene su vista CTkFrame
        para incrustarla en la misma ventana, sin ventanas emergentes.
        """
        try:
            import importlib.util
            script_abs = os.path.abspath(script_path)
            script_dir = os.path.dirname(script_abs)
            if script_dir not in sys.path:
                sys.path.insert(0, script_dir)

            mod_name = os.path.splitext(os.path.basename(script_path))[0]
            spec = importlib.util.spec_from_file_location(mod_name, script_abs)
            if spec and spec.loader:
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)

                # 1. Función estándar create_view(parent, on_back)
                if hasattr(module, "create_view"):
                    return module.create_view(self.main_container, on_back=self.show_dashboard_screen)

                # 2. Clases de vista conocidas
                for cls_name in ["PDFUnirView", "GeneDocumentosView", "FormatoVinculacionView", "PDFCompressorView", "ModuleView"]:
                    if hasattr(module, cls_name):
                        cls = getattr(module, cls_name)
                        return cls(self.main_container, on_back=self.show_dashboard_screen)

        except Exception as e:
            messagebox.showerror(
                "Error al cargar módulo",
                f"Ocurrió un error al cargar la vista de '{name}':\n{str(e)}"
            )
            return None

        return None

    def launch_module(self, name, script_filename):
        """Carga y muestra el módulo dentro de la misma ventana principal"""
        modules_dir_path = self.git_updater.modules_dir

        candidates = [
            script_filename,
            os.path.join(modules_dir_path, "Modulos", os.path.basename(script_filename)),
            os.path.join(modules_dir_path, os.path.basename(script_filename)),
            os.path.join(BASE_DIR, "Modulos", os.path.basename(script_filename)),
            os.path.join(BASE_DIR, script_filename),
            os.path.join(BASE_DIR, "Modulos_Lanzador", os.path.basename(script_filename))
        ]

        script_path = None
        for cand in candidates:
            if cand and os.path.isfile(cand):
                script_path = os.path.abspath(cand)
                break

        if not script_path:
            messagebox.showerror(
                "Archivo no encontrado",
                f"No se encontró el script del módulo:\n{script_filename}\nBuscado en: {modules_dir_path}"
            )
            return

        # Intentar incrustar la vista en la misma ventana principal
        self.clear_container()
        view = self.load_embedded_module(script_path, name)
        if view is not None:
            self.is_in_dashboard = False
            self.title(f"Automatización de Archivos - {name}")

            # Ajustar automáticamente la ventana al tamaño óptimo de este módulo
            pref_size = getattr(view, "preferred_size", None)
            min_size = getattr(view, "min_size", None)
            if pref_size:
                mw = min_size[0] if min_size else None
                mh = min_size[1] if min_size else None
                self.set_window_size_for_view("custom", custom_w=pref_size[0], custom_h=pref_size[1], min_w=mw, min_h=mh)
            elif "unir" in name.lower() or "pdf" in name.lower():
                self.set_window_size_for_view("pdf_unir")
            elif "documento" in name.lower() or "gene" in name.lower():
                self.set_window_size_for_view("gene_documentos")
            elif "vincula" in name.lower() or "formato" in name.lower() or "contrato" in name.lower():
                self.set_window_size_for_view("custom", custom_w=1040, custom_h=720, min_w=900, min_h=620)
            else:
                self.set_window_size_for_view("default_module")

            view.pack(fill="both", expand=True)
            return
        else:
            self.show_dashboard_screen()

    def safe_after(self, ms, callback):
        """Ejecuta callback en el hilo principal de forma segura si la ventana existe"""
        try:
            if self.winfo_exists():
                self.after(ms, callback)
        except Exception:
            pass

    # =========================================================================
    # ACTUALIZACIONES GIT / GITHUB DEL REPOSITORIO DE MÓDULOS
    # =========================================================================
    def start_background_git_check(self):
        """Comprueba e instala actualizaciones del repositorio de módulos en segundo plano"""
        def _bg():
            self.safe_after(0, lambda: self._update_git_ui("Git: Buscando actualizaciones en GitHub...", "#38bdf8"))
            try:
                def _on_status(msg):
                    self.safe_after(0, lambda: self._update_git_ui(f"Git: {msg}", "#38bdf8"))

                has_changes, message = self.git_updater.update_all(on_status_callback=_on_status)

                if has_changes:
                    self.safe_after(0, lambda: self._update_git_ui(f"Git: {message}", "#4ade80"))
                    self.safe_after(400, self.refresh_dashboard_if_visible)
                else:
                    self.safe_after(0, lambda: self._update_git_ui(f"Git: {message}", "#94a3b8"))
            except Exception as e:
                self.safe_after(0, lambda: self._update_git_ui(f"Git: {str(e)}", "#f87171"))

        threading.Thread(target=_bg, daemon=True).start()

    def refresh_dashboard_if_visible(self):
        """Refresca la pantalla de módulos si está actualmente activa"""
        if self.current_user and self.is_in_dashboard:
            self.show_dashboard_screen()

    def manual_git_sync(self):
        """Sincronización manual provocada por el usuario"""
        if hasattr(self, "git_status_lbl"):
            self.git_status_lbl.configure(text="Sincronizando con Git remoto...", text_color="#38bdf8")
        self.start_background_git_check()

    def _update_git_ui(self, text, color):
        if hasattr(self, "git_status_lbl") and self.git_status_lbl.winfo_exists():
            self.git_status_lbl.configure(text=text, text_color=color)

    # =========================================================================
    # CONFIGURACIÓN RÁPIDA DE GOOGLE APPS SCRIPT WEB APP
    # =========================================================================
    def open_config_dialog(self):
        """Ventana modal para editar la URL de la Web App de Apps Script"""
        dlg = ctk.CTkToplevel(self)
        dlg.title("Configuración de Servidor de Autenticación")
        dlg.geometry("560x360")
        dlg.transient(self)
        dlg.grab_set()

        lbl = ctk.CTkLabel(
            dlg,
            text="Configurar Google Apps Script Web App",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")
        )
        lbl.pack(pady=(20, 10))

        help_lbl = ctk.CTkLabel(
            dlg,
            text="Ingresa la URL de la Web App generada en Apps Script (terminada en /exec).\n"
                 "🔒 La hoja de cálculo permanece 100% privada y segura en tu cuenta.",
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8",
            wraplength=500
        )
        help_lbl.pack(pady=(0, 15))

        entry_url = ctk.CTkEntry(
            dlg,
            placeholder_text="https://script.google.com/macros/s/.../exec",
            width=500,
            height=40
        )
        entry_url.pack(padx=20, pady=(0, 15))
        entry_url.insert(0, self.config.get("apps_script_url", ""))

        status_lbl = ctk.CTkLabel(dlg, text="", font=ctk.CTkFont(size=11), text_color="#38bdf8")
        status_lbl.pack(pady=(0, 10))

        def _test_connection():
            url_to_test = entry_url.get().strip()
            if not url_to_test:
                status_lbl.configure(text="Por favor ingresa una URL primero.", text_color="#f87171")
                return
            status_lbl.configure(text="Probando conexión...", text_color="#38bdf8")
            temp_config = dict(self.config)
            temp_config["apps_script_url"] = url_to_test
            test_service = SheetAuth(temp_config)
            
            def _bg_test():
                ok, msg = test_service.verificar_conexion()
                if ok:
                    self.safe_after(0, lambda: status_lbl.configure(
                        text=f"✓ Conexión exitosa: {msg}",
                        text_color="#4ade80"
                    ))
                else:
                    self.safe_after(0, lambda: status_lbl.configure(
                        text=f"✗ Error al conectar: {msg}",
                        text_color="#f87171"
                    ))

            threading.Thread(target=_bg_test, daemon=True).start()

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(pady=10)

        btn_test = ctk.CTkButton(
            btn_row,
            text="Probar Conexión",
            command=_test_connection,
            fg_color="#182c47",
            hover_color="#243f66",
            text_color="#38bdf8",
            width=140
        )
        btn_test.pack(side="left", padx=8)

        def _save():
            new_url = entry_url.get().strip()
            self.config["apps_script_url"] = new_url
            save_config(self.config)
            self.auth_service = SheetAuth(self.config)
            messagebox.showinfo("Configuración", "URL de Google Apps Script guardada exitosamente.")
            dlg.destroy()

        btn_save = ctk.CTkButton(
            btn_row,
            text="Guardar Cambios",
            command=_save,
            fg_color="#00b4d8",
            hover_color="#0096c7",
            text_color="#02293f",
            width=140
        )
        btn_save.pack(side="left", padx=8)


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    app = LauncherApp()
    app.mainloop()
