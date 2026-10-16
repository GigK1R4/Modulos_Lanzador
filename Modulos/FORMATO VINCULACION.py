"""
=============================================================================
GENERADOR DE FORMATO DE VINCULACIÓN (FV) - MODERNO (CUSTOMTKINTER)
=============================================================================
Combina datos de empleados desde un archivo Excel (.xlsx, .xlsm) en la plantilla
oficial del Formato de Vinculación (FV) y exporta automáticamente a formato PDF
utilizando Microsoft Excel (COM).
Diseñado para integrarse dentro del Lanzador Modular o ejecutarse individualmente.
"""
import os
import sys
import re
import unicodedata
import warnings
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from openpyxl import load_workbook

# Silenciar advertencias de formatos de imagen no vectoriales en openpyxl (como WMF/EMF de logos)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --- COM / Excel automation ---
try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    pythoncom = None
    win32 = None


# -----------------------------------------------------------------------------
# RUTAS DE RECURSOS
# -----------------------------------------------------------------------------
def resource_path(relative_path: str) -> str:
    """Ruta correcta para archivos incluidos con PyInstaller o ejecución directa"""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# -----------------------------------------------------------------------------
# CONFIGURACIÓN DEL MAPEO DE CELDAS (FORMATO VINCULACIÓN)
# -----------------------------------------------------------------------------
MAPPING_CONFIG = {
    "FECHADELANOVEDAD": "F10",
    "DEPENDENCIA": "Q10",
    "NOMBREDELEMPLEADO": "I13",
    "CEDULA": "K16",
    "FECHADENACIMIENTO": "F19",
    "DIRECCION": "O19",
    "TIPODECONTRATO": "F22",
    "CARGOS": "M22",
    "ASIGNACIONSALARIAL": "T22",
    "SALUD": "J24",
    "PENSION": "T24",
}


# -----------------------------------------------------------------------------
# UTILIDADES Y LÓGICA DE NEGOCIO
# -----------------------------------------------------------------------------
def normalize_key(key: str) -> str:
    """Normaliza texto eliminando acentos, espacios y caracteres especiales para emparejamiento flexible"""
    if not key:
        return ""
    n = unicodedata.normalize("NFKD", str(key))
    cleaned = "".join(c for c in n if not unicodedata.combining(c))
    return re.sub(r"[^A-Za-z0-9]", "", cleaned).upper()


def safe_filename(text: str, max_len: int = 80) -> str:
    """Sanitiza una cadena para que sea un nombre de archivo válido en Windows"""
    text = str(text) if text is not None else ""
    text = text.strip()
    text = re.sub(r'[\\/*?:"<>|]+', "_", text)
    text = re.sub(r"\s+", " ", text)
    return (text[:max_len] if len(text) > max_len else text) or "documento"


def get_sheet_names(excel_path: str):
    """Devuelve la lista de nombres de hojas de un libro de Excel"""
    try:
        wb = load_workbook(excel_path, read_only=True, keep_links=False)
        names = list(wb.sheetnames)
        wb.close()
        return names
    except Exception:
        return []


def get_sheet_columns(excel_path: str, sheet_name: str, header_row: int = 1):
    """Obtiene los nombres de columnas de la fila de encabezados especificada"""
    try:
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active
        headers = []
        for row in ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True):
            for cell in row:
                if cell is not None and str(cell).strip():
                    headers.append(str(cell).strip())
        wb.close()
        return headers
    except Exception:
        return []


def iter_records_from_source(origen_xlsx: str, hoja_origen: str, header_row: int = 1, start_row: int = 2):
    """Generador que itera fila a fila los registros de datos del Excel de origen"""
    wb = load_workbook(origen_xlsx, data_only=True)
    if hoja_origen in wb.sheetnames:
        ws = wb[hoja_origen]
    else:
        ws = wb.active

    headers = []
    for cell in ws[header_row]:
        val = str(cell.value).strip() if cell.value is not None else ""
        headers.append(val)

    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, values_only=True):
        if all(v is None for v in row):
            continue
        record = {}
        for i, value in enumerate(row):
            if i < len(headers) and headers[i]:
                record[headers[i]] = value
        yield record

    wb.close()


def fill_template_for_record(plantilla_xlsx: str, hoja_destino: str, mapping: dict, record: dict, output_xlsx_path: str):
    """Rellena la plantilla de Excel con los datos del registro y la guarda en la ruta indicada"""
    wb = load_workbook(plantilla_xlsx)
    if hoja_destino in wb.sheetnames:
        ws = wb[hoja_destino]
    else:
        ws = wb.active

    # Diccionario con claves normalizadas para búsqueda flexible
    norm_record = {normalize_key(k): v for k, v in record.items()}

    for field, cell_addr in mapping.items():
        norm_field = normalize_key(field)
        val = record.get(field)
        if val is None:
            val = norm_record.get(norm_field)

        if val is not None:
            ws[cell_addr].value = val

    Path(output_xlsx_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_xlsx_path)
    wb.close()


# -----------------------------------------------------------------------------
# EXPORTADOR EXCEL -> PDF MEDIANTE COM
# -----------------------------------------------------------------------------
class ExcelPdfExporter:
    """Maneja la instancia de Microsoft Excel en segundo plano para exportar hojas a PDF"""
    def __init__(self):
        self.excel = None
        self.com_inited = False

    def start(self):
        if win32 is None or pythoncom is None:
            raise RuntimeError(
                "No se pudo cargar pywin32 (win32com/pythoncom).\n"
                "Instala la librería con: pip install pywin32"
            )
        pythoncom.CoInitialize()
        self.com_inited = True
        try:
            self.excel = win32.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
        except Exception as e:
            self.excel = None
            raise RuntimeError(
                "No se pudo iniciar Microsoft Excel para convertir a PDF.\n"
                "Verifica que Microsoft Excel esté instalado y activado en este equipo.\n\n"
                f"Detalle técnico: {repr(e)}"
            )

    def export(self, input_xlsx: str, output_pdf: str, sheet_name: str = None):
        if not self.excel:
            raise RuntimeError("Excel no está iniciado.")

        input_xlsx = str(Path(input_xlsx).resolve())
        output_pdf = str(Path(output_pdf).resolve())
        wb = None
        try:
            wb = self.excel.Workbooks.Open(input_xlsx)
            if sheet_name:
                try:
                    ws = wb.Worksheets(sheet_name)
                except Exception:
                    ws = wb.ActiveSheet
            else:
                ws = wb.ActiveSheet

            # Type=0 representa xlTypePDF en la API de Excel
            ws.ExportAsFixedFormat(Type=0, Filename=output_pdf, From=1, To=1)
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def close(self):
        if self.excel is not None:
            try:
                self.excel.Quit()
            except Exception:
                pass
            self.excel = None

        if self.com_inited and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self.com_inited = False


# -----------------------------------------------------------------------------
# VISTA PRINCIPAL (CUSTOMTKINTER - ESTÉTICA MODERNA IDÉNTICA AL LANZADOR)
# -----------------------------------------------------------------------------
class FormatoVinculacionView(ctk.CTkFrame):
    """
    Vista modular para la Generación de Formatos de Vinculación FV.
    Permite incrustarse sin parpadeos dentro del contenedor del Lanzador.
    """
    preferred_size = (1040, 720)
    min_size = (900, 620)

    def __init__(self, parent, on_back=None, *args, **kwargs):
        super().__init__(parent, fg_color="transparent", *args, **kwargs)
        self.on_back = on_back

        # Variables de estado
        self.path_origen = tk.StringVar(value="")
        self.path_plantilla = tk.StringVar(value="")
        self.path_destino = tk.StringVar(value="")
        self.hoja_origen_var = tk.StringVar(value="")
        self.hoja_plantilla_var = tk.StringVar(value="")
        self.col_nombre = tk.StringVar(value="NOMBREDELEMPLEADO")

        self.export_pdf_var = tk.BooleanVar(value=True)
        self.keep_xlsx_var = tk.BooleanVar(value=True)
        self.is_running = False

        self._build_ui()

    def safe_after(self, ms, callback):
        """Ejecuta una acción en el hilo principal de tkinter de forma segura"""
        try:
            if self.winfo_exists():
                self.after(ms, callback)
        except Exception:
            pass

    def _build_ui(self):
        # 1. Barra superior: botón Volver
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(fill="x", padx=25, pady=(15, 6))

        if self.on_back:
            btn_back = ctk.CTkButton(
                top_bar,
                text="← Volver al Menú Principal",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                width=190,
                height=32,
                corner_radius=8,
                fg_color="#182c47",
                hover_color="#243f63",
                text_color="#94a3b8",
                command=self.on_back
            )
            btn_back.pack(side="left")

        # 2. Tarjeta de Cabecera
        header_card = ctk.CTkFrame(
            self,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        header_card.pack(fill="x", padx=25, pady=(4, 12))

        lbl_title = ctk.CTkLabel(
            header_card,
            text="FORMATO DE VINCULACIÓN (FV)",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#ffffff"
        )
        lbl_title.pack(anchor="w", padx=20, pady=(12, 2))

        lbl_sub = ctk.CTkLabel(
            header_card,
            text="Rellene automáticamente la plantilla oficial FV con los datos de sus empleados en Excel y genere los PDFs al instante.",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#8ea3bf"
        )
        lbl_sub.pack(anchor="w", padx=20, pady=(0, 12))

        # 3. Contenedor central (2 Columnas: Archivos Fuente y Parámetros de Destino)
        center_frame = ctk.CTkFrame(self, fg_color="transparent")
        center_frame.pack(fill="both", expand=True, padx=25, pady=(0, 10))

        # --- Columna Izquierda: Archivos Fuente (Datos y Plantilla) ---
        left_card = ctk.CTkFrame(
            center_frame,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        left_card.pack(side="left", fill="both", expand=True, padx=(0, 8))

        lbl_sec1 = ctk.CTkLabel(
            left_card,
            text="1. Archivos de Entrada",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#00b4d8"
        )
        lbl_sec1.pack(anchor="w", padx=18, pady=(12, 8))

        # Botón Excel con Datos
        btn_orig = ctk.CTkButton(
            left_card,
            text="📊 Seleccionar Excel con Datos (.xlsx, .xlsm)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#182c47",
            hover_color="#243f63",
            text_color="#38bdf8",
            height=34,
            corner_radius=8,
            command=self.select_origen
        )
        btn_orig.pack(fill="x", padx=18, pady=(2, 2))

        self.lbl_orig_status = ctk.CTkLabel(
            left_card,
            text="Ningún archivo seleccionado",
            font=ctk.CTkFont(size=11),
            text_color="#64748b"
        )
        self.lbl_orig_status.pack(anchor="w", padx=22, pady=(0, 6))

        # Combobox Hoja de Datos
        lbl_h_orig = ctk.CTkLabel(
            left_card,
            text="Hoja de Datos a procesar:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_h_orig.pack(anchor="w", padx=18, pady=(2, 2))

        self.cb_hoja_origen = ctk.CTkComboBox(
            left_card,
            values=["(Cargue un Excel primero)"],
            height=32,
            corner_radius=8,
            fg_color="#071426",
            border_color="#1d3b61",
            text_color="#ffffff",
            button_color="#182c47",
            button_hover_color="#243f63",
            command=self.on_sheet_origen_changed
        )
        self.cb_hoja_origen.pack(fill="x", padx=18, pady=(0, 10))

        # Botón Plantilla Excel FV
        btn_tmpl = ctk.CTkButton(
            left_card,
            text="📑 Seleccionar Plantilla FV (.xlsx, .xlsm)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#182c47",
            hover_color="#243f63",
            text_color="#38bdf8",
            height=34,
            corner_radius=8,
            command=self.select_plantilla
        )
        btn_tmpl.pack(fill="x", padx=18, pady=(4, 2))

        self.lbl_tmpl_status = ctk.CTkLabel(
            left_card,
            text="Ninguna plantilla seleccionada",
            font=ctk.CTkFont(size=11),
            text_color="#64748b"
        )
        self.lbl_tmpl_status.pack(anchor="w", padx=22, pady=(0, 6))

        # Combobox Hoja a Rellenar
        lbl_h_tmpl = ctk.CTkLabel(
            left_card,
            text="Hoja a Rellenar en la Plantilla:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_h_tmpl.pack(anchor="w", padx=18, pady=(2, 2))

        self.cb_hoja_plantilla = ctk.CTkComboBox(
            left_card,
            values=["(Cargue la plantilla primero)"],
            height=32,
            corner_radius=8,
            fg_color="#071426",
            border_color="#1d3b61",
            text_color="#ffffff",
            button_color="#182c47",
            button_hover_color="#243f63"
        )
        self.cb_hoja_plantilla.pack(fill="x", padx=18, pady=(0, 12))

        # --- Columna Derecha: Destino, Opciones y Ejecución ---
        right_card = ctk.CTkFrame(
            center_frame,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        right_card.pack(side="left", fill="both", expand=True, padx=(8, 0))

        lbl_sec2 = ctk.CTkLabel(
            right_card,
            text="2. Destino y Configuración",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#00b4d8"
        )
        lbl_sec2.pack(anchor="w", padx=18, pady=(12, 8))

        # Carpeta Destino
        btn_dest = ctk.CTkButton(
            right_card,
            text="📁 Seleccionar Carpeta de Destino",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#10253f",
            hover_color="#18365c",
            text_color="#e2e8f0",
            height=34,
            corner_radius=8,
            command=self.select_destino
        )
        btn_dest.pack(fill="x", padx=18, pady=(2, 2))

        self.lbl_dest_status = ctk.CTkLabel(
            right_card,
            text="Carpeta: (no seleccionada)",
            font=ctk.CTkFont(size=11),
            text_color="#64748b"
        )
        self.lbl_dest_status.pack(anchor="w", padx=22, pady=(0, 8))

        # Columna ID para nombre de archivo
        lbl_col = ctk.CTkLabel(
            right_card,
            text="Columna Identificadora (Nombre del archivo):",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_col.pack(anchor="w", padx=18, pady=(2, 2))

        self.entry_col_id = ctk.CTkComboBox(
            right_card,
            values=["NOMBREDELEMPLEADO", "CEDULA", "NOMBRE"],
            variable=self.col_nombre,
            height=32,
            corner_radius=8,
            fg_color="#071426",
            border_color="#1d3b61",
            text_color="#ffffff",
            button_color="#182c47",
            button_hover_color="#243f63"
        )
        self.entry_col_id.pack(fill="x", padx=18, pady=(0, 10))

        # Opciones de generación (Switches)
        opt_frame = ctk.CTkFrame(right_card, fg_color="transparent")
        opt_frame.pack(fill="x", padx=18, pady=(0, 10))

        self.sw_pdf = ctk.CTkSwitch(
            opt_frame,
            text="Exportar automáticamente a PDF (con Excel)",
            variable=self.export_pdf_var,
            font=ctk.CTkFont(size=12),
            progress_color="#00b4d8",
            button_color="#ffffff",
            button_hover_color="#e2e8f0"
        )
        self.sw_pdf.pack(anchor="w", pady=3)

        self.sw_xlsx = ctk.CTkSwitch(
            opt_frame,
            text="Conservar archivos Excel (.xlsx) individuales",
            variable=self.keep_xlsx_var,
            font=ctk.CTkFont(size=12),
            progress_color="#00b4d8",
            button_color="#ffffff",
            button_hover_color="#e2e8f0"
        )
        self.sw_xlsx.pack(anchor="w", pady=3)

        # Botón de Inicio
        self.btn_run = ctk.CTkButton(
            right_card,
            text="⚡ GENERAR FORMATOS Y PDF",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color="#00b4d8",
            hover_color="#0096c7",
            text_color="#02293f",
            height=38,
            corner_radius=8,
            command=self.start_thread
        )
        self.btn_run.pack(fill="x", padx=18, pady=(10, 8))

        # 4. Sección Inferior: Barra de Progreso y Consola de Registro
        bottom_frame = ctk.CTkFrame(
            self,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        bottom_frame.pack(fill="x", padx=25, pady=(0, 14))

        # Barra de progreso y texto de estado
        prog_header = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        prog_header.pack(fill="x", padx=18, pady=(10, 4))

        self.lbl_prog_title = ctk.CTkLabel(
            prog_header,
            text="Registro de Actividad y Progreso",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_prog_title.pack(side="left")

        self.lbl_progress = ctk.CTkLabel(
            prog_header,
            text="Listo para procesar",
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        )
        self.lbl_progress.pack(side="right")

        self.progress_bar = ctk.CTkProgressBar(
            bottom_frame,
            height=10,
            corner_radius=5,
            progress_color="#00b4d8",
            fg_color="#182c47"
        )
        self.progress_bar.pack(fill="x", padx=18, pady=(0, 8))
        self.progress_bar.set(0.0)

        # Consola de Registro
        self.log_textbox = ctk.CTkTextbox(
            bottom_frame,
            height=110,
            corner_radius=8,
            fg_color="#071426",
            border_color="#1d3b61",
            border_width=1,
            text_color="#e2e8f0",
            font=ctk.CTkFont(family="Consolas", size=11)
        )
        self.log_textbox.pack(fill="x", padx=18, pady=(0, 12))

    # -------------------------------------------------------------------------
    # SELECCIÓN DE ARCHIVOS Y HOJAS
    # -------------------------------------------------------------------------
    def select_origen(self):
        f = filedialog.askopenfilename(
            title="Seleccionar Excel con Datos de Origen",
            filetypes=[("Archivos de Excel", "*.xlsx;*.xlsm"), ("Todos los archivos", "*.*")]
        )
        if f:
            self.path_origen.set(f)
            fname = os.path.basename(f)
            self.lbl_orig_status.configure(text=f"✓ {fname}", text_color="#4ade80")
            sheets = get_sheet_names(f)
            if sheets:
                self.cb_hoja_origen.configure(values=sheets)
                self.cb_hoja_origen.set(sheets[0])
                self.on_sheet_origen_changed(sheets[0])
            else:
                self.cb_hoja_origen.configure(values=["(No se leyeron hojas)"])

    def on_sheet_origen_changed(self, sheet_name):
        f = self.path_origen.get().strip()
        if f and os.path.isfile(f):
            cols = get_sheet_columns(f, sheet_name)
            if cols:
                self.entry_col_id.configure(values=cols)
                # Seleccionar automáticamente la columna de nombre o cédula si existe
                for c in cols:
                    if normalize_key(c) in ("NOMBREDELEMPLEADO", "NOMBRE", "EMPLEADO", "NOMBRES"):
                        self.col_nombre.set(c)
                        break

    def select_plantilla(self):
        f = filedialog.askopenfilename(
            title="Seleccionar Plantilla FV de Excel",
            filetypes=[("Archivos de Excel", "*.xlsx;*.xlsm"), ("Todos los archivos", "*.*")]
        )
        if f:
            self.path_plantilla.set(f)
            fname = os.path.basename(f)
            self.lbl_tmpl_status.configure(text=f"✓ {fname}", text_color="#4ade80")
            sheets = get_sheet_names(f)
            if sheets:
                self.cb_hoja_plantilla.configure(values=sheets)
                self.cb_hoja_plantilla.set(sheets[0])
            else:
                self.cb_hoja_plantilla.configure(values=["(No se leyeron hojas)"])

    def select_destino(self):
        d = filedialog.askdirectory(title="Seleccionar Carpeta para los Documentos Generados")
        if d:
            self.path_destino.set(d)
            self.lbl_dest_status.configure(text=f"✓ {d}", text_color="#4ade80")

    def log(self, message: str):
        """Registra un mensaje en la consola de la interfaz de forma thread-safe"""
        def _append():
            self.log_textbox.insert("end", f"{message}\n")
            self.log_textbox.see("end")
        self.safe_after(0, _append)

    # -------------------------------------------------------------------------
    # EJECUCIÓN EN SEGUNDO PLANO
    # -------------------------------------------------------------------------
    def start_thread(self):
        if self.is_running:
            return

        orig = self.path_origen.get().strip()
        tmpl = self.path_plantilla.get().strip()
        dest = self.path_destino.get().strip()

        if not orig or not os.path.isfile(orig):
            messagebox.showwarning("Atención", "Por favor seleccione el archivo Excel con los datos de origen.")
            return

        if not tmpl or not os.path.isfile(tmpl):
            messagebox.showwarning("Atención", "Por favor seleccione el archivo Excel de la plantilla oficial FV.")
            return

        if not dest or not os.path.isdir(dest):
            messagebox.showwarning("Atención", "Por favor seleccione una carpeta de destino válida.")
            return

        self.is_running = True
        self.btn_run.configure(state="disabled", text="⏳ PROCESANDO...", fg_color="#182c47", text_color="#94a3b8")
        self.progress_bar.set(0.0)
        self.lbl_progress.configure(text="Iniciando...", text_color="#38bdf8")

        threading.Thread(target=self._run_process, daemon=True).start()

    def _update_progress_ui(self, ratio, current, total):
        self.progress_bar.set(ratio)
        self.lbl_progress.configure(
            text=f"Procesando: {current} de {total} ({int(ratio * 100)}%)...",
            text_color="#38bdf8"
        )

    def _reset_ui_state(self):
        self.is_running = False
        self.btn_run.configure(
            state="normal",
            text="⚡ GENERAR FORMATOS Y PDF",
            fg_color="#00b4d8",
            text_color="#02293f"
        )

    def _run_process(self):
        exporter = None
        export_pdf = self.export_pdf_var.get()
        keep_xlsx = self.keep_xlsx_var.get()

        orig = self.path_origen.get().strip()
        h_orig = self.cb_hoja_origen.get().strip()
        tmpl = self.path_plantilla.get().strip()
        h_tmpl = self.cb_hoja_plantilla.get().strip()
        dest = self.path_destino.get().strip()
        col_id = self.col_nombre.get().strip() or "NOMBREDELEMPLEADO"

        try:
            self.log("=" * 60)
            self.log("🚀 Iniciando generación de Formatos de Vinculación (FV)...")

            self.log("⏳ Leyendo base de datos de origen...")
            records = list(iter_records_from_source(orig, h_orig))
            total = len(records)
            if total == 0:
                raise ValueError("No se encontraron registros de datos en la hoja seleccionada.")

            self.log(f"📋 Se encontraron {total} registro(s) listos para procesar.")

            if export_pdf:
                self.log("⚙️ Conectando con Microsoft Excel (COM) para exportación PDF...")
                exporter = ExcelPdfExporter()
                exporter.start()
                self.log("✓ Microsoft Excel inicializado correctamente.")

            errores = []
            exitos = 0

            for i, record in enumerate(records, start=1):
                norm_record = {normalize_key(k): v for k, v in record.items()}
                name_val = record.get(col_id)
                if name_val is None:
                    name_val = norm_record.get(normalize_key(col_id))

                base_name = safe_filename(name_val if name_val else f"formato_{i}")
                out_xlsx = os.path.join(dest, f"{base_name}.xlsx")
                out_pdf = os.path.join(dest, f"{base_name}.pdf")

                self.log(f"[{i}/{total}] Generando: {base_name}")
                ratio = i / total
                self.safe_after(0, lambda r=ratio, c=i, t=total: self._update_progress_ui(r, c, t))

                try:
                    fill_template_for_record(tmpl, h_tmpl, MAPPING_CONFIG, record, out_xlsx)

                    if export_pdf and exporter:
                        exporter.export(out_xlsx, out_pdf, sheet_name=h_tmpl)
                        if not keep_xlsx and os.path.exists(out_xlsx):
                            try:
                                os.remove(out_xlsx)
                            except Exception:
                                pass

                    exitos += 1
                except Exception as err:
                    err_msg = f"Fila {i} ({base_name}): {err}"
                    self.log(f"❌ Error en registro {i}: {err}")
                    errores.append(err_msg)

            self.safe_after(0, lambda: self._on_finished(exitos, total, errores, dest))

        except Exception as e:
            self.log(f"❌ Error crítico: {str(e)}")
            self.safe_after(0, lambda: messagebox.showerror("Error en el Proceso", str(e)))
            self.safe_after(0, self._reset_ui_state)
        finally:
            if exporter:
                exporter.close()

    def _on_finished(self, exitos, total, errores, dest_folder):
        self._reset_ui_state()
        self.progress_bar.set(1.0)
        self.lbl_progress.configure(
            text=f"✓ Finalizado: {exitos} de {total} documentos generados.",
            text_color="#4ade80"
        )
        self.log("=" * 60)
        self.log(f"🏁 Proceso completado: {exitos} generado(s) con éxito, {len(errores)} error(es).")

        if errores:
            msg = f"Se completó la generación de {exitos} de {total} documentos.\n\n"
            msg += f"Hubo {len(errores)} observación(es):\n" + "\n".join(errores[:8])
            if len(errores) > 8:
                msg += f"\n... y {len(errores) - 8} más."
            messagebox.showwarning("Finalizado con Observaciones", msg)
        else:
            ans = messagebox.askyesno(
                "¡Generación Exitosa!",
                f"Se han generado exitosamente los {total} documentos en:\n\n"
                f"{dest_folder}\n\n¿Desea abrir la carpeta contenedora ahora?"
            )
            if ans:
                try:
                    os.startfile(dest_folder)
                except Exception:
                    pass


# -----------------------------------------------------------------------------
# PUNTOS DE ENTRADA (INTEGRADO EN LANZADOR O STANDALONE)
# -----------------------------------------------------------------------------
def create_view(parent, on_back=None):
    """Punto de entrada estándar para incrustar en Lanzador.py"""
    return FormatoVinculacionView(parent, on_back=on_back)


def main():
    """Ejecución directa e independiente"""
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Automatización de Archivos - Formato de Vinculación (FV)")
    root.geometry("1040x720")
    root.minsize(900, 620)
    root.configure(fg_color="#0b1626")

    # Intentar asignar icono de la aplicación si existe
    for ico_cand in [
        resource_path("Logo_app.ico"),
        resource_path("pdf.ico"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "Logo_app.ico"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "pdf.ico"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logo_app.ico"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "pdf.ico"),
    ]:
        if os.path.exists(ico_cand):
            try:
                root.iconbitmap(ico_cand)
                break
            except Exception:
                pass

    view = FormatoVinculacionView(root, on_back=root.destroy)
    view.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()