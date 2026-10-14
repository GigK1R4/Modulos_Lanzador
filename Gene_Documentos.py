"""
=============================================================================
GENERACIÓN DE DOCUMENTOS MASIVOS (MODERNO - CUSTOMTKINTER)
=============================================================================
Genera múltiples documentos a partir de una plantilla Word (.docx) y una base
de datos Excel (.xlsx), con opción de conversión automática a PDF mediante Word.
Diseñado con CustomTkinter para ejecutarse dentro de la misma ventana del Lanzador.
"""
import os
import sys
import time
import threading
import pandas as pd
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

# --- COM / Word automation ---
try:
    import pythoncom
    import win32com.client  # pywin32
except Exception:
    pythoncom = None
    win32com = None


def resource_path(relative_path: str) -> str:
    """Ruta correcta para archivos incluidos con PyInstaller o ejecución directa"""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class WordPDFConverter:
    """Convierte DOCX -> PDF usando Microsoft Word (COM) de forma estable"""
    def __init__(self):
        self.word = None
        self.com_inited = False

    def start(self):
        if win32com is None or pythoncom is None:
            raise RuntimeError(
                "No se pudo cargar pywin32 (win32com/pythoncom).\n"
                "Instala: pip install pywin32"
            )

        # Inicializar COM en el hilo actual
        pythoncom.CoInitialize()
        self.com_inited = True

        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = 0
        except Exception as e:
            self.word = None
            raise RuntimeError(
                "No se pudo iniciar Microsoft Word para convertir a PDF.\n"
                "Verifica que Word esté instalado en este equipo.\n\n"
                f"Detalle: {repr(e)}"
            )

    def convert_docx_to_pdf(self, docx_path: str, pdf_path: str):
        if not self.word:
            raise RuntimeError("Word no está iniciado.")

        docx_path = os.path.abspath(docx_path)
        pdf_path = os.path.abspath(pdf_path)
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

        doc = None
        try:
            doc = self.word.Documents.Open(docx_path, ReadOnly=True)
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=17,  # wdExportFormatPDF
                OpenAfterExport=False,
                OptimizeFor=0,    # 0 = print
                CreateBookmarks=1
            )
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass

    def stop(self):
        if self.word is not None:
            try:
                self.word.Quit()
            except Exception:
                pass
            self.word = None

        if self.com_inited and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self.com_inited = False


class GeneDocumentosView(ctk.CTkFrame):
    """
    Vista de Generación Masiva con CustomTkinter idéntica a la estética del Lanzador.
    """
    preferred_size = (1060, 720)
    min_size = (920, 620)

    def __init__(self, parent, on_back=None, *args, **kwargs):
        super().__init__(parent, fg_color="transparent", *args, **kwargs)
        self.on_back = on_back

        # Rutas y estado
        self.template_path = ""
        self.excel_path = ""
        self.output_dir = ""
        self.df = None
        self.check_vars = {}
        self.is_generating = False

        self.generate_pdf_var = tk.BooleanVar(value=True)
        self.prefix_var = tk.StringVar(value="Contrato")

        self._build_ui()

    def _build_ui(self):
        # 1. Barra superior: botón Volver
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(fill="x", padx=25, pady=(15, 5))

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

        # Tarjeta de Cabecera
        header_card = ctk.CTkFrame(
            self,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        header_card.pack(fill="x", padx=25, pady=(5, 12))

        lbl_title = ctk.CTkLabel(
            header_card,
            text="GENERACIÓN DE DOCUMENTOS MASIVOS",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#ffffff"
        )
        lbl_title.pack(anchor="w", padx=20, pady=(12, 2))

        lbl_sub = ctk.CTkLabel(
            header_card,
            text="Combine una plantilla Word (.docx) con una base de datos Excel (.xlsx) y convierta a PDF automáticamente.",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#8ea3bf"
        )
        lbl_sub.pack(anchor="w", padx=20, pady=(0, 12))

        # 2. Contenedor central (Dos columnas: Configuración y Campos)
        center_frame = ctk.CTkFrame(self, fg_color="transparent")
        center_frame.pack(fill="both", expand=True, padx=25, pady=(0, 10))

        # --- Columna Izquierda: Configuración de Archivos ---
        left_card = ctk.CTkFrame(
            center_frame,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        left_card.pack(side="left", fill="both", expand=True, padx=(0, 10))

        lbl_sec1 = ctk.CTkLabel(
            left_card,
            text="1. Configuración de Archivos",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#00b4d8"
        )
        lbl_sec1.pack(anchor="w", padx=18, pady=(14, 10))

        # Plantilla Word
        btn_tmpl = ctk.CTkButton(
            left_card,
            text="📄 Seleccionar Plantilla Word (.docx)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#182c47",
            hover_color="#243f63",
            text_color="#38bdf8",
            height=34,
            corner_radius=8,
            command=self.load_template
        )
        btn_tmpl.pack(fill="x", padx=18, pady=(2, 2))

        self.lbl_tmpl_status = ctk.CTkLabel(
            left_card,
            text="Ninguna plantilla seleccionada",
            font=ctk.CTkFont(size=11),
            text_color="#64748b"
        )
        self.lbl_tmpl_status.pack(anchor="w", padx=22, pady=(0, 8))

        # Archivo Excel
        btn_xl = ctk.CTkButton(
            left_card,
            text="📊 Seleccionar Archivo Excel (.xlsx)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#182c47",
            hover_color="#243f63",
            text_color="#38bdf8",
            height=34,
            corner_radius=8,
            command=self.load_excel
        )
        btn_xl.pack(fill="x", padx=18, pady=(4, 2))

        self.lbl_xl_status = ctk.CTkLabel(
            left_card,
            text="Ningún archivo Excel seleccionado",
            font=ctk.CTkFont(size=11),
            text_color="#64748b"
        )
        self.lbl_xl_status.pack(anchor="w", padx=22, pady=(0, 8))

        # Pestaña de Excel
        lbl_sheet = ctk.CTkLabel(
            left_card,
            text="Hoja del Excel a procesar:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_sheet.pack(anchor="w", padx=18, pady=(4, 2))

        self.sheet_combo = ctk.CTkComboBox(
            left_card,
            values=["(Cargue un Excel primero)"],
            height=34,
            corner_radius=8,
            fg_color="#071426",
            border_color="#1d3b61",
            text_color="#ffffff",
            button_color="#182c47",
            button_hover_color="#243f63",
            command=self.on_sheet_selected
        )
        self.sheet_combo.pack(fill="x", padx=18, pady=(0, 10))

        # Carpeta de salida
        btn_out = ctk.CTkButton(
            left_card,
            text="📁 Seleccionar Carpeta de Destino",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#10253f",
            hover_color="#18365c",
            text_color="#e2e8f0",
            height=34,
            corner_radius=8,
            command=self.choose_output_folder
        )
        btn_out.pack(fill="x", padx=18, pady=(4, 2))

        self.lbl_out_status = ctk.CTkLabel(
            left_card,
            text="Carpeta: (no seleccionada)",
            font=ctk.CTkFont(size=11),
            text_color="#64748b"
        )
        self.lbl_out_status.pack(anchor="w", padx=22, pady=(0, 8))

        # Prefijo
        lbl_pref = ctk.CTkLabel(
            left_card,
            text="Prefijo para los nombres de archivo:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_pref.pack(anchor="w", padx=18, pady=(4, 2))

        self.prefix_entry = ctk.CTkEntry(
            left_card,
            textvariable=self.prefix_var,
            placeholder_text="Ej: Contrato, Certificado",
            height=34,
            corner_radius=8,
            fg_color="#071426",
            border_color="#1d3b61",
            text_color="#ffffff"
        )
        self.prefix_entry.pack(fill="x", padx=18, pady=(0, 10))

        # Opción PDF
        self.chk_pdf = ctk.CTkCheckBox(
            left_card,
            text="Generar también en formato PDF (requiere Word)",
            variable=self.generate_pdf_var,
            font=ctk.CTkFont(size=12),
            text_color="#e2e8f0",
            fg_color="#00b4d8",
            hover_color="#0096c7",
            border_color="#244770"
        )
        self.chk_pdf.pack(anchor="w", padx=18, pady=(4, 14))

        # --- Columna Derecha: Selección de Campos / Columnas ---
        right_card = ctk.CTkFrame(
            center_frame,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        right_card.pack(side="right", fill="both", expand=True, padx=(10, 0))

        lbl_sec2 = ctk.CTkLabel(
            right_card,
            text="2. Campos y Variables de la Hoja",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#00b4d8"
        )
        lbl_sec2.pack(anchor="w", padx=18, pady=(14, 2))

        lbl_sec2_sub = ctk.CTkLabel(
            right_card,
            text="Marque las columnas que coinciden con los marcadores {{ campo }} de su plantilla:",
            font=ctk.CTkFont(size=11),
            text_color="#8ea3bf",
            wraplength=340
        )
        lbl_sec2_sub.pack(anchor="w", padx=18, pady=(0, 6))

        # Botones rápidos de selección
        btn_box = ctk.CTkFrame(right_card, fg_color="transparent")
        btn_box.pack(fill="x", padx=18, pady=(0, 6))

        btn_all = ctk.CTkButton(
            btn_box,
            text="Marcar Todos",
            font=ctk.CTkFont(size=11),
            width=100,
            height=26,
            fg_color="#182c47",
            hover_color="#243f63",
            command=self.select_all_fields
        )
        btn_all.pack(side="left", padx=(0, 8))

        btn_none = ctk.CTkButton(
            btn_box,
            text="Desmarcar Todos",
            font=ctk.CTkFont(size=11),
            width=110,
            height=26,
            fg_color="#182c47",
            hover_color="#243f63",
            command=self.deselect_all_fields
        )
        btn_none.pack(side="left")

        # Scrollable Frame con los checkboxes
        self.scroll_fields = ctk.CTkScrollableFrame(
            right_card,
            fg_color="#071426",
            border_width=1,
            border_color="#1d3b61",
            corner_radius=10
        )
        self.scroll_fields.pack(fill="both", expand=True, padx=18, pady=(4, 14))

        self.lbl_no_fields = ctk.CTkLabel(
            self.scroll_fields,
            text="(Cargue un archivo Excel y elija una hoja para ver las columnas)",
            font=ctk.CTkFont(size=12),
            text_color="#64748b"
        )
        self.lbl_no_fields.pack(pady=30)

        # 3. Pie de página: Barra de progreso y botón de inicio
        footer_card = ctk.CTkFrame(
            self,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        footer_card.pack(fill="x", padx=25, pady=(0, 15))

        prog_box = ctk.CTkFrame(footer_card, fg_color="transparent")
        prog_box.pack(side="left", fill="x", expand=True, padx=18, pady=12)

        self.lbl_progress = ctk.CTkLabel(
            prog_box,
            text="Listo para iniciar.",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#8ea3bf"
        )
        self.lbl_progress.pack(anchor="w", pady=(0, 4))

        self.progress_bar = ctk.CTkProgressBar(
            prog_box,
            height=12,
            corner_radius=6,
            fg_color="#071426",
            progress_color="#00b4d8"
        )
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

        self.btn_generate = ctk.CTkButton(
            footer_card,
            text="⚡ INICIAR GENERACIÓN",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#00b4d8",
            hover_color="#0096c7",
            text_color="#02293f",
            width=230,
            height=44,
            corner_radius=10,
            command=self.start_generation
        )
        self.btn_generate.pack(side="right", padx=18, pady=12)

    # =========================================================================
    # LÓGICA DE CARGA DE ARCHIVOS
    # =========================================================================
    def load_template(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Plantilla Word",
            filetypes=[("Plantillas Word (*.docx)", "*.docx"), ("Todos los archivos", "*.*")]
        )
        if path:
            self.template_path = os.path.abspath(path)
            self.lbl_tmpl_status.configure(
                text=f"✓ {os.path.basename(self.template_path)}",
                text_color="#4ade80"
            )

    def load_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Archivo Excel",
            filetypes=[("Archivos Excel (*.xlsx;*.xls)", "*.xlsx;*.xls"), ("Todos los archivos", "*.*")]
        )
        if not path:
            return

        try:
            self.excel_path = os.path.abspath(path)
            xls = pd.ExcelFile(self.excel_path)
            sheets = xls.sheet_names
            if not sheets:
                messagebox.showerror("Error", "El archivo Excel no contiene hojas de cálculo.")
                return

            self.sheet_combo.configure(values=sheets)
            self.sheet_combo.set(sheets[0])
            self.lbl_xl_status.configure(
                text=f"✓ {os.path.basename(self.excel_path)} ({len(sheets)} hojas)",
                text_color="#4ade80"
            )
            self.on_sheet_selected(sheets[0])
        except Exception as e:
            messagebox.showerror("Error al leer Excel", f"No se pudo abrir el archivo:\n{e}")

    def on_sheet_selected(self, sheet_name=None):
        if not sheet_name:
            sheet_name = self.sheet_combo.get()

        if not self.excel_path or not sheet_name:
            return

        try:
            self.df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar la hoja '{sheet_name}':\n{e}")
            return

        # Limpiar contenedor de campos
        for w in self.scroll_fields.winfo_children():
            w.destroy()
        self.check_vars.clear()

        if self.df is None or len(self.df.columns) == 0:
            lbl = ctk.CTkLabel(self.scroll_fields, text="(La hoja seleccionada no tiene columnas)", text_color="#64748b")
            lbl.pack(pady=20)
            return

        # Generar checkbox estilizado por cada columna
        for col in self.df.columns:
            var = tk.BooleanVar(value=True)
            chk = ctk.CTkCheckBox(
                self.scroll_fields,
                text=str(col),
                variable=var,
                font=ctk.CTkFont(size=12),
                text_color="#ffffff",
                fg_color="#00b4d8",
                hover_color="#0096c7",
                border_color="#244770"
            )
            chk.pack(anchor="w", padx=10, pady=4)
            self.check_vars[col] = var

        self.lbl_progress.configure(
            text=f"Hoja '{sheet_name}' cargada: {len(self.df)} filas y {len(self.df.columns)} columnas.",
            text_color="#38bdf8"
        )

    def select_all_fields(self):
        for var in self.check_vars.values():
            var.set(True)

    def deselect_all_fields(self):
        for var in self.check_vars.values():
            var.set(False)

    def choose_output_folder(self):
        folder = filedialog.askdirectory(title="Seleccionar Carpeta de Destino")
        if folder:
            self.output_dir = os.path.abspath(folder)
            self.lbl_out_status.configure(
                text=f"✓ Destino: {self.output_dir}",
                text_color="#4ade80"
            )

    def _safe_filename_part(self, s: str) -> str:
        s = str(s).strip()
        for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
            s = s.replace(ch, "")
        return " ".join(s.split())

    # =========================================================================
    # GENERACIÓN EN SEGUNDO PLANO (THREAD-SAFE)
    # =========================================================================
    def start_generation(self):
        """Inicia el proceso de generación sin congelar la interfaz gráfica"""
        if self.is_generating:
            return

        if not self.template_path or not os.path.exists(self.template_path):
            messagebox.showerror("Plantilla requerida", "Por favor seleccione una plantilla Word (.docx).")
            return

        if self.df is None or len(self.df) == 0:
            messagebox.showerror("Datos requeridos", "Por favor cargue un archivo Excel con al menos una fila.")
            return

        if not self.output_dir or not os.path.isdir(self.output_dir):
            messagebox.showerror("Destino requerido", "Por favor seleccione la carpeta de salida.")
            return

        selected_fields = [col for col, var in self.check_vars.items() if var.get()]
        if not selected_fields:
            messagebox.showerror("Campos requeridos", "Debe marcar al menos un campo de la lista.")
            return

        prefix = self._safe_filename_part(self.prefix_var.get())
        if not prefix:
            messagebox.showerror("Prefijo requerido", "Escriba un prefijo para el nombre de los archivos.")
            return

        # Preparar UI
        self.is_generating = True
        self.btn_generate.configure(state="disabled", text="PROCESANDO...", fg_color="#182c47")
        self.progress_bar.set(0)

        # Iniciar en hilo secundario
        thread = threading.Thread(
            target=self._generation_worker,
            args=(selected_fields, prefix),
            daemon=True
        )
        thread.start()

    def _generation_worker(self, selected_fields, prefix):
        total = len(self.df)
        generate_pdf = self.generate_pdf_var.get()
        converter = None

        if generate_pdf:
            converter = WordPDFConverter()
            try:
                converter.start()
            except Exception as e:
                converter = None
                generate_pdf = False
                self.after(0, lambda: messagebox.showwarning(
                    "Conversión a PDF no disponible",
                    f"No se pudo inicializar Word:\n{e}\n\nSe generarán solo los archivos Word."
                ))

        errores = []

        try:
            for i in range(total):
                try:
                    doc = DocxTemplate(self.template_path)
                    context = {col: self.df.loc[i, col] for col in selected_fields}
                    doc.render(context)

                    parte_variable = self._safe_filename_part(self.df.loc[i, selected_fields[0]])
                    base_name = f"{prefix} {parte_variable}".strip()

                    docx_path = os.path.join(self.output_dir, f"{base_name}.docx")
                    pdf_path = os.path.join(self.output_dir, f"{base_name}.pdf")

                    doc.save(docx_path)

                    if generate_pdf and converter is not None:
                        time.sleep(0.04)
                        converter.convert_docx_to_pdf(docx_path, pdf_path)

                except Exception as e_row:
                    errores.append(f"Fila {i+1}: {str(e_row)}")

                # Actualizar progreso en la UI
                current_step = i + 1
                ratio = current_step / total
                self.after(0, lambda r=ratio, s=current_step: self._update_progress_ui(r, s, total))

        finally:
            if converter is not None:
                converter.stop()

        self.after(0, lambda: self._on_generation_finished(errores, total))

    def _update_progress_ui(self, ratio, current, total):
        self.progress_bar.set(ratio)
        self.lbl_progress.configure(
            text=f"Procesando: {current} de {total} documentos ({int(ratio * 100)}%)...",
            text_color="#38bdf8"
        )

    def _on_generation_finished(self, errores, total):
        self.is_generating = False
        self.btn_generate.configure(state="normal", text="⚡ INICIAR GENERACIÓN", fg_color="#00b4d8")

        if errores:
            msg = f"Se completó la generación de {total - len(errores)} documentos con {len(errores)} error(es):\n\n"
            msg += "\n".join(errores[:10])
            if len(errores) > 10:
                msg += f"\n... y {len(errores) - 10} más."
            messagebox.showwarning("Finalizado con observaciones", msg)
        else:
            ans = messagebox.askyesno(
                "¡Generación Exitosa!",
                f"Se han generado exitosamente los {total} documentos en:\n\n"
                f"{self.output_dir}\n\n¿Desea abrir la carpeta contenedora ahora?"
            )
            if ans:
                try:
                    os.startfile(self.output_dir)
                except Exception:
                    pass

        self.lbl_progress.configure(
            text=f"✓ Proceso finalizado ({total} registros procesados).",
            text_color="#4ade80"
        )


def create_view(parent, on_back=None):
    """Punto de entrada para incrustar en la misma ventana de Lanzador.py"""
    return GeneDocumentosView(parent, on_back=on_back)


def main():
    """Ejecución independiente de prueba"""
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Automatización de Archivos - Generador de Documentos")
    root.geometry("1020x700")
    root.minsize(920, 600)
    root.configure(fg_color="#0b1626")

    view = GeneDocumentosView(root, on_back=root.destroy)
    view.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()