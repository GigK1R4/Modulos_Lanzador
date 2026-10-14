import os
import sys
import time
import pandas as pd
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- COM / Word automation ---
try:
    import pythoncom
    import win32com.client  # pywin32
except Exception:
    pythoncom = None
    win32com = None


def resource_path(relative_path: str) -> str:
    """
    Ruta correcta para archivos incluidos con PyInstaller.
    - En modo .py: carpeta actual
    - En modo .exe onefile: carpeta temporal _MEIPASS
    """
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class WordPDFConverter:
    """
    Convierte DOCX -> PDF usando Microsoft Word (COM).
    Reusa 1 instancia de Word para convertir muchos archivos (más estable).
    """
    def __init__(self):
        self.word = None
        self.com_inited = False

    def start(self):
        if win32com is None or pythoncom is None:
            raise RuntimeError(
                "No se pudo cargar pywin32 (win32com/pythoncom).\n"
                "Instala: pip install pywin32"
            )

        # Inicializar COM (importante en .exe)
        pythoncom.CoInitialize()
        self.com_inited = True

        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = 0  # sin alerts
        except Exception as e:
            self.word = None
            raise RuntimeError(
                "No se pudo iniciar Microsoft Word para convertir a PDF.\n"
                "Verifica que Word esté instalado y que puedas abrirlo al menos una vez.\n\n"
                f"Detalle: {repr(e)}"
            )

    def convert_docx_to_pdf(self, docx_path: str, pdf_path: str):
        if not self.word:
            raise RuntimeError("Word no está iniciado.")

        docx_path = os.path.abspath(docx_path)
        pdf_path = os.path.abspath(pdf_path)

        # Asegurar carpeta destino
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

        doc = None
        try:
            # Abrir docx (ReadOnly para estabilidad)
            doc = self.word.Documents.Open(docx_path, ReadOnly=True)

            # ExportAsFixedFormat: 17 = wdExportFormatPDF
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,   # 0 = print
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


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, height=280, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.v_scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set)

        self.inner = ttk.Frame(self.canvas)

        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.v_scroll.pack(side="right", fill="y")

        self.configure(height=height)
        self.pack_propagate(False)
        self.canvas.configure(height=height)

        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.inner.bind("<Enter>", self._bind_mousewheel)
        self.inner.bind("<Leave>", self._unbind_mousewheel)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _bind_mousewheel(self, _event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        delta = int(-1 * (event.delta / 120))
        self.canvas.yview_scroll(delta, "units")


class CertificadosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Documentos")

        self._apply_system_theme()
        self._apply_icon("administrar.ico")

        self.root.geometry("700x770")
        self.root.minsize(680, 600)

        self.template_path = ""
        self.excel_path = ""
        self.output_dir = ""
        self.df = None
        self.check_vars = {}

        self.generate_pdf_var = tk.BooleanVar(value=True)
        self.prefix_var = tk.StringVar(value="Contrato")

        self.build_ui()

    def _apply_system_theme(self):
        style = ttk.Style()
        themes = set(style.theme_names())
        for preferred in ("vista", "xpnative", "aqua"):
            if preferred in themes:
                try:
                    style.theme_use(preferred)
                    return
                except Exception:
                    pass

    def _apply_icon(self, ico_filename: str):
        icon_path = resource_path(ico_filename)
        try:
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass

    def build_ui(self):
        frame = ttk.Frame(self.root, padding=15)
        frame.pack(fill="both", expand=True)

        ttk.Button(frame, text="Seleccionar plantilla Word", command=self.load_template).pack(fill="x", pady=5)
        ttk.Button(frame, text="Seleccionar archivo Excel", command=self.load_excel).pack(fill="x", pady=5)

        ttk.Label(frame, text="Seleccionar hoja del Excel:").pack(pady=(10, 2))

        self.sheet_combo = ttk.Combobox(frame, state="readonly")
        self.sheet_combo.pack(fill="x", pady=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.load_sheet)

        ttk.Button(frame, text="Seleccionar carpeta de salida", command=self.choose_output_folder).pack(fill="x", pady=5)

        self.output_label = ttk.Label(frame, text="Carpeta de salida: (no seleccionada)")
        self.output_label.pack(fill="x", pady=(0, 10))

        ttk.Label(frame, text="Prefijo del nombre del archivo (primera parte):").pack(pady=(10, 2))
        self.prefix_entry = ttk.Entry(frame, textvariable=self.prefix_var)
        self.prefix_entry.pack(fill="x", pady=5)

        ttk.Checkbutton(
            frame,
            text="Generar PDF además del Word (requiere Microsoft Word instalado)",
            variable=self.generate_pdf_var
        ).pack(anchor="w", pady=(5, 10))

        ttk.Label(frame, text="Campos a usar en la plantilla:").pack(pady=(10, 6))

        self.scroll_fields = ScrollableFrame(frame, height=280)
        self.scroll_fields.pack(fill="both", expand=True)

        ttk.Button(frame, text="Generar documentos", command=self.generate).pack(fill="x", pady=12)

        self.progress = ttk.Progressbar(frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=5)

        self.progress_label = ttk.Label(frame, text="")
        self.progress_label.pack(fill="x", pady=(0, 5))

        footer = ttk.Label(self.root, text="Realizado por DK Software", anchor="center", font=("Arial", 9))
        footer.pack(side="bottom", pady=5)

    def load_template(self):
        self.template_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if self.template_path:
            messagebox.showinfo("OK", "Plantilla cargada correctamente")

    def load_excel(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not self.excel_path:
            return
        try:
            xls = pd.ExcelFile(self.excel_path)
            self.sheet_combo["values"] = xls.sheet_names
            self.sheet_combo.set("")
            messagebox.showinfo("Excel cargado", "Seleccione la hoja que desea usar")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def choose_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir = folder
            self.output_label.config(text=f"Carpeta de salida: {self.output_dir}")

    def load_sheet(self, event=None):
        sheet_name = self.sheet_combo.get()
        if not sheet_name:
            return

        try:
            self.df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer la hoja:\n{e}")
            return

        for widget in self.scroll_fields.inner.winfo_children():
            widget.destroy()

        self.check_vars.clear()

        if self.df is None or len(self.df.columns) == 0:
            ttk.Label(self.scroll_fields.inner, text="(Esta hoja no tiene columnas)").pack(anchor="w")
            self.root.update_idletasks()
            return

        for col in self.df.columns:
            var = tk.BooleanVar(value=True)
            chk = ttk.Checkbutton(self.scroll_fields.inner, text=str(col), variable=var)
            chk.pack(anchor="w", padx=5, pady=2)
            self.check_vars[col] = var

        self.root.update_idletasks()

    def _safe_filename_part(self, s: str) -> str:
        s = str(s).strip()
        for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
            s = s.replace(ch, "")
        s = " ".join(s.split())
        return s

    def generate(self):
        if not self.template_path or self.df is None:
            messagebox.showerror("Error", "Debe cargar plantilla, Excel y hoja")
            return
        if not self.output_dir:
            messagebox.showerror("Error", "Debe seleccionar una carpeta de salida")
            return

        selected_fields = [col for col, var in self.check_vars.items() if var.get()]
        if not selected_fields:
            messagebox.showerror("Error", "Seleccione al menos un campo")
            return

        prefix = self._safe_filename_part(self.prefix_var.get())
        if not prefix:
            messagebox.showerror("Error", "Escriba un prefijo para el nombre del archivo")
            return

        total = len(self.df)
        if total == 0:
            messagebox.showinfo("Sin datos", "La hoja seleccionada no tiene filas")
            return

        generate_pdf = self.generate_pdf_var.get()

        # Inicializar convertidor Word (una sola vez)
        converter = None
        if generate_pdf:
            converter = WordPDFConverter()
            try:
                converter.start()
            except Exception as e:
                # Si Word falla, seguimos solo con DOCX
                messagebox.showwarning(
                    "PDF desactivado",
                    f"No se pudo iniciar Word para PDF.\n\n{e}\n\n"
                    "Se generarán solo los Word."
                )
                converter = None
                generate_pdf = False

        self.progress["value"] = 0
        self.progress["maximum"] = total
        self.progress_label.config(text=f"Procesando: 0 / {total}")
        self.root.update_idletasks()

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
                        # pequeña pausa para evitar carrera con Word en algunas PCs
                        time.sleep(0.05)
                        converter.convert_docx_to_pdf(docx_path, pdf_path)

                except Exception as e_row:
                    errores.append(f"Fila {i+1} -> {repr(e_row)}")

                self.progress["value"] = i + 1
                self.progress_label.config(text=f"Procesando: {i+1} / {total}")
                self.root.update_idletasks()

        finally:
            if converter is not None:
                converter.stop()

        if errores:
            messagebox.showwarning(
                "Finalizado con errores",
                "Se generaron documentos, pero hubo errores:\n\n" +
                "\n".join(errores[:12]) +
                ("\n\n... (hay más)" if len(errores) > 12 else "")
            )
        else:
            messagebox.showinfo("Finalizado", "Documentos Word y PDF generados correctamente")


if __name__ == "__main__":
    root = tk.Tk()
    app = CertificadosApp(root)
    root.mainloop()