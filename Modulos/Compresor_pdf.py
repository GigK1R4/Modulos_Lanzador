"""
=============================================================================
COMPRESOR DE DOCUMENTOS PDF (MODERNO - CUSTOMTKINTER)
=============================================================================
Permite optimizar y reducir el tamaño de archivos PDF utilizando Ghostscript
con diferentes niveles de compresión/calidad (Pantalla, Ebook, Impresión, etc.).
Diseñado con CustomTkinter para ejecutarse dentro de la ventana del Lanzador
o de forma independiente.
"""
import os
import sys
import glob
import shutil
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

# Opciones de compresión soportadas por Ghostscript
QUALITY_PRESETS = {
    "Ebook (150 dpi) - Balanceado [Recomendado]": "/ebook",
    "Pantalla (72 dpi) - Menor tamaño / máxima compresión": "/screen",
    "Impresora (300 dpi) - Alta calidad de impresión": "/printer",
    "Preprensa (300 dpi) - Calidad profesional": "/prepress",
    "Estándar (default) - Calidad general": "/default",
}


def resource_path(relative_path: str) -> str:
    """Ruta correcta para archivos incluidos con PyInstaller o ejecución directa"""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def default_gs_command() -> str:
    """
    Busca inteligentemente el ejecutable de Ghostscript en el sistema:
    1. En el PATH del sistema (gswin64c.exe, gswin32c.exe, gs.exe, gs).
    2. En rutas estándar de instalación en Windows (C:\\Program Files\\gs\\...\\bin).
    """
    if os.name == "nt":
        # 1. Probar en PATH
        for exe in ("gswin64c.exe", "gswin32c.exe", "gs.exe"):
            found = shutil.which(exe)
            if found:
                return found

        # 2. Buscar en Program Files
        search_patterns = [
            r"C:\Program Files\gs\gs*\bin\gswin64c.exe",
            r"C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe",
            r"C:\Program Files\gs\*\bin\*.exe",
        ]
        for pat in search_patterns:
            matches = glob.glob(pat)
            if matches:
                # Priorizar ejecutables de consola (*c.exe)
                console_exes = [m for m in matches if m.lower().endswith("c.exe")]
                return console_exes[0] if console_exes else matches[0]

        return "gswin64c.exe"  # Fallback
    else:
        for exe in ("gs", "ghostscript"):
            found = shutil.which(exe)
            if found:
                return found
        return "gs"


def format_file_size(size_bytes: int) -> str:
    """Formatea bytes a formato legible KB o MB"""
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.2f} MB"


def run_ghostscript(gs_exe: str, input_pdf: str, output_pdf: str, preset: str):
    """
    Ejecuta Ghostscript con parámetros de optimización para comprimir el archivo PDF.
    """
    cmd = [
        gs_exe,
        "-sDEVICE=pdfwrite",
        "-dSAFER",
        "-dBATCH",
        "-dNOPAUSE",
        "-dQUIET",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={preset}",
        "-dDetectDuplicateImages=true",
        "-dCompressFonts=true",
        "-dSubsetFonts=true",
        f"-sOutputFile={output_pdf}",
        input_pdf,
    ]

    creationflags = 0
    startupinfo = None
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    proc = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        creationflags=creationflags,
        startupinfo=startupinfo,
    )
    return proc.returncode, proc.stdout, proc.stderr, cmd


class PDFCompressorView(ctk.CTkFrame):
    """
    Vista modular para el Compresor de PDF integrada con CustomTkinter.
    """
    preferred_size = (1040, 700)
    min_size = (880, 580)

    def __init__(self, parent, on_back=None, *args, **kwargs):
        super().__init__(parent, fg_color="transparent", *args, **kwargs)
        self.on_back = on_back

        self.gs_path_var = tk.StringVar(value=default_gs_command())
        self.in_path_var = tk.StringVar(value="")
        self.out_path_var = tk.StringVar(value="")
        self.quality_key_var = tk.StringVar(value=list(QUALITY_PRESETS.keys())[0])

        self.is_compressing = False
        self._build_ui()

    def safe_after(self, ms, callback):
        try:
            if self.winfo_exists():
                self.after(ms, callback)
        except Exception:
            pass

    def _build_ui(self):
        # 1. Barra superior con botón volver
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
            text="COMPRESOR DE DOCUMENTOS PDF",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#ffffff"
        )
        lbl_title.pack(anchor="w", padx=20, pady=(12, 2))

        lbl_sub = ctk.CTkLabel(
            header_card,
            text="Reduzca el peso de sus archivos PDF sin perder legibilidad seleccionando el nivel de calidad deseado.",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#8ea3bf"
        )
        lbl_sub.pack(anchor="w", padx=20, pady=(0, 12))

        # 3. Contenedor central (2 Columnas: Archivos y Parámetros)
        center_frame = ctk.CTkFrame(self, fg_color="transparent")
        center_frame.pack(fill="both", expand=True, padx=25, pady=(0, 10))

        # --- Columna Izquierda: Selección de Archivos ---
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
            text="1. Archivos PDF",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#00b4d8"
        )
        lbl_sec1.pack(anchor="w", padx=18, pady=(12, 8))

        # Archivo Entrada
        btn_in = ctk.CTkButton(
            left_card,
            text="📄 Seleccionar Archivo PDF de Entrada",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#182c47",
            hover_color="#243f63",
            text_color="#38bdf8",
            height=34,
            corner_radius=8,
            command=self.select_input_pdf
        )
        btn_in.pack(fill="x", padx=18, pady=(2, 2))

        self.lbl_in_status = ctk.CTkLabel(
            left_card,
            text="Ningún PDF seleccionado",
            font=ctk.CTkFont(size=11),
            text_color="#64748b"
        )
        self.lbl_in_status.pack(anchor="w", padx=22, pady=(0, 10))

        # Archivo Salida
        btn_out = ctk.CTkButton(
            left_card,
            text="💾 Guardar PDF Comprimido Como...",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#10253f",
            hover_color="#18365c",
            text_color="#e2e8f0",
            height=34,
            corner_radius=8,
            command=self.select_output_pdf
        )
        btn_out.pack(fill="x", padx=18, pady=(2, 2))

        self.lbl_out_status = ctk.CTkLabel(
            left_card,
            text="Destino: (se sugerirá automáticamente)",
            font=ctk.CTkFont(size=11),
            text_color="#64748b"
        )
        self.lbl_out_status.pack(anchor="w", padx=22, pady=(0, 12))

        # Motor Ghostscript
        lbl_gs = ctk.CTkLabel(
            left_card,
            text="Motor Ghostscript detectado:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_gs.pack(anchor="w", padx=18, pady=(2, 2))

        gs_row = ctk.CTkFrame(left_card, fg_color="transparent")
        gs_row.pack(fill="x", padx=18, pady=(0, 10))

        self.entry_gs = ctk.CTkEntry(
            gs_row,
            textvariable=self.gs_path_var,
            font=ctk.CTkFont(family="Consolas", size=11),
            height=32,
            fg_color="#071426",
            border_color="#1d3b61",
            text_color="#e2e8f0"
        )
        self.entry_gs.pack(side="left", fill="x", expand=True, padx=(0, 6))

        btn_pick_gs = ctk.CTkButton(
            gs_row,
            text="Buscar...",
            width=80,
            height=32,
            fg_color="#182c47",
            hover_color="#243f63",
            text_color="#38bdf8",
            command=self.select_gs_exe
        )
        btn_pick_gs.pack(side="left")

        # --- Columna Derecha: Configuración y Ejecución ---
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
            text="2. Nivel de Compresión y Opciones",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#00b4d8"
        )
        lbl_sec2.pack(anchor="w", padx=18, pady=(12, 8))

        lbl_preset = ctk.CTkLabel(
            right_card,
            text="Nivel de Calidad / Compresión:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_preset.pack(anchor="w", padx=18, pady=(2, 2))

        self.cb_preset = ctk.CTkComboBox(
            right_card,
            values=list(QUALITY_PRESETS.keys()),
            variable=self.quality_key_var,
            height=34,
            corner_radius=8,
            fg_color="#071426",
            border_color="#1d3b61",
            text_color="#ffffff",
            button_color="#182c47",
            button_hover_color="#243f63"
        )
        self.cb_preset.pack(fill="x", padx=18, pady=(0, 15))

        # Información comparativa rápida
        self.info_box = ctk.CTkFrame(right_card, fg_color="#071426", corner_radius=8, border_width=1, border_color="#1d3b61")
        self.info_box.pack(fill="x", padx=18, pady=(0, 15))

        self.lbl_info_original = ctk.CTkLabel(
            self.info_box,
            text="Tamaño Original: --",
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        )
        self.lbl_info_original.pack(anchor="w", padx=14, pady=(8, 2))

        self.lbl_info_result = ctk.CTkLabel(
            self.info_box,
            text="Tamaño Comprimido: --",
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        )
        self.lbl_info_result.pack(anchor="w", padx=14, pady=(2, 8))

        # Botón de acción principal
        self.btn_compress = ctk.CTkButton(
            right_card,
            text="⚡ INICIAR COMPRESIÓN",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color="#00b4d8",
            hover_color="#0096c7",
            text_color="#02293f",
            height=38,
            corner_radius=8,
            command=self.start_compression_thread
        )
        self.btn_compress.pack(fill="x", padx=18, pady=(6, 8))

        # Botón de limpiar
        btn_clear = ctk.CTkButton(
            right_card,
            text="Limpiar Campos",
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            hover_color="#182c47",
            text_color="#94a3b8",
            height=28,
            command=self.clear_fields
        )
        btn_clear.pack(anchor="center", pady=(0, 4))

        # 4. Sección Inferior: Barra de Progreso y Consola de Registro
        bottom_frame = ctk.CTkFrame(
            self,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        bottom_frame.pack(fill="x", padx=25, pady=(0, 14))

        prog_header = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        prog_header.pack(fill="x", padx=18, pady=(10, 4))

        self.lbl_prog_title = ctk.CTkLabel(
            prog_header,
            text="Registro de Optimización",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_prog_title.pack(side="left")

        self.lbl_progress = ctk.CTkLabel(
            prog_header,
            text="Listo para comprimir",
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        )
        self.lbl_progress.pack(side="right")

        self.progress_bar = ctk.CTkProgressBar(
            bottom_frame,
            height=8,
            corner_radius=4,
            progress_color="#00b4d8",
            fg_color="#182c47"
        )
        self.progress_bar.pack(fill="x", padx=18, pady=(0, 8))
        self.progress_bar.set(0.0)

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

    def log(self, message: str):
        def _append():
            self.log_textbox.insert("end", f"{message}\n")
            self.log_textbox.see("end")
        self.safe_after(0, _append)

    def select_input_pdf(self):
        f = filedialog.askopenfilename(
            title="Seleccionar archivo PDF para comprimir",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
        )
        if f:
            self.in_path_var.set(f)
            fname = os.path.basename(f)
            size_bytes = os.path.getsize(f)
            size_str = format_file_size(size_bytes)
            self.lbl_in_status.configure(text=f"✓ {fname} ({size_str})", text_color="#4ade80")
            self.lbl_info_original.configure(text=f"Tamaño Original: {size_str}")

            # Sugerir salida automáticamente si no está establecida
            base, ext = os.path.splitext(f)
            suggested = f"{base}_comprimido.pdf"
            self.out_path_var.set(suggested)
            self.lbl_out_status.configure(text=f"✓ {os.path.basename(suggested)}", text_color="#38bdf8")

    def select_output_pdf(self):
        f = filedialog.asksaveasfilename(
            title="Guardar archivo PDF comprimido como",
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if f:
            self.out_path_var.set(f)
            self.lbl_out_status.configure(text=f"✓ {os.path.basename(f)}", text_color="#4ade80")

    def select_gs_exe(self):
        f = filedialog.askopenfilename(
            title="Seleccionar ejecutable de Ghostscript",
            filetypes=[("Ejecutables", "*.exe" if os.name == "nt" else "*"), ("Todos", "*.*")]
        )
        if f:
            self.gs_path_var.set(f)

    def clear_fields(self):
        self.in_path_var.set("")
        self.out_path_var.set("")
        self.lbl_in_status.configure(text="Ningún PDF seleccionado", text_color="#64748b")
        self.lbl_out_status.configure(text="Destino: (se sugerirá automáticamente)", text_color="#64748b")
        self.lbl_info_original.configure(text="Tamaño Original: --")
        self.lbl_info_result.configure(text="Tamaño Comprimido: --")
        self.progress_bar.set(0.0)
        self.lbl_progress.configure(text="Listo para comprimir", text_color="#94a3b8")
        self.log_textbox.delete("1.0", "end")

    def start_compression_thread(self):
        if self.is_compressing:
            return

        in_pdf = self.in_path_var.get().strip()
        out_pdf = self.out_path_var.get().strip()
        gs_exe = self.gs_path_var.get().strip()

        if not in_pdf or not os.path.isfile(in_pdf):
            messagebox.showwarning("Atención", "Por favor seleccione un archivo PDF de entrada válido.")
            return

        if not out_pdf:
            messagebox.showwarning("Atención", "Por favor defina la ruta del archivo PDF de salida.")
            return

        # Resolver ejecutable de Ghostscript
        resolved_gs = gs_exe
        if not os.path.isfile(resolved_gs):
            found = shutil.which(gs_exe)
            if found:
                resolved_gs = found
            else:
                resolved_gs = default_gs_command()

        if not os.path.isfile(resolved_gs) and not shutil.which(resolved_gs):
            messagebox.showerror(
                "Ghostscript no encontrado",
                "No se pudo encontrar el motor de compresión Ghostscript.\n\n"
                "Verifique que esté instalado en su equipo o use el botón 'Buscar...' "
                "para seleccionar gswin64c.exe."
            )
            return

        self.is_compressing = True
        self.btn_compress.configure(state="disabled", text="⏳ COMPRIMIENDO...", fg_color="#182c47", text_color="#94a3b8")
        self.progress_bar.set(0.4)
        self.lbl_progress.configure(text="Comprimiendo PDF...", text_color="#38bdf8")

        preset_key = self.quality_key_var.get()
        preset = QUALITY_PRESETS.get(preset_key, "/ebook")

        threading.Thread(
            target=self._run_compression,
            args=(resolved_gs, in_pdf, out_pdf, preset, preset_key),
            daemon=True
        ).start()

    def _reset_btn(self):
        self.is_compressing = False
        self.btn_compress.configure(
            state="normal",
            text="⚡ INICIAR COMPRESIÓN",
            fg_color="#00b4d8",
            text_color="#02293f"
        )

    def _run_compression(self, gs_exe, in_pdf, out_pdf, preset, preset_name):
        try:
            self.log("=" * 60)
            self.log(f"🚀 Iniciando compresión de: {os.path.basename(in_pdf)}")
            self.log(f"⚙️ Nivel seleccionado: {preset_name} ({preset})")
            self.log(f"🔧 Ghostscript: {gs_exe}")

            in_size = os.path.getsize(in_pdf)
            code, out, err, cmd = run_ghostscript(gs_exe, in_pdf, out_pdf, preset)

            if code != 0:
                self.log(f"❌ Error al ejecutar Ghostscript (código {code}):")
                self.log(err.strip() or "(sin mensaje de error)")
                self.safe_after(0, self._reset_btn)
                self.safe_after(0, lambda: self.lbl_progress.configure(text="Falló la compresión", text_color="#f87171"))
                self.safe_after(0, lambda: messagebox.showerror("Error", f"Ghostscript falló al comprimir:\n{err.strip()}"))
                return

            if not os.path.isfile(out_pdf):
                self.log("❌ No se encontró el archivo de salida tras la compresión.")
                self.safe_after(0, self._reset_btn)
                self.safe_after(0, lambda: self.lbl_progress.configure(text="Error al generar salida", text_color="#f87171"))
                return

            out_size = os.path.getsize(out_pdf)
            reduction = ((in_size - out_size) / in_size) * 100 if in_size > 0 else 0

            self.safe_after(0, lambda: self._on_success(in_size, out_size, reduction, out_pdf))

        except Exception as e:
            self.log(f"❌ Error inesperado: {str(e)}")
            self.safe_after(0, self._reset_btn)
            self.safe_after(0, lambda: messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{str(e)}"))

    def _on_success(self, in_size, out_size, reduction, out_pdf):
        self._reset_btn()
        self.progress_bar.set(1.0)

        in_str = format_file_size(in_size)
        out_str = format_file_size(out_size)
        reduc_str = f"{reduction:.1f}%" if reduction >= 0 else "0% (el archivo ya estaba optimizado)"

        self.lbl_info_result.configure(text=f"Tamaño Comprimido: {out_str} ({reduc_str} reducción)")
        self.lbl_progress.configure(
            text=f"✓ Optimizado exitosamente ({reduc_str} de reducción)",
            text_color="#4ade80"
        )

        self.log(f"✓ Archivo generado: {out_pdf}")
        self.log(f"📊 Original: {in_str} -> Comprimido: {out_str} (Reducción: {reduc_str})")
        self.log("=" * 60)

        ans = messagebox.askyesno(
            "¡Compresión Exitosa!",
            f"El PDF se comprimió correctamente:\n\n"
            f"• Original: {in_str}\n"
            f"• Comprimido: {out_str}\n"
            f"• Reducción lograda: {reduc_str}\n\n"
            f"Ubicación:\n{out_pdf}\n\n¿Desea abrir el archivo comprimido ahora?"
        )
        if ans:
            try:
                os.startfile(out_pdf)
            except Exception:
                pass


# -----------------------------------------------------------------------------
# PUNTOS DE ENTRADA
# -----------------------------------------------------------------------------
def check_for_pending_lanzador_update():
    """La actualización centralizada es gestionada de forma exclusiva por Lanzador.py"""
    pass


def create_view(parent, on_back=None):
    """Punto de entrada estándar para incrustar en el Lanzador"""
    check_for_pending_lanzador_update()
    return PDFCompressorView(parent, on_back=on_back)


def main():
    """Ejecución independiente del compresor"""
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Automatización de Archivos - Compresor de PDF")
    root.geometry("1040x700")
    root.minsize(880, 580)
    root.configure(fg_color="#0b1626")


    view = PDFCompressorView(root, on_back=root.destroy)
    view.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()
