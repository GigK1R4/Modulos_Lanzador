"""
=============================================================================
CONVERTIDOR Y UNIFICADOR DE WORD A PDF (TAMAÑO CARTA) - CUSTOMTKINTER
=============================================================================
Permite seleccionar múltiples documentos de Word (.docx, .doc),
convertirlos a PDF, normalizar cada página a tamaño Carta estándar (8.5 x 11 in)
manteniendo la proporción y centrando el contenido, y unificarlos en un solo archivo PDF.
Compatible con el Lanzador Modular o ejecutable de forma independiente.
=============================================================================
"""
import os
import sys
import tempfile
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from pypdf import PdfReader, PdfWriter, Transformation, PageObject


# Función de compatibilidad para actualización del lanzador
def check_for_pending_lanzador_update():
    """La actualización centralizada es gestionada de forma exclusiva por Lanzador.py"""
    pass


def convert_single_word_to_pdf(doc_path: str, pdf_path: str):
    """
    Convierte un documento Word a PDF utilizando docx2pdf o fallback directo a Word COM.
    """
    # 1. Intentar con docx2pdf
    try:
        from docx2pdf import convert
        convert(doc_path, pdf_path)
        if os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 0:
            return True
    except Exception:
        pass

    # 2. Fallback COM directo con Word (Windows nativo)
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(doc_path))
        # 17 = wdFormatPDF
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close(SaveChanges=False)
        word.Quit()
        pythoncom.CoUninitialize()
        return os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 0
    except Exception as e:
        raise RuntimeError(f"Error al convertir {os.path.basename(doc_path)}: {e}")


class WordToPdfMergerView(ctk.CTkFrame):
    """Vista moderna con CustomTkinter para convertir y unificar Word a PDF tamaño Carta"""
    def __init__(self, master, on_back=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.on_back = on_back

        self.preferred_size = (1040, 700)
        self.min_size = (880, 580)

        # Dimensiones Carta (Letter) en puntos PostScript (8.5 x 11 in)
        self.LETTER_WIDTH = 612.0
        self.LETTER_HEIGHT = 792.0

        self.files = []
        self.is_processing = False

        self._build_ui()

    def _build_ui(self):
        # 1. Barra de Navegación / Encabezado
        nav_frame = ctk.CTkFrame(self, fg_color="transparent")
        nav_frame.pack(fill="x", padx=30, pady=(18, 12))

        if self.on_back:
            btn_back = ctk.CTkButton(
                nav_frame,
                text="← Volver al Menú",
                width=140,
                height=32,
                corner_radius=8,
                fg_color="#182c47",
                hover_color="#243f66",
                text_color="#94a3b8",
                font=ctk.CTkFont(size=12, weight="bold"),
                command=self.on_back
            )
            btn_back.pack(side="left")

        title_lbl = ctk.CTkLabel(
            nav_frame,
            text="CONVERTIDOR Y UNIFICADOR A PDF (CARTA)",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#ffffff"
        )
        title_lbl.pack(side="left", padx=20)

        # 2. Contenedor Principal (Panel Dividido)
        content_frame = ctk.CTkFrame(self, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        # --- Panel Izquierdo: Lista de Archivos Word ---
        left_card = ctk.CTkFrame(
            content_frame,
            fg_color="#0e1f38",
            corner_radius=18,
            border_width=1,
            border_color="#1a3557"
        )
        left_card.pack(side="left", fill="both", expand=True, padx=(0, 15))

        left_header = ctk.CTkFrame(left_card, fg_color="transparent")
        left_header.pack(fill="x", padx=20, pady=(16, 10))

        lbl_docs = ctk.CTkLabel(
            left_header,
            text="📄 Documentos de Word a Procesar",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#ffffff"
        )
        lbl_docs.pack(side="left")

        # Botones de agregar y limpiar
        btn_box = ctk.CTkFrame(left_header, fg_color="transparent")
        btn_box.pack(side="right")

        btn_add = ctk.CTkButton(
            btn_box,
            text="➕ Agregar Word",
            width=130,
            height=32,
            corner_radius=8,
            fg_color="#00b4d8",
            hover_color="#0096c7",
            text_color="#031629",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.add_files
        )
        btn_add.pack(side="left", padx=6)

        btn_clear = ctk.CTkButton(
            btn_box,
            text="🗑️ Limpiar",
            width=90,
            height=32,
            corner_radius=8,
            fg_color="#213247",
            hover_color="#304763",
            text_color="#f87171",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.clear_list
        )
        btn_clear.pack(side="left")

        # Área de Lista y Botones de Reordenamiento
        list_container = ctk.CTkFrame(left_card, fg_color="transparent")
        list_container.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        # Listbox estilo nativo integrado
        list_sub = tk.Frame(list_container, bg="#071324", bd=0)
        list_sub.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_sub, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(
            list_sub,
            selectmode=tk.SINGLE,
            yscrollcommand=scrollbar.set,
            font=("Segoe UI", 11),
            bg="#071324",
            fg="#f1f5f9",
            selectbackground="#00b4d8",
            selectforeground="#031629",
            bd=0,
            highlightthickness=0,
            activestyle="none"
        )
        scrollbar.config(command=self.listbox.yview)
        self.listbox.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        scrollbar.pack(side="right", fill="y", pady=8)

        # Botones de ordenamiento (Subir, Bajar, Eliminar)
        order_bar = ctk.CTkFrame(list_container, fg_color="transparent", width=48)
        order_bar.pack(side="right", fill="y", padx=(10, 0))

        btn_up = ctk.CTkButton(
            order_bar,
            text="▲",
            width=42,
            height=42,
            corner_radius=10,
            fg_color="#182c47",
            hover_color="#243f66",
            text_color="#38bdf8",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.move_up
        )
        btn_up.pack(pady=(0, 8))

        btn_down = ctk.CTkButton(
            order_bar,
            text="▼",
            width=42,
            height=42,
            corner_radius=10,
            fg_color="#182c47",
            hover_color="#243f66",
            text_color="#38bdf8",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.move_down
        )
        btn_down.pack(pady=(0, 16))

        btn_del = ctk.CTkButton(
            order_bar,
            text="✕",
            width=42,
            height=42,
            corner_radius=10,
            fg_color="#2f1a24",
            hover_color="#451e2f",
            text_color="#f87171",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.remove_selected
        )
        btn_del.pack(pady=(0, 0))

        # --- Panel Derecho: Opciones de Formato y Acción ---
        right_card = ctk.CTkFrame(
            content_frame,
            fg_color="#0e1f38",
            corner_radius=18,
            border_width=1,
            border_color="#1a3557",
            width=320
        )
        right_card.pack(side="right", fill="y", padx=(0, 0))
        right_card.pack_propagate(False)

        lbl_opts = ctk.CTkLabel(
            right_card,
            text="⚙️ Configuración de Salida",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#ffffff"
        )
        lbl_opts.pack(anchor="w", padx=20, pady=(20, 14))

        # Información de tamaño y centrado
        info_box = ctk.CTkFrame(right_card, fg_color="#081528", corner_radius=12)
        info_box.pack(fill="x", padx=20, pady=(0, 18))

        info_lbl = ctk.CTkLabel(
            info_box,
            text="• Formato: Tamaño Carta (Letter)\n• Dimensiones: 612 x 792 pt\n• Escalado automático proporcional\n• Centrado de páginas con margen",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#94a3b8",
            justify="left"
        )
        info_lbl.pack(padx=14, pady=12, anchor="w")

        # Conteo de archivos
        self.lbl_count = ctk.CTkLabel(
            right_card,
            text="Archivos cargados: 0",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_count.pack(anchor="w", padx=20, pady=(0, 10))

        # Estado del proceso
        self.lbl_status = ctk.CTkLabel(
            right_card,
            text="Listo para iniciar.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#64748b",
            wraplength=270,
            justify="left"
        )
        self.lbl_status.pack(anchor="w", padx=20, pady=(0, 14))

        # Barra de progreso
        self.progress_bar = ctk.CTkProgressBar(right_card, height=10, corner_radius=5)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=20, pady=(0, 24))

        # Botón de Procesar y Unir
        self.btn_process = ctk.CTkButton(
            right_card,
            text="🚀 CONVERTIR Y UNIR",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            height=48,
            corner_radius=12,
            fg_color="#00b4d8",
            hover_color="#0096c7",
            text_color="#031629",
            command=self.start_processing
        )
        self.btn_process.pack(fill="x", padx=20, pady=(0, 15))

    # =========================================================================
    # GESTIÓN DE ARCHIVOS
    # =========================================================================
    def add_files(self):
        selected = filedialog.askopenfilenames(
            title="Seleccionar documentos de Word",
            filetypes=[("Documentos de Word", "*.docx *.doc")]
        )
        for path in selected:
            if path not in self.files:
                self.files.append(path)
                self.listbox.insert(tk.END, f"📄 {os.path.basename(path)}")
        self._update_status()

    def remove_selected(self):
        idx = self.listbox.curselection()
        if not idx:
            return
        index = idx[0]
        self.listbox.delete(index)
        del self.files[index]
        self._update_status()

    def clear_list(self):
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self._update_status()

    def move_up(self):
        idx = self.listbox.curselection()
        if not idx or idx[0] == 0:
            return
        i = idx[0]
        self.files[i - 1], self.files[i] = self.files[i], self.files[i - 1]
        text = self.listbox.get(i)
        self.listbox.delete(i)
        self.listbox.insert(i - 1, text)
        self.listbox.selection_set(i - 1)

    def move_down(self):
        idx = self.listbox.curselection()
        if not idx or idx[0] == len(self.files) - 1:
            return
        i = idx[0]
        self.files[i + 1], self.files[i] = self.files[i], self.files[i + 1]
        text = self.listbox.get(i)
        self.listbox.delete(i)
        self.listbox.insert(i + 1, text)
        self.listbox.selection_set(i + 1)

    def _update_status(self):
        count = len(self.files)
        self.lbl_count.configure(text=f"Archivos cargados: {count}")
        if count == 0:
            self.lbl_status.configure(text="Agregue al menos un documento Word.", text_color="#64748b")
        else:
            self.lbl_status.configure(text=f"{count} documento(s) listos para convertir.", text_color="#38bdf8")

    def _scale_page_to_letter(self, original_page):
        """Escala proporcionalmente la página y la centra en lienzo tamaño Carta"""
        orig_w = float(original_page.mediabox.width)
        orig_h = float(original_page.mediabox.height)

        scale = min(self.LETTER_WIDTH / orig_w, self.LETTER_HEIGHT / orig_h)
        scaled_w = orig_w * scale
        scaled_h = orig_h * scale

        offset_x = (self.LETTER_WIDTH - scaled_w) / 2
        offset_y = (self.LETTER_HEIGHT - scaled_h) / 2

        letter_page = PageObject.create_blank_page(
            width=self.LETTER_WIDTH,
            height=self.LETTER_HEIGHT
        )
        transform = Transformation().scale(scale).translate(offset_x, offset_y)
        letter_page.merge_transformed_page(original_page, transform)
        return letter_page

    # =========================================================================
    # PROCESAMIENTO CON HILOS
    # =========================================================================
    def start_processing(self):
        if not self.files:
            messagebox.showwarning("Sin archivos", "Por favor agregue al menos un documento de Word (.docx o .doc).")
            return

        if self.is_processing:
            return

        output_path = filedialog.asksaveasfilename(
            title="Guardar PDF unificado como",
            defaultextension=".pdf",
            filetypes=[("Archivo PDF", "*.pdf")]
        )
        if not output_path:
            return

        self.is_processing = True
        self.btn_process.configure(state="disabled", text="PROCESANDO...")
        self.progress_bar.set(0)

        threading.Thread(target=self._process_worker, args=(output_path,), daemon=True).start()

    def _process_worker(self, output_path):
        writer = PdfWriter()
        total = len(self.files)

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                for idx, doc_path in enumerate(self.files):
                    fname = os.path.basename(doc_path)
                    pct = idx / total
                    self._safe_update_ui(f"Convirtiendo ({idx + 1}/{total}): {fname}...", "#38bdf8", pct)

                    temp_pdf = os.path.join(temp_dir, f"temp_{idx}.pdf")
                    convert_single_word_to_pdf(doc_path, temp_pdf)

                    reader = PdfReader(temp_pdf)
                    for page in reader.pages:
                        letter_page = self._scale_page_to_letter(page)
                        writer.add_page(letter_page)

                self._safe_update_ui("Unificando y guardando archivo final...", "#38bdf8", 0.95)
                with open(output_path, "wb") as f_out:
                    writer.write(f_out)

            self._safe_update_ui("¡Completado con éxito!", "#4ade80", 1.0)
            self.after(200, lambda: self._on_success(output_path))

        except Exception as e:
            self._safe_update_ui(f"Error: {e}", "#f87171", 0)
            self.after(200, lambda: messagebox.showerror("Error al procesar", f"Ocurrió un error durante la conversión:\n{str(e)}"))
        finally:
            self.is_processing = False
            self.after(0, lambda: self.btn_process.configure(state="normal", text="🚀 CONVERTIR Y UNIR"))

    def _safe_update_ui(self, status_text, color, progress):
        try:
            if self.winfo_exists():
                self.after(0, lambda: self.lbl_status.configure(text=status_text, text_color=color))
                self.after(0, lambda: self.progress_bar.set(progress))
        except Exception:
            pass

    def _on_success(self, output_path):
        messagebox.showinfo("Proceso Completado", f"El archivo PDF unificado tamaño Carta se guardó con éxito en:\n\n{output_path}")
        try:
            dest_folder = os.path.dirname(os.path.abspath(output_path))
            if os.name == 'nt':
                os.startfile(dest_folder)
        except Exception:
            pass


# -----------------------------------------------------------------------------
# PUNTOS DE ENTRADA
# -----------------------------------------------------------------------------
def create_view(parent, on_back=None):
    """Punto de entrada estándar para incrustar en el Lanzador"""
    check_for_pending_lanzador_update()
    return WordToPdfMergerView(parent, on_back=on_back)


def main():
    """Ejecución independiente del convertidor"""
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Automatización de Archivos - Convertidor Word a PDF")
    root.geometry("1040x700")
    root.minsize(880, 580)
    root.configure(fg_color="#0b1626")

    view = WordToPdfMergerView(root, on_back=root.destroy)
    view.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()