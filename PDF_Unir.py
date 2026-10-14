"""
=============================================================================
UNIFICADOR DE DOCUMENTOS PDF (MODERNO - CUSTOMTKINTER)
=============================================================================
Permite seleccionar múltiples archivos PDF simultáneamente, organizarlos y
reordenarlos visualmente dentro de la misma ventana del Lanzador, y unificarlos.
"""
import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

# Soporte resiliente para pypdf o PyPDF2
try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        PdfReader = None
        PdfWriter = None

# Dimensiones tamaño carta en puntos tipográficos (8.5 x 11 pulgadas a 72 dpi)
LETTER_W, LETTER_H = 612, 792


def fit_page_to_letter(page):
    """
    Devuelve una página escalada y centrada para encajar en una hoja tamaño carta.
    """
    mb = page.mediabox
    w = float(mb.width)
    h = float(mb.height)

    if w <= 0 or h <= 0:
        return page

    # Escalar preservando la proporción original
    scale = min(LETTER_W / w, LETTER_H / h)
    new_w = w * scale
    new_h = h * scale

    # Centrar la página en la hoja carta
    tx = (LETTER_W - new_w) / 2
    ty = (LETTER_H - new_h) / 2

    page.scale_by(scale)
    page.add_transformation((1, 0, 0, 1, tx, ty))

    page.mediabox.lower_left = (0, 0)
    page.mediabox.upper_right = (LETTER_W, LETTER_H)

    if hasattr(page, "cropbox") and page.cropbox:
        page.cropbox.lower_left = (0, 0)
        page.cropbox.upper_right = (LETTER_W, LETTER_H)

    return page


def format_file_size(size_bytes):
    """Formatea bytes a formato legible KB o MB"""
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.2f} MB"


def get_pdf_metadata(filepath):
    """Obtiene el número de páginas y tamaño del archivo PDF"""
    try:
        size_bytes = os.path.getsize(filepath)
        size_str = format_file_size(size_bytes)
        if PdfReader is not None:
            reader = PdfReader(filepath)
            num_pages = len(reader.pages)
        else:
            num_pages = "?"
        return num_pages, size_str
    except Exception:
        return "?", "0 KB"


class PDFUnirView(ctk.CTkFrame):
    """
    Vista embebible del Unificador de PDFs con diseño CustomTkinter idéntico al Lanzador.
    """
    preferred_size = (1020, 680)
    min_size = (880, 580)

    def __init__(self, parent, on_back=None, *args, **kwargs):
        super().__init__(parent, fg_color="transparent", *args, **kwargs)
        self.on_back = on_back
        self.pdf_items = []

        self._setup_treeview_styles()
        self._build_ui()

    def _setup_treeview_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Configurar colores que armonizan con el tema oscuro del Lanzador
        style.configure(
            "PDF.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background="#0e223d",
            foreground="#38bdf8",
            relief="flat",
            padding=7
        )
        style.map(
            "PDF.Treeview.Heading",
            background=[("active", "#1b3a5f")]
        )
        style.configure(
            "PDF.Treeview",
            font=("Segoe UI", 10),
            rowheight=30,
            background="#071426",
            fieldbackground="#071426",
            foreground="#f1f5f9",
            bordercolor="#1d3b61"
        )
        style.map(
            "PDF.Treeview",
            background=[("selected", "#00b4d8")],
            foreground=[("selected", "#02293f")]
        )

    def _build_ui(self):
        # 1. Barra superior: botón Volver y títulos
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

        # Cabecera principal
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
            text="UNIFICADOR DE DOCUMENTOS PDF",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#ffffff"
        )
        lbl_title.pack(anchor="w", padx=20, pady=(12, 2))

        lbl_sub = ctk.CTkLabel(
            header_card,
            text="Selecciona múltiples archivos a la vez, organízalos en la lista y únelos en un único documento.",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#8ea3bf"
        )
        lbl_sub.pack(anchor="w", padx=20, pady=(0, 12))

        # 2. Contenedor central (Tabla y Panel lateral de ordenación)
        center_frame = ctk.CTkFrame(self, fg_color="transparent")
        center_frame.pack(fill="both", expand=True, padx=25, pady=(0, 10))

        # Contenedor de la tabla con borde
        table_container = ctk.CTkFrame(
            center_frame,
            fg_color="#071426",
            border_width=1,
            border_color="#1d3b61",
            corner_radius=14
        )
        table_container.pack(side="left", fill="both", expand=True, padx=(0, 14))

        columns = ("pos", "name", "pages", "size", "path")
        self.tree = ttk.Treeview(
            table_container,
            columns=columns,
            show="headings",
            style="PDF.Treeview",
            selectmode="browse"
        )

        self.tree.heading("pos", text="# Orden")
        self.tree.heading("name", text="Nombre del Archivo")
        self.tree.heading("pages", text="Páginas")
        self.tree.heading("size", text="Tamaño")
        self.tree.heading("path", text="Ubicación Completa")

        self.tree.column("pos", width=65, anchor="center")
        self.tree.column("name", width=250, anchor="w")
        self.tree.column("pages", width=75, anchor="center")
        self.tree.column("size", width=85, anchor="center")
        self.tree.column("path", width=260, anchor="w")

        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=(10, 0))
        vsb.grid(row=0, column=1, sticky="ns", pady=(10, 0), padx=(2, 6))
        hsb.grid(row=1, column=0, sticky="ew", padx=(10, 0), pady=(2, 6))

        table_container.rowconfigure(0, weight=1)
        table_container.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda e: self.open_selected_pdf())
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        # Panel lateral de controles (Organizar)
        sidebar = ctk.CTkFrame(
            center_frame,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14,
            width=200
        )
        sidebar.pack(side="right", fill="y")
        sidebar.pack_propagate(False)

        lbl_actions = ctk.CTkLabel(
            sidebar,
            text="Organización",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#ffffff"
        )
        lbl_actions.pack(pady=(16, 12))

        btn_add = ctk.CTkButton(
            sidebar,
            text="➕ Agregar PDFs",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color="#00b4d8",
            hover_color="#0096c7",
            text_color="#02293f",
            height=36,
            corner_radius=8,
            command=self.select_multiple_pdfs
        )
        btn_add.pack(fill="x", padx=14, pady=5)

        btn_up = ctk.CTkButton(
            sidebar,
            text="▲ Mover Arriba",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#182c47",
            hover_color="#243f63",
            text_color="#ffffff",
            height=32,
            corner_radius=8,
            command=self.move_up
        )
        btn_up.pack(fill="x", padx=14, pady=5)

        btn_down = ctk.CTkButton(
            sidebar,
            text="▼ Mover Abajo",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#182c47",
            hover_color="#243f63",
            text_color="#ffffff",
            height=32,
            corner_radius=8,
            command=self.move_down
        )
        btn_down.pack(fill="x", padx=14, pady=5)

        btn_view = ctk.CTkButton(
            sidebar,
            text="👁 Previsualizar",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#10253f",
            hover_color="#18365c",
            text_color="#38bdf8",
            height=32,
            corner_radius=8,
            command=self.open_selected_pdf
        )
        btn_view.pack(fill="x", padx=14, pady=5)

        btn_remove = ctk.CTkButton(
            sidebar,
            text="✖ Quitar Archivo",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#451a1a",
            hover_color="#7f1d1d",
            text_color="#fca5a5",
            height=32,
            corner_radius=8,
            command=self.remove_selected
        )
        btn_remove.pack(fill="x", padx=14, pady=(15, 5))

        btn_clear = ctk.CTkButton(
            sidebar,
            text="🗑 Limpiar Lista",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="transparent",
            hover_color="#182c47",
            text_color="#94a3b8",
            height=30,
            command=self.clear_all
        )
        btn_clear.pack(fill="x", padx=14, pady=5)

        # 3. Pie de página (Resumen, Checkbox carta y Botón de Unificar)
        footer_card = ctk.CTkFrame(
            self,
            fg_color="#0e1f38",
            border_width=1,
            border_color="#1a3557",
            corner_radius=14
        )
        footer_card.pack(fill="x", padx=25, pady=(0, 15))

        left_footer = ctk.CTkFrame(footer_card, fg_color="transparent")
        left_footer.pack(side="left", padx=18, pady=12)

        self.lbl_summary = ctk.CTkLabel(
            left_footer,
            text="● 0 archivos listados (0 páginas en total)",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#00b4d8"
        )
        self.lbl_summary.pack(anchor="w")

        self.fit_letter_var = tk.BooleanVar(value=True)
        self.chk_letter = ctk.CTkCheckBox(
            left_footer,
            text="Ajustar automáticamente páginas a tamaño carta (Letter)",
            variable=self.fit_letter_var,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#e2e8f0",
            fg_color="#00b4d8",
            hover_color="#0096c7",
            border_color="#244770"
        )
        self.chk_letter.pack(anchor="w", pady=(6, 0))

        btn_unify = ctk.CTkButton(
            footer_card,
            text="⚡ UNIFICAR Y GUARDAR PDF",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#00b4d8",
            hover_color="#0096c7",
            text_color="#02293f",
            width=240,
            height=44,
            corner_radius=10,
            command=self.unify_and_save
        )
        btn_unify.pack(side="right", padx=18, pady=12)

    # =========================================================================
    # LÓGICA DE GESTIÓN Y SELECCIÓN
    # =========================================================================
    def select_multiple_pdfs(self):
        """Permite seleccionar varios archivos PDF al mismo tiempo"""
        paths = filedialog.askopenfilenames(
            title="Selecciona uno o más archivos PDF a unificar",
            filetypes=[("Archivos PDF (*.pdf)", "*.pdf"), ("Todos los archivos", "*.*")]
        )
        if not paths:
            return

        added_count = 0
        for p in paths:
            abs_p = os.path.abspath(p)
            if any(item["path"] == abs_p for item in self.pdf_items):
                continue

            num_pages, size_str = get_pdf_metadata(abs_p)
            self.pdf_items.append({
                "path": abs_p,
                "name": os.path.basename(abs_p),
                "pages": num_pages,
                "size": size_str
            })
            added_count += 1

        self.refresh_table()
        if added_count > 0:
            self._select_tree_row(len(self.pdf_items) - 1)

    def refresh_table(self, keep_selected_index=None):
        """Actualiza el Treeview con la lista ordenada actual"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        total_pages = 0
        for idx, item in enumerate(self.pdf_items):
            pos_label = f"{idx + 1:02d}"
            pages_val = item.get("pages", "?")
            if isinstance(pages_val, int):
                total_pages += pages_val

            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(pos_label, item["name"], str(pages_val), item["size"], item["path"])
            )

        count = len(self.pdf_items)
        pages_text = f"{total_pages} páginas" if total_pages > 0 else "0 páginas"
        self.lbl_summary.configure(
            text=f"● {count} archivo(s) para unificar ({pages_text} en total)"
        )

        if keep_selected_index is not None and 0 <= keep_selected_index < len(self.pdf_items):
            self._select_tree_row(keep_selected_index)

    def _select_tree_row(self, index):
        iid = str(index)
        if self.tree.exists(iid):
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)

    def get_selected_index(self):
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except ValueError:
            return None

    def move_up(self):
        """Mueve el archivo seleccionado una posición hacia arriba"""
        idx = self.get_selected_index()
        if idx is None or idx <= 0:
            return
        self.pdf_items[idx], self.pdf_items[idx - 1] = self.pdf_items[idx - 1], self.pdf_items[idx]
        self.refresh_table(keep_selected_index=idx - 1)

    def move_down(self):
        """Mueve el archivo seleccionado una posición hacia abajo"""
        idx = self.get_selected_index()
        if idx is None or idx >= len(self.pdf_items) - 1:
            return
        self.pdf_items[idx], self.pdf_items[idx + 1] = self.pdf_items[idx + 1], self.pdf_items[idx]
        self.refresh_table(keep_selected_index=idx + 1)

    def remove_selected(self):
        """Quita el archivo seleccionado de la lista"""
        idx = self.get_selected_index()
        if idx is None:
            messagebox.showinfo("Aviso", "Selecciona un archivo de la lista para quitarlo.")
            return
        self.pdf_items.pop(idx)
        new_select = min(idx, len(self.pdf_items) - 1) if self.pdf_items else None
        self.refresh_table(keep_selected_index=new_select)

    def clear_all(self):
        """Limpia la lista completa"""
        if not self.pdf_items:
            return
        if messagebox.askyesno("Limpiar Lista", "¿Deseas vaciar la lista de PDFs?"):
            self.pdf_items.clear()
            self.refresh_table()

    def open_selected_pdf(self):
        """Abre el archivo PDF seleccionado"""
        idx = self.get_selected_index()
        if idx is None:
            messagebox.showinfo("Aviso", "Selecciona un archivo para previsualizar.")
            return

        path = self.pdf_items[idx]["path"]
        if os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
        else:
            messagebox.showerror("Error", "El archivo ya no existe en la ruta original.")

    # =========================================================================
    # UNIFICAR Y GUARDAR
    # =========================================================================
    def unify_and_save(self):
        """Ejecuta la unificación de los PDFs en el orden exacto de la tabla"""
        if not self.pdf_items:
            messagebox.showwarning(
                "Sin archivos",
                "Por favor selecciona al menos dos archivos PDF para unificar."
            )
            return

        if len(self.pdf_items) < 2:
            if not messagebox.askyesno(
                "Un solo archivo",
                "Solo has seleccionado 1 archivo PDF. ¿Deseas procesarlo de todos modos?"
            ):
                return

        if PdfWriter is None or PdfReader is None:
            messagebox.showerror(
                "Librería faltante",
                "No se encuentra la librería 'pypdf'. Instálala con:\npip install pypdf"
            )
            return

        default_name = "PDF_Unificado.pdf"
        out_path = filedialog.asksaveasfilename(
            title="Guardar PDF Unificado",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("Archivo PDF (*.pdf)", "*.pdf")]
        )
        if not out_path:
            return

        writer = PdfWriter()
        adjust_letter = self.fit_letter_var.get()
        total_pages_unified = 0

        try:
            for item in self.pdf_items:
                path = item["path"]
                if not os.path.exists(path):
                    messagebox.showerror(
                        "Archivo no encontrado",
                        f"No se encontró el archivo:\n{path}\nLa unificación ha sido cancelada."
                    )
                    return

                reader = PdfReader(path)
                for page in reader.pages:
                    if adjust_letter:
                        page = fit_page_to_letter(page)
                    writer.add_page(page)
                    total_pages_unified += 1

            with open(out_path, "wb") as f:
                writer.write(f)

            ans = messagebox.askyesno(
                "¡Unificación Exitosa!",
                f"El PDF unificado se guardó correctamente con {total_pages_unified} páginas en:\n\n"
                f"{out_path}\n\n¿Deseas abrir el archivo ahora?"
            )
            if ans:
                try:
                    os.startfile(out_path)
                except Exception:
                    pass

        except Exception as e:
            messagebox.showerror("Error al unificar", f"Ocurrió un problema durante la unificación:\n{str(e)}")


def create_view(parent, on_back=None):
    """Punto de entrada para incrustar en la misma ventana de Lanzador.py"""
    return PDFUnirView(parent, on_back=on_back)


def main():
    """Ejecución independiente de prueba"""
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Automatización de Archivos - Unificador de PDFs")
    root.geometry("1000x680")
    root.minsize(880, 560)
    root.configure(fg_color="#0b1626")

    # Si existe icono
    for ico in ["pdf.ico", os.path.join(os.path.dirname(__file__), "..", "pdf.ico")]:
        if os.path.exists(ico):
            try:
                root.iconbitmap(ico)
                break
            except Exception:
                pass

    view = PDFUnirView(root, on_back=root.destroy)
    view.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()
