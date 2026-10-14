"""
=============================================================================
UNIFICADOR DE DOCUMENTOS PDF (CON AJUSTE A TAMAÑO CARTA)
=============================================================================
Permite seleccionar múltiples archivos PDF simultáneamente, organizarlos y
reordenarlos visualmente en la interfaz gráfica, y unificarlos en un solo archivo.
"""
import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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


class PDFUnifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unificador de PDFs - Organizar y Unir")
        self.root.geometry("820x560")
        self.root.minsize(720, 480)

        # Establecer icono si existe
        self._set_window_icon()

        # Lista de archivos seleccionados: dict con path, name, pages, size
        self.pdf_items = []

        # Estilo de interfaz
        self._setup_styles()

        # Construir UI
        self._create_widgets()

        # Verificar dependencias
        if PdfReader is None:
            messagebox.showwarning(
                "Librería requerida",
                "No se encontró 'pypdf' o 'PyPDF2'.\n"
                "Instálalo con: pip install pypdf"
            )

    def _set_window_icon(self):
        """Busca el archivo pdf.ico en varias rutas posibles"""
        possible_paths = [
            "pdf.ico",
            os.path.join(os.path.dirname(__file__), "pdf.ico"),
            os.path.join(os.path.dirname(__file__), "..", "pdf.ico")
        ]
        for p in possible_paths:
            if os.path.exists(p):
                try:
                    self.root.iconbitmap(p)
                    break
                except Exception:
                    pass

    def _setup_styles(self):
        """Configuración de estilos ttk modernos"""
        style = ttk.Style()
        style.theme_use("clam")

        # Configurar colores de la cabecera y el Treeview
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background="#0e223d",
            foreground="#ffffff",
            relief="flat",
            padding=6
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#193c66")]
        )
        style.configure(
            "Treeview",
            font=("Segoe UI", 10),
            rowheight=28,
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#0f172a"
        )
        style.map(
            "Treeview",
            background=[("selected", "#00b4d8")],
            foreground=[("selected", "#ffffff")]
        )

    def _create_widgets(self):
        # 1. Cabecera superior moderna
        header_frame = tk.Frame(self.root, bg="#0b1626", padx=20, pady=14)
        header_frame.pack(fill="x")

        lbl_title = tk.Label(
            header_frame,
            text="UNIFICADOR DE DOCUMENTOS PDF",
            font=("Segoe UI", 16, "bold"),
            fg="#ffffff",
            bg="#0b1626"
        )
        lbl_title.pack(anchor="w")

        lbl_desc = tk.Label(
            header_frame,
            text="Selecciona varios archivos a la vez, ordénalos con las flechas y únelos en un solo PDF.",
            font=("Segoe UI", 10),
            fg="#94a3b8",
            bg="#0b1626"
        )
        lbl_desc.pack(anchor="w", pady=(2, 0))

        # 2. Barra de herramientas de selección
        toolbar = tk.Frame(self.root, padx=15, pady=10, bg="#f1f5f9")
        toolbar.pack(fill="x")

        btn_add = tk.Button(
            toolbar,
            text="➕ Seleccionar / Agregar PDFs",
            font=("Segoe UI", 10, "bold"),
            bg="#00b4d8",
            fg="#02293f",
            activebackground="#0096c7",
            activeforeground="#ffffff",
            padx=14,
            pady=6,
            relief="flat",
            cursor="hand2",
            command=self.select_multiple_pdfs
        )
        btn_add.pack(side="left", padx=(0, 10))

        btn_clear = tk.Button(
            toolbar,
            text="🗑 Limpiar Lista",
            font=("Segoe UI", 10),
            bg="#e2e8f0",
            fg="#334155",
            activebackground="#cbd5e1",
            padx=10,
            pady=6,
            relief="flat",
            cursor="hand2",
            command=self.clear_all
        )
        btn_clear.pack(side="left")

        # 3. Contenedor central (Tabla de archivos + Botones de ordenación lateral)
        center_frame = tk.Frame(self.root, padx=15, pady=5)
        center_frame.pack(fill="both", expand=True)

        # Frame tabla
        table_frame = tk.Frame(center_frame)
        table_frame.pack(side="left", fill="both", expand=True)

        columns = ("pos", "name", "pages", "size", "path")
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse"
        )

        self.tree.heading("pos", text="# Orden")
        self.tree.heading("name", text="Nombre del Archivo")
        self.tree.heading("pages", text="Páginas")
        self.tree.heading("size", text="Tamaño")
        self.tree.heading("path", text="Ubicación")

        self.tree.column("pos", width=70, anchor="center")
        self.tree.column("name", width=260, anchor="w")
        self.tree.column("pages", width=80, anchor="center")
        self.tree.column("size", width=90, anchor="center")
        self.tree.column("path", width=240, anchor="w")

        # Scrollbars para la tabla
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Eventos del Treeview
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        # Botones de ordenación a la derecha
        order_frame = tk.Frame(center_frame, padx=10, pady=5)
        order_frame.pack(side="right", fill="y")

        tk.Label(order_frame, text="Organizar:", font=("Segoe UI", 10, "bold"), fg="#475569").pack(anchor="w", pady=(0, 6))

        btn_up = tk.Button(
            order_frame,
            text="▲ Mover Arriba",
            font=("Segoe UI", 9, "bold"),
            bg="#f8fafc",
            fg="#0f172a",
            activebackground="#e2e8f0",
            relief="groove",
            width=15,
            pady=5,
            cursor="hand2",
            command=self.move_up
        )
        btn_up.pack(pady=4)

        btn_down = tk.Button(
            order_frame,
            text="▼ Mover Abajo",
            font=("Segoe UI", 9, "bold"),
            bg="#f8fafc",
            fg="#0f172a",
            activebackground="#e2e8f0",
            relief="groove",
            width=15,
            pady=5,
            cursor="hand2",
            command=self.move_down
        )
        btn_down.pack(pady=4)

        ttk.Separator(order_frame, orient="horizontal").pack(fill="x", pady=10)

        btn_remove = tk.Button(
            order_frame,
            text="✖ Quitar Archivo",
            font=("Segoe UI", 9),
            bg="#fee2e2",
            fg="#b91c1c",
            activebackground="#fecaca",
            relief="flat",
            width=15,
            pady=5,
            cursor="hand2",
            command=self.remove_selected
        )
        btn_remove.pack(pady=4)

        btn_open_file = tk.Button(
            order_frame,
            text="👁 Ver PDF",
            font=("Segoe UI", 9),
            bg="#f1f5f9",
            fg="#334155",
            activebackground="#e2e8f0",
            relief="flat",
            width=15,
            pady=5,
            cursor="hand2",
            command=self.open_selected_pdf
        )
        btn_open_file.pack(pady=(10, 4))

        # 4. Pie de página y opciones finales
        footer = tk.Frame(self.root, bg="#f8fafc", padx=20, pady=12, relief="groove", bd=1)
        footer.pack(fill="x", side="bottom")

        # Checkbox para tamaño carta
        self.fit_letter_var = tk.BooleanVar(value=True)
        chk_letter = tk.Checkbutton(
            footer,
            text="Ajustar automáticamente todas las páginas a tamaño carta (Letter)",
            variable=self.fit_letter_var,
            font=("Segoe UI", 10),
            bg="#f8fafc",
            activebackground="#f8fafc",
            fg="#1e293b"
        )
        chk_letter.pack(side="left")

        # Botón Unificar
        btn_unify = tk.Button(
            footer,
            text="⚡ UNIFICAR Y GUARDAR PDF",
            font=("Segoe UI", 11, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            padx=18,
            pady=8,
            relief="flat",
            cursor="hand2",
            command=self.unify_and_save
        )
        btn_unify.pack(side="right")

        # Etiqueta de resumen
        self.lbl_summary = tk.Label(
            self.root,
            text="0 archivos seleccionados (0 páginas en total)",
            font=("Segoe UI", 9),
            fg="#64748b",
            padx=20,
            pady=4
        )
        self.lbl_summary.pack(anchor="w", side="bottom")

    # =========================================================================
    # GESTIÓN Y SELECCIÓN DE ARCHIVOS
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
            # Evitar duplicados consecutivos exactos
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
            # Seleccionar el último elemento añadido
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

        # Actualizar resumen inferior
        count = len(self.pdf_items)
        pages_text = f"{total_pages} páginas" if total_pages > 0 else "0 páginas"
        self.lbl_summary.config(
            text=f"{count} archivo(s) listado(s) para unificar ({pages_text} en total)"
        )

        if keep_selected_index is not None and 0 <= keep_selected_index < len(self.pdf_items):
            self._select_tree_row(keep_selected_index)

    def _select_tree_row(self, index):
        """Selecciona una fila del Treeview por índice numérico"""
        iid = str(index)
        if self.tree.exists(iid):
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)

    def get_selected_index(self):
        """Devuelve el índice seleccionado o None"""
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except ValueError:
            return None

    # =========================================================================
    # ORGANIZACIÓN / REORDENAR ELEMENTOS EN LA INTERFAZ
    # =========================================================================
    def move_up(self):
        """Mueve el archivo seleccionado una posición hacia arriba"""
        idx = self.get_selected_index()
        if idx is None or idx <= 0:
            return

        # Intercambiar posiciones
        self.pdf_items[idx], self.pdf_items[idx - 1] = self.pdf_items[idx - 1], self.pdf_items[idx]
        self.refresh_table(keep_selected_index=idx - 1)

    def move_down(self):
        """Mueve el archivo seleccionado una posición hacia abajo"""
        idx = self.get_selected_index()
        if idx is None or idx >= len(self.pdf_items) - 1:
            return

        # Intercambiar posiciones
        self.pdf_items[idx], self.pdf_items[idx + 1] = self.pdf_items[idx + 1], self.pdf_items[idx]
        self.refresh_table(keep_selected_index=idx + 1)

    def remove_selected(self):
        """Quita el archivo seleccionado de la lista"""
        idx = self.get_selected_index()
        if idx is None:
            messagebox.showinfo("Aviso", "Selecciona un archivo de la lista para quitarlo.")
            return

        removed = self.pdf_items.pop(idx)
        new_select = min(idx, len(self.pdf_items) - 1) if self.pdf_items else None
        self.refresh_table(keep_selected_index=new_select)

    def clear_all(self):
        """Limpia la lista completa tras confirmación si hay elementos"""
        if not self.pdf_items:
            return
        if messagebox.askyesno("Limpiar Lista", "¿Deseas vaciar la lista de PDFs?"):
            self.pdf_items.clear()
            self.refresh_table()

    def open_selected_pdf(self):
        """Abre el archivo PDF seleccionado con el visor predeterminado del sistema"""
        idx = self.get_selected_index()
        if idx is None:
            messagebox.showinfo("Aviso", "Selecciona un archivo para abrir.")
            return

        path = self.pdf_items[idx]["path"]
        if os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
        else:
            messagebox.showerror("Error", "El archivo ya no existe en la ruta original.")

    def _on_double_click(self, event):
        """Doble clic en un archivo para visualizarlo"""
        self.open_selected_pdf()

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

        # Solicitar ruta de guardado al usuario
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

            # Escribir archivo final
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


def main():
    root = tk.Tk()
    app = PDFUnifierApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
