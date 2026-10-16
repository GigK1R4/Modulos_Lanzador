import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx2pdf import convert
from pypdf import PdfReader, PdfWriter, Transformation, PageObject

class DocxToPdfMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertidor y Unificador a PDF (Tamaño Carta)")
        self.root.geometry("620x460")
        self.root.resizable(True, True)

        # Dimensiones de tamaño Carta (Letter) en puntos PostScript (8.5 x 11 pulgadas)
        self.LETTER_WIDTH = 612.0
        self.LETTER_HEIGHT = 792.0

        self.files = []
        self._build_ui()

    def _build_ui(self):
        # Marco superior: botones de acción
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=tk.X)

        btn_add = ttk.Button(top_frame, text="➕ Agregar Archivos Word", command=self.add_files)
        btn_add.pack(side=tk.LEFT, padx=5)

        btn_clear = ttk.Button(top_frame, text="🗑️ Limpiar Lista", command=self.clear_list)
        btn_clear.pack(side=tk.LEFT, padx=5)

        # Marco central: Lista y botones para mover posición
        mid_frame = ttk.Frame(self.root, padding=10)
        mid_frame.pack(fill=tk.BOTH, expand=True)

        list_scroll = ttk.Scrollbar(mid_frame, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(
            mid_frame,
            selectmode=tk.SINGLE,
            yscrollcommand=list_scroll.set,
            font=("Segoe UI", 10),
        )
        list_scroll.config(command=self.listbox.yview)

        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.pack(side=tk.LEFT, fill=tk.Y)

        # Botones para ordenar
        move_frame = ttk.Frame(mid_frame, padding=(10, 0))
        move_frame.pack(side=tk.RIGHT, fill=tk.Y)

        btn_up = ttk.Button(move_frame, text="🔼 Subir", command=self.move_up)
        btn_up.pack(fill=tk.X, pady=5)

        btn_down = ttk.Button(move_frame, text="🔽 Bajar", command=self.move_down)
        btn_down.pack(fill=tk.X, pady=5)

        btn_remove = ttk.Button(move_frame, text="❌ Eliminar", command=self.remove_selected)
        btn_remove.pack(fill=tk.X, pady=15)

        # Marco inferior: barra de estado y botón procesar
        bottom_frame = ttk.Frame(self.root, padding=10)
        bottom_frame.pack(fill=tk.X)

        self.lbl_status = ttk.Label(bottom_frame, text="Archivos seleccionados: 0", foreground="gray")
        self.lbl_status.pack(side=tk.LEFT)

        btn_process = ttk.Button(bottom_frame, text="🚀 Convertir y Unir a Carta", command=self.process_files)
        btn_process.pack(side=tk.RIGHT)

    def add_files(self):
        selected = filedialog.askopenfilenames(
            title="Seleccionar documentos de Word",
            filetypes=[("Documentos de Word", "*.docx *.doc")]
        )
        for path in selected:
            if path not in self.files:
                self.files.append(path)
                self.listbox.insert(tk.END, os.path.basename(path))
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
        self.lbl_status.config(text=f"Archivos seleccionados: {len(self.files)}")

    def _scale_page_to_letter(self, original_page):
        """Escala proporcionalmente la página original y la centra en un lienzo tamaño Carta."""
        orig_w = float(original_page.mediabox.width)
        orig_h = float(original_page.mediabox.height)

        # Factor de escala manteniendo la relación de aspecto
        scale = min(self.LETTER_WIDTH / orig_w, self.LETTER_HEIGHT / orig_h)
        scaled_w = orig_w * scale
        scaled_h = orig_h * scale

        # Margen de centrado en la página tamaño carta
        offset_x = (self.LETTER_WIDTH - scaled_w) / 2
        offset_y = (self.LETTER_HEIGHT - scaled_h) / 2

        # Crear página en blanco tamaño Carta
        letter_page = PageObject.create_blank_page(
            width=self.LETTER_WIDTH,
            height=self.LETTER_HEIGHT
        )

        # Aplicar transformación (escala y traslación)
        transform = Transformation().scale(scale).translate(offset_x, offset_y)
        letter_page.merge_transformed_page(original_page, transform)
        return letter_page

    def process_files(self):
        if not self.files:
            messagebox.showwarning("Sin archivos", "Por favor agrega al menos un documento de Word.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Guardar PDF unificado como",
            defaultextension=".pdf",
            filetypes=[("Archivo PDF", "*.pdf")]
        )
        if not output_path:
            return

        self.root.config(cursor="wait")
        self.lbl_status.config(text="Procesando... por favor espera.", foreground="blue")
        self.root.update()

        writer = PdfWriter()

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                for idx, doc_path in enumerate(self.files):
                    temp_pdf = os.path.join(temp_dir, f"temp_{idx}.pdf")
                    
                    # 1. Convertir Word a PDF
                    convert(doc_path, temp_pdf)

                    # 2. Leer y normalizar cada página a tamaño Carta
                    reader = PdfReader(temp_pdf)
                    for page in reader.pages:
                        letter_page = self._scale_page_to_letter(page)
                        writer.add_page(letter_page)

                # 3. Guardar el archivo final unificado
                with open(output_path, "wb") as f_out:
                    writer.write(f_out)

            messagebox.showinfo("¡Éxito!", f"El PDF unificado se guardó correctamente en:\n{output_path}")
            self.lbl_status.config(text="Proceso completado con éxito.", foreground="green")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso:\n{str(e)}")
            self.lbl_status.config(text="Error al procesar.", foreground="red")
        finally:
            self.root.config(cursor="")


if __name__ == "__main__":
    root = tk.Tk()
    app = DocxToPdfMergerApp(root)
    root.mainloop()