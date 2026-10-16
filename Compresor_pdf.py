import os
import sys
import shutil
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


QUALITY_PRESETS = {
    "Pantalla (screen) - más pequeño / menor calidad": "/screen",
    "Ebook (ebook) - balanceado": "/ebook",
    "Impresora (printer) - mejor calidad": "/printer",
    "Preprensa (prepress) - máxima calidad": "/prepress",
    "Default (default) - estándar": "/default",
}


def default_gs_command():
    """
    Devuelve el ejecutable de Ghostscript típico según el SO.
    - Windows: gswin64c.exe o gswin32c.exe
    - Linux/macOS: gs
    """
    if os.name == "nt":
        for exe in ("gswin64c.exe", "gswin32c.exe", "gs.exe"):
            if shutil.which(exe):
                return exe
        return "gswin64c.exe"  # fallback
    else:
        return "gs"


def run_ghostscript(gs_exe, input_pdf, output_pdf, preset):
    """
    Ejecuta Ghostscript para comprimir.
    """
    # Ghostscript requiere -dSAFER para modo seguro
    # -dCompatibilityLevel=1.4 suele funcionar bien para compresión general
    # -dPDFSETTINGS define el preset de calidad/tamaño
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

    # En Windows, ocultar ventana de consola
    creationflags = 0
    startupinfo = None
    if os.name == "nt":
        creationflags = subprocess.CREATE_NO_WINDOW
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

def resource_path(relative_path):
    """
    Obtiene la ruta correcta tanto en modo script como en .exe (PyInstaller)
    """
    try:
        base_path = sys._MEIPASS  # cuando es .exe
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class PDFCompressorGUI:
    def __init__(self, root):
        self.root = root
        icon_path = resource_path("pdf.ico")
        root.iconbitmap(icon_path)
        self.root.title("Compresor de PDF (Ghostscript)")
        self.root.geometry("760x340")

        self.gs_path = tk.StringVar(value=default_gs_command())
        self.in_path = tk.StringVar(value="")
        self.out_path = tk.StringVar(value="")
        self.quality_label = tk.StringVar(value=list(QUALITY_PRESETS.keys())[1])  # ebook por defecto

        frm = tk.Frame(root, padx=12, pady=12)
        frm.pack(fill="both", expand=True)

        # Ghostscript path
        tk.Label(frm, text="Ghostscript (ejecutable):").grid(row=0, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.gs_path, width=58).grid(row=0, column=1, sticky="we", padx=8)
        tk.Button(frm, text="Buscar...", command=self.pick_gs).grid(row=0, column=2, sticky="e")

        # Input
        tk.Label(frm, text="PDF de entrada:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        tk.Entry(frm, textvariable=self.in_path, width=58).grid(row=1, column=1, sticky="we", padx=8, pady=(10, 0))
        tk.Button(frm, text="Seleccionar...", command=self.pick_input).grid(row=1, column=2, sticky="e", pady=(10, 0))

        # Output
        tk.Label(frm, text="PDF comprimido (salida):").grid(row=2, column=0, sticky="w", pady=(10, 0))
        tk.Entry(frm, textvariable=self.out_path, width=58).grid(row=2, column=1, sticky="we", padx=8, pady=(10, 0))
        tk.Button(frm, text="Guardar como...", command=self.pick_output).grid(row=2, column=2, sticky="e", pady=(10, 0))

        # Quality preset
        tk.Label(frm, text="Nivel de compresión/calidad:").grid(row=3, column=0, sticky="w", pady=(12, 0))
        self.quality_combo = ttk.Combobox(
            frm,
            values=list(QUALITY_PRESETS.keys()),
            textvariable=self.quality_label,
            state="readonly",
            width=55,
        )
        self.quality_combo.grid(row=3, column=1, sticky="w", padx=8, pady=(12, 0))

        # Buttons
        btns = tk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=3, sticky="w", pady=(16, 0))

        tk.Button(btns, text="Comprimir", command=self.compress).pack(side="left")
        tk.Button(btns, text="Limpiar", command=self.clear).pack(side="left", padx=10)
        tk.Button(btns, text="Salir", command=root.quit).pack(side="left")

        # Status + log
        self.status = tk.StringVar(value="Listo.")
        tk.Label(frm, textvariable=self.status, anchor="w").grid(row=5, column=0, columnspan=3, sticky="we", pady=(12, 4))

        self.log = tk.Text(frm, height=8, wrap="word")
        self.log.grid(row=6, column=0, columnspan=3, sticky="nsew")

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(6, weight=1)

    def write_log(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def pick_gs(self):
        path = filedialog.askopenfilename(
            title="Selecciona el ejecutable de Ghostscript",
            filetypes=[("Ejecutable", "*.exe" if os.name == "nt" else "*")],
        )
        if path:
            self.gs_path.set(path)

    def pick_input(self):
        path = filedialog.askopenfilename(
            title="Selecciona PDF de entrada",
            filetypes=[("PDF", "*.pdf")],
        )
        if not path:
            return
        self.in_path.set(path)

        # Sugerir salida automática
        base, ext = os.path.splitext(path)
        suggested = base + "_comprimido.pdf"
        if not self.out_path.get():
            self.out_path.set(suggested)

    def pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Guardar PDF comprimido como",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if path:
            self.out_path.set(path)

    def clear(self):
        self.in_path.set("")
        self.out_path.set("")
        self.quality_label.set(list(QUALITY_PRESETS.keys())[1])
        self.status.set("Listo.")
        self.log.delete("1.0", "end")

    def compress(self):
        gs_exe = self.gs_path.get().strip()
        in_pdf = self.in_path.get().strip()
        out_pdf = self.out_path.get().strip()
        preset = QUALITY_PRESETS.get(self.quality_label.get(), "/ebook")

        # Validaciones
        if not in_pdf or not os.path.isfile(in_pdf):
            messagebox.showerror("Error", "Selecciona un PDF de entrada válido.")
            return

        if not out_pdf:
            messagebox.showerror("Error", "Selecciona el archivo de salida.")
            return

        # Resolver ejecutable: si no es ruta, intentar por PATH
        gs_resolved = gs_exe
        if os.path.sep not in gs_exe and (os.name != "nt" or not gs_exe.lower().endswith(".exe")):
            # Puede ser 'gs' en linux/mac
            pass

        if os.path.isfile(gs_exe):
            gs_resolved = gs_exe
        else:
            found = shutil.which(gs_exe)
            if found:
                gs_resolved = found

        if not gs_resolved or (not os.path.isfile(gs_resolved) and not shutil.which(gs_exe)):
            messagebox.showerror(
                "Ghostscript no encontrado",
                "No se encontró Ghostscript.\n\n"
                "Solución:\n"
                "- Instala Ghostscript y asegúrate que esté en el PATH, o\n"
                "- Usa el botón 'Buscar...' para seleccionar el ejecutable (gs / gswin64c.exe)."
            )
            return

        # Ejecutar
        try:
            self.status.set("Comprimiendo... (esto puede tardar)")
            self.root.update_idletasks()

            # Tamaños para comparar
            in_size = os.path.getsize(in_pdf)

            code, out, err, cmd = run_ghostscript(gs_resolved, in_pdf, out_pdf, preset)

            self.write_log("Comando ejecutado:")
            self.write_log(" ".join(cmd))
            self.write_log("")

            if code != 0:
                self.write_log("ERROR (stderr):")
                self.write_log(err.strip() or "(sin salida)")
                self.status.set("Falló la compresión.")
                messagebox.showerror("Error", f"Ghostscript falló.\n\nDetalles:\n{err.strip() or 'Sin detalles'}")
                return

            # Verificar salida
            if not os.path.isfile(out_pdf):
                self.status.set("No se generó el archivo de salida.")
                messagebox.showerror("Error", "No se generó el PDF de salida.")
                return

            out_size = os.path.getsize(out_pdf)
            reduction = (1 - (out_size / in_size)) * 100 if in_size > 0 else 0

            self.write_log("OK.")
            self.write_log(f"Tamaño original: {in_size/1024:.1f} KB")
            self.write_log(f"Tamaño comprimido: {out_size/1024:.1f} KB")
            self.write_log(f"Reducción aproximada: {reduction:.1f}%")

            self.status.set("Listo. PDF comprimido creado.")
            messagebox.showinfo(
                "Éxito",
                f"PDF comprimido creado:\n{os.path.abspath(out_pdf)}\n\n"
                f"Original: {in_size/1024:.1f} KB\n"
                f"Comprimido: {out_size/1024:.1f} KB\n"
                f"Reducción: {reduction:.1f}%"
            )

        except Exception as e:
            self.status.set("Error.")
            messagebox.showerror("Error", f"Ocurrió un problema:\n{e}")


def main():
    # Ejecuta:
    #   pip install tk
    # Tkinter suele venir con Python en Windows/macOS, en Linux puede requerir:
    #   sudo apt-get install python3-tk
    root = tk.Tk()
    app = PDFCompressorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
