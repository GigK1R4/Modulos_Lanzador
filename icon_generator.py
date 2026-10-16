"""
Generador de iconos vectoriales/rasterizados con Pillow para el Lanzador.
Crea iconos elegantes en alta resolución con estilo minimalista moderno.
Soporta generación 100% en memoria (sin crear archivos en disco) o guardado en disco si se requiere.
"""
import os
import sys
from PIL import Image, ImageDraw, ImageFont

def set_hidden_attribute(path):
    """Establece el atributo de archivo/carpeta oculta en Windows si existe"""
    if os.name == 'nt' and os.path.exists(path):
        try:
            import ctypes
            # 0x02 = FILE_ATTRIBUTE_HIDDEN
            ctypes.windll.kernel32.SetFileAttributesW(os.path.abspath(path), 0x02)
        except Exception:
            pass

def ensure_icons_dir(base_dir=None):
    """Crea la estructura de carpetas de iconos solo si se requiere almacenar en disco"""
    if base_dir is None:
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(os.path.abspath(sys.executable))
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

    assets_dir = os.path.join(base_dir, "assets")
    os.makedirs(assets_dir, exist_ok=True)
    set_hidden_attribute(assets_dir)

    icons_dir = os.path.join(assets_dir, "icons")
    os.makedirs(icons_dir, exist_ok=True)
    set_hidden_attribute(icons_dir)

    return icons_dir

def create_unir_pdf_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de dos páginas de PDF fusionándose o documento con capas"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size
    
    # Documento trasero
    d.rounded_rectangle([w*0.18, h*0.15, w*0.62, h*0.75], radius=6, outline=color, width=5)
    d.line([w*0.30, h*0.35, w*0.50, h*0.35], fill=color, width=4)
    d.line([w*0.30, h*0.48, w*0.50, h*0.48], fill=color, width=4)
    
    # Documento delantero (superpuesto)
    d.rounded_rectangle([w*0.38, h*0.28, w*0.82, h*0.88], radius=6, fill="#ffffff", outline=color, width=5)
    d.line([w*0.48, h*0.45, w*0.72, h*0.45], fill=color, width=4)
    d.line([w*0.48, h*0.58, w*0.72, h*0.58], fill=color, width=4)
    d.line([w*0.48, h*0.70, w*0.65, h*0.70], fill=color, width=4)
    
    # Símbolo de suma/unión '+' en círculo pequeño
    cx, cy, r = w*0.35, h*0.78, 14
    d.ellipse([cx-r, cy-r, cx+r, cy+r], fill=color)
    d.line([cx-7, cy, cx+7, cy], fill="#ffffff", width=3)
    d.line([cx, cy-7, cx, cy+7], fill="#ffffff", width=3)
    
    if path:
        img.save(path, format="PNG")
    return img

def create_doc_masivos_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de documentos masivos / generación de archivos"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size
    
    # Hoja 1 atrás
    d.rounded_rectangle([w*0.15, h*0.12, w*0.58, h*0.68], radius=5, outline=color, width=4)
    # Hoja 2 medio
    d.rounded_rectangle([w*0.28, h*0.22, w*0.70, h*0.78], radius=5, fill="#ffffff", outline=color, width=4)
    # Hoja 3 frente
    d.rounded_rectangle([w*0.40, h*0.32, w*0.82, h*0.88], radius=5, fill="#ffffff", outline=color, width=5)
    
    # Líneas de texto en la hoja frontal
    d.line([w*0.50, h*0.48, w*0.72, h*0.48], fill=color, width=4)
    d.line([w*0.50, h*0.60, w*0.72, h*0.60], fill=color, width=4)
    d.line([w*0.50, h*0.72, w*0.65, h*0.72], fill=color, width=4)
    
    # Rayo o engranaje de automatización
    cx, cy, r = w*0.28, h*0.42, 12
    d.ellipse([cx-r, cy-r, cx+r, cy+r], fill=color)
    d.ellipse([cx-4, cy-4, cx+4, cy+4], fill="#ffffff")
    
    if path:
        img.save(path, format="PNG")
    return img

def create_edit_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de libreta con lápiz / editar"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size
    
    # Marco de libreta con esquina cortada/redondeada
    d.rounded_rectangle([w*0.20, h*0.25, w*0.75, h*0.85], radius=10, outline=color, width=6)
    # Lápiz diagonal
    d.line([w*0.40, h*0.60, w*0.78, h*0.22], fill=color, width=8)
    d.polygon([(w*0.34, h*0.66), (w*0.43, h*0.64), (w*0.36, h*0.57)], fill=color)
    if path:
        img.save(path, format="PNG")
    return img

def create_sync_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de actualización / flechas circulares"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size
    
    # Dos flechas curvas formando ciclo
    d.arc([w*0.22, h*0.22, w*0.78, h*0.78], start=30, end=150, fill=color, width=7)
    d.arc([w*0.22, h*0.22, w*0.78, h*0.78], start=210, end=330, fill=color, width=7)
    
    # Flecha punta 1 (arriba derecha)
    d.polygon([(w*0.65, h*0.20), (w*0.80, h*0.23), (w*0.75, h*0.38)], fill=color)
    # Flecha punta 2 (abajo izquierda)
    d.polygon([(w*0.35, h*0.80), (w*0.20, h*0.77), (w*0.25, h*0.62)], fill=color)
    
    if path:
        img.save(path, format="PNG")
    return img

def create_sheet_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de hoja de cálculo con 'X'"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size
    
    # Documento
    d.rounded_rectangle([w*0.22, h*0.15, w*0.78, h*0.85], radius=8, fill=color)
    # Letra X estilizada blanca
    d.line([w*0.40, h*0.42, w*0.60, h*0.65], fill="#ffffff", width=7)
    d.line([w*0.60, h*0.42, w*0.40, h*0.65], fill="#ffffff", width=7)
    if path:
        img.save(path, format="PNG")
    return img

def create_database_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de base de datos / discos"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size
    
    # Cilindro 1 (superior)
    d.ellipse([w*0.25, h*0.20, w*0.75, h*0.40], fill=color)
    # Cilindro 2 (medio)
    d.ellipse([w*0.25, h*0.40, w*0.75, h*0.60], fill=color)
    # Cilindro 3 (inferior)
    d.ellipse([w*0.25, h*0.60, w*0.75, h*0.80], fill=color)
    if path:
        img.save(path, format="PNG")
    return img

def create_script_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de script / código para módulos dinámicos (ventana con </>)"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size

    # Marco ventana de código
    d.rounded_rectangle([w*0.16, h*0.20, w*0.84, h*0.80], radius=8, outline=color, width=5)
    # Barra superior de la ventana
    d.line([w*0.16, h*0.35, w*0.84, h*0.35], fill=color, width=4)
    # 3 puntitos de ventana
    d.ellipse([w*0.24-3, h*0.275-3, w*0.24+3, h*0.275+3], fill=color)
    d.ellipse([w*0.32-3, h*0.275-3, w*0.32+3, h*0.275+3], fill=color)
    d.ellipse([w*0.40-3, h*0.275-3, w*0.40+3, h*0.275+3], fill=color)

    # Símbolo < / >
    # <
    d.line([w*0.38, h*0.57, w*0.30, h*0.57+5], fill=color, width=4)
    d.line([w*0.30, h*0.57+5, w*0.38, h*0.57+10], fill=color, width=4)
    # /
    d.line([w*0.53, h*0.50, w*0.47, h*0.72], fill=color, width=4)
    # >
    d.line([w*0.62, h*0.57, w*0.70, h*0.57+5], fill=color, width=4)
    d.line([w*0.70, h*0.57+5, w*0.62, h*0.57+10], fill=color, width=4)

    if path:
        img.save(path, format="PNG")
    return img

def create_lock_icon(path=None, size=(120, 120), color="#8892b0"):
    """Icono de candado para módulos bloqueados"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size
    
    # Arco del candado
    d.arc([w*0.35, h*0.20, w*0.65, h*0.55], start=180, end=0, fill=color, width=6)
    # Cuerpo del candado
    d.rounded_rectangle([w*0.28, h*0.45, w*0.72, h*0.82], radius=8, fill=color)
    # Ojo de la cerradura
    d.ellipse([w*0.46, h*0.56, w*0.54, h*0.66], fill="#ffffff")
    d.line([w*0.50, h*0.65, w*0.50, h*0.73], fill="#ffffff", width=4)
    if path:
        img.save(path, format="PNG")
    return img

def create_formato_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de formato oficial / vinculación con check y sello"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size

    # Hoja principal del formato
    d.rounded_rectangle([w*0.22, h*0.14, w*0.78, h*0.86], radius=7, fill="#ffffff", outline=color, width=5)
    
    # Líneas de texto del formato
    d.line([w*0.32, h*0.28, w*0.68, h*0.28], fill=color, width=4)
    d.line([w*0.32, h*0.40, w*0.68, h*0.40], fill=color, width=4)
    d.line([w*0.32, h*0.52, w*0.55, h*0.52], fill=color, width=4)
    
    # Círculo de sello de verificación abajo a la derecha
    cx, cy, r = w*0.66, h*0.68, 16
    d.ellipse([cx-r, cy-r, cx+r, cy+r], fill=color)
    # Checkmark blanco dentro del sello
    d.line([cx-7, cy, cx-2, cy+6], fill="#ffffff", width=3)
    d.line([cx-2, cy+6, cx+8, cy-5], fill="#ffffff", width=3)

    if path:
        img.save(path, format="PNG")
    return img

def create_compress_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de documento PDF con flechas de compresión hacia adentro"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size

    # Documento PDF base
    d.rounded_rectangle([w*0.25, h*0.14, w*0.75, h*0.86], radius=7, fill="#ffffff", outline=color, width=5)

    # Flecha superior apuntando hacia abajo
    d.line([w*0.50, h*0.24, w*0.50, h*0.44], fill=color, width=5)
    d.polygon([(w*0.41, h*0.37), (w*0.59, h*0.37), (w*0.50, h*0.47)], fill=color)

    # Flecha inferior apuntando hacia arriba
    d.line([w*0.50, h*0.76, w*0.50, h*0.56], fill=color, width=5)
    d.polygon([(w*0.41, h*0.63), (w*0.59, h*0.63), (w*0.50, h*0.53)], fill=color)

    # Línea central de compresión
    d.line([w*0.35, h*0.50, w*0.65, h*0.50], fill=color, width=4)

    if path:
        img.save(path, format="PNG")
    return img

def create_word_pdf_icon(path=None, size=(120, 120), color="#02293f"):
    """Icono de documento Word con 'W' y conversión a PDF"""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w, h = size

    # Hoja de documento base
    d.rounded_rectangle([w*0.20, h*0.14, w*0.80, h*0.86], radius=7, fill="#ffffff", outline=color, width=5)

    # Letra W estilizada
    pts = [
        (w*0.28, h*0.28),
        (w*0.35, h*0.54),
        (w*0.42, h*0.35),
        (w*0.49, h*0.54),
        (w*0.56, h*0.28)
    ]
    for i in range(len(pts)-1):
        d.line([pts[i], pts[i+1]], fill=color, width=4)

    # Flecha indicadora de conversión
    d.line([w*0.30, h*0.70, w*0.55, h*0.70], fill=color, width=4)
    d.polygon([(w*0.53, h*0.64), (w*0.65, h*0.70), (w*0.53, h*0.76)], fill=color)

    # Etiqueta 'PDF'
    d.rounded_rectangle([w*0.52, h*0.44, w*0.75, h*0.58], radius=4, fill=color)
    d.line([w*0.56, h*0.51, w*0.71, h*0.51], fill="#ffffff", width=2)

    if path:
        img.save(path, format="PNG")
    return img

def generate_all_icons(base_dir=None, save_to_disk=False):
    """
    Genera todos los iconos. Por defecto los crea en memoria (PIL Images)
    evitando crear carpetas o archivos en el disco del usuario.
    """
    icons = {
        "unir_pdf": create_unir_pdf_icon(),
        "pdf": create_unir_pdf_icon(),
        "doc_masivos": create_doc_masivos_icon(),
        "gene_documentos": create_doc_masivos_icon(),
        "formato_vinculacion": create_formato_icon(),
        "formato": create_formato_icon(),
        "vinculacion": create_formato_icon(),
        "compresor_pdf": create_compress_icon(),
        "compress": create_compress_icon(),
        "word_pdf": create_word_pdf_icon(),
        "word": create_word_pdf_icon(),
        "edit": create_edit_icon(),
        "sync": create_sync_icon(),
        "sheet": create_sheet_icon(),
        "excel": create_sheet_icon(),
        "backup": create_database_icon(),
        "database": create_database_icon(),
        "default_script": create_script_icon(),
        "script": create_script_icon(),
        "lock": create_lock_icon(),
    }

    if save_to_disk:
        icons_dir = ensure_icons_dir(base_dir)
        icons["unir_pdf"].save(os.path.join(icons_dir, "unir_pdf.png"))
        icons["doc_masivos"].save(os.path.join(icons_dir, "doc_masivos.png"))
        icons["formato_vinculacion"].save(os.path.join(icons_dir, "formato_vinculacion.png"))
        icons["compresor_pdf"].save(os.path.join(icons_dir, "compresor_pdf.png"))
        icons["word_pdf"].save(os.path.join(icons_dir, "word_pdf.png"))
        icons["edit"].save(os.path.join(icons_dir, "edit.png"))
        icons["sync"].save(os.path.join(icons_dir, "sync.png"))
        icons["sheet"].save(os.path.join(icons_dir, "sheet.png"))
        icons["backup"].save(os.path.join(icons_dir, "backup.png"))
        icons["default_script"].save(os.path.join(icons_dir, "default_script.png"))
        icons["lock"].save(os.path.join(icons_dir, "lock.png"))

    return icons

def get_icon_for_module(mod_name_or_script, icons_dict=None):
    """Devuelve la imagen PIL adecuada según el nombre o script del módulo"""
    if icons_dict is None:
        icons_dict = generate_all_icons(save_to_disk=False)

    text = str(mod_name_or_script).lower()
    if "comprim" in text or "compress" in text:
        return icons_dict.get("compresor_pdf", icons_dict["unir_pdf"])
    elif "vincula" in text or "formato" in text or "contrato" in text:
        return icons_dict.get("formato_vinculacion", icons_dict["doc_masivos"])
    elif "word" in text or "pds" in text:
        return icons_dict.get("word_pdf", icons_dict.get("doc_masivos", icons_dict["default_script"]))
    elif "unir" in text:
        return icons_dict.get("unir_pdf", icons_dict["default_script"])
    elif "doc" in text or "gene" in text:
        return icons_dict.get("doc_masivos", icons_dict["default_script"])
    elif "sheet" in text or "excel" in text or "tabla" in text:
        return icons_dict.get("sheet", icons_dict["default_script"])
    elif "backup" in text or "data" in text or "db" in text:
        return icons_dict.get("backup", icons_dict["default_script"])
    elif "edit" in text:
        return icons_dict.get("edit", icons_dict["default_script"])
    else:
        return icons_dict.get("default_script")

if __name__ == "__main__":
    generate_all_icons(save_to_disk=True)
    print("Iconos generados exitosamente.")
