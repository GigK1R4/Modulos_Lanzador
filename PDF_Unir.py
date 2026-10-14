import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert

def seleccionar_archivo():
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo",
        filetypes=(("Todos los archivos", "*.*"),
                   ("Archivos de texto", "*.txt"),
                   ("Archivos Excel", "*.xlsx;*.xls"))
    )
    if archivo:
        etiqueta.config(text=f"Archivo seleccionado:\n{archivo}")
        archivo_var.set(archivo)
        etiqueta.place(x=200 , y=80)
    else:
        messagebox.showwarning("Aviso", "No seleccionaste ningún archivo.")
    



def seleccionar_carpeta():
    carpeta = filedialog.askopenfilename(title="Seleccionar carpeta")
    
    if carpeta:
        etiqueta2.config(text=f"Carpeta seleccionada:\n{carpeta}")
        carpeta_var.set(carpeta)
    else:
        messagebox.showwarning("Aviso", "No seleccionaste ninguna carpeta.")


def certificados ():
    
    doc = DocxTemplate(archivo_var.get())
    df=pd.read_excel(carpeta_var.get())
    num_filas = len(df)

    for i in range(num_filas):
        context = {'Nombre':df.loc[i, 'Nombre'], 'CC':df.loc[i, 'Cedula'], 'Total':df.loc[i, 'Total'],
                'Efectivo':df.loc[i, 'Efectivo'],'Tranfe':df.loc[i, 'Tranferencia'],'Especie':df.loc[i, 'Especie'] ,
                'Letra':df.loc[i, 'Letra']}
        
        doc.render(context)
        doc.save(f"Certificado {df.loc[i, 'Nombre']}.docx")
        convert(f"Certificado {df.loc[i, 'Nombre']}.docx", f"Certificado {df.loc[i, 'Nombre']}.pdf")
        i += 1



# Crear ventana principal
ventana = tk.Tk()
ventana.title("Selector de Archivos y Carpetas")
ventana.geometry("500x500")
carpeta_var = tk.StringVar()
archivo_var = tk.StringVar()

# Botón para seleccionar archivo
boton_archivo = tk.Button(ventana, text="Seleccionar archivo", command=seleccionar_archivo)
boton_archivo.place(x=30 , y=100)

# Botón para seleccionar carpeta
boton_carpeta = tk.Button(ventana, text="Seleccionar carpeta", command=seleccionar_carpeta)
boton_carpeta.place(x=30 , y=150)

# Botón para Ejecutar Scrip
boton_certificados = tk.Button(ventana, text="Generar Certificado", command=certificados)
boton_certificados.place(x=150 , y=200)


# Etiqueta para mostrar la ruta
label_con_font = tk.Label(ventana, text="Generador De Certificados", font=("Bahnschrift SemiBold SemiConden", 20, "bold"),bg="lightblue",fg="red",cursor="hand2")
label_con_font.pack(pady=30) # Empaquetar la etiqueta con un poco de espacio vertical

etiqueta = tk.Label(ventana, text="No se ha seleccionado nada.", wraplength=500, justify="center")
etiqueta.place(x=200 , y=100)

etiqueta2 = tk.Label(ventana, text="No se ha seleccionado nada.", wraplength=500, justify="center")
etiqueta2.place(x=200 , y=150)
#Iniciar la app
ventana.mainloop()