# 🚀 Automatización de Archivos - Lanzador Modular

Sistema modular de automatización de archivos con interfaz moderna (**CustomTkinter**), autenticación contra **Google Sheets**, permisos por roles y sincronización automática desde **GitHub**.

---

## 🌟 Características Principales

- **Ejecución 100% Portable (Zero Clutter):** El ejecutable `Lanzador.exe` se puede ejecutar desde cualquier carpeta, Escritorio o memoria USB sin crear archivos ni carpetas adicionales en la máquina del usuario.
- **Sincronización Automática con GitHub:** Cada vez que se abre el lanzador, consulta este repositorio para descargar nuevos módulos o actualizaciones. Funciona incluso en computadores que no tienen Git instalado (descarga directa vía HTTP).
- **Ventana Única Integrada:** Los módulos no se abren en ventanas emergentes externas, sino que se integran y adaptan suavemente dentro de la misma ventana principal.
- **Control de Permisos Dinámico:** Consulta en tiempo real las credenciales y los permisos asignados a cada usuario desde una hoja de Google Sheets.
- **Iconos en Memoria:** Los gráficos vectoriales de los módulos se renderizan directamente en la memoria RAM con Pillow.

---

## 📦 Módulos Incluidos

1. **UNIR PDF:** Fusión organizada de múltiples archivos PDF, previsualización, ordenamiento dinámico y generación de certificados.
2. **DOCUMENTOS MASIVOS:** Generación masiva de documentos Word a partir de plantillas y bases de datos en Excel, con conversión automatizada a PDF.

> **💡 Para añadir un nuevo módulo:** Simplemente sube cualquier script `.py` nuevo a este repositorio. El lanzador lo detectará y creará su tarjeta automáticamente.

---

## ▶️ Uso Rápido

### Opción 1: Ejecutable directo (Recomendado para usuarios)
1. Descarga y haz doble clic en **`Lanzador.exe`**.
2. Ingresa tu usuario y contraseña registrados en la hoja de Google Sheets.

### Opción 2: Modo Desarrollador (Python)
1. Clona el repositorio:
   ```bash
   git clone https://github.com/GigK1R4/Modulos_Lanzador.git
   ```
2. Instala las dependencias:
   ```bash
   pip install customtkinter pillow requests pypdf reportlab pandas openpyxl docxtpl
   ```
3. Inicia la aplicación:
   ```bash
   python Lanzador.py
   ```

---

## 🛠️ Tecnologías

- **GUI:** [CustomTkinter](https://customtkinter.tomschimansky.com/) (Tema oscuro minimalista)
- **Gráficos:** Pillow (PIL)
- **Autenticación:** Google Sheets CSV API & Google Apps Script
- **Sincronización:** Git CLI + GitHub REST API (Fallback HTTP)

---

Desarrollado con ❤️ por **DKSoftware**
