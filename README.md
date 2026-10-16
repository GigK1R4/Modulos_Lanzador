# 🚀 Automatización de Archivos - Lanzador Modular

[![Version](https://img.shields.io/badge/version-1.0.4-blue.svg)](version.json)
[![GUI](https://img.shields.io/badge/GUI-CustomTkinter-blueviolet.svg)](https://customtkinter.tomschimansky.com/)
[![Python](https://img.shields.io/badge/Python-3.10%2B-green.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![Security](https://img.shields.io/badge/Security-Google%20Apps%20Script%20%28Private%29-success.svg)](GUIA_PASO_A_PASO.md)

Sistema modular y portable de automatización de documentos y procesamiento de archivos con interfaz moderna basada en **CustomTkinter**, autenticación segura contra **Google Sheets (100% Privado)**, control de permisos por roles, auto-actualizador integral de módulos y del ejecutable vía **GitHub**, y ventana única integrada.

---

## 📋 Tabla de Contenidos

1. [Novedades de la Versión 1.0.4](#-novedades-de-la-versión-104)
2. [Características Principales](#-características-principales)
3. [Módulos del Sistema](#-módulos-del-sistema)
4. [Arquitectura de Seguridad y Permisos](#-arquitectura-de-seguridad-y-permisos)
5. [Sistema de Auto-Actualización Integral](#-sistema-de-auto-actualización-integral)
6. [Estructura del Proyecto](#-estructura-del-proyecto)
7. [Instalación y Uso](#-instalación-y-uso)
8. [Configuración](#-configuración)
9. [Tecnologías Utilizadas](#-tecnologías-utilizadas)

---

## 🌟 Novedades de la Versión 1.0.4

- **Catálogo Completo de 5 Módulos Integrados:** Incorporación y consolidación de los módulos en la subcarpeta `Modulos/`, todos con soporte de ventana única y tema moderno CustomTkinter.
- **Nuevo Módulo "Word a PDF" (`Documento Word a PDF.py`):** Conversión por lotes de archivos Word a PDF, normalización estandarizada a tamaño Carta (8.5 × 11 in) con escalado proporcional, y opción de unificación en un solo documento PDF final.
- **Módulo "Comprimir PDF" Modernizado (`Compresor_pdf.py`):** Interfaz renovada con CustomTkinter, detección inteligente de Ghostscript en el sistema, preajustes de calidad (Ebook, Pantalla, Impresión, Preprensa) y visualización detallada del ahorro de peso y porcentaje en tiempo real.
- **Módulo "Formato Vinculación" (`FORMATO VINCULACION.py`):** Generación masiva automatizada de formatos FV combinando datos de planillas Excel con exportación directa a PDF mediante automatización COM de Microsoft Office.
- **Auto-Actualizador Integral de `Lanzador.exe`:**
  - Descarga e instalación automática de nuevas versiones del ejecutable desde GitHub.
  - Comparación de versiones semánticas basada en `version.json` (evita bucles o descargas innecesarias).
  - Sustitución silenciosa en segundo plano mediante script batch invisible sin ventanas negras de consola (CMD).
  - Confirmación interactiva al usuario antes de aplicar el reinicio con la nueva versión.
- **Identidad Visual y Experiencia de Usuario (UI/UX):**
  - Fondo ambiental translúcido integrado con la imagen corporativa `Logo_APP2.jpeg`.
  - Icono nativo de la aplicación integrado (`Logo_app.ico`) tanto en ventana como en el archivo ejecutable.
  - Panel flotante de configuración rápida de API (icono ⚙) con herramienta de prueba de conexión en vivo con Google Apps Script.

---

## ✨ Características Principales

- **Ejecución 100% Portable (Zero Clutter):** El ejecutable `Lanzador.exe` se puede ejecutar desde cualquier carpeta, Escritorio o memoria USB sin ensuciar el entorno de trabajo del usuario.
- **Ventana Única Integrada:** Todos los módulos se visualizan y operan dentro de la misma ventana principal, con adaptación suave de dimensiones y botón de retorno instantáneo ("Volver al Menú").
- **Doble Modo de Sincronización:** Si la máquina dispone de Git, sincroniza a través de Git CLI; en computadores sin Git, utiliza descarga directa y resiliente vía la API REST de GitHub / HTTP.
- **Autenticación Segura (Zero Exposición):** La base de datos de usuarios y permisos en Google Sheets permanece 100% privada (acceso restringido). La aplicación solo se comunica vía HTTPS contra una Web App de Google Apps Script.
- **Control de Acceso Granular:** Soporta permisos por módulo individual, comodín total (`*`), validación de estado de cuenta (Activo/Inactivo) y hashing de contraseñas mediante SHA-256.
- **Iconos Vectoriales Dinámicos en RAM:** Toda la iconografía se renderiza directamente en memoria con Pillow, garantizando alta resolución y sin generar residuos en disco.

---

## 📦 Módulos del Sistema

El sistema cuenta con **5 módulos especializados** ubicados en la carpeta [`Modulos/`](Modulos/):

| Módulo | Archivo Script | Tarjeta en Menú | Descripción Funcional |
| :--- | :--- | :--- | :--- |
| **1. Unir PDF** | `PDF_Unir.py` | `UNIR PDF` | Fusión organizada de múltiples documentos PDF, reordenamiento interactivo (subir/bajar), previsualización, ajuste de páginas a tamaño carta y generación unificada. |
| **2. Documentos Masivos** | `Gene_Documentos.py` | `DOCUMENTOS MASIVOS` | Creación automatizada en lote de documentos Word a partir de plantillas `.docx` combinadas con registros de bases de datos Excel (`.xlsx`), con conversión opcional a PDF. |
| **3. Formato Vinculación** | `FORMATO VINCULACION.py` | `FORMATO VINCULACIÓN` | Generación masiva automatizada de formatos de vinculación (FV) de personal tomando información de hojas de cálculo Excel y exportando directamente a PDF con Microsoft Excel COM. |
| **4. Comprimir PDF** | `Compresor_pdf.py` | `COMPRIMIR PDF` | Optimización avanzada del peso de archivos PDF utilizando el motor Ghostscript con perfiles ajustables (Ebook 150 DPI, Pantalla 72 DPI, Impresión 300 DPI, Preprensa) y reporte de ahorro. |
| **5. Word a PDF** | `Documento Word a PDF.py` | `WORD A PDF` | Conversión rápida por lotes de documentos `.docx` y `.doc` a PDF, normalizando cada página a tamaño Carta estándar (8.5 × 11 in) con escalado proporcional y opción de unificación. |

> 💡 **Arquitectura Abierta:** Para registrar un nuevo módulo, basta con añadir el archivo `.py` en la carpeta `Modulos/` o en el repositorio. El sistema lo detectará automáticamente y creará su tarjeta en el menú principal.

---

## 🛡️ Arquitectura de Seguridad y Permisos

La autenticación no descarga listas de usuarios ni expone las hojas de cálculo a los clientes finales:

```
┌─────────────────────┐         HTTPS POST          ┌─────────────────────────────┐
│    Lanzador.py      │ ──────────────────────────> │   Google Apps Script API    │
│  (Cliente Usuario)  │ <────────────────────────── │        (Web App /exec)      │
└─────────────────────┘      JSON Token + Permisos  └──────────────┬──────────────┘
                                                                   │ Acceso Privado
                                                                   ▼
                                                    ┌─────────────────────────────┐
                                                    │   Google Sheets (Privada)   │
                                                    │  (Columnas: Usuario,        │
                                                    │   Password, Modulos, Estado)│
                                                    └─────────────────────────────┘
```

1. **Hoja 100% Restringida:** Solo el administrador de Google Drive tiene acceso a la hoja de cálculo.
2. **Cero Exposición de Datos:** El cliente solo envía el usuario y la clave ingresada; el backend valida y responde con los módulos habilitados.
3. **Cifrado SHA-256:** Compatible tanto con contraseñas en texto plano como cifradas mediante hash SHA-256 (`=SHA256("contraseña")`).
4. Para más detalles sobre cómo desplegar el backend, consulta la [Guía Paso a Paso](GUIA_PASO_A_PASO.md).

---

## 🔄 Sistema de Auto-Actualización Integral

El actualizador central (`git_updater.py`) se encarga de mantener al día tanto el código como los ejecutables:

1. **Auto-actualización de Módulos:** Sincroniza en cada inicio las últimas versiones de los scripts en `Modulos/` y recursos visuales.
2. **Auto-actualización de `Lanzador.exe`:**
   - Comprueba el archivo remoto `version.json` contra la versión local.
   - Si existe una versión superior, descarga el nuevo binario en segundo plano.
   - Pregunta al usuario si desea actualizar en ese momento.
   - Aplica el cambio mediante un script batch silencioso que espera el cierre del ejecutable anterior, reemplaza el archivo y relanza la aplicación sin ventanas emergentes de consola.
3. **Operación Offline / Fallback Resiliente:** Si no hay conexión o falla GitHub, la aplicación continúa ejecutándose con los módulos y configuraciones locales existentes.

---

## 📁 Estructura del Proyecto

```text
Lanzador/
│
├── Modulos/                              # Directorio de módulos funcionales
│   ├── Compresor_pdf.py                  # Módulo 4: Compresión PDF con Ghostscript
│   ├── Documento Word a PDF.py           # Módulo 5: Conversión y unión Word a PDF Carta
│   ├── FORMATO VINCULACION.py            # Módulo 3: Formato de vinculación masivo
│   ├── Gene_Documentos.py                # Módulo 2: Generación masiva de documentos Word
│   ├── PDF_Unir.py                       # Módulo 1: Unificación y ajuste de PDFs
│   └── __init__.py
│
├── Lanzador.py                           # Aplicación principal y contenedor de GUI
├── Lanzador.exe                          # Ejecutable empaquetado autónomo para Windows
├── git_updater.py                        # Motor de actualización y sincronización
├── sheet_auth.py                         # Conector de autenticación con Google Apps Script
├── icon_generator.py                     # Generador vectorial de iconos en memoria
│
├── config.json                           # Configuración central (títulos, temas, módulos, URLs)
├── version.json                          # Metadatos de versión y enlace del ejecutable
├── requirements.txt                      # Dependencias de Python necesarias
│
├── Logo_APP2.jpeg                        # Fondo ambiental corporativo
├── Logo_app.ico                          # Icono oficial de la aplicación
├── GoogleAppsScript_Code.gs              # Código backend para Google Apps Script
├── GUIA_PASO_A_PASO.md                   # Documentación detallada de configuración de seguridad
├── iniciar_lanzador.bat                  # Script de ejecución rápida para entornos Python
└── iniciar_lanzador_debug.bat            # Script con consola para depuración
```

---

## 🚀 Instalación y Uso

### Opción 1: Ejecutable para Usuarios Finales (Recomendado)

1. Descarga el archivo compilado **`Lanzador.exe`** desde este repositorio o de los Releases.
2. Haz doble clic sobre **`Lanzador.exe`**.
3. Inicia sesión con tus credenciales asignadas en Google Sheets.
4. Selecciona el módulo con el que deseas trabajar.

### Opción 2: Modo Desarrollador (Python)

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/GigK1R4/Modulos_Lanzador.git
   cd Modulos_Lanzador
   ```

2. **Crear y activar un entorno virtual (recomendado):**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Ejecutar el Lanzador:**
   ```bash
   python Lanzador.py
   ```
   *O alternativamente haz doble clic en `iniciar_lanzador.bat`.*

---

## ⚙️ Configuración

El archivo [`config.json`](config.json) permite ajustar el comportamiento del sistema sin tocar código:

```json
{
  "app_title": "AUTOMATIZACIÓN DE ARCHIVOS",
  "app_subtitle": "Seleccione una opción para continuar",
  "apps_script_url": "https://script.google.com/macros/s/TU_SCRIPT_ID/exec",
  "modules_repo_url": "https://github.com/GigK1R4/Modulos_Lanzador.git",
  "modules_dir": "Modulos",
  "auto_git_update": true,
  "auto_discover_modules": true,
  "network_timeout": 10
}
```

- **`apps_script_url`**: URL de la Web App desplegada en Google Apps Script.
- **`modules_repo_url`**: Repositorio de GitHub para sincronización automática de módulos.
- **`auto_git_update`**: `true` para comprobar actualizaciones al iniciar la aplicación.
- **`auto_discover_modules`**: `true` para escanear y cargar módulos nuevos presentes en `Modulos/`.

---

## 🛠️ Tecnologías Utilizadas

- **Interfaz Gráfica:** [CustomTkinter](https://customtkinter.tomschimansky.com/) (diseño moderno, modo oscuro, widgets adaptativos).
- **Procesamiento Gráfico:** [Pillow (PIL)](https://pillow.readthedocs.io/) (renderizado de iconos vectoriales en memoria RAM y filtros ambientales).
- **Manipulación de PDFs:** [pypdf](https://pypdf.readthedocs.io/) & [ReportLab](https://www.reportlab.com/) (unión, escalado, transformación y dibujo de certificados).
- **Compresión de PDFs:** [Ghostscript](https://www.ghostscript.com/) (optimización y reducción de peso con perfiles DPI).
- **Procesamiento de Documentos:** [pandas](https://pandas.pydata.org/), [openpyxl](https://openpyxl.readthedocs.io/), [docxtpl](https://docxtpl.readthedocs.io/), [python-docx](https://python-docx.readthedocs.io/).
- **Automatización de Microsoft Office en Windows:** `pywin32` (Word y Excel COM API).
- **Backend de Autenticación:** [Google Apps Script](https://developers.google.com/apps-script) (API REST HTTPS privada sobre Google Sheets).
- **Control de Versiones y Distribución:** Git CLI + GitHub REST API.

---

Desarrollado con ❤️ por **DKSoftware**
