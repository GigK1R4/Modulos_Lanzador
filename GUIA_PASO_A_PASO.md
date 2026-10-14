# 🚀 Guía Paso a Paso: Lanzador con Google Sheets y Repositorio de Módulos Git

Este sistema proporciona un **Lanzador de Escritorio moderno** (diseñado con fondo azul oscuro `#0b1626`, tarjetas blancas con bordes redondeados y efecto hover, píldora de estado cian), conectado en tiempo real a **Google Sheets** para validar usuarios y permisos, con soporte para sincronización y actualización automática mediante el repositorio de Git:
🔗 **`https://github.com/GigK1R4/Modulos_Lanzador.git`**

---

## 🧩 Principio Fundamental: Cada Código es un Módulo

Todos los módulos del programa se descargan y ejecutan directamente desde el repositorio de GitHub:
- Carpeta local: `Modulos_Lanzador/`
- **Cada archivo `.py` que exista o se agregue a ese repositorio es automáticamente detectado como un módulo.**
- Al iniciar el programa o pulsar **`⟳ Sincronizar Git`**, el lanzador consulta el repositorio en GitHub y descarga automáticamente cualquier script nuevo o cambio en los scripts existentes.
- Si se añade un script (por ejemplo `Generar_Reportes.py`), el lanzador lo detecta sin necesidad de modificar el código del lanzador.

---

## 📂 Archivos en la Carpeta `Lanzandor`

1. **`Lanzador.py`**: Aplicación principal con la interfaz gráfica en CustomTkinter (Login + Menú de módulos con tarjetas interactivas).
2. **`Modulos_Lanzador/`**: Carpeta donde reside el repositorio Git clonado con los scripts de los módulos:
   - `PDF_Unir.py`: Script para unificación y generación de certificados en PDF.
   - `Gene_Documentos.py`: Script para generación masiva de plantillas Word y conversión a PDF.
   - *(Cualquier otro script `.py` nuevo que se añada al repositorio).*
3. **`sheet_auth.py`**: Servicio que conecta con Google Sheets y verifica credenciales y permisos.
4. **`git_updater.py`**: Servicio que clona y actualiza automáticamente `Modulos_Lanzador` con `git pull`.
5. **`icon_generator.py`**: Generador automático de iconos en alta definición con Pillow para las tarjetas del menú.
6. **`config.json`**: Archivo de configuración central (URL de Google Sheet, URL del repositorio Git, etc.).
7. **`iniciar_lanzador.bat`**: Acceso directo para iniciar el programa con doble clic sin consola.
8. **`iniciar_lanzador_debug.bat`**: Acceso directo con consola para ver mensajes de diagnóstico.
9. **`GoogleAppsScript_Code.gs`**: Código backend para Google Apps Script Web App.

---

## 📋 PASO 1: Configurar el Google Sheet de Usuarios y Permisos

Debes tener una hoja de cálculo en Google Sheets que servirá como base de datos de usuarios.

### 1.1 Estructura de la Pestaña `Usuarios`
Crea una pestaña llamada **`Usuarios`** con la siguiente estructura de columnas:

| Usuario | Password | Modulos | Estado |
| :--- | :--- | :--- | :--- |
| **admin** | `1234` | `*` | `Activo` |
| **juan.perez** | `clave2026` | `PDF_Unir, Gene_Documentos` | `Activo` |
| **marta.lopez** | `abc456` | `PDF_Unir` | `Activo` |
| **carlos.ruiz** | `7890` | `Gene_Documentos` | `Activo` |
| **usuario.demo** | `demo123` | `PDF_Unir` | `Inactivo` |

### 1.2 Reglas de Permisos:
- **`*`**: Acceso a **todos** los módulos disponibles (actuales y nuevos).
- **Nombres por script o amigables**: Puedes colocar el nombre del script o su nombre amigable, separados por coma:
  - `PDF_Unir` (o `Unir pdf`)
  - `Gene_Documentos` (o `Documentos Masivos`)
  - El nombre de cualquier nuevo script `.py` que subas a GitHub.
- Si un usuario no tiene asignado un módulo, la tarjeta aparecerá **bloqueada** con un icono de candado.
- Si la columna `Estado` dice `Inactivo` o `Bloqueado`, no podrá iniciar sesión.

### 1.3 Compartir la Hoja para Lectura:
1. Haz clic en el botón verde **Compartir** (arriba a la derecha en Google Sheets).
2. En "Acceso general", cambia a: **Cualquier persona con el enlace** con rol de **Lector**.
3. Copia el enlace de la hoja (por ejemplo: `https://docs.google.com/spreadsheets/d/1v5u7t.../edit?usp=sharing`).

---

## ⚙️ PASO 2: Configuración en `config.json`

El archivo [`config.json`](file:///d:/CR-SISTEMAS-99o/Documents/ProgramasAntig/Lanzandor/config.json) ya contiene la configuración del repositorio y de la hoja:

```json
{
  "app_title": "AUTOMATIZACIÓN DE ARCHIVOS",
  "app_subtitle": "Seleccione una opción para continuar",
  "sheet_url": "https://docs.google.com/spreadsheets/d/1v5u7tBZ9x_Ple_r_Maf2x3U8R4r4XMB0YJ-tu68X5jc/edit?usp=sharing/edit",
  "sheet_name": "Usuarios",
  "apps_script_url": "",
  "modules_repo_url": "https://github.com/GigK1R4/Lanzador_Modulos.git",
  "modules_dir": "Modulos_Lanzador",
  "auto_git_update": true,
  "auto_discover_modules": true
}
```

---

## 🔄 PASO 3: Sincronización Automática con GitHub (Zero Carpetas Locales)

1. **Al iniciar el programa:**
   - Se conecta automáticamente a GitHub (`https://github.com/GigK1R4/Lanzador_Modulos.git`) para buscar actualizaciones de los módulos.
   - **Funciona en cualquier PC:** Si el equipo tiene Git instalado, utiliza `git fetch/pull`. Si no tiene Git instalado, descarga los módulos directamente vía HTTP/Zipball desde GitHub.
   - **Cero desorden:** Toda la información de módulos y credenciales se almacena en el directorio seguro del sistema (`%LOCALAPPDATA%\LanzadorApp`). **No se crea ninguna carpeta ni archivo en el escritorio o donde se ubique el ejecutable**.
   - Los iconos se generan 100% en memoria RAM (no se crea carpeta `assets`).
2. **Botón `⟳ Sincronizar Git`:**
   - Ubicado en la barra inferior derecha. Permite refrescar y descargar al instante cualquier script nuevo o cambio subido a GitHub sin reiniciar.
3. **Añadir o modificar módulos:**
   - Simplemente sube cualquier archivo `.py` nuevo a tu repositorio `https://github.com/GigK1R4/Lanzador_Modulos.git`.
   - Cuando cualquier usuario abra su `Lanzador.exe` en cualquier PC, se descargará automáticamente y aparecerá en su pantalla.

---

## ▶️ PASO 4: Iniciar el Lanzador y Portabilidad

Tienes las siguientes opciones sencillas para ejecutarlo:

### Opción A (Recomendada): Ejecutable directo en cualquier PC
- Haz doble clic en **[`Lanzador.exe`](file:///d:/CR-SISTEMAS-99o/Documents/ProgramasAntig/Lanzandor/Lanzador.exe)**.
- Es un ejecutable autónomo y portable. Puedes copiarlo al Escritorio, Descargas o una memoria USB en cualquier PC; funcionará de inmediato sin requerir configuración ni crear carpetas visibles.

### Opción B: Modo Desarrollo
- Ejecuta [`iniciar_lanzador.bat`](file:///d:/CR-SISTEMAS-99o/Documents/ProgramasAntig/Lanzandor/iniciar_lanzador.bat) o `python Lanzador.py`.
- Para diagnósticos detallados por consola, utiliza [`iniciar_lanzador_debug.bat`](file:///d:/CR-SISTEMAS-99o/Documents/ProgramasAntig/Lanzandor/iniciar_lanzador_debug.bat).

---

## 🚀 PASO 5: Subir el Proyecto Completo a GitHub

El proyecto ya está completamente limpio de archivos temporales y cuenta con su archivo [`.gitignore`](file:///d:/CR-SISTEMAS-99o/Documents/ProgramasAntig/Lanzandor/.gitignore):

```bash
git init
git add .
git commit -m "Lanzador modular con sincronización automática y ejecución portable"
git branch -M main
git remote add origin https://github.com/GigK1R4/Lanzador_Modulos.git
git push -u origin main
```
