"""
=============================================================================
MÓDULO DE ACTUALIZACIÓN INTEGRAL (GIT / HTTP / AUTO-UPDATE DE EJECUTABLE)
=============================================================================
Soporta:
1. Sincronización completa de módulos y recursos (Modulos/*.py, Logo_APP2, etc.)
2. Auto-actualización transparente de Lanzador.exe desde GitHub:
   - Descarga la nueva versión en segundo plano.
   - Crea script batch de reemplazo seguro para evitar bloqueo de Windows.
   - Reinicia automáticamente con la versión más reciente.
3. Detección rápida y confiable tanto con Git como vía HTTP directa (sin Git).
4. No deja la interfaz bloqueada ni congelada en ningún estado.
=============================================================================
"""
import os
import sys
import io
import json
import shutil
import zipfile
import subprocess
import threading
import requests

CURRENT_VERSION = "1.0.3"
REPO_URL = "https://github.com/GigK1R4/Modulos_Lanzador.git"
RAW_EXE_URL = "https://raw.githubusercontent.com/GigK1R4/Modulos_Lanzador/main/Lanzador.exe"
RAW_VERSION_URL = "https://raw.githubusercontent.com/GigK1R4/Modulos_Lanzador/main/version.json"


def get_app_data_dir():
    """
    Retorna la ruta segura de almacenamiento de la aplicación en el sistema del usuario.
    En Windows utiliza %LOCALAPPDATA%/LanzadorApp para evitar ensuciar el escritorio o carpeta de ejecución.
    """
    if os.name == 'nt':
        base = os.environ.get("LOCALAPPDATA")
        if not base:
            base = os.path.join(os.path.expanduser("~"), "AppData", "Local")
    else:
        base = os.path.expanduser("~/.local/share")

    app_dir = os.path.join(base, "LanzadorApp")
    os.makedirs(app_dir, exist_ok=True)
    return app_dir


class GitUpdater:
    def __init__(self, base_dir=None, repo_url=None, modules_dir_name="Modulos"):
        self.is_frozen = getattr(sys, 'frozen', False)
        self.current_version = CURRENT_VERSION

        # Si está congelado (.exe), los módulos se guardan en AppData local para no crear carpetas visibles
        if self.is_frozen:
            self.storage_dir = get_app_data_dir()
            self.modules_dir = os.path.join(self.storage_dir, modules_dir_name)
        else:
            # En ejecución como script .py:
            local_base = base_dir or os.path.dirname(os.path.abspath(__file__))
            local_mods = os.path.join(local_base, modules_dir_name)
            fallback_mods = os.path.join(local_base, "Modulos")
            if os.path.isdir(local_mods):
                self.storage_dir = local_base
                self.modules_dir = local_mods
            elif os.path.isdir(fallback_mods):
                self.storage_dir = local_base
                self.modules_dir = fallback_mods
            else:
                self.storage_dir = local_base
                self.modules_dir = os.path.join(self.storage_dir, modules_dir_name)

        self.repo_url = repo_url or REPO_URL
        self.commit_cache_file = os.path.join(self.storage_dir, "last_commit.json")
        os.makedirs(self.modules_dir, exist_ok=True)

    def _has_git(self):
        """Comprueba si Git está disponible en la máquina"""
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            env = os.environ.copy()
            env["GIT_TERMINAL_PROMPT"] = "0"
            res = subprocess.run(["git", "--version"], capture_output=True, timeout=3, startupinfo=startupinfo, env=env)
            return res.returncode == 0
        except Exception:
            return False

    def _run_git_cmd(self, args, cwd=None, timeout=20):
        """Ejecuta un comando git de forma segura evitando prompts de terminal"""
        try:
            target_cwd = cwd or self.storage_dir
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            env = os.environ.copy()
            env["GIT_TERMINAL_PROMPT"] = "0"

            result = subprocess.run(
                ["git"] + args,
                cwd=target_cwd,
                capture_output=True,
                text=True,
                timeout=timeout,
                startupinfo=startupinfo,
                env=env
            )
            return result.returncode == 0, result.stdout.strip(), result.stderr.strip()
        except Exception as e:
            return False, "", str(e)

    def is_git_repo(self, path=None):
        target = path or self.modules_dir
        return os.path.isdir(os.path.join(target, ".git"))

    # =========================================================================
    # AUTO-ACTUALIZADOR DE EJECUTABLE (Lanzador.exe)
    # =========================================================================
    def check_exe_update_needed(self):
        """
        Comprueba de forma rápida si existe una versión más reciente de Lanzador.exe en GitHub.
        Compara tamaño de archivo (Content-Length) y versión en version.json.
        """
        try:
            target_exe = sys.executable if self.is_frozen else os.path.join(self.storage_dir, "Lanzador.exe")
            if not os.path.isfile(target_exe):
                return False, 0, ""

            current_size = os.path.getsize(target_exe)

            # 1. Comprobar Content-Length remoto del ejecutable
            head_resp = requests.head(RAW_EXE_URL, timeout=8)
            if head_resp.status_code == 200:
                remote_size = int(head_resp.headers.get("Content-Length", 0))
                # Si el archivo remoto tiene tamaño válido y difiere del local
                if remote_size > 15_000_000 and abs(remote_size - current_size) > 1024:
                    return True, remote_size, "Actualización disponible por tamaño de ejecutable"

            # 2. Comprobar version.json
            v_resp = requests.get(RAW_VERSION_URL, timeout=8)
            if v_resp.status_code == 200:
                v_data = v_resp.json()
                remote_ver = v_data.get("version", "")
                if remote_ver and remote_ver != self.current_version:
                    return True, 0, f"Nueva versión {remote_ver} disponible"

        except Exception as e:
            print(f"[GitUpdater] Error al comprobar actualización de ejecutable: {e}")

        return False, 0, ""

    def download_and_prepare_exe_update(self, on_status_callback=None):
        """
        Descarga la nueva versión de Lanzador.exe desde GitHub y genera el script de reinicio batch.
        Retorna (exito, nuevo_exe_path, bat_path, mensaje).
        """
        try:
            if on_status_callback:
                on_status_callback("Descargando actualización del Lanzador (0%)...")

            update_dir = get_app_data_dir()
            new_exe_path = os.path.join(update_dir, "Lanzador_update.exe")

            # Descarga por fragmentos con reporte de porcentaje
            resp = requests.get(RAW_EXE_URL, stream=True, timeout=60)
            if resp.status_code != 200:
                return False, None, None, f"Error HTTP {resp.status_code} al descargar ejecutable."

            total_size = int(resp.headers.get("Content-Length", 0))
            downloaded = 0

            with open(new_exe_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=262144): # 256KB
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_size > 0 and on_status_callback:
                            pct = int((downloaded / total_size) * 100)
                            on_status_callback(f"Descargando actualización del Lanzador ({pct}%)...")

            # Validación del ejecutable descargado
            if not os.path.isfile(new_exe_path) or os.path.getsize(new_exe_path) < 15_000_000:
                try:
                    os.remove(new_exe_path)
                except Exception:
                    pass
                return False, None, None, "La descarga del ejecutable está incompleta o corrupta."

            with open(new_exe_path, "rb") as f:
                header = f.read(2)
                if header != b'MZ':
                    try:
                        os.remove(new_exe_path)
                    except Exception:
                        pass
                    return False, None, None, "El archivo descargado no es un ejecutable PE válido."

            # Crear script batch de reemplazo seguro en Windows
            bat_path = os.path.join(update_dir, "apply_update.bat")
            bat_content = """@echo off
setlocal enabledelayedexpansion
set "PID=%~1"
set "NEW_EXE=%~2"
set "TARGET_EXE=%~3"

:: Esperar a que el proceso anterior cierre
set /a attempts=0
:WAIT_PID
tasklist /fi "PID eq !PID!" 2>nul | findstr /i "!PID!" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    set /a attempts+=1
    if !attempts! leq 35 goto WAIT_PID
)
timeout /t 1 /nobreak >nul

:: Reemplazar el ejecutable antiguo por el nuevo
copy /y "!NEW_EXE!" "!TARGET_EXE!" >nul
if errorlevel 1 (
    del /f /q "!TARGET_EXE!.old" 2>nul
    move /y "!TARGET_EXE!" "!TARGET_EXE!.old" >nul
    copy /y "!NEW_EXE!" "!TARGET_EXE!" >nul
)

:: Limpiar archivo temporal de descarga
del /f /q "!NEW_EXE!" 2>nul

:: Iniciar el ejecutable actualizado
start "" "!TARGET_EXE!"

:: Eliminar este script batch
(goto) 2>nul & del /f /q "%~f0"
"""
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(bat_content)

            return True, new_exe_path, bat_path, "Actualización lista para aplicar."

        except Exception as e:
            return False, None, None, f"Error al preparar actualización: {e}"

    # =========================================================================
    # MÉTODO HTTP / GITHUB DIRECTO PARA MÓDULOS Y RECURSOS
    # =========================================================================
    def _download_via_http(self, on_status_callback=None):
        """
        Descarga el repositorio como archivo ZIP desde GitHub y sincroniza:
        - Archivos de módulos en self.modules_dir
        - Recursos (Logo_APP2, config, etc.) en self.storage_dir
        """
        try:
            if on_status_callback:
                on_status_callback("Descargando módulos desde GitHub...")

            branches = ["main", "master"]
            zip_content = None
            for branch in branches:
                zip_url = f"https://github.com/GigK1R4/Modulos_Lanzador/archive/refs/heads/{branch}.zip"
                try:
                    resp = requests.get(zip_url, timeout=20)
                    if resp.status_code == 200:
                        zip_content = resp.content
                        break
                except Exception:
                    continue

            if not zip_content:
                return False, "No se pudo conectar con GitHub."

            # Extraer archivos al directorio correspondiente
            with zipfile.ZipFile(io.BytesIO(zip_content)) as zf:
                for member in zf.infolist():
                    filename = member.filename
                    parts = filename.split("/", 1)
                    if len(parts) > 1 and parts[1]:
                        rel_path = parts[1]
                        if not rel_path.endswith("/"):
                            # 1. Guardar en storage_dir
                            dest_storage = os.path.join(self.storage_dir, rel_path)
                            os.makedirs(os.path.dirname(dest_storage), exist_ok=True)
                            with zf.open(member) as source, open(dest_storage, "wb") as target:
                                shutil.copyfileobj(source, target)

                            # 2. Si es archivo de módulo, asegurar que esté en self.modules_dir
                            if rel_path.startswith("Modulos/"):
                                sub_name = rel_path.split("Modulos/", 1)[1]
                                if sub_name:
                                    dest_mod = os.path.join(self.modules_dir, sub_name)
                                    os.makedirs(os.path.dirname(dest_mod), exist_ok=True)
                                    shutil.copy2(dest_storage, dest_mod)

            return True, "Módulos y recursos sincronizados desde GitHub."
        except Exception as e:
            return False, f"Error en descarga directa: {e}"

    def check_http_updates(self):
        """Consulta si hay nuevos commits en GitHub a través de su API pública"""
        try:
            api_url = "https://api.github.com/repos/GigK1R4/Modulos_Lanzador/commits/main"
            resp = requests.get(api_url, timeout=8)
            if resp.status_code != 200:
                api_url = "https://api.github.com/repos/GigK1R4/Modulos_Lanzador/commits/master"
                resp = requests.get(api_url, timeout=8)

            if resp.status_code == 200:
                data = resp.json()
                latest_sha = data.get("sha", "")
                saved_sha = ""
                if os.path.exists(self.commit_cache_file):
                    try:
                        with open(self.commit_cache_file, "r", encoding="utf-8") as f:
                            saved_sha = json.load(f).get("sha", "")
                    except Exception:
                        pass

                if latest_sha and latest_sha != saved_sha:
                    return True, latest_sha
                return False, latest_sha
        except Exception:
            pass
        return False, ""

    def save_last_commit(self, sha):
        try:
            with open(self.commit_cache_file, "w", encoding="utf-8") as f:
                json.dump({"sha": sha}, f)
        except Exception:
            pass

    # =========================================================================
    # COPIA DE SEGURIDAD / ASSETS EMBEBIDOS
    # =========================================================================
    def ensure_embedded_fallback(self):
        """Si la carpeta de módulos está vacía, copia los scripts integrados si existen"""
        existing = [f for f in os.listdir(self.modules_dir) if f.endswith(".py")] if os.path.exists(self.modules_dir) else []
        if existing:
            return

        bundled_dirs = []
        if hasattr(sys, '_MEIPASS'):
            bundled_dirs.append(os.path.join(sys._MEIPASS, "Modulos"))
            bundled_dirs.append(os.path.join(sys._MEIPASS, "Modulos_Lanzador"))
        local_dev = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Modulos")
        bundled_dirs.append(local_dev)

        for bdir in bundled_dirs:
            if os.path.isdir(bdir):
                try:
                    for f in os.listdir(bdir):
                        if f.endswith(".py"):
                            shutil.copy2(os.path.join(bdir, f), os.path.join(self.modules_dir, f))
                    break
                except Exception:
                    pass

    # =========================================================================
    # MÉTODO PRINCIPAL DE ACTUALIZACIÓN INTEGRAL
    # =========================================================================
    def update_all(self, on_status_callback=None):
        """
        Comprueba y actualiza tanto los módulos como el ejecutable Lanzador.exe:
        Retorna dict con resultados detallados.
        """
        result = {
            "has_changes": False,
            "message": "",
            "exe_update_ready": False,
            "new_exe_path": None,
            "bat_path": None
        }

        # Asegurar módulos iniciales de respaldo por si no hay conexión
        self.ensure_embedded_fallback()

        # En modo script python local (desarrollo), solo hacer git pull si es repo
        if not self.is_frozen and self.is_git_repo(self.storage_dir):
            if on_status_callback:
                on_status_callback("Actualizando repositorio local...")
            ok_pull, _, _ = self._run_git_cmd(["pull", "--ff-only"], cwd=self.storage_dir, timeout=15)
            result["has_changes"] = ok_pull
            result["message"] = "[OK] Repositorio local sincronizado." if ok_pull else "[OK] Repositorio local al día."
            return result

        # 1. Comprobar si hay actualización del ejecutable Lanzador.exe
        if self.is_frozen:
            if on_status_callback:
                on_status_callback("Comprobando versión del ejecutable...")
            needs_exe_up, _, _ = self.check_exe_update_needed()
            if needs_exe_up:
                ok_down, new_exe, bat_file, down_msg = self.download_and_prepare_exe_update(on_status_callback=on_status_callback)
                if ok_down:
                    result["has_changes"] = True
                    result["exe_update_ready"] = True
                    result["new_exe_path"] = new_exe
                    result["bat_path"] = bat_file
                    result["message"] = "[OK] Nueva versión de Lanzador.exe lista para reiniciar."
                    return result

        # 2. Sincronizar módulos (Git o HTTP)
        if self._has_git():
            if not self.is_git_repo(self.modules_dir):
                if on_status_callback:
                    on_status_callback("Sincronizando módulos desde GitHub...")
                ok_http, msg_http = self._download_via_http(on_status_callback)
                result["has_changes"] = ok_http
                result["message"] = f"[OK] {msg_http}" if ok_http else "Módulos cargados (modo local)."
            else:
                if on_status_callback:
                    on_status_callback("Buscando cambios en GitHub...")
                ok_fetch, _, _ = self._run_git_cmd(["fetch", "--quiet"], cwd=self.modules_dir, timeout=12)
                if ok_fetch:
                    ok_rev, out_rev, _ = self._run_git_cmd(["rev-list", "HEAD..@{u}", "--count"], cwd=self.modules_dir, timeout=5)
                    if ok_rev and out_rev.isdigit() and int(out_rev) > 0:
                        if on_status_callback:
                            on_status_callback("Descargando cambios de módulos...")
                        ok_pull, _, _ = self._run_git_cmd(["pull", "--ff-only"], cwd=self.modules_dir, timeout=15)
                        result["has_changes"] = ok_pull
                        result["message"] = "[OK] Módulos actualizados con GitHub."
                    else:
                        result["message"] = "[OK] Módulos y sistema al día."
                else:
                    result["message"] = "[OK] Módulos locales activos."
        else:
            # Sin Git instalado
            if on_status_callback:
                on_status_callback("Comprobando actualizaciones en GitHub...")

            has_up, latest_sha = self.check_http_updates()
            existing_scripts = [f for f in os.listdir(self.modules_dir) if f.endswith(".py")] if os.path.exists(self.modules_dir) else []

            if has_up or not existing_scripts:
                ok_http, msg_http = self._download_via_http(on_status_callback)
                if ok_http:
                    result["has_changes"] = True
                    if latest_sha:
                        self.save_last_commit(latest_sha)
                    result["message"] = "[OK] Módulos actualizados desde GitHub."
                else:
                    result["message"] = f"[OK] Módulos locales activos ({len(existing_scripts)} disponibles)."
            else:
                result["message"] = "[OK] Módulos y sistema al día."

        return result
