"""
Módulo de actualización automática y descarga desde GitHub para el Lanzador.
Soporta:
1. Actualización nativa con Git (si git está instalado).
2. Descarga directa desde GitHub vía HTTP/Zipball con 'requests' (si git NO está instalado, ideal para cualquier PC).
3. Ubicación en AppData local para no crear carpetas en el escritorio o ruta del ejecutable.
"""
import os
import sys
import io
import json
import shutil
import zipfile
import tempfile
import subprocess
import threading

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
            elif self.is_git_repo(local_base):
                # Si la carpeta actual es un repositorio Git sin subcarpeta de módulos
                self.storage_dir = local_base
                self.modules_dir = local_base
            else:
                self.storage_dir = get_app_data_dir()
                self.modules_dir = os.path.join(self.storage_dir, modules_dir_name)

        self.repo_url = repo_url or "https://github.com/GigK1R4/Modulos_Lanzador.git"
        self.commit_cache_file = os.path.join(self.storage_dir, "last_commit.json")
        os.makedirs(self.modules_dir, exist_ok=True)

    def _has_git(self):
        """Comprueba si Git está disponible en la máquina"""
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            res = subprocess.run(["git", "--version"], capture_output=True, timeout=3, startupinfo=startupinfo)
            return res.returncode == 0
        except Exception:
            return False

    def _run_git_cmd(self, args, cwd=None, timeout=25):
        """Ejecuta un comando git de forma segura capturando la salida"""
        try:
            target_cwd = cwd or self.storage_dir
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            result = subprocess.run(
                ["git"] + args,
                cwd=target_cwd,
                capture_output=True,
                text=True,
                timeout=timeout,
                startupinfo=startupinfo
            )
            return result.returncode == 0, result.stdout.strip(), result.stderr.strip()
        except Exception as e:
            return False, "", str(e)

    def is_git_repo(self, path=None):
        target = path or self.modules_dir
        return os.path.isdir(os.path.join(target, ".git"))

    # =========================================================================
    # MÉTODO HTTP / GITHUB DIRECTO (Funciona en CUALQUIER PC sin Git instalado)
    # =========================================================================
    def _download_via_http(self, on_status_callback=None):
        """
        Descarga el repositorio como archivo ZIP desde GitHub y extrae los scripts en modules_dir.
        Funciona en cualquier ordenador sin necesidad de tener Git instalado.
        """
        try:
            import requests
            if on_status_callback:
                on_status_callback("Descargando módulos desde GitHub...")

            # Probar rama main, luego master
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
                return False, "No se pudo descargar el paquete desde GitHub."

            # Extraer archivos al directorio correspondiente
            target_extract = self.storage_dir if self.is_frozen else self.modules_dir
            with zipfile.ZipFile(io.BytesIO(zip_content)) as zf:
                for member in zf.infolist():
                    filename = member.filename
                    # Extraer solo archivos (evitando la carpeta contenedora inicial)
                    parts = filename.split("/", 1)
                    if len(parts) > 1 and parts[1]:
                        rel_path = parts[1]
                        if not rel_path.endswith("/"): # Es archivo
                            dest_path = os.path.join(target_extract, rel_path)
                            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                            with zf.open(member) as source, open(dest_path, "wb") as target:
                                shutil.copyfileobj(source, target)

            return True, "Módulos descargados exitosamente desde GitHub."
        except Exception as e:
            return False, f"Error en descarga directa: {e}"

    def check_http_updates(self):
        """Consulta si hay nuevos commits en GitHub a través de su API pública"""
        try:
            import requests
            api_url = "https://api.github.com/repos/GigK1R4/Modulos_Lanzador/commits/main"
            resp = requests.get(api_url, timeout=10)
            if resp.status_code != 200:
                api_url = "https://api.github.com/repos/GigK1R4/Modulos_Lanzador/commits/master"
                resp = requests.get(api_url, timeout=10)

            if resp.status_code == 200:
                data = resp.json()
                latest_sha = data.get("sha", "")
                
                # Comparar con el último commit guardado
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

        # Si está congelado en PyInstaller, los archivos pueden estar en sys._MEIPASS
        bundled_dirs = []
        if hasattr(sys, '_MEIPASS'):
            bundled_dirs.append(os.path.join(sys._MEIPASS, "Modulos_Lanzador"))
        local_dev = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Modulos_Lanzador")
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
    # MÉTODO PRINCIPAL DE ACTUALIZACIÓN
    # =========================================================================
    def update_all(self, on_status_callback=None):
        """
        Comprueba y actualiza los módulos de GitHub:
        1. Si tiene Git instalado y es repo git: usa git fetch/pull.
        2. Si no tiene Git o falla: usa descarga HTTP directa de GitHub.
        3. Si está sin conexión: asegura que existan los módulos embebidos de respaldo.
        """
        has_changes = False
        message = ""

        # Asegurar módulos iniciales de respaldo por si no hay conexión
        self.ensure_embedded_fallback()

        # En modo desarrollo dentro del propio repositorio Git, no sobreescribir ni clonar
        if not self.is_frozen and self.is_git_repo(self.storage_dir):
            if on_status_callback:
                on_status_callback("Repositorio local de desarrollo activo.")
            return False, "Modo desarrollo local activo (al día)."

        if self._has_git():
            # Flujo con Git
            if not self.is_git_repo(self.modules_dir):
                if on_status_callback:
                    on_status_callback("Conectando con GitHub (Git)...")
                # Si está vacío o no es git repo
                ok, _, err = self._run_git_cmd(["clone", self.repo_url, self.modules_dir], cwd=self.storage_dir, timeout=45)
                if ok:
                    has_changes = True
                    message = "[OK] Repositorio clonado al día desde GitHub."
                else:
                    # Fallback a HTTP
                    ok_http, msg_http = self._download_via_http(on_status_callback)
                    has_changes = ok_http
                    message = f"[OK] {msg_http}" if ok_http else f"Error: {err}"
            else:
                if on_status_callback:
                    on_status_callback("Buscando actualizaciones en GitHub...")
                ok_fetch, _, _ = self._run_git_cmd(["fetch", "--quiet"], cwd=self.modules_dir, timeout=15)
                if ok_fetch:
                    ok_rev, out_rev, _ = self._run_git_cmd(["rev-list", "HEAD..@{u}", "--count"], cwd=self.modules_dir, timeout=5)
                    if ok_rev and out_rev.isdigit() and int(out_rev) > 0:
                        if on_status_callback:
                            on_status_callback("Descargando cambios de GitHub...")
                        ok_pull, out_pull, _ = self._run_git_cmd(["pull", "--ff-only"], cwd=self.modules_dir, timeout=25)
                        if not ok_pull:
                            ok_pull, out_pull, _ = self._run_git_cmd(["pull"], cwd=self.modules_dir, timeout=25)
                        has_changes = ok_pull
                        message = "[OK] Módulos actualizados con los últimos cambios de GitHub."
                    else:
                        message = "[OK] Repositorio de módulos al día (GitHub)."
                else:
                    message = "[OK] Módulos cargados (modo local/sin conexión)."
        else:
            # Flujo sin Git (PC estándar sin herramientas de desarrollador)
            if on_status_callback:
                on_status_callback("Comprobando actualizaciones en GitHub...")

            has_up, latest_sha = self.check_http_updates()
            existing_scripts = [f for f in os.listdir(self.modules_dir) if f.endswith(".py")] if os.path.exists(self.modules_dir) else []

            if has_up or not existing_scripts:
                ok_http, msg_http = self._download_via_http(on_status_callback)
                if ok_http:
                    has_changes = True
                    if latest_sha:
                        self.save_last_commit(latest_sha)
                    message = "[OK] Módulos descargados y actualizados desde GitHub."
                else:
                    message = f"[OK] Módulos cargados ({len(existing_scripts)} disponibles)."
            else:
                message = "[OK] Repositorio de módulos al día (GitHub)."

        return has_changes, message

    def start_background_check(self, on_finished_callback=None, on_status_callback=None):
        """Ejecuta la comprobación en un hilo secundario sin congelar la interfaz"""
        def _thread_target():
            has_changes, msg = self.update_all(on_status_callback=on_status_callback)
            if on_finished_callback:
                on_finished_callback(has_changes, msg)

        t = threading.Thread(target=_thread_target, daemon=True)
        t.start()
        return t
