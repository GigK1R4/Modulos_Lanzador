"""
Módulo de autenticación y lectura de permisos desde Google Sheets.
Soporta:
1. Google Sheets público/compartido mediante descarga CSV directa en tiempo real.
2. Google Apps Script Web App (API JSON).
"""
import io
import re
import csv
import json
import urllib.parse
import requests

class SheetAuth:
    def __init__(self, config):
        self.config = config
        self.sheet_url = config.get("sheet_url", "")
        self.sheet_name = config.get("sheet_name", "Usuarios")
        self.api_url = config.get("apps_script_url", "")
        self.timeout = config.get("network_timeout", 10)

    @staticmethod
    def extract_sheet_id(url):
        """Extrae el ID del Google Sheet desde cualquier formato de enlace"""
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
        if match:
            return match.group(1)
        return url.strip()

    def _get_csv_export_url(self):
        """Construye la URL de exportación CSV directa de Google Sheets"""
        sheet_id = self.extract_sheet_id(self.sheet_url)
        quoted_name = urllib.parse.quote(self.sheet_name)
        # Formato gviz/tq devuelve el contenido de la pestaña especificada en CSV
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quoted_name}"

    def autenticar_via_csv(self, username, password):
        """
        Descarga la hoja de usuarios y valida las credenciales y permisos.
        """
        if not self.sheet_url:
            return False, [], "No se ha configurado la URL del Google Sheet en config.json."

        csv_url = self._get_csv_export_url()
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) LauncherAuth/1.0"
        }

        try:
            response = requests.get(csv_url, headers=headers, timeout=self.timeout)
            if response.status_code != 200:
                return False, [], f"Error al conectar con Google Sheets (Código HTTP {response.status_code}). Verifica los permisos de acceso."

            content = response.content.decode("utf-8-sig", errors="replace")
            reader = csv.DictReader(io.StringIO(content))
            
            # Normalizar nombres de columnas a minúsculas sin espacios
            if not reader.fieldnames:
                return False, [], "La hoja de cálculo está vacía o no tiene encabezados válidos."

            cleaned_fieldnames = {col: col.strip().lower().replace("ó", "o").replace("á", "a").replace("é", "e").replace("í", "i").replace("ú", "u") for col in reader.fieldnames if col}
            
            usuario_buscado = username.strip().lower()
            password_buscado = password.strip()

            encontrado = False
            for row in reader:
                # Mapear valores con columnas normalizadas
                normalized_row = {cleaned_fieldnames.get(k, k): (v or "").strip() for k, v in row.items()}
                
                # Buscar columnas de usuario
                user_val = (
                    normalized_row.get("usuario") or 
                    normalized_row.get("user") or 
                    normalized_row.get("username") or 
                    normalized_row.get("correo") or 
                    normalized_row.get("email") or ""
                )

                if user_val.lower() == usuario_buscado:
                    encontrado = True
                    
                    # Validar estado (si existe columna estado)
                    estado = (normalized_row.get("estado") or normalized_row.get("status") or "activo").lower()
                    if estado in ["inactivo", "bloqueado", "suspendido", "false", "no"]:
                        return False, [], f"El usuario '{username}' se encuentra desactivado. Contacta al administrador."

                    # Validar contraseña
                    pass_val = (
                        normalized_row.get("password") or 
                        normalized_row.get("contrasena") or 
                        normalized_row.get("clave") or 
                        normalized_row.get("pass") or ""
                    )

                    if pass_val != password_buscado:
                        return False, [], "Contraseña incorrecta."

                    # Obtener módulos asignados
                    modulos = []
                    
                    # Caso 1: Columna 'modulos' con lista separada por comas
                    raw_modulos = (
                        normalized_row.get("modulos") or 
                        normalized_row.get("permisos") or 
                        normalized_row.get("modules") or ""
                    )
                    
                    if raw_modulos:
                        if raw_modulos.strip() == "*":
                            modulos = ["*"]  # Todos
                        else:
                            modulos = [m.strip() for m in raw_modulos.replace(";", ",").split(",") if m.strip()]
                    
                    # Caso 2: Columnas independientes por cada módulo (ej. 'PDF_Unir': 'SI', 'Gene_Documentos': 'SI')
                    system_cols = {"usuario", "user", "username", "correo", "email", "password", "contrasena", "clave", "pass", "estado", "status", "modulos", "permisos", "modules"}
                    positive_values = {"SI", "TRUE", "1", "YES", "OK", "ACTIVO", "X"}

                    for col_orig, col_norm in cleaned_fieldnames.items():
                        if col_norm in system_cols:
                            continue
                        
                        val = normalized_row.get(col_norm, "").upper()
                        if val in positive_values:
                            # Añadir el nombre original de la columna como permiso
                            if col_orig.strip() not in modulos:
                                modulos.append(col_orig.strip())

                            # Reconocer sinónimos habituales para los módulos del repositorio
                            if "unir" in col_norm or "pdf" in col_norm:
                                for syn in ["PDF_Unir", "PDF Unir", "Unir pdf", "unir_pdf"]:
                                    if syn not in modulos:
                                        modulos.append(syn)
                            if "documento" in col_norm or "gene_" in col_norm or "gen_" in col_norm:
                                for syn in ["Gene_Documentos", "Gene Documentos", "Gen_Documentos", "Documentos Masivos", "documentos_masivos"]:
                                    if syn not in modulos:
                                        modulos.append(syn)

                    return True, modulos, "Autenticación exitosa."

            if not encontrado:
                return False, [], f"El usuario '{username}' no existe en el sistema."

        except requests.exceptions.Timeout:
            return False, [], "Tiempo de espera agotado al conectar con Google Sheets. Revisa tu conexión a internet."
        except requests.exceptions.RequestException as e:
            return False, [], f"Error de red al conectar con Google Sheets: {str(e)}"
        except Exception as e:
            return False, [], f"Error inesperado al validar usuario: {str(e)}"

    def autenticar_via_api(self, username, password):
        """
        Valida contra un endpoint de Google Apps Script Web App
        """
        if not self.api_url:
            return self.autenticar_via_csv(username, password)

        payload = {
            "action": "login",
            "usuario": username.strip(),
            "password": password.strip()
        }
        try:
            resp = requests.post(self.api_url, json=payload, timeout=self.timeout)
            data = resp.json()
            if data.get("success"):
                return True, data.get("modulos", []), "Autenticación exitosa."
            else:
                return False, [], data.get("message", "Credenciales incorrectas.")
        except Exception as e:
            # Fallback a CSV si falla la API
            return self.autenticar_via_csv(username, password)

    def autenticar(self, username, password):
        """Punto de entrada principal para autenticación"""
        if self.api_url:
            return self.autenticar_via_api(username, password)
        return self.autenticar_via_csv(username, password)

    # Alias para compatibilidad con llamadas en inglés
    authenticate = autenticar

