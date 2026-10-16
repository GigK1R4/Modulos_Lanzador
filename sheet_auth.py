"""
=============================================================================
MÓDULO DE AUTENTICACIÓN SEGURA MEDIANTE GOOGLE APPS SCRIPT WEB APP
=============================================================================
Este servicio se comunica exclusivamente con la API Web App de Google Apps Script.

Arquitectura de Seguridad:
1. No se conecta directamente a Google Sheets ni descarga archivos CSV en el cliente.
2. La hoja de cálculo permanece 100% PRIVADA en Google Drive del administrador.
3. El cliente envía credenciales mediante petición HTTPS POST hacia la Web App.
4. Apps Script valida las credenciales en el servidor y retorna únicamente:
   - success (bool)
   - modulos (list)
   - token (string con token temporal de sesión)
   - message (string)
=============================================================================
"""
import json
import requests

class SheetAuth:
    def __init__(self, config):
        self.config = config
        self.api_url = (config.get("apps_script_url") or "").strip()
        self.timeout = config.get("network_timeout", 10)
        self.session_token = None
        self.session_user = None

    def autenticar(self, username, password):
        """
        Valida las credenciales enviando una petición POST a la Web App de Apps Script.
        
        Retorna:
            (éxito: bool, modulos: list, mensaje: str)
        """
        user_clean = (username or "").strip()
        pass_clean = (password or "").strip()

        if not user_clean or not pass_clean:
            return False, [], "Por favor ingrese su usuario y contraseña."

        if not self.api_url:
            return (
                False,
                [],
                "No se ha configurado la URL de Google Apps Script.\n"
                "Por favor configure 'apps_script_url' en config.json o use el botón de configuración."
            )

        payload = {
            "action": "login",
            "usuario": user_clean,
            "password": pass_clean
        }

        headers = {
            "Content-Type": "application/json",
            "User-Agent": "LanzadorAuth/2.0"
        }

        try:
            # allow_redirects=True es indispensable porque Apps Script responde con redirección 302
            # hacia script.googleusercontent.com
            response = requests.post(
                self.api_url,
                json=payload,
                headers=headers,
                timeout=self.timeout,
                allow_redirects=True
            )

            if response.status_code != 200:
                return (
                    False,
                    [],
                    f"Error de conexión con el servidor (Código HTTP {response.status_code}). "
                    "Verifique la URL del despliegue en Apps Script."
                )

            try:
                data = response.json()
            except Exception:
                # Si no devolvió JSON directo, verificar el texto de respuesta
                text_preview = response.text[:200] if response.text else "Respuesta vacía"
                return (
                    False,
                    [],
                    f"El servidor no devolvió una respuesta válida. Contenido recibido: {text_preview}"
                )

            if data.get("success"):
                self.session_user = data.get("usuario", user_clean)
                self.session_token = data.get("token")
                modulos = data.get("modulos", [])
                mensaje = data.get("message", "Autenticación exitosa.")
                return True, modulos, mensaje
            else:
                mensaje_error = data.get("message", "Credenciales incorrectas o usuario no registrado.")
                return False, [], mensaje_error

        except requests.exceptions.Timeout:
            return (
                False,
                [],
                "Tiempo de espera agotado al conectar con el servidor de autenticación. "
                "Revise su conexión a internet."
            )
        except requests.exceptions.ConnectionError:
            return (
                False,
                [],
                "No se pudo establecer conexión con Google Apps Script. "
                "Verifique su conexión a internet o la URL configurada."
            )
        except Exception as e:
            return False, [], f"Error al procesar autenticación: {str(e)}"

    def verificar_conexion(self):
        """Prueba rápida para verificar si la URL de Apps Script está activa"""
        if not self.api_url:
            return False, "URL de Apps Script no configurada."

        try:
            payload = {"action": "test"}
            response = requests.post(
                self.api_url,
                json=payload,
                timeout=self.timeout,
                allow_redirects=True
            )
            if response.status_code == 200:
                data = response.json()
                return True, data.get("message", "API activa.")
            return False, f"Código HTTP: {response.status_code}"
        except Exception as e:
            return False, str(e)

    # Alias para compatibilidad
    authenticate = autenticar
