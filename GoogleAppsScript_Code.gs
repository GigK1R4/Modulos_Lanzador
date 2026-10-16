/**
 * =============================================================================
 * BACKEND SEGURO EN GOOGLE APPS SCRIPT PARA EL LANZADOR
 * =============================================================================
 * 
 * Este script actúa como intermediario seguro (API Web App) entre el Lanzador
 * de escritorio y la hoja de cálculo de Google Sheets.
 * 
 * 🔒 VENTAJA DE SEGURIDAD:
 * - La hoja de Google Sheets permanece 100% PRIVADA (Acceso restringido).
 * - Los usuarios finales NO tienen acceso a la hoja ni a Google Drive.
 * - El cliente (Lanzador.py) nunca descarga la hoja ni ve los datos de otros usuarios.
 * - El cliente envía únicamente usuario y contraseña por POST (HTTPS).
 * - Este script valida credenciales (en texto plano o con hash SHA-256) y responde
 *   únicamente con true/false, los módulos asignados y un token temporal.
 * 
 * 📋 INSTRUCCIONES DE DESPLIEGUE:
 * 1. Abre tu hoja de Google Sheets privada.
 * 2. Ve al menú superior: Extensiones > Apps Script.
 * 3. Borra el código existente y pega TODO el contenido de este archivo.
 * 4. Haz clic en el icono de Guardar (disquete).
 * 5. Arriba a la derecha, haz clic en "Implementar" (o "Deploy") > "Nueva implementación".
 * 6. En el engranaje "Seleccionar tipo", elige: "Aplicación web".
 * 7. Configura los siguientes campos exactamente:
 *    - Descripción: "API Lanzador Autenticación"
 *    - Ejecutar como: "Yo (tu_correo@gmail.com)"  <-- ¡CRÍTICO!
 *    - Quién tiene acceso: "Cualquier persona" (o "Cualquiera con cuenta Google / del dominio") <-- ¡CRÍTICO!
 * 8. Haz clic en "Implementar" y autoriza los permisos solicitados.
 * 9. Copia la "URL de la aplicación web" (termina en /exec) y pégala en config.json
 *    en el campo "apps_script_url".
 * =============================================================================
 */

// Nombre de la pestaña de usuarios en tu Google Sheet
const HOJA_USUARIOS = "Usuarios";

// ID opcional de la hoja (solo si el script es independiente y no está dentro de la hoja)
// Si el script lo abriste desde Extensiones > Apps Script en la misma hoja, déjalo vacío ("").
const ID_HOJA_CALCULO = "";

// Secreto interno para firma de tokens temporales (puedes cambiarlo si deseas)
const SECRETO_TOKEN = "Lanzador_Key_Secure_2026_@";

function doGet(e) {
  // Manejo de peticiones GET de prueba o verificación de estado
  return jsonResponse({
    success: true,
    message: "Servicio de Autenticación de Lanzador activo. Utilice peticiones POST para iniciar sesión."
  });
}

function doPost(e) {
  try {
    var params = {};
    if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (err) {
        params = e.parameter || {};
      }
    } else if (e && e.parameter) {
      params = e.parameter;
    }

    var action = (params.action || "login").toString().trim().toLowerCase();

    // Acción de comprobación de salud del backend
    if (action === "test" || action === "ping") {
      return jsonResponse({
        success: true,
        message: "API Google Apps Script operativa y respondiendo correctamente."
      });
    }

    // Acción de inicio de sesión
    if (action === "login") {
      var usuario = (params.usuario || params.user || params.username || "").toString().trim().toLowerCase();
      var password = (params.password || params.contrasena || params.clave || "").toString().trim();

      if (!usuario || !password) {
        return jsonResponse({
          success: false,
          message: "Debe ingresar usuario y contraseña."
        });
      }

      // Obtener la hoja de cálculo de forma segura
      var ss = null;
      if (ID_HOJA_CALCULO && ID_HOJA_CALCULO.trim() !== "") {
        ss = SpreadsheetApp.openById(ID_HOJA_CALCULO.trim());
      } else {
        ss = SpreadsheetApp.getActiveSpreadsheet();
      }

      if (!ss) {
        return jsonResponse({
          success: false,
          message: "Error interno: No se pudo conectar con la base de datos de usuarios."
        });
      }

      var sheet = ss.getSheetByName(HOJA_USUARIOS);
      if (!sheet) {
        return jsonResponse({
          success: false,
          message: "Error de configuración: No existe la pestaña '" + HOJA_USUARIOS + "' en la hoja."
        });
      }

      var data = sheet.getDataRange().getValues();
      if (!data || data.length < 2) {
        return jsonResponse({
          success: false,
          message: "No se encontraron usuarios registrados en el sistema."
        });
      }

      // Encabezados normalizados
      var headers = data[0].map(function(h) { return h.toString().trim().toLowerCase(); });
      var idxUser = headers.indexOf("usuario");
      if (idxUser === -1) idxUser = headers.indexOf("user");
      if (idxUser === -1) idxUser = headers.indexOf("correo");
      if (idxUser === -1) idxUser = headers.indexOf("email");

      var idxPass = headers.indexOf("password");
      if (idxPass === -1) idxPass = headers.indexOf("contrasena");
      if (idxPass === -1) idxPass = headers.indexOf("contraseña");
      if (idxPass === -1) idxPass = headers.indexOf("clave");
      if (idxPass === -1) idxPass = headers.indexOf("hash");

      var idxModulos = headers.indexOf("modulos");
      if (idxModulos === -1) idxModulos = headers.indexOf("permisos");
      if (idxModulos === -1) idxModulos = headers.indexOf("modules");

      var idxEstado = headers.indexOf("estado");
      if (idxEstado === -1) idxEstado = headers.indexOf("status");

      if (idxUser === -1 || idxPass === -1) {
        return jsonResponse({
          success: false,
          message: "Estructura de hoja inválida. Debe incluir columnas 'Usuario' y 'Password'."
        });
      }

      // Calcular hash SHA-256 de la contraseña recibida para permitir validación doble
      var passSha256 = calcularHashSHA256(password);

      for (var i = 1; i < data.length; i++) {
        var rowUser = data[i][idxUser].toString().trim().toLowerCase();
        if (rowUser === usuario) {
          var rowPass = data[i][idxPass].toString().trim();
          var rowEstado = idxEstado !== -1 ? data[i][idxEstado].toString().trim().toLowerCase() : "activo";

          // 1. Validar estado del usuario
          if (rowEstado === "inactivo" || rowEstado === "bloqueado" || rowEstado === "desactivado" || rowEstado === "false" || rowEstado === "no") {
            return jsonResponse({
              success: false,
              message: "El usuario está inactivo o suspendido. Contacta al administrador."
            });
          }

          // 2. Validar contraseña: soporta texto plano directo O hash SHA-256 almacenado
          var esValida = false;
          if (rowPass === password) {
            esValida = true;
          } else if (rowPass.toLowerCase() === passSha256.toLowerCase()) {
            esValida = true;
          }

          if (!esValida) {
            return jsonResponse({
              success: false,
              message: "Usuario o contraseña incorrectos."
            });
          }

          // 3. Procesar lista de módulos permitidos
          var modulos = [];
          if (idxModulos !== -1) {
            var rawMod = data[i][idxModulos].toString().trim();
            if (rawMod === "*") {
              modulos = ["*"];
            } else if (rawMod) {
              modulos = rawMod.replace(/;/g, ",").split(",").map(function(m) { return m.trim(); }).filter(Boolean);
            }
          }

          // Chequear columnas individuales por módulo si existen (ej. 'PDF_Unir': 'SI')
          var systemHeaders = ["usuario", "user", "username", "correo", "email", "password", "contrasena", "contraseña", "clave", "pass", "hash", "estado", "status", "modulos", "permisos", "modules"];
          var positiveVals = ["SI", "TRUE", "1", "YES", "OK", "ACTIVO", "X"];

          headers.forEach(function(h, colIdx) {
            if (systemHeaders.indexOf(h) !== -1) return;
            var val = data[i][colIdx].toString().trim().toUpperCase();
            if (positiveVals.indexOf(val) !== -1) {
              var rawHeader = data[0][colIdx].toString().trim();
              if (modulos.indexOf(rawHeader) === -1) modulos.push(rawHeader);

              if (h.indexOf("unir") !== -1 || h.indexOf("pdf") !== -1) {
                ["PDF_Unir", "PDF Unir", "Unir pdf"].forEach(function(syn) {
                  if (modulos.indexOf(syn) === -1) modulos.push(syn);
                });
              }
              if (h.indexOf("documento") !== -1 || h.indexOf("gene") !== -1 || h.indexOf("gen") !== -1) {
                ["Gene_Documentos", "Gene Documentos", "Gen_Documentos", "Documentos Masivos"].forEach(function(syn) {
                  if (modulos.indexOf(syn) === -1) modulos.push(syn);
                });
              }
            }
          });

          // 4. Generar token temporal de sesión seguro
          var expiraEn = new Date().getTime() + (8 * 60 * 60 * 1000); // 8 horas
          var tokenData = usuario + ":" + expiraEn;
          var tokenFirma = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(tokenData, SECRETO_TOKEN));
          var tokenTemporal = Utilities.base64EncodeWebSafe(tokenData) + "." + tokenFirma;

          // Responder ÚNICAMENTE con los datos necesarios (cero exposición de la hoja)
          return jsonResponse({
            success: true,
            usuario: rowUser,
            modulos: modulos,
            token: tokenTemporal,
            message: "Autenticación exitosa."
          });
        }
      }

      // Si no coincide el usuario
      return jsonResponse({
        success: false,
        message: "Usuario o contraseña incorrectos."
      });
    }

    return jsonResponse({
      success: false,
      message: "Acción no reconocida."
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      message: "Error interno del servidor de autenticación: " + error.toString()
    });
  }
}

/**
 * Calcula el Hash SHA-256 de una cadena de texto.
 * También se puede usar como fórmula personalizada dentro de Google Sheets:
 * Ejemplo en celda: =SHA256("MiContraseña123")
 */
function SHA256(texto) {
  return calcularHashSHA256(texto);
}

function calcularHashSHA256(input) {
  if (!input) return "";
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input.toString(), Utilities.Charset.UTF_8);
  var txtHash = "";
  for (var i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    var byteString = hashVal.toString(16);
    if (byteString.length == 1) byteString = "0" + byteString;
    txtHash += byteString;
  }
  return txtHash;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
