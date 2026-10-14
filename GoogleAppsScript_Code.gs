/**
 * API Backend en Google Apps Script para el Lanzador de Automatización de Archivos
 * 
 * Permite autenticar usuarios y obtener sus módulos asignados desde una pestaña 'Usuarios'.
 * 
 * Instrucciones de instalación:
 * 1. En tu hoja de cálculo de Google Sheets, ve a Extensiones > Apps Script.
 * 2. Pega este código.
 * 3. Haz clic en Desplegar > Nuevo despliegue > Tipo: "Aplicación web".
 * 4. Configura:
 *    - Ejecutar como: "Yo"
 *    - Quién tiene acceso: "Cualquier persona" (Anyone)
 * 5. Copia la URL del despliegue y pégala en `apps_script_url` dentro de config.json.
 */

const HOJA_USUARIOS = "Usuarios";

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
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

    var action = params.action || "test";

    if (action === "test") {
      return jsonResponse({
        success: true,
        message: "API del Lanzador activa y funcionando correctamente."
      });
    }

    if (action === "login") {
      var usuario = (params.usuario || "").toString().trim().toLowerCase();
      var password = (params.password || "").toString().trim();

      if (!usuario || !password) {
        return jsonResponse({
          success: false,
          message: "Debe ingresar usuario y contraseña."
        });
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(HOJA_USUARIOS);

      if (!sheet) {
        return jsonResponse({
          success: false,
          message: "No se encontró la pestaña '" + HOJA_USUARIOS + "' en la hoja de cálculo."
        });
      }

      var data = sheet.getDataRange().getValues();
      if (data.length < 2) {
        return jsonResponse({
          success: false,
          message: "No hay registros de usuarios en la hoja."
        });
      }

      var headers = data[0].map(function(h) { return h.toString().trim().toLowerCase(); });
      var idxUser = headers.indexOf("usuario");
      if (idxUser === -1) idxUser = headers.indexOf("user");

      var idxPass = headers.indexOf("password");
      if (idxPass === -1) idxPass = headers.indexOf("contrasena");
      if (idxPass === -1) idxPass = headers.indexOf("contraseña");

      var idxModulos = headers.indexOf("modulos");
      if (idxModulos === -1) idxModulos = headers.indexOf("permisos");

      var idxEstado = headers.indexOf("estado");

      if (idxUser === -1 || idxPass === -1) {
        return jsonResponse({
          success: false,
          message: "La hoja debe tener columnas 'Usuario' y 'Password'."
        });
      }

      for (var i = 1; i < data.length; i++) {
        var rowUser = data[i][idxUser].toString().trim().toLowerCase();
        var rowPass = data[i][idxPass].toString().trim();
        var rowEstado = idxEstado !== -1 ? data[i][idxEstado].toString().trim().toLowerCase() : "activo";

        if (rowUser === usuario) {
          if (rowEstado === "inactivo" || rowEstado === "bloqueado") {
            return jsonResponse({
              success: false,
              message: "El usuario está inactivo o bloqueado. Contacta al administrador."
            });
          }

          if (rowPass !== password) {
            return jsonResponse({
              success: false,
              message: "Contraseña incorrecta."
            });
          }

          // Procesar módulos permitidos
          var modulos = [];
          if (idxModulos !== -1) {
            var rawMod = data[i][idxModulos].toString().trim();
            if (rawMod === "*") {
              modulos = ["*"];
            } else {
              modulos = rawMod.split(",").map(function(m) { return m.trim(); }).filter(Boolean);
            }
          }

          // Chequear columnas directas por módulo (ej. 'PDF_Unir', 'Gene_Documentos' = 'SI')
          var systemHeaders = ["usuario", "user", "username", "correo", "email", "password", "contrasena", "contraseña", "clave", "pass", "estado", "status", "modulos", "permisos", "modules"];
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

          return jsonResponse({
            success: true,
            usuario: rowUser,
            modulos: modulos,
            message: "Autenticación exitosa."
          });
        }
      }

      return jsonResponse({
        success: false,
        message: "El usuario no se encuentra registrado."
      });
    }

    return jsonResponse({
      success: false,
      message: "Acción no reconocida."
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      message: "Error en el servidor Apps Script: " + error.toString()
    });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
