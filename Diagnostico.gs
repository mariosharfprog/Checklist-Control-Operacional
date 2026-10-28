/**
 * SISTEMA DE INSPECCIÓN DE PROCESOS - SHARF
 * Archivo: Diagnostico.gs — Diagnóstico temporal del arranque del checklist.
 *
 * CÓMO USARLO: en el editor de Apps Script, elige la función `diagnostico5S`
 * en el desplegable de arriba y pulsa Ejecutar. Luego abre "Registro de ejecución"
 * y copia todo lo que salga.
 *
 * Este archivo NO lo usa la aplicación: se puede borrar cuando ya no haga falta.
 */

function diagnostico5S() {
  var L = [];
  function log(msg) { L.push(msg); Logger.log(msg); }
  function cronometrar(nombre, fn) {
    var t0 = new Date().getTime();
    try {
      var r = fn();
      log("  OK   " + nombre + "  (" + (new Date().getTime() - t0) + " ms)");
      return r;
    } catch (e) {
      log("  FALLA " + nombre + "  (" + (new Date().getTime() - t0) + " ms)");
      log("        " + e.toString());
      if (e.stack) log("        " + String(e.stack).split("\n").slice(0, 3).join(" | "));
      throw e;
    }
  }

  log("===== DIAGNÓSTICO 5S =====");

  // 1) ¿Están presentes todos los archivos del proyecto?
  log("1) Funciones y constantes esperadas:");
  var requeridos = [
    "cargarDatosDesdeSheet", "cargarPlantillas5S", "cargarDetalle5S", "anexarMacro5S",
    "calcularPorcentajesS", "calcularPctActividad5S", "esHojaDetalle5S", "esHoja5S",
    "capitalizarNegocio", "normalizarTextoClave", "calcularAlertaResponsable",
    "procesarInspeccion", "obtenerUsuarioActual"
  ];
  var faltantes = [];
  for (var i = 0; i < requeridos.length; i++) {
    var nom = requeridos[i];
    var existe = false;
    try { existe = (eval("typeof " + nom) === "function"); } catch (e2) { existe = false; }
    if (!existe) faltantes.push(nom);
  }
  log(faltantes.length ? "  *** FALTAN: " + faltantes.join(", ") + " ***" : "  Todas presentes.");

  try {
    log("  CONFIG.HOJA_CL_5S      = " + CONFIG.HOJA_CL_5S);
    log("  CONFIG.UMBRAL_5S       = " + CONFIG.UMBRAL_5S);
    log("  CONFIG.HOJA_HISTORIAL_5S = " + CONFIG.HOJA_HISTORIAL_5S);
  } catch (e3) {
    log("  *** CONFIG incompleto: " + e3.toString() + " ***");
  }

  // 2) Hojas del libro y tamaño de cada rango
  var ss = cronometrar("SpreadsheetApp.getActiveSpreadsheet()", function() {
    return SpreadsheetApp.getActiveSpreadsheet();
  });
  var hojas = cronometrar("ss.getSheets()", function() { return ss.getSheets(); });

  log("2) Hojas (" + hojas.length + "):");
  for (var h = 0; h < hojas.length; h++) {
    var sh = hojas[h];
    var nom2 = sh.getName();
    var marca = "";
    try {
      if (typeof esHoja5S === "function" && esHoja5S(nom2)) marca = "  <- plantilla cl_5s";
      else if (typeof esHojaDetalle5S === "function" && esHojaDetalle5S(nom2)) marca = "  <- DESGLOSE 5S";
    } catch (e4) {}
    log("  " + nom2 + "  [" + sh.getLastRow() + " x " + sh.getLastColumn() + "]" + marca);
  }

  // 3) Lectura cronometrada de las hojas de 5S (aquí se ven los IMPORTRANGE lentos)
  log("3) Lectura de las hojas de 5S:");
  for (var k = 0; k < hojas.length; k++) {
    var sh2 = hojas[k];
    var n2 = sh2.getName();
    var esDet = false;
    try { esDet = (typeof esHojaDetalle5S === "function" && esHojaDetalle5S(n2)); } catch (e5) {}
    if (!esDet) continue;
    (function(hoja, nombre) {
      var vals = cronometrar("getDisplayValues() de " + nombre, function() {
        return hoja.getDataRange().getDisplayValues();
      });
      log("        filas leídas: " + vals.length);
      // Muestra la columna Cumplimiento de las primeras filas de S1 (detecta IMPORTRANGE roto)
      var muestras = 0;
      for (var r = 1; r < vals.length && muestras < 3; r++) {
        if (String(vals[r][2] || "").indexOf("1") === 0) {
          log("        S1 | op=" + vals[r][0] + " | id=" + vals[r][1] + " | Cumplimiento=\"" + vals[r][5] + "\"");
          muestras++;
        }
      }
    })(sh2, n2);
  }

  // 4) El pipeline completo, por etapas
  log("4) Pipeline:");
  var plantillas = cronometrar("cargarPlantillas5S(ss)", function() { return cargarPlantillas5S(ss); });
  log("        plantillas: " + (plantillas ? plantillas.length : "null"));

  var detalle = cronometrar("cargarDetalle5S(ss)", function() { return cargarDetalle5S(ss); });
  var claves = detalle ? Object.keys(detalle) : [];
  log("        claves de detalle: " + claves.length);
  for (var c = 0; c < Math.min(claves.length, 8); c++) {
    log("          " + claves[c] + " -> " + detalle[claves[c]].length + " puntos");
  }

  var res = cronometrar("cargarDatosDesdeSheet()  <-- LA QUE USA LA WEBAPP", function() {
    return cargarDatosDesdeSheet();
  });
  log("        status: " + (res && res.status));
  if (res && res.status !== "SUCCESS") log("        message: " + res.message);
  if (res && res.checklist) {
    log("        items: " + res.checklist.length);
    var negs = {}, cinco = 0;
    for (var q = 0; q < res.checklist.length; q++) {
      negs[res.checklist[q].negocio] = 1;
      if (String(res.checklist[q].actividad || "").trim() === "5S") cinco++;
    }
    log("        negocios: " + Object.keys(negs).join(", "));
    log("        items de 5S: " + cinco);
  }

  log("===== FIN =====");
  return L.join("\n");
}

/**
 * Compara el archivo Index.html tal como está guardado contra lo que HtmlService
 * realmente va a servir. Si HtmlService recorta o transforma el HTML, aquí se ve.
 *
 * CÓMO USARLO: en el editor, elige `diagnosticoHtmlServido` y pulsa Ejecutar.
 * Luego abre "Registro de ejecución" y copia todo.
 */
function diagnosticoHtmlServido() {
  var L = [];
  function log(m) { L.push(m); Logger.log(m); }

  log("===== HTML SERVIDO vs ARCHIVO =====");

  var contenido = HtmlService.createHtmlOutputFromFile('Index').getContent();
  log("longitud que se va a servir : " + contenido.length + " chars");
  log("longitud esperada del archivo: 124035 chars (lo que hay en el proyecto)");
  log(contenido.length < 124000
      ? "*** SE ESTÁ RECORTANDO: faltan " + (124035 - contenido.length) + " chars ***"
      : "longitud OK, no hay recorte");

  log("");
  log("etiquetas <script  : " + (contenido.match(/<script/gi) || []).length + "  (esperado 6)");
  log("etiquetas </script : " + (contenido.match(/<\/script>/gi) || []).length + "  (esperado 6)");
  log("¿tiene </body>?    : " + (contenido.indexOf("</body>") !== -1));
  log("¿tiene </html>?    : " + (contenido.indexOf("</html>") !== -1));
  log("¿tiene la sonda DESPUÉS? : " + (contenido.indexOf("__probeDespues") !== -1));
  log("¿tiene alRecibirDatos?   : " + (contenido.indexOf("function alRecibirDatos") !== -1));

  log("");
  log("ÚLTIMOS 200 caracteres de lo servido:");
  log(contenido.slice(-200));

  // Línea donde termina, para cruzarlo con el error del navegador
  var lineas = contenido.split("\n");
  log("");
  log("total de líneas servidas: " + lineas.length);
  log("línea 909 servida: " + (lineas[908] !== undefined ? lineas[908].slice(0, 160) : "(no existe: el HTML es más corto)"));

  return L.join("\n");
}
