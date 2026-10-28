/**
 * SISTEMA DE INSPECCIÓN DE PROCESOS - SHARF
 * Archivo: History.gs — Historial, alertas y formateo de listas/texto
 */

// Resolución de normalizarTextoClave (Config.gs) compatible con Node y Apps Script,
// siguiendo el mismo patrón usado en dashboard/InformeSemanalModelo.js.
var _cfgModHistory = (typeof module !== 'undefined' && module.exports)
  ? require('./Config.js')
  : { normalizarTextoClave: typeof normalizarTextoClave !== 'undefined' ? normalizarTextoClave : null };

function _normTextoClave(t) {
  if (_cfgModHistory && typeof _cfgModHistory.normalizarTextoClave === 'function') return _cfgModHistory.normalizarTextoClave(t);
  if (typeof normalizarTextoClave === 'function') return normalizarTextoClave(t);
  return String(t || "").toUpperCase().trim();
}

/**
 * Layouts de columnas (0-based, como entrega getDataRange().getValues()) usados por
 * calcularAlertaResponsable para leer registros previos. COLS_HISTORIAL_INSPECCIONES es
 * el layout histórico (default, no cambia). COLS_HISTORIAL_5S es el de la hoja Historial_5S,
 * donde la identidad del punto se arma como ID_S + "." + ID_Punto (idPunto no es único por sí solo).
 */
// Se agregan además `fecha`, `actividad` y `subactividad`: no se usan en calcularAlertaResponsable
// (que ignora claves que no consulta), pero sí en Memo.js (obtenerFallasCicloActual) para reconstruir
// la fecha y la "obligación comprometida" (Actividad – Subactividad) de cada falla previa del ciclo.
// Para Historial_5S, la obligación se arma con las columnas S (actividad) y Punto (subactividad).
// `detalle` (comentario) y `urls` (evidencias) se usan solo para mostrar la falla en la revisión de memos.
var COLS_HISTORIAL_INSPECCIONES = { negocio: 1, operacion: 2, idSub: 5, idPunto: null, estado: 7, responsable: 11, tipo: 12, alerta: 13, fecha: 0, actividad: 4, subactividad: 6, detalle: 8, urls: 10 };
var COLS_HISTORIAL_5S = { negocio: 1, operacion: 2, idSub: 4, idPunto: 6, estado: 9, responsable: 13, tipo: 14, alerta: 15, fecha: 0, actividad: 5, subactividad: 8, detalle: 10, urls: 12 };

/**
 * Predicado compartido: determina si una fila del historial cuenta como "falla previa" dentro
 * del ciclo de alertas de un responsable, para una subactividad/punto y negocio/operación dados.
 * Extraído de calcularAlertaResponsable para que Memo.js (obtenerFallasCicloActual) pueda
 * reconstruir el mismo conjunto de filas (para mostrar fecha + obligación en el memorando)
 * sin duplicar la lógica ni alterar su comportamiento.
 * normNeg/normOp ya deben venir normalizados con _normTextoClave; normSub con normalizarIdHistorico;
 * normResp ya normalizado con _normTextoClave.
 */
function filaCuentaComoFallaPrevia(row, c, normNeg, normOp, normSub, normResp) {
  if (!row) return false;

  var tipoHist = (c.tipo !== null && c.tipo !== undefined && row[c.tipo] !== undefined)
    ? String(row[c.tipo] || "").trim().toUpperCase()
    : "";
  var esOficial = (tipoHist === "OFICIAL" || tipoHist === "");
  if (!esOficial) return false;

  var negHist = _normTextoClave(row[c.negocio]);
  var opHist = _normTextoClave(row[c.operacion]);

  var idSubHistNorm = normalizarIdHistorico(row[c.idSub]);
  var subHist = idSubHistNorm;
  if (c.idPunto !== null && c.idPunto !== undefined) {
    var idPuntoHistNorm = normalizarIdHistorico(row[c.idPunto]);
    if (idPuntoHistNorm !== "") {
      subHist = idSubHistNorm + "." + idPuntoHistNorm;
    }
  }

  var estHist = String(row[c.estado] || "").trim().toUpperCase();
  var alertaHist = String(row[c.alerta] || "").trim();

  if (negHist === normNeg && opHist === normOp && subHist === normSub && estHist === "NO CUMPLE") {
    if (alertaHist !== "" && alertaHist !== "-" && alertaHist !== "0") {
      var respHistRaw = String(row[c.responsable] || "").trim();
      var listaRespHist = respHistRaw.split(/[,;]+/)
        .map(function(rh) { return _normTextoClave(rh); })
        .filter(function(rh) { return rh !== "" && rh !== "-" && rh !== "NAN"; });

      return listaRespHist.indexOf(normResp) !== -1;
    }
  }
  return false;
}

/**
 * Agrega el menú de sincronización en el Google Sheet
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Sistema SHARF')
      .addItem('🔄 Sincronizar CeCosOfi Ahora', 'sincronizarCeCosOfi')
      .addToUi();
  } catch (e) {
    Logger.log("Error en onOpen: " + e.toString());
  }
}

/**
 * Auxiliar para formatear listas en HTML
 */
function formatearListaHtml(items) {
  if (!items) return "-";
  var arreglo = [];
  if (Array.isArray(items)) {
    arreglo = items;
  } else if (typeof items === 'string') {
    if (items.trim() === "" || items.trim() === "-") return "-";
    arreglo = items.split(',').map(function(item) { return item.trim(); });
  }

  arreglo = arreglo.filter(function(item) { return item !== "" && item !== "-"; });
  if (arreglo.length === 0) return "-";

  var html = '<ul style="margin: 0; padding-left: 15px; text-align: left; list-style-type: disc; font-family: \'Segoe UI\', Tahoma, Geneva, Verdana, sans-serif;">';
  arreglo.forEach(function(item) {
    html += '<li style="margin-bottom: 3px;">' + item + '</li>';
  });
  html += '</ul>';
  return html;
}

/**
 * Construye URLs simples de fotos
 */
function construirTextoUrls(fotosList) {
  if (!fotosList || fotosList.length === 0) return "-";
  return fotosList.join("\n");
}

/**
 * Formatea el nombre de la actividad garantizando que no se eliminen números completos como "5s".
 */
function limpiarNombreActividad(texto) {
  if (!texto) return "PROCESO GENERAL";
  var textoStr = String(texto).trim();
  var limpio = textoStr.replace(/^\d+(\.\d+)*[\s\-\_]+/, '').trim();
  return limpio !== "" ? limpio.toUpperCase() : textoStr.toUpperCase();
}

/**
 * Formatea una fecha a formato corto para nombres de archivos (ej. '19ago26', '20ago26')
 */
function formatearFechaCortaArchivo(d) {
  var fechaObj = d ? new Date(d) : new Date();
  var dia = ("0" + fechaObj.getDate()).slice(-2);
  var meses = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  var mes = meses[fechaObj.getMonth()] || "mes";
  var anio = String(fechaObj.getFullYear()).slice(-2);
  return dia + mes + anio;
}

/**
 * Normaliza un ID de subactividad o punto leído del historial de Sheets o del checklist.
 * Resuelve el problema donde Sheets convierte cadenas como "4.2" a Date (4 de febrero)
 * o a Number (4.2).
 * - Si es Date: String(getDate()) + '.' + String(getMonth() + 1)
 * - Si es number: String(valor)
 * NOTA: Si una subactividad "4.10" fue guardada como número (4.1) en el historial antiguo de Sheets,
 * matemáticamente es irrecuperable diferenciarla de "4.1".
 */
function normalizarIdHistorico(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date || (typeof val === 'object' && typeof val.getDate === 'function')) {
    return String(val.getDate()) + "." + String(val.getMonth() + 1);
  }
  if (typeof val === 'number') {
    return String(val);
  }
  return String(val).trim();
}

/**
 * Calcula la alerta para el reporte PDF/Correo y para la celda de Google Sheet.
 * Reglas de negocio:
 * 1. Si noAlertar es true, no se emite ninguna alerta ("-").
 * 2. Si el estado no es "NO CUMPLE", no hay alerta ("-").
 * 3. Si no hay responsable asignado (vacío, "-", "NaN", etc.), no hay alerta ("-").
 * 4. Si hay responsable(s), calcula el número de fallas acumuladas para ESE responsable
 *    en esa misma subactividad (negocio + operación + subactividad) que hayan tenido alerta registrada.
 *    Solo cuentan como fallas previas las filas con Tipo = 'OFICIAL' (Tipo vacío o no definido en
 *    históricos antiguos se trata como OFICIAL por compatibilidad retroactiva).
 *    Ciclo de alertas de 3 niveles:
 *    Falla 1 -> "1era Alerta" (Número 1)
 *    Falla 2 -> "2da Alerta" (Número 2)
 *    Falla 3 -> "3era Alerta - Amonestación" (Número 3)
 *    Falla 4 -> "1era Alerta" (Número 1, vuelve a empezar el ciclo)
 *    Falla 5 -> "2da Alerta" (Número 2)
 *    Falla 6 -> "3era Alerta - Amonestación" (Número 3) ...
 * 5. La comparación de responsables se realiza por igualdad exacta del nombre normalizado
 *    (evita falsos positivos por subcadena como 'ANA' dentro de 'MARIANA'), separando
 *    por comas/punto y coma tanto el valor actual como el histórico.
 * 6. Si el envío actual es de PRUEBA (parámetro tipoEnvio):
 *    - textoPdf: Muestra la alerta calculada en el PDF para simulación y visualización del usuario.
 *    - valorSheet: Retorna "-" para que NO compute hacia amonestación en la base de datos de Sheets.
 * 7. En el reporte PDF/Correo se muestra solo la alerta ("1era Alerta" o "1era Alerta<br>2da Alerta")
 *    sin anteponer redundantemente los nombres de los responsables.
 * 8. En Google Sheets se almacena únicamente el valor numérico (1, 2, 3 o "1\n2") o "-" si no aplica / es prueba.
 *
 * Parámetro opcional `cols`: layout de columnas (0-based) de `registrosPrevios` a usar para
 * leer negocio/operación/idSub/idPunto/estado/responsable/tipo/alerta. Por defecto (sin pasar `cols`)
 * conserva el layout de Historial_Inspecciones (COLS_HISTORIAL_INSPECCIONES). Para los puntos de 5S
 * se le pasa COLS_HISTORIAL_5S, cuya identidad de punto se arma como ID_S + "." + ID_Punto.
 *
 * Parámetro opcional `tipoEnvio`: "OFICIAL" o "PRUEBA" (case-insensitive). Si es PRUEBA, valorSheet es "-".
 */
function calcularAlertaResponsable(registrosPrevios, negocio, operacion, idSubKey, responsableVal, estadoActual, noAlertar, cols, tipoEnvio) {
  if (noAlertar === true) return { textoPdf: "-", valorSheet: "-" };
  if (estadoActual !== "NO CUMPLE") return { textoPdf: "-", valorSheet: "-" };
  if (!responsableVal) return { textoPdf: "-", valorSheet: "-" };

  var respStr = String(responsableVal).trim();
  if (respStr === "" || respStr === "-" || respStr.toUpperCase() === "NAN") {
    return { textoPdf: "-", valorSheet: "-" };
  }

  var c = cols || COLS_HISTORIAL_INSPECCIONES;

  var normNeg = _normTextoClave(negocio);
  var normOp = _normTextoClave(operacion);
  var normSub = normalizarIdHistorico(idSubKey);

  var esPrueba = (String(tipoEnvio || "").trim().toLowerCase() === "prueba");

  // Desglosar por si se asignaron múltiples responsables a la misma subactividad
  var listaResp = respStr.split(/[,;]+/)
    .map(function(r) { return r.trim(); })
    .filter(function(r) { return r !== "" && r !== "-" && r.toUpperCase() !== "NAN"; });

  if (listaResp.length === 0) return { textoPdf: "-", valorSheet: "-" };

  var alertasTexto = [];
  var alertasNumeros = [];

  listaResp.forEach(function(nombreResp) {
    var normResp = _normTextoClave(nombreResp);
    var fallasPrevias = 0;

    if (Array.isArray(registrosPrevios)) {
      for (var r = 1; r < registrosPrevios.length; r++) {
        if (filaCuentaComoFallaPrevia(registrosPrevios[r], c, normNeg, normOp, normSub, normResp)) {
          fallasPrevias++;
        }
      }
    }

    // Ciclo de 3 alertas: 1 -> 2 -> 3 (Amonestación) -> 1 -> 2 -> 3 ...
    var numAlerta = (fallasPrevias % 3) + 1;
    var textoInd = "";
    if (numAlerta === 1) textoInd = "1era Alerta";
    else if (numAlerta === 2) textoInd = "2da Alerta";
    else textoInd = "3era Alerta - Amonestación";

    alertasTexto.push(textoInd);
    alertasNumeros.push(numAlerta);
  });

  var textoPdfFinal = alertasTexto.join("<br>");
  // Si el envío actual es de PRUEBA:
  // Se muestra la alerta calculada en el PDF (para simulación),
  // pero en Sheets se guarda "-" para que NO cuente hacia amonestación ni registre alerta en base de datos.
  var valorSheetFinal = esPrueba
    ? "-"
    : (alertasNumeros.length === 1 ? alertasNumeros[0] : alertasNumeros.join("\n"));

  return {
    textoPdf: textoPdfFinal,
    valorSheet: valorSheetFinal
  };
}

/**
 * Cuenta, para un negocio/operación, las fallas previas de cada responsable en cada subactividad/punto,
 * con el mismo criterio que calcularAlertaResponsable (reutiliza filaCuentaComoFallaPrevia).
 * Devuelve { "<idSub normalizado>": { "<RESPONSABLE NORMALIZADO>": n } }.
 * El checklist lo usa para anticipar, mientras se llena, qué responsable está por registrar su
 * 3era Alerta (n % 3 === 2) y mostrarle la bandera de memorando.
 */
function contarFallasPreviasPorResponsable(registrosPrevios, negocio, operacion, cols) {
  var c = cols || COLS_HISTORIAL_INSPECCIONES;
  var normNeg = _normTextoClave(negocio);
  var normOp = _normTextoClave(operacion);
  var conteo = {};
  if (!Array.isArray(registrosPrevios)) return conteo;

  for (var r = 1; r < registrosPrevios.length; r++) {
    var row = registrosPrevios[r];
    if (!row) continue;

    var sub = normalizarIdHistorico(row[c.idSub]);
    if (c.idPunto !== null && c.idPunto !== undefined) {
      var idPuntoNorm = normalizarIdHistorico(row[c.idPunto]);
      if (idPuntoNorm !== "") sub = sub + "." + idPuntoNorm;
    }

    var vistos = {};
    String(row[c.responsable] || "").split(/[,;]+/).forEach(function(rh) {
      var resp = _normTextoClave(rh);
      if (resp === "" || resp === "-" || resp === "NAN" || vistos[resp]) return;
      vistos[resp] = true;
      if (filaCuentaComoFallaPrevia(row, c, normNeg, normOp, sub, resp)) {
        if (!conteo[sub]) conteo[sub] = {};
        conteo[sub][resp] = (conteo[sub][resp] || 0) + 1;
      }
    });
  }
  return conteo;
}

/**
 * Endpoint (google.script.run) para el checklist: fallas previas por subactividad y responsable
 * de la operación seleccionada, separadas por hoja (Historial_Inspecciones y Historial_5S).
 */
function obtenerFallasPreviasAlertas(negocio, operacion) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaNormal = ss.getSheetByName(CONFIG.HOJA_HISTORIAL);
    var hoja5S = ss.getSheetByName(CONFIG.HOJA_HISTORIAL_5S);
    return {
      exito: true,
      normal: hojaNormal ? contarFallasPreviasPorResponsable(hojaNormal.getDataRange().getValues(), negocio, operacion, COLS_HISTORIAL_INSPECCIONES) : {},
      s5: hoja5S ? contarFallasPreviasPorResponsable(hoja5S.getDataRange().getValues(), negocio, operacion, COLS_HISTORIAL_5S) : {}
    };
  } catch (e) {
    Logger.log("Error en obtenerFallasPreviasAlertas: " + e.toString());
    return { exito: false, mensaje: e.toString(), normal: {}, s5: {} };
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    formatearListaHtml: formatearListaHtml,
    construirTextoUrls: construirTextoUrls,
    limpiarNombreActividad: limpiarNombreActividad,
    formatearFechaCortaArchivo: formatearFechaCortaArchivo,
    calcularAlertaResponsable: calcularAlertaResponsable,
    normalizarIdHistorico: normalizarIdHistorico,
    filaCuentaComoFallaPrevia: filaCuentaComoFallaPrevia,
    contarFallasPreviasPorResponsable: contarFallasPreviasPorResponsable,
    COLS_HISTORIAL_INSPECCIONES: COLS_HISTORIAL_INSPECCIONES,
    COLS_HISTORIAL_5S: COLS_HISTORIAL_5S
  };
}
