/**
 * SISTEMA DE INSPECCIÓN DE PROCESOS - SHARF
 * Desarrollado para Google Apps Script
 * Archivo: Config.gs — Configuración global y utilidades de estilo/carpetas
 */

// CONFIGURACIÓN GLOBAL
var CONFIG = {
  HOJA_CHECKLIST: "Config_Checklist",
  HOJA_PERSONAL: "Config_Personal",
  HOJA_PERSONAL_ALT: "PERSONAL",
  HOJA_HISTORIAL: "Historial_Inspecciones",
  HOJA_HISTORIAL_5S: "Historial_5S",
  HOJA_COMENTARIOS_EXTRA: "Comentarios_Extra",
  HOJA_CL_5S: "5s",
  UMBRAL_5S: 80,
  
  // ID del Logo SHARF en Google Drive - sharf_era_blanco (PDF) https://drive.google.com/file/d/1goE6UrKGg4Wgdmf8_UMp-nL_B66H9-_K/view?usp=sharing
  ID_LOGO_DRIVE: "1goE6UrKGg4Wgdmf8_UMp-nL_B66H9-_K",
  
  // Color Personalizado de Cabecera — Unificado con cabecera del checklist (#D452FF) oscurecido un par de tonos → #8E24AA
  COLOR_HEADER_CUSTOM: "#8E24AA",
  
  // Paleta Primaria Sharf
  COLOR_WHITE: "#FFFFFF",
  COLOR_BLACK: "#000000",
  COLOR_INK: "#111111",           // Texto Principal
  COLOR_INK_SOFT: "#3A3A3A",      // Texto Secundario / Captions
  COLOR_RED_BOLD: "#FF6568",      // Acento Sharf / Alertas
  COLOR_RED_DEEP: "#68002B",      // Texto Destacado Alerta / < 60%
  COLOR_RED_PALE: "#FFC1C2",      // Fondos Callout Alerta / < 60%
  COLOR_LINE: "#E6E6E6",          // Hairlines de tablas
  COLOR_LINE_SOFT: "#F1F1F1",     // Surfaces sutiles / Totales
  
  // Paleta Secundaria Sharf - Operaciones (Verde)
  COLOR_GREEN_BOLD: "#00CCAA",
  COLOR_GREEN_DEEP: "#003C46",    // >= 80%
  COLOR_GREEN_PALE: "#99EBDD",    // >= 80%

  // Paleta Preventiva (Amarillo / Ámbar)
  COLOR_YELLOW_DEEP: "#B8860B",   // [60%;80)
  COLOR_YELLOW_PALE: "#FFFACD",   // [60%;80)

  // MAPA DE CARPETAS DRIVE POR NEGOCIO Y OPERACIÓN
  // Nota: "HUB Cedros" reemplaza a "CALLAO" (se mantiene alias CALLAO por compatibilidad histórica)
  MAPA_CARPETAS_DRIVE: {
    "ALMACEN|PLC": "12dEH0ON3Dd7CzDW-xotE3xJbZEefGH0k",
    "ALMACEN|LURIN": "1pW-fT7A-iLFUe_6b_RuVOE6W1XLNV5Ot",
    "ALMACEN|LURÍN": "1pW-fT7A-iLFUe_6b_RuVOE6W1XLNV5Ot",
    "FEDEX|INBOUND": "1lo_DO5CBS5JVwpMRDeCyoeAJG6c7NBuZ",
    "FEDEX|OUTBOUND": "1IIL7WlmZ43DIm3MgdQjQ47QMtgP9Xg5C",
    "DISTRIBUCION|HUB CEDROS": "181ZWrDztaGmesmDtGpZ51Sf-kERd-xPu",
    "DISTRIBUCION|CALLAO": "181ZWrDztaGmesmDtGpZ51Sf-kERd-xPu",
    "DISTRIBUCION|PLC": "1QKrkYSarZ5cknWb0ftu1p6Srk-FMUUWg",
    "DISTRIBUCION|LURIN": "1vFouy_CZzGpeVFia5AT_V3WFMjZbWIny",
    "DISTRIBUCION|LURÍN": "1vFouy_CZzGpeVFia5AT_V3WFMjZbWIny",
    "DISTRIBUCION|CLARO REGULAR": "1JtgENnyt_xD-ni8_6rqvuAN0nn-CQlXu",
    "DISTRIBUCION|CLARO PROVINCIA": "13mMKAnjyng0rVltoSqcxccxGZLUdGGlI",
    "DISTRIBUCION|CLARO CHIPS LIMA": "13mMKAnjyng0rVltoSqcxccxGZLUdGGlI",
    "COURIER|COURIER": "1r_kzIth3sHOFM0JxX8TQ47qXn5rDn2Fs",
    "COURIER|GENERAL": "1r_kzIth3sHOFM0JxX8TQ47qXn5rDn2Fs",
    "DISTRIBUCION|COURIER": "1r_kzIth3sHOFM0JxX8TQ47qXn5rDn2Fs",
    "ALMACEN|COURIER": "1r_kzIth3sHOFM0JxX8TQ47qXn5rDn2Fs",
    "FEDEX|COURIER": "1r_kzIth3sHOFM0JxX8TQ47qXn5rDn2Fs"
  },

  // CARPETA RAÍZ DE REPORTES DE INSPECCIÓN - Jerarquía: Año / Negocio / Operación
  // https://drive.google.com/drive/folders/1-q7curdFAaeVDj3bda1e9esXETEoMRE-?usp=sharing
  ID_CARPETA_REPORTES: "1-q7curdFAaeVDj3bda1e9esXETEoMRE-",

  // MEMORANDOS DE AMONESTACIÓN (3era Alerta) — Ver Memo.js
  // Plantilla Google Doc con los marcadores {{...}} a reemplazar.
  MEMO_PLANTILLA_ID: "19Z8isDgjLT72GbKw-Ry1_BJsV6cnypbV8t3hG9WkwQ8",
  // Plantilla opcional para memos activados manualmente por el inspector en 1era/2da alerta
  // (bandera blanca), con los mismos marcadores {{...}} que MEMO_PLANTILLA_ID. Vacía ("") = se usa
  // MEMO_PLANTILLA_ID también para esos memos (ver elegirPlantillaMemo en Memo.js).
  MEMO_PLANTILLA_ID_MANUAL: "",
  // Fase de pruebas: el correo del memorando SOLO se envía a quienes tengan este Rol en Config_Personal.
  MEMO_ROL_DESTINATARIO: "Prueba_Memo",
  // Modo OFICIAL (MEMO_MODO_PRUEBA = false): el correo del memorando se envía "para" (to) a
  // quienes tengan este Rol ("Memos") en Config_Personal. Si ninguna fila tiene este Rol, el
  // memo NO se envía y se registra en la hoja Memorandos como "Pendiente - sin destinatarios".
  MEMO_ROL_OFICIAL: "Memos",
  // Hoja de bitácora de memorandos generados.
  HOJA_MEMORANDOS: "Memorandos",
  // Subcarpeta (dentro de ID_CARPETA_REPORTES) donde se guardan los Doc/PDF de memorandos.
  MEMO_NOMBRE_CARPETA: "Memorandos",
  // URL /exec del despliegue (Web App) del proyecto firma-memo (otro proyecto Apps Script, corre
  // como propietario) que atiende los botones Conforme / No conforme del correo del memorando.
  MEMO_URL_FIRMA: "https://script.google.com/a/macros/holasharf.com/s/AKfycbzB6D8GI_HpmY1GtmEoxIsNjb_0NJilyvBapviU3A3T5UC4SiRp2uZZFCNDehYlCIxjGg/exec",
  // false = modo OFICIAL: el memo se envía "para" (to) al Rol MEMO_ROL_OFICIAL ("Memos"), con
  // copia (cc) al colaborador amonestado, su jefe inmediato y el inspector que envió el
  // checklist. true = modo PRUEBAS: el correo va SOLO a los destinatarios de Rol
  // MEMO_ROL_DESTINATARIO ("Prueba_Memo"), sin cc ni bcc (para volver a pruebas, poner esto en true).
  MEMO_MODO_PRUEBA: false,
  // En modo OFICIAL, si es true se envía además copia oculta (bcc) a los correos de Rol
  // MEMO_ROL_DESTINATARIO ("Prueba_Memo") para monitoreo del flujo ya en producción.
  MEMO_BCC_MONITOREO: true,
  // Nombres (tal como aparecen en la columna Nombre de cecos) de quienes emiten/firman los
  // memorandos; se resuelven contra cecos con resolverEmisoras (Memo.js).
  MEMO_EMISORAS: ["MARGIE TARAZONA", "KATHERIN LEANDRO"],
  // Plazo (en horas) que tiene el colaborador para responder Conforme/No conforme desde el envío.
  MEMO_PLAZO_HORAS: 24,
  // Revisión humana previa al memo: al detectarse la 3era alerta la fila queda en este estado (sin
  // Token, no se envía nada) hasta que el inspector que envió el checklist decide generar o descartar.
  MEMO_ESTADO_PENDIENTE_REVISION: "Pendiente de revisión",
  MEMO_ESTADO_DESCARTADO: "Descartado por revisión"
};

/**
 * Función auxiliar para normalizar texto quitando tildes y caracteres especiales
 */
function normalizarTextoClave(texto) {
  if (!texto) return "";
  return String(texto)
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

/**
 * Normaliza nombre de operación sin hardcodear valores específicos.
 * El checklist (cl_* / Config_Checklist) es la fuente única de verdad;
 * el flujo no debe buscar explícitamente "CALLAO" ni ningún otro valor.
 * Esta función solo recorta espacios y preserva el valor tal cual viene
 * de la hoja, para mantener el sistema 100% genérico a futuros negocios/operaciones.
 */
function canonicalizarOperacion(operacion) {
  return String(operacion || "").trim();
}

/**
 * Obtiene la carpeta de destino en Google Drive según Negocio y Operación
 */
function obtenerCarpetaDestino(negocio, operacion) {
  var normNeg = normalizarTextoClave(negocio);
  var normOp = normalizarTextoClave(operacion);
  var clave = normNeg + "|" + normOp;

  var folderId = CONFIG.MAPA_CARPETAS_DRIVE[clave];

  // Fallback específico para Courier: cualquier combinación que contenga COURIER usa su carpeta dedicada
  if (!folderId && (normNeg === "COURIER" || normOp === "COURIER")) {
    folderId = "1r_kzIth3sHOFM0JxX8TQ47qXn5rDn2Fs";
  }

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      Logger.log("Error al abrir la carpeta con ID " + folderId + ": " + e.toString());
    }
  }

  // Si no se encuentra en la matriz o falla la ID, busca/crea una carpeta por defecto ordenada
  var nombreCarpetaFallback = "Evidencias_" + normNeg + "_" + normOp;
  var carpetas = DriveApp.getFoldersByName(nombreCarpetaFallback);
  if (carpetas.hasNext()) {
    return carpetas.next();
  } else {
    return DriveApp.createFolder(nombreCarpetaFallback);
  }
}

/**
 * Sanitiza un nombre para uso como nombre de archivo/carpeta en Drive
 * Reemplaza espacios por _ y elimina caracteres inválidos, colapsa _ múltiples
 */
function sanitizarNombreArchivo(nombre) {
  if (!nombre) return "";
  var s = String(nombre).trim();
  // Reemplazar espacios y guiones múltiples por _
  s = s.replace(/\s+/g, "_");
  // Mantener letras (incl. acentos), números, _ y -
  s = s.replace(/[^a-zA-Z0-9\u00C0-\u024F_\-]/g, "_");
  s = s.replace(/__+/g, "_");
  s = s.replace(/^_+|_+$/g, "");
  return s;
}

/**
 * Genera el nombre de archivo del reporte PDF según convención:
 * [Fecha]_[Operación]_[Negocio] ej. "21ago26_HUB_Cedros_Distribución.pdf"
 * Fecha en formato corto DDmmmAA (ej. 21ago26) ya calculado
 * Usa el valor de operación/negocio tal cual viene del checklist (sin hardcode).
 */
function generarNombreReporte(fechaCorta, operacion, negocio) {
  var opCanon = canonicalizarOperacion(operacion);
  var op = sanitizarNombreArchivo(opCanon) || "Operacion";
  var neg = sanitizarNombreArchivo(negocio) || "Negocio";
  return fechaCorta + "_" + op + "_" + neg + ".pdf";
}

/**
 * Busca o crea una subcarpeta hija dentro de una carpeta padre.
 * Búsqueda insensible a mayúsculas/acentos (normalizarTextoClave) para evitar duplicados.
 * Si no existe, crea la carpeta con el nombre deseado exacto (trim).
 */
function obtenerOCrearCarpetaHija(carpetaPadre, nombreDeseado) {
  var nombreTrim = String(nombreDeseado).trim();
  if (!nombreTrim) nombreTrim = "SinNombre";
  var normDeseado = normalizarTextoClave(nombreTrim);
  try {
    var carpetas = carpetaPadre.getFolders();
    while (carpetas.hasNext()) {
      var c = carpetas.next();
      if (normalizarTextoClave(c.getName()) === normDeseado) {
        return c;
      }
    }
  } catch (e) {
    // Fallback a getFoldersByName si getFolders falla
    try {
      var byName = carpetaPadre.getFoldersByName(nombreTrim);
      if (byName.hasNext()) return byName.next();
    } catch (e2) {}
  }
  // No existe -> crear
  return carpetaPadre.createFolder(nombreTrim);
}

/**
 * Obtiene (o crea) la jerarquía completa para reportes:
 * Raíz (ID_CARPETA_REPORTES) → Año (ej. 2026) → Negocio (ej. Fedex) → Operación (ej. Outbound)
 * Retorna la carpeta de Operación donde se debe guardar el PDF.
 * Usa fechaActual para extraer el año; si no se provee usa año actual.
 */
function obtenerCarpetaJerarquicaReporte(negocio, operacion, fechaActual) {
  var rootId = CONFIG.ID_CARPETA_REPORTES;
  if (!rootId) {
    Logger.log("CONFIG.ID_CARPETA_REPORTES no configurado");
    return null;
  }
  var carpetaRoot;
  try {
    carpetaRoot = DriveApp.getFolderById(rootId);
  } catch (e) {
    Logger.log("Error al abrir carpeta raíz de reportes " + rootId + ": " + e.toString());
    return null;
  }
  var anio = "2026";
  try {
    if (fechaActual && fechaActual.getFullYear) {
      anio = String(fechaActual.getFullYear());
    } else {
      anio = String(new Date().getFullYear());
    }
  } catch (eA) {
    anio = String(new Date().getFullYear());
  }
  var carpetaAnio = obtenerOCrearCarpetaHija(carpetaRoot, anio);
  var nombreNegocio = String(negocio || "General").trim();
  var carpetaNegocio = obtenerOCrearCarpetaHija(carpetaAnio, nombreNegocio);
  var nombreOperacion = canonicalizarOperacion(operacion || "General");
  var carpetaOperacion = obtenerOCrearCarpetaHija(carpetaNegocio, nombreOperacion);
  return carpetaOperacion;
}

/**
 * Guarda el blob PDF del reporte en la jerarquía Año/Negocio/Operación
 * y retorna el archivo creado (DriveApp File) o null si falla.
 * Aplica nomenclatura [Fecha]_[Operación]_[Negocio].pdf
 */
function guardarReportePDFEnJerarquia(blobPdf, negocio, operacion, fechaActual) {
  try {
    if (!blobPdf) return null;
    var fechaCorta = "";
    try {
      fechaCorta = formatearFechaCortaArchivo(fechaActual);
    } catch (eF) {
      var d = fechaActual || new Date();
      var dia = ("0" + d.getDate()).slice(-2);
      var meses = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];
      fechaCorta = dia + (meses[d.getMonth()] || "ene") + String(d.getFullYear()).slice(-2);
    }
    var nombreArchivo = generarNombreReporte(fechaCorta, operacion, negocio);
    blobPdf.setName(nombreArchivo);
    var carpetaDestino = obtenerCarpetaJerarquicaReporte(negocio, operacion, fechaActual);
    if (!carpetaDestino) {
      Logger.log("No se pudo obtener carpeta jerárquica para " + negocio + "/" + operacion);
      return null;
    }
    var archivo = carpetaDestino.createFile(blobPdf);
    // Intentar permisos heredados; no bloquear si falla
    try { archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (eS) {}
    return archivo;
  } catch (e) {
    Logger.log("Error en guardarReportePDFEnJerarquia: " + e.toString());
    return null;
  }
}

/**
 * Función auxiliar para obtener los colores correspondientes según el porcentaje de cumplimiento
 */
function obtenerEstiloCumplimiento(pct) {
  var valor = parseFloat(pct || 0);
  if (valor < 60) {
    return {
      deep: CONFIG.COLOR_RED_DEEP,
      pale: CONFIG.COLOR_RED_PALE
    };
  } else if (valor < 80) {
    return {
      deep: CONFIG.COLOR_YELLOW_DEEP,
      pale: CONFIG.COLOR_YELLOW_PALE
    };
  } else {
    return {
      deep: CONFIG.COLOR_GREEN_DEEP,
      pale: CONFIG.COLOR_GREEN_PALE
    };
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    CONFIG: CONFIG,
    normalizarTextoClave: normalizarTextoClave,
    canonicalizarOperacion: typeof canonicalizarOperacion !== 'undefined' ? canonicalizarOperacion : undefined,
    sanitizarNombreArchivo: typeof sanitizarNombreArchivo !== 'undefined' ? sanitizarNombreArchivo : undefined,
    generarNombreReporte: typeof generarNombreReporte !== 'undefined' ? generarNombreReporte : undefined,
    obtenerOCrearCarpetaHija: typeof obtenerOCrearCarpetaHija !== 'undefined' ? obtenerOCrearCarpetaHija : undefined,
    obtenerCarpetaJerarquicaReporte: typeof obtenerCarpetaJerarquicaReporte !== 'undefined' ? obtenerCarpetaJerarquicaReporte : undefined,
    guardarReportePDFEnJerarquia: typeof guardarReportePDFEnJerarquia !== 'undefined' ? guardarReportePDFEnJerarquia : undefined,
    obtenerCarpetaDestino: typeof obtenerCarpetaDestino !== 'undefined' ? obtenerCarpetaDestino : undefined,
    obtenerEstiloCumplimiento: obtenerEstiloCumplimiento
  };
}
