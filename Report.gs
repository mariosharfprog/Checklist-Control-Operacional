/**
 * SISTEMA DE INSPECCIÓN DE PROCESOS - SHARF
 * Archivo: Report.gs — Procesamiento de inspección, guardado y envío de correo/PDF
 */

// Resolución de calcularPorcentajesS / calcularPctActividad5S (Data.gs), compatible con Node
// y Apps Script, siguiendo el mismo patrón usado en dashboard/InformeSemanalModelo.js.
// No se reimplementan las reglas de cálculo de 5S: viven únicamente en Data.js.
var _dataMod5S = (typeof module !== 'undefined' && module.exports)
  ? require('./Data.js')
  : {
      calcularPorcentajesS: typeof calcularPorcentajesS !== 'undefined' ? calcularPorcentajesS : null,
      calcularPctActividad5S: typeof calcularPctActividad5S !== 'undefined' ? calcularPctActividad5S : null
    };

function _calcularPorcentajesS(items5S, estados) {
  if (_dataMod5S && typeof _dataMod5S.calcularPorcentajesS === 'function') return _dataMod5S.calcularPorcentajesS(items5S, estados);
  if (typeof calcularPorcentajesS === 'function') return calcularPorcentajesS(items5S, estados);
  return {};
}
function _calcularPctActividad5S(resumenPorS) {
  if (_dataMod5S && typeof _dataMod5S.calcularPctActividad5S === 'function') return _dataMod5S.calcularPctActividad5S(resumenPorS);
  if (typeof calcularPctActividad5S === 'function') return calcularPctActividad5S(resumenPorS);
  return null;
}

// Resolución de Memo.js (memorandos de amonestación por 3era Alerta), compatible con Node y Apps Script.
var _memoMod = (typeof module !== 'undefined' && module.exports)
  ? require('./Memo.js')
  : null;

function _extraerResponsablesConAlertaMemo(valorSheet, responsableStr, responsableEmailStr) {
  if (_memoMod && typeof _memoMod.extraerResponsablesConAlerta === 'function') return _memoMod.extraerResponsablesConAlerta(valorSheet, responsableStr, responsableEmailStr);
  if (typeof extraerResponsablesConAlerta === 'function') return extraerResponsablesConAlerta(valorSheet, responsableStr, responsableEmailStr);
  return [];
}
function _obtenerFallasCicloActualMemo(registrosPrevios, negocio, operacion, idSubKey, nombreResp, cols) {
  if (_memoMod && typeof _memoMod.obtenerFallasCicloActual === 'function') return _memoMod.obtenerFallasCicloActual(registrosPrevios, negocio, operacion, idSubKey, nombreResp, cols);
  if (typeof obtenerFallasCicloActual === 'function') return obtenerFallasCicloActual(registrosPrevios, negocio, operacion, idSubKey, nombreResp, cols);
  return [];
}
function _memoSeleccionadoParaEnvioMemo(nivel, nombre, memoOmitir, memoActivar) {
  if (_memoMod && typeof _memoMod.memoSeleccionadoParaEnvio === 'function') return _memoMod.memoSeleccionadoParaEnvio(nivel, nombre, memoOmitir, memoActivar);
  if (typeof memoSeleccionadoParaEnvio === 'function') return memoSeleccionadoParaEnvio(nivel, nombre, memoOmitir, memoActivar);
  return false;
}
function _generarYEnviarMemorandosMemo(amonestaciones, contexto) {
  if (_memoMod && typeof _memoMod.generarYEnviarMemorandos === 'function') return _memoMod.generarYEnviarMemorandos(amonestaciones, contexto);
  if (typeof generarYEnviarMemorandos === 'function') return generarYEnviarMemorandos(amonestaciones, contexto);
}
function _registrarAmonestacionesPendientesMemo(amonestaciones, contexto) {
  if (_memoMod && typeof _memoMod.registrarAmonestacionesPendientes === 'function') return _memoMod.registrarAmonestacionesPendientes(amonestaciones, contexto);
  if (typeof registrarAmonestacionesPendientes === 'function') return registrarAmonestacionesPendientes(amonestaciones, contexto);
}

/**
 * Transforma el id de una evidencia de 5S para su nombre de archivo: idGrupo + "_" + idPunto.
 * "7.2.11" -> "7.2_11" (solo cuando el id tiene EXACTAMENTE 3 segmentos numéricos).
 * Cualquier otro id (subactividades normales "1.3", grupos sin desglose "7.4", etc.) no cambia.
 * Transformación de string pura: no consulta el checklist.
 */
function transformarIdEvidencia5S(idSub) {
  var texto = String(idSub == null ? "" : idSub);
  var partes = texto.split(".");
  var esId5S = partes.length === 3 && partes.every(function(p) { return /^\d+$/.test(p); });
  if (esId5S) {
    return partes[0] + "." + partes[1] + "_" + partes[2];
  }
  return texto;
}

/**
 * Clasifica un item del checklist (ya resuelto por idSubactividad) según su destino de guardado:
 * "NORMAL"      -> subactividad de macroactividad normal, va a Historial_Inspecciones (sin cambios).
 * "PUNTO_5S"    -> punto MANUAL de la macro 5S, va a Historial_5S.
 * "DERIVADA_5S" -> placeholder de S5: no es un punto, no genera fila en Historial_5S
 *                  (queda representado por su fila de resumen de S en Historial_Inspecciones).
 */
function obtenerDestinoEvaluacion5S(itemInfo) {
  if (!itemInfo || !itemInfo.tipoS) return "NORMAL";
  if (itemInfo.tipoS === "DERIVADA") return "DERIVADA_5S";
  return "PUNTO_5S";
}

/**
 * Construye la fila (array, columnas A..Q) para una fila de punto en Historial_5S.
 * Reutiliza exactamente el mismo criterio de formateo de Evidencias/URLs/Tipo que
 * Historial_Inspecciones para una subactividad normal.
 */
function construirFilaHistorial5S(fechaActual, negocio, operacion, itemInfo, evalData, fotosList, textoUrls, resultadoAlerta, tipoEnvio, usuarioActivo) {
  return [
    fechaActual,
    negocio || "",
    operacion || "",
    String(itemInfo.idActividad || ""),          // ID_5S
    String(itemInfo.idGrupo || ""),              // ID_S
    itemInfo.grupo || "",                        // S
    String(itemInfo.idPunto || ""),              // ID_Punto
    itemInfo.area || "",                         // Área
    itemInfo.subactividadCorta || "",            // Punto
    evalData.estado || "",                       // Estado
    evalData.comentario || "",                   // Detalle
    (fotosList && fotosList.length > 0) ? fotosList.length + " Foto(s)" : "Sin evidencia", // Evidencias
    textoUrls,                                   // URLs
    evalData.responsable || "",                  // Responsable
    tipoEnvio === "prueba" ? "PRUEBA" : "OFICIAL", // Tipo
    resultadoAlerta.valorSheet,                  // Alerta
    usuarioActivo                                // Inspector / Usuario
  ];
}

/**
 * Construye las 5 filas de resumen por S (una por sIndex) que van a Historial_Inspecciones
 * bajo la macro 5S. NO agrega columnas: usa exactamente las 17 columnas ya existentes.
 * resumenS: resultado de calcularPorcentajesS (idGrupo -> {sIndex, grupo, pct, estado, ...}).
 * ctx: { fechaActual, negocio, operacion, tipoEnvio, usuarioActivo, cumploGenFormateado, pctFinalDecimal }
 * El texto largo de cada S se toma con el siguiente orden de preferencia:
 * 1. textoS del primer item de ese grupo (expuesto por Data.js desde cl_5s)
 * 2. it.subactividad del item con idPunto === "" (fallback genérico / DERIVADA)
 * 3. r.grupo como último recurso defensivo (nombre corto)
 */
function construirFilasResumen5S(items5SOperacion, resumenS, ctx) {
  var textoSPorGrupo = {};
  var textoPorGrupo = {};
  (items5SOperacion || []).forEach(function(it) {
    if (!it || !it.idGrupo) return;
    if (it.textoS && !textoSPorGrupo[it.idGrupo]) {
      textoSPorGrupo[it.idGrupo] = it.textoS;
    }
    if (it.idPunto === "" && !textoPorGrupo[it.idGrupo]) {
      textoPorGrupo[it.idGrupo] = it.subactividad;
    }
  });

  var idGrupos = Object.keys(resumenS || {}).sort(function(a, b) {
    return resumenS[a].sIndex - resumenS[b].sIndex;
  });

  var idActividad = (items5SOperacion && items5SOperacion.length > 0) ? items5SOperacion[0].idActividad : "";

  return idGrupos.map(function(idGrupo) {
    var r = resumenS[idGrupo];
    var subactividadTexto = textoSPorGrupo[idGrupo] || textoPorGrupo[idGrupo] || r.grupo;
    return [
      ctx.fechaActual,
      ctx.negocio || "",
      ctx.operacion || "",
      idActividad,                                     // ID_Actividad
      "5S",                                             // Actividad
      idGrupo,                                          // ID_Subactividad (7.1..7.5)
      subactividadTexto,                                // Subactividad (texto completo)
      r.estado,                                         // Estado
      "",                                               // Detalle
      "-",                                              // Evidencias
      "-",                                              // URLs
      "",                                               // Responsable (vacío: las filas de S no tienen responsable)
      ctx.tipoEnvio === "prueba" ? "PRUEBA" : "OFICIAL",// Tipo
      "-",                                              // Alerta (las filas de S no participan del ciclo)
      ctx.cumploGenFormateado,                          // % Cumplimiento General
      ctx.pctFinalDecimal,                              // % Cumplimiento Actividad (0..1, mismo valor en las 5)
      ctx.usuarioActivo                                 // Inspector / Usuario
    ];
  });
}

/**
 * Construye los "detalle" sintéticos (uno por S, intercalando sus puntos en NO CUMPLE)
 * para el bloque 5S del PDF/correo, con la misma forma que consumen las tablas existentes
 * (idSub/subactividad/estado/comentario/responsable/fotos/alerta/es5SPunto).
 * - Cada S va con es5SPunto: false
 * - Debajo de cada S van sus puntos en NO CUMPLE (ordenados por idPunto numérico ascendente),
 *   con es5SPunto: true
 * - Si una S no tiene puntos en NO CUMPLE, no se agrega ninguna fila extra bajo ella.
 */
function construirDetallesPdf5S(resumenS, puntosPorGrupo) {
  puntosPorGrupo = puntosPorGrupo || {};
  var idGrupos = Object.keys(resumenS || {}).sort(function(a, b) {
    return (resumenS[a].sIndex || 0) - (resumenS[b].sIndex || 0);
  });

  var listaDetalles = [];

  idGrupos.forEach(function(idGrupo) {
    var r = resumenS[idGrupo];
    var pctTexto = (r.pct != null && !isNaN(r.pct)) ? (Number(r.pct).toFixed(2) + "%") : "N/A";

    // Fila de la S
    listaDetalles.push({
      idSub: r.idGrupo || idGrupo,
      subactividad: "S" + r.sIndex + " — " + (r.grupo || "") + " (" + pctTexto + ")",
      estado: r.estado,
      comentario: "-",
      responsable: "-",
      fotos: [],
      alerta: "-",
      es5SPunto: false
    });

    // Puntos de esta S en NO CUMPLE
    var puntosS = puntosPorGrupo[idGrupo] || [];
    var puntosNoCumple = puntosS.filter(function(pt) {
      return pt.estado === "NO CUMPLE";
    });

    // Orden numérico ascendente por idPunto
    puntosNoCumple.sort(function(a, b) {
      var numA = parseFloat(a.idPunto);
      var numB = parseFloat(b.idPunto);
      if (!isNaN(numA) && !isNaN(numB) && numA !== numB) {
        return numA - numB;
      }
      return String(a.idPunto || "").localeCompare(String(b.idPunto || ""), undefined, { numeric: true });
    });

    puntosNoCumple.forEach(function(pt) {
      listaDetalles.push({
        idSub: pt.idSub || (pt.idPunto ? (idGrupo + "." + pt.idPunto) : idGrupo),
        subactividad: pt.subactividad || pt.subactividadCorta || "",
        estado: "NO CUMPLE",
        comentario: (pt.comentario != null && pt.comentario !== "") ? pt.comentario : "-",
        responsable: (pt.responsable != null && pt.responsable !== "") ? pt.responsable : "-",
        fotos: pt.fotos || [],
        alerta: (pt.alerta != null && pt.alerta !== "") ? pt.alerta : "-",
        es5SPunto: true
      });
    });
  });

  return listaDetalles;
}

function procesarInspeccion(payloadJson) {
  try {
    if (!payloadJson) {
      return { exito: false, mensaje: "Payload vacío recibido en el servidor." };
    }

    var payload = JSON.parse(payloadJson);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timeZone = ss.getSpreadsheetTimeZone() || "GMT-5";
    
    // Identificar el usuario emisor real de la sesión activa
    var usuarioActivo = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "Sesión activa";
    
    var cumploGenRaw = parseFloat(payload.cumplimientoGeneral);
    var tieneCumplimientoGeneral = isFinite(cumploGenRaw);
    var cumploGen = tieneCumplimientoGeneral ? cumploGenRaw : 0;
    var textoCumplimientoGeneral = tieneCumplimientoGeneral ? (cumploGen.toFixed(2) + "%") : "-";
    var cumploGenFormateado = tieneCumplimientoGeneral ? (cumploGenRaw / 100) : "-";
    // Fecha del checklist: usa payload.fechaChecklist (YYYY-MM-DD), solo fecha, nunca hora. Por defecto hoy si no viene.
    var fechaStr = payload.fechaChecklist ? String(payload.fechaChecklist).trim() : "";
    var fechaActual;
    if (fechaStr && /^\d{4}-\d{2}-\d{2}$/.test(fechaStr)) {
      var parts = fechaStr.split("-");
      fechaActual = new Date(parseInt(parts[0],10), parseInt(parts[1],10)-1, parseInt(parts[2],10), 12, 0, 0);
    } else {
      fechaActual = new Date();
      fechaActual.setHours(12,0,0,0);
    }
    // Validación: no permitir fecha futura
    var hoyLimite = new Date(); hoyLimite.setHours(12,0,0,0);
    if (fechaActual > hoyLimite) fechaActual = hoyLimite;
    var fechaCortaArchivo = formatearFechaCortaArchivo(fechaActual);

    // 1. Gestión Dinámica de Carpeta en Google Drive por Negocio y Operación
    var carpetaEvidencias = obtenerCarpetaDestino(payload.negocio, payload.operacion);

    var mapaUrlsFotos = {};

    // 2. Guardar Imágenes Base64 de Subactividades con Nomenclatura Estándar: Evidencia_[ID]_[Fecha].jpg
    if (payload.imagenes && carpetaEvidencias) {
      for (var idSub in payload.imagenes) {
        var listaFotosBase64 = payload.imagenes[idSub];
        mapaUrlsFotos[idSub] = [];
        
        if (Array.isArray(listaFotosBase64) && listaFotosBase64.length > 0) {
          listaFotosBase64.forEach(function(base64Data, index) {
            if (!base64Data) return;
            try {
              var partes = String(base64Data).split(",");
              var contenidoLimpio = partes.length > 1 ? partes[1] : partes[0];
              contenidoLimpio = contenidoLimpio.replace(/\s+/g, '');
              
              var mimeType = "image/jpeg";
              if (partes.length > 1 && partes[0].indexOf(":") > -1 && partes[0].indexOf(";") > -1) {
                mimeType = partes[0].split(":")[1].split(";")[0];
                if(!mimeType || mimeType.trim()==="") mimeType = "image/jpeg";
              }
              var ext = ".jpg";
              if(mimeType.indexOf("heic")!==-1 || mimeType.indexOf("heif")!==-1) ext = ".heic";
              else if(mimeType.indexOf("png")!==-1) ext = ".png";
              else if(mimeType.indexOf("webp")!==-1) ext = ".webp";

              // Para puntos de 5S (id de 3 segmentos, ej. "7.2.11") el nombre usa "x.y_z"
              var idSubEvidencia = transformarIdEvidencia5S(idSub);
              var nombreFoto = listaFotosBase64.length > 1
                ? "Evidencia_" + idSubEvidencia + "_" + fechaCortaArchivo + "_" + (index + 1) + ext
                : "Evidencia_" + idSubEvidencia + "_" + fechaCortaArchivo + ext;

              var blob = Utilities.newBlob(
                Utilities.base64Decode(contenidoLimpio), 
                mimeType, 
                nombreFoto
              );
              var archivo = carpetaEvidencias.createFile(blob);
              var urlArchivo = archivo.getUrl();
              
              // Guardar la URL de inmediato para asegurar que figure en el reporte
              mapaUrlsFotos[idSub].push(urlArchivo);

              // Intentar configurar permisos públicos o de dominio de manera aislada (no bloqueante)
              try {
                archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              } catch (eShare) {
                try {
                  archivo.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
                } catch (eDom) {
                  Logger.log("Aviso: No se pudo modificar setSharing en " + urlArchivo + " (hereda permisos de carpeta): " + eDom.toString());
                }
              }

              // Pequeña pausa preventiva para evitar límites de tasa al procesar múltiples archivos
              Utilities.sleep(80);
            } catch (errImg) {
              Logger.log("Error procesando imagen " + (index + 1) + " en ID " + idSub + ": " + errImg.toString());
            }
          });
        }
      }
    }

    // 2.1 Procesar Fotos de Comentarios Extra con Nomenclatura: Evidencia_CE_[Num]_[Fecha].jpg
    var comentariosExtraProcesados = [];
    if (Array.isArray(payload.comentariosExtra) && payload.comentariosExtra.length > 0) {
      payload.comentariosExtra.forEach(function(ceItem, idxCE) {
        var numCE = idxCE + 1;
        var textoCE = (ceItem.comentario || "").trim();
        var listaFotosCE = ceItem.imagenes || [];
        var urlsCE = [];
        var fileIdsCE = [];
        var base64CE = [];

        if (Array.isArray(listaFotosCE) && listaFotosCE.length > 0 && carpetaEvidencias) {
          listaFotosCE.forEach(function(base64Data, index) {
            if (!base64Data) return;
            try {
              var partes = String(base64Data).split(",");
              var contenidoLimpio = partes.length > 1 ? partes[1] : partes[0];
              contenidoLimpio = contenidoLimpio.replace(/\s+/g, '');
              
              var mimeType = "image/jpeg";
              if (partes.length > 1 && partes[0].indexOf(":") > -1 && partes[0].indexOf(";") > -1) {
                mimeType = partes[0].split(":")[1].split(";")[0];
                if(!mimeType || mimeType.trim()==="") mimeType = "image/jpeg";
              }
              var extCE = ".jpg";
              if(mimeType.indexOf("heic")!==-1 || mimeType.indexOf("heif")!==-1) extCE = ".heic";
              else if(mimeType.indexOf("png")!==-1) extCE = ".png";
              else if(mimeType.indexOf("webp")!==-1) extCE = ".webp";

              var nombreFotoCE = listaFotosCE.length > 1
                ? "Evidencia_CE_" + numCE + "_" + fechaCortaArchivo + "_" + (index + 1) + extCE
                : "Evidencia_CE_" + numCE + "_" + fechaCortaArchivo + extCE;

              var blob = Utilities.newBlob(
                Utilities.base64Decode(contenidoLimpio), 
                mimeType, 
                nombreFotoCE
              );
              var archivo = carpetaEvidencias.createFile(blob);
              urlsCE.push(archivo.getUrl());
              fileIdsCE.push(archivo.getId());
              base64CE.push(base64Data);

              try {
                archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              } catch (eShare) {
                try {
                  archivo.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
                } catch (eDom) {
                  Logger.log("Aviso setSharing foto CE: " + eDom.toString());
                }
              }
              Utilities.sleep(80);
            } catch (errImgCE) {
              Logger.log("Error procesando imagen CE " + numCE + ": " + errImgCE.toString());
            }
          });
        }

        var correosCopiaCE = (ceItem.correosCopia || "").trim();
        if (textoCE !== "" || urlsCE.length > 0 || correosCopiaCE !== "") {
          comentariosExtraProcesados.push({
            numero: numCE,
            comentario: textoCE,
            correosCopia: correosCopiaCE,
            urls: urlsCE,
            fileIds: fileIdsCE,
            base64List: base64CE
          });
        }
      });
    }

    // 3. Verificación / Creación de la hoja de Historial
    var sheetHistorial = ss.getSheetByName(CONFIG.HOJA_HISTORIAL);
    if (!sheetHistorial) {
      sheetHistorial = ss.insertSheet(CONFIG.HOJA_HISTORIAL);
      sheetHistorial.appendRow([
        'Fecha', 'Negocio', 'Operación', 'ID_Actividad', 'Actividad', 
        'ID_Subactividad', 'Subactividad', 'Estado', 'Detalle', 
        'Evidencias', 'URLs', 'Responsable', 'Tipo', 'Alerta', '% Cumplimiento General', '% Cumplimiento Actividad', 'Inspector / Usuario'
      ]);
      sheetHistorial.getRange("1:1").setFontWeight("bold").setBackground(CONFIG.COLOR_LINE_SOFT).setFontColor(CONFIG.COLOR_INK);
    }

    // Forzar formato texto en columnas de ID para evitar que Sheets interprete "7.1" como fecha
    try {
      var maxFilas = sheetHistorial.getMaxRows();
      sheetHistorial.getRange(2, 4, maxFilas - 1, 1).setNumberFormat('@');
      sheetHistorial.getRange(2, 6, maxFilas - 1, 1).setNumberFormat('@');
    } catch (eFmt) {
      Logger.log('Aviso: no se pudo forzar formato texto en Historial_Inspecciones: ' + eFmt.toString());
    }

    var datosChecklist = cargarDatosDesdeSheet().checklist || [];
    var registrosPrevios = sheetHistorial.getDataRange().getValues();

    function normalizarFecha(d) {
      var dateObj = new Date(d);
      return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
    }

    var fechaHoyLimpia = normalizarFecha(fechaActual);
    // El negocio/operación no se hardcodea: se toma tal cual del checklist (fuente única de verdad)
    // Solo se normaliza (mayúsculas + sin tildes) para comparación insensible a formato
    var negocioSeleccionado = (typeof normalizarTextoClave === 'function' && payload.negocio) ? normalizarTextoClave(payload.negocio) : (payload.negocio ? String(payload.negocio).trim().toUpperCase() : "");
    var operacionSeleccionada = "";
    if (payload.operacion) {
      var opCanonSel = (typeof canonicalizarOperacion === 'function') ? canonicalizarOperacion(payload.operacion) : String(payload.operacion).trim();
      operacionSeleccionada = (typeof normalizarTextoClave === 'function') ? normalizarTextoClave(opCanonSel) : String(opCanonSel).trim().toUpperCase();
    }

    // 4. Pre-procesar Actividades
    var resumenActividades = {};
    var listaEvaluacionesProcesadas = [];
    var puntos5SEvaluados = []; // items con tipoS MANUAL -> van a Historial_5S (no a Historial_Inspecciones)

    for (var idSubKey in payload.evaluaciones) {
      var evalData = payload.evaluaciones[idSubKey];
      if (!evalData || !evalData.estado) continue;

      var itemInfo = datosChecklist.find(function(c) {
        var cNeg = (typeof normalizarTextoClave === 'function') ? normalizarTextoClave(c.negocio) : String(c.negocio||"").trim().toUpperCase();
        var cOpRaw = c.operacion || "";
        var cOpCanon = (typeof canonicalizarOperacion === 'function') ? canonicalizarOperacion(cOpRaw) : String(cOpRaw).trim();
        var cOp = (typeof normalizarTextoClave === 'function') ? normalizarTextoClave(cOpCanon) : String(cOpCanon).trim().toUpperCase();
        return c.idSubactividad === idSubKey && cNeg === negocioSeleccionado && cOp === operacionSeleccionada;
      });

      if (!itemInfo) {
        itemInfo = { negocio: payload.negocio, operacion: payload.operacion, idActividad: "", actividad: "", subactividad: idSubKey };
      }

      // Clasificación por destino: 5S se procesa aparte (punto->Historial_5S; derivada->no genera fila de punto)
      var destino5S = obtenerDestinoEvaluacion5S(itemInfo);
      if (destino5S === "DERIVADA_5S") continue;
      if (destino5S === "PUNTO_5S") {
        puntos5SEvaluados.push({ idSubKey: idSubKey, evalData: evalData, itemInfo: itemInfo });
        continue;
      }

      var actNombre = limpiarNombreActividad(itemInfo.actividad);

      if (!resumenActividades[actNombre]) {
        resumenActividades[actNombre] = { cumple: 0, noCumple: 0, na: 0, total: 0, detalles: [] };
      }

      var estadoActual = evalData.estado;
      if (estadoActual === "CUMPLE") {
        resumenActividades[actNombre].cumple++;
        resumenActividades[actNombre].total++;
      } else if (estadoActual === "NO CUMPLE") {
        resumenActividades[actNombre].noCumple++;
        resumenActividades[actNombre].total++;
      } else if (estadoActual === "N/A") {
        resumenActividades[actNombre].na++;
      }

      listaEvaluacionesProcesadas.push({
        idSubKey: idSubKey,
        evalData: evalData,
        itemInfo: itemInfo,
        actNombre: actNombre,
        estadoActual: estadoActual
      });
    }

    // Amonestaciones (3era Alerta) detectadas durante el guardado del historial (normal + 5S),
    // para generar y enviar memorandos DESPUÉS de terminar todas las escrituras (ver más abajo).
    // "Enviar Memos" apagado en el checklist (enviarMemos === false): no se detecta ni se genera
    // ningún memorando. Payloads antiguos sin el campo mantienen el comportamiento previo.
    var esEnvioOficialParaMemo = (payload.tipoEnvio !== "prueba")
      && (payload.noAlertar !== true)
      && (payload.enviarMemos !== false);
    var amonestacionesDetectadas = [];

    // 5. Guardar Filas en Historial
    listaEvaluacionesProcesadas.forEach(function(item) {
      var idSubKey = item.idSubKey;
      var evalData = item.evalData;
      var itemInfo = item.itemInfo;
      var actNombre = item.actNombre;
      var estadoActual = item.estadoActual;

      var fotosList = mapaUrlsFotos[idSubKey] || [];
      var textoUrlsColumnaJ = construirTextoUrls(fotosList);

      // Calcular alerta específica por Responsable + Subactividad
      var resultadoAlerta = calcularAlertaResponsable(
        registrosPrevios, 
        payload.negocio, 
        payload.operacion, 
        idSubKey, 
        evalData.responsable, 
        estadoActual,
        payload.noAlertar === true,
        COLS_HISTORIAL_INSPECCIONES,
        payload.tipoEnvio
      );

      if (esEnvioOficialParaMemo) {
        var amonestacionesItem = _extraerResponsablesConAlertaMemo(resultadoAlerta.valorSheet, evalData.responsable, evalData.responsableEmail);
        amonestacionesItem.forEach(function(am) {
          // 3era alerta: se prepara el memo salvo que el inspector la apague (bandera roja);
          // 1era/2da alerta: solo se prepara si el inspector la activa (bandera blanca).
          if (!_memoSeleccionadoParaEnvioMemo(am.nivel, am.nombre, evalData.memoOmitir, evalData.memoActivar)) return;
          amonestacionesDetectadas.push({
            negocio: payload.negocio,
            operacion: payload.operacion,
            idSubKey: idSubKey,
            actividad: actNombre,
            subactividad: itemInfo.subactividad || idSubKey,
            fecha: fechaActual,
            detalle: evalData.comentario || "",
            evidencias: fotosList,
            nombre: am.nombre,
            email: am.email,
            nivelAlerta: am.nivel,
            fallasPrevias: _obtenerFallasCicloActualMemo(registrosPrevios, payload.negocio, payload.operacion, idSubKey, am.nombre, COLS_HISTORIAL_INSPECCIONES)
          });
        });
      }

      var infoAct = resumenActividades[actNombre];
      var pctActividadDecimal = infoAct.total > 0 ? (infoAct.cumple / infoAct.total) : 0;

      resumenActividades[actNombre].detalles.push({
        idSub: itemInfo.idSubactividad || idSubKey,
        subactividad: itemInfo.subactividad || idSubKey,
        estado: estadoActual,
        comentario: evalData.comentario || "-",
        responsable: evalData.responsable || "-",
        fotos: fotosList,
        alerta: resultadoAlerta.textoPdf
      });

      // appendRow es atómico (evita que envíos simultáneos se sobrescriban); el formato texto de las
      // columnas de ID ya se aplicó una sola vez a toda la hoja antes del bucle.
      sheetHistorial.appendRow([
        fechaActual,
        payload.negocio || "",
        payload.operacion || "",
        String(itemInfo.idActividad || ""),
        actNombre,
        String(idSubKey || ""),
        itemInfo.subactividad || idSubKey,
        estadoActual,
        evalData.comentario || "",
        fotosList.length > 0 ? fotosList.length + " Foto(s)" : "Sin evidencia",
        textoUrlsColumnaJ, 
        evalData.responsable || "",
        payload.tipoEnvio === "prueba" ? "PRUEBA" : "OFICIAL",
        resultadoAlerta.valorSheet,
        cumploGenFormateado, 
        pctActividadDecimal,
        usuarioActivo
      ]);

      var ultimaFilaInsertada = sheetHistorial.getLastRow();
      sheetHistorial.getRange(ultimaFilaInsertada, 11).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    });

    // 5.b Guardar puntos de 5S en Historial_5S + 5 filas de resumen por S en Historial_Inspecciones
    if (puntos5SEvaluados.length > 0) {
      var sheetHistorial5S = ss.getSheetByName(CONFIG.HOJA_HISTORIAL_5S);
      if (!sheetHistorial5S) {
        sheetHistorial5S = ss.insertSheet(CONFIG.HOJA_HISTORIAL_5S);
        sheetHistorial5S.appendRow([
          'Fecha', 'Negocio', 'Operación', 'ID_5S', 'ID_S', 'S', 'ID_Punto', 'Área', 'Punto',
          'Estado', 'Detalle', 'Evidencias', 'URLs', 'Responsable', 'Tipo', 'Alerta', 'Inspector / Usuario'
        ]);
        sheetHistorial5S.getRange("1:1").setFontWeight("bold").setBackground(CONFIG.COLOR_LINE_SOFT).setFontColor(CONFIG.COLOR_INK);
      }

      // Corregir cabecera si existe pero no coincide (bug: se copió de Historial_Inspecciones)
      var cabecera5SEsperada = [
        'Fecha', 'Negocio', 'Operación', 'ID_5S', 'ID_S', 'S', 'ID_Punto', 'Área', 'Punto',
        'Estado', 'Detalle', 'Evidencias', 'URLs', 'Responsable', 'Tipo', 'Alerta', 'Inspector / Usuario'
      ];
      try {
        var fila1Actual = sheetHistorial5S.getRange(1, 1, 1, cabecera5SEsperada.length).getValues()[0];
        var cabeceraCorrecta = cabecera5SEsperada.every(function(col, idx) {
          return String(fila1Actual[idx] || '').trim() === col;
        });
        if (!cabeceraCorrecta) {
          sheetHistorial5S.getRange(1, 1, 1, cabecera5SEsperada.length).setValues([cabecera5SEsperada]);
          sheetHistorial5S.getRange('1:1').setFontWeight('bold').setBackground(CONFIG.COLOR_LINE_SOFT).setFontColor(CONFIG.COLOR_INK);
        }
      } catch (eCab) {
        Logger.log('Aviso: no se pudo verificar/corregir cabecera de Historial_5S: ' + eCab.toString());
      }

      // Forzar formato texto en columnas de ID de Historial_5S
      try {
        var maxFilas5S = sheetHistorial5S.getMaxRows();
        sheetHistorial5S.getRange(2, 4, maxFilas5S - 1, 1).setNumberFormat('@');
        sheetHistorial5S.getRange(2, 5, maxFilas5S - 1, 1).setNumberFormat('@');
        sheetHistorial5S.getRange(2, 7, maxFilas5S - 1, 1).setNumberFormat('@');
      } catch (eFmt5S) {
        Logger.log('Aviso: no se pudo forzar formato texto en Historial_5S: ' + eFmt5S.toString());
      }

      var registrosPrevios5S = sheetHistorial5S.getDataRange().getValues();
      var estadosPlanos5S = {};
      puntos5SEvaluados.forEach(function(pt) {
        estadosPlanos5S[pt.itemInfo.idSubactividad || pt.idSubKey] = pt.evalData.estado;
      });

      var puntosPorGrupo = {};

      puntos5SEvaluados.forEach(function(pt) {
        var fotosList5S = mapaUrlsFotos[pt.idSubKey] || [];
        var textoUrls5S = construirTextoUrls(fotosList5S);
        // Identidad del punto para el ciclo de alertas: ID_S + "." + ID_Punto (idSubactividad ya viene así armado)
        var identidadPunto = pt.itemInfo.idSubactividad || pt.idSubKey;

        var resultadoAlerta5S = calcularAlertaResponsable(
          registrosPrevios5S,
          payload.negocio,
          payload.operacion,
          identidadPunto,
          pt.evalData.responsable,
          pt.evalData.estado,
          payload.noAlertar === true,
          COLS_HISTORIAL_5S,
          payload.tipoEnvio
        );

        if (esEnvioOficialParaMemo) {
          var amonestaciones5SItem = _extraerResponsablesConAlertaMemo(resultadoAlerta5S.valorSheet, pt.evalData.responsable, pt.evalData.responsableEmail);
          amonestaciones5SItem.forEach(function(am5S) {
            // 3era alerta: se prepara el memo salvo que el inspector la apague (bandera roja);
            // 1era/2da alerta: solo se prepara si el inspector la activa (bandera blanca).
            if (!_memoSeleccionadoParaEnvioMemo(am5S.nivel, am5S.nombre, pt.evalData.memoOmitir, pt.evalData.memoActivar)) return;
            amonestacionesDetectadas.push({
              negocio: payload.negocio,
              operacion: payload.operacion,
              idSubKey: identidadPunto,
              actividad: pt.itemInfo.grupo || "",
              subactividad: pt.itemInfo.subactividadCorta || pt.itemInfo.subactividad || "",
              fecha: fechaActual,
              detalle: pt.evalData.comentario || "",
              evidencias: fotosList5S,
              nombre: am5S.nombre,
              email: am5S.email,
              nivelAlerta: am5S.nivel,
              fallasPrevias: _obtenerFallasCicloActualMemo(registrosPrevios5S, payload.negocio, payload.operacion, identidadPunto, am5S.nombre, COLS_HISTORIAL_5S)
            });
          });
        }

        var fila5S = construirFilaHistorial5S(
          fechaActual, payload.negocio, payload.operacion, pt.itemInfo, pt.evalData,
          fotosList5S, textoUrls5S, resultadoAlerta5S, payload.tipoEnvio, usuarioActivo
        );
        sheetHistorial5S.appendRow(fila5S);
        var ultimaFila5SInsertada = sheetHistorial5S.getLastRow();
        sheetHistorial5S.getRange(ultimaFila5SInsertada, 13).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // URLs

        // Acumular datos de cada punto indexado por idGrupo para el reporte PDF/correo
        var idGrupo = pt.itemInfo.idGrupo;
        if (idGrupo) {
          if (!puntosPorGrupo[idGrupo]) {
            puntosPorGrupo[idGrupo] = [];
          }
          puntosPorGrupo[idGrupo].push({
            idSub: identidadPunto,
            idPunto: pt.itemInfo.idPunto,
            subactividad: pt.itemInfo.subactividadCorta || pt.itemInfo.subactividad || "",
            subactividadCorta: pt.itemInfo.subactividadCorta,
            estado: pt.evalData.estado,
            comentario: pt.evalData.comentario || "",
            responsable: pt.evalData.responsable || "",
            fotos: fotosList5S,
            alerta: resultadoAlerta5S.textoPdf
          });
        }
      });

      // Todos los items de 5S de esta (negocio, operación) -- evaluados o no -- para calcular % por S
      var items5SOperacion = datosChecklist.filter(function(c) {
        if (!c.tipoS) return false;
        var cNeg = (typeof normalizarTextoClave === 'function') ? normalizarTextoClave(c.negocio) : String(c.negocio||"").trim().toUpperCase();
        var cOpCanon = (typeof canonicalizarOperacion === 'function') ? canonicalizarOperacion(c.operacion || "") : String(c.operacion||"").trim();
        var cOp = (typeof normalizarTextoClave === 'function') ? normalizarTextoClave(cOpCanon) : String(cOpCanon).trim().toUpperCase();
        return cNeg === negocioSeleccionado && cOp === operacionSeleccionada;
      });

      var resumenS = _calcularPorcentajesS(items5SOperacion, estadosPlanos5S);
      var pctFinal5S = _calcularPctActividad5S(resumenS);
      var pctFinal5SDecimal = (pctFinal5S != null && !isNaN(pctFinal5S)) ? (pctFinal5S / 100) : 0;

      var filasResumen5S = construirFilasResumen5S(items5SOperacion, resumenS, {
        fechaActual: fechaActual,
        negocio: payload.negocio,
        operacion: payload.operacion,
        tipoEnvio: payload.tipoEnvio,
        usuarioActivo: usuarioActivo,
        cumploGenFormateado: cumploGenFormateado,
        pctFinalDecimal: pctFinal5SDecimal
      });
      filasResumen5S.forEach(function(fila) { sheetHistorial.appendRow(fila); });

      // Entrada sintética para el correo/PDF: 5 "detalles" (uno por S, intercalando sus puntos NO CUMPLE)
      // y el % real de la macroactividad (promedio de las 5 S), no cumple/total de puntos.
      var cumpleS = 0, totalS = 0;
      Object.keys(resumenS).forEach(function(k) {
        if (resumenS[k].estado === "CUMPLE") { cumpleS++; totalS++; }
        else if (resumenS[k].estado === "NO CUMPLE") { totalS++; }
      });
      resumenActividades["5S"] = {
        cumple: cumpleS,
        noCumple: totalS - cumpleS,
        na: 0,
        total: totalS,
        detalles: construirDetallesPdf5S(resumenS, puntosPorGrupo),
        pctOverride: pctFinal5S
      };
    }

    var ultimaFila = sheetHistorial.getLastRow();
    if (ultimaFila > 1) {
      sheetHistorial.getRange(2, 15, ultimaFila - 1, 2).setNumberFormat("0.00%");
    }

    // 5.0 Actualización automática instantánea de la hoja 'Porcentajes'
    try {
      if (typeof actualizarHojaPorcentajes === "function") {
        actualizarHojaPorcentajes(true);
      }
    } catch (ePorcentajes) {
      Logger.log("Aviso al auto-actualizar hoja Porcentajes: " + ePorcentajes.toString());
    }

    // 5.1 Guardar Comentarios Extra en Hoja Dedicada
    if (comentariosExtraProcesados.length > 0) {
      var sheetComentarios = ss.getSheetByName(CONFIG.HOJA_COMENTARIOS_EXTRA);
      if (!sheetComentarios) {
        sheetComentarios = ss.insertSheet(CONFIG.HOJA_COMENTARIOS_EXTRA);
        sheetComentarios.appendRow([
          'Fecha y Hora', 'Negocio', 'Operación', 'Comentario', 
          'Cantidad de Evidencias', 'Link Evidencia', 'Inspector / Usuario', 'Notificados (Copia)'
        ]);
        sheetComentarios.getRange("1:1").setFontWeight("bold").setBackground(CONFIG.COLOR_LINE_SOFT).setFontColor(CONFIG.COLOR_INK);
      } else if (sheetComentarios.getLastColumn() < 8) {
        sheetComentarios.getRange(1, 8).setValue('Notificados (Copia)').setFontWeight("bold").setBackground(CONFIG.COLOR_LINE_SOFT).setFontColor(CONFIG.COLOR_INK);
      }

      comentariosExtraProcesados.forEach(function(ce) {
        var textoUrlsCE = construirTextoUrls(ce.urls);
        sheetComentarios.appendRow([
          fechaActual,
          payload.negocio || "",
          payload.operacion || "",
          ce.comentario || "-",
          ce.urls.length > 0 ? ce.urls.length + " Foto(s)" : "Sin evidencia",
          textoUrlsCE,
          usuarioActivo,
          ce.correosCopia || "-"
        ]);
        var lastRowCE = sheetComentarios.getLastRow();
        sheetComentarios.getRange(lastRowCE, 6).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        sheetComentarios.getRange(lastRowCE, 8).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      });
    }

    // 6. Configurar Destinatarios Dinámicos
    var destinatarios = [];

    if (payload.tipoEnvio === "prueba") {
      destinatarios = obtenerCorreosPrueba(payload.negocio, payload.operacion);
    } else {
      destinatarios = payload.correosSeleccionados || [];
    }

    // Recolectar correos en copia desde los Comentarios Extra y Responsables
    // Para PRUEBA: aislamiento total — solo destinatarios PRUEBA, sin CC de oficiales/responsables
    var correosCopiaExtra = [];
    var correosResponsables = [];
    if (payload.tipoEnvio !== "prueba") {
      if (comentariosExtraProcesados.length > 0) {
        comentariosExtraProcesados.forEach(function(ce) {
          if (ce.correosCopia) {
            var listaCopia = String(ce.correosCopia).split(/[,;]+/);
            listaCopia.forEach(function(m) {
              var mailLimpio = m.trim().toLowerCase();
              if (mailLimpio && mailLimpio.indexOf("@") > -1 && !correosCopiaExtra.includes(mailLimpio) && !destinatarios.includes(mailLimpio)) {
                correosCopiaExtra.push(mailLimpio);
              }
            });
          }
        });
      }

      // Recolectar correos de Responsables desde evaluaciones (columna ahora es nombre corto, email en responsableEmail)
      if (payload.evaluaciones) {
        for (var idSubKey in payload.evaluaciones) {
          var evalResp = payload.evaluaciones[idSubKey];
          var respStr = (evalResp.responsableEmail && String(evalResp.responsableEmail).trim() !== "") ? evalResp.responsableEmail : (evalResp.responsable || "");
          var listaResp = String(respStr).split(/[,;]+/);
          listaResp.forEach(function(m) {
            var mail = m.trim().toLowerCase();
            if (mail && mail.indexOf("@") > -1 && !correosResponsables.includes(mail) && !destinatarios.includes(mail) && !correosCopiaExtra.includes(mail)) {
              correosResponsables.push(mail);
            }
          });
        }
      }
      // Defensa adicional: si por error frontend envió oficiales en payload.correosSeleccionados durante prueba (no debería), ya están aislados arriba; pero filtrar CC contra esa lista por si acaso
      var oficialesEnPayload = (payload.correosSeleccionados || []).map(function(m){ return String(m).trim().toLowerCase(); });
      if (oficialesEnPayload.length > 0) {
        correosCopiaExtra = correosCopiaExtra.filter(function(m){ return oficialesEnPayload.indexOf(m) === -1; });
        correosResponsables = correosResponsables.filter(function(m){ return oficialesEnPayload.indexOf(m) === -1; });
      }
    }

    if (!Array.isArray(destinatarios) || destinatarios.length === 0) {
      var msgSinDest = (payload.tipoEnvio === "prueba")
        ? "No hay correos con rol PRUEBA en Config_Personal para " + payload.negocio + " / " + payload.operacion + ". Agrégalos en la hoja para enviar la prueba."
        : "No se encontraron correos destinatarios válidos para procesar el envío.";
      return { 
        exito: false, 
        mensaje: msgSinDest
      };
    }

    // OBTENER COLOR DE SERVICIO DEFINIDO PARA EL CORREO
    function obtenerColorServicioCorreo(pct) {
      var valor = parseFloat(pct || 0);
      if (valor < 60) {
        return "#FF0000"; // Rojo bien definido
      } else if (valor < 80) {
        return "#B8860B"; // Amarillo/Dorado bien definido
      } else {
        return "#008000"; // Verde bien definido
      }
    }

    var participantesHtml = formatearListaHtml(payload.participantesSeleccionados);
    var fechaTexto = Utilities.formatDate(fechaActual, timeZone, "dd/MM/yyyy");

    // Color dinámico para cuadro Nivel de Cumplimiento General en correo y PDF: [0;60> #FF0000 rojo puro, [60;80> #B8860B amarillo, >=80 #00CCAA verde; sin datos (todo N/A) gris neutro
    var colorNivelGeneral = !tieneCumplimientoGeneral ? CONFIG.COLOR_INK_SOFT : (cumploGen < 60 ? "#FF0000" : (cumploGen < 80 ? "#B8860B" : "#00CCAA"));

    // OBTENCIÓN DE LOGO: Solo para PDF (logo removido del correo según requerimiento)
    var logoBase64 = obtenerBase64DesdeDrive(CONFIG.ID_LOGO_DRIVE);
    var logoPdfSrc = logoBase64 !== "" 
      ? logoBase64 
      : "https://lh3.googleusercontent.com/d/" + CONFIG.ID_LOGO_DRIVE;

    var logoPdfHtml = '<img src="' + logoPdfSrc + '" style="max-height: 28px; width: auto; display: block; margin: 0 auto;" alt="SHARF" />';

    // DISEÑO DE CORREO ELECTRÓNICO HTML (Ligero, sin base64 en el cuerpo para evitar recorte en Gmail)
    var htmlCorreo = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">
        <style>@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap'); body, table, td, th, p, h1, h3, span, div { font-family: 'Manrope', sans-serif; }</style>
      </head>
      <body style="margin: 0; padding: 0; background-color: #ffffff; font-family: 'Manrope', sans-serif;">
        <div style="max-width: 750px; margin: 20px auto; border: 1px solid ${CONFIG.COLOR_LINE}; border-radius: 8px; overflow: hidden; background-color: #ffffff; font-family: 'Manrope', sans-serif;">
          <div style="background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; padding: 22px 20px; text-align: center; font-family: 'Manrope', sans-serif;">
            <h1 style="margin: 0; font-size: 20px; font-weight: 800; color: #ffffff; text-transform: uppercase; letter-spacing: 0.5px; text-align: center; font-family: 'Manrope', sans-serif;">REPORTE DE INSPECCIÓN ${payload.tipoEnvio === "prueba" ? "(PRUEBA)" : ""}</h1>
            <p style="margin: 6px 0 0 0; font-size: 13px; font-weight: 700; color: #ffffff; text-transform: uppercase; text-align: center;">${payload.negocio} - (${payload.operacion})</p>
          </div>
          
          <div style="padding: 28px;">
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 24px; font-size: 13px; background-color: ${CONFIG.COLOR_LINE_SOFT}; border-radius: 6px;">
              <tr><td style="padding: 10px 14px; font-weight: 700; color: ${CONFIG.COLOR_INK_SOFT}; width: 35%;">Negocio:</td><td style="padding: 10px 14px; font-weight: 700; color: ${CONFIG.COLOR_INK};">${payload.negocio}</td></tr>
              <tr><td style="padding: 10px 14px; font-weight: 700; color: ${CONFIG.COLOR_INK_SOFT};">Sede / Operación:</td><td style="padding: 10px 14px; font-weight: 700; color: ${CONFIG.COLOR_INK};">${payload.operacion}</td></tr>
              <tr><td style="padding: 10px 14px; font-weight: 700; color: ${CONFIG.COLOR_INK_SOFT};">Fecha de Inspección:</td><td style="padding: 10px 14px; color: ${CONFIG.COLOR_INK};">${fechaTexto}</td></tr>
            </table>

            <div style="background-color: ${colorNivelGeneral}; color: #ffffff; padding: 20px; border-radius: 8px; text-align: center; margin-bottom: 28px; border: none;">
              <span style="font-size: 12px; text-transform: uppercase; font-weight: 800; display: block; margin-bottom: 4px; color: #ffffff;">NIVEL DE CUMPLIMIENTO GENERAL</span>
              <span style="font-size: 38px; font-weight: 800; color: #ffffff;">${textoCumplimientoGeneral}</span>
            </div>

            <h3 style="font-size: 14px; color: ${CONFIG.COLOR_INK}; border-bottom: 2px solid #000000; padding-bottom: 6px; margin-bottom: 16px; text-transform: uppercase; font-weight: 800;">RESUMEN POR MACROPROCESO</h3>
            
            <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
              <thead>
                <tr style="background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; color: #ffffff; font-weight: 700; font-family: 'Manrope', sans-serif;">
                  <th style="padding: 10px 12px; border: 1px solid ${CONFIG.COLOR_HEADER_CUSTOM}; text-align: center; width: 28%; color: #ffffff; background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; font-family: 'Manrope', sans-serif;">Proceso</th>
                  <th style="padding: 10px 12px; border: 1px solid ${CONFIG.COLOR_HEADER_CUSTOM}; text-align: center; width: 18%; color: #ffffff; background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; font-family: 'Manrope', sans-serif;">% Servicio</th>
                  <th style="padding: 10px 12px; border: 1px solid ${CONFIG.COLOR_HEADER_CUSTOM}; text-align: left; width: 54%; color: #ffffff; background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; font-family: 'Manrope', sans-serif;">Resumen</th>
                </tr>
              </thead>
              <tbody>
    `;

    var hayActividadesEvaluadas = false;
    for (var act in resumenActividades) {
      var infoAct = resumenActividades[act];
      // Si la actividad no tuvo items evaluables (todos fueron N/A), se omite del correo
      // Para 5S, pctOverride (el % real de la macro = promedio de las 5 S) manda sobre cumple/total de puntos
      if (infoAct.total === 0 && infoAct.pctOverride == null) continue;
      hayActividadesEvaluadas = true;

      var pctAct = (infoAct.pctOverride != null) ? Number(infoAct.pctOverride).toFixed(2) : ((infoAct.cumple / infoAct.total) * 100).toFixed(2);
      var colorServicio = obtenerColorServicioCorreo(pctAct);
      var resumenTexto = generarResumenProceso(infoAct);

      htmlCorreo += `
        <tr>
          <td style="padding: 10px 12px; border: 1px solid ${CONFIG.COLOR_LINE}; text-align: center; font-weight: 700; color: ${CONFIG.COLOR_INK}; vertical-align: top;">${act}</td>
          <td style="padding: 10px 12px; border: 1px solid ${CONFIG.COLOR_LINE}; text-align: center; font-weight: 800; font-size: 13px; color: ${colorServicio}; vertical-align: top;">${pctAct}%</td>
          <td style="padding: 10px 12px; border: 1px solid ${CONFIG.COLOR_LINE}; text-align: left; color: ${CONFIG.COLOR_INK}; vertical-align: top;">${resumenTexto}</td>
        </tr>
      `;
    }

    if (!hayActividadesEvaluadas) {
      htmlCorreo += `
        <tr>
          <td colspan="3" style="padding: 12px; border: 1px solid ${CONFIG.COLOR_LINE}; color: ${CONFIG.COLOR_INK_SOFT}; font-style: italic; text-align: center;">Todas las actividades fueron marcadas como No Aplica (N/A) para este turno.</td>
        </tr>
      `;
    }

    var htmlComentariosExtraCorreo = "";
    if (comentariosExtraProcesados.length > 0) {
      htmlComentariosExtraCorreo = `
        <div style="margin-top: 25px; border-top: 2px solid ${CONFIG.COLOR_BLACK}; padding-top: 15px;">
          <h3 style="font-size: 13px; color: ${CONFIG.COLOR_INK}; text-transform: uppercase; font-weight: 800; margin-bottom: 12px; letter-spacing: 0.5px;">
            COMENTARIOS ADICIONALES
          </h3>
          <ul style="padding-left: 18px; margin: 0;">
      `;

      comentariosExtraProcesados.forEach(function(ce) {
        var imgHtml = "";
        if (ce.fileIds && ce.fileIds.length > 0) {
          ce.fileIds.forEach(function(fId) {
            imgHtml += `
              <div style="margin-top: 8px; margin-bottom: 12px;">
                <img src="https://lh3.googleusercontent.com/d/${fId}" style="max-width: 580px; width: 100%; border-radius: 6px; border: 1px solid ${CONFIG.COLOR_LINE}; display: block;" alt="Evidencia Comentario Extra" />
              </div>
            `;
          });
        }

        var htmlNotificados = "";
        if (ce.correosCopia) {
          htmlNotificados = `<div style="font-size: 11px; color: ${CONFIG.COLOR_GREEN_DEEP}; font-weight: 700; margin-top: 4px;">En copia: ${ce.correosCopia}</div>`;
        }

        htmlComentariosExtraCorreo += `
          <li style="font-size: 13px; color: ${CONFIG.COLOR_INK}; line-height: 1.5; margin-bottom: 14px;">
            <div style="font-weight: 600;">${ce.comentario || "Sin comentario de texto."}</div>
            ${htmlNotificados}
            ${imgHtml}
          </li>
        `;
      });

      htmlComentariosExtraCorreo += `
          </ul>
        </div>
      `;
    }

    htmlCorreo += `
              </tbody>
            </table>

            ${htmlComentariosExtraCorreo}

            <p style="font-size: 11px; color: ${CONFIG.COLOR_INK_SOFT}; margin-top: 20px;">* Se adjunta a este mensaje el informe oficial en PDF conteniendo la matriz detallada de hallazgos y evidencias fotográficas.</p>
          </div>
          
          <div style="background-color: ${CONFIG.COLOR_LINE_SOFT}; padding: 14px; text-align: center; font-size: 11px; color: ${CONFIG.COLOR_INK_SOFT}; border-top: 1px solid ${CONFIG.COLOR_LINE};">
            Servicio Logístico SHARF — Mensaje generado automáticamente.
          </div>
        </div>
      </body>
      </html>
    `;

    // 6. GENERACIÓN DEL PDF ADJUNTO MULTI-PÁGINA CON ENCABEZADO Y CONTEO AUTOMÁTICO
    var bloquesMacroprocesos = [];
    for (var actK in resumenActividades) {
      var infoActK = resumenActividades[actK];
      // La sección de 5S muestra las 5 S (incluidas las N/A), sin filtrar por estado, y sin
      // el detalle punto por punto (infoActK.detalles ya son los 5 "detalle" sintéticos por S).
      var detallesEvaluados = (infoActK.pctOverride != null)
        ? infoActK.detalles
        : infoActK.detalles.filter(function(d) {
            return d.estado === 'CUMPLE' || d.estado === 'NO CUMPLE';
          });
      if (detallesEvaluados.length === 0) continue;

      var pctActK = (infoActK.pctOverride != null) ? Number(infoActK.pctOverride).toFixed(2) : (infoActK.total > 0 ? ((infoActK.cumple / infoActK.total) * 100).toFixed(2) : "0.00");
      var estiloActK = obtenerEstiloCumplimiento(pctActK);

      bloquesMacroprocesos.push({
        nombre: actK,
        cumplimiento: pctActK,
        colorDeep: estiloActK.deep,
        detalles: detallesEvaluados
      });
    }

    // Aplanar elementos para paginación precisa
    var listaPlanaPdf = [];
    bloquesMacroprocesos.forEach(function(b) {
      b.detalles.forEach(function(d) {
        listaPlanaPdf.push({
          actNombre: b.nombre,
          pct: b.cumplimiento,
          colorDeep: b.colorDeep,
          detalle: d
        });
      });
    });

    // Calcular páginas
    var paginasPdf = [];
    var indiceItem = 0;
    var totalItemsPdf = listaPlanaPdf.length;
    var numPagina = 1;

    while (indiceItem < totalItemsPdf) {
      var limiteItems = (numPagina === 1) ? 9 : 13;
      var grupo = listaPlanaPdf.slice(indiceItem, indiceItem + limiteItems);
      indiceItem += limiteItems;

      paginasPdf.push({
        numero: numPagina,
        items: grupo,
        tieneCE: false
      });
      numPagina++;
    }

    if (comentariosExtraProcesados.length > 0) {
      if (paginasPdf.length > 0 && paginasPdf[paginasPdf.length - 1].items.length <= (paginasPdf.length === 1 ? 4 : 7)) {
        paginasPdf[paginasPdf.length - 1].tieneCE = true;
      } else {
        paginasPdf.push({
          numero: numPagina,
          items: [],
          tieneCE: true
        });
      }
    }

    if (paginasPdf.length === 0) {
      paginasPdf.push({
        numero: 1,
        items: [],
        tieneCE: false
      });
    }

    var totalPaginasPdf = paginasPdf.length;

    var htmlPdf = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
          @page { size: A4 portrait; margin: 8mm; }
          * { font-family: 'Manrope', sans-serif; }
          html, body { -webkit-print-color-adjust: exact; print-color-adjust: exact; margin: 0; padding: 0; font-family: 'Manrope', sans-serif; color: ${CONFIG.COLOR_INK}; font-size: 8.5pt; }
          .pdf-page { width: 100%; box-sizing: border-box; }
          table.header-pdf-table { width: 100%; border-collapse: collapse; margin-bottom: 8px; border: none; }
          table.header-pdf-table td { border: none !important; vertical-align: middle; padding: 4px; background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; color: #ffffff; }
          .header-left-cell { width: 20%; text-align: center; background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; border: none !important; }
          .header-center-cell { width: 55%; text-align: center; font-weight: 800; color: #ffffff; font-size: 11pt; line-height: 1.2; background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; text-transform: uppercase; border: none !important; }
          .header-right-cell { width: 25%; padding: 0 !important; background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; border: none !important; }
          table.meta-box-table { width: 100%; border-collapse: collapse; font-size: 7.5pt; font-weight: 700; background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; color: #ffffff; border: none; }
          table.meta-box-table td { border: none !important; padding: 2px 4px !important; background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; color: #ffffff; text-align: center !important; }
          table.meta-table { width: 100%; margin-top: 4px; margin-bottom: 8px; border-collapse: collapse; font-size: 8.5pt; }
          table.meta-table td { padding: 4px 8px; border-bottom: 1px solid ${CONFIG.COLOR_LINE}; vertical-align: top; }
          .section-title { font-size: 9.5pt; font-weight: 800; color: #000000; margin-top: 8px; margin-bottom: 6px; border-bottom: 2px solid #000000; padding-bottom: 3px; text-transform: uppercase; }
          table.detail-table { width: 100%; border-collapse: collapse; margin-top: 4px; margin-bottom: 8px; table-layout: fixed; }
          table.detail-table th { background-color: ${CONFIG.COLOR_HEADER_CUSTOM}; color: #ffffff; padding: 5px 4px; text-align: center; font-size: 7.5pt; font-weight: 700; text-transform: uppercase; border: 1px solid ${CONFIG.COLOR_HEADER_CUSTOM}; }
          table.detail-table td { padding: 4px 4px; border: 1px solid ${CONFIG.COLOR_LINE}; font-size: 7.5pt; vertical-align: middle; text-align: center; word-wrap: break-word; }
          .col-sub { text-align: left !important; font-weight: 600; color: ${CONFIG.COLOR_INK}; }
          .col-sub-punto { padding-left: 14px !important; font-weight: normal !important; }
          .col-obs { text-align: center !important; }
          .badge-cumple { color: #003C46; font-weight: 800; background-color: #99EBDD; border-radius: 3px; padding: 2px 4px; }
          .badge-nocumple { color: #68002B; font-weight: 800; background-color: #FFC1C2; border-radius: 3px; padding: 2px 4px; }
          .badge-alerta { color: #68002B; font-weight: 800; text-align: center; }
          .img-link { color: #003C46; text-decoration: underline; font-weight: 700; }
        </style>
      </head>
      <body>
    `;

    paginasPdf.forEach(function(pag, idxPag) {
      var saltoEstilo = (idxPag < totalPaginasPdf - 1) ? 'page-break-after: always;' : '';
      
      htmlPdf += `
        <div class="pdf-page" style="${saltoEstilo}">
          <table class="header-pdf-table">
            <tr>
              <td class="header-left-cell">
                ${logoPdfHtml}
              </td>
              <td class="header-center-cell">
                Control Operacional - SIG
              </td>
              <td class="header-right-cell">
                <table class="meta-box-table">
                  <tr>
                    <td style="width: 40%; color: #ffffff; white-space: nowrap;">Código:</td>
                    <td style="width: 60%; font-weight: 400; color: #ffffff; white-space: nowrap;">SIG - RG - 320</td>
                  </tr>
                  <tr>
                    <td style="color: #ffffff; white-space: nowrap;">Versión:</td>
                    <td style="font-weight: 400; color: #ffffff; white-space: nowrap;">V.02</td>
                  </tr>
                  <tr>
                    <td style="color: #ffffff; white-space: nowrap;">Página:</td>
                    <td style="font-weight: 400; color: #ffffff; white-space: nowrap;">${pag.numero} de ${totalPaginasPdf}</td>
                  </tr>
                </table>  
              </td>
            </tr>
          </table>
      `;

      if (pag.numero === 1) {
        htmlPdf += `
          <table class="meta-table">
            <tr><td style="font-weight:700; width:28%; color:${CONFIG.COLOR_INK_SOFT};">Negocio:</td><td style="font-weight:700;">${payload.negocio}</td></tr>
            <tr><td style="font-weight:700; color:${CONFIG.COLOR_INK_SOFT};">Sede / Operación:</td><td style="font-weight:700;">${payload.operacion}</td></tr>
            <tr><td style="font-weight:700; color:${CONFIG.COLOR_INK_SOFT};">Fecha de Inspección:</td><td>${fechaTexto}</td></tr>
            <tr><td style="font-weight:700; color:${CONFIG.COLOR_INK_SOFT};">Participantes:</td><td>${participantesHtml}</td></tr>
            <tr>
              <td style="font-weight:700; color:${CONFIG.COLOR_INK_SOFT};">Cumplimiento General:</td>
              <td>
                <div style="background-color: ${colorNivelGeneral}; color: #ffffff; padding: 3px 10px; border-radius: 4px; display: inline-block; font-weight: 800; font-size: 9pt; border: none;">
                  ${textoCumplimientoGeneral}
                </div>
              </td>
            </tr>
          </table>
          <div class="section-title">EVALUACIÓN DETALLADA POR ACTIVIDAD</div>
        `;
      }

      if (pag.items && pag.items.length > 0) {
        var gruposActividad = {};
        pag.items.forEach(function(item) {
          var k = item.actNombre;
          if (!gruposActividad[k]) {
            gruposActividad[k] = { pct: item.pct, colorDeep: item.colorDeep, filas: [] };
          }
          gruposActividad[k].filas.push(item.detalle);
        });

        for (var nomAct in gruposActividad) {
          var g = gruposActividad[nomAct];
          htmlPdf += `
            <h4 style="margin-top: 8px; margin-bottom: 3px; color:${CONFIG.COLOR_INK}; font-size: 8.5pt; font-weight:800;">
              ${nomAct} — Cumplimiento: <span style="color:${g.colorDeep};">${g.pct}%</span>
            </h4>
            <table class="detail-table">
              <thead>
                <tr>
                  <th style="width: 30%;">Subactividad</th>
                  <th style="width: 12%;">Estado</th>
                  <th style="width: 22%;">Observación</th>
                  <th style="width: 14%;">Evidencia</th>
                  <th style="width: 11%;">Responsable</th>
                  <th style="width: 11%;">Alertas</th>
                </tr>
              </thead>
              <tbody>
          `;

          g.filas.forEach(function(d) {
            var clState = d.estado === 'CUMPLE' ? 'badge-cumple' : 'badge-nocumple';
            var respFormatted = formatearListaHtml(d.responsable);
            var obsText = (d.comentario && d.comentario.trim() !== "" && d.comentario.trim() !== "-") ? d.comentario : "-";

            var imgContent = "-";
            if (d.fotos && d.fotos.length > 0) {
              var links = [];
              d.fotos.forEach(function(url, idx) {
                links.push('<a class="img-link" href="' + url + '" target="_blank">Ver foto ' + (idx + 1) + '</a>');
              });
              imgContent = links.join("<br>");
            }

            var subTexto = d.es5SPunto ? ('↳ ' + d.subactividad) : d.subactividad;
            var estiloSub = d.es5SPunto ? ' style="padding-left: 14px; font-weight: normal;"' : '';
            var clSub = d.es5SPunto ? 'col-sub col-sub-punto' : 'col-sub';

            htmlPdf += `
              <tr>
                <td class="${clSub}"${estiloSub}>${subTexto}</td>
                <td class="${clState}">${d.estado}</td>
                <td class="col-obs">${obsText}</td>
                <td>${imgContent}</td>
                <td>${respFormatted}</td>
                <td class="badge-alerta">${d.alerta}</td>
              </tr>
            `;
          });

          htmlPdf += `</tbody></table>`;
        }
      } else if (pag.numero === 1 && !pag.tieneCE) {
        htmlPdf += `<p style="text-align: center; padding: 20px; color: ${CONFIG.COLOR_INK_SOFT}; font-style: italic;">No se registraron subactividades evaluables en este turno (todas fueron marcadas como No Aplica).</p>`;
      }

      if (pag.tieneCE && comentariosExtraProcesados.length > 0) {
        htmlPdf += `
          <div style="margin-top: 10px;">
            <h4 style="margin-top: 8px; margin-bottom: 4px; color:${CONFIG.COLOR_INK}; font-size: 9pt; font-weight:800; border-bottom: 1.5px solid ${CONFIG.COLOR_INK}; padding-bottom: 2px;">
              COMENTARIOS ADICIONALES
            </h4>
            <table class="detail-table" style="margin-top: 4px;">
              <thead>
                <tr>
                  <th style="width: 8%;">N°</th>
                  <th style="width: 52%;">Comentario / Hallazgo</th>
                  <th style="width: 40%;">Evidencia Fotográfica</th>
                </tr>
              </thead>
              <tbody>
        `;

        comentariosExtraProcesados.forEach(function(ce) {
          var linksCE = [];
          if (ce.urls && ce.urls.length > 0) {
            ce.urls.forEach(function(url, idx) {
              linksCE.push('<a class="img-link" href="' + url + '" target="_blank">Ver foto ' + (idx + 1) + '</a>');
            });
          }
          var imgContentCE = linksCE.length > 0 ? linksCE.join("<br>") : "-";

          var txtNotifPdf = ce.correosCopia
            ? "<br><span style='font-size: 6.5pt; color: #003C46; font-weight: 700;'>En copia: " + ce.correosCopia + "</span>"
            : "";

          htmlPdf += `
            <tr>
              <td style="font-weight: 700; text-align: center;">${ce.numero}</td>
              <td class="col-obs" style="text-align: left;">${ce.comentario || "-"}${txtNotifPdf}</td>
              <td>${imgContentCE}</td>
            </tr>
          `;
        });

        htmlPdf += `</tbody></table></div>`;
      }

      htmlPdf += `</div>`;
    });

    htmlPdf += `
      </body>
      </html>
    `;

    var blobPdf = Utilities.newBlob(htmlPdf, 'text/html', 'Reporte_Inspeccion_' + payload.operacion + '.html')
                             .getAs('application/pdf');

    // Nomenclatura oficial del reporte: [Fecha]_[Operación]_[Negocio].pdf  ej. 21ago26_Outbound_Fedex.pdf
    var nombreReporteOficial = "";
    try {
      if (typeof generarNombreReporte === 'function') {
        nombreReporteOficial = generarNombreReporte(fechaCortaArchivo, payload.operacion, payload.negocio);
      } else {
        var opSan = String(payload.operacion || "Operacion").trim().replace(/\s+/g, "_").replace(/[^a-zA-Z0-9_\-]/g, "_");
        var negSan = String(payload.negocio || "Negocio").trim().replace(/\s+/g, "_").replace(/[^a-zA-Z0-9_\-]/g, "_");
        nombreReporteOficial = fechaCortaArchivo + "_" + opSan + "_" + negSan + ".pdf";
      }
    } catch (eNom) {
      nombreReporteOficial = fechaCortaArchivo + "_" + String(payload.operacion).replace(/[^a-zA-Z0-9]/g, "_") + "_" + String(payload.negocio).replace(/[^a-zA-Z0-9]/g, "_") + ".pdf";
    }
    blobPdf.setName(nombreReporteOficial);

    // Guardar PDF en jerarquía Drive: Año / Negocio / Operación  (raíz: 1-q7curdFAaeVDj3bda1e9esXETEoMRE-)
    // Lógica: busca año, si no existe lo crea (ej. 2026); luego negocio; luego operación; si existen solo sube el reporte.
    var archivoReporteGuardado = null;
    var urlReporteGuardado = "";
    try {
      if (typeof guardarReportePDFEnJerarquia === 'function') {
        archivoReporteGuardado = guardarReportePDFEnJerarquia(blobPdf, payload.negocio, payload.operacion, fechaActual);
        if (archivoReporteGuardado) {
          try { urlReporteGuardado = archivoReporteGuardado.getUrl(); } catch (eUrl) {}
          Logger.log("Reporte PDF guardado en Drive jerárquico: " + nombreReporteOficial + " -> " + urlReporteGuardado);
        }
      } else if (typeof obtenerCarpetaJerarquicaReporte === 'function') {
        var carpetaRep = obtenerCarpetaJerarquicaReporte(payload.negocio, payload.operacion, fechaActual);
        if (carpetaRep) {
          archivoReporteGuardado = carpetaRep.createFile(blobPdf);
          try { urlReporteGuardado = archivoReporteGuardado.getUrl(); } catch (eUrl2) {}
        }
      }
    } catch (eSave) {
      Logger.log("Aviso: no se pudo guardar el reporte en la carpeta jerárquica: " + eSave.toString());
    }

    var sufijoAsuntoCumplimiento = tieneCumplimientoGeneral ? " (" + cumploGen.toFixed(2) + "%)" : "";
    var asunto = "Reporte de Inspección SHARF - " + payload.negocio + " / " + payload.operacion + sufijoAsuntoCumplimiento;

    var emailConfig = {
      to: destinatarios.join(","),
      subject: asunto,
      htmlBody: htmlCorreo,
      attachments: [blobPdf]
    };

    // Logo removido del correo según requerimiento — no se adjuntan inlineImages
    var todosCC = correosCopiaExtra.concat(correosResponsables);
    if (todosCC.length > 0) {
      emailConfig.cc = todosCC.join(",");
    }

    MailApp.sendEmail(emailConfig);

    // Memorandos de amonestación (3era Alerta): se registran como "Pendiente de revisión" y se avisa al
    // inspector DESPUÉS de todas las escrituras de historial y del envío del reporte principal. El memo
    // solo se genera/envía cuando el inspector lo aprueba (Revision.html). Un fallo aquí jamás debe
    // romper procesarInspeccion.
    var numMemosPreparados = 0;
    try {
      if (amonestacionesDetectadas.length > 0) {
        numMemosPreparados = _registrarAmonestacionesPendientesMemo(amonestacionesDetectadas, {
          negocio: payload.negocio,
          operacion: payload.operacion,
          fechaActual: fechaActual,
          inspector: usuarioActivo
        }) || 0;
      }
    } catch (eMemo) {
      Logger.log("Aviso: error al registrar amonestaciones pendientes de revisión: " + eMemo.toString());
    }

    var mensajeFinal = payload.tipoEnvio === "prueba"
      ? "¡Reporte de PRUEBA enviado exitosamente a: " + destinatarios.join(", ") + "!"
      : "¡Reporte OFICIAL enviado exitosamente con el informe PDF adjunto!";

    if (numMemosPreparados > 0) {
      mensajeFinal += " Además se prepararon " + numMemosPreparados +
        (numMemosPreparados === 1
          ? " memorando: revisa tu correo para confirmar su envío."
          : " memorandos: revisa tu correo para confirmar su envío.");
    }

    return {
      exito: true,
      mensaje: mensajeFinal,
      nombreReporte: nombreReporteOficial
    };

  } catch (e) {
    Logger.log("Error en procesarInspeccion: " + e.toString());
    return { exito: false, mensaje: "Error al procesar: " + e.toString() };
  }
}

/**
 * Genera el resumen conciso para la columna 'Resumen' de la tabla del correo.
 * Reglas de negocio:
 * 1. Si no hay fallas ni observaciones: 'Conforme. Sin observaciones.'
 * 2. Si hay puntos con observaciones (NO CUMPLE): 'Observaciones en los puntos X.XX, Y.YY. Revisar detalle en reporte.'
 * 3. Si hay fallas sin observación específica: 'N puntos no cumplen sin observación específica.'
 */
function generarResumenProceso(infoAct) {
  if (!infoAct || !infoAct.detalles || infoAct.detalles.length === 0) {
    return '<span style="color: #3A3A3A; font-style: italic;">Sin registro de actividades.</span>';
  }

  var puntosNoCumpleConObs = [];
  var puntosNoCumpleSinObs = 0;
  var puntosCumpleConObs = [];

  infoAct.detalles.forEach(function(d) {
    if (d.es5SPunto === true) return;

    var c = (d.comentario || "").trim();
    var tieneObs = (c !== "" && c !== "-" && c.toUpperCase() !== "NAN");
    var idDisplay = d.idSub ? String(d.idSub).trim() : "";
    
    // Si el idDisplay tiene formato de timestamp largo, extraer parte principal
    if (idDisplay.length > 10 && idDisplay.indexOf(" ") > -1) {
      idDisplay = idDisplay.split(" ")[0];
    }

    if (d.estado === "NO CUMPLE") {
      if (tieneObs) {
        puntosNoCumpleConObs.push(idDisplay);
      } else {
        puntosNoCumpleSinObs++;
      }
    } else if (d.estado === "CUMPLE" && tieneObs) {
      puntosCumpleConObs.push(idDisplay);
    }
  });

  // Caso 1: 100% CUMPLE / Sin fallas
  if (infoAct.noCumple === 0) {
    if (puntosCumpleConObs.length === 0) {
      return '<span style="color: #003C46; font-weight: 600;">Conforme. Sin observaciones.</span>';
    } else {
      var txtCumple = puntosCumpleConObs.length === 1 
        ? "el punto " + puntosCumpleConObs[0] 
        : "los puntos " + puntosCumpleConObs.join(", ");
      return '<span style="color: #003C46; font-weight: 600;">Conforme. Observación registrada en ' + txtCumple + '. Revisar detalle en reporte.</span>';
    }
  }

  // Caso 2: Puntos en NO CUMPLE con observación registrada
  if (puntosNoCumpleConObs.length > 0 && puntosNoCumpleSinObs === 0) {
    var txtNoCumple = puntosNoCumpleConObs.length === 1 
      ? "el punto " + puntosNoCumpleConObs[0] 
      : "los puntos " + puntosNoCumpleConObs.join(", ");
    return '<span style="color: #68002B; font-weight: 600;">Observaciones en ' + txtNoCumple + '. Revisar detalle en reporte.</span>';
  }

  // Caso 3: Puntos en NO CUMPLE sin observación específica
  if (puntosNoCumpleConObs.length === 0 && puntosNoCumpleSinObs > 0) {
    var txtSinObs = puntosNoCumpleSinObs === 1 
      ? "1 punto no cumple sin observación específica." 
      : puntosNoCumpleSinObs + " puntos no cumplen sin observación específica.";
    return '<span style="color: #68002B; font-weight: 600;">' + txtSinObs + '</span>';
  }

  // Caso 4: Mixto (algunos con observación y otros sin ella)
  var txtMixto = puntosNoCumpleConObs.length === 1 
    ? "el punto " + puntosNoCumpleConObs[0] 
    : "los puntos " + puntosNoCumpleConObs.join(", ");
  var txtExtra = puntosNoCumpleSinObs === 1
    ? "y 1 punto adicional sin observación"
    : "y " + puntosNoCumpleSinObs + " puntos adicionales sin observación";
  return '<span style="color: #68002B; font-weight: 600;">Observaciones en ' + txtMixto + ' (' + txtExtra + '). Revisar detalle en reporte.</span>';
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    transformarIdEvidencia5S: transformarIdEvidencia5S,
    obtenerDestinoEvaluacion5S: obtenerDestinoEvaluacion5S,
    construirFilaHistorial5S: construirFilaHistorial5S,
    construirFilasResumen5S: construirFilasResumen5S,
    construirDetallesPdf5S: construirDetallesPdf5S,
    generarResumenProceso: generarResumenProceso
  };
}
