/**
 * SISTEMA DE INSPECCIÓN DE PROCESOS - SHARF
 * Archivo: Memo.gs — Memorandos de amonestación (3era Alerta)
 *
 * Cuando un responsable llega a la 3era alerta del ciclo (ver History.js) en un envío OFICIAL
 * (y sin "No Alertar"), se genera un memorando de amonestación desde una plantilla Google Doc
 * y se envía por correo, con botones Conforme / No conforme que abren la web app de firma del
 * proyecto firma-memo (otro proyecto Apps Script, corre como propietario). Mientras
 * CONFIG.MEMO_MODO_PRUEBA sea true, el correo va SOLO a los destinatarios configurados en
 * Config_Personal con Rol CONFIG.MEMO_ROL_DESTINATARIO ("Prueba_Memo"), nunca al colaborador
 * ni a su jefe. Cuando es false (modo OFICIAL), el correo va "para" (to) al Rol
 * CONFIG.MEMO_ROL_OFICIAL ("Memos"), con copia (cc) al colaborador amonestado, su jefe
 * inmediato y el inspector; y copia oculta (bcc) al Rol "Prueba_Memo" si
 * CONFIG.MEMO_BCC_MONITOREO es true (ver construirDestinatariosMemo). La hoja "Memorandos" y el
 * contrato de columnas/URL de firma se comparten EXACTAMENTE con firma-memo (ver
 * construirCabeceraMemorandos y construirUrlRespuestaMemo).
 *
 * REVISIÓN PREVIA: procesarInspeccion NO genera el memo directamente. Registra la amonestación como
 * fila "Pendiente de revisión" (sin Token) y avisa al inspector que envió el checklist
 * (registrarAmonestacionesPendientes). El inspector decide en Revision.html (?vista=revision):
 * "generar" -> se genera/envía el memo ACTUALIZANDO esa misma fila; "descartar" -> queda
 * "Descartado por revisión" con motivo (decidirPendientes). generarYEnviarMemorandos se conserva
 * como generación directa de respaldo.
 *
 * Funciones puras (testeables en Node, sin dependencias de Apps Script) y funciones de
 * integración con Apps Script (DriveApp/DocumentApp/MailApp/SpreadsheetApp).
 */

// Resolución de utilidades de Config.js / History.js compatible con Node y Apps Script,
// siguiendo el mismo patrón usado en History.js / Report.js.
var _cfgModMemo = (typeof module !== 'undefined' && module.exports)
  ? require('./Config.js')
  : { CONFIG: typeof CONFIG !== 'undefined' ? CONFIG : null };

var _historyModMemo = (typeof module !== 'undefined' && module.exports)
  ? require('./History.js')
  : {
      filaCuentaComoFallaPrevia: typeof filaCuentaComoFallaPrevia !== 'undefined' ? filaCuentaComoFallaPrevia : null,
      normalizarIdHistorico: typeof normalizarIdHistorico !== 'undefined' ? normalizarIdHistorico : null
    };

function _configMemo() {
  return (_cfgModMemo && _cfgModMemo.CONFIG) ? _cfgModMemo.CONFIG : (typeof CONFIG !== 'undefined' ? CONFIG : {});
}

function _normTextoClaveMemo(t) {
  if (_cfgModMemo && typeof _cfgModMemo.normalizarTextoClave === 'function') return _cfgModMemo.normalizarTextoClave(t);
  if (typeof normalizarTextoClave === 'function') return normalizarTextoClave(t);
  return String(t || "").toUpperCase().trim();
}

function _filaCuentaComoFallaPreviaMemo(row, c, normNeg, normOp, normSub, normResp) {
  if (_historyModMemo && typeof _historyModMemo.filaCuentaComoFallaPrevia === 'function') {
    return _historyModMemo.filaCuentaComoFallaPrevia(row, c, normNeg, normOp, normSub, normResp);
  }
  if (typeof filaCuentaComoFallaPrevia === 'function') return filaCuentaComoFallaPrevia(row, c, normNeg, normOp, normSub, normResp);
  return false;
}

function _normalizarIdHistoricoMemo(val) {
  if (_historyModMemo && typeof _historyModMemo.normalizarIdHistorico === 'function') return _historyModMemo.normalizarIdHistorico(val);
  if (typeof normalizarIdHistorico === 'function') return normalizarIdHistorico(val);
  return String(val == null ? "" : val).trim();
}

/**
 * Normaliza texto (minúsculas, sin tildes, trim) para comparaciones/tokenización internas de Memo.js.
 */
function _normTextoMemo(t) {
  return String(t || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .trim();
}

/**
 * Tokeniza un texto en palabras normalizadas (sin tildes/mayúsculas, sin puntuación).
 */
function _tokensMemo(t) {
  return _normTextoMemo(t).split(/[^a-z0-9]+/).filter(function(x) { return x !== ""; });
}

/**
 * Convierte "ABARCA ZUBIETA, DIEGO MIGUEL" -> "Diego Miguel Abarca Zubieta" (Title Case, sin coma).
 * Si no hay coma, solo aplica Title Case al texto tal cual.
 */
function formatearNombreNatural(nombreCompleto) {
  if (!nombreCompleto) return "";
  var s = String(nombreCompleto).trim();
  var idxComa = s.indexOf(",");
  var resultado;
  if (idxComa !== -1) {
    var apellidos = s.slice(0, idxComa).trim();
    var nombres = s.slice(idxComa + 1).trim();
    resultado = (nombres + " " + apellidos).trim();
  } else {
    resultado = s;
  }
  return resultado
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .map(function(w) {
      if (!w) return w;
      return w.charAt(0).toUpperCase() + w.slice(1);
    })
    .join(" ");
}

/**
 * Normaliza un DNI a texto de 8 dígitos (rellena con ceros a la izquierda) si el valor es numérico
 * (número JS o cadena de solo dígitos, con o sin ".0" final por conversión de Sheets).
 * Si no es numérico, devuelve el texto tal cual (trim).
 */
function normalizarDni(valor) {
  if (valor === null || valor === undefined || valor === "") return "";
  var s = String(valor).trim();
  if (/^\d+(\.0+)?$/.test(s)) {
    s = s.split(".")[0];
    while (s.length < 8) s = "0" + s;
    return s;
  }
  return s;
}

/**
 * Convierte la matriz cruda de la hoja `cecos` (con cabecera, getDisplayValues) en objetos
 * {nombre, dni, cargo, area, sede, dependencia, email}. Lectura SIEMPRE por encabezado normalizado
 * (sin tildes, minúsculas, exacto) para no confundir "Área" con "Sub Área" ni "Nombre" con
 * "CeCo Denominación". `sede` = Centro Operaciones.
 */
function leerCecosComoObjetos(valores) {
  if (!Array.isArray(valores) || valores.length < 2) return [];
  var header = valores[0].map(function(h) { return _normTextoMemo(h); });
  function idx(nombreExacto) { return header.indexOf(nombreExacto); }

  var idxNombre = idx("nombre");
  var idxDni = idx("dni");
  var idxCargo = idx("cargo");
  var idxArea = idx("area"); // exacto: evita "sub area"
  var idxSede = idx("centro operaciones");
  var idxDependencia = idx("dependencia");
  var idxEmail = idx("email");
  var idxGerencia = idx("gerencia"); // columna G de cecos

  var out = [];
  for (var i = 1; i < valores.length; i++) {
    var row = valores[i];
    if (!row) continue;
    out.push({
      nombre: idxNombre > -1 ? String(row[idxNombre] || "").trim() : "",
      dni: idxDni > -1 ? normalizarDni(row[idxDni]) : "",
      cargo: idxCargo > -1 ? String(row[idxCargo] || "").trim() : "",
      area: idxArea > -1 ? String(row[idxArea] || "").trim() : "",
      sede: idxSede > -1 ? String(row[idxSede] || "").trim() : "",
      dependencia: idxDependencia > -1 ? String(row[idxDependencia] || "").trim() : "",
      email: idxEmail > -1 ? String(row[idxEmail] || "").trim() : "",
      gerencia: idxGerencia > -1 ? String(row[idxGerencia] || "").trim() : ""
    });
  }
  return out;
}

/**
 * Busca un colaborador en cecos: primero por email exacto (case-insensitive, requiere "@" válido);
 * si no hay coincidencia por email, busca por nombre: todos los tokens normalizados del nombre corto
 * deben estar entre los tokens del Nombre de cecos. Devuelve el registro solo si hay EXACTAMENTE
 * una coincidencia; si no, null.
 */
function buscarColaborador(cecos, email, nombreCorto) {
  cecos = Array.isArray(cecos) ? cecos : [];

  if (email) {
    var emailNorm = String(email).trim().toLowerCase();
    if (emailNorm.indexOf("@") > -1) {
      var tokensVerif = nombreCorto ? _tokensMemo(nombreCorto) : [];
      for (var i = 0; i < cecos.length; i++) {
        if (String(cecos[i].email || "").trim().toLowerCase() === emailNorm) {
          // Verificación cruzada: el email debe pertenecer a la persona indicada por el nombre corto.
          // Si no coincide (email desalineado), se ignora y se intenta por nombre.
          var tokensNombreEmail = _tokensMemo(cecos[i].nombre);
          var coincide = tokensVerif.every(function(tk) { return tokensNombreEmail.indexOf(tk) !== -1; });
          if (coincide) return cecos[i];
          break;
        }
      }
    }
  }

  if (nombreCorto) {
    var tokensCorto = _tokensMemo(nombreCorto);
    if (tokensCorto.length === 0) return null;
    var matches = cecos.filter(function(c) {
      var tokensNombre = _tokensMemo(c.nombre);
      return tokensCorto.every(function(tk) { return tokensNombre.indexOf(tk) !== -1; });
    });
    if (matches.length === 1) return matches[0];
  }

  return null;
}

/**
 * Busca al jefe inmediato en cecos a partir del texto de Dependencia ("NOMBRE APELLIDO"),
 * con el mismo criterio de tokens que buscarColaborador. Única coincidencia -> {nombre, email}; si no, null.
 */
function buscarJefe(cecos, dependencia) {
  if (!dependencia) return null;
  cecos = Array.isArray(cecos) ? cecos : [];
  var tokensDep = _tokensMemo(dependencia);
  if (tokensDep.length === 0) return null;
  var matches = cecos.filter(function(c) {
    var tokensNombre = _tokensMemo(c.nombre);
    return tokensDep.every(function(tk) { return tokensNombre.indexOf(tk) !== -1; });
  });
  if (matches.length === 1) {
    return { nombre: matches[0].nombre, email: matches[0].email };
  }
  return null;
}

/**
 * Resuelve una lista de nombres de emisoras/firmantes de memorandos (CONFIG.MEMO_EMISORAS) contra
 * `cecos`, con el MISMO criterio de tokens y coincidencia única que buscarJefe. Los nombres que no
 * resuelven a exactamente un registro se omiten. Devuelve [{nombre, email}, ...] (solo resueltos).
 */
function resolverEmisoras(cecos, nombresEmisoras) {
  var lista = Array.isArray(nombresEmisoras) ? nombresEmisoras : [];
  var out = [];
  lista.forEach(function(nombre) {
    var resuelto = buscarJefe(cecos, nombre);
    if (resuelto) out.push(resuelto);
  });
  return out;
}

/**
 * Construye la URL del botón de respuesta (Revisar detalle) hacia la web app de firma-memo.
 * Devuelve "" si `urlBase` está vacío (sin botones, ver _generarYEnviarUnMemo).
 */
function construirUrlRespuestaMemo(urlBase, idMemo, token, accion) {
  if (!urlBase) return "";
  return String(urlBase) + "?memo=" + encodeURIComponent(idMemo) + "&t=" + token + "&accion=" + accion;
}

/**
 * Lee Config_Personal (matriz con cabecera Correo/Nombre | Rol | Negocio | Operación) y devuelve
 * la lista de correos (sin duplicados, en minúsculas) de las filas cuyo Rol coincide (normalizado,
 * sin importar mayúsculas/tildes) con `rol`, SIN filtrar por Negocio/Operación (fase de pruebas:
 * comparación normalizada, sin importar Negocio/Operación).
 * Si "Correo/Nombre" contiene "@" se usa tal cual; si no, se busca ese nombre en `cecos` (por nombre,
 * única coincidencia) y se usa su Email si es válido.
 */
function obtenerDestinatariosMemo(valoresConfigPersonal, cecos, rol) {
  var out = [];
  if (!Array.isArray(valoresConfigPersonal) || valoresConfigPersonal.length < 2) return out;
  var rolNorm = _normTextoMemo(rol);
  if (rolNorm === "") return out;

  for (var i = 1; i < valoresConfigPersonal.length; i++) {
    var row = valoresConfigPersonal[i];
    if (!row) continue;
    var correoNombre = row[0] ? String(row[0]).trim() : "";
    var rolFila = row[1] ? String(row[1]).trim() : "";
    if (!correoNombre) continue;
    if (_normTextoMemo(rolFila) !== rolNorm) continue;

    var email = "";
    if (correoNombre.indexOf("@") > -1) {
      email = correoNombre.toLowerCase();
    } else {
      var colaborador = buscarColaborador(cecos, null, correoNombre);
      if (colaborador && colaborador.email && String(colaborador.email).indexOf("@") > -1) {
        email = String(colaborador.email).toLowerCase();
      }
    }
    if (email && out.indexOf(email) === -1) out.push(email);
  }
  return out;
}

/**
 * Construye {to, cc, bcc} (arrays de correos en minúsculas, sin duplicados ni vacíos) para el
 * correo del memorando, según el modo:
 *  - Modo PRUEBA (modoPrueba=true): to = correosPrueba (Rol Prueba_Memo); cc y bcc vacíos.
 *  - Modo OFICIAL (modoPrueba=false): to = correosOficiales (Rol Memos); cc = colaborador
 *    amonestado + jefe inmediato + inspector (sin los que ya estén en to); bcc = correosPrueba
 *    (Rol Prueba_Memo) SOLO si bccMonitoreo es true, sin los que ya estén en to/cc.
 * Función PURA (sin dependencias de Apps Script): decide destinatarios, no envía nada.
 */
function construirDestinatariosMemo(modoPrueba, correosOficiales, correosPrueba, emailColaborador, correoJefe, inspector, bccMonitoreo) {
  function limpiarLista(lista) {
    var out = [];
    (Array.isArray(lista) ? lista : []).forEach(function(email) {
      var e = String(email === null || email === undefined ? "" : email).trim().toLowerCase();
      if (e && out.indexOf(e) === -1) out.push(e);
    });
    return out;
  }

  if (modoPrueba) {
    return { to: limpiarLista(correosPrueba), cc: [], bcc: [] };
  }

  var to = limpiarLista(correosOficiales);

  var cc = limpiarLista([emailColaborador, correoJefe, inspector]).filter(function(e) {
    return to.indexOf(e) === -1;
  });

  var bcc = [];
  if (bccMonitoreo) {
    bcc = limpiarLista(correosPrueba).filter(function(e) {
      return to.indexOf(e) === -1 && cc.indexOf(e) === -1;
    });
  }

  return { to: to, cc: cc, bcc: bcc };
}

/**
 * Formatea {to, cc, bcc} (ver construirDestinatariosMemo) como texto legible para la columna
 * "Destinatarios" de la hoja Memorandos: "para: a@x, b@x | cc: c@x | cco: d@x" (cc/cco se omiten
 * si están vacíos).
 */
function formatearDestinatariosMemo(destResult) {
  destResult = destResult || {};
  var to = Array.isArray(destResult.to) ? destResult.to : [];
  var cc = Array.isArray(destResult.cc) ? destResult.cc : [];
  var bcc = Array.isArray(destResult.bcc) ? destResult.bcc : [];
  var partes = ["para: " + to.join(", ")];
  if (cc.length) partes.push("cc: " + cc.join(", "));
  if (bcc.length) partes.push("cco: " + bcc.join(", "));
  return partes.join(" | ");
}

/**
 * A partir del valorSheet devuelto por calcularAlertaResponsable (número 1/2/3, o texto "3\n1"
 * alineado con el orden de `responsableStr`), devuelve CADA responsable con alerta numérica 1, 2
 * o 3, como [{nombre, email, nivel}] (nombre/email tomados de responsableStr/responsableEmailStr
 * en el mismo orden). Antes solo existía para la 3era alerta (ver extraerAmonestaciones, que ahora
 * se reescribe sobre esta función); el memo de 1era/2da alerta (bandera blanca del inspector) la
 * necesita completa.
 */
function extraerResponsablesConAlerta(valorSheet, responsableStr, responsableEmailStr) {
  var listaNombres = String(responsableStr || "")
    .split(/[,;]+/)
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s !== "" && s !== "-" && s.toUpperCase() !== "NAN"; });

  // Index.html descarta los emails vacíos al armar responsableEmail, así que la lista solo está
  // alineada con los nombres si ambas tienen la misma cantidad; si no, no se usa ningún email.
  var listaEmails = String(responsableEmailStr || "")
    .split(/[,;]+/)
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s !== ""; });
  if (listaEmails.length !== listaNombres.length) listaEmails = [];

  var numeros;
  if (typeof valorSheet === "number") {
    numeros = [valorSheet];
  } else {
    var s = String(valorSheet == null ? "" : valorSheet).trim();
    if (s === "" || s === "-") {
      numeros = [];
    } else {
      numeros = s.split(/\n/).map(function(x) { return parseInt(x, 10); });
    }
  }

  var out = [];
  for (var i = 0; i < listaNombres.length; i++) {
    if (numeros[i] === 1 || numeros[i] === 2 || numeros[i] === 3) {
      out.push({ nombre: listaNombres[i], email: listaEmails[i] || "", nivel: numeros[i] });
    }
  }
  return out;
}

/**
 * Devuelve solo los responsables en 3era alerta, como [{nombre, email}] (MISMA forma que siempre:
 * sin `nivel`, para no romper a quien ya hace deepStrictEqual contra este resultado). Envoltorio
 * sobre extraerResponsablesConAlerta, que ahora concentra el parseo/alineación de nombres y emails.
 */
function extraerAmonestaciones(valorSheet, responsableStr, responsableEmailStr) {
  return extraerResponsablesConAlerta(valorSheet, responsableStr, responsableEmailStr)
    .filter(function(r) { return r.nivel === 3; })
    .map(function(r) { return { nombre: r.nombre, email: r.email }; });
}

/**
 * Indica si el inspector desactivó (bandera de memo en el checklist) el memorando de `nombre`
 * para esta subactividad. `memoOmitir` es la lista de nombres que envía Index.html en
 * evaluaciones[idSub].memoOmitir; la comparación es por nombre normalizado.
 */
function memoOmitidoPorInspector(memoOmitir, nombre) {
  if (!Array.isArray(memoOmitir) || memoOmitir.length === 0) return false;
  var clave = _normTextoClaveMemo(nombre);
  if (clave === "") return false;
  return memoOmitir.some(function(n) { return _normTextoClaveMemo(n) === clave; });
}

/**
 * Indica si el inspector ACTIVÓ (bandera blanca de memo en el checklist) el memorando de `nombre`
 * para esta subactividad, en una alerta de nivel 1 o 2 (en 3era alerta el memo ya se prepara por
 * defecto, ver memoSeleccionadoParaEnvio). `memoActivar` es la lista de nombres que envía
 * Index.html en evaluaciones[idSub].memoActivar; mismo criterio de comparación que
 * memoOmitidoPorInspector (nombre normalizado).
 */
function memoActivadoPorInspector(memoActivar, nombre) {
  if (!Array.isArray(memoActivar) || memoActivar.length === 0) return false;
  var clave = _normTextoClaveMemo(nombre);
  if (clave === "") return false;
  return memoActivar.some(function(n) { return _normTextoClaveMemo(n) === clave; });
}

/**
 * Decide si se prepara el memo de `nombre` para una alerta de `nivel` (1, 2 o 3), combinando las
 * dos banderas del checklist:
 *  - apagar (memoOmitir, bandera roja, solo aplica a 3era alerta) SIEMPRE gana: no enviar es lo
 *    seguro.
 *  - 3era alerta: true salvo que esté omitida.
 *  - 1era/2da alerta: el memo NO se prepara salvo que el inspector la haya activado (memoActivar,
 *    bandera blanca).
 * Un payload antiguo sin memoActivar (o con nivel que no es 1/2/3) se comporta exactamente como
 * hoy: solo la 3era alerta genera memo.
 */
function memoSeleccionadoParaEnvio(nivel, nombre, memoOmitir, memoActivar) {
  if (memoOmitidoPorInspector(memoOmitir, nombre)) return false;
  if (nivel === 3) return true;
  if (nivel === 1 || nivel === 2) return memoActivadoPorInspector(memoActivar, nombre);
  return false;
}

/**
 * Normaliza un nivel de alerta a 1, 2 o 3. Ausente o inválido -> 3 (las filas "Datos amonestación"
 * ya existentes, todas previas a este cambio, son todas de 3era alerta).
 */
function normalizarNivelAlertaMemo(valor) {
  var n = parseInt(valor, 10);
  if (n === 1 || n === 2 || n === 3) return n;
  return 3;
}

/**
 * Texto de nivel de alerta para mostrar en el correo del memo (misma nomenclatura de
 * "1era/2da/3era Alerta" que calcularAlertaResponsable en History.js, sin el sufijo
 * "- Amonestación" que ese texto usa para el PDF/correo del reporte).
 */
function textoNivelAlertaMemo(nivel) {
  var n = normalizarNivelAlertaMemo(nivel);
  if (n === 1) return "1era Alerta";
  if (n === 2) return "2da Alerta";
  return "3era Alerta";
}

/**
 * Construye el texto de "obligación comprometida" de una fila: "<Actividad> – <Subactividad>".
 * Si falta alguno de los dos, devuelve solo el que exista.
 */
function _construirObligacion(actividad, subactividad) {
  var a = String(actividad == null ? "" : actividad).trim();
  var s = String(subactividad == null ? "" : subactividad).trim();
  if (a && s) return a + " – " + s;
  return a || s || "";
}

/**
 * Extrae las URLs http(s) de un texto (columna "URLs" del historial: una por línea, o "-"/vacío).
 */
function _extraerUrlsMemo(texto) {
  var m = String(texto == null ? "" : texto).match(/https?:\/\/[^\s,;]+/g);
  return m ? m : [];
}

/**
 * Busca en `registrosPrevios` las filas que cuentan como falla previa del ciclo actual para
 * (negocio, operacion, idSubKey, nombreResp) usando EXACTAMENTE el mismo criterio que
 * calcularAlertaResponsable (vía el predicado compartido filaCuentaComoFallaPrevia de History.js),
 * y devuelve SOLO las fallas del CICLO ACTUAL (en orden cronológico tal como aparecen en el
 * historial): las últimas (coincidencias.length % 3), es decir ninguna si el ciclo recién empieza
 * (0, 3, 6... previas totales), la última si va por su 2da alerta (1, 4, 7... previas) y las 2
 * últimas si va por su 3era alerta (2, 5, 8... previas). Con 4 previas totales (1 ciclo completo +
 * 1 del ciclo actual) devuelve solo esa última; con 2 o 5 previas (3era alerta) es idéntico a
 * devolver siempre "las 2 últimas", que es lo único que hacía esta función antes de este cambio
 * (incorrecto para 1era/2da alerta: con 4 previas devolvía una falla del ciclo ANTERIOR). Cada
 * elemento: {fecha, actividad, subactividad, obligacion}. `cols` es el layout de columnas
 * (COLS_HISTORIAL_INSPECCIONES o COLS_HISTORIAL_5S, extendido con fecha/actividad/subactividad)
 * usado para leerlas.
 */
function obtenerFallasCicloActual(registrosPrevios, negocio, operacion, idSubKey, nombreResp, cols) {
  var out = [];
  if (!Array.isArray(registrosPrevios) || !cols) return out;

  var normNeg = _normTextoClaveMemo(negocio);
  var normOp = _normTextoClaveMemo(operacion);
  var normSub = _normalizarIdHistoricoMemo(idSubKey);
  var normResp = _normTextoClaveMemo(nombreResp);

  var coincidencias = [];
  for (var r = 1; r < registrosPrevios.length; r++) {
    var row = registrosPrevios[r];
    if (_filaCuentaComoFallaPreviaMemo(row, cols, normNeg, normOp, normSub, normResp)) {
      var actividad = (cols.actividad !== null && cols.actividad !== undefined) ? row[cols.actividad] : "";
      var subactividad = (cols.subactividad !== null && cols.subactividad !== undefined) ? row[cols.subactividad] : "";
      var fecha = (cols.fecha !== null && cols.fecha !== undefined) ? row[cols.fecha] : "";
      var detalle = (cols.detalle !== null && cols.detalle !== undefined && row[cols.detalle] != null) ? String(row[cols.detalle]).trim() : "";
      var textoUrls = (cols.urls !== null && cols.urls !== undefined && row[cols.urls] != null) ? String(row[cols.urls]) : "";
      coincidencias.push({
        fecha: fecha,
        actividad: actividad,
        subactividad: subactividad,
        obligacion: _construirObligacion(actividad, subactividad),
        detalle: detalle,
        evidencias: _extraerUrlsMemo(textoUrls)
      });
    }
  }

  // Solo las fallas del ciclo actual (ver docstring): las últimas (coincidencias.length % 3).
  var enCicloActual = coincidencias.length % 3;
  return coincidencias.slice(Math.max(0, coincidencias.length - enCicloActual));
}

/**
 * Formatea una fecha a "dd/mm/yyyy". Acepta un objeto Date (o Date-like con getDate) o texto
 * "dd/mm/yyyy hh:mm[:ss]" / "dd/mm/yyyy". Si no reconoce el formato, intenta parsear como Date;
 * si tampoco, devuelve el texto original.
 */
function formatearFechaMemo(valor) {
  if (valor === null || valor === undefined || valor === "") return "";

  if (valor instanceof Date || (typeof valor === "object" && typeof valor.getDate === "function")) {
    var d = valor;
    var dd = ("0" + d.getDate()).slice(-2);
    var mm = ("0" + (d.getMonth() + 1)).slice(-2);
    var yyyy = d.getFullYear();
    return dd + "/" + mm + "/" + yyyy;
  }

  var s = String(valor).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    var dia = ("0" + m[1]).slice(-2);
    var mes = ("0" + m[2]).slice(-2);
    return dia + "/" + mes + "/" + m[3];
  }

  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return formatearFechaMemo(parsed);
  }

  return s;
}

/**
 * Distribuye hasta 3 fechas de incumplimiento (textos "dd/mm/yyyy" vía formatearFechaMemo, o ""
 * si no aplica) en los 3 slots {{Fecha Incumplimiento 1/2/3}} de la plantilla, según el nivel de
 * la alerta: n = normalizarNivelAlertaMemo(nivel); se toman las últimas (n-1) de `fallasPrevias`
 * (ninguna si n=1); el slot i (0-based) con i < n-1 recibe previas[i].fecha (o "" si falta); el
 * slot n-1 recibe `fechaActual`; el resto queda "". Para una 3era alerta (nivel 3, el único caso
 * que existía antes de este cambio) con las 2 previas de siempre, el resultado es idéntico al que
 * construirReemplazosMemo armaba antes directamente.
 */
function distribuirFechasIncumplimiento(fallasPrevias, fechaActual, nivel) {
  var n = normalizarNivelAlertaMemo(nivel);
  var fallas = Array.isArray(fallasPrevias) ? fallasPrevias : [];
  var cantidadPrevias = n - 1;
  var previas = cantidadPrevias > 0 ? fallas.slice(Math.max(0, fallas.length - cantidadPrevias)) : [];

  var out = ["", "", ""];
  for (var i = 0; i < cantidadPrevias; i++) {
    out[i] = previas[i] ? formatearFechaMemo(previas[i].fecha) : "";
  }
  out[n - 1] = formatearFechaMemo(fechaActual);
  return out;
}

/**
 * Arma la frase "la fecha A" / "las fechas A y B" / "las fechas A, B y C" a partir de los textos
 * no vacíos de `fechas` (ver distribuirFechasIncumplimiento); "" si no hay ninguna. La usa
 * _generarMemoFila para reemplazar la frase completa del párrafo de ANTECEDENTES de la plantilla
 * (ver fraseFechasIncumplimientoPlantillaMemo) cuando el memo es de 1era o 2da alerta y no las 3
 * fechas aplican.
 */
function construirFraseFechasIncumplimiento(fechas) {
  var noVacias = (Array.isArray(fechas) ? fechas : []).filter(function(f) { return f !== "" && f != null; });
  if (noVacias.length === 0) return "";
  if (noVacias.length === 1) return "la fecha " + noVacias[0];
  if (noVacias.length === 2) return "las fechas " + noVacias[0] + " y " + noVacias[1];
  return "las fechas " + noVacias.slice(0, -1).join(", ") + " y " + noVacias[noVacias.length - 1];
}

/**
 * Construye el mapa marcador -> valor para reemplazar en la plantilla del memorando.
 * `datos`: {
 *   nombreNatural, cargo, gerencia, dni, fechaEnvio,
 *   fallasPrevias: [{fecha, obligacion}, ...] (0, 1 o 2 elementos, en orden cronológico),
 *   fallaActual: {fecha, obligacion},
 *   correoJefe,
 *   nivelAlerta (1, 2 o 3; ausente = 3, ver normalizarNivelAlertaMemo)
 * }
 * NOTA: {{Firma Colaborador}} y {{Fecha y hora de recepción}} se DEJAN literales (sin entrada en
 * este mapa): los llena el flujo de firma (proyecto firma-memo) cuando el colaborador responde.
 * El PDF inicial se genera aparte, desde una copia temporal donde esos dos marcadores sí se
 * reemplazan por cadena vacía (ver _generarYEnviarUnMemo).
 */
function construirReemplazosMemo(datos) {
  datos = datos || {};
  var fallas = Array.isArray(datos.fallasPrevias) ? datos.fallasPrevias : [];
  var fallaActual = datos.fallaActual || {};
  var obligacionActual = fallaActual.obligacion || "";
  var fechasIncumplimiento = distribuirFechasIncumplimiento(fallas, fallaActual.fecha, datos.nivelAlerta);

  return {
    "{{Nombres, apellidos colaborador}}": datos.nombreNatural || "",
    "{{Cargo}}": datos.cargo || "",
    "{{Gerencia}}": datos.gerencia || "",
    "{{N.° de documento de identidad}}": datos.dni || "",
    "{{Fecha de envío}}": formatearFechaMemo(datos.fechaEnvio || new Date()),
    "{{Fecha Incumplimiento 1}}": fechasIncumplimiento[0],
    "{{Fecha  Incumplimiento  1}}": fechasIncumplimiento[0],
    "{{Fecha Incumplimiento 2}}": fechasIncumplimiento[1],
    "{{Fecha Incumplimiento 3}}": fechasIncumplimiento[2],
    "{{Actividad de la operación comprometida}}": obligacionActual,
    "{{Correo del jefe inmediato}}": datos.correoJefe || ""
  };
}

/**
 * Elige el ID de plantilla Google Doc del memo según el nivel de alerta: para 1era/2da alerta
 * (memo activado manualmente por el inspector) usa CONFIG.MEMO_PLANTILLA_ID_MANUAL si está
 * configurada (misma estructura de marcadores que la plantilla de 3era alerta, pero puede tener
 * un texto distinto); si está vacía, o la alerta es de 3era, usa siempre MEMO_PLANTILLA_ID. PURA:
 * recibe `cfg` (normalmente CONFIG) en vez de leerlo directamente, para poder testearla en Node.
 */
function elegirPlantillaMemo(nivel, cfg) {
  var CFG = cfg || {};
  var n = normalizarNivelAlertaMemo(nivel);
  if (n !== 3 && CFG.MEMO_PLANTILLA_ID_MANUAL) return CFG.MEMO_PLANTILLA_ID_MANUAL;
  return CFG.MEMO_PLANTILLA_ID;
}

// ============================================================================
// Funciones de integración con Apps Script (Drive/Docs/Mail/Sheets).
// No se ejecutan en Node; las globals de Apps Script se resuelven dentro de cada función
// (nunca al cargar el módulo) para que `require('./Memo.js')` no falle en Node.
// ============================================================================

/**
 * Escapa un marcador de plantilla ("{{Marcador}}") para usarlo como regex en Body.replaceText,
 * tolerando espacios múltiples/dobles entre palabras (algunos marcadores de la plantilla real
 * tienen doble espacio, ej. "{{Fecha  Incumplimiento  1}}") y escapando caracteres regex
 * especiales (incluye "." usado en "N.°").
 */
function _escaparRegexMarcadorMemo(marcador) {
  var esc = String(marcador).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  esc = esc.replace(/ +/g, "\\s+");
  return esc;
}

/**
 * Reemplaza todos los marcadores del mapa `reemplazos` en un elemento de Documento (Body, Header o
 * Footer) que exponga replaceText, tolerando variantes con espacios múltiples.
 */
function _reemplazarMarcadoresEnElemento(elemento, reemplazos) {
  if (!elemento) return;
  Object.keys(reemplazos).forEach(function(marcador) {
    try {
      elemento.replaceText(_escaparRegexMarcadorMemo(marcador), String(reemplazos[marcador] == null ? "" : reemplazos[marcador]));
    } catch (e) {
      Logger.log("Memo: no se pudo reemplazar marcador " + marcador + ": " + e.toString());
    }
  });
}

/**
 * Frase literal del párrafo ANTECEDENTES de la plantilla del memo que menciona las 3 fechas de
 * incumplimiento: "Acumulación de alertas por incumplimiento en las fechas {{Fecha Incumplimiento
 * 1}}, {{Fecha Incumplimiento 2}} y {{Fecha Incumplimiento 3}}, en relación con sus obligaciones
 * como {{Cargo}}." En un memo de 1era o 2da alerta, reemplazar solo los 3 marcadores individuales
 * dejaría huecos sueltos (ej. "en las fechas 01/10/2026,  y ,"), así que _generarMemoFila
 * reemplaza esta frase COMPLETA (ver construirFraseFechasIncumplimiento) ANTES del reemplazo por
 * marcador de siempre. Si la plantilla cambia y ya no trae esta frase literal, el replaceText de
 * esta frase simplemente no encuentra nada y el reemplazo por marcador de siempre sigue
 * funcionando igual.
 */
function fraseFechasIncumplimientoPlantillaMemo() {
  return "las fechas {{Fecha Incumplimiento 1}}, {{Fecha Incumplimiento 2}} y {{Fecha Incumplimiento 3}}";
}

/**
 * Obtiene (o crea) la subcarpeta CONFIG.MEMO_NOMBRE_CARPETA dentro de la carpeta raíz de reportes.
 * Si no se puede resolver la raíz, usa la raíz de reportes directamente (o null si tampoco existe).
 */
function _obtenerCarpetaMemos() {
  var CFG = _configMemo();
  try {
    if (!CFG.ID_CARPETA_REPORTES) return null;
    var carpetaRoot = DriveApp.getFolderById(CFG.ID_CARPETA_REPORTES);
    try {
      return obtenerOCrearCarpetaHija(carpetaRoot, CFG.MEMO_NOMBRE_CARPETA || "Memorandos");
    } catch (eHija) {
      Logger.log("Memo: no se pudo crear/obtener subcarpeta de memorandos, se usa la raíz: " + eHija.toString());
      return carpetaRoot;
    }
  } catch (e) {
    Logger.log("Memo: no se pudo abrir la carpeta raíz de reportes: " + e.toString());
    return null;
  }
}

/**
 * Cabecera completa (23 columnas originales + 17 columnas del flujo de firma, contrato compartido
 * con el proyecto firma-memo, que localiza columnas POR NOMBRE de encabezado + 4 columnas de la
 * revisión previa del inspector + 1 columna de nivel de alerta, agregadas AL FINAL). Función pura
 * para poder testearla y reutilizarla en la migración idempotente de _obtenerHojaMemorandos.
 */
function construirCabeceraMemorandos() {
  return [
    'Fecha generación', 'Negocio', 'Operación', 'ID_Subactividad', 'Actividad', 'Subactividad',
    'Colaborador', 'Email colaborador', 'DNI', 'Cargo', 'Área', 'Sede', 'Jefe inmediato',
    'Correo jefe', 'Fecha incumplimiento 1', 'Fecha incumplimiento 2', 'Fecha incumplimiento 3',
    'Destinatarios', 'Estado', 'URL Doc', 'URL PDF', 'Detalle', 'Inspector',
    'ID Memo', 'Token', 'ID Doc', 'ID PDF', 'Gerencia', 'Fecha envío ISO', 'Estado respuesta',
    'Fecha/hora respuesta', 'Fuera de plazo', 'Tipo firma', 'Nombre firmante', 'DNI firmante',
    'Cuenta que respondió', 'URL imagen firma', 'URL PDF firmado', 'Descargo', 'Notificados respuesta',
    'Datos amonestación', 'Revisor', 'Fecha revisión', 'Motivo descarte', 'Nivel alerta'
  ];
}

/**
 * Asegura que exista la hoja CONFIG.HOJA_MEMORANDOS con su cabecera (en negrita) y devuelve el Sheet.
 * Migración idempotente: si la hoja ya existe con la cabecera antigua (23 columnas), agrega al
 * final las columnas nuevas del flujo de firma, SIN tocar las filas existentes.
 */
function _obtenerHojaMemorandos(ss) {
  var CFG = _configMemo();
  var nombreHoja = CFG.HOJA_MEMORANDOS || "Memorandos";
  var sheet = ss.getSheetByName(nombreHoja);
  var cabecera = construirCabeceraMemorandos();
  if (!sheet) {
    sheet = ss.insertSheet(nombreHoja);
    sheet.appendRow(cabecera);
    sheet.getRange("1:1").setFontWeight("bold");
    try { sheet.getRange(2, 9, sheet.getMaxRows() - 1, 1).setNumberFormat('@'); } catch (eFmt) {} // DNI como texto
  } else {
    try {
      var anchoActual = sheet.getLastColumn();
      var cabeceraActual = anchoActual > 0 ? sheet.getRange(1, 1, 1, anchoActual).getValues()[0] : [];
      if (cabeceraActual.length < cabecera.length) {
        var faltantes = cabecera.slice(cabeceraActual.length);
        sheet.getRange(1, cabeceraActual.length + 1, 1, faltantes.length).setValues([faltantes]);
        sheet.getRange(1, 1, 1, cabecera.length).setFontWeight("bold");
      }
    } catch (eMig) {
      Logger.log("Memo: no se pudo migrar la cabecera de la hoja de memorandos: " + eMig.toString());
    }
  }
  return sheet;
}

function _registrarMemoEnHoja(ss, fila) {
  try {
    var sheet = _obtenerHojaMemorandos(ss);
    sheet.appendRow(fila);
    var ultimaFila = sheet.getLastRow();
    try { sheet.getRange(ultimaFila, 9, 1, 1).setNumberFormat('@'); } catch (eFmt) {}
  } catch (e) {
    Logger.log("Memo: no se pudo registrar en la hoja de memorandos: " + e.toString());
  }
}

// ============================================================================
// Filas de la hoja "Memorandos" (funciones PURAS: construyen los valores, no escriben nada).
// Se separan de la escritura para poder reutilizarlas tanto al AGREGAR una fila (generación normal)
// como al ACTUALIZAR la fila "Pendiente de revisión" existente (ver decidirPendientes).
// ============================================================================

/**
 * Fila (23 columnas) del memo cuyo colaborador no se identificó de forma única en cecos.
 * ctx: { fechaActual, inspector }
 */
function construirFilaMemoNoIdentificado(amon, ctx) {
  return [
    ctx.fechaActual, amon.negocio || "", amon.operacion || "", amon.idSubKey || "",
    amon.actividad || "", amon.subactividad || "", amon.nombre || "", amon.email || "",
    "", "", "", "", "", "", "", "", "", "", "Pendiente - colaborador no identificado",
    "", "", "No se encontró exactamente una coincidencia en cecos para este responsable.", ctx.inspector || ""
  ];
}

/**
 * Fila (23 columnas) del memo que no se envió porque no hay destinatarios del Rol activo.
 */
function construirFilaMemoSinDestinatarios(amon, colaborador, ctx, rolDestinatarioActivo) {
  return [
    ctx.fechaActual, amon.negocio || "", amon.operacion || "", amon.idSubKey || "",
    amon.actividad || "", amon.subactividad || "", colaborador.nombre || amon.nombre || "",
    colaborador.email || amon.email || "", colaborador.dni || "", colaborador.cargo || "",
    colaborador.area || "", colaborador.sede || "", "", "", "", "", "", "",
    "Pendiente - sin destinatarios " + rolDestinatarioActivo + " configurados",
    "", "", "Configure filas con Rol '" + rolDestinatarioActivo + "' en Config_Personal.",
    ctx.inspector || ""
  ];
}

/**
 * Fila (23 columnas) cuando la generación/envío falló. `error` es el texto del error.
 */
function construirFilaMemoError(amon, ctx, error) {
  return [
    ctx.fechaActual, amon.negocio || "", amon.operacion || "", amon.idSubKey || "",
    amon.actividad || "", amon.subactividad || "", amon.nombre || "", amon.email || "",
    "", "", "", "", "", "", "", "", "", "", "Error - " + error, "", "", "", ctx.inspector || ""
  ];
}

/**
 * Fila completa (40 columnas) de un memo generado y enviado.
 * d: { amon, colaborador, jefe, destResult, fechaActual, inspector, obligacionActual, urlDoc,
 *      urlPdf, idMemo, token, idDoc, idPdf, fechaEnvioIso }
 */
function construirFilaMemoEnviado(d) {
  var amon = d.amon;
  var colaborador = d.colaborador;
  var jefe = d.jefe;
  var fechasIncumplimiento = distribuirFechasIncumplimiento(amon.fallasPrevias, amon.fecha, amon.nivelAlerta);
  return [
    d.fechaActual, amon.negocio || "", amon.operacion || "", amon.idSubKey || "",
    amon.actividad || "", amon.subactividad || "", colaborador.nombre || "", colaborador.email || "",
    colaborador.dni || "", colaborador.cargo || "", colaborador.area || "", colaborador.sede || "",
    jefe ? jefe.nombre : "", jefe ? jefe.email : "",
    fechasIncumplimiento[0], fechasIncumplimiento[1], fechasIncumplimiento[2],
    formatearDestinatariosMemo(d.destResult), "Enviado", d.urlDoc, d.urlPdf, d.obligacionActual, d.inspector || "",
    d.idMemo, d.token, d.idDoc, d.idPdf, colaborador.gerencia || "", d.fechaEnvioIso, "Pendiente",
    "", "", "", "", "", "", "", "", ""
  ];
}

/**
 * Combina una fila nueva (resultado de la generación, columnas posicionales de la cabecera canónica)
 * sobre una fila EXISTENTE de la hoja: las columnas que la fila nueva cubre se sobrescriben; las
 * demás (ID Memo si la fila nueva es corta, "Datos amonestación", etc.) se conservan.
 * `extras` = { "Nombre de columna": valor } se aplica por nombre de cabecera (Revisor, Fecha revisión...).
 * Devuelve un array con el ancho de `cabecera`. PURA.
 */
function combinarFilaMemoExistente(cabecera, filaExistente, filaNueva, extras) {
  var out = [];
  var existente = Array.isArray(filaExistente) ? filaExistente : [];
  var nueva = Array.isArray(filaNueva) ? filaNueva : [];
  for (var i = 0; i < cabecera.length; i++) {
    out[i] = i < existente.length ? existente[i] : "";
  }
  for (var j = 0; j < nueva.length && j < cabecera.length; j++) {
    out[j] = nueva[j];
  }
  Object.keys(extras || {}).forEach(function(nombre) {
    var idx = cabecera.indexOf(nombre);
    if (idx > -1) out[idx] = extras[nombre];
  });
  return out;
}

/**
 * Construye una fila completa (ancho de la cabecera canónica) a partir de { "Columna": valor }.
 */
function construirFilaMemoPorNombre(valoresPorNombre) {
  var cab = construirCabeceraMemorandos();
  return cab.map(function(nombre) {
    var v = valoresPorNombre ? valoresPorNombre[nombre] : undefined;
    return v === undefined || v === null ? "" : v;
  });
}

// ============================================================================
// Revisión humana previa (estado "Pendiente de revisión"): funciones PURAS.
// ============================================================================

/**
 * Convierte una fecha (Date o texto) a texto que sobrevive a JSON: los Date van como
 * "dd/MM/yyyy HH:mm" (hora local del script, America/Lima); el texto se devuelve recortado.
 * formatearFechaMemo entiende ese formato.
 */
function fechaATextoMemo(valor) {
  if (valor === null || valor === undefined || valor === "") return "";
  if (valor instanceof Date || (typeof valor === "object" && typeof valor.getDate === "function")) {
    if (isNaN(valor.getTime())) return "";
    return ("0" + valor.getDate()).slice(-2) + "/" + ("0" + (valor.getMonth() + 1)).slice(-2) + "/" +
      valor.getFullYear() + " " + ("0" + valor.getHours()).slice(-2) + ":" + ("0" + valor.getMinutes()).slice(-2);
  }
  return String(valor).trim();
}

function _listaTextosMemo(lista) {
  return (Array.isArray(lista) ? lista : []).map(function(u) { return String(u == null ? "" : u).trim(); })
    .filter(function(u) { return u !== "" && u !== "-"; });
}

function _faltaASerializableMemo(f) {
  f = f || {};
  return {
    fecha: fechaATextoMemo(f.fecha),
    actividad: String(f.actividad == null ? "" : f.actividad),
    subactividad: String(f.subactividad == null ? "" : f.subactividad),
    obligacion: f.obligacion != null && f.obligacion !== "" ? String(f.obligacion) : _construirObligacion(f.actividad, f.subactividad),
    detalle: String(f.detalle == null ? "" : f.detalle).trim(),
    evidencias: _listaTextosMemo(f.evidencias)
  };
}

/**
 * Serializa la amonestación completa (incluidas sus fallasPrevias) a JSON con las fechas como texto,
 * para guardarla en la columna "Datos amonestación" y poder generar el memo después.
 */
function serializarDatosAmonestacion(amon) {
  amon = amon || {};
  return JSON.stringify({
    negocio: amon.negocio || "",
    operacion: amon.operacion || "",
    idSubKey: amon.idSubKey == null ? "" : String(amon.idSubKey),
    actividad: amon.actividad || "",
    subactividad: amon.subactividad || "",
    fecha: fechaATextoMemo(amon.fecha),
    detalle: String(amon.detalle == null ? "" : amon.detalle).trim(),
    evidencias: _listaTextosMemo(amon.evidencias),
    nombre: amon.nombre || "",
    email: amon.email || "",
    nivelAlerta: normalizarNivelAlertaMemo(amon.nivelAlerta),
    fallasPrevias: (Array.isArray(amon.fallasPrevias) ? amon.fallasPrevias : []).map(_faltaASerializableMemo)
  });
}

/**
 * Inverso de serializarDatosAmonestacion. Devuelve el objeto amonestación (fechas como texto
 * "dd/MM/yyyy HH:mm", válidas para formatearFechaMemo) o null si el texto no es un JSON válido.
 * `nivelAlerta` siempre se normaliza (1, 2 o 3); si el JSON es de antes de este campo, queda 3
 * (todas las filas "Datos amonestación" previas a este cambio son de 3era alerta).
 */
function deserializarDatosAmonestacion(texto) {
  var obj;
  try {
    obj = JSON.parse(String(texto == null ? "" : texto));
  } catch (e) {
    return null;
  }
  if (!obj || typeof obj !== "object" || Array.isArray(obj)) return null;
  obj.fallasPrevias = Array.isArray(obj.fallasPrevias) ? obj.fallasPrevias : [];
  obj.evidencias = Array.isArray(obj.evidencias) ? obj.evidencias : [];
  obj.nivelAlerta = normalizarNivelAlertaMemo(obj.nivelAlerta);
  return obj;
}

/**
 * Fila de "Memorandos" en estado "Pendiente de revisión": SIN Token (así firma-memo no acepta
 * respuestas de un memo que aún no se envió). ctx: { fechaActual, inspector, idMemo }.
 */
function construirFilaMemoPendiente(amon, ctx) {
  var CFG = _configMemo();
  return construirFilaMemoPorNombre({
    'Fecha generación': ctx.fechaActual,
    'Negocio': amon.negocio || "",
    'Operación': amon.operacion || "",
    'ID_Subactividad': amon.idSubKey || "",
    'Actividad': amon.actividad || "",
    'Subactividad': amon.subactividad || "",
    'Colaborador': amon.nombre || "",
    'Email colaborador': amon.email || "",
    'Estado': CFG.MEMO_ESTADO_PENDIENTE_REVISION || "Pendiente de revisión",
    'Inspector': ctx.inspector || "",
    'ID Memo': ctx.idMemo,
    'Datos amonestación': serializarDatosAmonestacion(amon),
    'Nivel alerta': normalizarNivelAlertaMemo(amon.nivelAlerta)
  });
}

function _valorCeldaTextoMemo(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date || (typeof v === "object" && typeof v.getDate === "function")) return fechaATextoMemo(v);
  return String(v).trim();
}

/**
 * Convierte una fila de la hoja en objeto { "Columna": valor } usando la cabecera de la hoja.
 */
function objetoDesdeFilaMemo(cabecera, fila) {
  var o = {};
  for (var i = 0; i < cabecera.length; i++) {
    var nombre = String(cabecera[i] == null ? "" : cabecera[i]).trim();
    if (nombre) o[nombre] = fila[i] === undefined ? "" : fila[i];
  }
  return o;
}

function _esEstadoPendienteRevisionMemo(estado) {
  var CFG = _configMemo();
  return _normTextoMemo(estado) === _normTextoMemo(CFG.MEMO_ESTADO_PENDIENTE_REVISION || "Pendiente de revisión");
}

function _mismaCuentaMemo(inspector, cuentaActiva) {
  var a = String(inspector == null ? "" : inspector).trim().toLowerCase();
  var b = String(cuentaActiva == null ? "" : cuentaActiva).trim().toLowerCase();
  return a !== "" && a === b;
}

/**
 * Arma los datos que se MUESTRAN de una fila pendiente. Nunca incluye DNI ni datos de cecos.
 * INTEGRACIÓN OPCIÓN A: Re-calcula dinámicamente las faltas leyendo Historial_Inspecciones en tiempo real.
 */
function construirVistaPendienteMemo(obj) {
  var amon = deserializarDatosAmonestacion(obj['Datos amonestación']);
  var faltas = [];
  if (amon) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var layoutHistorial = (typeof resolverLayoutHistorial === "function") 
        ? resolverLayoutHistorial(amon.negocio, ss) 
        : null;
        
      if (layoutHistorial && layoutHistorial.sheet) {
        var registrosHistorial = layoutHistorial.sheet.getDataRange().getValues();
        var fallasActualizadas = obtenerFallasCicloActual(
          registrosHistorial, 
          amon.negocio, 
          amon.operacion, 
          amon.idSubKey, 
          amon.nombre, 
          layoutHistorial.cols
        );
        amon.fallasPrevias = fallasActualizadas;
      }
    } catch (eRecalc) {
      Logger.log("Memo: Error recalculando fallas en tiempo real: " + eRecalc.toString());
    }

    amon.fallasPrevias.slice(0, 2).forEach(function(f) {
      faltas.push({
        fecha: formatearFechaMemo(f.fecha),
        obligacion: f.obligacion || "",
        detalle: f.detalle || "",
        evidencias: _listaTextosMemo(f.evidencias)
      });
    });
    faltas.push({
      fecha: formatearFechaMemo(amon.fecha),
      obligacion: _construirObligacion(amon.actividad, amon.subactividad),
      detalle: amon.detalle || "",
      evidencias: _listaTextosMemo(amon.evidencias)
    });
  }
  return {
    idMemo: String(obj['ID Memo'] || ""),
    fecha: _valorCeldaTextoMemo(obj['Fecha generación']),
    negocio: _valorCeldaTextoMemo(obj['Negocio']),
    operacion: _valorCeldaTextoMemo(obj['Operación']),
    colaborador: _valorCeldaTextoMemo(obj['Colaborador']),
    actividad: _valorCeldaTextoMemo(obj['Actividad']),
    subactividad: _valorCeldaTextoMemo(obj['Subactividad']),
    faltas: faltas,
    generable: !!amon,
    nivelAlerta: amon ? amon.nivelAlerta : 3
  };
}

/**
 * De la matriz completa de la hoja (con cabecera) devuelve los memos en estado "Pendiente de revisión".
 * Si esSupervisor es true, devuelve los pendientes de TODOS los inspectores.
 * Si esSupervisor es false, sólo devuelve los del inspector activo.
 */
function filtrarPendientesRevision(valores, cuentaActiva, esSupervisor) {
  if (!Array.isArray(valores) || valores.length < 2) return [];
  var cab = valores[0];
  var out = [];
  for (var i = 1; i < valores.length; i++) {
    var fila = valores[i];
    if (!fila) continue;
    var obj = objetoDesdeFilaMemo(cab, fila);
    if (!_esEstadoPendienteRevisionMemo(obj['Estado'])) continue;
    
    // Si NO es supervisor, restringimos a sus propios memos
    if (!esSupervisor && !_mismaCuentaMemo(obj['Inspector'], cuentaActiva)) continue;
    if (!obj['ID Memo']) continue;
    
    out.push(construirVistaPendienteMemo(obj));
  }
  return out;
}

/**
 * Busca por "ID Memo" en la matriz completa. Devuelve { indice (0-based en `valores`), obj } o null.
 */
function buscarFilaMemoPorId(valores, idMemo) {
  var id = String(idMemo == null ? "" : idMemo).trim();
  if (!id || !Array.isArray(valores) || valores.length < 2) return null;
  var cab = valores[0];
  var idx = -1;
  for (var c = 0; c < cab.length; c++) {
    if (String(cab[c] == null ? "" : cab[c]).trim() === 'ID Memo') { idx = c; break; }
  }
  if (idx === -1) return null;
  for (var i = 1; i < valores.length; i++) {
    if (valores[i] && String(valores[i][idx] == null ? "" : valores[i][idx]).trim() === id) {
      return { indice: i, obj: objetoDesdeFilaMemo(cab, valores[i]) };
    }
  }
  return null;
}

function validarDecisionRevision(decision, filaObj, cuentaActiva, esSupervisor) {
  if (!decision || typeof decision !== "object" || !String(decision.idMemo == null ? "" : decision.idMemo).trim()) {
    return { ok: false, codigo: "rechazada", mensaje: "Decisión sin identificador de memo." };
  }
  var tipo = String(decision.decision == null ? "" : decision.decision).trim().toLowerCase();
  if (tipo !== "generar" && tipo !== "descartar") {
    return { ok: false, codigo: "rechazada", mensaje: "La decisión debe ser 'generar' o 'descartar'." };
  }
  if (tipo === "descartar" && String(decision.motivo == null ? "" : decision.motivo).trim().length < 5) {
    return { ok: false, codigo: "rechazada", mensaje: "Para no generar el memo debes indicar un motivo (mínimo 5 caracteres)." };
  }
  if (!filaObj) {
    return { ok: false, codigo: "omitida", mensaje: "No se encontró el memo pendiente." };
  }
  if (!_esEstadoPendienteRevisionMemo(filaObj['Estado'])) {
    return { ok: false, codigo: "omitida", mensaje: "Este memo ya no está pendiente de revisión (estado: " + _valorCeldaTextoMemo(filaObj['Estado']) + ")." };
  }
  
  // Si no es supervisor, validamos que la cuenta coincida con el inspector original
  if (!esSupervisor && !_mismaCuentaMemo(filaObj['Inspector'], cuentaActiva)) {
    return { ok: false, codigo: "omitida", mensaje: "Este memo corresponde a otro inspector." };
  }
  return { ok: true, codigo: "", mensaje: "" };
}

function _escHtmlMemo(t) {
  return String(t == null ? "" : t)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * Arma asunto y cuerpo HTML del correo de CONFIRMACIÓN al inspector con los memorandos que se
 * están preparando (estado "Pendiente de revisión") a partir de UN envío de checklist. Ninguno se
 * ha enviado todavía: el inspector debe entrar a Revision.html, marcar a quiénes sí se les envía y
 * presionar "Enviar memos". opts: { urlRevision ("" si no hay), modoPrueba, negocio, operacion }.
 * Estilo sobrio SHARF (igual que el correo del memo). PURA.
 */

/**
 * Verifica si un correo tiene un rol específico (o supervisión) en Config_Personal.
 * Permite saber si el usuario puede ver todos los memos pendientes de todos los inspectores.
 */
function esSupervisorMemo(correo, cecosData, valoresConfigPersonal) {
  if (!correo || !Array.isArray(valoresConfigPersonal)) return false;
  var correoNorm = String(correo).trim().toLowerCase();
  var CFG = _configMemo();
  var rolSupervisorNorm = _normTextoMemo(CFG.MEMO_ROL_SUPERVISOR || "Supervisores_Memo");

  for (var i = 1; i < valoresConfigPersonal.length; i++) {
    var row = valoresConfigPersonal[i];
    if (!row) continue;
    var correoNombreFila = row[0] ? String(row[0]).trim() : "";
    var rolFila = row[1] ? String(row[1]).trim() : "";

    if (!correoNombreFila) continue;
    if (_normTextoMemo(rolFila) !== rolSupervisorNorm) continue;

    if (correoNombreFila.indexOf("@") > -1) {
      if (correoNombreFila.toLowerCase() === correoNorm) return true;
    } else {
      var colab = buscarColaborador(cecosData, null, correoNombreFila);
      if (colab && colab.email && String(colab.email).toLowerCase() === correoNorm) return true;
    }
  }
  return false;
}

function construirCorreoRevisionMemo(amonestaciones, opts) {
  opts = opts || {};
  var CFG = _configMemo();
  var colorInk = CFG.COLOR_INK || "#111111";
  var colorBoton = CFG.COLOR_GREEN_DEEP || "#003C46";
  var lista = Array.isArray(amonestaciones) ? amonestaciones : [];
  var destino = (opts.negocio || "") + (opts.operacion ? "/" + opts.operacion : "");
  var asunto = (opts.modoPrueba ? "[PRUEBA] " : "") + "Confirma el envío de memorandos" + (destino ? " - " + destino : "");

  var items = lista.map(function(amon) {
    var fechasTexto = distribuirFechasIncumplimiento(amon.fallasPrevias, amon.fecha, amon.nivelAlerta)
      .filter(function(f) { return f !== ""; });
    var nivelNorm = normalizarNivelAlertaMemo(amon.nivelAlerta);
    var textoNivel = textoNivelAlertaMemo(nivelNorm) + (nivelNorm < 3 ? " · activado por ti" : "");
    return '<li style="margin-bottom:8px;">Para: <b>' + _escHtmlMemo(amon.nombre) + '</b><br>' +
      _escHtmlMemo(_construirObligacion(amon.actividad, amon.subactividad) || "-") + '<br>' +
      _escHtmlMemo(textoNivel) + '<br>' +
      'Fechas de incumplimiento: ' + _escHtmlMemo(fechasTexto.join(", ") || "-") + '</li>';
  }).join("");

  var bloqueAccion;
  if (opts.urlRevision) {
    bloqueAccion =
      '<p>Entra, marca a quiénes sí se les enviará el memorando y presiona <b>Enviar memos</b>.</p>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:24px 0;"><tr><td>' +
        '<a href="' + _escHtmlMemo(opts.urlRevision) + '" target="_blank" style="display:inline-block;background-color:' + colorBoton + ';color:#FFFFFF;text-decoration:none;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;padding:12px 24px;border-radius:4px;">Revisar y enviar memos</a>' +
      '</td></tr></table>';
  } else {
    bloqueAccion =
      '<p>No se pudo generar el enlace directo. Abre la aplicación del checklist y agrega <b>?vista=revision</b> al final de la dirección para revisar los memorandos pendientes.</p>';
  }

  var cuerpo =
    '<div style="background-color:#FFFFFF;font-family:Arial,Helvetica,sans-serif;color:' + colorInk + ';font-size:14px;line-height:1.5;padding:16px;">' +
      '<p>Se están generando los siguientes memorandos a partir del checklist que enviaste' + (destino ? ' (' + _escHtmlMemo(destino) + ')' : '') +
      '. Todavía no se ha enviado ninguno:</p>' +
      '<ul style="padding-left:20px;">' + items + '</ul>' +
      bloqueAccion +
      '<p style="color:#3A3A3A;font-size:12px;">Mientras no presiones Enviar memos, no se envía ningún memorando.</p>' +
    '</div>';

  return { asunto: asunto, htmlBody: cuerpo };
}

/**
 * Amonestaciones de EJEMPLO (2 responsables en 3era alerta, con sus 2 faltas previas del ciclo) para
 * probarCorreoConfirmacionMemos. Las fechas se calculan hacia atrás desde `fechaBase` (hoy, -7 y -14
 * días). Datos ficticios: no salen del historial ni de cecos. PURA.
 */
function construirAmonestacionesEjemploMemo(fechaBase) {
  var base = (fechaBase instanceof Date) ? fechaBase : new Date();
  function haceDias(n) { return new Date(base.getFullYear(), base.getMonth(), base.getDate() - n); }
  function ejemplo(nombre, actividad, subactividad) {
    return {
      negocio: "Distribución", operacion: "PLC", idSubKey: "0.0",
      actividad: actividad, subactividad: subactividad,
      fecha: base, nombre: nombre, email: "", nivelAlerta: 3,
      fallasPrevias: [
        { fecha: haceDias(14), actividad: actividad, subactividad: subactividad, obligacion: _construirObligacion(actividad, subactividad) },
        { fecha: haceDias(7), actividad: actividad, subactividad: subactividad, obligacion: _construirObligacion(actividad, subactividad) }
      ]
    };
  }
  return [
    ejemplo("Juan Pérez", "ORDEN Y LIMPIEZA", "Pasillos y zonas de tránsito despejados"),
    ejemplo("Ana Torres", "DESPACHO", "Verificación de guías antes de la salida del vehículo")
  ];
}

// ============================================================================
// Funciones de integración con Apps Script (Drive/Docs/Mail/Sheets).
// ============================================================================

/**
 * Actualiza UNA fila existente de la hoja (número de fila 1-based) con `fila` completa.
 */
function _escribirFilaMemoExistente(sheet, numeroFila, fila) {
  try { sheet.getRange(numeroFila, 9, 1, 1).setNumberFormat('@'); } catch (eFmt) {} // DNI como texto
  sheet.getRange(numeroFila, 1, 1, fila.length).setValues([fila]);
}

/**
 * Lee cecos y destinatarios (Prueba_Memo / Memos) y arma el contexto compartido de generación.
 */
function _prepararContextoMemo(ss, contexto) {
  var CFG = _configMemo();

  var cecosData = [];
  try {
    var shCecos = (typeof detectarHojaCecos === "function") ? detectarHojaCecos(ss) : null;
    if (shCecos) cecosData = leerCecosComoObjetos(shCecos.getDataRange().getDisplayValues());
  } catch (eCecos) {
    Logger.log("Memo: error leyendo hoja cecos: " + eCecos.toString());
  }

  var destinatariosPrueba = [];
  var destinatariosOficiales = [];
  try {
    var shPersonal = ss.getSheetByName(CFG.HOJA_PERSONAL) || ss.getSheetByName(CFG.HOJA_PERSONAL_ALT);
    if (shPersonal) {
      var valoresPersonal = shPersonal.getDataRange().getDisplayValues();
      destinatariosPrueba = obtenerDestinatariosMemo(valoresPersonal, cecosData, CFG.MEMO_ROL_DESTINATARIO);
      destinatariosOficiales = obtenerDestinatariosMemo(valoresPersonal, cecosData, CFG.MEMO_ROL_OFICIAL);
    }
  } catch (eDest) {
    Logger.log("Memo: error obteniendo destinatarios Prueba_Memo/Memos: " + eDest.toString());
  }

  return {
    ss: ss,
    fechaActual: (contexto && contexto.fechaActual) || new Date(),
    inspector: contexto && contexto.inspector,
    cecosData: cecosData,
    destinatariosPrueba: destinatariosPrueba,
    destinatariosOficiales: destinatariosOficiales
  };
}

/**
 * Genera (Doc + PDF) y envía el memorando de UNA amonestación y DEVUELVE la fila resultante
 * (sin escribirla en la hoja). Estado de la fila: "Enviado", "Pendiente - colaborador no
 * identificado" o "Pendiente - sin destinatarios ...". Lanza excepción si algo falla.
 *
 * amon: {
 *   negocio, operacion, idSubKey, actividad, subactividad, fecha (fecha del incumplimiento actual),
 *   nombre, email (responsable), fallasPrevias: [{fecha, obligacion}, ...]
 * }
 * contexto: { ss, negocio, operacion, fechaActual, inspector, cecosData, destinatarios, idMemo? }
 * `contexto.idMemo` (opcional) conserva el ID Memo de una fila pendiente de revisión.
 */
function _generarMemoFila(amon, contexto) {
  var CFG = _configMemo();
  var cecosData = contexto.cecosData || [];
  var destinatariosPrueba = contexto.destinatariosPrueba || [];
  var destinatariosOficiales = contexto.destinatariosOficiales || [];
  var esModoPrueba = !!CFG.MEMO_MODO_PRUEBA;

  var obligacionActual = _construirObligacion(amon.actividad, amon.subactividad);
  var fallaActual = { fecha: amon.fecha, obligacion: obligacionActual };

  var colaborador = buscarColaborador(cecosData, amon.email, amon.nombre);
  if (!colaborador) {
    return construirFilaMemoNoIdentificado(amon, contexto);
  }

  // Rol de destinatarios "para" (to) según el modo: en pruebas, Prueba_Memo; en oficial, Memos.
  var rolDestinatarioActivo = esModoPrueba ? (CFG.MEMO_ROL_DESTINATARIO || "Prueba_Memo") : (CFG.MEMO_ROL_OFICIAL || "Memos");
  var destinatariosActivos = esModoPrueba ? destinatariosPrueba : destinatariosOficiales;
  if (!destinatariosActivos || destinatariosActivos.length === 0) {
    return construirFilaMemoSinDestinatarios(amon, colaborador, contexto, rolDestinatarioActivo);
  }

  var jefe = buscarJefe(cecosData, colaborador.dependencia);
  var nombreNatural = formatearNombreNatural(colaborador.nombre);

  // Fechas de incumplimiento (1/2/3) según el nivel de esta amonestación: para 3era alerta ocupa
  // los 3 slots igual que siempre; para 1era/2da (memo activado manualmente por el inspector) deja
  // vacíos los slots que no aplican (ver distribuirFechasIncumplimiento).
  var fechasIncumplimientoTexto = distribuirFechasIncumplimiento(amon.fallasPrevias || [], amon.fecha, amon.nivelAlerta);

  var reemplazos = construirReemplazosMemo({
    nombreNatural: nombreNatural,
    cargo: colaborador.cargo,
    gerencia: colaborador.gerencia,
    dni: colaborador.dni,
    fechaEnvio: contexto.fechaActual,
    fallasPrevias: amon.fallasPrevias || [],
    fallaActual: fallaActual,
    correoJefe: jefe ? jefe.email : "",
    nivelAlerta: amon.nivelAlerta
  });

  var idPlantilla = elegirPlantillaMemo(amon.nivelAlerta, CFG);
  var carpeta = _obtenerCarpetaMemos();
  var nombreDoc = "Memorando_" + sanitizarNombreArchivo(nombreNatural) + "_" + sanitizarNombreArchivo(amon.negocio) + "_" + sanitizarNombreArchivo(amon.operacion);

  var copia = carpeta
    ? DriveApp.getFileById(idPlantilla).makeCopy(nombreDoc, carpeta)
    : DriveApp.getFileById(idPlantilla).makeCopy(nombreDoc);
  var idDoc = copia.getId();

  // El Doc del memo CONSERVA literalmente {{Firma Colaborador}} y {{Fecha y hora de recepción}}
  // (los llena firma-memo cuando el colaborador responde); no están en el mapa `reemplazos`.
  var doc = DocumentApp.openById(idDoc);
  var body = doc.getBody();
  // ANTES del reemplazo por marcador: si la plantilla trae la frase completa del párrafo de
  // ANTECEDENTES, se reemplaza por la frase ya armada con solo las fechas que aplican (evita
  // "en las fechas 01/10/2026,  y ," en memos de 1era/2da alerta). Si la plantilla cambia y la
  // frase no está, replaceText no encuentra nada y el reemplazo por marcador de siempre sigue
  // funcionando igual; para 3era alerta la frase queda idéntica a la de hoy.
  var mapaFraseFechas = {};
  mapaFraseFechas[fraseFechasIncumplimientoPlantillaMemo()] = construirFraseFechasIncumplimiento(fechasIncumplimientoTexto);
  _reemplazarMarcadoresEnElemento(body, mapaFraseFechas);
  try { _reemplazarMarcadoresEnElemento(doc.getHeader(), mapaFraseFechas); } catch (eHF) {}
  try { _reemplazarMarcadoresEnElemento(doc.getFooter(), mapaFraseFechas); } catch (eFF) {}
  _reemplazarMarcadoresEnElemento(body, reemplazos);
  try { _reemplazarMarcadoresEnElemento(doc.getHeader(), reemplazos); } catch (eH) {}
  try { _reemplazarMarcadoresEnElemento(doc.getFooter(), reemplazos); } catch (eF) {}
  doc.saveAndClose();

  var urlDoc = copia.getUrl();
  var blobPdf = null;
  var urlPdf = "";
  var idPdf = "";
  try {
    // El PDF inicial se exporta desde una COPIA TEMPORAL del Doc ya rellenado, en la que
    // {{Firma Colaborador}} y {{Fecha y hora de recepción}} se reemplazan por "" (el Doc
    // original los conserva literales). La copia temporal se manda a la papelera al terminar.
    var copiaTemp = carpeta
      ? copia.makeCopy(nombreDoc + "_tmp_pdf", carpeta)
      : copia.makeCopy(nombreDoc + "_tmp_pdf");
    var docTemp = DocumentApp.openById(copiaTemp.getId());
    var marcadoresFirmaVacios = {
      "{{Firma Colaborador}}": "",
      "{{Fecha y hora de recepción}}": ""
    };
    _reemplazarMarcadoresEnElemento(docTemp.getBody(), marcadoresFirmaVacios);
    try { _reemplazarMarcadoresEnElemento(docTemp.getHeader(), marcadoresFirmaVacios); } catch (eHT) {}
    try { _reemplazarMarcadoresEnElemento(docTemp.getFooter(), marcadoresFirmaVacios); } catch (eFT) {}
    docTemp.saveAndClose();

    blobPdf = copiaTemp.getAs(MimeType.PDF);
    var archivoPdf = carpeta ? carpeta.createFile(blobPdf) : DriveApp.createFile(blobPdf);
    archivoPdf.setName(nombreDoc + ".pdf");
    urlPdf = archivoPdf.getUrl();
    idPdf = archivoPdf.getId();

    copiaTemp.setTrashed(true);
  } catch (ePdf) {
    Logger.log("Memo: no se pudo exportar/guardar el PDF del memorando: " + ePdf.toString());
  }

  // Identificadores del flujo de firma (ver contrato compartido con firma-memo).
  var idMemo = contexto.idMemo || Utilities.getUuid();
  var token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");

  var prefijoAsunto = esModoPrueba ? "[PRUEBA] " : "";
  var asunto = prefijoAsunto + "Memorando de Amonestación Escrita - " + nombreNatural + " - " + (amon.negocio || "") + "/" + (amon.operacion || "");

  var colorInk = CFG.COLOR_INK || "#111111";
  var colorRevisarDetalleFondo = CFG.COLOR_GREEN_DEEP || "#003C46";
  var colorNoConformeTexto = CFG.COLOR_RED_DEEP || "#68002B";
  var plazoHoras = CFG.MEMO_PLAZO_HORAS || 24;

  var urlBase = CFG.MEMO_URL_FIRMA;
  var tieneBotones = !!urlBase;
  var botonesHtml = "";
  if (tieneBotones) {
    var urlRevisarDetalle = construirUrlRespuestaMemo(urlBase, idMemo, token, "revisar detalle");
    botonesHtml =
      '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:24px 0;"><tr>' +
        '<td style="padding-right:12px;">' +
          '<a href="' + urlRevisarDetalle + '" target="_blank" style="display:inline-block;background-color:' + colorRevisarDetalleFondo + ';color:#FFFFFF;text-decoration:none;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;padding:12px 24px;border-radius:4px;">RevisarDetalle</a>' +
        '</td>' +
      '</tr></table>';
  } else {
    Logger.log("Memo: CONFIG.MEMO_URL_FIRMA está vacío; el correo se envía sin botón Revisar Detalle.");
  }

  // Nivel de alerta a mostrar en el correo ("3era Alerta", o "1era/2da Alerta (memo activado por
  // el inspector)" cuando el memo no es por la 3era alerta del ciclo). Solo se listan las fechas
  // de incumplimiento que aplican (las vacías de distribuirFechasIncumplimiento se omiten).
  var nivelAlertaNormalizado = normalizarNivelAlertaMemo(amon.nivelAlerta);
  var textoNivelCorreoMemo = textoNivelAlertaMemo(nivelAlertaNormalizado) +
    (nivelAlertaNormalizado < 3 ? ' (memo activado por el inspector)' : '');
  var liFechasIncumplimientoMemo = "";
  fechasIncumplimientoTexto.forEach(function(f, idxFecha) {
    if (f === "") return;
    liFechasIncumplimientoMemo += '<li><b>Fecha incumplimiento ' + (idxFecha + 1) + ':</b> ' + f + '</li>';
  });

  var cuerpoHtml =
    '<div style="background-color:#FFFFFF;font-family:Arial,Helvetica,sans-serif;color:' + colorInk + ';font-size:14px;line-height:1.5;padding:16px;">' +
      (esModoPrueba ? '<p style="color:' + colorNoConformeTexto + ';font-weight:bold;">Correo generado en FASE DE PRUEBAS (los destinatarios reales — colaborador y jefe inmediato — no lo reciben todavía).</p>' : '') +
      '<p>Se ha generado un memorando de amonestación escrita con el siguiente detalle:</p>' +
      '<ul>' +
        '<li><b>Colaborador:</b> ' + nombreNatural + '</li>' +
        '<li><b>Gerencia:</b> ' + (colaborador.gerencia || '-') + '</li>' +
        '<li><b>Subactividad:</b> ' + (obligacionActual || '-') + '</li>' +
        '<li><b>Nivel de alerta:</b> ' + textoNivelCorreoMemo + '</li>' +
        liFechasIncumplimientoMemo +
      '</ul>' +
      (tieneBotones
        ? '<p>El colaborador tiene un plazo de <b>' + plazoHoras + ' horas</b> desde este correo para responder.</p>' + botonesHtml
        : '<p>Aún no se ha configurado el enlace de firma (CONFIG.MEMO_URL_FIRMA); este correo se envía sin botones de respuesta.</p>') +
      '<p>Se adjunta el memorando en PDF.</p>' +
    '</div>';

  var destResult = construirDestinatariosMemo(
    esModoPrueba,
    destinatariosOficiales,
    destinatariosPrueba,
    colaborador.email || amon.email || "",
    jefe ? jefe.email : "",
    contexto.inspector || "",
    !!CFG.MEMO_BCC_MONITOREO
  );

  var emailConfig = { to: destResult.to.join(","), subject: asunto, htmlBody: cuerpoHtml };
  if (destResult.cc.length) emailConfig.cc = destResult.cc.join(",");
  if (destResult.bcc.length) emailConfig.bcc = destResult.bcc.join(",");
  if (blobPdf) emailConfig.attachments = [blobPdf];
  MailApp.sendEmail(emailConfig);

  var fechaEnvioIso = Utilities.formatDate(contexto.fechaActual, "America/Lima", "yyyy-MM-dd'T'HH:mm:ssXXX");

  return construirFilaMemoEnviado({
    amon: amon, colaborador: colaborador, jefe: jefe, destResult: destResult,
    fechaActual: contexto.fechaActual, inspector: contexto.inspector,
    obligacionActual: obligacionActual, urlDoc: urlDoc, urlPdf: urlPdf,
    idMemo: idMemo, token: token, idDoc: idDoc, idPdf: idPdf, fechaEnvioIso: fechaEnvioIso
  });
}

/**
 * Genera y envía el memorando de UNA amonestación y AGREGA su fila a la hoja "Memorandos"
 * (flujo directo, sin revisión). Se ejecuta dentro de su propio try/catch desde
 * generarYEnviarMemorandos: un fallo aquí no debe afectar a las demás amonestaciones.
 */
function _generarYEnviarUnMemo(amon, contexto) {
  var fila = _generarMemoFila(amon, contexto);
  _registrarMemoEnHoja(contexto.ss, fila);
}

/**
 * Genera y envía un memorando por cada amonestación detectada (responsable que llegó a la 3era
 * alerta), en su propio try/catch para que un fallo individual no afecte a las demás.
 * `amonestaciones`: array de objetos como los que arma Report.js (ver _generarMemoFila).
 * `contexto`: { negocio, operacion, fechaActual, inspector }.
 * NOTA: procesarInspeccion ya no la llama (usa registrarAmonestacionesPendientes + revisión del
 * inspector); se conserva como respaldo de generación directa.
 */
function generarYEnviarMemorandos(amonestaciones, contexto) {
  if (!amonestaciones || amonestaciones.length === 0) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ctxCompleto = _prepararContextoMemo(ss, contexto);

  amonestaciones.forEach(function(amon) {
    try {
      _generarYEnviarUnMemo(amon, ctxCompleto);
    } catch (eUno) {
      Logger.log("Memo: error generando/enviando memorando para " + (amon && amon.nombre) + ": " + eUno.toString());
      try {
        _registrarMemoEnHoja(ss, construirFilaMemoError(amon, ctxCompleto, eUno.toString()));
      } catch (eReg) {
        Logger.log("Memo: no se pudo registrar el error en la hoja de memorandos: " + eReg.toString());
      }
    }
  });
}

// ============================================================================
// Revisión humana: registro de pendientes + aviso al inspector + funciones de la página Revision.
// ============================================================================

/**
 * Registra cada amonestación como fila "Pendiente de revisión" en la hoja Memorandos (SIN Token:
 * firma-memo exige Token para aceptar una respuesta, así que un memo no enviado no puede ser
 * respondido) y envía UN correo al inspector con la lista y el botón "Revisar memos".
 * Sin amonestaciones no hace nada. Sin inspector (correo) solo registra y deja Logger.log.
 * `contexto`: { fechaActual, inspector, negocio, operacion }.
 */
function registrarAmonestacionesPendientes(amonestaciones, contexto) {
  if (!amonestaciones || amonestaciones.length === 0) return 0;
  var CFG = _configMemo();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = _obtenerHojaMemorandos(ss);
  var fechaActual = (contexto && contexto.fechaActual) || new Date();
  var inspector = (contexto && contexto.inspector) || "";

  var registradas = [];
  amonestaciones.forEach(function(amon) {
    try {
      var fila = construirFilaMemoPendiente(amon, { fechaActual: fechaActual, inspector: inspector, idMemo: Utilities.getUuid() });
      sheet.appendRow(fila);
      try { sheet.getRange(sheet.getLastRow(), 9, 1, 1).setNumberFormat('@'); } catch (eFmt) {}
      registradas.push(amon);
    } catch (eReg) {
      Logger.log("Memo: no se pudo registrar la amonestación pendiente de " + (amon && amon.nombre) + ": " + eReg.toString());
    }
  });
  if (registradas.length === 0) return 0;

  if (String(inspector).indexOf("@") === -1) {
    Logger.log("Memo: " + registradas.length + " amonestación(es) registradas como pendientes de revisión, pero no hay correo de inspector para avisar.");
    return registradas.length;
  }

  var urlApp = "";
  try { urlApp = ScriptApp.getService().getUrl() || ""; } catch (eUrl) { urlApp = ""; }
  var urlRevision = urlApp ? urlApp + (urlApp.indexOf("?") === -1 ? "?" : "&") + "vista=revision" : "";

  var correo = construirCorreoRevisionMemo(registradas, {
    urlRevision: urlRevision,
    modoPrueba: !!CFG.MEMO_MODO_PRUEBA,
    negocio: contexto && contexto.negocio,
    operacion: contexto && contexto.operacion
  });
  MailApp.sendEmail({ to: inspector, subject: correo.asunto, htmlBody: correo.htmlBody });
  return registradas.length;
}

/**
 * PRUEBA MANUAL (ejecutar desde el editor de Apps Script: seleccionar esta función y "Ejecutar").
 * Envía a los correos con Rol CONFIG.MEMO_ROL_DESTINATARIO ("Prueba_Memo") de Config_Personal un
 * ejemplo del correo de confirmación que recibe el inspector, con 2 responsables de ejemplo en 3era
 * alerta (construirAmonestacionesEjemploMemo). No escribe en la hoja Memorandos ni genera Doc/PDF:
 * el botón "Revisar y enviar memos" abre la página real, que mostrará solo los pendientes reales.
 */
function probarCorreoConfirmacionMemos() {
  var CFG = _configMemo();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var destinatarios = _prepararContextoMemo(ss, {}).destinatariosPrueba;
  if (!destinatarios || destinatarios.length === 0) {
    throw new Error("No hay correos con Rol '" + (CFG.MEMO_ROL_DESTINATARIO || "Prueba_Memo") + "' en Config_Personal.");
  }

  var urlApp = "";
  try { urlApp = ScriptApp.getService().getUrl() || ""; } catch (eUrl) { urlApp = ""; }
  var urlRevision = urlApp ? urlApp + (urlApp.indexOf("?") === -1 ? "?" : "&") + "vista=revision" : "";

  var ejemplos = construirAmonestacionesEjemploMemo(new Date());
  var correo = construirCorreoRevisionMemo(ejemplos, {
    urlRevision: urlRevision,
    modoPrueba: true,
    negocio: ejemplos[0].negocio,
    operacion: ejemplos[0].operacion
  });
  var aviso = '<p style="font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:bold;color:' + (CFG.COLOR_RED_DEEP || "#68002B") +
    ';background-color:' + (CFG.COLOR_RED_PALE || "#FFC1C2") + ';padding:10px 16px;margin:0;">Correo de EJEMPLO con datos ficticios: así lo recibirá el inspector. No se registró ni se enviará ningún memorando.</p>';

  MailApp.sendEmail({ to: destinatarios.join(","), subject: correo.asunto, htmlBody: aviso + correo.htmlBody });
  Logger.log("Memo: correo de ejemplo de confirmación enviado a " + destinatarios.join(", "));
  return destinatarios;
}

function _cuentaActivaMemo() {
  var email = "";
  try { email = Session.getActiveUser().getEmail(); } catch (e) { email = ""; }
  return String(email || "").trim().toLowerCase();
}

/**
 * Para la página Revision.html: Memos "Pendiente de revisión".
 * Si la cuenta activa es Supervisor/Revisor, obtiene los de todos los inspectores.
 */
function obtenerPendientesRevision() {
  try {
    var cuenta = _cuentaActivaMemo();
    if (!cuenta) {
      return { ok: false, error: "No se pudo identificar tu cuenta de Google. Abre la aplicación iniciando sesión con tu cuenta corporativa." };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ctxBase = _prepararContextoMemo(ss, {});
    var CFG = _configMemo();
    
    // Verificar permiso de supervisión desde Config_Personal
    var shPersonal = ss.getSheetByName(CFG.HOJA_PERSONAL) || ss.getSheetByName(CFG.HOJA_PERSONAL_ALT);
    var valoresPersonal = shPersonal ? shPersonal.getDataRange().getDisplayValues() : [];
    var esSuper = esSupervisorMemo(cuenta, ctxBase.cecosData, valoresPersonal);

    var sheet = _obtenerHojaMemorandos(ss);
    var valores = sheet.getLastRow() < 2 ? [] : sheet.getDataRange().getValues();
    
    return { 
      ok: true, 
      cuenta: cuenta, 
      esSupervisor: esSuper, 
      pendientes: filtrarPendientesRevision(valores, cuenta, esSuper) 
    };
  } catch (e) {
    Logger.log("Memo: error en obtenerPendientesRevision: " + e.toString());
    return { ok: false, error: "No se pudieron cargar los memorandos pendientes. Intenta nuevamente." };
  }
}

function _copiarContextoMemo(ctx) {
  var out = {};
  Object.keys(ctx).forEach(function(k) { out[k] = ctx[k]; });
  return out;
}

/**
 * Escribe valores por nombre de columna en una fila existente (número de fila 1-based).
 */
function _escribirCeldasMemoPorNombre(sheet, cabecera, numeroFila, valores) {
  Object.keys(valores).forEach(function(nombre) {
    var idx = cabecera.indexOf(nombre);
    if (idx > -1) sheet.getRange(numeroFila, idx + 1).setValue(valores[nombre]);
  });
}

function _procesarDecisionRevision(decision, cuenta, ss, sheet, obtenerCtxBase) {
  var CFG = _configMemo();
  var idMemo = String(decision && decision.idMemo != null ? decision.idMemo : "").trim();
  var res = { idMemo: idMemo, ok: false, estado: "", mensaje: "" };
  try {
    if (!cuenta) {
      res.mensaje = "No se pudo identificar tu cuenta de Google.";
      return res;
    }
    var valores = sheet.getLastRow() < 2 ? [] : sheet.getDataRange().getValues();
    var encontrada = buscarFilaMemoPorId(valores, idMemo);
    var validacion = validarDecisionRevision(decision, encontrada ? encontrada.obj : null, cuenta);
    if (!validacion.ok) {
      res.estado = validacion.codigo;
      res.mensaje = validacion.mensaje;
      return res;
    }

    var cabecera = valores[0].map(function(h) { return String(h == null ? "" : h).trim(); });
    var numeroFila = encontrada.indice + 1;
    var ahora = new Date();
    var tipo = String(decision.decision).trim().toLowerCase();

    if (tipo === "descartar") {
      _escribirCeldasMemoPorNombre(sheet, cabecera, numeroFila, {
        'Estado': CFG.MEMO_ESTADO_DESCARTADO || "Descartado por revisión",
        'Revisor': cuenta,
        'Fecha revisión': ahora,
        'Motivo descarte': String(decision.motivo).trim()
      });
      res.ok = true;
      res.estado = CFG.MEMO_ESTADO_DESCARTADO || "Descartado por revisión";
      res.mensaje = "Memo descartado.";
      return res;
    }

    // tipo === "generar": se reconstruye la amonestación y se usa el flujo de generación existente,
    // actualizando ESTA misma fila (conserva el ID Memo).
    var amon = deserializarDatosAmonestacion(encontrada.obj['Datos amonestación']);
    if (!amon) {
      res.mensaje = "Los datos de la amonestación están incompletos o dañados; no se puede generar el memo.";
      return res;
    }

    // INTEGRACIÓN OPCIÓN A: Re-calcula dinámicamente el historial antes de generar el memo oficial.
    try {
      var layoutHistorial = (typeof resolverLayoutHistorial === "function") 
        ? resolverLayoutHistorial(amon.negocio, ss) 
        : null;
        
      if (layoutHistorial && layoutHistorial.sheet) {
        var registrosHistorial = layoutHistorial.sheet.getDataRange().getValues();
        amon.fallasPrevias = obtenerFallasCicloActual(
          registrosHistorial, 
          amon.negocio, 
          amon.operacion, 
          amon.idSubKey, 
          amon.nombre, 
          layoutHistorial.cols
        );
      }
    } catch (eRecalcGen) {
      Logger.log("Memo: Error recalculando fallas para envío: " + eRecalcGen.toString());
    }

    var ctx = _copiarContextoMemo(obtenerCtxBase());
    ctx.fechaActual = ahora;
    ctx.inspector = _valorCeldaTextoMemo(encontrada.obj['Inspector']);
    ctx.idMemo = idMemo;

    var filaNueva;
    var huboError = false;
    try {
      filaNueva = _generarMemoFila(amon, ctx);
    } catch (eGen) {
      Logger.log("Memo: error generando/enviando memorando (revisión) para " + amon.nombre + ": " + eGen.toString());
      filaNueva = construirFilaMemoError(amon, ctx, eGen.toString());
      huboError = true;
    }

    var filaFinal = combinarFilaMemoExistente(cabecera, valores[encontrada.indice], filaNueva, {
      'Revisor': cuenta,
      'Fecha revisión': ahora
    });
    _escribirFilaMemoExistente(sheet, numeroFila, filaFinal);

    var idxEstado = cabecera.indexOf('Estado');
    res.estado = idxEstado > -1 ? String(filaFinal[idxEstado]) : "";
    res.ok = !huboError && res.estado === "Enviado";
    res.mensaje = res.ok
      ? "Memo generado y enviado."
      : "El memo no se envió (" + res.estado + "). Revisa la hoja Memorandos.";
    return res;
  } catch (e) {
    Logger.log("Memo: error procesando decisión de revisión " + idMemo + ": " + e.toString());
    res.ok = false;
    res.mensaje = "Error inesperado: " + e.toString() + ". Verifica el estado del memo en la hoja Memorandos antes de reintentar.";
    return res;
  }
}

/**
 * Para la página Revision.html: aplica las decisiones del inspector.
 * decisiones: [{ idMemo, decision: "generar" | "descartar", motivo }].
 * Devuelve [{ idMemo, ok, estado, mensaje }] en el mismo orden. Cada decisión va en su propio
 * try/catch; se valida que la fila siga "Pendiente de revisión" y que el Inspector sea la cuenta activa.
 */
function decidirPendientes(decisiones) {
  if (!Array.isArray(decisiones) || decisiones.length === 0) return [];
  var cuenta = _cuentaActivaMemo();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (eLock) {
    return decisiones.map(function(d) {
      return { idMemo: String(d && d.idMemo != null ? d.idMemo : ""), ok: false, estado: "", mensaje: "El sistema está ocupado. Intenta nuevamente en unos segundos." };
    });
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = _obtenerHojaMemorandos(ss);
    var ctxBase = null;
    var obtenerCtxBase = function() {
      if (!ctxBase) ctxBase = _prepararContextoMemo(ss, {});
      return ctxBase;
    };
    return decisiones.map(function(d) {
      return _procesarDecisionRevision(d, cuenta, ss, sheet, obtenerCtxBase);
    });
  } finally {
    lock.releaseLock();
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    formatearNombreNatural: formatearNombreNatural,
    normalizarDni: normalizarDni,
    leerCecosComoObjetos: leerCecosComoObjetos,
    buscarColaborador: buscarColaborador,
    buscarJefe: buscarJefe,
    obtenerDestinatariosMemo: obtenerDestinatariosMemo,
    extraerAmonestaciones: extraerAmonestaciones,
    extraerResponsablesConAlerta: extraerResponsablesConAlerta,
    memoOmitidoPorInspector: memoOmitidoPorInspector,
    memoActivadoPorInspector: memoActivadoPorInspector,
    memoSeleccionadoParaEnvio: memoSeleccionadoParaEnvio,
    normalizarNivelAlertaMemo: normalizarNivelAlertaMemo,
    textoNivelAlertaMemo: textoNivelAlertaMemo,
    obtenerFallasCicloActual: obtenerFallasCicloActual,
    formatearFechaMemo: formatearFechaMemo,
    distribuirFechasIncumplimiento: distribuirFechasIncumplimiento,
    construirFraseFechasIncumplimiento: construirFraseFechasIncumplimiento,
    fraseFechasIncumplimientoPlantillaMemo: fraseFechasIncumplimientoPlantillaMemo,
    elegirPlantillaMemo: elegirPlantillaMemo,
    construirAmonestacionesEjemploMemo: construirAmonestacionesEjemploMemo,
    probarCorreoConfirmacionMemos: probarCorreoConfirmacionMemos,
    construirReemplazosMemo: construirReemplazosMemo,
    resolverEmisoras: resolverEmisoras,
    construirDestinatariosMemo: construirDestinatariosMemo,
    formatearDestinatariosMemo: formatearDestinatariosMemo,
    construirUrlRespuestaMemo: construirUrlRespuestaMemo,
    construirCabeceraMemorandos: construirCabeceraMemorandos,
    generarYEnviarMemorandos: typeof generarYEnviarMemorandos !== 'undefined' ? generarYEnviarMemorandos : undefined,
    // Revisión humana previa al memo
    registrarAmonestacionesPendientes: registrarAmonestacionesPendientes,
    obtenerPendientesRevision: obtenerPendientesRevision,
    decidirPendientes: decidirPendientes,
    fechaATextoMemo: fechaATextoMemo,
    serializarDatosAmonestacion: serializarDatosAmonestacion,
    deserializarDatosAmonestacion: deserializarDatosAmonestacion,
    construirFilaMemoPendiente: construirFilaMemoPendiente,
    construirFilaMemoPorNombre: construirFilaMemoPorNombre,
    construirFilaMemoNoIdentificado: construirFilaMemoNoIdentificado,
    construirFilaMemoSinDestinatarios: construirFilaMemoSinDestinatarios,
    construirFilaMemoError: construirFilaMemoError,
    construirFilaMemoEnviado: construirFilaMemoEnviado,
    combinarFilaMemoExistente: combinarFilaMemoExistente,
    objetoDesdeFilaMemo: objetoDesdeFilaMemo,
    construirVistaPendienteMemo: construirVistaPendienteMemo,
    filtrarPendientesRevision: filtrarPendientesRevision,
    buscarFilaMemoPorId: buscarFilaMemoPorId,
    validarDecisionRevision: validarDecisionRevision,
    construirCorreoRevisionMemo: construirCorreoRevisionMemo,
    _obtenerHojaMemorandos: _obtenerHojaMemorandos,
    _generarMemoFila: _generarMemoFila,
    _escaparRegexMarcadorMemo: _escaparRegexMarcadorMemo
  };
}