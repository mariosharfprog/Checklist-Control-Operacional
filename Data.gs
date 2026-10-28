/**
 * SISTEMA DE INSPECCIÓN DE PROCESOS - SHARF
 * Archivo: Data.gs — Carga de datos desde Sheets y utilidades de consulta
 */

/**
 * Obtiene el correo del usuario que está utilizando la aplicación web
 */
function obtenerUsuarioActual() {
  try {
    var emailActive = Session.getActiveUser().getEmail();
    var emailEffective = Session.getEffectiveUser().getEmail();
    var emailFinal = emailActive || emailEffective || "";
    
    if (emailFinal && emailFinal.indexOf("@") !== -1) {
      return {
        exito: true,
        email: emailFinal
      };
    } else {
      return {
        exito: false,
        email: "Sesión activa"
      };
    }
  } catch (e) {
    Logger.log("Error en obtenerUsuarioActual: " + e.toString());
    return {
      exito: false,
      email: "Sesión activa"
    };
  }
}

function capitalizarNegocio(s) {
  const m = { almacen:"Almacén", distribucion:"Distribución", fedex:"Fedex" };
  const k = String(s||"").toLowerCase().replace(/[^a-z]/g,"");
  return m[k] || String(s).charAt(0).toUpperCase()+String(s).slice(1).toLowerCase();
}
function normalizarTextoBusqueda(s){ return String(s||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/,/g," ").replace(/\s+/g," ").trim(); }
function esIdFechaLike(val){
  if (val == null || val === "") return false;
  if (Object.prototype.toString.call(val) === '[object Date]') return true;
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return true;
  if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) return true;
  if (/^\d{4}\/\d{1,2}\/\d{1,2}/.test(s)) return true;
  return false;
}
function normalizarIdSubactividad(val, fallback){
  var s = String(val||"").trim();
  if (s === "") return String(fallback||"");
  // Quitar punto final "1.1." -> "1.1"
  if (/^\d+(\.\d+)*\.$/.test(s)) s = s.slice(0,-1);
  return s;
}
function detectarHojaCecos(ss){
  const cand = ["cecos","CECOS","CeCosOfi","CECOSOFI","Cecos","cecosofí"];
  for(const n of cand){ const sh=ss.getSheetByName(n); if(sh) return sh; }
  const sheets=ss.getSheets();
  for(const sh of sheets){ if(String(sh.getName()).toLowerCase().includes("ceco")) return sh; }
  return null;
}
function detectarHojaCorreosExtra(ss){
  const cand = ["correosExtra","correosextra","CorreosExtra","CORREOSEXTRA","correos_extra","correos extra"];
  for(const n of cand){ const sh=ss.getSheetByName(n); if(sh) return sh; }
  const sheets=ss.getSheets();
  for(const sh of sheets){
    const nm = String(sh.getName()).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[_ ]/g,"");
    if(nm.includes("correos") && nm.includes("extra")) return sh;
    if(nm === "correosextra") return sh;
  }
  return null;
}
function cargarCorreosExtraMatriz(ss){
  let matriz=[];
  const sh = detectarHojaCorreosExtra(ss);
  if(!sh) return matriz;
  const vals = sh.getDataRange().getDisplayValues();
  if(!vals || vals.length < 2) return matriz;
  const header = vals[0].map(function(h){ return String(h||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,""); });
  let idxNombre = header.findIndex(function(h){ return h.includes("nombre"); });
  let idxEmail = header.findIndex(function(h){ return h.includes("email")||h.includes("correo"); });
  if(idxNombre < 0) idxNombre = 0;
  if(idxEmail < 0) idxEmail = 1;
  // Si solo hay 1 columna y parece email, asumir nombre = email local part
  for(let j=1;j<vals.length;j++){
    let nombre = String(vals[j][idxNombre]||"").trim();
    let email = String(vals[j][idxEmail]||"").trim();
    // Fallback: si idxEmail fuera de rango o vacío pero nombre parece email
    if(!email && nombre && nombre.indexOf("@") !== -1){ email = nombre; }
    if(!nombre && email) nombre = email.split("@")[0].replace(/[._-]+/g," ");
    if(!nombre && !email) continue;
    if(email && email.indexOf("@") === -1) continue; // solo filas con email válido o sin email pero con nombre
    const searchIndex = normalizarTextoBusqueda(nombre+" "+email);
    matriz.push({id:"extra_"+j, nombre:nombre, cargo:"", area:"", email:email.toLowerCase(), searchIndex:searchIndex, palabras: searchIndex.split(/\s+/).filter(Boolean), origen:"correosExtra"});
  }
  return matriz;
}

function normalizarClave5S(s) {
  return String(s || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[_\s\-]/g, "")
    .trim();
}

function normalizarParaComparar(s) {
  return String(s || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function esHoja5S(nombreHoja) {
  var base = (typeof CONFIG !== 'undefined' && CONFIG.HOJA_CL_5S) ? CONFIG.HOJA_CL_5S : "cl_5s";
  var normBase = normalizarClave5S(base).replace(/^CL/, "");
  var normNombre = normalizarClave5S(nombreHoja);
  var normSinCl = normNombre.replace(/^CL/, "");
  return normSinCl === "5S" || (normBase !== "" && normSinCl === normBase) || normNombre === normalizarClave5S(base);
}

function esHojaDetalle5S(nombreHoja) {
  return /^(?:cl[_\-]?)?5s[_\-]+(.+)$/i.test(String(nombreHoja || "").trim());
}

function esActividad5S(nombreAct) {
  return normalizarClave5S(nombreAct) === "5S";
}

function cargarPlantillas5S(ss) {
  var plantillas = [];
  if (!ss) return plantillas;
  var sh = null;
  var nombreConf = (typeof CONFIG !== 'undefined' && CONFIG.HOJA_CL_5S) ? CONFIG.HOJA_CL_5S : "cl_5s";
  try {
    sh = ss.getSheetByName(nombreConf);
  } catch (e) {}
  if (!sh) {
    try {
      var sheets = ss.getSheets();
      for (var s = 0; s < sheets.length; s++) {
        if (esHoja5S(sheets[s].getName())) {
          sh = sheets[s];
          break;
        }
      }
    } catch (e2) {}
  }
  if (!sh) return plantillas;

  var vals = null;
  try {
    vals = sh.getDataRange().getDisplayValues();
  } catch (e3) {
    return plantillas;
  }
  if (!vals || vals.length < 2) return plantillas;

  for (var i = 1; i < vals.length; i++) {
    var fila = vals[i];
    if (!fila) continue;
    var esVacia = true;
    for (var c = 0; c < fila.length; c++) {
      if (String(fila[c] || "").trim() !== "") {
        esVacia = false;
        break;
      }
    }
    if (esVacia) continue;

    var opRaw = fila[0] ? String(fila[0]).trim() : "";
    var opCanon = (opRaw !== "" && typeof canonicalizarOperacion === 'function') ? canonicalizarOperacion(opRaw) : opRaw;
    var actRaw = fila[2] ? String(fila[2]).trim() : "5S";
    var subRaw = fila[4] ? String(fila[4]).trim() : "";

    plantillas.push({
      operacion: opCanon,
      actividad: actRaw || "5S",
      subactividad: subRaw,
      sIndex: i
    });
  }
  return plantillas;
}

function normalizarIdPunto5S(idRaw, fallbackSecuencial) {
  if (idRaw != null && String(idRaw).trim() !== "") {
    var s = String(idRaw).trim();
    var n = parseFloat(s);
    if (!isNaN(n) && isFinite(n)) {
      return String(Math.round(n));
    }
  }
  return String(fallbackSecuencial);
}

function cargarDetalle5S(ss) {
  var detalle = {};
  if (!ss) return detalle;
  var sheets = null;
  try {
    sheets = ss.getSheets();
  } catch (e) {
    return detalle;
  }
  if (!sheets || !sheets.length) return detalle;

  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var nombreHoja = String(sh.getName() || "").trim();
    if (!esHojaDetalle5S(nombreHoja)) continue;
    var m = nombreHoja.match(/^(?:cl[_\-]?)?5s[_\-]+(.+)$/i);
    if (!m) continue;
    var slug = m[1];
    var negocio = capitalizarNegocio(slug);

    var vals = null;
    try {
      vals = sh.getDataRange().getDisplayValues();
    } catch (e2) {
      continue;
    }
    if (!vals || vals.length < 2) continue;

    var valsRaw = null;
    try {
      valsRaw = sh.getDataRange().getValues();
    } catch (eRaw) {
      valsRaw = null;
    }

    // Columnas fijas por POSICIÓN: 0 Operación | 1 id | 2 S | 3 Área | 4 Punto | 5 Cumplimiento
    for (var i = 1; i < vals.length; i++) {
      var fila = vals[i];
      if (!fila) continue;
      var filaRaw = (valsRaw && valsRaw[i]) ? valsRaw[i] : fila;

      var esVacia = true;
      for (var c = 0; c < fila.length; c++) {
        if (String(fila[c] || "").trim() !== "") {
          esVacia = false;
          break;
        }
      }
      if (esVacia) continue;

      var opRaw = fila[0] != null ? String(fila[0]).trim() : "";
      var puntoRaw = fila[4] != null ? String(fila[4]).trim() : "";
      if (opRaw === "" || puntoRaw === "") continue;

      var sRaw = fila[2] != null ? String(fila[2]).trim() : "";
      var sNum = parseInt(sRaw, 10);
      if (isNaN(sNum) || !isFinite(sNum) || sNum <= 0) continue;

      var opCanon = (typeof canonicalizarOperacion === 'function') ? canonicalizarOperacion(opRaw) : opRaw;
      var clave = normalizarParaComparar(negocio) + "|" + normalizarParaComparar(opCanon) + "|" + Number(sNum);

      if (!detalle[clave]) {
        detalle[clave] = [];
      }

      var areaTrim = fila[3] != null ? String(fila[3]).trim() : "";
      var puntoTrim = puntoRaw.trim();
      var textoPunto = areaTrim !== "" ? (areaTrim + " — " + puntoTrim) : puntoTrim;

      var seqIndex = detalle[clave].length + 1;
      var idRaw = (filaRaw[1] != null && String(filaRaw[1]).trim() !== "") ? filaRaw[1] : fila[1];
      var idPunto = normalizarIdPunto5S(idRaw, seqIndex);

      detalle[clave].push({
        idPunto: idPunto,
        area: areaTrim,
        punto: puntoTrim,
        textoPunto: textoPunto
      });
    }
  }
  return detalle;
}

function anexarMacro5S(checklist, plantillas5S, contadorPorMacro, detalle5S) {
  if (!checklist || !plantillas5S || plantillas5S.length === 0) {
    return checklist || [];
  }
  if (!detalle5S) detalle5S = {};

  var paresMap = {};
  var paresOrdenados = [];

  for (var i = 0; i < checklist.length; i++) {
    var item = checklist[i];
    if (!item) continue;
    var nNorm = normalizarParaComparar(item.negocio);
    var oNorm = normalizarParaComparar(item.operacion);
    var clave = nNorm + "|" + oNorm;

    if (!paresMap[clave]) {
      paresMap[clave] = {
        negocio: item.negocio,
        operacion: item.operacion,
        lastIndex: i,
        maxMacro: 0,
        yaTiene5S: false
      };
      paresOrdenados.push(clave);
    } else {
      paresMap[clave].lastIndex = i;
    }

    if (esActividad5S(item.actividad)) {
      paresMap[clave].yaTiene5S = true;
    }

    var primerSegmento = String(item.idActividad || "").split(".")[0].trim();
    var numMacro = parseInt(primerSegmento, 10);
    if (!isNaN(numMacro) && numMacro > paresMap[clave].maxMacro) {
      paresMap[clave].maxMacro = numMacro;
    }
  }

  var inserciones = [];

  for (var p = 0; p < paresOrdenados.length; p++) {
    var clavePar = paresOrdenados[p];
    var info = paresMap[clavePar];

    if (info.yaTiene5S) continue;

    var plantillasPar = [];
    for (var t = 0; t < plantillas5S.length; t++) {
      var pl = plantillas5S[t];
      var plOp = String(pl.operacion || "").trim();
      if (plOp === "" || normalizarParaComparar(plOp) === normalizarParaComparar(info.operacion)) {
        plantillasPar.push(pl);
      }
    }
    if (plantillasPar.length === 0) continue;

    var nMacro = (info.maxMacro > 0) ? (info.maxMacro + 1) : 1;
    var idActFinal = String(nMacro);

    var itemsGenerados = [];
    for (var k = 0; k < plantillasPar.length; k++) {
      var pl = plantillasPar[k];
      var sIdx = (pl.sIndex != null) ? pl.sIndex : (k + 1);
      var plSub = String(pl.subactividad || "").trim();
      var colonIdx = plSub.indexOf(":");
      var nombreGrupo = (colonIdx !== -1) ? plSub.slice(0, colonIdx).trim() : plSub;
      if (nombreGrupo === "") nombreGrupo = plSub;

      var idGrupo = idActFinal + "." + sIdx;

      // S5 (sIndex === 5) NUNCA lee detalle: siempre un único item placeholder DERIVADA
      if (sIdx === 5) {
        var subactividadCorta5 = (colonIdx !== -1) ? plSub.slice(colonIdx + 1).trim() : plSub;
        if (subactividadCorta5 === "") subactividadCorta5 = plSub;

        itemsGenerados.push({
          negocio: info.negocio,
          operacion: info.operacion,
          idActividad: idActFinal,
          actividad: String(pl.actividad || "5S").trim(),
          idGrupo: idGrupo,
          grupo: nombreGrupo,
          sIndex: 5,
          tipoS: "DERIVADA",
          idPunto: "",
          area: "",
          idSubactividad: idGrupo,
          subactividad: plSub,
          subactividadCorta: subactividadCorta5,
          nombreEvidencia: idGrupo,
          textoS: plSub
        });
        continue;
      }

      // S1..S4: buscar detalle
      var claveDetalle = normalizarParaComparar(info.negocio) + "|" + normalizarParaComparar(info.operacion) + "|" + Number(sIdx);
      var puntos = detalle5S[claveDetalle];

      if (puntos && puntos.length > 0) {
        // CASO A: con detalle
        var tipoS = "MANUAL";
        for (var j = 0; j < puntos.length; j++) {
          var puntoObj = puntos[j];
          var idPunto = "";
          var areaPunto = "";
          var textoPunto = "";

          if (typeof puntoObj === 'object' && puntoObj !== null) {
            idPunto = String(puntoObj.idPunto != null ? puntoObj.idPunto : (j + 1)).trim();
            areaPunto = String(puntoObj.area || "").trim();
            textoPunto = puntoObj.textoPunto ? String(puntoObj.textoPunto).trim() : (areaPunto ? (areaPunto + " — " + String(puntoObj.punto || "").trim()) : String(puntoObj.punto || "").trim());
          } else {
            // Retrocompatibilidad si punto viene como string
            idPunto = String(j + 1);
            areaPunto = "";
            textoPunto = String(puntoObj || "").trim();
          }

          var idSub = idGrupo + "." + idPunto;
          var subact = nombreGrupo + " — " + textoPunto;
          var subCorta = textoPunto;
          var nomEv = idPunto ? (idGrupo + "_" + idPunto) : idGrupo;

          itemsGenerados.push({
            negocio: info.negocio,
            operacion: info.operacion,
            idActividad: idActFinal,
            actividad: String(pl.actividad || "5S").trim(),
            idGrupo: idGrupo,
            grupo: nombreGrupo,
            sIndex: sIdx,
            tipoS: tipoS,
            idPunto: idPunto,
            area: areaPunto,
            idSubactividad: idSub,
            subactividad: subact,
            subactividadCorta: subCorta,
            nombreEvidencia: nomEv,
            textoS: plSub
          });
        }
      } else {
        // CASO B: sin detalle (fallback genérico)
        var subactividadCortaGen = (colonIdx !== -1) ? plSub.slice(colonIdx + 1).trim() : plSub;
        if (subactividadCortaGen === "") subactividadCortaGen = plSub;

        itemsGenerados.push({
          negocio: info.negocio,
          operacion: info.operacion,
          idActividad: idActFinal,
          actividad: String(pl.actividad || "5S").trim(),
          idGrupo: idGrupo,
          grupo: nombreGrupo,
          sIndex: sIdx,
          tipoS: "MANUAL",
          idPunto: "",
          area: "",
          idSubactividad: idGrupo,
          subactividad: plSub,
          subactividadCorta: subactividadCortaGen,
          nombreEvidencia: idGrupo,
          textoS: plSub,
          sinDesglose: true
        });
      }
    }

    if (contadorPorMacro) {
      var claveContador = String(info.negocio || "") + "|" + String(info.operacion || "") + "|" + idActFinal;
      contadorPorMacro[claveContador] = plantillasPar.length;
    }

    inserciones.push({
      insertIndex: info.lastIndex + 1,
      items: itemsGenerados
    });
  }

  inserciones.sort(function(a, b) {
    return b.insertIndex - a.insertIndex;
  });

  for (var j = 0; j < inserciones.length; j++) {
    var ins = inserciones[j];
    Array.prototype.splice.apply(checklist, [ins.insertIndex, 0].concat(ins.items));
  }

  return checklist;
}

/**
 * CÁLCULO DE PORCENTAJES Y ESTADOS 5S
 * items5S: los items de UNA (negocio, operación) cuya actividad es 5S
 * estados: objeto { idSubactividad: "CUMPLE"|"NO CUMPLE"|"N/A"|"" }
 */
function calcularPorcentajesS(items5S, estados) {
  var resumen = {};
  if (!items5S || !items5S.length) {
    return resumen;
  }
  if (!estados) estados = {};

  var umbral = (typeof CONFIG !== 'undefined' && CONFIG.UMBRAL_5S != null) ? CONFIG.UMBRAL_5S : 80;

  var gruposMap = {};
  var listaGrupos = [];

  for (var i = 0; i < items5S.length; i++) {
    var it = items5S[i];
    if (!it) continue;
    var idG = it.idGrupo || String((it.idActividad || "") + "." + (it.sIndex || ""));
    if (!gruposMap[idG]) {
      gruposMap[idG] = {
        idGrupo: idG,
        grupo: it.grupo || "",
        sIndex: Number(it.sIndex) || 0,
        tipoS: it.tipoS || "MANUAL",
        items: []
      };
      listaGrupos.push(gruposMap[idG]);
    }
    gruposMap[idG].items.push(it);
  }

  listaGrupos.sort(function(a, b) {
    return a.sIndex - b.sIndex;
  });

  // Paso 1: Calcular S1..S4 (MANUAL)
  for (var g = 0; g < listaGrupos.length; g++) {
    var grp = listaGrupos[g];
    if (grp.tipoS === "DERIVADA" || grp.sIndex === 5) {
      continue;
    }

    var pct = null;
    var cumple = 0;
    var noCumple = 0;
    for (var k = 0; k < grp.items.length; k++) {
      var idSub = grp.items[k].idSubactividad;
      var est = String(estados[idSub] || "").trim().toUpperCase();
      if (est === "CUMPLE") {
        cumple++;
      } else if (est === "NO CUMPLE") {
        noCumple++;
      }
    }
    var total = cumple + noCumple;
    if (total > 0) {
      pct = (cumple / total) * 100;
    }

    var estado = "N/A";
    if (pct !== null) {
      estado = (pct >= umbral) ? "CUMPLE" : "NO CUMPLE";
    }

    grp.pct = pct;
    grp.estado = estado;
  }

  // Paso 2: Calcular S5 (DERIVADA) como promedio de S1..S4 no nulos
  for (var d = 0; d < listaGrupos.length; d++) {
    var grpD = listaGrupos[d];
    if (grpD.tipoS === "DERIVADA" || grpD.sIndex === 5) {
      var sumaS = 0;
      var cantS = 0;
      for (var s = 0; s < listaGrupos.length; s++) {
        var otro = listaGrupos[s];
        if (otro.sIndex !== 5 && otro.tipoS !== "DERIVADA") {
          if (otro.pct !== null && !isNaN(otro.pct) && isFinite(otro.pct)) {
            sumaS += otro.pct;
            cantS++;
          }
        }
      }
      var pctDerivada = (cantS > 0) ? (sumaS / cantS) : null;
      var estadoDerivada = "N/A";
      if (pctDerivada !== null) {
        estadoDerivada = (pctDerivada >= umbral) ? "CUMPLE" : "NO CUMPLE";
      }
      grpD.pct = pctDerivada;
      grpD.estado = estadoDerivada;
    }
  }

  // Devolver en orden de sIndex
  for (var r = 0; r < listaGrupos.length; r++) {
    var itemGrp = listaGrupos[r];
    resumen[itemGrp.idGrupo] = {
      idGrupo: itemGrp.idGrupo,
      grupo: itemGrp.grupo,
      sIndex: itemGrp.sIndex,
      tipoS: itemGrp.tipoS,
      pct: itemGrp.pct,
      estado: itemGrp.estado
    };
  }

  return resumen;
}

/**
 * Calcula el % Cumplimiento Actividad de 5S (promedio de los % de las 5 S que no sean null)
 * resumenPorS: objeto retornado por calcularPorcentajesS
 */
function calcularPctActividad5S(resumenPorS) {
  if (!resumenPorS) return null;
  var keys = Object.keys(resumenPorS);
  if (keys.length === 0) return null;

  var suma = 0;
  var cant = 0;
  for (var i = 0; i < keys.length; i++) {
    var it = resumenPorS[keys[i]];
    if (!it) continue;
    var p = (typeof it === 'object' && it !== null && 'pct' in it) ? it.pct : it;
    if (p != null && !isNaN(p) && isFinite(p)) {
      suma += Number(p);
      cant++;
    }
  }

  if (cant === 0) return null;
  return suma / cant;
}

/**
 * CARGAR DATOS DESDE LAS HOJAS DE CONFIGURACIÓN
 * Dinámico cl_* + cecos con fallback legacy
 */
function cargarDatosDesdeSheet(){
  try{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var checklist=[];
    var sheets=ss.getSheets();
    var prefijo = (typeof CONFIG !== 'undefined' && CONFIG.PREFIJO_CL) ? CONFIG.PREFIJO_CL : "cl_";
    var hojasCl=sheets.filter(function(s){
      var nombre = String(s.getName());
      return nombre.toLowerCase().startsWith(prefijo.toLowerCase()) && !esHoja5S(nombre) && !esHojaDetalle5S(nombre);
    });
    var plantillas5S = cargarPlantillas5S(ss);
    var detalle5S = cargarDetalle5S(ss);
    // Contador secuencial por macro para corregir IDs con fecha (ej. 2026-01-01) -> 1.1, 1.2 ...
    var contadorPorMacro = {};
    function siguienteIdSecuencial(negocio, operacion, idAct, filaIdx){
      var clave = String(negocio||"")+"|"+String(operacion||"")+"|"+String(idAct||"");
      var actual = contadorPorMacro[clave] || 0;
      actual += 1;
      contadorPorMacro[clave] = actual;
      var base = String(idAct||"").trim();
      if (base === "" || base === "0") base = "1";
      // Normalizar base sin punto final
      if (/^\d+(\.\d+)*\.$/.test(base)) base = base.slice(0,-1);
      return base + "." + actual;
    }
    if(hojasCl.length>0){
      hojasCl.forEach(function(sh){
        const negocio=capitalizarNegocio(String(sh.getName()).replace(/^cl_/i,""));
        const vals=sh.getDataRange().getDisplayValues();
        // También leer valores crudos para detectar Date objects (getValues vs getDisplayValues)
        var valsRaw = null;
        try { valsRaw = sh.getDataRange().getValues(); } catch(e){ valsRaw = vals; }
        for(let i=1;i<vals.length;i++){
          const fila=vals[i];
          const filaRaw=valsRaw ? valsRaw[i] : fila;
          if(!fila[0]||String(fila[0]).trim()==="") continue;
          var opRawCl = String(fila[0]).trim();
          var opCanonCl = (typeof canonicalizarOperacion === 'function') ? canonicalizarOperacion(opRawCl) : opRawCl;
          var idActRaw = fila[1]?String(fila[1]).trim():"";
          var actRaw = fila[2]?String(fila[2]).trim():"GENERAL";
          var idSubRawDisp = fila[3];
          var idSubRawVal = filaRaw ? filaRaw[3] : idSubRawDisp;
          var subRaw = fila[4]?String(fila[4]).trim():"";
          // Normalizar ID: si es fecha o vacío o no parece ID tipo "1.1", generar secuencial
          var idSubNorm = normalizarIdSubactividad(idSubRawDisp, "");
          var esFecha = esIdFechaLike(idSubRawVal) || esIdFechaLike(idSubRawDisp);
          var pareceIdValido = /^\d+(\.\d+)+$/.test(idSubNorm) || /^\d+$/.test(idSubNorm);
          var idSubFinal;
          if (!esFecha && pareceIdValido && idSubNorm !== "") {
            idSubFinal = idSubNorm;
            // Actualizar contador para no colisionar con futuros autogenerados
            var claveCnt = String(negocio||"")+"|"+String(opCanonCl||"")+"|"+String(idActRaw||"");
            var partes = idSubFinal.split(".");
            var suf = parseInt(partes[partes.length-1],10);
            if (!isNaN(suf)) {
              var cur = contadorPorMacro[claveCnt] || 0;
              if (suf > cur) contadorPorMacro[claveCnt] = suf;
            }
          } else {
            // Generar secuencial dentro de la macro
            idSubFinal = siguienteIdSecuencial(negocio, opCanonCl, idActRaw, i);
          }
          var subFinal = subRaw !== "" ? subRaw : (idSubFinal || "Punto "+i);
          // Evitar que subactividad quede como fecha cuando el ID era fecha
          if (esIdFechaLike(subFinal) && !esIdFechaLike(idSubFinal)) {
            subFinal = "Punto "+idSubFinal;
          }
          checklist.push({
            negocio: negocio,
            operacion: opCanonCl,
            idActividad: idActRaw,
            actividad: actRaw,
            idSubactividad: idSubFinal,
            subactividad: subFinal
          });
        }
      });
      checklist = anexarMacro5S(checklist, plantillas5S, contadorPorMacro, detalle5S);
    } else {
      // fallback legacy Config_Checklist
      var sheetChecklist = ss.getSheetByName(CONFIG.HOJA_CHECKLIST) || ss.getSheetByName("Checklist") || ss.getSheetByName("CHECKLIST") || ss.getSheetByName("checklist");
      if(!sheetChecklist) return {status:"ERROR", message:"No se encontró Config_Checklist ni cl_*"};
      var datosChecklist=sheetChecklist.getDataRange().getDisplayValues();
      var valsRawLeg = null;
      try { valsRawLeg = sheetChecklist.getDataRange().getValues(); } catch(e){ valsRawLeg = datosChecklist; }
      for(let i=1;i<datosChecklist.length;i++){
        const fila=datosChecklist[i];
        const filaRawLeg = valsRawLeg ? valsRawLeg[i] : fila;
        if(fila[0]&&fila[1]&&(fila[5]||fila[4])) {
          var opRawLeg = String(fila[1]).trim();
          var opCanonLeg = (typeof canonicalizarOperacion === 'function') ? canonicalizarOperacion(opRawLeg) : opRawLeg;
          var idActLeg = String(fila[2]||"").trim();
          var actLeg = String(fila[3]||"GENERAL").trim();
          var idSubDispLeg = fila[4];
          var idSubRawLegVal = filaRawLeg ? filaRawLeg[4] : idSubDispLeg;
          var subDispLeg = fila[5]?String(fila[5]).trim():"";
          var idSubNormLeg = normalizarIdSubactividad(idSubDispLeg, "");
          var esFechaLeg = esIdFechaLike(idSubRawLegVal) || esIdFechaLike(idSubDispLeg);
          var pareceIdLeg = /^\d+(\.\d+)+$/.test(idSubNormLeg) || /^\d+$/.test(idSubNormLeg);
          var idSubFinalLeg;
          var negLeg = String(fila[0]).trim();
          if (!esFechaLeg && pareceIdLeg && idSubNormLeg !== "") {
            idSubFinalLeg = idSubNormLeg;
            var claveLeg = negLeg+"|"+opCanonLeg+"|"+idActLeg;
            var partesLeg = idSubFinalLeg.split(".");
            var sufLeg = parseInt(partesLeg[partesLeg.length-1],10);
            if (!isNaN(sufLeg)) {
              var curLeg = contadorPorMacro[claveLeg] || 0;
              if (sufLeg > curLeg) contadorPorMacro[claveLeg] = sufLeg;
            }
          } else {
            idSubFinalLeg = siguienteIdSecuencial(negLeg, opCanonLeg, idActLeg, i);
          }
          var subFinalLeg = subDispLeg !== "" ? subDispLeg : (idSubFinalLeg || "Punto "+i);
          if (esIdFechaLike(subFinalLeg) && !esIdFechaLike(idSubFinalLeg)) subFinalLeg = "Punto "+idSubFinalLeg;
          checklist.push({negocio:negLeg, operacion:opCanonLeg, idActividad:idActLeg, actividad:actLeg, idSubactividad:idSubFinalLeg, subactividad:subFinalLeg});
        }
      }
      checklist = anexarMacro5S(checklist, plantillas5S, contadorPorMacro, detalle5S);
    }
    // cecos
    let cecosMatriz=[];
    const shCecos=detectarHojaCecos(ss);
    if(shCecos){
      const vals=shCecos.getDataRange().getDisplayValues();
      if(vals && vals.length > 0){
        const headerNorm=vals[0].map(function(h){ 
          return String(h||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim(); 
        });

        // Detección precisa de encabezados para evitar colisiones (ej. Nombre vs Código, Área vs Sub Área)
        let idxNombre = headerNorm.findIndex(function(h){ return h === "nombre" || h === "nombre completo"; });
        if (idxNombre < 0) idxNombre = headerNorm.findIndex(function(h){ return h.includes("nombre") && !h.includes("codigo") && !h.includes("ceco"); });

        let idxCargo = headerNorm.findIndex(function(h){ return h === "cargo" || h === "puesto"; });
        if (idxCargo < 0) idxCargo = headerNorm.findIndex(function(h){ return h.includes("cargo"); });

        let idxArea = headerNorm.findIndex(function(h){ return h === "area"; });
        if (idxArea < 0) idxArea = headerNorm.findIndex(function(h){ return h.includes("area") && !h.includes("sub"); });

        let idxEmail = headerNorm.findIndex(function(h){ return h === "email" || h === "correo" || h === "correo electronico"; });
        if (idxEmail < 0) idxEmail = headerNorm.findIndex(function(h){ return h.includes("email") || h.includes("correo"); });

        // Con la copia completa de CeCosOfi (más de 6 columnas) las posiciones fijas apuntarían a
        // Código/DNI/Empresa/División: si falta un encabezado se omite cecos (sin romper la carga del checklist).
        const cabeceraIncompleta = headerNorm.length > 6 && (idxNombre < 0 || idxCargo < 0 || idxArea < 0 || idxEmail < 0);
        if (cabeceraIncompleta) {
          Logger.log("Aviso: la hoja cecos no tiene los encabezados Nombre, Cargo, Área y Email; se omite la lista de personas.");
        }
        if(idxNombre<0) idxNombre=0; if(idxCargo<0) idxCargo=1; if(idxArea<0) idxArea=2; if(idxEmail<0) idxEmail=3;

        for(let j=1;j<vals.length && !cabeceraIncompleta;j++){
          const nombre=String(vals[j][idxNombre]||"").trim();
          const cargo=String(vals[j][idxCargo]||"").trim();
          const area=String(vals[j][idxArea]||"").trim();
          const email=String(vals[j][idxEmail]||"").trim();
          const emailLower = email.toLowerCase();

          // Omitir filas sin email válido (mismo criterio que tenía la sincronización)
          if (!email || 
              emailLower === "no aplica" || 
              emailLower === "nan" || 
              emailLower === "-" || 
              email.indexOf("@") === -1) {
            continue;
          }

          if (!nombre) continue;

          const searchIndex=normalizarTextoBusqueda(nombre+" "+cargo+" "+area+" "+emailLower);
          cecosMatriz.push({
            id: j, 
            nombre: nombre, 
            cargo: cargo, 
            area: area, 
            email: emailLower, 
            searchIndex: searchIndex, 
            palabras: searchIndex.split(/\s+/).filter(Boolean)
          });
        }
      }
    } else {
      // fallback legacy personal
      var sheetPersonal = ss.getSheetByName(CONFIG.HOJA_PERSONAL)||ss.getSheetByName(CONFIG.HOJA_PERSONAL_ALT) || ss.getSheetByName("PERSONAL") || ss.getSheetByName("Personal") || ss.getSheetByName("personal") || ss.getSheetByName("config_personal");
      if(sheetPersonal){
        const vals=sheetPersonal.getDataRange().getDisplayValues();
        for(let j=1;j<vals.length;j++){
          const nombre=String(vals[j][0]||"").trim(); if(!nombre) continue;
          cecosMatriz.push({id:j, nombre:nombre, cargo:String(vals[j][1]||""), area:"", email: nombre.includes("@")?nombre:"", searchIndex: normalizarTextoBusqueda(nombre+" "+vals[j][1]), palabras: normalizarTextoBusqueda(nombre).split(/\s+/).filter(Boolean)});
        }
      }
    }
    let correosExtraMatriz = cargarCorreosExtraMatriz(ss);
    return {status:"SUCCESS", checklist:checklist, cecosMatriz:cecosMatriz, correosExtraMatriz: correosExtraMatriz, personalMatriz: []};
  }catch(e){ Logger.log("Error "+e); return {status:"ERROR", message:e.toString()};}
}

/**
 * Obtiene los correos de prueba dinámicos según el Negocio y Operación seleccionados.
 * Lee directamente la hoja Config_Personal (Correo/Nombre | Rol | Negocio | Operación)
 * filtrando filas cuyo Rol contenga "PRUEBA", con wildcards TODOS/TODAS.
 * Si no hay coincidencias retorna [] (no usa correo default).
 */
function obtenerCorreosPrueba(negocio, operacion) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPersonal = ss.getSheetByName(CONFIG.HOJA_PERSONAL) ||
                        ss.getSheetByName(CONFIG.HOJA_PERSONAL_ALT) ||
                        ss.getSheetByName("PERSONAL") ||
                        ss.getSheetByName("Personal") ||
                        ss.getSheetByName("personal") ||
                        ss.getSheetByName("config_personal");

    var nUpper = (negocio || "").toUpperCase();
    var oUpper = (operacion || "").toUpperCase();

    var correosPrueba = [];

    if (sheetPersonal) {
      var datosPersonal = sheetPersonal.getDataRange().getDisplayValues();
      for (var i = 1; i < datosPersonal.length; i++) {
        var fila = datosPersonal[i];
        var nombrePersona = fila[0] ? String(fila[0]).trim() : "";
        var rolPersona = fila[1] ? String(fila[1]).trim().toUpperCase() : "";
        var negocioPersona = fila[2] ? String(fila[2]).trim().toUpperCase() : "TODOS";
        var operacionPersona = fila[3] ? String(fila[3]).trim().toUpperCase() : "TODAS";

        if (!nombrePersona) continue;
        if (!rolPersona.includes("PRUEBA")) continue;

        var cumpleNegocio = (negocioPersona === "TODOS" || negocioPersona === "TODAS" || negocioPersona === nUpper);
        var cumpleOperacion = (operacionPersona === "TODAS" || operacionPersona === "TODOS" || operacionPersona === oUpper);

        if (cumpleNegocio && cumpleOperacion) {
          var splitted = nombrePersona.split(/[,;]+/);
          splitted.forEach(function(c) {
            var mailLimpio = c.trim();
            if (mailLimpio !== "" && mailLimpio.indexOf("@") !== -1 && !correosPrueba.includes(mailLimpio)) {
              correosPrueba.push(mailLimpio);
            }
          });
        }
      }
    }

    return correosPrueba;
  } catch (e) {
    Logger.log("Error en obtenerCorreosPrueba: " + e.toString());
    return [];
  }
}

/**
 * Convierte un archivo de imagen de Google Drive a Data URI en Base64
 */
function obtenerBase64DesdeDrive(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var bytes = blob.getBytes();
    var base64 = Utilities.base64Encode(bytes);
    return "data:" + blob.getContentType() + ";base64," + base64;
  } catch (e) {
    Logger.log("Error al obtener la imagen en Base64 desde Drive: " + e.toString());
    return "";
  }
}
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    cargarDatosDesdeSheet: cargarDatosDesdeSheet,
    capitalizarNegocio: capitalizarNegocio,
    normalizarTextoBusqueda: normalizarTextoBusqueda,
    detectarHojaCecos: detectarHojaCecos,
    detectarHojaCorreosExtra: detectarHojaCorreosExtra,
    cargarCorreosExtraMatriz: cargarCorreosExtraMatriz,
    obtenerUsuarioActual: obtenerUsuarioActual,
    obtenerCorreosPrueba: obtenerCorreosPrueba,
    obtenerBase64DesdeDrive: obtenerBase64DesdeDrive,
    esHoja5S: esHoja5S,
    esHojaDetalle5S: esHojaDetalle5S,
    cargarPlantillas5S: cargarPlantillas5S,
    cargarDetalle5S: cargarDetalle5S,
    anexarMacro5S: anexarMacro5S,
    calcularPorcentajesS: calcularPorcentajesS,
    calcularPctActividad5S: calcularPctActividad5S
  };
}
