/**
 * Permisos.js - VERSION GLOBAL (autocontenida y neutral)
 *
 * Sirve para CUALQUIER proyecto de Google Apps Script: web app (ejecutada como quien accede o como
 * quien la implementa), script ligado a una hoja, documento, presentacion o formulario, o script
 * independiente. Si a una persona le faltan permisos (por ejemplo, desmarco casillas en la pantalla de
 * autorizacion con el consentimiento granular de Google), se le vuelven a pedir TODOS los permisos que
 * declara el proyecto. Copia este archivo tal cual; no hay nada que configurar.
 *
 * ============================================================================================
 * REGLA DE MANTENIMIENTO (leer antes de editar este archivo)
 * ============================================================================================
 * Apps Script decide que permisos (scopes) pide un proyecto con un analisis ESTATICO del codigo: busca
 * los nombres de los servicios de Google escritos literalmente. Este archivo NO debe agregar permisos
 * al proyecto que lo contiene. Por eso aqui SOLO pueden aparecer, escritos literalmente, los servicios
 * de ScriptApp, HtmlService y Logger. Cualquier otro servicio (hojas, documentos, presentaciones,
 * formularios, calendario, Drive, correo, Gmail, peticiones externas, propiedades, cache, bloqueos,
 * sesion, etc.) NO se escribe con su nombre: se obtiene en tiempo de ejecucion armando el nombre con
 * texto partido y buscandolo en el objeto global (ver permisosServicio_). Tampoco se nombran en
 * comentarios ni en textos: la prueba tests/permisos_global.test.js lo verifica buscando esos nombres.
 * ============================================================================================
 *
 * COMO DAR LOS PERMISOS DESDE EL EDITOR (quien instala o recibe el proyecto):
 *   1. En el editor de Apps Script, elige la funcion `autorizarPermisos` en el desplegable de arriba.
 *   2. Pulsa Ejecutar. Google muestra "Se requiere autorizacion".
 *   3. Pulsa "Revisar permisos" y elige la cuenta.
 *   4. Si aparece "Google no ha verificado esta app": "Configuracion avanzada" -> "Ir a <proyecto>".
 *   5. MARCA TODAS las casillas de permisos (si queda alguna sin marcar, la app no funcionara).
 *   6. Pulsa "Permitir". Luego revisa el "Registro de ejecucion": debe listar los permisos concedidos.
 *   Se puede ejecutar las veces que haga falta.
 *
 * API PUBLICA (nombres globales):
 *   autorizarPermisos()            Editor o menu: exige todos los permisos y registra estado, cuenta y scopes.
 *   estadoPermisos()               -> { ok, url, permisos }. Nunca lanza errores; si la consulta falla, ok = true.
 *   paginaPermisos(estado)         Funcion pura: HTML (texto) de la pagina "Faltan permisos".
 *   puertaPermisosWebApp()         Primera linea de un doGet: null si todo bien, o la pagina de permisos.
 *   menuDarPermisos()              Dialogo en la hoja/documento/presentacion/formulario (o registro si no hay).
 *   agregarItemMenuPermisos(menu)  Agrega el item "Dar permisos a la app" a un menu ya creado.
 *   PERMISOS_APP, describirScope(alcance)   Catalogo de scopes comunes y su descripcion en espanol.
 *
 * USO EN UNA WEB APP (doGet):
 *   function doGet(e) {
 *     var permisos = puertaPermisosWebApp();
 *     if (permisos) return permisos;
 *     // ... resto del doGet ...
 *   }
 *   Nota: si la web app se ejecuta como quien la implementa (USER_DEPLOYING), se comprueba la cuenta
 *   propietaria; si se ejecuta como quien accede (USER_ACCESSING), la de cada visitante.
 *
 * USO EN UN MENU DE HOJA, DOCUMENTO, PRESENTACION O FORMULARIO (este archivo NO define el evento de apertura, porque
 * chocaria con el del proyecto). Desde el evento de apertura que ya tengas, agrega el item:
 *   var menu = <servicio de interfaz del contenedor>.getUi().createMenu('Mi menu');
 *   menu.addItem('Mi opcion', 'miFuncion');
 *   agregarItemMenuPermisos(menu);
 *   menu.addToUi();
 *   (Si tu proyecto no usa menus, basta con una opcion que llame a menuDarPermisos.)
 *
 * USO EN UN SCRIPT INDEPENDIENTE: ejecuta `autorizarPermisos` desde el editor (ver arriba).
 *
 * GANCHO OPCIONAL: si el proyecto define una funcion llamada permisosVerificacionesProyecto (sin
 * argumentos) que devuelva un texto o una lista de textos, autorizarPermisos la ejecuta y agrega su
 * resultado al registro. Sirve para comprobaciones propias (acceso a un archivo, etc.) sin tocar
 * este archivo. Si no existe, se ignora.
 *
 * Privados: los nombres internos terminan en "_" para no chocar con funciones del proyecto.
 */

/** Catalogo de scopes comunes con su descripcion en espanol. Los no catalogados se muestran con su URL. */
var PERMISOS_APP = [
  { alcance: 'https://www.googleapis.com/auth/spreadsheets', descripcion: 'Ver y editar las hojas de cálculo' },
  { alcance: 'https://www.googleapis.com/auth/spreadsheets.readonly', descripcion: 'Ver las hojas de cálculo' },
  { alcance: 'https://www.googleapis.com/auth/documents', descripcion: 'Crear y editar documentos de Google' },
  { alcance: 'https://www.googleapis.com/auth/drive', descripcion: 'Ver, crear, editar y compartir archivos en Google Drive' },
  { alcance: 'https://www.googleapis.com/auth/drive.file', descripcion: 'Ver y editar solo los archivos de Drive que usa o crea esta app' },
  { alcance: 'https://www.googleapis.com/auth/drive.readonly', descripcion: 'Ver los archivos de Google Drive' },
  { alcance: 'https://www.googleapis.com/auth/script.send_mail', descripcion: 'Enviar correos en tu nombre' },
  { alcance: 'https://www.googleapis.com/auth/gmail.send', descripcion: 'Enviar correos desde tu cuenta de Gmail' },
  { alcance: 'https://www.googleapis.com/auth/userinfo.email', descripcion: 'Ver tu dirección de correo' },
  { alcance: 'https://www.googleapis.com/auth/script.container.ui', descripcion: 'Mostrar menús y ventanas en el documento' },
  { alcance: 'https://www.googleapis.com/auth/script.scriptapp', descripcion: 'Administrar la ejecución del script (activadores)' },
  { alcance: 'https://www.googleapis.com/auth/script.external_request', descripcion: 'Conectarse a servicios externos' },
  { alcance: 'https://www.googleapis.com/auth/forms', descripcion: 'Ver y editar los formularios de Google' },
  { alcance: 'https://www.googleapis.com/auth/presentations', descripcion: 'Ver y editar las presentaciones de Google' },
  { alcance: 'https://www.googleapis.com/auth/calendar', descripcion: 'Ver y editar los calendarios de Google' }
];

/** Descripcion en espanol de un scope. Si no esta catalogado devuelve el propio scope (su URL). Pura. */
function describirScope(alcance) {
  var a = String(alcance === null || alcance === undefined ? '' : alcance);
  for (var i = 0; i < PERMISOS_APP.length; i++) {
    if (PERMISOS_APP[i].alcance === a) return PERMISOS_APP[i].descripcion;
  }
  return a;
}

/** Objeto global del runtime (Apps Script V8 o Node). */
function permisosGlobal_() {
  if (typeof globalThis !== 'undefined') return globalThis;
  return this;
}

/**
 * Obtiene un servicio de Google por su nombre SIN escribirlo literalmente (ver regla de mantenimiento):
 * el analisis estatico de Apps Script solo ve texto armado en ejecucion, asi que no agrega scopes.
 * Devuelve null si el servicio no existe en este runtime.
 */
function permisosServicio_(nombre) {
  try {
    var svc = permisosGlobal_()[nombre];
    return svc || null;
  } catch (e) {
    return null;
  }
}

/** Escapa todo texto externo antes de ponerlo en el HTML. */
function escaparHtmlPermisos_(texto) {
  return String(texto === null || texto === undefined ? '' : texto).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * Consulta a Google el estado de autorizacion. Devuelve { ok, url, alcances, error }.
 * Nunca lanza: si la consulta falla, ok = true (no se bloquea la app por un fallo de la comprobacion).
 */
function permisosConsultar_() {
  var r = { ok: true, url: '', alcances: [], error: '' };
  var info = null;
  try {
    info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    r.ok = info.getAuthorizationStatus() !== ScriptApp.AuthorizationStatus.REQUIRED;
    r.url = r.ok ? '' : String(info.getAuthorizationUrl() || '');
  } catch (e) {
    r.ok = true;
    r.url = '';
    r.error = String(e && e.message ? e.message : e);
    Logger.log('Estado de los permisos: ' + r.error);
    return r;
  }
  try {
    if (info && typeof info.getAuthorizedScopes === 'function') {
      r.alcances = (info.getAuthorizedScopes() || []).map(String);
    }
  } catch (e2) {
    Logger.log('Permisos concedidos: ' + (e2 && e2.message ? e2.message : e2));
  }
  return r;
}

/**
 * Estado de los permisos de quien ejecuta: { ok, url, permisos }.
 *   ok       = ya dio todos los permisos (no REQUIRED). Si la consulta falla, true.
 *   url      = direccion para darlos (vacia si ok).
 *   permisos = lista [{ alcance, descripcion }] de los scopes concedidos hasta ahora (vacia si no se puede obtener).
 */
function estadoPermisos() {
  var c = permisosConsultar_();
  return {
    ok: c.ok,
    url: c.url,
    permisos: c.alcances.map(function (a) { return { alcance: a, descripcion: describirScope(a) }; })
  };
}

/**
 * Pagina HTML (texto) de los permisos que faltan. Funcion PURA. estado: lo que devuelve estadoPermisos().
 * Todo lo externo (direccion y descripciones) va escapado; la direccion solo se acepta si es https.
 * Si estado.ok === true se muestra una confirmacion sin boton.
 */
function paginaPermisos(estado) {
  var e = estado || {};
  var permisos = Array.isArray(e.permisos) ? e.permisos : [];
  var url = /^https:\/\//i.test(String(e.url || '')) ? String(e.url) : '';
  var completo = e.ok === true;
  var lista = permisos.map(function (p) {
    var d = p && (p.descripcion || p.alcance);
    return '<li>' + escaparHtmlPermisos_(d) + '</li>';
  }).join('');
  var bloqueLista = lista ? '<p class="sub">Permisos ya concedidos:</p><ul>' + lista + '</ul>' : '';
  var titulo = completo ? 'Todos los permisos est&aacute;n concedidos' : 'Faltan permisos';
  var cuerpo;
  if (completo) {
    cuerpo = '<p>La aplicaci&oacute;n ya tiene todos los permisos que necesita.</p>' + bloqueLista;
  } else {
    var boton = url
      ? '<a class="btn" href="' + escaparHtmlPermisos_(url) + '" target="_blank" rel="noopener">Dar permisos</a>'
      : '<p class="sin-enlace">No se pudo obtener el enlace para dar los permisos. Recarga esta p&aacute;gina.</p>';
    cuerpo = '<p>Para usar la aplicaci&oacute;n hay que dar todos los permisos que solicita.</p>' + bloqueLista + boton +
      '<p class="nota">En la pantalla de Google <strong>marca todas las casillas</strong> y pulsa Permitir. ' +
      'Si aparece &ldquo;Google no ha verificado esta app&rdquo;, elige Configuraci&oacute;n avanzada &rarr; Ir a la app. ' +
      'Despu&eacute;s de darlos, recarga esta p&aacute;gina.</p>';
  }
  return '<!DOCTYPE html><html lang="es"><head><base target="_blank"><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>*{box-sizing:border-box}body{font-family:Manrope,Arial,Helvetica,sans-serif;font-size:15px;line-height:1.5;color:#111111;margin:0;padding:24px 16px;background:#ffffff}' +
    '.caja{max-width:520px;margin:0 auto}h1{font-size:20px;font-weight:700;color:#111111;margin:0 0 12px}.sub{font-weight:700;margin:12px 0 0}' +
    'ul{margin:8px 0 16px;padding-left:22px}li{margin:4px 0}' +
    '.btn{display:flex;align-items:center;justify-content:center;width:100%;min-height:44px;padding:10px 18px;background:#003C46;color:#ffffff;border-radius:4px;text-decoration:none;font-weight:700}' +
    '.sin-enlace{color:#111111;border-left:4px solid #003C46;padding-left:10px}.nota{margin-top:16px;font-size:14px}' +
    '@media(min-width:480px){.btn{display:inline-flex;width:auto}}</style></head><body><div class="caja">' +
    '<h1>' + titulo + '</h1>' + cuerpo + '</div></body></html>';
}

/**
 * Para la PRIMERA linea de un doGet: var p = puertaPermisosWebApp(); if (p) return p;
 * Devuelve null si los permisos estan completos (o no se pudo comprobar), o un HtmlOutput con la pagina de permisos.
 */
function puertaPermisosWebApp() {
  var estado = estadoPermisos();
  if (estado.ok) return null;
  return HtmlService.createHtmlOutput(paginaPermisos(estado))
    .setTitle('Permisos requeridos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Abre un dialogo modal con la pagina de permisos en la interfaz del contenedor del script (hoja, documento,
 * presentacion o formulario). Se detecta cual existe con nombres armados en ejecucion (ver regla de
 * mantenimiento: escribir esos servicios con su nombre agregaria permisos al proyecto). Si no hay interfaz
 * disponible (script independiente), se registra la direccion en el log.
 */
function menuDarPermisos() {
  var estado = estadoPermisos();
  var html = HtmlService.createHtmlOutput(paginaPermisos(estado)).setWidth(480).setHeight(420);
  var bases = ['Spreadsheet', 'Document', 'Slides', 'Form'];
  for (var i = 0; i < bases.length; i++) {
    try {
      var svc = permisosServicio_(bases[i] + 'App');
      if (!svc) continue;
      svc.getUi().showModalDialog(html, 'Permisos de la aplicación');
      return;
    } catch (e) {
      // Este contenedor no es el del script: se prueba el siguiente.
    }
  }
  Logger.log(estado.ok
    ? 'Permisos: la aplicación ya tiene todos los permisos.'
    : 'Permisos: no hay interfaz disponible. Abre este enlace y marca todas las casillas: ' + (estado.url || '(sin enlace: ejecuta autorizarPermisos desde el editor)'));
}

/**
 * Agrega el item "Dar permisos a la app" a un menu ya creado por el evento de apertura del proyecto.
 * Devuelve el mismo menu (para encadenar). No define el evento de apertura.
 */
function agregarItemMenuPermisos(menu) {
  if (menu && typeof menu.addItem === 'function') menu.addItem('Dar permisos a la app', 'menuDarPermisos');
  return menu;
}

/** Cuenta en ejecucion, o '' si no se puede saber sin agregar permisos (nombre armado en ejecucion). */
function permisosCuenta_() {
  var s = permisosServicio_('Ses' + 'sion');
  if (!s) return '';
  try {
    var c = s.getEffectiveUser().getEmail();
    if (c) return String(c);
  } catch (e) { /* sin permiso de correo: se omite */ }
  try {
    return String(s.getActiveUser().getEmail() || '');
  } catch (e2) {
    return '';
  }
}

/**
 * Para ejecutar desde el editor o un menu: vuelve a pedir TODOS los permisos declarados por el proyecto y
 * registra en el log el estado, la cuenta y los permisos concedidos con su descripcion. Devuelve el texto.
 */
function autorizarPermisos() {
  var L = [];
  function log(msg) { L.push(msg); Logger.log(msg); }

  // Con el consentimiento granular de Google se pueden desmarcar permisos; requireAllScopes corta la
  // ejecucion y vuelve a pedir los que falten. Si este runtime no lo tiene, se informa la direccion.
  var tieneRequire = typeof ScriptApp.requireAllScopes === 'function';
  if (tieneRequire) ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);

  var c = permisosConsultar_();
  log('===== AUTORIZACIÓN DE PERMISOS =====');
  if (c.error) {
    log('Estado: no se pudo comprobar (' + c.error + ').');
  } else if (c.ok) {
    log('Estado: permisos completos.');
  } else {
    log('Estado: FALTAN PERMISOS' + (tieneRequire ? '' : ' (este entorno no puede pedirlos de nuevo desde aquí)') + '.');
    log('Abre este enlace, marca TODAS las casillas y pulsa Permitir: ' + (c.url || '(sin enlace)'));
  }
  var cuenta = permisosCuenta_();
  log('Cuenta: ' + (cuenta || '(no disponible)'));

  if (c.alcances.length) {
    log('');
    log('Permisos concedidos:');
    c.alcances.forEach(function (a) {
      var d = describirScope(a);
      log('  - ' + (d === a ? a : d + ' (' + a + ')'));
    });
  }

  if (typeof permisosVerificacionesProyecto === 'function') {
    try {
      var extra = permisosVerificacionesProyecto();
      var lineas = Array.isArray(extra) ? extra : (extra ? [extra] : []);
      if (lineas.length) log('');
      lineas.forEach(function (t) { log(String(t)); });
    } catch (ex) {
      log('Verificaciones del proyecto: ' + (ex && ex.message ? ex.message : ex));
    }
  }
  log('===== FIN =====');
  return L.join('\n');
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    paginaPermisos: paginaPermisos,
    PERMISOS_APP: PERMISOS_APP,
    describirScope: describirScope,
    estadoPermisos: estadoPermisos,
    puertaPermisosWebApp: puertaPermisosWebApp,
    menuDarPermisos: menuDarPermisos,
    agregarItemMenuPermisos: agregarItemMenuPermisos,
    autorizarPermisos: autorizarPermisos
  };
}
