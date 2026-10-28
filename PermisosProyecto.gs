/**
 * PermisosProyecto.js - Comprobaciones ESPECIFICAS del checklist para el traspaso del proyecto.
 *
 * Permisos.js es el archivo GLOBAL (idéntico en todos los proyectos): define `autorizarPermisos`, que pide
 * todos los permisos y, al final, ejecuta el gancho `permisosVerificacionesProyecto` definido aquí.
 *
 * CÓMO USARLO: en el editor de Apps Script, elige `autorizarPermisos` y pulsa Ejecutar. Luego abre
 * "Registro de ejecución": cada comprobación debe salir OK (si sale FALLA, el mensaje dice qué revisar).
 * Solo lee (nombres de libro/carpeta/plantilla y cuota de correo): no modifica nada ni envía correos ni
 * hace peticiones a la red.
 */

/**
 * Devuelve la lista de líneas de resultado (texto) que autorizarPermisos agrega al registro.
 * Cada comprobación va en su propio try/catch: una falla no impide revisar las demás.
 */
function permisosVerificacionesProyecto() {
  var CFG = (typeof CONFIG !== 'undefined') ? CONFIG : {};
  var idCecos = (typeof SYNC_CONFIG !== 'undefined') ? SYNC_CONFIG.ORIGEN_SPREADSHEET_ID : '';

  var pruebas = [
    ['Cuenta en sesión', 'sin acceso a la sesión: vuelve a ejecutar y acepta todos los permisos.', function() {
      return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '(sin correo visible)';
    }],
    ['Google Sheets - libro del checklist', 'el script debe estar ligado al libro del checklist (abre el editor desde Extensiones ▸ Apps Script de ese libro).', function() {
      return SpreadsheetApp.getActiveSpreadsheet().getName();
    }],
    ['Google Sheets - libro origen CeCosOfi', 'revisa SYNC_CONFIG.ORIGEN_SPREADSHEET_ID en SyncCeCos.js o pide acceso a ese libro.', function() {
      return SpreadsheetApp.openById(idCecos).getName();
    }],
    ['Google Drive - carpeta de reportes', 'revisa CONFIG.ID_CARPETA_REPORTES en Config.js o pide acceso a esa carpeta.', function() {
      return DriveApp.getFolderById(CFG.ID_CARPETA_REPORTES).getName();
    }],
    ['Google Docs - plantilla del memorando', 'revisa CONFIG.MEMO_PLANTILLA_ID en Config.js o pide acceso a esa plantilla.', function() {
      return DocumentApp.openById(CFG.MEMO_PLANTILLA_ID).getName();
    }],
    ['Gmail - envío de correos (MailApp)', 'vuelve a ejecutar y marca la casilla de enviar correos.', function() {
      return 'cuota restante hoy: ' + MailApp.getRemainingDailyQuota() + ' correos';
    }]
  ];

  var lineas = ['===== COMPROBACIONES DEL CHECKLIST ====='];
  var fallas = 0;
  pruebas.forEach(function(p) {
    try {
      lineas.push('  OK     ' + p[0] + '  ->  ' + p[2]());
    } catch (e) {
      fallas++;
      lineas.push('  FALLA  ' + p[0] + '  ->  ' + (e && e.message ? e.message : String(e)) + '  (Qué revisar: ' + p[1] + ')');
    }
  });

  if (CFG.MEMO_URL_FIRMA) {
    lineas.push('  INFO   URL de firma de memorandos (firma-memo): ' + CFG.MEMO_URL_FIRMA + '  ->  no se comprueba desde aquí; debe apuntar a la implementación vigente del proyecto firma-memo.');
  }

  lineas.push(fallas === 0
    ? 'Listo: la cuenta tiene todos los permisos que usa el checklist.'
    : '*** ' + fallas + ' comprobación(es) con FALLA: si es de permisos, vuelve a ejecutar y marca todas las casillas; si dice que no existe o no tienes acceso, revisa el ID indicado o pide acceso a ese archivo. ***');
  return lineas;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { permisosVerificacionesProyecto: permisosVerificacionesProyecto };
}
