/**
 * SISTEMA DE INSPECCIÓN DE PROCESOS - SHARF
 * Archivo: Server.gs — Punto de entrada de la aplicación Web (doGet)
 */

/**
 * Servidor Web Application UI.
 * Con ?vista=revision sirve la página de revisión de memorandos (Revision.html, ver Memo.js);
 * en cualquier otro caso sirve Index.html, igual que siempre.
 */
function doGet(e) {
  try {
    // Si a quien entra le faltan permisos (p. ej. desmarcó casillas en la pantalla de Google),
    // se muestra la página "Dar permisos" en lugar de fallar más adelante. Ver Permisos.js.
    var paginaPermisos = (typeof puertaPermisosWebApp === "function") ? puertaPermisosWebApp() : null;
    if (paginaPermisos) return paginaPermisos;

    if (e && e.parameter && e.parameter.vista === "revision") {
      var paginaRevision = HtmlService.createTemplateFromFile('Revision').evaluate();
      paginaRevision.setTitle('Revisión de memorandos - SHARF');
      paginaRevision.addMetaTag('viewport', 'width=device-width, initial-scale=1');
      paginaRevision.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      return paginaRevision;
    }
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Index');
    htmlOutput.setTitle('Inspección de Procesos - SHARF');
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return htmlOutput;
  } catch (err) {
    Logger.log("Error en doGet: " + err.toString());
    return HtmlService.createHtmlOutput("Error al cargar la interfaz: " + err.toString());
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { doGet: doGet };
}
