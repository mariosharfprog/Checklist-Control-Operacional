/**
 * SISTEMA SHARF - SINCRONIZACIÓN DIRECTA DE CECOSOFI
 * Archivo: SyncCeCos.gs
 * 
 * Copia la hoja completa CeCosOfi desde el libro origen a la pestaña 'cecos',
 * preservando todas sus filas, columnas y formatos numéricos (DNI, fechas, etc.)
 * sin congelar la hoja de cálculo.
 */

var SYNC_CONFIG = {
  // ID del libro CeCosOfi original
  ORIGEN_SPREADSHEET_ID: "12mfCkJdIOn8P6hQr4FN3Rk2VnmnKHClr95WldqBSDm8",
  ORIGEN_HOJA_NOMBRE: "CeCosOfi",
  
  // Nombre de la pestaña destino en tu libro actual
  DESTINO_HOJA_NOMBRE: "cecos"
};

/**
 * Función principal para extraer y copiar la hoja completa CeCosOfi
 */
function sincronizarCeCosOfi() {
  try {
    var ssDestino = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Conectar al libro origen CeCosOfi
    var ssOrigen = SpreadsheetApp.openById(SYNC_CONFIG.ORIGEN_SPREADSHEET_ID);
    var sheetOrigen = ssOrigen.getSheetByName(SYNC_CONFIG.ORIGEN_HOJA_NOMBRE) || ssOrigen.getSheets()[0];
    
    if (!sheetOrigen) {
      throw new Error("No se encontró la pestaña '" + SYNC_CONFIG.ORIGEN_HOJA_NOMBRE + "' en el libro origen.");
    }
    
    // 2. Leer datos y formatos numéricos del origen en memoria RAM del servidor
    var rangeOrigen = sheetOrigen.getDataRange();
    var valores = rangeOrigen.getValues();
    var formatos = rangeOrigen.getNumberFormats();
    if (valores.length <= 1) {
      throw new Error("La hoja CeCosOfi no contiene registros suficientes.");
    }

    var numFilas = valores.length;
    var numCols = valores[0].length;
    
    // 3. Preparar la pestaña de destino en tu libro actual
    var sheetDestino = ssDestino.getSheetByName(SYNC_CONFIG.DESTINO_HOJA_NOMBRE);
    if (!sheetDestino) {
      sheetDestino = ssDestino.insertSheet(SYNC_CONFIG.DESTINO_HOJA_NOMBRE);
    }
    
    var maxFilas = sheetDestino.getMaxRows();
    var maxCols = sheetDestino.getMaxColumns();
    if (numFilas > maxFilas) sheetDestino.insertRowsAfter(maxFilas, numFilas - maxFilas);
    if (numCols > maxCols) sheetDestino.insertColumnsAfter(maxCols, numCols - maxCols);

    // 4. Limpiar el destino por completo (contenido y formatos) para no dejar residuos
    sheetDestino.clear();

    // 5. Escribir formatos y valores con el mismo tamaño en una sola llamada cada una
    var rangeDestino = sheetDestino.getRange(1, 1, numFilas, numCols);
    rangeDestino.setNumberFormats(formatos);
    rangeDestino.setValues(valores);
    
    // 6. Estilos visuales para las cabeceras (toda la fila 1)
    sheetDestino.getRange(1, 1, 1, numCols)
      .setFontWeight("bold")
      .setBackground("#F1F1F1")
      .setFontColor("#111111");
    
    SpreadsheetApp.flush();
    
    var mensajeExito = "¡Sincronización exitosa! Se cargaron " + (numFilas - 1) + " registros en la pestaña '" + SYNC_CONFIG.DESTINO_HOJA_NOMBRE + "'.";
    Logger.log(mensajeExito);
    
    try {
      SpreadsheetApp.getUi().alert(mensajeExito);
    } catch(eUi) {}
    
    return { exito: true, total: numFilas - 1 };
    
  } catch (error) {
    var errorMsg = "Error al sincronizar CeCosOfi: " + error.toString();
    Logger.log(errorMsg);
    try {
      SpreadsheetApp.getUi().alert(errorMsg);
    } catch(eUi) {}
    return { exito: false, error: errorMsg };
  }
}

/**
 * Menú personalizado en Google Sheets para ejecutar con 1 clic
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Sistema SHARF")
      .addItem("🔄 Sincronizar CeCosOfi Ahora", "sincronizarCeCosOfi")
      .addToUi();
  } catch(e) {}
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    sincronizarCeCosOfi: sincronizarCeCosOfi,
    SYNC_CONFIG: SYNC_CONFIG
  };
}
