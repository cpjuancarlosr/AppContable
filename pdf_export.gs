//== File: pdf_export.gs (salida a PDF con portada)

/**
 * Exports a package of reports to a single PDF file in a designated Drive folder.
 */
function exportarPaquetePDF() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActive();
    const folderId = getCfg("ID_CARPETA_PDF");
    if (!folderId) {
      throw new Error("El ID de la carpeta de PDFs no está configurado en la hoja 'CFG'.");
    }
    const folder = DriveApp.getFolderById(folderId);

    // For now, just export the 'Inicio' sheet as a proof of concept.
    const shInicio = ss.getSheetByName("Inicio");
    const periodo = "YYYY-MM"; // Placeholder for actual period

    const blob = exportSheetAsPdfBlob_(ss, shInicio, {portrait: false, size: 'LETTER'});
    const fileName = `Reporte Financiero - ${periodo}.pdf`;
    const pdfFile = folder.createFile(blob).setName(fileName);

    ui.alert(`Reporte exportado exitosamente.`, `El archivo "${fileName}" ha sido guardado en la carpeta "${folder.getName()}".\n\nURL: ${pdfFile.getUrl()}`);

  } catch (e) {
    log_(`Error en exportarPaquetePDF: ${e.message}`);
    ui.alert(`Error al exportar: ${e.message}`);
  }
}

/**
 * Creates a PDF blob from a given sheet with specified options.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to export.
 * @param {object} options PDF export options.
 * @return {GoogleAppsScript.Base.Blob} The generated PDF blob.
 */
function exportSheetAsPdfBlob_(ss, sheet, options = {}) {
  const spreadsheetId = ss.getId();
  const sheetId = sheet.getSheetId();

  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    `format=pdf` +
    `&gid=${sheetId}` +
    `&portrait=${options.portrait || 'true'}` +
    `&size=${options.size || 'A4'}` +
    `&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`
    }
  });

  return response.getBlob().setName(`${sheet.getName()}.pdf`);
}
