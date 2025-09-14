//== File: styles.gs (tema, formatos, fuente Rubik, colores)

/**
 * Applies the visual theme to the entire spreadsheet.
 * This includes setting the font, colors, and basic formatting.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 */
function aplicarTemaVisual(ss) {
  const paleta = {
    naranja: '#DD9F86',
    azul: '#A6C2DB',
    texto: '#111111',
    fondo: '#FFFFFF',
    grisSuave: '#F5F5F5'
  };

  // Set the overall spreadsheet theme. This affects charts, etc.
  // This requires RGB integer values.
  const theme = ss.getSpreadsheetTheme()
    .setConcreteColor(SpreadsheetApp.ThemeColorType.TEXT, 17, 17, 17)
    .setConcreteColor(SpreadsheetApp.ThemeColorType.BACKGROUND, 255, 255, 255)
    .setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT1, 221, 159, 134)
    .setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT2, 166, 194, 219)
    .setFontFamily("Rubik");
  ss.setSpreadsheetTheme(theme);

  // Apply specific formatting directly to cells, which uses hex strings.
  ss.getSheets().forEach(sh => {
    const sheetRange = sh.getRange("A1:Z1000");
    sheetRange.setFontFamily("Rubik").setFontSize(10);

    const headerRange = sh.getRange(1, 1, 1, sh.getMaxColumns());
    headerRange
      .setFontSize(11)
      .setFontWeight('bold')
      .setBackground(paleta.azul)
      .setFontColor(paleta.texto)
      .setBorder(null, null, true, null, null, null, paleta.texto, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Apply zebra striping to sheets that have more than a header
    if (sh.getMaxRows() > 1) {
        try {
            const dataRange = sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns());
            const banding = dataRange.getBandings()[0] || dataRange.applyRowBanding();
            banding.setFirstRowColor(paleta.fondo)
                    .setSecondRowColor(paleta.grisSuave)
                    .setHeaderRowColor(paleta.azul);
        } catch(e) {
            log_(`Could not apply banding to sheet: ${sh.getName()}. Error: ${e.message}`);
        }
    }
  });
}
