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

  const theme = ss.getSpreadsheetTheme()
    .setConcreteColor(SpreadsheetApp.ThemeColorType.TEXT, paleta.texto)
    .setConcreteColor(SpreadsheetApp.ThemeColorType.BACKGROUND, paleta.fondo)
    .setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT1, paleta.naranja)
    .setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT2, paleta.azul)
    .setFontFamily("Rubik");

  ss.setSpreadsheetTheme(theme);

  ss.getSheets().forEach(sh => {
    // Set default font size for the entire sheet
    sh.getRange("A1:Z1000").setFontSize(10).setFontFamily("Rubik");

    // Style headers
    const headerRange = sh.getRange(1, 1, 1, sh.getMaxColumns());
    headerRange
      .setFontSize(11)
      .setFontWeight('bold')
      .setBackground(paleta.azul)
      .setFontColor(paleta.texto)
      .setBorder(null, null, true, null, null, null, paleta.texto, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Apply zebra striping
    try {
      const dataRange = sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns());
      const banding = dataRange.getBandings()[0] || dataRange.applyRowBanding();
      banding.setFirstRowColor(paleta.fondo).setSecondRowColor(paleta.grisSuave).setHeaderRowColor(paleta.azul);
    } catch(e) {
      // Ignore errors on sheets where banding is not applicable
    }
  });
}
