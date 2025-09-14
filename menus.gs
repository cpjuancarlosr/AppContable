//== File: menus.gs (menú “JC Contable” y asignación de botones)

/**
 * Builds the main "JC Contable" custom menu.
 */
function buildMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('JC Contable')
      .addSubMenu(ui.createMenu('SAT')
          .addItem('Configurar FIEL', 'configurarFiel_')
          .addItem('Descargar CFDI (Directo)', 'descargarCFDI_')
          .addSeparator()
          .addItem('Importar XML desde Carpeta Drive', 'importarDesdeCarpetaDrive'))
      .addSubMenu(ui.createMenu('Bancos')
          .addItem('Subir Estado de Cuenta', 'subirMovimientosBancarios_'))
      .addSubMenu(ui.createMenu('Procesos')
          .addItem('Armar Pólizas', 'generarPolizas')
          .addItem('Conciliar', 'conciliarMovimientos_')
          .addItem('Recalcular Estados', 'recalcularEstadosFinancieros'))
      .addSubMenu(ui.createMenu('Exportar')
          .addItem('Exportar Paquete PDF', 'exportarPaquetePDF'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Sistema')
          .addItem('Instalar/Actualizar Sistema', 'setupSistema')
          .addItem('Ayuda', 'mostrarAyuda_'))
      .addToUi();
}

/**
 * Shows a dialog to upload a bank statement file.
 */
function subirMovimientosBancarios_() {
  const html = HtmlService.createHtmlOutputFromFile('FileUpload')
      .setWidth(400)
      .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Subir Estado de Cuenta');
}


// Placeholder functions for menu items
function configurarFiel_() {
  SpreadsheetApp.getUi().alert('Función "Configurar FIEL" no implementada aún.');
}
function descargarCFDI_() {
  SpreadsheetApp.getUi().alert('Función "Descargar CFDI" no implementada aún.');
}
function conciliarMovimientos_() {
    SpreadsheetApp.getUi().alert('Función "Conciliar" no implementada aún.');
}
function mostrarAyuda_() {
    SpreadsheetApp.getUi().alert('Guía Rápida:\n1. Use Sistema -> Instalar para configurar.\n2. Configure su FIEL en SAT -> Configurar FIEL.\n3. Descargue CFDI y suba estados de cuenta.\n4. Genere pólizas y concilie desde el menú Procesos.');
}
