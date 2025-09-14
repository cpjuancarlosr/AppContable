//== File: Code.gs (entrypoint, onOpen, rutas de menú y botones)

/**
 * @OnlyCurrentDoc
 * The above comment directs App Script to limit the scope of authorization
 * to only the current spreadsheet.
 */

/**
 * Runs when the spreadsheet is opened.
 * @param {object} e The event object.
 */
function onOpen(e) {
  buildMenu_();
}

/**
 * Main setup function to initialize the entire spreadsheet.
 */
function setupSistema() {
  const ss = SpreadsheetApp.getActive();
  crearHojas_(ss);
  setupConfiguracion_(ss.getSheetByName("CFG"));
  aplicarTemaVisual(ss);
  // To be implemented: setupDashboard_(ss.getSheetByName("Inicio"));
  SpreadsheetApp.getUi().alert("El sistema ha sido instalado y configurado con el nuevo tema visual.");
}

/**
 * Creates all necessary sheets with their headers and basic formatting.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 */
function crearHojas_(ss) {
  const hojas = {
    "Inicio": [["KPI", "Valor", "Tendencia"], ["", "", ""]],
    "CFG": [["Clave", "Valor", "Descripción"]],
    "Cuentas": [["Codigo", "Nombre", "Naturaleza", "Nivel", "Agrupador SAT"]],
    "Entidades": [["RFC", "Razón Social", "Tipo", "Cuenta Contable", "Clave SAT Prod/Serv", "Unidad SAT", "Tasa IVA Predet."]],
    "XML": [["UUID", "Emisor RFC", "Emisor Nombre", "Receptor RFC", "Receptor Nombre", "Tipo", "Fecha Emisión", "Fecha Timbrado", "Método Pago", "Forma Pago", "Subtotal", "IVA 16%", "IVA 8%", "IVA 0%", "IVA Exento", "Total", "Uso CFDI", "Moneda", "UUIDs Relacionados", "Link"]],
    "Bancos": [["Fecha", "Concepto", "Referencia", "Monto", "Naturaleza", "Cuenta Banco", "Etiquetas", "Folio Factura", "UUID Conciliado", "Calidad Dato", "Link"]],
    "Polizas": [["Fecha", "Tipo", "Folio", "Cuenta", "Subcuenta", "Concepto", "Referencia", "Debe", "Haber", "UUID", "Origen", "Estatus", "Link"]]
  };

  // Ensure 'Inicio' is the first sheet
  const sheetNames = [
    "Inicio", "CFG", "Cuentas", "Entidades", "XML", "Bancos", "Polizas"
  ];

  sheetNames.forEach((nombre, index) => {
    let sh = ss.getSheetByName(nombre);
    if (!sh) {
      sh = ss.insertSheet(nombre, index);
    }
    sh.clear();
    const headers = hojas[nombre];
    sh.getRange(1, 1, headers.length, headers[0].length).setValues(headers).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (ss.getActiveSheet().getSheetId() !== sh.getSheetId()) {
      sh.hideSheet();
    }
  });

  // Activate the 'Inicio' sheet
  const inicioSheet = ss.getSheetByName("Inicio");
  if (inicioSheet) {
    inicioSheet.showSheet().activate();
  }
}
