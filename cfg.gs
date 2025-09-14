//== File: cfg.gs (constantes, rangos con nombre, mapeos por banco)

const CFG_DEFAULTS = {
  "ID_CARPETA_XML": "",
  "ID_CARPETA_PDF": "",
  "ID_CARPETA_BANCOS": "",
  "RUTA_FIEL_CER": "",
  "RUTA_FIEL_KEY": "",
  "CONTRASENA_FIEL": "NO GUARDAR AQUI - USAR MENU",
  "CFDI_TIPO_DESCARGA": "Recibidos",
  "CFDI_VERSION": "4.0",
  "IVA_TASA_GENERAL": 0.16,
  "IVA_TASA_FRONTERA": 0.08,
  "ISR_COEF_UTILIDAD": 0.15,
  "TRIGGER_DIARIO_ACTIVO": "FALSE"
};

const CFG_DESCRIPTIONS = {
  "ID_CARPETA_XML": "ID de la carpeta en Drive para guardar los XML descargados.",
  "ID_CARPETA_PDF": "ID de la carpeta en Drive para exportar los reportes PDF.",
  "ID_CARPETA_BANCOS": "ID de la carpeta en Drive donde se buscan los estados de cuenta.",
  "RUTA_FIEL_CER": "Ruta completa en Drive al archivo .cer de la FIEL.",
  "RUTA_FIEL_KEY": "Ruta completa en Drive al archivo .key de la FIEL.",
  "CONTRASENA_FIEL": "NUNCA guarde la contraseña aquí. Use el menú 'SAT > Configurar FIEL' para almacenarla de forma segura.",
  "CFDI_TIPO_DESCARGA": "Tipo de CFDI a descargar por defecto (Emitidos o Recibidos).",
  "CFDI_VERSION": "Versión del CFDI a descargar (ej. 4.0).",
  "IVA_TASA_GENERAL": "Tasa de IVA general (ej. 0.16 para 16%).",
  "IVA_TASA_FRONTERA": "Tasa de IVA para zona fronteriza (ej. 0.08 para 8%).",
  "ISR_COEF_UTILIDAD": "Coeficiente de utilidad para el cálculo de ISR de Pagos Provisionales.",
  "TRIGGER_DIARIO_ACTIVO": "Poner en TRUE para activar la descarga automática diaria de CFDI."
};


/**
 * Populates the 'CFG' sheet with default parameters.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh The 'CFG' sheet.
 */
function setupConfiguracion_(sh) {
  const data = Object.keys(CFG_DEFAULTS).map(key => {
    return [key, CFG_DEFAULTS[key], CFG_DESCRIPTIONS[key] || ""];
  });

  sh.getRange(2, 1, data.length, 3).setValues(data);
  sh.autoResizeColumns(1, 3);

  // Add data validation for specific fields
  const tipoDescargaRule = SpreadsheetApp.newDataValidation().requireValueInList(['Emitidos', 'Recibidos']).build();
  sh.getRange("B2:B").setDataValidation(tipoDescargaRule); // Assuming CFDI_TIPO_DESCARGA is in row 2

  const triggerRule = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE']).build();
  sh.getRange("B12:B12").setDataValidation(triggerRule); // Assuming TRIGGER_DIARIO_ACTIVO is in row 12

  // Protect the password cell to warn users
  const protection = sh.getRange("B6").protect(); // Assuming CONTRASENA_FIEL is in row 6
  protection.setDescription('No editar directamente. Usar el menú.');
  protection.setWarningOnly(true);
}


// Caching mechanism to reduce sheet reads
let CFG_CACHE = null;

function getCfg(key) {
  if (CFG_CACHE === null) {
    const sh = SpreadsheetApp.getActive().getSheetByName("CFG");
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    CFG_CACHE = data.reduce((obj, row) => {
      if (row[0]) obj[row[0]] = row[1];
      return obj;
    }, {});
  }
  return CFG_CACHE[key];
}

function setCfg(key, value) {
    const sh = SpreadsheetApp.getActive().getSheetByName("CFG");
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const rowIndex = data.findIndex(row => row[0] === key);
    if (rowIndex !== -1) {
        sh.getRange(rowIndex + 2, 2).setValue(value);
        if(CFG_CACHE) CFG_CACHE[key] = value; // Update cache
    } else {
        sh.appendRow([key, value, ""]); // Add if not found
    }
}
