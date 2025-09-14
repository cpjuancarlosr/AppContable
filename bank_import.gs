//== File: bank_import.gs (PDF/CSV bancario → Bancos)

/**
 * Entry point for importing bank transactions from a file uploaded via the UI.
 * @param {object} fileData An object containing the fileName, mimeType, and base64Data of the uploaded file.
 * @return {string} A status message for the user.
 */
function importarMovimientosBancarios(fileData) {
  try {
    const { fileName, mimeType, base64Data } = fileData;
    if (!fileName || !mimeType || !base64Data) {
      throw new Error("Datos del archivo incompletos.");
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);

    if (mimeType === 'application/pdf') {
      return parsearPdfBancario_(blob);
    } else if (mimeType.includes('csv') || mimeType === 'application/vnd.ms-excel') {
      return parsearCsvBancario_(blob);
    } else {
      throw new Error(`Tipo de archivo no soportado: ${mimeType}. Por favor, suba un PDF o CSV.`);
    }
  } catch (e) {
    log_(`Error en importarMovimientosBancarios: ${e.message}`);
    throw new Error(`Error en el servidor: ${e.message}`);
  }
}

/**
 * Parses a PDF bank statement by extracting its text and applying bank-specific regex.
 * @param {GoogleAppsScript.Base.Blob} pdfBlob The PDF file blob.
 * @return {string} A status message.
 */
function parsearPdfBancario_(pdfBlob) {
  const sh = SpreadsheetApp.getActive().getSheetByName("Bancos");
  const textContent = pdfBlob.getDataAsString('UTF-8');

  if (textContent.trim().length < 100) {
    return `El archivo PDF "${pdfBlob.getName()}" parece no contener texto extraíble. Por favor, intente con un archivo CSV.`;
  }

  let bank = 'desconocido';
  if (textContent.match(/BBVA/gi)) bank = 'bbva';
  else if (textContent.match(/Banorte/gi)) bank = 'banorte';
  else if (textContent.match(/Santander/gi)) bank = 'santander';
  else if (textContent.match(/NU MEXICO/gi)) bank = 'nu';

  const regexPatterns = {
    bbva: /(\d{2}\s\w{3})\s+(.*?)\s+([\d,]+\.\d{2})\s?([\d,]+\.\d{2})?/,
    banorte: /(\d{2}\/\d{2}\/\d{4})\s+(.*?)\s+([\d,]+\.\d{2})/,
    // Add other bank patterns here
  };

  const pattern = regexPatterns[bank];
  if (!pattern) {
    return `Banco no reconocido o no soportado para el archivo: ${pdfBlob.getName()}.`;
  }

  const lines = textContent.split('\n');
  const nuevasFilas = [];
  let currentYear = new Date().getFullYear(); // Assume current year if not specified

  lines.forEach(line => {
    const match = line.match(pattern);
    if (match) {
      try {
        const fechaStr = match[1];
        const concepto = match[2].trim();
        const monto1 = parseFloat(match[3].replace(/,/g, '')) || 0;
        const monto2 = parseFloat(match[4]?.replace(/,/g, '')) || 0;

        const cargo = (bank === 'bbva' && monto2 > 0) ? monto1 : (monto1 > 0 ? monto1 : 0);
        const abono = (bank === 'bbva' && monto2 > 0) ? monto2 : (monto1 < 0 ? Math.abs(monto1) : 0);

        // Basic date parsing, needs improvement
        const fecha = new Date(`${fechaStr} ${currentYear}`);
        const naturaleza = cargo > 0 ? 'Cargo' : 'Abono';
        const importe = cargo > 0 ? cargo : abono;

        nuevasFilas.push([
          fecha, concepto, '', importe, naturaleza, '', bank.toUpperCase(), '', '', 'PDF', ''
        ]);
      } catch(e) {
        // Ignore lines that match but fail to parse
      }
    }
  });

  if (nuevasFilas.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
    return `Importación de PDF exitosa. Se agregaron ${nuevasFilas.length} movimientos desde ${pdfBlob.getName()}.`;
  } else {
    return `No se encontraron movimientos que coincidieran con el patrón del banco '${bank}' en ${pdfBlob.getName()}.`;
  }
}


/**
 * Parses a CSV bank statement and adds the data to the 'Bancos' sheet.
 * @param {GoogleAppsScript.Base.Blob} csvBlob The CSV file blob.
 * @return {string} A status message.
 */
function parsearCsvBancario_(csvBlob) {
  const sh = SpreadsheetApp.getActive().getSheetByName("Bancos");
  const csvData = csvBlob.getDataAsString('UTF-8');
  const rows = Utilities.parseCsv(csvData);

  if (rows.length < 2) {
    return "El archivo CSV está vacío o no contiene datos.";
  }

  const headers = rows.shift().map(h => h.toLowerCase().trim());
  const headerMap = {
    fecha: headers.indexOf('fecha'),
    descripcion: headers.indexOf('concepto') !== -1 ? headers.indexOf('concepto') : headers.indexOf('descripcion'),
    referencia: headers.indexOf('referencia'),
    cargo: headers.indexOf('cargo'),
    abono: headers.indexOf('abono'),
    monto: headers.indexOf('monto')
  };

  if (headerMap.fecha === -1) {
    throw new Error("El CSV debe contener una columna de 'Fecha'.");
  }

  const nuevasFilas = rows.map(row => {
    const fecha = new Date(row[headerMap.fecha]);
    let cargo = parseFloat(row[headerMap.cargo] || 0);
    let abono = parseFloat(row[headerMap.abono] || 0);
    const monto = parseFloat(row[headerMap.monto] || 0);

    if (monto !== 0 && cargo === 0 && abono === 0) {
        if (monto < 0) cargo = Math.abs(monto);
        else abono = monto;
    }

    const naturaleza = (cargo > 0) ? 'Cargo' : 'Abono';
    const importe = cargo > 0 ? cargo : abono;

    return [
      fecha, row[headerMap.descripcion] || '', row[headerMap.referencia] || '',
      importe, naturaleza, '', '', '', '', 'OK', ''
    ];
  }).filter(row => !isNaN(row[0].getTime()));

  if (nuevasFilas.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
    return `Importación exitosa. Se agregaron ${nuevasFilas.length} movimientos desde ${csvBlob.getName()}.`;
  } else {
    return "No se encontraron filas de datos válidas en el CSV.";
  }
}
