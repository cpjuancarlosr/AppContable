//== File: polizas.gs (reglas contables y subcuentas dinámicas)

/**
 * Main function to generate journal entries (pólizas) for all unprocessed transactions.
 */
function generarPolizas() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert("Iniciando la generación de pólizas. Este proceso puede tardar.");

    const ss = SpreadsheetApp.getActive();
    const shXml = ss.getSheetByName("XML");
    const shPolizas = ss.getSheetByName("Polizas");

    const xmlData = shXml.getRange(2, 1, shXml.getLastRow() - 1, shXml.getLastColumn()).getValues();
    const polizasGeneradas = shPolizas.getRange(2, 1, shPolizas.getLastRow() - 1, shPolizas.getLastColumn()).getValues();

    // Create a set of UUIDs that already have a póliza
    const uuidsProcesados = new Set(polizasGeneradas.map(p => p[9]).filter(Boolean));

    const nuevasPolizas = [];

    xmlData.forEach((row, index) => {
      const uuid = row[0];
      if (!uuid || uuidsProcesados.has(uuid)) {
        return; // Skip if no UUID or already processed
      }

      const asiento = crearAsientoDesdeXML_(row);
      if (asiento.length > 0) {
        nuevasPolizas.push(...asiento);
      }
    });

    if (nuevasPolizas.length > 0) {
      shPolizas.getRange(shPolizas.getLastRow() + 1, 1, nuevasPolizas.length, nuevasPolizas[0].length).setValues(nuevasPolizas);
      ui.alert(`${nuevasPolizas.length / 2} pólizas han sido generadas exitosamente.`); // Assuming 2 rows per entry
    } else {
      ui.alert("No se encontraron nuevas transacciones para generar pólizas.");
    }

  } catch (e) {
    log_(`Error en generarPolizas: ${e.message}`);
    ui.alert(`Error: ${e.message}`);
  }
}

/**
 * Creates a journal entry for a given row of XML data based on simple rules.
 * This is a simplified version and should be expanded with a proper rule engine.
 * @param {Array} xmlRow A single row of data from the 'XML' sheet.
 * @return {Array<Array>} An array of rows representing the journal entry.
 */
function crearAsientoDesdeXML_(xmlRow) {
  const [uuid, rfcEmisor, , rfcReceptor, , tipo, fecha, , metodoPago, formaPago, subtotal, iva16, iva8, iva0, , total] = xmlRow;

  const esIngreso = (tipo === 'I');
  const esEgreso = (tipo === 'E');
  const esPago = (tipo === 'P');

  let asiento = [];
  const concepto = `CFDI ${tipo} ${uuid.substring(0, 8)}`;
  const ivaTotal = iva16 + iva8 + iva0;

  if (esIngreso) {
    // Basic Sale Entry
    const cuentaCliente = obtenerSubcuentaDinamica_(rfcReceptor, 'Cliente');
    asiento.push([fecha, 'Ingreso', uuid.substring(0,8), '105', cuentaCliente, concepto, '', total, 0, uuid, 'XML', 'OK', '']); // Clientes (Cargo)
    asiento.push([fecha, 'Ingreso', uuid.substring(0,8), '400-000', '', concepto, '', 0, subtotal, uuid, 'XML', 'OK', '']); // Ventas (Abono)
    if (ivaTotal > 0) {
      asiento.push([fecha, 'Ingreso', uuid.substring(0,8), '240-200', '', concepto, '', 0, ivaTotal, uuid, 'XML', 'OK', '']); // IVA Trasladado (Abono)
    }
  } else if (esEgreso) {
    // Basic Expense/Purchase Entry
    const cuentaProveedor = obtenerSubcuentaDinamica_(rfcEmisor, 'Proveedor');
    asiento.push([fecha, 'Egreso', uuid.substring(0,8), '510-000', '', concepto, '', subtotal, 0, uuid, 'XML', 'OK', '']); // Gasto (Cargo)
    if (ivaTotal > 0) {
      asiento.push([fecha, 'Egreso', uuid.substring(0,8), '240-300', '', concepto, '', ivaTotal, 0, uuid, 'XML', 'OK', '']); // IVA Acreditable (Cargo)
    }
     asiento.push([fecha, 'Egreso', uuid.substring(0,8), '201', cuentaProveedor, concepto, '', 0, total, uuid, 'XML', 'OK', '']); // Proveedores (Abono)
  }

  return asiento;
}


/**
 * Gets or creates a dynamic sub-account for a client or provider.
 * @param {string} rfc The RFC of the entity.
 * @param {string} tipo 'Cliente' or 'Proveedor'.
 * @return {string} The sub-account code.
 */
function obtenerSubcuentaDinamica_(rfc, tipo) {
    const ss = SpreadsheetApp.getActive();
    const shEntidades = ss.getSheetByName("Entidades");
    const data = shEntidades.getRange(2, 1, shEntidades.getLastRow() - 1, 4).getValues();

    let entidad = data.find(row => row[0] === rfc);

    if (entidad && entidad[3]) {
        return entidad[3]; // Return existing account if present
    }

    // If not found or no account, create a new one
    const prefijo = (tipo === 'Cliente') ? '105' : '201';
    const nuevoCodigo = prefijo + '.' + (data.filter(row => row[3] && row[3].startsWith(prefijo)).length + 1).toString().padStart(2, '0');

    if (entidad) {
        // Update existing entity
        const rowIndex = data.findIndex(row => row[0] === rfc) + 2;
        shEntidades.getRange(rowIndex, 4).setValue(nuevoCodigo);
    } else {
        // This case should be handled by the XML import creating the entity first.
        // For safety, we just return a temporary code.
    }

    return nuevoCodigo;
}
