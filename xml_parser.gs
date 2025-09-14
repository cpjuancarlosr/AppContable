//== File: xml_parser.gs (normalización a hoja XML)

/**
 * Gets all existing UUIDs from the XML sheet to prevent duplicates.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh The 'XML' sheet.
 * @return {Set<string>} A Set containing all existing UUIDs.
 */
function obtenerUuidsExistentes_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return new Set();
  }
  const uuids = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  return new Set(uuids.filter(Boolean));
}

/**
 * Gets all existing RFCs from the Entidades sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh The 'Entidades' sheet.
 * @return {Set<string>} A Set containing all existing RFCs.
 */
function obtenerRfcEntidades_(sh) {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
        return new Set();
    }
    const rfcs = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    return new Set(rfcs.filter(Boolean));
}


/**
 * Parses a CFDI XML string and extracts relevant data into a structured object.
 * @param {string} xmlContent The XML content of the CFDI.
 * @return {object} An object containing the structured CFDI data.
 */
function parsearCFDI_(xmlContent) {
  const xml = XmlService.parse(xmlContent);
  const root = xml.getRootElement();
  const ns = root.getNamespace();
  const nsc = XmlService.getNamespace('cfdi', 'http://www.sat.gob.mx/cfd/4');
  const nst = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');

  const A = (el, attr) => el.getAttribute(attr)?.getValue() || '';

  const emisorEl = root.getChild('Emisor', nsc);
  const receptorEl = root.getChild('Receptor', nsc);
  const timbreEl = root.getChild('Complemento', nsc)?.getChild('TimbreFiscalDigital', nst);

  if (!timbreEl) {
    throw new Error('El XML no contiene un Timbre Fiscal Digital. No es un CFDI válido.');
  }

  const impuestosEl = root.getChild('Impuestos', nsc);
  let totalTraslados = { '0.160000': 0, '0.080000': 0, '0.000000': 0, exento: 0 };
  if (impuestosEl) {
      const trasladosEl = impuestosEl.getChild('Traslados', nsc);
      if (trasladosEl) {
          trasladosEl.getChildren('Traslado', nsc).forEach(t => {
              const tasa = A(t, 'TasaOCuota');
              const importe = parseFloat(A(t, 'Importe'));
              if (A(t, 'TipoFactor').toLowerCase() === 'tasa' && totalTraslados.hasOwnProperty(tasa)) {
                  totalTraslados[tasa] += importe;
              } else if (A(t, 'TipoFactor').toLowerCase() === 'exento') {
                  // This is an approximation. The base amount should be used.
                  totalTraslados.exento += parseFloat(A(t, 'Base'));
              }
          });
      }
  }


  const cfdi = {
    UUID: A(timbreEl, 'UUID'),
    Emisor: {
      Rfc: A(emisorEl, 'Rfc'),
      Nombre: A(emisorEl, 'Nombre')
    },
    Receptor: {
      Rfc: A(receptorEl, 'Rfc'),
      Nombre: A(receptorEl, 'Nombre')
    },
    Tipo: A(root, 'TipoDeComprobante'),
    FechaEmision: new Date(A(root, 'Fecha')),
    FechaTimbrado: new Date(A(timbreEl, 'FechaTimbrado')),
    MetodoPago: A(root, 'MetodoPago'),
    FormaPago: A(root, 'FormaPago'),
    Subtotal: parseFloat(A(root, 'SubTotal')),
    Total: parseFloat(A(root, 'Total')),
    UsoCFDI: A(receptorEl, 'UsoCFDI'),
    Moneda: A(root, 'Moneda'),
    Relacionados: root.getChildren('CfdiRelacionados', nsc).map(rel => A(rel, 'UUID')).join(', '),
    Impuestos: totalTraslados
  };

  cfdi.filaHoja = [
      cfdi.UUID, cfdi.Emisor.Rfc, cfdi.Emisor.Nombre, cfdi.Receptor.Rfc, cfdi.Receptor.Nombre,
      cfdi.Tipo, cfdi.FechaEmision, cfdi.FechaTimbrado, cfdi.MetodoPago, cfdi.FormaPago,
      cfdi.Subtotal, cfdi.Impuestos['0.160000'], cfdi.Impuestos['0.080000'], cfdi.Impuestos['0.000000'],
      cfdi.Impuestos.exento, cfdi.Total, cfdi.UsoCFDI, cfdi.Moneda, cfdi.Relacionados, ''
  ];

  return cfdi;
}
