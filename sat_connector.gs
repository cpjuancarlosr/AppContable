//== File: sat_connector.gs (FIEL, descarga SAT, control de sesión)

/**
 * Imports XML files from the Drive folder specified in the CFG sheet.
 * This is the fallback mechanism if direct SAT download is not possible.
 */
function importarDesdeCarpetaDrive() {
  const ui = SpreadsheetApp.getUi();
  try {
    const folderId = getCfg("ID_CARPETA_XML");
    if (!folderId) {
      throw new Error("El ID de la carpeta de XML no está configurado en la hoja 'CFG'.");
    }
    const folder = DriveApp.getFolderById(folderId);
    ui.alert(`Iniciando importación desde la carpeta de Drive: "${folder.getName()}". Esto puede tardar varios minutos.`);
    const stats = procesarCarpetaCFDI_(folder);
    ui.alert(`Importación completada.\n\nResultados:\n- ${stats.ok} CFDI procesados correctamente.\n- ${stats.dup} duplicados omitidos.\n- ${stats.err} errores encontrados.\n- ${stats.fuera} fuera del período de trabajo.`);
  } catch (e) {
    ui.alert(`Error durante la importación: ${e.message}`);
    log_(`Error en importarDesdeCarpetaDrive: ${e.stack}`);
  }
}


/**
 * Processes all XML files within a given Drive folder.
 * @param {GoogleAppsScript.Drive.Folder} folder The folder to process.
 * @return {object} Statistics about the import process.
 */
function procesarCarpetaCFDI_(folder) {
  const ss = SpreadsheetApp.getActive();
  const shXml = ss.getSheetByName("XML");
  const shEntidades = ss.getSheetByName("Entidades");

  const uuidsExistentes = obtenerUuidsExistentes_(shXml);
  const rfcEntidades = obtenerRfcEntidades_(shEntidades);

  const archivos = recolectarXMLs_(folder);
  let stats = { ok: 0, dup: 0, err: 0, fuera: 0 };
  const nuevasFilasXml = [];
  const nuevasEntidades = new Map();

  // TODO: Get a proper period from a cell in 'Inicio' sheet
  const periodo = {
      ini: new Date(new Date().getFullYear(), new Date().getMonth(), 1),
      fin: new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0)
  };


  archivos.forEach(archivo => {
    try {
      const xmlContent = archivo.getText('UTF-8');
      const cfdi = parsearCFDI_(xmlContent);

      if (cfdi.FechaEmision < periodo.ini || cfdi.FechaEmision > periodo.fin) {
        stats.fuera++;
        return;
      }

      if (uuidsExistentes.has(cfdi.UUID)) {
        stats.dup++;
        return;
      }

      nuevasFilasXml.push(cfdi.filaHoja);
      uuidsExistentes.add(cfdi.UUID);

      // Check and add new entities (clients/providers)
      [cfdi.Emisor, cfdi.Receptor].forEach(entidad => {
        if (entidad.Rfc && !rfcEntidades.has(entidad.Rfc) && !nuevasEntidades.has(entidad.Rfc)) {
            const tipo = (entidad.Rfc === cfdi.Emisor.Rfc) ? 'Proveedor' : 'Cliente';
            nuevasEntidades.set(entidad.Rfc, [entidad.Rfc, entidad.Nombre, tipo, '', '', '', 0.16]);
        }
      });

      stats.ok++;
    } catch (e) {
      stats.err++;
      log_(`Error procesando el archivo ${archivo.getName()}: ${e.message}`);
    }
  });

  if (nuevasFilasXml.length > 0) {
    shXml.getRange(shXml.getLastRow() + 1, 1, nuevasFilasXml.length, nuevasFilasXml[0].length).setValues(nuevasFilasXml);
  }

  if (nuevasEntidades.size > 0) {
      const filasEntidades = Array.from(nuevasEntidades.values());
      shEntidades.getRange(shEntidades.getLastRow() + 1, 1, filasEntidades.length, filasEntidades[0].length).setValues(filasEntidades);
  }

  return stats;
}


// Placeholder functions for menu items that will be implemented later
function configurarFiel_() {
  SpreadsheetApp.getUi().alert('Función "Configurar FIEL" no implementada aún.');
}
function descargarCFDI() {
  SpreadsheetApp.getUi().alert('Función "Descargar CFDI" no implementada aún.');
}
