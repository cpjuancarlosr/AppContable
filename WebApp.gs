/********************  UI WRAPPERS  ********************/
// Estas funciones sirven como una capa intermedia para devolver mensajes de texto simples a la UI del sidebar.

function importarCFDIIngresosUI() {
  importarCFDIIngresos();
  return "Importación de ingresos iniciada. Revise los logs para ver el detalle.";
}

function importarCFDIEgresosUI() {
  importarCFDIEgresos();
  return "Importación de egresos iniciada. Revise los logs para ver el detalle.";
}

function generarPolizasDesdeMovimientosUI() {
  generarPolizasDesdeMovimientos();
  return "Proceso de generación de pólizas finalizado.";
}

function recalcularEstadosUI() {
  recalcularEstados();
  return "Estados financieros recalculados con fórmulas.";
}

function conciliarBancariaAvanzadaUI() {
  const count = conciliarBancariaAvanzada();
  return `Conciliación automática finalizada. Se encontraron ${count} coincidencias.`;
}

function calcularIVA_MensualUI() {
  const resultado = calcularIVA_Mensual();
  return `Cálculo de IVA finalizado. IVA a pagar/favor: ${resultado.neto}`;
}

function calcularISR_MensualUI() {
  const resultado = calcularISR_Mensual();
  if (resultado.error) {
    return `Error: ${resultado.error}`;
  }
  return `Cálculo de ISR (${resultado.regimen}) finalizado. Impuesto: ${resultado.impuesto}`;
}

function exportarPDFsEstadosUI() {
  exportarPDFsEstados();
  return "PDFs exportados a su carpeta de Google Drive.";
}

function enviarPaqueteFiscalUI() {
  enviarPaqueteFiscal();
  return "Paquete fiscal enviado por correo.";
}


/********************  WEB APP & CONFIG API  ********************/
function doGet(){
  return HtmlService.createHtmlOutputFromFile('index');
}

function getPeriodosUI(){
  const res=[]; const d=new Date(); d.setDate(1);
  for(let i=0;i<24;i++){
    const y=d.getFullYear(); const m=String(d.getMonth()+1).padStart(2,'0');
    res.push(`${y}-${m}`);
    d.setMonth(d.getMonth()-1);
  }
  return res;
}

function apiGetConfig(clave){
  return getCfg(clave);
}
function apiSetConfig(clave, valor){
  setCfg_(clave, valor);
  return true;
}

function getBankCsvFiles() {
  const folderId = getCfg("CARPETA_BANCOS_ID"); // Assuming CARPETA_BANCOS_ID is stored in config
  if (!folderId) {
    // Let's try to find it if it's not in config
    try {
      const ss = SpreadsheetApp.getActive();
      const file = DriveApp.getFileById(ss.getId());
      const parent = file.getParents().next();
      const folderIt = parent.getFoldersByName("BANCOS");
      if (folderIt.hasNext()) {
        const folder = folderIt.next();
        setCfg_("CARPETA_BANCOS_ID", folder.getId());
        return getCsvsFromFolder(folder);
      } else {
        return []; // No folder found
      }
    } catch (e) {
      log("Error finding BANCOS folder: " + e.message);
      return [];
    }
  }
  const folder = DriveApp.getFolderById(folderId);
  return getCsvsFromFolder(folder);
}

function getCsvsFromFolder(folder) {
  const files = folder.getFilesByType(MimeType.CSV);
  const fileList = [];
  while (files.hasNext()) {
    const file = files.next();
    fileList.push(file.getName());
  }
  return fileList;
}
