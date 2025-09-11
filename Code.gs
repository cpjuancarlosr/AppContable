/**
 * SISTEMA CONTABLE MX — JC (Operativo ++ México) — FULL LIMPIO
 * Crea carpetas CFDI_EMITIDAS / CFDI_RECIBIDAS / PDFs / BANCOS automáticamente
 * y guarda sus IDs en la hoja Config. Menú seguro sin fallar en contextos sin UI.
 *
 * Requisitos opcionales: Tasks API (servicio avanzado), Drive API (para atajos), UrlFetch.
 */

/********************  PALETA Y ESTILO  ********************/
const PALETA = { baseBg: "#FFFFFF", baseText: "#000000", blue: "#7CA7EB", choc: "#402924", camel: "#CBE290", navy: "#000B26" };

/********************  CONFIGURACIÓN  ********************/
const CFG = {
  NOMBRE_SISTEMA: "📘 Consultoría Contable",
  MONEDA: "MXN",
  IVA_TASAS: { "16": 0.16, "8": 0.08, "0": 0.00 },
  ISR_PM_TASA: 0.30,
  COEF_UTILIDAD: 0.25,
  DIAS_TOLERANCIA_CONCILIACION: 4,
  LISTA_TASKS: "Contabilidad JC",
  CALENDARIO_NOMBRE: "Obligaciones Fiscales JC",
  CARPETA_PDFS_ID: "",
  CARPETA_CFDI_ING_ID: "",
  CARPETA_CFDI_EGR_ID: "",
  CORREOS_DESTINO: ""
};

// Cache para configuraciones; evita lecturas repetidas de la hoja Config
let CFG_CACHE = null;

/********************  MENÚ UI (SEGURO)  ********************/
const SHEETS_TO_IGNORE = ["Dashboard", "Config"];

function onOpen(e){
  if (!canUseUi_()) return;
  buildMenu_();
  hideAllSheets_();
}

/**
 * Oculta todas las hojas excepto las que están en la lista SHEETS_TO_IGNORE.
 */
function hideAllSheets_() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sh => {
    if (SHEETS_TO_IGNORE.indexOf(sh.getName()) === -1) {
      sh.hideSheet();
    }
  });
  mostrarDashboard(); // Asegura que el dashboard sea la hoja activa.
}

/**
 * Muestra una hoja específica por su nombre.
 * @param {string} sheetName - El nombre de la hoja a mostrar.
 */
function showSheet(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (sh) {
    sh.showSheet().activate();
  } else {
    safeAlert_(`La hoja "${sheetName}" no existe.`);
  }
}

function buildMenu_(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu(CFG.NOMBRE_SISTEMA);

  menu.addItem("🚀 Mostrar Dashboard", "mostrarDashboard");
  menu.addSeparator();

  const subMenuSetup = ui.createMenu("1. Configuración");
  subMenuSetup.addItem("⚙️ Instalar/Actualizar Sistema Completo", "setupSistemaMX");
  subMenuSetup.addItem("📂 Crear Carpetas en Drive", "crearCarpetasSistema");
  menu.addSubMenu(subMenuSetup);

  const subMenuImport = ui.createMenu("2. Importación");
  subMenuImport.addItem("⬇️ Importar CFDI Ingresos", "importarCFDIIngresos");
  subMenuImport.addItem("⬇️ Importar CFDI Egresos", "importarCFDIEgresos");
  menu.addSubMenu(subMenuImport);

  const subMenuProcess = ui.createMenu("3. Procesamiento");
  subMenuProcess.addItem("🧩 Generar Pólizas", "generarPolizasDesdeMovimientos");
  subMenuProcess.addItem("📐 Recalcular Estados + KPIs", "recalcularEstados");
  subMenuProcess.addItem("🏦 Conciliar Bancos", "conciliarBancariaAvanzada");
  menu.addSubMenu(subMenuProcess);

  const subMenuReports = ui.createMenu("4. Reportes y Consultas");
  subMenuReports.addItem("🧾 Calcular IVA Mensual", "calcularIVA_Mensual");
  subMenuReports.addItem("💸 Calcular ISR PM Mensual", "calcularISR_PM_Mensual");
  subMenuReports.addItem("🧩 Generar DIOT (CSV)", "generarDIOT_CSV");
  subMenuReports.addSeparator();
  subMenuReports.addItem("🖨️ Exportar Estados a PDF", "exportarPDFsEstados");
  subMenuReports.addItem("📧 Enviar Paquete Fiscal", "enviarPaqueteFiscal");
  menu.addSubMenu(subMenuReports);

  const subMenuView = ui.createMenu("5. Ver Hojas de Datos");
  subMenuView.addItem("Ver Ingresos", "showIngresos");
  subMenuView.addItem("Ver Egresos", "showEgresos");
  subMenuView.addItem("Ver Pólizas", "showPolizas");
  subMenuView.addItem("Ver Balanza", "showBalanza");
  subMenuView.addItem("Ver Estado de Resultados", "showER");
  subMenuView.addItem("Ver Balance General", "showBG");
  subMenuView.addSeparator();
  subMenuView.addItem("Ocultar Hojas", "hideAllSheets_");
  menu.addSubMenu(subMenuView);

  menu.addToUi();
}

// Funciones wrapper para los menús
function showIngresos() { showSheet("Ingresos"); }
function showEgresos() { showSheet("Egresos"); }
function showPolizas() { showSheet("Polizas"); }
function showBalanza() { showSheet("Balanza"); }
function showER() { showSheet("EstadoResultados"); }
function showBG() { showSheet("BalanceGeneral"); }

/**
 * Activa y muestra la hoja del Dashboard.
 */
function mostrarDashboard() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName("Dashboard");
  if (dash) {
    dash.activate();
  } else {
    safeAlert_("El Dashboard no se encuentra. Por favor, ejecute la instalación del sistema.");
  }
}
function canUseUi_(){ try{ SpreadsheetApp.getUi(); return true; }catch(e){ return false; } }
function safeAlert_(msg){
  try{ SpreadsheetApp.getUi().alert(msg); }
  catch(e){ try{ SpreadsheetApp.getActive().toast(msg); }catch(_){ Logger.log("[ALERTA] "+msg); } }
}

/********************  INSTALACIÓN  ********************/
function setupSistemaMX(){
  const ss = SpreadsheetApp.getActive();
  const hojas = [
    "Config","Clientes","Proveedores","CatCuentas","Ingresos","Egresos","Bancos",
    "Polizas","Mayor","Balanza","EstadoResultados","BalanceGeneral","KPIs",
    "Plantillas","PagosImpuestos","DIOT","Logs","Dashboard"
  ];
  hojas.forEach(h => getOrCreateSheet(ss,h));

  prepararConfig();
  prepararMaestros();
  prepararCatCuentas();
  prepararIngresosEgresos();
  prepararBancos();
  prepararPolizasMayor();
  prepararEstadosPlantillas();
  prepararPeriodos();
  aplicarTemaVisual();

  // Crear carpetas del sistema y guardar IDs en Config
  crearCarpetasSistema();

  safeAlert_("Sistema MX listo. Llena Config y usa el menú.");
}

/********************  CREACIÓN DE CARPETAS  ********************/
function crearCarpetasSistema(){
  const ss = SpreadsheetApp.getActive();
  const file = DriveApp.getFileById(ss.getId());
  let parent;
  try{ parent = file.getParents().next(); }catch(e){
    // Si no hay padre (poco común), crea uno
    parent = DriveApp.createFolder(`Contabilidad_${ss.getName()}`);
    parent.addFile(file); DriveApp.getRootFolder().removeFile(file);
  }

  const folCFDI_Emit = getOrCreateSub_(parent, "CFDI_EMITIDAS");
  const folCFDI_Rec  = getOrCreateSub_(parent, "CFDI_RECIBIDAS");
  const folPDFs      = getOrCreateSub_(parent, "PDFs");
  const folBancos    = getOrCreateSub_(parent, "BANCOS");

  setCfg_("CARPETA_CFDI_ING_ID", folCFDI_Emit.getId());
  setCfg_("CARPETA_CFDI_EGR_ID", folCFDI_Rec.getId());
  setCfg_("CARPETA_PDFS_ID", folPDFs.getId());

  log(`Carpetas listas → Emitidas:${folCFDI_Emit.getName()} · Recibidas:${folCFDI_Rec.getName()} · PDFs:${folPDFs.getName()} · Bancos:${folBancos.getName()}`);
  safeAlert_("Carpetas creadas y guardadas en Config.");
}
function getOrCreateSub_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext()? it.next(): parent.createFolder(name);
}

/********************  PREPARACIÓN DE HOJAS  ********************/
function prepararConfig(){
  const sh = SpreadsheetApp.getActive().getSheetByName("Config"); sh.clear();
  sh.getRange(1,1,1,3).setValues([["Clave","Valor","Notas"]]).setFontWeight("bold");
  const data = [
    ["CARPETA_PDFS_ID", CFG.CARPETA_PDFS_ID, "Destino PDF"],
    ["CARPETA_CFDI_ING_ID", CFG.CARPETA_CFDI_ING_ID, "Folder XML emitidos"],
    ["CARPETA_CFDI_EGR_ID", CFG.CARPETA_CFDI_EGR_ID, "Folder XML recibidos"],
    ["CORREOS_DESTINO", CFG.CORREOS_DESTINO, "Separados por coma"],
    ["LISTA_TASKS", CFG.LISTA_TASKS, "Lista de Tasks"],
    ["CALENDARIO_NOMBRE", CFG.CALENDARIO_NOMBRE, "Calendario de obligaciones"],
    ["IVA_16", CFG.IVA_TASAS["16"], "Tasa general"],
    ["IVA_8", CFG.IVA_TASAS["8"], "Frontera"],
    ["ISR_PM_TASA", CFG.ISR_PM_TASA, "Tasa PM"],
    ["COEF_UTILIDAD", CFG.COEF_UTILIDAD, "Coeficiente anual"],
    ["DIAS_TOLERANCIA_CONCILIACION", CFG.DIAS_TOLERANCIA_CONCILIACION, "± días"],
    ["MONEDA", CFG.MONEDA, ""],
    ["FECHA_CORTE", new Date(), "Mes de trabajo"],
    ["PERIODO_LABEL", "", "YYYY-MM o fecha"]
  ];
  sh.getRange(2,1,data.length,3).setValues(data); sh.autoResizeColumns(1,3);
}

function prepararMaestros(){
  const cli = SpreadsheetApp.getActive().getSheetByName("Clientes");
  cli.clear(); cli.getRange(1,1,1,6).setValues([["ID","Nombre","RFC","Email","Tipo","Cuenta Contable"]]).setFontWeight("bold");
  const prov = SpreadsheetApp.getActive().getSheetByName("Proveedores");
  prov.clear(); prov.getRange(1,1,1,8).setValues([["ID","Nombre","RFC","Email","Tipo","Retención ISR","Tasa IVA","Cuenta Contable"]]).setFontWeight("bold");
}

function prepararCatCuentas(){
  const sh = SpreadsheetApp.getActive().getSheetByName("CatCuentas"); sh.clear();
  sh.getRange(1,1,1,7).setValues([["Codigo","Nombre","Tipo","SAT","Nivel","Padre","Naturaleza"]]).setFontWeight("bold");
  const base = [
    ["100-000","Activo Circulante","Activo","",1,"","Deudora"],
    ["110-100","Bancos","Activo","",2,"100-000","Deudora"],
    ["110-110","Banco MXN","Activo","",3,"110-100","Deudora"],
    ["120-000","Clientes","Activo","",2,"100-000","Deudora"],
    ["130-000","Inventarios","Activo","",2,"100-000","Deudora"],
    ["200-000","Pasivo Corto Plazo","Pasivo","",1,"","Acreedora"],
    ["210-100","Proveedores","Pasivo","",2,"200-000","Acreedora"],
    ["240-200","IVA Trasladado 16%","Impuesto","",3,"200-000","Acreedora"],
    ["240-210","IVA Trasladado 8%","Impuesto","",3,"200-000","Acreedora"],
    ["240-300","IVA Acreditable 16%","Impuesto","",3,"200-000","Deudora"],
    ["240-310","IVA Acreditable 8%","Impuesto","",3,"200-000","Deudora"],
    ["240-400","ISR Retenido","Impuesto","",2,"200-000","Acreedora"],
    ["300-000","Capital","Capital","",1,"","Acreedora"],
    ["400-000","Ingresos","Resultado","",1,"","Acreedora"],
    ["500-000","Costo/Ventas","Resultado","",1,"","Deudora"],
    ["510-000","Gastos","Resultado","",1,"","Deudora"],
    ["520-000","Gastos Honorarios","Resultado","",2,"510-000","Deudora"]
  ];
  sh.getRange(2,1,base.length,7).setValues(base);
}

function prepararIngresosEgresos(){
  const head = ["Fecha","Folio","Cliente/Proveedor","RFC","Concepto","Tasa IVA","Subtotal","IVA","Retenciones","Total","FormaPago","Banco/Cuenta","EstadoPago","UUID","Cuenta Contable","PolizaID","Conciliado"];
  const shI = SpreadsheetApp.getActive().getSheetByName("Ingresos"); shI.clear(); shI.getRange(1,1,1,head.length).setValues([head]).setFontWeight("bold");
  const shE = SpreadsheetApp.getActive().getSheetByName("Egresos"); shE.clear(); shE.getRange(1,1,1,head.length).setValues([head]).setFontWeight("bold");
}

function prepararBancos(){
  const sh = SpreadsheetApp.getActive().getSheetByName("Bancos"); sh.clear();
  sh.getRange(1,1,1,11).setValues([["Fecha","Descripcion","Cargo","Abono","Importe","Referencia","Banco","Cuenta","FolioFactura","UUID","Conciliado"]]).setFontWeight("bold");
}

function prepararPolizasMayor(){
  const p = SpreadsheetApp.getActive().getSheetByName("Polizas"); p.clear();
  p.getRange(1,1,1,12).setValues([["Fecha","Tipo","Ref","Cuenta","Nombre Cuenta","Descripcion","Debe","Haber","UUID","Origen","CentroCosto","Proyecto"]]).setFontWeight("bold");
  const m = SpreadsheetApp.getActive().getSheetByName("Mayor"); m.clear(); m.getRange(1,1,1,7).setValues([["Cuenta","Fecha","Ref","Descripcion","Debe","Haber","Saldo"]]).setFontWeight("bold");
}

function prepararEstadosPlantillas(){
  const bal = SpreadsheetApp.getActive().getSheetByName("Balanza"); bal.clear(); bal.getRange(1,1,1,5).setValues([["Cuenta","Nombre","Debe","Haber","Saldo"]]).setFontWeight("bold");
  const er = SpreadsheetApp.getActive().getSheetByName("EstadoResultados"); er.clear(); er.getRange(1,1,1,3).setValues([["Rubro","Cuenta","Importe"]]).setFontWeight("bold");
  const bg = SpreadsheetApp.getActive().getSheetByName("BalanceGeneral"); bg.clear(); bg.getRange(1,1,1,3).setValues([["Rubro","Cuenta","Importe"]]).setFontWeight("bold");
  const kpi = SpreadsheetApp.getActive().getSheetByName("KPIs"); kpi.clear(); kpi.getRange(1,1,1,2).setValues([["Indicador","Valor"]]).setFontWeight("bold");
  const pl = SpreadsheetApp.getActive().getSheetByName("Plantillas"); pl.clear(); pl.getRange(1,1,1,3).setValues([["Codigo","Asunto","HTML"]]).setFontWeight("bold");
  pl.getRange(2,1,2,3).setValues([
    ["AVISO_IMPUESTOS","Aviso fiscal {{periodo}}","<div style='font-family:Rubik'><h2 style='color:"+PALETA.navy+"'>Obligaciones — {{periodo}}</h2><p>IVA neto: <b>{{iva_neto}}</b> · ISR estimado: <b>{{isr_estimado}}</b></p><p>Fecha límite: <b>{{fecha_limite}}</b></p></div>"],
    ["ENTREGA_ESTADOS","Estados financieros {{periodo}} — {{empresa}}","<div style='font-family:Rubik'><h2 style='color:"+PALETA.choc+"'>Estados {{periodo}}</h2><ul><li>ER</li><li>BG</li><li>Balanza</li></ul><p>KPI principal: <b>{{kpi_principal}}</b></p></div>"]
  ]);
  const pagos = SpreadsheetApp.getActive().getSheetByName("PagosImpuestos"); pagos.clear(); pagos.getRange(1,1,1,7).setValues([["Periodo","Impuesto","Base","Tasa","Importe","Fecha Pago","Acuse URL"]]).setFontWeight("bold");
  const diot = SpreadsheetApp.getActive().getSheetByName("DIOT"); diot.clear(); diot.getRange(1,1,1,8).setValues([["RFC","Proveedor","Tipo","Base 16%","IVA 16%","Base 8%","IVA 8%","Exento"]]).setFontWeight("bold");
}

function prepararPeriodos(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Periodos") || ss.insertSheet("Periodos");
  sh.clear();
  sh.getRange(1,1,1,2).setValues([["Periodo","Estado"]]).setFontWeight("bold");
  const now = new Date();
  const rows=[];
  for (let i=-12;i<=12;i++){
    const d=new Date(now.getFullYear(), now.getMonth()+i,1);
    const label=`${d.getFullYear()}-${("0"+(d.getMonth()+1)).slice(-2)}`;
    rows.push([label,"Abierto"]);
  }
  if (rows.length) sh.getRange(2,1,rows.length,2).setValues(rows);
  const cfg = ss.getSheetByName("Config");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getRange(2,1,rows.length,1), true).build();
  const row = getCfgRow_("PERIODO_LABEL");
  if (row>0) cfg.getRange(row,2).setDataValidation(rule);
}

/********************  TEMA VISUAL  ********************/
function aplicarTemaVisual(){
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach((sh,i)=>{
    const head = sh.getRange(1,1,1,Math.min(12, sh.getMaxColumns()));
    head.setBackgrounds([new Array(head.getNumColumns()).fill(PALETA.baseBg)])
        .setFontColor(PALETA.baseText).setFontWeight("bold")
        .setBorder(true,true,true,true,false,false,PALETA.navy,SpreadsheetApp.BorderStyle.SOLID_THICK);
    try{ sh.setTabColor([PALETA.blue,PALETA.navy,PALETA.camel,PALETA.choc][i%4]); }catch(e){}
  });

  // Rediseño del Dashboard
  const dash = ss.getSheetByName("Dashboard");
  dash.clear();
  dash.getRange("A1:F10").clear(); // Limpia un área más grande

  // Título y Periodo
  dash.getRange("A1").setValue(CFG.NOMBRE_SISTEMA).setFontSize(18).setFontWeight("bold").setFontColor(PALETA.navy);
  dash.getRange("A2").setValue("Periodo Activo:");
  dash.getRange("B2").setFormula("=Config!B14"); // Vincula al periodo en Config

  // Sección de KPIs
  dash.getRange("D1").setValue("KPIs Financieros").setFontWeight("bold");
  dash.getRange("D2:E5").setValues([
    ["Margen Neto", "=KPIs!B2"],
    ["Liquidez", "=KPIs!B3"],
    ["Caja", "=KPIs!B7"],
    ["Runway (meses)", "=KPIs!B9"]
  ]);

  // Secciones de Acciones (placeholders, los botones se agregarán después)
  dash.getRange("A4").setValue("1. Configuración e Importación").setFontWeight("bold");
  dash.getRange("A5").setValue("· Instalar / Actualizar Sistema");
  dash.getRange("A6").setValue("· Importar CFDI Ingresos");
  dash.getRange("A7").setValue("· Importar CFDI Egresos");

  dash.getRange("C4").setValue("2. Procesamiento").setFontWeight("bold");
  dash.getRange("C5").setValue("· Generar Pólizas");
  dash.getRange("C6").setValue("· Recalcular Estados Financieros");

  dash.getRange("E4").setValue("3. Reportes y Consultas").setFontWeight("bold");
  dash.getRange("E5").setValue("· Ver Balanza de Comprobación");
  dash.getRange("E6").setValue("· Ver Estado de Resultados");
  dash.getRange("E7").setValue("· Ver Balance General");

  // Formato
  dash.getRange("A1:F1").merge().setHorizontalAlignment("center");
  dash.getRange("A4:B4, C4:D4, E4:F4").merge();
  dash.getRange("A1:F10").setFontFamily("Arial");
  dash.getRange("A4,C4,E4,D1").setFontColor(PALETA.choc);
  dash.getRange("D2:E5").setBorder(true, true, true, true, true, true);
  dash.autoResizeColumns(1, 6);
}

/********************  IMPORTACIÓN DE CFDI (XML)  ********************/
/**
 * Funciones wrapper para iniciar la importación de CFDI de ingresos o egresos.
 * Toman el ID de la carpeta de Google Drive desde la hoja de Configuración.
 */
function importarCFDIIngresos(){ importarCFDI_(getCfg("CARPETA_CFDI_ING_ID"), true); }
function importarCFDIEgresos(){ importarCFDI_(getCfg("CARPETA_CFDI_EGR_ID"), false); }

/**
 * Lógica principal de importación de CFDI.
 * Procesa todos los archivos XML de una carpeta de Drive, extrae los datos,
 * evita duplicados y los inserta en las hojas 'Ingresos' o 'Egresos'.
 * @param {string} folderId - El ID de la carpeta de Google Drive que contiene los XML.
 * @param {boolean} esIngreso - True si son CFDI de ingresos, false si son de egresos.
 */
function importarCFDI_(folderId, esIngreso){
  if (!folderId) throw new Error("Config: CARPETA_CFDI_* no definida");

  // Obtiene el periodo de trabajo actual para filtrar facturas
  const pr = periodo_();
  const fol = DriveApp.getFolderById(folderId);

  // Recolecta todos los archivos XML, incluso dentro de zips y subcarpetas.
  const fuentes = recolectarXMLs_(fol);
  // Obtiene un mapa de UUIDs ya existentes para evitar duplicados.
  const ya = uuidsExistentes_();
  const rowsI = [], rowsE = [];
  let totFuentes=0, ok=0, dup=0, err=0, fuera=0, ruteoI=0, ruteoE=0;

  fuentes.forEach(src => {
    totFuentes++;
    try{
      let xml;
      try{ xml = XmlService.parse(cleanXmlText_(src.getText('UTF-8'))); }
      catch(e1){ xml = XmlService.parse(cleanXmlText_(src.getText())); }
      // Parsea el contenido del archivo XML.
      const root = xml.getRootElement();
      const A = attrMap_(root); // Convierte atributos del nodo raíz a un objeto.

      // Extrae datos primarios del CFDI. Si la fecha está fuera del periodo, lo ignora.
      const Fecha = parseDate_(A.Fecha||A.fecha); if(!(Fecha>=pr.ini && Fecha<=pr.fin)){ fuera++; return; }
      const Folio = A.Folio||A.folio||""; const Serie=A.Serie||A.serie||"";
      const Subtotal = toNum(A.SubTotal||A.Subtotal); const Total = toNum(A.Total);
      const Tipo = (A.TipoDeComprobante||A.Tipo||"").toString().toUpperCase();
      const Moneda = A.Moneda||"MXN"; const Metodo=A.MetodoPago||""; const Forma=A.FormaPago||"";

      // Recorre los nodos descendientes para encontrar datos clave como Emisor, Receptor y Timbre Fiscal (UUID).
      let UUID="", Emisor={}, Receptor={}, uso=""; let rels=[];
      const desc=root.getDescendants();
      for (let n of desc){ const el=n.asElement&&n.asElement(); if(!el) continue; const nm=el.getName();
        if(/^Emisor$/i.test(nm)) Emisor=attrMap_(el);
        else if(/^Receptor$/i.test(nm)) { Receptor=attrMap_(el); uso = Receptor.UsoCFDI || uso; }
        else if(/TimbreFiscalDigital/i.test(nm)) { const a=attrMap_(el); if(a.UUID) UUID=a.UUID; } // El UUID es el identificador único de la factura.
        else if(/CfdiRelacionado/i.test(nm)) { const a=attrMap_(el); if(a.UUID) rels.push(a.UUID); }
        else if(/DoctoRelacionado/i.test(nm)) { const a=attrMap_(el); if(a.IdDocumento) rels.push(a.IdDocumento); }
      }

      const key=(UUID||"").trim();
      if(key && ya[key]){
        dup++;
        log(`UUID duplicado omitido: ${key}`);
        return;
      }
      const tasa = detectarTasaDesdeImpuestos_(root);
      const ivaCalc = round2(Total - Subtotal);
      const origen = esIngreso? 'Emitidas':'Recibidas';
      const rowMes=[origen, Tipo, Fecha, Serie, Folio, UUID, Emisor.Rfc||'', Emisor.Nombre||'', Receptor.Rfc||'', Receptor.Nombre||'', Moneda, Subtotal, ivaCalc, Total, Metodo, Forma, uso, rels.join('|'), '', ''];
      registrarEnHojaMes_(pr.label, rowMes);

      upsertMaestrosDesdeCFDI_(origen, Emisor, Receptor);

      const tercero = esIngreso? (Receptor.Nombre||Receptor.Rfc) : (Emisor.Nombre||Emisor.Rfc);
      const rfc     = esIngreso? (Receptor.Rfc||"")             : (Emisor.Rfc||"");
      const concepto = `CFDI ${src.name || ''} [Tipo:${Tipo||'?'}]`;
      const rowIE = [Fecha, Folio, tercero, rfc, concepto, tasaLabel_(tasa), Subtotal, ivaCalc, "", Total, "Transferencia","Banco MXN","Pendiente", UUID, "", "", "No"];
      if (esIngreso){ rowsI.push(rowIE); ruteoI++; } else { rowsE.push(rowIE); ruteoE++; }
      if(key) ya[key]=true; ok++;
    }catch(e){ err++; log(`Error XML fuente ${src.name||"(sin nombre)"}: ${e}`); }
  });

  const ss = SpreadsheetApp.getActive();
  if (rowsI.length){ const shI = ss.getSheetByName("Ingresos"); shI.getRange(shI.getLastRow()+1,1,rowsI.length,rowsI[0].length).setValues(rowsI); }
  if (rowsE.length){ const shE = ss.getSheetByName("Egresos"); shE.getRange(shE.getLastRow()+1,1,rowsE.length,rowsE[0].length).setValues(rowsE); }

  vincularDocumentosPeriodo_(pr.label);
  aplicarPagosDesdeHojaMes_(pr.label);
  resaltarPendientesMaestros_();

  log(`Importación CFDI — fuentes:${totFuentes} ok:${ok} dup:${dup} err:${err} fuera:${fuera} → I:${ruteoI} E:${ruteoE}`);
}

/**
 * Recorre una carpeta de Drive y sus subcarpetas de forma recursiva.
 * Extrae todos los archivos .xml, incluso si están dentro de archivos .zip.
 * @param {Folder} folder - El objeto Folder de DriveApp a procesar.
 * @returns {Array<Object>} Un arreglo de objetos, donde cada objeto representa un archivo XML y tiene las propiedades {name, getText}.
 */
function recolectarXMLs_(folder){
  const out = [];
  const MAX_ZIP_MB = 50; // Límite para evitar procesar archivos ZIP demasiado grandes.
  const crawl = (fol) => {
    const files = fol.getFiles();
    while (files.hasNext()){
      let file = files.next();
      try{ file = resolveShortcutFile_(file); }catch(e){}
      const mt = (file.getMimeType()||"").toLowerCase();
      const name = file.getName();
      if (mt.indexOf('zip')!==-1 || /\.zip$/i.test(name)){
        try{
          const sizeMB = (file.getSize()||0)/(1024*1024);
          if (sizeMB>MAX_ZIP_MB){ log(`ZIP omitido por tamaño (${sizeMB.toFixed(1)} MB): ${name}`); continue; }
          const blobs = Utilities.unzip(file.getBlob());
          blobs.forEach(b=>{ if(/\.xml$/i.test(b.getName())) out.push({name: b.getName(), getText: (enc)=> b.getDataAsString(enc||'UTF-8')}); });
        }catch(e){ log(`ZIP con error ${name}: ${e}`); }
      } else if (mt.indexOf('xml')!==-1 || /\.xml$/i.test(name)){
        out.push({name, getText: (enc)=> file.getBlob().getDataAsString(enc||'UTF-8')});
      }
    }
    const subs = fol.getFolders();
    while (subs.hasNext()) crawl(subs.next());
  };
  crawl(folder);
  return out;
}

/**
 * Si el archivo es un atajo de Google Drive, devuelve el archivo real al que apunta.
 * Requiere que el servicio avanzado de Drive API esté habilitado en el proyecto.
 * @param {File} file - El objeto File de DriveApp.
 * @returns {File} El archivo original o el archivo apuntado por el atajo.
 */
function resolveShortcutFile_(file){
  const MIME_SHORTCUT = 'application/vnd.google-apps.shortcut';
  try{
    if (file.getMimeType() !== MIME_SHORTCUT) return file;
    const meta = Drive.Files.get(file.getId());
    const targetId = meta.shortcutDetails && meta.shortcutDetails.targetId;
    if (targetId){ return DriveApp.getFileById(targetId); }
  }catch(e){ /* si no hay servicio avanzado, seguimos con file */ }
  return file;
}

/**
 * Lee las hojas de Ingresos y Egresos para crear un conjunto (Set) de todos los UUIDs
 * que ya han sido importados. Esto es crucial para evitar duplicar facturas.
 * @returns {Object} Un objeto que funciona como un Set, con los UUIDs como claves.
 */
function uuidsExistentes_(){
  const ss = SpreadsheetApp.getActive();
  const set = {};
  ["Ingresos","Egresos"].forEach(n=>{
    const sh = ss.getSheetByName(n); const lr = sh.getLastRow();
    if (lr>1){ sh.getRange(2,14,lr-1,1).getValues().forEach(r=>{ const u = (r[0]||"").toString().trim(); if(u) set[u]=true; }); }
  });
  return set;
}

/**
 * Intenta detectar la tasa de IVA (16% u 8%) analizando los nodos de impuestos del XML.
 * Busca un impuesto de tipo "Traslado" con código "002" (IVA).
 * @param {Element} root - El elemento raíz del documento XML.
 * @returns {number} La tasa de IVA detectada (0.16, 0.08, o 0.00).
 */
function detectarTasaDesdeImpuestos_(root){
  const d = root.getDescendants();
  for (let n of d){ try{ const el=n.asElement(); if(!el) continue; if(/Traslado/i.test(el.getName())){ const a=attrMap_(el); if(a.Impuesto=="002"&&a.TasaOCuota){ const t=parseFloat(a.TasaOCuota); return t>=0.15?0.16:(t>=0.07?0.08:0); } } }catch(e){} }
  return 0.16; // Devuelve 16% por defecto si no se encuentra.
}

/********************  MOTOR DE PÓLIZAS (MX)  ********************/
/**
 * Orquesta la generación de pólizas contables.
 * Lee los movimientos de las hojas 'Ingresos' y 'Egresos' y llama a la función
 * que procesa cada una. Finalmente, recalcula los estados financieros.
 */
function generarPolizasDesdeMovimientos(){
  const ss=SpreadsheetApp.getActive();
  const shI=ss.getSheetByName("Ingresos");
  const shE=ss.getSheetByName("Egresos");
  const p=ss.getSheetByName("Polizas");

  // Limpia la hoja de Pólizas (excepto el encabezado) para evitar duplicados.
  const lr = p.getLastRow();
  if (lr > 1) {
    p.getRange(2, 1, lr - 1, p.getLastColumn()).clearContent();
  }

  log("Hoja de Pólizas limpiada. Generando nuevas pólizas...");

  if (shI.getLastRow()>1) procesarMovimientosMX(shI,p,true);
  if (shE.getLastRow()>1) procesarMovimientosMX(shE,p,false);
  recalcularEstados();
  safeAlert_("Pólizas generadas y estados financieros actualizados.");
}

/**
 * Procesa una hoja de movimientos (Ingresos o Egresos) y genera los asientos
 * contables correspondientes en la hoja 'Polizas'.
 * @param {Sheet} shMov - La hoja de Ingresos o Egresos.
 * @param {Sheet} shPol - La hoja de Pólizas donde se escribirán los asientos.
 * @param {boolean} esIngreso - True si se está procesando la hoja de Ingresos.
 */
function procesarMovimientosMX(shMov, shPol, esIngreso){
  const values = shMov.getRange(2,1,Math.max(0, shMov.getLastRow()-1), shMov.getLastColumn()).getValues();
  const out=[]; // Arreglo para almacenar las filas de las nuevas pólizas.
  values.forEach(r=>{
    const [fecha, folio, tercero, rfc, concepto, tasaLabel, subtotal0, iva0, ret0, total0, , , , uuid] = r;
    if(!fecha||!total0) return;
    const tipo = tipoDesdeConcepto_(concepto);
    if (tipo==="P") return;
    const factor = (tipo==="E"?-1:1);
    const subtotal = Number(subtotal0||0)*factor;
    const iva      = Number(iva0||0)*factor;
    const ret      = Number(ret0||0)*factor;
    const total    = Number(total0||0)*factor;
    const tasa = (String(tasaLabel)==="8"?0.08:(String(tasaLabel)==="0"?0:0.16));
    const ref = `${esIngreso?"I":"E"}-${folio||uuid||Utilities.getUuid().slice(0,8)}`;

    // Asignación de cuentas contables según el catálogo predefinido.
    const ctaTercero = esIngreso? "120-000" : "210-100"; // Clientes o Proveedores
    const ctaIngreso = "400-000"; // Cuenta de Ingresos
    const ctaGasto   = (/honorario/i.test(String(concepto))?"520-000":"510-000"); // Gasto general o por honorarios
    const ctaIvaTras = (tasa===0.08?"240-210":"240-200"); // IVA Trasladado (cobrado)
    const ctaIvaAcred= (tasa===0.08?"240-310":"240-300"); // IVA Acreditable (pagado)

    // Genera los asientos de la póliza según si es ingreso o egreso.
    // Genera los asientos de la póliza según si es ingreso o egreso.
    if (esIngreso){
      // Asiento de Ingreso:
      // Cargo (Debe) a Clientes por el total.
      // Abono (Haber) a Ingresos por el subtotal.
      // Abono (Haber) a IVA Trasladado por el IVA.
      // Cargo (Debe) a ISR Retenido si aplica (es un activo).
      out.push([fecha,"Ingreso",ref,ctaTercero,"Clientes",`Cobro ${concepto}`,total,0,uuid,"Ingresos","",""]); // DEBE
      out.push([fecha,"Ingreso",ref,ctaIngreso,"Ingresos",`Venta ${concepto}`,0,subtotal||0,uuid,"Ingresos","",""]); // HABER
      if (tasa>0) out.push([fecha,"Ingreso",ref,ctaIvaTras,"IVA Trasladado",`IVA ${tasa*100}%`,0,iva||tasa*(subtotal||0),uuid,"Ingresos","",""]); // HABER
      if (ret>0) out.push([fecha,"Ingreso",ref,"240-400","ISR Retenido","Retención ISR",ret,0,uuid,"Ingresos","",""]); // DEBE
    } else {
      // Asiento de Egreso:
      // Cargo (Debe) a Gastos por el subtotal.
      // Cargo (Debe) a IVA Acreditable por el IVA.
      // Abono (Haber) a Proveedores por el total.
      // Abono (Haber) a Retenciones si aplican.
      const prov = buscarProveedorPorRFC_(rfc); const retCfg = calcRetenciones_(prov, Math.abs(subtotal)||0, tasa);
      const isrR = ret>0? Number(ret): retCfg.isr; const ivaR = retCfg.iva;
      const ivaMonto = iva||tasa*(subtotal||0);
      out.push([fecha,"Egreso",ref,ctaGasto,"Gasto",`Gasto ${concepto}`,Math.abs(subtotal)||0,0,uuid,"Egresos","",""]); // DEBE
      if (tasa>0) out.push([fecha,"Egreso",ref,ctaIvaAcred,"IVA Acreditable",`IVA ${tasa*100}%`,Math.abs(ivaMonto),0,uuid,"Egresos","",""]); // DEBE
      if (isrR>0) out.push([fecha,"Egreso",ref,"240-400","ISR Retenido","Ret ISR a PF",0,isrR,uuid,"Egresos","",""]); // HABER
      if (ivaR>0) out.push([fecha,"Egreso",ref,"240-300","IVA Retenido","Ret IVA a PF",0,ivaR,uuid,"Egresos","",""]); // HABER
      out.push([fecha,"Egreso",ref,ctaTercero,"Proveedores",`Pago a ${tercero}`,0,Math.abs(total),uuid,"Egresos","",""]); // HABER
    }
  });
  if (out.length) shPol.getRange(shPol.getLastRow()+1,1,out.length,12).setValues(out);
}
function tipoDesdeConcepto_(concepto){ const m=/\[Tipo:([IEPN])\]/i.exec(String(concepto)||""); return m? m[1].toUpperCase(): ""; }

function buscarProveedorPorRFC_(rfc){
  const sh=SpreadsheetApp.getActive().getSheetByName("Proveedores");
  const vals=sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),8).getValues();
  return vals.find(r=> r[2]===rfc) || null;
}

function calcRetenciones_(prov, subtotal, tasa){
  if (!prov) return {isr:0, iva:0};
  const tipo = String(prov[4]||"").toUpperCase();
  if (tipo.indexOf("HONOR")>=0 || tipo.indexOf("ARREN")>=0){
    const isr = 0.10*subtotal; const iva = (2/3)*(tasa*subtotal);
    return {isr:round2(isr), iva:round2(iva)};
  }
  return {isr:0, iva:0};
}

/********************  MAYOR, BALANZA, ESTADOS, KPIs  ********************/
/**
 * Orquesta el recálculo de todos los reportes financieros.
 * Limpia las hojas de reportes y las vuelve a generar desde la hoja 'Polizas'.
 * 1. Genera el Libro Mayor.
 * 2. Genera la Balanza de Comprobación.
 * 3. Genera el Estado de Resultados.
 * 4. Genera el Balance General.
 * 5. Calcula los KPIs financieros.
 */
function recalcularEstados(){
  const ss=SpreadsheetApp.getActive(); const p=ss.getSheetByName("Polizas"); const mayor=ss.getSheetByName("Mayor"); const bal=ss.getSheetByName("Balanza"); const er=ss.getSheetByName("EstadoResultados"); const bg=ss.getSheetByName("BalanceGeneral"); const kpi=ss.getSheetByName("KPIs"); const cat=ss.getSheetByName("CatCuentas");

  // 1. Generar Libro Mayor y calcular saldos por cuenta
  mayor.getRange(2,1,mayor.getLastRow(), mayor.getLastColumn()).clearContent();
  const pols=p.getRange(2,1,Math.max(0,p.getLastRow()-1),12).getValues();
  const mrows=[]; const saldos={}; // 'saldos' acumulará el total de debe/haber por cuenta.
  pols.forEach(r=>{ const [fecha,,ref,cta,,desc,debe,haber]=[r[0],r[1],r[2],r[3],r[4],r[5],Number(r[6]||0),Number(r[7]||0)]; if(!cta) return; if(!saldos[cta]) saldos[cta]={debe:0,haber:0}; saldos[cta].debe+=debe; saldos[cta].haber+=haber; const saldo=saldos[cta].debe - saldos[cta].haber; mrows.push([cta,fecha,ref,desc,debe,haber,saldo]); });
  if (mrows.length) mayor.getRange(2,1,mrows.length,7).setValues(mrows);

  // 2. Generar Balanza de Comprobación
  bal.getRange(2,1,bal.getLastRow(), bal.getLastColumn()).clearContent();
  const nombres = mapearNombreCuenta_(cat); // Mapea códigos de cuenta a nombres.
  const brows = Object.keys(saldos).map(cta=>[cta, nombres[cta]||"", saldos[cta].debe, saldos[cta].haber, saldos[cta].debe - saldos[cta].haber]);
  if (brows.length) bal.getRange(2,1,brows.length,5).setValues(brows);

  // 3. Generar Estado de Resultados
  er.getRange(2,1,er.getLastRow(), er.getLastColumn()).clearContent();
  bg.getRange(2,1,bg.getLastRow(), bg.getLastColumn()).clearContent();
  const sumPref=(pref,neg)=> brows.filter(r=> String(r[0]).startsWith(pref)).reduce((a,b)=> a + Number(b[4]||0), 0)*(neg?-1:1);
  const ingresos = sumPref("400-", true)*-1; // Ingresos (naturaleza acreedora, se multiplica por -1 para mostrar positivo)
  const costos   = Math.abs(sumPref("500-", false));
  const gastos   = Math.abs(sumPref("510-", false));
  const utilidad = ingresos - costos - gastos;
  er.getRange(2,1,4,3).setValues([["Ingresos","400-***", ingresos],["Costos","500-***", -costos],["Gastos","510-***", -gastos],["Utilidad","—", utilidad]]);

  // 4. Generar Balance General
  const activo  = brows.filter(r=> String(r[0]).startsWith("1")).reduce((a,b)=> a+Number(b[4]||0),0);
  const pasivo  = brows.filter(r=> String(r[0]).startsWith("2")).reduce((a,b)=> a+Number(b[4]||0),0);
  const capital = brows.filter(r=> String(r[0]).startsWith("3")).reduce((a,b)=> a+Number(b[4]||0),0) + utilidad; // El capital incluye la utilidad del ejercicio.
  bg.getRange(2,1,3,3).setValues([["Activo","1xx", activo],["Pasivo","2xx", pasivo],["Capital","3xx+U", capital]]);

  // 5. Calcular KPIs
  kpi.getRange(2,1,kpi.getLastRow(),2).clearContent();
  const caja = saldoPorCuenta_(brows,"110-110"); // Saldo de la cuenta de caja/bancos.
  const liquidez = pasivo? caja/pasivo : 0;
  const acida = (caja + saldoPorCuenta_(brows,"120-000"))/Math.max(1,pasivo); // Prueba ácida = (Activo Corriente - Inventario) / Pasivo Corriente
  const margen = ingresos? utilidad/ingresos:0; // Margen de utilidad neta
  const roa = activo? utilidad/activo:0; // Retorno sobre activos
  const roe = capital? utilidad/capital:0; // Retorno sobre capital
  const gastosMens = gastos;
  const runway = gastosMens? caja/gastosMens:0; // Meses de "supervivencia" si no hay ingresos.
  const ks = [["Margen Neto", round2(margen)],["Liquidez", round2(liquidez)],["Prueba Ácida", round2(acida)],["ROA", round2(roa)],["ROE", round2(roe)],["Caja", round2(caja)],["Gastros Mensuales", round2(gastosMens)],["Runway (meses)", round2(runway)]];
  kpi.getRange(2,1,ks.length,2).setValues(ks);
}

function mapearNombreCuenta_(cat){ const v=cat.getRange(2,1,Math.max(0,cat.getLastRow()-1),2).getValues(); const m={}; v.forEach(a=>m[a[0]]=a[1]); return m; }
function saldoPorCuenta_(rows, cta){ const r=rows.find(x=> x[0]===cta); return r? Number(r[4]||0):0; }

/********************  CONCILIACIÓN AVANZADA  ********************/
function conciliarBancariaAvanzada(){
  const ss=SpreadsheetApp.getActive(); const shB=ss.getSheetByName("Bancos"); const shI=ss.getSheetByName("Ingresos"); const shE=ss.getSheetByName("Egresos"); const tol=Number(getCfg("DIAS_TOLERANCIA_CONCILIACION"))||CFG.DIAS_TOLERANCIA_CONCILIACION;
  const bank=shB.getRange(2,1,Math.max(0,shB.getLastRow()-1), shB.getLastColumn()).getValues();
  const idxI=indexarMovs_(shI); const idxE=indexarMovs_(shE); let c=0;
  bank.forEach((r,i)=>{
    const [fecha,,cargo,abono,importe,ref] = r;
    const monto=Number(importe||cargo||abono||0); if(!fecha||!monto) return;
    const target=(Number(cargo||0)>0)? idxE: idxI;
    const hit=buscarMatch_(target, fecha, monto, ref, tol);
    if(hit){
      shB.getRange(i+2,9).setValue(hit.folio||"");
      shB.getRange(i+2,10).setValue(hit.uuid||"");
      shB.getRange(i+2,11).setValue("Sí");
      hit.sh.getRange(hit.row,17).setValue("Sí");
      c++;
    }
  });
  log(`Conciliación avanzada: ${c} match.`);
}

function indexarMovs_(sh){
  const vals=sh.getRange(2,1,Math.max(0,sh.getLastRow()-1), sh.getLastColumn()).getValues();
  const arr=[];
  vals.forEach((r,i)=>{
    const fecha=r[0], folio=r[1], total=r[9], uuid=r[13];
    if(!fecha||!total) return;
    arr.push({fecha:new Date(fecha), monto:Number(total), folio, uuid, row:i+2, sh});
  });
  return arr;
}
function buscarMatch_(idx, fecha, monto, referencia, tolDias){
  const f=new Date(fecha); const ref=(referencia||"").toString();
  for(let it of idx){
    const okF=Math.abs(Math.round((f-it.fecha)/(1000*60*60*24)))<=tolDias;
    const okM=Math.abs(it.monto-Number(monto))<0.5;
    const okR=(ref && (ref.indexOf(it.uuid)>=0 || ref.indexOf(it.folio)>=0));
    if((okF&&okM)||okR) return it;
  }
  return null;
}

/********************  IVA, ISR PM, DIOT  ********************/
function periodo_(){
  const raw = getCfg("PERIODO_LABEL");
  let y, m;
  if (raw) {
    if (Object.prototype.toString.call(raw) === '[object Date]') {
      const d = new Date(raw); y = d.getFullYear(); m = d.getMonth()+1;
    } else {
      const txt = String(raw).trim(); let mm, yy;
      let a = txt.match(/^(\d{4})[-\/](\d{1,2})$/);
      let b = txt.match(/^(\d{1,2})[-\/](\d{4})$/);
      if (a) { yy = +a[1]; mm = +a[2]; }
      else if (b) { yy = +b[2]; mm = +b[1]; }
      else { const f=getCfg("FECHA_CORTE")||new Date(); const d=new Date(f); yy=d.getFullYear(); mm=d.getMonth()+1; }
      y = yy; m = Math.min(12, Math.max(1, +mm||1));
    }
  } else {
    const f=getCfg("FECHA_CORTE")||new Date(); const d=new Date(f); y=d.getFullYear(); m=d.getMonth()+1;
  }
  const label = `${y}-${("0"+m).slice(-2)}`;
  const ini=new Date(y,m-1,1), fin=new Date(y,m,0); fin.setHours(23,59,59,999);
  const cerrado = String(getCfg("PERIODO_CERRADO")||"") === "1";
  return {ini,fin,label,cerrado};
}

function periodosSheet_(){
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName("Periodos") || ss.insertSheet("Periodos");
}

function abrirPeriodo(label){
  const sh = periodosSheet_();
  const per = label || getCfg("PERIODO_LABEL");
  if (label) setCfg_("PERIODO_LABEL", per);
  const vals = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),2).getValues();
  for (let i=0;i<vals.length;i++){
    if (vals[i][0] === per){ sh.getRange(i+2,2).setValue("Abierto"); break; }
  }
  setCfg_("PERIODO_CERRADO", "");
  log(`Periodo ${per} abierto`);
}

function cerrarPeriodo(label){
  const sh = periodosSheet_();
  const per = label || getCfg("PERIODO_LABEL");
  if (label) setCfg_("PERIODO_LABEL", per);
  const vals = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),2).getValues();
  for (let i=0;i<vals.length;i++){
    if (vals[i][0] === per){ sh.getRange(i+2,2).setValue("Cerrado"); break; }
  }
  setCfg_("PERIODO_CERRADO", "1");
  log(`Periodo ${per} cerrado`);
}

function calcularIVA_Mensual(){
  const pr=periodo_();
  const pol=SpreadsheetApp.getActive().getSheetByName("Polizas").getRange(2,1,Math.max(0,SpreadsheetApp.getActive().getSheetByName("Polizas").getLastRow()-1),12).getValues();
  let tras16=0, tras8=0, acred16=0, acred8=0, retIVA=0;
  pol.forEach(r=>{
    const fecha=new Date(r[0]); if(fecha<pr.ini||fecha>pr.fin) return;
    const cta=String(r[3]); const debe=Number(r[6]||0), haber=Number(r[7]||0);
    if(cta==="240-200") tras16+=haber; if(cta==="240-210") tras8+=haber; if(cta==="240-300") acred16+=debe; if(cta==="240-310") acred8+=debe;
  });
  const neto = round2((tras16+tras8) - (acred16+acred8) - retIVA);
  log(`IVA ${pr.label}: Trasl ${(tras16+tras8)} - Acred ${(acred16+acred8)} = Neto ${neto}`);
  return {periodo:pr.label, trasladado:tras16+tras8, acreditable:acred16+acred8, neto};
}

function calcularISR_PM_Mensual(){
  const pr=periodo_(); const coef=Number(getCfg("COEF_UTILIDAD")||CFG.COEF_UTILIDAD); const tasa=Number(getCfg("ISR_PM_TASA")||CFG.ISR_PM_TASA);
  const pol=SpreadsheetApp.getActive().getSheetByName("Polizas").getRange(2,1,Math.max(0,SpreadsheetApp.getActive().getSheetByName("Polizas").getLastRow()-1),12).getValues();
  let ingresosAcum=0;
  pol.forEach(r=>{ const fecha=new Date(r[0]); if(fecha.getFullYear()!==pr.ini.getFullYear()) return; const cta=String(r[3]); const haber=Number(r[7]||0); if(cta.startsWith("400-")) ingresosAcum+=haber; });
  const utilidadEst=ingresosAcum*coef; const impuesto=round2(utilidadEst*tasa);
  log(`ISR PM ${pr.label}: Ingresos Acum ${ingresosAcum} * coef ${coef} * tasa ${tasa} = ${impuesto}`);
  return {periodo:pr.label, ingresosAcum, coef, tasa, impuesto};
}

function generarDIOT_CSV(){
  const ss=SpreadsheetApp.getActive(); const egr=ss.getSheetByName("Egresos");
  if(egr.getLastRow()<2){ safeAlert_("No hay egresos"); return; }
  const vals=egr.getRange(2,1,egr.getLastRow()-1, egr.getLastColumn()).getValues();
  const mapa={};
  vals.forEach(r=>{
    const rfc=r[3], prov=r[2], tasa=String(r[5]||"16"), sub=Number(r[6]||0), iva=Number(r[7]||0);
    if(!rfc) return;
    if(!mapa[rfc]) mapa[rfc]={prov, b16:0,i16:0,b8:0,i8:0,exe:0};
    if(tasa==="16"){ mapa[rfc].b16+=sub; mapa[rfc].i16+=iva; }
    else if(tasa==="8"){ mapa[rfc].b8+=sub; mapa[rfc].i8+=iva; }
    else { mapa[rfc].exe+=sub; }
  });
  const out=Object.keys(mapa).map(k=>[k,mapa[k].prov,"NACIONAL", round2(mapa[k].b16), round2(mapa[k].i16), round2(mapa[k].b8), round2(mapa[k].i8), round2(mapa[k].exe)]);
  const sh=ss.getSheetByName("DIOT");
  sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),8).clearContent();
  if(out.length) sh.getRange(2,1,out.length,8).setValues(out);
  const csv=["RFC,Proveedor,Tipo,Base16,IVA16,Base8,IVA8,Exento"].concat(out.map(r=>r.join(","))).join("\n");
  const blobCsv = Utilities.newBlob(csv, 'text/csv', `DIOT_${periodo_().label}.csv`);
  const file = DriveApp.createFile(blobCsv);
  log(`DIOT CSV: ${file.getUrl()}`);
}

/********************  PDF + CORREO + RECORDATORIOS  ********************/
function exportarPDFsEstados(){
  const ss=SpreadsheetApp.getActive();
  const folderId=getCfg("CARPETA_PDFS_ID");
  const folder=folderId? DriveApp.getFolderById(folderId): DriveApp.createFolder(`PDFs_${Date.now()}`);
  const ids=[];
  ["EstadoResultados","BalanceGeneral","Balanza","KPIs"].forEach(name=>{
    const sh=ss.getSheetByName(name);
    const blob=exportSheetAsPdfBlob_(ss, sh);
    const f=folder.createFile(blob).setName(`${name}_${periodo_().label}.pdf`);
    ids.push(f.getId());
  });
  log(`PDFs listos en ${folder.getName()}`);
  return ids;
}

function enviarPaqueteFiscal(){
  const ss=SpreadsheetApp.getActive();
  const ids=exportarPDFsEstados();
  const iva=calcularIVA_Mensual();
  const isr=calcularISR_PM_Mensual();
  const tpl=getPlantilla("ENTREGA_ESTADOS");
  const asunto=render(tpl.asunto,{periodo:periodo_().label, empresa:ss.getName()});
  const kpiPrincipal=SpreadsheetApp.getActive().getSheetByName("KPIs").getRange(2,1,1,2).getValues()[0][1]||"";
  const html=render(tpl.html,{periodo:periodo_().label, empresa:ss.getName(), kpi_principal:kpiPrincipal});
  const to=getCfg("CORREOS_DESTINO")||Session.getActiveUser().getEmail();
  const files=ids.map(id=> DriveApp.getFileById(id).getBlob());
  GmailApp.sendEmail(to, asunto, "HTML", {htmlBody: html, attachments: files});
  log(`Correo enviado a ${to}`);
}

/********************  TASKS + NOTA TIPO KEEP  ********************/
function programarRecordatoriosMensuales(){
  ScriptApp.getProjectTriggers().forEach(t=> ScriptApp.deleteTrigger(t));
  crearTriggerDia_(10, "calcularIVA_Mensual");
  crearTriggerDia_(10, "calcularISR_PM_Mensual");
  crearTriggerDia_(12, "enviarPaqueteFiscal");
  crearTriggerDia_(16, "recordatorioVencimiento_");
  log("Recordatorios programados");
}
function crearTriggerDia_(dia, func){ const now=new Date(); const run=new Date(now.getFullYear(), now.getMonth(), dia, 9, 0, 0); ScriptApp.newTrigger(func).timeBased().at(run).create(); }
function recordatorioVencimiento_(){
  const iva=calcularIVA_Mensual();
  const isr=calcularISR_PM_Mensual();
  const tpl=getPlantilla("AVISO_IMPUESTOS");
  const html=render(tpl.html,{periodo:periodo_().label, iva_neto:iva.neto, isr_estimado:isr.impuesto, fecha_limite: fechaVencimientoSAT_()});
  const asunto=render(tpl.asunto,{periodo:periodo_().label});
  const to=getCfg("CORREOS_DESTINO")||Session.getActiveUser().getEmail();
  GmailApp.sendEmail(to, asunto, "HTML", {htmlBody: html});
}
function crearTareasObligaciones(){ const lista=asegurarListaTasks_(getCfg("LISTA_TASKS")||CFG.LISTA_TASKS); const pr=periodo_(); const venc=fechaVencimientoSAT_(); [{t:`Revisión balanza ${pr.label}`,n:"Validar saldos",d:venc},{t:`Presentación IVA ${pr.label}`,n:"DIOT y determinación",d:venc},{t:`Pago ISR ${pr.label}`,n:"Provisional",d:venc}].forEach(it=> Tasks.Tasks.insert({title:it.t, notes:it.n, due: toRfc3339_(it.d)}, lista.id)); log("Tareas creadas"); }
function asegurarListaTasks_(nombre){ const listas=Tasks.Tasklists.list().items||[]; let l=listas.find(x=>x.title===nombre); if(!l) l=Tasks.Tasklists.insert({title:nombre}); return l; }
function generarNotaKeep(){ const pr=periodo_(); const doc=DocumentApp.create(`Nota — ${pr.label}`); doc.getBody().appendParagraph("Nota Contable").setHeading(DocumentApp.ParagraphHeading.HEADING1); doc.getBody().appendParagraph(`Periodo ${pr.label}. IVA/ISR listos. PDFs generados.`); doc.saveAndClose(); log(`Nota Doc: ${doc.getUrl()}`); }

/********************  UTILIDADES  ********************/
function getOrCreateSheet(ss,name){ return ss.getSheetByName(name)||ss.insertSheet(name); }

// --- Configuración con caché ---
function configSheet_(){ return SpreadsheetApp.getActive().getSheetByName("Config"); }
function loadCfg_(){
  const sh = configSheet_();
  if (!sh) return {};
  const vals = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < vals.length; i++){
    const key = vals[i][0];
    if (key) map[key] = vals[i][1];
  }
  return map;
}
function refreshCfgCache(){ CFG_CACHE = null; }
function getCfg(clave, def = null){
  if (CFG_CACHE === null) CFG_CACHE = loadCfg_();
  return Object.prototype.hasOwnProperty.call(CFG_CACHE, clave) ? CFG_CACHE[clave] : def;
}
function getCfgRow_(clave){
  const sh = configSheet_();
  if (!sh) return -1;
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2,1,last-1,1).getValues();
  for (let i=0;i<vals.length;i++){
    if (vals[i][0] === clave) return i+2;
  }
  return -1;
}
function setCfg_(clave, valor){
  const sh = configSheet_();
  if (!sh) return;
  const row = getCfgRow_(clave);
  if (row>0){ sh.getRange(row,2).setValue(valor); }
  else {
    const lr = sh.getLastRow();
    sh.getRange(lr+1,1,1,3).setValues([[clave,valor,""]]);
  }
  if (CFG_CACHE) CFG_CACHE[clave] = valor;
}
function getPlantilla(cod){ const sh=SpreadsheetApp.getActive().getSheetByName("Plantillas"); const v=sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),3).getValues(); for(let r of v){ if(r[0]===cod) return {asunto:r[1], html:r[2]}; } return {asunto:"Aviso", html:"<p>Contenido</p>"}; }
function render(tpl,ctx){ return tpl.replace(/{{(\w+)}}/g,(_,k)=> ctx[k]!==undefined? ctx[k]: ""); }
function exportSheetAsPdfBlob_(ss, sheet){
  const url=`https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&portrait=false&size=A4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&gid=${sheet.getSheetId()}`;
  const token=ScriptApp.getOAuthToken();
  const resp=UrlFetchApp.fetch(url,{headers:{Authorization:`Bearer ${token}`}});
  const blob = resp.getBlob().setName(`${sheet.getName()}.pdf`).setContentType('application/pdf');
  return blob;
}
function toRfc3339_(date){ return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'"); }
function cleanXmlText_(s){ if(!s) return s; s = s.replace(/^﻿/, ''); const i = s.indexOf('<'); if(i>0) s = s.slice(i); return s; }
function fechaVencimientoSAT_(){ const pr=periodo_(); const d=new Date(pr.ini); d.setMonth(d.getMonth()+1); d.setDate(17); d.setHours(12,0,0,0); return d; }
function log(msg){
  try{
    const sh = SpreadsheetApp.getActive().getSheetByName("Logs");
    if (sh) sh.appendRow([new Date(), String(msg)]);
    else Logger.log(String(msg));
  }catch(e){
    Logger.log(String(msg));
  }
}
function attrMap_(el){ const m={}; el.getAttributes().forEach(a=> m[a.getName()]=a.getValue()); return m; }
function parseDate_(s){ return s? new Date(String(s).replace("T"," ")): new Date(); }
function toNum(n){ return Number(n||0); }
function round2(n){ return Math.round(Number(n)*100)/100; }
function tasaLabel_(t){ if(t===0.16) return "16"; if(t===0.08) return "8"; return "0"; }

/********************  HOJA MENSUAL + VÍNCULOS + MAESTROS  ********************/
function registrarEnHojaMes_(label, rowMes){
  const name = `CFDI_${label}`;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh){ sh = ss.insertSheet(name); sh.getRange(1,1,1,20).setValues([["Origen","Tipo","Fecha","Serie","Folio","UUID","RFC Emisor","Nombre Emisor","RFC Receptor","Nombre Receptor","Moneda","SubTotal","IVA","Total","MetodoPago","FormaPago","UsoCFDI","UUIDs Rel","Relacionado","Notas"]]).setFontWeight("bold"); }
  sh.appendRow(rowMes);
}

function vincularDocumentosPeriodo_(label){
  const sh = SpreadsheetApp.getActive().getSheetByName(`CFDI_${label}`); if(!sh) return;
  const lr=sh.getLastRow(); if(lr<2) return;
  const vals=sh.getRange(2,1,lr-1,19).getValues();
  const idxI={};
  vals.forEach(r=>{ if((r[1]||"")==="I"){ const u=(r[5]||"").toString().trim(); if(u) idxI[u]=true; }});
  const out=vals.map(r=>{ const tipo=(r[1]||""); const rels=(r[17]||"").toString().split(/[\s,;|]+/).filter(Boolean); let ok=false; if(tipo!=="I"&&rels.length){ ok=rels.some(u=> idxI[u]); } return [ok?"Sí":"No"]; });
  sh.getRange(2,19,out.length,1).setValues(out);
}

function aplicarPagosDesdeHojaMes_(label){
  const ss=SpreadsheetApp.getActive();
  const shMes=ss.getSheetByName(`CFDI_${label}`); if(!shMes) return;
  const lr=shMes.getLastRow(); if(lr<2) return;
  const rows=shMes.getRange(2,1,lr-1,19).getValues();
  const pagos=rows.filter(r=> String(r[1]||"")==="P");
  if(!pagos.length) return;
  const shI=ss.getSheetByName("Ingresos"); const shE=ss.getSheetByName("Egresos");
  const idx={};
  [[shI,"I"],[shE,"E"]].forEach(([sh,_])=>{
    const l=sh.getLastRow(); if(l>1){ sh.getRange(2,1,l-1, sh.getLastColumn()).getValues().forEach((r,i)=>{ const u=(r[13]||"").toString(); if(u) idx[u]={sh,row:i+2}; }); }
  });
  let n=0;
  pagos.forEach(p=>{
    const rels=(p[17]||"").toString().split(/[\s,;|]+/).filter(Boolean);
    rels.forEach(u=>{ const hit=idx[u]; if(hit){ hit.sh.getRange(hit.row,13).setValue("Pagado"); n++; } });
  });
  if(n) log(`Pagos aplicados (P → EstadoPago): ${n}`);
}

/********************  IMPORTACIÓN BANCOS (CSV)  ********************/
function importarBancoDesdeCsv(nombreArchivo){
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName("Bancos");
  const parent=DriveApp.getFileById(ss.getId()).getParents().next();
  const files=parent.getFilesByName(nombreArchivo);
  if(!files.hasNext()) throw new Error("CSV no encontrado");
  const csv=files.next().getBlob().getDataAsString();
  const rows=Utilities.parseCsv(csv);
  sh.getRange(sh.getLastRow()+1,1,rows.length, Math.min(11, rows[0].length)).setValues(rows);
  log(`Banco importado: ${nombreArchivo}`);
}
// —— Maestros: alta/actualización automática + resaltado de cuentas faltantes
function upsertMaestrosDesdeCFDI_(origen, Emisor, Receptor){
  const ss=SpreadsheetApp.getActive();
  if (origen==='Emitidas'){ // clientes = Receptor
    const sh=ss.getSheetByName("Clientes");
    const rfc=(Receptor.Rfc||"").toString().trim(); if(!rfc) return;
    const lr=sh.getLastRow(); const vals=lr>1? sh.getRange(2,1,lr-1,6).getValues():[];
    const idx=vals.findIndex(r=> (r[2]||"").toString().trim()===rfc);
    if(idx>=0){
      if(!vals[idx][1] && Receptor.Nombre) sh.getRange(idx+2,2).setValue(Receptor.Nombre);
    } else {
      const id="CLI-"+("000"+(lr)).slice(-3);
      sh.getRange(lr+1,1,1,6).setValues([[id, Receptor.Nombre||rfc, rfc, "", "Cliente", ""]]);
      sh.getRange(lr+1,6).setBackground("#CBE290"); // camel → cuenta pendiente
    }
  } else { // Recibidas → proveedores = Emisor
    const sh=ss.getSheetByName("Proveedores");
    const rfc=(Emisor.Rfc||"").toString().trim(); if(!rfc) return;
    const lr=sh.getLastRow(); const vals=lr>1? sh.getRange(2,1,lr-1,8).getValues():[];
    const idx=vals.findIndex(r=> (r[2]||"").toString().trim()===rfc);
    if(idx>=0){
      if(!vals[idx][1] && Emisor.Nombre) sh.getRange(idx+2,2).setValue(Emisor.Nombre);
    } else {
      const id="PROV-"+("000"+(lr)).slice(-3);
      // Tipo por defecto “General”, IVA 16; puedes ajustar luego
      sh.getRange(lr+1,1,1,8).setValues([[id, Emisor.Nombre||rfc, rfc, "", "General", "", "16", ""]]);
      sh.getRange(lr+1,8).setBackground("#CBE290");
    }
  }
}

function resaltarPendientesMaestros_(){
  const ss=SpreadsheetApp.getActive();
  const cli=ss.getSheetByName("Clientes");
  const prov=ss.getSheetByName("Proveedores");
  marcarPendientes_(cli, 6, "CLI");
  marcarPendientes_(prov, 8, "PROV");
}

function marcarPendientes_(sh, colCuenta, pref){
  const lr=sh.getLastRow(); if(lr<2) return;
  const ids=sh.getRange(2,1,lr-1,1).getValues();
  const cuentas=sh.getRange(2,colCuenta,lr-1,1).getValues();
  const bgs=[];
  for(let i=0;i<lr-1;i++){
    if(!ids[i][0]) sh.getRange(i+2,1).setValue(pref+"-"+("000"+(i+1)).slice(-3));
    const vacio=!cuentas[i][0];
    bgs.push([vacio? "#CBE290" : "#FFFFFF"]); // camel si falta cuenta
  }
  sh.getRange(2,colCuenta,lr-1,1).setBackgrounds(bgs);
}
function importarCFDI_(folderId, esIngreso){
  if (!folderId) throw new Error("Config: CARPETA_CFDI_* no definida");

  const pr = periodo_();
  const fol = DriveApp.getFolderById(folderId);

  // 1) sets para evitar duplicados
  const setUUID = uuidsExistentes_();                          // por UUID
  const ss = SpreadsheetApp.getActive();
  const shI = ss.getSheetByName("Ingresos");
  const shE = ss.getSheetByName("Egresos");
  const setI = buildKeysSet_(shI);                             // llave compuesta Ingresos
  const setE = buildKeysSet_(shE);                             // llave compuesta Egresos

  // 2) leer fuentes
  const fuentes = recolectarXMLs_(fol);
  const rowsI = [], rowsE = [];
  let totFuentes=0, ok=0, dup=0, err=0, fuera=0, ruteoI=0, ruteoE=0;

  fuentes.forEach(src => {
    totFuentes++;
    try{
      let xml;
      try{ xml = XmlService.parse(cleanXmlText_(src.getText('UTF-8'))); }
      catch(e1){ xml = XmlService.parse(cleanXmlText_(src.getText())); }
      const root = xml.getRootElement();
      const A = attrMap_(root);

      const Fecha = parseDate_(A.Fecha||A.fecha); 
      if(!(Fecha>=pr.ini && Fecha<=pr.fin)){ fuera++; return; }

      const Folio = A.Folio||A.folio||""; 
      const Serie = A.Serie||A.serie||"";
      const Subtotal = toNum(A.SubTotal||A.Subtotal); 
      const Total = toNum(A.Total);
      const Tipo = (A.TipoDeComprobante||A.Tipo||"").toString().toUpperCase();
      const Moneda = A.Moneda||"MXN"; 
      const Metodo = A.MetodoPago||""; 
      const Forma  = A.FormaPago||"";

      let UUID="", Emisor={}, Receptor={}, uso=""; 
      let rels=[];
      const desc=root.getDescendants();
      for (let n of desc){ 
        const el=n.asElement&&n.asElement(); if(!el) continue; 
        const nm=el.getName();
        if(/^Emisor$/i.test(nm)) Emisor=attrMap_(el);
        else if(/^Receptor$/i.test(nm)) { Receptor=attrMap_(el); uso = Receptor.UsoCFDI || uso; }
        else if(/TimbreFiscalDigital/i.test(nm)) { const a=attrMap_(el); if(a.UUID) UUID=a.UUID; }
        else if(/CfdiRelacionado/i.test(nm))     { const a=attrMap_(el); if(a.UUID) rels.push(a.UUID); }
        else if(/DoctoRelacionado/i.test(nm))    { const a=attrMap_(el); if(a.IdDocumento) rels.push(a.IdDocumento); }
      }

      // —— bloqueo por UUID
      const keyUUID=(UUID||"").trim().toUpperCase();
      if(keyUUID && setUUID[keyUUID]){ dup++; return; }

      const tasa = detectarTasaDesdeImpuestos_(root);
      const ivaCalc = round2(Total - Subtotal);
      const origen = esIngreso? 'Emitidas':'Recibidas';

      // Hoja mensual (para auditoría de periodo)
      const rowMes=[origen, Tipo, Fecha, Serie, Folio, UUID, Emisor.Rfc||'', Emisor.Nombre||'', Receptor.Rfc||'', Receptor.Nombre||'', Moneda, Subtotal, ivaCalc, Total, Metodo, Forma, uso, rels.join('|'), '', ''];
      registrarEnHojaMes_(pr.label, rowMes);

      // Maestros
      upsertMaestrosDesdeCFDI_(origen, Emisor, Receptor);

      // Movimiento a Ingresos/Egresos (con llave compuesta anti-dup)
      const tercero = esIngreso? (Receptor.Nombre||Receptor.Rfc) : (Emisor.Nombre||Emisor.Rfc);
      const rfc     = esIngreso? (Receptor.Rfc||"")             : (Emisor.Rfc||"");
      const concepto = `CFDI ${src.name || ''} [Tipo:${Tipo||'?'}]`;
      const rowIE = [Fecha, Folio, tercero, rfc, concepto, tasaLabel_(tasa), Subtotal, ivaCalc, "", Total, "Transferencia","Banco MXN","Pendiente", UUID, "", "", "No"];

      const k = keyMovimientoFromRow_(rowIE);
      if (esIngreso){
        if(setI[k]){ dup++; }
        else { rowsI.push(rowIE); setI[k]=true; ruteoI++; }
      } else {
        if(setE[k]){ dup++; }
        else { rowsE.push(rowIE); setE[k]=true; ruteoE++; }
      }

      if(keyUUID) setUUID[keyUUID]=true; // marca como visto
      ok++;
    }catch(e){ err++; log(`Error XML fuente ${src.name||"(sin nombre)"}: ${e}`); }
  });

  // 3) volcar en hojas
  if (rowsI.length){ shI.getRange(shI.getLastRow()+1,1,rowsI.length,rowsI[0].length).setValues(rowsI); }
  if (rowsE.length){ shE.getRange(shE.getLastRow()+1,1,rowsE.length,rowsE[0].length).setValues(rowsE); }

  // 4) vínculos y resaltado
  vincularDocumentosPeriodo_(pr.label);
  aplicarPagosDesdeHojaMes_(pr.label);
  resaltarPendientesMaestros_();

  log(`Importación CFDI — fuentes:${totFuentes} ok:${ok} dup:${dup} err:${err} fuera:${fuera} → I:${ruteoI} E:${ruteoE}`);
}
function keyMovimientoFromRow_(row){
  const tz = Session.getScriptTimeZone();
  const fecha = row[0] ? Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd') : '';
  const folio = (row[1]||'').toString().trim();
  const tercero = (row[2]||'').toString().trim().toUpperCase();
  const rfc = (row[3]||'').toString().trim().toUpperCase();
  const total = Number(row[9]||0).toFixed(2);
  const uuid = (row[13]||'').toString().trim().toUpperCase();
  if (uuid) return `U:${uuid}`;
  return `B:${fecha}|${folio}|${tercero}|${rfc}|${total}`;
}

function buildKeysSet_(sh){
  const set={};
  const lr=sh.getLastRow();
  if(lr>1){
    const vals=sh.getRange(2,1,lr-1, sh.getLastColumn()).getValues();
    vals.forEach(r=>{ const k=keyMovimientoFromRow_(r); if(k) set[k]=true; });
  }
  return set;
}

