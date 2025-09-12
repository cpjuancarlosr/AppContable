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
  subMenuProcess.addItem("🧩 Generar Pólizas de Nuevas Transacciones", "generarPolizasDesdeMovimientos");
  subMenuProcess.addItem("📐 Recalcular Estados + KPIs", "recalcularEstados");
  subMenuProcess.addSeparator();
  subMenuProcess.addItem("🏦 Conciliar Bancos (Automático)", "conciliarBancariaAvanzada");
  menu.addSubMenu(subMenuProcess);

  const subMenuReports = ui.createMenu("4. Reportes y Consultas");
  subMenuReports.addItem("🧾 Calcular IVA Mensual", "calcularIVA_Mensual");
  subMenuReports.addItem("💸 Calcular ISR Mensual", "calcularISR_Mensual");
  subMenuReports.addItem("🧩 Generar DIOT (CSV)", "generarDIOT_CSV");
  subMenuReports.addSeparator();
  subMenuReports.addItem("📊 Ver Reporte para Cliente", "verReporteCliente");
  subMenuReports.addItem("📝 Generar Guión de Reunión", "generarGuionReunion");
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

/**
 * Aplica formato estándar a una hoja: congela la primera fila y aplica bandas de colores.
 * @param {Sheet} sh - La hoja a formatear.
 */
function formatSheet_(sh) {
  sh.setFrozenRows(1);
  const range = sh.getRange(2, 1, sh.getMaxRows(), sh.getMaxColumns());
  try {
    const banding = range.getBandings()[0] || range.applyRowBanding();
    banding.setHeaderRowColor(PALETA.navy)
           .setFirstRowColor(PALETA.baseBg)
           .setSecondRowColor("#F0F0F0");
  } catch(e) {
    // Las bandas pueden fallar en hojas muy grandes o vacías, ignorar error.
    log(`No se pudo aplicar banding a la hoja ${sh.getName()}: ${e.message}`);
  }
}

/********************  INSTALACIÓN  ********************/
function setupSistemaMX(){
  const ss = SpreadsheetApp.getActive();
  const hojas = [
    "Config","Clientes","Proveedores","CatCuentas","Ingresos","Egresos","Bancos",
    "Polizas","Mayor","Balanza","EstadoResultados","BalanceGeneral","KPIs","AnalisisFinanciero",
    "Plantillas","PagosImpuestos","DIOT","Logs","Dashboard", "Asientos"
  ];
  hojas.forEach(h => getOrCreateSheet(ss,h));

  prepararConfig();
  prepararMaestros();
  prepararCatCuentas();
  prepararIngresosEgresos();
  prepararBancos();
  prepararPolizasMayor();
  prepararEstadosPlantillas();
  prepararAnalisisFinanciero();
  prepararPeriodos();
  importarReglasDeAsientos();
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
    ["REGIMEN_FISCAL", "PM_General", "Ej: PM_General, RESICO, PFAE"],
    ["TIPO_PERSONA", "Moral", "Moral o Fisica"],
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
  const ss = SpreadsheetApp.getActive();
  const cli = ss.getSheetByName("Clientes");
  cli.clear(); cli.getRange(1,1,1,6).setValues([["ID","Nombre","RFC","Email","Tipo","Cuenta Contable"]]).setFontWeight("bold");
  const prov = ss.getSheetByName("Proveedores");
  prov.clear(); prov.getRange(1,1,1,8).setValues([["ID","Nombre","RFC","Email","Tipo","Retención ISR","Tasa IVA","Cuenta Contable"]]).setFontWeight("bold");
  formatSheet_(cli);
  formatSheet_(prov);
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
  formatSheet_(sh);
}

function prepararIngresosEgresos(){
  const head = ["Fecha","Folio","Cliente/Proveedor","RFC","Concepto","Tasa IVA","SUBTOTAL","IVA","IVA RETENIDO","ISR RETENIDO","Total","MetodoPago","FormaPago","Banco/Cuenta","EstadoPago","UUID","UUIDs Relacionados","Cuenta Contable","PolizaID","Conciliado","Poliza Generada?"];
  const ss = SpreadsheetApp.getActive();
  const shI = ss.getSheetByName("Ingresos"); shI.clear(); shI.getRange(1,1,1,head.length).setValues([head]).setFontWeight("bold");
  const shE = ss.getSheetByName("Egresos"); shE.clear(); shE.getRange(1,1,1,head.length).setValues([head]).setFontWeight("bold");
  formatSheet_(shI);
  formatSheet_(shE);
}

function prepararBancos(){
  const sh = SpreadsheetApp.getActive().getSheetByName("Bancos"); sh.clear();
  sh.getRange(1,1,1,11).setValues([["Fecha","Descripcion","Cargo","Abono","Importe","Referencia","Banco","Cuenta","FolioFactura","UUID","Conciliado"]]).setFontWeight("bold");
  formatSheet_(sh);
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

function prepararAnalisisFinanciero() {
  const sh = SpreadsheetApp.getActive().getSheetByName("AnalisisFinanciero");
  sh.clear();
  sh.getRange("A1:D1").setValues([["Grupo", "Indicador", "Valor", "Interpretación"]]).setFontWeight("bold");

  const data = [
    // --- Liquidez ---
    ["Liquidez", "Razón Circulante", `=IFERROR(KPIs!B8, 0)`, "Mide la capacidad de cubrir deudas a corto plazo. Un valor > 1 es generalmente bueno."],
    ["Liquidez", "Prueba Ácida", `=IFERROR(KPIs!B9, 0)`, "Similar a la razón circulante, pero excluyendo el inventario, que es menos líquido."],
    // --- Rentabilidad ---
    ["Rentabilidad", "Margen Bruto", `=IFERROR(EstadoResultados!C3 / EstadoResultados!C2, 0)`, "Porcentaje de ingresos que queda después de cubrir el costo de ventas."],
    ["Rentabilidad", "Margen Operativo", `=IFERROR(EstadoResultados!C5 / EstadoResultados!C2, 0)`, "Eficiencia de la operación principal del negocio antes de intereses e impuestos."],
    ["Rentabilidad", "Margen Neto", `=IFERROR(KPIs!B2, 0)`, "El porcentaje de cada peso de venta que se convierte en ganancia neta."],
    ["Rentabilidad", "ROA (Return on Assets)", `=IFERROR(KPIs!B3, 0)`, "Qué tan eficientemente se usan los activos para generar ganancias."],
    ["Rentabilidad", "ROE (Return on Equity)", `=IFERROR(KPIs!B4, 0)`, "Rendimiento generado sobre la inversión de los accionistas."],
    // --- Actividad / Eficiencia ---
    ["Eficiencia", "Rotación de Cuentas por Cobrar", `=IFERROR(EstadoResultados!C2 / IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 3)="120")), 0), 0)`, "Cuántas veces la empresa convierte sus cuentas por cobrar en efectivo en un periodo."],
    ["Eficiencia", "Días de Cartera", `=IFERROR(365 / IFERROR(VLOOKUP("Rotación de Cuentas por Cobrar", B:C, 2, FALSE), 1), "")`, "Número promedio de días que tarda una empresa en cobrar sus ventas a crédito."],
    ["Eficiencia", "Ciclo de Conversión de Efectivo", "Próximamente", "Días que tarda la empresa en convertir sus inversiones en inventarios y otros recursos en efectivo."],
    // --- Apalancamiento ---
    ["Apalancamiento", "Razón de Endeudamiento", `=IFERROR(BalanceGeneral!C3 / BalanceGeneral!C2, 0)`, "Porcentaje de los activos que se financian a través de deuda."],
    ["Apalancamiento", "Deuda a Capital", `=IFERROR(BalanceGeneral!C3 / (BalanceGeneral!C4+BalanceGeneral!C5), 0)`, "Compara la deuda total con el capital de los accionistas."]
  ];

  const values = data.map(row => [row[0], row[1], "", row[3]]);
  const formulas = data.map(row => [row[2]]);

  sh.getRange(2, 1, data.length, 4).setValues(values);
  sh.getRange(2, 3, formulas.length, 1).setFormulas(formulas);

  sh.getRange("C:C").setNumberFormat("0.00");
  sh.getRange("B:B").setFontWeight("bold");
  sh.autoResizeColumns(1, 4);
  formatSheet_(sh);
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

  const dash = ss.getSheetByName("Dashboard");
  if (!dash) return;

  dash.clear();
  dash.getRange("A1").setValue(CFG.NOMBRE_SISTEMA).setFontSize(20).setFontWeight("bold").setFontColor(PALETA.navy);

  // KPIs Section
  dash.getRange("A3").setValue("Resumen Financiero").setFontSize(14).setFontWeight("bold");
  dash.getRange("A4:B6").setValues([
    ["Ingresos Totales", `=IFERROR(SUM(Ingresos!G:G), 0)`],
    ["Egresos Totales", `=IFERROR(SUM(Egresos!G:G), 0)`],
    ["Utilidad Bruta", "=B4-B5"]
  ]).setNumberFormat("$#,##0.00");
  dash.getRange("A4:A6").setFontWeight("bold");

  // Navigation Section
  dash.getRange("D3").setValue("Navegación Rápida").setFontSize(14).setFontWeight("bold");
  const shI = ss.getSheetByName("Ingresos");
  const shE = ss.getSheetByName("Egresos");
  const shP = ss.getSheetByName("Polizas");
  const shB = ss.getSheetByName("Balanza");

  const links = [
    ["Ver Ingresos", shI ? `=HYPERLINK("#gid=${shI.getSheetId()}", "Ir a Ingresos")` : "No encontrada"],
    ["Ver Egresos", shE ? `=HYPERLINK("#gid=${shE.getSheetId()}", "Ir a Egresos")` : "No encontrada"],
    ["Ver Pólizas", shP ? `=HYPERLINK("#gid=${shP.getSheetId()}", "Ir a Pólizas")` : "No encontrada"],
    ["Ver Balanza", shB ? `=HYPERLINK("#gid=${shB.getSheetId()}", "Ir a Balanza")` : "No encontrada"]
  ];
  dash.getRange("D4:E7").setValues(links);
  dash.getRange("D4:D7").setFontWeight("bold");

  // Formatting
  dash.getRange("A3:B6").setBorder(true, true, true, true, true, true);
  dash.getRange("D3:E7").setBorder(true, true, true, true, true, true);
  dash.autoResizeColumns(1, 5);
}

/********************  IMPORTACIÓN DE CFDI (XML)  ********************/
function importarCFDIIngresos(){ importarCFDI_(getCfg("CARPETA_CFDI_ING_ID"), true); }
function importarCFDIEgresos(){ importarCFDI_(getCfg("CARPETA_CFDI_EGR_ID"), false); }


// Recorre carpeta y subcarpetas. Devuelve arreglo de {name, getText()}
function recolectarXMLs_(folder){
  const out = [];
  const MAX_ZIP_MB = 50;
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

// Si es atajo de Drive, devuelve el archivo real (requiere Servicio Avanzado de Drive habilitado)
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

// Conjunto de UUID ya cargados en Ingresos/Egresos para evitar duplicados
function uuidsExistentes_(){
  const ss = SpreadsheetApp.getActive();
  const set = {};
  ["Ingresos","Egresos"].forEach(n=>{
    const sh = ss.getSheetByName(n); const lr = sh.getLastRow();
    if (lr>1){ sh.getRange(2, 16, lr-1, 1).getValues().forEach(r=>{ const u = (r[0]||"").toString().trim(); if(u) set[u]=true; }); }
  });
  return set;
}

function detectarTasaDesdeImpuestos_(root){
  const d = root.getDescendants();
  for (let n of d){ try{ const el=n.asElement(); if(!el) continue; if(/Traslado/i.test(el.getName())){ const a=attrMap_(el); if(a.Impuesto=="002"&&a.TasaOCuota){ const t=parseFloat(a.TasaOCuota); return t>=0.15?0.16:(t>=0.07?0.08:0); } } }catch(e){} }
  return 0.16;
}

function generarPolizasDesdeMovimientos() {
  const ss = SpreadsheetApp.getActive();
  const polizasSheet = ss.getSheetByName("Polizas");

  log("Iniciando generación de pólizas para nuevas transacciones...");
  const reglas = getReglasAsientos_();
  if (!reglas || reglas.length === 0) {
    safeAlert_("No se encontraron reglas de asientos en la hoja 'Asientos'. El proceso no puede continuar.");
    return;
  }

  const allNuevasPolizas = [];
  let totalProcesadas = 0;

  const processSheet = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet.getLastRow() <= 1) return;

    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const data = dataRange.getValues();

    const unprocessedRows = [];
    const processedRowIndices = [];

    data.forEach((row, index) => {
      if (row[20] !== "Sí") {
        unprocessedRows.push(row);
        processedRowIndices.push(index);
      }
    });

    if (unprocessedRows.length === 0) {
      log(`No hay transacciones nuevas que procesar en la hoja '${sheetName}'.`);
      return;
    }

    const nuevasPolizas = generarPolizasParaMovimientos_(unprocessedRows, reglas);
    allNuevasPolizas.push(...nuevasPolizas);

    processedRowIndices.forEach(index => {
      data[index][20] = "Sí";
    });

    dataRange.setValues(data);
    totalProcesadas += unprocessedRows.length;
    log(`Se procesaron ${unprocessedRows.length} transacciones y se generaron ${nuevasPolizas.length} asientos desde la hoja '${sheetName}'.`);
  };

  processSheet("Ingresos");
  processSheet("Egresos");

  if (allNuevasPolizas.length > 0) {
    polizasSheet.getRange(polizasSheet.getLastRow() + 1, 1, allNuevasPolizas.length, allNuevasPolizas[0].length).setValues(allNuevasPolizas);
  }

  if (totalProcesadas > 0) {
    recalcularEstados();
    safeAlert_(`Proceso completado. Se procesaron ${totalProcesadas} nuevas transacciones.`);
  } else {
    safeAlert_("No se encontraron nuevas transacciones para procesar.");
  }
}

function generarPolizasParaMovimientos_(movimientos, reglas) {
  const polizas = [];
  const processedUuids = {}; // Para no procesar el mismo movimiento dos veces

  for (const mov of movimientos) {
    const uuid = mov[15]; // Corrected index for UUID
    if (uuid && processedUuids[uuid]) {
      continue;
    }

    for (const regla of reglas) {
      if (matchRule_(mov, regla)) {
        const asiento = generarAsientoDesdeRegla_(mov, regla);
        polizas.push(...asiento);
        if (uuid) {
          processedUuids[uuid] = true;
        }
        break; // Pasa al siguiente movimiento una vez que se encuentra una regla
      }
    }
  }
  return polizas;
}

function matchRule_(mov, regla) {
  const tipoCfdiMov = tipoDesdeConcepto_(mov[4]); // Concepto
  const metodoPagoMov = mov[11];                   // MetodoPago
  const relatedUuids = mov[16];                    // UUIDs Relacionados

  // Comprobaciones básicas
  if (regla.tipoCfdi !== '*' && regla.tipoCfdi !== tipoCfdiMov) return false;
  if (regla.metodoPago !== '*' && regla.metodoPago !== metodoPagoMov) return false;

  // Comprobaciones de Condición
  const condicion = regla.condicion.toLowerCase();
  if (condicion.includes("escomplementopago=si") && tipoCfdiMov !== 'P') {
    return false;
  }
  if (condicion.includes("sin complemento") && tipoCfdiMov === 'P') {
    return false;
  }
  if (condicion.includes("relacionada a uuid") && !relatedUuids) {
    return false;
  }

  return true;
}

function generarAsientoDesdeRegla_(mov, regla) {
  // Unpack the movement array with the new 20-column structure
  const [
    fecha, folio, tercero, rfc, concepto,
    , // Tasa IVA
    subtotal, iva, ivaRetenido, isrRetenido, total,
    , , , , // Metodo, Forma, Banco, Estado
    uuid
  ] = mov.map(val => val || 0);

  const ref = `${tipoDesdeConcepto_(concepto)}-${folio||(uuid && uuid.slice(0,8))||''}`;
  const poliza = [];

  const getMontoParaCuenta = (cuenta) => {
    const ctaLower = cuenta.toLowerCase();
    if (ctaLower.includes("cliente") || ctaLower.includes("proveedor") || ctaLower.includes("banco")) return Math.abs(total);
    if (ctaLower.includes("iva retenido")) return Math.abs(ivaRetenido);
    if (ctaLower.includes("isr retenido")) return Math.abs(isrRetenido);
    if (ctaLower.includes("iva")) return Math.abs(iva);
    // Por defecto, es el subtotal
    return Math.abs(subtotal);
  };

  const crearFila = (cuenta, desc, debe, haber) => {
    let cuentaFinal = cuenta.replace(/\[RFC\]/g, rfc).replace(/\[Cuenta\]/g, "Bancos"); // Placeholder
    return [fecha, regla.tipoPoliza, ref, cuentaFinal, "", desc, debe, haber, uuid, "Motor de Reglas", "", ""];
  };

  regla.debe.forEach(cuenta => {
    if (!cuenta) return;
    const monto = getMontoParaCuenta(cuenta);
    poliza.push(crearFila(cuenta, regla.concepto, monto, 0));
  });

  regla.haber.forEach(cuenta => {
    if (!cuenta) return;
    const monto = getMontoParaCuenta(cuenta);
    poliza.push(crearFila(cuenta, regla.concepto, 0, monto));
  });

  return poliza;
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
function recalcularEstados(){
  const ss = SpreadsheetApp.getActive();
  const bal = ss.getSheetByName("Balanza");
  const er = ss.getSheetByName("EstadoResultados");
  const bg = ss.getSheetByName("BalanceGeneral");
  const kpi = ss.getSheetByName("KPIs");
  const mayor = ss.getSheetByName("Mayor");
  const p = ss.getSheetByName("Polizas");

  // --- Mayor (Ledger) Calculation (Remains code-driven for detailed drill-down) ---
  const lastMayorRow = mayor.getLastRow();
  if (lastMayorRow > 1) {
    mayor.getRange(2, 1, lastMayorRow - 1, mayor.getLastColumn()).clearContent();
  }
  const lastPolizaRow = p.getLastRow();
  if (lastPolizaRow > 1) {
    const pols = p.getRange(2, 1, lastPolizaRow - 1, 12).getValues();
    const mrows = [];
    const saldos = {};
    pols.forEach(r => {
      const [fecha, , ref, cta, , desc, debe, haber] = [r[0], r[1], r[2], r[3], r[4], r[5], Number(r[6] || 0), Number(r[7] || 0)];
      if (!cta) return;
      if (!saldos[cta]) saldos[cta] = { debe: 0, haber: 0, saldo: 0 };
      saldos[cta].debe += debe;
      saldos[cta].haber += haber;
      saldos[cta].saldo = saldos[cta].debe - saldos[cta].haber;
      mrows.push([cta, fecha, ref, desc, debe, haber, saldos[cta].saldo]);
    });
    if (mrows.length) mayor.getRange(2, 1, mrows.length, 7).setValues(mrows);
  }

  // --- Balanza de Comprobación (Trial Balance) with Formulas ---
  const lastBalanzaRow = bal.getLastRow();
  if (lastBalanzaRow > 1) {
    bal.getRange(2, 1, lastBalanzaRow - 1, 5).clearContent();
  }
  bal.getRange("A2").setFormula(`=IFERROR(SORT(UNIQUE(FILTER(Polizas!D2:D, Polizas!D2:D<>""))), "")`);
  bal.getRange("B2").setFormula(`=ARRAYFORMULA(IF(A2:A<>"", IFERROR(VLOOKUP(A2:A, CatCuentas!A:B, 2, FALSE), "Sin nombre"), ""))`);
  bal.getRange("C2").setFormula(`=ARRAYFORMULA(IF(A2:A<>"", SUMIF(Polizas!D:D, A2:A, Polizas!G:G), ""))`);
  bal.getRange("D2").setFormula(`=ARRAYFORMULA(IF(A2:A<>"", SUMIF(Polizas!D:D, A2:A, Polizas!H:H), ""))`);
  bal.getRange("E2").setFormula(`=ARRAYFORMULA(IF(A2:A<>"", N(C2:C)-N(D2:D), ""))`);

  // --- Estado de Resultados (Income Statement) with Formulas ---
  const lastErRow = er.getLastRow();
  if (lastErRow > 1) {
    er.getRange(2, 1, lastErRow - 1, 3).clearContent();
  }
  const erData = [
    ["Ingresos", "4*", ""],
    ["Costos de Venta", "500-*", ""],
    ["Utilidad Bruta", "", ""],
    ["Gastos de Operación", "510-*", ""],
    ["Utilidad de Operación", "", ""],
    ["Resultado Integral de Financiamiento", "520-*", ""],
    ["Utilidad antes de Impuestos", "", ""],
    ["Impuestos", "530-*", ""],
    ["Utilidad Neta", "", ""]
  ];
  const erFormulas = [
    [`=IFERROR(-SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 1)="4")), 0)`],
    [`=IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 3)="500")), 0)`],
    [`=C2-C3`],
    [`=IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 3)="510")), 0)`],
    [`=C4-C5`],
    [`=IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 3)="520")), 0)`],
    [`=C6-C7`],
    [`=IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 3)="530")), 0)`],
    [`=C8-C9`]
  ];
  er.getRange(2, 1, erData.length, 2).setValues(erData.map(row => [row[0], row[1]]));
  er.getRange(2, 3, erFormulas.length, 1).setFormulas(erFormulas);

  // --- Balance General (Balance Sheet) with Formulas ---
  const lastBgRow = bg.getLastRow();
  if (lastBgRow > 1) {
    bg.getRange(2, 1, lastBgRow - 1, 3).clearContent();
  }
  const bgData = [
    ["Activo", "1*", ""],
    ["Pasivo", "2*", ""],
    ["Capital Contable", "3*", ""],
    ["Utilidad del Ejercicio", "", ""],
    ["Total Pasivo + Capital", "", ""],
    ["Suma de Verificación", "", ""]
  ];
  const bgFormulas = [
    [`=IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 1)="1")), 0)`],
    [`=IFERROR(-SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 1)="2")), 0)`],
    [`=IFERROR(-SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 1)="3")), 0)`],
    [`=IFERROR(EstadoResultados!C10, 0)`], // Utilidad Neta from ER
    [`=C3+C4+C5`],
    [`=C2-C6`]
  ];
  bg.getRange(2, 1, bgData.length, 2).setValues(bgData.map(row => [row[0], row[1]]));
  bg.getRange(2, 3, bgFormulas.length, 1).setFormulas(bgFormulas);

  // --- KPIs with Formulas ---
  const lastKpiRow = kpi.getLastRow();
  if (lastKpiRow > 1) {
    kpi.getRange(2, 1, lastKpiRow - 1, 2).clearContent();
  }
  const kpiData = [
    ["Margen Neto", ""],
    ["ROA (Return on Assets)", ""],
    ["ROE (Return on Equity)", ""],
    ["Razón Circulante", ""],
    ["Prueba Ácida", ""],
    ["Apalancamiento", ""],
    ["Días de Cartera", ""],
    ["Días de Inventario", ""],
    ["Días de Proveedores", ""]
  ];
  const kpiFormulas = [
    [`=IFERROR(BalanceGeneral!C5 / EstadoResultados!C2, 0)`],
    [`=IFERROR(BalanceGeneral!C5 / BalanceGeneral!C2, 0)`],
    [`=IFERROR(BalanceGeneral!C5 / (BalanceGeneral!C4 + BalanceGeneral!C5), 0)`],
    [`=IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 2)="11")) / -SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 2)="21")), 0)`],
    [`=IFERROR((SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 2)="11")) - IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 3)="130")), 0)) / -SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 2)="21")), 0)`],
    [`=IFERROR(BalanceGeneral!C2 / (BalanceGeneral!C4 + BalanceGeneral!C5), 0)`],
    [`=IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 3)="120")) / (EstadoResultados!C2 / 365), 0)`],
    [`=IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 3)="130")) / (EstadoResultados!C3 / 365), 0)`],
    [`=IFERROR(SUM(FILTER(Balanza!E:E, LEFT(Balanza!A:A, 3)="210")) / (EstadoResultados!C3 / 365), 0)`]
  ];
  kpi.getRange(2, 1, kpiData.length, 1).setValues(kpiData.map(row => [row[0]]));
  kpi.getRange(2, 2, kpiFormulas.length, 1).setFormulas(kpiFormulas);

  log("Estados financieros recalculados con fórmulas dinámicas.");
  safeAlert_("Los estados financieros han sido actualizados y ahora usan fórmulas dinámicas.");
}

function mapearNombreCuenta_(cat){ const v=cat.getRange(2,1,Math.max(0,cat.getLastRow()-1),2).getValues(); const m={}; v.forEach(a=>m[a[0]]=a[1]); return m; }
function saldoPorCuenta_(rows, cta){ const r=rows.find(x=> x[0]===cta); return r? Number(r[4]||0):0; }

/********************  CONCILIACIÓN AVANZADA  ********************/
function conciliarBancariaAvanzada(){
  const ss=SpreadsheetApp.getActive(); const shB=ss.getSheetByName("Bancos"); const shI=ss.getSheetByName("Ingresos"); const shE=ss.getSheetByName("Egresos"); const tol=Number(getCfg("DIAS_TOLERANCIA_CONCILIACION"))||CFG.DIAS_TOLERANCIA_CONCILIACION;
  const bankData = shB.getRange(2,1,Math.max(0,shB.getLastRow()-1), shB.getLastColumn()).getValues();

  const bankRowsToUpdate = [];
  const accountingRowsToUpdate = {};

  const idxI=indexarMovs_(shI); const idxE=indexarMovs_(shE); let c=0;

  bankData.forEach((row, index) => {
    const isReconciled = row[10]; // "Conciliado" column
    if (isReconciled === "Sí") return;

    const [fecha,,cargo,abono,importe,ref] = row;
    const monto = Number(importe||cargo||abono||0);
    if(!fecha || !monto) return;

    const target = (Number(cargo||0)>0) ? idxE : idxI;
    const hit = buscarMatch_(target, fecha, monto, ref, tol);

    if(hit){
      // Mark bank row for update
      bankRowsToUpdate.push({row: index + 2, folio: hit.folio || "", uuid: hit.uuid || ""});

      // Mark accounting row for update
      const sheetName = hit.sh.getName();
      if (!accountingRowsToUpdate[sheetName]) {
        accountingRowsToUpdate[sheetName] = [];
      }
      accountingRowsToUpdate[sheetName].push(hit.row);

      c++;
    }
  });

  // Batch-update rows for efficiency
  bankRowsToUpdate.forEach(update => {
    shB.getRange(update.row, 9, 1, 3).setValues([[update.folio, update.uuid, "Sí"]]);
  });

  for (const sheetName in accountingRowsToUpdate) {
    const sheet = ss.getSheetByName(sheetName);
    accountingRowsToUpdate[sheetName].forEach(rowNum => {
      sheet.getRange(rowNum, 20).setValue("Sí"); // Corrected column index for "Conciliado"
    });
  }

  log(`Conciliación avanzada: ${c} nuevas coincidencias.`);
  return c; // Return the count
}

function indexarMovs_(sh){
  const vals=sh.getRange(2,1,Math.max(0,sh.getLastRow()-1), sh.getLastColumn()).getValues();
  const arr=[];
  vals.forEach((r,i)=>{
    const isReconciled = r[19]; // "Conciliado" column
    if (isReconciled === "Sí") return;

    const fecha=r[0], folio=r[1], total=r[10], uuid=r[15];
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

/**
 * Calcula el ISR mensual llamando a la función específica del régimen fiscal configurado.
 * @returns {Object} Un objeto con los detalles del cálculo del ISR.
 */
function calcularISR_Mensual() {
  const regimen = getCfg("REGIMEN_FISCAL", "PM_General");
  let resultado;

  switch (regimen) {
    case "PM_General":
      resultado = calcularISR_Regimen_PM_General_();
      break;
    case "RESICO":
      resultado = calcularISR_Regimen_RESICO_();
      break;
    case "PFAE":
      resultado = calcularISR_Regimen_PFAE_();
      break;
    default:
      const msg = `El régimen fiscal "${regimen}" no es soportado o no es válido.`;
      log(msg);
      safeAlert_(msg);
      return { error: msg };
  }

  log(`Cálculo de ISR para ${regimen} finalizado.`);
  return resultado;
}

/**
 * Calcula el ISR para el régimen de Personas Morales (Régimen General).
 * Basado en coeficiente de utilidad.
 */
function calcularISR_Regimen_PM_General_(){
  const pr = periodo_();
  const coef = Number(getCfg("COEF_UTILIDAD") || CFG.COEF_UTILIDAD);
  const tasa = Number(getCfg("ISR_PM_TASA") || CFG.ISR_PM_TASA);

  const pol = SpreadsheetApp.getActive().getSheetByName("Polizas").getRange(2, 1, Math.max(0, SpreadsheetApp.getActive().getSheetByName("Polizas").getLastRow() - 1), 12).getValues();
  let ingresosAcum = 0;
  pol.forEach(r => {
    const fecha = new Date(r[0]);
    if (fecha.getFullYear() !== pr.ini.getFullYear() || fecha > pr.fin) return;
    const cta = String(r[3]);
    const haber = Number(r[7] || 0);
    if (cta.startsWith("400-")) {
      ingresosAcum += haber;
    }
  });

  const utilidadEst = ingresosAcum * coef;
  const impuesto = round2(utilidadEst * tasa);
  const logMsg = `ISR PM ${pr.label}: Ingresos Acum ${ingresosAcum} * Coef. ${coef} * Tasa ${tasa} = ${impuesto}`;
  log(logMsg);
  safeAlert_(logMsg);

  return { periodo: pr.label, ingresosAcum, coef, tasa, impuesto, regimen: "PM_General" };
}

/**
 * Calcula el ISR para el Régimen Simplificado de Confianza (RESICO).
 * Se basa en el flujo de efectivo (cobrado/pagado) y distingue entre Persona Moral y Física.
 */
function calcularISR_Regimen_RESICO_() {
  const ss = SpreadsheetApp.getActive();
  const pr = periodo_();
  const tipoPersona = getCfg("TIPO_PERSONA", "Moral");

  const shI = ss.getSheetByName("Ingresos");
  const shE = ss.getSheetByName("Egresos");

  const allIngresos = shI.getRange(2, 1, Math.max(0, shI.getLastRow() - 1), shI.getLastColumn()).getValues();
  const allEgresos = shE.getRange(2, 1, Math.max(0, shE.getLastRow() - 1), shE.getLastColumn()).getValues();

  // Filtra por flujo de efectivo (pagado) y por el periodo actual del año
  const ingresosEfectivos = allIngresos.filter(r => {
    const fecha = new Date(r[0]);
    return r[14] === 'Pagado' && fecha.getFullYear() === pr.ini.getFullYear() && fecha <= pr.fin;
  }).reduce((acc, r) => acc + toNum(r[6]), 0); // Suma el subtotal

  const egresosEfectivos = allEgresos.filter(r => {
    const fecha = new Date(r[0]);
    return r[14] === 'Pagado' && fecha.getFullYear() === pr.ini.getFullYear() && fecha <= pr.fin;
  }).reduce((acc, r) => acc + toNum(r[6]), 0); // Suma el subtotal

  let impuesto = 0;
  let tasa = 0;
  let base = 0;
  let logMsg = "";

  if (tipoPersona === 'Moral') {
    base = ingresosEfectivos - egresosEfectivos;
    tasa = 0.30; // Tasa fija para PM RESICO
    impuesto = round2(Math.max(0, base) * tasa);
    logMsg = `ISR RESICO PM ${pr.label}: (Ingresos ${ingresosEfectivos} - Egresos ${egresosEfectivos}) * ${tasa * 100}% = ${impuesto}`;
  } else { // Persona Física
    base = ingresosEfectivos;
    if (base <= 300000) tasa = 0.01;
    else if (base <= 600000) tasa = 0.011;
    else if (base <= 1000000) tasa = 0.015;
    else if (base <= 2500000) tasa = 0.02;
    else tasa = 0.025;

    impuesto = round2(base * tasa);
    logMsg = `ISR RESICO PF ${pr.label}: Ingresos Acum ${base} * Tasa ${tasa * 100}% = ${impuesto}`;
  }

  log(logMsg);
  safeAlert_(logMsg);

  return { periodo: pr.label, base, tasa, impuesto, regimen: `RESICO ${tipoPersona}` };
}

/**
 * Calcula el ISR para el régimen de Personas Físicas con Actividades Empresariales (PFAE).
 * Se basa en la utilidad (ingresos - egresos) y aplica la tarifa progresiva mensual.
 */
function calcularISR_Regimen_PFAE_() {
  const ss = SpreadsheetApp.getActive();
  const pr = periodo_();

  const shI = ss.getSheetByName("Ingresos");
  const shE = ss.getSheetByName("Egresos");

  const allIngresos = shI.getRange(2, 1, Math.max(0, shI.getLastRow() - 1), shI.getLastColumn()).getValues();
  const allEgresos = shE.getRange(2, 1, Math.max(0, shE.getLastRow() - 1), shE.getLastColumn()).getValues();

  const ingresosDelMes = allIngresos.filter(r => {
    const fecha = new Date(r[0]);
    return r[14] === 'Pagado' && fecha >= pr.ini && fecha <= pr.fin;
  }).reduce((acc, r) => acc + toNum(r[6]), 0);

  const egresosDelMes = allEgresos.filter(r => {
    const fecha = new Date(r[0]);
    return r[14] === 'Pagado' && fecha >= pr.ini && fecha <= pr.fin;
  }).reduce((acc, r) => acc + toNum(r[6]), 0);

  const base = ingresosDelMes - egresosDelMes;

  // Tarifa Mensual ISR 2023 para PFAE
  const tablaISR = [
    { limInf: 0.01, limSup: 746.04, cuotaFija: 0.00, porciento: 0.0192 },
    { limInf: 746.05, limSup: 6332.05, cuotaFija: 14.32, porciento: 0.0640 },
    { limInf: 6332.06, limSup: 11128.01, cuotaFija: 371.83, porciento: 0.1088 },
    { limInf: 11128.02, limSup: 12935.82, cuotaFija: 893.63, porciento: 0.1600 },
    { limInf: 12935.83, limSup: 15487.71, cuotaFija: 1182.88, porciento: 0.1792 },
    { limInf: 15487.72, limSup: 31236.49, cuotaFija: 1640.18, porciento: 0.2136 },
    { limInf: 31236.50, limSup: 49233.00, cuotaFija: 5004.12, porciento: 0.2352 },
    { limInf: 49233.01, limSup: 93993.90, cuotaFija: 9236.89, porciento: 0.3000 },
    { limInf: 93993.91, limSup: 125325.20, cuotaFija: 22665.17, porciento: 0.3200 },
    { limInf: 125325.21, limSup: 375975.61, cuotaFija: 32691.18, porciento: 0.3400 },
    { limInf: 375975.62, limSup: Infinity, cuotaFija: 117912.32, porciento: 0.3500 }
  ];

  let impuesto = 0;
  if (base > 0) {
    const rango = tablaISR.find(r => base >= r.limInf && base <= r.limSup);
    if (rango) {
      const excedente = base - rango.limInf;
      impuesto = round2((excedente * rango.porciento) + rango.cuotaFija);
    }
  }

  const logMsg = `ISR PFAE ${pr.label}: (Ingresos ${ingresosDelMes} - Egresos ${egresosDelMes}) = Base ${base}. Impuesto: ${impuesto}`;
  log(logMsg);
  safeAlert_(logMsg);

  return { periodo: pr.label, base, impuesto, regimen: "PFAE" };
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
    const l=sh.getLastRow(); if(l>1){ sh.getRange(2,1,l-1, sh.getLastColumn()).getValues().forEach((r,i)=>{ const u=(r[15]||"").toString(); if(u) idx[u]={sh,row:i+2}; }); }
  });
  let n=0;
  pagos.forEach(p=>{
    const rels=(p[17]||"").toString().split(/[\s,;|]+/).filter(Boolean);
    rels.forEach(u=>{ const hit=idx[u]; if(hit){ hit.sh.getRange(hit.row,15).setValue("Pagado"); n++; } });
  });
  if(n) log(`Pagos aplicados (P → EstadoPago): ${n}`);
}

/********************  IMPORTACIÓN BANCOS (CSV)  ********************/
function conciliarBancariaAvanzadaUI() {
  const count = conciliarBancariaAvanzada();
  return `Conciliación automática finalizada. Se encontraron ${count} coincidencias.`;
}

function importarEstadoDeCuentaCsv(nombreArchivo) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Bancos");
  let parentFolder;
  try {
    parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  } catch (e) {
    throw new Error("No se pudo encontrar la carpeta contenedora de este archivo.");
  }

  const files = parentFolder.getFilesByName(nombreArchivo);
  if (!files.hasNext()) {
    throw new Error(`Archivo CSV no encontrado en Google Drive: ${nombreArchivo}`);
  }

  const csvData = files.next().getBlob().getDataAsString();
  const rows = Utilities.parseCsv(csvData);

  if (rows.length < 1) {
    return "El archivo CSV está vacío.";
  }

  const headers = rows.shift().map(h => h.toLowerCase().trim());

  // Mapeo inteligente de cabeceras
  const headerMap = {};
  const mapping = {
    fecha: ['fecha', 'date'],
    descripcion: ['descripción', 'descripcion', 'description', 'memo'],
    cargo: ['cargo', 'debit', 'retiro'],
    abono: ['abono', 'credit', 'deposito'],
    importe: ['importe', 'monto', 'amount']
  };

  for (const key in mapping) {
    for (const alias of mapping[key]) {
      const index = headers.indexOf(alias);
      if (index !== -1) {
        headerMap[key] = index;
        break;
      }
    }
  }

  if (headerMap.fecha === undefined) {
    throw new Error("No se pudo encontrar una columna de 'Fecha' en el CSV.");
  }

  const processedRows = rows.map(row => {
    const fecha = new Date(row[headerMap.fecha]);
    const descripcion = row[headerMap.descripcion] || '';
    const cargo = toNum(row[headerMap.cargo]);
    const abono = toNum(row[headerMap.abono]);
    // Si hay una columna 'importe' y no hay cargo/abono, usarla.
    let importe = toNum(row[headerMap.importe]);
    if(importe !== 0 && cargo === 0 && abono === 0) {
      if (importe < 0) {
        cargo = Math.abs(importe);
      } else {
        abono = importe;
      }
    }

    return [fecha, descripcion, cargo, abono, '', '', '', '', '', '', 'No'];
  }).filter(row => row[0] instanceof Date && !isNaN(row[0])); // Filtrar filas sin fecha válida

  if (processedRows.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, processedRows.length, processedRows[0].length).setValues(processedRows);
    const message = `Importación exitosa. Se agregaron ${processedRows.length} transacciones bancarias.`;
    log(message);
    return message;
  } else {
    return "No se encontraron filas válidas para importar en el CSV.";
  }
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
/**
 * Extrae los montos de IVA trasladado, IVA retenido e ISR retenido de un nodo XML de CFDI.
 * @param {GoogleAppsScript.XML_Service.Element} root - El elemento raíz del XML.
 * @returns {{iva: number, retIva: number, retIsr: number}} - Un objeto con los totales de impuestos.
 */
function extraerImpuestos_(root) {
  let iva = 0, retIva = 0, retIsr = 0;
  const descendents = root.getDescendants();

  for (const d of descendents) {
    try {
      const el = d.asElement();
      if (!el) continue;
      const elName = el.getName();

      if (/Retencion/i.test(elName)) {
        const attrs = attrMap_(el);
        const importe = toNum(attrs.Importe || attrs.importe);
        if (attrs.Impuesto === '001' || attrs.impuesto === '001') retIsr += importe;
        if (attrs.Impuesto === '002' || attrs.impuesto === '002') retIva += importe;
      } else if (/Traslado/i.test(elName)) {
        const attrs = attrMap_(el);
        // Solo sumar IVA (002)
        if (attrs.Impuesto === '002' || attrs.impuesto === '002') {
          iva += toNum(attrs.Importe || attrs.importe);
        }
      }
    } catch (e) { /* Ignorar nodos que no son elementos */ }
  }

  // Si no se encontró IVA explícito en un nodo de Traslados, usar el cálculo simple como fallback.
  if (iva === 0 && retIva === 0 && retIsr === 0) {
    const A = attrMap_(root);
    const Subtotal = toNum(A.SubTotal || A.Subtotal);
    const Total = toNum(A.Total);
    // Este es un cálculo aproximado y puede no ser preciso si hay otros impuestos.
    iva = Total - Subtotal;
  }

  return { iva: round2(iva), retIva: round2(retIva), retIsr: round2(retIsr) };
}

function importarCFDI_(folderId, esIngreso){
  if (!folderId) throw new Error("Config: CARPETA_CFDI_* no definida");

  const pr = periodo_();
  const fol = DriveApp.getFolderById(folderId);

  // 1) sets para evitar duplicados
  const setUUID = uuidsExistentes_();
  const ss = SpreadsheetApp.getActive();
  const shI = ss.getSheetByName("Ingresos");
  const shE = ss.getSheetByName("Egresos");
  const setI = buildKeysSet_(shI);
  const setE = buildKeysSet_(shE);

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
      const impuestos = extraerImpuestos_(root);
      const origen = esIngreso? 'Emitidas':'Recibidas';

      // Hoja mensual (para auditoría de periodo)
      const rowMes=[origen, Tipo, Fecha, Serie, Folio, UUID, Emisor.Rfc||'', Emisor.Nombre||'', Receptor.Rfc||'', Receptor.Nombre||'', Moneda, Subtotal, impuestos.iva, Total, Metodo, Forma, uso, rels.join('|'), '', ''];
      registrarEnHojaMes_(pr.label, rowMes);

      // Maestros
      upsertMaestrosDesdeCFDI_(origen, Emisor, Receptor);

      // Movimiento a Ingresos/Egresos (con llave compuesta anti-dup)
      const tercero = esIngreso? (Receptor.Nombre||Receptor.Rfc) : (Emisor.Nombre||Emisor.Rfc);
      const rfc     = esIngreso? (Receptor.Rfc||"")             : (Emisor.Rfc||"");
      const concepto = `CFDI ${src.name || ''} [Tipo:${Tipo||'?'}]`;
      const relsJoined = rels.join('|');
      const rowIE = [
        Fecha, Folio, tercero, rfc, concepto,
        tasaLabel_(tasa), Subtotal, impuestos.iva, impuestos.retIva, impuestos.retIsr,
        Total, Metodo, Forma, "", "Pendiente", UUID,
        relsJoined, "", "", "No", "No"
      ];

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
  const total = Number(row[10]||0).toFixed(2); // Adjusted index for Total
  const uuid = (row[15]||'').toString().trim().toUpperCase(); // Adjusted index for UUID
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

/********************  MOTOR DE ASIENTOS BASADO EN REGLAS  ********************/
/**
 * Lee y parsea las reglas de la hoja "Asientos" para su uso en el motor de pólizas.
 * Combina las líneas de una misma regla en un solo objeto.
 * @returns {Array<Object>} Un arreglo de objetos, donde cada objeto es una regla de asiento.
 */
function getReglasAsientos_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Asientos");
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  const reglas = [];
  let currentRule = null;

  values.forEach(row => {
    const asientoID = row[0];
    if (asientoID) { // Es una nueva regla
      if (currentRule) {
        reglas.push(currentRule);
      }
      currentRule = {
        id: row[0],
        nombre: row[1],
        tipoCfdi: row[2],
        metodoPago: row[3],
        formaPago: row[4],
        condicion: row[5],
        aplicaPf: row[6],
        aplicaPm: row[7],
        debe: [row[8]],
        haber: [row[9]],
        iva: row[10],
        ret: row[11],
        tipoPoliza: row[12],
        concepto: row[13]
      };
    } else { // Es una continuación de la regla anterior
      if (currentRule) {
        if (row[8]) currentRule.debe.push(row[8]);
        if (row[9]) currentRule.haber.push(row[9]);
      }
    }
  });

  if (currentRule) {
    reglas.push(currentRule); // Agrega la última regla
  }

  return reglas;
}


/**
 * Importa las reglas de asientos desde el archivo Asientos.csv ubicado en la misma
 * carpeta que la hoja de cálculo y las vuelca en la hoja "Asientos".
 */
function importarReglasDeAsientos() {
  const ss = SpreadsheetApp.getActive();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();
  const files = parentFolder.getFilesByName("Asientos.csv");

  if (!files.hasNext()) {
    const msg = "No se encontró el archivo 'Asientos.csv' en la misma carpeta que este Google Sheet. No se pueden cargar las reglas de asientos.";
    log(msg);
    safeAlert_(msg);
    return;
  }

  const csvFile = files.next();
  const csvData = csvFile.getBlob().getDataAsString();

  const sheet = getOrCreateSheet(ss, "Asientos");
  const data = Utilities.parseCsv(csvData);

  if (data.length === 0) {
    log("El archivo 'Asientos.csv' está vacío. No se cargaron reglas.");
    return;
  }

  sheet.clear();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  formatSheet_(sheet);
  log("Reglas de asientos importadas correctamente desde Asientos.csv.");
}


/********************  REPORTERÍA AVANZADA  ********************/
function verReporteCliente() {
  const htmlOutput = generarReporteHTML_();
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Reporte Financiero para Cliente");
}

function generarReporteHTML_() {
  const ss = SpreadsheetApp.getActive();
  const analisisSheet = ss.getSheetByName("AnalisisFinanciero");
  const erSheet = ss.getSheetByName("EstadoResultados");
  const bgSheet = ss.getSheetByName("BalanceGeneral");

  // Leer datos usando getDisplayValues para conservar el formato
  const analisisData = analisisSheet.getRange("A2:D" + analisisSheet.getLastRow()).getDisplayValues();
  const erData = erSheet.getRange("A2:C" + erSheet.getLastRow()).getDisplayValues();
  const bgData = bgSheet.getRange("A2:C" + bgSheet.getLastRow()).getDisplayValues();

  const template = HtmlService.createTemplateFromFile('ReporteCliente');
  template.periodo = getCfg("PERIODO_LABEL", new Date().toLocaleDateString());
  template.analisis = analisisData;
  template.er = erData;
  template.bg = bgData;

  return template.evaluate().setWidth(900).setHeight(600);
}

function generarGuionReunion() {
  const periodo = getCfg("PERIODO_LABEL", "este periodo");
  const nombreCliente = SpreadsheetApp.getActive().getName().replace("Contabilidad_", "");
  const nombreDoc = `Guión de Reunión Financiera - ${nombreCliente} - ${periodo}`;

  const textoGuion = generarTextoGuion_(periodo);

  const doc = DocumentApp.create(nombreDoc);
  doc.getBody().setText(textoGuion);

  const url = doc.getUrl();
  const msg = `Se ha creado el documento con el guión de la reunión. Puedes acceder a él aquí: ${url}`;
  safeAlert_(msg);
  log(msg);
}

function generarTextoGuion_() {
  const ss = SpreadsheetApp.getActive();
  const analisisSheet = ss.getSheetByName("AnalisisFinanciero");
  const analisisData = analisisSheet.getRange("A2:D" + analisisSheet.getLastRow()).getDisplayValues();

  let guion = `Guión para la Reunión de Análisis Financiero\n`;
  guion += `========================================\n\n`;

  guion += `Agenda:\n`;
  guion += `1. Resumen de Desempeño del Periodo.\n`;
  guion += `2. Análisis de Indicadores Clave (KPIs).\n`;
  guion += `3. Discusión de Puntos Relevantes.\n`;
  guion += `4. Definición del Plan de Acción.\n\n`;

  guion += `Análisis de Indicadores Clave (KPIs)\n`;
  guion += `------------------------------------\n`;

  let currentGroup = "";
  analisisData.forEach(row => {
    const [grupo, indicador, valor, interpretacion] = row;
    if (grupo !== currentGroup) {
      guion += `\n**${grupo.toUpperCase()}**\n`;
      currentGroup = grupo;
    }
    guion += `- **${indicador}:** ${valor}\n`;
    guion += `  *Interpretación:* ${interpretacion}\n`;
  });

  guion += `\nTemas a Discutir\n`;
  guion += `------------------\n`;
  guion += `1. ¿Qué factores (internos/externos) explican los resultados de rentabilidad de este periodo?\n`;
  guion += `2. La liquidez de la empresa es [adecuada/preocupante]. ¿Tenemos suficiente efectivo para las operaciones de los próximos 3 meses?\n`;
  guion += `3. Los días de cartera son de [X días]. ¿Estamos conformes con este plazo o necesitamos optimizar la cobranza?\n`;
  guion += `4. El nivel de endeudamiento es [bajo/alto]. ¿Cómo se alinea esto con nuestra estrategia de crecimiento?\n\n`;

  guion += `Plan de Acción\n`;
  guion += `--------------\n`;
  guion += `- Tarea 1: [Definir responsable y fecha]\n`;
  guion += `- Tarea 2: [Definir responsable y fecha]\n`;
  guion += `- Tarea 3: [Definir responsable y fecha]\n`;

  return guion;
}


/********************  MANUAL_RECONCILIATION  ********************/
function getUnmatchedBankTxs() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Bancos");
  const data = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), sh.getLastColumn()).getValues();
  const unmatched = [];
  data.forEach((row, index) => {
    if (row[10] !== 'Sí') {
      unmatched.push({
        rowNum: index + 2, // 1-based index for sheet, plus 1 for header
        fecha: row[0],
        descripcion: row[1],
        cargo: row[2],
        abono: row[3]
      });
    }
  });
  return unmatched;
}

function getUnmatchedAccountTxs() {
  const ss = SpreadsheetApp.getActive();
  const unmatched = [];
  const sheets = ["Ingresos", "Egresos"];

  sheets.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    const data = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), sh.getLastColumn()).getValues();
    data.forEach((row, index) => {
      // Assuming 'Conciliado' is in column 17 (index 16) for both sheets
      if (row[19] !== 'Sí') { // Corregido a columna 20 (índice 19) que es 'Conciliado'
        unmatched.push({
          rowNum: index + 2,
          sheet: sheetName,
          fecha: row[0],
          tercero: row[2],
          total: row[10]
        });
      }
    });
  });
  return unmatched;
}

function reconcileManualMatch_(bankTxIds, accountTxIds) {
  const ss = SpreadsheetApp.getActive();
  const shBancos = ss.getSheetByName("Bancos");

  bankTxIds.forEach(id => {
    const rowNum = parseInt(id.split('|')[0]);
    shBancos.getRange(rowNum, 11).setValue("Sí"); // Columna 'Conciliado'
  });

  const accountUpdates = {};
  accountTxIds.forEach(id => {
    const [rowNum, sheetName] = id.split('|');
    if (!accountUpdates[sheetName]) {
      accountUpdates[sheetName] = [];
    }
    accountUpdates[sheetName].push(parseInt(rowNum));
  });

  for (const sheetName in accountUpdates) {
    const sh = ss.getSheetByName(sheetName);
    accountUpdates[sheetName].forEach(rowNum => {
      sh.getRange(rowNum, 20).setValue("Sí"); // Corregido a columna 20 (índice 19)
    });
  }

  const total = bankTxIds.length + accountTxIds.length;
  log(`${total} items marcados como conciliados manualmente.`);
  return `${total} items conciliados exitosamente.`;
}

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
