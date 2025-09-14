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

/********************  INSTALACIÓN  ********************/
function setupSistemaMX(){
  const ss = SpreadsheetApp.getActive();
  const hojas = [
    "Config","Clientes","Proveedores","CatCuentas","Ingresos","Egresos","Bancos",
    "Polizas","Mayor","Balanza","EstadoResultados","BalanceGeneral","KPIs","AnalisisFinanciero",
    "Plantillas","PagosImpuestos","DIOT","Logs","Dashboard", "Asientos", "Periodos", "TaxTables"
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
  prepararTaxTables();
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
  sh.getRange(1,1,1,8).setValues([["Codigo","Nombre","Tipo","SAT","Nivel","Padre","Naturaleza", "PalabrasClave"]]).setFontWeight("bold");
  const base = [
    ["100-000","Activo Circulante","Activo","",1,"","Deudora", ""],
    ["110-100","Bancos","Activo","",2,"100-000","Deudora", ""],
    ["110-110","Banco MXN","Activo","",3,"110-100","Deudora", ""],
    ["120-000","Clientes","Activo","",2,"100-000","Deudora", ""],
    ["130-000","Inventarios","Activo","",2,"100-000","Deudora", "inventario,materia prima"],
    ["200-000","Pasivo Corto Plazo","Pasivo","",1,"","Acreedora", ""],
    ["210-100","Proveedores","Pasivo","",2,"200-000","Acreedora", ""],
    ["240-200","IVA Trasladado 16%","Impuesto","",3,"200-000","Acreedora", ""],
    ["240-210","IVA Trasladado 8%","Impuesto","",3,"200-000","Acreedora", ""],
    ["240-300","IVA Acreditable 16%","Impuesto","",3,"200-000","Deudora", ""],
    ["240-310","IVA Acreditable 8%","Impuesto","",3,"200-000","Deudora", ""],
    ["240-400","ISR Retenido","Impuesto","",2,"200-000","Acreedora", ""],
    ["300-000","Capital","Capital","",1,"","Acreedora", ""],
    ["400-000","Ingresos","Resultado","",1,"","Acreedora", "ingreso,venta"],
    ["500-000","Costo/Ventas","Resultado","",1,"","Deudora", ""],
    ["510-000","Gastos","Resultado","",1,"","Deudora", "gasto,comision,gasolina"],
    ["520-000","Gastos Honorarios","Resultado","",2,"510-000","Deudora", "honorarios,consultoria"]
  ];
  sh.getRange(2,1,base.length,8).setValues(base);
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

function prepararTaxTables() {
  const sh = SpreadsheetApp.getActive().getSheetByName("TaxTables") || SpreadsheetApp.getActive().insertSheet("TaxTables");
  sh.clear();
  sh.getRange("A1:E1").setValues([["Regimen", "LimiteInferior", "LimiteSuperior", "CuotaFija", "PorcentajeExcedente"]]).setFontWeight("bold");

  const data = [
    ["PFAE_2023", 0.01, 746.04, 0.00, 0.0192],
    ["PFAE_2023", 746.05, 6332.05, 14.32, 0.0640],
    ["PFAE_2023", 6332.06, 11128.01, 371.83, 0.1088],
    ["PFAE_2023", 11128.02, 12935.82, 1182.88, 0.1600],
    ["PFAE_2023", 12935.83, 15487.71, 1640.18, 0.1792],
    ["PFAE_2023", 15487.72, 31236.49, 5004.12, 0.2136],
    ["PFAE_2023", 31236.50, 49233.00, 9236.89, 0.2352],
    ["PFAE_2023", 49233.01, 93993.90, 22665.17, 0.3000],
    ["PFAE_2023", 93993.91, 125325.20, 32691.18, 0.3200],
    ["PFAE_2023", 125325.21, 375975.61, 117912.32, 0.3400],
    ["PFAE_2023", 375975.62, 999999999, 117912.32, 0.3500]
  ];

  sh.getRange(2, 1, data.length, 5).setValues(data).setNumberFormat("0.00");
  sh.autoResizeColumns(1, 5);
  formatSheet_(sh);
  log("Hoja de tablas de impuestos preparada.");
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
