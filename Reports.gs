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
