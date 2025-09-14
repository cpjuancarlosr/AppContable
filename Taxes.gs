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

  // Leer la tabla de ISR desde la hoja de cálculo
  const tablaSheet = SpreadsheetApp.getActive().getSheetByName("TaxTables");
  if (!tablaSheet) {
    const msg = "La hoja 'TaxTables' no se encuentra. No se puede calcular el ISR de PFAE.";
    log(msg);
    safeAlert_(msg);
    return { error: msg };
  }

  const tablaData = tablaSheet.getRange(2, 1, tablaSheet.getLastRow() - 1, 5).getValues();
  const tablaISR = tablaData.map(row => ({
    regimen: row[0],
    limInf: parseFloat(row[1]),
    limSup: parseFloat(row[2]),
    cuotaFija: parseFloat(row[3]),
    porciento: parseFloat(row[4])
  })).filter(r => r.regimen === "PFAE_2023"); // Filtrar por el régimen adecuado

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
