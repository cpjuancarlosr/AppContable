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
