//== File: reports.gs (matrices IVA/ISR, Balanza, ER, Flujo)

/**
 * Recalculates all financial reports and KPIs by placing formulas in the 'Inicio' sheet.
 */
function recalcularEstadosFinancieros() {
  const ss = SpreadsheetApp.getActive();
  const shInicio = ss.getSheetByName("Inicio");

  if (!shInicio) {
    throw new Error("La hoja 'Inicio' no se encuentra.");
  }

  // Clear previous reports area (e.g., from row 5 downwards)
  shInicio.getRange("A5:Z100").clearContent();

  generarBalanza_(shInicio);
  generarEstadoResultados_(shInicio);
  generarFlujoEfectivo_(shInicio); // Placeholder for now

  SpreadsheetApp.getUi().alert("Los reportes y KPIs en la hoja 'Inicio' han sido actualizados.");
}

/**
 * Generates the trial balance (Balanza de Comprobación) using formulas.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh The sheet to place the report on.
 */
function generarBalanza_(sh) {
  sh.getRange("A5").setValue("Balanza de Comprobación");
  sh.getRange("A5:E5").merge().setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange("A6:E6").setValues([["Cuenta", "Nombre", "Debe", "Haber", "Saldo Final"]]).setFontWeight('bold');

  // Formula to get unique accounts from Polizas
  const formulaCuentas = `=SORT(UNIQUE(Polizas!D2:D))`;
  sh.getRange("A7").setFormula(formulaCuentas);

  // Formula to get account names
  const formulaNombres = `=ARRAYFORMULA(IF(A7:A<>"", IFERROR(VLOOKUP(A7:A, Cuentas!A:B, 2, FALSE), "N/A"), ""))`;
  sh.getRange("B7").setFormula(formulaNombres);

  // Formula for Debits
  const formulaDebe = `=ARRAYFORMULA(IF(A7:A<>"", SUMIF(Polizas!D:D, A7:A, Polizas!H:H), ""))`;
  sh.getRange("C7").setFormula(formulaDebe);

  // Formula for Credits
  const formulaHaber = `=ARRAYFORMULA(IF(A7:A<>"", SUMIF(Polizas!D:D, A7:A, Polizas!I:I), ""))`;
  sh.getRange("D7").setFormula(formulaHaber);

  // Formula for Final Balance
  const formulaSaldo = `=ARRAYFORMULA(IF(A7:A<>"", IFERROR(VLOOKUP(A7:A, Cuentas!A:C, 3, FALSE)) * (C7:C - D7:D), C7:C - D7:D))`;
   sh.getRange("E7").setFormula(formulaSaldo);
}

/**
 * Generates the Income Statement (Estado de Resultados) using formulas.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh The sheet to place the report on.
 */
function generarEstadoResultados_(sh) {
  sh.getRange("G5").setValue("Estado de Resultados");
  sh.getRange("G5:H5").merge().setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange("G6:H6").setValues([["Concepto", "Monto"]]).setFontWeight('bold');

  const data = [
    ["Ingresos", `=SUMIF(Balanza!B:B, "Ingresos", Balanza!E:E)`],
    ["- Costo de Ventas", `=SUMIF(Balanza!B:B, "Costo de Ventas", Balanza!E:E)`],
    ["= Utilidad Bruta", "=H7+H8"],
    ["- Gastos de Operación", `=SUMIF(Balanza!B:B, "Gastos", Balanza!E:E)`],
    ["= Utilidad de Operación", "=H9+H10"],
    ["- Impuestos", `=SUMIF(Balanza!B:B, "Impuestos", Balanza!E:E)`],
    ["= Utilidad Neta", "=H11+H12"]
  ];

  sh.getRange("G7:H13").setValues(data);
  sh.getRange("H7:H13").setNumberFormat("$#,##0.00;($#,##0.00)");
}


/**
 * Generates the Cash Flow Statement (placeholder).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh The sheet to place the report on.
 */
function generarFlujoEfectivo_(sh) {
  sh.getRange("J5").setValue("Estado de Flujo de Efectivo (Simplificado)");
  sh.getRange("J5:K5").merge().setFontWeight('bold').setHorizontalAlignment('center');

  // This is a simplified placeholder. A real cash flow statement is much more complex.
  sh.getRange("J6:K6").setValues([["Concepto", "Monto"]]).setFontWeight('bold');
  sh.getRange("J7:K9").setValues([
      ["Saldo Inicial de Efectivo", "0.00"],
      ["+ Entradas / - Salidas de Efectivo", `=SUMIF(Balanza!B:B, "Bancos", Balanza!E:E)`],
      ["= Saldo Final de Efectivo", "=K7+K8"]
  ]);
}
