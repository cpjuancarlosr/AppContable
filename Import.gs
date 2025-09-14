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

function importarCFDI_(folderId, esIngreso){
  if (!folderId) throw new Error("Config: CARPETA_CFDI_* no definida");

  const pr = periodo_();
  const fol = DriveApp.getFolderById(folderId);
  const keywordMap = buildKeywordMap_(); // Build the map once

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

      // --- Nueva Lógica de Mapeo de Cuentas ---
      let cuentaContable = findAccountByKeywords_(concepto, keywordMap);
      // Si no se encuentra por keyword, se deja vacío para mapeo manual posterior por proveedor/cliente.

      const rowIE = [
        Fecha, Folio, tercero, rfc, concepto,
        tasaLabel_(tasa), Subtotal, impuestos.iva, impuestos.retIva, impuestos.retIsr,
        Total, Metodo, Forma, "", "Pendiente", UUID,
        relsJoined, cuentaContable || "", "", "No", "No" // Se inserta la cuenta encontrada
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

/********************  IMPORTACIÓN BANCOS (CSV)  ********************/
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
