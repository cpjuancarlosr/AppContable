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

function buildKeywordMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName("CatCuentas");
  const data = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), 8).getValues();
  const keywordMap = {};

  data.forEach(row => {
    const accountCode = row[0];
    const keywords = (row[7] || "").split(',');
    keywords.forEach(kw => {
      const trimmedKw = kw.trim().toLowerCase();
      if (trimmedKw) {
        keywordMap[trimmedKw] = accountCode;
      }
    });
  });

  return keywordMap;
}

function findAccountByKeywords_(description, keywordMap) {
  if (!description) return null;
  const descLower = description.toLowerCase();
  for (const keyword in keywordMap) {
    if (descLower.includes(keyword)) {
      return keywordMap[keyword];
    }
  }
  return null;
}
