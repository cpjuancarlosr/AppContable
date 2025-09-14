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
