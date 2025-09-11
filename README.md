# AppContable: Sistema Contable para Google Sheets

Este proyecto implementa un sistema contable completo directamente en Google Sheets, utilizando Google Apps Script. Está diseñado principalmente para la contabilidad en México, con soporte para importación de CFDI, cálculo de impuestos como IVA e ISR, y generación de reportes financieros.

## Características Principales

- **Instalación Automatizada:** Un menú permite instalar y configurar todas las hojas de cálculo, carpetas de Drive y configuraciones necesarias con un solo clic.
- **Interfaz Web Integrada:** El sistema cuenta con una interfaz de usuario web (sidebar o standalone) para una interacción más amigable.
- **Importación de CFDI:** Importa masivamente facturas XML (CFDI 3.3 y 4.0) desde una carpeta de Google Drive, extrayendo todos los datos relevantes y clasificándolos en ingresos y egresos.
- **Generación de Pólizas Contables:** Crea automáticamente los asientos contables (pólizas) a partir de los registros de ingresos y egresos.
- **Estados Financieros:** Calcula y presenta en tiempo real el Estado de Resultados, Balance General, Balanza de Comprobación y KPIs financieros.
- **Cálculo de Impuestos (MX):** Módulos específicos para el cálculo de IVA mensual, ISR para Personas Morales y la generación del archivo para la DIOT.
- **Conciliación Bancaria:** Herramienta para cruzar movimientos bancarios con los registros contables.
- **Exportación y Envío:** Permite exportar los reportes a PDF y enviarlos por correo electrónico.

## Cómo Usar el Sistema

### 1. Instalación

1.  Abre el editor de Google Apps Script en tu hoja de cálculo de Google.
2.  Copia el contenido de `Code.gs` y `index.html` en sus respectivos archivos dentro del editor de Apps Script.
3.  Guarda los cambios.
4.  Recarga la hoja de cálculo. Aparecerá un nuevo menú con el nombre "📘 Consultoría Contable".
5.  Desde el menú, selecciona **⚙️ Instalar/Actualizar Sistema MX**. Esto creará todas las hojas y carpetas necesarias.

### 2. Despliegue como Web App (Opcional pero recomendado)

Para usar la interfaz web, debes desplegar el script:

1.  En el editor de Apps Script, haz clic en **Deploy > New deployment**.
2.  Selecciona el tipo de despliegue **Web app**.
3.  En la configuración:
    -   **Description:** `Versión 1.0 del sistema contable`.
    -   **Execute as:** `Me (your email)`.
    -   **Who has access:** `Only myself` (o quien desees que tenga acceso).
4.  Haz clic en **Deploy**.
5.  Copia la URL de la Web app. Puedes usarla para acceder a la interfaz.

### 3. Flujo de Trabajo Típico

1.  **Configuración:** Llena la información en la hoja `Config`, como los correos de destino, coeficientes, etc.
2.  **Importar:** Sube tus archivos XML a las carpetas `CFDI_EMITIDAS` y `CFDI_RECIBIDAS` en Google Drive. Luego, usa los botones en el menú o la interfaz web para importarlos.
3.  **Contabilizar:** Genera las pólizas y recalcula los estados financieros.
4.  **Revisar:** Analiza los resultados en las hojas de `Balanza`, `EstadoResultados`, `BalanceGeneral` y `KPIs`.
5.  **Impuestos:** Calcula los impuestos del periodo.
6.  **Cerrar Periodo:** Una vez finalizado el trabajo del mes, puedes cerrar el periodo desde el menú para evitar modificaciones accidentales.

## Desarrollo

El código fuente se divide en:

-   `Code.gs`: Contiene toda la lógica del backend (Google Apps Script).
-   `index.html`: Define la interfaz de usuario para la Web App.

**Nota sobre la verificación de sintaxis:** El comando `node --check` no funciona con archivos `.gs` ya que es una extensión específica de Google. El código está escrito en JavaScript moderno (ES5/ES6) y su validez se comprueba directamente en el editor de Google Apps Script.
