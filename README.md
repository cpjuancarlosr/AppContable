# AppContable

Sistema contable para Google Sheets.

## Características
- Gestión de periodos contables desde la hoja **Periodos** con opciones de abrir y cerrar.
- Acceso a configuración mediante caché para minimizar lecturas de la hoja.
- Registro de logs con respaldo cuando la hoja `Logs` no está disponible.

## Desarrollo
El código principal se encuentra en `Code.gs` y el frontend en `index.html`. Para verificar la sintaxis del backend:

```bash
node --check Code.gs
```
