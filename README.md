# Transformadores María Grande

Primera etapa de migración del libro Excel a una aplicación web con Google Apps Script y Google Sheets. El archivo `.xlsm` original permanece intacto en este repositorio y es el respaldo.

## Alcance actual

- Consulta y búsqueda del inventario (`Maria Grande`), movimientos históricos (`Consolidado`) y enviados sin informar (`Enviados sin informar 2020`).
- Paginación en el navegador; actualización manual y aviso si la hoja excede 20.000 filas.
- **Solo lectura**: no crea ni modifica datos. No reproduce todavía los formularios, altas, bajas, movimientos, impresiones, validaciones ni otras automatizaciones VBA. Deben analizarse y probarse antes de habilitar escritura; no utilizar esta versión como reemplazo integral del Excel.
- Se muestran hasta 21 columnas del inventario, 12 del consolidado y 20 de enviados. Los valores se muestran tal como aparecen en Sheets; no se importan controles ni macros VBA.

## Preparación de datos

1. Conservar una copia del Excel original. Subir una copia a Google Drive y abrirla con Google Sheets; crear una copia en formato Google Sheets (Archivo → Guardar como Hojas de cálculo de Google, si está disponible).
2. Verificar nombres de hojas, cantidades de registros, códigos, fechas, fórmulas y campos en la copia antes de usarla. Las macros VBA no se convierten. Si la conversión falla por límites de celdas o formato excesivo, preparar una copia de trabajo con solo las filas utilizadas; no modificar el respaldo.
3. Copiar el ID de la planilla desde su URL (`/spreadsheets/d/ID/edit`). No publicar el ID ni datos sensibles si el repositorio es público.

## Despliegue

1. Crear un proyecto independiente en https://script.google.com. Copiar `Code.gs`, `Index.html`, `Styles.html` y `Script.html` al proyecto (los tres últimos como archivos HTML). Activar en Configuración del proyecto la visualización del archivo de manifiesto y copiar el contenido de `appsscript.json`.
2. En Configuración del proyecto → Propiedades del script, crear `SPREADSHEET_ID` con el ID de la copia en Sheets.
3. Implementar → Nueva implementación → Aplicación web. Elegir cuidadosamente quién ejecuta y quién tiene acceso. Si se ejecuta como propietario, cualquier usuario autorizado a acceder a la aplicación puede consultar los datos del propietario; restringir el acceso a usuarios de confianza. Si se ejecuta como usuario, cada uno necesita acceso de lectura a la planilla y autorizar el script. No elegir acceso público para datos internos.
4. Autorizar, abrir la URL de implementación y comparar los datos con el Excel. Para actualizar código después de los cambios, crear una nueva versión de la implementación.

## Pendientes antes de uso operativo

Relevar los cinco formularios VBA y eventos del libro; definir reglas de negocio y permisos; acordar identificadores, historial y concurrencia; implementar y probar flujos de altas, bajas, movimientos, consultas e informes con copias de datos. Validar con usuarios y respaldar regularmente. La aplicación aún no escribe en Sheets ni genera commits automáticos.
