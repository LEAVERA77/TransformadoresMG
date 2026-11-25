# Prueba de ingreso de transformadores

**Versión experimental: usar una copia de la planilla, nunca datos operativos.** No probada aún en Google Apps Script. Hacer una copia de respaldo adicional. Mantener el acceso a la aplicación restringido: todo usuario con acceso a una aplicación ejecutada como propietario puede registrar ingresos con los permisos del propietario.

## Instalación para probar

1. En Google Sheets crear una copia de la planilla actual y copiar el ID de la copia.
2. Preferentemente crear un proyecto Apps Script de prueba independiente. Copiar todos los archivos de `apps-script`; en Configuración → Propiedades del script establecer SPREADSHEET_ID al ID de la copia. Si se reutiliza el proyecto existente, cambiar esta propiedad afecta todas sus implementaciones y debe restaurarse luego.
3. Implementar como aplicación web solo para el usuario de prueba; autorizar permiso `spreadsheets` (lectura y escritura). No permitir acceso público.
4. Usar un equipo ficticio, sin datos reales. Completar ingreso y comprobar las nuevas filas en `Maria Grande`, `Mov` y `Consolidado`; comprobar almacén seleccionado, campos, fecha y matrícula. Probar duplicado marca+número, fechas futuras y datos faltantes; no deben insertarse filas. Probar cada almacén y los tres estados.
5. Ante error, cierre o timeout, verificar las tres hojas antes de reintentar; no asumir que la operación no se guardó. El bloqueo coordina llamadas de este proyecto, no ediciones directas en Sheets ni otros proyectos. El intento de reversión no es una transacción atómica y puede fallar.

## Reglas implementadas y diferencias

El destino seleccionado es la ubicación del inventario y también el destino/ubicación de la entrada en los movimientos. Ubicación física es opcional y se guarda aparte (columna N del inventario). El motivo se guarda en columna L y en el historial. Matrícula Nuevo=0, Reparado/Usado=5. Se rechazan duplicados por marca y número (comparación sin mayúsculas y espacios extremos). Se permiten los tres almacenes acordados; fecha válida no futura; potencia positiva. Las filas nuevas se agregan luego de la última fila utilizada en cada hoja, que puede incluir listas auxiliares. No se trasladaron formatos VBA, opciones dinámicas, ni otros flujos. No se modifican fórmulas ni el Excel del repositorio.

El resto de funciones del VBA sigue pendiente. Validar con usuarios antes de utilizar con datos reales.
