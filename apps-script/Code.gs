/** Configurar SPREADSHEET_ID en Propiedades del script, no en el código. */
const FUENTES = {
  inventario: { hoja: 'Maria Grande', columnas: 21, encabezado: 1 },
  movimientos: { hoja: 'Consolidado', columnas: 12, encabezado: 1 },
  enviados: { hoja: 'Enviados sin informar 2020', columnas: 20, encabezado: 1 }
};
const MAX_FILAS = 20000;

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Transformadores · María Grande');
}
function incluir(nombre) { return HtmlService.createHtmlOutputFromFile(nombre).getContent(); }

/** Lectura deliberadamente sin escritura mientras se validan los procesos VBA. */
function obtenerDatos(tipo) {
  const fuente = FUENTES[tipo];
  if (!fuente) throw new Error('Sección desconocida.');
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Falta configurar SPREADSHEET_ID en las propiedades del script.');
  const libro = SpreadsheetApp.openById(id);
  const hoja = libro.getSheetByName(fuente.hoja);
  if (!hoja) throw new Error('No se encontró la hoja «' + fuente.hoja + '».');
  const ultima = hoja.getLastRow();
  if (!ultima) return { encabezados: [], filas: [], truncado: false };
  const valores = hoja.getRange(1, 1, Math.min(ultima, MAX_FILAS), fuente.columnas).getDisplayValues();
  const encabezados = valores[fuente.encabezado - 1].map((v, i) => v || 'Columna ' + (i + 1));
  const filas = valores.slice(fuente.encabezado).map((valores, i) => ({ fila: i + fuente.encabezado + 1, valores: valores }))
    .filter(r => r.valores.some(v => v !== ''));
  return { encabezados, filas, truncado: ultima > MAX_FILAS, limite: MAX_FILAS };
}
