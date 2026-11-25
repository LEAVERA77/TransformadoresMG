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

/** Ingreso: usar exclusivamente sobre una copia de prueba hasta validar con el usuario. */
function registrarIngreso(datos) {
  if (!datos || typeof datos !== 'object') throw new Error('Faltan los datos del ingreso.');
  function texto(campo, obligatorio) {
    const v = String(datos[campo] == null ? '' : datos[campo]).trim();
    if (obligatorio && !v) throw new Error('Completá el campo ' + campo + '.');
    if (v.length > 300 || /^[=+@\-]/.test(v)) throw new Error('Valor no permitido en ' + campo + '.');
    return v;
  }
  const fecha = texto('fecha', true), potencia = texto('potencia', true), relacion = texto('relacion', true);
  const marca = texto('marca', true), numero = texto('numero', true), origen = texto('origen', true);
  const destino = texto('destino', true), motivo = texto('motivo', true), estado = texto('estado', true);
  const ubicacion = texto('ubicacion', false);
  if (!/^\d+(?:[.,]\d+)?$/.test(potencia) || Number(potencia.replace(',', '.')) <= 0) throw new Error('La potencia debe ser un número positivo.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Ingresá una fecha válida.');
  const partes = fecha.split('-').map(Number), fechaObj = new Date(partes[0], partes[1]-1, partes[2], 12);
  if (fechaObj.getFullYear() !== partes[0] || fechaObj.getMonth() !== partes[1]-1 || fechaObj.getDate() !== partes[2]) throw new Error('La fecha no es válida.');
  if (fecha > Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')) throw new Error('La fecha no puede ser futura.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) throw new Error('Falta configurar SPREADSHEET_ID.');
    const libro = SpreadsheetApp.openById(id);
    const inv = libro.getSheetByName('Maria Grande'), mov = libro.getSheetByName('Mov'), cons = libro.getSheetByName('Consolidado');
    if (!inv || !mov || !cons) throw new Error('Falta alguna hoja: Maria Grande, Mov o Consolidado.');
    if (inv.getLastRow() > MAX_FILAS) throw new Error('El inventario excede el límite de revisión de duplicados; no se guardó el ingreso.');
    const normalizar = v => String(v).trim().toLocaleUpperCase();
    if (inv.getLastRow() >= 2) {
      const claves = inv.getRange(2, 10, inv.getLastRow()-1, 2).getDisplayValues();
      if (claves.some(r => normalizar(r[0]) === normalizar(marca) && normalizar(r[1]) === normalizar(numero))) throw new Error('Ya existe un transformador con esa marca y número.');
    }
    const estadoNormal = estado.toLocaleUpperCase();
    const matricula = estadoNormal === 'NUEVO' ? '0' : (['REPARADO','USADO'].includes(estadoNormal) ? '5' : '');
    const movimiento = [fechaObj, Number(potencia.replace(',', '.')), relacion, marca, numero, origen, destino, motivo, destino, estado, matricula, ''];
    const inventario = ['María Grande','','','',destino,'','',Number(potencia.replace(',', '.')),relacion,marca,numero,motivo,estado,ubicacion,'Actualizado el ' + Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm')];
    // Filas nuevas, verificadas; el bloqueo coordina solicitudes de esta aplicación, no ediciones manuales.
    const operaciones = [
      {hoja:mov, fila:Math.max(6,mov.getLastRow()+1), valores:movimiento},
      {hoja:cons, fila:Math.max(2,cons.getLastRow()+1), valores:movimiento},
      {hoja:inv, fila:Math.max(2,inv.getLastRow()+1), valores:inventario}
    ];
    // Ampliar si es necesario antes de obtener los rangos.
    operaciones.forEach(o => {const faltan=o.fila-o.hoja.getMaxRows(); if(faltan>0)o.hoja.insertRowsAfter(o.hoja.getMaxRows(),faltan);});
    operaciones.forEach(o => {o.rango=o.hoja.getRange(o.fila,1,1,o.valores.length); if(o.rango.getValues()[0].some(v=>v!=='')) throw new Error('La fila de destino ya no está vacía.');});
    const escritos=[];
    try {
      operaciones.forEach(o => {escritos.push(o);o.rango.setValues([o.valores]);});
      operaciones.slice(0,2).forEach(o => o.hoja.getRange(o.fila,1).setNumberFormat('dd/MM/yyyy'));
      SpreadsheetApp.flush();
    } catch(e) {
      let fallo=false; escritos.reverse().forEach(o=>{try{o.rango.clearContent();}catch(_){fallo=true;}});
      throw new Error('No se completó el ingreso. ' + (fallo ? 'ATENCIÓN: la reversión falló; revisá las tres hojas antes de reintentar. ' : 'Se intentó revertir la operación; verificá las tres hojas antes de reintentar. ') + e.message);
    }
    return { mensaje:'Ingreso registrado. Verificá el inventario y los movimientos.', filaInventario:operaciones[2].fila };
  } finally { lock.releaseLock(); }
}

/** Opciones sugeridas: valores existentes, sin impedir otros valores. */
function obtenerOpcionesIngreso() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Falta configurar SPREADSHEET_ID.');
  const libro = SpreadsheetApp.openById(id);
  const inv = libro.getSheetByName('Maria Grande'), cons = libro.getSheetByName('Consolidado');
  if (!inv || !cons) throw new Error('Faltan hojas para cargar las opciones.');
  if (inv.getLastRow() > MAX_FILAS || cons.getLastRow() > MAX_FILAS) throw new Error('Una hoja supera el límite de ' + MAX_FILAS + ' filas para cargar opciones.');
  const listas = {potencia:[], relacion:[], marca:[], origen:[], destino:['Almacén María Grande','Almacén Bovril','Almacén Santa Elena'], motivo:[], estado:['Nuevo','Reparado','Usado'], ubicacion:[]};
  function agregar(clave, valor) { const v = String(valor || '').trim(); if (v && v.length <= 300 && !/^[=+@\-]/.test(v)) listas[clave].push(v); }
  if (inv.getLastRow() > 1) inv.getRange(2,1,inv.getLastRow()-1,15).getDisplayValues().forEach(r => {
    agregar('potencia',r[7]); agregar('relacion',r[8]); agregar('marca',r[9]); agregar('motivo',r[11]); agregar('estado',r[12]); agregar('ubicacion',r[13]);
  });
  if (cons.getLastRow() > 1) cons.getRange(2,1,cons.getLastRow()-1,10).getDisplayValues().forEach(r => {
    agregar('potencia',r[1]); agregar('relacion',r[2]); agregar('marca',r[3]); agregar('origen',r[5]); agregar('destino',r[6]); agregar('motivo',r[7]); agregar('estado',r[9]);
  });
  Object.keys(listas).forEach(k => {
    const encontrados = new Map(); listas[k].forEach(v => { const clave=v.toLocaleUpperCase(); if(!encontrados.has(clave)) encontrados.set(clave,v); });
    listas[k] = Array.from(encontrados.values()).sort((a,b)=>a.localeCompare(b,'es',{numeric:true}));
  });
  return listas;
}
