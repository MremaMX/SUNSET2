const CONFIG = {
  FUENTE_HOJA: 'JOIA Paraíso Resort',
  COL_HAB: 2,
  COL_CHECKOUT: 8,

  HOJA_DESTINO: 'COCKTAIL SUNSET SPAM',
  FILA_LINK: 3,
  COL_LINK: 3,
  COL_LINK_ANCHO: 8,

  FILA_FECHA: 6,
  COL_INICIO: 2, // B
  ROWS_PER_BLOCK: 29,
  GAP_ROWS_BETWEEN_TABLES: 3,

  HEADER_FONT_SIZE: 13,
  HAB_FONT_SIZE: 13,
  FECHA_FONT_SIZE: 25,
  CHECKBOX_FONT_SIZE: 15,
  YA_ENVIADO_WIDTH: 130,

  COLOR_BG_GENERAL: '#FCFCE3',
  COLOR_TOP_BAND: '#134f5c',
  COLOR_FECHA: '#0C2C55',
  COLOR_HEADER_HAB: '#009690',
  COLOR_HEADER_ENVIADO: '#296374',
  COLOR_HEADER_YAENVIADO: '#629FAD',
  COLOR_HEADER_SOCIOS: '#EDEDCE',

  URL_WEB_APP: 'https://script.google.com/a/macros/iberostar.com/s/AKfycbxyGOheNRg1h1DLUKEy6uY6S_OCmg0OdILhf6Jdc53gpP48MDfsi50K4NxbjeAAk_oB/exec'
};

/* =========================
   MENU
========================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('COCKTAIL SUNSET')
    .addItem('Agregar tabla del día', 'actualizarTablaActual') // nombre que quieres conservar
    .addItem('Abrir flyer', 'abrirFlyer')
    .addToUi();
}

/* =========================
   PRINCIPAL
========================= */
function actualizarTablaActual() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No hay hoja activa.');

    const sheet = ss.getSheetByName(CONFIG.HOJA_DESTINO) || ss.getActiveSheet();

    prepararHojaSiHaceFalta(sheet);

    const linkValue = sheet.getRange(CONFIG.FILA_LINK, CONFIG.COL_LINK).getValue();
    const id = extraerId(String(linkValue || ''));
    if (!id) throw new Error('Pega un link o ID válido en C3.');

    const habs = obtenerHabitaciones(id);
    if (!habs.length) throw new Error('No se encontraron habitaciones válidas.');

    const fechaHoyTxt = formatearFecha(new Date());

    // Nueva tabla arriba
    crearNuevoBloqueArriba(sheet, habs, fechaHoyTxt);

    // Reaplicar formato a TODAS para que histórico no se rompa
    reFormatearTodasLasTablas(sheet);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

/* =========================
   CREAR TABLA NUEVA ARRIBA
========================= */
function crearNuevoBloqueArriba(sheet, habs, fechaTxt) {
  const bloques = Math.ceil(habs.length / CONFIG.ROWS_PER_BLOCK) || 1;
  const groupWidth = 4; // HAB, ENVIADO, YA ENVIADO, SOCIOS
  const columnas = bloques * groupWidth;
  const filasTabla = 1 + CONFIG.ROWS_PER_BLOCK; // header + data
  const altoBloque = 1 + filasTabla; // fecha + tabla

  const rowFecha = CONFIG.FILA_FECHA;
  const rowHeader = rowFecha + 1;
  const rowDataStart = rowFecha + 2;

  // Empuja histórico hacia abajo
  sheet.insertRowsBefore(rowFecha, altoBloque + CONFIG.GAP_ROWS_BETWEEN_TABLES);

  asegurarTamanoHoja(
    sheet,
    rowDataStart + CONFIG.ROWS_PER_BLOCK - 1,
    CONFIG.COL_INICIO + columnas + 2
  );

  // Construir valores
  const matriz = Array.from({ length: filasTabla }, () => Array(columnas).fill(''));

  for (let b = 0; b < bloques; b++) {
    const c = b * groupWidth;
    matriz[0][c] = 'HAB';
    matriz[0][c + 1] = 'ENVIADO';
    matriz[0][c + 2] = 'YA ENVIADO';
    matriz[0][c + 3] = 'SOCIOS';
  }

  for (let i = 0; i < habs.length; i++) {
    const b = Math.floor(i / CONFIG.ROWS_PER_BLOCK);
    const r = 1 + (i % CONFIG.ROWS_PER_BLOCK);
    const c = b * groupWidth;
    matriz[r][c] = Number(habs[i]);
    matriz[r][c + 1] = false;
    matriz[r][c + 2] = false;
    matriz[r][c + 3] = false;
  }

  // Fecha
  escribirFechaEn(sheet, rowFecha, columnas, fechaTxt);

  // Escribir tabla
  sheet.getRange(rowHeader, CONFIG.COL_INICIO, filasTabla, columnas).setValues(matriz);

  // Checkboxes reales
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (let b = 0; b < bloques; b++) {
    const c = CONFIG.COL_INICIO + b * groupWidth;
    sheet.getRange(rowDataStart, c + 1, CONFIG.ROWS_PER_BLOCK, 3).setDataValidation(cb);
  }

  // Formato del bloque nuevo
  formatearBloque(sheet, rowFecha, bloques);
}

/* =========================
   REFORMATEAR TODAS LAS TABLAS
========================= */
function reFormatearTodasLasTablas(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.FILA_FECHA) return;

  const vals = sheet
    .getRange(CONFIG.FILA_FECHA, CONFIG.COL_INICIO, lastRow - CONFIG.FILA_FECHA + 1, 1)
    .getValues();

  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'HAB') {
      const rowHeader = CONFIG.FILA_FECHA + i;
      const rowFecha = rowHeader - 1;
      const bloques = contarBloquesEnHeader(sheet, rowHeader);
      if (bloques > 0) formatearBloque(sheet, rowFecha, bloques);
    }
  }
}

function contarBloquesEnHeader(sheet, rowHeader) {
  const lastCol = sheet.getLastColumn();
  const row = sheet.getRange(rowHeader, CONFIG.COL_INICIO, 1, lastCol - CONFIG.COL_INICIO + 1).getValues()[0];

  let bloques = 0;
  for (let i = 0; i < row.length; i += 4) {
    if (String(row[i] || '').trim() === 'HAB') bloques++;
    else break;
  }
  return bloques;
}

/* =========================
   FORMATO DE BLOQUE
========================= */
function formatearBloque(sheet, rowFecha, bloques) {
  const columnas = bloques * 4;
  const rowHeader = rowFecha + 1;
  const rowDataStart = rowHeader + 1;
  const filasTabla = 1 + CONFIG.ROWS_PER_BLOCK; // header + data
  const totalRows = 1 + filasTabla;             // fecha + (header+data)
  const colStart = CONFIG.COL_INICIO;
  const colEnd = colStart + columnas - 1;
  const rowBottom = rowFecha + totalRows - 1;

  const bloque = sheet.getRange(rowFecha, colStart, totalRows, columnas);

  // Base
  bloque.setBackground('#FFFFFF')
    .setFontColor('#000000')
    .setFontStyle('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(false);

  // Fecha
  const fecha = sheet.getRange(rowFecha, colStart, 1, columnas);
  fecha.setBackground(CONFIG.COLOR_FECHA)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setFontSize(CONFIG.FECHA_FONT_SIZE)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // Headers / body por bloque
  for (let b = 0; b < bloques; b++) {
    const c = colStart + b * 4;

    sheet.getRange(rowHeader, c, 1, 1).setBackground(CONFIG.COLOR_HEADER_HAB).setFontColor('#FFFFFF');
    sheet.getRange(rowHeader, c + 1, 1, 1).setBackground(CONFIG.COLOR_HEADER_ENVIADO).setFontColor('#FFFFFF');
    sheet.getRange(rowHeader, c + 2, 1, 1).setBackground(CONFIG.COLOR_HEADER_YAENVIADO).setFontColor('#FFFFFF');
    sheet.getRange(rowHeader, c + 3, 1, 1).setBackground(CONFIG.COLOR_HEADER_SOCIOS).setFontColor('#000000');

    sheet.getRange(rowHeader, c, 1, 4)
      .setFontWeight('bold')
      .setFontStyle('normal')
      .setFontSize(CONFIG.HEADER_FONT_SIZE);

    sheet.getRange(rowDataStart, c, CONFIG.ROWS_PER_BLOCK, 4)
      .setBackground('#FFFFFF')
      .setFontColor('#000000')
      .setFontStyle('normal');

    sheet.getRange(rowDataStart, c, CONFIG.ROWS_PER_BLOCK, 1).setFontSize(CONFIG.HAB_FONT_SIZE);
    sheet.getRange(rowDataStart, c + 1, CONFIG.ROWS_PER_BLOCK, 3).setFontSize(CONFIG.CHECKBOX_FONT_SIZE);

    sheet.setColumnWidth(c + 2, CONFIG.YA_ENVIADO_WIDTH);
  }

  // Limpiar bordes internos
  bloque.setBorder(false, false, false, false, false, false);

  // ---- 1) BORDE EXTERNO NEGRO GRUESO CERRADO (fecha incluida) ----
  bloque.setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Quitar unión fecha/header sin romper borde exterior
  sheet.getRange(rowFecha, colStart, 1, columnas)
    .setBorder(null, null, false, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(rowHeader, colStart, 1, columnas)
    .setBorder(false, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);

// ---- 2) LÍNEA PUNTEADA ENTRE SOCIOS / HAB (en cada separación de bloque) ----
for (let b = 0; b < bloques - 1; b++) {
  const sepCol = colStart + (b * 4) + 3; // columna SOCIOS del bloque
  sheet.getRange(rowHeader, sepCol, filasTabla, 1) // header + TODAS las filas de datos
    .setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.DOTTED);
}

  // Reforzar SOLO borde inferior grueso al final (por si algo lo pisa)
  sheet.getRange(rowBottom, colStart, 1, columnas)
    .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Alturas
  sheet.setRowHeight(rowFecha, 42);
  sheet.setRowHeight(rowHeader, 30);
  for (let r = rowDataStart; r < rowDataStart + CONFIG.ROWS_PER_BLOCK; r++) {
    sheet.setRowHeight(r, 24);
  }

  // Separación histórica
  if (CONFIG.GAP_ROWS_BETWEEN_TABLES > 0) {
    sheet.getRange(rowFecha + totalRows, colStart, CONFIG.GAP_ROWS_BETWEEN_TABLES, columnas)
      .setBackground('#FCFCE3')
      .setBorder(false, false, false, false, false, false);
  }

  // Re-refuerzo exterior completo
  bloque.setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
}

function escribirFechaEn(sheet, rowFecha, columnas, textoFecha) {
  const celda = sheet.getRange(rowFecha, CONFIG.COL_INICIO, 1, columnas);
  celda.breakApart();
  celda.clear();
  celda.merge();
  celda.setValue(textoFecha || formatearFecha(new Date()));
}

/* =========================
   PREPARAR HOJA
========================= */
function prepararHojaSiHaceFalta(sheet) {
  sheet.setHiddenGridlines(true);

  const maxRows = Math.max(sheet.getMaxRows(), 350);
  const maxCols = Math.max(sheet.getMaxColumns(), 40);

  if (maxRows > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), maxRows - sheet.getMaxRows());
  }
  if (maxCols > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxCols - sheet.getMaxColumns());
  }

  const flag = sheet.getRange('A1');
  const init = String(flag.getNote() || '') === 'LAYOUT_OK';

  // Solo una vez
  if (!init) {
    sheet.getRange(1, 1, maxRows, maxCols).setBackground(CONFIG.COLOR_BG_GENERAL);
    sheet.getRange(1, 1, 5, maxCols).setBackground(CONFIG.COLOR_TOP_BAND);
    flag.setNote('LAYOUT_OK');
  }

  // C3 combinada blanca
  const linkValue = sheet.getRange(CONFIG.FILA_LINK, CONFIG.COL_LINK).getValue();
  const linkRange = sheet.getRange(CONFIG.FILA_LINK, CONFIG.COL_LINK, 1, CONFIG.COL_LINK_ANCHO);
  linkRange.breakApart();
  linkRange.clear();
  linkRange.merge();
  linkRange.setValue(linkValue);
  linkRange.setBackground('#FFFFFF');
  linkRange.setFontColor('#000000');
  linkRange.setHorizontalAlignment('left');
  linkRange.setVerticalAlignment('middle');
  linkRange.setFontSize(11);
  linkRange.setBorder(true, true, true, true, null, null, '#BBBBBB', SpreadsheetApp.BorderStyle.SOLID);
}

/* =========================
   FUENTE EXTERNA
========================= */
function resolverHojaFuente(ss) {
  const hojas = ss.getSheets();
  if (!hojas || hojas.length === 0) throw new Error('El archivo externo no tiene pestañas.');

  const report = hojas.find(h => /^Report\b/i.test(String(h.getName() || '').trim()));
  if (report) return report;

  if (CONFIG.FUENTE_HOJA) {
    const fija = ss.getSheetByName(CONFIG.FUENTE_HOJA);
    if (fija) return fija;
  }

  return hojas[0];
}

function obtenerHabitaciones(id) {
  let ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (_) {
    throw new Error('No se pudo abrir el archivo externo. Verifica C3.');
  }

  const hoja = resolverHojaFuente(ss);
  const total = hoja.getLastRow();
  if (total < 2) return [];

  const datos = hoja.getRange(2, 1, total - 1, Math.max(CONFIG.COL_HAB, CONFIG.COL_CHECKOUT)).getValues();
  const viernes = getViernesSemanaActual();

  const vistos = new Set();
  const resultado = [];

  for (const f of datos) {
    const numero = extraerNumero(String(f[CONFIG.COL_HAB - 1] || ''));
    if (!numero || !/^7[234]/.test(numero)) continue;

    const checkout = normalizarFecha(f[CONFIG.COL_CHECKOUT - 1]);
    if (!checkout || checkout <= viernes) continue;

    if (vistos.has(numero)) continue;
    vistos.add(numero);
    resultado.push(numero);
  }

  return resultado.sort((a, b) => Number(a) - Number(b));
}

/* =========================
   FLYER
========================= */
function abrirFlyer() {
  const html = `
    <!DOCTYPE html>
    <html>
    <head><base target="_top"><meta charset="utf-8"></head>
    <body>
      <script>
        const u = '${CONFIG.URL_WEB_APP}?v=' + Date.now();
        window.open(
          u,
          'FlyerWindow',
          'width=980,height=900,left=120,top=40,resizable=yes,scrollbars=no'
        );
        google.script.host.close();
      <\/script>
    </body>
    </html>
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1).setHeight(1),
    ' '
  );
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Flyer')
    .setTitle('COCKTAIL SUNSET FLYER')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================
   HELPERS
========================= */
function asegurarTamanoHoja(sheet, minLastRow, minLastCol) {
  if (minLastRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minLastRow - sheet.getMaxRows());
  }
  if (minLastCol > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), minLastCol - sheet.getMaxColumns());
  }
}

function getViernesSemanaActual() {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const diff = 5 - hoy.getDay(); // viernes
  return new Date(hoy.setDate(hoy.getDate() + diff));
}

function formatearFecha(fecha) {
  const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  return `${dias[fecha.getDay()]}, ${('0' + fecha.getDate()).slice(-2)} de ${meses[fecha.getMonth()]} del ${fecha.getFullYear()}`;
}

function extraerId(urlOId) {
  const input = String(urlOId || '').trim();
  const m = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{30,}$/.test(input)) return input;
  return '';
}

function extraerNumero(str) {
  const m = String(str).match(/\d+/);
  return m ? m[0] : null;
}

function normalizarFecha(valor) {
  if (!valor) return null;

  if (valor instanceof Date) {
    const d = new Date(valor);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  const n = Number(valor);
  if (!isNaN(n) && n > 30000 && n < 100000) {
    const d = new Date((n - 25569) * 86400 * 1000);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  const partes = String(valor).split(/[\/\-.]/);
  if (partes.length === 3) {
    let d = new Date(parseInt(partes[2], 10), parseInt(partes[1], 10) - 1, parseInt(partes[0], 10));
    if (isNaN(d.getTime())) d = new Date(parseInt(partes[2], 10), parseInt(partes[0], 10) - 1, parseInt(partes[1], 10));
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }
  }

  return null;
}