/**
 * ============================================================================
 *  PORTADA - STATUS DE RENDICIONES   (v3 - corrige combinacion de celdas)
 * ----------------------------------------------------------------------------
 *  Genera (o regenera) una hoja "Portada" al principio del libro con una
 *  barra de progreso por responsable: % de comprobantes ADJUNTADOS sobre el
 *  total cargado.
 *
 *  - El % se calcula con COUNTIF sobre la columna "Estado Comprobante" (F),
 *    por lo que se ACTUALIZA SOLO a medida que agregas filas en cada hoja.
 *  - Si una persona todavia no tiene rendiciones cargadas, cuenta como 100%.
 *  - Para sumar una persona nueva: agrega una linea a CFG.personas y volve
 *    a correr crearPortada().
 *
 *  Como usarlo:
 *    1) Google Sheets -> Extensiones -> Apps Script.
 *    2) Borra el codigo anterior, pega este, guarda.
 *    3) Ejecuta crearPortada() (o usa el menu "Portada" al reabrir).
 * ============================================================================
 */

// --- Configuracion ---------------------------------------------------------
var CFG = {
  hojaPortada: 'Portada',
  personas: [
    { nombre: 'Gabi', hoja: 'RENDICIONES GABI' },
    { nombre: 'Fede', hoja: 'RENDICIONES FEDE' }
    // , { nombre: 'Nuevo', hoja: 'RENDICIONES NUEVO' }   // <- sumar aca
  ],
  colEstado: 'F',
  valAdjuntado: 'Adjuntado',
  valPendiente: 'No Adjuntado',
  titulo: 'STATUS DE RENDICIONES',
  subtitulo: 'Seguimiento de comprobantes adjuntados por responsable'
};

// --- Paleta ----------------------------------------------------------------
var C = {
  morado:    '#5B3F86',
  moradoOsc: '#4A3270',
  track:     '#E8E6EF',
  lila:      '#EDE7F6',
  lilaMed:   '#D1C3E7',
  gris:      '#7A7A7A',
  texto:     '#33313A',
  blanco:    '#FFFFFF',
  borde:     '#E0DAF0',
  zebra:     '#F5F3FA'
};
var FONT = 'Lato';

// --- Menu -------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Portada')
    .addItem('Crear / Actualizar portada', 'crearPortada')
    .addToUi();
}

// --- Actualiza la fecha sola al editar una hoja de rendiciones --------------
// (los numeros y barras ya se actualizan solos porque son formulas COUNTIF)
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var nombreHoja = e.range.getSheet().getName();
    var esRendicion = CFG.personas.some(function (p) { return p.hoja === nombreHoja; });
    if (!esRendicion) return;
    var sh = e.source.getSheetByName(CFG.hojaPortada);
    if (!sh) return;
    sh.getRange('B4').setValue('  ' + CFG.subtitulo + '   -   Actualizado ' + hoy());
  } catch (err) {
    // Silencioso: nunca interrumpe la edicion del usuario
  }
}

// --- Funcion principal ------------------------------------------------------
function crearPortada() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sh = ss.getSheetByName(CFG.hojaPortada);
  if (!sh) sh = ss.insertSheet(CFG.hojaPortada);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);

  // Limpieza total: contenido, formatos y COMBINACIONES previas
  sh.clear();
  sh.clearNotes();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.setHiddenGridlines(true);
  if (sh.getMaxRows() > 60) sh.deleteRows(61, sh.getMaxRows() - 60);
  if (sh.getMaxColumns() > 9) sh.deleteColumns(10, sh.getMaxColumns() - 9);

  var anchos = { 1: 26, 2: 200, 3: 150, 4: 130, 5: 130, 6: 130, 7: 130, 8: 26 };
  for (var col in anchos) sh.setColumnWidth(Number(col), anchos[col]);

  var n = CFG.personas.length;

  // -- Layout (sin solapamientos) -------------------------------------------
  var L = {};
  L.banner    = 2;   // banner ocupa filas 2-3
  L.sub       = 4;
  L.gTitle    = 6;   // "AVANCE GLOBAL" (B6:H6)
  L.gBig      = 7;   // % grande (B7:C9)  |  derecha: label/barra/caption filas 7-9
  L.gBar      = 8;
  L.gCap      = 9;
  L.gLabel    = 10;  // "AVANCE GENERAL" bajo el numero
  L.prTitle   = 12;  // "POR RESPONSABLE"
  L.cardStart = 13;
  L.detTitle  = L.cardStart + n * 3 + 1;
  L.detHead   = L.detTitle + 1;
  L.detFirst  = L.detHead + 1;
  L.detGlobal = L.detFirst + n;

  var dir = {};
  for (var i = 0; i < n; i++) {
    var rp = L.detFirst + i;
    dir[CFG.personas[i].nombre] = { adj: 'C' + rp, pen: 'D' + rp, tot: 'E' + rp, pct: 'F' + rp };
  }
  var dirGlobal = {
    adj: 'C' + L.detGlobal, pen: 'D' + L.detGlobal,
    tot: 'E' + L.detGlobal, pct: 'F' + L.detGlobal
  };

  // -- Banner ---------------------------------------------------------------
  sh.setRowHeight(1, 8);
  merge(sh, L.banner, 2, 2, 7)
    .setValue('  ' + CFG.titulo)
    .setBackground(C.morado).setFontColor(C.blanco)
    .setFontFamily(FONT).setFontSize(26).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(L.banner, 42);
  sh.setRowHeight(L.banner + 1, 18);

  merge(sh, L.sub, 2, 1, 7)
    .setValue('  ' + CFG.subtitulo + '   -   Actualizado ' + hoy())
    .setBackground(C.moradoOsc).setFontColor('#D9CEEF')
    .setFontFamily(FONT).setFontSize(11)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(L.sub, 26);
  sh.setRowHeight(5, 16);

  // -- Avance global --------------------------------------------------------
  seccion(sh, L.gTitle, 'AVANCE GLOBAL');   // B6:H6

  // Numero grande (izquierda, B7:C9)
  merge(sh, L.gBig, 2, 3, 2)
    .setFormula('=' + dirGlobal.pct)
    .setNumberFormat('0.0%')
    .setFontFamily(FONT).setFontSize(46).setFontWeight('bold')
    .setFontColor(C.morado)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(L.gBig, 34);

  merge(sh, L.gLabel, 2, 1, 2)
    .setValue('AVANCE GENERAL')
    .setFontFamily(FONT).setFontSize(10).setFontWeight('bold')
    .setFontColor(C.gris).setHorizontalAlignment('center');

  // Derecha: label (fila 7) / barra (fila 8) / caption (fila 9), columnas D:G
  merge(sh, L.gBig, 4, 1, 4)
    .setValue('Progreso global')
    .setFontFamily(FONT).setFontSize(12).setFontWeight('bold')
    .setFontColor(C.texto).setVerticalAlignment('middle');
  barra(merge(sh, L.gBar, 4, 1, 4), dirGlobal.pct);
  sh.setRowHeight(L.gBar, 30);
  merge(sh, L.gCap, 4, 1, 4)
    .setFormula(captionFormula(dirGlobal))
    .setFontFamily(FONT).setFontSize(10).setFontColor(C.gris)
    .setVerticalAlignment('middle');

  sh.setRowHeight(11, 16);

  // -- Por responsable ------------------------------------------------------
  seccion(sh, L.prTitle, 'POR RESPONSABLE');

  for (var j = 0; j < n; j++) {
    var p = CFG.personas[j];
    var rName = L.cardStart + j * 3;
    var rCap  = rName + 1;
    var d = dir[p.nombre];

    sh.getRange(rName, 2)
      .setValue(p.nombre)
      .setFontFamily(FONT).setFontSize(14).setFontWeight('bold')
      .setFontColor(C.texto).setVerticalAlignment('middle');

    sh.getRange(rName, 3)
      .setFormula('=' + d.pct).setNumberFormat('0.0%')
      .setFontFamily(FONT).setFontSize(22).setFontWeight('bold')
      .setFontColor(C.morado)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');

    barra(merge(sh, rName, 4, 1, 4), d.pct);
    sh.setRowHeight(rName, 34);

    merge(sh, rCap, 4, 1, 4)
      .setFormula(captionFormula(d))
      .setFontFamily(FONT).setFontSize(10).setFontColor(C.gris)
      .setVerticalAlignment('middle');
    sh.setRowHeight(rCap, 18);
    sh.setRowHeight(rName + 2, 10);
  }

  // -- Detalle de calculo ---------------------------------------------------
  merge(sh, L.detTitle, 2, 1, 5)
    .setValue('DETALLE DE CALCULO')
    .setFontFamily(FONT).setFontSize(11).setFontWeight('bold')
    .setFontColor(C.morado);

  sh.getRange(L.detHead, 2, 1, 5)
    .setValues([['Responsable', 'Adjuntados', 'Pendientes', 'Total', '% Adjuntado']])
    .setBackground(C.morado).setFontColor(C.blanco)
    .setFontFamily(FONT).setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');

  for (var k = 0; k < n; k++) {
    var pp = CFG.personas[k];
    var r = L.detFirst + k;
    var ref = "'" + pp.hoja + "'!" + CFG.colEstado + ':' + CFG.colEstado;
    sh.getRange(r, 2).setValue(pp.nombre);
    sh.getRange(r, 3).setFormula('=COUNTIF(' + ref + ',"' + CFG.valAdjuntado + '")');
    sh.getRange(r, 4).setFormula('=COUNTIF(' + ref + ',"' + CFG.valPendiente + '")');
    sh.getRange(r, 5).setFormula('=C' + r + '+D' + r);
    sh.getRange(r, 6).setFormula('=IF(E' + r + '=0,1,C' + r + '/E' + r + ')').setNumberFormat('0.0%');
  }

  var rg = L.detGlobal;
  var first = L.detFirst, last = L.detFirst + n - 1;
  sh.getRange(rg, 2).setValue('Global').setFontWeight('bold');
  sh.getRange(rg, 3).setFormula('=SUM(C' + first + ':C' + last + ')');
  sh.getRange(rg, 4).setFormula('=SUM(D' + first + ':D' + last + ')');
  sh.getRange(rg, 5).setFormula('=SUM(E' + first + ':E' + last + ')');
  sh.getRange(rg, 6).setFormula('=IF(E' + rg + '=0,1,C' + rg + '/E' + rg + ')').setNumberFormat('0.0%');

  sh.getRange(L.detFirst, 2, n + 1, 5)
    .setFontFamily(FONT).setFontSize(10)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, C.borde, SpreadsheetApp.BorderStyle.SOLID);
  for (var z = 0; z < n; z++) {
    sh.getRange(L.detFirst + z, 2, 1, 5).setBackground(z % 2 === 0 ? C.blanco : C.zebra);
  }
  sh.getRange(rg, 2, 1, 5).setBackground(C.lila).setFontWeight('bold');
  sh.getRange(L.detFirst, 2, n + 1, 1).setHorizontalAlignment('left');

  sh.setActiveSelection('A1');
  SpreadsheetApp.getUi().alert('Portada creada/actualizada correctamente.');
}

// --- Helpers ----------------------------------------------------------------
function merge(sh, row, col, numRows, numCols) {
  var rng = sh.getRange(row, col, numRows, numCols);
  rng.merge();
  return rng;
}

function seccion(sh, row, texto) {
  merge(sh, row, 2, 1, 7)
    .setValue(texto)
    .setFontFamily(FONT).setFontSize(11).setFontWeight('bold')
    .setFontColor(C.morado);
  sh.setRowHeight(row, 22);
}

function barra(rng, celdaPct) {
  var f = '=SPARKLINE({' + celdaPct + ',1-' + celdaPct + '},' +
          '{"charttype","bar";"max",1;"color1","' + C.morado + '";"color2","' + C.track + '"})';
  rng.setFormula(f);
}

function captionFormula(d) {
  return '=IF(' + d.tot + '=0,"Sin rendiciones cargadas - 100% al dia",' +
         d.adj + '&" de "&' + d.tot + '&" comprobantes adjuntados - "&' + d.pen + '&" pendientes")';
}

function hoy() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

