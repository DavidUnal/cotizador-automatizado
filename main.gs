/************************************************************
 * AOP - Code.gs (UNICO) - Cotizacion + Sidebar + Menu
 ************************************************************/

/***********************
 * CONFIG
 ***********************/
const CFG = {
  // Hojas
  inicioSheet: "Inicio",
  cotSheet: "Cotizacion",

  // PDF
  pdfNameContains: "Cotizacion",

  // Excel Cliente
  excelClienteRangeA1: "A10:H1083",
  excelClienteNewSheetBaseName: "Cliente"
};

const CFG_COT = {
  sheetName: "Cotizacion",
  inputCellA1: "A4",
  firstItemRow: 25,
  lastCol: 16,              // A:P  (hasta columna P)
  templateRow: 25,          // fila plantilla para formato
  subtotalLabel: "SUBTOTAL",
  subtotalLabelCol: 5,      // columna E
  ivaRate: 0.19,

  // Columnas manuales por fila ITEM (se limpian al crear fila nueva / limpiar)
  // A=ITEM (lo pone el script), E/F/J/L/O/P son formulas (NO borrar)
  manualCols: [2, 3, 4, 7, 8, 9, 11, 13, 14] // B,C,D,G,H,I,K,M,N
};


/***********************
 * HELPERS BASICOS
 ***********************/
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function shCot_() {
  const sh = ss_().getSheetByName(CFG_COT.sheetName);
  if (!sh) throw new Error(`No existe la hoja "${CFG_COT.sheetName}".`);
  return sh;
}


function getSheetOrThrow_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error(`No existe la hoja: "${name}"`);
  return sh;
}

function normalize_(v) {
  return String(v ?? "").trim();
}

function norm_(v) {
  return String(v ?? "").trim();
}

/** Evita errores de getUi() cuando pruebas desde el editor */
function toast_(msg, title = "AOP", seconds = 3) {
  ss_().toast(msg, title, seconds);
}

function colToLetter_(col) {
  let temp = "";
  let letter = "";
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

/***********************
 * MENU (UNICO onOpen)
 ***********************/
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("AOP")
    .addItem("Abrir panel Inicio", "AOP_AbrirSidebarInicio")
    .addSeparator()
    .addItem("Ir a Inicio", "goInicio")
    .addItem("Ir a Cotizacion", "mostrarcotizacion")
    .addItem("Ir a Impresion", "mostrarimpresion")
    .addItem("Ir a Personal", "mostrarpersonal")
    .addItem("Ir a Dinero", "mostrardinero")
    .addItem("Ir a Taller", "mostrartaller")
    .addSeparator()
    .addItem("Ir a Material", "mostrarmaterial")
    .addItem("Ir a Orden Castano", "mostrarcastano")
    .addItem("Ir a Transporte", "mostrartransporte")
    .addItem("Ir a Proveedores", "mostrarproveedores")
    .addItem("Ir a Ferreteria", "mostrarferreteria")
    .addSeparator()
    .addItem("Enviar por correo (PDF)", "BotonEmailPDF")


    .addSeparator()
    .addItem("Autorizar correo (1ra vez)", "AutorizarCorreo")
    .addSeparator()

    .addToUi();

  // Abre sidebar al abrir (si no lo quieres, comenta esta línea)
  AOP_AbrirSidebarInicio();
}

/***********************
 * SIDEBAR
 ***********************/
function AOP_AbrirSidebarInicio() {
  const html = HtmlService.createHtmlOutput(getSidebarInicioHtml_())
    .setTitle("AOP - Inicio");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSidebarInicioHtml_() {
  return `
  <div style="font-family:Arial,sans-serif;padding:12px;">
    <div style="font-size:16px;font-weight:bold;margin-bottom:10px;">Panel AOP</div>

    <div style="display:grid;grid-template-columns:1fr;gap:8px;">
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrarcotizacion')">Crear Cotizacion</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrarimpresion')">Solicitud de Impresion</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrarpersonal')">Solicitud de Personal</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrardinero')">Solicitud de Dinero</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrartaller')">Solicitud de Taller</button>
      <hr/>
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrarmaterial')">Solicitud de Material</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrarcastano')">Orden Castano</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrartransporte')">Orden Transporte</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrarproveedores')">Orden Proveedores</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('mostrarferreteria')">Orden de Ferreteria</button>
      <hr/>
      <button style="padding:10px;cursor:pointer;" onclick="run('BotonIngresar')">Ingresar ITEM</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('BotonEliminar')">Eliminar ITEM</button>
      <button style="padding:10px;cursor:pointer;" onclick="run('BotonLimpiar')">Limpiar ITEMS</button>
    </div>

    <div id="status" style="margin-top:12px;font-size:12px;color:#444;"></div>

    <script>
      function run(fnName){
        document.getElementById('status').innerText = 'Ejecutando: ' + fnName + ' ...';
        google.script.run
          .withSuccessHandler(() => document.getElementById('status').innerText = 'Listo: ' + fnName)
          .withFailureHandler(err => document.getElementById('status').innerText = 'Error: ' + (err && err.message ? err.message : err))
          [fnName]();
      }
    </script>
  </div>`;
}

/***********************
 * SHOW / HIDE SHEETS
 ***********************/
function goInicio() { getSheetOrThrow_(CFG.inicioSheet).activate(); }

function _ocultarYVolverInicio_(nombreHoja) {
  getSheetOrThrow_(CFG.inicioSheet).activate();
  getSheetOrThrow_(nombreHoja).hideSheet();
}

function _mostrarYActivar_(nombreHoja) {
  const sh = getSheetOrThrow_(nombreHoja);
  sh.showSheet();
  sh.activate();
}

function ocultarcotizacion() { _ocultarYVolverInicio_("Cotizacion"); }
function mostrarcotizacion() { _mostrarYActivar_("Cotizacion"); }

function ocultarimpresion() { _ocultarYVolverInicio_("Solicitud de Impresion"); }
function mostrarimpresion() { _mostrarYActivar_("Solicitud de Impresion"); }

function ocultarpersonal() { _ocultarYVolverInicio_("Solicitud de Personal"); }
function mostrarpersonal() { _mostrarYActivar_("Solicitud de Personal"); }

function ocultardinero() { _ocultarYVolverInicio_("Solicitud de Dinero"); }
function mostrardinero() { _mostrarYActivar_("Solicitud de Dinero"); }

function ocultarcastano() { _ocultarYVolverInicio_("Orden Castano"); }
function mostrarcastano() { _mostrarYActivar_("Orden Castano"); }

function ocultarmaterial() { _ocultarYVolverInicio_("Solicitud de Material"); }
function mostrarmaterial() { _mostrarYActivar_("Solicitud de Material"); }

function ocultartransporte() { _ocultarYVolverInicio_("Orden Transporte"); }
function mostrartransporte() { _mostrarYActivar_("Orden Transporte"); }

function ocultarproveedores() { _ocultarYVolverInicio_("Orden Proveedores"); }
function mostrarproveedores() { _mostrarYActivar_("Orden Proveedores"); }

function ocultarferreteria() { _ocultarYVolverInicio_("Orden de Ferreteria"); }
function mostrarferreteria() { _mostrarYActivar_("Orden de Ferreteria"); }

function ocultartaller() { _ocultarYVolverInicio_("Solicitud de Taller"); }
function mostrartaller() { _mostrarYActivar_("Solicitud de Taller"); }

/***********************
 * COTIZACION: LIMITAR A AREA DE ITEMS
 * (desde fila 25 hasta justo antes del rótulo "SUBTOTAL")
 ***********************/
function findLabelCell_(sh, text) {
  const cell = sh.createTextFinder(text).matchCase(false).matchEntireCell(false).findNext();
  if (!cell) throw new Error(`No encontré "${text}" en la hoja "${sh.getName()}".`);
  return cell;
}

function findLastRowByMarker_(sh, markerText, extraRows) {
  const t = String(markerText || "").trim();
  if (!t) return null;

  const cell = sh.createTextFinder(t).matchCase(false).findNext();
  if (!cell) return null;

  return cell.getRow() + (extraRows ?? 0);
}


function findLastRowByMarkerContentWindow_(sh, markerText, endCol, windowRows) {
  const t = String(markerText || "").trim();
  if (!t) return null;

  const cell = sh.createTextFinder(t).matchCase(false).findNext();
  if (!cell) return null;

  const startRow = cell.getRow();
  const lastPossible = Math.min(sh.getLastRow(), startRow + (windowRows ?? 30));
  const numRows = lastPossible - startRow + 1;

  // Escanea A..endCol desde la fila del marcador hacia abajo
  const values = sh.getRange(startRow, 1, numRows, endCol).getDisplayValues();

  let lastMeaningful = startRow;
  for (let i = 0; i < values.length; i++) {
    if (values[i].some(isMeaningfulCell_)) {
      lastMeaningful = startRow + i;
    }
  }
  return lastMeaningful;
}






function findSubtotalRowCot_(sh) {
  const cell = sh.createTextFinder(CFG_COT.subtotalLabel)
    .matchCase(false)
    .matchEntireCell(false)
    .findNext();

  if (!cell) throw new Error(`No encontré "${CFG_COT.subtotalLabel}" en la hoja ${CFG_COT.sheetName}.`);
  return cell.getRow();
}

function firstEmptyItemRowCotizacion_(sh) {
  const start = CFG_COT.startRow;
  const subtotalRow = findSubtotalRowCotizacion_(sh);
  const end = subtotalRow - 1;
  if (end < start) return start;

  const values = sh.getRange(start, COT_COL.ITEM, end - start + 1, 1).getValues().flat();
  for (let i = 0; i < values.length; i++) {
    if (normalize_(values[i]) === "") return start + i;
  }
  return null; // no hay huecos
}

function lastItemRowCotizacion_(sh) {
  const start = CFG_COT.startRow;
  const subtotalRow = findSubtotalRowCotizacion_(sh);
  const end = subtotalRow - 1;
  if (end < start) return start;

  const values = sh.getRange(start, COT_COL.ITEM, end - start + 1, 1).getValues().flat();
  let last = start - 1;
  for (let i = 0; i < values.length; i++) {
    if (normalize_(values[i]) !== "") last = start + i;
  }
  return Math.max(last, start);
}

function applyTemplateFormatOnly_(sh, row) {
  // Copia SOLO formato de la fila plantilla A:P
  sh.getRange(CFG_COT.templateRow, 1, 1, CFG_COT.lastCol)
    .copyTo(sh.getRange(row, 1, 1, CFG_COT.lastCol), { formatOnly: true });
}

function ensureSummaryFormulasCotizacion_(sh) {
  // Encuentra las celdas de rótulo
  const subLabel = findLabelCell_(sh, "SUBTOTAL");
  const ivaLabel = findLabelCell_(sh, "IVA");
  const totLabel = findLabelCell_(sh, "TOTAL");

  // Determina dónde va el valor (si hay una celda "$" al lado, el valor queda dos columnas a la derecha)
  function valueCellFromLabel_(labelCell) {
    const r = labelCell.getRow();
    const c = labelCell.getColumn();
    const right = (c + 1 <= sh.getMaxColumns()) ? sh.getRange(r, c + 1).getDisplayValue() : "";
    const valueCol = (String(right).trim() === "$") ? (c + 2) : (c + 1);
    return sh.getRange(r, valueCol);
  }

  const subVal = valueCellFromLabel_(subLabel);
  const ivaVal = valueCellFromLabel_(ivaLabel);
  const totVal = valueCellFromLabel_(totLabel);

  const labelColLetter = colToLetter_(subLabel.getColumn());

  // SUBTOTAL: suma F desde fila 25 hasta la fila anterior a donde esté "SUBTOTAL"
  subVal.setFormula(
    `=SUM(F$${CFG_COT.startRow}:INDEX(F:F, MATCH("SUBTOTAL", ${labelColLetter}:${labelColLetter}, 0)-1))`
  );

  // IVA (19%)
  ivaVal.setFormula(`=${subVal.getA1Notation()}*0.19`);

  // TOTAL
  totVal.setFormula(`=${subVal.getA1Notation()}+${ivaVal.getA1Notation()}`);
}

/***********************
 * BOTON: INGRESAR (CORREGIDO)
 * - Escribe en la primera fila vacía dentro del bloque de ITEMS
 * - Si no hay espacio, inserta una fila justo antes del SUBTOTAL
 * - NO copia contenidos (evita “recuadros” y #REF), solo formato
 * - Ajusta fórmula de Valor Total por fila (col F)
 ***********************/
function BotonIngresar() {
  const sh = shCot_();
  sh.activate();

  const input = sh.getRange(CFG_COT.inputCellA1);
  const item = norm_(input.getValue());

  if (!item) {
    toast_("Agregue un ITEM en A4.", "AOP", 4);
    sh.setActiveRange(input);
    return;
  }

  const baseCell = sh.getRange(CFG_COT.firstItemRow, 1); // A25
  const baseEmpty = norm_(baseCell.getValue()) === "";

  let targetRow = CFG_COT.firstItemRow;

  if (baseEmpty) {
    // Primer ITEM
    baseCell.setValue(item);
    applyCotRowFormulas_(sh, targetRow);
  } else {
    // Siguiente ITEM: insertar debajo del último item real (antes de SUBTOTAL)
    const lastItemRow = getLastItemRowCot_(sh);
    sh.insertRowsAfter(lastItemRow, 1);
    targetRow = lastItemRow + 1;

    // Copiar formato/estructura A:P desde la plantilla
    sh.getRange(CFG_COT.templateRow, 1, 1, CFG_COT.lastCol)
      .copyTo(sh.getRange(targetRow, 1, 1, CFG_COT.lastCol), { contentsOnly: false });

    // Limpiar manuales para que no se copie nada raro
    clearManualCellsCotRow_(sh, targetRow, { clearItem: true });

    // Set ITEM + formulas
    sh.getRange(targetRow, 1).setValue(item);
    applyCotRowFormulas_(sh, targetRow);
  }

  // Totales siempre consistentes
  applyCotTotalsFormulas_(sh);

  // Limpiar input y enfocar
  input.clearContent();
  sh.setActiveRange(input);

  toast_(`ITEM agregado en fila ${targetRow}.`, "AOP", 3);
}

/***********************
 * BOTON: ELIMINAR (solo dentro del bloque ITEMS)
 ***********************/
function BotonEliminar() {
  const sh = shCot_();
  const cell = sh.getActiveCell();

  if (cell.getColumn() !== 1) {
    toast_("Selecciona un ITEM en la columna A para eliminar.", "AOP", 4);
    sh.setActiveRange(sh.getRange(CFG_COT.inputCellA1));
    return;
  }

  const row = cell.getRow();
  const subtotalRow = findSubtotalRowCot_(sh);

  // Validar que esté dentro del área de items
  if (row < CFG_COT.firstItemRow || row >= subtotalRow) {
    toast_("Selecciona una fila válida de ITEM (desde la fila 25).", "AOP", 4);
    sh.setActiveRange(sh.getRange(CFG_COT.inputCellA1));
    return;
  }
  const lastItemRow = getLastItemRowCot_(sh);

  // Si solo existe A25 como item
  if (lastItemRow === CFG_COT.firstItemRow && row === CFG_COT.firstItemRow) {
    clearManualCellsCotRow_(sh, CFG_COT.firstItemRow, { clearItem: true });
    applyCotRowFormulas_(sh, CFG_COT.firstItemRow);
    applyCotTotalsFormulas_(sh);
    sh.setActiveRange(sh.getRange(CFG_COT.inputCellA1));
    toast_("ITEM eliminado (se limpió la fila 25).", "AOP", 3);
    return;
  }

  // Si hay más, borrar la fila completa
  sh.deleteRow(row);

  // Recalcular totales
  applyCotTotalsFormulas_(sh);

  sh.setActiveRange(sh.getRange(CFG_COT.inputCellA1));
  toast_("ITEM eliminado.", "AOP", 3);
}


/***********************
 * BOTON: LIMPIAR (SOLO ITEMS desde fila 25 hasta antes de SUBTOTAL)
 * - No toca encabezados ni secciones de arriba
 * - Evita el error de SpreadsheetApp.ZgetUi (era un typo)
 ***********************/
function BotonLimpiar() {
  const sh = shCot_();
  sh.activate();

  const lastItemRow = getLastItemRowCot_(sh);

  // Borra filas agregadas (26..lastItemRow)
  if (lastItemRow > CFG_COT.firstItemRow) {
    const start = CFG_COT.firstItemRow + 1;
    const howMany = lastItemRow - CFG_COT.firstItemRow;
    sh.deleteRows(start, howMany);
  }

  // Limpia fila base (A25) pero conserva formulas
  clearManualCellsCotRow_(sh, CFG_COT.firstItemRow, { clearItem: true });
  applyCotRowFormulas_(sh, CFG_COT.firstItemRow);

  // Totales
  applyCotTotalsFormulas_(sh);

  // Input
  const input = sh.getRange(CFG_COT.inputCellA1);
  input.clearContent();
  sh.setActiveRange(input);

  toast_("Cotización limpia (solo ITEMS).", "AOP", 3);
}

function RepararFormulasCotizacion() {
  const sh = shCot_();
  const last = getLastItemRowCot_(sh);

  for (let r = CFG_COT.firstItemRow; r <= last; r++) {
    if (norm_(sh.getRange(r, 1).getValue()) !== "") {
      applyCotRowFormulas_(sh, r);
    }
  }
  applyCotTotalsFormulas_(sh);
  toast_("Fórmulas reparadas en filas de ITEMS.", "AOP", 4);
}

/***********************
 * PDF
 ***********************/
const CFG_PDF_COT = {
  sheetName: "Cotizacion",
  startRow: 10,      // <-- antes 1. Sube/baja si quieres más/menos encabezado
  startCol: 1,
  endCol: 7,        // <-- SOLO A..G
  maxScanRows: 200,
  fileNamePrefix: "Cotizacion"
};





/**
 * BOTON PDF (Cotizacion)
 * Exporta SOLO el rango A2:P{lastRowConContenido}
 */
function BotonPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_PDF_COT.sheetName);
  if (!sh) throw new Error(`No existe la hoja "${CFG_PDF_COT.sheetName}".`);

  // FORZAR solo hasta columna G
  const END_COL_G = 7;

  // 1) lastRow por contenido (tu método)
  const lastByContent = getLastMeaningfulRowInRect_(
    sh,
    CFG_PDF_COT.startRow,
    CFG_PDF_COT.startCol,
    END_COL_G,
    CFG_PDF_COT.maxScanRows
  );

  // 2) lastRow "seguro" por marcador de notas (evita 2da página)
  const lastByMarker = findLastRowByMarkerContentWindow_(sh, "Forma de pago", END_COL_G, 30);

  // si existe el marcador, nos quedamos con ese límite (más estable)
  const lastRow = lastByMarker ? Math.min(lastByContent, lastByMarker) : lastByContent;

  // 3) Export bounds
  const r1 = CFG_PDF_COT.startRow - 1;
  const c1 = CFG_PDF_COT.startCol - 1;
  const r2 = lastRow;
  const c2 = END_COL_G;

  // 4) Generar PDF (NO guardar en Drive) y enviar al modal para descargar
  // Tomar valores desde la hoja (rango combinado B..G)
  const consecutivo = String(sh.getRange("B21").getDisplayValue()).trim();
  const actividad   = String(sh.getRange("B16").getDisplayValue()).trim();

  // Sanitizar para nombre de archivo (sin caracteres inválidos)
  const clean = (s) => s
    .replace(/[\\\/:*?"<>|]/g, "")   // inválidos en Windows
    .replace(/\s+/g, " ")            // colapsa espacios
    .trim();

  const cons = clean(consecutivo || "SIN_CONSECUTIVO");
  const act  = clean(actividad   || "SIN_ACTIVIDAD");

  // Nombre final
  const pdfName = `${cons} - ${act}.pdf`;

  
  const pdfBlob = exportSheetRectToPdfBlob_(
    ss.getId(),
    sh.getSheetId(),
    pdfName,
    r1,
    r2,
    c1,
    c2
  );

  const tpl = HtmlService.createTemplateFromFile("DownloadPdf");
  tpl.pdfName = pdfName;
  tpl.pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.NATIVE)
      .setWidth(480)
      .setHeight(280),
    "Descargar PDF"
  );

  ss.toast(`PDF listo para descargar: ${pdfName}`, "AOP", 5);
}





/**
 * Escanea A:G desde startRow hasta startRow+maxScanRows-1 y retorna la última fila con contenido real.
 */
function getLastMeaningfulRowInRect_(sh, startRow, startCol, endCol, maxScanRows) {
  const lastPossible = Math.min(sh.getLastRow(), startRow + maxScanRows - 1);
  if (lastPossible < startRow) return startRow;

  const numRows = lastPossible - startRow + 1;
  const numCols = endCol - startCol + 1;

  const values = sh.getRange(startRow, startCol, numRows, numCols).getDisplayValues();

  let lastMeaningful = startRow;

  for (let r = 0; r < values.length; r++) {
    const rowVals = values[r];
    const hasMeaning = rowVals.some(isMeaningfulCell_);
    if (hasMeaning) lastMeaningful = startRow + r;
  }

  return lastMeaningful;
}

function isMeaningfulCell_(v) {
  const s = String(v ?? "").trim();
  if (s === "") return false;
  if (s === "-") return false;
  // ignora ceros decorativos ($0, 0, 0.00, etc.)
  if (/^\$?\s*0+([.,]0+)?$/.test(s)) return false;
  return true;
}

function copyOverGridImagesByUrl_(srcSheet, dstSheet, maxColToCopy) {
  const imgs = srcSheet.getImages(); // OverGridImage[]
  if (!imgs || !imgs.length) return;

  const token = ScriptApp.getOAuthToken();

  imgs.forEach(img => {
    const anchor = img.getAnchorCell();
    const row = anchor.getRow();
    const col = anchor.getColumn();

    // solo copiar si la imagen cae en A..G (o el maxCol que definas)
    if (col > maxColToCopy) return;

    const url = img.getContentUrl(); // <- clave
    if (!url) return;

    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) return;

    const blob = resp.getBlob();
    const newImg = dstSheet.insertImage(blob, col, row);

    // Mantener tamaño aproximado
    try {
      newImg.setWidth(img.getWidth()).setHeight(img.getHeight());
    } catch (e) {}

    // Mantener offsets si tu versión lo soporta
    try {
      newImg.setAnchorCellXOffset(img.getAnchorCellXOffset());
      newImg.setAnchorCellYOffset(img.getAnchorCellYOffset());
    } catch (e) {}
  });
}










function copyOverGridImages_(srcSheet, dstSheet, maxColToCopy) {
  const imgs = srcSheet.getImages(); // imágenes "sobre celdas"
  imgs.forEach(img => {
    const anchor = img.getAnchorCell();
    const col = anchor.getColumn();
    const row = anchor.getRow();

    // Solo copiar imágenes ancladas dentro de A..G (para evitar copiar cosas de H en adelante)
    if (col > maxColToCopy) return;

    const newImg = dstSheet.insertImage(img.getBlob(), col, row);
    newImg.setWidth(img.getWidth()).setHeight(img.getHeight());

    // Mantener offsets si existen (mejora la posición)
    try {
      newImg.setAnchorCellXOffset(img.getAnchorCellXOffset());
      newImg.setAnchorCellYOffset(img.getAnchorCellYOffset());
    } catch (e) {
      // si tu versión no soporta offsets, lo ignora sin romper
    }
  });
}








/**
 * Decide si un valor “cuenta” como contenido real para definir el final del PDF.
 * - Ignora: vacío, "-", "0", "$ 0", "0.00", etc.
 * - Cuenta: texto, números relevantes, encabezados, etc.
 */


function getLastItemRowCot_(sh) {
  const start = CFG_COT.firstItemRow;
  const subtotalRow = findSubtotalRowCot_(sh);
  const end = subtotalRow - 1;

  if (end < start) return start;

  const values = sh.getRange(start, 1, end - start + 1, 1).getValues().flat();
  for (let i = values.length - 1; i >= 0; i--) {
    if (norm_(values[i]) !== "") return start + i;
  }
  return start;
}


/** Limpia SOLO celdas manuales en una fila (no borra formulas) */
function clearManualCellsCotRow_(sh, row, { clearItem = false } = {}) {
  if (clearItem) sh.getRange(row, 1).clearContent(); // A
  CFG_COT.manualCols.forEach(col => sh.getRange(row, col).clearContent());
}

function applyCotTotalsFormulas_(sh) {
  const subtotalRow = findSubtotalRowCot_(sh);

  // SUBTOTAL (VENTA) en F
  sh.getRange(subtotalRow, 6).setFormula(
    `=SUM(F${CFG_COT.firstItemRow}:INDEX(F:F, MATCH("${CFG_COT.subtotalLabel}",$E:$E,0)-1))`
  );

  // IVA en F (fila siguiente)
  sh.getRange(subtotalRow + 1, 6).setFormula(
    `=IF(F${subtotalRow}="","",F${subtotalRow}*${CFG_COT.ivaRate})`
  );

  // TOTAL en F (fila siguiente)
  sh.getRange(subtotalRow + 2, 6).setFormula(
    `=IF(F${subtotalRow}="","",F${subtotalRow}+F${subtotalRow + 1})`
  );

  // SUBTOTAL COSTOS (columna J) en la misma fila del SUBTOTAL
  sh.getRange(subtotalRow, 10).setFormula(
    `=SUM(J${CFG_COT.firstItemRow}:INDEX(J:J, MATCH("${CFG_COT.subtotalLabel}",$E:$E,0)-1))`
  );
}



function OpenSaveAsSidebar(pdfName, pdfBase64) {
  const tpl = HtmlService.createTemplateFromFile("SaveAsSidebar");
  tpl.pdfName = pdfName;
  tpl.pdfBase64 = pdfBase64;

  SpreadsheetApp.getUi().showSidebar(
    tpl.evaluate().setTitle("Guardar PDF (Elegir carpeta)")
  );
}





function applyCotRowFormulas_(sh, row) {
  // E: Valor Unitario = ROUNDUP(I/(1-K), -3)
  sh.getRange(row, 5).setFormula(
    `=IF($A${row}="","",ROUNDUP($I${row}/(1-$K${row}),-3))`
  );

  // F: Valor Total = C * D * E
  sh.getRange(row, 6).setFormula(
    `=IF($A${row}="","",$C${row}*$D${row}*$E${row})`
  );

  // J: Costo Total = C * D * I
  sh.getRange(row, 10).setFormula(
    `=IF($A${row}="","",$C${row}*$D${row}*$I${row})`
  );

  // L: Margen $ = F - J
  sh.getRange(row, 12).setFormula(
    `=IF($A${row}="","",$F${row}-$J${row})`
  );

  // O: Costo Total 4 = C * D * N
  sh.getRange(row, 15).setFormula(
    `=IF($A${row}="","",$C${row}*$D${row}*$N${row})`
  );

  // P: Margen (ejecutado) = F - O
  sh.getRange(row, 16).setFormula(
    `=IF($A${row}="","",$F${row}-$O${row})`
  );
}

/**
 * Devuelve la última fila (dentro del rectángulo) que tenga al menos 1 celda con contenido “real”.
 * Escanea A:P desde startRow hacia abajo, y busca desde el final hacia arriba.
 */
function getLastMeaningfulRowInRect_(sh, startRow, startCol, endCol, maxScanRows) {
  const lastRowSheet = Math.max(sh.getLastRow(), startRow);
  const lastToCheck = Math.min(lastRowSheet, startRow + maxScanRows - 1);

  const numRows = lastToCheck - startRow + 1;
  const numCols = endCol - startCol + 1;

  const values = sh.getRange(startRow, startCol, numRows, numCols).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    if (row.some(isMeaningfulCell_)) {
      return startRow + i;
    }
  }

  // si no encontró nada “real”, exporta mínimo la fila startRow
  return startRow;
}

function exportSheetRectToPdfBlob_(spreadsheetId, sheetId, filename, r1, r2, c1, c2) {
  const base =
    "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?format=pdf" +
    "&gid=" + sheetId +

    // recorte de rango
    "&r1=" + r1 +
    "&r2=" + r2 +
    "&c1=" + c1 +
    "&c2=" + c2 +

    // opciones de impresión
    "&portrait=true" +

    // >>> CAMBIO ESTRICTAMENTE NECESARIO: forzar 1 sola página <<<
    // scale=4 => "Fit to page" (ajusta TODO a 1 página)
    "&scale=4" +

    // (deja fitw fuera para que no interfiera)
    // "&fitw=true"  <-- eliminado

    "&gridlines=false" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&fzr=false" +
    "&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25";

  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(base, { headers: { Authorization: "Bearer " + token } });
  return resp.getBlob().setName(filename);
}


/***********************
 * EXCEL CLIENTE (crear hoja)
 ***********************/
/***********************
 * EXCEL CLIENTE (crear hoja) - FIX
 * - Copia EXACTO lo visible del PDF (desde fila 9, columnas A:G)
 * - No copia botones (se eliminan filas superiores)
 * - No genera #REF porque CLONA la hoja primero (con fórmulas e imagen)
 * - Deja VALORES (pega contentsOnly) para que el cliente vea números finales
 ***********************/
/***********************
 * EXCEL CLIENTE (crear hoja) - FIX REAL
 * - NO copia botones: elimina DRAWINGS (botones son drawings)
 * - NO genera #REF: primero "congela" valores (con todas las columnas aún),
 *   y DESPUÉS recorta a A:G
 * - Mantiene logo: NO elimina imágenes, solo drawings
 * - Recorta filas/columnas para que quede como el PDF (sin basura extra)
 ***********************/
function hojaCliente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SRC_SHEET = "Cotizacion";
  const NEW_BASE_NAME = "Cliente";

  // Lo que el cliente debe ver: A..G
  const END_COL = 7;           // A:G

  // Para congelar valores sin #REF necesitamos conservar columnas de cálculo (hasta P)
  const FREEZE_COLS = 16;      // A:P (ajusta si tu cotización usa más columnas)

  const MAX_SCAN_ROWS = 2500;

  const src = ss.getSheetByName(SRC_SHEET);
  if (!src) throw new Error(`No existe la hoja "${SRC_SHEET}".`);

  // 1) Definir última fila visible (igual lógica del PDF)
  const lastByContent = getLastMeaningfulRowInRect_(src, 1, 1, END_COL, MAX_SCAN_ROWS);
  const lastByMarker = findLastRowByMarkerContentWindow_(src, "Forma de pago", END_COL, 30);

  const LAST_ROW_SRC = lastByMarker ? Math.min(lastByContent, lastByMarker) : lastByContent;

  // 2) Clonar hoja completa (mantiene formatos + logo)
  const newName = buildUniqueSheetName_(ss, NEW_BASE_NAME);
  const newSheet = src.copyTo(ss).setName(newName);

  // Moverla justo después de Cotizacion
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(src.getIndex() + 1);

  // 3) ELIMINAR BOTONES: los botones son "drawings" (NO imágenes)
  try {
    const drawings = newSheet.getDrawings();
    drawings.forEach(d => d.remove());
  } catch (e) {
    // si no soporta getDrawings en tu entorno, no rompe
  }

  // 4) Determinar desde qué fila arrancar para NO cortar el logo.
  //    Usamos la fila más alta donde haya una imagen (logo), y si no hay, arrancamos desde 1.
  let topKeepRow = 1;
  try {
    const imgs = newSheet.getImages();
    if (imgs && imgs.length) {
      topKeepRow = imgs
        .map(img => img.getAnchorCell().getRow())
        .reduce((a, b) => Math.min(a, b), 999999);
    }
  } catch (e) {}

  // Si por alguna razón el logo está más abajo, igual queremos mantenerlo,
  // pero NO queremos filas sobrantes arriba (botonera). Entonces:
  // - si topKeepRow > 1, cortamos filas 1..topKeepRow-1
  // (eso elimina “ranuras” de arriba sin tocar el logo)
  // Nota: esto NO elimina imágenes, solo desplaza hacia arriba.
  if (topKeepRow > 1) {
    newSheet.deleteRows(1, topKeepRow - 1);
  }

  // Ajuste: como ya borramos filas arriba, la LAST_ROW_SRC también se desplaza
  // (pero solo si topKeepRow > 1)
  const lastRowAfterTopCut = LAST_ROW_SRC - (topKeepRow > 1 ? (topKeepRow - 1) : 0);

  // 5) Recortar filas inferiores (para que quede como el PDF)
  const maxRows = newSheet.getMaxRows();
  if (maxRows > lastRowAfterTopCut) {
    newSheet.deleteRows(lastRowAfterTopCut + 1, maxRows - lastRowAfterTopCut);
  }

  // 6) 🔥 CLAVE: Congelar valores ANTES de borrar columnas.
  //    Esto evita #REF porque la hoja todavía tiene todas las columnas necesarias para calcular.
  const rowsToFreeze = newSheet.getLastRow();
  if (rowsToFreeze > 0) {
    const colsToFreeze = Math.min(FREEZE_COLS, newSheet.getMaxColumns());
    const rngFreeze = newSheet.getRange(1, 1, rowsToFreeze, colsToFreeze);
    rngFreeze.copyTo(rngFreeze, { contentsOnly: true }); // deja valores finales
  }

  // 7) Ahora sí: recortar columnas a A:G (sin riesgo de #REF)
  const maxCols = newSheet.getMaxColumns();
  if (maxCols > END_COL) {
    newSheet.deleteColumns(END_COL + 1, maxCols - END_COL);
  }

  // 8) Ajustes visuales
  newSheet.setHiddenGridlines(true);

  // Cursor en un lugar “bonito”
  try { newSheet.setActiveRange(newSheet.getRange("B6")); } catch (e) {}

  ss.toast(`Hoja "${newName}" creada (sin botones, sin #REF, con valores finales).`, "AOP", 5);
}



function buildUniqueSheetName_(ss, baseName) {
  const existing = new Set(ss.getSheets().map(s => s.getName()));
  if (!existing.has(baseName)) return baseName;
  let i = 2;
  while (existing.has(`${baseName} (${i})`)) i++;
  return `${baseName} (${i})`;
}

function ensureSheetSize_(sh, minRows, minCols) {
  const curRows = sh.getMaxRows();
  const curCols = sh.getMaxColumns();
  if (curRows < minRows) sh.insertRowsAfter(curRows, minRows - curRows);
  if (curCols < minCols) sh.insertColumnsAfter(curCols, minCols - curCols);
}
