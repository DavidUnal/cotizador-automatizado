const CFG_TRANSPORTE = {
  sheetName: "Orden Transporte",
  inputCellA1: "A4",

  tableStartRow: 16,     // primera fila de items
  tableItemCol: 1,       // col A
  tableLastCol: 6,       // A:F
  templateRow: 16,       // fila plantilla a copiar

  totalLabel: "TOTAL",
  totalLabelCol: 4,      // columna donde está la palabra TOTAL (D)
  totalValueCol: 5,      // columna del valor total (E)

  // columnas a limpiar cuando se borra la última fila (deja E intacta porque suele tener fórmula)
  clearColsA1: ["A16","B16","C16","D16","F16"]
};

/** UI seguro (evita errores si ejecutas desde Run en el editor) */
function getUiSafe_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}
function uiAlertSafe_(title, msg) {
  const ui = getUiSafe_();
  if (ui) ui.alert(title, msg, ui.ButtonSet.OK);
  else Logger.log(`${title}: ${msg}`);
}
function normalize_(v) {
  return String(v ?? "").trim();
}

/** Encuentra la fila donde está el rótulo TOTAL (para usarla como límite inferior de la tabla) */
function getTotalRowTransporte_(sh) {
  const finder = sh
    .createTextFinder(CFG_TRANSPORTE.totalLabel)
    .matchCase(false)
    .matchEntireCell(true);

  const cell = finder.findNext();
  if (!cell) throw new Error(`No encontré "${CFG_TRANSPORTE.totalLabel}" en la hoja "${CFG_TRANSPORTE.sheetName}".`);

  return cell.getRow();
}

/** Última fila con ítem (col A) pero SOLO hasta antes de TOTAL */
function getLastRowTransporte_(sh) {
  const start = CFG_TRANSPORTE.tableStartRow;
  const totalRow = getTotalRowTransporte_(sh);
  const end = totalRow - 1;

  if (end < start) return start;

  const values = sh
    .getRange(start, CFG_TRANSPORTE.tableItemCol, end - start + 1, 1)
    .getValues()
    .flat();

  for (let i = values.length - 1; i >= 0; i--) {
    if (normalize_(values[i]) !== "") return start + i;
  }
  return start;
}

/** ✅ Asegura fórmula por fila en E: Cantidad(B) * NumDias(C) * ValorUnitario(D) */
function ensureRowFormulaTransporte_(sh, row) {
  // E = B * C * D (si A está vacío, deja vacío)
  sh.getRange(row, 5).setFormula(`=IF($A${row}="","",$B${row}*$C${row}*$D${row})`);
}

/** ✅ Asegura fórmula del TOTAL (celda gris en E de la fila donde dice TOTAL en D) */
function ensureTotalFormulaTransporte_(sh) {
  const totalRow = getTotalRowTransporte_(sh);
  const start = CFG_TRANSPORTE.tableStartRow;

  // Suma E16 : E(fila anterior a TOTAL)
  // INDEX(E:E, ROW()-1) devuelve E(totalRow-1) cuando la fórmula está en la fila totalRow
  sh.getRange(totalRow, CFG_TRANSPORTE.totalValueCol).setFormula(
    `=IFERROR(SUM(E${start}:INDEX(E:E,ROW()-1)),0)`
  );
}

function BotonIngresartransporte() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_TRANSPORTE.sheetName);
  if (!sh) throw new Error(`No existe la hoja "${CFG_TRANSPORTE.sheetName}"`);

  const input = sh.getRange(CFG_TRANSPORTE.inputCellA1);
  const desc = normalize_(input.getValue());

  if (!desc) {
    uiAlertSafe_("Aplicativo AOP", "Agregue una Descripción.");
    sh.setActiveRange(input);
    return;
  }

  // Re-asegura TOTAL (por si alguien borró fórmulas)
  ensureTotalFormulaTransporte_(sh);

  const baseCell = sh.getRange(CFG_TRANSPORTE.tableStartRow, CFG_TRANSPORTE.tableItemCol); // A16
  const baseIsEmpty = normalize_(baseCell.getValue()) === "";

  if (baseIsEmpty) {
    baseCell.setValue(desc);

    // ✅ fórmula por fila en E16
    ensureRowFormulaTransporte_(sh, CFG_TRANSPORTE.tableStartRow);

  } else {
    // Fila donde está el rótulo TOTAL (insertaremos justo antes)
    const totalRow = getTotalRowTransporte_(sh);

    // ✅ Insertar SIEMPRE antes de TOTAL
    sh.insertRowBefore(totalRow);
    const newRow = totalRow; // la fila insertada ocupa el antiguo totalRow

    // Copiar plantilla (A:F) desde templateRow (no desde lastRow)
    sh.getRange(CFG_TRANSPORTE.templateRow, 1, 1, CFG_TRANSPORTE.tableLastCol)
      .copyTo(sh.getRange(newRow, 1, 1, CFG_TRANSPORTE.tableLastCol), { contentsOnly: false });

    // Limpia celdas manuales (no tocar E porque la vamos a forzar con fórmula)
    sh.getRange(newRow, 2).clearContent(); // B
    sh.getRange(newRow, 3).clearContent(); // C
    sh.getRange(newRow, 4).clearContent(); // D
    sh.getRange(newRow, 6).clearContent(); // F

    // Escribe descripción
    sh.getRange(newRow, 1).setValue(desc);

    // ✅ fórmula por fila en E(newRow)
    ensureRowFormulaTransporte_(sh, newRow);

    // ✅ Re-asegura TOTAL porque la fila TOTAL se movió 1 hacia abajo
    ensureTotalFormulaTransporte_(sh);
  }

  input.setValue("");
  sh.setActiveRange(input);
}

function BotonEliminartransporte() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_TRANSPORTE.sheetName);
  if (!sh) throw new Error(`No existe la hoja "${CFG_TRANSPORTE.sheetName}"`);

  const input = sh.getRange(CFG_TRANSPORTE.inputCellA1);
  const cell = sh.getActiveCell();

  if (cell.getColumn() !== CFG_TRANSPORTE.tableItemCol) {
    uiAlertSafe_("AOP", "Seleccione la descripción que desea eliminar (columna A).");
    sh.setActiveRange(input);
    return;
  }

  const row = cell.getRow();
  const totalRow = getTotalRowTransporte_(sh);
  const lastRow = getLastRowTransporte_(sh);

  // validar: debe estar dentro del bloque de items (A16 .. antes de TOTAL)
  if (row < CFG_TRANSPORTE.tableStartRow || row > (totalRow - 1)) {
    uiAlertSafe_("AOP", "Seleccione una fila válida dentro de la tabla (antes de TOTAL).");
    sh.setActiveRange(input);
    return;
  }

  if (lastRow === CFG_TRANSPORTE.tableStartRow) {
    // solo queda la fila base -> limpiar celdas (sin tocar E si es fórmula)
    CFG_TRANSPORTE.clearColsA1.forEach(a1 => sh.getRange(a1).setValue(""));
    // reponer fórmula de E16 y TOTAL por seguridad
    ensureRowFormulaTransporte_(sh, CFG_TRANSPORTE.tableStartRow);
    ensureTotalFormulaTransporte_(sh);

    sh.setActiveRange(input);
    return;
  }

  sh.deleteRow(row);

  // Re-asegura total y deja fórmula correcta en filas restantes
  ensureTotalFormulaTransporte_(sh);

  sh.setActiveRange(input);
}


function Abrir_Formulario_Transporte() {
  const url = "https://docs.google.com/forms/d/e/1FAIpQLScgKYVBdb-1ORorLmeAOow70u7hXNWm9nhfqwRsu6UctjppiA/viewform";

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:14px;">
      <div style="font-size:14px;font-weight:700;margin-bottom:10px;">
        Formulario Transporte
      </div>

      <button
        style="padding:10px 14px;cursor:pointer;border:1px solid #ccc;border-radius:8px;"
        onclick="window.open('${url}', '_blank'); google.script.host.close();">
        Abrir formulario en nueva pestaña
      </button>

      <div style="margin-top:12px;font-size:12px;color:#444;">
        Si tu navegador bloquea la ventana emergente, usa este link:
        <div style="margin-top:6px;">
          <a href="${url}" target="_blank" rel="noopener noreferrer">${url}</a>
        </div>
      </div>
    </div>
  `).setWidth(520).setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, "AOP");
}


