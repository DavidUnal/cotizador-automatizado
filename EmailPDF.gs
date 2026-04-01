/***********************
 * BOTÓN: abrir modal de email
 ***********************/
function BotonEmailPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMMdd_HHmmss");
  const suggestedSubject = `Cotización - ${ts}`;

  const tpl = HtmlService.createTemplateFromFile("EmailDialog");
  tpl.defaultSubject = suggestedSubject;

  SpreadsheetApp.getUi().showModalDialog(
    tpl.evaluate().setWidth(520).setHeight(420),
    "Enviar cotización por correo"
  );
}


function pingEmailDialog() {
  return { ok: true, msg: "Conectado con servidor" };
}



function sendOrDraftPdfEmail(payload) {
  try {
    const to = String((payload && payload.to) || "").trim();
    const subject = String((payload && payload.subject) || "").trim() || "Cotización";
    const extraMsg = String((payload && payload.message) || "");
    const makeDraft = Boolean(payload && payload.makeDraft);

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
      return { ok: false, error: "Correo destinatario inválido." };
    }

    // Log para ver en Executions que sí llegó
    console.log("sendOrDraftPdfEmail llamado. Draft:", makeDraft, "To:", to);

    const pdfBlob = makeCotizacionPdfBlob_();

    const autoHtml = buildAutoEmailHtml_();
    const extraHtml = extraMsg
      ? `<p><b>Mensaje:</b><br>${escapeHtml_(extraMsg).replace(/\n/g, "<br>")}</p>`
      : "";

    const signatureHtml = getGmailSignatureHtml_();

    const htmlBody = `
      ${autoHtml}
      ${extraHtml}
      ${signatureHtml ? "<br><br>" + signatureHtml : ""}
    `.trim();

    const options = {
      htmlBody,
      attachments: [pdfBlob]
    };

    if (makeDraft) {
      GmailApp.createDraft(to, subject, " ", options);
      return { ok: true, mode: "draft" };
    } else {
      GmailApp.sendEmail(to, subject, " ", options);
      return { ok: true, mode: "sent" };
    }

  } catch (e) {
    console.log("ERROR sendOrDraftPdfEmail:", e);
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}



/***********************
 * ACCIÓN: enviar o crear borrador con PDF adjunto
 ***********************/
function sendOrDraftPdfEmail_(payload) {
  try {
    // Normalizar inputs
    const to = String((payload && payload.to) || "").trim();
    const subject = String((payload && payload.subject) || "").trim() || "Cotización";
    const extraMsg = String((payload && payload.message) || "");
    const makeDraft = Boolean(payload && payload.makeDraft);

    // Validación básica email
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
      return { ok: false, error: "Correo destinatario inválido." };
    }

    // 1) PDF adjunto
    const pdfBlob = makeCotizacionPdfBlob_();

    // 2) Mensaje automático + adicional
    const autoHtml = buildAutoEmailHtml_();
    const extraHtml = extraMsg
      ? `<p><b>Mensaje:</b><br>${escapeHtml_(extraMsg).replace(/\n/g, "<br>")}</p>`
      : "";

    // 3) Firma real (HTML) desde Gmail (si está disponible)
    const signatureHtml = getGmailSignatureHtml_();

    const htmlBody = `
      ${autoHtml}
      ${extraHtml}
      ${signatureHtml ? "<br><br>" + signatureHtml : ""}
    `.trim();

    const options = {
      htmlBody: htmlBody,
      attachments: [pdfBlob]
    };

    // 4) Draft vs Send
    if (makeDraft) {
      GmailApp.createDraft(to, subject, " ", options);
      return { ok: true, mode: "draft" };
    } else {
      GmailApp.sendEmail(to, subject, " ", options);
      return { ok: true, mode: "sent" };
    }

  } catch (e) {
    // Siempre devolver algo serializable al HTML
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

/***********************
 * HELPER: generar PDF Cotización como Blob
 ***********************/
function makeCotizacionPdfBlob_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG_PDF_COT.sheetName);
  if (!sh) throw new Error(`No existe la hoja "${CFG_PDF_COT.sheetName}".`);

  // FORZAR solo hasta columna G
  const END_COL_G = 7;

  // 1) lastRow por contenido
  const lastByContent = getLastMeaningfulRowInRect_(
    sh,
    CFG_PDF_COT.startRow,
    CFG_PDF_COT.startCol,
    END_COL_G,
    CFG_PDF_COT.maxScanRows
  );

  // 2) lastRow por marcador
  const lastByMarker = findLastRowByMarker_(sh, "Forma de pago", /*extraRows=*/2);
  const lastRow = lastByMarker ? Math.min(lastByContent, lastByMarker) : lastByContent;

  // Export bounds
  const r1 = CFG_PDF_COT.startRow - 1;
  const c1 = CFG_PDF_COT.startCol - 1;
  const r2 = lastRow;
  const c2 = END_COL_G;


  const consecutivo = String(sh.getRange("B21").getDisplayValue()).trim(); // B21:G21 (normalmente combinado)
  const actividad   = String(sh.getRange("B16").getDisplayValue()).trim(); // B16:G16 (normalmente combinado)

  let baseName = [consecutivo, actividad].filter(Boolean).join(" - ");

  // Sanitizar para evitar caracteres inválidos en Windows / Gmail
  baseName = baseName
    .replace(/[\\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMMdd_HHmmss");
  const pdfName = baseName ? `${baseName}.pdf` : `${CFG_PDF_COT.fileNamePrefix}_${ts}.pdf`;


  const pdfBlob = exportSheetRectToPdfBlob_(
    ss.getId(),
    sh.getSheetId(),
    pdfName,
    r1,
    r2,
    c1,
    c2
  );

  return pdfBlob.setName(pdfName);
}

/***********************
 * HELPER: mensaje automático estándar
 ***********************/
function buildAutoEmailHtml_() {
  return `
    <p>Hola,</p>
    <p>
      Te comparto la cotización adjunta en PDF.
      Si necesitas algún ajuste o tienes preguntas, con gusto lo revisamos.
    </p>
    <p>Quedo atento(a).</p>
  `.trim();
}

/***********************
 * HELPER: escape HTML para el mensaje adicional
 ***********************/
function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/***********************
 * HELPER: firma real desde Gmail Settings (Gmail API Advanced Service)
 * - Si no está habilitada o no hay permisos, devuelve "" y no rompe.
 ***********************/
function getGmailSignatureHtml_() {
  try {
    // Email principal del usuario
    const profile = Gmail.Users.getProfile("me");
    const sendAsId = profile.emailAddress;

    // Settings del "send as" principal
    const sendAs = Gmail.Users.Settings.SendAs.get("me", sendAsId);
    return (sendAs && sendAs.signature) ? sendAs.signature : "";
  } catch (e) {
    return "";
  }
}

/***********************
 * (Opcional) Test para forzar permisos una vez
 ***********************/
function testAuthEmail() {
  // Cambia por TU correo real
  GmailApp.createDraft(Session.getActiveUser().getEmail(), "Test permisos", "ok");
}
