// ============================================================
//  CANAL 10 — Apps Script para Pedidos de Edición
//  Pegá este código en: Extensions > Apps Script del Sheets
//  Luego desplegalo como Web App (ver instrucciones abajo)
// ============================================================

var ENCABEZADOS = [
  "N° Ticket",
  "Fecha y hora",
  "Área",
  "Tipo de pieza",
  "Prioridad",
  "Duración",
  "Deadline",
  "Material / Ubicación",
  "Descripción",
  "Solicitante",
  "Estado"
];

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Si la hoja está vacía, escribir encabezados
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(ENCABEZADOS);

      // Estilo de encabezados
      var headerRange = sheet.getRange(1, 1, 1, ENCABEZADOS.length);
      headerRange.setBackground("#C0392B");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      headerRange.setFontSize(11);
      sheet.setFrozenRows(1);

      // Anchos de columna
      sheet.setColumnWidth(1, 110);  // N° Ticket
      sheet.setColumnWidth(2, 140);  // Fecha y hora
      sheet.setColumnWidth(3, 110);  // Área
      sheet.setColumnWidth(4, 140);  // Tipo
      sheet.setColumnWidth(5, 140);  // Prioridad
      sheet.setColumnWidth(6, 100);  // Duración
      sheet.setColumnWidth(7, 140);  // Deadline
      sheet.setColumnWidth(8, 180);  // Material
      sheet.setColumnWidth(9, 320);  // Descripción
      sheet.setColumnWidth(10, 140); // Solicitante
      sheet.setColumnWidth(11, 110); // Estado
    }

    // Parsear datos del formulario
    var datos = JSON.parse(e.postData.contents);

    // Generar número de ticket
    var ultimaFila = sheet.getLastRow();
    var anio = new Date().getFullYear();
    var num = ultimaFila; // fila 1 = encabezado, fila 2 = primer pedido (#1)
    var ticket = "C10-" + anio + "-" + String(num).padStart(3, "0");

    // Fecha y hora actual Argentina
    var ahora = Utilities.formatDate(
      new Date(),
      "America/Argentina/Buenos_Aires",
      "dd/MM/yyyy HH:mm"
    );

    // Agregar fila
    var fila = [
      ticket,
      ahora,
      datos.area || "",
      datos.tipo || "",
      datos.prio || "",
      datos.dur || "",
      datos.deadline || "",
      datos.ubic || "",
      datos.desc || "",
      datos.nombre || "",
      "PENDIENTE"
    ];

    sheet.appendRow(fila);

    // Color de fila según prioridad
    var nuevaFila = sheet.getLastRow();
    var rangoFila = sheet.getRange(nuevaFila, 1, 1, ENCABEZADOS.length);

    var colorFila = "#FFFFFF";
    if (datos.prio && datos.prio.includes("P1")) colorFila = "#FADBD8"; // rojo suave
    else if (datos.prio && datos.prio.includes("P2")) colorFila = "#FDEBD0"; // naranja suave
    else if (datos.prio && datos.prio.includes("P3")) colorFila = "#EAFAF1"; // verde suave
    else if (datos.prio && datos.prio.includes("P4")) colorFila = "#EBF5FB"; // azul suave

    rangoFila.setBackground(colorFila);

    // Celda Estado con formato
    var celdaEstado = sheet.getRange(nuevaFila, 11);
    celdaEstado.setBackground("#F9E79F");
    celdaEstado.setFontWeight("bold");

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, ticket: ticket }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Permite testear desde el navegador
function doGet(e) {
  return ContentService
    .createTextOutput("Canal 10 — Apps Script activo ✓")
    .setMimeType(ContentService.MimeType.TEXT);
}
