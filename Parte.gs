function sendRangeAsPDF() {
  // ID 
  const spreadsheetId = "1M8F8qLhuTd9wZ9nq1cEEuzIA5eDniVcd7S9572aYcY0";
  const sheetName = "Parte Diario"; // Nombre de la hoja original
  const range = "A1:J45"; // Rango de celdas que deseas exportar

  // Abre el archivo y la hoja especificada
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const originalSheet = spreadsheet.getSheetByName(sheetName);

  // Verificar si la hoja especificada existe
  if (!originalSheet) {
    console.error("La hoja especificada no existe: " + sheetName);
    return;
  }

  // Obtener el valor de la celda K3 para incluir en el asunto
  const asuntoExtra = originalSheet.getRange("J1:J2").getDisplayValue();
  const subject = "Parte Diario GDI - " + asuntoExtra; // Concatenar con el valor de K3

  // Crea una nueva hoja temporal en el mismo archivo
  const tempSheet = spreadsheet.insertSheet("TempSheet");

  // Copia el rango especificado a la hoja temporal
  const rangeData = originalSheet.getRange(range);
  rangeData.copyTo(tempSheet.getRange("A1"), { contentsOnly: false }); // Mantiene los formatos

  // Confirmar que los datos han sido copiados
  Logger.log(tempSheet.getDataRange().getValues());

  // Configurar la URL de exportación en PDF para la hoja temporal
  const url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId +
              "/export?exportFormat=pdf&format=pdf&size=letter&portrait=true&fitw=true" +
              "&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&gid=" + tempSheet.getSheetId() +
              "&range=" + encodeURIComponent(range); // Asegúrate de que el rango esté definido

  try {
    // Dowload and autenthication of PDF
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: "Bearer " + token
      }
    });

    // Crea el archivo PDF a partir de la respuesta
    const pdfBlob = response.getBlob().setName("RangoSeleccionado.pdf");

    // It configurates the mail
    const email = "rrhh@contactogarantido.com";
    const ccEmails = "mariana.cruz@contactogarantido.com,sergio.bellino@contactogarantido.com,maximiliano.ortiz@gestiondeingresos.com,anabela.lopez@gestiondeingresos.com,leonel.sanagua@gestiondeingresos.com,enzo.aguirre@gestiondeingresos.com"; 
    const message = "Estimados muy buenas! <br><br>" + "Comparto el parte diario a la fecha durante el turno corriente <br><br>" + "Quedo a disposición<br><br>" + "Saludos.<br><br>";


    // Envía el correo con el archivo PDF adjunto
    MailApp.sendEmail({
      to: email,
      cc: ccEmails, // Agregar destinatarios en copia
      subject: subject,
      htmlBody: message,
      attachments: [pdfBlob]
    });
  } catch (error) {
    console.error("Error al exportar el PDF o enviar el correo: ", error);
  } finally {
    // Deletes temporal sheet
    spreadsheet.deleteSheet(tempSheet);
  }
}
