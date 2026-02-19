// --- VARIABLES DE ENTORNO ---
const SPREADSHEET_ID = '1OGGX4C_BmP-pD8_G0hDpYZ2sMFN2_5gXoaGcuFTD53g'
const SHEET_NAME = 'Respuestas'
const UPLOAD_FOLDER_ID = '1UGOjmXS474Cokr3MEySfnm8tTT_fzBw7'
// Webhook URL (Pega tu URL aquí)
const WEBHOOK_URL =
  'https://n8n.balticec.com/webhook-test/ade824be-ccdf-4779-a05c-c0888edc8a27'
// ----------------------------

function doPost(e) {
  try {
    // Si los datos vienen como JSON string en el postData
    var params = JSON.parse(e.postData.contents)
    var resultado = processExternalForm(params)

    return ContentService.createTextOutput(
      JSON.stringify(resultado),
    ).setMimeType(ContentService.MimeType.JSON)
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: 'Error en doPost: ' + error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON)
  }
}

function processExternalForm(data) {
  // 1. Validar acceso a Spreadsheet
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID)
  let sheet = ss.getSheetByName(SHEET_NAME)
  if (!sheet) {
    sheet = ss.getSheets()[0]
  }

  // 2. Manejar archivo (Base64 desde el cliente)
  let fileUrl = ''
  if (data.constancia && data.constancia.data) {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID)
    // data.constancia.data debe ser el string base64 sin el encabezado 'data:application/pdf;base64,'
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.constancia.data),
      data.constancia.mimeType,
      data.constancia.name,
    )

    const newName = `${data.dni || 'SIN_DNI'}_${data.apellidos || ''}_Constancia.pdf`
    const file = folder.createFile(blob).setName(newName)
    fileUrl = file.getUrl()

    // Agregamos la URL y el ID al objeto data para enviarla al webhook/email
    data.fileUrl = fileUrl
    data.fileId = file.getId()

    // Eliminamos el objeto constancia (base64 pesado) para no enviarlo al webhook ni correo
    delete data.constancia
  }

  // 3. Preparar fila
  const rowData = [
    new Date(),
    data.nombres || '',
    data.apellidos || '',
    "'" + (data.dni || ''),
    data.correo_estudiantil || '',
    data.programa || '',
    data.ciclo || '',

    // Organización
    data.razon_social || '',
    "'" + (data.ruc || ''),
    data.jefe || '',
    data.fecha_inicio || '',
    data.horas || '',
    data.desc || '',

    // Descripción
    data.cargo || '',
    data.actividades || '',
    fileUrl,
  ]

  sheet.appendRow(rowData)

  // 4. Enviar notificación al correo del estudiante
  if (data.correo_estudiantil) {
    sendConfirmationEmail(data)
  }

  // 5. Enviar a Webhook
  if (WEBHOOK_URL) {
    sendToWebhook(data)
  }

  return { success: true, message: 'Registro guardado exitosamente.' }
}

function sendConfirmationEmail(data) {
  const subject = '✅ Confirmación de Registro EFSRT - Instituto Neumann'

  // Plantilla HTML simple y limpia
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f9fafb; padding: 20px; border-radius: 10px;">
      <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <div style="text-align: center; margin-bottom: 20px;">
           <h2 style="color: #7c3aed; margin: 0;">Registro Exitoso</h2>
           <p style="color: #6b7280; font-size: 14px;">Experiencia Formativa en Situaciones Reales de Trabajo</p>
        </div>
        
        <p style="color: #374151; font-size: 16px;">Hola <strong>${data.nombres}</strong>,</p>
        
        <p style="color: #374151; line-height: 1.6;">
          Tu formulario de registro EFSRT ha sido recibido correctamente. A continuación, un resumen de la información registrada:
        </p>
        
        <div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <ul style="list-style: none; padding: 0; margin: 0; color: #4b5563; font-size: 14px;">
            <li style="margin-bottom: 8px;"><strong>Programa:</strong> ${data.programa}</li>
            <li style="margin-bottom: 8px;"><strong>Ciclo:</strong> ${data.ciclo}</li>
            <li style="margin-bottom: 8px;"><strong>Empresa:</strong> ${data.razon_social}</li>
            <li style="margin-bottom: 8px;"><strong>Cargo:</strong> ${data.cargo}</li>
            <li><strong>Fecha Inicio:</strong> ${data.fecha_inicio}</li>
          </ul>
        </div>
        
        <p style="color: #374151; font-size: 14px;">
          Hemos adjuntado tu constancia en nuestros registros. Si tienes alguna duda, contacta con tu coordinador de carrera.
        </p>
        
        <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 30px 0;">
        
        <div style="text-align: center; color: #9ca3af; font-size: 12px;">
          <p>© ${new Date().getFullYear()} Instituto Neumann. Todos los derechos reservados.</p>
        </div>
      </div>
    </div>
  `

  MailApp.sendEmail({
    to: data.correo_estudiantil,
    subject: subject,
    htmlBody: htmlBody,
  })
}

function sendToWebhook(data) {
  try {
    const payload = JSON.stringify(data)
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: payload,
      muteHttpExceptions: true, // Para que no falle el script si el webhook responde error
    }
    UrlFetchApp.fetch(WEBHOOK_URL, options)
  } catch (e) {
    Logger.log('Error enviando webhook: ' + e.toString())
  }
}
