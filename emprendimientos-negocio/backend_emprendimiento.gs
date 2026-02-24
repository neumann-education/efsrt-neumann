// --- CONFIGURACIÓN ---
const SPREADSHEET_ID = '1iH0GGY9um05qRxpXVH7hhrwhLoBp91CA3oAHgIM1gzg'
const SHEET_NAME = 'Respuestas'
const UPLOAD_FOLDER_ID = '1k-iHQ8NT_EDxEybKhQXbkaAUuGoVmDDt'
const WEBHOOK_URL =
  'https://n8n.balticec.com/webhook-test/ade824be-ccdf-4779-a05c-c0888edc8a27'

// ---------------------------------------------------

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents)
    const resultado = processForm(params)

    return ContentService.createTextOutput(
      JSON.stringify(resultado),
    ).setMimeType(ContentService.MimeType.JSON)
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: 'Error en el servidor: ' + error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON)
  }
}

// ---------------------------------------------------

function processForm(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID)
  let sheet = ss.getSheetByName(SHEET_NAME)
  if (!sheet) sheet = ss.getSheets()[0]

  // ---------- 1. ARCHIVO ----------
  let fileUrl = ''
  let fileId = ''

  if (data.adjunto && data.adjunto.data) {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID)

    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.adjunto.data),
      data.adjunto.mimeType,
      data.adjunto.name,
    )

    const ext = data.adjunto.name.split('.').pop()
    const tipoArchivo =
      data.tipo === 'Iniciativa de negocio' ? 'PlanNegocio' : 'Evidencias'

    const fileName = `${data.dni || 'SINDNI'}_${data.apellidos || 'Estudiante'}_${tipoArchivo}.${ext}`

    const file = folder.createFile(blob).setName(fileName)
    fileUrl = file.getUrl()
    fileId = file.getId()
  }

  // ---------- 2. REGISTRO ----------
  const idRegistro = Utilities.getUuid()

  sheet.appendRow([
    new Date(), // Fecha
    idRegistro, // ID
    data.nombres || '',
    data.apellidos || '',
    "'" + (data.dni || ''),
    data.correo_estudiantil || '',
    data.programa || '',
    data.ciclo || '',
    data.tipo || '',
    data.nombre_empresa || '-',
    data.marca || '-',
    data.mision || '-',
    data.vision || '-',
    data.estructura_legal || '-',
    fileUrl,
  ])

  // ---------- 3. EMAIL ----------
  if (data.correo_estudiantil) {
    sendEmail(data, idRegistro)
  }

  // ---------- 4. WEBHOOK ----------
  if (WEBHOOK_URL) {
    const payload = {
      ...data,
      id: idRegistro,
      fileId: fileId,
      fileUrl: fileUrl,
    }

    delete payload.adjunto // NUNCA enviar base64

    sendWebhook(payload)
  }

  return {
    success: true,
    id: idRegistro,
    fileId: fileId,
    fileUrl: fileUrl,
  }
}

// ---------------------------------------------------

function sendEmail(data, id) {
  const subject = `✅ Confirmación EFSRT - Emprendimiento (${id})`

  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f9fafb; padding: 20px; border-radius: 10px;">
      <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <div style="text-align: center; margin-bottom: 20px;">
           <h2 style="color: #7c3aed; margin: 0;">Registro Recibido</h2>
           <p style="color: #6b7280; font-size: 14px;">EFSRT - Emprendimiento e Iniciativas de Negocio</p>
        </div>
        
        <p style="color: #374151; font-size: 16px;">Hola <strong>${data.nombres}</strong>,</p>
        
        <p style="color: #374151; line-height: 1.6;">
          Hemos recibido tu solicitud de registro para EFSRT en la modalidad de Emprendimiento. A continuación el detalle:
        </p>
        
        <div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <ul style="list-style: none; padding: 0; margin: 0; color: #4b5563; font-size: 14px;">
            <li style="margin-bottom: 8px;"><strong>ID Registro:</strong> ${id}</li>
            <li style="margin-bottom: 8px;"><strong>Tipo:</strong> ${data.tipo}</li>
            <li style="margin-bottom: 8px;"><strong>Programa:</strong> ${data.programa}</li>
            ${data.nombre_empresa ? `<li style="margin-bottom: 8px;"><strong>Empresa:</strong> ${data.nombre_empresa}</li>` : ''}
          </ul>
        </div>

        <p style="color: #374151; font-size: 14px;">
          Tu archivo adjunto (Plan de Negocio o Evidencias) ha sido almacenado correctamente.
        </p>
        
        <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 30px 0;">
        
        <div style="text-align: center; color: #9ca3af; font-size: 12px;">
          <p>© ${new Date().getFullYear()} Instituto Neumann</p>
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

// ---------------------------------------------------

function sendWebhook(payload) {
  try {
    UrlFetchApp.fetch(WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    })
  } catch (e) {
    console.log('Error webhook:', e)
  }
}
