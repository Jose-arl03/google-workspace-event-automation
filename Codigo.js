/**
 * Script de Automatización de Procesos (BPA)
 * Proyecto: Turing IA Workspace - Día 2
 * Descripción: Integra Sheets, Calendar y Gmail con manejo de errores.
 */

function procesarNuevosRegistros() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  
  // Se asume el calendario principal del usuario que autoriza el script
  const calendarId = "primary"; 

  // Iteramos desde 1 para saltar los encabezados
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const nombre = row[0];
    const correo = row[1];
    const titulo = row[2];
    const fecha = row[3]; 
    const hora = row[4];
    const estatus = row[5];

    // Validación de Idempotencia: Solo procesar si el estatus está vacío
    if (estatus !== "PROCESADO" && nombre !== "" && correo !== "") {
      try {
        // 1. Formatear Fecha y Hora para Calendar
        // Nota: Asegúrate de que el formato en Sheets coincida con el constructor de Date
        const fechaInicio = new Date(`${fecha}T${hora}:00`);
        const fechaFin = new Date(fechaInicio.getTime() + (60 * 60 * 1000)); // Evento de 1 hora
        
        // 2. Crear Evento en Google Calendar
        CalendarApp.getCalendarById(calendarId).createEvent(titulo, fechaInicio, fechaFin, {
          description: `Evento generado automáticamente desde Workspace Automation para ${nombre}.`,
          guests: correo,
          sendInvites: true
        });

        // 3. Enviar Notificación (Con fallback para cuentas sin servicio de Gmail)
        const asunto = `[Turing IA] Confirmación de Evento: ${titulo}`;
        const mensaje = `Hola ${nombre},\n\nEl sistema de automatización ha creado exitosamente el evento "${titulo}".\n\nRevisa tu calendario.\n\nSaludos,\nEquipo de Automatización.`;
        
        try {
          // Intentamos usar MailApp, que a veces es más flexible que GmailApp
          MailApp.sendEmail(correo, asunto, mensaje);
        } catch (mailError) {
          // Si la cuenta no tiene correo habilitado, evitamos que el script colapse
          Logger.log(`Advertencia controlada: No se pudo enviar el correo a ${correo} (Cuenta sin Gmail). El flujo continúa.`);
        }

        // 4. Actualizar Estatus para evitar bucles (Idempotencia)
        sheet.getRange(i + 1, 6).setValue("PROCESADO");
        Logger.log(`Fila ${i + 1} procesada exitosamente para ${correo}`);
        
      } catch (error) {
        // Manejo de errores
        Logger.log(`Error crítico en la fila ${i + 1}: ${error.message}`);
        sheet.getRange(i + 1, 6).setValue("ERROR");
      }
    }
  }
}