// correo.js - Funciones para el cierre de turno y envío de correos.

/**
 * Cierra el turno actual, archiva los registros por correo y limpia la hoja de registro.
 * Esta función es típicamente llamada al final de una jornada.
 *
 * @param {string} usuario - El usuario o sistema que ejecuta el cierre de turno.
 * @returns {Object} Resultado de la operación: { exito: boolean, mensaje: string }.
 */
function cerrarTurno(usuario) {
  try {
    const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
    if (!hojaRegistro) {
      const mensajeError = `Hoja de registro "${SHEET_REGISTRO}" no encontrada. No se pudo cerrar el turno.`;
      console.error(`❌ ${mensajeError}`);
      if (typeof registrarLog === 'function') {
        registrarLog('ERROR', 'Fallo al cerrar turno', { error: mensajeError, usuario: usuario });
      }
      return { exito: false, mensaje: mensajeError };
    }

    const datos = hojaRegistro.getDataRange().getValues();

    // Obtener el destinatario del correo desde la configuración del sistema
    let destinatario = 'admin@ejemplo.com'; // Destinatario por defecto si no se encuentra en la configuración
    if (typeof obtenerConfiguracion === 'function') {
      const configResult = obtenerConfiguracion();
      if (configResult.success && configResult.config && configResult.config.NOTIFICACIONES_EMAIL) {
        destinatario = configResult.config.NOTIFICACIONES_EMAIL;
      } else {
        console.warn(`⚠️ No se pudo obtener el email de notificación desde la configuración. Usando default: ${destinatario}`);
      }
    } else {
      console.warn(`⚠️ La función 'obtenerConfiguracion' no está disponible. Usando email default: ${destinatario}`);
    }

    // Preparar el asunto y cuerpo del correo
    const fechaCierre = new Date();
    const asunto = `Cierre de turno - ${Utilities.formatDate(fechaCierre, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')}`;
    const cuerpo = `Adjunto el registro completo del día hasta el cierre de turno.\n\nUsuario que cerró el turno: ${usuario || 'N/A'}`;

    // Convertir los datos a formato CSV
    const csvContent = datos.map(fila => fila.join(",")).join("\n");
    const archivoBlob = Utilities.newBlob(csvContent, 'text/csv', `registro_cierre_turno_${Utilities.formatDate(fechaCierre, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.csv`);

    // Enviar el correo electrónico
    GmailApp.sendEmail(destinatario, asunto, cuerpo, { attachments: [archivoBlob] });
    console.log(`✅ Correo de cierre de turno enviado a: ${destinatario}`);

    // Limpiar la hoja de registro principal después de enviar el correo
    // CORRECCIÓN: Se llama a 'limpiarRegistros()' en lugar de la función inexistente 'limpiarRegistro()'
    if (typeof limpiarRegistros === 'function') {
      const limpiezaResult = limpiarRegistros();
      if (!limpiezaResult.exito) {
        console.warn(`⚠️ Problema al limpiar registros después de enviar el correo: ${limpiezaResult.mensaje}`);
      }
    } else {
      console.error("❌ La función 'limpiarRegistros' no está definida o accesible.");
      // Considerar una limpieza manual de la hoja si la función no es accesible.
    }

    // Registrar el evento en el sistema de logs (si está disponible)
    if (typeof registrarLog === 'function') {
      registrarLog('INFO', 'Turno cerrado y enviado por correo', { usuario: usuario, destinatario: destinatario, registros_enviados: datos.length - 1 }, usuario);
    } else {
      console.log(`📝 Cierre de turno: Turno cerrado y correo enviado por ${usuario} a ${destinatario}.`);
    }

    return { exito: true, mensaje: 'Turno cerrado y registros enviados por correo.' };

  } catch (error) {
    console.error(`❌ Error en 'cerrarTurno': ${error.message}`);
    // Registrar el error en el sistema de logs (si está disponible)
    if (typeof registrarLog === 'function') {
      registrarLog('ERROR', 'Error al cerrar el turno', { error: error.message, stack: error.stack, usuario: usuario });
    }
    return { exito: false, mensaje: `Error al cerrar el turno: ${error.message}` };
  }
}

/**
 * Envía un correo electrónico con un asunto, cuerpo y adjunto opcional.
 * Esta es una función auxiliar generalizada para el envío de correos.
 *
 * @param {string} destinatario - La dirección de correo electrónico del destinatario.
 * @param {string} asunto - El asunto del correo.
 * @param {string} cuerpo - El cuerpo del mensaje del correo.
 * @param {GoogleAppsScript.Base.BlobSource} [adjunto] - Un archivo Blob para adjuntar al correo (opcional).
 * @returns {Object} Resultado de la operación: { exito: boolean, mensaje: string }.
 */
function enviarCorreoResumen(destinatario, asunto, cuerpo, adjunto) {
  try {
    const opciones = adjunto ? { attachments: [adjunto] } : {};
    GmailApp.sendEmail(destinatario, asunto, cuerpo, opciones);
    console.log(`✅ Correo resumen enviado a: ${destinatario} con asunto: "${asunto}"`);

    // Registrar el evento en el sistema de logs (si está disponible)
    if (typeof registrarLog === 'function') {
      registrarLog('INFO', 'Correo resumen enviado', { destinatario: destinatario, asunto: asunto, adjunto: !!adjunto });
    }

    return { exito: true, mensaje: 'Correo enviado correctamente.' };
  } catch (error) {
    console.error(`❌ Error en 'enviarCorreoResumen': ${error.message}`);
    // Registrar el error en el sistema de logs (si está disponible)
    if (typeof registrarLog === 'function') {
      registrarLog('ERROR', 'Error al enviar correo resumen', { error: error.message, destinatario: destinatario, asunto: asunto });
    }
    return { exito: false, mensaje: `Error al enviar correo: ${error.message}` };
  }
}