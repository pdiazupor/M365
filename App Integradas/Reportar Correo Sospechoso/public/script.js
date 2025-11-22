/* script.js */

Office.onReady(function (info) {
  // Podríamos hacer algo al inicializar si es necesario
});

/**
 * Funcion llamada por el comando "Reportar Correo Sospechoso"
 * definida en el manifest (FunctionName = reportarCorreoSospechoso)
 */
function reportarCorreoSospechoso(event) {
  try {
    const item = Office.context.mailbox.item;

    // Asunto original del correo
    const originalSubject = item.subject || "(sin asunto)";

    // Nuevo asunto con la nomenclatura definida
    const nuevoAsunto = `[SGSI Correo Sospechoso] – ${originalSubject}`;

    // Crear correo nuevo hacia la casilla de Seguridad TI
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["pdiaz@ultraport.cl"],          // solo buzón de seguridad
      subject: nuevoAsunto,
      htmlBody:
        "<p>Correo sospechoso reportado por el usuario.</p>" +
        "<p>Por favor revise el mensaje original en la bandeja del usuario.</p>" +
        "<hr/>"
    });

  } catch (e) {
    // En caso de error mostramos un aviso simple
    Office.context.mailbox.item.notificationMessages.replaceAsync("sgsiError", {
      type: "errorMessage",
      message: "No se pudo iniciar el reporte de correo sospechoso."
    });
  } finally {
    // SIEMPRE llamar a completed() para que Outlook sepa que terminamos
    event.completed();
  }
}

// Registrar la función para los comandos (Outlook moderno)
Office.actions && Office.actions.associate &&
  Office.actions.associate("reportarCorreoSospechoso", reportarCorreoSospechoso);
