Office.onReady(() => {});

function forwardAndSend(event) {
    const item = Office.context.mailbox.item;

    // 1. Forward nativo (mantiene el cuerpo exacto)
    item.forwardAsync((result) => {
        const forwardMessage = result.value;

        // 2. Ajustar asunto
        forwardMessage.subject = "[SGSI Correo Sospechoso] " + item.subject;

        // 3. Agregar texto arriba del correo reenviado
        forwardMessage.body.setAsync(
            "Reporte automático desde Outlook.\n\n",
            { coercionType: Office.CoercionType.Text },
            () => {
                // 4. Obtener .eml del mensaje original
                item.getAttachmentContentAsync(item.itemId, (contentResult) => {
                    // Esto NO funciona directamente porque la API no permite leer el item como .eml
                    // Necesitamos usar EWS o Graph API para obtener el .eml real.
                    // Para la prueba se adjunta como mensaje original del forwarding.

                    // 5. Configurar envío
                    forwardMessage.to.setAsync(["pdiaz@ultraport.cl"]);
                    forwardMessage.saveAsync(() => {
                        forwardMessage.sendAsync();
                        event.completed();
                    });
                });
            }
        );
    });
}
