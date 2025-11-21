Office.onReady(() => {});

function reportarCorreoSospechoso(event) {
    const item = Office.context.mailbox.item;

    // 1. Forward nativo
    item.forwardAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const forwardMessage = result.value;

            // 2. Ajustar asunto
            forwardMessage.subject = "[SGSI Correo Sospechoso] " + item.subject;

            // 3. Agregar texto arriba del correo reenviado
            forwardMessage.body.setAsync(
                "Reporte automático desde Outlook.\n\n",
                { coercionType: Office.CoercionType.Text },
                () => {
                    // 4. Configurar destinatario
                    forwardMessage.to.setAsync(["pdiaz@ultraport.cl"], () => {
                        // 5. Guardar y enviar
                        forwardMessage.saveAsync(() => {
                            forwardMessage.sendAsync(() => {
                                event.completed();
                            });
                        });
                    });
                }
            );
        } else {
            console.error("Error al reenviar:", result.error);
            event.completed();
        }
    });
}
