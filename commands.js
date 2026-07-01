Office.onReady();

// ======================================================
// CONFIGURACIÓN — AJUSTA ESTOS VALORES A TUS CLIENTES REALES
// ======================================================
const REGLAS_CLIENTES = [
  {
    // Palabra clave que debe aparecer en el nombre del archivo adjunto
    keywordAdjunto: "ClienteA",
    // Dominio (o correo específico) que debe estar en Para/CC
    dominioCorreo: "clientea.com"
  },
  {
    keywordAdjunto: "ClienteB",
    dominioCorreo: "clienteb.com"
  }
];
// ======================================================

function validarDestinatario(event) {
  const item = Office.context.mailbox.item;

  item.getAttachmentsAsync(function (attResult) {
    if (attResult.status !== Office.AsyncResultStatus.Succeeded) {
      // Si no se pueden leer adjuntos, no bloqueamos el envío
      event.completed({ allowEvent: true });
      return;
    }

    const attachments = attResult.value || [];

    // Detecta qué reglas de cliente aplican según los adjuntos presentes
    const reglasDetectadas = REGLAS_CLIENTES.filter(regla =>
      attachments.some(att =>
        att.name.toLowerCase().includes(regla.keywordAdjunto.toLowerCase())
      )
    );

    if (reglasDetectadas.length === 0) {
      // Ningún adjunto coincide con las palabras clave configuradas
      event.completed({ allowEvent: true });
      return;
    }

    item.to.getAsync(function (toResult) {
      item.cc.getAsync(function (ccResult) {
        const destinatarios = []
          .concat(toResult.value || [])
          .concat(ccResult.value || [])
          .map(r => (r.emailAddress || "").toLowerCase());

        const destinatariosStr = destinatarios.join(";");

        const reglaIncumplida = reglasDetectadas.find(
          regla => !destinatariosStr.includes(regla.dominioCorreo.toLowerCase())
        );

        if (reglaIncumplida) {
          event.completed({
            allowEvent: false,
            errorMessage:
              "El archivo adjunto corresponde a '" +
              reglaIncumplida.keywordAdjunto +
              "' pero el destinatario no coincide con ese cliente (" +
              reglaIncumplida.dominioCorreo +
              "). Verifica antes de enviar."
          });
        } else {
          event.completed({ allowEvent: true });
        }
      });
    });
  });
}

Office.actions = Office.actions || {};
Office.actions.associate("validarDestinatario", validarDestinatario);
