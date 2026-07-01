Office.onReady();

// ======================================================
// CONFIGURACIÓN
// Cada cliente tiene una keyword (visible, va en el nombre del
// adjunto) y un hash SHA-256 de su dominio de correo (NO revela
// el dominio real, solo se usa para comparar).
//
// Para obtener el hash de tu dominio real, usa el archivo
// calculadora-hash-LOCAL-NO-SUBIR.html en tu computadora
// (nunca subas ese archivo a GitHub).
// ======================================================
const REGLAS_CLIENTES = [
  {
    keyword: "Tenant 01",
    dominioHash: "PEGA_AQUI_EL_HASH_DEL_DOMINIO_TENANT_01"
  },
  {
    keyword: "Tenant 02",
    dominioHash: "PEGA_AQUI_EL_HASH_DEL_DOMINIO_TENANT_02"
  }
];

async function sha256(texto) {
  const buffer = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(texto.trim().toLowerCase())
  );
  return Array.from(new Uint8Array(buffer))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

function obtenerDominio(correo) {
  const partes = (correo || "").toLowerCase().split("@");
  return partes.length === 2 ? partes[1] : "";
}
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
        att.name.toLowerCase().includes(regla.keyword.toLowerCase())
      )
    );

    if (reglasDetectadas.length === 0) {
      // Ningún adjunto coincide con las palabras clave configuradas
      event.completed({ allowEvent: true });
      return;
    }

    item.to.getAsync(function (toResult) {
      item.cc.getAsync(async function (ccResult) {
        const destinatarios = []
          .concat(toResult.value || [])
          .concat(ccResult.value || [])
          .map(r => obtenerDominio(r.emailAddress))
          .filter(Boolean);

        // Calcula el hash de cada dominio de destinatario presente
        const hashesDestinatarios = await Promise.all(
          destinatarios.map(sha256)
        );

        let reglaIncumplida = null;
        for (const regla of reglasDetectadas) {
          if (!hashesDestinatarios.includes(regla.dominioHash)) {
            reglaIncumplida = regla;
            break;
          }
        }

        if (reglaIncumplida) {
          event.completed({
            allowEvent: false,
            errorMessage:
              "El archivo adjunto corresponde a '" +
              reglaIncumplida.keyword +
              "' pero el destinatario no coincide con ese cliente. Verifica antes de enviar."
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
