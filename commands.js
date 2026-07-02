Office.onReady();

// ======================================================
// CONFIGURACIÓN
// ======================================================
const REGLAS_CLIENTES = [
  {
    keyword: "Tenant 01",
    dominioHash: "52bf949d857eefa0d0e8546331f2af6eda6e48120c2ee5839566ba53c431430b"
  },
  {
    keyword: "Tenant 02",
    dominioHash: "c4ed401f810c02c4d5af59e6e52e78cbf324b4c20f4dbdf24afd834649710646"
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
  console.log("[Validador] validarDestinatario() SE EJECUTÓ");
  const item = Office.context.mailbox.item;

  item.getAttachmentsAsync(function (attResult) {
    console.log("[Validador] getAttachmentsAsync status:", attResult.status);

    if (attResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.warn("[Validador] No se pudieron leer adjuntos, se permite el envío por defecto.");
      event.completed({ allowEvent: true });
      return;
    }

    const attachments = attResult.value || [];
    console.log("[Validador] Adjuntos detectados:", attachments.map(a => a.name));

    const reglasDetectadas = REGLAS_CLIENTES.filter(regla =>
      attachments.some(att =>
        att.name.toLowerCase().includes(regla.keyword.toLowerCase())
      )
    );
    console.log("[Validador] Reglas detectadas por keyword:", reglasDetectadas.map(r => r.keyword));

    if (reglasDetectadas.length === 0) {
      console.log("[Validador] Ningún adjunto coincide con las keywords configuradas -> se permite el envío.");
      event.completed({ allowEvent: true });
      return;
    }

    item.to.getAsync(function (toResult) {
      console.log("[Validador] item.to.getAsync status:", toResult.status, toResult.value);

      item.cc.getAsync(async function (ccResult) {
        console.log("[Validador] item.cc.getAsync status:", ccResult.status, ccResult.value);

        try {
          const destinatarios = []
            .concat(toResult.value || [])
            .concat(ccResult.value || [])
            .map(r => obtenerDominio(r.emailAddress))
            .filter(Boolean);
          console.log("[Validador] Dominios de destinatarios:", destinatarios);

          const hashesDestinatarios = await Promise.all(destinatarios.map(sha256));
          console.log("[Validador] Hashes calculados de destinatarios:", hashesDestinatarios);

          let reglaIncumplida = null;
          for (const regla of reglasDetectadas) {
            console.log(
              "[Validador] Comparando regla:", regla.keyword,
              "| hash esperado:", regla.dominioHash,
              "| ¿coincide con algún destinatario?",
              hashesDestinatarios.includes(regla.dominioHash)
            );
            if (!hashesDestinatarios.includes(regla.dominioHash)) {
              reglaIncumplida = regla;
              break;
            }
          }

          if (reglaIncumplida) {
            console.log("[Validador] BLOQUEANDO envío por regla incumplida:", reglaIncumplida.keyword);
            event.completed({
              allowEvent: false,
              errorMessage:
                "El archivo adjunto corresponde a '" +
                reglaIncumplida.keyword +
                "' pero el destinatario no coincide con ese cliente. Verifica antes de enviar."
            });
          } else {
            console.log("[Validador] Todas las reglas se cumplen -> se permite el envío.");
            event.completed({ allowEvent: true });
          }
        } catch (err) {
          console.error("[Validador] ERROR inesperado durante la validación:", err);
          // Si algo truena aquí y no llamamos event.completed, Outlook
          // eventualmente deja pasar el envío sin avisar.
          event.completed({ allowEvent: true, errorMessage: "Error interno del validador: " + err.message });
        }
      });
    });
  });
}

Office.actions = Office.actions || {};
Office.actions.associate("validarDestinatario", validarDestinatario);