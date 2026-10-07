/* Llenar cuerpo WF - Complemento de Outlook
 * Edita SOLO la tabla AGENCIAS para agregar/cambiar agencias.
 * La clave es una palabra/frase que debe aparecer en el nombre de la lista de distribución (sin importar mayúsculas ni acentos). */

const AGENCIAS = [
  // [ frase(s) que identifican el nombre de la lista, OSCAC, CAAT ]
  { match: ["wonder fields nogales", "aduanal nogales wonder", "nogales wonder"], oscac: "PYSR", caat: "32VT" },
  { match: ["wonder fields mcallen", "wonder mcallen"], oscac: "PJDM", caat: "3J50" },
  { match: ["aduanal mcallen lipman", "mcallen lipman", "lipman"], oscac: "PJDM", caat: "3J50" },
];

const CAJA = "MEXICANA";
const MEDIDA = "53´";
const SALTAR_FIN_DE_SEMANA = false; // true = si mañana es sábado/domingo, pone el lunes

function norm(s) {
  return (s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim();
}

function fechaCruce() {
  const d = new Date();
  d.setDate(d.getDate() + 1);
  if (SALTAR_FIN_DE_SEMANA) {
    if (d.getDay() === 6) d.setDate(d.getDate() + 2);
    if (d.getDay() === 0) d.setDate(d.getDate() + 1);
  }
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getFullYear()}`;
}

function economicoDeAsunto(asunto) {
  const m = (asunto || "").match(/^\s*\d+\s*\/\s*\d+\s*\/\s*(\d+)/);
  return m ? m[1] : null;
}

function buscarAgencia(nombres) {
  for (const n of nombres) {
    const nn = norm(n);
    for (const a of AGENCIAS) {
      if (a.match.some((k) => nn.includes(norm(k)))) return a;
    }
  }
  return null;
}

function construirHtml(fecha, a, eco) {
  return (
    `<div><span style="background-color:#FFFF00"><u>CRUCE MAÑANA ${fecha}</u></span></div>` +
    `<div><br></div>` +
    `<div>OSCAC: ${a.oscac}</div>` +
    `<div>CAAT: ${a.caat}</div>` +
    `<div>ECONOMICO: ${eco}</div>` +
    `<div>CAJA:&nbsp; ${CAJA}</div>` +
    `<div>MEDIDA: ${MEDIDA}</div>` +
    `<div><br></div>` +
    `<div>Anexo documentos.</div>` +
    `<div><br></div><div><br></div>`
  );
}

function aviso(event, mensaje, tipo) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("wf", {
    type: tipo === "ok" ? "informationalMessage" : "errorMessage",
    message: mensaje,
    icon: "icon16",
    persistent: false,
  }, () => event.completed());
}

function getAsync(fn) {
  return new Promise((res, rej) => fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error))));
}

async function llenarCuerpo(event) {
  try {
    const item = Office.context.mailbox.item;
    const asunto = await getAsync((cb) => item.subject.getAsync(cb));
    const eco = economicoDeAsunto(asunto);
    if (!eco) return aviso(event, "Asunto no válido. Formato: 6943/3443/1045 Exportación...", "err");

    const para = await getAsync((cb) => item.to.getAsync(cb));
    const nombres = para.map((r) => r.displayName).concat(para.map((r) => r.emailAddress));
    const ag = buscarAgencia(nombres);
    if (!ag) return aviso(event, "No reconozco la lista en PARA: " + (para[0] ? para[0].displayName : "(vacío)"), "err");

    const texto = await getAsync((cb) => item.body.getAsync(Office.CoercionType.Text, cb));
    if (/CRUCE MA.ANA/i.test(texto)) return aviso(event, "El cuerpo ya tiene datos de CRUCE. No se duplicó.", "err");

    await getAsync((cb) => item.body.prependAsync(construirHtml(fechaCruce(), ag, eco), { coercionType: Office.CoercionType.Html }, cb));
    aviso(event, `Listo: ECONOMICO ${eco}, OSCAC ${ag.oscac}, CAAT ${ag.caat}.`, "ok");
  } catch (e) {
    aviso(event, "Error: " + (e && e.message ? e.message : e), "err");
  }
}

Office.onReady(() => {
  Office.actions.associate("llenarCuerpo", llenarCuerpo);
});

if (typeof module !== "undefined") module.exports = { norm, fechaCruce, economicoDeAsunto, buscarAgencia, construirHtml };
