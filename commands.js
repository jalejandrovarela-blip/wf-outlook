/* Llenar cuerpo WF - Complemento de Outlook (v2)
 * Edita SOLO la tabla AGENCIAS para agregar/cambiar agencias.
 * match = frases que deben aparecer en el nombre de la lista en PARA (sin importar mayúsculas ni acentos). */

const AGENCIAS = [
  { match: ["wonder fields nogales", "aduanal nogales wonder", "nogales wonder"], oscac: "PYSR", caat: "32VT" },
  { match: ["wonder fields mcallen", "wonder mcallen", "aduanal mcallen wonder", "mcallen wonder"], oscac: "PJDM", caat: "3J50" },
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
  return dd + "/" + mm + "/" + d.getFullYear();
}

// Tercer dato del asunto: acepta 1045 y también R6, etc.
function economicoDeAsunto(asunto) {
  const m = (asunto || "").match(/^\s*\d+\s*\/\s*\d+\s*\/\s*([^\s\/]+)/);
  return m ? m[1] : null;
}

function buscarAgencia(nombres) {
  for (const n of nombres) {
    const nn = norm(n);
    if (!nn) continue;
    for (const a of AGENCIAS) {
      if (a.match.some((k) => nn.includes(norm(k)))) return a;
    }
  }
  return null;
}

function construirHtml(fecha, a, eco) {
  return (
    '<div><span style="background-color:#FFFF00"><u>CRUCE MAÑANA ' + fecha + "</u></span></div>" +
    "<div><br></div>" +
    "<div>OSCAC: " + a.oscac + "</div>" +
    "<div>CAAT: " + a.caat + "</div>" +
    "<div>ECONOMICO: " + eco + "</div>" +
    "<div>CAJA:&nbsp; " + CAJA + "</div>" +
    "<div>MEDIDA: " + MEDIDA + "</div>" +
    "<div><br></div>" +
    "<div>Anexo documentos.</div>" +
    "<div><br></div><div><br></div>"
  );
}

// Cierra SIEMPRE la función (aunque el aviso falle o nunca responda).
function terminar(event, mensaje) {
  let listo = false;
  const fin = () => {
    if (listo) return;
    listo = true;
    try { event.completed(); } catch (e) {}
  };
  setTimeout(fin, 1500); // red de seguridad
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "wf",
      { type: "informationalMessage", message: String(mensaje).slice(0, 150), icon: "Icon16", persistent: false },
      fin
    );
  } catch (e) {
    fin();
  }
}

function getAsync(fn) {
  return new Promise((res, rej) =>
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error)))
  );
}

async function llenarCuerpo(event) {
  try {
    const item = Office.context.mailbox.item;
    const asunto = await getAsync((cb) => item.subject.getAsync(cb));
    const eco = economicoDeAsunto(asunto);
    if (!eco) return terminar(event, "⚠ Asunto no válido. Formato: 6943/3443/1045 Exportación...");

    const para = await getAsync((cb) => item.to.getAsync(cb));
    const nombres = para.map((r) => r.displayName).concat(para.map((r) => r.emailAddress));
    const ag = buscarAgencia(nombres);
    if (!ag) return terminar(event, "⚠ No reconozco la lista en PARA: " + (para[0] ? para[0].displayName : "(vacío)"));

    const texto = await getAsync((cb) => item.body.getAsync(Office.CoercionType.Text, cb));
    if (/CRUCE MA.ANA/i.test(texto)) return terminar(event, "⚠ El cuerpo ya tiene datos de CRUCE. No se duplicó.");

    await getAsync((cb) =>
      item.body.prependAsync(construirHtml(fechaCruce(), ag, eco), { coercionType: Office.CoercionType.Html }, cb)
    );
    terminar(event, "Listo: ECONOMICO " + eco + ", OSCAC " + ag.oscac + ", CAAT " + ag.caat + ".");
  } catch (e) {
    terminar(event, "⚠ Error: " + (e && e.message ? e.message : JSON.stringify(e)));
  }
}

// Registro a nivel global (no dentro de onReady) para que la función exista cuando Outlook la invoque.
if (typeof Office !== "undefined") {
  Office.onReady(() => {});
  Office.actions.associate("llenarCuerpo", llenarCuerpo);
}

if (typeof module !== "undefined") module.exports = { norm, fechaCruce, economicoDeAsunto, buscarAgencia, construirHtml };
