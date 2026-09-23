/* Excel Merger — logica de la interfaz web local (v4.0.0, fase 2).
   Sin dependencias: vanilla JS, todo contra la API del servidor Java. */
"use strict";

const TOKEN_HEADER = "X-Em-Token";
const POLL_MS = 900;
const MAX_LINES = 2000;

/* El token llega en ?token=; se guarda en sessionStorage y se limpia de
   la barra de direcciones: recargar sigue funcionando (se relee del
   almacenamiento), pero la URL expuesta no conserva el secreto. */
function readToken() {
  const fromUrl = new URLSearchParams(window.location.search).get("token");
  if (fromUrl) {
    try { sessionStorage.setItem("em-token", fromUrl); } catch (e) { /* modo restringido */ }
    try { window.history.replaceState(null, "", window.location.pathname); } catch (e) { /* file:/ antiguo */ }
    return fromUrl;
  }
  try { return sessionStorage.getItem("em-token") || ""; } catch (e) { return ""; }
}

const token = readToken();
const logEl = document.getElementById("log");
const statusEl = document.getElementById("status");
const resultEl = document.getElementById("result");
const buildEl = document.getElementById("build");
const pathsEl = document.getElementById("paths");
const rowInputEl = document.getElementById("row-input");
const rowOutputEl = document.getElementById("row-output");
const pathInputEl = document.getElementById("path-input");
const pathOutputEl = document.getElementById("path-output");
const countEl = document.getElementById("log-count");
const jumpEl = document.getElementById("log-bottom");
const btnMerge = document.getElementById("btn-merge");
const btnCompare = document.getElementById("btn-compare");
const btnExit = document.getElementById("btn-exit");
const exitDialog = document.getElementById("exit-dialog");
const exitCancel = document.getElementById("exit-cancel");
const exitConfirm = document.getElementById("exit-confirm");
const offEl = document.getElementById("off");

/* Ajustes de salida (v4.1.0). */
const settingsEl = document.getElementById("settings");
const settingsNoteEl = document.getElementById("settings-note");
const modeEl = document.getElementById("setting-mode");
const summaryEl = document.getElementById("setting-summary");
const byRespEl = document.getElementById("setting-by-responsible");
const previewEl = document.getElementById("setting-preview");

const LABELS = {
  merge: "Fusión de Excel",
  compare: "Comprobador de discrepancias contra CSV",
};

/* Exit codes de App.java (0-4); el texto recupera las denominaciones
   que el propio programa loguea (CONFIGURACION, ENTRADA INVALIDA...). */
const EXIT_TEXT = {
  1: "error en la ejecución",
  2: "error de configuración",
  3: "entrada inválida",
  4: "fichero de salida inválido",
};

let cursor = 0;
let busy = false;
let stopping = false; /* true durante/desde el cierre: corta el sondeo */
let lastStatus = "";
let stick = true; /* seguir el final del registro automaticamente */
let lines = 0;

/* ---------- API ---------- */

async function api(path, method) {
  const response = await fetch(path, {
    method: method || "GET",
    headers: { [TOKEN_HEADER]: token },
  });
  let data = {};
  try {
    data = await response.json();
  } catch (e) {
    data = {};
  }
  if (!response.ok) {
    const error = new Error(data.error || "HTTP " + response.status);
    error.status = response.status;
    throw error;
  }
  return data;
}

/* ---------- Estado ---------- */

function setStatus(text, tone) {
  const key = text + "|" + (tone || "");
  if (key === lastStatus) {
    return;
  }
  lastStatus = key;
  statusEl.textContent = text;
  statusEl.className = "status" + (tone ? " " + tone : "");
}

function showResult(kind, exitCode, ms) {
  const label = LABELS[kind];
  const ok = exitCode === 0;
  resultEl.className = "result " + (ok ? "ok" : "err");
  resultEl.textContent = "";
  const text = ok
    ? label + " completada. "
    : label + " no completada — " + (EXIT_TEXT[exitCode] || "error") + ". ";
  resultEl.appendChild(document.createTextNode(text));
  resultEl.appendChild(Object.assign(document.createElement("span"), {
    className: "badge",
    textContent: "código " + exitCode + " · " + ms + " ms",
  }));
  resultEl.hidden = false;
}

function showError(message) {
  resultEl.className = "result err";
  resultEl.textContent = message;
  resultEl.hidden = false;
}

/* ---------- Registro ---------- */

/* Los eventos llegan por canal (OUT/ERR); el nivel real de logback viaja
   como prefijo en el propio mensaje (formato %-5level), y asi se colorea. */
function sniffLevel(message) {
  const match = /^\[?(TRACE|DEBUG|INFO|WARN|WARNING|ERROR|FATAL|FAIL)\b/.exec(message);
  return match ? match[1] : "";
}

function levelClass(event) {
  let level = event.level;
  if (level === "OUT" || level === "ERR") {
    level = sniffLevel(event.msg) || (level === "ERR" ? "ERROR" : "");
  }
  if (level === "ERROR" || level === "FATAL" || level === "ERR") {
    return "line err";
  }
  if (level === "WARN" || level === "WARNING") {
    return "line warn";
  }
  return "line";
}

function updateCount() {
  countEl.textContent = lines + (lines === 1 ? " línea" : " líneas");
}

function setEmptyMessage(text) {
  logEl.textContent = "";
  const p = document.createElement("p");
  p.className = "log-empty";
  p.textContent = text;
  logEl.appendChild(p);
  lines = 0;
  updateCount();
}

function appendEvent(event) {
  const empty = logEl.querySelector(".log-empty");
  if (empty) {
    empty.remove();
  }
  const div = document.createElement("div");
  div.className = levelClass(event);
  div.textContent = event.msg;
  logEl.appendChild(div);
  lines += 1;
  while (logEl.childElementCount > MAX_LINES) {
    logEl.removeChild(logEl.firstElementChild);
    lines -= 1;
  }
  updateCount();
  if (stick) {
    logEl.scrollTop = logEl.scrollHeight;
  }
}

/* Pegarse al final salvo que el usuario haya leido hacia arriba. */

logEl.addEventListener("scroll", function () {
  const nearBottom = logEl.scrollHeight - logEl.scrollTop - logEl.clientHeight < 48;
  stick = nearBottom;
  jumpEl.hidden = nearBottom;
});

jumpEl.addEventListener("click", function () {
  stick = true;
  jumpEl.hidden = true;
  logEl.scrollTop = logEl.scrollHeight;
  logEl.focus({ preventScroll: true });
});

async function pollLogs() {
  if (stopping) {
    return;
  }
  let data;
  try {
    data = await api("/api/logs?after=" + cursor);
  } catch (error) {
    if (stopping) {
      return;
    }
    if (error.status === 403) {
      setStatus("Sin acceso (token ausente o inválido)", "err");
    } else {
      setStatus("Sin conexión con el proceso", "err");
    }
    return;
  }
  if (stopping) {
    return;
  }
  if (typeof data.next === "number") {
    if (data.next < cursor) {
      /* El proceso se relanzo: buffer nuevo con numeracion mas baja. */
      setEmptyMessage("Registro reiniciado: el proceso se relanzó.");
      cursor = 0;
    } else {
      cursor = data.next;
    }
  }
  if (data.events && data.events.length > 0) {
    for (const event of data.events) {
      appendEvent(event);
    }
  }
  if (!busy && lastStatus.indexOf("Sin ") === 0) {
    setStatus("Listo");
  }
}

/* ---------- Acciones (opciones 1 y 2 del antiguo menu) ---------- */

async function runAction(kind) {
  if (busy || stopping) {
    return;
  }
  busy = true;
  const button = kind === "merge" ? btnMerge : btnCompare;
  const hadFocus = document.activeElement === button;
  button.classList.add("busy");
  button.setAttribute("aria-busy", "true");
  btnMerge.disabled = true;
  btnCompare.disabled = true;
  btnExit.disabled = true;
  resultEl.hidden = true;
  setStatus("Ejecutando " + LABELS[kind].toLowerCase() + "…");
  try {
    const data = await api("/api/" + kind, "POST");
    showResult(kind, data.exitCode, data.ms);
    setStatus(data.exitCode === 0 ? "Completado" : "Con errores",
      data.exitCode === 0 ? "ok" : "err");
  } catch (error) {
    if (error.status === 409) {
      setStatus("Ya hay una ejecución en curso", "warn");
    } else if (error.status === 403) {
      showError("Sin acceso: falta el token. Abre la URL completa que imprimió la terminal.");
      setStatus("Sin acceso", "err");
    } else {
      showError(error.message);
      setStatus("Error", "err");
    }
  } finally {
    busy = false;
    button.classList.remove("busy");
    button.removeAttribute("aria-busy");
    btnMerge.disabled = false;
    btnCompare.disabled = false;
    btnExit.disabled = false;
    if (hadFocus) {
      button.focus({ preventScroll: true });
    }
  }
}

/* ---------- Salir (opcion 3: exit code 0) ---------- */

async function shutdown() {
  stopping = true; /* corta el sondeo: sin parpadeos de error al morir */
  btnExit.disabled = true;
  btnMerge.disabled = true;
  btnCompare.disabled = true;
  setStatus("Cerrando…", "warn");
  try {
    await api("/api/shutdown", "POST");
  } catch (error) {
    /* El servidor puede caer justo al responder: da igual, ya sale. */
  }
  setStatus("Proceso finalizado", "ok");
  offEl.hidden = false;
  window.setTimeout(function () {
    try { window.close(); } catch (e) { /* el navegador puede negarlo: queda el cartel */ }
  }, 700);
}

/* ---------- Ajustes de salida (v4.1.0) ---------- */

/* Que hojas genera cada modo. El chip «Resumen» se recalcula con los
   interruptores; «Deuda *» solo aparece si se aporta el 3er fichero. */
const MODE_INFO = {
  cierre: {
    label: "Resumen clásico",
    base: ["Cierre", "Extracción", "Deuda *", "Resultado"],
    desc: "La salida de siempre: cierre, extracción, deuda (si se aporta) y resultado.",
  },
  responsables: {
    label: "Por responsable",
    base: ["Cierre", "Extracción", "Resultado", "+ 1 hoja por responsable"],
    desc: "Una hoja con los totales por cada responsable. Sin Deuda ni Resumen.",
  },
  completo: {
    label: "Completo",
    base: ["Cierre", "Extracción", "Deuda *", "Resultado", "+ 1 hoja por responsable"],
    desc: "Todo: el resumen clásico más una hoja por responsable.",
  },
};

function renderSettingsPreview() {
  const mode = modeEl.value;
  const info = MODE_INFO[mode] || MODE_INFO.completo;
  const chips = info.base.slice();
  const resumen = summaryEl.checked;
  const matriz = resumen && byRespEl.checked;
  const soportaResumen = mode !== "responsables";

  if (soportaResumen && resumen) {
    chips.push(matriz ? "Resumen + matriz por responsable" : "Resumen");
  }

  previewEl.textContent = "";
  const title = document.createElement("p");
  title.className = "preview-title";
  title.textContent = "Este resultado tendrá:";
  previewEl.appendChild(title);

  const row = document.createElement("p");
  row.className = "preview-chips";
  for (const chip of chips) {
    const span = document.createElement("span");
    span.className = "chip";
    span.textContent = chip;
    row.appendChild(span);
  }
  previewEl.appendChild(row);

  const line = document.createElement("p");
  line.className = "preview-desc";
  let desc = info.desc + " ";
  if (soportaResumen && !resumen) {
    desc = "Sin la hoja Resumen: " + info.desc + " ";
  }
  if (!soportaResumen && (resumen || matriz)) {
    desc += "(en este modo no se genera Resumen) ";
  }
  line.textContent = desc.trim();
  previewEl.appendChild(line);

  if (chips.indexOf("Deuda *") >= 0) {
    const foot = document.createElement("p");
    foot.className = "preview-note";
    foot.textContent = "* Solo si aportas el fichero de Deuda.";
    previewEl.appendChild(foot);
  }
}

function onSettingsChanged() {
  if (stopping) {
    return;
  }
  /* Misma regla del servidor, resuelta en origen para que no se
     notifique al usuario con un cambio no pedido: la matriz por
     responsable exige la hoja Resumen. */
  if (byRespEl.checked && !summaryEl.checked) {
    summaryEl.checked = true;
  } else if (!summaryEl.checked) {
    byRespEl.checked = false;
  }
  renderSettingsPreview();
  saveSettings();
}

let settingsSaving = false;
let settingsNoteTimer = null;

async function saveSettings() {
  if (settingsSaving || stopping) {
    return;
  }
  settingsSaving = true;
  const body = JSON.stringify({
    mode: modeEl.value,
    summaryEnabled: summaryEl.checked,
    byResponsibleEnabled: byRespEl.checked,
  });
  try {
    const response = await fetch("/api/settings", {
      method: "POST",
      headers: { [TOKEN_HEADER]: token, "Content-Type": "application/json" },
      body: body,
    });
    if (!response.ok) {
      let message = "HTTP " + response.status;
      try {
        const data = await response.json();
        if (data.error) {
          message = data.error;
        }
      } catch (e) {
        /* sin cuerpo JSON: se queda el mensaje HTTP */
      }
      throw new Error(message);
    }
    const data = await response.json();
    /* Valores autoritativos del servidor (el acoplamiento ya resuelto). */
    modeEl.value = data.mode;
    summaryEl.checked = !!data.summaryEnabled;
    byRespEl.checked = !!data.byResponsibleEnabled;
    renderSettingsPreview();
    flashNote("Guardado · se aplica a la próxima fusión");
  } catch (error) {
    flashNote("No se pudo guardar: " + error.message, true);
  } finally {
    settingsSaving = false;
  }
}

function flashNote(text, isError) {
  settingsNoteEl.textContent = text;
  settingsNoteEl.classList.toggle("note-err", !!isError);
  if (settingsNoteTimer) {
    window.clearTimeout(settingsNoteTimer);
  }
  settingsNoteTimer = window.setTimeout(function () {
    settingsNoteEl.textContent = "Se guarda automáticamente para la próxima vez.";
    settingsNoteEl.classList.remove("note-err");
    settingsNoteTimer = null;
  }, 3000);
}

async function loadSettings() {
  try {
    const data = await api("/api/settings");
    modeEl.value = data.mode || "completo";
    summaryEl.checked = !!data.summaryEnabled;
    byRespEl.checked = !!data.byResponsibleEnabled;
    settingsEl.hidden = false;
  } catch (error) {
    /* Proceso sin ajustes: se omite el panel y la fusión usa la
       configuración de siempre. */
    settingsEl.hidden = true;
  }
  renderSettingsPreview();
}

modeEl.addEventListener("change", onSettingsChanged);
summaryEl.addEventListener("change", onSettingsChanged);
byRespEl.addEventListener("change", onSettingsChanged);

/* ---------- Arranque ---------- */

async function loadInfo() {
  try {
    const info = await api("/api/info");
    buildEl.textContent = info.build || ("v" + info.version) || "Excel Merger";
    if (info.inputDir) {
      pathInputEl.textContent = info.inputDir;
      rowInputEl.hidden = false;
      pathsEl.hidden = false;
    }
    if (info.outputDir) {
      pathOutputEl.textContent = info.outputDir;
      rowOutputEl.hidden = false;
      pathsEl.hidden = false;
    }
  } catch (error) {
    buildEl.textContent = "Versión no disponible";
    if (error.status === 403) {
      showError("Sin acceso: falta el token. Abre la URL completa que imprimió la terminal (incluye ?token=…).");
      setStatus("Sin acceso", "err");
    }
  }
}

btnMerge.addEventListener("click", function () {
  runAction("merge");
});
btnCompare.addEventListener("click", function () {
  runAction("compare");
});

btnExit.addEventListener("click", function () {
  if (stopping) {
    return;
  }
  if (typeof exitDialog.showModal === "function") {
    exitDialog.showModal();
  } else if (window.confirm("¿Salir? El proceso terminará con código 0.")) {
    shutdown();
  }
});
exitCancel.addEventListener("click", function () {
  exitDialog.close();
});
exitConfirm.addEventListener("click", function () {
  exitDialog.close();
  shutdown();
});

setEmptyMessage("Esperando actividad…");
loadInfo();
loadSettings();
pollLogs();
window.setInterval(pollLogs, POLL_MS);
