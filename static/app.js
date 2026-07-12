// Facturador Henkel — frontend (SALIDAS + DESTRUCCIÓN + INGRESOS).
// UI mínima: calendario, Generar (con barra de avance), Descargar Excel y problemas.
// Sin dependencias externas. Sondea /api/progress para mostrar avance y errores en vivo.

const $ = (id) => document.getElementById(id);
const POLL_MS = 700;

// ---- Fechas: input type=date usa yyyy-mm-dd; la API usa dd/mm/yyyy ----
function toDdMmYyyy(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
function toIso(ddmmyyyy) {
  if (!ddmmyyyy) return "";
  const [d, m, y] = ddmmyyyy.split("/");
  return `${y}-${m}-${d}`;
}
function fmtElapsed(sec) {
  if (sec == null || isNaN(sec)) return "";
  if (sec < 60) return `${sec.toFixed(1)} s`;
  const m = Math.floor(sec / 60);
  const r = Math.round(sec - m * 60);
  return `${m} min ${r} s`;
}

// ---- Estado/UI ----
let pollTimer = null;

function setStatus(state, text) {
  const pill = $("statusPill");
  pill.textContent = text;
  pill.className = "pill " + state;
}
function showBanner(type, html) {
  const b = $("banner");
  b.className = "banner banner-" + type;
  b.innerHTML = html;
}
function clearBanner() {
  $("banner").className = "banner hidden";
}
function setButtons({ generar = true, exportar = false } = {}) {
  $("btnGenerar").disabled = !generar;
  $("btnExportar").disabled = !exportar;
}

// ---- Carga inicial: precargar el rango disponible ----
async function init() {
  setStatus("pill-idle", "Listo");
  try {
    const res = await fetch("/api/daterange/default").then((r) => r.json());
    if (res.ok && res.data.start && res.data.end) {
      $("startDate").value = toIso(res.data.start);
      $("endDate").value = toIso(res.data.end);
    }
  } catch {
    /* el usuario escribirá las fechas manualmente */
  }
}

// ---- Generar: arranca el proceso y sondea el avance ----
async function generar() {
  const startIso = $("startDate").value;
  const endIso = $("endDate").value;
  if (!startIso || !endIso) {
    showBanner("error", "Selecciona ambas fechas (Desde y Hasta).");
    return;
  }
  const start = toDdMmYyyy(startIso);
  const end = toDdMmYyyy(endIso);

  clearBanner();
  setButtons({ generar: false, exportar: false });
  setStatus("pill-busy", "Procesando…");
  $("resultCard").classList.remove("hidden");
  $("issuesBox").classList.add("hidden");
  $("issuesBox").innerHTML = "";
  $("progressNote").innerHTML = "";
  $("progressNote").className = "progress-note";
  $("progressFill").className = "progress-fill";
  $("progressFill").style.width = "0%";
  $("progressPct").textContent = "0%";
  $("progressStage").textContent = "Iniciando…";

  let res;
  try {
    res = await fetch("/api/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ start, end }),
    }).then((r) => r.json());
  } catch (err) {
    setStatus("pill-error", "Error");
    showBanner("error", `No se pudo conectar con el servidor: ${err}`);
    setButtons({ generar: true });
    return;
  }

  if (!res.ok) {
    setStatus("pill-error", "Error");
    showBanner("error", `<strong>${res.error}</strong>${res.detail ? `<br>${res.detail}` : ""}`);
    setButtons({ generar: true });
    return;
  }

  startPolling();
}

function startPolling() {
  stopPolling();
  pollTimer = setInterval(pollOnce, POLL_MS);
  pollOnce(); // consulta inmediata
}

function stopPolling() {
  if (pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
}

async function pollOnce() {
  let state;
  try {
    const res = await fetch("/api/progress").then((r) => r.json());
    state = res.ok ? res.data : null;
  } catch {
    return; // reintenta en el próximo tick
  }
  if (!state) return;

  renderProgress(state);
  renderIssues(state.issues || []);

  if (state.done) {
    stopPolling();
    onFinish(state);
  }
}

function renderProgress(state) {
  const pct = Math.max(0, Math.min(100, state.percent || 0));
  $("progressPct").textContent = `${pct}%`;
  $("progressFill").style.width = `${pct}%`;
  $("progressStage").textContent = state.stage || "Procesando…";
  $("progressFill").className = "progress-fill" + (state.blocked ? " blocked" : state.done ? " done" : "");
  if (!state.done) {
    $("progressNote").className = "progress-note";
    $("progressNote").textContent = `Tiempo transcurrido: ${fmtElapsed(state.elapsed_seconds)}`;
  }
}

function onFinish(state) {
  renderIssues(state.issues || []);
  const t = fmtElapsed(state.elapsed_seconds);
  if (state.blocked) {
    setStatus("pill-error", "Detenido");
    $("progressNote").className = "progress-blocked";
    $("progressNote").innerHTML =
      `<strong>Proceso detenido${t ? ` tras ${t}` : ""} por un error de archivo.</strong>` +
      `${state.error ? `<br>${escapeHtml(state.error)}` : ""}` +
      `<br>Corrige el archivo indicado y vuelve a pulsar Generar.`;
    setButtons({ generar: true, exportar: false });
    return;
  }
  if (state.has_result) {
    setStatus("pill-idle", "Listo");
    $("progressNote").className = "progress-note progress-ok";
    $("progressNote").textContent = `Proceso terminado${t ? ` en ${t}` : ""}. Ya puedes descargar el Excel.`;
    setButtons({ generar: true, exportar: true });
    return;
  }
  // Terminó sin bloqueo pero sin resultado (p. ej. nada que facturar en el rango).
  const msg = (state.issues || []).find((i) => i.severity === "error");
  setStatus("pill-error", "Sin datos");
  $("progressNote").className = "progress-blocked";
  $("progressNote").innerHTML =
    `<strong>No se generaron servicios${t ? ` (${t})` : ""}.</strong>` +
    `${msg ? `<br>${escapeHtml(msg.msg)}` : ""}<br>Revisa el rango de fechas y las fuentes.`;
  setButtons({ generar: true, exportar: false });
}

// ---- Exportar (fetch+blob para mostrar errores) ----
async function exportar() {
  const start = toDdMmYyyy($("startDate").value);
  const end = toDdMmYyyy($("endDate").value);
  if (!start || !end) return;

  clearBanner();
  setStatus("pill-busy", "Generando Excel…");
  setButtons({ generar: false, exportar: false });

  try {
    const resp = await fetch(`/api/export?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`);
    if (!resp.ok) {
      const body = await resp.json().catch(() => ({}));
      setStatus("pill-idle", "Listo");
      showBanner("error", `<strong>${body.error || "No se pudo exportar."}</strong>${body.detail ? `<br>${body.detail}` : ""}`);
      setButtons({ generar: true, exportar: true });
      return;
    }
    const blob = await resp.blob();
    const filename = (resp.headers.get("content-disposition") || "").match(/filename="?([^"]+)"?/i);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename ? filename[1] : "facturacion.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setStatus("pill-idle", "Listo");
    setButtons({ generar: true, exportar: true });
  } catch (err) {
    setStatus("pill-error", "Error");
    showBanner("error", `No se pudo descargar: ${err}`);
    setButtons({ generar: true, exportar: true });
  }
}

// ---- Render de problemas (errores / advertencias) ----
function renderIssues(issues) {
  const box = $("issuesBox");
  if (!issues.length) {
    box.classList.add("hidden");
    box.innerHTML = "";
    return;
  }
  const errors = issues.filter((i) => i.severity === "error");
  const warnings = issues.filter((i) => i.severity !== "error");

  const parts = [];
  if (errors.length) {
    parts.push(
      `<div class="issue-group issue-error"><div class="issue-head">⛔ Errores (${errors.length})</div>` +
        errors.map((i) => `<div class="issue">${issueText(i)}</div>`).join("") +
        `</div>`
    );
  }
  if (warnings.length) {
    parts.push(
      `<div class="issue-group issue-warn"><div class="issue-head">⚠ Advertencias (${warnings.length})</div>` +
        warnings.map((i) => `<div class="issue">${issueText(i)}</div>`).join("") +
        `</div>`
    );
  }
  box.classList.remove("hidden");
  box.innerHTML = parts.join("");
}

function issueText(i) {
  const file = i.file ? `<span class="issue-file">${escapeHtml(i.file)}:</span> ` : "";
  return `${file}${escapeHtml(i.msg)}`;
}

function escapeHtml(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[c]);
}

// ---- Eventos ----
document.addEventListener("DOMContentLoaded", () => {
  $("btnGenerar").addEventListener("click", generar);
  $("btnExportar").addEventListener("click", exportar);
  init();
});
