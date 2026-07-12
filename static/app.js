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
  validateSources(); // checklist de fuentes al cargar
}

// ---- Validación previa de fuentes (checklist antes de facturar) ----
async function validateSources() {
  $("sourcesSummary").textContent = "Revisando archivos…";
  $("sourcesList").innerHTML = "";
  try {
    const res = await fetch("/api/validate").then((r) => r.json());
    if (res.ok) renderSources(res.data);
    else $("sourcesSummary").textContent = "No se pudieron revisar las fuentes.";
  } catch {
    $("sourcesSummary").textContent = "No se pudieron revisar las fuentes.";
  }
}

function renderSources(data) {
  const s = data.summary;
  const parts = [`<span class="src-ok">✅ ${s.present} presentes</span>`];
  if (s.missing_optional) parts.push(`<span class="src-warn">⚠️ ${s.missing_optional} opcionales faltan</span>`);
  if (s.missing_required) parts.push(`<span class="src-req">⛔ ${s.missing_required} requeridos faltan</span>`);
  const nota = s.missing_required
    ? `<span class="src-note">Faltan archivos requeridos: la facturación se detendrá.</span>`
    : s.missing_optional
      ? `<span class="src-note">Algunos pasos opcionales no tendrán datos (se omiten con aviso).</span>`
      : `<span class="src-note">Todo en orden. Listo para facturar.</span>`;
  $("sourcesSummary").innerHTML = parts.join(" ") + " " + nota;

  const groups = {};
  for (const it of data.items) (groups[it.group] = groups[it.group] || []).push(it);
  const order = ["Requeridos", "Pasos", "Auxiliares"];
  const icon = { present: "✅", missing_optional: "⚠️", missing_required: "⛔" };
  const cls = { present: "src-ok", missing_optional: "src-warn", missing_required: "src-req" };
  let html = "";
  for (const g of order) {
    if (!groups[g]) continue;
    html += `<div class="src-group"><div class="src-group-head">${g}</div>`;
    for (const it of groups[g]) {
      const cnt = it.count > 0
        ? `<span class="src-count">${it.count} archivo${it.count === 1 ? "" : "s"}</span>`
        : `<span class="src-count src-missing">no encontrado</span>`;
      html += `<div class="src-row ${cls[it.status]}"><span class="src-ico">${icon[it.status]}</span>`
        + `<span class="src-label">${escapeHtml(it.label)}</span>`
        + `<span class="src-where">${escapeHtml(it.where)}</span>${cnt}</div>`;
    }
    html += `</div>`;
  }
  $("sourcesList").innerHTML = html;
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
    launchConfetti();
    notifyDone(`Facturación lista${t ? ` (${t})` : ""}. Ya puedes descargar el Excel.`);
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

// ---- Confetti (vanilla JS, sin dependencias) ----
let confettiCanvas = null;
let confettiRAF = null;
function launchConfetti(durationMs = 3500) {
  if (!confettiCanvas) {
    confettiCanvas = document.createElement("canvas");
    confettiCanvas.id = "confetti-canvas";
    document.body.appendChild(confettiCanvas);
    window.addEventListener("resize", () => {
      confettiCanvas.width = window.innerWidth;
      confettiCanvas.height = window.innerHeight;
    });
  }
  const canvas = confettiCanvas;
  const ctx = canvas.getContext("2d");
  canvas.width = window.innerWidth;
  canvas.height = window.innerHeight;

  const colors = ["#0d7377", "#1e3a8a", "#16a34a", "#f59e0b", "#ef4444", "#8b5cf6", "#ec4899"];
  const N = 170;
  const pieces = [];
  for (let i = 0; i < N; i++) {
    pieces.push({
      x: Math.random() * canvas.width,
      y: -20 - Math.random() * canvas.height * 0.6,
      r: 6 + Math.random() * 7,
      color: colors[(Math.random() * colors.length) | 0],
      vx: -2 + Math.random() * 4,
      vy: 2 + Math.random() * 3.5,
      rot: Math.random() * Math.PI,
      vrot: -0.2 + Math.random() * 0.4,
      shape: Math.random() < 0.5 ? "rect" : "circle",
    });
  }

  if (confettiRAF) cancelAnimationFrame(confettiRAF);
  const start = performance.now();
  const tick = (now) => {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    let alive = 0;
    for (const p of pieces) {
      p.x += p.vx;
      p.y += p.vy;
      p.vy += 0.05; // gravedad suave
      p.rot += p.vrot;
      if (p.y < canvas.height + 40) alive++;
      ctx.save();
      ctx.translate(p.x, p.y);
      ctx.rotate(p.rot);
      ctx.fillStyle = p.color;
      if (p.shape === "rect") ctx.fillRect(-p.r / 2, -p.r / 2, p.r, p.r * 0.6);
      else { ctx.beginPath(); ctx.arc(0, 0, p.r / 2, 0, Math.PI * 2); ctx.fill(); }
      ctx.restore();
    }
    if (now - start < durationMs && alive > 0) {
      confettiRAF = requestAnimationFrame(tick);
    } else {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      confettiRAF = null;
    }
  };
  confettiRAF = requestAnimationFrame(tick);
}

// ---- Toast de éxito + notificación de escritorio (si ya está permitida) ----
let toastTimer = null;
function showToast(html, ms = 6000) {
  const t = $("toast");
  t.innerHTML = html;
  t.classList.add("show");
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove("show"), ms);
}
function notifyDone(msg) {
  showToast("✅ " + msg);
  if ("Notification" in window && Notification.permission === "granted") {
    try { new Notification("Facturador Henkel", { body: msg }); } catch { /* ignora */ }
  }
}

// ---- Eventos ----
document.addEventListener("DOMContentLoaded", () => {
  $("btnGenerar").addEventListener("click", generar);
  $("btnExportar").addEventListener("click", exportar);
  $("btnValidate").addEventListener("click", validateSources);
  init();
});
