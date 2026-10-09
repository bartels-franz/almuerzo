(() => {
  "use strict";

  // ---- Datos tomados del Excel "PEDIDO-SÁBADO 10-10-26" ----
  const GUESTS = [
    [1, "Myriam"], [2, "Sebastían"], [3, "Andrea"], [4, "Sofía"], [5, "Valentina", true],
    [6, "Daniel"], [7, "Johana"], [8, "Nicolas"], [9, "Sofía"], [10, "Mariana"],
    [11, "Arturo"], [12, "Nicoll", true], [13, "Paula"], [14, "Richard"], [15, "Martina", true],
    [16, "Pilar"], [17, "Oscar"], [18, "Matías", true], [19, "Ana María"], [20, "Javier"],
    [21, "Cristina"], [22, "Hanna", true], [23, "Alejandra"], [24, "Andrés"], [25, "Thiago", true],
    [26, "Sofía"], [27, "Franz"], [28, "Juan"], [29, "Camila"], [30, "Julián"], [31, "Juan Pablo"],
  ].map(([id, nombre, nino]) => ({ id, nombre, nino: !!nino }));

  const LABELS = {
    ajiaco_normal: "Ajiaco plato normal",
    ajiaco_pequeno: "Ajiaco plato pequeño",
    frijol: "Frijol",
    verdura: "Verdura con carne molida",
    pasta: "Pasta con carne molida",
    lomo: "Lomo de cerdo a la plancha",
    bistec: "Carne en bisteck",
    pechuga: "Pechuga a la plancha",
    pollo: "Pollo al horno",
    mora: "Mora",
    mango: "Mango",
  };
  const TOTAL_GROUPS = [
    ["Sopita", ["ajiaco_normal", "ajiaco_pequeno"]],
    ["Principio", ["frijol", "verdura", "pasta"]],
    ["Proteína", ["lomo", "bistec", "pechuga", "pollo"]],
    ["Jugo", ["mora", "mango"]],
  ];

  const API = (window.PEDIDOS_API || "").trim();
  const DEMO_KEY = "almuerzo-10-10-26-demo";
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];

  let orders = {};      // id -> pedido
  let current = null;   // invitado seleccionado
  let sending = false;

  // ---------- Almacenamiento ----------
  function demoLoad() {
    try { return JSON.parse(localStorage.getItem(DEMO_KEY) || "{}"); } catch { return {}; }
  }
  function demoSave(all) {
    try { localStorage.setItem(DEMO_KEY, JSON.stringify(all)); } catch {}
  }

  async function fetchOrders() {
    if (!API) return demoLoad();
    const res = await fetch(API + (API.includes("?") ? "&" : "?") + "t=" + Date.now(), { method: "GET" });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || "Error al leer pedidos");
    const map = {};
    (data.orders || []).forEach(o => { map[o.id] = o; });
    return map;
  }

  async function saveOrder(order) {
    if (!API) {
      const all = demoLoad();
      all[order.id] = { ...order, fecha: new Date().toISOString() };
      demoSave(all);
      return all[order.id];
    }
    const res = await fetch(API, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: "save", order }),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || "No se pudo guardar");
    return data.order || order;
  }

  // ---------- Utilidades ----------
  const norm = s => s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const dupNames = (() => {
    const c = {}; GUESTS.forEach(g => (c[g.nombre] = (c[g.nombre] || 0) + 1));
    return new Set(Object.keys(c).filter(k => c[k] > 1));
  })();
  const display = g => dupNames.has(g.nombre) ? `${g.nombre} (#${g.id})` : g.nombre;

  let toastTimer;
  function toast(msg) {
    const t = $("#toast"); t.textContent = msg; t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), 3200);
  }

  function principioText(o) {
    const p = [];
    if (o.frijol) p.push("Frijol");
    if (o.acomp) p.push(LABELS[o.acomp]);
    return p.join(" + ");
  }

  // ---------- Paso 1: invitados ----------
  function renderGuests() {
    const q = norm($("#guestSearch").value.trim());
    const grid = $("#guestGrid");
    const list = GUESTS.filter(g => !q || norm(g.nombre).includes(q));
    if (!list.length) { grid.innerHTML = `<p class="empty">No encontramos ese nombre en la lista de invitados.</p>`; return; }
    grid.innerHTML = list.map(g => {
      const done = !!orders[g.id];
      const avatar = done
        ? `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>`
        : esc(g.nombre.charAt(0));
      return `<button type="button" role="listitem" class="guest${done ? " is-done" : ""}" data-id="${g.id}"
        aria-label="${esc(display(g))}${done ? ", pedido registrado" : ""}">
        <span class="guest__avatar">${avatar}</span>
        <span class="guest__name"><b>${esc(g.nombre)}${g.nino ? '<span class="kid">Niño/a</span>' : ""}</b>
        <small>Invitado #${g.id}${done ? " · ya pidió" : ""}</small></span>
      </button>`;
    }).join("");
  }

  function updateStats() {
    const n = Object.keys(orders).length;
    $("#statDone").textContent = n;
    $("#statTotal").textContent = GUESTS.length;
  }

  // ---------- Paso 2: menú ----------
  function selectGuest(id) {
    current = GUESTS.find(g => g.id === id);
    if (!current) return;
    $("#menuGuestName").textContent = display(current);
    const f = $("#orderForm"); f.reset();
    $$(".course").forEach(c => c.classList.remove("is-missing"));
    $("#formError").hidden = true;
    const o = orders[id];
    if (o) {
      if (o.sopa) $(`#o-${o.sopa}`).checked = true;
      $("#o-frijol").checked = !!o.frijol;
      if (o.acomp) $(`#o-${o.acomp}`).checked = true;
      if (o.proteina) $(`#o-${o.proteina}`).checked = true;
      if (o.jugo) $(`#o-${o.jugo}`).checked = true;
      $("#obs").value = o.obs || "";
    }
    showStep("menu");
    updateProgress();
    $("#stepMenu").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function readForm() {
    const f = $("#orderForm");
    const v = n => (f.querySelector(`input[name="${n}"]:checked`) || {}).value || "";
    return {
      id: current.id,
      nombre: current.nombre,
      nino: current.nino,
      sopa: v("sopa"),
      frijol: $("#o-frijol").checked,
      acomp: v("acomp"),
      proteina: v("proteina"),
      jugo: v("jugo"),
      obs: $("#obs").value.trim().slice(0, 300),
    };
  }

  function status(o) {
    return {
      sopa: !!o.sopa,
      principio: o.frijol || !!o.acomp,
      proteina: !!o.proteina,
      jugo: !!o.jugo,
    };
  }

  function updateProgress() {
    if (!current) return;
    const s = status(readForm());
    $$("#progress li").forEach(li => li.classList.toggle("is-done", s[li.dataset.k]));
    $$(".course[data-group]").forEach(c => {
      const ok = s[c.dataset.group];
      c.classList.toggle("is-complete", ok);
      if (ok) c.classList.remove("is-missing");
    });
    if (Object.values(s).every(Boolean)) $("#formError").hidden = true;
    $("#submitBtn").textContent = orders[current.id] ? "Actualizar pedido" : "Enviar pedido";
  }

  // Permite desmarcar un radio tocándolo de nuevo (p. ej. verdura/pasta es opcional junto al frijol)
  function enableRadioToggle() {
    $$('#orderForm input[type="radio"]').forEach(r => {
      let was = false;
      r.closest("label").addEventListener("pointerdown", () => { was = r.checked; });
      r.addEventListener("click", () => {
        if (was) { r.checked = false; updateProgress(); }
        was = false;
      });
    });
  }

  async function submit() {
    if (!current || sending) return;
    const o = readForm();
    const s = status(o);
    const missing = Object.keys(s).filter(k => !s[k]);
    if (missing.length) {
      const names = { sopa: "Sopita", principio: "Principio", proteina: "Proteína", jugo: "Jugo" };
      $$(".course[data-group]").forEach(c => c.classList.toggle("is-missing", missing.includes(c.dataset.group)));
      const err = $("#formError");
      err.textContent = "Falta elegir: " + missing.map(k => names[k]).join(", ") + ".";
      err.hidden = false;
      $(`.course[data-group="${missing[0]}"]`).scrollIntoView({ behavior: "smooth", block: "center" });
      return;
    }
    sending = true;
    const btn = $("#submitBtn"); const label = btn.textContent;
    btn.disabled = true; btn.textContent = "Enviando…";
    try {
      const saved = await saveOrder(o);
      orders[o.id] = { ...o, ...saved };
      updateStats(); renderGuests();
      showDone(orders[o.id]);
    } catch (e) {
      toast("No se pudo enviar. Revisa tu conexión e inténtalo de nuevo.");
      console.error(e);
    } finally {
      sending = false; btn.disabled = false; btn.textContent = label;
    }
  }

  function showDone(o) {
    $("#doneName").textContent = current.nombre;
    const rows = [
      ["Sopita", LABELS[o.sopa]],
      ["Principio", principioText(o)],
      ["Proteína", LABELS[o.proteina]],
      ["Jugo", LABELS[o.jugo]],
    ];
    if (o.obs) rows.push(["Observación", o.obs]);
    $("#doneTicket").innerHTML = rows.map(([k, v]) => `<li><span>${k}</span><span>${esc(v)}</span></li>`).join("");
    showStep("done");
    window.scrollTo({ top: $("#stepDone").offsetTop - 16, behavior: "smooth" });
  }

  function showStep(which) {
    $("#stepGuest").hidden = which !== "guest";
    $("#stepMenu").hidden = which !== "menu";
    $("#stepDone").hidden = which !== "done";
    $("#dock").hidden = which !== "menu";
  }

  // ---------- Resumen ----------
  function renderSummary() {
    const list = Object.values(orders);
    const count = k => list.filter(o => o.sopa === k || o.proteina === k || o.jugo === k || o.acomp === k || (k === "frijol" && o.frijol)).length;
    $("#totals").innerHTML = TOTAL_GROUPS.map(([title, keys]) => `
      <div class="tgroup"><h4>${title}</h4>
        ${keys.map(k => `<div class="trow"><span>${LABELS[k]}</span><b>${count(k)}</b></div>`).join("")}
      </div>`).join("");
    const kids = list.filter(o => GUESTS.find(g => g.id === Number(o.id))?.nino).length;
    $("#sumMeta").textContent = `${list.length} de ${GUESTS.length} invitados han registrado su pedido` + (kids ? ` (${kids} niño${kids > 1 ? "s" : ""}).` : ".");
    $("#ordersTable tbody").innerHTML = GUESTS.map(g => {
      const o = orders[g.id];
      if (!o) return `<tr class="pending"><td>${g.id}</td><td>${esc(g.nombre)}${g.nino ? ' <span class="kid">Niño/a</span>' : ""}</td><td colspan="5">Pendiente</td></tr>`;
      return `<tr><td>${g.id}</td><td><b>${esc(g.nombre)}</b>${g.nino ? ' <span class="kid">Niño/a</span>' : ""}</td>
        <td>${esc(LABELS[o.sopa] || "—")}</td><td>${esc(principioText(o) || "—")}</td>
        <td>${esc(LABELS[o.proteina] || "—")}</td><td>${esc(LABELS[o.jugo] || "—")}</td><td>${esc(o.obs || "")}</td></tr>`;
    }).join("");
  }

  async function refresh(showToast) {
    try {
      orders = await fetchOrders();
      updateStats(); renderGuests(); renderSummary();
      if (showToast) toast("Resumen actualizado");
    } catch (e) {
      console.error(e);
      $("#sumMeta").textContent = "No se pudieron cargar los pedidos. Toca Actualizar para reintentar.";
    }
  }

  function setView(v) {
    $$(".tab").forEach(t => t.classList.toggle("is-active", t.dataset.view === v));
    $("#viewPedido").hidden = v !== "pedido";
    $("#viewResumen").hidden = v !== "resumen";
    $("#dock").hidden = v !== "pedido" || $("#stepMenu").hidden;
    if (v === "resumen") refresh(false);
  }

  // ---------- Eventos ----------
  $("#guestSearch").addEventListener("input", renderGuests);
  $("#guestGrid").addEventListener("click", e => {
    const b = e.target.closest(".guest"); if (b) selectGuest(Number(b.dataset.id));
  });
  $("#orderForm").addEventListener("change", updateProgress);
  $("#orderForm").addEventListener("submit", e => { e.preventDefault(); submit(); });
  $("#submitBtn").addEventListener("click", submit);
  $("#changeGuest").addEventListener("click", () => { current = null; showStep("guest"); $("#stepGuest").scrollIntoView({ behavior: "smooth" }); });
  $("#editAgain").addEventListener("click", () => selectGuest(current.id));
  $("#otherGuest").addEventListener("click", () => { current = null; $("#guestSearch").value = ""; renderGuests(); showStep("guest"); $("#stepGuest").scrollIntoView({ behavior: "smooth" }); });
  $$(".tab").forEach(t => t.addEventListener("click", () => setView(t.dataset.view)));
  $("#refreshBtn").addEventListener("click", () => refresh(true));

  enableRadioToggle();
  $("#demoBanner").hidden = !!API;
  renderGuests(); updateStats(); renderSummary();
  refresh(false);
  if (location.hash === "#resumen") setView("resumen");
})();
