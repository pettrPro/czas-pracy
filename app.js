/*
 * app.js — Rejestr czasu pracy (OneDrive + Excel)
 */
(function () {
  "use strict";

  const CFG = Object.assign(
    { clientId: "", shareUrl: "", admins: [], appName: "Rejestr czasu pracy", maxHoursPerDayWarning: 12 },
    window.CZAS_PRACY_CONFIG || {}
  );
  const { XlsxBook, serialToIso, XLSX_MIME } = window.XlsxStore;
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["User.Read", "Files.ReadWrite.All"];
  const T_ENTRIES = "tblWpisy";
  const T_PROJECTS = "tblProjekty";

  // kolumny tabeli tblWpisy (klucz -> nazwa kolumny w Excelu)
  const COLS = {
    id: "ID",
    date: "Data",
    name: "Imię i nazwisko",
    project: "ID projektu",
    desc: "Opis",
    hours: "Godziny",
    created: "Utworzono",
    user: "Użytkownik",
  };
  const PCOLS = { id: "ID projektu", name: "Nazwa projektu", active: "Aktywny" };

  // ---------- narzędzia ----------
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const norm = (s) =>
    String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/ł/g, "l").replace(/\s+/g, " ").trim();
  const esc = (s) =>
    String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const pad = (n) => String(n).padStart(2, "0");
  const todayIso = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
  const nowText = () => { const d = new Date(); return `${todayIso()} ${pad(d.getHours())}:${pad(d.getMinutes())}`; };
  const fmtHours = (h) => (Math.round(h * 100) / 100).toLocaleString("pl-PL", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  const fmtDate = (iso) => {
    if (!iso) return "";
    const [y, m, d] = iso.split("-");
    const wd = new Date(Date.UTC(+y, +m - 1, +d)).toLocaleDateString("pl-PL", { weekday: "short", timeZone: "UTC" });
    return `${d}.${m}.${y} <span class="muted">${wd}</span>`;
  };
  const newId = () => "W" + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 6).toUpperCase();
  const store = {
    get(k) { try { return localStorage.getItem(k); } catch { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch { /* brak dostępu */ } },
    del(k) { try { localStorage.removeItem(k); } catch { /* brak dostępu */ } },
  };

  class ConflictError extends Error {}
  class LockedError extends Error {}

  function toIsoDate(v) {
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === "number") return serialToIso(v);
    const s = String(v).trim();
    let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
    if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
    m = /^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/.exec(s);
    if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
    return "";
  }
  function toNumber(v) {
    if (typeof v === "number") return v;
    const n = parseFloat(String(v ?? "").replace(",", ".").replace(/\s/g, ""));
    return isNaN(n) ? 0 : n;
  }

  // ---------- warstwa dostępu do pliku ----------
  function makeGraphStore(getToken) {
    let ids = null;
    const cacheKey = "cp_item_" + CFG.shareUrl;

    async function gfetch(path, opts = {}) {
      const token = await getToken();
      const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
        ...opts,
        headers: { Authorization: "Bearer " + token, ...(opts.headers || {}) },
      });
      return res;
    }
    async function errText(res) {
      try { const j = await res.json(); return j.error?.message || res.statusText; } catch { return res.statusText; }
    }
    function encodeShare(url) {
      const b64 = btoa(unescape(encodeURIComponent(url)));
      return "u!" + b64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
    }
    async function resolve() {
      if (ids) return ids;
      const cached = store.get(cacheKey);
      if (cached) { ids = JSON.parse(cached); return ids; }
      const res = await gfetch(`/shares/${encodeShare(CFG.shareUrl)}/driveItem?$select=id,name,parentReference,webUrl`, {
        headers: { Prefer: "redeemSharingLinkIfNecessary" },
      });
      if (!res.ok) {
        const msg = await errText(res);
        if (res.status === 403 || res.status === 404)
          throw new Error(`Brak dostępu do pliku Excel (${res.status}). Poproś właściciela o udostępnienie pliku Twojemu kontu z prawem edycji. Szczegóły: ${msg}`);
        throw new Error("Nie udało się otworzyć linku do pliku: " + msg);
      }
      const item = await res.json();
      ids = { driveId: item.parentReference.driveId, itemId: item.id };
      store.set(cacheKey, JSON.stringify(ids));
      return ids;
    }
    const itemPath = () => `/drives/${ids.driveId}/items/${ids.itemId}`;

    return {
      async load() {
        await resolve();
        let res = await gfetch(itemPath());
        if (res.status === 404 || res.status === 403) {
          store.del(cacheKey); ids = null; await resolve();
          res = await gfetch(itemPath());
        }
        if (!res.ok) throw new Error("Nie udało się pobrać informacji o pliku: " + (await errText(res)));
        const item = await res.json();
        let bytes;
        const dl = item["@microsoft.graph.downloadUrl"];
        const cres = dl ? await fetch(dl, { cache: "no-store" }) : await gfetch(itemPath() + "/content");
        if (!cres.ok) throw new Error("Nie udało się pobrać pliku Excel (" + cres.status + ").");
        bytes = await cres.arrayBuffer();
        return { bytes, eTag: item.eTag, webUrl: item.webUrl, name: item.name };
      },
      async save(blob, eTag) {
        const res = await gfetch(itemPath() + "/content", {
          method: "PUT",
          headers: { "Content-Type": XLSX_MIME, "If-Match": eTag },
          body: blob,
        });
        if (res.status === 412) throw new ConflictError("Plik został zmieniony przez inną osobę.");
        if (res.status === 423)
          throw new LockedError("Plik jest zablokowany — prawdopodobnie ktoś ma go otwartego do edycji w aplikacji Excel. Zamknij plik w Excelu i spróbuj ponownie.");
        if (!res.ok) throw new Error("Zapis do OneDrive nie powiódł się: " + (await errText(res)));
        const item = await res.json();
        return { eTag: item.eTag };
      },
    };
  }

  function makeDemoStore() {
    const KEY = "cp_demo_file";
    let bytes = null, version = 1;
    function b64ToBuf(b64) {
      const bin = atob(b64); const u = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
      return u.buffer;
    }
    function bufToB64(buf) {
      const u = new Uint8Array(buf); let s = "";
      for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
      return btoa(s);
    }
    return {
      demo: true,
      async load() {
        if (!bytes) {
          const saved = store.get(KEY);
          if (saved) bytes = b64ToBuf(saved);
          else {
            const res = await fetch("CzasPracy.xlsx", { cache: "no-store" });
            if (!res.ok) throw new Error("Nie znaleziono pliku CzasPracy.xlsx obok aplikacji.");
            bytes = await res.arrayBuffer();
          }
        }
        return { bytes: bytes.slice(0), eTag: "demo-" + version, webUrl: null, name: "CzasPracy.xlsx (demo)" };
      },
      async save(blob, eTag) {
        if (eTag !== "demo-" + version) throw new ConflictError("konflikt");
        bytes = await blob.arrayBuffer();
        version++;
        store.set(KEY, bufToB64(bytes));
        return { eTag: "demo-" + version };
      },
      reset() { store.del(KEY); bytes = null; version++; },
      currentBlob() { return new Blob([bytes], { type: XLSX_MIME }); },
    };
  }

  // ---------- stan ----------
  const S = {
    demo: false,
    fileStore: null,
    user: { name: "", email: "" },
    isAdmin: false,
    projects: [],
    entries: [],
    webUrl: null,
    eTag: null,
    lastLoad: 0,
    busy: false,
    editingId: null,
  };

  // ---------- parsowanie tabel ----------
  function headerMap(headers, wanted, tableName) {
    const byNorm = new Map(headers.map((h) => [norm(h), h]));
    const map = {}, missing = [];
    for (const [k, name] of Object.entries(wanted)) {
      const real = byNorm.get(norm(name));
      if (real) map[k] = real; else missing.push(name);
    }
    return { map, missing };
  }

  async function parseBook(book) {
    const pt = await book.readTable(T_PROJECTS);
    const pm = headerMap(pt.headers, PCOLS);
    if (!pm.map.id) throw new Error(`W tabeli ${T_PROJECTS} brakuje kolumny "${PCOLS.id}".`);
    const projects = pt.rows
      .map((r) => {
        const a = pm.map.active ? norm(r.values[pm.map.active]) : "t";
        return {
          id: String(r.values[pm.map.id] ?? "").trim(),
          name: pm.map.name ? String(r.values[pm.map.name] ?? "").trim() : "",
          active: !["n", "nie", "0", "false", "no"].includes(a),
        };
      })
      .filter((p) => p.id);

    const et = await book.readTable(T_ENTRIES);
    const em = headerMap(et.headers, COLS);
    const required = ["id", "date", "name", "project", "desc", "hours"];
    const miss = required.filter((k) => !em.map[k]).map((k) => COLS[k]);
    if (miss.length) throw new Error(`W tabeli ${T_ENTRIES} brakuje kolumn: ${miss.join(", ")}.`);
    return { projects, table: et, map: em.map };
  }

  function recordsToEntries(rows, map) {
    return rows.map((r, i) => ({
      id: String(r.values[map.id] ?? "").trim() || "#wiersz" + r.rowNumber,
      hasId: !!String(r.values[map.id] ?? "").trim(),
      date: toIsoDate(r.values[map.date]),
      name: String(r.values[map.name] ?? "").trim(),
      project: String(r.values[map.project] ?? "").trim(),
      desc: String(r.values[map.desc] ?? ""),
      hours: toNumber(r.values[map.hours]),
      created: map.created ? String(r.values[map.created] ?? "") : "",
      userEmail: map.user ? String(r.values[map.user] ?? "").trim() : "",
      _order: i,
    }));
  }

  // ---------- operacje na danych ----------
  async function loadAll() {
    const f = await S.fileStore.load();
    const book = await XlsxBook.open(f.bytes);
    const { projects, table, map } = await parseBook(book);
    S.projects = projects;
    S.entries = recordsToEntries(table.rows, map);
    S.eTag = f.eTag;
    S.webUrl = f.webUrl;
    S.lastLoad = Date.now();
  }

  /** Pobiera świeżą wersję pliku, stosuje zmianę i zapisuje (z ponowieniem przy konflikcie). */
  async function mutate(change) {
    for (let attempt = 1; attempt <= 4; attempt++) {
      const f = await S.fileStore.load();
      const book = await XlsxBook.open(f.bytes);
      const { projects, table, map } = await parseBook(book);
      const rows = table.rows.map((r) => ({ values: { ...r.values }, styles: r.styles, rowNumber: r.rowNumber }));
      for (const r of rows) if (!String(r.values[map.id] ?? "").trim()) r.values[map.id] = newId(); // nadaj ID wierszom dopisanym ręcznie w Excelu
      const legacyIds = new Map(table.rows.map((r, i) => ["#wiersz" + r.rowNumber, rows[i].values[map.id]]));
      const result = change(rows, map, legacyIds);
      await book.writeTable(T_ENTRIES, rows, [map.date]);
      const blob = await book.toBlob();
      try {
        const saved = await S.fileStore.save(blob, f.eTag);
        S.projects = projects;
        S.entries = recordsToEntries(rows, map);
        S.eTag = saved.eTag;
        S.webUrl = f.webUrl || S.webUrl;
        S.lastLoad = Date.now();
        return result;
      } catch (e) {
        if (e instanceof ConflictError && attempt < 4) { await new Promise((r) => setTimeout(r, 300 * attempt)); continue; }
        if (e instanceof ConflictError) throw new Error("Nie udało się zapisać — plik jest w tej chwili często zmieniany. Spróbuj ponownie.");
        throw e;
      }
    }
  }

  const canEdit = (e) => S.isAdmin || (e.userEmail && e.userEmail.toLowerCase() === S.user.email.toLowerCase());

  // ---------- UI: status, powiadomienia ----------
  function setStatus(kind, text) {
    const el = $("#sync");
    el.className = "sync " + kind;
    el.querySelector("span").textContent = text;
  }
  let toastTimer;
  function toast(msg, kind = "ok") {
    const el = $("#toast");
    el.textContent = msg;
    el.className = "toast show " + kind;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (el.className = "toast " + kind), kind === "err" ? 7000 : 3000);
  }
  function showView(id) {
    for (const v of ["#setup", "#login", "#app", "#loading"]) $(v).hidden = v !== id;
  }

  async function withBusy(label, fn) {
    if (S.busy) return;
    S.busy = true;
    document.body.classList.add("busy");
    setStatus("work", label);
    try {
      const r = await fn();
      setStatus("ok", "Zsynchronizowano " + new Date().toLocaleTimeString("pl-PL", { hour: "2-digit", minute: "2-digit" }));
      return r;
    } catch (e) {
      console.error(e);
      setStatus("err", "Błąd synchronizacji");
      toast(e.message || String(e), "err");
      throw e;
    } finally {
      S.busy = false;
      document.body.classList.remove("busy");
    }
  }

  // ---------- UI: formularz ----------
  function projectLabel(id) {
    const p = S.projects.find((x) => x.id === id);
    return p && p.name ? `${id} — ${p.name}` : id;
  }

  function fillProjectSelect(selected) {
    const sel = $("#f-project");
    const active = S.projects.filter((p) => p.active || p.id === selected);
    sel.innerHTML =
      '<option value="">— wybierz projekt —</option>' +
      active.map((p) => `<option value="${esc(p.id)}">${esc(p.id)}${p.name ? " — " + esc(p.name) : ""}${p.active ? "" : " (nieaktywny)"}</option>`).join("");
    sel.value = selected || "";
  }

  function fillNameList() {
    const names = [...new Set(S.entries.map((e) => e.name).filter(Boolean))].sort((a, b) => a.localeCompare(b, "pl"));
    $("#names").innerHTML = names.map((n) => `<option value="${esc(n)}">`).join("");
  }

  function resetForm(keepContext = true) {
    const project = keepContext ? $("#f-project").value : "";
    S.editingId = null;
    $("#entry-form").reset();
    $("#f-date").value = todayIso();
    $("#f-name").value = store.get("cp_name_" + S.user.email) || S.user.name;
    fillProjectSelect(project);
    $("#form-title").textContent = "Nowy wpis";
    $("#f-submit").textContent = "Zapisz wpis";
    $("#f-cancel").hidden = true;
    $("#form-card").classList.remove("editing");
  }

  function startEdit(id, copy = false) {
    const e = S.entries.find((x) => x.id === id);
    if (!e) return;
    S.editingId = copy ? null : id;
    $("#f-date").value = copy ? todayIso() : e.date;
    $("#f-name").value = copy ? $("#f-name").value || e.name : e.name;
    fillProjectSelect(e.project);
    $("#f-desc").value = e.desc;
    $("#f-hours").value = String(e.hours).replace(".", ",");
    $("#form-title").textContent = copy ? "Nowy wpis (kopia)" : "Edycja wpisu";
    $("#f-submit").textContent = copy ? "Zapisz wpis" : "Zapisz zmiany";
    $("#f-cancel").hidden = false;
    $("#form-card").classList.toggle("editing", !copy);
    $("#form-card").scrollIntoView({ behavior: "smooth", block: "start" });
    (copy ? $("#f-hours") : $("#f-desc")).focus();
  }

  async function onSubmit(ev) {
    ev.preventDefault();
    const date = $("#f-date").value;
    const name = $("#f-name").value.trim();
    const project = $("#f-project").value;
    const desc = $("#f-desc").value.trim();
    const hours = Math.round(toNumber($("#f-hours").value) * 100) / 100;

    const errs = [];
    if (!date) errs.push("datę");
    if (!name) errs.push("imię i nazwisko");
    if (!project) errs.push("projekt");
    if (!desc) errs.push("opis");
    if (errs.length) return toast("Uzupełnij: " + errs.join(", ") + ".", "err");
    if (!(hours > 0) || hours > 24) return toast("Liczba godzin musi być większa od 0 i nie większa niż 24.", "err");

    const dayTotal = S.entries
      .filter((e) => e.id !== S.editingId && e.date === date && norm(e.name) === norm(name))
      .reduce((a, e) => a + e.hours, 0) + hours;
    if (dayTotal > 24) return toast(`Suma godzin tej osoby w dniu ${date} przekroczyłaby 24 (${fmtHours(dayTotal)} h).`, "err");
    if (dayTotal > CFG.maxHoursPerDayWarning &&
        !confirm(`Suma godzin w dniu ${date} wyniesie ${fmtHours(dayTotal)} h. Zapisać mimo to?`)) return;

    store.set("cp_name_" + S.user.email, name);
    const editingId = S.editingId;
    try {
      await withBusy("Zapisywanie…", () =>
        mutate((rows, map, legacyIds) => {
          const values = {
            [map.date]: date,
            [map.name]: name,
            [map.project]: project,
            [map.desc]: desc,
            [map.hours]: hours,
          };
          if (editingId) {
            const targetId = legacyIds.get(editingId) || editingId;
            const row = rows.find((r) => r.values[map.id] === targetId);
            if (!row) throw new Error("Ten wpis został w międzyczasie usunięty lub zmieniony w Excelu. Odśwież listę.");
            Object.assign(row.values, values);
          } else {
            rows.push({
              values: {
                [map.id]: newId(),
                ...values,
                ...(map.created ? { [map.created]: nowText() } : {}),
                ...(map.user ? { [map.user]: S.user.email } : {}),
              },
            });
          }
        })
      );
      toast(editingId ? "Zmiany zapisane." : "Wpis zapisany.");
      resetForm(true);
      render();
    } catch { /* komunikat już wyświetlony */ }
  }

  async function onDelete(id) {
    const e = S.entries.find((x) => x.id === id);
    if (!e) return;
    if (!confirm(`Usunąć wpis z dnia ${e.date} (${projectLabel(e.project)}, ${fmtHours(e.hours)} h)?`)) return;
    try {
      await withBusy("Usuwanie…", () =>
        mutate((rows, map, legacyIds) => {
          const targetId = legacyIds.get(id) || id;
          const i = rows.findIndex((r) => r.values[map.id] === targetId);
          if (i < 0) throw new Error("Ten wpis został już usunięty.");
          rows.splice(i, 1);
        })
      );
      if (S.editingId === id) resetForm();
      toast("Wpis usunięty.");
      render();
    } catch { /* komunikat już wyświetlony */ }
  }

  // ---------- UI: filtry, lista, podsumowanie ----------
  function fillFilters() {
    const personSel = $("#flt-person");
    const cur = personSel.value;
    const names = [...new Set(S.entries.map((e) => e.name).filter(Boolean))].sort((a, b) => a.localeCompare(b, "pl"));
    personSel.innerHTML = '<option value="">Wszystkie osoby</option>' + names.map((n) => `<option value="${esc(n)}">${esc(n)}</option>`).join("");
    if (cur === "__init__") {
      const mine = store.get("cp_name_" + S.user.email) || S.user.name;
      personSel.value = names.includes(mine) ? mine : "";
    } else personSel.value = names.includes(cur) ? cur : "";

    const projSel = $("#flt-project");
    const curP = projSel.value;
    const used = new Set(S.entries.map((e) => e.project));
    const plist = S.projects.filter((p) => used.has(p.id) || p.active).map((p) => p.id);
    for (const id of used) if (id && !plist.includes(id)) plist.push(id);
    projSel.innerHTML = '<option value="">Wszystkie projekty</option>' + plist.map((id) => `<option value="${esc(id)}">${esc(projectLabel(id))}</option>`).join("");
    projSel.value = plist.includes(curP) ? curP : "";
  }

  function filtered() {
    const month = $("#flt-month").value; // RRRR-MM lub ""
    const person = $("#flt-person").value;
    const project = $("#flt-project").value;
    return S.entries
      .filter((e) => (!month || e.date.startsWith(month)) && (!person || e.name === person) && (!project || e.project === project))
      .sort((a, b) => (b.date || "").localeCompare(a.date || "") || (b.created || "").localeCompare(a.created || "") || b._order - a._order);
  }

  function renderList() {
    const list = filtered();
    const total = list.reduce((a, e) => a + e.hours, 0);
    $("#list-count").textContent = `${list.length} ${list.length === 1 ? "wpis" : list.length % 10 >= 2 && list.length % 10 <= 4 && (list.length % 100 < 10 || list.length % 100 >= 20) ? "wpisy" : "wpisów"}`;
    $("#list-total").textContent = fmtHours(total) + " h";
    const tbody = $("#entries tbody");
    if (!list.length) {
      tbody.innerHTML = `<tr class="empty"><td colspan="6">Brak wpisów dla wybranych filtrów.</td></tr>`;
      return;
    }
    let lastDate = null;
    tbody.innerHTML = list
      .map((e) => {
        const dayHead = e.date !== lastDate;
        lastDate = e.date;
        const editable = canEdit(e);
        const p = S.projects.find((x) => x.id === e.project);
        return `<tr class="${dayHead ? "day-first" : ""}${S.editingId === e.id ? " is-editing" : ""}">
          <td class="c-date" data-label="Data">${fmtDate(e.date)}</td>
          <td class="c-name" data-label="Osoba">${esc(e.name)}</td>
          <td class="c-proj" data-label="Projekt"><span class="tag">${esc(e.project)}</span>${p && p.name ? `<span class="pname">${esc(p.name)}</span>` : ""}</td>
          <td class="c-desc" data-label="Opis">${esc(e.desc)}</td>
          <td class="c-hours num" data-label="Godziny">${fmtHours(e.hours)}</td>
          <td class="c-act">
            <button class="icon" data-act="copy" data-id="${esc(e.id)}" title="Kopiuj jako nowy wpis" aria-label="Kopiuj">⧉</button>
            ${editable ? `<button class="icon" data-act="edit" data-id="${esc(e.id)}" title="Edytuj" aria-label="Edytuj">✎</button>
            <button class="icon danger" data-act="del" data-id="${esc(e.id)}" title="Usuń" aria-label="Usuń">🗑</button>` : ""}
          </td></tr>`;
      })
      .join("");
  }

  function renderSummary() {
    const list = filtered();
    const total = list.reduce((a, e) => a + e.hours, 0);
    const group = (keyFn) => {
      const m = new Map();
      for (const e of list) m.set(keyFn(e), (m.get(keyFn(e)) || 0) + e.hours);
      return [...m.entries()].sort((a, b) => b[1] - a[1]);
    };
    const bars = (rows, labelFn) => {
      if (!rows.length) return `<p class="muted">Brak danych.</p>`;
      const max = rows[0][1] || 1;
      return rows
        .map(([k, v]) => `<div class="bar-row"><div class="bar-label">${labelFn(k)}</div>
          <div class="bar-track"><div class="bar" style="width:${Math.max(2, (v / max) * 100)}%"></div></div>
          <div class="bar-val num">${fmtHours(v)} h <span class="muted">${total ? Math.round((v / total) * 100) : 0}%</span></div></div>`)
        .join("");
    };
    const month = $("#flt-month").value;
    const title = month
      ? new Date(month + "-01T12:00:00").toLocaleDateString("pl-PL", { month: "long", year: "numeric" })
      : "cały okres";
    const persons = new Set(list.map((e) => e.name)).size;
    const days = new Set(list.map((e) => e.name + "|" + e.date)).size;
    $("#summary").innerHTML = `
      <div class="kpis">
        <div class="kpi"><div class="kpi-v num">${fmtHours(total)}</div><div class="kpi-l">godzin · ${esc(title)}</div></div>
        <div class="kpi"><div class="kpi-v num">${persons}</div><div class="kpi-l">${persons === 1 ? "osoba" : "osób"}</div></div>
        <div class="kpi"><div class="kpi-v num">${days}</div><div class="kpi-l">osobodni</div></div>
        <div class="kpi"><div class="kpi-v num">${days ? fmtHours(total / days) : 0}</div><div class="kpi-l">śr. godzin / osobodzień</div></div>
      </div>
      <div class="sum-grid">
        <section><h3>Wg projektu</h3>${bars(group((e) => e.project), (k) => `<span class="tag">${esc(k)}</span> ${esc(S.projects.find((p) => p.id === k)?.name || "")}`)}</section>
        <section><h3>Wg osoby</h3>${bars(group((e) => e.name), (k) => esc(k))}</section>
      </div>`;
  }

  function render() {
    fillNameList();
    fillFilters();
    renderList();
    renderSummary();
    const link = $("#open-excel");
    link.hidden = !S.webUrl;
    if (S.webUrl) link.href = S.webUrl;
  }

  async function refresh(silent = false) {
    try {
      await withBusy("Wczytywanie…", loadAll);
      const keepProject = $("#f-project").value;
      fillProjectSelect(keepProject);
      render();
      if (!silent) toast("Dane odświeżone.");
    } catch { /* komunikat już wyświetlony */ }
  }

  // ---------- start aplikacji ----------
  function wireUi() {
    $("#entry-form").addEventListener("submit", onSubmit);
    $("#f-cancel").addEventListener("click", () => { resetForm(); renderList(); });
    $("#entries").addEventListener("click", (ev) => {
      const b = ev.target.closest("button[data-act]");
      if (!b || S.busy) return;
      const { act, id } = b.dataset;
      if (act === "edit") { startEdit(id); renderList(); }
      else if (act === "copy") startEdit(id, true);
      else if (act === "del") onDelete(id);
    });
    for (const id of ["#flt-month", "#flt-person", "#flt-project"]) $(id).addEventListener("change", () => { renderList(); renderSummary(); });
    $("#flt-clear").addEventListener("click", () => {
      $("#flt-month").value = ""; $("#flt-person").value = ""; $("#flt-project").value = "";
      renderList(); renderSummary();
    });
    $("#month-prev").addEventListener("click", () => shiftMonth(-1));
    $("#month-next").addEventListener("click", () => shiftMonth(1));
    $("#refresh").addEventListener("click", () => refresh());
    for (const t of $$(".tabs button")) t.addEventListener("click", () => {
      $$(".tabs button").forEach((x) => x.setAttribute("aria-selected", String(x === t)));
      $("#tab-list").hidden = t.dataset.tab !== "list";
      $("#tab-summary").hidden = t.dataset.tab !== "summary";
    });
    window.addEventListener("focus", () => {
      if (!S.busy && S.fileStore && Date.now() - S.lastLoad > 60000) refresh(true);
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && S.editingId) { resetForm(); renderList(); }
    });
  }

  function shiftMonth(delta) {
    const el = $("#flt-month");
    const base = el.value || todayIso().slice(0, 7);
    const [y, m] = base.split("-").map(Number);
    const d = new Date(y, m - 1 + delta, 1);
    el.value = `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
    renderList(); renderSummary();
  }

  async function enterApp() {
    $("#app-title").textContent = CFG.appName;
    document.title = CFG.appName;
    $("#user-name").textContent = S.user.name || S.user.email;
    $("#user-email").textContent = S.user.email;
    $("#admin-badge").hidden = !S.isAdmin;
    $("#demo-banner").hidden = !S.demo;
    $("#flt-month").value = todayIso().slice(0, 7);
    $("#flt-person").innerHTML = '<option value="__init__"></option>';
    $("#flt-person").value = "__init__";
    showView("#app");
    resetForm(false);
    try {
      await withBusy("Wczytywanie…", loadAll);
      resetForm(false);
      render();
    } catch { /* komunikat już wyświetlony */ }
  }

  function startDemo() {
    S.demo = true;
    const ds = makeDemoStore();
    S.fileStore = ds;
    S.user = { name: "Jan Kowalski", email: "jan.kowalski@przyklad.pl" };
    S.isAdmin = true;
    $("#logout").textContent = "Zakończ demo";
    $("#logout").onclick = () => { location.href = location.pathname; };
    $("#demo-download").hidden = false;
    $("#demo-download").onclick = () => {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(ds.currentBlob());
      a.download = "CzasPracy-demo.xlsx";
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
    };
    $("#demo-reset").hidden = false;
    $("#demo-reset").onclick = () => {
      if (!confirm("Przywrócić dane demonstracyjne do stanu początkowego?")) return;
      ds.reset(); refresh();
    };
    enterApp();
  }

  async function startMsal() {
    if (!window.msal) throw new Error("Nie wczytano biblioteki logowania (lib/msal-browser.min.js).");
    const redirectUri = location.origin + location.pathname;
    const pca = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/common", redirectUri, postLogoutRedirectUri: redirectUri },
      cache: { cacheLocation: "localStorage" },
    });
    await pca.initialize();
    const res = await pca.handleRedirectPromise();
    let account = res?.account || pca.getActiveAccount() || pca.getAllAccounts()[0];

    $("#login-btn").onclick = () => pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
    if (!account) { showView("#login"); return; }
    pca.setActiveAccount(account);

    const getToken = async () => {
      try {
        const r = await pca.acquireTokenSilent({ scopes: SCOPES, account });
        return r.accessToken;
      } catch (e) {
        if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|consent_required|login_required/.test(e.errorCode || "")) {
          await pca.acquireTokenRedirect({ scopes: SCOPES, account });
          return new Promise(() => {}); // strona przekieruje się do logowania
        }
        throw e;
      }
    };
    S.fileStore = makeGraphStore(getToken);
    const email = (account.username || account.idTokenClaims?.email || account.idTokenClaims?.preferred_username || "").toLowerCase();
    S.user = { name: account.name || email, email };
    S.isAdmin = (CFG.admins || []).map((a) => a.toLowerCase()).includes(email);
    $("#logout").onclick = () => pca.logoutRedirect({ account });
    await enterApp();
  }

  async function main() {
    wireUi();
    $("#redirect-uri").textContent = location.origin + location.pathname;
    $("#demo-btn").addEventListener("click", startDemo);
    const params = new URLSearchParams(location.search);
    if (params.has("demo")) return startDemo();
    if (!CFG.clientId || !CFG.shareUrl) {
      $("#setup-missing").textContent = [!CFG.clientId && "clientId", !CFG.shareUrl && "shareUrl"].filter(Boolean).join(" i ");
      showView("#setup");
      return;
    }
    try {
      await startMsal();
    } catch (e) {
      console.error(e);
      showView("#login");
      $("#login-error").hidden = false;
      $("#login-error").textContent = "Błąd logowania: " + (e.message || e);
    }
  }

  document.addEventListener("DOMContentLoaded", main);
})();
