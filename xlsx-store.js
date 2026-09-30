/*
 * xlsx-store.js — odczyt i zapis tabel w pliku .xlsx bezpośrednio na poziomie XML.
 * Zmieniane są tylko wiersze danych wskazanej tabeli (oraz jej zakres),
 * więc formatowanie, pozostałe arkusze, formuły i walidacje zostają nietknięte.
 * Wymaga: JSZip (lib/jszip.min.js).
 */
(function (global) {
  "use strict";

  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const DAY_MS = 86400000;
  const EPOCH = Date.UTC(1899, 11, 30);

  // ---------- pomocnicze ----------
  const byName = (node, name) => Array.from(node.getElementsByTagNameNS("*", name));
  const firstByName = (node, name) => byName(node, name)[0] || null;
  const childrenByName = (node, name) =>
    Array.from(node.childNodes).filter((n) => n.nodeType === 1 && n.localName === name);

  function colToIdx(col) {
    let n = 0;
    for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function idxToCol(i) {
    let s = "";
    i += 1;
    while (i > 0) {
      const m = (i - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  }
  function parseCellRef(ref) {
    const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref);
    if (!m) throw new Error("Niepoprawny adres komórki: " + ref);
    return { c: colToIdx(m[1]), r: parseInt(m[2], 10) };
  }
  function parseRange(ref) {
    const [a, b] = ref.split(":");
    const s = parseCellRef(a);
    const e = parseCellRef(b || a);
    return { c1: s.c, r1: s.r, c2: e.c, r2: e.r };
  }
  function rangeToRef(r) {
    return idxToCol(r.c1) + r.r1 + ":" + idxToCol(r.c2) + r.r2;
  }
  function resolvePath(baseFile, target) {
    if (target.startsWith("/")) return target.slice(1);
    const parts = baseFile.split("/");
    parts.pop();
    for (const seg of target.split("/")) {
      if (seg === "..") parts.pop();
      else if (seg !== ".") parts.push(seg);
    }
    return parts.join("/");
  }
  function relsPathFor(file) {
    const parts = file.split("/");
    const name = parts.pop();
    return (parts.length ? parts.join("/") + "/" : "") + "_rels/" + name + ".rels";
  }

  function isoToSerial(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    return (Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS;
  }
  function serialToIso(serial) {
    const d = new Date(EPOCH + Math.round(serial * DAY_MS));
    return d.toISOString().slice(0, 10);
  }

  // ---------- klasa skoroszytu ----------
  class XlsxBook {
    constructor(zip) {
      this.zip = zip;
      this.docs = new Map(); // ścieżka -> XMLDocument (zmodyfikowane są zapisywane)
      this.dirty = new Set();
    }

    static async open(data) {
      if (!global.JSZip) throw new Error("Brak biblioteki JSZip.");
      const zip = await global.JSZip.loadAsync(data);
      const book = new XlsxBook(zip);
      await book._init();
      return book;
    }

    async _xml(path) {
      if (this.docs.has(path)) return this.docs.get(path);
      const f = this.zip.file(path);
      if (!f) return null;
      const text = await f.async("string");
      const doc = new DOMParser().parseFromString(text, "application/xml");
      if (byName(doc, "parsererror").length) throw new Error("Uszkodzony XML w pliku: " + path);
      this.docs.set(path, doc);
      return doc;
    }

    async _rels(file) {
      const doc = await this._xml(relsPathFor(file));
      const map = {};
      if (!doc) return map;
      for (const r of byName(doc, "Relationship")) {
        map[r.getAttribute("Id")] = {
          target: resolvePath(file, r.getAttribute("Target")),
          type: r.getAttribute("Type") || "",
        };
      }
      return map;
    }

    async _init() {
      const ct = await this._xml("[Content_Types].xml");
      let wbPath = "xl/workbook.xml";
      if (ct) {
        for (const o of byName(ct, "Override")) {
          if (/sheet\.main\+xml$/.test(o.getAttribute("ContentType") || "")) {
            wbPath = o.getAttribute("PartName").replace(/^\//, "");
          }
        }
      }
      this.wbPath = wbPath;
      const wb = await this._xml(wbPath);
      if (!wb) throw new Error("To nie jest poprawny plik Excel (.xlsx).");
      const rels = await this._rels(wbPath);
      this.sheets = {};
      for (const s of byName(wb, "sheet")) {
        const rid =
          s.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") ||
          s.getAttribute("r:id");
        if (rels[rid]) this.sheets[s.getAttribute("name")] = rels[rid].target;
      }
      this.stylesPath = Object.values(rels).find((r) => /\/styles$/.test(r.type))?.target || null;
      const sstRel = Object.values(rels).find((r) => /\/sharedStrings$/.test(r.type));
      this.sharedStrings = [];
      if (sstRel) {
        const sst = await this._xml(sstRel.target);
        for (const si of childrenByName(sst.documentElement, "si")) {
          this.sharedStrings.push(this._richText(si));
        }
      }
      // Tabele: nazwa -> {sheetPath, tablePath}
      this.tables = {};
      for (const [sheetName, sheetPath] of Object.entries(this.sheets)) {
        const srels = await this._rels(sheetPath);
        for (const r of Object.values(srels)) {
          if (!/\/table$/.test(r.type)) continue;
          const tdoc = await this._xml(r.target);
          const t = tdoc.documentElement;
          const name = t.getAttribute("displayName") || t.getAttribute("name");
          this.tables[name.toLowerCase()] = { name, sheetName, sheetPath, tablePath: r.target };
        }
      }
    }

    _richText(node) {
      // tekst z <si>/<is>, pomijając fonetykę (<rPh>)
      let out = "";
      for (const t of byName(node, "t")) {
        let p = t.parentNode, skip = false;
        while (p && p !== node) {
          if (p.localName === "rPh") { skip = true; break; }
          p = p.parentNode;
        }
        if (!skip) out += t.textContent;
      }
      return out;
    }

    _cellValue(c) {
      const t = c.getAttribute("t") || "n";
      const v = firstByName(c, "v");
      if (t === "inlineStr") {
        const is = firstByName(c, "is");
        return is ? this._richText(is) : "";
      }
      if (!v) return null;
      const raw = v.textContent;
      switch (t) {
        case "s": return this.sharedStrings[parseInt(raw, 10)] ?? "";
        case "str": case "e": return raw;
        case "b": return raw === "1";
        default: {
          const n = parseFloat(raw);
          return isNaN(n) ? raw : n;
        }
      }
    }

    tableNames() {
      return Object.values(this.tables).map((t) => t.name);
    }

    /**
     * Czyta tabelę. Zwraca { headers:[...], rows:[{ values:{nagłówek:wartość}, styles:{nagłówek:s} }] }
     */
    async readTable(tableName) {
      const info = this.tables[tableName.toLowerCase()];
      if (!info) throw new Error(`Nie znaleziono tabeli "${tableName}" w pliku Excel.`);
      const tdoc = await this._xml(info.tablePath);
      const range = parseRange(tdoc.documentElement.getAttribute("ref"));
      const headers = byName(tdoc, "tableColumn").map((c) => c.getAttribute("name"));
      const sdoc = await this._xml(info.sheetPath);
      const rows = [];
      for (const row of byName(sdoc, "row")) {
        const r = parseInt(row.getAttribute("r"), 10);
        if (r <= range.r1 || r > range.r2) continue;
        const values = {}, styles = {};
        let any = false;
        for (const c of childrenByName(row, "c")) {
          const ref = parseCellRef(c.getAttribute("r"));
          const ci = ref.c - range.c1;
          if (ci < 0 || ci >= headers.length) continue;
          const h = headers[ci];
          const val = this._cellValue(c);
          if (c.getAttribute("s")) styles[h] = c.getAttribute("s");
          if (val !== null && val !== "") { values[h] = val; any = true; }
        }
        if (any) rows.push({ values, styles, rowNumber: r });
      }
      return { headers, rows };
    }

    async _dateStyleIndex() {
      if (this._dateStyle !== undefined) return this._dateStyle;
      const doc = this.stylesPath ? await this._xml(this.stylesPath) : null;
      if (!doc) return (this._dateStyle = null);
      const ns = doc.documentElement.namespaceURI;
      const numFmtsEl = firstByName(doc, "numFmts");
      const custom = {};
      if (numFmtsEl) for (const f of childrenByName(numFmtsEl, "numFmt")) custom[f.getAttribute("numFmtId")] = f.getAttribute("formatCode");
      const cellXfs = firstByName(doc, "cellXfs");
      const xfs = childrenByName(cellXfs, "xf");
      const isDateFmt = (id) => {
        if (["14", "15", "16", "17"].includes(id)) return true;
        const code = (custom[id] || "").toLowerCase();
        return /y/.test(code) && /d/.test(code) && !/h/.test(code);
      };
      for (let i = 0; i < xfs.length; i++) {
        if (isDateFmt(xfs[i].getAttribute("numFmtId") || "0")) return (this._dateStyle = String(i));
      }
      // brak – dodajemy format rrrr-mm-dd
      let fmtsEl = numFmtsEl;
      if (!fmtsEl) {
        fmtsEl = doc.createElementNS(ns, "numFmts");
        doc.documentElement.insertBefore(fmtsEl, doc.documentElement.firstChild);
      }
      const ids = Object.keys(custom).map(Number);
      const newId = String(Math.max(163, ...ids) + 1);
      const nf = doc.createElementNS(ns, "numFmt");
      nf.setAttribute("numFmtId", newId);
      nf.setAttribute("formatCode", "yyyy-mm-dd");
      fmtsEl.appendChild(nf);
      fmtsEl.setAttribute("count", String(childrenByName(fmtsEl, "numFmt").length));
      const xf = doc.createElementNS(ns, "xf");
      xf.setAttribute("numFmtId", newId);
      xf.setAttribute("fontId", "0"); xf.setAttribute("fillId", "0"); xf.setAttribute("borderId", "0");
      xf.setAttribute("xfId", "0"); xf.setAttribute("applyNumberFormat", "1");
      cellXfs.appendChild(xf);
      cellXfs.setAttribute("count", String(xfs.length + 1));
      this.dirty.add(this.stylesPath);
      return (this._dateStyle = String(xfs.length));
    }

    /**
     * Zastępuje wszystkie wiersze danych tabeli.
     * rows: tablica obiektów { values:{nagłówek:wartość}, styles?:{nagłówek:s} }
     * dateColumns: nazwy kolumn, w których wartość "RRRR-MM-DD" ma być zapisana jako data Excela.
     */
    async writeTable(tableName, rows, dateColumns = []) {
      const info = this.tables[tableName.toLowerCase()];
      if (!info) throw new Error(`Nie znaleziono tabeli "${tableName}" w pliku Excel.`);
      const tdoc = await this._xml(info.tablePath);
      const tEl = tdoc.documentElement;
      const range = parseRange(tEl.getAttribute("ref"));
      const headers = byName(tdoc, "tableColumn").map((c) => c.getAttribute("name"));
      const dateStyle = dateColumns.length ? await this._dateStyleIndex() : null;

      const sdoc = await this._xml(info.sheetPath);
      const ns = sdoc.documentElement.namespaceURI;
      const sheetData = firstByName(sdoc, "sheetData");

      // usuń stare wiersze danych tabeli (i wszystko poniżej w tym zakresie kolumn)
      for (const row of childrenByName(sheetData, "row")) {
        const r = parseInt(row.getAttribute("r"), 10);
        if (r <= range.r1) continue;
        const cells = childrenByName(row, "c");
        const outside = cells.filter((c) => {
          const ci = parseCellRef(c.getAttribute("r")).c;
          return ci < range.c1 || ci > range.c2;
        });
        if (r > range.r2 && outside.length === cells.length) continue; // wiersz spoza tabeli – zostaw
        if (outside.length === 0) sheetData.removeChild(row);
        else for (const c of cells) if (!outside.includes(c)) row.removeChild(c);
      }

      const existingRows = new Map(childrenByName(sheetData, "row").map((r) => [parseInt(r.getAttribute("r"), 10), r]));
      const newLast = range.r1 + Math.max(rows.length, 1);

      rows.forEach((rec, i) => {
        const rn = range.r1 + 1 + i;
        let rowEl = existingRows.get(rn);
        if (!rowEl) {
          rowEl = sdoc.createElementNS(ns, "row");
          rowEl.setAttribute("r", String(rn));
          // wstaw w odpowiednie miejsce (wiersze muszą być posortowane)
          const next = childrenByName(sheetData, "row").find((r) => parseInt(r.getAttribute("r"), 10) > rn);
          sheetData.insertBefore(rowEl, next || null);
        }
        headers.forEach((h, ci) => {
          let v = rec.values[h];
          if (v === undefined || v === null || v === "") return;
          const colIdx = range.c1 + ci;
          const c = sdoc.createElementNS(ns, "c");
          c.setAttribute("r", idxToCol(colIdx) + rn);
          let style = rec.styles && rec.styles[h];
          if (dateColumns.includes(h) && typeof v === "string" && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
            v = isoToSerial(v);
            style = dateStyle || style;
          }
          if (style) c.setAttribute("s", style);
          if (typeof v === "number") {
            const ve = sdoc.createElementNS(ns, "v");
            ve.textContent = String(v);
            c.appendChild(ve);
          } else if (typeof v === "boolean") {
            c.setAttribute("t", "b");
            const ve = sdoc.createElementNS(ns, "v");
            ve.textContent = v ? "1" : "0";
            c.appendChild(ve);
          } else {
            c.setAttribute("t", "inlineStr");
            const is = sdoc.createElementNS(ns, "is");
            const t = sdoc.createElementNS(ns, "t");
            t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
            t.textContent = String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
            is.appendChild(t);
            c.appendChild(is);
          }
          // komórki w wierszu posortowane wg kolumny
          const next = childrenByName(rowEl, "c").find((x) => parseCellRef(x.getAttribute("r")).c > colIdx);
          rowEl.insertBefore(c, next || null);
        });
        rowEl.removeAttribute("spans");
      });

      // zakres tabeli i autofiltra
      const newRange = { c1: range.c1, r1: range.r1, c2: range.c2, r2: newLast };
      tEl.setAttribute("ref", rangeToRef(newRange));
      const af = firstByName(tdoc, "autoFilter");
      if (af) af.setAttribute("ref", rangeToRef(newRange));
      // wymiar arkusza
      const dim = firstByName(sdoc, "dimension");
      if (dim) {
        const rowsNow = childrenByName(sheetData, "row");
        let maxR = 1, maxC = 0;
        for (const r of rowsNow) {
          maxR = Math.max(maxR, parseInt(r.getAttribute("r"), 10));
          for (const c of childrenByName(r, "c")) maxC = Math.max(maxC, parseCellRef(c.getAttribute("r")).c);
        }
        maxR = Math.max(maxR, newLast);
        maxC = Math.max(maxC, range.c2);
        dim.setAttribute("ref", "A1:" + idxToCol(maxC) + maxR);
      }
      this.dirty.add(info.tablePath);
      this.dirty.add(info.sheetPath);

      // wymuś przeliczenie formuł po otwarciu w Excelu
      const wb = await this._xml(this.wbPath);
      let calcPr = firstByName(wb, "calcPr");
      if (!calcPr) {
        calcPr = wb.createElementNS(wb.documentElement.namespaceURI, "calcPr");
        wb.documentElement.appendChild(calcPr);
      }
      calcPr.setAttribute("fullCalcOnLoad", "1");
      this.dirty.add(this.wbPath);
    }

    async toBlob() {
      const ser = new XMLSerializer();
      for (const path of this.dirty) {
        let xml = ser.serializeToString(this.docs.get(path));
        if (!xml.startsWith("<?xml")) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
        this.zip.file(path, xml);
      }
      return this.zip.generateAsync({ type: "blob", mimeType: XLSX_MIME, compression: "DEFLATE" });
    }
  }

  global.XlsxStore = { XlsxBook, isoToSerial, serialToIso, XLSX_MIME };
})(window);
