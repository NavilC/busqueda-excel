(() => {
  "use strict";

  let workbook = null;
  let sheets = {};
  let summaryResults = [];
  let detailResults = [];
  let allAssignments = [];

  const $ = (id) => document.getElementById(id);
  const fileInput = $("fileInput");
  const dropZone = $("dropZone");
  const pickFileBtn = $("pickFileBtn");
  const changeFileBtn = $("changeFileBtn");
  const fileCard = $("fileCard");
  const searchCard = $("searchCard");
  const resultsCard = $("resultsCard");
  const previewCard = $("previewCard");
  const fileName = $("fileName");
  const fileStats = $("fileStats");
  const sheetChips = $("sheetChips");
  const searchInput = $("searchInput");
  const matchMode = $("matchMode");
  const sheetFilter = $("sheetFilter");
  const projectFilter = $("projectFilter");
  const contractorFilter = $("contractorFilter");
  const roleFilter = $("roleFilter");
  const searchBtn = $("searchBtn");
  const summaryBtn = $("summaryBtn");
  const clearBtn = $("clearBtn");
  const exportBtn = $("exportBtn");
  const resultsSummary = $("resultsSummary");
  const fieldChips = $("fieldChips");
  const emptyResults = $("emptyResults");
  const resultsWrap = $("resultsWrap");
  const resultsTable = $("resultsTable");
  const previewSheet = $("previewSheet");
  const previewTable = $("previewTable");

  const normalize = (value) => String(value ?? "")
    .trim().toUpperCase().normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");

  const safeText = (value) => String(value ?? "");

  function excelColumn(index) {
    let n = index + 1, out = "";
    while (n > 0) {
      const r = (n - 1) % 26;
      out = String.fromCharCode(65 + r) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }

  function columnIndexFromRef(ref) {
    const m = String(ref || "").match(/^([A-Z]+)/i);
    if (!m) return 0;
    let n = 0;
    for (const ch of m[1].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  pickFileBtn.addEventListener("click", (e) => { e.stopPropagation(); fileInput.click(); });
  changeFileBtn.addEventListener("click", () => fileInput.click());
  dropZone.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => fileInput.files?.[0] && readExcel(fileInput.files[0]));

  ["dragenter", "dragover"].forEach(evt => dropZone.addEventListener(evt, e => {
    e.preventDefault(); dropZone.classList.add("dragover");
  }));
  ["dragleave", "drop"].forEach(evt => dropZone.addEventListener(evt, e => {
    e.preventDefault(); dropZone.classList.remove("dragover");
  }));
  dropZone.addEventListener("drop", e => {
    const f = e.dataTransfer.files?.[0];
    if (f) readExcel(f);
  });

  searchBtn.addEventListener("click", executeSearch);
  summaryBtn.addEventListener("click", executeSummary);
  clearBtn.addEventListener("click", clearSearch);
  exportBtn.addEventListener("click", exportCSV);
  previewSheet.addEventListener("change", () => renderPreview(previewSheet.value));

  [sheetFilter, projectFilter, contractorFilter, roleFilter].forEach(el => {
    el.addEventListener("change", () => {
      if (allAssignments.length && !resultsCard.classList.contains("hidden")) executeSummary();
    });
  });

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("Tu navegador no soporta descompresión local. Usa Chrome o Edge actualizado.");
    }
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzipXlsx(arrayBuffer) {
    const data = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);
    let eocd = -1;
    const min = Math.max(0, data.length - 65557);
    for (let i = data.length - 22; i >= min; i--) {
      if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("No se encontró la estructura ZIP del archivo XLSX.");

    const total = view.getUint16(eocd + 10, true);
    let pos = view.getUint32(eocd + 16, true);
    const decoder = new TextDecoder("utf-8");
    const entries = new Map();

    for (let i = 0; i < total; i++) {
      if (view.getUint32(pos, true) !== 0x02014b50) throw new Error("Directorio ZIP inválido.");
      const method = view.getUint16(pos + 10, true);
      const compSize = view.getUint32(pos + 20, true);
      const nameLen = view.getUint16(pos + 28, true);
      const extraLen = view.getUint16(pos + 30, true);
      const commentLen = view.getUint16(pos + 32, true);
      const localOffset = view.getUint32(pos + 42, true);
      const name = decoder.decode(data.slice(pos + 46, pos + 46 + nameLen));

      if (view.getUint32(localOffset, true) !== 0x04034b50) throw new Error("Entrada ZIP inválida.");
      const localNameLen = view.getUint16(localOffset + 26, true);
      const localExtraLen = view.getUint16(localOffset + 28, true);
      const start = localOffset + 30 + localNameLen + localExtraLen;
      const compressed = data.slice(start, start + compSize);

      let content;
      if (method === 0) content = compressed;
      else if (method === 8) content = await inflateRaw(compressed);
      else throw new Error(`Método ZIP no soportado (${method}).`);

      entries.set(name.replace(/\\/g, "/"), content);
      pos += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  function xmlFromEntry(entries, name) {
    const bytes = entries.get(name);
    if (!bytes) return null;
    const xmlText = new TextDecoder("utf-8").decode(bytes);
    const xml = new DOMParser().parseFromString(xmlText, "application/xml");
    if (xml.querySelector("parsererror")) throw new Error(`XML inválido: ${name}`);
    return xml;
  }

  function directChildByLocalName(parent, localName) {
    return [...parent.children].find(el => el.localName === localName) || null;
  }

  function allTextNodes(parent, localName) {
    return [...parent.getElementsByTagNameNS("*", localName)].map(x => x.textContent || "").join("");
  }

  function parseSharedStrings(entries) {
    const xml = xmlFromEntry(entries, "xl/sharedStrings.xml");
    if (!xml) return [];
    return [...xml.getElementsByTagNameNS("*", "si")].map(si => allTextNodes(si, "t"));
  }

  function parseWorkbook(entries) {
    const wbXml = xmlFromEntry(entries, "xl/workbook.xml");
    const relXml = xmlFromEntry(entries, "xl/_rels/workbook.xml.rels");
    if (!wbXml || !relXml) throw new Error("El archivo no contiene un libro Excel válido.");

    const rels = new Map();
    [...relXml.getElementsByTagNameNS("*", "Relationship")].forEach(rel => {
      rels.set(rel.getAttribute("Id"), rel.getAttribute("Target"));
    });

    const list = [];
    [...wbXml.getElementsByTagNameNS("*", "sheet")].forEach(sheet => {
      const name = sheet.getAttribute("name") || "Hoja";
      let rid = sheet.getAttribute("r:id");
      if (!rid) for (const attr of [...sheet.attributes]) if (attr.localName === "id") rid = attr.value;
      let target = rels.get(rid) || "";
      target = target.replace(/^\//, "");
      if (!target.startsWith("xl/")) target = "xl/" + target.replace(/^\.\//, "");
      target = target.replace(/\/\.\//g, "/");
      list.push({ name, path: target });
    });
    return list;
  }

  function parseWorksheet(xml, sharedStrings) {
    const rows = [];
    const rowNodes = [...xml.getElementsByTagNameNS("*", "row")];

    rowNodes.forEach((rowNode, rowOrder) => {
      const excelRow = Math.max(1, parseInt(rowNode.getAttribute("r") || (rowOrder + 1), 10));
      const arr = [];
      const cells = [...rowNode.getElementsByTagNameNS("*", "c")];

      cells.forEach(cell => {
        const ref = cell.getAttribute("r") || "A1";
        const col = columnIndexFromRef(ref);
        const type = cell.getAttribute("t") || "n";
        let value = "";

        if (type === "inlineStr") {
          const is = directChildByLocalName(cell, "is");
          value = is ? allTextNodes(is, "t") : "";
        } else {
          const v = directChildByLocalName(cell, "v");
          const raw = v ? v.textContent || "" : "";
          if (type === "s") value = sharedStrings[parseInt(raw || "0", 10)] ?? raw;
          else if (type === "b") value = raw === "1" ? "VERDADERO" : "FALSO";
          else value = raw;
        }
        arr[col] = value;
      });
      rows[excelRow - 1] = arr;
    });
    return rows;
  }

  async function readExcel(file) {
    const lower = file.name.toLowerCase();
    if (!lower.endsWith(".xlsx") && !lower.endsWith(".xlsm")) {
      alert("Admite .xlsx y .xlsm. Si tienes .xls, guárdalo primero como .xlsx desde Excel.");
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const entries = await unzipXlsx(buffer);
      const sharedStrings = parseSharedStrings(entries);
      const bookSheets = parseWorkbook(entries);
      sheets = {};

      for (const sh of bookSheets) {
        const xml = xmlFromEntry(entries, sh.path);
        if (xml) sheets[sh.name] = parseWorksheet(xml, sharedStrings);
      }

      workbook = { SheetNames: Object.keys(sheets) };
      if (!workbook.SheetNames.length) throw new Error("No se encontraron hojas legibles.");

      allAssignments = extractAllAssignments();
      renderFileInfo(file);
      buildSheetSelectors();
      buildDataFilters();
      renderPreview(workbook.SheetNames[0]);
      fileCard.classList.remove("hidden");
      searchCard.classList.remove("hidden");
      previewCard.classList.remove("hidden");
      resultsCard.classList.add("hidden");
      searchInput.focus();
    } catch (err) {
      console.error(err);
      alert("No se pudo leer el Excel. " + (err.message || "Verifica que el archivo no esté dañado o protegido."));
    }
  }

  function nonEmptyRowCount(rows) {
    return rows.filter(row => row && row.some(v => safeText(v).trim() !== "")).length;
  }

  function renderFileInfo(file) {
    fileName.textContent = file.name;
    const totalRows = workbook.SheetNames.reduce((acc, name) => acc + nonEmptyRowCount(sheets[name] || []), 0);
    fileStats.textContent = `${workbook.SheetNames.length} hoja(s) · ${totalRows.toLocaleString()} fila(s) con información · ${allAssignments.length.toLocaleString()} asignación(es) detectada(s)`;
    sheetChips.innerHTML = "";
    workbook.SheetNames.forEach(name => {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = `${name} · ${nonEmptyRowCount(sheets[name] || [])} filas`;
      sheetChips.appendChild(chip);
    });
  }

  function buildSheetSelectors() {
    sheetFilter.innerHTML = '<option value="__ALL__">Todas las hojas</option>';
    previewSheet.innerHTML = "";
    workbook.SheetNames.forEach(name => {
      for (const select of [sheetFilter, previewSheet]) {
        const op = document.createElement("option");
        op.value = name; op.textContent = name; select.appendChild(op);
      }
    });
  }

  function fillSelect(select, values, allLabel) {
    const current = select.value;
    select.innerHTML = "";
    const all = document.createElement("option");
    all.value = "__ALL__";
    all.textContent = allLabel;
    select.appendChild(all);
    [...new Set(values.filter(Boolean))].sort((a,b) => a.localeCompare(b, "es")).forEach(value => {
      const op = document.createElement("option");
      op.value = value;
      op.textContent = value;
      select.appendChild(op);
    });
    if ([...select.options].some(o => o.value === current)) select.value = current;
  }

  function buildDataFilters() {
    fillSelect(projectFilter, allAssignments.map(x => x.project), "Todos los proyectos");
    fillSelect(contractorFilter, allAssignments.map(x => x.contractor), "Todos los contratistas");
  }

  function getTerms() {
    return [...new Set(searchInput.value.split(/[\n,;]+/).map(normalize).filter(Boolean))];
  }

  function rowText(row) {
    return (row || []).map(safeText).filter(Boolean).join(" | ");
  }

  function getContractor(row) {
    for (const cell of (row || [])) {
      const text = safeText(cell).trim();
      const m = text.match(/CONTRATISTA\s*[:\-]\s*(.+)$/i);
      if (m) return m[1].trim();
    }
    return "";
  }

  function formatProjectValue(label, value) {
    const clean = safeText(value).trim();
    if (!clean) return "";

    const labelN = normalize(label);
    const valueN = normalize(clean);

    // En las planillas del usuario el campo suele llamarse "Proyecto OP".
    // Si el valor es solo el número/código (ej. 0524), lo mostramos como "OP 0524".
    if (labelN.includes("PROYECTO OP")) {
      if (valueN.startsWith("OP ")) return clean;
      return `OP ${clean}`;
    }

    return clean;
  }

  function getProjectFromRow(row) {
    for (let i = 0; i < (row || []).length; i++) {
      const text = safeText(row[i]).trim();
      const n = normalize(text);

      // Casos en una sola celda:
      // "Proyecto OP: 0524" -> "OP 0524"
      // "Proyecto: Hospital Santa Bárbara" -> "Hospital Santa Bárbara"
      let m = text.match(/(PROYECTO\s*OP|PROYECTO|OBRA)\s*[:\-]\s*(.+)$/i);
      if (m) return formatProjectValue(m[1], m[2]);

      // Casos donde el encabezado está en una celda y el valor en la siguiente:
      // [Proyecto OP] [0524]
      if (n === "PROYECTO OP" || n === "PROYECTO" || n === "OBRA") {
        const next = safeText(row[i + 1]).trim();
        if (next) return formatProjectValue(text, next);
      }
    }
    return "";
  }

  function inferProject(rows, blockStart) {
    for (let r = blockStart; r >= Math.max(0, blockStart - 40); r--) {
      const p = getProjectFromRow(rows[r]);
      if (p) return p;
    }
    for (let r = blockStart + 1; r <= Math.min(rows.length - 1, blockStart + 15); r++) {
      const p = getProjectFromRow(rows[r]);
      if (p) return p;
    }
    return "SIN PROYECTO DETECTADO";
  }

  function getMarkColumn(row) {
    let fallback = -1;
    for (let i = 0; i < (row || []).length; i++) {
      const n = normalize(row[i]);
      if (!n) continue;
      if (n.includes("MARCA") && n.includes("ITEM")) return i;
      if (n === "MARCA" || n.startsWith("MARCA/")) return i;
      if (n === "ITEM" || n.includes("MARCA")) fallback = i;
    }
    return fallback;
  }

  function descriptionColumn(row) {
    for (let i = 0; i < (row || []).length; i++) {
      const n = normalize(row[i]);
      if (n.includes("DESCRIPCION")) return i;
    }
    return 1;
  }

  function quantityColumn(row) {
    let fallback = -1;
    for (let i = 0; i < (row || []).length; i++) {
      const n = normalize(row[i]);
      if (!n) continue;
      if (n === "CANTIDAD" || n === "CANT." || n === "CANT") return i;
      if (n.includes("CANTIDAD") || n.startsWith("CANT")) fallback = i;
      if (n === "QTY") return i;
    }
    return fallback;
  }

  function parseQuantity(value) {
    const raw = safeText(value).trim().replace(/\s/g, "");
    if (!raw) return 1;
    let normalized = raw;
    if (/^\d{1,3}(\.\d{3})+,\d+$/.test(raw)) normalized = raw.replace(/\./g, "").replace(",", ".");
    else if (/^\d+,\d+$/.test(raw)) normalized = raw.replace(",", ".");
    const n = Number(normalized);
    return Number.isFinite(n) && n >= 0 ? n : 1;
  }

  function scoreRole(text) {
    const n = normalize(text);
    let paint = 0, fab = 0;

    const paintWords = ["PINTURA", "PINTADO", "PINTAR", "PINTADAS", "PINTADOS", "ESMALTE", "PRIMER", "FLEXEADO Y PINTADO", "APLICAR NEGRO", "APLICAR GRIS"];
    const fabWords = ["ARMADO", "ARMAR", "SOLDADO", "SOLDAR", "SOLDADURA", "FABRICACION", "FABRICAR", "CORTE", "PERFOR", "ESMERIL", "ENSAMBL", "EMPATE DE VIGA", "MODIFICACION"];

    paintWords.forEach(w => { if (n.includes(w)) paint += w.includes("PINT") ? 4 : 2; });
    fabWords.forEach(w => { if (n.includes(w)) fab += (w.includes("SOLD") || w.includes("ARMAD")) ? 4 : 2; });
    return { paint, fab };
  }

  function buildContractorBlocks(rows) {
    const contractorRows = [];
    for (let r = 0; r < rows.length; r++) {
      const name = getContractor(rows[r]);
      if (name) contractorRows.push({ r, name });
    }

    const blocks = [];
    contractorRows.forEach((item, idx) => {
      const start = item.r;
      const end = idx + 1 < contractorRows.length ? contractorRows[idx + 1].r - 1 : rows.length - 1;
      let markCol = -1, descCol = 1, qtyCol = -1, headerRow = -1;

      for (let r = start + 1; r <= Math.min(end, start + 15); r++) {
        const mc = getMarkColumn(rows[r]);
        if (mc >= 0) {
          markCol = mc;
          descCol = descriptionColumn(rows[r]);
          qtyCol = quantityColumn(rows[r]);
          headerRow = r;
          break;
        }
      }

      let paintScore = 0, fabScore = 0;
      const from = headerRow >= 0 ? headerRow + 1 : start + 1;
      for (let r = from; r <= end; r++) {
        const text = safeText(rows[r]?.[descCol] ?? rowText(rows[r]));
        const sc = scoreRole(text);
        paintScore += sc.paint;
        fabScore += sc.fab;
      }

      let role = "RESPONSABLE";
      if (paintScore > fabScore) role = "PINTOR";
      else if (fabScore > paintScore) role = "ARMADOR / SOLDADOR";

      blocks.push({
        contractor: item.name,
        project: inferProject(rows, start),
        start, end, markCol, descCol, qtyCol, headerRow, role, paintScore, fabScore
      });
    });
    return blocks;
  }

  function rowRole(block, description) {
    const sc = scoreRole(description);
    if (sc.paint > sc.fab && sc.paint > 0) return "PINTOR";
    if (sc.fab > sc.paint && sc.fab > 0) return "ARMADOR / SOLDADOR";
    return block.role;
  }

  function extractAllAssignments() {
    const out = [];
    for (const sheetName of workbook.SheetNames) {
      const rows = sheets[sheetName] || [];
      const blocks = buildContractorBlocks(rows);

      for (const block of blocks) {
        if (block.markCol < 0) continue;
        const from = block.headerRow >= 0 ? block.headerRow + 1 : block.start + 1;

        for (let r = from; r <= block.end; r++) {
          const mark = safeText(rows[r]?.[block.markCol]).trim();
          if (!mark) continue;

          const desc = safeText(rows[r]?.[block.descCol]).trim();
          const qty = block.qtyCol >= 0 ? parseQuantity(rows[r]?.[block.qtyCol]) : 1;

          out.push({
            project: block.project,
            contractor: block.contractor,
            role: rowRole(block, desc),
            mark,
            quantity: qty,
            description: desc,
            sheet: sheetName,
            row: r + 1
          });
        }
      }
    }
    return out;
  }

  function escapeRegex(text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function markMatches(value, term, mode) {
    const a = normalize(value);
    const b = normalize(term);
    if (!a || !b) return false;
    if (mode === "contains") return a.includes(b);
    if (a === b) return true;
    const rx = new RegExp(`(^|[^A-Z0-9])${escapeRegex(b)}([^A-Z0-9]|$)`);
    return rx.test(a);
  }

  function passesFilters(item) {
    if (sheetFilter.value !== "__ALL__" && item.sheet !== sheetFilter.value) return false;
    if (projectFilter.value !== "__ALL__" && item.project !== projectFilter.value) return false;
    if (contractorFilter.value !== "__ALL__" && item.contractor !== contractorFilter.value) return false;
    if (roleFilter.value !== "__ALL__" && item.role !== roleFilter.value) return false;
    return true;
  }

  function executeSearch() {
    if (!workbook) return;
    const terms = getTerms();
    if (!terms.length) return alert("Escribe al menos una marca. Ejemplo: CT-19");

    detailResults = [];
    for (const item of allAssignments) {
      if (!passesFilters(item)) continue;
      for (const term of terms) {
        if (!markMatches(item.mark, term, matchMode.value)) continue;
        detailResults.push({ ...item, searched: term });
      }
    }

    summaryResults = terms.map(term => {
      const items = detailResults.filter(x => x.searched === term);
      const contractors = [...new Set(items.map(x => x.contractor))];
      const projects = [...new Set(items.map(x => x.project))];
      const roles = [...new Set(items.map(x => x.role))];
      return {
        mark: term,
        projects,
        contractors,
        roles,
        quantity: items.reduce((sum, x) => sum + x.quantity, 0),
        appearances: items.length
      };
    });

    renderSearchResults(terms);
  }

  function executeSummary() {
    if (!workbook) return;

    detailResults = allAssignments.filter(passesFilters);
    const grouped = new Map();

    detailResults.forEach(item => {
      const key = [item.project, item.contractor, item.role].join("|||");
      if (!grouped.has(key)) {
        grouped.set(key, {
          project: item.project,
          contractor: item.contractor,
          role: item.role,
          quantity: 0,
          marks: new Set(),
          rows: 0
        });
      }
      const g = grouped.get(key);
      g.quantity += item.quantity;
      g.marks.add(item.mark);
      g.rows++;
    });

    summaryResults = [...grouped.values()].map(g => ({
      ...g,
      marks: [...g.marks].sort((a,b) => a.localeCompare(b, "es"))
    })).sort((a,b) =>
      a.project.localeCompare(b.project, "es") ||
      a.contractor.localeCompare(b.contractor, "es") ||
      a.role.localeCompare(b.role, "es")
    );

    renderAssignmentSummary();
  }

  function renderMetrics(assignments) {
    const old = resultsCard.querySelector(".summary-grid");
    if (old) old.remove();

    const grid = document.createElement("div");
    grid.className = "summary-grid";

    const totalQty = assignments.reduce((s, x) => s + x.quantity, 0);
    const projects = new Set(assignments.map(x => x.project)).size;
    const contractors = new Set(assignments.map(x => x.contractor)).size;
    const marks = new Set(assignments.map(x => x.mark)).size;

    [
      ["Cantidad asignada", formatNumber(totalQty)],
      ["Proyectos", projects],
      ["Contratistas", contractors],
      ["Marcas", marks]
    ].forEach(([label, value]) => {
      const card = document.createElement("div");
      card.className = "metric-card";
      card.innerHTML = `<span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong>`;
      grid.appendChild(card);
    });

    resultsCard.insertBefore(grid, resultsCard.querySelector(".chips"));
  }

  function formatNumber(n) {
    return Number(n).toLocaleString("es-HN", { maximumFractionDigits: 2 });
  }

  function escapeHtml(value) {
    return safeText(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function prepareResults() {
    resultsCard.classList.remove("hidden");
    resultsTable.innerHTML = "";
    fieldChips.innerHTML = "";
    const old = resultsCard.querySelector(".summary-grid");
    if (old) old.remove();
  }

  function renderSearchResults(terms) {
    prepareResults();

    const foundCount = summaryResults.filter(r => r.appearances > 0).length;
    resultsSummary.textContent = `${foundCount} de ${terms.length} marca(s) encontradas · ${detailResults.length} asignación(es) localizadas.`;

    if (!detailResults.length) {
      emptyResults.classList.remove("hidden");
      resultsWrap.classList.add("hidden");
      exportBtn.disabled = true;
      return;
    }

    emptyResults.classList.add("hidden");
    resultsWrap.classList.remove("hidden");
    exportBtn.disabled = false;
    renderMetrics(detailResults);

    ["Proyecto", "Contratista", "Marca", "Proceso", "Cantidad"].forEach(label => {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = label;
      fieldChips.appendChild(chip);
    });

    const headers = ["Marca buscada", "Marca encontrada", "Proyecto", "Contratista", "Proceso", "Cantidad", "Hoja", "Fila", "Descripción"];
    const thead = document.createElement("thead");
    const hr = document.createElement("tr");
    headers.forEach(label => {
      const th = document.createElement("th");
      th.textContent = label;
      hr.appendChild(th);
    });
    thead.appendChild(hr);

    const tbody = document.createElement("tbody");
    detailResults.forEach(item => {
      const tr = document.createElement("tr");
      [
        item.searched, item.mark, item.project, item.contractor,
        item.role, formatNumber(item.quantity), item.sheet, item.row, item.description
      ].forEach((value, i) => {
        const td = document.createElement("td");
        td.textContent = safeText(value);
        if (i === 0 || i === 1) td.classList.add("match");
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    resultsTable.append(thead, tbody);
  }

  function renderAssignmentSummary() {
    prepareResults();

    resultsSummary.textContent = `${summaryResults.length} grupo(s) · ${detailResults.length} registro(s) de asignación filtrados.`;

    if (!detailResults.length) {
      emptyResults.classList.remove("hidden");
      resultsWrap.classList.add("hidden");
      exportBtn.disabled = true;
      return;
    }

    emptyResults.classList.add("hidden");
    resultsWrap.classList.remove("hidden");
    exportBtn.disabled = false;
    renderMetrics(detailResults);

    ["Proyecto", "Contratista", "Proceso", "Cantidad asignada", "Marcas"].forEach(label => {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = label;
      fieldChips.appendChild(chip);
    });

    const headers = ["Proyecto", "Contratista", "Proceso", "Cantidad asignada", "Marcas únicas", "Marcas"];
    const thead = document.createElement("thead");
    const hr = document.createElement("tr");
    headers.forEach(label => {
      const th = document.createElement("th");
      th.textContent = label;
      hr.appendChild(th);
    });
    thead.appendChild(hr);

    const tbody = document.createElement("tbody");
    summaryResults.forEach(item => {
      const tr = document.createElement("tr");
      [
        item.project,
        item.contractor,
        item.role,
        formatNumber(item.quantity),
        item.marks.length,
        item.marks.join(", ")
      ].forEach(value => {
        const td = document.createElement("td");
        td.textContent = safeText(value);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    resultsTable.append(thead, tbody);
  }

  function renderPreview(sheetName) {
    const allRows = sheets[sheetName] || [];
    previewSheet.value = sheetName;
    previewTable.innerHTML = "";
    if (!allRows.length) return;

    const rows = allRows.slice(0, 60);
    const width = Math.max(0, ...rows.map(r => (r || []).length));
    const thead = document.createElement("thead"), hr = document.createElement("tr");
    const nh = document.createElement("th"); nh.textContent = "#"; hr.appendChild(nh);
    for (let c = 0; c < width; c++) {
      const th = document.createElement("th"); th.textContent = excelColumn(c); hr.appendChild(th);
    }
    thead.appendChild(hr);

    const tbody = document.createElement("tbody");
    rows.forEach((row, r) => {
      const tr = document.createElement("tr");
      const rn = document.createElement("td"); rn.textContent = r + 1; tr.appendChild(rn);
      for (let c = 0; c < width; c++) {
        const td = document.createElement("td"); td.textContent = safeText(row?.[c]); tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    previewTable.append(thead, tbody);
  }

  function clearSearch() {
    searchInput.value = "";
    summaryResults = [];
    detailResults = [];
    resultsCard.classList.add("hidden");
    exportBtn.disabled = true;
    searchInput.focus();
  }

  function exportCSV() {
    if (!detailResults.length) return;
    const esc = v => `"${String(v ?? "").replace(/"/g, '""')}"`;

    const isSummary = summaryResults.length && summaryResults[0] && "marks" in summaryResults[0];
    const lines = [];

    if (isSummary) {
      lines.push(["Proyecto", "Contratista", "Proceso", "Cantidad asignada", "Marcas únicas", "Marcas"].map(esc).join(","));
      summaryResults.forEach(r => {
        lines.push([
          r.project, r.contractor, r.role, r.quantity, r.marks.length, r.marks.join(" | ")
        ].map(esc).join(","));
      });
      lines.push("");
      lines.push(["DETALLE DE ASIGNACIONES"].map(esc).join(","));
    }

    lines.push(["Proyecto", "Contratista", "Marca", "Proceso", "Cantidad", "Hoja", "Fila", "Descripción"].map(esc).join(","));
    detailResults.forEach(r => {
      lines.push([
        r.project, r.contractor, r.mark, r.role, r.quantity, r.sheet, r.row, r.description
      ].map(esc).join(","));
    });

    const blob = new Blob(["\uFEFF" + lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob), a = document.createElement("a");
    a.href = url;
    a.download = isSummary ? "resumen_asignaciones.csv" : "busqueda_marcas.csv";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }
})();