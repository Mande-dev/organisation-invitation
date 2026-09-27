/**
 * Recherche de noms — charge le Excel via « Choisir le fichier ».
 * Le fichier est mémorisé (IndexedDB) : un rafraîchissement le recharge.
 * Double-clic sur index.html OK (pas de fetch).
 */

const DB_NAME = "organisation-invitation";
const DB_STORE = "excel";
const DB_KEY = "current";

/**
 * @typedef {{ place: string, name: string }} Guest
 * @typedef {{ id: string, guests: Guest[] }} Table
 */

/** @type {Table[]} */
let tables = [];

const searchInput = document.getElementById("search-input");
const searchBtn = document.getElementById("search-btn");
const statusEl = document.getElementById("status");
const resultsEl = document.getElementById("results");
const fileInput = document.getElementById("file-input");
const fileNameEl = document.getElementById("file-name");
const fileBtnLabel = document.getElementById("file-btn-label");

function setStatus(message, kind) {
  statusEl.textContent = message;
  statusEl.className = "status" + (kind ? ` is-${kind}` : "");
}

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(DB_STORE)) {
        db.createObjectStore(DB_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("IndexedDB indisponible."));
  });
}

async function saveExcelToDb(fileName, buffer) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readwrite");
    tx.oncomplete = () => {
      db.close();
      resolve();
    };
    tx.onerror = () => {
      db.close();
      reject(tx.error || new Error("Échec de la sauvegarde locale."));
    };
    tx.objectStore(DB_STORE).put(
      { name: fileName, buffer, savedAt: Date.now() },
      DB_KEY
    );
  });
}

async function loadExcelFromDb() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readonly");
    const req = tx.objectStore(DB_STORE).get(DB_KEY);
    req.onsuccess = () => {
      db.close();
      resolve(req.result || null);
    };
    req.onerror = () => {
      db.close();
      reject(req.error || new Error("Échec de la lecture locale."));
    };
  });
}

function normalize(text) {
  return String(text ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function cellToString(value) {
  if (value == null || value === "") return "";
  if (value instanceof Date) return value.toLocaleDateString("fr-FR");
  return String(value).trim();
}

function isTableHeader(value) {
  return /^table\s*\d+/i.test(cellToString(value));
}

function parseSeatingChart(workbook) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];

  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: true,
  });

  const headers = [];
  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r] || [];
    for (let c = 0; c < row.length; c++) {
      const raw = cellToString(row[c]);
      if (isTableHeader(raw)) {
        headers.push({ row: r, col: c, id: raw.replace(/\s+/g, " ").trim() });
      }
    }
  }

  const parsed = [];
  for (const header of headers) {
    const nameCol = header.col + 1;
    const nextSameCol = headers
      .filter((h) => h.col === header.col && h.row > header.row)
      .sort((a, b) => a.row - b.row)[0];
    const endRow = nextSameCol ? nextSameCol.row : matrix.length;

    const guests = [];
    let emptyStreak = 0;

    for (let r = header.row + 1; r < endRow; r++) {
      const row = matrix[r] || [];
      const place = cellToString(row[header.col]);
      const name = cellToString(row[nameCol]);

      if (isTableHeader(place) || isTableHeader(name)) break;

      if (!name) {
        emptyStreak += 1;
        if (emptyStreak >= 2 && guests.length) break;
        continue;
      }

      emptyStreak = 0;
      guests.push({ place, name });
    }

    if (guests.length) parsed.push({ id: header.id, guests });
  }

  parsed.sort((a, b) => {
    const na = parseInt(a.id.replace(/\D/g, ""), 10) || 0;
    const nb = parseInt(b.id.replace(/\D/g, ""), 10) || 0;
    return na - nb;
  });

  return parsed;
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Une entrée par personne : { tableId, name, guests }. */
function searchTables(query) {
  const needle = normalize(query);
  if (!needle) return [];

  const hits = [];
  for (const table of tables) {
    for (const guest of table.guests) {
      if (normalize(guest.name).includes(needle)) {
        hits.push({
          tableId: table.id,
          name: guest.name,
          guests: table.guests,
        });
      }
    }
  }
  return hits;
}

function renderResults(hits, query) {
  resultsEl.innerHTML = "";
  if (!query.trim()) return;

  if (!hits.length) {
    resultsEl.innerHTML = `
      <div class="empty-state">
        Aucun résultat pour « ${escapeHtml(query)} ».
      </div>`;
    setStatus("Aucun résultat.", "empty");
    return;
  }

  setStatus(
    `${hits.length} résultat${hits.length > 1 ? "s" : ""} — cliquez pour voir la table.`,
    "ready"
  );

  hits.forEach((hit, index) => {
    const block = document.createElement("article");
    block.className = "result-card";
    block.dataset.index = String(index);

    const summary = document.createElement("button");
    summary.type = "button";
    summary.className = "result-summary";
    summary.setAttribute("aria-expanded", "false");
    summary.innerHTML = `
      <span class="result-main">
        <span class="result-table">${escapeHtml(hit.tableId)}</span>
        <span class="result-sep" aria-hidden="true">—</span>
        <span class="result-name">${escapeHtml(hit.name)}</span>
      </span>
      <span class="result-hint">Voir la table</span>
    `;

    const detail = document.createElement("div");
    detail.className = "result-detail";
    detail.hidden = true;

    const listWrap = document.createElement("div");
    listWrap.className = "table-wrap";

    const tableEl = document.createElement("table");
    tableEl.innerHTML = "<thead><tr><th>Place</th><th>Nom</th></tr></thead>";
    const tbody = document.createElement("tbody");

    for (const guest of hit.guests) {
      const tr = document.createElement("tr");
      const isMatch = guest.name === hit.name;
      if (isMatch) tr.classList.add("is-match");

      const tdPlace = document.createElement("td");
      tdPlace.textContent = guest.place;

      const tdName = document.createElement("td");
      tdName.textContent = guest.name;
      if (isMatch) tdName.classList.add("mark");

      tr.appendChild(tdPlace);
      tr.appendChild(tdName);
      tbody.appendChild(tr);
    }

    tableEl.appendChild(tbody);
    listWrap.appendChild(tableEl);
    detail.appendChild(listWrap);

    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = "result-close";
    closeBtn.textContent = "Fermer";
    detail.appendChild(closeBtn);

    function setOpen(open) {
      block.classList.toggle("is-open", open);
      detail.hidden = !open;
      summary.setAttribute("aria-expanded", open ? "true" : "false");
      summary.querySelector(".result-hint").textContent = open
        ? "Masquer"
        : "Voir la table";
    }

    summary.addEventListener("click", () => {
      setOpen(detail.hidden);
    });

    closeBtn.addEventListener("click", () => {
      setOpen(false);
      summary.focus();
    });

    block.appendChild(summary);
    block.appendChild(detail);
    resultsEl.appendChild(block);
  });
}

function runSearch() {
  if (!tables.length) return;
  renderResults(searchTables(searchInput.value), searchInput.value);
}

function markFileLoaded(fileLabel) {
  fileNameEl.textContent = fileLabel || "";
  if (fileBtnLabel) {
    fileBtnLabel.textContent = "Changer le fichier Excel";
  }
}

function enableSearch(message) {
  searchInput.disabled = false;
  searchBtn.disabled = false;
  const guestCount = tables.reduce((n, t) => n + t.guests.length, 0);
  setStatus(message || `Prêt — ${tables.length} tables, ${guestCount} noms.`, "ready");
  searchInput.focus();
}

function applyTables(parsed, fileLabel) {
  if (!parsed.length) {
    throw new Error("Aucune table détectée dans le fichier Excel (en-têtes « TABLE 1 », …).");
  }
  tables = parsed;
  resultsEl.innerHTML = "";
  markFileLoaded(fileLabel);
  enableSearch(
    fileLabel
      ? `Prêt — « ${fileLabel} » : ${tables.length} tables, ${tables.reduce((n, t) => n + t.guests.length, 0)} noms.`
      : undefined
  );
  if (searchInput.value.trim()) runSearch();
}

function ensureSheetJs() {
  if (typeof XLSX === "undefined") {
    throw new Error(
      "SheetJS n’est pas chargé (CDN). Vérifiez votre connexion internet."
    );
  }
}

function parseBuffer(buffer, fileLabel) {
  ensureSheetJs();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  applyTables(parseSeatingChart(workbook), fileLabel);
}

async function loadFromFile(file) {
  setStatus("Lecture du fichier Excel…", "loading");
  const buffer = await file.arrayBuffer();
  parseBuffer(buffer, file.name);
  try {
    await saveExcelToDb(file.name, buffer);
  } catch (err) {
    console.warn("Fichier chargé, mais non mémorisé :", err);
  }
}

async function restoreSavedExcel() {
  setStatus("Restauration du fichier mémorisé…", "loading");
  const saved = await loadExcelFromDb();
  if (!saved || !saved.buffer) return false;
  parseBuffer(saved.buffer, saved.name);
  return true;
}

searchBtn.addEventListener("click", runSearch);
searchInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter") runSearch();
});

let debounceTimer;
searchInput.addEventListener("input", () => {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(runSearch, 180);
});

fileInput.addEventListener("change", async () => {
  const file = fileInput.files && fileInput.files[0];
  if (!file) return;

  fileNameEl.textContent = file.name;
  try {
    await loadFromFile(file);
  } catch (err) {
    console.error(err);
    tables = [];
    searchInput.disabled = true;
    searchBtn.disabled = true;
    setStatus(err.message || "Erreur de lecture.", "error");
    resultsEl.innerHTML = `<div class="error-state">${escapeHtml(err.message || "Erreur")}</div>`;
  }
});

async function init() {
  try {
    const restored = await restoreSavedExcel();
    if (!restored) {
      setStatus("Choisissez le fichier Excel pour commencer.", "loading");
    }
  } catch (err) {
    console.warn(err);
    setStatus("Choisissez le fichier Excel pour commencer.", "loading");
  }
}

init();
