/*
 * MBS Supporting Documents Add-in
 * taskpane.js — all business logic
 */

"use strict";

// ── Constants ─────────────────────────────────────────────────────────────────
const EMBED_SHEET   = "EMBEDDOCS";
const SUPDOCS_SHEET = "SUPDOCS";
const SUPPORTED_EXT = new Set([".pdf",".png",".jpg",".jpeg",".bmp",".tif",".tiff",".docx",".xlsx"]);
const EXT_COLORS    = {
  ".pdf":"#c0392b",".docx":"#1565c0",".xlsx":"#1a5c3a",
  ".png":"#6a0dad",".jpg":"#6a0dad",".jpeg":"#6a0dad",
  ".bmp":"#e67e22",".tif":"#e67e22",".tiff":"#e67e22"
};

// ── State ─────────────────────────────────────────────────────────────────────
let currentSheet = "";
let allDocs = {};   // { sheetName: [{fname, ext, b64}] }

// ── Init ──────────────────────────────────────────────────────────────────────
Office.onReady(function(info) {
  if (info.host !== Office.HostType.Excel) return;

  document.getElementById("add-btn")   .addEventListener("click", onAddClick);
  document.getElementById("file-input").addEventListener("change", onFilesSelected);
  document.getElementById("save-btn")  .addEventListener("click", onSave);
  document.getElementById("export-btn").addEventListener("click", onExport);
  document.getElementById("drop-zone") .addEventListener("click", onAddClick);

  setupDragDrop();
  Excel.run(initAddin);
});

async function initAddin(ctx) {
  try {
    // Check EMBEDDOCS sheet exists
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    const names = sheets.items.map(s => s.name);
    if (!names.includes(EMBED_SHEET)) {
      show("error-state");
      hide("loading");
      return;
    }

    // Load all embedded documents
    await loadAllDocs(ctx);

    // Get current sheet name
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    await ctx.sync();
    currentSheet = ws.name;

    // Listen for sheet changes
    ctx.workbook.worksheets.onActivated.add(onSheetActivated);
    await ctx.sync();

    updateUI();
    hide("loading");
    show("main");

  } catch(e) {
    showStatus("Error loading add-in: " + e.message, "error");
    hide("loading");
    show("main");
  }
}

// ── Sheet activation ──────────────────────────────────────────────────────────
async function onSheetActivated(event) {
  currentSheet = event.worksheetId || event.worksheetName;
  // Get the actual name if we only have an id
  await Excel.run(async ctx => {
    try {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();
      currentSheet = ws.name;
      updateSheetBanner();
    } catch(e) {}
  });
}

// ── Load all docs from EMBEDDOCS sheet ─────────────────────────────────────────
async function loadAllDocs(ctx) {
  allDocs = {};
  try {
    const ws = ctx.workbook.worksheets.getItem(EMBED_SHEET);
    const used = ws.getUsedRange();
    used.load("values");
    await ctx.sync();

    const vals = used.values;
    if (!vals || vals.length < 2) return;

    for (let i = 1; i < vals.length; i++) {
      const row = vals[i];
      const wp    = String(row[0] || "").trim();
      const fname = String(row[1] || "").trim();
      const ext   = String(row[2] || "").trim().toLowerCase();
      const b64   = String(row[3] || "").trim().replace(/\|/g, "");
      if (!wp || !b64) continue;
      if (!allDocs[wp]) allDocs[wp] = [];
      allDocs[wp].push({ fname, ext, b64 });
    }
  } catch(e) {
    // EMBEDDOCS might be empty — that's fine
  }
}

// ── UI update ─────────────────────────────────────────────────────────────────
function updateUI() {
  updateSheetBanner();
  renderDocList();
  updateSummary();
  hideStatus();
}

function updateSheetBanner() {
  const nameEl  = document.getElementById("sheet-name");
  const countEl = document.getElementById("sheet-banner-count");
  nameEl.textContent = currentSheet || "—";
  const n = (allDocs[currentSheet] || []).length;
  countEl.textContent = n > 0 ? `${n} doc${n!==1?"s":""}` : "";
}

function renderDocList() {
  const list = document.getElementById("doc-list");
  list.innerHTML = "";
  const docs = allDocs[currentSheet] || [];

  if (docs.length === 0) {
    list.innerHTML = '<div class="no-docs">No documents attached to this workpaper yet</div>';
    return;
  }

  docs.forEach((doc, idx) => {
    const ext   = doc.ext || ".file";
    const color = EXT_COLORS[ext] || "#6c7a8d";
    const label = ext.replace(".","").toUpperCase().slice(0,4);
    const sizeKB = Math.round((doc.b64.length * 3/4) / 1024);
    const sizeStr = sizeKB > 1024 ? `${(sizeKB/1024).toFixed(1)} MB` : `${sizeKB} KB`;

    const card = document.createElement("div");
    card.className = "doc-card";
    card.innerHTML = `
      <div class="doc-badge" style="background:${color}">${label}</div>
      <div class="doc-info">
        <div class="doc-name" title="${esc(doc.fname)}">${esc(doc.fname)}</div>
        <div class="doc-meta">${sizeStr} · embedded ✓</div>
      </div>
      <button class="doc-remove" title="Remove" data-idx="${idx}">✕</button>
    `;
    card.querySelector(".doc-remove").addEventListener("click", () => removeDoc(idx));
    list.appendChild(card);
  });
}

function updateSummary() {
  const total  = Object.values(allDocs).reduce((s,v) => s+v.length, 0);
  const sheets = Object.keys(allDocs).filter(k => allDocs[k].length > 0).length;
  document.getElementById("total-docs").textContent   = total;
  document.getElementById("total-sheets").textContent = sheets;
}

// ── File add ──────────────────────────────────────────────────────────────────
function onAddClick() {
  document.getElementById("file-input").click();
}

function onFilesSelected(e) {
  processFiles(Array.from(e.target.files));
  e.target.value = "";  // reset so same file can be added again
}

function setupDragDrop() {
  const zone = document.getElementById("drop-zone");
  zone.addEventListener("dragover",  e => { e.preventDefault(); zone.classList.add("drag-over"); });
  zone.addEventListener("dragleave", e => zone.classList.remove("drag-over"));
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    processFiles(Array.from(e.dataTransfer.files));
  });
}

function processFiles(files) {
  if (!currentSheet) { showStatus("No workpaper selected.", "error"); return; }
  if (!allDocs[currentSheet]) allDocs[currentSheet] = [];

  const existing = new Set(allDocs[currentSheet].map(d => d.fname));
  let added = 0, skipped = [];

  const readers = files.map(file => new Promise((resolve) => {
    const ext = "." + file.name.split(".").pop().toLowerCase();
    if (!SUPPORTED_EXT.has(ext)) {
      skipped.push(file.name);
      resolve();
      return;
    }
    // Deduplicate filename
    let fname = file.name;
    if (existing.has(fname)) {
      const base = file.name.replace(/\.[^.]+$/, "");
      const sfx  = ext;
      let i = 2;
      while (existing.has(`${base}_${i}${sfx}`)) i++;
      fname = `${base}_${i}${sfx}`;
    }

    const reader = new FileReader();
    reader.onload = e => {
      // e.target.result is data:...;base64,XXXX — strip the prefix
      const b64 = e.target.result.split(",")[1];
      allDocs[currentSheet].push({ fname, ext, b64 });
      existing.add(fname);
      added++;
      resolve();
    };
    reader.onerror = () => { skipped.push(file.name); resolve(); };
    reader.readAsDataURL(file);
  }));

  Promise.all(readers).then(() => {
    renderDocList();
    updateSummary();
    updateSheetBanner();
    if (skipped.length > 0) {
      showStatus(`Added ${added} file${added!==1?"s":""}. Skipped: ${skipped.join(", ")}`, "info");
    } else if (added > 0) {
      showStatus(`${added} file${added!==1?"s":""} added — click Save & Embed All to save`, "info");
    }
  });
}

// ── Remove doc ────────────────────────────────────────────────────────────────
function removeDoc(idx) {
  if (!allDocs[currentSheet]) return;
  const doc = allDocs[currentSheet][idx];
  if (!doc) return;
  if (!confirm(`Remove "${doc.fname}" from ${currentSheet}?`)) return;
  allDocs[currentSheet].splice(idx, 1);
  renderDocList();
  updateSummary();
  updateSheetBanner();
  showStatus("Document removed — click Save & Embed All to save", "info");
}

// ── Save & Embed ──────────────────────────────────────────────────────────────
async function onSave() {
  setProgress("Preparing…", 5, true);
  setBtns(false);

  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(EMBED_SHEET);

      // Clear existing data (keep row 1 header)
      setProgress("Clearing old data…", 15);
      try {
        const used = ws.getUsedRange();
        used.load("rowCount");
        await ctx.sync();
        if (used.rowCount > 1) {
          ws.getRangeByIndexes(1, 0, used.rowCount - 1, 4).delete(Excel.DeleteShiftDirection.up);
          await ctx.sync();
        }
      } catch(e) {
        // Sheet might be empty — fine
      }

      // Write header
      ws.getRangeByIndexes(0, 0, 1, 4).values = [
        ["WorkpaperSheet","FileName","FileExt","Base64Content"]
      ];

      // Write all docs
      const entries = [];
      for (const [sheet, docs] of Object.entries(allDocs)) {
        for (const doc of docs) {
          if (!doc.b64) continue;
          // Split base64 into 30k-char chunks joined by |
          const chunks = [];
          for (let i = 0; i < doc.b64.length; i += 30000) {
            chunks.push(doc.b64.slice(i, i+30000));
          }
          entries.push([sheet, doc.fname, doc.ext, chunks.join("|")]);
        }
      }

      const total = entries.length;
      if (total > 0) {
        // Write in batches of 50 rows to avoid timeout
        const BATCH = 50;
        for (let i = 0; i < total; i += BATCH) {
          const batch = entries.slice(i, i+BATCH);
          ws.getRangeByIndexes(i+1, 0, batch.length, 4).values = batch;
          await ctx.sync();
          setProgress(`Saving document ${Math.min(i+BATCH, total)} of ${total}…`,
                      15 + Math.round(75 * (i+BATCH) / total));
        }
      }

      await ctx.sync();
    });

    setProgress("Saved ✓", 100);
    setTimeout(() => {
      hideProgress();
      showStatus(`✓ ${Object.values(allDocs).reduce((s,v)=>s+v.length,0)} documents embedded in workbook`, "success");
      setBtns(true);
    }, 1200);

  } catch(e) {
    hideProgress();
    showStatus("Save failed: " + e.message, "error");
    setBtns(true);
  }
}

// ── Export PDF ────────────────────────────────────────────────────────────────
async function onExport() {
  showStatus(
    "To export PDF: use the MBS_Export_with_SupportingDocs.py script on your computer.\n\n" +
    "It reads the embedded documents from this workbook and merges everything into one PDF.",
    "info"
  );
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function show(id) { document.getElementById(id).style.display = ""; }
function hide(id) { document.getElementById(id).style.display = "none"; }

function showStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className   = `status ${type}`;
  el.style.display = "";
}
function hideStatus() {
  document.getElementById("status").style.display = "none";
}

function setProgress(label, pct, show_=false) {
  const wrap = document.getElementById("progress-wrap");
  if (show_) wrap.style.display = "";
  document.getElementById("progress-label").textContent = label;
  document.getElementById("progress-bar").style.width   = pct + "%";
}
function hideProgress() {
  document.getElementById("progress-wrap").style.display = "none";
}

function setBtns(enabled) {
  document.getElementById("save-btn").disabled   = !enabled;
  document.getElementById("export-btn").disabled = !enabled;
  document.getElementById("add-btn").disabled    = !enabled;
}

function esc(str) {
  return String(str)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
