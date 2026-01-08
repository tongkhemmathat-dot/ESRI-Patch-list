let DATA = {
  "All_Enterprise": {
    "columns": ["Released", "Release Month", "Component", "Security", "Patch Name", "Support Page", "PatchFiles", "__row"],
    "rows": []
  }
};

const PATCHES_JSON_URL = "./data/patches.json";
let lastPatchFetchTs = 0;
let activeSheet = "All_Enterprise";

const themeBtn = document.getElementById("themeBtn");
function applyTheme(mode) {
  document.documentElement.setAttribute("data-theme", mode);
  const label = (mode === "dark") ? "Switch to light mode" : "Switch to dark mode";
  themeBtn.setAttribute("aria-label", label);
  themeBtn.setAttribute("title", label);
  themeBtn.setAttribute("aria-pressed", String(mode === "dark"));
  try { localStorage.setItem("theme", mode); } catch (e) {}
}
const storedTheme = (() => { try { return localStorage.getItem("theme"); } catch (e) { return null; } })();
const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
applyTheme(storedTheme || (prefersDark ? "dark" : "light"));
themeBtn.addEventListener("click", () => {
  const current = document.documentElement.getAttribute("data-theme") || "light";
  applyTheme(current === "dark" ? "light" : "dark");
});

const versionSel = document.getElementById("versionSel");

function rebuildPatchSheetOptions() {
  if (!versionSel) return;
  const sheets = Object.keys(DATA || {});
  const hasAll = sheets.includes("All_Enterprise");
  const versionSheets = sheets
    .filter(s => s !== "All_Enterprise")
    .sort((a,b) => a.localeCompare(b, undefined, {numeric:true}));
  const ordered = [];
  if (hasAll) ordered.push("All_Enterprise");
  ordered.push(...versionSheets);
  versionSel.innerHTML = "";
  for (const s of ordered) {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = (s === "All_Enterprise") ? "All versions" : s.replaceAll("_", ".");
    versionSel.appendChild(opt);
  }
  const nextActive = ordered.includes(activeSheet)
    ? activeSheet
    : (hasAll ? "All_Enterprise" : ordered[0]);
  activeSheet = nextActive || "All_Enterprise";
  versionSel.value = activeSheet;
}

function rebuildPatchComponentOptions() {
  const compSel = document.getElementById("componentSel");
  if (!compSel) return;
  const sheet = DATA?.All_Enterprise || DATA?.[activeSheet];
  if (!sheet?.columns || !sheet?.rows) return;
  const compIdx = sheet.columns.indexOf("Component");
  if (compIdx < 0) return;
  const comps = Array.from(new Set(sheet.rows.map(r => r[compIdx]).filter(Boolean))).sort((a,b) => String(a).localeCompare(String(b)));
  if (!comps.length) return;
  compSel.innerHTML = '<option value="">All components</option>';
  for (const c of comps) {
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = c;
    compSel.appendChild(opt);
  }
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function splitUrls(text) {
  const t = String(text || "").trim();
  if (!t) return [];
  return t.split(/\s+/).filter(x => x.toLowerCase().startsWith("http"));
}

function renderLinksCell(val) {
  const urls = splitUrls(val);
  if (!urls.length) return "";
  const maxShow = 2;
  const shown = urls.slice(0, maxShow);
  const rest = urls.slice(maxShow);

  let out = '<div class="actions">';
  for (const u of shown) {
    const esc = escapeHtml(u);
    out += '<a class="pill open" href="' + esc + '" target="_blank" rel="noopener noreferrer">Open</a>';
    out += '<a class="pill dl" href="' + esc + '" target="_blank" rel="noopener noreferrer" download>Download</a>';
  }
  if (rest.length) {
    out += '<a class="pill" href="#" onclick="alert(\'More links:\\n\\n\' + ' + JSON.stringify(rest.join('\n')) + '); return false;">More (' + rest.length + ')</a>';
  }
  out += '</div>';
  return out;
}

function renderSupportCell(val) {
  const urls = splitUrls(val);
  if (!urls.length) {
    const t = String(val||"").trim();
    if (t.toLowerCase().startsWith("http")) urls.push(t);
  }
  if (!urls.length) return "";
  const esc = escapeHtml(urls[0]);
  return '<a class="pill open" href="' + esc + '" target="_blank" rel="noopener noreferrer">Open support page</a>';
}

function extractFirstHttpUrl(val) {
  if (Array.isArray(val)) {
    for (const v of val) {
      const found = extractFirstHttpUrl(v);
      if (found) return found;
    }
    return null;
  }
  const s = String(val || "").trim();
  if (!s) return null;
  const parts = s.split(/\s+/);
  for (const p of parts) {
    if (p.toLowerCase().startsWith("http")) return p;
  }
  return null;
}

function normalizeVersionToToken(versionStr) {
  const s = String(versionStr || "").trim();
  if (!s) return "";
  return s.replaceAll(".", "");
}

function pickPatchFileByVersion(patchFilesArray, versionStr) {
  const token = normalizeVersionToToken(versionStr);
  if (!token) return null;
  const files = Array.isArray(patchFilesArray) ? patchFilesArray : [];
  if (!files.length) return null;
  const patterns = [
    new RegExp(`ArcGIS-${token}-`, "i"),
    new RegExp(`/PFA-${token}-`, "i"),
    new RegExp(`/S-${token}-`, "i")
  ];
  for (const fileUrl of files) {
    const url = String(fileUrl || "").trim();
    if (!url) continue;
    if (!/gisupdates\.esri\.com/i.test(url)) continue;
    if (patterns.some(re => re.test(url))) return url;
  }
  return null;
}

function getRowVersion(rowObj) {
  const fromRow = rowObj?.version || rowObj?.Version || "";
  if (fromRow) return String(fromRow);
  if (String(activeSheet).startsWith("v")) {
    return activeSheet.slice(1).replaceAll("_", ".");
  }
  return "";
}

function getDirectPatchDownloadUrl(rowObj) {
  const raw = rowObj?._raw || rowObj || {};
  const versionStr = getRowVersion(rowObj);
  const patchFiles = raw?.PatchFiles || raw?.patchFiles || raw?.patchfiles;
  if (Array.isArray(patchFiles) && patchFiles.length) {
    const match = pickPatchFileByVersion(patchFiles, versionStr);
    return match || null;
  }
  const preferredKeys = [
    "download_url","download url","download","direct download","download link","file url","file_url","qfe_url","qfe url","url"
  ];
  let directUrl = null;
  const keyMap = new Map(Object.keys(raw).map(k => [String(k).toLowerCase(), k]));
  for (const key of preferredKeys) {
    const rawKey = keyMap.get(key);
    if (!rawKey) continue;
    directUrl = extractFirstHttpUrl(raw[rawKey]);
    if (directUrl) break;
  }
  if (directUrl && /gisupdates\.esri\.com/i.test(directUrl)) return directUrl;
  return null;
}

function renderPatchActionsCell(row, colIndex, columns) {
  const rowObj = colIndex.has("__row")
    ? row[colIndex.get("__row")]
    : Object.fromEntries(columns.map((c, i) => [c, row[i]]));
  const directUrl = getDirectPatchDownloadUrl(rowObj);
  let out = '<div class="actions">';
  if (directUrl) {
    out += '<a class="pill dl" href="' + escapeHtml(directUrl) + '" target="_blank" rel="noopener noreferrer" download>Download patch</a>';
  } else {
    out += '<span class="pill disabled" title="No direct download URL available in data/patches.json">Direct DL N/A</span>';
  }
  out += '</div>';
  return out;
}

function normalizeReleaseDate(val) {
  const s = String(val || "").trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const mm = m[1].padStart(2, "0");
    const dd = m[2].padStart(2, "0");
    return `${m[3]}-${mm}-${dd}`;
  }
  return s;
}

function releaseMonthFromDate(val) {
  const s = String(val || "").trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s.slice(0, 7);
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[1].padStart(2, "0")}`;
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  return "";
}

function inferPatchComponent(products) {
  const text = String(products || "").toLowerCase();
  if (text.includes("portal")) return "Portal";
  if (text.includes("data store")) return "Data Store";
  if (text.includes("notebook")) return "Notebook";
  if (text.includes("geoevent")) return "GeoEvent";
  if (text.includes("server")) return "ArcGIS Server";
  return "Other";
}

function normalizeSecurity(val) {
  const s = String(val || "").toLowerCase();
  if (s.includes("security") || s === "y" || s === "yes" || s === "true") return "Y";
  return "N";
}

function normalizePatchesJson(raw) {
  if (raw?.All_Enterprise?.columns && raw?.All_Enterprise?.rows) return raw;
  if (raw?.columns && raw?.rows) return { All_Enterprise: raw };
  const columns = ["Released", "Release Month", "Component", "Security", "Patch Name", "Support Page", "PatchFiles", "__row"];
  const data = {};
  const addRow = (sheetKey, row) => {
    if (!data[sheetKey]) data[sheetKey] = { columns: [...columns], rows: [] };
    data[sheetKey].rows.push(row);
  };
  const groups = [];
  if (Array.isArray(raw)) {
    groups.push({ version: null, patches: raw });
  } else if (raw && Array.isArray(raw.Product)) {
    groups.push(...raw.Product);
  } else if (raw && Array.isArray(raw.patches)) {
    groups.push({ version: raw.version || null, patches: raw.patches });
  } else if (raw && typeof raw === "object") {
    for (const [key, val] of Object.entries(raw)) {
      if (Array.isArray(val)) groups.push({ version: key, patches: val });
    }
  }

  for (const g of groups) {
    const groupVersion = g?.version ? String(g.version).trim() : "";
    const patches = Array.isArray(g?.patches) ? g.patches : [];
    for (const p of patches) {
      const rowVersion = groupVersion || String(p?.version || p?.Version || "").trim();
      const sheetKey = rowVersion ? `v${rowVersion.replaceAll(".", "_")}` : "All_Enterprise";
      const released = normalizeReleaseDate(p?.ReleaseDate || p?.Released || p?.released || "");
      const month = releaseMonthFromDate(released);
      const component = inferPatchComponent(p?.Products || p?.Component || p?.component);
      const security = normalizeSecurity(p?.Critical || p?.Security || p?.security);
      const name = p?.Name || p?.["Patch Name"] || p?.patch || "";
      const support = p?.url || p?.["Support Page"] || p?.support || "";
      const patchFiles = p?.PatchFiles || p?.patchFiles || p?.patchfiles || "";
      const rowObj = {
        "Released": released,
        "Release Month": month,
        "Component": component,
        "Security": security,
        "Patch Name": name,
        "Support Page": support,
        "PatchFiles": patchFiles,
        "version": rowVersion,
        _raw: p
      };
      const row = [
        rowObj["Released"],
        rowObj["Release Month"],
        rowObj["Component"],
        rowObj["Security"],
        rowObj["Patch Name"],
        rowObj["Support Page"],
        rowObj["PatchFiles"],
        rowObj
      ];
      addRow("All_Enterprise", row);
      if (sheetKey !== "All_Enterprise") addRow(sheetKey, row);
    }
  }
  return data;
}

async function loadPatchesLatest(force = false) {
  const now = Date.now();
  if (!force && now - lastPatchFetchTs < 3000) return;
  lastPatchFetchTs = now;
  const statusEl = document.getElementById("status");
  if (statusEl) statusEl.textContent = "Loading patches...";
  try {
    const url = `${PATCHES_JSON_URL}?t=${Date.now()}`;
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error("HTTP " + res.status);
    const json = await res.json();
    DATA = normalizePatchesJson(json);
    rebuildPatchSheetOptions();
    rebuildPatchComponentOptions();
    renderTable();
    if (statusEl) statusEl.textContent = "";
  } catch (e) {
    const msg = "Error loading patches: " + (e?.message || e);
    if (statusEl) statusEl.textContent = msg;
    showPatchTableMessage(msg);
  }
}

function showPatchTableMessage(message) {
  const sheet = DATA?.[activeSheet] || DATA?.All_Enterprise;
  const columns = sheet?.columns?.length ? sheet.columns : ["Released", "Component", "Security", "Patch Name", "Support Page"];
  const preferredOrder = ["Released", "Component", "Security", "Patch Name", "Support Page", "Download"];
  const visibleColumns = preferredOrder.filter(c => c === "Download" || columns.includes(c));
  const thead = document.getElementById("thead");
  if (thead) {
    thead.innerHTML = "";
    for (const c of visibleColumns) {
      const th = document.createElement("th");
      th.textContent = c;
      thead.appendChild(th);
    }
  }
  const tbody = document.getElementById("tbody");
  if (!tbody) return;
  tbody.innerHTML = "";
  const tr = document.createElement("tr");
  const td = document.createElement("td");
  td.colSpan = Math.max(1, visibleColumns.length);
  td.textContent = message;
  td.className = "wrap";
  tr.appendChild(td);
  tbody.appendChild(tr);
  const verBadge = document.getElementById("verBadge");
  if (verBadge) {
    verBadge.textContent = (activeSheet === "All_Enterprise") ? "Version: All" : ("Version: " + activeSheet.replaceAll("_","."));
  }
  const countBadge = document.getElementById("countBadge");
  if (countBadge) countBadge.textContent = "Rows: 0";
}

function renderTable() {
  const sheet = DATA?.[activeSheet] || DATA?.All_Enterprise;
  if (!sheet) return;
  const {columns, rows} = sheet;
  const preferredOrder = ["Released", "Component", "Security", "Patch Name", "Support Page", "Download"];
  const visibleColumns = preferredOrder.filter(c => c === "Download" || columns.includes(c));
  const colIndex = new Map(columns.map((c,i)=>[c,i]));
  const q = document.getElementById("q").value.trim().toLowerCase();
  const comp = document.getElementById("componentSel").value.trim().toLowerCase();
  const sec = document.getElementById("securitySel").value.trim();
  const compIdx = colIndex.has("Component") ? colIndex.get("Component") : -1;
  const secIdx = colIndex.has("Security") ? colIndex.get("Security") : -1;

  const thead = document.getElementById("thead");
  if (thead) {
    thead.innerHTML = "";
    for (const c of visibleColumns) {
      const th = document.createElement("th");
      th.textContent = c;
      thead.appendChild(th);
    }
  }

  const filtered = rows.filter(r => {
    const norm = (v) => String(v || "").trim();
    const releasedIdx = colIndex.get("Released");
    const nameIdx = colIndex.get("Patch Name");
    const supportIdx = colIndex.get("Support Page");
    if (releasedIdx !== undefined && nameIdx !== undefined && supportIdx !== undefined) {
      const isHeaderRow = norm(r[releasedIdx]) === "Released" &&
        norm(r[nameIdx]) === "Patch Name" &&
        norm(r[supportIdx]) === "Support Page";
      if (isHeaderRow) return false;
    }
    if (comp && compIdx >= 0) {
      const v = String(r[compIdx]||"").toLowerCase();
      if (v !== comp) return false;
    }
    if (sec && secIdx >= 0) {
      const v = String(r[secIdx]||"");
      if (v !== sec) return false;
    }
    if (q) {
      const joined = r.map(x=>String(x||"")).join(" ").toLowerCase();
      if (!joined.includes(q)) return false;
    }
    return true;
  });

  const HARD_LIMIT = 4000;
  const renderRows = filtered.slice(0, HARD_LIMIT);
  const tbody = document.getElementById("tbody");
  tbody.innerHTML = "";

  for (const r of renderRows) {
    const tr = document.createElement("tr");
    for (const col of visibleColumns) {
      const i = colIndex.get(col);
      const td = document.createElement("td");
      const val = r[i] ?? "";
      const colLow = col.toLowerCase();

      if (colLow === "download" || colLow === "links") {
        td.innerHTML = renderPatchActionsCell(r, colIndex, columns);
      } else if (colLow.includes("download") && colLow.includes("url")) {
        td.innerHTML = renderLinksCell(val);
      } else if (colLow.includes("support")) {
        td.innerHTML = renderSupportCell(val);
      } else if (colLow === "security") {
        const v = String(val||"");
        td.textContent = v;
        td.className = (v === "Y") ? "nowrap secY" : "nowrap secN";
      } else if (colLow === "released") {
        td.textContent = String(val||"");
        td.className = "nowrap col-date";
      } else if (colLow === "component") {
        td.textContent = String(val||"");
        td.className = "nowrap col-component";
      } else if (colLow.includes("summary") || colLow.includes("patch name") || colLow.includes("name")) {
        td.textContent = String(val||"");
        td.className = "wrap col-name";
      } else {
        td.textContent = String(val||"");
        td.className = "nowrap";
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  document.getElementById("verBadge").textContent =
    (activeSheet === "All_Enterprise") ? "Version: All" : ("Version: " + activeSheet.replaceAll("_","."));
  document.getElementById("countBadge").textContent =
    "Rows: " + filtered.length.toLocaleString() + " (rendered " + renderRows.length.toLocaleString() + ")";
  document.getElementById("status").textContent =
    (filtered.length > HARD_LIMIT) ? ("Showing first " + HARD_LIMIT.toLocaleString() + " rows. Refine filters.") : "";
}

versionSel.addEventListener("change", () => {
  activeSheet = versionSel.value;
  document.getElementById("q").value = "";
  renderTable();
});
document.getElementById("q").addEventListener("input", () => renderTable());
document.getElementById("componentSel").addEventListener("change", () => renderTable());
document.getElementById("securitySel").addEventListener("change", () => renderTable());
document.getElementById("clearBtn").addEventListener("click", () => {
  document.getElementById("q").value = "";
  document.getElementById("componentSel").value = "";
  document.getElementById("securitySel").value = "";
  renderTable();
});

document.addEventListener("DOMContentLoaded", () => loadPatchesLatest(true));

// ===== Software download view (from Google Sheet CSV) =====
const SOFTWARE_CSV_URL = 'https://docs.google.com/spreadsheets/d/1XU3pUMOnohtXhxOKH0OJqe5nPWOQHwgdl6NDbhPtwD0/gviz/tq?tqx=out:csv&sheet=ArcGIS_DirectDL_Recursive';
let SOFTWARE_ROWS = [];

function parseCsv(text) {
  const rows = [];
  let row = [], cur = "", inQ = false;
  for (let i=0;i<text.length;i++) {
    const ch = text[i], nx = text[i+1];
    if (ch === '"' && inQ && nx === '"') { cur += '"'; i++; continue; }
    if (ch === '"') { inQ = !inQ; continue; }
    if (!inQ && ch === ',') { row.push(cur); cur=""; continue; }
    if (!inQ && (ch === '\n' || ch === '\r')) {
      if (ch === '\r' && nx === '\n') i++;
      row.push(cur); rows.push(row);
      row = []; cur=""; continue;
    }
    cur += ch;
  }
  if (cur.length || row.length) { row.push(cur); rows.push(row); }
  return rows.map(r => r.map(c => (c ?? '').trim()));
}

function normalizeProVersion(digits) {
  const s = String(digits || '');
  if (s.length === 2) return `${s[0]}.${s[1]}`;
  if (s.length === 3) return `${s[0]}.${s[1]}.${s[2]}`;
  if (s.length === 4) return `${s[0]}.${s[1]}.${s.slice(2)}`;
  return '';
}

function parseYearSubVersion_(s) {
  const text = String(s || '');
  let m = text.match(/(?:^|[^0-9])(20\d{2})[_\-. ](\d{1,2})(?:[_\-. ]|[^0-9]|$)/);
  if (m) return `${m[1]}.${m[2]}`;
  m = text.match(/(?:^|[^0-9])(20\d{2})(?:[^0-9]|$)/);
  if (m) return `${m[1]}`;
  return null;
}

function parseEnterprise12x_(filename, folderPath) {
  const fn = String(filename || '');
  const fp = String(folderPath || '');
  let m = fn.match(/(?:^|[^0-9])(12\d)(?:[_\-. ]|[^0-9]|$)/);
  if (!m) m = fp.match(/(?:^|[^0-9])(12\d)(?:[_\-. ]|[^0-9]|$)/);
  if (!m) return null;
  const digits = m[1];
  const num = Number(digits);
  if (num < 120 || num > 129) return null;
  return `12.${digits[2]}`;
}

function inferVersion(filename, folderPath, component) {
  const fp = (folderPath || '');
  const fn = (filename || '');
  let m = fp.match(/ArcGISPro(\d{2,4})\b/i) || fn.match(/ArcGISPro_(\d{2,4})\b/i) || fn.match(/ArcGIS_Pro_(\d{2,4})\b/i);
  if (m) {
    const v = normalizeProVersion(m[1]);
    if (v) return v;
  }

  if (component === 'License Manager' || component === 'ArcGIS Monitor' || component === 'ArcGIS Insights') {
    const ym = parseYearSubVersion_(fn) || parseYearSubVersion_(fp);
    if (ym) return ym;
  }

  const enterpriseComponents = new Set(['ArcGIS Server', 'Portal', 'Data Store', 'Notebook', 'Web Adaptor']);
  if (enterpriseComponents.has(component)) {
    const v12 = parseEnterprise12x_(fn, fp);
    if (v12) return v12;
  }

  m = fp.match(/\b11\.(\d)\b/);
  if (m) return `11.${m[1]}`;
  m = fp.match(/ArcGIS(1\d{2,4})\b/i);
  if (m) {
    const s = m[1];
    if (s.length === 3 && s.startsWith('11')) return `11.${s[2]}`;
    if (s.length === 4 && s.startsWith('10')) return `10.${s[2]}.${s[3]}`;
    if (s.length === 5 && s.startsWith('10')) return `10.${s[2]}.${s.slice(3)}`;
  }

  m = fp.match(/\b(\d{1,2}\.\d{1,2}\.\d{1,2})\b/);
  if (m) return m[1];
  m = fp.match(/\b(\d{1,2}\.\d{1,2})\b/);
  if (m) return m[1];
  return 'Unknown';
}

// Temporary checks (uncomment to verify)
// console.assert(inferVersion("ArcGIS_Server_Windows_120_111111.exe", "", "ArcGIS Server") === "12.0", "12.0 infer failed");
// console.assert(inferVersion("ArcGIS_Server_Windows_121_111111.exe", "", "ArcGIS Server") === "12.1", "12.1 infer failed");
// console.assert(inferVersion("ArcGIS_Insights_Windows_2023_1_185917.exe", "", "ArcGIS Insights") === "2023.1", "Insights year infer failed");

function inferComponent(filename, folderPath) {
  const n = (filename || '').toLowerCase();
  const fp = (folderPath || '').toLowerCase();
  if (n.includes('arcgis_insights') || n.includes('insights')) return 'ArcGIS Insights';
  if (n.includes('arcgis_monitor') || n.startsWith('arcgis_monitor') || fp.includes('arcgis_monitor') || fp.includes('arcgis monitor')) return 'ArcGIS Monitor';
  if (n.includes('license_manager') || n.includes('licensemanager') || fp.includes('license manager') || fp.includes('license_manager') || fp.includes('licensemanager')) return 'License Manager';
  if (n.includes('arcgis_server')) return 'ArcGIS Server';
  if (n.startsWith('portal')) return 'Portal';
  if (n.includes('datastore')) return 'Data Store';
  if (n.includes('notebook')) return 'Notebook';
  if (n.includes('web_adaptor') || n.includes('webadaptor') || n.includes('web adaptor')) return 'Web Adaptor';
  if (n.includes('desktop')) return 'Desktop';
  if (n.includes('arcgis_pro') || n.includes('arcgispro') || n.includes('pro_')) return 'Pro';
  if (n.includes('enterprise')) return 'Enterprise';
  if (n.includes('server')) return 'ArcGIS Server';
  return 'Other';
}

async function loadSoftwareFromSheet() {
  const swStatus = document.getElementById('swStatus');
  swStatus.textContent = 'Loading software list...';
  const res = await fetch(SOFTWARE_CSV_URL, { cache: 'no-store' });
  if (!res.ok) throw new Error('Failed to fetch CSV: HTTP ' + res.status);
  const txt = await res.text();
  const rows = parseCsv(txt);
  const header = rows[0] || [];

  const idx = {
    folder: header.indexOf('Folder Path'),
    name: header.indexOf('Filename'),
    size: header.indexOf('Size (GB)'),
    dl: header.indexOf('Direct Download'),
    view: header.indexOf('Web View'),
    updated: header.indexOf('Last Updated'),
  };
  if (idx.name < 0 || idx.dl < 0) {
    throw new Error("CSV header missing required columns: 'Filename' and 'Direct Download'");
  }

  const out = [];
  for (let i=1;i<rows.length;i++) {
    const r = rows[i];
    const filename = (r[idx.name] || '').trim();
    if (!filename) continue;
    if (!filename.startsWith('ArcGIS') && !filename.startsWith('Portal')) continue;

    const folderPath = idx.folder >= 0 ? (r[idx.folder] || '') : '';
    const component = inferComponent(filename, folderPath);
    const version = inferVersion(filename, folderPath, component) || 'Unknown';

    out.push({
      version,
      component,
      filename,
      sizeGB: idx.size >= 0 ? (r[idx.size] || '') : '',
      direct: (r[idx.dl] || ''),
      view: idx.view >= 0 ? (r[idx.view] || '') : '',
      updated: idx.updated >= 0 ? (r[idx.updated] || '') : '',
      folderPath
    });
  }

  SOFTWARE_ROWS = out;
  buildVersionDropdown();
  buildSoftwareComponentDropdown();
  renderSoftware();
  swStatus.textContent = `Loaded ${out.length} items.`;
}

function versionSortKey(v) {
  if (v === 'Unknown') return { type: 9 };
  const parts = String(v).split('.').map(p => Number(p));
  if (parts.every(p => Number.isFinite(p))) return { type: 1, parts };
  return { type: 5, text: String(v) };
}

function compareVersions(a, b) {
  const ka = versionSortKey(a);
  const kb = versionSortKey(b);
  if (ka.type !== kb.type) return ka.type - kb.type;
  if (ka.type === 1) {
    const len = Math.max(ka.parts.length, kb.parts.length);
    for (let i = 0; i < len; i++) {
      const da = ka.parts[i] ?? 0;
      const db = kb.parts[i] ?? 0;
      if (da !== db) return da - db;
    }
    return 0;
  }
  if (ka.type === 5) return ka.text.localeCompare(kb.text);
  return 0;
}

function buildVersionDropdown() {
  const sel = document.getElementById('swVersionSel');
  if (!sel) return;
  const prev = sel.value;
  const versions = Array.from(new Set(SOFTWARE_ROWS.map(x => x.version).filter(Boolean)));
  const known = versions.filter(v => v !== 'Unknown')
    .sort(compareVersions);
  const ordered = known.concat(versions.includes('Unknown') ? ['Unknown'] : []);
  sel.innerHTML = '<option value="">All versions</option>';
  for (const v of ordered) {
    const o = document.createElement('option');
    o.value = v; o.textContent = v;
    sel.appendChild(o);
  }
  sel.value = ordered.includes(prev) ? prev : '';
}

function buildSoftwareComponentDropdown() {
  const compSel = document.getElementById('swCompSel');
  if (!compSel) return;
  const prevComp = compSel.value;
  const comps = Array.from(new Set(SOFTWARE_ROWS.map(r => r.component).filter(Boolean)));
  const preferred = [
    'ArcGIS Server',
    'Portal',
    'Data Store',
    'Notebook',
    'Web Adaptor',
    'ArcGIS Monitor',
    'License Manager',
    'Pro',
    'Desktop',
    'Enterprise',
    'Other'
  ];
  const remaining = comps.filter(c => !preferred.includes(c)).sort((a,b) => a.localeCompare(b));
  const orderedComps = preferred.filter(c => comps.includes(c)).concat(remaining);
  compSel.innerHTML = '<option value="">All components</option>';
  for (const c of orderedComps) {
    const o = document.createElement('option');
    o.value = c; o.textContent = c;
    compSel.appendChild(o);
  }
  compSel.value = orderedComps.includes(prevComp) ? prevComp : '';
}

function renderSoftware() {
  const v = document.getElementById('swVersionSel').value;
  const c = document.getElementById('swCompSel').value;
  const q = (document.getElementById('swQ').value || '').toLowerCase().trim();

  let rows = SOFTWARE_ROWS.slice();
  if (v) rows = rows.filter(r => r.version === v);
  if (c) rows = rows.filter(r => r.component === c);
  if (q) rows = rows.filter(r => (r.filename || '').toLowerCase().includes(q));

  document.getElementById('swBadge').textContent = v ? `Software version: ${v}` : 'Software: All versions';
  document.getElementById('swCount').textContent = `Items: ${rows.length}`;

  const tb = document.getElementById('swTbody');
  tb.innerHTML = rows.map(r => `
    <tr>
      <td class="nowrap">${escapeHtml(r.version === 'Unknown' ? '—' : (r.version || ''))}</td>
      <td class="nowrap">${escapeHtml(r.component || '')}</td>
      <td class="wrap">${escapeHtml(r.filename || '')}</td>
      <td class="nowrap">${escapeHtml(r.sizeGB || '')}</td>
      <td>
        <div class="actions">
          ${r.view ? `<a class="pill open" href="${escapeHtml(r.view)}" target="_blank" rel="noopener noreferrer">Open</a>` : ''}
          ${r.direct ? `<a class="pill dl" href="${escapeHtml(r.direct)}" target="_blank" rel="noopener noreferrer" download>Download</a>` : ''}
        </div>
      </td>
    </tr>
  `).join('');
}

// Tabs + menu
const menuView = document.getElementById('menuView');
const menuBtn = document.getElementById('menuBtn');
const menuPatch = document.getElementById('menuPatch');
const menuSoftware = document.getElementById('menuSoftware');
const menuDownload = document.getElementById('menuDownload');

const tabPatch = document.getElementById('tabPatch');
const tabSoftware = document.getElementById('tabSoftware');
const patchView = document.getElementById('patchView');
const softwareView = document.getElementById('softwareView');

function setView(view, updateHash = true) {
  const isMenu = view === 'menu';
  const isPatch = view === 'patch';
  const isSoftware = view === 'software';

  menuView.style.display = isMenu ? '' : 'none';
  patchView.style.display = isPatch ? '' : 'none';
  softwareView.style.display = isSoftware ? '' : 'none';

  document.body.classList.toggle('view-menu', isMenu);
  document.body.classList.toggle('view-patch', isPatch);
  document.body.classList.toggle('view-software', isSoftware);
  tabPatch.classList.toggle('toggle', isPatch);
  tabSoftware.classList.toggle('toggle', isSoftware);
  tabPatch.setAttribute('aria-pressed', String(isPatch));
  tabSoftware.setAttribute('aria-pressed', String(isSoftware));

  document.body.classList.toggle('menu-active', isMenu);
  if (isPatch) loadPatchesLatest(true);
  if (updateHash) {
    const hash = isMenu ? '#menu' : (isPatch ? '#patch' : '#software');
    history.replaceState(null, '', hash);
  }
  window.scrollTo({ top: 0, behavior: 'auto' });
}

async function showSoftware() {
  setView('software');
  if (!SOFTWARE_ROWS.length) {
    try { await loadSoftwareFromSheet(); }
    catch (e) {
      document.getElementById('swStatus').textContent = 'Error: ' + (e?.message || e);
    }
  }
}

menuBtn?.addEventListener('click', () => setView('menu'));
menuPatch?.addEventListener('click', () => setView('patch'));
menuSoftware?.addEventListener('click', showSoftware);
menuDownload?.addEventListener('click', () => {
  window.location.href = './scripts.html';
});
tabPatch?.addEventListener('click', () => setView('patch'));
tabSoftware?.addEventListener('click', showSoftware);

document.getElementById('swVersionSel')?.addEventListener('change', renderSoftware);
document.getElementById('swCompSel')?.addEventListener('change', renderSoftware);
document.getElementById('swQ')?.addEventListener('input', renderSoftware);
document.getElementById('swReload')?.addEventListener('click', loadSoftwareFromSheet);

const downloadBtn = document.getElementById('downloadReg');
function applyHashView() {
  const hash = (location.hash || '').toLowerCase();
  if (hash === '#software') {
    showSoftware();
  } else if (hash === '#patch') {
    setView('patch', false);
  } else {
    setView('menu', false);
  }
}

applyHashView();
window.addEventListener('hashchange', applyHashView);

downloadBtn?.addEventListener('click', async () => {
  const url = downloadBtn.getAttribute('data-url');
  if (!url) return;
  try {
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const blob = await res.blob();
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'DisableArcGISProUpdates.reg';
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
  } catch (e) {
    window.open(url, '_blank', 'noopener,noreferrer');
  }
});
