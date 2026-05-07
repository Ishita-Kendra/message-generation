/* ── File state ──────────────────────────────────────────────────────────── */
let intFile    = null;
let orgFile    = null;
let refFile    = null;
let otherFiles = [];  // array of File objects

/* ── Drag-and-drop helpers ───────────────────────────────────────────────── */
function onDragOver(e, zoneId) {
  e.preventDefault();
  document.getElementById(zoneId).classList.add('drag-over');
}
function onDragLeave(zoneId) {
  document.getElementById(zoneId).classList.remove('drag-over');
}

function onDrop(e, inputId, zoneId, metaId) {
  e.preventDefault();
  onDragLeave(zoneId);
  const file = e.dataTransfer.files[0];
  if (!file) return;
  setFile(inputId, metaId, zoneId, file);
}

function onDropMulti(e, inputId, zoneId, metaId) {
  e.preventDefault();
  onDragLeave(zoneId);
  const files = Array.from(e.dataTransfer.files).filter(f => f.name.match(/\.xlsx?$/i));
  if (!files.length) return;
  files.forEach(f => addOtherFile(f));
  renderOtherFiles();
  document.getElementById(zoneId).classList.add('has-file');
}

function onFileSelect(inputId, metaId, zoneId) {
  const input = document.getElementById(inputId);
  if (!input.files[0]) return;
  setFile(inputId, metaId, zoneId, input.files[0]);
}

function onMultiFileSelect(inputId, metaId, zoneId) {
  const input = document.getElementById(inputId);
  Array.from(input.files).forEach(f => addOtherFile(f));
  renderOtherFiles();
  if (otherFiles.length) {
    document.getElementById(zoneId).classList.add('has-file');
  }
}

function setFile(inputId, metaId, zoneId, file) {
  const meta = document.getElementById(metaId);
  meta.textContent = `${file.name}  (${fmt_size(file.size)})`;
  meta.classList.add('show');
  document.getElementById(zoneId).classList.add('has-file');
  if (inputId === 'intFile')  intFile  = file;
  if (inputId === 'orgFile')  orgFile  = file;
  if (inputId === 'refFile')  refFile  = file;
}

function addOtherFile(file) {
  if (!otherFiles.find(f => f.name === file.name)) {
    otherFiles.push(file);
  }
}

function removeOtherFile(name) {
  otherFiles = otherFiles.filter(f => f.name !== name);
  renderOtherFiles();
  if (!otherFiles.length) {
    document.getElementById('othersDrop').classList.remove('has-file');
  }
}

function renderOtherFiles() {
  const list = document.getElementById('othersFileList');
  list.innerHTML = otherFiles.map(f =>
    `<div class="file-pill">
      <span class="file-pill-name" title="${f.name}">${f.name}</span>
      <span class="file-pill-remove" onclick="removeOtherFile('${f.name}')">✕</span>
    </div>`
  ).join('');
  const meta = document.getElementById('othersMeta');
  if (otherFiles.length) {
    meta.textContent = `${otherFiles.length} file${otherFiles.length > 1 ? 's' : ''} selected`;
    meta.classList.add('show');
  } else {
    meta.classList.remove('show');
  }
}

function fmt_size(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1024 / 1024).toFixed(1) + ' MB';
}

/* ── Status helpers ──────────────────────────────────────────────────────── */
function showStatus(id, msg, type) {
  const el = document.getElementById(id);
  el.textContent = msg;
  el.className = `status-bar ${type}`;
}
function hideStatus(id) {
  document.getElementById(id).className = 'status-bar hidden';
}

/* ── Message Generator ───────────────────────────────────────────────────── */
async function runPreview() {
  if (!intFile || !orgFile) {
    showStatus('msgStatus', 'Please upload both the Interested List and Organization List first.', 'error');
    return;
  }
  hideStatus('msgStatus');
  showStatus('msgStatus', 'Processing files…', 'loading');

  const btn = document.querySelector('.btn-primary');
  btn.classList.add('loading');

  const fd = new FormData();
  fd.append('interested', intFile);
  fd.append('organization', orgFile);
  fd.append('template', document.getElementById('msgTemplate').value);
  fd.append('fuzzy', document.getElementById('fuzzySlider').value);

  try {
    const res = await fetch('/api/preview', { method: 'POST', body: fd });
    const data = await res.json();
    btn.classList.remove('loading');

    if (!data.ok) {
      showStatus('msgStatus', `Error: ${data.error}`, 'error');
      return;
    }

    hideStatus('msgStatus');
    renderMsgPreview(data);
    document.getElementById('msgDownloadBtn').disabled = false;

  } catch (err) {
    btn.classList.remove('loading');
    showStatus('msgStatus', `Request failed: ${err.message}`, 'error');
  }
}

function renderMsgPreview(data) {
  // Stats
  const statsEl = document.getElementById('msgStats');
  statsEl.innerHTML = `
    <div class="stat-chip info"><span class="sc-label">Total Rows</span><span class="sc-val">${data.total}</span></div>
    <div class="stat-chip good"><span class="sc-label">Matched</span><span class="sc-val">${data.matched}</span></div>
    <div class="stat-chip warn"><span class="sc-label">No Match</span><span class="sc-val">${data.unmatched}</span></div>
    <div class="stat-chip good"><span class="sc-label">Match Rate</span><span class="sc-val">${data.total ? Math.round(data.matched/data.total*100) : 0}%</span></div>
  `;

  // Table
  const tbody = document.getElementById('msgTbody');
  tbody.innerHTML = data.preview.map(row => {
    const badge = row.match_type === 'none'  ? '<span class="badge badge-none">—</span>'
                : row.match_type === 'exact' ? '<span class="badge badge-exact">Exact</span>'
                : `<span class="badge badge-fuzzy">${row.match_type}</span>`;
    const rowCls = row.has_match ? 'row-match' : row.match_type === 'none' ? 'row-nokey' : '';
    return `<tr class="${rowCls}">
      <td>${esc(row.name)}</td>
      <td>${esc(row.company)}</td>
      <td>${esc(row.matched_company)}</td>
      <td>${badge}</td>
      <td>${esc(row.contacts)}</td>
      <td>${esc(row.message)}</td>
    </tr>`;
  }).join('');

  const note = document.getElementById('msgNote');
  note.textContent = data.total > 100
    ? `Showing first 100 of ${data.total} rows. Download Excel to see all.`
    : `Showing all ${data.total} rows.`;

  document.getElementById('msgPreviewArea').classList.remove('hidden');
}

async function runGenerate() {
  if (!intFile || !orgFile) return;

  const btn = document.getElementById('msgDownloadBtn');
  btn.classList.add('loading');
  showStatus('msgStatus', 'Generating Excel file…', 'loading');

  const fd = new FormData();
  fd.append('interested', intFile);
  fd.append('organization', orgFile);
  fd.append('template', document.getElementById('msgTemplate').value);
  fd.append('fuzzy', document.getElementById('fuzzySlider').value);

  try {
    const res = await fetch('/api/generate', { method: 'POST', body: fd });
    btn.classList.remove('loading');

    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      showStatus('msgStatus', `Error: ${data.error || res.statusText}`, 'error');
      return;
    }

    hideStatus('msgStatus');
    const blob = await res.blob();
    triggerDownload(blob, 'feaam_messages.xlsx');
    showStatus('msgStatus', 'Excel file downloaded successfully.', 'success');

  } catch (err) {
    btn.classList.remove('loading');
    showStatus('msgStatus', `Download failed: ${err.message}`, 'error');
  }
}

/* ── Deduplication ───────────────────────────────────────────────────────── */
async function runDedupPreview() {
  if (!refFile) {
    showStatus('dedupStatus', 'Please upload a Reference Sheet first.', 'error');
    return;
  }
  if (!otherFiles.length) {
    showStatus('dedupStatus', 'Please upload at least one "Other Sheet" to check against.', 'error');
    return;
  }
  hideStatus('dedupStatus');
  showStatus('dedupStatus', 'Checking for duplicates…', 'loading');

  const fd = buildDedupForm();

  try {
    const res = await fetch('/api/dedup-preview', { method: 'POST', body: fd });
    const data = await res.json();

    if (!data.ok) {
      showStatus('dedupStatus', `Error: ${data.error}`, 'error');
      return;
    }

    hideStatus('dedupStatus');
    renderDedupPreview(data);
    document.getElementById('dedupDownloadBtn').disabled = false;

  } catch (err) {
    showStatus('dedupStatus', `Request failed: ${err.message}`, 'error');
  }
}

function renderDedupPreview(data) {
  const statsEl = document.getElementById('dedupStats');
  const dupPct  = data.total ? Math.round(data.duplicates / data.total * 100) : 0;
  statsEl.innerHTML = `
    <div class="stat-chip info"><span class="sc-label">Reference Rows</span><span class="sc-val">${data.total}</span></div>
    <div class="stat-chip info"><span class="sc-label">Files Checked</span><span class="sc-val">${data.files_checked}</span></div>
    <div class="stat-chip ${data.duplicates > 0 ? 'bad' : 'good'}"><span class="sc-label">Duplicates</span><span class="sc-val">${data.duplicates}</span></div>
    <div class="stat-chip good"><span class="sc-label">Unique</span><span class="sc-val">${data.unique}</span></div>
    <div class="stat-chip ${dupPct > 20 ? 'bad' : 'warn'}"><span class="sc-label">Dup Rate</span><span class="sc-val">${dupPct}%</span></div>
  `;

  const dupArea = document.getElementById('dedupDupArea');
  if (data.duplicates === 0) {
    dupArea.classList.add('hidden');
  } else {
    dupArea.classList.remove('hidden');
    const tbody = document.getElementById('dedupTbody');
    tbody.innerHTML = data.preview.map(row =>
      `<tr class="row-dup">
        <td>${esc(row.name)}</td>
        <td>${esc(row.email)}</td>
        <td><span class="badge badge-dup">${esc(row.reason)}</span></td>
      </tr>`
    ).join('');
  }

  document.getElementById('dedupPreviewArea').classList.remove('hidden');
}

async function runDedupDownload() {
  if (!refFile || !otherFiles.length) return;

  const btn = document.getElementById('dedupDownloadBtn');
  btn.classList.add('loading');
  showStatus('dedupStatus', 'Generating deduplication report…', 'loading');

  const fd = buildDedupForm();

  try {
    const res = await fetch('/api/dedup-download', { method: 'POST', body: fd });
    btn.classList.remove('loading');

    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      showStatus('dedupStatus', `Error: ${data.error || res.statusText}`, 'error');
      return;
    }

    hideStatus('dedupStatus');
    const blob = await res.blob();
    triggerDownload(blob, 'feaam_deduplication.xlsx');
    showStatus('dedupStatus', 'Deduplication report downloaded successfully.', 'success');

  } catch (err) {
    btn.classList.remove('loading');
    showStatus('dedupStatus', `Download failed: ${err.message}`, 'error');
  }
}

function buildDedupForm() {
  const fd = new FormData();
  fd.append('reference', refFile);
  otherFiles.forEach(f => fd.append('others', f));
  return fd;
}

/* ── Utilities ───────────────────────────────────────────────────────────── */
function esc(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
