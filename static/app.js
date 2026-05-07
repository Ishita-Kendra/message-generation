/* ── State ───────────────────────────────────────────────────────────────── */
let intFile      = null;   // File obj or null
let orgFile      = null;
let dedupRefFile = null;
let otherFiles   = [];     // File[]

// Selected reference filenames (string) — overrides uploaded file
let intRefName      = '';
let orgRefName      = '';
let dedupRefName    = '';
let othersRefNames  = new Set();

// Active tab in Reference Files section
let activeRefTab = 'reference';

/* ── Reference Files — load & render ─────────────────────────────────────── */
async function loadReferenceFiles() {
  try {
    const res  = await fetch('/api/reference-files');
    const data = await res.json();
    if (data.ok) renderReferenceFiles(data.files);
  } catch (e) {
    console.error('Could not load reference files', e);
  }
}

function fileIcon(name) {
  const ext = (name || '').split('.').pop().toLowerCase();
  if (ext === 'pdf') return 'PDF';
  if (ext === 'docx' || ext === 'doc') return 'DOC';
  return 'XLS';
}

function renderReferenceFiles(files) {
  const refFiles = files.filter(f => !f.type || f.type === 'reference');
  const outFiles = files.filter(f => f.type === 'output');

  renderFileGrid(refFiles, 'refFileGrid', 'refEmpty');
  renderFileGrid(outFiles, 'outRefFileGrid', 'outRefEmpty');
  syncRefSelects(refFiles);
}

function renderFileGrid(files, gridId, emptyId) {
  const grid  = document.getElementById(gridId);
  const empty = document.getElementById(emptyId);
  [...grid.querySelectorAll('.ref-file-card')].forEach(c => c.remove());

  if (!files.length) {
    empty.style.display = '';
    return;
  }
  empty.style.display = 'none';
  files.forEach(f => {
    const card = document.createElement('div');
    card.className = 'ref-file-card';
    card.id = `rfc-${f.name}`;
    card.innerHTML = `
      <div class="ref-file-icon">${fileIcon(f.name)}</div>
      <div class="ref-file-info">
        <div class="ref-file-name" title="${esc(f.original_name || f.name)}">${esc(f.original_name || f.name)}</div>
        <div class="ref-file-meta">${fmtSize(f.size)} · ${fmtDate(f.uploaded_at)}</div>
      </div>
      <div class="ref-file-delete" title="Delete file" onclick="deleteRefFile('${f.name}')">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
          <path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/>
        </svg>
      </div>
    `;
    grid.appendChild(card);
  });
}

/* ── Tab switching ────────────────────────────────────────────────────────── */
function switchRefTab(tab) {
  activeRefTab = tab;
  document.getElementById('tabRefFiles').classList.toggle('active', tab === 'reference');
  document.getElementById('tabOutRef').classList.toggle('active', tab === 'output');
  document.getElementById('refGlobalDrop').classList.toggle('hidden', tab !== 'reference');
  document.getElementById('outRefGlobalDrop').classList.toggle('hidden', tab !== 'output');
}

function syncRefSelects(files) {
  // Single-select dropdowns
  ['intRefSelect', 'orgRefSelect', 'dedupRefSelect'].forEach(id => {
    const sel = document.getElementById(id);
    const cur = sel.value;
    sel.innerHTML = '<option value="">— Select from Reference Files —</option>';
    files.forEach(f => {
      const opt = document.createElement('option');
      opt.value = f.name;
      opt.textContent = f.original_name || f.name;
      if (f.name === cur) opt.selected = true;
      sel.appendChild(opt);
    });
  });

  // Multi-checkbox list for dedup others
  const checklist = document.getElementById('othersRefChecklist');
  checklist.innerHTML = '';
  if (!files.length) {
    checklist.innerHTML = '<div style="font-size:11px;color:var(--muted)">No reference files yet</div>';
    return;
  }
  files.forEach(f => {
    const item = document.createElement('label');
    item.className = 'ref-check-item' + (othersRefNames.has(f.name) ? ' checked' : '');
    item.innerHTML = `
      <input type="checkbox" ${othersRefNames.has(f.name) ? 'checked' : ''}
             onchange="toggleOtherRef('${f.name}', this.checked, this.parentElement)"/>
      <span>${esc(f.original_name || f.name)}</span>
    `;
    checklist.appendChild(item);
  });
}

async function deleteRefFile(name) {
  if (!confirm('Remove this file from the reference library?')) return;
  try {
    const res  = await fetch(`/api/reference-files/${encodeURIComponent(name)}`, { method: 'DELETE' });
    const data = await res.json();
    if (data.ok) {
      // Clear any selections pointing to this file
      if (intRefName === name)      { intRefName = '';      document.getElementById('intRefSelect').value = ''; }
      if (orgRefName === name)      { orgRefName = '';      document.getElementById('orgRefSelect').value = ''; }
      if (dedupRefName === name)    { dedupRefName = '';    document.getElementById('dedupRefSelect').value = ''; }
      othersRefNames.delete(name);
      await loadReferenceFiles();
    }
  } catch (e) {
    alert('Delete failed: ' + e.message);
  }
}

async function uploadRefFiles(input) {
  if (!input.files.length) return;
  const fd = new FormData();
  Array.from(input.files).forEach(f => fd.append('files', f));
  fd.append('type', activeRefTab);
  try {
    await fetch('/api/reference-files', { method: 'POST', body: fd });
    await loadReferenceFiles();
  } catch (e) {
    alert('Upload failed: ' + e.message);
  }
  input.value = '';
}

/* ── Reference section drag-and-drop ─────────────────────────────────────── */
function onRefDragOver(e) {
  e.preventDefault();
  document.getElementById('refGlobalDrop').classList.add('drag-over');
}
function onRefDragLeave() {
  document.getElementById('refGlobalDrop').classList.remove('drag-over');
}
async function onRefDrop(e) {
  e.preventDefault();
  onRefDragLeave();
  if (!e.dataTransfer.files.length) return;
  const fd = new FormData();
  Array.from(e.dataTransfer.files).forEach(f => fd.append('files', f));
  fd.append('type', 'reference');
  try {
    await fetch('/api/reference-files', { method: 'POST', body: fd });
    await loadReferenceFiles();
  } catch (err) {
    alert('Upload failed: ' + err.message);
  }
}

function onOutRefDragOver(e) {
  e.preventDefault();
  document.getElementById('outRefGlobalDrop').classList.add('drag-over');
}
function onOutRefDragLeave() {
  document.getElementById('outRefGlobalDrop').classList.remove('drag-over');
}
async function onOutRefDrop(e) {
  e.preventDefault();
  onOutRefDragLeave();
  if (!e.dataTransfer.files.length) return;
  const fd = new FormData();
  Array.from(e.dataTransfer.files).forEach(f => fd.append('files', f));
  fd.append('type', 'output');
  try {
    await fetch('/api/reference-files', { method: 'POST', body: fd });
    await loadReferenceFiles();
    switchRefTab('output');
  } catch (err) {
    alert('Upload failed: ' + err.message);
  }
}

/* ── Section file selection helpers ──────────────────────────────────────── */
function onRefSelectChange(selectId, zoneId, metaId) {
  const val = document.getElementById(selectId).value;
  if (selectId === 'intRefSelect')      intRefName   = val;
  if (selectId === 'orgRefSelect')      orgRefName   = val;
  if (selectId === 'dedupRefSelect')    dedupRefName = val;

  const zone = document.getElementById(zoneId);
  const meta = document.getElementById(metaId);
  if (val) {
    zone.classList.add('has-file');
    meta.textContent = '✓ Using reference file';
    meta.classList.add('show');
  } else {
    zone.classList.remove('has-file');
    meta.classList.remove('show');
  }
}

function toggleOtherRef(name, checked, labelEl) {
  if (checked) {
    othersRefNames.add(name);
    labelEl.classList.add('checked');
  } else {
    othersRefNames.delete(name);
    labelEl.classList.remove('checked');
  }
}

/* ── Upload zone drag-and-drop (per-card) ────────────────────────────────── */
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
  if (file) setFile(inputId, metaId, zoneId, file);
}
function onDropMulti(e, inputId, zoneId, metaId) {
  e.preventDefault();
  onDragLeave(zoneId);
  Array.from(e.dataTransfer.files)
    .filter(f => f.name.match(/\.xlsx?$/i))
    .forEach(f => addOtherFile(f));
  renderOtherFiles();
  if (otherFiles.length) document.getElementById(zoneId).classList.add('has-file');
}
function onFileSelect(inputId, metaId, zoneId, selectId) {
  const input = document.getElementById(inputId);
  if (!input.files[0]) return;
  const file = input.files[0];
  setFile(inputId, metaId, zoneId, file);
  // Clear any reference selection for this slot
  if (selectId) {
    document.getElementById(selectId).value = '';
    if (selectId === 'intRefSelect')   intRefName   = '';
    if (selectId === 'orgRefSelect')   orgRefName   = '';
    if (selectId === 'dedupRefSelect') dedupRefName = '';
  }
}
function onMultiFileSelect(inputId, metaId, zoneId) {
  const input = document.getElementById(inputId);
  Array.from(input.files).forEach(f => addOtherFile(f));
  renderOtherFiles();
  if (otherFiles.length) document.getElementById(zoneId).classList.add('has-file');
}

function setFile(inputId, metaId, zoneId, file) {
  document.getElementById(metaId).textContent = `${file.name}  (${fmtSize(file.size)})`;
  document.getElementById(metaId).classList.add('show');
  document.getElementById(zoneId).classList.add('has-file');
  if (inputId === 'intFile')       intFile      = file;
  if (inputId === 'orgFile')       orgFile      = file;
  if (inputId === 'dedupRefFile')  dedupRefFile = file;
}

function addOtherFile(file) {
  if (!otherFiles.find(f => f.name === file.name)) otherFiles.push(file);
}
function removeOtherFile(name) {
  otherFiles = otherFiles.filter(f => f.name !== name);
  renderOtherFiles();
  if (!otherFiles.length) document.getElementById('othersDrop').classList.remove('has-file');
}
function renderOtherFiles() {
  const list = document.getElementById('othersFileList');
  list.innerHTML = otherFiles.map(f =>
    `<div class="file-pill">
      <span class="file-pill-name" title="${esc(f.name)}">${esc(f.name)}</span>
      <span class="file-pill-remove" onclick="removeOtherFile('${esc(f.name)}')">✕</span>
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

/* ── Build FormData helpers ───────────────────────────────────────────────── */
function buildMsgForm() {
  const fd = new FormData();
  if (intRefName)  fd.append('interested_ref', intRefName);
  else if (intFile) fd.append('interested', intFile);

  if (orgRefName)  fd.append('organization_ref', orgRefName);
  else if (orgFile) fd.append('organization', orgFile);
  return fd;
}

function buildDedupForm() {
  const fd = new FormData();
  if (dedupRefName)  fd.append('reference_ref', dedupRefName);
  else if (dedupRefFile) fd.append('reference', dedupRefFile);

  othersRefNames.forEach(n => fd.append('others_ref', n));
  otherFiles.forEach(f => fd.append('others', f));
  return fd;
}

/* ── Message Generator ───────────────────────────────────────────────────── */
async function runPreview() {
  const hasInt = intRefName || intFile;
  const hasOrg = orgRefName || orgFile;
  if (!hasInt || !hasOrg) {
    showStatus('msgStatus', 'Please select or upload both an Interested List and an Organization List.', 'error');
    return;
  }
  showStatus('msgStatus', 'Processing files…', 'loading');
  const btn = document.querySelector('button[onclick="runPreview()"]');
  btn.classList.add('loading');

  try {
    const res  = await fetch('/api/preview', { method: 'POST', body: buildMsgForm() });
    const data = await res.json();
    btn.classList.remove('loading');
    if (!data.ok) { showStatus('msgStatus', `Error: ${data.error}`, 'error'); return; }
    hideStatus('msgStatus');
    renderMsgPreview(data);
    document.getElementById('msgDownloadBtn').disabled = false;
    // Refresh reference files (newly uploaded files get added)
    await loadReferenceFiles();
  } catch (err) {
    btn.classList.remove('loading');
    showStatus('msgStatus', `Request failed: ${err.message}`, 'error');
  }
}

function renderMsgPreview(data) {
  document.getElementById('msgStats').innerHTML = `
    <div class="stat-chip info"><span class="sc-label">Total Rows</span><span class="sc-val">${data.total}</span></div>
    <div class="stat-chip good"><span class="sc-label">Matched</span><span class="sc-val">${data.matched}</span></div>
    <div class="stat-chip warn"><span class="sc-label">No Match</span><span class="sc-val">${data.unmatched}</span></div>
    <div class="stat-chip good"><span class="sc-label">Match Rate</span><span class="sc-val">${data.total ? Math.round(data.matched/data.total*100) : 0}%</span></div>
  `;
  document.getElementById('msgTbody').innerHTML = data.preview.map(row => {
    const badge = row.match_type === 'none'  ? '<span class="badge badge-none">—</span>'
                : row.match_type === 'exact' ? '<span class="badge badge-exact">Exact</span>'
                : `<span class="badge badge-fuzzy">${esc(row.match_type)}</span>`;
    return `<tr class="${row.has_match ? 'row-match' : ''}">
      <td>${esc(row.int_name)}</td>
      <td>${esc(row.company)}</td>
      <td>${esc(row.new_name)}</td>
      <td style="max-width:180px">${esc(row.new_title)}</td>
      <td>${badge}</td>
      <td>${esc(row.subject)}</td>
      <td style="color:var(--muted2);font-style:italic">${esc(row.body_short)}</td>
    </tr>`;
  }).join('');
  document.getElementById('msgNote').textContent = data.total > 100
    ? `Showing first 100 of ${data.total} rows. Download Excel for full messages.`
    : `Showing all ${data.total} rows. Download Excel for full messages.`;
  document.getElementById('msgPreviewArea').classList.remove('hidden');
}

async function runGenerate() {
  showStatus('msgStatus', 'Generating Excel…', 'loading');
  const btn = document.getElementById('msgDownloadBtn');
  btn.classList.add('loading');
  try {
    const res = await fetch('/api/generate', { method: 'POST', body: buildMsgForm() });
    btn.classList.remove('loading');
    if (!res.ok) {
      const d = await res.json().catch(() => ({}));
      showStatus('msgStatus', `Error: ${d.error || res.statusText}`, 'error');
      return;
    }
    hideStatus('msgStatus');
    triggerDownload(await res.blob(), 'feaam_messages.xlsx');
    showStatus('msgStatus', 'Excel downloaded.', 'success');
    await loadReferenceFiles();
  } catch (err) {
    btn.classList.remove('loading');
    showStatus('msgStatus', `Failed: ${err.message}`, 'error');
  }
}

/* ── Deduplication ───────────────────────────────────────────────────────── */
async function runDedupPreview() {
  const hasRef    = dedupRefName || dedupRefFile;
  const hasOthers = othersRefNames.size || otherFiles.length;
  if (!hasRef)    { showStatus('dedupStatus', 'Please select or upload a Reference Sheet.', 'error'); return; }
  if (!hasOthers) { showStatus('dedupStatus', 'Please select or upload at least one Other Sheet.', 'error'); return; }

  showStatus('dedupStatus', 'Checking duplicates…', 'loading');
  try {
    const res  = await fetch('/api/dedup-preview', { method: 'POST', body: buildDedupForm() });
    const data = await res.json();
    if (!data.ok) { showStatus('dedupStatus', `Error: ${data.error}`, 'error'); return; }
    hideStatus('dedupStatus');
    renderDedupPreview(data);
    document.getElementById('dedupDownloadBtn').disabled = false;
    await loadReferenceFiles();
  } catch (err) {
    showStatus('dedupStatus', `Failed: ${err.message}`, 'error');
  }
}

function renderDedupPreview(data) {
  const dupPct = data.total ? Math.round(data.duplicates / data.total * 100) : 0;
  document.getElementById('dedupStats').innerHTML = `
    <div class="stat-chip info"><span class="sc-label">Reference Rows</span><span class="sc-val">${data.total}</span></div>
    <div class="stat-chip ${data.duplicates > 0 ? 'bad' : 'good'}"><span class="sc-label">Duplicates</span><span class="sc-val">${data.duplicates}</span></div>
    <div class="stat-chip good"><span class="sc-label">Unique</span><span class="sc-val">${data.unique}</span></div>
    <div class="stat-chip ${dupPct > 20 ? 'bad' : 'warn'}"><span class="sc-label">Dup Rate</span><span class="sc-val">${dupPct}%</span></div>
  `;
  const dupArea = document.getElementById('dedupDupArea');
  if (!data.duplicates) { dupArea.classList.add('hidden'); }
  else {
    dupArea.classList.remove('hidden');
    document.getElementById('dedupTbody').innerHTML = data.preview.map(row =>
      `<tr class="row-dup">
        <td>${esc(row.name)}</td><td>${esc(row.email)}</td>
        <td><span class="badge badge-dup">${esc(row.reason)}</span></td>
      </tr>`
    ).join('');
  }
  document.getElementById('dedupPreviewArea').classList.remove('hidden');
}

async function runDedupDownload() {
  showStatus('dedupStatus', 'Generating report…', 'loading');
  const btn = document.getElementById('dedupDownloadBtn');
  btn.classList.add('loading');
  try {
    const res = await fetch('/api/dedup-download', { method: 'POST', body: buildDedupForm() });
    btn.classList.remove('loading');
    if (!res.ok) {
      const d = await res.json().catch(() => ({}));
      showStatus('dedupStatus', `Error: ${d.error || res.statusText}`, 'error');
      return;
    }
    hideStatus('dedupStatus');
    triggerDownload(await res.blob(), 'feaam_deduplication.xlsx');
    showStatus('dedupStatus', 'Report downloaded.', 'success');
    await loadReferenceFiles();
  } catch (err) {
    btn.classList.remove('loading');
    showStatus('dedupStatus', `Failed: ${err.message}`, 'error');
  }
}

/* ── Utilities ───────────────────────────────────────────────────────────── */
function showStatus(id, msg, type) {
  const el = document.getElementById(id);
  el.textContent = msg;
  el.className = `status-bar ${type}`;
}
function hideStatus(id) {
  document.getElementById(id).className = 'status-bar hidden';
}
function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}
function esc(str) {
  return String(str || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function fmtSize(bytes) {
  if (!bytes) return '0 B';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes/1024).toFixed(1) + ' KB';
  return (bytes/1048576).toFixed(1) + ' MB';
}
function fmtDate(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleDateString(undefined, {month:'short',day:'numeric',year:'numeric'}); }
  catch { return iso; }
}

/* ── Init ────────────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', loadReferenceFiles);
