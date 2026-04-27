/* global XLSX, math */
(function () {
const MAX_SIZE = 20 * 1024 * 1024;
const state = {
  step: 1,
  workbook: null,
  sheetName: null,
  sourceCols: [],
  rows: [],
  sheets: {}, // { sheetName: { cols: [...], rows: [...] } }
  mappings: [],
  targetFields: ['id','name','date','amount','currency','rate','converted'],
  formulas: [], // { name, expr, sheet, cols, joins: [{ sheet, joinFromCol, joinToCol, valueCol, alias }] }
  fx: [], // [{ column, year, quarter, month, resolved }]
  fxTable: {
    2023: { year: 82.5, Q1:82.0, Q2:82.3, Q3:82.7, Q4:83.0,
            Jan:82.0,Feb:82.1,Mar:82.2,Apr:82.3,May:82.3,Jun:82.4,
            Jul:82.5,Aug:82.7,Sep:82.8,Oct:83.0,Nov:83.1,Dec:83.2 },
    2024: { year: 83.4, Q1:83.0, Q2:83.3, Q3:83.6, Q4:83.8,
            Jan:82.9,Feb:83.0,Mar:83.1,Apr:83.2,May:83.3,Jun:83.4,
            Jul:83.5,Aug:83.6,Sep:83.7,Oct:83.8,Nov:83.85,Dec:83.9 },
    2025: { year: 84.0, Q1:83.9, Q2:84.0, Q3:84.1, Q4:84.2,
            Jan:83.85,Feb:83.9,Mar:83.95,Apr:84.0,May:84.0,Jun:84.0,
            Jul:84.05,Aug:84.1,Sep:84.15,Oct:84.2,Nov:84.2,Dec:84.25 },
    2026: { year: 84.5, Q1:84.3, Q2:84.5, Q3:84.6, Q4:84.7,
            Jan:84.25,Feb:84.3,Mar:84.35,Apr:84.45,May:84.5,Jun:84.55 }
  },
  processed: null,
};

function $(id){ return document.getElementById(id); }

function gotoStep(n) {
  state.step = n;
  document.querySelectorAll('[data-pane]').forEach(el => el.classList.add('d-none'));
  const pane = document.querySelector('[data-pane="'+n+'"]'); if (pane) pane.classList.remove('d-none');
  document.querySelectorAll('.step-pill').forEach(p => {
    const s = +p.dataset.step;
    p.classList.toggle('active', s === n);
    p.classList.toggle('done', s < n);
  });
  $('btnPrev').disabled = n === 1;
  $('btnNext').classList.toggle('d-none', n === 6);
  if (n === 2) renderMapping();
  if (n === 3) renderFormulas();
  if (n === 4) renderFx();
  if (n === 5) renderPreviewArea();
}

document.querySelectorAll('.step-pill').forEach(p => p.onclick = () => gotoStep(+p.dataset.step));
$('btnPrev').onclick = () => gotoStep(Math.max(1, state.step - 1));
$('btnNext').onclick = () => {
  if (state.step === 1 && !state.rows.length) { alert('Please upload a file first.'); return; }
  if (state.step === 2) {
    const missing = state.mappings.filter(m => m.mandatory && (!m.source || !m.target));
    if (missing.length) { $('mapWarning').textContent = 'Mandatory fields must be fully mapped.'; return; }
    const targets = state.mappings.map(m => m.target).filter(Boolean);
    const dupes = targets.filter((t,i) => targets.indexOf(t) !== i);
    if (dupes.length) { $('mapWarning').textContent = 'Duplicate target mapping: ' + [...new Set(dupes)].join(', '); return; }
    $('mapWarning').textContent = '';
  }
  gotoStep(Math.min(6, state.step + 1));
};

/* Upload */
const dropZone = $('dropZone');
const fileInput = $('fileInput');
['dragenter','dragover'].forEach(e => dropZone.addEventListener(e, ev => { ev.preventDefault(); dropZone.classList.add('dragover'); }));
['dragleave','drop'].forEach(e => dropZone.addEventListener(e, ev => { ev.preventDefault(); dropZone.classList.remove('dragover'); }));
dropZone.addEventListener('drop', ev => { if (ev.dataTransfer.files[0]) handleFile(ev.dataTransfer.files[0]); });
fileInput.addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });

function setStatus(html, type) {
  type = type || 'info';
  $('uploadStatus').innerHTML = '<div class="alert alert-'+type+' py-2 mb-0">'+html+'</div>';
}

function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls','csv'].includes(ext)) return setStatus('❌ Unsupported format. Use .xlsx, .xls, or .csv', 'danger');
  if (file.size === 0) return setStatus('❌ File is empty.', 'danger');
  if (file.size > MAX_SIZE) return setStatus('❌ File exceeds 20 MB limit.', 'danger');
  setStatus('⏳ Reading <strong>'+file.name+'</strong> ('+(file.size/1024).toFixed(1)+' KB)...');
  const reader = new FileReader();
  reader.onerror = () => setStatus('❌ Could not read file (corrupted?).', 'danger');
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      if (!wb.SheetNames.length) throw new Error('No sheets found');
      state.workbook = wb;
      state.sheets = {};
      const picker = $('sheetPicker');
      const sel = $('sheetSelect');
      // Find "raw" sheet (case-insensitive) and put it first
      const rawIdx = wb.SheetNames.findIndex(n => n.toLowerCase() === 'raw');
      const defaultSheet = rawIdx >= 0 ? wb.SheetNames[rawIdx] : wb.SheetNames[0];
      sel.innerHTML = wb.SheetNames.map(n => '<option '+(n===defaultSheet?'selected':'')+'>'+n+'</option>').join('');
      picker.classList.remove('d-none');
      sel.onchange = () => loadSheet(sel.value);
      // Pre-cache every sheet so the formula builder can pick from any of them
      wb.SheetNames.forEach(n => {
        const ws = wb.Sheets[n];
        const j = XLSX.utils.sheet_to_json(ws, { defval: '' });
        const cols = j.length ? Object.keys(j[0]) : [];
        state.sheets[n] = { cols, rows: j };
      });
      loadSheet(defaultSheet);
      const rawNote = rawIdx >= 0 ? ' Default sheet: <strong>'+defaultSheet+'</strong>.' : '';
      setStatus('✅ Loaded <strong>'+file.name+'</strong> — '+wb.SheetNames.length+' sheet(s).'+rawNote, 'success');
    } catch (err) {
      setStatus('❌ Corrupted or unreadable file: ' + err.message, 'danger');
    }
  };
  reader.readAsArrayBuffer(file);
}

function loadSheet(name) {
  state.sheetName = name;
  const ws = state.workbook.Sheets[name];
  const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
  if (!json.length) { setStatus('❌ Selected sheet is empty.', 'danger'); return; }
  const rawHeaders = Object.keys(json[0]);
  const seen = {}; const headers = [];
  rawHeaders.forEach(h => { if (seen[h]) { let i=2; while (seen[h+'_'+i]) i++; const n=h+'_'+i; seen[n]=1; headers.push(n);} else { seen[h]=1; headers.push(h);} });
  const renamed = headers.some((h,i) => h !== rawHeaders[i]);
  state.rows = renamed ? json.map(r => { const o={}; rawHeaders.forEach((h,i)=> o[headers[i]] = r[h]); return o; }) : json;
  state.sourceCols = headers;
  state.sheets[name] = { cols: headers, rows: state.rows };
  state.mappings = headers.slice(0, 8).map(h => ({ source: h, target: h.toLowerCase().replace(/\s+/g,'_'), mandatory: false }));
  renderFilePreview();
}

function renderFilePreview() {
  const cols = state.sourceCols;
  const sample = state.rows.slice(0, 5);
  $('filePreview').innerHTML =
    '<div class="alert alert-light border"><strong>Rows:</strong> '+state.rows.length.toLocaleString()+
    ' &nbsp;|&nbsp; <strong>Columns:</strong> '+cols.length+'</div>'+
    '<div class="table-responsive table-scroll"><table class="table table-sm table-striped">'+
    '<thead><tr>'+cols.map(c=>'<th>'+c+'</th>').join('')+'</tr></thead>'+
    '<tbody>'+sample.map(r=>'<tr>'+cols.map(c=>'<td>'+(r[c]==null?'':r[c])+'</td>').join('')+'</tr>').join('')+'</tbody>'+
    '</table></div>';
}

/* Mapping */
function renderMapping() {
  const tbody = $('mapTableBody');
  tbody.innerHTML = state.mappings.map((m, i) =>
    '<tr class="map-row">'+
      '<td><select class="form-select form-select-sm" data-i="'+i+'" data-k="source">'+
        '<option value="">— select —</option>'+
        state.sourceCols.map(c => '<option '+(c===m.source?'selected':'')+'>'+c+'</option>').join('')+
      '</select></td>'+
      '<td><input list="targetList" class="form-control form-control-sm" value="'+(m.target||'')+'" data-i="'+i+'" data-k="target" /></td>'+
      '<td class="text-center"><input type="checkbox" class="form-check-input" '+(m.mandatory?'checked':'')+' data-i="'+i+'" data-k="mandatory" /></td>'+
      '<td><button class="btn btn-sm btn-outline-danger" data-del="'+i+'"><i class="bi bi-trash"></i></button></td>'+
    '</tr>'
  ).join('') + '<datalist id="targetList">'+state.targetFields.map(t=>'<option>'+t+'</option>').join('')+'</datalist>';

  tbody.querySelectorAll('select,input').forEach(el => {
    el.onchange = () => {
      const i = +el.dataset.i, k = el.dataset.k;
      state.mappings[i][k] = el.type === 'checkbox' ? el.checked : el.value;
    };
  });
  tbody.querySelectorAll('[data-del]').forEach(b => b.onclick = () => { state.mappings.splice(+b.dataset.del,1); renderMapping(); });
}
$('btnAddMapRow').onclick = () => { state.mappings.push({source:'',target:'',mandatory:false}); renderMapping(); };
$('btnAddTarget').onclick = () => { const n = prompt('New target field name:'); if (n) { state.targetFields.push(n); renderMapping(); } };

/* ============================================================
   Formulas — with multi-sheet JOIN support
============================================================ */
const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const QUARTERS = ['Q1','Q2','Q3','Q4'];

function sheetNamesList() { return Object.keys(state.sheets); }
function colsForSheet(sheet) { return (state.sheets[sheet] && state.sheets[sheet].cols) || []; }
function defaultSheet() {
  const sheets = sheetNamesList();
  const raw = sheets.find(n => n.toLowerCase() === 'raw');
  return raw || state.sheetName || sheets[0] || '';
}

// Build a lookup index for joins: { sheetName: { joinCol: { keyValue: row } } }
function buildJoinIndex(sheet, joinCol) {
  const sd = state.sheets[sheet]; if (!sd) return {};
  const idx = {};
  sd.rows.forEach(r => {
    const k = String(r[joinCol] == null ? '' : r[joinCol]);
    if (!(k in idx)) idx[k] = r;
  });
  return idx;
}

function renderFormulas() {
  const wrap = $('formulasList');
  const sheets = sheetNamesList();
  if (!sheets.length) { wrap.innerHTML = '<div class="alert alert-warning">Upload a file first.</div>'; return; }

  wrap.innerHTML = state.formulas.map((f, i) => {
    const sheet = f.sheet || defaultSheet();
    const cols = colsForSheet(sheet);
    const selected = f.cols || [];
    const joins = f.joins || [];

    const joinsHtml = joins.map((j, ji) => {
      const jSheet = j.sheet || sheets.find(s => s !== sheet) || sheets[0];
      const jCols = colsForSheet(jSheet);
      return (
        '<div class="border rounded p-2 mb-2 bg-light">'+
          '<div class="row g-2 align-items-end">'+
            '<div class="col-md-3"><label class="form-label small mb-1">Join with sheet</label>'+
              '<select class="form-select form-select-sm" data-i="'+i+'" data-ji="'+ji+'" data-jk="sheet">'+
                sheets.filter(s => s !== sheet).map(s => '<option '+(s===jSheet?'selected':'')+'>'+s+'</option>').join('')+
              '</select></div>'+
            '<div class="col-md-3"><label class="form-label small mb-1">'+sheet+' key</label>'+
              '<select class="form-select form-select-sm" data-i="'+i+'" data-ji="'+ji+'" data-jk="joinFromCol">'+
                '<option value="">—</option>'+cols.map(c => '<option '+(c===j.joinFromCol?'selected':'')+'>'+c+'</option>').join('')+
              '</select></div>'+
            '<div class="col-md-3"><label class="form-label small mb-1">'+jSheet+' key</label>'+
              '<select class="form-select form-select-sm" data-i="'+i+'" data-ji="'+ji+'" data-jk="joinToCol">'+
                '<option value="">—</option>'+jCols.map(c => '<option '+(c===j.joinToCol?'selected':'')+'>'+c+'</option>').join('')+
              '</select></div>'+
            '<div class="col-md-2"><label class="form-label small mb-1">Pull column</label>'+
              '<select class="form-select form-select-sm" data-i="'+i+'" data-ji="'+ji+'" data-jk="valueCol">'+
                '<option value="">—</option>'+jCols.map(c => '<option '+(c===j.valueCol?'selected':'')+'>'+c+'</option>').join('')+
              '</select></div>'+
            '<div class="col-md-1 text-end"><button class="btn btn-sm btn-outline-danger" data-deljoin-i="'+i+'" data-deljoin-ji="'+ji+'"><i class="bi bi-x"></i></button></div>'+
            '<div class="col-12"><label class="form-label small mb-1">Alias (use in formula)</label>'+
              '<input class="form-control form-control-sm" placeholder="e.g. fxRate" value="'+(j.alias||'').replace(/"/g,'&quot;')+'" data-i="'+i+'" data-ji="'+ji+'" data-jk="alias" /></div>'+
          '</div>'+
        '</div>'
      );
    }).join('');

    return (
      '<div class="card mb-3 border-primary-subtle"><div class="card-body p-3">'+
        '<div class="row g-2 align-items-start">'+
          '<div class="col-md-3"><label class="form-label small mb-1">Column name</label>'+
            '<input class="form-control form-control-sm" placeholder="e.g. total" value="'+(f.name||'')+'" data-i="'+i+'" data-k="name" /></div>'+
          '<div class="col-md-3"><label class="form-label small mb-1">Source sheet</label>'+
            '<select class="form-select form-select-sm" data-i="'+i+'" data-k="sheet">'+
              sheets.map(s => '<option '+(s===sheet?'selected':'')+'>'+s+'</option>').join('')+
            '</select></div>'+
          '<div class="col-md-4"><label class="form-label small mb-1">Columns from <em>'+sheet+'</em> (Ctrl/Cmd to multi-select)</label>'+
            '<select multiple class="form-select form-select-sm" size="4" data-i="'+i+'" data-k="cols">'+
              cols.map(c => '<option '+(selected.includes(c)?'selected':'')+'>'+c+'</option>').join('')+
            '</select></div>'+
          '<div class="col-md-2 text-end"><button class="btn btn-sm btn-outline-danger mt-4" data-del="'+i+'"><i class="bi bi-trash"></i> Remove</button></div>'+

          '<div class="col-12 mt-2">'+
            '<div class="d-flex justify-content-between align-items-center mb-1">'+
              '<label class="form-label small mb-0 fw-bold">Joins (lookup from other sheets)</label>'+
              '<button class="btn btn-sm btn-outline-secondary" data-addjoin="'+i+'"><i class="bi bi-plus"></i> Add Join</button>'+
            '</div>'+
            (joinsHtml || '<div class="small text-muted">No joins. Add one to pull a column from another sheet matched by a key.</div>')+
          '</div>'+

          '<div class="col-12">'+
            '<label class="form-label small mb-1">Formula <span class="text-muted">(use column names + join aliases)</span></label>'+
            '<input class="form-control form-control-sm" placeholder="e.g. amount * fxRate" maxlength="200" value="'+(f.expr||'').replace(/"/g,'&quot;')+'" data-i="'+i+'" data-k="expr" /></div>'+
          '<div class="col-12"><span class="small" id="fxPrev_'+i+'">—</span></div>'+
        '</div>'+
      '</div></div>'
    );
  }).join('');

  // Bind formula-level fields
  wrap.querySelectorAll('[data-k]').forEach(el => {
    const handler = () => {
      const i = +el.dataset.i, k = el.dataset.k;
      if (k === 'cols') state.formulas[i].cols = Array.from(el.selectedOptions).map(o => o.value);
      else if (k === 'sheet') {
        state.formulas[i].sheet = el.value;
        state.formulas[i].cols = [];
        state.formulas[i].joins = []; // reset joins when source sheet changes
        renderFormulas(); return;
      } else state.formulas[i][k] = el.value;
      previewFormula(i);
    };
    el.onchange = handler;
    if (el.tagName === 'INPUT') el.oninput = handler;
  });
  // Bind join-level fields
  wrap.querySelectorAll('[data-jk]').forEach(el => {
    const handler = () => {
      const i = +el.dataset.i, ji = +el.dataset.ji, jk = el.dataset.jk;
      const j = state.formulas[i].joins[ji];
      j[jk] = el.value;
      if (jk === 'sheet') { j.joinToCol = ''; j.valueCol = ''; renderFormulas(); return; }
      previewFormula(i);
    };
    el.onchange = handler;
    if (el.tagName === 'INPUT') el.oninput = handler;
  });
  wrap.querySelectorAll('[data-del]').forEach(b => b.onclick = () => { state.formulas.splice(+b.dataset.del,1); renderFormulas(); });
  wrap.querySelectorAll('[data-addjoin]').forEach(b => b.onclick = () => {
    const i = +b.dataset.addjoin;
    state.formulas[i].joins = state.formulas[i].joins || [];
    state.formulas[i].joins.push({ sheet:'', joinFromCol:'', joinToCol:'', valueCol:'', alias:'' });
    renderFormulas();
  });
  wrap.querySelectorAll('[data-deljoin-i]').forEach(b => b.onclick = () => {
    const i = +b.dataset.deljoinI, ji = +b.dataset.deljoinJi;
    state.formulas[i].joins.splice(ji, 1); renderFormulas();
  });

  state.formulas.forEach((_,i)=>previewFormula(i));
}

function buildScopeForRow(f, row) {
  const sheet = f.sheet || defaultSheet();
  const sd = state.sheets[sheet];
  const useCols = (f.cols && f.cols.length) ? f.cols : (sd ? sd.cols : []);
  const scope = {};
  useCols.forEach(c => { const v = Number(row[c]); scope[c] = isNaN(v) ? 0 : v; });
  // Joins
  (f.joins || []).forEach(j => {
    if (!j.sheet || !j.joinFromCol || !j.joinToCol || !j.valueCol || !j.alias) return;
    const idx = buildJoinIndex(j.sheet, j.joinToCol);
    const key = String(row[j.joinFromCol] == null ? '' : row[j.joinFromCol]);
    const matched = idx[key];
    const v = matched ? Number(matched[j.valueCol]) : NaN;
    scope[j.alias] = isNaN(v) ? 0 : v;
  });
  return scope;
}

function previewFormula(i) {
  const f = state.formulas[i]; const out = $('fxPrev_'+i); if (!out) return;
  if (!f.expr) { out.textContent = '—'; out.className='small text-muted'; return; }
  const sheet = f.sheet || defaultSheet();
  const sd = state.sheets[sheet];
  if (!sd || !sd.rows.length) { out.textContent='⚠ No data in selected sheet'; out.className='small text-danger'; return; }
  try {
    const scope = buildScopeForRow(f, sd.rows[0]);
    const v = math.evaluate(f.expr, scope);
    out.textContent = '= ' + v + '  (vars: ' + Object.keys(scope).join(', ') + ')';
    out.className = 'small text-success';
  } catch (e) { out.textContent = '⚠ ' + e.message; out.className='small text-danger'; }
}

$('btnAddFormula').onclick = () => {
  if (state.formulas.length >= 50) return alert('Max 50 derived columns.');
  state.formulas.push({ name:'', expr:'', sheet: defaultSheet(), cols: [], joins: [] });
  renderFormulas();
};

/* Export / Import formulas as JSON */
$('btnExportFormulas').onclick = () => {
  const payload = { version: 2, exportedAt: new Date().toISOString(), formulas: state.formulas };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'derived-columns.json'; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};
$('btnImportFormulas').onclick = () => $('importFormulasInput').click();
$('importFormulasInput').onchange = (e) => {
  const file = e.target.files[0]; if (!file) return;
  const r = new FileReader();
  r.onload = ev => {
    try {
      const data = JSON.parse(ev.target.result);
      const arr = Array.isArray(data) ? data : (data.formulas || []);
      if (!Array.isArray(arr)) throw new Error('Invalid JSON: expected array of formulas');
      state.formulas = arr.map(f => ({
        name: String(f.name || ''),
        expr: String(f.expr || ''),
        sheet: f.sheet || defaultSheet(),
        cols: Array.isArray(f.cols) ? f.cols : [],
        joins: Array.isArray(f.joins) ? f.joins.map(j => ({
          sheet: j.sheet || '', joinFromCol: j.joinFromCol || '',
          joinToCol: j.joinToCol || '', valueCol: j.valueCol || '', alias: j.alias || ''
        })) : [],
      })).slice(0, 50);
      renderFormulas();
      alert('Imported '+state.formulas.length+' derived column(s).');
    } catch (err) { alert('Failed to import JSON: ' + err.message); }
    finally { e.target.value = ''; }
  };
  r.readAsText(file);
};

/* ============================================================
   FX — per-column exchange rates
============================================================ */
function resolveFxRate(item) {
  if (item.resolved != null && !isNaN(Number(item.resolved))) return Number(item.resolved);
  const y = item.year; if (!y || !state.fxTable[y]) return null;
  const t = state.fxTable[y];
  return (item.month && t[item.month]) || (item.quarter && t[item.quarter]) || t.year || null;
}
function renderFx() {
  $('fxTablePreview').textContent = JSON.stringify(state.fxTable, null, 2);
  const wrap = $('fxList');
  const sheet = defaultSheet();
  const cols = colsForSheet(sheet);
  if (!cols.length) { wrap.innerHTML = '<div class="alert alert-warning">Upload a file first.</div>'; return; }
  const years = Object.keys(state.fxTable);
  wrap.innerHTML = state.fx.map((it, i) => {
    const auto = resolveFxRate(it);
    return (
      '<div class="card mb-2"><div class="card-body p-3">'+
        '<div class="row g-2 align-items-end">'+
          '<div class="col-md-3"><label class="form-label small mb-1">Column (from <em>'+sheet+'</em>)</label>'+
            '<select class="form-select form-select-sm" data-i="'+i+'" data-k="column">'+
              '<option value="">—</option>'+cols.map(c => '<option '+(c===it.column?'selected':'')+'>'+c+'</option>').join('')+
            '</select></div>'+
          '<div class="col-md-2"><label class="form-label small mb-1">Year</label>'+
            '<select class="form-select form-select-sm" data-i="'+i+'" data-k="year">'+
              '<option value="">—</option>'+years.map(y => '<option '+(y===String(it.year)?'selected':'')+'>'+y+'</option>').join('')+
            '</select></div>'+
          '<div class="col-md-2"><label class="form-label small mb-1">Quarter</label>'+
            '<select class="form-select form-select-sm" data-i="'+i+'" data-k="quarter">'+
              '<option value="">—</option>'+QUARTERS.map(q => '<option '+(q===it.quarter?'selected':'')+'>'+q+'</option>').join('')+
            '</select></div>'+
          '<div class="col-md-2"><label class="form-label small mb-1">Month</label>'+
            '<select class="form-select form-select-sm" data-i="'+i+'" data-k="month">'+
              '<option value="">—</option>'+MONTHS.map(m => '<option '+(m===it.month?'selected':'')+'>'+m+'</option>').join('')+
            '</select></div>'+
          '<div class="col-md-2"><label class="form-label small mb-1">Override Rate</label>'+
            '<input type="number" step="0.0001" class="form-control form-control-sm" placeholder="'+(auto==null?'no rate':auto)+'" value="'+(it.resolved==null?'':it.resolved)+'" data-i="'+i+'" data-k="resolved" /></div>'+
          '<div class="col-md-1 text-end"><button class="btn btn-sm btn-outline-danger" data-delfx="'+i+'"><i class="bi bi-trash"></i></button></div>'+
          '<div class="col-12"><span class="small '+(auto==null?'text-warning':'text-success')+'">'+
            (auto==null ? '⚠ No rate matched — provide an override.' : '→ Effective rate: '+auto)+'</span></div>'+
        '</div>'+
      '</div></div>'
    );
  }).join('') || '<div class="small text-muted">No column rates configured. Click "Add Column Rate" to start.</div>';

  wrap.querySelectorAll('[data-k]').forEach(el => {
    const handler = () => {
      const i = +el.dataset.i, k = el.dataset.k;
      let v = el.value;
      if (k === 'resolved') v = v === '' ? null : Number(v);
      state.fx[i][k] = v;
      renderFx();
    };
    el.onchange = handler;
    if (el.tagName === 'INPUT') el.oninput = handler;
  });
  wrap.querySelectorAll('[data-delfx]').forEach(b => b.onclick = () => { state.fx.splice(+b.dataset.delfx,1); renderFx(); });
}
$('btnAddFx').onclick = () => {
  state.fx.push({ column:'', year:'', quarter:'', month:'', resolved:null });
  renderFx();
};

/* ============================================================
   Process
============================================================ */
$('btnProcess').onclick = function runProcessing() {
  const t0 = performance.now();
  const errors = [];
  const baseSheet = defaultSheet();
  const baseRows = (state.sheets[baseSheet] && state.sheets[baseSheet].rows) || state.rows;

  // Pre-resolve FX rates
  const fxRules = state.fx
    .filter(it => it.column)
    .map(it => ({ column: it.column, rate: resolveFxRate(it) }));

  const outRows = baseRows.map((row, idx) => {
    const out = {};
    state.mappings.forEach(m => { if (!m.target) return; out[m.target] = m.source ? row[m.source] : ''; });
    fxRules.forEach(rule => {
      const v = Number(row[rule.column]);
      if (!isNaN(v) && rule.rate != null) out[rule.column + '_converted'] = v * rule.rate;
    });
    state.formulas.forEach(f => {
      if (!f.name || !f.expr) return;
      try {
        const fSheet = f.sheet || baseSheet;
        const sd = state.sheets[fSheet];
        const fRow = (sd && sd.rows[idx]) ? sd.rows[idx] : row;
        const scope = buildScopeForRow(f, fRow);
        let val = math.evaluate(f.expr, scope);
        if (typeof val === 'number' && !isFinite(val)) { val = 'DIV/0'; throw new Error('Division by zero'); }
        out[f.name] = val;
      } catch (e) {
        out[f.name] = '#ERR';
        errors.push({ row: idx + 2, column: f.name, message: e.message });
      }
    });
    return out;
  });

  const cols = outRows.length ? Object.keys(outRows[0]) : [];
  state.processed = { rows: outRows, errors, columns: cols, ms: Math.round(performance.now() - t0) };
  renderPreviewArea();
};

function renderPreviewArea() {
  const stats = $('processStats');
  const wrap = $('previewTableWrap');
  const errWrap = $('errorTableWrap');
  if (!state.processed) {
    stats.innerHTML = '<div class="alert alert-info py-2 mb-0">Click <strong>Run Processing</strong> to apply mappings, formulas, and exchange rate.</div>';
    wrap.innerHTML = ''; errWrap.innerHTML = ''; return;
  }
  const p = state.processed;
  stats.innerHTML = '<div class="alert alert-success py-2 mb-0">✅ Processed <strong>'+p.rows.length.toLocaleString()+
    '</strong> rows in <strong>'+p.ms+' ms</strong>. Errors: <strong>'+p.errors.length+'</strong>.</div>';
  const sample = p.rows.slice(0, 50);
  wrap.innerHTML =
    '<table class="table table-sm table-striped"><thead><tr>'+p.columns.map(c=>'<th>'+c+'</th>').join('')+'</tr></thead>'+
    '<tbody>'+sample.map(r=>'<tr>'+p.columns.map(c=>'<td>'+(r[c]==null?'':r[c])+'</td>').join('')+'</tr>').join('')+'</tbody></table>'+
    '<div class="small-muted">Showing first '+sample.length+' of '+p.rows.length+' rows.</div>';
  if (p.errors.length) {
    errWrap.innerHTML = '<h6 class="mt-3">Error Log</h6><table class="table table-sm table-bordered">'+
      '<thead class="table-warning"><tr><th>Row</th><th>Column</th><th>Message</th></tr></thead><tbody>'+
      p.errors.slice(0,100).map(e=>'<tr><td>'+e.row+'</td><td>'+e.column+'</td><td>'+e.message+'</td></tr>').join('')+
      '</tbody></table>';
  } else errWrap.innerHTML = '';
}

/* Download */
$('btnDownload').onclick = () => {
  if (!state.processed) { alert('Please run processing first.'); return; }
  const p = state.processed;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(p.rows, { header: p.columns }), 'Data');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(p.errors.length ? p.errors : [{ row:'', column:'', message:'No errors' }]), 'Errors');
  const summary = [
    { metric:'Total rows', value: p.rows.length },
    { metric:'Total columns', value: p.columns.length },
    { metric:'Mapping rules', value: state.mappings.length },
    { metric:'Derived columns', value: state.formulas.length },
    { metric:'FX rules', value: state.fx.length },
    { metric:'Errors', value: p.errors.length },
    { metric:'Generated', value: new Date().toISOString() },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'Summary');
  XLSX.writeFile(wb, 'master_output.xlsx');
  $('downloadStatus').textContent = '✅ Downloaded master_output.xlsx';
};

/* Save / Load */
$('btnSaveCfg').onclick = () => {
  const name = prompt('Configuration name:'); if (!name) return;
  const all = JSON.parse(localStorage.getItem('excelToolConfigs') || '{}');
  const versions = all[name] || [];
  versions.push({
    version: versions.length + 1, savedAt: new Date().toISOString(),
    mappings: state.mappings, formulas: state.formulas, fx: state.fx, targetFields: state.targetFields,
  });
  all[name] = versions;
  localStorage.setItem('excelToolConfigs', JSON.stringify(all));
  alert('Saved "'+name+'" (v'+versions.length+').');
};
$('btnLoadCfg').onclick = () => {
  const all = JSON.parse(localStorage.getItem('excelToolConfigs') || '{}');
  const names = Object.keys(all); if (!names.length) return alert('No saved configurations.');
  const name = prompt('Load which? ' + names.join(', ')); if (!name || !all[name]) return;
  const versions = all[name];
  const v = prompt('Versions 1..'+versions.length+' (default latest):', versions.length);
  const cfg = versions[(+v || versions.length) - 1]; if (!cfg) return;
  state.mappings = cfg.mappings; state.formulas = cfg.formulas; state.fx = Array.isArray(cfg.fx) ? cfg.fx : []; state.targetFields = cfg.targetFields;
  alert('Loaded "'+name+'" v'+cfg.version+'.');
  gotoStep(2);
};

gotoStep(1);
})();
