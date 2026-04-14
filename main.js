// ══════════════════════════════════════════════
// ÉTAT
// ══════════════════════════════════════════════
let allColumns        = [];   // [{ id, label, kind, choices?, refTable?, refLabelField? }]
let allRecords        = [];
let currentRecord     = null;
let pendingChanges    = {};
let pendingSavedValues = null; // valeurs sauvegardées en attente de confirmation par onRecord
let editMode          = false;
let tableId           = null;

// Données des tables référencées : { tableId: { columns: [], rows: [{id, ...fields}] } }
let refData = {};

// Layout item : { id, kind, colId?, label?, span?, refLabelField?, collapsed?, bgColor? }
let layout = [];

// Palette Notion (clés utilisées comme data-color et pour le color picker)
const NOTION_COLORS = [
  { key:'gray',   bg:'#F1F1EF' },
  { key:'brown',  bg:'#F4EEEE' },
  { key:'orange', bg:'#FAEBDD' },
  { key:'yellow', bg:'#FBF3DB' },
  { key:'green',  bg:'#EDF3EC' },
  { key:'blue',   bg:'#E7F3F8' },
  { key:'purple', bg:'#F4F0F9' },
  { key:'pink',   bg:'#FBE2E9' },
  { key:'red',    bg:'#FDEBEC' },
];
let _idCounter = 1;
function newId() { return _idCounter++; }

let dragSrcIdx  = null;
let dropTargIdx = null;

// ══════════════════════════════════════════════
// PERSISTANCE
// ══════════════════════════════════════════════
async function saveLayout() {
  try { await grist.setOption('layout', JSON.stringify(layout)); } catch(e) {
    try { localStorage.setItem('sdpc_layout_v7', JSON.stringify(layout)); } catch(_) {}
  }
}
async function loadLayout() {
  try {
    const s = await grist.getOption('layout');
    if (s) { layout = JSON.parse(s); _idCounter = Math.max(...layout.map(i=>i.id||0), 0)+1; return; }
  } catch(e) {}
  try {
    const s = localStorage.getItem('sdpc_layout_v7');
    if (s) { layout = JSON.parse(s); _idCounter = Math.max(...layout.map(i=>i.id||0), 0)+1; }
  } catch(_) {}
}

// ══════════════════════════════════════════════
// HELPERS DOM
// ══════════════════════════════════════════════
const el   = id => document.getElementById(id);
const show = id => { const e=el(id); if(e) e.style.display=''; };
const hide = id => { const e=el(id); if(e) e.style.display='none'; };

function showToast(msg, d=2500) {
  const t = el('toast');
  t.textContent = msg; t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), d);
}

function escHtml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function guessKind(val) {
  if (typeof val === 'boolean') return 'bool';
  if (typeof val === 'number')  return 'number';
  if (typeof val === 'string' && val.length > 80) return 'longtext';
  return 'text';
}

// ══════════════════════════════════════════════
// MÉTADONNÉES COLONNES GRIST
// ══════════════════════════════════════════════
async function loadColumnMeta() {
  try {
    const colMeta = await grist.docApi.fetchTable('_grist_Tables_column');
    const tblMeta = await grist.docApi.fetchTable('_grist_Tables');

    const tableRefToId = {};
    tblMeta.id.forEach((id, i) => { tableRefToId[id] = tblMeta.tableId[i]; });
    const myTableRef = tblMeta.id[tblMeta.tableId.indexOf(tableId)];

    colMeta.id.forEach((_, i) => {
      if (colMeta.parentId[i] !== myTableRef) return;
      const colId = colMeta.colId[i];
      const col   = allColumns.find(c => c.id === colId);
      if (!col) return;

      const type = colMeta.type[i] || '';
      col.gristType = type;

      if (type === 'Choice' || type === 'ChoiceList') {
        try {
          const opts = JSON.parse(colMeta.widgetOptions[i] || '{}');
          col.choices = opts.choices || [];
        } catch(e) { col.choices = []; }
        col.kind = type === 'ChoiceList' ? 'choiceList' : 'choice';

      } else if (type.startsWith('Ref:') || type.startsWith('RefList:')) {
        const refTableId = type.replace(/^Ref(List)?:/, '');
        col.refTable = refTableId;
        col.kind     = type.startsWith('RefList') ? 'refList' : 'ref';
        loadRefTable(refTableId);

      } else if (type === 'Bool')              col.kind = 'bool';
      else if (type === 'Numeric'||type==='Int') col.kind = 'number';
      else if (type === 'Text'||type==='Any')  col.kind = 'text';
    });

    renderForm();
  } catch(e) { console.warn('Métadonnées colonnes :', e); }
}

async function loadRefTable(refTableId) {
  if (refData[refTableId]) return;
  try {
    const rows = await grist.docApi.fetchTable(refTableId);
    const keys = Object.keys(rows).filter(k => k !== 'id' && !k.startsWith('_'));
    refData[refTableId] = { columns: keys, rows: rows };
    renderForm();
  } catch(e) { console.warn('Table référencée :', refTableId, e); }
}

// Récupère la valeur label d'une ligne d'une table référencée
function getRefLabel(refTableId, rowId, labelField) {
  const data = refData[refTableId];
  if (!data) return String(rowId);
  const idx = data.rows.id.indexOf(rowId);
  if (idx === -1) return String(rowId);
  const field = labelField || data.columns[0];
  return field ? String(data.rows[field]?.[idx] ?? rowId) : String(rowId);
}

// Retourne toutes les lignes d'une ref table comme [{id, label}]
function getRefRows(refTableId, labelField) {
  const data = refData[refTableId];
  if (!data) return [];
  const field = labelField || data.columns[0];
  return data.rows.id.map((id, i) => ({
    id,
    label: field ? String(data.rows[field]?.[i] ?? id) : String(id)
  })).filter(r => r.label && r.label !== 'undefined');
}

// ══════════════════════════════════════════════
// GRIST API
// ══════════════════════════════════════════════
grist.ready({ requiredAccess: 'full' });
loadLayout();

grist.onRecords(async (records) => {
  allRecords = records;
  if (!tableId) { try { tableId = await grist.getSelectedTableId(); } catch(e) {} }

  if (records.length > 0 && allColumns.length === 0) {
    const sample = records[0];
    allColumns = Object.keys(sample)
      .filter(k => k !== 'id' && !k.startsWith('_'))
      .map(k => ({ id: k, label: k.replace(/_/g,' '), kind: guessKind(sample[k]) }));
    if (layout.length === 0) {
      layout = allColumns.map(c => ({ id: newId(), kind: 'field', colId: c.id, label: c.label, span: 3 }));
      saveLayout();
    }
    if (tableId) loadColumnMeta();
  }
  populateSelect();
  hide('loading');
  if (!currentRecord) show('empty-state');
});

grist.onRecord((record) => {
  if (hasPendingChanges() && !confirm('Des modifications non enregistrées seront perdues. Continuer ?')) return;
  // Si la ligne change, les valeurs sauvegardées ne s'appliquent plus
  if (pendingSavedValues && currentRecord && record.id !== currentRecord.id) {
    pendingSavedValues = null;
  }
  // Fusionner les valeurs sauvegardées pour protéger contre un onRecord périmé de Grist
  if (pendingSavedValues) {
    record = { ...record, ...pendingSavedValues };
  }
  hide('loading'); hide('empty-state');
  el('product-form').style.display = 'block';
  currentRecord = record;
  pendingChanges = {};
  updateSaveButtons();
  renderForm();
  el('product-select').value = record.id;
});

// ══════════════════════════════════════════════
// SÉLECTEUR PRODUIT
// ══════════════════════════════════════════════
function populateSelect() {
  const sel = el('product-select'), prev = sel.value;
  sel.innerHTML = '<option value="">— Sélectionner —</option>';
  const nf = allColumns.find(c => /nom|name|titre|title|label|produit|product|designation|ref/i.test(c.id));
  allRecords.forEach(r => {
    const o = document.createElement('option');
    o.value = r.id;
    o.textContent = nf ? (r[nf.id] || `#${r.id}`) : `Enregistrement #${r.id}`;
    sel.appendChild(o);
  });
  sel.value = prev || (currentRecord ? currentRecord.id : '');
  const n = allRecords.length;
  el('product-count').textContent = `${n} produit${n>1?'s':''}`;
}
el('product-select').addEventListener('change', async (e) => {
  const id = parseInt(e.target.value);
  if (!id) { el('product-form').style.display='none'; show('empty-state'); return; }
  await grist.setCursorPos({ rowId: id });
});

// ══════════════════════════════════════════════
// MODIFICATIONS & SAUVEGARDE
// ══════════════════════════════════════════════
function hasPendingChanges() { return Object.keys(pendingChanges).length > 0; }
function markDirty(colId, value) { pendingChanges[colId] = value; if (pendingSavedValues) delete pendingSavedValues[colId]; updateSaveButtons(); }
function updateSaveButtons() {
  const d = hasPendingChanges();
  el('btn-save').classList.toggle('visible', d);
  el('btn-discard').classList.toggle('visible', d);
}
async function saveChanges() {
  if (!currentRecord || !tableId || !hasPendingChanges()) return;
  const toSave = { ...pendingChanges };
  pendingSavedValues = { ...pendingSavedValues, ...toSave }; // accumule les valeurs sauvegardées
  pendingChanges = {}; updateSaveButtons();
  try {
    await grist.docApi.applyUserActions([['UpdateRecord', tableId, currentRecord.id, toSave]]);
    Object.assign(currentRecord, toSave);
    renderForm();
    showToast('✓ Modifications enregistrées');
  } catch(e) { pendingSavedValues = null; pendingChanges = toSave; updateSaveButtons(); showToast('Erreur : '+e.message, 4000); }
}
function discardChanges() { pendingChanges = {}; updateSaveButtons(); renderForm(); }

// ══════════════════════════════════════════════
// MODE CONFIGURATION
// ══════════════════════════════════════════════
function toggleEditMode() {
  editMode = !editMode;
  el('btn-edit').classList.toggle('active', editMode);
  el('product-form').classList.toggle('edit-mode', editMode);
  el('edit-banner').classList.toggle('visible', editMode);
  el('add-bar').style.display = editMode ? 'flex' : 'none';
  renderForm();
}

// ══════════════════════════════════════════════
// RENDU DU FORMULAIRE
// ══════════════════════════════════════════════
function renderForm() {
  const form = el('product-form');
  form.innerHTML = '';
  if (editMode) form.classList.add('edit-mode');
  else form.classList.remove('edit-mode');

  if (layout.length === 0) {
    form.innerHTML = `<div style="padding:28px 20px;color:var(--text-muted);font-size:13px;text-align:center;">
      Aucun champ. Cliquez sur <strong>Configurer</strong> puis ajoutez des éléments.</div>`;
    return;
  }

  let isFirst          = true;   // pour le style du premier élément (pas de border-top)
  let grid             = null;   // la ligne <div> en cours de construction
  let usedCols         = 0;      // nb de colonnes occupées dans la ligne en cours (max 6)
  let currentContainer = form;   // où appender les items non-titre (form ou section-body)

  // Ajoute la ligne en cours au conteneur courant et remet les compteurs à zéro
  const flushGrid = () => {
    if (grid) {
      currentContainer.appendChild(grid);
      grid     = null;
      usedCols = 0;
    }
  };

  // Retourne la ligne en cours, ou en crée une nouvelle
  const getGrid = () => {
    if (!grid) {
      grid = document.createElement('div');
      grid.className = 'form-grid';
    }
    return grid;
  };

  layout.forEach((item, idx) => {

    // ── TITRE DE SECTION ──────────────────────────────────────────
    if (item.kind === 'title') {
      flushGrid();

      // Migration inline : s'assure que les nouvelles propriétés existent
      if (item.collapsed === undefined) item.collapsed = false;
      if (item.bgColor   === undefined) item.bgColor   = null;

      // Bloc section (wrapper)
      const block = document.createElement('div');
      block.className = 'section-block' + (item.bgColor ? ' has-color' : '');
      if (item.bgColor) block.dataset.color = item.bgColor;
      if (!editMode && item.collapsed) block.classList.add('collapsed');

      // En-tête
      const header = document.createElement('div');
      header.className = 'section-header' + (isFirst ? ' first-el' : '') + (!editMode ? ' toggleable' : '');
      header.dataset.idx = idx;

      // Flèche toggle
      const arrow = document.createElement('span');
      arrow.className = 'section-toggle-arrow' + (editMode || !item.collapsed ? ' expanded' : '');
      arrow.innerHTML = `<svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M2 4l4 4 4-4"/></svg>`;
      header.appendChild(arrow);

      if (editMode) {
        const inp = inlineInput('700', '13px', 'var(--text)', item.label || 'Section');
        inp.addEventListener('change', e => { layout[idx].label = e.target.value; saveLayout(); });
        header.appendChild(inp);
        header.classList.add('draggable-el');
        setupDragEvents(header, idx);
        addOverlay(header, idx);
      } else {
        const lbl = document.createElement('span');
        lbl.textContent = item.label || 'Section';
        header.appendChild(lbl);
        header.addEventListener('click', () => toggleSection(idx, block, arrow));
      }

      // Corps de section
      const body = document.createElement('div');
      body.className = 'section-body';

      block.appendChild(header);
      block.appendChild(body);
      form.appendChild(block);

      currentContainer = body;
      isFirst = false;

    // ── TEXTE DESCRIPTIF ──────────────────────────────────────────
    } else if (item.kind === 'desc') {
      flushGrid();
      const div = makeEl('div', 'form-section-desc', idx);
      if (editMode) {
        const ta = inlineTextarea('400', '12px', 'var(--text-label)', item.label || '');
        ta.addEventListener('change', e => { layout[idx].label = e.target.value; saveLayout(); });
        div.appendChild(ta);
      } else {
        div.textContent = item.label || '';
      }
      if (editMode) addOverlay(div, idx);
      currentContainer.appendChild(div);

    // ── SÉPARATEUR HORIZONTAL ─────────────────────────────────────
    } else if (item.kind === 'separator') {
      flushGrid();
      const div = makeEl('div', 'form-separator', idx);
      if (editMode) addOverlay(div, idx);
      currentContainer.appendChild(div);

    // ── CHAMP ─────────────────────────────────────────────────────
    } else if (item.kind === 'field') {
      const span = Math.min(item.span || 3, 6);

      if (usedCols + span > 6) flushGrid();

      const cell = buildFieldCell(item, idx);
      cell.classList.add(`span-${span}`);
      getGrid().appendChild(cell);
      usedCols += span;
      isFirst = false;
    }
  });

  flushGrid();
}

// Toggle collapse/expand d'une section en view mode
function toggleSection(idx, blockEl, arrowEl) {
  if (editMode) return;
  layout[idx].collapsed = !layout[idx].collapsed;
  saveLayout();
  blockEl.classList.toggle('collapsed', layout[idx].collapsed);
  arrowEl.classList.toggle('expanded',  !layout[idx].collapsed);
}

// Popover de sélection de couleur pour une section
function showColorPicker(idx, anchorBtn) {
  document.querySelectorAll('.color-picker-popover').forEach(p => p.remove());

  const item = layout[idx];
  const popover = document.createElement('div');
  popover.className = 'color-picker-popover';

  // Swatch "aucune couleur"
  const noneSwatch = document.createElement('div');
  noneSwatch.className = 'color-swatch color-swatch-none' + (!item.bgColor ? ' selected' : '');
  noneSwatch.title = 'Aucune couleur';
  noneSwatch.addEventListener('click', e => {
    e.stopPropagation();
    layout[idx].bgColor = null;
    saveLayout(); renderForm();
  });
  popover.appendChild(noneSwatch);

  // Swatches Notion
  NOTION_COLORS.forEach(({ key, bg }) => {
    const sw = document.createElement('div');
    sw.className = 'color-swatch' + (item.bgColor === key ? ' selected' : '');
    sw.style.background = bg;
    sw.title = key;
    sw.addEventListener('click', e => {
      e.stopPropagation();
      layout[idx].bgColor = key;
      saveLayout(); renderForm();
    });
    popover.appendChild(sw);
  });

  // Position fixed calée sur le bouton
  const rect = anchorBtn.getBoundingClientRect();
  popover.style.top  = (rect.bottom + 4) + 'px';
  popover.style.left = Math.max(4, rect.right - 176) + 'px';
  document.body.appendChild(popover);

  // Fermeture sur clic extérieur
  const close = e => {
    if (!popover.contains(e.target)) {
      popover.remove();
      document.removeEventListener('click', close, true);
    }
  };
  setTimeout(() => document.addEventListener('click', close, true), 0);
}

function inlineInput(weight, size, color, value) {
  const inp = document.createElement('input');
  Object.assign(inp.style, { border:'none', outline:'none', background:'transparent',
    fontWeight:weight, fontSize:size, color, width:'100%', fontFamily:'var(--font)' });
  inp.value = value; return inp;
}
function inlineTextarea(weight, size, color, value) {
  const ta = document.createElement('textarea');
  Object.assign(ta.style, { border:'none', outline:'none', background:'transparent', resize:'none',
    fontWeight:weight, fontSize:size, color, width:'100%', fontFamily:'var(--font)', lineHeight:'1.55' });
  ta.value=value; ta.rows=2; return ta;
}
function makeEl(tag, cls, idx) {
  const div = document.createElement(tag);
  div.className = cls+(editMode?' draggable-el':'');
  div.dataset.idx = idx;
  if (editMode) setupDragEvents(div, idx);
  return div;
}

// ══════════════════════════════════════════════
// CELLULE CHAMP
// ══════════════════════════════════════════════
function buildFieldCell(item, idx) {
  const col    = allColumns.find(c => c.id === item.colId) || { id: item.colId, label: item.colId, kind: 'text' };
  const rawVal = currentRecord ? currentRecord[item.colId] : '';
  const val    = pendingChanges.hasOwnProperty(item.colId) ? pendingChanges[item.colId]
                 : (rawVal !== null && rawVal !== undefined ? rawVal : '');
  const kind   = col.kind || guessKind(rawVal);

  const isEmpty = kind !== 'bool' && (
    val === null || val === undefined || val === '' ||
    (Array.isArray(val) && val.length === 0) ||
    ((kind === 'ref' || kind === 'refList') && (val === 0 || val === null))
  );

  const cell = document.createElement('div');
  cell.className = 'form-field'+(editMode?' draggable-el':'');
  if (isEmpty && !editMode) cell.classList.add('field-empty');
  cell.dataset.idx = idx;
  if (editMode) setupDragEvents(cell, idx);

  const label = document.createElement('div');
  label.className   = 'form-field-label';
  label.textContent = item.label || col.label || item.colId;
  cell.appendChild(label);

  if (editMode) {
    // Aperçu désactivé en mode config
    const preview = document.createElement('div');
    preview.style.cssText = 'font-size:13px;color:var(--text-muted);font-style:italic;padding:3px 0;';
    preview.textContent = formatValPreview(val, col, item.refLabelField);
    cell.appendChild(preview);
    addOverlay(cell, idx);
    return cell;
  }

  if      (kind === 'bool')       cell.appendChild(buildBool(col, val));
  else if (kind === 'choice')     cell.appendChild(buildChoiceSelect(col, val, v => markDirty(item.colId, v)));
  else if (kind === 'choiceList') cell.appendChild(buildChoiceList(col, val,   v => markDirty(item.colId, v)));
  else if (kind === 'ref')        cell.appendChild(buildRefSearch(col, val, false, item.refLabelField, v => markDirty(item.colId, v)));
  else if (kind === 'refList')    cell.appendChild(buildRefSearch(col, val, true,  item.refLabelField, v => markDirty(item.colId, v)));
  else if (kind === 'longtext' || (typeof rawVal==='string' && rawVal.length>80)) cell.appendChild(buildLongText(col, val));
  else if (kind === 'number')     cell.appendChild(buildNumber(col, val));
  else                            cell.appendChild(buildText(col, val));

  return cell;
}

function formatValPreview(val, col, labelField) {
  if (val===null||val===undefined||val==='') return 'Non renseigné';
  if (col.kind==='ref')     return getRefLabel(col.refTable, val, labelField);
  if (col.kind==='refList') return (Array.isArray(val)?val.filter(id=>typeof id==='number'&&id>0):[val]).map(id=>getRefLabel(col.refTable,id,labelField)).join(', ');
  if (Array.isArray(val))   return val.join(', ');
  return String(val);
}

// ══════════════════════════════════════════════
// WIDGETS PAR TYPE
// ══════════════════════════════════════════════

function buildBool(col, val) {
  const wrap = document.createElement('div');
  wrap.className = 'form-field-check-wrap';
  const chk = document.createElement('input');
  chk.type='checkbox'; chk.className='grist-check'; chk.checked=!!val;
  chk.addEventListener('change', () => markDirty(col.id, chk.checked));
  wrap.appendChild(chk); return wrap;
}

function buildText(col, val) {
  const inp = document.createElement('input');
  inp.type='text'; inp.className='form-field-input'; inp.value=val||''; inp.placeholder='Non renseigné';
  inp.addEventListener('input', () => markDirty(col.id, inp.value));
  return inp;
}

function buildNumber(col, val) {
  const inp = document.createElement('input');
  inp.type='number'; inp.className='form-field-input'; inp.value=val||''; inp.placeholder='Non renseigné';
  inp.addEventListener('input', () => markDirty(col.id, inp.value===''?null:parseFloat(inp.value)));
  return inp;
}

function buildLongText(col, val) {
  const ta = document.createElement('textarea');
  ta.className='form-field-textarea'; ta.value=val||''; ta.rows=3; ta.placeholder='Non renseigné';
  ta.addEventListener('input', () => {
    ta.style.height='auto'; ta.style.height=ta.scrollHeight+'px'; markDirty(col.id, ta.value);
  });
  setTimeout(() => { ta.style.height='auto'; ta.style.height=ta.scrollHeight+'px'; }, 0);
  return ta;
}

// ── Choice (simple) ──
function buildChoiceSelect(col, currentVal, onChange) {
  const sel = document.createElement('select');
  sel.className = 'form-field-input';
  sel.style.cssText = 'cursor:pointer;appearance:auto;padding:2px 0;';
  const empty = document.createElement('option');
  empty.value=''; empty.textContent='Non renseigné'; sel.appendChild(empty);
  (col.choices||[]).forEach(c => {
    const o = document.createElement('option');
    o.value=c; o.textContent=c; if(c===currentVal) o.selected=true; sel.appendChild(o);
  });
  if (!currentVal) sel.value='';
  sel.addEventListener('change', () => onChange(sel.value||null));
  return sel;
}

// ── ChoiceList (multi) — tags + bouton "+" ──
function buildChoiceList(col, currentVal, onChange) {
  let selected = Array.isArray(currentVal)
    ? currentVal.filter(v => typeof v === 'string' && v !== 'L')
    : (currentVal && typeof currentVal === 'string')
      ? currentVal.split(',').map(s => s.trim()).filter(Boolean)
      : [];

  const wrap = document.createElement('div');
  wrap.style.cssText = 'display:flex;flex-wrap:wrap;align-items:center;gap:4px;padding:3px 0;position:relative;';

  const dropdown = document.createElement('div');
  dropdown.style.cssText = `
    display:none;position:absolute;top:calc(100% + 4px);left:0;z-index:150;min-width:160px;
    background:var(--surface);border:1px solid var(--border);border-radius:var(--r);
    box-shadow:0 4px 16px rgba(0,0,0,0.12);padding:4px;
  `;
  // Le dropdown est ajouté UNE SEULE FOIS au wrap, ici, et n'en bouge plus jamais
  wrap.appendChild(dropdown);

  const renderAll = () => {
    // On vide wrap SAUF le dropdown
    Array.from(wrap.children).forEach(c => { if (c !== dropdown) c.remove(); });

    // Tags des valeurs sélectionnées — ajoutés avant le dropdown via appendChild
    // mais comme dropdown est déjà dans wrap, ils se placent avant lui
    selected.forEach((s, i) => {
      const tag = makePill(s, '#e8f0fe', '#1a73e8', () => {
        selected.splice(i, 1); renderAll(); onChange(selected.length ? ['L', ...selected] : null);
      });
      // On insère avant le dropdown qui est déjà dans le wrap
      wrap.insertBefore(tag, dropdown);
    });

    // Bouton "+"
    const remaining = (col.choices || []).filter(c => !selected.includes(c));
    if (remaining.length > 0) {
      const addBtn = document.createElement('button');
      addBtn.style.cssText = `
        width:20px;height:20px;border-radius:50%;border:1.5px solid var(--accent);
        background:var(--accent-light);color:var(--accent);font-size:14px;line-height:1;
        cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;
        font-family:var(--font);padding:0;transition:background .1s;
      `;
      addBtn.textContent = '+';
      addBtn.title = 'Ajouter une valeur';
      addBtn.addEventListener('mouseenter', () => { addBtn.style.background = 'var(--accent)'; addBtn.style.color = '#fff'; });
      addBtn.addEventListener('mouseleave', () => { addBtn.style.background = 'var(--accent-light)'; addBtn.style.color = 'var(--accent)'; });

      // Contenu du dropdown
      dropdown.innerHTML = '';
      remaining.forEach(choice => {
        const item = document.createElement('div');
        item.style.cssText = 'padding:5px 8px;cursor:pointer;font-size:12px;border-radius:2px;transition:background .1s;';
        item.textContent = choice;
        item.addEventListener('mouseenter', () => item.style.background = 'var(--bg)');
        item.addEventListener('mouseleave', () => item.style.background = '');
        item.addEventListener('click', e => {
          e.stopPropagation();
          selected.push(choice);
          renderAll();
          onChange(selected.length ? ['L', ...selected] : null);
          dropdown.style.display = 'none';
        });
        dropdown.appendChild(item);
      });

      addBtn.addEventListener('click', e => {
        e.stopPropagation();
        dropdown.style.display = dropdown.style.display === 'block' ? 'none' : 'block';
      });

      // Insère le bouton "+" avant le dropdown
      wrap.insertBefore(addBtn, dropdown);
    }
  };

  // Ferme le dropdown si clic en dehors
  document.addEventListener('click', (e) => {
    if (!wrap.contains(e.target)) dropdown.style.display = 'none';
  });

  renderAll();
  return wrap;
}

// ── Référence avec recherche ──
function buildRefSearch(col, currentVal, isMulti, labelField, onChange) {
  // Pour RefList, Grist retourne ["L", id1, id2, ...] où "L" est un marqueur de type.
  // Ce marqueur peut être la chaîne "L", null, ou un objet selon la version.
  // On garde uniquement les entiers positifs valides.
  const toRefIds = (val) => {
    if (!Array.isArray(val)) return (val && typeof val === 'number' && val > 0) ? [val] : [];
    return val.filter(id => typeof id === 'number' && id > 0);
  };
  let selected = isMulti
    ? toRefIds(currentVal)
    : (currentVal ? [currentVal] : []);

  const wrap = document.createElement('div');
  wrap.style.cssText = 'display:flex;flex-direction:column;gap:4px;position:relative;';

  // Zone de tags
  const tagsRow = document.createElement('div');
  tagsRow.style.cssText = 'display:flex;flex-wrap:wrap;gap:3px;min-height:20px;';

  const renderTags = () => {
    tagsRow.innerHTML = '';
    selected.forEach((id, i) => {
      const lbl = getRefLabel(col.refTable, id, labelField);
      const tag = makePill(lbl, '#e8f0fe', '#1a73e8', () => {
        selected.splice(i,1); renderTags();
        onChange(isMulti ? ['L', ...selected] : (selected[0]??null));
      });
      tagsRow.appendChild(tag);
    });
  };
  renderTags();

  // Barre de recherche
  const searchWrap = document.createElement('div');
  searchWrap.style.cssText = 'position:relative;';

  const searchInp = document.createElement('input');
  const rows = getRefRows(col.refTable, labelField);
  searchInp.type='text'; searchInp.className='form-field-input';
  searchInp.placeholder = rows.length ? 'Rechercher…' : 'Chargement…';

  const dropdown = document.createElement('div');
  dropdown.style.cssText = `
    display:none;position:absolute;top:100%;left:0;right:0;z-index:150;
    background:var(--surface);border:1px solid var(--border);border-radius:var(--r);
    box-shadow:0 4px 16px rgba(0,0,0,0.12);max-height:200px;overflow-y:auto;
  `;

  const renderDropdown = (query) => {
    dropdown.innerHTML = '';
    const currentRows = getRefRows(col.refTable, labelField);
    const q = query.toLowerCase();
    const filtered = currentRows.filter(r => r.label.toLowerCase().includes(q) && !selected.includes(r.id)).slice(0,50);

    if (!filtered.length) {
      const empty = document.createElement('div');
      empty.style.cssText='padding:10px 12px;color:var(--text-muted);font-size:12px;';
      empty.textContent = query ? 'Aucun résultat' : (currentRows.length ? 'Toutes les entrées sont déjà sélectionnées' : 'Chargement…');
      dropdown.appendChild(empty); return;
    }

    filtered.forEach(row => {
      const item = document.createElement('div');
      item.style.cssText='padding:7px 12px;cursor:pointer;font-size:13px;transition:background .1s;';
      if (query) {
        const qi = row.label.toLowerCase().indexOf(q);
        item.innerHTML = qi>=0
          ? escHtml(row.label.slice(0,qi))+`<strong style="color:var(--accent)">${escHtml(row.label.slice(qi,qi+q.length))}</strong>`+escHtml(row.label.slice(qi+q.length))
          : escHtml(row.label);
      } else { item.textContent = row.label; }
      item.addEventListener('mouseenter', () => item.style.background='var(--bg)');
      item.addEventListener('mouseleave', () => item.style.background='');
      item.addEventListener('mousedown', e => {
        e.preventDefault();
        if (isMulti) selected.push(row.id); else selected=[row.id];
        renderTags(); onChange(isMulti ? ['L', ...selected] : selected[0]);
        searchInp.value=''; dropdown.style.display='none';
      });
      dropdown.appendChild(item);
    });
  };

  searchInp.addEventListener('focus',  () => { renderDropdown(searchInp.value); dropdown.style.display='block'; });
  searchInp.addEventListener('input',  () => { renderDropdown(searchInp.value); dropdown.style.display='block'; });
  searchInp.addEventListener('blur',   () => { setTimeout(() => dropdown.style.display='none', 150); });

  searchWrap.appendChild(searchInp);
  searchWrap.appendChild(dropdown);
  wrap.appendChild(tagsRow);
  wrap.appendChild(searchWrap);
  return wrap;
}

// Fabrique une pastille-tag avec croix
function makePill(text, bg, color, onRemove) {
  const tag = document.createElement('span');
  tag.style.cssText = `display:inline-flex;align-items:center;gap:4px;padding:2px 8px 2px 9px;
    border-radius:10px;background:${bg};color:${color};font-size:11px;font-weight:600;`;
  const lbl = document.createElement('span');
  lbl.textContent = text;
  const rm = document.createElement('span');
  rm.textContent='×';
  rm.style.cssText='cursor:pointer;font-size:13px;line-height:1;opacity:.7;';
  rm.addEventListener('mouseenter', () => rm.style.opacity='1');
  rm.addEventListener('mouseleave', () => rm.style.opacity='.7');
  rm.addEventListener('click', e => { e.stopPropagation(); onRemove(); });
  tag.appendChild(lbl); tag.appendChild(rm);
  return tag;
}

// ══════════════════════════════════════════════
// OVERLAY + TOOLBAR (mode configuration)
// ══════════════════════════════════════════════
function addOverlay(div, idx) {
  const item = layout[idx];

  const ov = document.createElement('div');
  ov.className='el-overlay'; div.appendChild(ov);

  const tb = document.createElement('div');
  tb.className='el-toolbar';

  // Poignée
  const handle = document.createElement('div');
  handle.className='drag-handle';
  handle.innerHTML=`<svg width="10" height="14" viewBox="0 0 10 16" fill="currentColor">
    <circle cx="3" cy="3" r="1.2"/><circle cx="7" cy="3" r="1.2"/>
    <circle cx="3" cy="8" r="1.2"/><circle cx="7" cy="8" r="1.2"/>
    <circle cx="3" cy="13" r="1.2"/><circle cx="7" cy="13" r="1.2"/>
  </svg>`;
  tb.appendChild(handle);

  if (item.kind === 'title') {
    // Bouton color picker pour les titres de section
    const colorBtn = tbBtn('🎨 Couleur', false);
    colorBtn.addEventListener('click', e => {
      e.stopPropagation();
      showColorPicker(idx, colorBtn);
    });
    tb.appendChild(colorBtn);

  } else if (item.kind === 'field') {
    const col = allColumns.find(c => c.id === item.colId);

    // Boutons largeur
    [{l:'1/3',s:2},{l:'1/2',s:3},{l:'2/3',s:4},{l:'Plein',s:6}].forEach(opt => {
      const btn = tbBtn(opt.l, item.span===opt.s);
      btn.addEventListener('click', e => { e.stopPropagation(); layout[idx].span=opt.s; saveLayout(); renderForm(); });
      tb.appendChild(btn);
    });

    // Bouton champ de recherche (Ref / RefList uniquement)
    if (col && (col.kind==='ref'||col.kind==='refList') && col.refTable && refData[col.refTable]) {
      const pickBtn = tbBtn('🔍 Champ de recherche', false);
      pickBtn.addEventListener('click', e => {
        e.stopPropagation();
        showRefFieldPicker(idx, col.refTable, item.refLabelField);
      });
      tb.appendChild(pickBtn);
    }

    // Renommer
    const renBtn = tbBtn('✏ Renommer', false);
    renBtn.addEventListener('click', e => {
      e.stopPropagation();
      const v = prompt('Nom affiché :', item.label||item.colId);
      if (v!==null) { layout[idx].label=v; saveLayout(); renderForm(); }
    });
    tb.appendChild(renBtn);
  }
  // desc et separator : uniquement drag handle + supprimer

  // Supprimer
  const delBtn = tbBtn('✕ Supprimer', false, true);
  delBtn.addEventListener('click', e => { e.stopPropagation(); layout.splice(idx,1); saveLayout(); renderForm(); });
  tb.appendChild(delBtn);

  div.appendChild(tb);
}

function tbBtn(label, active, danger=false) {
  const btn = document.createElement('button');
  btn.className = 'el-tb-btn'+(active?' w-active':'')+(danger?' danger':'');
  btn.textContent = label;
  btn.addEventListener('mousedown', e => e.stopPropagation());
  return btn;
}

// ── Mini-modal pour choisir le champ de recherche d'une référence ──
function showRefFieldPicker(layoutIdx, refTableId, currentField) {
  const data = refData[refTableId];
  if (!data) return;

  // Supprimer l'ancien picker s'il existe
  const old = document.getElementById('ref-field-picker');
  if (old) old.remove();

  const modal = document.createElement('div');
  modal.id = 'ref-field-picker';
  modal.style.cssText = `
    position:fixed;inset:0;z-index:400;background:rgba(0,0,0,0.22);
    display:flex;align-items:center;justify-content:center;
  `;

  const box = document.createElement('div');
  box.style.cssText = `
    background:var(--surface);border:1px solid var(--border);border-radius:4px;
    box-shadow:0 4px 20px rgba(0,0,0,0.14);width:320px;overflow:hidden;
  `;

  const header = document.createElement('div');
  header.style.cssText='height:42px;padding:0 14px;border-bottom:1px solid var(--border);font-size:13px;font-weight:600;display:flex;align-items:center;justify-content:space-between;';
  header.innerHTML = `<span>Champ d'affichage / recherche</span>`;
  const close = document.createElement('button');
  close.textContent='×'; close.style.cssText='border:none;background:none;cursor:pointer;font-size:18px;color:var(--text-muted);';
  close.addEventListener('click', () => modal.remove());
  header.appendChild(close);
  box.appendChild(header);

  const sub = document.createElement('div');
  sub.style.cssText='padding:8px 14px 4px;font-size:11px;color:var(--text-muted);';
  sub.textContent = `Table : ${refTableId} — choisissez la colonne utilisée pour afficher et rechercher les entrées`;
  box.appendChild(sub);

  const list = document.createElement('div');
  list.style.cssText='padding:6px;max-height:280px;overflow-y:auto;';

  data.columns.forEach(colId => {
    const btn = document.createElement('button');
    btn.style.cssText='display:flex;align-items:center;gap:8px;width:100%;padding:7px 10px;border:none;background:none;font-family:var(--font);font-size:13px;color:var(--text);cursor:pointer;border-radius:var(--r);text-align:left;transition:background .1s;';
    btn.innerHTML = `
      <span style="flex:1">${escHtml(colId.replace(/_/g,' '))}</span>
      ${colId===currentField ? `<span style="color:var(--accent);font-size:11px;font-weight:700;">✓ Actif</span>` : ''}
    `;
    btn.addEventListener('mouseenter', () => btn.style.background='var(--bg)');
    btn.addEventListener('mouseleave', () => btn.style.background='');
    btn.addEventListener('click', () => {
      layout[layoutIdx].refLabelField = colId;
      saveLayout(); renderForm(); modal.remove();
    });
    list.appendChild(btn);
  });

  box.appendChild(list);
  modal.appendChild(box);
  modal.addEventListener('click', e => { if(e.target===modal) modal.remove(); });
  document.body.appendChild(modal);
}

// ══════════════════════════════════════════════
// DRAG & DROP INLINE
// ══════════════════════════════════════════════
function setupDragEvents(div, idx) {
  div.draggable = true;
  div.addEventListener('dragstart', e => {
    dragSrcIdx=idx; e.dataTransfer.effectAllowed='move';
    e.dataTransfer.setData('text/plain', String(idx));
    setTimeout(() => div.classList.add('dragging'), 0);
  });
  div.addEventListener('dragend', () => {
    div.classList.remove('dragging');
    document.querySelectorAll('.draggable-el').forEach(d => d.classList.remove('drag-over'));
    if (dragSrcIdx!==null && dropTargIdx!==null && dragSrcIdx!==dropTargIdx) {
      const moved = layout.splice(dragSrcIdx,1)[0];
      const insertAt = dropTargIdx>dragSrcIdx ? dropTargIdx-1 : dropTargIdx;
      layout.splice(insertAt,0,moved);
      saveLayout(); renderForm();
    }
    dragSrcIdx=null; dropTargIdx=null;
  });
  div.addEventListener('dragover', e => {
    e.preventDefault(); e.dataTransfer.dropEffect='move';
    document.querySelectorAll('.draggable-el').forEach(d => d.classList.remove('drag-over'));
    div.classList.add('drag-over'); dropTargIdx=idx;
  });
  div.addEventListener('dragleave', () => div.classList.remove('drag-over'));
  div.addEventListener('drop', e => { e.preventDefault(); div.classList.remove('drag-over'); });
}

// ══════════════════════════════════════════════
// AJOUTER DES ÉLÉMENTS
// ══════════════════════════════════════════════
function addItem(kind) {
  const defaults = { title:'Nouvelle section', desc:'Description…', separator:'' };
  const item = { id:newId(), kind, label:defaults[kind]??'', span:kind==='field'?3:undefined };
  if (kind === 'title') { item.collapsed = false; item.bgColor = null; }
  layout.push(item);
  saveLayout(); renderForm();
}

// ══════════════════════════════════════════════
// PICKER DE CHAMPS
// ══════════════════════════════════════════════
function openFieldPicker() {
  el('field-picker').classList.add('open');
  el('picker-search-input').value='';
  renderPickerList('');
  el('picker-search-input').focus();
}
function closeFieldPicker() { el('field-picker').classList.remove('open'); }
el('field-picker').addEventListener('click', e => { if(e.target===el('field-picker')) closeFieldPicker(); });
function filterPickerFields() { renderPickerList(el('picker-search-input').value.toLowerCase()); }

function renderPickerList(query) {
  const list=el('picker-list'); list.innerHTML='';
  const used = new Set(layout.filter(i=>i.kind==='field').map(i=>i.colId));
  const TYPE_LABEL = { choice:'Choice', choiceList:'ChoiceList', ref:'Réf.', refList:'Réf.+', bool:'Bool', number:'Num.', longtext:'Long', text:'Texte' };
  allColumns
    .filter(c => !query || c.id.toLowerCase().includes(query) || c.label.toLowerCase().includes(query))
    .forEach(col => {
      const btn = document.createElement('button');
      btn.className='picker-field-btn'+(used.has(col.id)?' already-added':'');
      btn.innerHTML=`<span style="flex:1">${escHtml(col.label||col.id)}</span><span class="field-type-tag">${TYPE_LABEL[col.kind]||'Texte'}</span>`;
      if (!used.has(col.id)) btn.addEventListener('click', () => {
        layout.push({ id:newId(), kind:'field', colId:col.id, label:col.label||col.id, span:3 });
        saveLayout(); renderForm(); renderPickerList(query);
      });
      list.appendChild(btn);
    });
  if (!list.children.length)
    list.innerHTML=`<div style="text-align:center;padding:20px;color:var(--text-muted);font-size:12px;">Aucun champ trouvé</div>`;
}