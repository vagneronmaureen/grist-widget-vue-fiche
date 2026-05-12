// ══════════════════════════════════════════════
// ÉTAT
// ══════════════════════════════════════════════
let allColumns        = [];   // [{ id, label, kind, choices?, refTable?, refLabelField? }]
let allRecords        = [];
let currentRecord     = null;
let pendingChanges    = {};
let pendingSavedValues = null; // valeurs sauvegardées en attente de confirmation par onRecord
let editMode          = false;  // mode configuration layout (owners uniquement)
let dataEditMode      = false;  // mode édition des données (editors+)
let selectionMode     = 'internal'; // 'internal' | 'linked'
let userAccess        = null;   // 'owners' | 'editors' | 'viewers' | null | 'unchecked'
// 'unchecked' = vérification en cours, boutons masqués par défaut
let canWriteTable     = false;  // accès en écriture à la table courante (vérifié via ACL Grist)
let tableWriteChecked = null;   // tableId pour laquelle l'accès a déjà été testé
let tableId           = null;

// Données des tables référencées : { tableId: { columns: [], rows: [{id, ...fields}] } }
let refData = {};

// Valeurs brutes de la table courante : { recordId: { colId: rawValue } }
// fetchTable() retourne les vrais IDs pour Ref/RefList, contrairement à onRecord
// qui peut envoyer les valeurs d'affichage (ex: "Jean Dupont") ou null.
let rawRecords = {};

// Layout item : { id, kind, colId?, label?, span?, refLabelField?, collapsed?, bgColor? }
let layout = [];

// Tailles de police pour titres et descriptions
const TITLE_SIZES = { s:'11px', m:'13px', l:'16px', xl:'20px' };
const DESC_SIZES  = { s:'10px', m:'12px', l:'14px', xl:'18px' };

// Palette DSFR — tons pastels conformes au Système de Design de l'État
// Tokens officiels DSFR v1
// bg = tinte 975 (très clair), bgDark = tinte 950 (en-tête), border = Sun (couleur principale)
const DSFR_COLORS = [
  { key:'gray',   bg:'#f6f6f6', bgDark:'#eeeeee', border:'#3a3a3a', label:'Gris'           },
  { key:'blue',   bg:'#f5f5fe', bgDark:'#ececfe', border:'#000091', label:'Bleu France'    },
  { key:'green',  bg:'#e3fdeb', bgDark:'#d6fce2', border:'#18753c', label:'Vert émeraude'  },
  { key:'teal',   bg:'#e8fdff', bgDark:'#c7faf5', border:'#009099', label:'Écume'          },
  { key:'red',    bg:'#fff0f0', bgDark:'#fee9e9', border:'#c9191e', label:'Rouge Marianne' },
  { key:'orange', bg:'#fff3de', bgDark:'#ffdeb8', border:'#b34000', label:'Orange'         },
  { key:'yellow', bg:'#fef7da', bgDark:'#fef7ba', border:'#716043', label:'Jaune'          },
  { key:'purple', bg:'#fee7fc', bgDark:'#fcd5f9', border:'#6e445a', label:'Pourpre'        },
  { key:'pink',   bg:'#ffe9e6', bgDark:'#ffdbd1', border:'#8d533e', label:'Rose'           },
];

// Convertit un hex en rgba(r,g,b,alpha)
function hexToRgba(hex, alpha) {
  const r = parseInt(hex.slice(1,3), 16);
  const g = parseInt(hex.slice(3,5), 16);
  const b = parseInt(hex.slice(5,7), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

// Calcule la position left d'une popup fixed pour qu'elle reste dans le viewport.
// Si elle déborderait à droite, on l'ancre depuis la droite du bouton déclencheur.
function safePopLeft(rect, popWidth) {
  const vw = document.documentElement.clientWidth;
  return (rect.left + popWidth + 4 > vw)
    ? Math.max(4, rect.right - popWidth)
    : rect.left;
}
let _idCounter = 1;
function newId() { return _idCounter++; }

let dragSrcIdx  = null;
let dropTargIdx = null;

// Palette de couleurs pour les tags ref/refList
const TAG_COLORS = [
  { key: 'blue',   bg: '#e8f0fe', text: '#1a73e8', label: 'Bleu'   },
  { key: 'green',  bg: '#e6f4ea', text: '#137333', label: 'Vert'   },
  { key: 'purple', bg: '#f3e8fd', text: '#7c3aed', label: 'Violet' },
  { key: 'teal',   bg: '#e6f4f1', text: '#0d7377', label: 'Cyan'   },
  { key: 'orange', bg: '#fef3e2', text: '#b45309', label: 'Orange' },
  { key: 'red',    bg: '#fce8e6', text: '#c5221f', label: 'Rouge'  },
  { key: 'pink',   bg: '#fce8f3', text: '#b4276f', label: 'Rose'   },
  { key: 'gray',   bg: '#f1f3f4', text: '#5f6368', label: 'Gris'   },
];

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
// ACCÈS UTILISATEUR
// ══════════════════════════════════════════════
// 'unchecked' → vérification en cours, on masque par défaut
// null → API indisponible (hors Grist / dev) → on laisse tout visible
function isOwner()  { return userAccess === null || userAccess === 'owners'; }
function canWrite() { return isOwner() || userAccess === 'editors'; }

async function checkUserAccess() {
  userAccess = 'unchecked'; // masque les boutons pendant la vérification
  updateAccessUI();
  try {
    const session = await grist.getUserSession();
    // session.access : 'owners' | 'editors' | 'viewers' | undefined
    userAccess = session ? (session.access || null) : null;
  } catch(e) {
    userAccess = null; // API indisponible → mode dev, tout visible
  }
  updateAccessUI();
}

// Teste l'accès en écriture sur la table via un UpdateRecord à champs vides.
// C'est un no-op (rien n'est modifié) mais les règles ACL avancées de Grist sont
// vérifiées, ce qui permet de détecter si l'utilisateur peut écrire dans cette table.
async function checkTableWriteAccess() {
  if (!tableId || allRecords.length === 0) return;
  if (tableWriteChecked === tableId) return; // déjà vérifié pour cette table
  tableWriteChecked = tableId;
  try {
    await grist.docApi.applyUserActions([['UpdateRecord', tableId, allRecords[0].id, {}]]);
    canWriteTable = true;
  } catch(e) {
    canWriteTable = false;
  }
  updateAccessUI();
  updateTopbarButtons();
}

function updateAccessUI() {
  const btnEdit     = el('btn-edit');
  const btnDataEdit = el('btn-data-edit');
  if (btnEdit)     btnEdit.style.display     = isOwner()      ? '' : 'none';
  if (btnDataEdit) btnDataEdit.style.display = canWriteTable  ? '' : 'none';
  updateLinkBtn();
}

// ── Mode édition des données ──
function toggleDataEditMode() {
  if (!canWriteTable) return; // droits insuffisants
  if (dataEditMode) {
    if (hasPendingChanges() && !confirm('Des modifications non enregistrées seront perdues. Continuer ?')) return;
    dataEditMode = false;
    pendingChanges = {};
  } else {
    dataEditMode = true;
  }
  updateSaveButtons();
  updateTopbarButtons();
  renderForm();
}

function updateTopbarButtons() {
  const btnDataEdit = el('btn-data-edit');
  if (!btnDataEdit) return;
  if (editMode) {
    btnDataEdit.style.display = 'none';
    return;
  }
  btnDataEdit.style.display = canWriteTable ? '' : 'none';
  if (dataEditMode) {
    btnDataEdit.innerHTML = `<svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor"><path d="M13.78 4.22a.75.75 0 010 1.06l-7.25 7.25a.75.75 0 01-1.06 0L2.22 9.28a.75.75 0 011.06-1.06L6 10.94l6.72-6.72a.75.75 0 011.06 0z"/></svg> Enregistrer`;
    btnDataEdit.classList.add('active');
  } else {
    btnDataEdit.innerHTML = `<svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor"><path d="M11.013 1.427a1.75 1.75 0 012.474 0l1.086 1.086a1.75 1.75 0 010 2.474l-8.61 8.61c-.21.21-.47.364-.756.445l-3.251.93a.75.75 0 01-.927-.928l.929-3.25c.081-.286.235-.547.445-.758l8.61-8.61zm1.414 1.06a.25.25 0 00-.354 0L10.811 3.75l1.439 1.44 1.263-1.263a.25.25 0 000-.354l-1.086-1.086zM11.189 6.25L9.75 4.81l-6.286 6.287a.25.25 0 00-.064.108l-.558 1.953 1.953-.558a.25.25 0 00.108-.065L11.19 6.25z"/></svg> Éditer`;
    btnDataEdit.classList.remove('active');
  }
}

// ══════════════════════════════════════════════
// MODE SÉLECTION : interne vs vue liée
// ══════════════════════════════════════════════

// Charge le mode depuis les options du widget (persisté dans Grist)
async function loadSelectionMode() {
  try {
    const saved = await grist.getOption('selectionMode');
    if (saved === 'linked' || saved === 'internal') selectionMode = saved;
  } catch(e) {}
  syncTopbarZone();
  updateLinkBtn();
}

// Bascule entre les deux modes et persiste le choix (owners en mode config uniquement)
function toggleSelectionMode() {
  if (!isOwner() || !editMode) return;
  selectionMode = selectionMode === 'internal' ? 'linked' : 'internal';
  try { grist.setOption('selectionMode', selectionMode); } catch(e) {}
  syncTopbarZone();
  updateLinkBtn();
  updateEmptyStateMsg();
}

// Synchronise quelle zone du topbar est visible
function syncTopbarZone() {
  const inEdit   = editMode;
  const isLinked = selectionMode === 'linked';
  el('topbar-product').style.display      = (!inEdit && !isLinked) ? 'flex' : 'none';
  el('topbar-linked').style.display       = (!inEdit &&  isLinked) ? 'flex' : 'none';
  el('topbar-edit-actions').style.display = inEdit ? 'flex' : 'none';
}

// Met à jour l'apparence du bouton Vue liée / Interne
function updateLinkBtn() {
  const btn = el('btn-link-mode');
  if (!btn) return;
  const linked = selectionMode === 'linked';
  btn.style.display = (editMode && isOwner()) ? '' : 'none';
  if (linked) {
    btn.innerHTML = `<svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor"><path d="M7.775 3.275a.75.75 0 001.06 1.06l1.25-1.25a2 2 0 112.83 2.83l-2.5 2.5a2 2 0 01-2.83 0 .75.75 0 00-1.06 1.06 3.5 3.5 0 004.95 0l2.5-2.5a3.5 3.5 0 00-4.95-4.95l-1.25 1.25zm-4.69 9.64a2 2 0 010-2.83l2.5-2.5a2 2 0 012.83 0 .75.75 0 001.06-1.06 3.5 3.5 0 00-4.95 0l-2.5 2.5a3.5 3.5 0 004.95 4.95l1.25-1.25a.75.75 0 00-1.06-1.06l-1.25 1.25a2 2 0 01-2.83 0z"/></svg> ↩ Interne`;
    btn.classList.add('linked-active');
  } else {
    btn.innerHTML = `<svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor"><path d="M7.775 3.275a.75.75 0 001.06 1.06l1.25-1.25a2 2 0 112.83 2.83l-2.5 2.5a2 2 0 01-2.83 0 .75.75 0 00-1.06 1.06 3.5 3.5 0 004.95 0l2.5-2.5a3.5 3.5 0 00-4.95-4.95l-1.25 1.25zm-4.69 9.64a2 2 0 010-2.83l2.5-2.5a2 2 0 012.83 0 .75.75 0 001.06-1.06 3.5 3.5 0 00-4.95 0l-2.5 2.5a3.5 3.5 0 004.95 4.95l1.25-1.25a.75.75 0 00-1.06-1.06l-1.25 1.25a2 2 0 01-2.83 0z"/></svg> Vue liée`;
    btn.classList.remove('linked-active');
  }
}

// Met à jour le nom de l'enregistrement affiché dans la zone Vue liée
function updateLinkedBadge() {
  const span = el('linked-record-name');
  if (!span) return;
  if (currentRecord) {
    const p = _productLabels.find(p => p.id === currentRecord.id);
    span.textContent = p ? p.label : getProductLabel(currentRecord);
    span.style.fontStyle = 'normal';
    span.style.color = 'var(--text)';
  } else {
    span.textContent = '— sélectionnez une ligne —';
    span.style.fontStyle = 'italic';
    span.style.color = 'var(--text-muted)';
  }
}

// Adapte le message et l'icône de l'état vide selon le mode
function updateEmptyStateMsg() {
  const msg  = el('empty-state-msg');
  const icoI = el('empty-icon-internal');
  const icoL = el('empty-icon-linked');
  if (!msg) return;
  const linked = selectionMode === 'linked';
  msg.textContent = linked
    ? 'Sélectionnez une ligne dans la vue liée'
    : 'Sélectionnez un produit';
  if (icoI) icoI.style.display = linked ? 'none' : '';
  if (icoL) icoL.style.display = linked ? ''     : 'none';
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

    // Debug : clés disponibles dans _grist_Tables_column
    console.log('[Widget] _grist_Tables_column keys:', Object.keys(colMeta));

    colMeta.id.forEach((_, i) => {
      if (colMeta.parentId[i] !== myTableRef) return;
      const colId = colMeta.colId[i];
      const col   = allColumns.find(c => c.id === colId);
      if (!col) return;

      // Colonne formule pure (isFormula = true) → non modifiable dans le widget
      col.isFormula = !!(colMeta.isFormula && colMeta.isFormula[i]);

      // Description de la colonne (renseignée dans Grist)
      // Le champ peut s'appeler 'description' ou ne pas être retourné si vide
      const rawDesc = colMeta.description ? colMeta.description[i] : undefined;
      col.description = (rawDesc && typeof rawDesc === 'string') ? rawDesc.trim() : '';
      if (col.description) console.log(`[Widget] description colonne "${colId}":`, col.description);

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
      else if (type === 'Date')                col.kind = 'date';
      else if (type === 'DateTime')            col.kind = 'datetime';
      else if (type === 'Attachments')         col.kind = 'attachment';
      else if (type === 'Text'||type==='Any')  col.kind = 'text';
    });

    renderForm();
    loadRawRecords(); // charge les vrais IDs pour Ref/RefList en arrière-plan
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

// Charge les valeurs brutes de la table courante via fetchTable.
// Contrairement à onRecord (qui envoie les valeurs d'affichage pour Ref/RefList),
// fetchTable retourne les vrais IDs stockés : ex. ['L', 5, 7] pour un RefList.
async function loadRawRecords() {
  if (!tableId) return;
  try {
    const rows = await grist.docApi.fetchTable(tableId);
    rawRecords = {};
    rows.id.forEach((id, i) => {
      rawRecords[id] = {};
      Object.keys(rows).forEach(key => {
        if (key !== 'id') rawRecords[id][key] = rows[key][i];
      });
    });
    // Si un record est déjà affiché, mettre à jour ses champs ref/refList et re-rendre
    if (currentRecord && rawRecords[currentRecord.id]) {
      const raw = rawRecords[currentRecord.id];
      let changed = false;
      allColumns.forEach(col => {
        if (col.kind !== 'ref' && col.kind !== 'refList') return;
        if (pendingSavedValues && col.id in pendingSavedValues) return;
        const rawVal = raw[col.id];
        if (rawVal !== null && rawVal !== undefined) {
          currentRecord[col.id] = rawVal;
          changed = true;
        }
      });
      if (changed) renderForm();
    }
  } catch(e) { console.warn('loadRawRecords :', e); }
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
checkUserAccess();
loadSelectionMode();

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
  // Vérifier l'accès en écriture dès qu'on a la table et au moins un enregistrement
  if (tableId && records.length > 0) checkTableWriteAccess();
  hide('loading');
  if (!currentRecord) show('empty-state');
});

grist.onRecord((record) => {
  // En mode vue liée, Grist peut envoyer null quand aucune ligne n'est sélectionnée
  if (!record || !record.id) {
    currentRecord = null;
    el('product-form').style.display = 'none';
    updateEmptyStateMsg();
    show('empty-state');
    if (selectionMode === 'linked') updateLinkedBadge();
    return;
  }

  if (hasPendingChanges() && !confirm('Des modifications non enregistrées seront perdues. Continuer ?')) return;
  const prev = currentRecord;
  // Si la ligne change, les valeurs sauvegardées ne s'appliquent plus
  if (pendingSavedValues && prev && record.id !== prev.id) {
    pendingSavedValues = null;
  }
  // Fusionner les valeurs sauvegardées pour protéger contre un onRecord périmé de Grist
  if (pendingSavedValues) {
    record = { ...record, ...pendingSavedValues };
  }
  // Grist envoie les valeurs d'affichage (ex: "Jean Dupont") dans onRecord pour
  // les champs Ref/RefList, pas les vrais IDs. On patch avec rawRecords si disponible
  // (rawRecords vient de fetchTable qui retourne les vraies valeurs brutes).
  if (rawRecords[record.id]) {
    const raw = rawRecords[record.id];
    const patched = { ...record };
    allColumns.forEach(col => {
      if (col.kind !== 'ref' && col.kind !== 'refList') return;
      if (pendingSavedValues && col.id in pendingSavedValues) return;
      const rawVal = raw[col.id];
      if (rawVal !== null && rawVal !== undefined) {
        patched[col.id] = rawVal;
      }
    });
    record = patched;
  }
  hide('loading'); hide('empty-state');
  el('product-form').style.display = 'block';
  currentRecord = record;
  pendingChanges = {};
  updateSaveButtons();
  renderForm();
  // Met à jour l'input ou le badge selon le mode
  if (selectionMode === 'linked') {
    updateLinkedBadge();
  } else {
    const _p = _productLabels.find(p => p.id === record.id);
    el('product-search-input').value = _p ? _p.label : getProductLabel(record);
  }
});

// ══════════════════════════════════════════════
// SÉLECTEUR PRODUIT (recherche avec dropdown)
// ══════════════════════════════════════════════
let _productLabels = []; // cache [{id, label}]

function getProductLabel(r) {
  const nf = allColumns.find(c => /nom|name|titre|title|label|produit|product|designation|ref/i.test(c.id));
  return nf ? (r[nf.id] || `#${r.id}`) : `Enregistrement #${r.id}`;
}

function populateSelect() {
  _productLabels = allRecords.map(r => ({ id: r.id, label: getProductLabel(r) }));
  const n = allRecords.length;
  el('product-count').textContent = `${n} produit${n>1?'s':''}`;
  // Met à jour l'input si un record est sélectionné
  if (currentRecord) {
    const p = _productLabels.find(p => p.id === currentRecord.id);
    el('product-search-input').value = p ? p.label : '';
  }
}

function setupProductSearch() {
  const input = el('product-search-input');
  const dropdown = el('product-dropdown');

  const renderDropdown = (query) => {
    dropdown.innerHTML = '';
    const q = (query || '').toLowerCase();
    const locale = { sensitivity: 'base' };
    const cmp = (a, b) => a.label.localeCompare(b.label, 'fr', locale);
    let items;
    if (!q) {
      items = [..._productLabels].sort(cmp);
    } else {
      // Groupe 1 : commence par la recherche — Groupe 2 : contient la recherche ailleurs
      const starts   = _productLabels.filter(p =>  p.label.toLowerCase().startsWith(q)).sort(cmp);
      const contains = _productLabels.filter(p => !p.label.toLowerCase().startsWith(q) && p.label.toLowerCase().includes(q)).sort(cmp);
      items = [...starts, ...contains];
    }

    if (!items.length) {
      const empty = document.createElement('div');
      empty.style.cssText = 'padding:10px 12px;color:var(--text-muted);font-size:12px;text-align:center;';
      empty.textContent = q ? 'Aucun résultat' : 'Aucun enregistrement';
      dropdown.appendChild(empty);
    } else {
      items.forEach(p => {
        const isCurrent = currentRecord && currentRecord.id === p.id;
        const row = document.createElement('div');
        row.style.cssText = `padding:7px 12px;cursor:pointer;font-size:13px;transition:background .1s;
          ${isCurrent ? 'background:var(--accent-light);color:var(--accent);font-weight:600;' : ''}`;
        // Surlignage de la recherche
        if (q) {
          const qi = p.label.toLowerCase().indexOf(q);
          row.innerHTML = qi >= 0
            ? escHtml(p.label.slice(0,qi)) + `<strong style="color:var(--accent)">${escHtml(p.label.slice(qi,qi+q.length))}</strong>` + escHtml(p.label.slice(qi+q.length))
            : escHtml(p.label);
        } else {
          row.textContent = p.label;
        }
        row.addEventListener('mouseenter', () => { if (!isCurrent) row.style.background = 'var(--bg)'; });
        row.addEventListener('mouseleave', () => { if (!isCurrent) row.style.background = ''; });
        row.addEventListener('mousedown', async e => {
          e.preventDefault();
          dropdown.style.display = 'none';
          if (!p.id) { el('product-form').style.display='none'; show('empty-state'); return; }
          await grist.setCursorPos({ rowId: p.id });
        });
        dropdown.appendChild(row);
      });
    }
    dropdown.style.display = 'block';
  };

  input.addEventListener('focus', () => {
    input.value = ''; // vide l'entrée pour montrer toute la liste
    renderDropdown('');
  });
  input.addEventListener('input', () => renderDropdown(input.value));
  input.addEventListener('blur', () => {
    setTimeout(() => {
      dropdown.style.display = 'none';
      // Restaure le label du record courant si aucune sélection n'a été faite
      if (currentRecord) {
        const p = _productLabels.find(p => p.id === currentRecord.id);
        input.value = p ? p.label : '';
      } else {
        input.value = '';
      }
    }, 160);
  });
  input.addEventListener('keydown', e => {
    if (e.key === 'Escape') { dropdown.style.display = 'none'; input.blur(); }
  });
  document.addEventListener('click', e => {
    if (!el('product-search-wrap').contains(e.target)) dropdown.style.display = 'none';
  });
}

// ══════════════════════════════════════════════
// MODIFICATIONS & SAUVEGARDE
// ══════════════════════════════════════════════
function hasPendingChanges() { return Object.keys(pendingChanges).length > 0; }
function markDirty(colId, value) { pendingChanges[colId] = value; if (pendingSavedValues) delete pendingSavedValues[colId]; updateSaveButtons(); }
function updateSaveButtons() {
  // Les boutons Annuler / Enregistrer ne s'affichent que si :
  // - on est en mode édition données
  // - il y a des changements en attente
  // - l'utilisateur a les droits d'écriture sur la table
  const show = hasPendingChanges() && dataEditMode && canWriteTable;
  el('btn-save').classList.toggle('visible', show);
  el('btn-discard').classList.toggle('visible', show);

  // Quand Annuler/Enregistrer sont visibles, masquer btn-data-edit pour éviter le doublon
  const btnDataEdit = el('btn-data-edit');
  if (btnDataEdit && dataEditMode && !editMode) {
    btnDataEdit.style.display = show ? 'none' : (canWriteTable ? '' : 'none');
  }
}

async function saveChanges() {
  if (!currentRecord || !tableId || !hasPendingChanges() || !canWriteTable) return;
  const toSave = { ...pendingChanges };
  pendingSavedValues = { ...pendingSavedValues, ...toSave };
  pendingChanges = {};
  try {
    await grist.docApi.applyUserActions([['UpdateRecord', tableId, currentRecord.id, toSave]]);
    Object.assign(currentRecord, toSave);
    // Sauvegarde réussie → sortir du mode édition
    dataEditMode = false;
    updateSaveButtons();
    updateTopbarButtons();
    renderForm();
    showToast('✓ Modifications enregistrées');
  } catch(e) {
    pendingSavedValues = null;
    pendingChanges = toSave;
    updateSaveButtons();
    showToast('Erreur : ' + e.message, 4000);
  }
}

function discardChanges() {
  pendingChanges = {};
  // Annuler → sortir du mode édition également
  dataEditMode = false;
  updateSaveButtons();
  updateTopbarButtons();
  renderForm();
}

// Garde : empêche d'entrer en mode édition sans les droits
function toggleDataEditModeGuard() {
  if (!canWriteTable) return;
  toggleDataEditMode();
}

// ══════════════════════════════════════════════
// MODE CONFIGURATION
// ══════════════════════════════════════════════
function toggleEditMode() {
  editMode = !editMode;
  if (editMode) dataEditMode = false; // config mode désactive l'édition données
  el('product-form').classList.toggle('edit-mode', editMode);
  el('edit-banner').classList.toggle('visible', editMode);

  // Topbar : zones produit/liée/config + bouton vue liée
  syncTopbarZone();
  updateLinkBtn();

  // Bouton : "Configurer" ↔ "Sauvegarder" (vert)
  const btn = el('btn-edit');
  if (editMode) {
    btn.innerHTML = `<svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor"><path d="M13.78 4.22a.75.75 0 010 1.06l-7.25 7.25a.75.75 0 01-1.06 0L2.22 9.28a.75.75 0 011.06-1.06L6 10.94l6.72-6.72a.75.75 0 011.06 0z"/></svg> Sauvegarder`;
    btn.classList.add('config-mode');
  } else {
    btn.innerHTML = `<svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor"><path d="M11.013 1.427a1.75 1.75 0 012.474 0l1.086 1.086a1.75 1.75 0 010 2.474l-8.61 8.61c-.21.21-.47.364-.756.445l-3.251.93a.75.75 0 01-.927-.928l.929-3.25c.081-.286.235-.547.445-.758l8.61-8.61zm1.414 1.06a.25.25 0 00-.354 0L10.811 3.75l1.439 1.44 1.263-1.263a.25.25 0 000-.354l-1.086-1.086zM11.189 6.25L9.75 4.81l-6.286 6.287a.25.25 0 00-.064.108l-.558 1.953 1.953-.558a.25.25 0 00.108-.065L11.19 6.25z"/></svg> Configurer`;
    btn.classList.remove('config-mode');
  }
  updateTopbarButtons();
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

  let isFirst          = true;
  let grid             = null;
  let usedCols         = 0;
  let currentContainer = form;
  let halfWidthBlocks  = [];    // sections demi-largeur en attente d'un wrapper

  const flushGrid = () => {
    if (grid) { currentContainer.appendChild(grid); grid = null; usedCols = 0; }
  };
  const getGrid = () => {
    if (!grid) { grid = document.createElement('div'); grid.className = 'form-grid'; }
    return grid;
  };
  const flushHalfWidths = () => {
    if (halfWidthBlocks.length) {
      const wrapper = document.createElement('div');
      wrapper.className = 'section-columns';
      halfWidthBlocks.forEach(b => wrapper.appendChild(b));
      form.appendChild(wrapper);
      halfWidthBlocks = [];
    }
  };

  layout.forEach((item, idx) => {

    // ── TITRE DE SECTION ──────────────────────────────────────────
    if (item.kind === 'title') {
      flushGrid();
      if (!item.halfWidth) flushHalfWidths();

      // Migration inline
      if (item.collapsed    === undefined) item.collapsed    = false;
      if (item.bgColor      === undefined) item.bgColor      = null;
      if (item.sectionStyle === undefined) item.sectionStyle = 'fill';
      if (item.noToggle     === undefined) item.noToggle     = false;

      // Couleur DSFR associée
      const dsfrColor = DSFR_COLORS.find(c => c.key === item.bgColor);

      // Bloc section
      const block = document.createElement('div');
      const hasColor = item.bgColor && item.sectionStyle !== 'none';
      block.className = 'section-block' + (hasColor ? ' has-color' : '')
                       + (item.sectionStyle === 'border' ? ' style-contour' : '');
      if (hasColor && dsfrColor) {
        if (item.sectionStyle === 'border') {
          block.style.borderLeftColor = dsfrColor.border;
        }
        // 'fill' : pas de fond sur le bloc lui-même,
        // l'en-tête et le corps ont chacun leur propre fond
      }
      if (!editMode && item.collapsed && !item.noToggle) block.classList.add('collapsed');

      // En-tête
      const canToggle = !editMode && !item.noToggle;
      const header = document.createElement('div');
      header.className = 'section-header' + (isFirst ? ' first-el' : '') + (canToggle ? ' toggleable' : '');
      header.dataset.idx = idx;
      if (hasColor && dsfrColor && item.sectionStyle === 'fill') {
        header.style.background = dsfrColor.bg;
      }

      // Flèche toggle (cachée si noToggle)
      const arrow = document.createElement('span');
      arrow.className = 'section-toggle-arrow' + (editMode || !item.collapsed || item.noToggle ? ' expanded' : '');
      if (item.noToggle) arrow.style.display = 'none';
      arrow.innerHTML = `<svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M2 4l4 4 4-4"/></svg>`;
      header.appendChild(arrow);

      header.style.fontSize = TITLE_SIZES[item.fontSize] || TITLE_SIZES.m;

      if (editMode) {
        const ce = document.createElement('div');
        ce.contentEditable = 'true';
        ce.innerHTML = item.label || 'Section';
        ce.dataset.layoutIdx = idx;
        ce.style.cssText = 'outline:none;flex:1;min-width:0;';
        ce.addEventListener('keydown', e => { if (e.key === 'Enter') e.preventDefault(); });
        ce.addEventListener('input', () => { layout[idx].label = ce.innerHTML; saveLayout(); });
        header.appendChild(ce);
        header.classList.add('draggable-el');
        setupDragEvents(header, idx);
        addOverlay(header, idx);
      } else {
        const lbl = document.createElement('span');
        lbl.innerHTML = item.label || 'Section';
        header.appendChild(lbl);
        if (canToggle) header.addEventListener('click', () => toggleSection(idx, block, arrow));
      }

      const body = document.createElement('div');
      body.className = 'section-body';
      // Fond des champs à 5 % d'opacité de la couleur principale
      if (hasColor && dsfrColor && item.sectionStyle === 'fill') {
        body.style.background = hexToRgba(dsfrColor.border, 0.05);
      }
      block.appendChild(header);
      block.appendChild(body);

      // Demi-largeur ou pleine largeur
      if (item.halfWidth) {
        halfWidthBlocks.push(block);
      } else {
        form.appendChild(block);
      }

      currentContainer = body;
      isFirst = false;

    // ── TEXTE DESCRIPTIF ──────────────────────────────────────────
    } else if (item.kind === 'desc') {
      flushGrid();
      const div = makeEl('div', 'form-section-desc', idx);
      const descSize = DESC_SIZES[item.fontSize] || DESC_SIZES.m;
      div.style.fontSize = descSize;
      if (editMode) {
        const ce = document.createElement('div');
        ce.contentEditable = 'true';
        ce.innerHTML = item.label || '';
        ce.dataset.layoutIdx = idx;
        ce.style.cssText = 'outline:none;width:100%;min-height:1.4em;';
        ce.addEventListener('input', () => { layout[idx].label = ce.innerHTML; saveLayout(); });
        div.appendChild(ce);
      } else {
        div.innerHTML = item.label || '';
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
  flushHalfWidths();
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

  // Swatches DSFR
  DSFR_COLORS.forEach(({ key, bg, border, label }) => {
    const sw = document.createElement('div');
    sw.className = 'color-swatch' + (item.bgColor === key ? ' selected' : '');
    // Fond + liseré de la couleur d'accentuation
    sw.style.cssText = `background:${bg};outline:2px solid ${border};outline-offset:-4px;`;
    sw.title = label;
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

  // Pour refList, Grist stocke un champ vide comme ['L'] (marqueur seul, sans IDs).
  // On filtre les IDs numériques > 0 pour détecter correctement l'état vide.
  const refIds = (kind === 'refList' || kind === 'ref')
    ? (Array.isArray(val) ? val.filter(v => typeof v === 'number' && v > 0) : (val && typeof val === 'number' && val > 0 ? [val] : []))
    : null;

  const isEmpty = kind !== 'bool' && (
    val === null || val === undefined || val === '' ||
    (refIds !== null ? refIds.length === 0 : (Array.isArray(val) && val.length === 0))
  );

  const cell = document.createElement('div');
  cell.className = 'form-field'+(editMode?' draggable-el':'');
  if (isEmpty && !editMode) cell.classList.add('field-empty');
  cell.dataset.idx = idx;
  if (editMode) setupDragEvents(cell, idx);

  const label = document.createElement('div');
  label.className = 'form-field-label';
  if (item.emoji) {
    const ico = document.createElement('span');
    ico.textContent = item.emoji;
    ico.style.cssText = 'margin-right:5px;font-style:normal;';
    label.appendChild(ico);
  }
  label.appendChild(document.createTextNode(item.label || col.label || item.colId));

  // Icône ⓘ si la colonne a une description renseignée dans Grist
  if (col.description) {
    const infoBtn = document.createElement('i');
    infoBtn.className = 'field-info-btn';
    infoBtn.textContent = 'i';
    infoBtn.setAttribute('aria-label', col.description);
    const desc = col.description;
    infoBtn.addEventListener('mouseenter', () => {
      let tip = document.getElementById('field-info-tooltip');
      if (!tip) { tip = document.createElement('div'); tip.id = 'field-info-tooltip'; document.body.appendChild(tip); }
      tip.textContent = desc;
      tip.style.display = 'block';
      const r = infoBtn.getBoundingClientRect();
      const tipW = 260;
      tip.style.left = Math.min(r.left, window.innerWidth - tipW - 8) + 'px';
      tip.style.top  = (r.bottom + 6) + 'px';
    });
    infoBtn.addEventListener('mouseleave', () => {
      const tip = document.getElementById('field-info-tooltip');
      if (tip) tip.style.display = 'none';
    });
    label.appendChild(infoBtn);
  }

  cell.appendChild(label);

  const emptyText = item.emptyText || 'Non renseigné';

  if (editMode) {
    // Mode configuration : aperçu du placeholder + draggable
    const previewText = isEmpty ? emptyText : formatValPreview(val, col, item.refLabelField);
    const preview = document.createElement('div');
    preview.dataset.preview = '1';
    preview.style.cssText = `font-size:13px;padding:3px 0;${isEmpty ? 'color:var(--danger);font-style:italic;background:#fdf0f0;border-radius:3px;padding:2px 6px;display:inline-block;' : 'color:var(--text-muted);font-style:italic;'}`;
    preview.textContent = previewText;
    cell.appendChild(preview);
    addOverlay(cell, idx);
    return cell;
  }

  if (isEmpty) cell.classList.add('field-empty');

  if (!dataEditMode) {
    // Mode vue (présentation) : affichage statique
    if (isEmpty) {
      const emptySpan = document.createElement('span');
      emptySpan.className = 'field-empty-display';
      emptySpan.textContent = emptyText;
      cell.appendChild(emptySpan);
    } else {
      const displayEl = buildFieldDisplay(item, col, val, kind);
      if (displayEl) cell.appendChild(displayEl);
    }
    return cell;
  }

  // Champ formule pure : non modifiable en mode édition données
  if (col.isFormula) {
    if (isEmpty) {
      const emptySpan = document.createElement('span');
      emptySpan.className = 'field-empty-display';
      emptySpan.textContent = emptyText;
      cell.appendChild(emptySpan);
    } else {
      const displayEl = buildFieldDisplay(item, col, val, kind);
      if (displayEl) cell.appendChild(displayEl);
    }
    const badge = document.createElement('span');
    badge.textContent = 'FORMULE';
    badge.title = 'Ce champ est calculé automatiquement et ne peut pas être modifié';
    badge.style.cssText = 'font-size:10px;font-weight:600;letter-spacing:.05em;color:var(--text-muted);background:var(--surface-alt,#f0f0f0);border:1px solid var(--border);border-radius:3px;padding:1px 5px;margin-left:6px;vertical-align:middle;float:right;';
    label.appendChild(badge);
    return cell;
  }

  // Mode édition données : inputs interactifs
  if      (kind === 'bool')       cell.appendChild(buildBool(col, val));
  else if (kind === 'choice')     cell.appendChild(buildChoiceSelect(col, val, emptyText, v => markDirty(item.colId, v)));
  else if (kind === 'choiceList') cell.appendChild(buildChoiceList(col, val, v => markDirty(item.colId, v)));
  else if (kind === 'ref')        cell.appendChild(buildRefSearch(col, val, false, item.refLabelField, item.tagColor, emptyText, v => markDirty(item.colId, v)));
  else if (kind === 'refList')    cell.appendChild(buildRefSearch(col, val, true,  item.refLabelField, item.tagColor, emptyText, v => markDirty(item.colId, v)));
  else if (kind === 'date')       cell.appendChild(buildDate(col, val, false, emptyText));
  else if (kind === 'datetime')   cell.appendChild(buildDate(col, val, true, emptyText));
  else if (kind === 'attachment') cell.appendChild(buildAttachment(col, val, emptyText));
  else if (kind === 'longtext' || (typeof rawVal==='string' && rawVal.length>80)) cell.appendChild(buildLongText(col, val, emptyText));
  else if (kind === 'number')     cell.appendChild(buildNumber(col, val, emptyText));
  else                            cell.appendChild(buildText(col, val, emptyText));

  return cell;
}

function formatValPreview(val, col, labelField) {
  if (val===null||val===undefined||val==='') return 'Non renseigné';
  if (col.kind==='ref')        return getRefLabel(col.refTable, val, labelField);
  if (col.kind==='refList')    return (Array.isArray(val)?val.filter(id=>typeof id==='number'&&id>0):[val]).map(id=>getRefLabel(col.refTable,id,labelField)).join(', ');
  if (col.kind==='date')       return formatDatePreview(val, false);
  if (col.kind==='datetime')   return formatDatePreview(val, true);
  if (col.kind==='attachment') { const ids=Array.isArray(val)?val.filter(v=>typeof v==='number'&&v>0):[]; return ids.length?`${ids.length} pièce(s) jointe(s)`:'Aucune'; }
  if (Array.isArray(val))      return val.join(', ');
  return String(val);
}

/**
 * Convertit une valeur date Grist en millisecondes JS.
 * Grist peut stocker les dates en secondes, en jours ou (par bug) en ms.
 * On auto-détecte l'unité selon l'ordre de grandeur :
 *   < 100 000          → jours depuis epoch  (× 86 400 000)
 *   < 2 × 10^10        → secondes Unix       (× 1 000)
 *   ≥ 2 × 10^10        → millisecondes JS    (valeur brute)
 */
function gristDateToMs(val) {
  if (!val || typeof val !== 'number') return null;
  const abs = Math.abs(val);
  if (abs < 100000)   return val * 86400000; // jours
  if (abs < 2e10)     return val * 1000;     // secondes (format natif Grist)
  return val;                                 // ms (bug d'ancien save)
}

function formatDatePreview(val, withTime) {
  if (!val) return 'Non renseigné';
  try {
    const ms = gristDateToMs(val);
    if (ms === null) return String(val);
    const d = new Date(ms);
    if (isNaN(d.getTime())) return String(val);
    // Pour les dates seules, forcer UTC pour éviter le décalage timezone
    return withTime ? d.toLocaleString('fr-FR') : d.toLocaleDateString('fr-FR', { timeZone: 'UTC' });
  } catch(e) { return String(val); }
}

// ══════════════════════════════════════════════
// AFFICHAGE EN MODE VUE (lecture seule)
// ══════════════════════════════════════════════
function buildFieldDisplay(item, col, val, kind) {
  // Bool : checkbox désactivée
  if (kind === 'bool') {
    const wrap = document.createElement('div');
    wrap.style.cssText = 'padding:3px 0;';
    const chk = document.createElement('input');
    chk.type = 'checkbox'; chk.checked = !!val; chk.disabled = true;
    chk.className = 'grist-check';
    wrap.appendChild(chk);
    return wrap;
  }
  // Ref / RefList : tags colorés
  if (kind === 'ref' || kind === 'refList') {
    const refIds = Array.isArray(val) ? val.filter(v => typeof v === 'number' && v > 0)
                                      : (val && typeof val === 'number' ? [val] : []);
    if (!refIds.length) return null;
    const tagPalette = TAG_COLORS.find(c => c.key === item.tagColor) || TAG_COLORS[0];
    const wrap = document.createElement('div');
    wrap.style.cssText = 'display:flex;flex-wrap:wrap;gap:3px;padding:2px 0;';
    refIds.forEach(id => {
      const lbl = getRefLabel(col.refTable, id, item.refLabelField);
      const tag = document.createElement('span');
      tag.style.cssText = `display:inline-flex;align-items:center;padding:2px 8px;border-radius:10px;background:${tagPalette.bg};color:${tagPalette.text};font-size:11px;font-weight:600;`;
      tag.textContent = lbl;
      wrap.appendChild(tag);
    });
    return wrap;
  }
  // Choice / ChoiceList : tags bleus
  if (kind === 'choice' || kind === 'choiceList') {
    const values = kind === 'choice'
      ? (val ? [val] : [])
      : (Array.isArray(val) ? val.filter(v => v !== 'L') : []);
    if (!values.length) return null;
    const wrap = document.createElement('div');
    wrap.style.cssText = 'display:flex;flex-wrap:wrap;gap:3px;padding:2px 0;';
    values.forEach(v => {
      const tag = document.createElement('span');
      tag.style.cssText = 'display:inline-flex;padding:2px 8px;border-radius:10px;background:#e8f0fe;color:#1a73e8;font-size:11px;font-weight:600;';
      tag.textContent = v;
      wrap.appendChild(tag);
    });
    return wrap;
  }
  // Date / DateTime
  if (kind === 'date' || kind === 'datetime') {
    if (!val) return null;
    const span = document.createElement('span');
    span.style.cssText = 'font-size:13px;color:var(--text);';
    span.textContent = formatDatePreview(val, kind === 'datetime');
    return span;
  }
  // Pièces jointes
  if (kind === 'attachment') {
    const ids = Array.isArray(val) ? val.filter(v => typeof v === 'number' && v > 0) : [];
    if (!ids.length) return null;
    const wrap = document.createElement('div');
    wrap.className = 'attachment-list';
    ids.forEach((_, i) => {
      const chip = document.createElement('span');
      chip.className = 'attachment-chip';
      chip.innerHTML = `<svg width="10" height="10" viewBox="0 0 16 16" fill="currentColor"><path d="M3 2a1 1 0 011-1h7.586a1 1 0 01.707.293l1.414 1.414A1 1 0 0114 3.414V14a1 1 0 01-1 1H4a1 1 0 01-1-1V2z"/></svg> Fichier ${i+1}`;
      wrap.appendChild(chip);
    });
    return wrap;
  }
  // Texte / Nombre / LongText
  const text = (val !== null && val !== undefined) ? String(val) : '';
  if (!text) return null;
  const span = document.createElement('span');
  span.style.cssText = 'font-size:13px;color:var(--text);white-space:pre-wrap;line-height:1.5;';
  span.textContent = text;
  return span;
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

function buildText(col, val, emptyText = 'Non renseigné') {
  const ta = document.createElement('textarea');
  ta.className='form-field-textarea'; ta.value=val||''; ta.rows=1; ta.placeholder=emptyText;
  ta.style.minHeight = '22px';
  ta.addEventListener('input', () => {
    ta.style.height='auto'; ta.style.height=ta.scrollHeight+'px';
    markDirty(col.id, ta.value);
  });
  setTimeout(() => { ta.style.height='auto'; ta.style.height=ta.scrollHeight+'px'; }, 0);
  return ta;
}

function buildNumber(col, val, emptyText = 'Non renseigné') {
  const inp = document.createElement('input');
  inp.type='number'; inp.className='form-field-input'; inp.value=val||''; inp.placeholder=emptyText;
  inp.addEventListener('input', () => markDirty(col.id, inp.value===''?null:parseFloat(inp.value)));
  return inp;
}

function buildLongText(col, val, emptyText = 'Non renseigné') {
  const ta = document.createElement('textarea');
  ta.className='form-field-textarea'; ta.value=val||''; ta.rows=3; ta.placeholder=emptyText;
  ta.addEventListener('input', () => {
    ta.style.height='auto'; ta.style.height=ta.scrollHeight+'px'; markDirty(col.id, ta.value);
  });
  setTimeout(() => { ta.style.height='auto'; ta.style.height=ta.scrollHeight+'px'; }, 0);
  return ta;
}

// ── Date / DateTime ──
function buildDate(col, val, withTime, emptyText = 'Non renseigné') {
  const inp = document.createElement('input');
  inp.type = withTime ? 'datetime-local' : 'date';
  inp.className = 'form-field-input';
  inp.placeholder = emptyText;

  if (val && typeof val === 'number') {
    try {
      const ms = gristDateToMs(val);
      if (ms !== null) {
        const d = new Date(ms);
        if (!isNaN(d.getTime())) {
          if (withTime) {
            // datetime-local attend "YYYY-MM-DDTHH:MM" en heure locale
            const pad = n => String(n).padStart(2,'0');
            inp.value = `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
          } else {
            // Date seule : toISOString() donne toujours la date UTC
            inp.value = d.toISOString().split('T')[0];
          }
        }
      }
    } catch(e) {}
  }

  inp.addEventListener('change', () => {
    if (!inp.value) { markDirty(col.id, null); return; }
    const d = new Date(inp.value);
    if (isNaN(d.getTime())) { markDirty(col.id, null); return; }
    // Toujours sauvegarder en secondes Unix (format natif Grist)
    markDirty(col.id, Math.round(d.getTime() / 1000));
  });
  return inp;
}

// ── Pièce jointe ──
function buildAttachment(col, val, emptyText = 'Non renseigné') {
  const ids = Array.isArray(val) ? val.filter(v => typeof v === 'number' && v > 0) : [];

  const wrap = document.createElement('div');
  wrap.className = 'attachment-wrap';

  // Affichage des pièces jointes existantes
  if (ids.length > 0) {
    const list = document.createElement('div');
    list.className = 'attachment-list';
    ids.forEach((id, i) => {
      const chip = document.createElement('span');
      chip.className = 'attachment-chip';
      chip.innerHTML = `<svg width="10" height="10" viewBox="0 0 16 16" fill="currentColor"><path d="M3 2a1 1 0 011-1h7.586a1 1 0 01.707.293l1.414 1.414A1 1 0 0114 3.414V14a1 1 0 01-1 1H4a1 1 0 01-1-1V2z"/></svg> Fichier ${i+1}`;
      // Croix de suppression
      const rm = document.createElement('span');
      rm.textContent = '×'; rm.style.cssText = 'cursor:pointer;font-size:13px;opacity:.7;';
      rm.addEventListener('click', e => {
        e.stopPropagation();
        const newIds = ids.filter((_, j) => j !== i);
        markDirty(col.id, newIds.length ? ['L', ...newIds] : null);
        const cell = wrap.closest('.form-field');
        if (cell) cell.classList.toggle('field-empty', newIds.length === 0);
      });
      chip.appendChild(rm);
      list.appendChild(chip);
    });
    wrap.appendChild(list);
  } else {
    const empty = document.createElement('span');
    empty.textContent = emptyText;
    empty.style.cssText = 'font-size:12px;color:var(--text-muted);font-style:italic;';
    wrap.appendChild(empty);
  }

  // Bouton ajouter
  const fileInput = document.createElement('input');
  fileInput.type = 'file'; fileInput.multiple = true; fileInput.style.display = 'none';
  fileInput.addEventListener('change', async () => {
    const files = Array.from(fileInput.files);
    if (!files.length) return;
    try {
      const newIds = [];
      for (const file of files) {
        const id = await grist.docApi.uploadAttachment(file);
        newIds.push(id);
      }
      const allIds = [...ids, ...newIds];
      markDirty(col.id, ['L', ...allIds]);
      showToast(`${files.length} fichier(s) ajouté(s)`);
    } catch(e) {
      showToast('Gestion des PJ via Grist directement — ' + e.message, 4000);
    }
    fileInput.value = '';
  });

  const addBtn = document.createElement('button');
  addBtn.className = 'attachment-add-btn';
  addBtn.innerHTML = `<svg width="10" height="10" viewBox="0 0 16 16" fill="currentColor"><path d="M7.75 2a.75.75 0 01.75.75V7h4.25a.75.75 0 010 1.5H8.5v4.25a.75.75 0 01-1.5 0V8.5H2.75a.75.75 0 010-1.5H7V2.75A.75.75 0 017.75 2z"/></svg> Ajouter un fichier`;
  addBtn.addEventListener('click', () => fileInput.click());

  wrap.appendChild(fileInput);
  wrap.appendChild(addBtn);
  return wrap;
}

// ── Choice (simple) ──
function buildChoiceSelect(col, currentVal, emptyText = 'Non renseigné', onChange) {
  const sel = document.createElement('select');
  sel.className = 'form-field-input';
  sel.style.cssText = 'cursor:pointer;appearance:auto;padding:2px 0;';
  const empty = document.createElement('option');
  empty.value=''; empty.textContent=emptyText; sel.appendChild(empty);
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
        const cell = wrap.closest('.form-field');
        if (cell && !editMode) cell.classList.toggle('field-empty', selected.length === 0);
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
function buildRefSearch(col, currentVal, isMulti, labelField, tagColorKey, emptyText = 'Non renseigné', onChange) {
  // Pour RefList, Grist retourne ["L", id1, id2, ...] où "L" est un marqueur de type.
  const toRefIds = (val) => {
    if (!Array.isArray(val)) return (val && typeof val === 'number' && val > 0) ? [val] : [];
    return val.filter(id => typeof id === 'number' && id > 0);
  };
  // Pour refList : si Grist envoie une chaîne d'affichage (ex: "Jean Dupont") au lieu des IDs,
  // on l'accepte telle quelle comme ref le fait, plutôt que de filtrer et tout perdre.
  const toInitialSelected = (val) => {
    if (!isMulti) return val ? [val] : [];
    if (!Array.isArray(val)) return val ? [val] : [];
    const ids = toRefIds(val);
    return ids.length > 0 ? ids : [];
  };
  const tagPalette = TAG_COLORS.find(c => c.key === tagColorKey) || TAG_COLORS[0];

  let selected = toInitialSelected(currentVal);

  const wrap = document.createElement('div');
  wrap.style.cssText = 'display:flex;flex-direction:column;gap:4px;position:relative;';

  // Zone de tags
  const tagsRow = document.createElement('div');
  tagsRow.style.cssText = 'display:flex;flex-wrap:wrap;gap:3px;min-height:20px;';

  const updateCellEmpty = () => {
    const cell = wrap.closest('.form-field');
    if (cell && !editMode) cell.classList.toggle('field-empty', selected.length === 0);
  };

  const renderTags = () => {
    tagsRow.innerHTML = '';
    if (selected.length === 0) {
      const placeholder = document.createElement('span');
      placeholder.textContent = emptyText;
      placeholder.style.cssText = 'font-size:12px;color:var(--text-muted);font-style:italic;padding:2px 0;';
      tagsRow.appendChild(placeholder);
      return;
    }
    selected.forEach((id, i) => {
      const lbl = getRefLabel(col.refTable, id, labelField);
      const tag = makePill(lbl, tagPalette.bg, tagPalette.text, () => {
        selected.splice(i,1); renderTags();
        onChange(isMulti ? ['L', ...selected] : (selected[0]??null));
        updateCellEmpty();
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
    tb.classList.add('el-toolbar-right');
    // Taille de police
    [{l:'S',s:'s'},{l:'M',s:'m'},{l:'L',s:'l'},{l:'XL',s:'xl'}].forEach(opt => {
      const btn = tbBtn(opt.l, (item.fontSize||'m') === opt.s);
      btn.addEventListener('click', e => { e.stopPropagation(); layout[idx].fontSize=opt.s; saveLayout(); renderForm(); });
      tb.appendChild(btn);
    });
    // Style de section : fond / bordure / aucun
    const style = item.sectionStyle || 'fill';
    [{l:'Fond',s:'fill'},{l:'Aucun',s:'none'}].forEach(opt => {
      const btn = tbBtn(opt.l, style === opt.s);
      btn.addEventListener('click', e => { e.stopPropagation(); layout[idx].sectionStyle=opt.s; saveLayout(); renderForm(); });
      tb.appendChild(btn);
    });
    // Couleur DSFR
    const colorBtn = tbBtn('🎨 Couleur', false);
    colorBtn.addEventListener('click', e => { e.stopPropagation(); showColorPicker(idx, colorBtn); });
    tb.appendChild(colorBtn);
    // Toggle repliable
    const toggleBtn = tbBtn(item.noToggle ? '▶ Fixe' : '▼ Repliable', false);
    toggleBtn.addEventListener('click', e => {
      e.stopPropagation(); layout[idx].noToggle = !item.noToggle; saveLayout(); renderForm();
    });
    tb.appendChild(toggleBtn);
    // Demi-largeur
    const halfBtn = tbBtn(item.halfWidth ? '↔½' : '↔ Plein', item.halfWidth);
    halfBtn.addEventListener('click', e => {
      e.stopPropagation(); layout[idx].halfWidth = !item.halfWidth; saveLayout(); renderForm();
    });
    tb.appendChild(halfBtn);

  } else if (item.kind === 'desc') {
    tb.classList.add('el-toolbar-right');
    // Taille de police (inclut XL)
    [{l:'S',s:'s'},{l:'M',s:'m'},{l:'L',s:'l'},{l:'XL',s:'xl'}].forEach(opt => {
      const btn = tbBtn(opt.l, (item.fontSize||'m') === opt.s);
      btn.addEventListener('click', e => { e.stopPropagation(); layout[idx].fontSize=opt.s; saveLayout(); renderForm(); });
      tb.appendChild(btn);
    });

  } else if (item.kind === 'field') {
    const col = allColumns.find(c => c.id === item.colId);

    // Sélecteur de largeur compact
    const spanSel = document.createElement('select');
    spanSel.style.cssText = 'padding:2px 4px;border:1px solid var(--border);border-radius:4px;font-size:11px;font-family:var(--font);background:var(--surface);color:var(--text);cursor:pointer;height:22px;';
    [{l:'⅓ Largeur',s:2},{l:'½ Largeur',s:3},{l:'⅔ Largeur',s:4},{l:'↔ Plein',s:6}].forEach(opt => {
      const o = document.createElement('option');
      o.value = opt.s; o.textContent = opt.l;
      if (item.span === opt.s) o.selected = true;
      spanSel.appendChild(o);
    });
    spanSel.addEventListener('mousedown', e => e.stopPropagation());
    spanSel.addEventListener('change', e => { e.stopPropagation(); layout[idx].span = parseInt(spanSel.value); saveLayout(); renderForm(); });
    tb.appendChild(spanSel);

    // Boutons ref/refList
    if (col && (col.kind === 'ref' || col.kind === 'refList') && col.refTable) {
      if (refData[col.refTable]) {
        const pickBtn = tbBtn('🔍 Champ affiché', false);
        pickBtn.addEventListener('click', e => {
          e.stopPropagation();
          showRefFieldPicker(idx, col.refTable, item.refLabelField);
        });
        tb.appendChild(pickBtn);
      }
      const tagColorBtn = tbBtn('🏷 Couleur tag', false);
      tagColorBtn.addEventListener('click', e => {
        e.stopPropagation();
        showTagColorPicker(idx, tagColorBtn, item.tagColor || 'blue');
      });
      tb.appendChild(tagColorBtn);
    }

    // Emoji
    const emojiBtn = tbBtn((item.emoji || '☐') + ' Icône', false);
    emojiBtn.addEventListener('click', e => {
      e.stopPropagation();
      showEmojiPicker(idx, emojiBtn, item.emoji || '');
    });
    tb.appendChild(emojiBtn);

    // Placeholder (texte si champ vide)
    const emptyBtn = tbBtn('💬 Placeholder', false);
    emptyBtn.addEventListener('click', e => {
      e.stopPropagation();
      showEmptyTextPicker(idx, emptyBtn, item.emptyText || '');
    });
    tb.appendChild(emptyBtn);

    // Renommer (dans le widget uniquement, pas dans Grist)
    const renBtn = tbBtn('✏ Renommer', false);
    renBtn.addEventListener('click', e => {
      e.stopPropagation();
      showRenamePicker(idx, renBtn, item.label || item.colId);
    });
    tb.appendChild(renBtn);
  }
  // desc et separator : uniquement drag handle + masquer

  // Masquer (retire l'élément du layout widget uniquement)
  const delBtn = tbBtn('✕ Masquer', false, true);
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

// ── Picker emoji pour les champs ──
function showEmojiPicker(layoutIdx, anchorBtn, currentEmoji) {
  const old = document.getElementById('emoji-picker');
  if (old) { old.remove(); return; }

  const EMOJIS = [
    '👤','👥','🏢','📍','📌','📅','🗓','⏰','💶','💰','📊','📈',
    '🎯','🚀','💡','🔑','📝','💬','✅','❌','⚠️','🔗','📎','📂',
    '📁','🗂','📋','🔧','⚙️','🌐','📧','📱','💻','🏠','🏗','🔢',
    '🏷','🎨','📦','🚚',
  ];

  const pop = document.createElement('div');
  pop.id = 'emoji-picker';
  const rect = anchorBtn.getBoundingClientRect();
  pop.style.cssText = `
    position:fixed;top:${rect.bottom + 6}px;left:${safePopLeft(rect, 230)}px;z-index:500;
    background:var(--surface);border:1px solid var(--border);border-radius:8px;
    box-shadow:0 4px 16px rgba(0,0,0,0.14);padding:10px;width:230px;box-sizing:border-box;
  `;
  // Empêche les clics à l'intérieur de fermer le popover
  pop.addEventListener('click', e => e.stopPropagation());

  // Input libre
  const row = document.createElement('div');
  row.style.cssText = 'display:flex;gap:6px;margin-bottom:8px;align-items:center;';
  const inp = document.createElement('input');
  inp.type = 'text'; inp.placeholder = 'Coller ou taper une emoji…';
  inp.value = currentEmoji;
  inp.style.cssText = 'flex:1;min-width:0;padding:4px 8px;border:1px solid var(--border);border-radius:4px;font-size:16px;font-family:var(--font);box-sizing:border-box;';
  const clearBtn = document.createElement('button');
  clearBtn.textContent = '✕';
  clearBtn.title = 'Retirer l\'icône';
  clearBtn.style.cssText = 'flex-shrink:0;padding:4px 8px;border:1px solid var(--border);border-radius:4px;background:none;cursor:pointer;color:var(--text-muted);';
  clearBtn.addEventListener('click', e => {
    e.stopPropagation();
    layout[layoutIdx].emoji = '';
    saveLayout(); renderForm(); pop.remove();
  });
  row.appendChild(inp); row.appendChild(clearBtn);
  pop.appendChild(row);

  // Grille rapide
  const grid = document.createElement('div');
  grid.style.cssText = 'display:flex;flex-wrap:wrap;gap:4px;';
  EMOJIS.forEach(em => {
    const btn = document.createElement('button');
    btn.textContent = em;
    btn.title = em;
    btn.style.cssText = `
      width:30px;height:30px;border:1.5px solid ${em === currentEmoji ? 'var(--accent)' : 'transparent'};
      border-radius:4px;background:${em === currentEmoji ? 'var(--accent-light)' : 'none'};
      cursor:pointer;font-size:16px;line-height:1;transition:background .1s;
    `;
    btn.addEventListener('click', e => {
      e.stopPropagation();
      layout[layoutIdx].emoji = em;
      saveLayout(); renderForm(); pop.remove();
    });
    grid.appendChild(btn);
  });
  pop.appendChild(grid);

  // Valider l'input libre
  inp.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.stopPropagation();
      layout[layoutIdx].emoji = inp.value.trim();
      saveLayout(); renderForm(); pop.remove();
    }
  });

  document.addEventListener('click', () => pop.remove(), { once: true });
  document.body.appendChild(pop);
}

// ── Picker texte vide (remplace prompt natif) ──
function showEmptyTextPicker(layoutIdx, anchorBtn, currentText) {
  const OLD_ID = 'empty-text-picker';
  const old = document.getElementById(OLD_ID);
  if (old) { old.remove(); return; }

  const pop = document.createElement('div');
  pop.id = OLD_ID;
  const rect = anchorBtn.getBoundingClientRect();
  pop.style.cssText = `
    position:fixed;top:${rect.bottom + 6}px;left:${safePopLeft(rect, 240)}px;z-index:500;
    background:var(--surface);border:1px solid var(--border);border-radius:8px;
    box-shadow:0 4px 16px rgba(0,0,0,0.14);padding:10px;width:240px;box-sizing:border-box;
  `;
  pop.addEventListener('click', e => e.stopPropagation());

  const label = document.createElement('div');
  label.textContent = 'Texte si champ vide :';
  label.style.cssText = 'font-size:11px;color:var(--text-label);margin-bottom:6px;';
  pop.appendChild(label);

  const row = document.createElement('div');
  row.style.cssText = 'display:flex;gap:6px;align-items:center;';

  const inp = document.createElement('input');
  inp.type = 'text';
  inp.placeholder = 'Non renseigné';
  inp.value = currentText;
  inp.style.cssText = 'flex:1;min-width:0;padding:5px 8px;border:1px solid var(--border);border-radius:4px;font-size:13px;font-family:var(--font);box-sizing:border-box;';

  const okBtn = document.createElement('button');
  okBtn.textContent = '✓';
  okBtn.title = 'Valider';
  okBtn.style.cssText = 'flex-shrink:0;padding:5px 10px;border:none;border-radius:4px;background:var(--accent);color:#fff;font-size:13px;cursor:pointer;';

  // Preview live dans le formulaire (mode config)
  const applyPreview = (text) => {
    layout[layoutIdx].emptyText = text;
    // Met à jour seulement le preview du champ concerné sans tout re-rendre
    const previewEl = el('product-form')
      .querySelector(`.form-field[data-idx="${layoutIdx}"] [data-preview]`);
    if (previewEl) previewEl.textContent = text || 'Non renseigné';
  };

  const save = () => {
    layout[layoutIdx].emptyText = inp.value.trim();
    saveLayout(); pop.remove();
  };

  inp.addEventListener('input', () => applyPreview(inp.value.trim() || 'Non renseigné'));
  okBtn.addEventListener('click', e => { e.stopPropagation(); save(); });
  inp.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.stopPropagation(); save(); }
    if (e.key === 'Escape') { e.stopPropagation(); pop.remove(); }
  });

  row.appendChild(inp);
  row.appendChild(okBtn);
  pop.appendChild(row);

  document.addEventListener('click', () => { save(); }, { once: true });
  document.body.appendChild(pop);

  // Focus automatique
  setTimeout(() => { inp.focus(); inp.select(); }, 0);
}

// ── Picker renommage (widget uniquement, ne modifie pas Grist) ──
function showRenamePicker(layoutIdx, anchorBtn, currentLabel) {
  const OLD_ID = 'rename-picker';
  const old = document.getElementById(OLD_ID);
  if (old) { old.remove(); return; }

  const pop = document.createElement('div');
  pop.id = OLD_ID;
  const rect = anchorBtn.getBoundingClientRect();
  pop.style.cssText = `
    position:fixed;top:${rect.bottom + 6}px;left:${safePopLeft(rect, 260)}px;z-index:500;
    background:var(--surface);border:1px solid var(--border);border-radius:8px;
    box-shadow:0 4px 16px rgba(0,0,0,0.14);padding:10px;width:260px;box-sizing:border-box;
  `;
  pop.addEventListener('click', e => e.stopPropagation());

  const lbl = document.createElement('div');
  lbl.textContent = 'Nom affiché dans le widget :';
  lbl.style.cssText = 'font-size:11px;color:var(--text-label);margin-bottom:4px;';
  const sub = document.createElement('div');
  sub.textContent = '(ne modifie pas le nom de la colonne dans Grist)';
  sub.style.cssText = 'font-size:10px;color:var(--text-muted);margin-bottom:8px;font-style:italic;';
  pop.appendChild(lbl);
  pop.appendChild(sub);

  const row = document.createElement('div');
  row.style.cssText = 'display:flex;gap:6px;align-items:center;';
  const inp = document.createElement('input');
  inp.type = 'text'; inp.value = currentLabel;
  inp.style.cssText = 'flex:1;min-width:0;padding:5px 8px;border:1px solid var(--border);border-radius:4px;font-size:13px;font-family:var(--font);box-sizing:border-box;';
  const okBtn = document.createElement('button');
  okBtn.textContent = '✓'; okBtn.title = 'Valider';
  okBtn.style.cssText = 'flex-shrink:0;padding:5px 10px;border:none;border-radius:4px;background:var(--accent);color:#fff;font-size:13px;cursor:pointer;';

  const save = () => {
    const v = inp.value.trim();
    if (v) { layout[layoutIdx].label = v; saveLayout(); renderForm(); }
    pop.remove();
  };
  okBtn.addEventListener('click', e => { e.stopPropagation(); save(); });
  inp.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.stopPropagation(); save(); }
    if (e.key === 'Escape') { e.stopPropagation(); pop.remove(); }
  });
  row.appendChild(inp); row.appendChild(okBtn);
  pop.appendChild(row);

  document.addEventListener('click', () => pop.remove(), { once: true });
  document.body.appendChild(pop);
  setTimeout(() => { inp.focus(); inp.select(); }, 0);
}

// ── Picker couleur pour les tags ref/refList ──
function showTagColorPicker(layoutIdx, anchorBtn, currentKey) {
  const old = document.getElementById('tag-color-picker');
  if (old) { old.remove(); return; }

  const pop = document.createElement('div');
  pop.id = 'tag-color-picker';
  const rect = anchorBtn.getBoundingClientRect();
  pop.style.cssText = `
    position:fixed;top:${rect.bottom + 6}px;left:${safePopLeft(rect, 186)}px;z-index:500;
    background:var(--surface);border:1px solid var(--border);border-radius:6px;
    box-shadow:0 4px 16px rgba(0,0,0,0.14);padding:8px;
    display:flex;flex-wrap:wrap;gap:6px;width:186px;
  `;

  TAG_COLORS.forEach(c => {
    const swatch = document.createElement('button');
    swatch.title = c.label;
    swatch.style.cssText = `
      width:32px;height:32px;border-radius:6px;border:2px solid ${c.key === currentKey ? c.text : 'transparent'};
      background:${c.bg};cursor:pointer;display:flex;align-items:center;justify-content:center;
      font-size:11px;font-weight:700;color:${c.text};transition:border .1s;
    `;
    swatch.textContent = 'Aa';
    swatch.addEventListener('click', e => {
      e.stopPropagation();
      layout[layoutIdx].tagColor = c.key;
      saveLayout(); renderForm(); pop.remove();
    });
    pop.appendChild(swatch);
  });

  document.addEventListener('click', () => pop.remove(), { once: true });
  document.body.appendChild(pop);
}

// ── Mini-modal pour choisir le champ de recherche d'une référence ──
function showRefFieldPicker(layoutIdx, refTableId, currentField) {
  const data = refData[refTableId];
  if (!data) return;

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
  header.style.cssText = 'height:42px;padding:0 14px;border-bottom:1px solid var(--border);font-size:13px;font-weight:600;display:flex;align-items:center;justify-content:space-between;';
  header.innerHTML = `<span>Champ d'affichage / recherche</span>`;
  const close = document.createElement('button');
  close.textContent = '×'; close.style.cssText = 'border:none;background:none;cursor:pointer;font-size:18px;color:var(--text-muted);';
  close.addEventListener('click', () => modal.remove());
  header.appendChild(close);
  box.appendChild(header);

  const sub = document.createElement('div');
  sub.style.cssText = 'padding:8px 14px 4px;font-size:11px;color:var(--text-muted);';
  sub.textContent = `Table : ${refTableId} — choisissez la colonne utilisée pour afficher et rechercher les entrées`;
  box.appendChild(sub);

  const list = document.createElement('div');
  list.style.cssText = 'padding:6px;max-height:280px;overflow-y:auto;';

  data.columns.forEach(colId => {
    const btn = document.createElement('button');
    btn.style.cssText = 'display:flex;align-items:center;gap:8px;width:100%;padding:7px 10px;border:none;background:none;font-family:var(--font);font-size:13px;color:var(--text);cursor:pointer;border-radius:var(--r);text-align:left;transition:background .1s;';
    btn.innerHTML = `
      <span style="flex:1">${escHtml(colId.replace(/_/g,' '))}</span>
      ${colId === currentField ? `<span style="color:var(--accent);font-size:11px;font-weight:700;">✓ Actif</span>` : ''}
    `;
    btn.addEventListener('mouseenter', () => btn.style.background = 'var(--bg)');
    btn.addEventListener('mouseleave', () => btn.style.background = '');
    btn.addEventListener('click', () => {
      layout[layoutIdx].refLabelField = colId;
      saveLayout(); renderForm(); modal.remove();
    });
    list.appendChild(btn);
  });

  box.appendChild(list);
  modal.appendChild(box);
  modal.addEventListener('click', e => { if (e.target === modal) modal.remove(); });
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
function resetLayout() {
  if (!confirm('Retirer tous les champs et éléments du formulaire ?')) return;
  layout = [];
  saveLayout(); renderForm();
}

function addItem(kind) {
  const defaults = { title:'Nouvelle section', desc:'Description…', separator:'' };
  const item = { id:newId(), kind, label:defaults[kind]??'', span:kind==='field'?3:undefined };
  if (kind === 'title') {
    item.collapsed = false; item.bgColor = null;
    item.sectionStyle = 'fill'; item.noToggle = false; item.halfWidth = false;
  }
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

// ══════════════════════════════════════════════
// RICH TEXT TOOLBAR
// ══════════════════════════════════════════════
function initRichToolbar() {
  const tb = el('rich-toolbar');
  if (!tb) return;

  // Exécuter une commande de formatage au clic (mousedown pour conserver la sélection)
  tb.querySelectorAll('[data-cmd]').forEach(btn => {
    btn.addEventListener('mousedown', e => {
      e.preventDefault(); // empêche la perte de la sélection
      document.execCommand(btn.dataset.cmd, false, null);
      // Sauvegarder le contenu mis à jour
      const sel = window.getSelection();
      if (sel && sel.rangeCount) {
        const node = sel.getRangeAt(0).commonAncestorContainer;
        const ce = (node.nodeType === 3 ? node.parentElement : node).closest('[contenteditable="true"]');
        if (ce && ce.dataset.layoutIdx !== undefined) {
          layout[parseInt(ce.dataset.layoutIdx)].label = ce.innerHTML;
          saveLayout();
        }
      }
      updateRichToolbarState();
    });
  });

  // Afficher/repositionner la toolbar à chaque changement de sélection
  document.addEventListener('selectionchange', () => {
    if (!editMode) { tb.style.display = 'none'; return; }
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || !sel.rangeCount) { tb.style.display = 'none'; return; }
    const node = sel.getRangeAt(0).commonAncestorContainer;
    const ce = (node.nodeType === 3 ? node.parentElement : node).closest('[contenteditable="true"]');
    if (!ce) { tb.style.display = 'none'; return; }

    const rect = sel.getRangeAt(0).getBoundingClientRect();
    tb.style.display = 'flex';
    tb.style.left = Math.max(4, rect.left + rect.width / 2 - 42) + 'px';
    tb.style.top  = Math.max(4, rect.top - 34) + 'px';
    updateRichToolbarState();
  });

  // Masquer si clic en dehors d'un contenteditable ou de la toolbar
  document.addEventListener('mousedown', e => {
    if (!tb.contains(e.target) && !e.target.closest('[contenteditable="true"]')) {
      tb.style.display = 'none';
    }
  });
}

function updateRichToolbarState() {
  const tb = el('rich-toolbar');
  if (!tb) return;
  tb.querySelectorAll('[data-cmd]').forEach(btn => {
    btn.classList.toggle('rt-active', document.queryCommandState(btn.dataset.cmd));
  });
}

// Initialisation au chargement
document.addEventListener('DOMContentLoaded', () => {
  initRichToolbar();
  setupProductSearch();
  updateTopbarButtons();
  updateAccessUI();
});