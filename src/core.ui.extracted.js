// ===== CSV COLUMN MAP =====
const CSV_COL_MAP = {
  'Account Classification': 'category',
  'Company Name':           'company',
  'Priority Focus':         'priority',
  'Status / Next Steps':    'status',
  'Interacted With?':       'interactedWith',
  'Next Meeting Date':      'meeting',
  'Cluster POC':            'clusterPOC',
  'Key Account POC':        'keyPOC',
  'Key AccountPOC':          'keyPOC',
  'Client POC':             'clientPOC',
  'Existing Vendors':       'existingVendors',
  'Existing Work Done':     'existingWorkDone',
  'Vendor Pipeline Status': 'pipeline',
  'Action Items':           'action',
  'Last Updated Date':      'lastUpdatedAt',
  'Last Updated By':        'lastUpdatedBy',
  'Notes':                  'nextsteps',
};

const LS_ACCOUNTS_KEY = 'ec_fsm_accounts_v2';
const LS_UPLOAD_KEY   = 'ec_fsm_upload_label';

function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 3000);
}

function parseCSV(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const headers = splitCSVRow(lines[0]);
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const vals = splitCSVRow(lines[i]);
    const obj = {};
    headers.forEach((h, idx) => { obj[h.trim()] = (vals[idx] || '').trim(); });
    rows.push(obj);
  }
  return rows;
}

function splitCSVRow(row) {
  const result = [];
  let cur = '', inQuote = false;
  for (let i = 0; i < row.length; i++) {
    const ch = row[i];
    if (ch === '"') {
      if (inQuote && row[i+1] === '"') { cur += '"'; i++; }
      else inQuote = !inQuote;
    } else if (ch === ',' && !inQuote) {
      result.push(cur); cur = '';
    } else {
      cur += ch;
    }
  }
  result.push(cur);
  return result;
}

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv') {
    handleCSVUpload(event);
  } else {
    handleXLSXUpload(event);
  }
}

function handleXLSXUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  event.target.value = '';
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb = XLSX.read(e.target.result, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];

      // Get all rows as raw arrays to find the header row
      const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      // Find the row that contains "Company Name"
      let headerRowIdx = -1;
      for (let r = 0; r < Math.min(allRows.length, 15); r++) {
        if (allRows[r].some(v => String(v).trim().toLowerCase() === 'company name')) {
          headerRowIdx = r;
          break;
        }
      }
      if (headerRowIdx === -1) {
        const sample = allRows.slice(0,5).map(r => r.filter(Boolean).join(' | ')).join('\n');
        showToast('⚠ Could not find header row. First rows: ' + sample.substring(0,120));
        return;
      }

      // Build header→field map (case-insensitive, trimmed)
      const headers = allRows[headerRowIdx].map(h => String(h).trim());
      const colIndexMap = {};
      headers.forEach((h, i) => {
        const match = Object.keys(CSV_COL_MAP).find(k => k.trim().toLowerCase() === h.toLowerCase());
        if (match) colIndexMap[i] = CSV_COL_MAP[match];
      });

      // Parse data rows
      const parsed = [];
      for (let r = headerRowIdx + 1; r < allRows.length; r++) {
        const row = allRows[r];
        if (row.every(v => !String(v).trim())) continue; // skip blank rows
        const acct = { id: parsed.length + 1 };
        Object.entries(colIndexMap).forEach(([idx, field]) => {
          acct[field] = String(row[idx] || '').trim();
        });
        if (acct.company) parsed.push(acct);
      }

      if (!parsed.length) {
        showToast('⚠ Headers found at row ' + (headerRowIdx+1) + ' but no data rows matched.');
        return;
      }
      accounts = parsed;
      nextId = parsed.length + 1;
      const label = file.name + ' · ' + new Date().toLocaleDateString();
      localStorage.setItem(LS_ACCOUNTS_KEY, JSON.stringify(accounts));
      localStorage.setItem(LS_UPLOAD_KEY, label);
      setUploadLabel(label);
      renderTable();
      showToast('✓ Loaded ' + parsed.length + ' accounts from ' + file.name);
    } catch(err) {
      showToast('⚠ Error: ' + err.message);
      console.error(err);
    }
  };
  reader.readAsBinaryString(file);
}
function handleCSVUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const rows = parseCSV(e.target.result);
      const parsed = rows.map((row, i) => {
        const acct = { id: i + 1 };
        Object.entries(CSV_COL_MAP).forEach(([col, field]) => {
          acct[field] = row[col] !== undefined ? row[col] : '';
        });
        return acct;
      }).filter(a => a.company);
      if (!parsed.length) { showToast('⚠ No accounts found — check column headers match'); return; }
      accounts = parsed;
      nextId = parsed.length + 1;
      const label = file.name + ' · ' + new Date().toLocaleDateString();
      localStorage.setItem(LS_ACCOUNTS_KEY, JSON.stringify(accounts));
      localStorage.setItem(LS_UPLOAD_KEY, label);
      setUploadLabel(label);
      renderTable();
      showToast('✓ Loaded ' + parsed.length + ' accounts from ' + file.name);
    } catch(err) {
      showToast('⚠ Could not read CSV: ' + err.message);
    }
  };
  reader.readAsText(file);
  event.target.value = '';
}

function setUploadLabel(label) {
  const el = document.getElementById('uploadLabel');
  if (label) { el.textContent = '📄 ' + label; el.style.display = ''; }
  else { el.style.display = 'none'; }
}

function loadCache() {
  try {
    const raw = localStorage.getItem(LS_ACCOUNTS_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && parsed.length > 0) {
      accounts = parsed;
      nextId = Math.max(...parsed.map(a => Number(a.id) || 0), 0) + 1;
      setUploadLabel(localStorage.getItem(LS_UPLOAD_KEY) || '');
    }
  } catch(e) {}
}

// ===== DATA =====
let accounts = [];

const alliances = [
  {
    vendor: "Microsoft",
    internalPOC: "Tom Kirby (tokirby@deloitte.com), Thomas Hale (thhale@deloitte.com)",
    vendorPOC: "TBD",
    comment: "",
    update: "[10/29] Thomas Hale to run through research lead in USI to validate. ~1 week turnaround needed."
  },
  {
    vendor: "ServiceNow",
    internalPOC: "Simmi Mehta (spmehta@deloitte.com), Josh Lehrberger (jlehrberger@deloitte.com), Andrew Obey (aobey@deloitte.com)",
    vendorPOC: "TBD",
    comment: "",
    update: "[10/29] Simmi confirmed only Alcoa US Corp on FS Cross-Sell target list. Brett to confirm if additional call with account team needed. Pending confirmation on whether other 35 accounts should be included in campaign."
  },
  {
    vendor: "Salesforce",
    internalPOC: "Michael E Moore (mimoore@deloitte.com), Tannyr Deakins (tdeakins@deloitte.com)",
    vendorPOC: "Nathan (Alliance Dir), Bill & Ryan (RVPs)",
    comment: "",
    update: "[10/29] Tannyr to ask SF alliance counterpart for FS pipe breakdown at included accounts in next 1:1. Michael to run through Salesforce Alliance Dir Nathan and coordinate with RVPs. Tannyr to support on Sales Ops."
  }
];

let editingId = null;
let sortField = null;
let sortDir = 1;
let nextId = 1;
let activeCardFilter = null; // 'high' | 'action' | 'engaged' | null

// ===== USERNAME =====

function setUsername() {
  const current = localStorage.getItem('ec_fsm_username') || '';
  const name = prompt('Your name (shown when you edit accounts):', current);
  if (name !== null && name.trim()) {
    localStorage.setItem('ec_fsm_username', name.trim());
    updateUserChip();
  }
}

function updateUserChip() {
  const name = localStorage.getItem('ec_fsm_username');
  const chip = document.getElementById('userChip');
  const label = document.getElementById('userLabel');
  if (name) {
    label.textContent = name;
    chip.classList.add('named');
    chip.title = 'Click to change name';
  } else {
    label.textContent = 'Set your name';
    chip.classList.remove('named');
    chip.title = 'Click to set your name';
  }
}

function getUsername() {
  return localStorage.getItem('ec_fsm_username') || 'Unknown';
}

function todayStr() {
  const d = new Date();
  return (d.getMonth()+1).toString().padStart(2,'0') + '/' +
         d.getDate().toString().padStart(2,'0') + '/' +
         d.getFullYear().toString().slice(2);
}

// ===== RENDERING =====

function catClass(cat) {
  if (!cat) return '';
  if (cat.includes('ACP')) return 'cat-acp';
  if (cat.includes('Upstream')) return 'cat-upstream';
  if (cat.includes('Mid')) return 'cat-mid';
  if (cat.includes('ETS')) return 'cat-ets';
  if (cat.includes('Mining')) return 'cat-mining';
  return '';
}
function priClass(p) {
  if (p === 'HIGH') return 'badge-high';
  if (p === 'MEDIUM') return 'badge-medium';
  if (p === 'LOW') return 'badge-low';
  return 'badge-tbd';
}
function esc(s) {
  return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');
}

function filterByCard(type) {
  activeCardFilter = activeCardFilter === type ? null : type;
  renderTable();
}

function renderSummary() {
  const total = accounts.length;
  const high = accounts.filter(a=>a.priority==='HIGH').length;
  const withAction = accounts.filter(a=>a.action).length;
  const engaged = accounts.filter(a=>a.status && a.status !== 'TBD').length;
  const af = activeCardFilter;
  document.getElementById('summaryRow').innerHTML = `
    <div class="summary-card green${af===null?'':''}" onclick="filterByCard(null)" title="Show all accounts"><div class="summary-num">${total}</div><div class="summary-label">Total Accounts</div></div>
    <div class="summary-card red${af==='high'?' active-filter':''}" onclick="filterByCard('high')" title="Filter: HIGH Priority"><div class="summary-num">${high}</div><div class="summary-label">HIGH Priority</div></div>
    <div class="summary-card amber${af==='action'?' active-filter':''}" onclick="filterByCard('action')" title="Filter: Open Action Items"><div class="summary-num">${withAction}</div><div class="summary-label">Open Action Items</div></div>
    <div class="summary-card blue${af==='engaged'?' active-filter':''}" onclick="filterByCard('engaged')" title="Filter: Active Engagements"><div class="summary-num">${engaged}</div><div class="summary-label">Active Engagements</div></div>
  `;
}

function renderTable() {
  const q = document.getElementById('searchInput').value.toLowerCase();
  const fCat = document.getElementById('filterCat').value;
  const fPri = document.getElementById('filterPriority').value;

  let rows = accounts.filter(a => {
    const text = [a.company, a.status, a.nextsteps, a.clusterPOC, a.keyPOC, a.clientPOC, a.pipeline, a.action, a.interactedWith, a.existingVendors, a.existingWorkDone].join(' ').toLowerCase();
    const cardOk = !activeCardFilter
      || (activeCardFilter === 'high'    && a.priority === 'HIGH')
      || (activeCardFilter === 'action'  && !!a.action)
      || (activeCardFilter === 'engaged' && a.status && a.status !== 'TBD');
    return (!q || text.includes(q))
      && (!fCat || a.category === fCat)
      && (!fPri || a.priority === fPri)
      && cardOk;
  });

  if (sortField) {
    rows.sort((a,b) => {
      const av = (a[sortField]||'').toLowerCase();
      const bv = (b[sortField]||'').toLowerCase();
      return av < bv ? -sortDir : av > bv ? sortDir : 0;
    });
  }

  document.getElementById('countBadge').textContent = rows.length + ' accounts';
  renderSummary();

  const tbody = document.getElementById('tableBody');
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="9"><div class="empty-state"><div class="icon">🔍</div>No accounts match your filters.</div></td></tr>`;
    return;
  }

  tbody.innerHTML = rows.map(a => {
    // POC: first name only (before comma or parenthesis)
    const firstName = s => s ? s.split(',')[0].split('(')[0].trim() : '';

    // Status + Notes cell
    const statusCell = `
      ${a.status ? `<div class="cell-section"><div class="cell-clamp">${esc(a.status)}</div></div>` : ''}
      ${a.nextsteps ? `<div class="next-steps"><div class="cell-label">Notes</div><div class="cell-clamp-2">${esc(a.nextsteps)}</div></div>` : ''}
      ${!a.status && !a.nextsteps ? '<span class="cell-empty">No update yet</span>' : ''}`;

    // POC cell — every field uses same label+value block; omitted fields are simply absent
    const pocRow = (label, val, chipClass) => val
      ? `<div class="cell-section"><div class="cell-label">${label}</div><span class="poc-chip${chipClass ? ' ' + chipClass : ''}">${esc(val)}</span></div>`
      : '';
    const pocCell = [
      pocRow('Cluster',   firstName(a.clusterPOC), ''),
      pocRow('Key Acct',  firstName(a.keyPOC),     ''),
      pocRow('Client',    a.clientPOC,              'client'),
      a.interactedWith
        ? `<div class="cell-section"><div class="cell-label">Talked To</div><div class="cell-clamp-2" style="font-size:11px">${esc(a.interactedWith)}</div></div>`
        : '',
    ].filter(Boolean).join('') || '<span class="cell-empty">—</span>';

    // Pipeline cell
    const pipelineCell = `
      ${a.pipeline ? `<div class="cell-section"><div class="cell-clamp-2">${esc(a.pipeline)}</div></div>` : ''}
      ${a.existingWorkDone ? `<div class="cell-section"><div class="cell-label">Work Done</div><div class="cell-clamp-2">${esc(a.existingWorkDone)}</div></div>` : ''}
      ${!a.pipeline && !a.existingWorkDone ? '<span class="cell-empty">—</span>' : ''}`;

    // Vendors cell
    const vendorsCell = a.existingVendors
      ? a.existingVendors.split(',').map(v => `<span class="vendor-chip">${esc(v.trim())}</span>`).join('')
      : '<span class="cell-empty">—</span>';

    // Action cell
    const actionCell = `
      ${a.action ? `<div class="cell-section" style="display:flex;align-items:baseline;gap:4px"><span class="action-dot${a.action.toLowerCase().includes('urgent') ? ' urgent' : ''}"></span><div class="cell-clamp-2">${esc(a.action)}</div></div>` : '<span class="cell-empty">—</span>'}
      ${a.lastUpdatedBy ? `<div class="updated-by">✎ ${esc(a.lastUpdatedBy)}${a.lastUpdatedAt ? ' · ' + esc(a.lastUpdatedAt) : ''}</div>` : ''}`;

    return `<tr onclick="openModal(${a.id})">
      <td><span class="cat-badge ${catClass(a.category)}">${a.category||'—'}</span></td>
      <td style="font-weight:600;min-width:130px">${esc(a.company)}</td>
      <td style="min-width:70px">${a.priority ? `<span class="badge ${priClass(a.priority)}">${a.priority}</span>` : '<span class="cell-empty">—</span>'}</td>
      <td class="status-cell">${statusCell}</td>
      <td style="min-width:140px">${pocCell}</td>
      <td style="min-width:160px;max-width:200px">${pipelineCell}</td>
      <td style="min-width:100px">${vendorsCell}</td>
      <td style="min-width:150px">${actionCell}</td>
      <td><button class="btn btn-secondary btn-sm" onclick="event.stopPropagation();openModal(${a.id})">Edit</button></td>
    </tr>`;
  }).join('');
}

const VENDOR_LOGOS = {
  Microsoft: `<svg class="alliance-logo" version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 337.6 72">
<path fill="#737373" d="M140.4,14.4v43.2h-7.5V23.7h-0.1l-13.4,33.9h-5l-13.7-33.9h-0.1v33.9h-6.9V14.4h10.8l12.4,32h0.2l13.1-32H140.4 z M146.6,17.7c0-1.2,0.4-2.2,1.3-3c0.9-0.8,1.9-1.2,3.1-1.2c1.3,0,2.4,0.4,3.2,1.2s1.3,1.8,1.3,3c0,1.2-0.4,2.2-1.3,3 c-0.9,0.8-1.9,1.2-3.2,1.2s-2.3-0.4-3.1-1.2C147.1,19.8,146.6,18.8,146.6,17.7z M154.7,26.6v31h-7.3v-31H154.7z M176.8,52.3 c1.1,0,2.3-0.2,3.6-0.8c1.3-0.5,2.5-1.2,3.6-2v6.8c-1.2,0.7-2.5,1.2-4,1.5c-1.5,0.3-3.1,0.5-4.9,0.5c-4.6,0-8.3-1.4-11.1-4.3 c-2.9-2.9-4.3-6.6-4.3-11c0-5,1.5-9.1,4.4-12.3c2.9-3.2,7-4.8,12.4-4.8c1.4,0,2.8,0.2,4.1,0.5c1.4,0.3,2.5,0.8,3.3,1.2v7 c-1.1-0.8-2.3-1.5-3.4-1.9c-1.2-0.4-2.4-0.7-3.6-0.7c-2.9,0-5.2,0.9-7,2.8s-2.6,4.4-2.6,7.6c0,3.1,0.9,5.6,2.6,7.3 C171.6,51.4,173.9,52.3,176.8,52.3z M204.7,26.1c0.6,0,1.1,0,1.6,0.1s0.9,0.2,1.2,0.3v7.4c-0.4-0.3-0.9-0.6-1.7-0.8 s-1.6-0.4-2.7-0.4c-1.8,0-3.3,0.8-4.5,2.3s-1.9,3.8-1.9,7v15.6h-7.3v-31h7.3v4.9h0.1c0.7-1.7,1.7-3,3-4 C201.2,26.6,202.8,26.1,204.7,26.1z M207.9,42.6c0-5.1,1.5-9.2,4.3-12.2c2.9-3,6.9-4.5,12-4.5c4.8,0,8.6,1.4,11.3,4.3 s4.1,6.8,4.1,11.7c0,5-1.5,9-4.3,12c-2.9,3-6.8,4.5-11.8,4.5c-4.8,0-8.6-1.4-11.4-4.2C209.3,51.3,207.9,47.4,207.9,42.6z M215.5,42.3c0,3.2,0.7,5.7,2.2,7.4s3.6,2.6,6.3,2.6c2.6,0,4.7-0.8,6.1-2.6c1.4-1.7,2.1-4.2,2.1-7.6c0-3.3-0.7-5.8-2.1-7.6 c-1.4-1.7-3.5-2.6-6-2.6c-2.7,0-4.7,0.9-6.2,2.7C216.2,36.5,215.5,39,215.5,42.3z M250.5,34.8c0,1,0.3,1.9,1,2.5 c0.7,0.6,2.1,1.3,4.4,2.2c2.9,1.2,5,2.5,6.1,3.9c1.2,1.5,1.8,3.2,1.8,5.3c0,2.9-1.1,5.2-3.4,7c-2.2,1.8-5.3,2.6-9.1,2.6 c-1.3,0-2.7-0.2-4.3-0.5c-1.6-0.3-2.9-0.7-4-1.2v-7.2c1.3,0.9,2.8,1.7,4.3,2.2c1.5,0.5,2.9,0.8,4.2,0.8c1.6,0,2.9-0.2,3.6-0.7 c0.8-0.5,1.2-1.2,1.2-2.3c0-1-0.4-1.8-1.2-2.6c-0.8-0.7-2.4-1.5-4.6-2.4c-2.7-1.1-4.6-2.4-5.7-3.8s-1.7-3.2-1.7-5.4 c0-2.8,1.1-5.1,3.3-6.9c2.2-1.8,5.1-2.7,8.6-2.7c1.1,0,2.3,0.1,3.6,0.4s2.5,0.6,3.4,0.9V34c-1-0.6-2.1-1.2-3.4-1.7 c-1.3-0.5-2.6-0.7-3.8-0.7c-1.4,0-2.5,0.3-3.2,0.8C250.9,33.1,250.5,33.8,250.5,34.8z M266.9,42.6c0-5.1,1.5-9.2,4.3-12.2 c2.9-3,6.9-4.5,12-4.5c4.8,0,8.6,1.4,11.3,4.3s4.1,6.8,4.1,11.7c0,5-1.5,9-4.3,12c-2.9,3-6.8,4.5-11.8,4.5c-4.8,0-8.6-1.4-11.4-4.2 C268.4,51.3,266.9,47.4,266.9,42.6z M274.5,42.3c0,3.2,0.7,5.7,2.2,7.4s3.6,2.6,6.3,2.6c2.6,0,4.7-0.8,6.1-2.6 c1.4-1.7,2.1-4.2,2.1-7.6c0-3.3-0.7-5.8-2.1-7.6c-1.4-1.7-3.5-2.6-6-2.6c-2.7,0-4.7,0.9-6.2,2.7C275.3,36.5,274.5,39,274.5,42.3z M322.9,32.6h-10.9v25h-7.4v-25h-5.2v-6h5.2v-4.3c0-3.2,1.1-5.9,3.2-8s4.8-3.1,8.1-3.1c0.9,0,1.7,0.1,2.4,0.1s1.3,0.2,1.8,0.4v6.3 c-0.2-0.1-0.7-0.3-1.3-0.5c-0.6-0.2-1.3-0.3-2.1-0.3c-1.5,0-2.7,0.5-3.5,1.4c-0.8,0.9-1.2,2.4-1.2,4.2v3.7h10.9v-7l7.3-2.2v9.2h7.4 v6h-7.4v14.5c0,1.9,0.4,3.2,1,4c0.7,0.8,1.8,1.2,3.3,1.2c0.4,0,0.9-0.1,1.5-0.3c0.6-0.2,1.1-0.4,1.5-0.7v6c-0.5,0.3-1.2,0.5-2.3,0.7 c-1.1,0.2-2.1,0.3-3.2,0.3c-3.1,0-5.4-0.8-6.9-2.4c-1.5-1.6-2.3-4.1-2.3-7.4L322.9,32.6L322.9,32.6z"/>
<rect fill="#F25022" width="34.2" height="34.2"/>
<rect x="37.8" fill="#7FBA00" width="34.2" height="34.2"/>
<rect y="37.8" fill="#00A4EF" width="34.2" height="34.2"/>
<rect x="37.8" y="37.8" fill="#FFB900" width="34.2" height="34.2"/>
</svg>`,
  ServiceNow: `<svg class="alliance-logo" width="129.8" height="19.25" enable-background="new 0 0 132.5 20" version="1.1" viewBox="0 0 130.3 19.25" xml:space="preserve" xmlns="http://www.w3.org/2000/svg" xmlns:cc="http://creativecommons.org/ns#" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><title>ServiceNow logo</title><desc>A cloud computing and enterprise software provider based in Santa Clara, California, United States</desc><metadata><rdf:RDF><cc:Work rdf:about=""><dc:format>image/svg+xml</dc:format><dc:type rdf:resource="http://purl.org/dc/dcmitype/StillImage"/><dc:title/></cc:Work></rdf:RDF></metadata><g clip-rule="evenodd" fill-rule="evenodd"><g fill="#293e40">
<path d="m31.7 5.862c-1.4 0-2.6 0.5-3.6 1.3v-1.1h-3.4v12.9h3.5v-8.2c0.5-0.7 1.7-1.6 3.2-1.6 0.5 0 1 0.1 1.4 0.2v-3.3c-0.4-0.1-0.8-0.2-1.1-0.2"/>
<path d="m1.6 14.96c0.9 0.8 2.3 1.3 3.5 1.3 0.9 0 1.7-0.5 1.7-1.1 0-2-6.3-1.3-6.3-5.4 0-2.5 2.4-4 4.9-4 1.7 0 3.5 0.6 4.4 1.3l-1.6 2.5c-0.7-0.4-1.5-0.9-2.5-0.9s-1.7 0.4-1.7 1.1c0 1.7 6.3 1 6.3 5.5 0 2.5-2.4 4-5.1 4-1.8 0-3.7-0.6-5.2-1.7z"/>
<path d="m23.2 12.36c0-3.6-2.5-6.6-6.1-6.6-3.8 0-6.3 3.2-6.3 6.7 0 4.1 2.9 6.7 6.7 6.7 2 0 4-0.8 5.3-2.3l-2-2c-0.6 0.7-1.8 1.5-3.2 1.5-1.8 0-3.3-1.3-3.5-3.1h8.9c0.2-0.2 0.2-0.5 0.2-0.9zm-8.8-1.5c0.2-1.2 1.4-2.3 2.8-2.3s2.4 1.1 2.6 2.3z"/>
<polygon transform="translate(-.3 -.5379)" points="48.3 6.6 42.4 19.5 40 19.5 34.1 6.6 37.6 6.6 41.2 14.5 44.7 6.6"/>
<path d="m50.9-0.03788c1.3 0 2.3 1 2.3 2.2 0 1.3-1 2.2-2.3 2.2s-2.2-1-2.2-2.2c0-1.3 0.9-2.2 2.2-2.2"/>
<rect x="49.2" y="6.062" width="3.5" height="12.9"/>
<path d="m66.9 16.26c-1.5 2-3.4 2.9-5.8 2.9-4 0-6.9-3-6.9-6.7 0-3.8 3-6.7 6.9-6.7 2.3 0 4.2 1.1 5.4 2.6l-2.3 2.1c-0.7-0.9-1.7-1.5-2.9-1.5-2 0-3.5 1.6-3.5 3.5 0 2 1.4 3.5 3.5 3.5 1.4 0 2.5-0.8 3.1-1.7z"/>
<path d="m79.4 16.86c-1.3 1.5-3.3 2.3-5.3 2.3-3.8 0-6.7-2.6-6.7-6.7 0-3.6 2.4-6.7 6.3-6.7 3.5 0 6.1 3 6.1 6.6 0 0.4 0 0.7-0.1 1h-9c0.2 1.8 1.7 3.1 3.5 3.1 1.4 0 2.6-0.8 3.2-1.5zm-3.2-6c-0.1-1.1-1.1-2.3-2.6-2.3-1.4 0-2.6 1.1-2.8 2.3z"/>
<path d="m81.1 18.96v-12.9h3.3v1.1c1-0.8 2.2-1.3 3.6-1.3 1.8 0 3.4 0.8 4.5 2.1 0.8 1 1.4 2.3 1.4 4.5v6.6h-3.5v-6.9c0-1.3-0.3-2-0.8-2.4-0.5-0.5-1.1-0.8-2-0.8-1.4 0-2.6 0.9-3.2 1.6v8.4h-3.3z"/>
</g><path d="m102.8 5.762c-4.2 0-7.5 3.3-7.5 7.5 0 2.2 0.9 4.2 2.3 5.6 0.5 0.5 1.4 0.5 2 0.1 0.8-0.7 2-1.1 3.2-1.1 1.3 0 2.3 0.4 3.2 1.1 0.6 0.5 1.4 0.4 2-0.2 1.4-1.4 2.3-3.3 2.3-5.5-0.1-4.1-3.4-7.5-7.5-7.5m-0.1 11.4c-2.3 0-3.8-1.7-3.8-3.8s1.5-3.8 3.8-3.8 3.8 1.7 3.8 3.8-1.5 3.8-3.8 3.8" fill="#81b5a1"/>
<polygon transform="translate(-.3 -.5379)" points="109.7 6.6 113.1 6.6 115.9 14 118.7 6.6 121.6 6.6 124.3 14 127.1 6.6 130.6 6.6 125.5 19.5 122.9 19.5 120.1 12.2 117.4 19.5 114.8 19.5" fill="#293e40"/>
</g></svg>`,
  Salesforce: `<svg class="alliance-logo" version="1.1" viewBox="0 0 273 191" xmlns="http://www.w3.org/2000/svg" xmlns:cc="http://creativecommons.org/ns#" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#" xmlns:xlink="http://www.w3.org/1999/xlink">
<title>Salesforce.com logo</title>
<desc>A cloud computing company based in San Francisco, California, United States</desc>
 <metadata>
  <rdf:RDF>
   <cc:Work rdf:about="">
    <dc:format>image/svg+xml</dc:format>
    <dc:type rdf:resource="http://purl.org/dc/dcmitype/StillImage"/>
    <dc:title/>
   </cc:Work>
  </rdf:RDF>
 </metadata>
 <defs>
  <path id="a" d="m0.06 0.5h272v190h-272z"/>
 </defs>
 <g fill-rule="evenodd">
  <mask id="b" fill="#fff">
   <use xlink:href="#a"/>
  </mask>
  <path d="m113 21.3c8.78-9.14 21-14.8 34.5-14.8 18 0 33.6 10 42 24.9a58 58 0 0 1 23.7-5.05c32.4 0 58.7 26.5 58.7 59.2s-26.3 59.2-58.7 59.2c-3.96 0-7.82-0.398-11.6-1.15-7.35 13.1-21.4 22-37.4 22a42.7 42.7 0 0 1-18.8-4.32c-7.45 17.5-24.8 29.8-45 29.8-21.1 0-39-13.3-45.9-32a45.1 45.1 0 0 1-9.34 0.972c-25.1 0-45.4-20.6-45.4-45.9 0-17 9.14-31.8 22.7-39.8a52.6 52.6 0 0 1-4.35-21c0-29.2 23.7-52.8 52.9-52.8 17.1 0 32.4 8.15 42 20.8" fill="#00A1E0" mask="url(#b)"/>
  <path d="m39.4 99.3c-0.171 0.446 0.061 0.539 0.116 0.618 0.511 0.37 1.03 0.638 1.55 0.939 2.78 1.47 5.4 1.9 8.14 1.9 5.58 0 9.05-2.97 9.05-7.75v-0.094c0-4.42-3.92-6.03-7.58-7.18l-0.479-0.155c-2.77-0.898-5.16-1.68-5.16-3.5v-0.093c0-1.56 1.4-2.71 3.56-2.71 2.4 0 5.26 0.799 7.09 1.81 0 0 0.542 0.35 0.739-0.173 0.107-0.283 1.04-2.78 1.14-3.06 0.106-0.293-0.08-0.514-0.271-0.628-2.1-1.28-5-2.15-8-2.15l-0.557 2e-3c-5.11 0-8.68 3.09-8.68 7.51v0.095c0 4.66 3.94 6.18 7.62 7.23l0.592 0.184c2.68 0.824 5 1.54 5 3.42v0.094c0 1.73-1.51 3.02-3.93 3.02-0.941 0-3.94-0.016-7.19-2.07-0.393-0.229-0.617-0.394-0.92-0.579-0.16-0.097-0.56-0.272-0.734 0.252l-1.1 3.06m81.7 0c-0.171 0.446 0.061 0.539 0.118 0.618 0.509 0.37 1.03 0.638 1.55 0.939 2.78 1.47 5.4 1.9 8.14 1.9 5.58 0 9.05-2.97 9.05-7.75v-0.094c0-4.42-3.91-6.03-7.58-7.18l-0.479-0.155c-2.77-0.898-5.16-1.68-5.16-3.5v-0.093c0-1.56 1.4-2.71 3.56-2.71 2.4 0 5.25 0.799 7.09 1.81 0 0 0.542 0.35 0.74-0.173 0.106-0.283 1.04-2.78 1.13-3.06 0.107-0.293-0.08-0.514-0.27-0.628-2.1-1.28-5-2.15-8-2.15l-0.558 2e-3c-5.11 0-8.68 3.09-8.68 7.51v0.095c0 4.66 3.94 6.18 7.62 7.23l0.591 0.184c2.69 0.824 5 1.54 5 3.42v0.094c0 1.73-1.51 3.02-3.93 3.02-0.943 0-3.95-0.016-7.19-2.07-0.393-0.229-0.623-0.387-0.921-0.579-0.101-0.064-0.572-0.248-0.733 0.252l-1.1 3.06m55.8-9.36c0 2.7-0.504 4.83-1.49 6.34-0.984 1.49-2.47 2.22-4.54 2.22s-3.55-0.724-4.52-2.21c-0.977-1.5-1.47-3.64-1.47-6.34 0-2.7 0.496-4.82 1.47-6.31 0.968-1.48 2.44-2.19 4.52-2.19s3.56 0.717 4.54 2.19c0.992 1.49 1.49 3.61 1.49 6.31m4.66-5.01c-0.459-1.55-1.17-2.91-2.12-4.05-0.951-1.14-2.15-2.06-3.58-2.72-1.42-0.665-3.1-1-5-1s-3.57 0.337-5 1c-1.42 0.664-2.63 1.58-3.58 2.72-0.948 1.14-1.66 2.5-2.12 4.05-0.455 1.54-0.686 3.22-0.686 5.01 0 1.79 0.231 3.47 0.686 5.01 0.457 1.55 1.17 2.91 2.12 4.05 0.951 1.14 2.16 2.05 3.58 2.7 1.43 0.648 3.11 0.978 5 0.978 1.89 0 3.57-0.33 4.99-0.978 1.42-0.648 2.63-1.56 3.58-2.7 0.949-1.14 1.66-2.5 2.12-4.05 0.454-1.54 0.685-3.22 0.685-5.01 0-1.78-0.231-3.47-0.685-5.01m38.3 12.8c-0.153-0.453-0.595-0.282-0.595-0.282-0.677 0.259-1.4 0.499-2.17 0.619-0.776 0.122-1.64 0.183-2.55 0.183-2.25 0-4.05-0.671-5.33-2-1.29-1.33-2.01-3.47-2-6.37 7e-3 -2.64 0.645-4.62 1.79-6.14 1.13-1.5 2.87-2.28 5.17-2.28 1.92 0 3.39 0.223 4.93 0.705 0 0 0.365 0.159 0.54-0.322 0.409-1.13 0.711-1.94 1.15-3.18 0.124-0.355-0.18-0.505-0.291-0.548-0.604-0.236-2.03-0.623-3.11-0.786-1.01-0.154-2.18-0.234-3.5-0.234-1.96 0-3.7 0.335-5.19 0.999-1.49 0.663-2.75 1.58-3.75 2.72-1 1.14-1.76 2.5-2.27 4.05-0.505 1.54-0.76 3.23-0.76 5.02 0 3.86 1.04 6.99 3.1 9.28 2.06 2.3 5.16 3.46 9.2 3.46 2.39 0 4.84-0.483 6.6-1.18 0 0 0.336-0.162 0.19-0.554l-1.15-3.16m8.15-10.4c0.223-1.5 0.634-2.75 1.28-3.72 0.967-1.48 2.44-2.29 4.51-2.29 2.07 0 3.44 0.814 4.42 2.29 0.65 0.975 0.934 2.27 1.04 3.72l-11.3-2e-3zm15.7-3.3c-0.397-1.49-1.38-3-2.02-3.69-1.02-1.09-2.01-1.86-3-2.28a11.5 11.5 0 0 0-4.52-0.917c-1.97 0-3.76 0.333-5.21 1.01-1.45 0.682-2.67 1.61-3.63 2.77-0.959 1.16-1.68 2.53-2.14 4.1-0.46 1.55-0.692 3.25-0.692 5.03 0 1.82 0.241 3.51 0.715 5.04 0.479 1.54 1.25 2.89 2.29 4.01 1.04 1.13 2.37 2.01 3.97 2.63 1.59 0.615 3.52 0.934 5.73 0.927 4.56-0.015 6.96-1.03 7.94-1.58 0.175-0.098 0.34-0.267 0.134-0.754l-1.03-2.89c-0.158-0.431-0.594-0.275-0.594-0.275-1.13 0.422-2.73 1.18-6.48 1.17-2.45-4e-3 -4.26-0.727-5.4-1.86-1.16-1.16-1.74-2.85-1.83-5.25l15.8 0.012s0.416-4e-3 0.459-0.41c0.017-0.168 0.541-3.24-0.471-6.79zm-142 3.3c0.223-1.5 0.635-2.75 1.28-3.72 0.968-1.48 2.44-2.29 4.51-2.29 2.07 0 3.44 0.814 4.42 2.29 0.649 0.975 0.933 2.27 1.04 3.72l-11.3-2e-3zm15.7-3.3c-0.396-1.49-1.38-3-2.02-3.69-1.02-1.09-2.01-1.86-3-2.28a11.5 11.5 0 0 0-4.52-0.917c-1.97 0-3.76 0.333-5.21 1.01-1.45 0.682-2.67 1.61-3.63 2.77-0.957 1.16-1.68 2.53-2.14 4.1-0.459 1.55-0.69 3.25-0.69 5.03 0 1.82 0.239 3.51 0.716 5.04 0.478 1.54 1.25 2.89 2.28 4.01 1.04 1.13 2.37 2.01 3.97 2.63 1.59 0.615 3.51 0.934 5.73 0.927 4.56-0.015 6.96-1.03 7.94-1.58 0.174-0.098 0.34-0.267 0.133-0.754l-1.03-2.89c-0.159-0.431-0.595-0.275-0.595-0.275-1.13 0.422-2.73 1.18-6.48 1.17-2.44-4e-3 -4.26-0.727-5.4-1.86-1.16-1.16-1.74-2.85-1.83-5.25l15.8 0.012s0.416-4e-3 0.459-0.41c0.017-0.168 0.541-3.24-0.472-6.79zm-49.8 13.6c-0.619-0.494-0.705-0.615-0.91-0.936-0.313-0.483-0.473-1.17-0.473-2.05 0-1.38 0.46-2.38 1.41-3.05-0.01 2e-3 1.36-1.18 4.58-1.14a32 32 0 0 1 4.28 0.365v7.17h2e-3s-2 0.431-4.26 0.567c-3.21 0.193-4.63-0.924-4.62-0.921zm6.28-11.1c-0.64-0.047-1.47-0.07-2.46-0.07-1.35 0-2.66 0.168-3.88 0.498-1.23 0.332-2.34 0.846-3.29 1.53a7.63 7.63 0 0 0-2.29 2.6c-0.559 1.04-0.844 2.26-0.844 3.64 0 1.4 0.243 2.61 0.723 3.6a6.54 6.54 0 0 0 2.06 2.47c0.877 0.638 1.96 1.11 3.21 1.39 1.24 0.283 2.64 0.426 4.18 0.426 1.62 0 3.23-0.136 4.79-0.399a95.1 95.1 0 0 0 3.97-0.772c0.526-0.121 1.11-0.28 1.11-0.28 0.39-0.099 0.36-0.516 0.36-0.516l-9e-3 -14.4c0-3.16-0.844-5.51-2.51-6.96-1.66-1.45-4.09-2.18-7.24-2.18-1.18 0-3.09 0.16-4.23 0.389 0 0-3.44 0.668-4.86 1.78 0 0-0.312 0.192-0.142 0.627l1.12 3c0.139 0.389 0.518 0.256 0.518 0.256s0.119-0.047 0.259-0.13c3.03-1.65 6.87-1.6 6.87-1.6 1.7 0 3.02 0.345 3.9 1.02 0.861 0.661 1.3 1.66 1.3 3.76v0.667c-1.35-0.196-2.6-0.309-2.6-0.309zm127-8.13a0.428 0.428 0 0 0-0.237-0.568c-0.269-0.102-1.61-0.385-2.64-0.449-1.98-0.124-3.08 0.21-4.07 0.654-0.978 0.441-2.06 1.15-2.66 1.97l-2e-3 -1.92c0-0.264-0.187-0.477-0.453-0.477h-4.04c-0.262 0-0.452 0.213-0.452 0.477v23.5a0.48 0.48 0 0 0 0.479 0.479h4.14a0.479 0.479 0 0 0 0.478-0.479v-11.8c0-1.58 0.174-3.15 0.521-4.14 0.342-0.979 0.807-1.76 1.38-2.32a4.79 4.79 0 0 1 1.95-1.17 7.68 7.68 0 0 1 2.12-0.298c0.825 0 1.73 0.212 1.73 0.212 0.304 0.034 0.473-0.152 0.576-0.426 0.271-0.721 1.04-2.88 1.19-3.31" fill="#FFFFFE"/>
  <path d="M162.201 67.548a13.258 13.258 0 0 0-1.559-.37 12.217 12.217 0 0 0-2.144-.166c-2.853 0-5.102.806-6.681 2.398-1.568 1.58-2.635 3.987-3.17 7.154l-.193 1.069h-3.581s-.437-.018-.529.459l-.588 3.28c-.041.314.094.51.514.508h3.486l-3.537 19.743c-.277 1.59-.594 2.898-.945 3.889-.346.978-.684 1.711-1.1 2.243-.403.515-.785.894-1.444 1.115-.544.183-1.17.267-1.856.267-.382 0-.89-.064-1.265-.139-.375-.074-.57-.158-.851-.276 0 0-.409-.156-.57.254-.131.335-1.06 2.89-1.17 3.206-.112.312.045.558.243.629.464.166.809.272 1.441.421.878.207 1.618.22 2.311.22 1.452 0 2.775-.204 3.872-.6 1.104-.399 2.065-1.094 2.915-2.035.919-1.015 1.497-2.078 2.05-3.528.547-1.437 1.013-3.221 1.386-5.3l3.554-20.109h5.196s.438.016.529-.459l.588-3.28c.041-.314-.093-.51-.515-.508h-5.043c.025-.114.254-1.888.833-3.558.247-.713.712-1.288 1.106-1.683a3.273 3.273 0 0 1 1.321-.822 5.48 5.48 0 0 1 1.693-.244c.475 0 .941.057 1.296.131.489.104.679.159.807.197.514.157.583.005.684-.244l1.206-3.312c.124-.356-.178-.506-.29-.55m-70.474 34.117c0 .264-.188.479-.452.479h-4.183c-.265 0-.453-.215-.453-.479V67.997c0-.263.188-.476.453-.476h4.183c.264 0 .452.213.452.476v33.668" fill="#FFFFFE"/>
 </g>
</svg>`,
};

function renderAlliances() {
  const placeholder = '<span class="alliance-empty">—</span>';
  const val = v => (v && v.trim() && v.trim().toUpperCase() !== 'TBD') ? esc(v) : placeholder;
  document.getElementById('allianceGrid').innerHTML = alliances.map(a => `
    <div class="alliance-card">
      <div class="alliance-header">
        ${VENDOR_LOGOS[a.vendor] || `<div class="alliance-name">${esc(a.vendor)}</div>`}
      </div>
      <div class="alliance-row">
        <div class="alliance-label">Internal POC</div>
        <div class="alliance-val">${val(a.internalPOC)}</div>
      </div>
      <div class="alliance-row">
        <div class="alliance-label">Vendor POC</div>
        <div class="alliance-val">${val(a.vendorPOC)}</div>
      </div>
      <div class="alliance-update">
        <div class="alliance-label" style="margin-bottom:4px">Latest Update</div>
        ${a.update ? esc(a.update) : '<span class="alliance-empty">No updates yet</span>'}
      </div>
    </div>
  `).join('');
}

// ===== SORT =====
function sortBy(field) {
  if (sortField === field) sortDir *= -1;
  else { sortField = field; sortDir = 1; }
  document.querySelectorAll('thead th').forEach(th => th.classList.remove('sort-asc','sort-desc'));
  renderTable();
}

// ===== TABS =====
function switchTab(tab, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('tab-accounts').style.display = tab === 'accounts' ? '' : 'none';
  document.getElementById('tab-alliances').style.display = tab === 'alliances' ? '' : 'none';
  if (tab === 'alliances') renderAlliances();
}

// ===== MODAL =====
function openModal(id) {
  editingId = id || null;
  const a = id ? accounts.find(x => x.id === id) : null;
  document.getElementById('modalTitle').textContent = a ? `Edit: ${a.company}` : 'Add Account';
  document.getElementById('modalSubtitle').textContent = a ? a.category : 'E&C Field Service Tracker';
  document.getElementById('deleteBtn').style.display = a ? '' : 'none';

  document.getElementById('f_company').value = a?.company || '';
  document.getElementById('f_category').value = a?.category || '';
  document.getElementById('f_priority').value = a?.priority || '';
  document.getElementById('f_meeting').value = a?.meeting || '';
  document.getElementById('f_status').value = a?.status || '';
  document.getElementById('f_nextsteps').value = a?.nextsteps || '';
  document.getElementById('f_cluster').value = a?.clusterPOC || '';
  document.getElementById('f_keypoc').value = a?.keyPOC || '';
  document.getElementById('f_clientpoc').value = a?.clientPOC || '';
  document.getElementById('f_pipeline').value = a?.pipeline || '';
  document.getElementById('f_existingwork').value = a?.existingWorkDone || '';
  document.getElementById('f_existingvendors').value = a?.existingVendors || '';
  document.getElementById('f_interactedwith').value = a?.interactedWith || '';
  document.getElementById('f_action').value = a?.action || '';

  document.getElementById('modalOverlay').classList.add('open');
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('open');
  editingId = null;
}

function saveAccount() {
  const company = document.getElementById('f_company').value.trim();
  const category = document.getElementById('f_category').value;
  if (!company || !category) { alert('Company name and category are required.'); return; }

  const data = {
    company, category,
    priority: document.getElementById('f_priority').value,
    meeting: document.getElementById('f_meeting').value,
    status: document.getElementById('f_status').value.trim(),
    nextsteps: document.getElementById('f_nextsteps').value.trim(),
    clusterPOC: document.getElementById('f_cluster').value.trim(),
    keyPOC: document.getElementById('f_keypoc').value.trim(),
    clientPOC: document.getElementById('f_clientpoc').value.trim(),
    pipeline: document.getElementById('f_pipeline').value.trim(),
    existingWorkDone: document.getElementById('f_existingwork').value.trim(),
    existingVendors: document.getElementById('f_existingvendors').value.trim(),
    interactedWith: document.getElementById('f_interactedwith').value.trim(),
    action: document.getElementById('f_action').value.trim(),
    lastUpdatedBy: getUsername(),
    lastUpdatedAt: todayStr(),
  };

  if (editingId) {
    const idx = accounts.findIndex(a => a.id === editingId);
    accounts[idx] = { ...accounts[idx], ...data };
  } else {
    accounts.push({ id: nextId++, ...data });
  }

  closeModal();
  renderTable();
}

function confirmDelete() {
  if (!editingId) return;
  if (!confirm('Delete this account? This cannot be undone.')) return;
  accounts = accounts.filter(a => a.id !== editingId);
  closeModal();
  renderTable();
}

// ===== EXPORT =====
function exportCSV() {
  const headers = ['Account Classification','Company Name','Priority Focus','Status / Next Steps','Interacted With?','Next Meeting Date','Cluster POC','Key Account POC','Client POC','Existing Vendors','Existing Work Done','Vendor Pipeline Status','Action Items','Last Updated Date','Last Updated By','Notes'];
  const rows = accounts.map(a => [a.category,a.company,a.priority,a.status,a.interactedWith,a.meeting,a.clusterPOC,a.keyPOC,a.clientPOC,a.existingVendors,a.existingWorkDone,a.pipeline,a.action,a.lastUpdatedAt,a.lastUpdatedBy,a.nextsteps].map(v => `"${(v||'').replace(/"/g,'""')}"`));
  const csv = [headers, ...rows].map(r => r.join(',')).join('\n');
  const blob = new Blob([csv], {type:'text/csv'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'EC_FieldService_Tracker.csv'; a.click();
  URL.revokeObjectURL(url);
}

// Close modal on overlay click
document.getElementById('modalOverlay').addEventListener('click', function(e) {
  if (e.target === this) closeModal();
});

// ===== INIT =====
updateUserChip();
loadCache();
renderTable();