/* ui.js — DOM rendering and event wiring. Expects XLSX and EC globals. */
(function() {
  'use strict';

  // ── State ─────────────────────────────────────────────────────────────────
  var state = {
    accounts:          [],
    alliances:         [],
    activeView:        null,
    selectedAccount:   null,
    filters: {
      search:          '',
      classifications: [],
      priorities:      [],
      alliances:       [],
      interacted:      null,
    },
    sort: { field: 'company', dir: 'asc' },
    chipsBuilt:        false,
  };

  // ── Utility ───────────────────────────────────────────────────────────────

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function classTag(classification) {
    var map = {
      'acp accounts':    'acp',
      'upstream':        'upstream',
      'mid / downstream':'mid',
      'ets / ofs':       'ets',
      'mining':          'mining',
    };
    var key = (classification || '').toLowerCase().trim();
    return 'tag tag--' + (map[key] || 'default');
  }

  function priorityPill(p) {
    var label = p === 'UNSET' ? 'No Priority' : (p || 'UNSET');
    var cls   = ['HIGH','MEDIUM','LOW'].indexOf(p) !== -1 ? p : 'UNSET';
    return '<span class="pill pill--' + cls + '">' + esc(label) + '</span>';
  }

  function showToast(msg, type) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'toast toast--' + (type || 'success') + ' toast--visible';
    setTimeout(function() { t.className = 'toast'; }, 3500);
  }

  // ── Navigation ────────────────────────────────────────────────────────────

  function enableNav() {
    document.querySelectorAll('.nav-item').forEach(function(el) {
      el.classList.remove('nav-item--disabled');
    });
    document.getElementById('emptyState').style.display = 'none';
  }

  function navigate(view) {
    state.activeView = view;
    document.querySelectorAll('.view').forEach(function(el) { el.style.display = 'none'; });
    document.querySelectorAll('.nav-item').forEach(function(el) {
      el.classList.toggle('nav-item--active', el.dataset.view === view);
    });
    document.getElementById('view-' + view).style.display = '';
    closePanel();
    if (view === 'home')           renderHome();
    else if (view === 'accounts')  renderAccounts();
    else if (view === 'alliances') renderAlliances();
  }

  // ── File import ───────────────────────────────────────────────────────────

  function handleFile(file) {
    if (!file) return;
    if (!/\.xlsx?$/i.test(file.name)) {
      showToast('Please upload an .xlsx file', 'error');
      return;
    }
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var wb     = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        var parsed = EC.parseWorkbook(wb);
        if (!parsed.accounts.length) {
          showToast('No accounts found — check the "Lean View" sheet exists', 'error');
          return;
        }
        state.accounts  = parsed.accounts;
        state.alliances = parsed.alliances;
        state.chipsBuilt = false; // rebuild chips on new import
        localStorage.setItem('ec_last_file', file.name);
        document.getElementById('importFilename').textContent = file.name;
        enableNav();
        showToast('Imported ' + parsed.accounts.length + ' accounts' +
          (parsed.alliances.length ? ', ' + parsed.alliances.length + ' alliances' : ''));
        navigate('home');
      } catch (err) {
        showToast('Failed to read file: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ── Slide-in panel ────────────────────────────────────────────────────────

  function openPanel(account) {
    state.selectedAccount = account;
    renderPanel(account);
    document.getElementById('detailPanel').classList.add('open');
    document.getElementById('app').classList.add('panel-open');
  }

  function closePanel() {
    state.selectedAccount = null;
    document.getElementById('detailPanel').classList.remove('open');
    document.getElementById('app').classList.remove('panel-open');
    document.querySelectorAll('.row--selected').forEach(function(el) {
      el.classList.remove('row--selected');
    });
  }

  function renderPanel(account) {
    function pocLine(val) {
      var emails = EC.extractEmails(val);
      if (emails.length) {
        return emails.map(function(e) {
          return '<a href="mailto:' + esc(e) + '">' + esc(e) + '</a>';
        }).join(', ');
      }
      return esc(val);
    }
    function field(label, val) {
      if (!val) return '';
      return '<div class="panel-field">' +
        '<div class="panel-field-label">' + esc(label) + '</div>' +
        '<div class="panel-field-value">' + esc(val) + '</div>' +
        '</div>';
    }
    function pocField(label, val) {
      if (!val) return '';
      return '<div class="panel-field">' +
        '<div class="panel-field-label">' + esc(label) + '</div>' +
        '<div class="panel-field-value">' + pocLine(val) + '</div>' +
        '</div>';
    }

    var bullets = EC.parsePipelineBullets(account.pipeline);
    var actions = EC.parseActionItems(account.action);
    var alTags  = EC.splitAlliance(account.alliance).map(function(al) {
      return '<span class="tag tag--default">' + esc(al) + '</span>';
    }).join(' ');

    var leftCol =
      '<div class="panel-section-title">Details</div>' +
      field('Next Meeting', EC.formatDate(account.nextMeeting)) +
      pocField('Cluster POC', account.clusterPOC) +
      pocField('Key Account POC', account.keyPOC) +
      pocField('Client POC', account.clientPOC) +
      field('Existing Vendors', account.existingVendors) +
      field('Existing Work', account.existingWork);

    var rightCol = '';
    if (account.status) {
      rightCol += '<div class="panel-section-title">Status / Next Steps</div>' +
        '<div class="panel-status-text">' + esc(account.status) + '</div>';
    }
    if (bullets.length) {
      rightCol += '<div class="panel-section-title">Pipeline</div>' +
        '<ul class="panel-bullets">' +
        bullets.map(function(b) { return '<li>' + esc(b) + '</li>'; }).join('') +
        '</ul>';
    }
    if (actions.length) {
      rightCol += '<div class="panel-section-title">Action Items</div>' +
        actions.map(function(a) {
          return '<div class="panel-action-item">' +
            '<span class="panel-action-date">' + esc(a.date || '') + '</span>' +
            '<span>' + esc(a.text) + '</span>' +
            '</div>';
        }).join('');
    }

    var footer = '';
    if (alTags || account.lastUpdatedDate || account.lastUpdatedBy) {
      footer = '<div class="panel-footer">' +
        (alTags ? '<div style="margin-bottom:5px">' + alTags + '</div>' : '') +
        'Last updated ' + EC.formatDate(account.lastUpdatedDate) +
        (account.lastUpdatedBy ? ' by ' + esc(account.lastUpdatedBy) : '') +
        '</div>';
    }

    document.getElementById('panelBody').innerHTML =
      '<div class="panel-company">' + esc(account.company) + '</div>' +
      '<div class="panel-badges">' +
        '<span class="' + classTag(account.classification) + '">' + esc(account.classification || '—') + '</span>' +
        priorityPill(account.priority) +
        (account.interacted
          ? '<span class="pill" style="background:#E0F2FE;color:#0369A1">✓ Engaged</span>'
          : '') +
      '</div>' +
      '<div class="panel-cols">' +
        '<div>' + leftCol + '</div>' +
        '<div>' + rightCol + '</div>' +
      '</div>' +
      footer;
  }

  // ── Shared table builder ──────────────────────────────────────────────────

  var COMPACT_COLS = ['company','classification','priority','nextMeeting','lastUpdatedDate'];
  var FULL_COLS    = ['company','classification','priority','interacted','nextMeeting','lastUpdatedDate','alliance'];
  var COL_LABELS   = {
    company: 'Company', classification: 'Classification', priority: 'Priority',
    interacted: 'Engaged', nextMeeting: 'Next Meeting',
    lastUpdatedDate: 'Last Updated', alliance: 'Alliance',
  };

  function buildTable(accounts, compact) {
    var cols = compact ? COMPACT_COLS : FULL_COLS;
    if (!accounts.length) {
      return '<table class="data-table"><tbody><tr><td class="empty-table-cell" colspan="' +
        cols.length + '">No accounts found</td></tr></tbody></table>';
    }

    var thead = '<thead><tr>' + cols.map(function(c) {
      var sortCls = (!compact && state.sort.field === c) ? ' class="sort-' + state.sort.dir + '"' : '';
      return '<th data-col="' + c + '"' + sortCls + '>' + COL_LABELS[c] + '</th>';
    }).join('') + '</tr></thead>';

    var tbody = '<tbody>' + accounts.map(function(a, idx) {
      var cells = cols.map(function(c) {
        if (c === 'priority') {
          return '<td>' + priorityPill(a.priority) + '</td>';
        }
        if (c === 'classification') {
          return '<td><span class="' + classTag(a.classification) + '">' + esc(a.classification || '—') + '</span></td>';
        }
        if (c === 'interacted') {
          return '<td style="color:' + (a.interacted ? 'var(--green)' : 'var(--muted)') + ';font-weight:600">' +
            (a.interacted ? '✓' : '—') + '</td>';
        }
        if (c === 'nextMeeting' || c === 'lastUpdatedDate') {
          return '<td>' + esc(EC.formatDate(a[c])) + '</td>';
        }
        if (c === 'alliance') {
          var tags = EC.splitAlliance(a.alliance).map(function(t) {
            return '<span class="tag tag--default">' + esc(t) + '</span>';
          }).join(' ');
          return '<td>' + (tags || '<span style="color:var(--muted)">—</span>') + '</td>';
        }
        return '<td>' + esc(a[c] || '—') + '</td>';
      }).join('');
      return '<tr data-idx="' + idx + '">' + cells + '</tr>';
    }).join('') + '</tbody>';

    return '<table class="data-table">' + thead + tbody + '</table>';
  }

  function wireTableClicks(containerId, accounts) {
    var wrap = document.getElementById(containerId);
    if (!wrap) return;

    // Row click → panel
    wrap.querySelectorAll('tbody tr[data-idx]').forEach(function(tr) {
      tr.addEventListener('click', function() {
        var account = accounts[parseInt(tr.dataset.idx, 10)];
        if (!account) return;
        document.querySelectorAll('.row--selected').forEach(function(r) {
          r.classList.remove('row--selected');
        });
        tr.classList.add('row--selected');
        openPanel(account);
      });
    });

    // Sort on th click (full table only)
    wrap.querySelectorAll('th[data-col]').forEach(function(th) {
      th.addEventListener('click', function() {
        var col = th.dataset.col;
        if (state.sort.field === col) {
          state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
        } else {
          state.sort.field = col;
          state.sort.dir = 'asc';
        }
        applyFilters();
      });
    });
  }

  // ── Home view ─────────────────────────────────────────────────────────────

  function renderHome() {
    var kpis = EC.computeKPIs(state.accounts);
    renderKPIs(kpis);
    renderCharts();
    renderHomeTable();
  }

  function renderKPIs(kpis) {
    var defs = [
      {
        label: 'Total Accounts',
        value: kpis.total,
        sub:   kpis.highPriority + ' high priority',
        icon:  '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg>',
      },
      {
        label: 'High Priority',
        value: kpis.highPriority,
        sub:   kpis.highPct + '% of total',
        icon:  '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>',
      },
      {
        label: 'Engaged',
        value: kpis.engagedPct + '%',
        sub:   (100 - kpis.engagedPct) + '% not yet engaged',
        icon:  '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg>',
      },
      {
        label: 'Meetings ≤30d',
        value: kpis.meetingsIn30,
        sub:   kpis.nextMeeting ? 'next: ' + EC.formatDate(kpis.nextMeeting) : 'none scheduled',
        icon:  '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',
      },
    ];
    document.getElementById('kpiGrid').innerHTML = defs.map(function(k) {
      return '<div class="kpi-card">' +
        '<div class="kpi-icon">' + k.icon + '</div>' +
        '<div class="kpi-value">' + esc(String(k.value)) + '</div>' +
        '<div class="kpi-label">' + esc(k.label) + '</div>' +
        '<div class="kpi-sub">' + esc(k.sub) + '</div>' +
        '</div>';
    }).join('');
  }

  function renderCharts() {
    var total      = state.accounts.length || 1;
    var byClass    = EC.groupBy(state.accounts, 'classification');
    var byPriority = EC.groupBy(state.accounts, 'priority');
    var byAlliance = EC.groupByAlliance(state.accounts);

    function barRows(data, filterKey) {
      var entries = Object.keys(data).map(function(k) { return [k, data[k]]; });
      entries.sort(function(a, b) { return b[1] - a[1]; });
      return entries.map(function(pair) {
        var label = pair[0], count = pair[1];
        var pct   = Math.round(count / total * 100);
        return '<div class="chart-bar-row" data-filter-key="' + esc(filterKey) + '" data-filter-val="' + esc(label) + '">' +
          '<span class="chart-bar-label" title="' + esc(label) + '">' + esc(label) + '</span>' +
          '<div class="chart-bar-track"><div class="chart-bar-fill" style="width:' + pct + '%"></div></div>' +
          '<span class="chart-bar-count">' + count + '</span>' +
          '</div>';
      }).join('');
    }

    // Priority stacked bar
    var priColors = { HIGH: 'var(--green)', MEDIUM: 'var(--amber)', LOW: '#9CA3AF', UNSET: '#E5E7EB' };
    var priOrder  = ['HIGH', 'MEDIUM', 'LOW', 'UNSET'];
    var stackedSegs = priOrder.map(function(p) {
      var count = byPriority[p] || 0;
      var w     = Math.round(count / total * 100);
      return w > 0
        ? '<div class="stacked-seg" title="' + p + ': ' + count + '" style="width:' + w + '%;background:' + priColors[p] + '"></div>'
        : '';
    }).join('');
    var priLegend = priOrder.map(function(p) {
      var count = byPriority[p] || 0;
      return count > 0
        ? '<div class="stacked-legend-item"><div class="stacked-legend-dot" style="background:' + priColors[p] + '"></div>' + p + ' ' + count + '</div>'
        : '';
    }).join('');

    var allianceHtml = Object.keys(byAlliance).length
      ? barRows(byAlliance, 'alliances')
      : '<div style="color:var(--muted);font-size:12px;line-height:1.5">No alliance tags — add an Alliance column in Excel</div>';

    document.getElementById('chartsRow').innerHTML =
      '<div class="chart-card">' +
        '<div class="chart-title">By Classification</div>' +
        barRows(byClass, 'classifications') +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Priority Mix</div>' +
        '<div class="stacked-bar">' + stackedSegs + '</div>' +
        '<div class="stacked-legend">' + priLegend + '</div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Alliance Split</div>' +
        allianceHtml +
      '</div>';

    // Click chart bar → navigate to Accounts with that filter pre-set
    document.getElementById('chartsRow').querySelectorAll('.chart-bar-row').forEach(function(el) {
      el.addEventListener('click', function() {
        var key = el.dataset.filterKey;
        var val = el.dataset.filterVal;
        if (key && val) {
          state.filters = { search: '', classifications: [], priorities: [], alliances: [], interacted: null };
          if (key === 'classifications') state.filters.classifications = [val];
          else if (key === 'priorities')  state.filters.priorities  = [val];
          else if (key === 'alliances')   state.filters.alliances   = [val];
          state.chipsBuilt = false;
          navigate('accounts');
        }
      });
    });
  }

  function renderHomeTable() {
    var high = EC.sortAccounts(
      state.accounts.filter(function(a) { return a.priority === 'HIGH'; }),
      'nextMeeting', 'asc'
    );
    document.getElementById('homeTable').innerHTML = buildTable(high, true);
    wireTableClicks('homeTable', high);
  }

  // ── Accounts view ─────────────────────────────────────────────────────────

  function renderAccounts() {
    buildFilterChips();
    syncChipState();
    document.getElementById('searchInput').value = state.filters.search;
    applyFilters();
  }

  function buildFilterChips() {
    if (state.chipsBuilt) return;
    state.chipsBuilt = true;
    var chips = document.getElementById('filterChips');
    chips.innerHTML = '';

    var priorities      = ['HIGH', 'MEDIUM', 'LOW', 'UNSET'];
    var classifications = [];
    var seen = {};
    state.accounts.forEach(function(a) {
      if (a.classification && !seen[a.classification]) {
        seen[a.classification] = true;
        classifications.push(a.classification);
      }
    });
    classifications.sort();
    var allianceNames = Object.keys(EC.groupByAlliance(state.accounts)).sort();

    function addChips(values, filterKey) {
      values.forEach(function(v) {
        var el = document.createElement('button');
        el.className = 'chip';
        el.textContent = v;
        el.dataset.value = v;
        el.dataset.filter = filterKey;
        el.addEventListener('click', function() {
          var arr = state.filters[filterKey];
          var idx = arr.indexOf(v);
          if (idx === -1) arr.push(v); else arr.splice(idx, 1);
          el.classList.toggle('chip--active', arr.indexOf(v) !== -1);
          applyFilters();
        });
        chips.appendChild(el);
      });
    }

    addChips(priorities, 'priorities');
    addChips(classifications, 'classifications');
    if (allianceNames.length) addChips(allianceNames, 'alliances');
  }

  function syncChipState() {
    // Reflect state.filters back onto chip active classes (e.g. after chart click nav)
    document.querySelectorAll('#filterChips .chip').forEach(function(el) {
      var arr = state.filters[el.dataset.filter] || [];
      el.classList.toggle('chip--active', arr.indexOf(el.dataset.value) !== -1);
    });
  }

  function wireSearch() {
    var input = document.getElementById('searchInput');
    if (input.dataset.wired) return;
    input.dataset.wired = '1';
    var timer;
    input.addEventListener('input', function() {
      clearTimeout(timer);
      timer = setTimeout(function() {
        state.filters.search = input.value;
        applyFilters();
      }, 200);
    });
  }

  function wireClearFilters() {
    var btn = document.getElementById('clearFilters');
    if (btn.dataset.wired) return;
    btn.dataset.wired = '1';
    btn.addEventListener('click', function() {
      state.filters = { search: '', classifications: [], priorities: [], alliances: [], interacted: null };
      document.getElementById('searchInput').value = '';
      document.querySelectorAll('#filterChips .chip--active').forEach(function(el) {
        el.classList.remove('chip--active');
      });
      applyFilters();
    });
  }

  function applyFilters() {
    var filtered = EC.filterAccounts(state.accounts, state.filters);
    var sorted   = EC.sortAccounts(filtered, state.sort.field, state.sort.dir);

    var countEl = document.getElementById('recordCount');
    if (countEl) {
      countEl.textContent = sorted.length === state.accounts.length
        ? sorted.length + ' accounts'
        : sorted.length + ' of ' + state.accounts.length + ' accounts';
    }

    var hasFilters = state.filters.search ||
      state.filters.classifications.length ||
      state.filters.priorities.length ||
      state.filters.alliances.length ||
      state.filters.interacted !== null;

    var clearBtn = document.getElementById('clearFilters');
    if (clearBtn) clearBtn.style.display = hasFilters ? '' : 'none';

    document.getElementById('accountsTable').innerHTML = buildTable(sorted, false);
    wireTableClicks('accountsTable', sorted);

    // Restore sort indicators
    document.querySelectorAll('#accountsTable th[data-col]').forEach(function(th) {
      if (state.sort.field === th.dataset.col) {
        th.className = 'sort-' + state.sort.dir;
      } else {
        th.className = '';
      }
    });
  }

  // ── Alliances view ────────────────────────────────────────────────────────

  function renderAlliances() {
    document.getElementById('allianceCards').style.display = '';
    document.getElementById('allianceDetail').style.display = 'none';

    if (!state.alliances.length) {
      document.getElementById('allianceCards').innerHTML =
        '<div style="color:var(--muted);padding:32px;font-size:14px">' +
        'No Alliances sheet found in this workbook.</div>';
      return;
    }

    var allianceCounts = EC.groupByAlliance(state.accounts);

    document.getElementById('allianceCards').innerHTML = state.alliances.map(function(al, idx) {
      var count     = allianceCounts[al.vendor] || 0;
      var highCount = state.accounts.filter(function(a) {
        return a.priority === 'HIGH' && EC.splitAlliance(a.alliance).indexOf(al.vendor) !== -1;
      }).length;
      var email = EC.extractEmails(al.internalPOC)[0] || al.internalPOC || '—';
      return '<div class="alliance-card" data-idx="' + idx + '">' +
        '<div class="alliance-card-name">' + esc(al.vendor) + '</div>' +
        '<div class="alliance-card-stat">' + count + ' tagged accounts · ' + highCount + ' high priority</div>' +
        '<div class="alliance-card-stat">' + esc(email) + '</div>' +
        (al.nextAction ? '<div class="alliance-card-action">' + esc(al.nextAction) + '</div>' : '') +
        '</div>';
    }).join('');

    document.getElementById('allianceCards').querySelectorAll('.alliance-card').forEach(function(card) {
      card.addEventListener('click', function() {
        showAllianceDetail(state.alliances[parseInt(card.dataset.idx, 10)]);
      });
    });
  }

  function showAllianceDetail(al) {
    document.getElementById('allianceCards').style.display = 'none';
    document.getElementById('allianceDetail').style.display = '';

    var tagged = state.accounts.filter(function(a) {
      return EC.splitAlliance(a.alliance).indexOf(al.vendor) !== -1;
    });
    tagged = EC.sortAccounts(tagged, 'priority', 'asc');

    function pocBlock(label, val) {
      if (!val) return '';
      var emails = EC.extractEmails(val);
      var display = emails.length
        ? emails.map(function(e) { return '<a href="mailto:' + esc(e) + '">' + esc(e) + '</a>'; }).join(', ')
        : esc(val);
      return '<div class="panel-field">' +
        '<div class="panel-field-label">' + esc(label) + '</div>' +
        '<div class="panel-field-value">' + display + '</div>' +
        '</div>';
    }

    var rows = tagged.map(function(a, i) {
      return '<tr data-idx="' + i + '">' +
        '<td>' + esc(a.company) + '</td>' +
        '<td>' + priorityPill(a.priority) + '</td>' +
        '<td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--muted)">' +
          esc((a.status || '—').split('\n')[0]) + '</td>' +
        '<td>' + esc(EC.formatDate(a.lastUpdatedDate)) + '</td>' +
        '</tr>';
    }).join('');

    document.getElementById('allianceDetailContent').innerHTML =
      '<h2 style="font-size:20px;font-weight:700;margin-bottom:20px">' + esc(al.vendor) + '</h2>' +
      pocBlock('Internal POC', al.internalPOC) +
      pocBlock('Vendor POC', al.vendorPOC) +
      (al.comment ? '<div class="panel-field"><div class="panel-field-label">Comment</div><div class="panel-field-value">' + esc(al.comment) + '</div></div>' : '') +
      (al.nextAction ? '<div class="panel-field"><div class="panel-field-label">Next Action</div><div class="panel-field-value">' + esc(al.nextAction) + '</div></div>' : '') +
      (tagged.length
        ? '<div style="margin-top:28px">' +
            '<h3 style="font-size:15px;font-weight:600;margin-bottom:12px">Tagged Accounts (' + tagged.length + ')</h3>' +
            '<table class="data-table"><thead><tr>' +
              '<th>Company</th><th>Priority</th><th>Status</th><th>Last Updated</th>' +
            '</tr></thead><tbody>' + rows + '</tbody></table>' +
          '</div>'
        : '<div style="color:var(--muted);margin-top:16px;font-size:14px">No accounts tagged with this alliance.</div>');

    document.getElementById('allianceDetailContent').querySelectorAll('tbody tr').forEach(function(tr) {
      tr.style.cursor = 'pointer';
      tr.addEventListener('click', function() {
        openPanel(tagged[parseInt(tr.dataset.idx, 10)]);
      });
    });
  }

  // ── Boot ──────────────────────────────────────────────────────────────────

  function init() {
    // Import button
    document.getElementById('importBtn').addEventListener('click', function() {
      document.getElementById('fileInput').click();
    });
    document.getElementById('fileInput').addEventListener('change', function(e) {
      handleFile(e.target.files[0]);
      e.target.value = '';
    });

    // Panel close
    document.getElementById('panelClose').addEventListener('click', closePanel);
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') closePanel();
    });

    // View all link
    document.getElementById('viewAllLink').addEventListener('click', function(e) {
      e.preventDefault();
      state.filters = { search: '', classifications: [], priorities: [], alliances: [], interacted: null };
      state.chipsBuilt = false;
      navigate('accounts');
    });

    // Alliance back
    document.getElementById('allianceBack').addEventListener('click', renderAlliances);

    // Sidebar nav
    document.querySelectorAll('.nav-item').forEach(function(el) {
      el.addEventListener('click', function(e) {
        e.preventDefault();
        if (!el.classList.contains('nav-item--disabled')) navigate(el.dataset.view);
      });
    });

    // Drag-and-drop on main content area
    var main = document.getElementById('main');
    main.addEventListener('dragover', function(e) {
      e.preventDefault();
      main.classList.add('drag-over');
    });
    main.addEventListener('dragleave', function(e) {
      if (!main.contains(e.relatedTarget)) main.classList.remove('drag-over');
    });
    main.addEventListener('drop', function(e) {
      e.preventDefault();
      main.classList.remove('drag-over');
      handleFile(e.dataTransfer.files[0]);
    });

    // Wire search + clear (these are always in DOM)
    wireSearch();
    wireClearFilters();

    // Restore last filename from localStorage
    var last = localStorage.getItem('ec_last_file');
    if (last) document.getElementById('importFilename').textContent = last;
  }

  document.addEventListener('DOMContentLoaded', init);
})();
