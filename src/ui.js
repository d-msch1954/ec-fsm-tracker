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
      search:             '',
      clusters:           [],
      priorities:         [],
      engagementStatuses: [],
    },
    sort: { field: 'company', dir: 'asc' },
    sidebarBuilt:      false,
  };

  // ── Utility ───────────────────────────────────────────────────────────────

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function clusterTag(cluster) {
    var map = {
      'acp':            'acp',
      'upstream':       'upstream',
      'mid-downstream': 'mid',
      'mid / downstream': 'mid',
      'ets-ofs':        'ets',
      'ets / ofs':      'ets',
      'mining':         'mining',
    };
    var key = (cluster || '').toLowerCase().trim();
    return 'tag tag--' + (map[key] || 'default');
  }

  function statusBadge(status) {
    var map = {
      'active':      'status--active',
      'monitoring':  'status--monitoring',
      'on hold':     'status--hold',
      'restricted':  'status--restricted',
      'not started': 'status--notstarted',
    };
    var key = (status || '').toLowerCase().trim();
    return '<span class="status-badge ' + (map[key] || 'status--notstarted') + '">' + esc(status || '—') + '</span>';
  }

  function pipelineStatusBadge(val) {
    if (!val || val === '—') return '<span style="color:var(--muted)">—</span>';
    var lc = val.toLowerCase();
    if (lc === 'yes' || lc === 'true') {
      return '<span class="pill" style="background:#D1FAE5;color:#065F46">' + esc(val) + '</span>';
    }
    if (lc.indexOf('block') !== -1 || lc.indexOf('competitor') !== -1) {
      return '<span class="pill" style="background:#FEE2E2;color:#991B1B">' + esc(val) + '</span>';
    }
    return '<span style="color:var(--muted);font-size:12px">' + esc(val) + '</span>';
  }

  function logoHtml(company, size) {
    size = size || 32;
    var domain   = EC.inferDomain(company);
    var initials = EC.companyInitials(company);
    var color    = EC.nameToColor(company);
    var avatarCss = 'width:' + size + 'px;height:' + size + 'px;border-radius:6px;background:' + color +
      ';color:#fff;display:inline-flex;align-items:center;justify-content:center;' +
      'font-size:' + Math.round(size * 0.38) + 'px;font-weight:700;flex-shrink:0;letter-spacing:-0.5px';
    var avatarHtml = '<span style="' + avatarCss + '">' + esc(initials) + '</span>';
    if (!domain) return avatarHtml;
    return '<img src="https://logo.clearbit.com/' + esc(domain) + '"' +
      ' style="width:' + size + 'px;height:' + size + 'px;border-radius:6px;object-fit:contain;' +
      'background:#fff;border:1px solid var(--border);flex-shrink:0"' +
      ' loading="lazy" alt=""' +
      ' onerror="this.onerror=null;this.outerHTML=this.dataset.fb"' +
      ' data-fb="' + esc(avatarHtml) + '">';
  }

  function priorityPill(p) {
    var label = p === 'UNSET' ? 'TBD' : (p || 'TBD');
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
          showToast('No accounts found — check the sheet exists with correct column headers', 'error');
          return;
        }
        state.accounts  = parsed.accounts;
        state.alliances = parsed.alliances;
        state.sidebarBuilt = false;
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
    function htmlField(label, html) {
      if (!html) return '';
      return '<div class="panel-field">' +
        '<div class="panel-field-label">' + esc(label) + '</div>' +
        '<div class="panel-field-value">' + html + '</div>' +
        '</div>';
    }

    var leftCol =
      '<div class="panel-section-title">Engagement</div>' +
      htmlField('Status', statusBadge(account.engagementStatus)) +
      field('Type', account.engagementType) +
      field('Action Owner', account.actionOwner) +
      (account.additionalInfo
        ? '<div class="panel-field"><div class="panel-field-label">Additional Info</div>' +
          '<div class="panel-field-value"><span class="pill" style="background:#FEF3C7;color:#92400E">Needed</span></div></div>'
        : '') +
      '<div class="panel-section-title" style="margin-top:16px">People</div>' +
      pocField('Deloitte POC', account.deloittePOC) +
      pocField('Cluster POC', account.clusterPOC) +
      '<div class="panel-section-title" style="margin-top:16px">Existing</div>' +
      field('Vendors', account.existingVendors) +
      field('Work Done', account.existingWork);

    var rightCol = '';
    if (account.nextSteps) {
      rightCol += '<div class="panel-section-title">Next Steps</div>' +
        '<div class="panel-status-text">' + esc(account.nextSteps) + '</div>';
    }
    if (account.pipelineStatus || account.pipelineNotes) {
      rightCol += '<div class="panel-section-title" style="margin-top:16px">Pipeline</div>';
      if (account.pipelineStatus) {
        rightCol += '<div style="margin-bottom:8px">' + pipelineStatusBadge(account.pipelineStatus) + '</div>';
      }
      if (account.pipelineNotes) {
        rightCol += '<div class="panel-status-text">' + esc(account.pipelineNotes) + '</div>';
      }
    }
    if (account.notes) {
      rightCol += '<div class="panel-section-title" style="margin-top:16px">Notes</div>' +
        '<div class="panel-status-text">' + esc(account.notes) + '</div>';
    }

    var footer = account.lastUpdated
      ? '<div class="panel-footer">Last updated ' + EC.formatDate(account.lastUpdated) + '</div>'
      : '';

    document.getElementById('panelBody').innerHTML =
      '<div class="company-cell" style="gap:12px;margin-bottom:8px">' +
        logoHtml(account.company, 40) +
        '<div class="panel-company" style="margin-bottom:0">' + esc(account.company) + '</div>' +
      '</div>' +
      '<div class="panel-badges">' +
        '<span class="' + clusterTag(account.cluster) + '">' + esc(account.cluster || '—') + '</span>' +
        priorityPill(account.priority) +
      '</div>' +
      '<div class="panel-cols">' +
        '<div>' + leftCol + '</div>' +
        '<div>' + rightCol + '</div>' +
      '</div>' +
      footer;
  }

  // ── Shared table builder ──────────────────────────────────────────────────

  var COMPACT_COLS = ['company', 'cluster', 'priority', 'engagementStatus', 'actionOwner'];
  var FULL_COLS    = ['company', 'cluster', 'priority', 'engagementStatus', 'engagementType', 'actionOwner', 'pipelineStatus', 'lastUpdated'];
  var COL_LABELS   = {
    company:          'Company',
    cluster:          'Cluster',
    priority:         'Priority',
    engagementStatus: 'Status',
    engagementType:   'Type',
    actionOwner:      'Action Owner',
    pipelineStatus:   'Pipeline',
    lastUpdated:      'Last Updated',
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
        if (c === 'cluster') {
          return '<td><span class="' + clusterTag(a.cluster) + '">' + esc(a.cluster || '—') + '</span></td>';
        }
        if (c === 'engagementStatus') {
          return '<td>' + statusBadge(a.engagementStatus) + '</td>';
        }
        if (c === 'pipelineStatus') {
          return '<td>' + pipelineStatusBadge(a.pipelineStatus) + '</td>';
        }
        if (c === 'lastUpdated') {
          return '<td>' + esc(EC.formatDate(a.lastUpdated)) + '</td>';
        }
        if (c === 'company') {
          return '<td><div class="company-cell">' + logoHtml(a.company, 32) + '<span>' + esc(a.company || '—') + '</span></div></td>';
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
        label: 'Active Engagements',
        value: kpis.active,
        sub:   kpis.activePct + '% of total',
        icon:  '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
      },
      {
        label: 'Needs Info',
        value: kpis.needsInfo,
        sub:   'additional info required',
        icon:  '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>',
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
    var total         = state.accounts.length || 1;
    var byCluster     = EC.groupBy(state.accounts, 'cluster');
    var byPriority    = EC.groupBy(state.accounts, 'priority');
    var byStatus      = EC.groupBy(state.accounts, 'engagementStatus');

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

    document.getElementById('chartsRow').innerHTML =
      '<div class="chart-card">' +
        '<div class="chart-title">By Cluster</div>' +
        barRows(byCluster, 'clusters') +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Priority Mix</div>' +
        '<div class="stacked-bar">' + stackedSegs + '</div>' +
        '<div class="stacked-legend">' + priLegend + '</div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Engagement Status</div>' +
        barRows(byStatus, 'engagementStatuses') +
      '</div>';

    document.getElementById('chartsRow').querySelectorAll('.chart-bar-row').forEach(function(el) {
      el.addEventListener('click', function() {
        var key = el.dataset.filterKey;
        var val = el.dataset.filterVal;
        if (key && val) {
          state.filters = { search: '', clusters: [], priorities: [], engagementStatuses: [] };
          if (key === 'clusters')            state.filters.clusters            = [val];
          else if (key === 'priorities')     state.filters.priorities          = [val];
          else if (key === 'engagementStatuses') state.filters.engagementStatuses = [val];
          state.sidebarBuilt = false;
          navigate('accounts');
        }
      });
    });
  }

  function renderHomeTable() {
    var high = EC.sortAccounts(
      state.accounts.filter(function(a) { return a.priority === 'HIGH'; }),
      'company', 'asc'
    );
    document.getElementById('homeTable').innerHTML = buildTable(high, true);
    wireTableClicks('homeTable', high);
  }

  // ── Accounts view ─────────────────────────────────────────────────────────

  function renderAccounts() {
    buildFilterSidebar();
    syncSidebarState();
    document.getElementById('searchInput').value = state.filters.search;
    applyFilters();
  }

  function buildFilterSidebar() {
    if (state.sidebarBuilt) return;
    state.sidebarBuilt = true;
    var sidebar = document.getElementById('filterSidebar');
    sidebar.innerHTML = '';

    var priorities = ['HIGH', 'MEDIUM', 'LOW', 'UNSET'];
    var clusters = [], statuses = [];
    var seenC = {}, seenS = {};
    state.accounts.forEach(function(a) {
      if (a.cluster && !seenC[a.cluster]) { seenC[a.cluster] = true; clusters.push(a.cluster); }
      if (a.engagementStatus && !seenS[a.engagementStatus]) { seenS[a.engagementStatus] = true; statuses.push(a.engagementStatus); }
    });
    clusters.sort();
    statuses.sort();

    var header = document.createElement('div');
    header.className = 'filter-sidebar-header';
    header.innerHTML = '<span class="filter-sidebar-title">Filters</span>' +
      '<button class="filter-clear-all" id="clearAllFilters">Clear all</button>';
    sidebar.appendChild(header);

    document.getElementById('clearAllFilters').addEventListener('click', function() {
      state.filters = { search: '', clusters: [], priorities: [], engagementStatuses: [] };
      document.getElementById('searchInput').value = '';
      syncSidebarState();
      applyFilters();
    });

    function addSection(title, values, filterKey, labelFn) {
      var section = document.createElement('div');
      section.className = 'filter-section';
      var titleEl = document.createElement('div');
      titleEl.className = 'filter-section-title';
      titleEl.dataset.section = filterKey;
      titleEl.textContent = title;
      section.appendChild(titleEl);

      values.forEach(function(v) {
        var label = document.createElement('label');
        label.className = 'filter-check';
        var cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.dataset.value = v;
        cb.dataset.filter = filterKey;
        cb.addEventListener('change', function() {
          var arr = state.filters[filterKey];
          var idx = arr.indexOf(v);
          if (cb.checked && idx === -1) arr.push(v);
          else if (!cb.checked && idx !== -1) arr.splice(idx, 1);
          syncSidebarState();
          applyFilters();
        });
        label.appendChild(cb);
        label.appendChild(document.createTextNode(labelFn ? labelFn(v) : v));
        section.appendChild(label);
      });

      sidebar.appendChild(section);
    }

    addSection('Priority', priorities, 'priorities', function(v) { return v === 'UNSET' ? 'TBD' : v; });
    addSection('Cluster', clusters, 'clusters', null);
    addSection('Engagement Status', statuses, 'engagementStatuses', null);
  }

  function syncSidebarState() {
    var sidebar = document.getElementById('filterSidebar');
    if (!sidebar) return;

    sidebar.querySelectorAll('input[type="checkbox"]').forEach(function(cb) {
      var arr = state.filters[cb.dataset.filter] || [];
      cb.checked = arr.indexOf(cb.dataset.value) !== -1;
    });

    sidebar.querySelectorAll('.filter-section-title').forEach(function(titleEl) {
      var filterKey = titleEl.dataset.section;
      var count = (state.filters[filterKey] || []).length;
      var badge = titleEl.querySelector('.filter-badge');
      if (count > 0) {
        if (!badge) {
          badge = document.createElement('span');
          badge.className = 'filter-badge';
          titleEl.appendChild(badge);
        }
        badge.textContent = count;
      } else if (badge) {
        badge.remove();
      }
    });

    var hasFilters = state.filters.search ||
      state.filters.clusters.length ||
      state.filters.priorities.length ||
      state.filters.engagementStatuses.length;

    var clearBtn = document.getElementById('clearAllFilters');
    if (clearBtn) clearBtn.style.display = hasFilters ? '' : 'none';
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
      state.filters.clusters.length ||
      state.filters.priorities.length ||
      state.filters.engagementStatuses.length;

    syncSidebarState();

    document.getElementById('accountsTable').innerHTML = buildTable(sorted, false);
    wireTableClicks('accountsTable', sorted);

    document.querySelectorAll('#accountsTable th[data-col]').forEach(function(th) {
      th.className = state.sort.field === th.dataset.col ? 'sort-' + state.sort.dir : '';
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

    document.getElementById('allianceCards').innerHTML = state.alliances.map(function(al, idx) {
      var email = EC.extractEmails(al.internalPOC)[0] || al.internalPOC || '—';
      return '<div class="alliance-card" data-idx="' + idx + '">' +
        '<div class="alliance-card-name">' + esc(al.vendor) + '</div>' +
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

    document.getElementById('allianceDetailContent').innerHTML =
      '<h2 style="font-size:20px;font-weight:700;margin-bottom:20px">' + esc(al.vendor) + '</h2>' +
      pocBlock('Internal POC', al.internalPOC) +
      pocBlock('Vendor POC', al.vendorPOC) +
      (al.comment ? '<div class="panel-field"><div class="panel-field-label">Comment</div><div class="panel-field-value">' + esc(al.comment) + '</div></div>' : '') +
      (al.nextAction ? '<div class="panel-field"><div class="panel-field-label">Next Action</div><div class="panel-field-value">' + esc(al.nextAction) + '</div></div>' : '');
  }

  // ── Boot ──────────────────────────────────────────────────────────────────

  function init() {
    document.getElementById('importBtn').addEventListener('click', function() {
      document.getElementById('fileInput').click();
    });
    document.getElementById('fileInput').addEventListener('change', function(e) {
      handleFile(e.target.files[0]);
      e.target.value = '';
    });

    document.getElementById('panelClose').addEventListener('click', closePanel);
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') closePanel();
    });

    document.getElementById('viewAllLink').addEventListener('click', function(e) {
      e.preventDefault();
      state.filters = { search: '', clusters: [], priorities: [], engagementStatuses: [] };
      state.sidebarBuilt = false;
      navigate('accounts');
    });

    document.getElementById('allianceBack').addEventListener('click', renderAlliances);

    document.querySelectorAll('.nav-item').forEach(function(el) {
      el.addEventListener('click', function(e) {
        e.preventDefault();
        if (!el.classList.contains('nav-item--disabled')) navigate(el.dataset.view);
      });
    });

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

    wireSearch();

    var last = localStorage.getItem('ec_last_file');
    if (last) document.getElementById('importFilename').textContent = last;
  }

  document.addEventListener('DOMContentLoaded', init);
})();
