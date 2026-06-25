/* core.js — pure data logic, no DOM. UMD: works in Node and browser. */
(function(root, factory) {
  if (typeof module !== 'undefined') module.exports = factory();
  else root.EC = factory();
})(typeof self !== 'undefined' ? self : this, function() {
  'use strict';

  var STALENESS_DAYS = 30;

  var KNOWN_HEADERS = [
    'cluster', 'company', 'priority',
    'engagement status', 'engagement type', 'additional info required',
    'next steps', 'action owner', 'deloitte poc',
    'cluster poc', 'existing vendors', 'existing work done',
    'pipeline status', 'pipeline notes', 'last updated', 'notes'
  ];

  // ── Parsing ───────────────────────────────────────────────────────────────

  function parseWorkbook(wb) {
    return {
      accounts:  parseAccountsSheet(wb),
      alliances: parseAlliancesSheet(wb),
    };
  }

  var ACCOUNT_SHEET_NAMES = ['lean view', 'account tracker', 'accounts', 'tracker'];

  function parseAccountsSheet(wb) {
    var sheetName = wb.SheetNames.find(function(n) {
      return ACCOUNT_SHEET_NAMES.indexOf(n.trim().toLowerCase()) !== -1;
    }) || wb.SheetNames[0];
    var ws = wb.Sheets[sheetName];
    if (!ws) return [];

    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Find header row: scan first 6 rows for ≥2 known column labels
    var headerIdx = -1;
    var colMap = {};
    for (var i = 0; i < Math.min(6, rows.length); i++) {
      var cells = rows[i].map(function(c) { return String(c).toLowerCase().trim(); });
      var hits = 0;
      for (var k = 0; k < KNOWN_HEADERS.length; k++) {
        if (cells.indexOf(KNOWN_HEADERS[k]) !== -1) hits++;
      }
      if (hits >= 2) {
        headerIdx = i;
        cells.forEach(function(cell, j) { colMap[cell] = j; });
        break;
      }
    }
    if (headerIdx === -1) return [];

    function get(row, header) {
      var idx = colMap[header.toLowerCase()];
      return idx !== undefined ? String(row[idx] || '').trim() : '';
    }

    return rows.slice(headerIdx + 1)
      .filter(function(row) {
        return row.some(function(c) { return String(c).trim() !== ''; });
      })
      .map(function(row) {
        var pri = get(row, 'priority').toUpperCase();
        if (['HIGH', 'MEDIUM', 'LOW'].indexOf(pri) === -1) pri = 'UNSET';
        return {
          cluster:          get(row, 'cluster'),
          company:          get(row, 'company'),
          priority:         pri,
          engagementStatus: get(row, 'engagement status'),
          engagementType:   get(row, 'engagement type'),
          additionalInfo:   get(row, 'additional info required').toLowerCase() === 'yes',
          nextSteps:        get(row, 'next steps'),
          actionOwner:      get(row, 'action owner'),
          deloittePOC:      get(row, 'deloitte poc'),
          clusterPOC:       get(row, 'cluster poc'),
          existingVendors:  get(row, 'existing vendors'),
          existingWork:     get(row, 'existing work done'),
          pipelineStatus:   get(row, 'pipeline status'),
          pipelineNotes:    get(row, 'pipeline notes'),
          lastUpdated:      parseDate(get(row, 'last updated')),
          notes:            get(row, 'notes'),
        };
      })
      .filter(function(a) { return a.company; });
  }

  function parseAlliancesSheet(wb) {
    var sheetName = wb.SheetNames.find(function(n) {
      return n.trim().toLowerCase() === 'alliances';
    });
    if (!sheetName) return [];
    var ws = wb.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    return rows.map(function(r) {
      return {
        vendor:      String(r['Vendor']       || '').trim(),
        internalPOC: String(r['Internal POC'] || '').trim(),
        vendorPOC:   String(r['Vendor POC']   || '').trim(),
        comment:     String(r['Comment']      || '').trim(),
        nextAction:  String(r['Next Action']  || '').trim(),
      };
    }).filter(function(a) { return a.vendor; });
  }

  function parseDate(val) {
    if (!val || val === '') return null;
    // Excel serial number (days since 1899-12-30)
    var n = Number(val);
    if (!isNaN(n) && n > 1000 && n < 100000) {
      var d = new Date((n - 25569) * 86400000);
      if (!isNaN(d.getTime())) return d;
    }
    var d2 = new Date(val);
    return isNaN(d2.getTime()) ? null : d2;
  }

  // ── KPIs ──────────────────────────────────────────────────────────────────

  function computeKPIs(accounts) {
    var total        = accounts.length;
    var highPriority = accounts.filter(function(a) { return a.priority === 'HIGH'; }).length;
    var active       = accounts.filter(function(a) {
      return (a.engagementStatus || '').toLowerCase() === 'active';
    }).length;
    var needsInfo    = accounts.filter(function(a) { return a.additionalInfo; }).length;

    return {
      total:        total,
      highPriority: highPriority,
      highPct:      total ? Math.round(highPriority / total * 100) : 0,
      active:       active,
      activePct:    total ? Math.round(active / total * 100) : 0,
      needsInfo:    needsInfo,
    };
  }

  // ── Grouping ──────────────────────────────────────────────────────────────

  function groupBy(accounts, field) {
    return accounts.reduce(function(acc, a) {
      var key = a[field] || 'Unset';
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});
  }

  // ── Filtering & sorting ───────────────────────────────────────────────────

  var SEARCH_FIELDS = [
    'cluster', 'company', 'engagementStatus', 'engagementType',
    'nextSteps', 'actionOwner', 'deloittePOC', 'clusterPOC',
    'pipelineStatus', 'pipelineNotes', 'notes'
  ];

  function filterAccounts(accounts, filters) {
    var search             = (filters.search || '').toLowerCase();
    var clusters           = filters.clusters || [];
    var priorities         = filters.priorities || [];
    var engagementStatuses = filters.engagementStatuses || [];

    return accounts.filter(function(a) {
      if (search && !SEARCH_FIELDS.some(function(f) {
        return (a[f] || '').toLowerCase().indexOf(search) !== -1;
      })) return false;

      if (clusters.length && clusters.indexOf(a.cluster) === -1) return false;
      if (priorities.length && priorities.indexOf(a.priority) === -1) return false;
      if (engagementStatuses.length && engagementStatuses.indexOf(a.engagementStatus) === -1) return false;

      return true;
    });
  }

  function sortAccounts(accounts, field, dir) {
    var d = dir === 'desc' ? -1 : 1;
    return accounts.slice().sort(function(a, b) {
      var av = a[field], bv = b[field];
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      if (av instanceof Date && bv instanceof Date) return d * (av - bv);
      return d * String(av).localeCompare(String(bv));
    });
  }

  // ── Text helpers ──────────────────────────────────────────────────────────

  function extractEmails(text) {
    if (!text) return [];
    return (text.match(/[\w.+\-]+@[\w.\-]+\.\w+/g) || []);
  }

  function formatDate(d) {
    if (!d) return '—';
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }

  // ── Logo helpers ──────────────────────────────────────────────────────────

  var DOMAIN_MAP = {
    'slb':               'slb.com',
    'totalenergies':     'totalenergies.com',
    'phillips 66':       'phillips66.com',
    'hf sinclair':       'hfsinclair.com',
    'tc energy':         'tcenergy.com',
    'technipfmc':        'technipfmc.com',
    'baker hughes':      'bakerhughes.com',
    'conocophillips':    'conocophillips.com',
    'thyssenkrupp':      'thyssenkrupp.com',
    'rio tinto':         'riotinto.com',
    'exxonmobil':        'exxonmobil.com',
    'marathon petroleum':'marathonpetroleum.com',
    'halliburton':       'halliburton.com',
    'chevron':           'chevron.com',
    'suncor':            'suncor.com',
    'equinor':           'equinor.com',
    'enbridge':          'enbridge.com',
    'cenovus':           'cenovus.com',
    'citgo':             'citgo.com',
    'oceaneering':       'oceaneering.com',
    'nabors':            'nabors.com',
    'weatherford':       'weatherford.com',
    'alcoa':             'alcoa.com',
    'acerinox':          'acerinox.com',
  };

  function inferDomain(name) {
    if (!name) return null;
    var key = name.toLowerCase().trim();
    if (DOMAIN_MAP[key]) return DOMAIN_MAP[key];
    return key.replace(/[^a-z0-9]/g, '') + '.com';
  }

  function companyInitials(name) {
    if (!name) return '?';
    var words = name.trim().split(/\s+/).filter(Boolean);
    if (words.length === 1) return words[0].slice(0, 2).toUpperCase();
    return (words[0][0] + words[1][0]).toUpperCase();
  }

  function nameToColor(name) {
    var hash = 0, s = name || '?';
    for (var i = 0; i < s.length; i++) {
      hash = ((hash << 5) - hash) + s.charCodeAt(i); hash |= 0;
    }
    return 'hsl(' + (Math.abs(hash) % 360) + ',50%,38%)';
  }

  // ── Public API ────────────────────────────────────────────────────────────

  return {
    parseWorkbook:   parseWorkbook,
    computeKPIs:     computeKPIs,
    groupBy:         groupBy,
    filterAccounts:  filterAccounts,
    sortAccounts:    sortAccounts,
    extractEmails:   extractEmails,
    formatDate:      formatDate,
    inferDomain:     inferDomain,
    companyInitials: companyInitials,
    nameToColor:     nameToColor,
    STALENESS_DAYS:  STALENESS_DAYS,
  };
});
