/* core.js — pure data logic, no DOM. UMD: works in Node and browser. */
(function(root, factory) {
  if (typeof module !== 'undefined') module.exports = factory();
  else root.EC = factory();
})(typeof self !== 'undefined' ? self : this, function() {
  'use strict';

  var STALENESS_DAYS    = 30;
  var MEETING_WINDOW_DAYS = 30;

  var KNOWN_HEADERS = [
    'company name', 'account classification', 'priority focus',
    'status / next steps', 'interacted with?', 'next meeting date',
    'cluster poc', 'key account poc', 'client poc',
    'existing vendors', 'existing work done', 'vendor pipeline status',
    'action items', 'last updated date', 'last updated by', 'notes', 'alliance'
  ];

  // ── Parsing ───────────────────────────────────────────────────────────────

  function parseWorkbook(wb) {
    return {
      accounts:  parseAccountsSheet(wb),
      alliances: parseAlliancesSheet(wb),
    };
  }

  function parseAccountsSheet(wb) {
    var sheetName = wb.SheetNames.find(function(n) {
      return n.trim().toLowerCase() === 'lean view';
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
        var pri = get(row, 'priority focus').toUpperCase();
        if (['HIGH', 'MEDIUM', 'LOW'].indexOf(pri) === -1) pri = 'UNSET';
        return {
          company:         get(row, 'company name'),
          classification:  get(row, 'account classification'),
          priority:        pri,
          status:          get(row, 'status / next steps'),
          interacted:      get(row, 'interacted with?').toLowerCase() === 'yes',
          nextMeeting:     parseDate(get(row, 'next meeting date')),
          clusterPOC:      get(row, 'cluster poc'),
          keyPOC:          get(row, 'key account poc'),
          clientPOC:       get(row, 'client poc'),
          existingVendors: get(row, 'existing vendors'),
          existingWork:    get(row, 'existing work done'),
          pipeline:        get(row, 'vendor pipeline status'),
          action:          get(row, 'action items'),
          lastUpdatedDate: parseDate(get(row, 'last updated date')),
          lastUpdatedBy:   get(row, 'last updated by'),
          notes:           get(row, 'notes'),
          alliance:        get(row, 'alliance'),
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
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var in30 = new Date(today); in30.setDate(today.getDate() + MEETING_WINDOW_DAYS);

    var total        = accounts.length;
    var highPriority = accounts.filter(function(a) { return a.priority === 'HIGH'; }).length;
    var engaged      = accounts.filter(function(a) { return a.interacted; }).length;
    var upcoming     = accounts.filter(function(a) {
      return a.nextMeeting && a.nextMeeting >= today && a.nextMeeting <= in30;
    });
    upcoming.sort(function(a, b) { return a.nextMeeting - b.nextMeeting; });

    return {
      total:        total,
      highPriority: highPriority,
      highPct:      total ? Math.round(highPriority / total * 100) : 0,
      engaged:      engaged,
      engagedPct:   total ? Math.round(engaged / total * 100) : 0,
      meetingsIn30: upcoming.length,
      nextMeeting:  upcoming.length ? upcoming[0].nextMeeting : null,
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

  function groupByAlliance(accounts) {
    var result = {};
    accounts.forEach(function(a) {
      if (!a.alliance) return;
      a.alliance.split(/[,;]/).forEach(function(s) {
        var al = s.trim();
        if (al) result[al] = (result[al] || 0) + 1;
      });
    });
    return result;
  }

  // ── Filtering & sorting ───────────────────────────────────────────────────

  var SEARCH_FIELDS = [
    'company', 'classification', 'status',
    'clusterPOC', 'keyPOC', 'clientPOC', 'pipeline', 'notes'
  ];

  function filterAccounts(accounts, filters) {
    var search          = (filters.search || '').toLowerCase();
    var classifications = filters.classifications || [];
    var priorities      = filters.priorities || [];
    var alliances       = filters.alliances || [];
    var interacted      = (filters.interacted !== undefined && filters.interacted !== null)
                          ? filters.interacted : null;

    return accounts.filter(function(a) {
      if (search && !SEARCH_FIELDS.some(function(f) {
        return (a[f] || '').toLowerCase().indexOf(search) !== -1;
      })) return false;

      if (classifications.length && classifications.indexOf(a.classification) === -1) return false;
      if (priorities.length && priorities.indexOf(a.priority) === -1) return false;

      if (alliances.length) {
        var tags = splitAlliance(a.alliance);
        if (!alliances.some(function(al) { return tags.indexOf(al) !== -1; })) return false;
      }

      if (interacted !== null && a.interacted !== interacted) return false;
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

  function splitAlliance(text) {
    if (!text) return [];
    return text.split(/[,;]/).map(function(s) { return s.trim(); }).filter(Boolean);
  }

  function parsePipelineBullets(text) {
    if (!text) return [];
    var idx = text.toLowerCase().indexOf('pipeline:');
    var src = idx >= 0 ? text.slice(idx + 9) : text;
    return src.split(/[\n\r]+/)
      .map(function(l) { return l.replace(/^[-•*]\s*/, '').trim(); })
      .filter(Boolean);
  }

  function parseActionItems(text) {
    if (!text) return [];
    var re = /\[(\d{2}\/\d{2})\]\s*([^\[]+)/g;
    var items = [], m;
    while ((m = re.exec(text)) !== null) {
      items.push({ date: m[1], text: m[2].trim() });
    }
    if (items.length) return items;
    // Fallback: plain lines
    return text.split(/[\n\r]+/)
      .map(function(l) { return l.trim(); })
      .filter(Boolean)
      .map(function(t) { return { date: null, text: t }; });
  }

  function extractEmails(text) {
    if (!text) return [];
    return (text.match(/[\w.+\-]+@[\w.\-]+\.\w+/g) || []);
  }

  function formatDate(d) {
    if (!d) return '—';
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }

  // ── Public API ────────────────────────────────────────────────────────────

  return {
    parseWorkbook:        parseWorkbook,
    computeKPIs:          computeKPIs,
    groupBy:              groupBy,
    groupByAlliance:      groupByAlliance,
    filterAccounts:       filterAccounts,
    sortAccounts:         sortAccounts,
    splitAlliance:        splitAlliance,
    parsePipelineBullets: parsePipelineBullets,
    parseActionItems:     parseActionItems,
    extractEmails:        extractEmails,
    formatDate:           formatDate,
    STALENESS_DAYS:       STALENESS_DAYS,
    MEETING_WINDOW_DAYS:  MEETING_WINDOW_DAYS,
  };
});
