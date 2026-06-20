// ============================================================
// EC FSM Opportunity Tracker — Google Apps Script Backend
//
// SETUP INSTRUCTIONS:
//   1. Go to script.google.com → New project, paste this code
//   2. Set SHEET_ID to your Google Sheet's ID (from the URL:
//      docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit)
//   3. Set API_KEY to a secret string — copy it, same value
//      goes in CONFIG.apiKey in EC_FieldService_Tracker.html
//   4. Create a Google Sheet; name the first tab "Accounts"
//      with this header in row 1:
//      id | company | category | priority | status | nextsteps |
//      clusterPOC | keyPOC | clientPOC | pipeline | action | meeting
//   5. Deploy → New deployment → Web app
//      Execute as: Me  |  Who has access: Anyone
//      Copy the Web App URL → paste into CONFIG.sheetUrl in the HTML
//   6. From the script editor, run migrateData() once to seed
//      the sheet with the 27 starter accounts
// ============================================================

var SHEET_ID  = 'YOUR_GOOGLE_SHEET_ID_HERE';
var SHEET_TAB = 'Accounts';
var API_KEY   = 'YOUR_SECRET_KEY_HERE';

// Order matches Excel columns exactly (id is internal, appended at col A)
var COLS = ['id','category','company','priority','status','interactedWith','meeting',
            'clusterPOC','keyPOC','clientPOC','existingVendors','existingWorkDone',
            'pipeline','action','lastUpdatedAt','lastUpdatedBy','nextsteps'];

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
}

function rowToObj_(row) {
  var obj = {};
  COLS.forEach(function(c, i) {
    obj[c] = (row[i] !== undefined && row[i] !== null) ? String(row[i]) : '';
  });
  return obj;
}

function makeJson_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (!e || e.parameter.key !== API_KEY) {
    return makeJson_({ error: 'Unauthorized' });
  }
  var sheet = getSheet_();
  var data  = sheet.getDataRange().getValues();
  var rows  = data.slice(1)
    .filter(function(r) { return r[0] !== ''; })
    .map(rowToObj_);
  return makeJson_(rows);
}

function doPost(e) {
  var payload;
  try { payload = JSON.parse(e.postData.contents); }
  catch (err) { return makeJson_({ error: 'Invalid JSON' }); }

  if (payload.key !== API_KEY) {
    return makeJson_({ error: 'Unauthorized' });
  }

  var sheet = getSheet_();
  var data  = sheet.getDataRange().getValues();

  // DELETE
  if (payload.action === 'delete') {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.id)) {
        sheet.deleteRow(i + 1);
        return makeJson_({ ok: true });
      }
    }
    return makeJson_({ ok: true, note: 'row not found' });
  }

  // UPSERT (insert or update)
  var acct    = payload.account;
  var rowData = COLS.map(function(c) { return acct[c] || ''; });
  var found   = false;

  for (var j = 1; j < data.length; j++) {
    if (String(data[j][0]) === String(acct.id)) {
      sheet.getRange(j + 1, 1, 1, COLS.length).setValues([rowData]);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow(rowData);
  }

  return makeJson_({ ok: true });
}

// Run this ONCE from the Apps Script editor (Run → migrateData)
// to seed the sheet with all 27 starter accounts.
function migrateData() {
  var accounts = [
    { id:1,  category:"ACP Accounts",    company:"Koch Industries",    priority:"",       status:"",                                                                                                              nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:2,  category:"ACP Accounts",    company:"SLB",                priority:"",       status:"",                                                                                                              nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:3,  category:"ACP Accounts",    company:"Halliburton",        priority:"LOW",    status:"Brett staying engaged with Account Team; Field Service rollout tied to SAP S4.",                               nextsteps:"Potential MSFT Project Ops Opps",                                                                                                                                                           clusterPOC:"Nick Karam, Peter B, Salesforce Account Team",                  keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:4,  category:"ACP Accounts",    company:"Marathon Petroleum", priority:"LOW",    status:"Continued expansion of existing Field Service Project",                                                        nextsteps:"Engage account team / Existing NGMM team on next phase",                                                                                                                                    clusterPOC:"Account team / Existing NGMM team",                             keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:5,  category:"ACP Accounts",    company:"Chevron",            priority:"MEDIUM", status:"Potential outreach, not many details yet",                                                                     nextsteps:"1. Meet with Cluster lead and present one-pager\n2. Upon their sign-off, plan next steps",                                                                                                       clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:6,  category:"ACP Accounts",    company:"ExxonMobil",         priority:"",       status:"",                                                                                                              nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:7,  category:"Upstream",        company:"ConocoPhillips",     priority:"MEDIUM", status:"Luis to follow-up; call not yet scheduled",                                                                    nextsteps:"Target account for SF – greenfield. Mulesoft pipe, Informatica takeout, Tableau in pipeline",                                                                                                    clusterPOC:"Pete Buettgen (pbuettgen@deloitte.com)",                        keyPOC:"",                                                   clientPOC:"", pipeline:"Pipeline: Mulesoft pipe, Informatica takeout, Tableau. Not vendor-specific yet – SFDC TBD",           action:"",                                                                                         meeting:"" },
    { id:8,  category:"Upstream",        company:"Duncan Oil",         priority:"",       status:"SMB account for SF",                                                                                           nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:9,  category:"Upstream",        company:"Murphy Oil",         priority:"",       status:"Target account for SF – greenfield",                                                                           nextsteps:"Field service opp in March – no partner tag. PWC tagged to other opps.",                                                                                                                         clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"Field service opp March, PWC tagged to other opps",                                                action:"",                                                                                         meeting:"" },
    { id:10, category:"Upstream",        company:"Suncor",             priority:"HIGH",   status:"Brett waiting for update from Anand; CIO mentioned Field Service as key pillar for 2026",                     nextsteps:"Follow up with account team on field service opp (next summer, no partner tag yet)",                                                                                                          clusterPOC:"Anand Radia",                                                   keyPOC:"",                                                   clientPOC:"", pipeline:"Field service opp for next summer – no partner tag",                                               action:"[06/04] Brett to follow up with account team",                                             meeting:"" },
    { id:11, category:"Upstream",        company:"Equinor",            priority:"HIGH",   status:"Luis to follow up with Andy",                                                                                  nextsteps:"EMEA account – get Andy's email and schedule intro",                                                                                                                                             clusterPOC:"",                                                              keyPOC:"Andy (?) – Email Needed",                            clientPOC:"", pipeline:"AGNOSTIC – EMEA account",                                                                               action:"",                                                                                         meeting:"" },
    { id:12, category:"Upstream",        company:"TotalEnergies",      priority:"",       status:"EMEA account for SF ($5–10M AOV)",                                                                             nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:13, category:"Mid / Downstream",company:"HF Sinclair",        priority:"HIGH",   status:"Meeting with Susan scheduled; Gariel involved in RFP",                                                         nextsteps:"Lack partner strategy – prime target. POV for field service turnarounds/refining, Agentforce for onboarding, Digital/Loyalty RFP",                                                                clusterPOC:"Susan Cattozzo (scattozzo@deloitte.com)",                       keyPOC:"Susan (portfolio), Gariel (Salesforce)",             clientPOC:"", pipeline:"Field service opp 'turnarounds/refining', Agentforce onboarding, Digital/Loyalty RFP",              action:"",                                                                                         meeting:"" },
    { id:14, category:"Mid / Downstream",company:"Enbridge",           priority:"HIGH",   status:"Canadian account for SF. Salesforce preparing POV/POC – pending audit green light",                           nextsteps:"Confirm audit status; proceed with POV/POC once cleared",                                                                                                                                       clusterPOC:"McClintock, Mike (mimcclintock@deloitte.com)",                  keyPOC:"",                                                   clientPOC:"", pipeline:"Salesforce POV/POC pending audit approval",                                                       action:"",                                                                                         meeting:"" },
    { id:15, category:"Mid / Downstream",company:"TC Energy",          priority:"",       status:"Opportunities exist; decisions led from Canada",                                                               nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:16, category:"Mid / Downstream",company:"Cenovus",            priority:"",       status:"",                                                                                                              nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:17, category:"Mid / Downstream",company:"CITGO",              priority:"HIGH",   status:"Restricted account (unknown reasons) – Ryan Watson's account. Existing POV deck; client engagement needed",   nextsteps:"Luis/Brett to follow up with Salesforce team; leverage existing POV deck",                                                                                                                  clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"Existing POV deck ready",                                                                          action:"[06/04] Brett to follow up with account team",                                             meeting:"" },
    { id:18, category:"Mid / Downstream",company:"Phillips 66",        priority:"HIGH",   status:"Luis has pending conversation with P66 Lubricants lead. S4 implementation ongoing. Proficient very saturated – Deloitte needs to be all-in on Salesforce to win", nextsteps:"Leverage P66 POV as baseline for other accounts. Field Service POC opp July 2026 – no partner tag. Harish Kumbhare working opportunity", clusterPOC:"",                                                  keyPOC:"Sonia (?) – Email Needed",                           clientPOC:"", pipeline:"Field Service POC opp July 2026 – no partner tag. S4 implementation in progress.",               action:"[06/04] On hold – S4 implementation ongoing, Harish Kumbhare working the opp",             meeting:"" },
    { id:19, category:"ETS / OFS",       company:"TechnipFMC",         priority:"HIGH",   status:"Luis to start conversation; tailored POV to be created",                                                       nextsteps:"Reach out to Linh to start contact. Existing SFDC + ServiceMax (Asset 360) stack – key differentiator",                                                                                           clusterPOC:"Michelle Mackelmore (nichellemclemore@deloitte.com) – OFS",     keyPOC:"Nguyen, Linh (linhnguyen3@deloitte.com)",            clientPOC:"", pipeline:"SFDC + ServiceMax (Asset 360) existing",                                                             action:"[06/04] Brett to follow up to Linh to start contact",                                      meeting:"" },
    { id:20, category:"ETS / OFS",       company:"Oceaneering",        priority:"LOW",    status:"TBD",                                                                                                           nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:21, category:"ETS / OFS",       company:"Nabors",             priority:"HIGH",   status:"Brett followed up with Nick; definite field service opportunity – EP planning to lead",                        nextsteps:"Coordinate with Nick Karam and EP team on approach",                                                                                                                                          clusterPOC:"Nick Karam",                                                    keyPOC:"Nicholas Karan (nkaram@deloitte.com), Jeff Croxen (jcroxen@deloitte.com)", clientPOC:"", pipeline:"Field service opp – EP planning to lead",                                                    action:"[06/04] Brett to follow up with account team",                                             meeting:"" },
    { id:22, category:"ETS / OFS",       company:"Weatherford",        priority:"",       status:"PWC dominant here",                                                                                             nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"No – PWC is huge here",                                                                           action:"",                                                                                         meeting:"" },
    { id:23, category:"ETS / OFS",       company:"Baker Hughes",       priority:"HIGH",   status:"Similar to Phillips 66 – Deloitte needs to be all-in on Salesforce to compete",                               nextsteps:"Define all-in Salesforce strategy before engaging",                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:24, category:"Mining",          company:"Rio Tinto",          priority:"",       status:"TBD",                                                                                                           nextsteps:"",                                                                                                                                                                                          clusterPOC:"Sam Das (samratdas7@deloitte.com)",                             keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:25, category:"Mining",          company:"Alcoa",              priority:"",       status:"TBD",                                                                                                           nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:26, category:"Mining",          company:"Thyssenkrupp",       priority:"",       status:"TBD",                                                                                                           nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" },
    { id:27, category:"Mining",          company:"Acerinox",           priority:"",       status:"TBD",                                                                                                           nextsteps:"",                                                                                                                                                                                          clusterPOC:"",                                                              keyPOC:"",                                                   clientPOC:"", pipeline:"",                                                                                                    action:"",                                                                                         meeting:"" }
  ];

  var sheet = getSheet_();

  // Add header row if sheet is empty (display names matching Excel)
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['id','Account Classification','Company Name','Priority Focus',
      'Status / Next Steps','Interacted With?','Next Meeting Date',
      'Cluster POC','Key Account POC','Client POC','Existing Vendors',
      'Existing Work Done','Vendor Pipeline Status','Action Items',
      'Last Updated Date','Last Updated By','Notes']);
  }

  // Clear existing data rows (keep header)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  // Seed all accounts
  accounts.forEach(function(a) {
    sheet.appendRow(COLS.map(function(c) { return a[c] || ''; }));
  });

  Logger.log('Migration complete: ' + accounts.length + ' accounts seeded.');
}
