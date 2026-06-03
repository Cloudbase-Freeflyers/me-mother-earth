/**
 * AMAZON ADS BULK SHEET ANALYZER
 * =================================
 * Place this script in any Google Sheet that has a "Main" tab.
 * The user enters a Google Drive URL/ID for an Amazon Ads bulk
 * editor XLSX in cell B3 of the Main tab, then clicks "Run Analysis"
 * (or runs the menu item Amazon Ads > Run Analysis).
 *
 * The script:
 *   1. Resolves the Drive file from the URL/ID
 *   2. Converts the XLSX into a temporary Google Sheet (Drive API)
 *   3. Reads every relevant tab from the converted sheet:
 *        - Sponsored Products Campaigns
 *        - Sponsored Brands Campaigns
 *        - SB Multi Ad Group Campaigns
 *        - Sponsored Display Campaigns
 *        - SP Search Term Report
 *        - SB Search Term Report
 *        - Portfolios
 *   4. Builds a multi-tab dashboard back into THIS sheet:
 *        - Dashboard          (headline KPIs + channel mix)
 *        - SP Campaigns       (campaign-level rollup)
 *        - SB & SD Campaigns  (cross-channel rollup)
 *        - Top ASINs          (advertised products)
 *        - Search Terms       (top 100 by sales)
 *        - Negative KWs       (the search-term analysis from the skill)
 *        - Top Performers     (winning search terms to harvest)
 *        - Match Types        (auto vs phrase vs broad vs exact)
 *        - Recommendations    (action-oriented summary)
 *
 * Notes on attribution:
 *   - Sponsored Display sales use the "Sales (Views & Clicks)" column
 *     when present, falling back to "Sales" otherwise.
 *   - Sponsored Brands Multi Ad Group is a duplicate view of the
 *     legacy SB tab; we use SB MAG when present (more recent
 *     attribution) and the legacy SB tab as a fallback. We do NOT
 *     sum them.
 *
 * Required: enable "Drive API v2" in Services (Resources > Advanced
 * Google services). The first run will prompt for OAuth scopes.
 */

// ============================================================
// CONSTANTS
// ============================================================
const MAIN_TAB_NAME = 'Main';
const URL_CELL = 'B3';
const STATUS_CELL = 'B5';

// Tabs we read from the bulk editor file
const TAB_SP_CAMPAIGNS  = 'Sponsored Products Campaigns';
const TAB_SB_CAMPAIGNS  = 'Sponsored Brands Campaigns';
const TAB_SB_MAG        = 'SB Multi Ad Group Campaigns';
const TAB_SD_CAMPAIGNS  = 'Sponsored Display Campaigns';
const TAB_SP_STR        = 'SP Search Term Report';
const TAB_SB_STR        = 'SB Search Term Report';
const TAB_PORTFOLIOS    = 'Portfolios';

// Dashboard output tab names
const OUT_DASHBOARD     = '📊 Dashboard';
const OUT_SP_CAMPAIGNS  = 'SP Campaigns';
const OUT_SB_SD         = 'SB & SD Campaigns';
const OUT_TOP_ASINS     = 'Top ASINs';
const OUT_BEST_KEYWORDS = '🎯 Best Keywords';
const OUT_BEST_TARGETS  = '🎯 Best Targets';
const OUT_PLACEMENTS    = '📍 Placements';
const OUT_SEARCH_TERMS  = 'Search Terms';
const OUT_NEGATIVES     = '⚠️ Negative KWs';
const OUT_TOP_PERF      = '🏆 Top Performers';
const OUT_MATCH_TYPES   = 'Match Types';
const OUT_RECOMMEND     = '💡 Recommendations';

// Style constants
const COLOR_PRIMARY  = '#1F3864';
const COLOR_ACCENT   = '#2E75B6';
const COLOR_SUCCESS  = '#375623';
const COLOR_DANGER   = '#C00000';
const COLOR_WARNING  = '#BF8F00';
const COLOR_LIGHT    = '#D9E1F2';
const COLOR_LIGHTER  = '#EDF2F9';
const COLOR_GREEN_LT = '#E2EFDA';
const COLOR_RED_LT   = '#FCE4E4';
const COLOR_GOLD_LT  = '#FFF2CC';
const COLOR_TEXT     = '#262626';
const COLOR_MUTED    = '#595959';
const COLOR_BORDER   = '#BFBFBF';

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Amazon Ads')
    .addItem('▶ Run Analysis', 'runAnalysis')
    .addSeparator()
    .addItem('⚙ Setup Main Tab', 'setupMainTab')
    .addItem('🧹 Clear Output Tabs', 'clearOutputTabs')
    .addToUi();
}

// ============================================================
// SETUP MAIN TAB (for first-time use)
// ============================================================
function setupMainTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let main = ss.getSheetByName(MAIN_TAB_NAME);
  if (!main) main = ss.insertSheet(MAIN_TAB_NAME, 0);
  main.clear();
  main.setHiddenGridlines(true);

  // Title
  main.getRange('A1').setValue('Amazon Ads Bulk Sheet Analyzer')
      .setFontSize(20).setFontWeight('bold').setFontColor(COLOR_PRIMARY);
  main.getRange('A1:E1').merge();

  // Subtitle
  main.getRange('A2').setValue('Paste the Google Drive URL of your Amazon Ads bulk editor XLSX, then run the analysis.')
      .setFontSize(11).setFontColor(COLOR_MUTED).setFontStyle('italic');
  main.getRange('A2:E2').merge();

  // URL input
  main.getRange('A3').setValue('Drive URL / File ID:').setFontWeight('bold').setFontColor(COLOR_PRIMARY);
  main.getRange('B3').setValue('').setBackground('#FFF2CC').setBorder(true, true, true, true, false, false);
  main.getRange('B3:E3').merge();

  // Status row
  main.getRange('A5').setValue('Status:').setFontWeight('bold').setFontColor(COLOR_PRIMARY);
  main.getRange('B5').setValue('Ready').setFontStyle('italic').setFontColor(COLOR_MUTED);
  main.getRange('B5:E5').merge();

  // Instructions
  main.getRange('A7').setValue('How to use')
      .setFontSize(14).setFontWeight('bold').setFontColor(COLOR_PRIMARY);

  const instructions = [
    ['1.', 'Upload your Amazon Ads bulk editor file (.xlsx) to Google Drive.'],
    ['2.', 'Right-click the file in Drive → Share → Anyone with link → Copy link.'],
    ['3.', 'Paste that link into cell B3 above.'],
    ['4.', 'Click the menu: Amazon Ads → Run Analysis.'],
    ['5.', 'When done, review the new tabs that appear at the bottom of this sheet.'],
  ];
  main.getRange(8, 1, instructions.length, 2).setValues(instructions)
      .setFontSize(11).setVerticalAlignment('top').setWrap(true);
  main.getRange(8, 1, instructions.length, 1).setFontWeight('bold').setHorizontalAlignment('center');

  // Output tabs preview
  main.getRange('A15').setValue('Output Tabs Generated')
      .setFontSize(14).setFontWeight('bold').setFontColor(COLOR_PRIMARY);
  const outputs = [
    ['📊 Dashboard',     'Account-level KPIs, channel mix, snapshot insights'],
    ['SP Campaigns',     'Sponsored Products campaign rollup with ROAS / ACOS'],
    ['SB & SD Campaigns','Sponsored Brands and Display campaign breakdown'],
    ['🎯 Best Keywords',  'Your targeting keywords (SP + SB), branded vs non-branded'],
    ['🎯 Best Targets',   'Product targets (SP) and audience targets (SD)'],
    ['📍 Placements',     'Top of Search / Product Pages / Rest of Search performance'],
    ['Top ASINs',        'Highest-performing advertised products'],
    ['Search Terms',     'Top customer search terms by sales'],
    ['⚠️ Negative KWs',  'Negative keyword candidates (zero-sale spend)'],
    ['🏆 Top Performers','Winning search terms to harvest into exact match'],
    ['Match Types',      'Auto vs Phrase vs Broad vs Exact comparison'],
    ['💡 Recommendations','Action-oriented strategic recommendations'],
  ];
  main.getRange(16, 1, outputs.length, 2).setValues(outputs).setFontSize(11);
  main.getRange(16, 1, outputs.length, 1).setFontWeight('bold').setFontColor(COLOR_PRIMARY);

  // Column widths
  main.setColumnWidth(1, 180);
  main.setColumnWidth(2, 350);
  main.setColumnWidths(3, 3, 120);

  // Activate
  main.getRange('B3').activate();
  SpreadsheetApp.getUi().alert('Main tab is set up. Paste a Drive URL into B3 and run the analysis.');
}

// ============================================================
// CLEAR OUTPUT TABS
// ============================================================
function clearOutputTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputTabs = [
    OUT_DASHBOARD, OUT_SP_CAMPAIGNS, OUT_SB_SD, OUT_TOP_ASINS,
    OUT_BEST_KEYWORDS, OUT_BEST_TARGETS, OUT_PLACEMENTS,
    OUT_SEARCH_TERMS, OUT_NEGATIVES, OUT_TOP_PERF, OUT_MATCH_TYPES, OUT_RECOMMEND
  ];
  outputTabs.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
  SpreadsheetApp.getUi().alert('Cleared output tabs.');
}

// ============================================================
// MAIN ENTRY POINT
// ============================================================
function runAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName(MAIN_TAB_NAME);
  if (!main) {
    SpreadsheetApp.getUi().alert('No "Main" tab found. Run "Setup Main Tab" first from the Amazon Ads menu.');
    return;
  }

  const urlOrId = String(main.getRange(URL_CELL).getValue() || '').trim();
  if (!urlOrId) {
    setStatus('❌ No URL provided. Paste a Drive URL into B3.', COLOR_DANGER);
    return;
  }

  let convertedFileId = null;
  try {
    setStatus('⏳ Resolving Drive file…', COLOR_ACCENT);
    const sourceFileId = extractFileId_(urlOrId);
    const sourceFile = DriveApp.getFileById(sourceFileId);
    const sourceName = sourceFile.getName();

    setStatus(`⏳ Converting "${sourceName}" to a Google Sheet…`, COLOR_ACCENT);
    convertedFileId = convertXlsxToSheet_(sourceFile);
    const tempSpreadsheet = SpreadsheetApp.openById(convertedFileId);

    setStatus('⏳ Reading bulk-editor tabs…', COLOR_ACCENT);
    const data = readBulkData_(tempSpreadsheet);

    setStatus('⏳ Building dashboard…', COLOR_ACCENT);
    buildDashboard_(ss, data, sourceName);
    buildSPCampaignsTab_(ss, data);
    buildSBSDTab_(ss, data);
    buildBestKeywordsTab_(ss, data);
    buildBestTargetsTab_(ss, data);
    buildPlacementsTab_(ss, data);
    buildTopAsinsTab_(ss, data);
    buildSearchTermsTab_(ss, data);
    buildNegativesTab_(ss, data);
    buildTopPerformersTab_(ss, data);
    buildMatchTypesTab_(ss, data);
    buildRecommendationsTab_(ss, data);

    setStatus(`✅ Done — analyzed "${sourceName}" at ${new Date().toLocaleString()}`, COLOR_SUCCESS);

    // Activate the dashboard tab
    const dash = ss.getSheetByName(OUT_DASHBOARD);
    if (dash) ss.setActiveSheet(dash);

  } catch (err) {
    setStatus(`❌ Error: ${err.message}`, COLOR_DANGER);
    Logger.log(err.stack || err);
    SpreadsheetApp.getUi().alert('Analysis failed:\n\n' + err.message);
  } finally {
    // Clean up the temporary converted Google Sheet
    if (convertedFileId) {
      try {
        DriveApp.getFileById(convertedFileId).setTrashed(true);
      } catch (e) {
        Logger.log('Could not trash temp file: ' + e);
      }
    }
  }
}

function setStatus(msg, color) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName(MAIN_TAB_NAME);
  if (!main) return;
  main.getRange(STATUS_CELL).setValue(msg)
    .setFontColor(color || COLOR_TEXT)
    .setFontWeight('bold');
  SpreadsheetApp.flush();
}

// ============================================================
// DRIVE / XLSX CONVERSION
// ============================================================
function extractFileId_(input) {
  // Accept full URL or bare ID
  if (/^[a-zA-Z0-9_-]{20,}$/.test(input)) return input;
  const patterns = [
    /\/d\/([a-zA-Z0-9_-]{20,})/,        // /d/<id>/
    /[?&]id=([a-zA-Z0-9_-]{20,})/,      // ?id=<id>
    /\/file\/d\/([a-zA-Z0-9_-]{20,})/,  // /file/d/<id>
  ];
  for (const re of patterns) {
    const m = input.match(re);
    if (m) return m[1];
  }
  throw new Error('Could not parse a Drive file ID from: ' + input);
}

function convertXlsxToSheet_(sourceFile) {
  // Use Drive API v2 to convert XLSX -> Google Sheet
  const blob = sourceFile.getBlob();
  const resource = {
    title: '__temp_amazon_ads_' + Date.now(),
    mimeType: MimeType.GOOGLE_SHEETS,
  };
  const inserted = Drive.Files.insert(resource, blob, { convert: true });
  return inserted.id;
}

// ============================================================
// READ BULK DATA — heart of the parser
// ============================================================
function readBulkData_(ss) {
  const sheets = {};
  ss.getSheets().forEach(s => { sheets[s.getName()] = s; });

  const data = {
    spCampaigns: [],
    spKeywords: [],
    spProductTargets: [],
    spProductAds: [],
    spPlacements: [],
    sbCampaigns: [],
    sbKeywords: [],
    sbMagCampaigns: [],
    sdCampaigns: [],
    sdProductAds: [],
    sdAudienceTargets: [],
    spSearchTerms: [],
    sbSearchTerms: [],
    portfolios: [],
    spMatchTypes: {},
    sourceTabs: Object.keys(sheets),
    dateRange: null,
  };

  // ---- SP Campaigns ----
  if (sheets[TAB_SP_CAMPAIGNS]) {
    parseSpCampaignTab_(sheets[TAB_SP_CAMPAIGNS], data);
  }
  // ---- SB Campaigns (legacy) ----
  if (sheets[TAB_SB_CAMPAIGNS]) {
    parseSbCampaignTab_(sheets[TAB_SB_CAMPAIGNS], data, false);
  }
  // ---- SB Multi Ad Group ----
  if (sheets[TAB_SB_MAG]) {
    parseSbCampaignTab_(sheets[TAB_SB_MAG], data, true);
  }
  // ---- SD Campaigns ----
  if (sheets[TAB_SD_CAMPAIGNS]) {
    parseSdCampaignTab_(sheets[TAB_SD_CAMPAIGNS], data);
  }
  // ---- SP Search Term Report ----
  if (sheets[TAB_SP_STR]) {
    parseSpSearchTermTab_(sheets[TAB_SP_STR], data);
  }
  // ---- SB Search Term Report ----
  if (sheets[TAB_SB_STR]) {
    parseSbSearchTermTab_(sheets[TAB_SB_STR], data);
  }
  // ---- Portfolios ----
  if (sheets[TAB_PORTFOLIOS]) {
    parsePortfoliosTab_(sheets[TAB_PORTFOLIOS], data);
  }

  return data;
}

// Helper — reads sheet to 2D array, builds header→index map, filters rows where Impressions>0
function readSheetWithHeaders_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [], idx: {} };
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[h] = i; });
  const rows = values.slice(1);
  return { headers, rows, idx };
}

function n_(v) {
  if (v === null || v === undefined || v === '') return 0;
  const f = parseFloat(v);
  return isNaN(f) ? 0 : f;
}

function s_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function getCol_(row, idx, name) {
  if (!(name in idx)) return null;
  return row[idx[name]];
}

function getNum_(row, idx, ...names) {
  for (const name of names) {
    if (name in idx) {
      const v = row[idx[name]];
      if (v !== '' && v !== null && v !== undefined) return n_(v);
    }
  }
  return 0;
}

function getStr_(row, idx, ...names) {
  for (const name of names) {
    if (name in idx) {
      const v = row[idx[name]];
      if (v !== '' && v !== null && v !== undefined) return s_(v);
    }
  }
  return '';
}

function hasImpressions_(row, idx) {
  return getNum_(row, idx, 'Impressions') > 0;
}

// -------- Parsers ------------------------------------------------
function parseSpCampaignTab_(sheet, data) {
  const { rows, idx } = readSheetWithHeaders_(sheet);
  rows.forEach(row => {
    if (!hasImpressions_(row, idx)) return;
    const ent = getStr_(row, idx, 'Entity');
    const common = {
      Campaign: getStr_(row, idx, 'Campaign Name (Informational only)', 'Campaign Name'),
      Portfolio: getStr_(row, idx, 'Portfolio Name (Informational only)'),
      Impressions: getNum_(row, idx, 'Impressions'),
      Clicks: getNum_(row, idx, 'Clicks'),
      Spend: getNum_(row, idx, 'Spend'),
      Sales: getNum_(row, idx, 'Sales'),
      Orders: getNum_(row, idx, 'Orders'),
      Units: getNum_(row, idx, 'Units'),
    };
    if (ent === 'Campaign') {
      data.spCampaigns.push(Object.assign({
        TargetingType: getStr_(row, idx, 'Targeting Type'),
        State: getStr_(row, idx, 'Campaign State (Informational only)', 'State'),
        DailyBudget: getNum_(row, idx, 'Daily Budget'),
      }, common));
    } else if (ent === 'Keyword') {
      const kw = getStr_(row, idx, 'Keyword Text');
      if (kw) {
        data.spKeywords.push(Object.assign({
          Keyword: kw,
          MatchType: getStr_(row, idx, 'Match Type'),
          Bid: getNum_(row, idx, 'Bid'),
        }, common));
      }
    } else if (ent === 'Product Targeting') {
      data.spProductTargets.push(Object.assign({
        TargetingExpr: getStr_(row, idx, 'Product Targeting Expression'),
        ResolvedExpr: getStr_(row, idx, 'Resolved Product Targeting Expression (Informational only)'),
      }, common));
    } else if (ent === 'Product Ad') {
      data.spProductAds.push(Object.assign({
        ASIN: getStr_(row, idx, 'ASIN'),
        SKU: getStr_(row, idx, 'SKU'),
        AdGroup: getStr_(row, idx, 'Ad Group Name (Informational only)'),
      }, common));
    } else if (ent === 'Bidding Adjustment') {
      const placement = getStr_(row, idx, 'Placement');
      if (placement) {
        data.spPlacements.push(Object.assign({
          Placement: placement,
          BiddingStrategy: getStr_(row, idx, 'Bidding Strategy'),
          Percentage: getNum_(row, idx, 'Percentage'),
        }, common));
      }
    }
  });
}

function parseSbCampaignTab_(sheet, data, isMag) {
  const { rows, idx } = readSheetWithHeaders_(sheet);
  rows.forEach(row => {
    if (!hasImpressions_(row, idx)) return;
    const ent = getStr_(row, idx, 'Entity');
    const common = {
      Campaign: getStr_(row, idx, 'Campaign Name (Informational only)', 'Campaign Name'),
      Impressions: getNum_(row, idx, 'Impressions'),
      Clicks: getNum_(row, idx, 'Clicks'),
      Spend: getNum_(row, idx, 'Spend'),
      Sales: getNum_(row, idx, 'Sales'),
      Orders: getNum_(row, idx, 'Orders'),
    };
    if (ent === 'Campaign' || ent === 'Draft Campaign') {
      const camp = Object.assign({
        AdFormat: getStr_(row, idx, 'Ad Format (Informational only)', 'Ad Format'),
        Budget: getNum_(row, idx, 'Budget'),
        IsMag: isMag,
      }, common);
      if (isMag) data.sbMagCampaigns.push(camp);
      else data.sbCampaigns.push(camp);
    } else if (ent === 'Keyword') {
      const kw = getStr_(row, idx, 'Keyword Text');
      if (kw) {
        data.sbKeywords.push(Object.assign({
          Keyword: kw,
          MatchType: getStr_(row, idx, 'Match Type'),
        }, common));
      }
    }
  });
}

function parseSdCampaignTab_(sheet, data) {
  const { rows, idx } = readSheetWithHeaders_(sheet);
  rows.forEach(row => {
    if (!hasImpressions_(row, idx)) return;
    const ent = getStr_(row, idx, 'Entity');
    const common = {
      Campaign: getStr_(row, idx, 'Campaign Name (Informational only)', 'Campaign Name'),
      Impressions: getNum_(row, idx, 'Impressions'),
      Clicks: getNum_(row, idx, 'Clicks'),
      Spend: getNum_(row, idx, 'Spend'),
      // Use Views & Clicks attribution when available, else fall back to Sales
      Sales: getNum_(row, idx, 'Sales (Views & Clicks)', 'Sales'),
      Orders: getNum_(row, idx, 'Orders (Views & Clicks)', 'Orders'),
    };
    if (ent === 'Campaign') {
      data.sdCampaigns.push(Object.assign({
        Tactic: getStr_(row, idx, 'Tactic'),
        Budget: getNum_(row, idx, 'Budget'),
      }, common));
    } else if (ent === 'Product Ad') {
      data.sdProductAds.push(Object.assign({
        ASIN: getStr_(row, idx, 'ASIN'),
        SKU: getStr_(row, idx, 'SKU'),
        Units: getNum_(row, idx, 'Units'),
      }, common));
    } else if (ent === 'Audience Targeting') {
      data.sdAudienceTargets.push(Object.assign({
        TargetingExpr: getStr_(row, idx, 'Targeting Expression'),
        ResolvedExpr: getStr_(row, idx, 'Resolved Targeting Expression (Informational only)'),
      }, common));
    }
  });
}

function parseSpSearchTermTab_(sheet, data) {
  const { rows, idx } = readSheetWithHeaders_(sheet);
  // Aggregate as we read — a given search term can appear in many rows
  const cstAgg = {};
  const mtAgg = {};
  rows.forEach(row => {
    if (!hasImpressions_(row, idx)) return;
    const cst = getStr_(row, idx, 'Customer Search Term').toLowerCase();
    if (!cst) return;
    if (!cstAgg[cst]) cstAgg[cst] = {
      Term: cst,
      Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0,
      Campaigns: new Set(), Keywords: new Set(), MatchTypes: new Set(),
    };
    const a = cstAgg[cst];
    a.Impressions += getNum_(row, idx, 'Impressions');
    a.Clicks      += getNum_(row, idx, 'Clicks');
    a.Spend       += getNum_(row, idx, 'Spend');
    a.Sales       += getNum_(row, idx, 'Sales');
    a.Orders      += getNum_(row, idx, 'Orders');
    const camp = getStr_(row, idx, 'Campaign Name (Informational only)', 'Campaign Name');
    if (camp) a.Campaigns.add(camp);
    const kw = getStr_(row, idx, 'Keyword Text');
    if (kw) a.Keywords.add(kw);
    const mt = getStr_(row, idx, 'Match Type') || 'Auto';
    a.MatchTypes.add(mt);
    if (!mtAgg[mt]) mtAgg[mt] = { MatchType: mt, Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0 };
    mtAgg[mt].Impressions += getNum_(row, idx, 'Impressions');
    mtAgg[mt].Clicks      += getNum_(row, idx, 'Clicks');
    mtAgg[mt].Spend       += getNum_(row, idx, 'Spend');
    mtAgg[mt].Sales       += getNum_(row, idx, 'Sales');
    mtAgg[mt].Orders      += getNum_(row, idx, 'Orders');
  });
  // Convert sets to comma-strings, push to data
  Object.values(cstAgg).forEach(a => {
    a.Campaigns = Array.from(a.Campaigns).join(', ');
    a.Keywords = Array.from(a.Keywords).join(', ');
    a.MatchTypes = Array.from(a.MatchTypes).join(', ');
    data.spSearchTerms.push(a);
  });
  data.spMatchTypes = mtAgg;
}

function parseSbSearchTermTab_(sheet, data) {
  const { rows, idx } = readSheetWithHeaders_(sheet);
  const cstAgg = {};
  rows.forEach(row => {
    if (!hasImpressions_(row, idx)) return;
    const cst = getStr_(row, idx, 'Customer Search Term').toLowerCase();
    if (!cst) return;
    if (!cstAgg[cst]) cstAgg[cst] = {
      Term: cst,
      Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0,
    };
    const a = cstAgg[cst];
    a.Impressions += getNum_(row, idx, 'Impressions');
    a.Clicks      += getNum_(row, idx, 'Clicks');
    a.Spend       += getNum_(row, idx, 'Spend');
    a.Sales       += getNum_(row, idx, 'Sales');
    a.Orders      += getNum_(row, idx, 'Orders');
  });
  Object.values(cstAgg).forEach(a => data.sbSearchTerms.push(a));
}

function parsePortfoliosTab_(sheet, data) {
  const { rows, idx } = readSheetWithHeaders_(sheet);
  rows.forEach(row => {
    const name = getStr_(row, idx, 'Portfolio Name');
    if (name) {
      data.portfolios.push({
        Name: name,
        Budget: getNum_(row, idx, 'Budget Amount'),
        State: getStr_(row, idx, 'State (Informational only)'),
      });
    }
  });
}
