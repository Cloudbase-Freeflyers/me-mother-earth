/**
 * DASHBOARD BUILDERS
 * ==================
 * Each function builds one of the output tabs.
 */

// ============================================================
// HELPERS — shared formatting
// ============================================================
function getOrCreateTab_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) {
    sh.clear();
    sh.clearConditionalFormatRules();
  } else {
    sh = ss.insertSheet(name);
  }
  sh.setHiddenGridlines(true);
  return sh;
}

function setTitle_(sh, row, text, span) {
  sh.getRange(row, 1).setValue(text)
    .setFontSize(18).setFontWeight('bold').setFontColor(COLOR_PRIMARY);
  if (span > 1) sh.getRange(row, 1, 1, span).merge();
  sh.setRowHeight(row, 32);
}

function setSubtitle_(sh, row, text, span) {
  sh.getRange(row, 1).setValue(text)
    .setFontSize(10).setFontStyle('italic').setFontColor(COLOR_MUTED);
  if (span > 1) sh.getRange(row, 1, 1, span).merge();
}

function setSection_(sh, row, text, span) {
  sh.getRange(row, 1).setValue(text)
    .setFontSize(13).setFontWeight('bold').setFontColor(COLOR_PRIMARY);
  if (span > 1) sh.getRange(row, 1, 1, span).merge();
  sh.setRowHeight(row, 26);
}

function styleHeaderRow_(sh, row, startCol, numCols, bgColor) {
  const range = sh.getRange(row, startCol, 1, numCols);
  range.setBackground(bgColor || COLOR_PRIMARY)
    .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 28);
}

function styleDataBlock_(sh, startRow, startCol, numRows, numCols) {
  if (numRows <= 0) return;
  const range = sh.getRange(startRow, startCol, numRows, numCols);
  range.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, COLOR_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  // Alternating row backgrounds
  for (let i = 0; i < numRows; i++) {
    if (i % 2 === 1) {
      sh.getRange(startRow + i, startCol, 1, numCols).setBackground(COLOR_LIGHTER);
    }
  }
}

function setKpiCard_(sh, row, col, label, value, color) {
  const labelCell = sh.getRange(row, col);
  const valueCell = sh.getRange(row + 1, col);
  labelCell.setValue(label).setFontSize(10).setFontColor('#FFFFFF')
    .setBackground(color || COLOR_PRIMARY).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, false, true, false, false, COLOR_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  valueCell.setValue(value).setFontSize(18).setFontColor(color || COLOR_PRIMARY)
    .setBackground('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, true, true, true, false, false, COLOR_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 24);
  sh.setRowHeight(row + 1, 36);
}

function fmt$(n)   { return '$' + Math.round(n).toLocaleString(); }
function fmt$$(n)  { return '$' + n.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}); }
function fmtNum(n) { return Math.round(n).toLocaleString(); }
function fmtPct(n) { return (n*100).toFixed(1) + '%'; }
function fmtRoas(n){ return n.toFixed(2) + 'x'; }
function safeDiv(a, b) { return b === 0 ? 0 : a / b; }


// ============================================================
// TAB 1 — DASHBOARD
// ============================================================
function buildDashboard_(ss, data, sourceName) {
  const sh = getOrCreateTab_(ss, OUT_DASHBOARD);

  // Compute aggregates
  const sp = aggregate_(data.spCampaigns);
  const sbLegacy = aggregate_(data.sbCampaigns);
  const sbMag = aggregate_(data.sbMagCampaigns);
  // SB MAG is duplicate of legacy SB — use MAG when present (more recent attribution)
  const sb = (data.sbMagCampaigns.length > 0) ? sbMag : sbLegacy;
  const sd = aggregate_(data.sdCampaigns);

  const total = {
    Impressions: sp.Impressions + sb.Impressions + sd.Impressions,
    Clicks:      sp.Clicks      + sb.Clicks      + sd.Clicks,
    Spend:       sp.Spend       + sb.Spend       + sd.Spend,
    Sales:       sp.Sales       + sb.Sales       + sd.Sales,
    Orders:      sp.Orders      + sb.Orders      + sd.Orders,
  };
  const roas = safeDiv(total.Sales, total.Spend);
  const acos = safeDiv(total.Spend, total.Sales);
  const ctr  = safeDiv(total.Clicks, total.Impressions);
  const cpc  = safeDiv(total.Spend, total.Clicks);
  const cvr  = safeDiv(total.Orders, total.Clicks);
  const aov  = safeDiv(total.Sales, total.Orders);

  // Title
  setTitle_(sh, 1, 'Amazon Ads — Performance Dashboard', 8);
  setSubtitle_(sh, 2, `Source: ${sourceName}  •  Generated: ${new Date().toLocaleString()}`, 8);

  // KPI cards — 4 across, 2 rows
  // Row 4-5: Spend, Sales, ROAS, ACOS
  setKpiCard_(sh, 4, 1, 'AD SPEND',    fmt$(total.Spend),  COLOR_PRIMARY);
  setKpiCard_(sh, 4, 3, 'AD SALES',    fmt$(total.Sales),  COLOR_SUCCESS);
  setKpiCard_(sh, 4, 5, 'ROAS',        fmtRoas(roas),      COLOR_ACCENT);
  setKpiCard_(sh, 4, 7, 'ACOS',        fmtPct(acos),       COLOR_WARNING);
  // Row 7-8: Orders, Clicks, CPC, CTR
  setKpiCard_(sh, 7, 1, 'ORDERS',      fmtNum(total.Orders),       COLOR_PRIMARY);
  setKpiCard_(sh, 7, 3, 'CLICKS',      fmtNum(total.Clicks),       COLOR_PRIMARY);
  setKpiCard_(sh, 7, 5, 'AVG CPC',     fmt$$(cpc),                 COLOR_ACCENT);
  setKpiCard_(sh, 7, 7, 'CTR',         fmtPct(ctr),                COLOR_ACCENT);
  // Row 10-11: Impressions, AOV, CVR
  setKpiCard_(sh, 10, 1, 'IMPRESSIONS', fmtNum(total.Impressions), COLOR_PRIMARY);
  setKpiCard_(sh, 10, 3, 'AVG ORDER VALUE', fmt$$(aov),            COLOR_PRIMARY);
  setKpiCard_(sh, 10, 5, 'CONVERSION RATE', fmtPct(cvr),           COLOR_SUCCESS);

  // Channel mix table starting row 13
  setSection_(sh, 13, 'Channel Mix', 8);
  const chanHeaders = ['Channel', 'Spend', '% Spend', 'Sales', '% Sales', 'Orders', 'ROAS', 'ACOS'];
  sh.getRange(14, 1, 1, chanHeaders.length).setValues([chanHeaders]);
  styleHeaderRow_(sh, 14, 1, chanHeaders.length);

  const chanData = [
    ['Sponsored Products', sp.Spend, safeDiv(sp.Spend, total.Spend), sp.Sales, safeDiv(sp.Sales, total.Sales), sp.Orders, safeDiv(sp.Sales, sp.Spend), safeDiv(sp.Spend, sp.Sales)],
    ['Sponsored Brands',   sb.Spend, safeDiv(sb.Spend, total.Spend), sb.Sales, safeDiv(sb.Sales, total.Sales), sb.Orders, safeDiv(sb.Sales, sb.Spend), safeDiv(sb.Spend, sb.Sales)],
    ['Sponsored Display',  sd.Spend, safeDiv(sd.Spend, total.Spend), sd.Sales, safeDiv(sd.Sales, total.Sales), sd.Orders, safeDiv(sd.Sales, sd.Spend), safeDiv(sd.Spend, sd.Sales)],
    ['TOTAL',              total.Spend, 1, total.Sales, 1, total.Orders, roas, acos],
  ];
  sh.getRange(15, 1, chanData.length, chanData[0].length).setValues(chanData);
  styleDataBlock_(sh, 15, 1, chanData.length, chanHeaders.length);

  // Number formats
  sh.getRange(15, 2, 4, 1).setNumberFormat('"$"#,##0');     // Spend
  sh.getRange(15, 3, 4, 1).setNumberFormat('0.0%');         // % Spend
  sh.getRange(15, 4, 4, 1).setNumberFormat('"$"#,##0');     // Sales
  sh.getRange(15, 5, 4, 1).setNumberFormat('0.0%');         // % Sales
  sh.getRange(15, 6, 4, 1).setNumberFormat('#,##0');        // Orders
  sh.getRange(15, 7, 4, 1).setNumberFormat('0.00"x"');      // ROAS
  sh.getRange(15, 8, 4, 1).setNumberFormat('0.0%');         // ACOS

  // Highlight TOTAL row
  const totalRow = 15 + chanData.length - 1;
  sh.getRange(totalRow, 1, 1, chanHeaders.length)
    .setBackground(COLOR_LIGHT).setFontWeight('bold');

  // Conditional ROAS color highlighting
  const roasRange = sh.getRange(15, 7, 3, 1);
  const sales = sh.getRange(15, 4, 3, 1);
  // Highlight strong ROAS (>10x) green
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(10)
    .setBackground(COLOR_GREEN_LT).setFontColor(COLOR_SUCCESS).setBold(true)
    .setRanges([roasRange]).build();
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(5)
    .setBackground(COLOR_RED_LT).setFontColor(COLOR_DANGER).setBold(true)
    .setRanges([roasRange]).build();
  sh.setConditionalFormatRules([rule1, rule2]);

  // Quick insights (right-side commentary)
  setSection_(sh, 21, '🎯 Account Snapshot', 8);
  const insights = generateAccountInsights_(total, sp, sb, sd, data);
  insights.forEach((ins, i) => {
    sh.getRange(22 + i, 1).setValue('•').setFontWeight('bold').setFontColor(COLOR_PRIMARY).setHorizontalAlignment('center');
    sh.getRange(22 + i, 2).setValue(ins).setWrap(true).setVerticalAlignment('top');
    sh.getRange(22 + i, 2, 1, 7).merge();
  });

  // Column widths
  sh.setColumnWidth(1, 40);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 90);
  sh.setColumnWidth(6, 100);
  sh.setColumnWidth(7, 90);
  sh.setColumnWidth(8, 90);
  // After KPI cards, widen columns 1, 3, 5, 7 (label cols of cards)
  sh.setColumnWidths(1, 1, 160);
  sh.setColumnWidths(2, 1, 30);
  sh.setColumnWidths(3, 1, 160);
  sh.setColumnWidths(4, 1, 30);
  sh.setColumnWidths(5, 1, 160);
  sh.setColumnWidths(6, 1, 30);
  sh.setColumnWidths(7, 1, 160);
  sh.setColumnWidths(8, 1, 60);

  // Then re-size the cols used in the channel mix below (overrides the KPI sizing for those cells beyond row 14)
  // Use actual column adjustments based on channel mix data
  // We'll keep the KPI sizes — the channel mix headers are wide enough at 160px
}

function aggregate_(rows) {
  const a = { Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0 };
  rows.forEach(r => {
    a.Impressions += r.Impressions || 0;
    a.Clicks      += r.Clicks      || 0;
    a.Spend       += r.Spend       || 0;
    a.Sales       += r.Sales       || 0;
    a.Orders      += r.Orders      || 0;
  });
  return a;
}

function generateAccountInsights_(total, sp, sb, sd, data) {
  const out = [];
  const roas = safeDiv(total.Sales, total.Spend);

  // Headline
  out.push(`Total: ${fmt$(total.Sales)} in ad sales on ${fmt$(total.Spend)} spend — a ${fmtRoas(roas)} ROAS across ${fmtNum(total.Orders)} orders.`);

  // Channel observation
  if (sd.Spend > 0 && safeDiv(sd.Sales, sd.Spend) > roas * 1.5) {
    out.push(`Sponsored Display is the most efficient channel at ${fmtRoas(safeDiv(sd.Sales, sd.Spend))} — well above the ${fmtRoas(roas)} blended account ROAS. Worth scaling.`);
  }

  // Search terms summary
  if (data.spSearchTerms.length > 0) {
    const sortedSt = data.spSearchTerms.slice().sort((a, b) => b.Sales - a.Sales);
    const top = sortedSt[0];
    out.push(`Top customer search term: "${top.Term}" — ${fmt$(top.Sales)} in sales, ${fmtRoas(safeDiv(top.Sales, top.Spend))} ROAS.`);
  }

  // Wasted spend headline
  const wasted = data.spSearchTerms.filter(t => t.Clicks >= 5 && t.Sales === 0);
  if (wasted.length > 0) {
    const wastedSpend = wasted.reduce((a, t) => a + t.Spend, 0);
    out.push(`Wasted spend opportunity: ${wasted.length} search terms with 5+ clicks and zero sales totalling ${fmt$(wastedSpend)}. See the "⚠️ Negative KWs" tab.`);
  }

  // Top SP campaign
  if (data.spCampaigns.length > 0) {
    const sortedC = data.spCampaigns.slice().sort((a, b) => b.Sales - a.Sales);
    const top = sortedC[0];
    out.push(`Top SP campaign: "${top.Campaign}" — ${fmt$(top.Sales)} in sales at ${fmtRoas(safeDiv(top.Sales, top.Spend))} ROAS.`);
  }

  // Branded vs Non-branded SP keyword (auto-detect brand from most-spent keyword)
  const brandHint = detectBrandHint_(data);
  if (brandHint) {
    const branded = data.spKeywords.filter(k => k.Keyword.toLowerCase().includes(brandHint));
    const nonBranded = data.spKeywords.filter(k => !k.Keyword.toLowerCase().includes(brandHint));
    const brSp = branded.reduce((a, k) => a + k.Spend, 0);
    const brSa = branded.reduce((a, k) => a + k.Sales, 0);
    const nbSp = nonBranded.reduce((a, k) => a + k.Spend, 0);
    const nbSa = nonBranded.reduce((a, k) => a + k.Sales, 0);
    if (brSp > 0 && nbSp > 0) {
      out.push(`Branded SP keywords (containing "${brandHint}") returned ${fmtRoas(safeDiv(brSa, brSp))} vs non-branded at ${fmtRoas(safeDiv(nbSa, nbSp))}. Branded sales = ${fmtPct(safeDiv(brSa, brSa+nbSa))} of keyword-driven sales.`);
    }
  }

  return out;
}

function detectBrandHint_(data) {
  // Best heuristic: take the keyword with the highest sales and find the most common single-word token
  if (data.spKeywords.length === 0) return null;
  const sorted = data.spKeywords.slice().sort((a, b) => b.Sales - a.Sales);
  // Look at top 20 keywords by sales — find the most common non-trivial token
  const tokenCount = {};
  const stopWords = new Set(['for', 'the', 'and', 'with', 'bra', 'bras', 'women', 'men', 'plus', 'size', 'in', 'a', 'an', 'of', 'to']);
  sorted.slice(0, 20).forEach(k => {
    k.Keyword.toLowerCase().split(/\s+/).forEach(tok => {
      if (tok.length < 3 || stopWords.has(tok)) return;
      tokenCount[tok] = (tokenCount[tok] || 0) + (k.Sales || 1);
    });
  });
  let bestTok = null, bestScore = 0;
  Object.entries(tokenCount).forEach(([tok, score]) => {
    if (score > bestScore) { bestScore = score; bestTok = tok; }
  });
  return bestTok;
}


// ============================================================
// TAB 2 — SP CAMPAIGNS
// ============================================================
function buildSPCampaignsTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_SP_CAMPAIGNS);
  setTitle_(sh, 1, 'Sponsored Products — Campaign Performance', 8);

  if (data.spCampaigns.length === 0) {
    sh.getRange(3, 1).setValue('No Sponsored Products campaign data found.').setFontStyle('italic');
    return;
  }

  const headers = ['Campaign Name', 'Targeting', 'State', 'Daily Budget', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'CTR', 'CPC', 'CVR', 'ACOS', 'ROAS'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, 3, 1, headers.length);

  // Sort campaigns by sales descending
  const sorted = data.spCampaigns.slice().sort((a, b) => b.Sales - a.Sales);
  const rows = sorted.map(c => [
    c.Campaign,
    c.TargetingType || '',
    c.State || '',
    c.DailyBudget || 0,
    c.Impressions, c.Clicks, c.Spend, c.Sales, c.Orders,
    safeDiv(c.Clicks, c.Impressions),
    safeDiv(c.Spend, c.Clicks),
    safeDiv(c.Orders, c.Clicks),
    safeDiv(c.Spend, c.Sales),
    safeDiv(c.Sales, c.Spend),
  ]);

  // Total row
  const tot = aggregate_(data.spCampaigns);
  rows.push([
    'TOTAL', '', '', '',
    tot.Impressions, tot.Clicks, tot.Spend, tot.Sales, tot.Orders,
    safeDiv(tot.Clicks, tot.Impressions),
    safeDiv(tot.Spend, tot.Clicks),
    safeDiv(tot.Orders, tot.Clicks),
    safeDiv(tot.Spend, tot.Sales),
    safeDiv(tot.Sales, tot.Spend),
  ]);

  sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  styleDataBlock_(sh, 4, 1, rows.length, headers.length);

  // Number formats
  sh.getRange(4, 4, rows.length, 1).setNumberFormat('"$"#,##0');     // Daily Budget
  sh.getRange(4, 5, rows.length, 1).setNumberFormat('#,##0');        // Impressions
  sh.getRange(4, 6, rows.length, 1).setNumberFormat('#,##0');        // Clicks
  sh.getRange(4, 7, rows.length, 1).setNumberFormat('"$"#,##0.00');  // Spend
  sh.getRange(4, 8, rows.length, 1).setNumberFormat('"$"#,##0.00');  // Sales
  sh.getRange(4, 9, rows.length, 1).setNumberFormat('#,##0');        // Orders
  sh.getRange(4, 10, rows.length, 1).setNumberFormat('0.00%');       // CTR
  sh.getRange(4, 11, rows.length, 1).setNumberFormat('"$"#,##0.00'); // CPC
  sh.getRange(4, 12, rows.length, 1).setNumberFormat('0.0%');        // CVR
  sh.getRange(4, 13, rows.length, 1).setNumberFormat('0.0%');        // ACOS
  sh.getRange(4, 14, rows.length, 1).setNumberFormat('0.00"x"');     // ROAS

  // Highlight total row
  sh.getRange(4 + rows.length - 1, 1, 1, headers.length)
    .setBackground(COLOR_LIGHT).setFontWeight('bold');

  // Color-code ROAS column
  const roasRange = sh.getRange(4, 14, rows.length - 1, 1);
  const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(10)
    .setBackground(COLOR_GREEN_LT).setFontColor(COLOR_SUCCESS).setBold(true)
    .setRanges([roasRange]).build();
  const ruleLow = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(3)
    .setBackground(COLOR_RED_LT).setFontColor(COLOR_DANGER).setBold(true)
    .setRanges([roasRange]).build();
  sh.setConditionalFormatRules([ruleHigh, ruleLow]);

  sh.setColumnWidth(1, 350);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 80);
  for (let c = 4; c <= 9; c++) sh.setColumnWidth(c, 100);
  for (let c = 10; c <= 14; c++) sh.setColumnWidth(c, 80);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);
}


// ============================================================
// TAB 3 — SB & SD CAMPAIGNS
// ============================================================
function buildSBSDTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_SB_SD);
  setTitle_(sh, 1, 'Sponsored Brands & Sponsored Display — Campaigns', 9);
  setSubtitle_(sh, 2, 'SB uses Multi Ad Group attribution when available. SD uses Views & Clicks attribution.', 9);

  let row = 4;
  const headers = ['Campaign', 'Channel', 'Format/Tactic', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'ROAS'];

  // SB section
  setSection_(sh, row, 'Sponsored Brands', 9); row++;
  sh.getRange(row, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, row, 1, headers.length);
  row++;

  // Use SB MAG if present, otherwise legacy
  const sbSource = data.sbMagCampaigns.length > 0 ? data.sbMagCampaigns : data.sbCampaigns;
  const sbSorted = sbSource.slice().sort((a, b) => b.Sales - a.Sales);
  const sbRows = sbSorted.map(c => [
    c.Campaign,
    'SB',
    c.AdFormat || (c.IsMag ? 'Multi Ad Group' : ''),
    c.Impressions, c.Clicks, c.Spend, c.Sales, c.Orders,
    safeDiv(c.Sales, c.Spend),
  ]);
  if (sbRows.length === 0) sbRows.push(['(no SB campaigns)', '', '', 0,0,0,0,0,0]);

  // Add total
  const sbTot = aggregate_(sbSource);
  sbRows.push(['TOTAL SB', '', '', sbTot.Impressions, sbTot.Clicks, sbTot.Spend, sbTot.Sales, sbTot.Orders, safeDiv(sbTot.Sales, sbTot.Spend)]);

  sh.getRange(row, 1, sbRows.length, headers.length).setValues(sbRows);
  styleDataBlock_(sh, row, 1, sbRows.length, headers.length);
  applyCampaignNumberFormats_(sh, row, sbRows.length);
  // Highlight total
  sh.getRange(row + sbRows.length - 1, 1, 1, headers.length).setBackground(COLOR_LIGHT).setFontWeight('bold');
  row += sbRows.length + 2;

  // SD section
  setSection_(sh, row, 'Sponsored Display', 9); row++;
  sh.getRange(row, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, row, 1, headers.length);
  row++;

  const sdSorted = data.sdCampaigns.slice().sort((a, b) => b.Sales - a.Sales);
  const sdRows = sdSorted.map(c => [
    c.Campaign,
    'SD',
    c.Tactic || '',
    c.Impressions, c.Clicks, c.Spend, c.Sales, c.Orders,
    safeDiv(c.Sales, c.Spend),
  ]);
  if (sdRows.length === 0) sdRows.push(['(no SD campaigns)', '', '', 0,0,0,0,0,0]);
  const sdTot = aggregate_(data.sdCampaigns);
  sdRows.push(['TOTAL SD', '', '', sdTot.Impressions, sdTot.Clicks, sdTot.Spend, sdTot.Sales, sdTot.Orders, safeDiv(sdTot.Sales, sdTot.Spend)]);

  sh.getRange(row, 1, sdRows.length, headers.length).setValues(sdRows);
  styleDataBlock_(sh, row, 1, sdRows.length, headers.length);
  applyCampaignNumberFormats_(sh, row, sdRows.length);
  sh.getRange(row + sdRows.length - 1, 1, 1, headers.length).setBackground(COLOR_LIGHT).setFontWeight('bold');

  sh.setColumnWidth(1, 320);
  sh.setColumnWidth(2, 70);
  sh.setColumnWidth(3, 130);
  for (let c = 4; c <= 8; c++) sh.setColumnWidth(c, 100);
  sh.setColumnWidth(9, 80);
  sh.setFrozenRows(3);
}

function applyCampaignNumberFormats_(sh, startRow, numRows) {
  // Columns: 1 Campaign, 2 Channel, 3 Format, 4 Impressions, 5 Clicks, 6 Spend, 7 Sales, 8 Orders, 9 ROAS
  sh.getRange(startRow, 4, numRows, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, 5, numRows, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, 6, numRows, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(startRow, 7, numRows, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(startRow, 8, numRows, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, 9, numRows, 1).setNumberFormat('0.00"x"');
}


// ============================================================
// TAB 4 — TOP ASINS
// ============================================================
function buildTopAsinsTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_TOP_ASINS);
  setTitle_(sh, 1, 'Top Advertised ASINs (Sponsored Products)', 8);

  // Aggregate ASINs across all SP product ad rows
  const asinAgg = {};
  data.spProductAds.forEach(a => {
    if (!a.ASIN) return;
    if (!asinAgg[a.ASIN]) asinAgg[a.ASIN] = {
      ASIN: a.ASIN, SKU: a.SKU || '',
      Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0, Units: 0,
      Campaigns: new Set(),
    };
    const x = asinAgg[a.ASIN];
    x.Impressions += a.Impressions;
    x.Clicks += a.Clicks;
    x.Spend += a.Spend;
    x.Sales += a.Sales;
    x.Orders += a.Orders;
    x.Units += a.Units || 0;
    if (a.Campaign) x.Campaigns.add(a.Campaign);
    if (!x.SKU && a.SKU) x.SKU = a.SKU;
  });

  const sorted = Object.values(asinAgg).sort((a, b) => b.Sales - a.Sales).slice(0, 50);

  if (sorted.length === 0) {
    sh.getRange(3, 1).setValue('No advertised ASIN data found.').setFontStyle('italic');
    return;
  }

  const headers = ['ASIN', 'SKU', 'Campaigns', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'ROAS'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, 3, 1, headers.length);

  const rows = sorted.map(a => [
    a.ASIN, a.SKU,
    a.Campaigns.size + (a.Campaigns.size === 1 ? ' campaign' : ' campaigns'),
    a.Impressions, a.Clicks, a.Spend, a.Sales, a.Orders,
    safeDiv(a.Sales, a.Spend),
  ]);
  sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  styleDataBlock_(sh, 4, 1, rows.length, headers.length);
  sh.getRange(4, 4, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 5, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 6, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(4, 7, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(4, 8, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 9, rows.length, 1).setNumberFormat('0.00"x"');

  // ROAS heat map
  const roasRange = sh.getRange(4, 9, rows.length, 1);
  const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(10).setBackground(COLOR_GREEN_LT).setFontColor(COLOR_SUCCESS).setBold(true)
    .setRanges([roasRange]).build();
  const ruleLow = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(3).setBackground(COLOR_RED_LT).setFontColor(COLOR_DANGER).setBold(true)
    .setRanges([roasRange]).build();
  sh.setConditionalFormatRules([ruleHigh, ruleLow]);

  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 130);
  for (let c = 4; c <= 9; c++) sh.setColumnWidth(c, 100);
  sh.setFrozenRows(3);
}


// ============================================================
// TAB 5 — SEARCH TERMS (top 100 SP)
// ============================================================
function buildSearchTermsTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_SEARCH_TERMS);
  setTitle_(sh, 1, 'Top Customer Search Terms (Sponsored Products)', 9);
  setSubtitle_(sh, 2, `Aggregated from ${data.spSearchTerms.length.toLocaleString()} unique search terms.`, 9);

  if (data.spSearchTerms.length === 0) {
    sh.getRange(3, 1).setValue('No SP search term data found.').setFontStyle('italic');
    return;
  }

  const sorted = data.spSearchTerms.slice().sort((a, b) => b.Sales - a.Sales).slice(0, 100);

  const headers = ['Search Term', 'Match Types', 'Impressions', 'Clicks', 'CTR', 'Spend', 'Sales', 'Orders', 'CVR', 'ACOS', 'ROAS'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, 3, 1, headers.length);

  const rows = sorted.map(t => [
    t.Term, t.MatchTypes,
    t.Impressions, t.Clicks,
    safeDiv(t.Clicks, t.Impressions),
    t.Spend, t.Sales, t.Orders,
    safeDiv(t.Orders, t.Clicks),
    safeDiv(t.Spend, t.Sales),
    safeDiv(t.Sales, t.Spend),
  ]);
  sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  styleDataBlock_(sh, 4, 1, rows.length, headers.length);

  sh.getRange(4, 3, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 4, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 5, rows.length, 1).setNumberFormat('0.00%');
  sh.getRange(4, 6, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(4, 7, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(4, 8, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 9, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(4,10, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(4,11, rows.length, 1).setNumberFormat('0.00"x"');

  const roasRange = sh.getRange(4, 11, rows.length, 1);
  const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(10).setBackground(COLOR_GREEN_LT).setFontColor(COLOR_SUCCESS).setBold(true)
    .setRanges([roasRange]).build();
  const ruleLow = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(3).setBackground(COLOR_RED_LT).setFontColor(COLOR_DANGER).setBold(true)
    .setRanges([roasRange]).build();
  sh.setConditionalFormatRules([ruleHigh, ruleLow]);

  sh.setColumnWidth(1, 360);
  sh.setColumnWidth(2, 130);
  for (let c = 3; c <= 11; c++) sh.setColumnWidth(c, 90);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);
}


// ============================================================
// TAB 6 — NEGATIVE KEYWORD CANDIDATES (the skill's analysis)
// ============================================================
function buildNegativesTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_NEGATIVES);
  setTitle_(sh, 1, '⚠️ Negative Keyword Candidates', 8);
  setSubtitle_(sh, 2, 'Search terms burning spend without converting. Add the high-priority items as negatives in your campaigns.', 8);

  if (data.spSearchTerms.length === 0) {
    sh.getRange(3, 1).setValue('No SP search term data found.').setFontStyle('italic');
    return;
  }

  // Detect brand for filtering
  const brandHint = detectBrandHint_(data);

  // Helper: ASIN-format detector (b0XXXXXXXX)
  const isAsinTerm = t => /^b0[a-z0-9]{8}$/i.test(t);

  // High Priority: 5+ clicks, $0 sales, NOT branded, NOT ASIN
  const high = data.spSearchTerms
    .filter(t => t.Clicks >= 5 && t.Sales === 0 && !isAsinTerm(t.Term) && (!brandHint || !t.Term.toLowerCase().includes(brandHint)))
    .sort((a, b) => b.Spend - a.Spend);

  // Medium Priority: ASIN-format with 5+ clicks and $0 sales
  const mediumAsin = data.spSearchTerms
    .filter(t => t.Clicks >= 5 && t.Sales === 0 && isAsinTerm(t.Term))
    .sort((a, b) => b.Spend - a.Spend);

  // Low Priority: 2-4 clicks, $0 sales (watch list)
  const low = data.spSearchTerms
    .filter(t => t.Clicks >= 2 && t.Clicks < 5 && t.Sales === 0 && !isAsinTerm(t.Term))
    .sort((a, b) => b.Spend - a.Spend)
    .slice(0, 50);

  // High ACOS (sales > 0 but ACOS too high — cut bids rather than fully negate)
  const highAcos = data.spSearchTerms
    .filter(t => t.Sales > 0 && t.Clicks >= 5 && safeDiv(t.Spend, t.Sales) > 0.5)
    .sort((a, b) => safeDiv(b.Spend, b.Sales) - safeDiv(a.Spend, a.Sales))
    .slice(0, 30);

  // Summary box at top
  const totalWasted = data.spSearchTerms.filter(t => t.Sales === 0 && t.Clicks >= 1).reduce((a, t) => a + t.Spend, 0);
  const highWasted = high.reduce((a, t) => a + t.Spend, 0);
  const mediumWasted = mediumAsin.reduce((a, t) => a + t.Spend, 0);
  const totalSpSpend = data.spSearchTerms.reduce((a, t) => a + t.Spend, 0);

  setSection_(sh, 4, '💰 Wasted Spend Summary', 8);
  const summaryHeaders = ['Tier', 'Description', '# Terms', 'Wasted Spend', '% of SP Spend', 'Action'];
  sh.getRange(5, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
  styleHeaderRow_(sh, 5, 1, summaryHeaders.length, COLOR_DANGER);
  const summaryRows = [
    ['🔴 High Priority', '5+ clicks, $0 sales, non-branded, non-ASIN', high.length, highWasted, safeDiv(highWasted, totalSpSpend), 'Negate immediately'],
    ['🟡 Medium Priority', '5+ clicks, $0 sales, competitor ASINs', mediumAsin.length, mediumWasted, safeDiv(mediumWasted, totalSpSpend), 'Negate the highest-spend ASINs'],
    ['🟠 Low Priority', '2-4 clicks, $0 sales (watch list)', low.length, low.reduce((a,t)=>a+t.Spend,0), safeDiv(low.reduce((a,t)=>a+t.Spend,0), totalSpSpend), 'Monitor; negate if spending continues'],
    ['⚠️ High ACOS', 'Selling but ACOS > 50%', highAcos.length, highAcos.reduce((a,t)=>a+t.Spend,0), safeDiv(highAcos.reduce((a,t)=>a+t.Spend,0), totalSpSpend), 'Reduce bids rather than negate'],
    ['TOTAL Wasted', 'All terms with clicks but $0 sales', data.spSearchTerms.filter(t=>t.Sales===0&&t.Clicks>=1).length, totalWasted, safeDiv(totalWasted, totalSpSpend), '—'],
  ];
  sh.getRange(6, 1, summaryRows.length, summaryHeaders.length).setValues(summaryRows);
  styleDataBlock_(sh, 6, 1, summaryRows.length, summaryHeaders.length);
  sh.getRange(6, 3, summaryRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(6, 4, summaryRows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(6, 5, summaryRows.length, 1).setNumberFormat('0.0%');
  // Bold the total row
  sh.getRange(6 + summaryRows.length - 1, 1, 1, summaryHeaders.length)
    .setBackground(COLOR_LIGHT).setFontWeight('bold');

  let row = 6 + summaryRows.length + 2;

  // ----- HIGH PRIORITY TABLE -----
  setSection_(sh, row, `🔴 HIGH PRIORITY — Negate These Now (${high.length} terms, ${fmt$(highWasted)} wasted)`, 8); row++;
  const detailHeaders = ['Search Term', 'Impressions', 'Clicks', 'CPC', 'Spend', 'Reason', 'Match Types', 'Campaigns'];
  sh.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders]);
  styleHeaderRow_(sh, row, 1, detailHeaders.length, COLOR_DANGER);
  row++;

  const highRows = high.slice(0, 100).map(t => [
    t.Term, t.Impressions, t.Clicks, safeDiv(t.Spend, t.Clicks), t.Spend,
    classifyNegativeReason_(t.Term, brandHint),
    t.MatchTypes, t.Campaigns,
  ]);
  if (highRows.length === 0) highRows.push(['(no high-priority candidates)', 0,0,0,0,'','','']);
  sh.getRange(row, 1, highRows.length, detailHeaders.length).setValues(highRows);
  styleDataBlock_(sh, row, 1, highRows.length, detailHeaders.length);
  applyNegativeNumberFormats_(sh, row, highRows.length);
  row += highRows.length + 2;

  // ----- MEDIUM PRIORITY TABLE -----
  setSection_(sh, row, `🟡 MEDIUM PRIORITY — Competitor ASINs (${mediumAsin.length} terms, ${fmt$(mediumWasted)} wasted)`, 8); row++;
  sh.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders]);
  styleHeaderRow_(sh, row, 1, detailHeaders.length, COLOR_WARNING);
  row++;

  const medRows = mediumAsin.slice(0, 50).map(t => [
    t.Term.toUpperCase(), t.Impressions, t.Clicks, safeDiv(t.Spend, t.Clicks), t.Spend,
    'Competitor ASIN — auto-target not converting',
    t.MatchTypes, t.Campaigns,
  ]);
  if (medRows.length === 0) medRows.push(['(no competitor-ASIN candidates)', 0,0,0,0,'','','']);
  sh.getRange(row, 1, medRows.length, detailHeaders.length).setValues(medRows);
  styleDataBlock_(sh, row, 1, medRows.length, detailHeaders.length);
  applyNegativeNumberFormats_(sh, row, medRows.length);
  row += medRows.length + 2;

  // ----- LOW PRIORITY (WATCH LIST) -----
  setSection_(sh, row, `🟠 LOW PRIORITY — Watch List (top 50 by spend)`, 8); row++;
  sh.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders]);
  styleHeaderRow_(sh, row, 1, detailHeaders.length, COLOR_ACCENT);
  row++;

  const lowRows = low.map(t => [
    t.Term, t.Impressions, t.Clicks, safeDiv(t.Spend, t.Clicks), t.Spend,
    classifyNegativeReason_(t.Term, brandHint),
    t.MatchTypes, t.Campaigns,
  ]);
  if (lowRows.length === 0) lowRows.push(['(no watch-list terms)', 0,0,0,0,'','','']);
  sh.getRange(row, 1, lowRows.length, detailHeaders.length).setValues(lowRows);
  styleDataBlock_(sh, row, 1, lowRows.length, detailHeaders.length);
  applyNegativeNumberFormats_(sh, row, lowRows.length);
  row += lowRows.length + 2;

  // ----- HIGH ACOS (BID DOWN) -----
  setSection_(sh, row, `⚠️ HIGH ACOS — Reduce Bids (top 30, ACOS > 50%)`, 8); row++;
  const acosHeaders = ['Search Term', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'ACOS', 'ROAS'];
  sh.getRange(row, 1, 1, acosHeaders.length).setValues([acosHeaders]);
  styleHeaderRow_(sh, row, 1, acosHeaders.length, COLOR_WARNING);
  row++;
  const acosRows = highAcos.map(t => [
    t.Term, t.Impressions, t.Clicks, t.Spend, t.Sales, t.Orders,
    safeDiv(t.Spend, t.Sales),
    safeDiv(t.Sales, t.Spend),
  ]);
  if (acosRows.length === 0) acosRows.push(['(no high-ACOS converting terms)', 0,0,0,0,0,0,0]);
  sh.getRange(row, 1, acosRows.length, acosHeaders.length).setValues(acosRows);
  styleDataBlock_(sh, row, 1, acosRows.length, acosHeaders.length);
  sh.getRange(row, 2, acosRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(row, 3, acosRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(row, 4, acosRows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(row, 5, acosRows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(row, 6, acosRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(row, 7, acosRows.length, 1).setNumberFormat('0.0%');
  sh.getRange(row, 8, acosRows.length, 1).setNumberFormat('0.00"x"');

  // Column widths
  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 80);
  sh.setColumnWidth(4, 80);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 220);
  sh.setColumnWidth(7, 120);
  sh.setColumnWidth(8, 200);
  sh.setFrozenRows(3);
}

function applyNegativeNumberFormats_(sh, startRow, numRows) {
  // Cols: 1 Term, 2 Impressions, 3 Clicks, 4 CPC, 5 Spend
  sh.getRange(startRow, 2, numRows, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, 3, numRows, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, 4, numRows, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(startRow, 5, numRows, 1).setNumberFormat('"$"#,##0.00');
}

function classifyNegativeReason_(term, brandHint) {
  const t = term.toLowerCase();
  if (/^b0[a-z0-9]{8}$/i.test(t)) return 'Competitor ASIN';
  // Common competitor brand names (extend as needed)
  const competitors = [
    'wacoal','panache','freya','fantasie','curvy kate','chantelle','natori','glamorise',
    'playtex','bali','warner','calvin klein','victoria','hanes','triumph','simone perele',
    'spanx','maidenform','le mystere','dkny','tommy','ralph lauren','goddess','cake'
  ];
  for (const c of competitors) {
    if (t.includes(c)) return `Competitor brand: ${c}`;
  }
  // Off-category attribute hints
  if (/minimi[sz]er/.test(t)) return 'Wrong product attribute (minimizer)';
  if (/nursing|maternity|pumping/.test(t)) return 'Wrong product attribute (nursing/maternity)';
  if (/strapless/.test(t)) return 'Possibly wrong product attribute (strapless)';
  if (/sport[s]?\s|athletic|workout|gym|running|yoga/.test(t)) return 'Wrong category (sport/athletic)';
  if (/men[\'s]*\s|male/.test(t)) return 'Wrong audience (men)';
  // Length-based heuristic
  if (t.split(/\s+/).length <= 2) return 'Too generic';
  return 'Non-converting — review';
}


// ============================================================
// TAB 7 — TOP PERFORMERS (winning search terms to harvest)
// ============================================================
function buildTopPerformersTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_TOP_PERF);
  setTitle_(sh, 1, '🏆 Top Performing Search Terms — Harvest These', 9);
  setSubtitle_(sh, 2, 'Search terms with strong sales and ROAS. Move the best ones into exact-match keyword campaigns to lock in the win.', 9);

  if (data.spSearchTerms.length === 0) {
    sh.getRange(3, 1).setValue('No SP search term data found.').setFontStyle('italic');
    return;
  }

  // Top 50 by sales, with ROAS >= 3x and at least 1 order
  const top = data.spSearchTerms
    .filter(t => t.Orders >= 1 && safeDiv(t.Sales, t.Spend) >= 3)
    .sort((a, b) => b.Sales - a.Sales)
    .slice(0, 50);

  const headers = ['Search Term', 'Match Types', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'CVR', 'ACOS', 'ROAS', 'Action'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, 3, 1, headers.length, COLOR_SUCCESS);

  const rows = top.map(t => {
    const roas = safeDiv(t.Sales, t.Spend);
    const acos = safeDiv(t.Spend, t.Sales);
    let action = 'Harvest as exact match';
    if (t.MatchTypes && t.MatchTypes.toLowerCase().includes('exact')) action = 'Already exact — increase bid';
    else if (roas > 15) action = '🔥 High ROAS — harvest as exact, scale aggressively';
    return [
      t.Term, t.MatchTypes,
      t.Impressions, t.Clicks, t.Spend, t.Sales, t.Orders,
      safeDiv(t.Orders, t.Clicks),
      acos, roas,
      action,
    ];
  });
  if (rows.length === 0) rows.push(['(no qualifying top performers — all converting terms have ROAS < 3x)', '',0,0,0,0,0,0,0,0,'']);

  sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  styleDataBlock_(sh, 4, 1, rows.length, headers.length);
  sh.getRange(4, 3, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 4, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 5, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(4, 6, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(4, 7, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 8, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(4, 9, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(4,10, rows.length, 1).setNumberFormat('0.00"x"');

  // Highlight ROAS column with a heat map
  const roasRange = sh.getRange(4, 10, rows.length, 1);
  const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(15).setBackground(COLOR_GREEN_LT).setFontColor(COLOR_SUCCESS).setBold(true)
    .setRanges([roasRange]).build();
  sh.setConditionalFormatRules([ruleHigh]);

  sh.setColumnWidth(1, 320);
  sh.setColumnWidth(2, 110);
  for (let c = 3; c <= 10; c++) sh.setColumnWidth(c, 90);
  sh.setColumnWidth(11, 280);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);
}


// ============================================================
// TAB 8 — MATCH TYPES
// ============================================================
function buildMatchTypesTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_MATCH_TYPES);
  setTitle_(sh, 1, 'Match Type Performance Comparison', 8);
  setSubtitle_(sh, 2, 'Aggregated from the SP Search Term Report. Reveals where the budget is flowing across Auto, Phrase, Broad, and Exact.', 8);

  if (Object.keys(data.spMatchTypes).length === 0) {
    sh.getRange(3, 1).setValue('No match type data found.').setFontStyle('italic');
    return;
  }

  const headers = ['Match Type', 'Impressions', 'Clicks', 'CTR', 'Spend', 'Sales', 'Orders', 'CVR', 'ACOS', 'ROAS'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, 3, 1, headers.length);

  const sorted = Object.values(data.spMatchTypes).sort((a, b) => b.Sales - a.Sales);
  const rows = sorted.map(m => [
    m.MatchType, m.Impressions, m.Clicks,
    safeDiv(m.Clicks, m.Impressions),
    m.Spend, m.Sales, m.Orders,
    safeDiv(m.Orders, m.Clicks),
    safeDiv(m.Spend, m.Sales),
    safeDiv(m.Sales, m.Spend),
  ]);

  // Total row
  const tot = sorted.reduce((a, m) => ({
    Impressions: a.Impressions + m.Impressions,
    Clicks:      a.Clicks + m.Clicks,
    Spend:       a.Spend + m.Spend,
    Sales:       a.Sales + m.Sales,
    Orders:      a.Orders + m.Orders,
  }), { Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0 });
  rows.push([
    'TOTAL', tot.Impressions, tot.Clicks,
    safeDiv(tot.Clicks, tot.Impressions),
    tot.Spend, tot.Sales, tot.Orders,
    safeDiv(tot.Orders, tot.Clicks),
    safeDiv(tot.Spend, tot.Sales),
    safeDiv(tot.Sales, tot.Spend),
  ]);

  sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  styleDataBlock_(sh, 4, 1, rows.length, headers.length);
  sh.getRange(4, 2, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 3, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 4, rows.length, 1).setNumberFormat('0.00%');
  sh.getRange(4, 5, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(4, 6, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(4, 7, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(4, 8, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(4, 9, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(4, 10, rows.length, 1).setNumberFormat('0.00"x"');

  // Total row formatting
  sh.getRange(4 + rows.length - 1, 1, 1, headers.length)
    .setBackground(COLOR_LIGHT).setFontWeight('bold');

  // ROAS heat map
  const roasRange = sh.getRange(4, 10, rows.length - 1, 1);
  const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(10).setBackground(COLOR_GREEN_LT).setFontColor(COLOR_SUCCESS).setBold(true)
    .setRanges([roasRange]).build();
  const ruleLow = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(3).setBackground(COLOR_RED_LT).setFontColor(COLOR_DANGER).setBold(true)
    .setRanges([roasRange]).build();
  sh.setConditionalFormatRules([ruleHigh, ruleLow]);

  sh.setColumnWidth(1, 140);
  for (let c = 2; c <= 10; c++) sh.setColumnWidth(c, 110);
  sh.setFrozenRows(3);
}


// ============================================================
// TAB — BEST KEYWORDS (your targeting keywords across SP + SB)
// ============================================================
function buildBestKeywordsTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_BEST_KEYWORDS);
  setTitle_(sh, 1, '🎯 Best Performing Keywords (Your Targeting Keywords)', 11);
  setSubtitle_(sh, 2, 'Targeting keywords from your campaigns — sorted by sales. Includes SP and SB. To see what shoppers actually typed, use the Search Terms tab.', 11);

  // Combine SP and SB keywords; aggregate by lowercased keyword + match type
  const agg = {};
  const addRow = (k, channel) => {
    const key = (k.Keyword.toLowerCase() + '|' + (k.MatchType || '') + '|' + channel);
    if (!agg[key]) agg[key] = {
      Keyword: k.Keyword.toLowerCase(),
      MatchType: k.MatchType || '',
      Channel: channel,
      Type: '',
      Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0,
      Campaigns: new Set(),
    };
    const a = agg[key];
    a.Impressions += k.Impressions || 0;
    a.Clicks      += k.Clicks      || 0;
    a.Spend       += k.Spend       || 0;
    a.Sales       += k.Sales       || 0;
    a.Orders      += k.Orders      || 0;
    if (k.Campaign) a.Campaigns.add(k.Campaign);
  };
  data.spKeywords.forEach(k => addRow(k, 'SP'));
  data.sbKeywords.forEach(k => addRow(k, 'SB'));

  if (Object.keys(agg).length === 0) {
    sh.getRange(3, 1).setValue('No keyword data found.').setFontStyle('italic');
    return;
  }

  // Auto-detect brand for branded vs non-branded labelling
  const brandHint = detectBrandHint_(data);
  Object.values(agg).forEach(a => {
    a.Type = (brandHint && a.Keyword.includes(brandHint)) ? 'Branded' : 'Non-Branded';
  });

  // Compute totals for branded vs non-branded summary
  const allKws = Object.values(agg);
  const branded = allKws.filter(k => k.Type === 'Branded');
  const nonBranded = allKws.filter(k => k.Type === 'Non-Branded');
  const sumOf = (arr, f) => arr.reduce((s, x) => s + f(x), 0);

  // ---- Summary block ----
  setSection_(sh, 4, 'Branded vs Non-Branded Keyword Summary', 11);
  const summHeaders = ['Type', '# Keywords', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'CTR', 'CVR', 'ACOS', 'ROAS'];
  sh.getRange(5, 1, 1, summHeaders.length).setValues([summHeaders]);
  styleHeaderRow_(sh, 5, 1, summHeaders.length);

  function totalsRow(label, arr) {
    const im = sumOf(arr, x => x.Impressions);
    const cl = sumOf(arr, x => x.Clicks);
    const sp = sumOf(arr, x => x.Spend);
    const sa = sumOf(arr, x => x.Sales);
    const or_ = sumOf(arr, x => x.Orders);
    return [label, arr.length, im, cl, sp, sa, or_,
      safeDiv(cl, im), safeDiv(or_, cl), safeDiv(sp, sa), safeDiv(sa, sp)];
  }
  const summaryRows = [
    totalsRow(`Branded${brandHint ? ' (contains "' + brandHint + '")' : ''}`, branded),
    totalsRow('Non-Branded', nonBranded),
    totalsRow('TOTAL', allKws),
  ];
  sh.getRange(6, 1, summaryRows.length, summHeaders.length).setValues(summaryRows);
  styleDataBlock_(sh, 6, 1, summaryRows.length, summHeaders.length);
  // Number formats
  sh.getRange(6, 2, summaryRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(6, 3, summaryRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(6, 4, summaryRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(6, 5, summaryRows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(6, 6, summaryRows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(6, 7, summaryRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(6, 8, summaryRows.length, 1).setNumberFormat('0.00%');
  sh.getRange(6, 9, summaryRows.length, 1).setNumberFormat('0.0%');
  sh.getRange(6, 10, summaryRows.length, 1).setNumberFormat('0.0%');
  sh.getRange(6, 11, summaryRows.length, 1).setNumberFormat('0.00"x"');
  sh.getRange(6 + summaryRows.length - 1, 1, 1, summHeaders.length)
    .setBackground(COLOR_LIGHT).setFontWeight('bold');

  let row = 6 + summaryRows.length + 2;

  // ---- Top 50 keywords by sales ----
  setSection_(sh, row, 'Top 50 Keywords by Sales', 11); row++;
  const headers = ['Keyword', 'Type', 'Channel', 'Match Type', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'ACOS', 'ROAS'];
  sh.getRange(row, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, row, 1, headers.length);
  row++;

  const sorted = allKws.slice().sort((a, b) => b.Sales - a.Sales).slice(0, 50);
  const rows = sorted.map(k => [
    k.Keyword, k.Type, k.Channel, k.MatchType,
    k.Impressions, k.Clicks, k.Spend, k.Sales, k.Orders,
    safeDiv(k.Spend, k.Sales),
    safeDiv(k.Sales, k.Spend),
  ]);
  if (rows.length === 0) rows.push(['(no keywords with data)','','','',0,0,0,0,0,0,0]);
  sh.getRange(row, 1, rows.length, headers.length).setValues(rows);
  styleDataBlock_(sh, row, 1, rows.length, headers.length);
  sh.getRange(row, 5, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(row, 6, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(row, 7, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(row, 8, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(row, 9, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(row,10, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(row,11, rows.length, 1).setNumberFormat('0.00"x"');

  // ROAS heat map
  const roasRange = sh.getRange(row, 11, rows.length, 1);
  const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(10).setBackground(COLOR_GREEN_LT).setFontColor(COLOR_SUCCESS).setBold(true)
    .setRanges([roasRange]).build();
  const ruleLow = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(3).setBackground(COLOR_RED_LT).setFontColor(COLOR_DANGER).setBold(true)
    .setRanges([roasRange]).build();
  sh.setConditionalFormatRules([ruleHigh, ruleLow]);

  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 70);
  sh.setColumnWidth(4, 90);
  for (let c = 5; c <= 11; c++) sh.setColumnWidth(c, 95);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);
}


// ============================================================
// TAB — BEST TARGETS (Product Targets + Audience Targets)
// ============================================================
function buildBestTargetsTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_BEST_TARGETS);
  setTitle_(sh, 1, '🎯 Best Performing Targets (Product & Audience)', 9);
  setSubtitle_(sh, 2, 'SP product targets and SD audience targets — your non-keyword targeting expressions. Useful for finding which competitor ASINs and audience segments work.', 9);

  let row = 4;

  // ---- SP PRODUCT TARGETS ----
  setSection_(sh, row, 'Sponsored Products — Product Targets', 9); row++;
  const ptHeaders = ['Targeting Expression', 'Resolved Expression', 'Campaign', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'ROAS'];
  sh.getRange(row, 1, 1, ptHeaders.length).setValues([ptHeaders]);
  styleHeaderRow_(sh, row, 1, ptHeaders.length);
  row++;

  // Aggregate SP product targets by resolved expression
  const ptAgg = {};
  data.spProductTargets.forEach(t => {
    const key = t.ResolvedExpr || t.TargetingExpr;
    if (!key) return;
    if (!ptAgg[key]) ptAgg[key] = {
      Expr: t.TargetingExpr, Resolved: t.ResolvedExpr,
      Campaign: t.Campaign,
      Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0,
    };
    const a = ptAgg[key];
    a.Impressions += t.Impressions || 0;
    a.Clicks += t.Clicks || 0;
    a.Spend += t.Spend || 0;
    a.Sales += t.Sales || 0;
    a.Orders += t.Orders || 0;
  });
  const ptSorted = Object.values(ptAgg).sort((a, b) => b.Sales - a.Sales).slice(0, 50);
  const ptRows = ptSorted.map(t => [
    t.Expr || '', t.Resolved || '', t.Campaign || '',
    t.Impressions, t.Clicks, t.Spend, t.Sales, t.Orders,
    safeDiv(t.Sales, t.Spend),
  ]);
  if (ptRows.length === 0) ptRows.push(['(no SP product targets)', '', '', 0,0,0,0,0,0]);
  sh.getRange(row, 1, ptRows.length, ptHeaders.length).setValues(ptRows);
  styleDataBlock_(sh, row, 1, ptRows.length, ptHeaders.length);
  applyTargetNumberFormats_(sh, row, ptRows.length);
  row += ptRows.length + 2;

  // ---- SD AUDIENCE TARGETS ----
  setSection_(sh, row, 'Sponsored Display — Audience Targets', 9); row++;
  sh.getRange(row, 1, 1, ptHeaders.length).setValues([ptHeaders]);
  styleHeaderRow_(sh, row, 1, ptHeaders.length);
  row++;

  const atAgg = {};
  data.sdAudienceTargets.forEach(t => {
    const key = t.ResolvedExpr || t.TargetingExpr;
    if (!key) return;
    if (!atAgg[key]) atAgg[key] = {
      Expr: t.TargetingExpr, Resolved: t.ResolvedExpr,
      Campaign: t.Campaign,
      Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0,
    };
    const a = atAgg[key];
    a.Impressions += t.Impressions || 0;
    a.Clicks += t.Clicks || 0;
    a.Spend += t.Spend || 0;
    a.Sales += t.Sales || 0;
    a.Orders += t.Orders || 0;
  });
  const atSorted = Object.values(atAgg).sort((a, b) => b.Sales - a.Sales).slice(0, 50);
  const atRows = atSorted.map(t => [
    t.Expr || '', t.Resolved || '', t.Campaign || '',
    t.Impressions, t.Clicks, t.Spend, t.Sales, t.Orders,
    safeDiv(t.Sales, t.Spend),
  ]);
  if (atRows.length === 0) atRows.push(['(no SD audience targets)', '', '', 0,0,0,0,0,0]);
  sh.getRange(row, 1, atRows.length, ptHeaders.length).setValues(atRows);
  styleDataBlock_(sh, row, 1, atRows.length, ptHeaders.length);
  applyTargetNumberFormats_(sh, row, atRows.length);

  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 280);
  sh.setColumnWidth(3, 180);
  for (let c = 4; c <= 8; c++) sh.setColumnWidth(c, 95);
  sh.setColumnWidth(9, 80);
  sh.setFrozenRows(3);
}

function applyTargetNumberFormats_(sh, startRow, numRows) {
  // Cols: 1 Expr, 2 Resolved, 3 Campaign, 4 Impressions, 5 Clicks, 6 Spend, 7 Sales, 8 Orders, 9 ROAS
  sh.getRange(startRow, 4, numRows, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, 5, numRows, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, 6, numRows, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(startRow, 7, numRows, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(startRow, 8, numRows, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, 9, numRows, 1).setNumberFormat('0.00"x"');
}


// ============================================================
// TAB — PLACEMENTS (Top of Search / Rest of Search / Product Pages)
// ============================================================
function buildPlacementsTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_PLACEMENTS);
  setTitle_(sh, 1, '📍 Placement Performance', 10);
  setSubtitle_(sh, 2, 'Where your Sponsored Products ads are showing — Top of Search, Product Pages, Rest of Search, Amazon Business. Use this to set placement bid multipliers.', 10);

  if (data.spPlacements.length === 0) {
    sh.getRange(3, 1).setValue('No placement (Bidding Adjustment) data found.').setFontStyle('italic');
    return;
  }

  // Aggregate placements by placement type (across all campaigns)
  const placeAgg = {};
  data.spPlacements.forEach(p => {
    const key = p.Placement;
    if (!placeAgg[key]) placeAgg[key] = {
      Placement: key,
      Impressions: 0, Clicks: 0, Spend: 0, Sales: 0, Orders: 0,
    };
    const a = placeAgg[key];
    a.Impressions += p.Impressions || 0;
    a.Clicks      += p.Clicks      || 0;
    a.Spend       += p.Spend       || 0;
    a.Sales       += p.Sales       || 0;
    a.Orders      += p.Orders      || 0;
  });

  // ---- Account-level placement summary ----
  setSection_(sh, 4, 'Account-Level Placement Summary', 10);
  const summHeaders = ['Placement', 'Impressions', '% Impressions', 'Clicks', 'CTR', 'Spend', 'Sales', 'Orders', 'ACOS', 'ROAS'];
  sh.getRange(5, 1, 1, summHeaders.length).setValues([summHeaders]);
  styleHeaderRow_(sh, 5, 1, summHeaders.length);

  const placeArr = Object.values(placeAgg).sort((a, b) => b.Sales - a.Sales);
  const totalImpr = placeArr.reduce((s, p) => s + p.Impressions, 0);
  const totalClicks = placeArr.reduce((s, p) => s + p.Clicks, 0);
  const totalSpend = placeArr.reduce((s, p) => s + p.Spend, 0);
  const totalSales = placeArr.reduce((s, p) => s + p.Sales, 0);
  const totalOrders = placeArr.reduce((s, p) => s + p.Orders, 0);

  const placeRows = placeArr.map(p => [
    p.Placement, p.Impressions,
    safeDiv(p.Impressions, totalImpr),
    p.Clicks, safeDiv(p.Clicks, p.Impressions),
    p.Spend, p.Sales, p.Orders,
    safeDiv(p.Spend, p.Sales),
    safeDiv(p.Sales, p.Spend),
  ]);
  // Total row
  placeRows.push([
    'TOTAL', totalImpr, 1, totalClicks,
    safeDiv(totalClicks, totalImpr),
    totalSpend, totalSales, totalOrders,
    safeDiv(totalSpend, totalSales),
    safeDiv(totalSales, totalSpend),
  ]);
  sh.getRange(6, 1, placeRows.length, summHeaders.length).setValues(placeRows);
  styleDataBlock_(sh, 6, 1, placeRows.length, summHeaders.length);
  sh.getRange(6, 2, placeRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(6, 3, placeRows.length, 1).setNumberFormat('0.0%');
  sh.getRange(6, 4, placeRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(6, 5, placeRows.length, 1).setNumberFormat('0.00%');
  sh.getRange(6, 6, placeRows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(6, 7, placeRows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(6, 8, placeRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(6, 9, placeRows.length, 1).setNumberFormat('0.0%');
  sh.getRange(6,10, placeRows.length, 1).setNumberFormat('0.00"x"');
  sh.getRange(6 + placeRows.length - 1, 1, 1, summHeaders.length)
    .setBackground(COLOR_LIGHT).setFontWeight('bold');

  // ROAS heat map for the data rows (excluding total)
  if (placeRows.length > 1) {
    const roasRange = sh.getRange(6, 10, placeRows.length - 1, 1);
    const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(15).setBackground(COLOR_GREEN_LT).setFontColor(COLOR_SUCCESS).setBold(true)
      .setRanges([roasRange]).build();
    const ruleLow = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(5).setBackground(COLOR_RED_LT).setFontColor(COLOR_DANGER).setBold(true)
      .setRanges([roasRange]).build();
    sh.setConditionalFormatRules([ruleHigh, ruleLow]);
  }

  let row = 6 + placeRows.length + 2;

  // ---- Per-campaign placement performance ----
  setSection_(sh, row, 'Per-Campaign Placement Performance', 10); row++;
  const detailHeaders = ['Campaign', 'Placement', 'Bid Strategy', 'Multiplier', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'ROAS'];
  sh.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders]);
  styleHeaderRow_(sh, row, 1, detailHeaders.length);
  row++;

  // Sort: campaign asc, then placement
  const placementOrder = { 'Placement Top': 1, 'Placement Rest Of Search': 2, 'Placement Product Page': 3, 'Placement Amazon Business': 4 };
  const sorted = data.spPlacements.slice().sort((a, b) => {
    if (a.Campaign !== b.Campaign) return a.Campaign.localeCompare(b.Campaign);
    return (placementOrder[a.Placement] || 99) - (placementOrder[b.Placement] || 99);
  });

  // Render with subtle visual grouping by campaign
  const detailRows = sorted.map(p => [
    p.Campaign, p.Placement, p.BiddingStrategy, p.Percentage / 100,
    p.Impressions, p.Clicks, p.Spend, p.Sales, p.Orders,
    safeDiv(p.Sales, p.Spend),
  ]);
  if (detailRows.length === 0) detailRows.push(['(no per-campaign placement data)', '','',0,0,0,0,0,0,0]);
  sh.getRange(row, 1, detailRows.length, detailHeaders.length).setValues(detailRows);
  styleDataBlock_(sh, row, 1, detailRows.length, detailHeaders.length);
  sh.getRange(row, 4, detailRows.length, 1).setNumberFormat('+0%;-0%;0%');  // multiplier (bid adjustment)
  sh.getRange(row, 5, detailRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(row, 6, detailRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(row, 7, detailRows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(row, 8, detailRows.length, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(row, 9, detailRows.length, 1).setNumberFormat('#,##0');
  sh.getRange(row,10, detailRows.length, 1).setNumberFormat('0.00"x"');

  // Visual grouping: alternate background by campaign
  let lastCampaign = '';
  let stripe = false;
  for (let i = 0; i < detailRows.length; i++) {
    const camp = detailRows[i][0];
    if (camp !== lastCampaign) {
      stripe = !stripe;
      lastCampaign = camp;
    }
    if (stripe) {
      sh.getRange(row + i, 1, 1, detailHeaders.length).setBackground(COLOR_LIGHTER);
    } else {
      sh.getRange(row + i, 1, 1, detailHeaders.length).setBackground('#FFFFFF');
    }
  }

  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 90);
  for (let c = 5; c <= 9; c++) sh.setColumnWidth(c, 100);
  sh.setColumnWidth(10, 80);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);
}


// ============================================================
// TAB 9 — RECOMMENDATIONS
// ============================================================
function buildRecommendationsTab_(ss, data) {
  const sh = getOrCreateTab_(ss, OUT_RECOMMEND);
  setTitle_(sh, 1, '💡 Strategic Recommendations', 4);
  setSubtitle_(sh, 2, 'Action-oriented next steps based on the data analyzed.', 4);

  // Calculate inputs
  const sp = aggregate_(data.spCampaigns);
  const sb = data.sbMagCampaigns.length > 0 ? aggregate_(data.sbMagCampaigns) : aggregate_(data.sbCampaigns);
  const sd = aggregate_(data.sdCampaigns);
  const total = {
    Spend: sp.Spend + sb.Spend + sd.Spend,
    Sales: sp.Sales + sb.Sales + sd.Sales,
  };
  const blendedRoas = safeDiv(total.Sales, total.Spend);

  const recs = [];

  // 1. Wasted spend / negatives
  const wasted = data.spSearchTerms.filter(t => t.Clicks >= 5 && t.Sales === 0);
  const wastedSpend = wasted.reduce((a, t) => a + t.Spend, 0);
  if (wastedSpend > 0) {
    recs.push({
      priority: 1,
      title: `Eliminate ${fmt$(wastedSpend)} of wasted spend by adding negatives`,
      detail: `${wasted.length} search terms received 5+ clicks but produced zero sales. Adding these as negatives is the single highest-ROI optimization. See the "⚠️ Negative KWs" tab for the prioritized list.`,
      impact: 'High',
    });
  }

  // 2. Top performers harvest
  const harvestable = data.spSearchTerms
    .filter(t => t.Orders >= 2 && safeDiv(t.Sales, t.Spend) >= 5 && (!t.MatchTypes || !t.MatchTypes.toLowerCase().includes('exact')))
    .sort((a, b) => b.Sales - a.Sales);
  if (harvestable.length > 0) {
    const harvestSales = harvestable.slice(0, 30).reduce((a, t) => a + t.Sales, 0);
    recs.push({
      priority: 2,
      title: `Harvest ${harvestable.length} top-performing search terms into exact-match`,
      detail: `These terms are converting strongly through Auto, Phrase, or Broad campaigns. Moving them to dedicated exact-match keyword campaigns lets you control bids precisely and protect the volume. Top 30 represent ${fmt$(harvestSales)} in sales already proven. See "🏆 Top Performers" tab.`,
      impact: 'High',
    });
  }

  // 3. SD scale-up if it's the most efficient channel
  if (sd.Spend > 0) {
    const sdRoas = safeDiv(sd.Sales, sd.Spend);
    if (sdRoas > blendedRoas * 1.3 && sd.Spend < total.Spend * 0.20) {
      const incrementalSpend = sd.Spend * 0.5;
      const incrementalSales = incrementalSpend * sdRoas * 0.6; // assume 60% retention
      recs.push({
        priority: 3,
        title: `Scale Sponsored Display by ~50%`,
        detail: `SD is delivering ${fmtRoas(sdRoas)} ROAS — well above the ${fmtRoas(blendedRoas)} account average — and currently represents only ${fmtPct(safeDiv(sd.Spend, total.Spend))} of spend. A 50% increase (~${fmt$(incrementalSpend)} additional spend) could deliver roughly ${fmt$(incrementalSales)} incremental sales even at 60% efficiency retention.`,
        impact: 'High',
      });
    }
  }

  // 4. Match type observation
  const mt = data.spMatchTypes;
  if (mt['Auto'] && mt['Phrase']) {
    const autoRoas = safeDiv(mt['Auto'].Sales, mt['Auto'].Spend);
    const phraseRoas = safeDiv(mt['Phrase'].Sales, mt['Phrase'].Spend);
    if (autoRoas > phraseRoas * 1.3) {
      recs.push({
        priority: 4,
        title: 'Auto campaigns are outperforming Phrase — mine them harder',
        detail: `Auto: ${fmtRoas(autoRoas)} ROAS  |  Phrase: ${fmtRoas(phraseRoas)} ROAS. Run a weekly "search term harvest" routine on Auto campaigns to convert their winners into structured Phrase/Exact targets, then negative-out the harvested terms in the Auto campaigns to avoid cannibalization.`,
        impact: 'Medium',
      });
    }
  }

  // 5. Brand defense
  const brandHint = detectBrandHint_(data);
  if (brandHint) {
    const branded = data.spSearchTerms.filter(t => t.Term.toLowerCase().includes(brandHint));
    const brSales = branded.reduce((a, t) => a + t.Sales, 0);
    const brShare = safeDiv(brSales, total.Sales);
    if (brShare > 0.5) {
      recs.push({
        priority: 5,
        title: `Branded search ("${brandHint}") drives ${fmtPct(brShare)} of sales — protect it`,
        detail: `Maintain top-of-page branded coverage on Sponsored Brands, defend against competitor bidding on your branded terms, and ensure SP and SD are not allowing competitors to outbid you on your own brand SERP.`,
        impact: 'Medium',
      });
    }
  }

  // 6. Underperforming campaigns
  const underperformers = data.spCampaigns
    .filter(c => c.Spend > 100 && safeDiv(c.Sales, c.Spend) < 2)
    .sort((a, b) => b.Spend - a.Spend);
  if (underperformers.length > 0) {
    const upSpend = underperformers.reduce((a, c) => a + c.Spend, 0);
    recs.push({
      priority: 6,
      title: `Review ${underperformers.length} SP campaigns running below 2x ROAS`,
      detail: `These campaigns combined consumed ${fmt$(upSpend)} but returned less than $2 for every $1 spent. Either reduce bids, narrow targeting, or pause low performers. See the "SP Campaigns" tab for the full list (any rows highlighted red).`,
      impact: 'Medium',
    });
  }

  // 7. Placement bid multiplier opportunity
  if (data.spPlacements.length > 0) {
    const placeAgg = {};
    data.spPlacements.forEach(p => {
      if (!placeAgg[p.Placement]) placeAgg[p.Placement] = { Spend: 0, Sales: 0 };
      placeAgg[p.Placement].Spend += p.Spend;
      placeAgg[p.Placement].Sales += p.Sales;
    });
    const placeRoas = {};
    Object.entries(placeAgg).forEach(([k, v]) => { placeRoas[k] = safeDiv(v.Sales, v.Spend); });
    const top = placeRoas['Placement Top'] || 0;
    const rest = placeRoas['Placement Rest Of Search'] || 0;
    const pp = placeRoas['Placement Product Page'] || 0;
    const sortedP = Object.entries(placeRoas).filter(([k,v]) => v > 0).sort((a,b) => b[1] - a[1]);
    if (sortedP.length >= 2 && sortedP[0][1] > sortedP[sortedP.length-1][1] * 1.5) {
      recs.push({
        priority: 7,
        title: `Apply placement bid multipliers — ${sortedP[0][0].replace('Placement ','')} is your strongest`,
        detail: `${sortedP[0][0].replace('Placement ','')} returns ${fmtRoas(sortedP[0][1])} vs ${sortedP[sortedP.length-1][0].replace('Placement ','')} at ${fmtRoas(sortedP[sortedP.length-1][1])}. Apply a positive bid adjustment (+25–50%) to the strongest placement and a negative adjustment to the weakest. See the "📍 Placements" tab for the full breakdown.`,
        impact: 'Medium',
      });
    }
  }

  // 8. Generic fallback
  if (recs.length < 3) {
    recs.push({
      priority: recs.length + 1,
      title: 'Run a weekly search-term review',
      detail: 'Re-run this analysis weekly. The negatives accumulate, the harvest list grows, and the account compounds in efficiency — that\'s the playbook for moving ROAS from average to excellent.',
      impact: 'Habit',
    });
  }

  // Render recommendations
  setSection_(sh, 4, 'Prioritised Action List', 4);
  const headers = ['#', 'Recommendation', 'Detail', 'Impact'];
  sh.getRange(5, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sh, 5, 1, headers.length);

  const rows = recs.map(r => [r.priority, r.title, r.detail, r.impact]);
  sh.getRange(6, 1, rows.length, headers.length).setValues(rows);
  styleDataBlock_(sh, 6, 1, rows.length, headers.length);
  sh.getRange(6, 1, rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontColor(COLOR_PRIMARY);
  sh.getRange(6, 2, rows.length, 1).setFontWeight('bold');
  sh.getRange(6, 3, rows.length, 1).setWrap(true).setVerticalAlignment('top');
  sh.getRange(6, 4, rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold');

  // Color-code Impact column
  const impactRange = sh.getRange(6, 4, rows.length, 1);
  const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('High')
    .setBackground(COLOR_GREEN_LT).setFontColor(COLOR_SUCCESS)
    .setRanges([impactRange]).build();
  const ruleMed = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Medium')
    .setBackground(COLOR_GOLD_LT).setFontColor(COLOR_WARNING)
    .setRanges([impactRange]).build();
  sh.setConditionalFormatRules([ruleHigh, ruleMed]);

  // Set row heights based on detail length
  rows.forEach((r, i) => {
    const detailLen = String(r[2]).length;
    sh.setRowHeight(6 + i, Math.max(50, Math.min(150, 25 + Math.ceil(detailLen / 80) * 18)));
  });

  sh.setColumnWidth(1, 50);
  sh.setColumnWidth(2, 320);
  sh.setColumnWidth(3, 540);
  sh.setColumnWidth(4, 90);
  sh.setFrozenRows(5);
}
