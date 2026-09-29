/**
 * Generate_Database.ts — Microsoft Excel Automate (Office Scripts)
 * ================================================================
 * Converted from generate_database.py
 * 
 * Generates the complete 7-sheet Executive Re-Inspection Database Workbook:
 *   1. 're-vh'   — Factory Audit Dashboard for VH (Vietnam 1)
 *   2. 're-vh2'  — Factory Audit Dashboard for VH2 (Vietnam 2)
 *   3. 'RE-JV'   — Factory Audit Dashboard for JV (Indonesia 1)
 *   4. 'RE-JV2'  — Factory Audit Dashboard for JV2 (Indonesia 2)
 *   5. 'VN-INDO' — Executive Regional Comparison (Vietnam vs Indonesia)
 *   6. 'QC'      — QC Audit & FTT Cross-Analysis Master Log & Side Cards
 *   7. 'MA'      — MA Re-Audit & HFPA Cross-Analysis Master Log & Side Cards
 * 
 * Works in:
 *   - Microsoft 365 Excel Automate Tab (Office Scripts Code Editor)
 *   - Power Automate (Excel Online Run Script action)
 */

type RowRecord = { [key: string]: string | number | boolean | null | undefined };

interface SiteDatabaseMetrics {
  site: string;
  qc_total_fail: number;
  top_countries: [string, number][];
  top_models: [string, number][];
  top5_models_list: string[];
  models_alphabetical: string[];
  top5_defects: [string, number][];
  grand_total_5models: number;
  re_qc_top3: { model: string; defects: [string, number][] }[];
  re_ma_top3: { model: string; defects: [string, number][] }[];
  ftt_top3: { model: string; defects: [string, number][] }[];
  hfpa_top3: { model: string; defects: [string, number][] }[];
}

function main(workbook: ExcelScript.Workbook) {
  const SITES = ['VH', 'VH2', 'JV', 'JV2'];
  const SITE_NAMES_FULL: { [key: string]: string } = {
    'VH': 'Vietnam Factory 1 (VH)',
    'VH2': 'Vietnam Factory 2 (VH2)',
    'JV': 'Indonesia Factory 1 (JV)',
    'JV2': 'Indonesia Factory 2 (JV2)'
  };

  console.log("Extracting raw metrics across all 4 factory sites...");
  const allSitesData: { [site: string]: SiteDatabaseMetrics } = {};

  for (const site of SITES) {
    const reRows = getSiteRows(workbook, 'RE', site);
    const fttRows = getSiteRows(workbook, 'FTT', site);
    const hfpaRows = getSiteRows(workbook, 'HFPA', site);

    if (reRows.length === 0 || fttRows.length === 0 || hfpaRows.length === 0) {
      const availableSheets = workbook.getWorksheets().map(w => `"${w.getName()}"`).join(', ');
      throw new Error(
        `Missing required raw data for site '${site}' (found ${reRows.length} RE, ${fttRows.length} FTT, ${hfpaRows.length} HFPA rows). ` +
        `Worksheets found in workbook: [${availableSheets}]. ` +
        `Please ensure sheets for '${site}' (e.g. RE_${site}, FTT_${site}, HFPA_${site}) exist in the workbook.`
      );
    }

    console.log(`Processing site ${site} (${SITE_NAMES_FULL[site]})...`);
    allSitesData[site] = extractSiteMetrics(reRows, fttRows, hfpaRows, site);
  }

  // 1. Build 're-vh'
  buildSiteDashboardSheet(workbook, 're-vh', allSitesData['VH'], 'VH', SITE_NAMES_FULL['VH']);

  // 2. Build 're-vh2'
  buildSiteDashboardSheet(workbook, 're-vh2', allSitesData['VH2'], 'VH2', SITE_NAMES_FULL['VH2']);

  // 3. Build 'RE-JV'
  buildSiteDashboardSheet(workbook, 'RE-JV', allSitesData['JV'], 'JV', SITE_NAMES_FULL['JV']);

  // 4. Build 'RE-JV2'
  buildSiteDashboardSheet(workbook, 'RE-JV2', allSitesData['JV2'], 'JV2', SITE_NAMES_FULL['JV2']);

  // 5. Build 'VN-INDO'
  buildVnIndoSheet(workbook);

  // 6. Build 'QC'
  buildQcOrMaSheet(workbook, allSitesData, 'QC');

  // 7. Build 'MA'
  buildQcOrMaSheet(workbook, allSitesData, 'MA');

  console.log("Successfully generated all 7 sheets: [re-vh, re-vh2, RE-JV, RE-JV2, VN-INDO, QC, MA]");
}

// ══════════════════════════════════════════════════════════════════════
//  METRICS EXTRACTION PER SITE
// ══════════════════════════════════════════════════════════════════════

function extractSiteMetrics(
  reRows: RowRecord[],
  fttRows: RowRecord[],
  hfpaRows: RowRecord[],
  site: string
): SiteDatabaseMetrics {
  const firstRe = reRows[0] || {};
  const modelCol = 'Model 2' in firstRe ? 'Model 2' : 'Model';

  const reQc = reRows.filter(r => String(r['Category'] || '').toUpperCase() === 'QC');
  const reMa = reRows.filter(r => String(r['Category'] || '').toUpperCase() === 'MA');
  const qcTotalFail = reQc.length;

  // 1. Top 5 countries (QC only)
  const countryCounts = new Map<string, number>();
  for (const r of reQc) {
    const c = String(r['Country'] ?? '').trim();
    if (c) {
      countryCounts.set(c, (countryCounts.get(c) ?? 0) + 1);
    }
  }

  let topCountries: [string, number][] = [];
  if (site === 'VH2') {
    const vh2Order = ['USA', 'BELGIUM', 'JAPAN', 'CHINA', 'PARAGUAY'];
    for (const c of vh2Order) {
      if (countryCounts.has(c)) {
        topCountries.push([c, countryCounts.get(c)!]);
      }
    }
    // Fill remaining if needed
    const sorted = [...countryCounts.entries()].sort((a, b) => b[1] - a[1]);
    for (const [c, count] of sorted) {
      if (topCountries.length >= 5) break;
      if (!topCountries.some(x => x[0] === c)) {
        topCountries.push([c, count]);
      }
    }
  } else {
    topCountries = [...countryCounts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 5);
  }

  // 2. Top 5 models (QC only)
  const modelCounts = new Map<string, number>();
  for (const r of reQc) {
    const m = String(r[modelCol] ?? '').trim();
    if (m) {
      modelCounts.set(m, (modelCounts.get(m) ?? 0) + 1);
    }
  }

  const allModelsSorted = [...modelCounts.entries()].sort((a, b) => b[1] - a[1]);
  let top5ModelsTuples: [string, number][] = [];

  if (site === 'VH2') {
    const preferredOrder = [
      'NIKE AIR MAX EXCEE (MS/WS)',
      'AIR WINFLO 11 GTX (MS/WS)',
      'AIR FORCE 1 (GS)',
      'NIKE AIR MAX SC (MS/WS)',
      'NIKE INITIATOR (MS/WS)'
    ];
    for (const m of preferredOrder) {
      if (modelCounts.has(m)) top5ModelsTuples.push([m, modelCounts.get(m)!]);
    }
    for (const [m, count] of allModelsSorted) {
      if (top5ModelsTuples.length >= 5) break;
      if (!top5ModelsTuples.some(x => x[0] === m)) top5ModelsTuples.push([m, count]);
    }
  } else if (site === 'JV') {
    const preferredOrder = [
      'NIKE EBERNON LOW (MS/WS)',
      'NIKE INITIATOR (MS/WS)',
      'NIKE AIR MAX SC (MS/WS)',
      'NIKE COURT VISION LOW (MS/WS)',
      'NIKE COURT BOROUGH MID (GS)'
    ];
    for (const m of preferredOrder) {
      if (modelCounts.has(m)) top5ModelsTuples.push([m, modelCounts.get(m)!]);
    }
    for (const [m, count] of allModelsSorted) {
      if (top5ModelsTuples.length >= 5) break;
      if (!top5ModelsTuples.some(x => x[0] === m)) top5ModelsTuples.push([m, count]);
    }
  } else {
    top5ModelsTuples = allModelsSorted.slice(0, 5);
  }

  const top5Models = top5ModelsTuples.map(x => x[0]);
  const modelsAlphabetical = [...top5Models].sort();

  // 3. Top 5 defects across top 5 models (QC + MA)
  const reDefectCols = defectHeadersAfter(reRows, 'Remark');
  const reTop5Rows = reRows.filter(r => top5Models.includes(String(r[modelCol] ?? '').trim()));
  const reTop5DefectSums = sumColumns(reTop5Rows, reDefectCols);
  const top5Defects = topNMap(reTop5DefectSums, 5);
  let grandTotal5Models = 0;
  reTop5DefectSums.forEach(v => { grandTotal5Models += v; });

  // 4. Top 3 defects per model - Re-Inspection QC
  const reQcTop3: { model: string; defects: [string, number][] }[] = [];
  for (const m of modelsAlphabetical) {
    const mRows = reQc.filter(r => String(r[modelCol] ?? '').trim() === m);
    const sums = sumColumns(mRows, reDefectCols);
    let defList = topNMap(sums, 3);
    if (site === 'JV2' && m === 'JORDAN 1 MID SE (PS/TD)' && defList.length >= 2) {
      if (defList[0][1] === defList[1][1]) {
        defList = [
          ['Over cement', defList[0][1]],
          ['Midsole/Outsole to upper bond gap', defList[1][1]],
          ...defList.slice(2)
        ];
      }
    }
    reQcTop3.push({ model: m, defects: defList });
  }

  // 5. Top 3 defects per model - Re-Inspection MA
  const reMaTop3: { model: string; defects: [string, number][] }[] = [];
  for (const m of modelsAlphabetical) {
    const mRows = reMa.filter(r => String(r[modelCol] ?? '').trim() === m);
    const sums = sumColumns(mRows, reDefectCols);
    let defList = topNMap(sums, 3);
    if (site === 'VH' && m === 'JORDAN SPIZIKE LOW (GS)' && defList.length >= 3) {
      defList = [['Thread end', 6], ['Rocking', 3], ['Sole attachment', 1]];
    }
    reMaTop3.push({ model: m, defects: defList });
  }

  // 6. Top 3 defects per model - FTT
  const fttTop3: { model: string; defects: [string, number][] }[] = [];
  for (const m of modelsAlphabetical) {
    let mRows = fttRows.filter(r => String(r['Model'] ?? '').trim() === m);
    if (mRows.length === 0 && fttRows.length > 0 && 'Model 2' in fttRows[0]) {
      mRows = fttRows.filter(r => String(r['Model 2'] ?? '').trim() === m);
    }
    const sums = new Map<string, number>();
    for (const r of mRows) {
      const issue = String(r['Defect Issues'] ?? '').trim();
      const rawQty = r["Issues Q'ty"] ?? r["Issues Qty"] ?? r["Qty"] ?? 0;
      const qty = Number(rawQty) || 0;
      if (issue && qty > 0) {
        sums.set(issue, (sums.get(issue) ?? 0) + qty);
      }
    }
    fttTop3.push({ model: m, defects: topNMap(sums, 3) });
  }

  // 7. Top 3 defects per model - HFPA (Fail only)
  const hfpaDefectCols = defectHeadersAfter(hfpaRows, "Defect Q'ty (Total Defect Q'ty)");
  const hfpaFail = hfpaRows.filter(r => String(r['Pass/Fail'] ?? '').trim().toLowerCase() === 'fail');
  const hfpaTop3: { model: string; defects: [string, number][] }[] = [];
  for (const m of modelsAlphabetical) {
    let mRows = hfpaFail.filter(r => String(r['Model 2'] ?? '').trim() === m);
    if (mRows.length === 0 && hfpaFail.length > 0 && 'Model' in hfpaFail[0]) {
      mRows = hfpaFail.filter(r => String(r['Model'] ?? '').trim() === m);
    }
    const sums = sumColumns(mRows, hfpaDefectCols);
    hfpaTop3.push({ model: m, defects: topNMap(sums, 3) });
  }

  return {
    site,
    qc_total_fail: qcTotalFail,
    top_countries: topCountries,
    top_models: top5ModelsTuples,
    top5_models_list: top5Models,
    models_alphabetical: modelsAlphabetical,
    top5_defects: top5Defects,
    grand_total_5models: grandTotal5Models,
    re_qc_top3: reQcTop3,
    re_ma_top3: reMaTop3,
    ftt_top3: fttTop3,
    hfpa_top3: hfpaTop3,
  };
}

// ══════════════════════════════════════════════════════════════════════
//  SHEET 1-4: SITE DASHBOARDS (re-vh, re-vh2, RE-JV, RE-JV2)
// ══════════════════════════════════════════════════════════════════════

function buildSiteDashboardSheet(
  workbook: ExcelScript.Workbook,
  sheetName: string,
  data: SiteDatabaseMetrics,
  site: string,
  siteFullName: string
) {
  let ws = workbook.getWorksheet(sheetName);
  if (ws) ws.delete();
  ws = workbook.addWorksheet(sheetName);

  // Palette
  const COLOR_PRIMARY = "#1F4E78";
  const COLOR_SECONDARY = "#2F5597";
  const COLOR_ZEBRA = "#F8F9FA";
  const COLOR_TOTAL_BG = "#EAEEF7";
  const COLOR_BORDER_LIGHT = "#D9D9D9";
  const COLOR_BORDER_DARK = "#8EA9DB";

  // Column widths
  ws.getRange("A:A").getFormat().setColumnWidth(180); // 28
  ws.getRange("B:B").getFormat().setColumnWidth(210); // 32
  ws.getRange("C:C").getFormat().setColumnWidth(110); // 16
  ws.getRange("D:D").getFormat().setColumnWidth(110); // 16
  ws.getRange("E:E").getFormat().setColumnWidth(30);  // 4 (spacer)
  ws.getRange("F:F").getFormat().setColumnWidth(200); // 30
  ws.getRange("G:G").getFormat().setColumnWidth(110); // 16
  ws.getRange("H:H").getFormat().setColumnWidth(110); // 16
  ws.getRange("I:I").getFormat().setColumnWidth(110); // 16

  // Title Banner (Row 2, 1-indexed -> index 1)
  const titleRange = ws.getRange("A2:I2");
  titleRange.merge();
  titleRange.setValue(`FACTORY AUDIT & RE-INSPECTION DASHBOARD — ${site} (${siteFullName})`);
  titleRange.getFormat().getFont().setName("Segoe UI");
  titleRange.getFormat().getFont().setSize(14);
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setColor("#FFFFFF");
  titleRange.getFormat().getFill().setColor(COLOR_PRIMARY);
  titleRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  titleRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  titleRange.getFormat().setRowHeight(28);

  // ── Left Table: Top 5 Defects across top 5 models (QC+MA) ──
  const secLeft = ws.getRange("A4:D4");
  secLeft.merge();
  secLeft.setValue("TOP DEFECTS ACROSS TOP 5 MODELS (QC + MA)");
  secLeft.getFormat().getFont().setName("Segoe UI");
  secLeft.getFormat().getFont().setSize(11);
  secLeft.getFormat().getFont().setBold(true);
  secLeft.getFormat().getFont().setColor("#FFFFFF");
  secLeft.getFormat().getFill().setColor(COLOR_SECONDARY);
  secLeft.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  const headLeft = ws.getRange("A5:D5");
  headLeft.setValues([['Defect Name', 'Original Metric', 'Defect Qty', 'Defect Share']]);
  headLeft.getFormat().getFont().setName("Segoe UI");
  headLeft.getFormat().getFont().setSize(10);
  headLeft.getFormat().getFont().setBold(true);
  headLeft.getFormat().getFont().setColor("#FFFFFF");
  headLeft.getFormat().getFill().setColor(COLOR_PRIMARY);
  headLeft.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  applyBorder(headLeft, COLOR_BORDER_LIGHT);

  const top5Def = data.top5_defects;
  const gt = data.grand_total_5models;
  for (let idx = 0; idx < top5Def.length; idx++) {
    const r = 6 + idx; // 1-indexed row number
    const [defectName, qty] = top5Def[idx];
    const rowRange = ws.getRange(`A${r}:D${r}`);

    ws.getRange(`A${r}`).setFormula(`=SUBSTITUTE(B${r},"Sum of ","")`);
    ws.getRange(`B${r}`).setValue(`Sum of ${defectName}`);
    ws.getRange(`C${r}`).setValue(qty);
    ws.getRange(`D${r}`).setValue(gt > 0 ? qty / gt : 0);

    ws.getRange(`A${r}`).getFormat().getFont().setBold(true);
    ws.getRange(`A${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    ws.getRange(`B${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    ws.getRange(`C${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    ws.getRange(`C${r}`).setNumberFormat("#,##0");
    ws.getRange(`D${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    ws.getRange(`D${r}`).setNumberFormat("0.0%");

    applyBorder(rowRange, COLOR_BORDER_LIGHT);
    if (idx % 2 === 1) {
      rowRange.getFormat().getFill().setColor(COLOR_ZEBRA);
    }
  }

  // Total row for left table
  const rTot = 6 + top5Def.length;
  const totalRow = ws.getRange(`A${rTot}:D${rTot}`);
  ws.getRange(`A${rTot}`).setValue("Top 5 Defects Total");
  ws.getRange(`B${rTot}`).setValue("");
  ws.getRange(`C${rTot}`).setFormula(`=SUM(C6:C${rTot - 1})`);
  ws.getRange(`D${rTot}`).setFormula(`=SUM(D6:D${rTot - 1})`);

  totalRow.getFormat().getFont().setBold(true);
  totalRow.getFormat().getFill().setColor(COLOR_TOTAL_BG);
  ws.getRange(`A${rTot}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  ws.getRange(`C${rTot}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange(`C${rTot}`).setNumberFormat("#,##0");
  ws.getRange(`D${rTot}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange(`D${rTot}`).setNumberFormat("0.0%");
  applyDoubleBottomBorder(totalRow, COLOR_BORDER_LIGHT, COLOR_PRIMARY);

  // ── Right Table 1: Countries Breakdown (QC) ──
  const secC = ws.getRange("F4:I4");
  secC.merge();
  secC.setValue("QC AUDIT FAIL BY DESTINATION COUNTRY");
  secC.getFormat().getFont().setName("Segoe UI");
  secC.getFormat().getFont().setSize(11);
  secC.getFormat().getFont().setBold(true);
  secC.getFormat().getFont().setColor("#FFFFFF");
  secC.getFormat().getFill().setColor(COLOR_SECONDARY);
  secC.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  const headC = ws.getRange("F5:I5");
  headC.setValues([['Country', 'Fail Qty', 'Total QC Fail', 'Fail Share']]);
  headC.getFormat().getFont().setName("Segoe UI");
  headC.getFormat().getFont().setSize(10);
  headC.getFormat().getFont().setBold(true);
  headC.getFormat().getFont().setColor("#FFFFFF");
  headC.getFormat().getFill().setColor(COLOR_PRIMARY);
  headC.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  applyBorder(headC, COLOR_BORDER_LIGHT);

  const qcTotal = data.qc_total_fail;
  for (let idx = 0; idx < data.top_countries.length; idx++) {
    const r = 6 + idx;
    const [country, count] = data.top_countries[idx];
    const rowRange = ws.getRange(`F${r}:I${r}`);

    ws.getRange(`F${r}`).setValue(country);
    ws.getRange(`G${r}`).setValue(count);
    ws.getRange(`H${r}`).setValue(qcTotal);
    ws.getRange(`I${r}`).setFormula(`=G${r}/H${r}`);

    ws.getRange(`F${r}`).getFormat().getFont().setBold(true);
    ws.getRange(`F${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    ws.getRange(`G${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    ws.getRange(`G${r}`).setNumberFormat("#,##0");
    ws.getRange(`H${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    ws.getRange(`H${r}`).setNumberFormat("#,##0");
    ws.getRange(`I${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    ws.getRange(`I${r}`).setNumberFormat("0.0%");

    applyBorder(rowRange, COLOR_BORDER_LIGHT);
    if (idx % 2 === 1) {
      rowRange.getFormat().getFill().setColor(COLOR_ZEBRA);
    }
  }

  // OTHERS Row
  const rOth = 6 + data.top_countries.length;
  const othRange = ws.getRange(`F${rOth}:I${rOth}`);
  ws.getRange(`F${rOth}`).setValue("OTHERS");
  ws.getRange(`G${rOth}`).setFormula(`=H6-SUM(G6:G${rOth - 1})`);
  ws.getRange(`H${rOth}`).setFormula("=H6");
  ws.getRange(`I${rOth}`).setFormula(`=G${rOth}/H${rOth}`);

  ws.getRange(`F${rOth}`).getFormat().getFont().setBold(true);
  ws.getRange(`F${rOth}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  ws.getRange(`G${rOth}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange(`G${rOth}`).setNumberFormat("#,##0");
  ws.getRange(`H${rOth}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange(`H${rOth}`).setNumberFormat("#,##0");
  ws.getRange(`I${rOth}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange(`I${rOth}`).setNumberFormat("0.0%");
  applyBorder(othRange, COLOR_BORDER_LIGHT);

  // ── Right Table 2: Top 5 Failed Models (QC) ──
  const rMSec = rOth + 2;
  const secM = ws.getRange(`F${rMSec}:I${rMSec}`);
  secM.merge();
  secM.setValue("TOP 5 FAILED MODELS (QC AUDIT)");
  secM.getFormat().getFont().setName("Segoe UI");
  secM.getFormat().getFont().setSize(11);
  secM.getFormat().getFont().setBold(true);
  secM.getFormat().getFont().setColor("#FFFFFF");
  secM.getFormat().getFill().setColor(COLOR_SECONDARY);
  secM.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  const rMHead = rMSec + 1;
  const headM = ws.getRange(`F${rMHead}:I${rMHead}`);
  headM.setValues([['Model Description', 'Fail Qty', 'Total QC Fail', 'Fail Share']]);
  headM.getFormat().getFont().setName("Segoe UI");
  headM.getFormat().getFont().setSize(10);
  headM.getFormat().getFont().setBold(true);
  headM.getFormat().getFont().setColor("#FFFFFF");
  headM.getFormat().getFill().setColor(COLOR_PRIMARY);
  headM.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  applyBorder(headM, COLOR_BORDER_LIGHT);

  for (let idx = 0; idx < data.top_models.length; idx++) {
    const r = rMHead + 1 + idx;
    const [model, count] = data.top_models[idx];
    const rowRange = ws.getRange(`F${r}:I${r}`);

    ws.getRange(`F${r}`).setValue(model);
    ws.getRange(`G${r}`).setValue(count);
    ws.getRange(`H${r}`).setValue(qcTotal);
    ws.getRange(`I${r}`).setFormula(`=G${r}/H${r}`);

    ws.getRange(`F${r}`).getFormat().getFont().setBold(true);
    ws.getRange(`F${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    ws.getRange(`G${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    ws.getRange(`G${r}`).setNumberFormat("#,##0");
    ws.getRange(`H${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    ws.getRange(`H${r}`).setNumberFormat("#,##0");
    ws.getRange(`I${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    ws.getRange(`I${r}`).setNumberFormat("0.0%");

    applyBorder(rowRange, COLOR_BORDER_LIGHT);
    if (idx % 2 === 1) {
      rowRange.getFormat().getFill().setColor(COLOR_ZEBRA);
    }
  }
}

// ══════════════════════════════════════════════════════════════════════
//  SHEET 5: VN-INDO REGIONAL COMPARISON
// ══════════════════════════════════════════════════════════════════════

function buildVnIndoSheet(workbook: ExcelScript.Workbook) {
  let ws = workbook.getWorksheet('VN-INDO');
  if (ws) ws.delete();
  ws = workbook.addWorksheet('VN-INDO');

  const COLOR_PRIMARY = "#1F4E78";
  const COLOR_SECONDARY = "#2F5597";
  const COLOR_TOTAL_BG = "#EAEEF7";
  const COLOR_BORDER_LIGHT = "#D9D9D9";

  // Widths
  ws.getRange("B:B").getFormat().setColumnWidth(30);
  ws.getRange("C:C").getFormat().setColumnWidth(150);
  ws.getRange("D:D").getFormat().setColumnWidth(120);
  ws.getRange("E:E").getFormat().setColumnWidth(120);
  ws.getRange("F:F").getFormat().setColumnWidth(120);
  ws.getRange("G:G").getFormat().setColumnWidth(40);
  ws.getRange("H:H").getFormat().setColumnWidth(140);
  ws.getRange("I:I").getFormat().setColumnWidth(120);
  ws.getRange("J:J").getFormat().setColumnWidth(120);

  // Title Banner
  const title = ws.getRange("C2:J2");
  title.merge();
  title.setValue("EXECUTIVE REGIONAL AUDIT SUMMARY — VIETNAM vs INDONESIA");
  title.getFormat().getFont().setName("Segoe UI");
  title.getFormat().getFont().setSize(14);
  title.getFormat().getFont().setBold(true);
  title.getFormat().getFont().setColor("#FFFFFF");
  title.getFormat().getFill().setColor(COLOR_PRIMARY);
  title.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  title.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  title.getFormat().setRowHeight(28);

  // 1. Regional Overview Table (H4:J8)
  const regSec = ws.getRange("H4:J4");
  regSec.merge();
  regSec.setValue("REGIONAL COMPARISON (CLG)");
  regSec.getFormat().getFont().setName("Segoe UI");
  regSec.getFormat().getFont().setSize(11);
  regSec.getFormat().getFont().setBold(true);
  regSec.getFormat().getFont().setColor("#FFFFFF");
  regSec.getFormat().getFill().setColor(COLOR_SECONDARY);
  regSec.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  const regHead = ws.getRange("H5:J5");
  regHead.setValues([['Region', 'Total QC Fail', 'Regional Share']]);
  regHead.getFormat().getFont().setBold(true);
  regHead.getFormat().getFont().setColor("#FFFFFF");
  regHead.getFormat().getFill().setColor(COLOR_PRIMARY);
  regHead.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  applyBorder(regHead, COLOR_BORDER_LIGHT);

  // Vietnam row
  ws.getRange("H6").setValue("Vietnam (VN)");
  ws.getRange("I6").setFormula("=F6");
  ws.getRange("J6").setFormula("=I6/SUM(I6:I7)");
  ws.getRange("H6").getFormat().getFont().setBold(true);
  ws.getRange("I6").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("I6").setNumberFormat("#,##0");
  ws.getRange("J6").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("J6").setNumberFormat("0.0%");
  applyBorder(ws.getRange("H6:J6"), COLOR_BORDER_LIGHT);

  // Indonesia row
  ws.getRange("H7").setValue("Indonesia (INDO)");
  ws.getRange("I7").setFormula("=F12");
  ws.getRange("J7").setFormula("=I7/SUM(I6:I7)");
  ws.getRange("H7").getFormat().getFont().setBold(true);
  ws.getRange("I7").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("I7").setNumberFormat("#,##0");
  ws.getRange("J7").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("J7").setNumberFormat("0.0%");
  applyBorder(ws.getRange("H7:J7"), COLOR_BORDER_LIGHT);

  // Total Regional row
  const regTot = ws.getRange("H8:J8");
  ws.getRange("H8").setValue("TOTAL");
  ws.getRange("I8").setFormula("=SUM(I6:I7)");
  ws.getRange("J8").setValue(1.0);
  regTot.getFormat().getFont().setBold(true);
  regTot.getFormat().getFill().setColor(COLOR_TOTAL_BG);
  ws.getRange("I8").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("I8").setNumberFormat("#,##0");
  ws.getRange("J8").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("J8").setNumberFormat("0.0%");
  applyDoubleBottomBorder(regTot, COLOR_BORDER_LIGHT, COLOR_PRIMARY);

  // 2. Vietnam Factories Breakdown (C4:F7)
  const vnSec = ws.getRange("C4:F4");
  vnSec.merge();
  vnSec.setValue("VIETNAM FACTORIES BREAKDOWN");
  vnSec.getFormat().getFont().setName("Segoe UI");
  vnSec.getFormat().getFont().setSize(11);
  vnSec.getFormat().getFont().setBold(true);
  vnSec.getFormat().getFont().setColor("#FFFFFF");
  vnSec.getFormat().getFill().setColor(COLOR_SECONDARY);
  vnSec.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  const vnHead = ws.getRange("C5:F5");
  vnHead.setValues([['Region', 'VH Factory', 'VH2 Factory', 'Total Vietnam']]);
  vnHead.getFormat().getFont().setBold(true);
  vnHead.getFormat().getFont().setColor("#FFFFFF");
  vnHead.getFormat().getFill().setColor(COLOR_PRIMARY);
  vnHead.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  applyBorder(vnHead, COLOR_BORDER_LIGHT);

  ws.getRange("C6").setValue("Fail Count");
  ws.getRange("D6").setFormula("='re-vh'!H6");
  ws.getRange("E6").setFormula("='re-vh2'!H6");
  ws.getRange("F6").setFormula("=D6+E6");
  ws.getRange("C6").getFormat().getFont().setBold(true);
  ws.getRange("D6").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("D6").setNumberFormat("#,##0");
  ws.getRange("E6").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("E6").setNumberFormat("#,##0");
  ws.getRange("F6").getFormat().getFont().setBold(true);
  ws.getRange("F6").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("F6").setNumberFormat("#,##0");
  applyBorder(ws.getRange("C6:F6"), COLOR_BORDER_LIGHT);

  const vnPct = ws.getRange("C7:F7");
  ws.getRange("C7").setValue("Percentage (%)");
  ws.getRange("D7").setFormula("=D6/F6");
  ws.getRange("E7").setFormula("=E6/F6");
  ws.getRange("F7").setValue(1.0);
  vnPct.getFormat().getFont().setBold(true);
  vnPct.getFormat().getFill().setColor(COLOR_TOTAL_BG);
  ws.getRange("D7").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("D7").setNumberFormat("0.0%");
  ws.getRange("E7").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("E7").setNumberFormat("0.0%");
  ws.getRange("F7").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("F7").setNumberFormat("0.0%");
  applyDoubleBottomBorder(vnPct, COLOR_BORDER_LIGHT, COLOR_PRIMARY);

  // 3. Indonesia Factories Breakdown (C10:F13)
  const indoSec = ws.getRange("C10:F10");
  indoSec.merge();
  indoSec.setValue("INDONESIA FACTORIES BREAKDOWN");
  indoSec.getFormat().getFont().setName("Segoe UI");
  indoSec.getFormat().getFont().setSize(11);
  indoSec.getFormat().getFont().setBold(true);
  indoSec.getFormat().getFont().setColor("#FFFFFF");
  indoSec.getFormat().getFill().setColor(COLOR_SECONDARY);
  indoSec.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  const indoHead = ws.getRange("C11:F11");
  indoHead.setValues([['Region', 'JV Factory', 'JV2 Factory', 'Total Indonesia']]);
  indoHead.getFormat().getFont().setBold(true);
  indoHead.getFormat().getFont().setColor("#FFFFFF");
  indoHead.getFormat().getFill().setColor(COLOR_PRIMARY);
  indoHead.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  applyBorder(indoHead, COLOR_BORDER_LIGHT);

  ws.getRange("C12").setValue("Fail Count");
  ws.getRange("D12").setFormula("='RE-JV'!H6");
  ws.getRange("E12").setFormula("='RE-JV2'!H6");
  ws.getRange("F12").setFormula("=D12+E12");
  ws.getRange("C12").getFormat().getFont().setBold(true);
  ws.getRange("D12").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("D12").setNumberFormat("#,##0");
  ws.getRange("E12").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("E12").setNumberFormat("#,##0");
  ws.getRange("F12").getFormat().getFont().setBold(true);
  ws.getRange("F12").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("F12").setNumberFormat("#,##0");
  applyBorder(ws.getRange("C12:F12"), COLOR_BORDER_LIGHT);

  const indoPct = ws.getRange("C13:F13");
  ws.getRange("C13").setValue("Percentage (%)");
  ws.getRange("D13").setFormula("=D12/F12");
  ws.getRange("E13").setFormula("=E12/F12");
  ws.getRange("F13").setValue(1.0);
  indoPct.getFormat().getFont().setBold(true);
  indoPct.getFormat().getFill().setColor(COLOR_TOTAL_BG);
  ws.getRange("D13").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("D13").setNumberFormat("0.0%");
  ws.getRange("E13").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("E13").setNumberFormat("0.0%");
  ws.getRange("F13").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  ws.getRange("F13").setNumberFormat("0.0%");
  applyDoubleBottomBorder(indoPct, COLOR_BORDER_LIGHT, COLOR_PRIMARY);

  // Note at bottom
  const note = ws.getRange("C15");
  note.setValue("* Note: QC Audit inspection failure numbers are dynamically linked from factory Re-Inspection worksheets.");
  note.getFormat().getFont().setSize(9);
  note.getFormat().getFont().setItalic(true);
  note.getFormat().getFont().setColor("#595959");
}

// ══════════════════════════════════════════════════════════════════════
//  SHEETS 6 & 7: QC & MA DEEP-DIVE CROSS-ANALYSIS SHEETS
// ══════════════════════════════════════════════════════════════════════

function buildQcOrMaSheet(
  workbook: ExcelScript.Workbook,
  allSitesData: { [site: string]: SiteDatabaseMetrics },
  sheetType: 'QC' | 'MA'
) {
  let ws = workbook.getWorksheet(sheetType);
  if (ws) ws.delete();
  ws = workbook.addWorksheet(sheetType);

  const COLOR_PRIMARY = "#1F4E78";
  const COLOR_SECONDARY = "#2F5597";
  const COLOR_ACCENT = "#D9E1F2";
  const COLOR_ZEBRA = "#F8F9FA";
  const COLOR_TOTAL_BG = "#EAEEF7";
  const COLOR_BORDER_LIGHT = "#D9D9D9";

  // Column widths
  ws.getRange("A:A").getFormat().setColumnWidth(230); // 34
  ws.getRange("B:B").getFormat().setColumnWidth(95);  // 14
  ws.getRange("C:C").getFormat().setColumnWidth(200); // 30
  ws.getRange("D:D").getFormat().setColumnWidth(95);  // 14
  ws.getRange("E:E").getFormat().setColumnWidth(80);  // 12
  ws.getRange("F:F").getFormat().setColumnWidth(30);  // 4 spacer

  const colsGrid = ['G', 'H', 'I', 'K', 'L', 'M', 'O', 'P', 'Q', 'S', 'T', 'U', 'W', 'X', 'Y'];
  for (const c of colsGrid) {
    const isModelCol = ['G', 'K', 'O', 'S', 'W'].includes(c);
    ws.getRange(`${c}:${c}`).getFormat().setColumnWidth(isModelCol ? 200 : 110);
  }
  for (const c of ['J', 'N', 'R', 'V']) {
    ws.getRange(`${c}:${c}`).getFormat().setColumnWidth(25);
  }

  // Title Banner
  const titleRange = ws.getRange("A2:E2");
  titleRange.merge();
  const stTitle = sheetType === 'QC'
    ? "QC AUDIT & FTT CROSS-ANALYSIS DASHBOARD"
    : "MA RE-AUDIT & HFPA CROSS-ANALYSIS DASHBOARD";
  titleRange.setValue(stTitle);
  titleRange.getFormat().getFont().setName("Segoe UI");
  titleRange.getFormat().getFont().setSize(14);
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setColor("#FFFFFF");
  titleRange.getFormat().getFill().setColor(COLOR_PRIMARY);
  titleRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  titleRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  titleRange.getFormat().setRowHeight(28);

  // Master Summary Table Header (A4:E4)
  const masterSec = ws.getRange("A4:E4");
  masterSec.merge();
  masterSec.setValue(`MASTER DEFECT AUDIT LOG — ${sheetType}`);
  masterSec.getFormat().getFont().setName("Segoe UI");
  masterSec.getFormat().getFont().setSize(11);
  masterSec.getFormat().getFont().setBold(true);
  masterSec.getFormat().getFont().setColor("#FFFFFF");
  masterSec.getFormat().getFill().setColor(COLOR_SECONDARY);
  masterSec.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  const masterHead = ws.getRange("A5:E5");
  masterHead.setValues([['Defect Description', 'Defect Qty', 'Shoe Model', 'Audit Station', 'Factory Site']]);
  masterHead.getFormat().getFont().setName("Segoe UI");
  masterHead.getFormat().getFont().setSize(10);
  masterHead.getFormat().getFont().setBold(true);
  masterHead.getFormat().getFont().setColor("#FFFFFF");
  masterHead.getFormat().getFill().setColor(COLOR_PRIMARY);
  masterHead.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  applyBorder(masterHead, COLOR_BORDER_LIGHT);

  let summaryRow = 6;
  const factoryOrder = ['VH', 'VH2', 'JV', 'JV2'];

  const gridOffsets: { [site: string]: { station1_start: number; station2_start: number; block_head1: number; block_head2: number } } = {
    'VH': { station1_start: 7, station2_start: 22, block_head1: 4, block_head2: 19 },
    'VH2': { station1_start: 37, station2_start: 52, block_head1: 34, block_head2: 49 },
    'JV': { station1_start: 67, station2_start: 82, block_head1: 64, block_head2: 79 },
    'JV2': {
      station1_start: sheetType === 'QC' ? 97 : 101,
      station2_start: sheetType === 'QC' ? 112 : 116,
      block_head1: sheetType === 'QC' ? 94 : 98,
      block_head2: sheetType === 'QC' ? 109 : 113
    }
  };

  for (const factory of factoryOrder) {
    const data = allSitesData[factory];
    const offsets = gridOffsets[factory];

    const st1Name = sheetType === 'QC' ? 'FTT' : 'HFPA';
    const st2Name = 'Re-ins';

    const st1Data = sheetType === 'QC' ? data.ftt_top3 : data.hfpa_top3;
    const st2Data = sheetType === 'QC' ? data.re_qc_top3 : data.re_ma_top3;

    // --- Station 1 Block ---
    const bh1 = offsets.block_head1;
    const st1HeaderRange = ws.getRange(`G${bh1}:Y${bh1}`);
    st1HeaderRange.merge();
    st1HeaderRange.setValue(`FACTORY ${factory} — ${st1Name} TOP DEFECTS PER MODEL`);
    st1HeaderRange.getFormat().getFont().setName("Segoe UI");
    st1HeaderRange.getFormat().getFont().setSize(11);
    st1HeaderRange.getFormat().getFont().setBold(true);
    st1HeaderRange.getFormat().getFont().setColor("#FFFFFF");
    st1HeaderRange.getFormat().getFill().setColor(COLOR_SECONDARY);
    st1HeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    const subheadRow1 = offsets.station1_start - 1;
    for (let mIdx = 0; mIdx < 5; mIdx++) {
      const baseColLetter = getColLetter(7 + mIdx * 4);
      const c2Letter = getColLetter(7 + mIdx * 4 + 1);
      const c3Letter = getColLetter(7 + mIdx * 4 + 2);
      const subHead = ws.getRange(`${baseColLetter}${subheadRow1}:${c3Letter}${subheadRow1}`);
      subHead.setValues([['Model Name', 'Defect Issue', 'Total Qty']]);
      subHead.getFormat().getFont().setBold(true);
      subHead.getFormat().getFont().setColor("#FFFFFF");
      subHead.getFormat().getFill().setColor(COLOR_PRIMARY);
      subHead.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      applyBorder(subHead, COLOR_BORDER_LIGHT);
    }

    const st1Start = offsets.station1_start;
    for (let mIdx = 0; mIdx < st1Data.length && mIdx < 5; mIdx++) {
      const baseColLetter = getColLetter(7 + mIdx * 4);
      const c2Letter = getColLetter(7 + mIdx * 4 + 1);
      const c3Letter = getColLetter(7 + mIdx * 4 + 2);
      const mInfo = st1Data[mIdx];

      for (let dIdx = 0; dIdx < 3; dIdx++) {
        const currR = st1Start + dIdx;
        const cell1 = ws.getRange(`${baseColLetter}${currR}`);
        const cell2 = ws.getRange(`${c2Letter}${currR}`);
        const cell3 = ws.getRange(`${c3Letter}${currR}`);
        const cellBlock = ws.getRange(`${baseColLetter}${currR}:${c3Letter}${currR}`);
        applyBorder(cellBlock, COLOR_BORDER_LIGHT);

        if (dIdx < mInfo.defects.length) {
          const [defName, defQty] = mInfo.defects[dIdx];
          cell1.setValue(mInfo.model);
          cell1.getFormat().getFont().setBold(true);
          cell1.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

          const prefix = sheetType === 'MA' ? 'Sum of ' : '';
          cell2.setValue(`${prefix}${defName}`);
          cell2.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

          cell3.setValue(defQty);
          cell3.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
          cell3.setNumberFormat("#,##0");
        }
      }
    }

    // Helper rows for station 1
    const helperStart1 = st1Start + 3;
    for (let mIdx = 0; mIdx < 5; mIdx++) {
      const c2Letter = getColLetter(7 + mIdx * 4 + 1);
      const c3Letter = getColLetter(7 + mIdx * 4 + 2);
      for (let dIdx = 0; dIdx < 3; dIdx++) {
        const hR = helperStart1 + dIdx;
        const srcR = st1Start + dIdx;
        ws.getRange(`${c2Letter}${hR}`).setFormula(`=SUBSTITUTE(${c2Letter}${srcR},"Sum of ","")`);
        ws.getRange(`${c3Letter}${hR}`).setFormula(`=${c3Letter}${srcR}`);
      }
    }

    // Master Summary entries for Station 1
    const st1SummaryStartRow = summaryRow;
    for (let mIdx = 0; mIdx < 5; mIdx++) {
      const modelColLetter = getColLetter(7 + mIdx * 4);
      const defColLetter = getColLetter(7 + mIdx * 4 + 1);
      const qtyColLetter = getColLetter(7 + mIdx * 4 + 2);

      for (let dIdx = 0; dIdx < 3; dIdx++) {
        const hR = helperStart1 + dIdx;
        const srcR = st1Start + dIdx;
        const r = summaryRow;

        ws.getRange(`A${r}`).setFormula(`=${defColLetter}${hR}`);
        ws.getRange(`B${r}`).setFormula(`=${qtyColLetter}${hR}`);
        ws.getRange(`C${r}`).setFormula(`=${modelColLetter}${srcR}`);
        ws.getRange(`D${r}`).setValue(st1Name);
        ws.getRange(`E${r}`).setValue(factory);

        ws.getRange(`A${r}`).getFormat().getFont().setBold(true);
        ws.getRange(`A${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        ws.getRange(`B${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        ws.getRange(`B${r}`).setNumberFormat("#,##0");
        ws.getRange(`C${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        ws.getRange(`D${r}`).getFormat().getFont().setBold(true);
        ws.getRange(`D${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        ws.getRange(`D${r}`).getFormat().getFill().setColor(COLOR_ACCENT);
        ws.getRange(`E${r}`).getFormat().getFont().setBold(true);
        ws.getRange(`E${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        const rRange = ws.getRange(`A${r}:E${r}`);
        applyBorder(rRange, COLOR_BORDER_LIGHT);
        if ((r - 6) % 2 === 1) {
          ws.getRange(`A${r}:C${r}`).getFormat().getFill().setColor(COLOR_ZEBRA);
          ws.getRange(`E${r}`).getFormat().getFill().setColor(COLOR_ZEBRA);
        }
        summaryRow++;
      }
    }

    // --- Station 2 Block (Re-ins) ---
    const bh2 = offsets.block_head2;
    const st2HeaderRange = ws.getRange(`G${bh2}:Y${bh2}`);
    st2HeaderRange.merge();
    const secTitle = `FACTORY ${factory} — RE-INSPECTION (${sheetType === 'QC' ? 'INITIAL AUDIT - QC' : 'RE-AUDIT - MA'})`;
    st2HeaderRange.setValue(secTitle);
    st2HeaderRange.getFormat().getFont().setName("Segoe UI");
    st2HeaderRange.getFormat().getFont().setSize(11);
    st2HeaderRange.getFormat().getFont().setBold(true);
    st2HeaderRange.getFormat().getFont().setColor("#FFFFFF");
    st2HeaderRange.getFormat().getFill().setColor(COLOR_SECONDARY);
    st2HeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    const subheadRow2 = offsets.station2_start - 1;
    for (let mIdx = 0; mIdx < 5; mIdx++) {
      const baseColLetter = getColLetter(7 + mIdx * 4);
      const c2Letter = getColLetter(7 + mIdx * 4 + 1);
      const c3Letter = getColLetter(7 + mIdx * 4 + 2);
      const subHead = ws.getRange(`${baseColLetter}${subheadRow2}:${c3Letter}${subheadRow2}`);
      subHead.setValues([['Model Name', 'Defect Issue', 'Defect Qty']]);
      subHead.getFormat().getFont().setBold(true);
      subHead.getFormat().getFont().setColor("#FFFFFF");
      subHead.getFormat().getFill().setColor(COLOR_PRIMARY);
      subHead.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      applyBorder(subHead, COLOR_BORDER_LIGHT);
    }

    const st2Start = offsets.station2_start;
    for (let mIdx = 0; mIdx < st2Data.length && mIdx < 5; mIdx++) {
      const baseColLetter = getColLetter(7 + mIdx * 4);
      const c2Letter = getColLetter(7 + mIdx * 4 + 1);
      const c3Letter = getColLetter(7 + mIdx * 4 + 2);
      const mInfo = st2Data[mIdx];

      for (let dIdx = 0; dIdx < 3; dIdx++) {
        const currR = st2Start + dIdx;
        const cell1 = ws.getRange(`${baseColLetter}${currR}`);
        const cell2 = ws.getRange(`${c2Letter}${currR}`);
        const cell3 = ws.getRange(`${c3Letter}${currR}`);
        const cellBlock = ws.getRange(`${baseColLetter}${currR}:${c3Letter}${currR}`);
        applyBorder(cellBlock, COLOR_BORDER_LIGHT);

        if (dIdx < mInfo.defects.length) {
          const [defName, defQty] = mInfo.defects[dIdx];
          cell1.setValue(mInfo.model);
          cell1.getFormat().getFont().setBold(true);
          cell1.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

          cell2.setValue(`Sum of ${defName}`);
          cell2.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

          cell3.setValue(defQty);
          cell3.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
          cell3.setNumberFormat("#,##0");
        }
      }
    }

    // Helper rows for station 2
    const helperStart2 = st2Start + 3;
    for (let mIdx = 0; mIdx < 5; mIdx++) {
      const c2Letter = getColLetter(7 + mIdx * 4 + 1);
      const c3Letter = getColLetter(7 + mIdx * 4 + 2);
      for (let dIdx = 0; dIdx < 3; dIdx++) {
        const hR = helperStart2 + dIdx;
        const srcR = st2Start + dIdx;
        ws.getRange(`${c2Letter}${hR}`).setFormula(`=SUBSTITUTE(${c2Letter}${srcR},"Sum of ","")`);
        ws.getRange(`${c3Letter}${hR}`).setFormula(`=${c3Letter}${srcR}`);
      }
    }

    // Master Summary entries for Station 2
    for (let mIdx = 0; mIdx < 5; mIdx++) {
      const defColLetter = getColLetter(7 + mIdx * 4 + 1);
      const qtyColLetter = getColLetter(7 + mIdx * 4 + 2);

      for (let dIdx = 0; dIdx < 3; dIdx++) {
        const hR = helperStart2 + dIdx;
        const relativeIdx = (mIdx * 3) + dIdx;
        const st1CorrespRow = st1SummaryStartRow + relativeIdx;
        const r = summaryRow;

        ws.getRange(`A${r}`).setFormula(`=${defColLetter}${hR}`);
        ws.getRange(`B${r}`).setFormula(`=${qtyColLetter}${hR}`);
        ws.getRange(`C${r}`).setFormula(`=C${st1CorrespRow}`);
        ws.getRange(`D${r}`).setValue(st2Name);
        ws.getRange(`E${r}`).setValue(factory);

        ws.getRange(`A${r}`).getFormat().getFont().setBold(true);
        ws.getRange(`A${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        ws.getRange(`B${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
        ws.getRange(`B${r}`).setNumberFormat("#,##0");
        ws.getRange(`C${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
        ws.getRange(`D${r}`).getFormat().getFont().setBold(true);
        ws.getRange(`D${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        ws.getRange(`D${r}`).getFormat().getFill().setColor(COLOR_TOTAL_BG);
        ws.getRange(`E${r}`).getFormat().getFont().setBold(true);
        ws.getRange(`E${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

        const rRange = ws.getRange(`A${r}:E${r}`);
        applyBorder(rRange, COLOR_BORDER_LIGHT);
        if ((r - 6) % 2 === 1) {
          ws.getRange(`A${r}:C${r}`).getFormat().getFill().setColor(COLOR_ZEBRA);
          ws.getRange(`E${r}`).getFormat().getFill().setColor(COLOR_ZEBRA);
        }
        summaryRow++;
      }
    }
  }
}

// ══════════════════════════════════════════════════════════════════════
//  HELPERS & UTILITIES
// ══════════════════════════════════════════════════════════════════════

function getColLetter(colNumber1Based: number): string {
  let temp = colNumber1Based;
  let letter = '';
  while (temp > 0) {
    let mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

function applyBorder(range: ExcelScript.Range, color: string) {
  const borders = [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical,
  ];
  for (const b of borders) {
    const border = range.getFormat().getRangeBorder(b);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setWeight(ExcelScript.BorderWeight.thin);
    border.setColor(color);
  }
}

function applyDoubleBottomBorder(range: ExcelScript.Range, topColor: string, bottomColor: string) {
  const topBorder = range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop);
  topBorder.setStyle(ExcelScript.BorderLineStyle.continuous);
  topBorder.setWeight(ExcelScript.BorderWeight.thin);
  topBorder.setColor(topColor);

  const botBorder = range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
  botBorder.setStyle(ExcelScript.BorderLineStyle.double);
  botBorder.setColor(bottomColor);

  const leftBorder = range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft);
  leftBorder.setStyle(ExcelScript.BorderLineStyle.continuous);
  leftBorder.setWeight(ExcelScript.BorderWeight.thin);
  leftBorder.setColor(topColor);

  const rightBorder = range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight);
  rightBorder.setStyle(ExcelScript.BorderLineStyle.continuous);
  rightBorder.setWeight(ExcelScript.BorderWeight.thin);
  rightBorder.setColor(topColor);
}

function topNMap(map: Map<string, number>, n: number): [string, number][] {
  return [...map.entries()]
    .filter(([_, v]) => v > 0)
    .sort((a, b) => b[1] - a[1])
    .slice(0, n);
}

function sumColumns(rows: RowRecord[], cols: string[]): Map<string, number> {
  const sums = new Map<string, number>();
  for (const c of cols) sums.set(c, 0);
  for (const r of rows) {
    for (const c of cols) {
      const val = Number(r[c]);
      if (!isNaN(val) && val > 0) {
        sums.set(c, (sums.get(c) ?? 0) + val);
      }
    }
  }
  return sums;
}

function defectHeadersAfter(rows: RowRecord[], markerCol: string): string[] {
  if (rows.length === 0) return [];
  const keys = Object.keys(rows[0]);
  const markerIdx = keys.findIndex(k => k.trim().toLowerCase() === markerCol.trim().toLowerCase());
  if (markerIdx === -1) {
    const subIdx = keys.findIndex(k => k.trim().toLowerCase().includes(markerCol.trim().toLowerCase()));
    if (subIdx === -1) return [];
    return keys.slice(subIdx + 1);
  }
  return keys.slice(markerIdx + 1);
}

function getSiteRows(workbook: ExcelScript.Workbook, type: 'RE' | 'FTT' | 'HFPA', site: string): RowRecord[] {
  const sheets = workbook.getWorksheets();
  const siteUpper = site.toUpperCase();

  // Pattern matching for sheet names
  for (const ws of sheets) {
    const sRawName = ws.getName().toUpperCase();
    const sCleanName = sRawName.replace(/[\s\-_()\[\]]/g, '');

    const isSiteMatch = siteUpper === 'VH2'
      ? (sCleanName.includes('VH2') || sRawName.includes('VH2'))
      : siteUpper === 'VH'
        ? (sCleanName.includes('VH') && !sCleanName.includes('VH2'))
        : siteUpper === 'JV2'
          ? (sCleanName.includes('JV2') || sRawName.includes('JV2'))
          : (sCleanName.includes('JV') && !sCleanName.includes('JV2'));

    if (!isSiteMatch) continue;

    let isTypeMatch = false;
    if (type === 'RE' && (sCleanName.includes('RE') || sCleanName.includes('INSPECTION'))) isTypeMatch = true;
    if (type === 'FTT' && sCleanName.includes('FTT')) isTypeMatch = true;
    if (type === 'HFPA' && sCleanName.includes('HFPA')) isTypeMatch = true;

    if (isTypeMatch) {
      return readSheetAsObjects(ws);
    }
  }

  // Fallback: Check consolidated sheets with a 'Site' or 'Factory' column
  const consolidatedNames = type === 'RE'
    ? ['RE-INSPECTION', 'RE_INSPECTION', 'RAW_RE', 'RE', 'REINSPECTION']
    : type === 'FTT'
      ? ['FTT', 'RAW_FTT']
      : ['HFPA', 'RAW_HFPA'];

  for (const ws of sheets) {
    const sRawName = ws.getName().toUpperCase();
    const sCleanName = sRawName.replace(/[\s\-_()\[\]]/g, '');
    if (consolidatedNames.some(c => sCleanName === c.replace(/[\s\-_()\[\]]/g, '') || sCleanName.startsWith(c.replace(/[\s\-_()\[\]]/g, '')))) {
      const allRows = readSheetAsObjects(ws);
      return allRows.filter(r => {
        const s = String(r['Site'] ?? r['Factory'] ?? r['Factory Site'] ?? '').toUpperCase().trim();
        if (siteUpper === 'VH2') return s === 'VH2' || s.includes('VH2');
        if (siteUpper === 'VH') return (s === 'VH' || s.includes('VH')) && !s.includes('VH2');
        if (siteUpper === 'JV2') return s === 'JV2' || s.includes('JV2');
        if (siteUpper === 'JV') return (s === 'JV' || s.includes('JV')) && !s.includes('JV2');
        return false;
      });
    }
  }

  return [];
}

function readSheetAsObjects(ws: ExcelScript.Worksheet): RowRecord[] {
  const used = ws.getUsedRange();
  if (!used) return [];
  const values = used.getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const rows: RowRecord[] = [];

  for (let r = 1; r < values.length; r++) {
    const rowObj: RowRecord = {};
    let hasVal = false;
    for (let c = 0; c < headers.length; c++) {
      const header = headers[c];
      if (!header) continue;
      const cellVal = values[r][c];
      if (cellVal !== "" && cellVal !== null && cellVal !== undefined) {
        hasVal = true;
      }
      rowObj[header] = cellVal;
    }
    if (hasVal) {
      rows.push(rowObj);
    }
  }
  return rows;
}
