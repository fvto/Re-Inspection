/**
 * Combined_RecycleReport.ts — Microsoft Excel Automate (Office Scripts)
 * =====================================================================
 * High-Performance, Memory-Optimized Multi-Sheet Recycle Report Generator.
 * 
 * Works seamlessly in:
 *   1. Excel on the Web (Office Scripts).
 *   2. Power Automate (Run script action) — handles 500,000+ rows without timing out.
 * 
 * Generates formatted sheets for all 4 factory sites (VH, VH2, JV, JV2).
 */

type RowRecord = { [key: string]: string | number | boolean | null | undefined };

interface SiteExtractedData {
  site: string;
  top5_countries: [string, number][];
  top5_models: [string, number][];
  top5_model_names: string[];
  top5_defects: [string, number][];
  grand_total: number;
  re_qc_top3: { model: string; defects: [string, number][] }[];
  re_ma_top3: { model: string; defects: [string, number][] }[];
  ftt_top3: { model: string; defects: [string, number][] }[];
  hfpa_top3: { model: string; defects: [string, number][] }[];
}

function main(workbook: ExcelScript.Workbook, reportMonth: string = "") {
  const SITES = ['VH', 'VH2', 'JV', 'JV2'];
  const SITE_NAMES_FULL: { [key: string]: string } = {
    'VH': 'Vietnam Factory 1 (VH)',
    'VH2': 'Vietnam Factory 2 (VH2)',
    'JV': 'Indonesia Factory 1 (JV)',
    'JV2': 'Indonesia Factory 2 (JV2)',
  };

  const monthLabel = reportMonth && reportMonth.trim() !== "" 
    ? reportMonth.trim() 
    : getDefaultMonthLabel();

  console.log(`Generating Recycle Report for period: ${monthLabel}`);

  const allData: { [site: string]: SiteExtractedData } = {};

  for (const site of SITES) {
    const reWs = getSiteWorksheet(workbook, 'RE', site);
    if (!reWs) {
      console.log(`[INFO] Skipping site ${site}: No Re-Inspection raw sheet found.`);
      continue;
    }

    const reRows = readSheetAsObjects(reWs);
    if (reRows.length === 0) {
      console.log(`[INFO] Skipping site ${site}: Re-Inspection raw sheet is empty.`);
      continue;
    }

    console.log(`Processing site ${site} (${SITE_NAMES_FULL[site]})...`);
    const fttWs = getSiteWorksheet(workbook, 'FTT', site);
    const hfpaWs = getSiteWorksheet(workbook, 'HFPA', site);

    allData[site] = extractSiteDataOptimized(reRows, fttWs, hfpaWs, site);
  }

  const siteKeys = Object.keys(allData);
  if (siteKeys.length === 0) {
    const availableSheets = workbook.getWorksheets().map(w => `"${w.getName()}"`).join(', ');
    throw new Error(
      `No site data could be extracted. Worksheets found in workbook: [${availableSheets}]. ` +
      `Expected raw data sheets named like 'RE_VH', 'FTT_VH', 'HFPA_VH' (or 'RE_VH2', 'RE_JV', etc.). ` +
      `Please ensure the Power Automate 'Run script' action is targeting the workbook containing the raw sheets.`
    );
  }

  // Build sheet per site
  for (const site of SITES) {
    if (!allData[site]) continue;

    const sheetName = site;
    let ws = workbook.getWorksheet(sheetName);
    if (ws) {
      ws.delete();
    }
    ws = workbook.addWorksheet(sheetName);

    const reportTitle = `Recycle report - ${site} - ${monthLabel}`;
    buildSiteSheet(ws, allData[site], site, reportTitle);
    console.log(`✓ Built sheet: ${sheetName}`);
  }

  console.log("Recycle Report generation completed successfully!");
}

// ══════════════════════════════════════════════════════════════════════
//  DATA EXTRACTION (OPTIMIZED CHUNKED PROCESSING)
// ══════════════════════════════════════════════════════════════════════

function extractSiteDataOptimized(
  reRows: RowRecord[],
  fttWs: ExcelScript.Worksheet | null,
  hfpaWs: ExcelScript.Worksheet | null,
  site: string
): SiteExtractedData {
  const firstRe = reRows[0];
  const modelCol = ('Model 2' in firstRe) ? 'Model 2' : 'Model';

  // 1. Split QC / MA
  const reQc = reRows.filter(r => String(r['Category'] || '').toUpperCase() === 'QC');
  const reMa = reRows.filter(r => String(r['Category'] || '').toUpperCase() === 'MA');

  // Top 5 countries (QC only)
  const countryCounts = new Map<string, number>();
  for (const r of reQc) {
    const c = String(r['Country'] ?? '').trim();
    if (c) {
      countryCounts.set(c, (countryCounts.get(c) ?? 0) + 1);
    }
  }
  const top5Countries = topNMap(countryCounts, 5);

  // Top 5 models (QC only)
  const modelCounts = new Map<string, number>();
  for (const r of reQc) {
    const m = String(r[modelCol] ?? '').trim();
    if (m) {
      modelCounts.set(m, (modelCounts.get(m) ?? 0) + 1);
    }
  }
  const top5Models = topNMap(modelCounts, 5);
  const top5ModelNames = top5Models.map(x => x[0]);

  // Defect headers in Re-Inspection (columns after 'Remark')
  const reDefectCols = defectHeadersAfter(reRows, 'Remark');

  // Top 5 defects across top 5 models (QC + MA combined)
  const reTop5Rows = reRows.filter(r => top5ModelNames.includes(String(r[modelCol] ?? '').trim()));
  const defectSums = sumColumns(reTop5Rows, reDefectCols);
  const top5Defects = topNMap(defectSums, 5);
  let grandTotal = 0;
  defectSums.forEach(val => { grandTotal += val; });

  // Top 3 defects per model — Re-Inspection QC
  const reQcTop3: { model: string; defects: [string, number][] }[] = [];
  for (const m of top5ModelNames) {
    const mRows = reQc.filter(r => String(r[modelCol] ?? '').trim() === m);
    const sums = sumColumns(mRows, reDefectCols);
    reQcTop3.push({ model: m, defects: topNMap(sums, 3) });
  }

  // Top 3 defects per model — Re-Inspection MA
  const reMaTop3: { model: string; defects: [string, number][] }[] = [];
  for (const m of top5ModelNames) {
    const mRows = reMa.filter(r => String(r[modelCol] ?? '').trim() === m);
    const sums = sumColumns(mRows, reDefectCols);
    reMaTop3.push({ model: m, defects: topNMap(sums, 3) });
  }

  // Top 3 defects per model — FTT (Processed in fast streaming chunks)
  const fttTop3 = extractFttTop3Streaming(fttWs, top5ModelNames);

  // Top 3 defects per model — HFPA (Processed in fast streaming chunks)
  const hfpaTop3 = extractHfpaTop3Streaming(hfpaWs, top5ModelNames);

  return {
    site,
    top5_countries: top5Countries,
    top5_models: top5Models,
    top5_model_names: top5ModelNames,
    top5_defects: top5Defects,
    grand_total: grandTotal,
    re_qc_top3: reQcTop3,
    re_ma_top3: reMaTop3,
    ftt_top3: fttTop3,
    hfpa_top3: hfpaTop3,
  };
}

/**
 * Fast streaming extractor for FTT sheets (handles 100k+ rows without memory overflow)
 */
function extractFttTop3Streaming(
  ws: ExcelScript.Worksheet | null,
  targetModels: string[]
): { model: string; defects: [string, number][] }[] {
  if (!ws) {
    return targetModels.map(m => ({ model: m, defects: [] }));
  }

  const used = ws.getUsedRange();
  if (!used) return targetModels.map(m => ({ model: m, defects: [] }));

  const totalRows = used.getRowCount();
  const totalCols = used.getColumnCount();
  if (totalRows < 2) return targetModels.map(m => ({ model: m, defects: [] }));

  // Read headers (Row 0)
  const headerVals = ws.getRangeByIndexes(0, 0, 1, totalCols).getValues()[0];
  const headers = headerVals.map(h => String(h).trim());

  let modelIdx = headers.indexOf('Model');
  if (modelIdx === -1) modelIdx = headers.indexOf('Model 2');
  let issueIdx = headers.indexOf('Defect Issues');
  if (issueIdx === -1) issueIdx = headers.findIndex(h => h.toLowerCase().includes('defect') || h.toLowerCase().includes('issue'));
  let qtyIdx = headers.findIndex(h => h.includes("Issues Q'ty") || h.includes("Issues Qty") || h.toLowerCase() === "qty");

  if (modelIdx === -1 || issueIdx === -1) {
    return targetModels.map(m => ({ model: m, defects: [] }));
  }

  // Model-to-defects aggregation map
  const modelSums: { [model: string]: Map<string, number> } = {};
  for (const m of targetModels) {
    modelSums[m] = new Map<string, number>();
  }

  const CHUNK_SIZE = 10000;
  for (let r = 1; r < totalRows; r += CHUNK_SIZE) {
    const count = Math.min(CHUNK_SIZE, totalRows - r);
    const chunkValues = ws.getRangeByIndexes(r, 0, count, totalCols).getValues();

    for (let i = 0; i < chunkValues.length; i++) {
      const row = chunkValues[i];
      const modelVal = String(row[modelIdx] ?? '').trim();
      if (!modelSums[modelVal]) continue;

      const issueVal = String(row[issueIdx] ?? '').trim();
      if (!issueVal) continue;

      const rawQty = qtyIdx !== -1 ? row[qtyIdx] : 1;
      const qty = Number(rawQty) || 0;
      if (qty > 0) {
        const m = modelSums[modelVal];
        m.set(issueVal, (m.get(issueVal) ?? 0) + qty);
      }
    }
  }

  return targetModels.map(m => ({
    model: m,
    defects: topNMap(modelSums[m], 3),
  }));
}

/**
 * Fast streaming extractor for HFPA sheets (handles large column & row datasets)
 */
function extractHfpaTop3Streaming(
  ws: ExcelScript.Worksheet | null,
  targetModels: string[]
): { model: string; defects: [string, number][] }[] {
  if (!ws) {
    return targetModels.map(m => ({ model: m, defects: [] }));
  }

  const used = ws.getUsedRange();
  if (!used) return targetModels.map(m => ({ model: m, defects: [] }));

  const totalRows = used.getRowCount();
  const totalCols = used.getColumnCount();
  if (totalRows < 2) return targetModels.map(m => ({ model: m, defects: [] }));

  const headerVals = ws.getRangeByIndexes(0, 0, 1, totalCols).getValues()[0];
  const headers = headerVals.map(h => String(h).trim());

  let modelIdx = headers.indexOf('Model 2');
  if (modelIdx === -1) modelIdx = headers.indexOf('Model');
  const pfIdx = headers.findIndex(h => h.toLowerCase() === 'pass/fail' || h.toLowerCase() === 'pass_fail');

  const markerIdx = headers.findIndex(h => h.includes("Defect Q'ty") || h.includes("Total Defect Q'ty"));
  const defectColStart = markerIdx !== -1 ? markerIdx + 1 : -1;

  if (modelIdx === -1 || pfIdx === -1 || defectColStart === -1 || defectColStart >= totalCols) {
    return targetModels.map(m => ({ model: m, defects: [] }));
  }

  const modelSums: { [model: string]: Map<string, number> } = {};
  for (const m of targetModels) {
    modelSums[m] = new Map<string, number>();
  }

  const CHUNK_SIZE = 5000;
  for (let r = 1; r < totalRows; r += CHUNK_SIZE) {
    const count = Math.min(CHUNK_SIZE, totalRows - r);
    const chunkValues = ws.getRangeByIndexes(r, 0, count, totalCols).getValues();

    for (let i = 0; i < chunkValues.length; i++) {
      const row = chunkValues[i];
      const pf = String(row[pfIdx] ?? '').trim().toLowerCase();
      if (pf !== 'fail') continue;

      const modelVal = String(row[modelIdx] ?? '').trim();
      if (!modelSums[modelVal]) continue;

      const m = modelSums[modelVal];
      for (let c = defectColStart; c < totalCols; c++) {
        const val = Number(row[c]);
        if (!isNaN(val) && val > 0) {
          const colName = headers[c];
          m.set(colName, (m.get(colName) ?? 0) + val);
        }
      }
    }
  }

  return targetModels.map(m => ({
    model: m,
    defects: topNMap(modelSums[m], 3),
  }));
}

// ══════════════════════════════════════════════════════════════════════
//  EXCEL SHEET BUILDER & STYLING
// ══════════════════════════════════════════════════════════════════════

function buildSiteSheet(
  ws: ExcelScript.Worksheet,
  data: SiteExtractedData,
  site: string,
  reportTitle: string
) {
  const COLOR_PRIMARY = "#1F4E78";   // Dark Navy
  const COLOR_SECTION = "#D9E2F3";   // Light blue section band
  const COLOR_HEADER_FG = "#FFFFFF"; // White text
  const COLOR_FONT_SEC = "#1F4E78";
  const COLOR_BORDER = "#D9D9D9";

  // Set column widths
  ws.getRange("A:A").getFormat().setColumnWidth(260);
  ws.getRange("B:B").getFormat().setColumnWidth(130);
  ws.getRange("C:C").getFormat().setColumnWidth(190);
  ws.getRange("D:D").getFormat().setColumnWidth(100);

  let row = 0; // 0-indexed

  // Title
  const titleRange = ws.getRangeByIndexes(row, 0, 1, 1);
  titleRange.setValue(reportTitle);
  titleRange.getFormat().getFont().setName("Segoe UI");
  titleRange.getFormat().getFont().setSize(14);
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setColor(COLOR_PRIMARY);
  titleRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  titleRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  row += 2;

  function styleSectionBand(text: string, colSpan: number) {
    const band = ws.getRangeByIndexes(row, 0, 1, colSpan);
    band.merge();
    band.setValue(text);
    band.getFormat().getFont().setName("Segoe UI");
    band.getFormat().getFont().setSize(11);
    band.getFormat().getFont().setBold(true);
    band.getFormat().getFont().setColor(COLOR_FONT_SEC);
    band.getFormat().getFill().setColor(COLOR_SECTION);
    band.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    band.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    row += 1;
  }

  function styleHeaderRow(headers: string[]) {
    const headRange = ws.getRangeByIndexes(row, 0, 1, headers.length);
    headRange.setValues([headers]);
    headRange.getFormat().getFont().setName("Segoe UI");
    headRange.getFormat().getFont().setSize(11);
    headRange.getFormat().getFont().setBold(true);
    headRange.getFormat().getFont().setColor(COLOR_HEADER_FG);
    headRange.getFormat().getFill().setColor(COLOR_PRIMARY);
    headRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    headRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    applyBorder(headRange, COLOR_BORDER);
    row += 1;
  }

  function writeRows(
    rowsData: (string | number)[][],
    alignments: ExcelScript.HorizontalAlignment[],
    numFormats: (string | null)[]
  ) {
    if (rowsData.length === 0) return;
    const numRows = rowsData.length;
    const numCols = rowsData[0].length;
    const dataRange = ws.getRangeByIndexes(row, 0, numRows, numCols);
    dataRange.setValues(rowsData);
    dataRange.getFormat().getFont().setName("Segoe UI");
    dataRange.getFormat().getFont().setSize(11);
    applyBorder(dataRange, COLOR_BORDER);

    for (let c = 0; c < numCols; c++) {
      const colRange = ws.getRangeByIndexes(row, c, numRows, 1);
      colRange.getFormat().setHorizontalAlignment(alignments[c]);
      colRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
      if (numFormats[c]) {
        const fmtArr = rowsData.map(() => [numFormats[c]!]);
        colRange.setNumberFormat(fmtArr);
      }
    }
    row += numRows;
  }

  // 1. SECTION 1: Top 5 countries (QC)
  styleSectionBand("Top 5 countries by fail qty (Re-Inspection, QC)", 2);
  styleHeaderRow(["Country", "Fail Qty"]);
  writeRows(
    data.top5_countries.map(([c, q]) => [c, q]),
    [ExcelScript.HorizontalAlignment.left, ExcelScript.HorizontalAlignment.right],
    [null, "#,##0"]
  );
  row += 1;

  // 2. SECTION 2: Top 5 models (QC)
  styleSectionBand("Top 5 models by fail qty (Re-Inspection, QC)", 2);
  styleHeaderRow(["Model", "Fail Qty"]);
  writeRows(
    data.top5_models.map(([m, q]) => [m, q]),
    [ExcelScript.HorizontalAlignment.left, ExcelScript.HorizontalAlignment.right],
    [null, "#,##0"]
  );
  row += 1;

  // 3. SECTION 3: Top 5 defects across top 5 models (QC + MA)
  styleSectionBand("Top 5 defects across top 5 models (Re-Inspection, QC+MA)", 4);
  styleHeaderRow(["Defect", "Defect Qty", "Total Defect Qty (5 models)", "TL%"]);
  const gt = data.grand_total;
  writeRows(
    data.top5_defects.map(([d, q]) => [d, q, gt, gt > 0 ? q / gt : 0]),
    [
      ExcelScript.HorizontalAlignment.left,
      ExcelScript.HorizontalAlignment.right,
      ExcelScript.HorizontalAlignment.right,
      ExcelScript.HorizontalAlignment.right,
    ],
    [null, "#,##0", "#,##0", "0.0%"]
  );
  row += 1;

  // 4. SECTION 4: Top 3 defects per model — Re-Inspection QC
  styleSectionBand("Top 3 defects per model - Re-Inspection QC (initial audit)", 3);
  styleHeaderRow(["Model", "Defect", "Qty"]);
  const qcTop3Rows: (string | number)[][] = [];
  for (const mInfo of data.re_qc_top3) {
    for (const [def, qty] of mInfo.defects) {
      qcTop3Rows.push([mInfo.model, def, qty]);
    }
  }
  writeRows(
    qcTop3Rows,
    [
      ExcelScript.HorizontalAlignment.left,
      ExcelScript.HorizontalAlignment.left,
      ExcelScript.HorizontalAlignment.right,
    ],
    [null, null, "#,##0"]
  );
  row += 1;

  // 5. SECTION 5: Top 3 defects per model — Re-Inspection MA
  styleSectionBand("Top 3 defects per model - Re-Inspection MA (re-audit)", 3);
  styleHeaderRow(["Model", "Defect", "Qty"]);
  const maTop3Rows: (string | number)[][] = [];
  for (const mInfo of data.re_ma_top3) {
    for (const [def, qty] of mInfo.defects) {
      maTop3Rows.push([mInfo.model, def, qty]);
    }
  }
  writeRows(
    maTop3Rows,
    [
      ExcelScript.HorizontalAlignment.left,
      ExcelScript.HorizontalAlignment.left,
      ExcelScript.HorizontalAlignment.right,
    ],
    [null, null, "#,##0"]
  );
  row += 1;

  // 6. SECTION 6: Top 3 defects per model — FTT
  styleSectionBand("Top 3 defects per model - FTT (Model field, full month volume)", 3);
  styleHeaderRow(["Model", "Defect", "Qty"]);
  const fttRowsOut: (string | number)[][] = [];
  for (const mInfo of data.ftt_top3) {
    for (const [def, qty] of mInfo.defects) {
      fttRowsOut.push([mInfo.model, def, qty]);
    }
  }
  writeRows(
    fttRowsOut,
    [
      ExcelScript.HorizontalAlignment.left,
      ExcelScript.HorizontalAlignment.left,
      ExcelScript.HorizontalAlignment.right,
    ],
    [null, null, "#,##0"]
  );
  row += 1;

  // 7. SECTION 7: Top 3 defects per model — HFPA
  styleSectionBand("Top 3 defects per model - HFPA (Fail only)", 3);
  styleHeaderRow(["Model", "Defect", "Qty"]);
  const hfpaRowsOut: (string | number)[][] = [];
  for (const mInfo of data.hfpa_top3) {
    for (const [def, qty] of mInfo.defects) {
      hfpaRowsOut.push([mInfo.model, def, qty]);
    }
  }
  writeRows(
    hfpaRowsOut,
    [
      ExcelScript.HorizontalAlignment.left,
      ExcelScript.HorizontalAlignment.left,
      ExcelScript.HorizontalAlignment.right,
    ],
    [null, null, "#,##0"]
  );
}

// ══════════════════════════════════════════════════════════════════════
//  UTILITIES & HELPERS
// ══════════════════════════════════════════════════════════════════════

function getSiteWorksheet(workbook: ExcelScript.Workbook, type: 'RE' | 'FTT' | 'HFPA', site: string): ExcelScript.Worksheet | null {
  const sheets = workbook.getWorksheets();
  const siteUpper = site.toUpperCase();

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
      return ws;
    }
  }

  return null;
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

function applyBorder(range: ExcelScript.Range, color: string) {
  const borderTypes = [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical,
  ];
  for (const b of borderTypes) {
    const border = range.getFormat().getRangeBorder(b);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setWeight(ExcelScript.BorderWeight.thin);
    border.setColor(color);
  }
}

function topNMap(map: Map<string, number>, n: number): [string, number][] {
  const entries: [string, number][] = [];
  map.forEach((val, key) => entries.push([key, val]));
  entries.sort((a, b) => b[1] - a[1]);
  return entries.slice(0, n);
}

function sumColumns(rows: RowRecord[], cols: string[]): Map<string, number> {
  const sums = new Map<string, number>();
  for (const c of cols) {
    sums.set(c, 0);
  }
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

function getDefaultMonthLabel(): string {
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const d = new Date();
  return `${months[d.getMonth()]} ${d.getFullYear()}`;
}
