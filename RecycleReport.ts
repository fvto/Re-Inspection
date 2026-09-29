/**
 * RecycleReport.ts — Microsoft Excel Automate (Office Scripts)
 * =============================================================
 * Multi-Site Automated Recycle Report Generator for Power Automate.
 * 
 * Works on:
 *   - Freshly created blank Excel workbooks (e.g. containing only 'Trang_tính1' / 'Sheet1').
 *   - Direct JSON row arrays passed from Power Automate "List rows present in a table".
 * 
 * Capabilities:
 *   1. All-in-One Multi-Site (Default): Pass consolidated Re-Inspection, FTT, and HFPA row arrays.
 *      The script will automatically detect and split data for VH, VH2, JV, and JV2,
 *      and generate 4 beautifully styled sheets ['VH', 'VH2', 'JV', 'JV2'].
 *   2. Per-Site Mode: Specify a site (e.g. "VH") to generate a single site sheet.
 *   3. Automatically deletes the default blank sheet ('Trang_tính1' / 'Sheet1') upon completion.
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

function main(
  workbook: ExcelScript.Workbook,
  reportMonth: string = "",
  reInspectionRows: object[] = [],
  fttRows: object[] = [],
  hfpaRows: object[] = [],
  targetSite: string = "" // Optional: Leave blank to process ALL sites (VH, VH2, JV, JV2)
) {
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

  console.log(`Generating Multi-Site Recycle Report for period: ${monthLabel}`);

  const allRe = (reInspectionRows || []) as RowRecord[];
  const allFtt = (fttRows || []) as RowRecord[];
  const allHfpa = (hfpaRows || []) as RowRecord[];

  if (allRe.length === 0) {
    throw new Error("reInspectionRows is empty. Please pass raw Re-Inspection rows from Power Automate.");
  }

  // Determine which sites to build
  const sitesToProcess = (targetSite && targetSite.trim() !== "")
    ? [targetSite.trim().toUpperCase()]
    : SITES;

  let generatedCount = 0;

  for (const site of sitesToProcess) {
    // Filter rows for this specific site
    const siteRe = filterRowsForSite(allRe, site);
    const siteFtt = filterRowsForSite(allFtt, site);
    const siteHfpa = filterRowsForSite(allHfpa, site);

    if (siteRe.length === 0) {
      console.log(`[INFO] Skipping site ${site}: No Re-Inspection records found.`);
      continue;
    }

    console.log(`Processing site ${site} (${SITE_NAMES_FULL[site] ?? site})...`);
    const siteData = extractSiteData(siteRe, siteFtt, siteHfpa, site);

    // Create / Overwrite sheet for this site
    const sheetName = site;
    let ws = workbook.getWorksheet(sheetName);
    if (ws) {
      ws.delete();
    }
    ws = workbook.addWorksheet(sheetName);

    const reportTitle = `Recycle report - ${site} - ${monthLabel}`;
    buildSiteSheet(ws, siteData, site, reportTitle);
    console.log(`✓ Built sheet: ${sheetName}`);
    generatedCount++;
  }

  if (generatedCount === 0) {
    throw new Error(`Could not generate any sheets for sites [${sitesToProcess.join(', ')}]. Please check that the raw data rows contain a 'Site' or 'Factory' column matching VH, VH2, JV, or JV2.`);
  }

  // Clean up default blank sheets (e.g. Trang_tính1 / Sheet1 / Sheet)
  const defaultSheetNames = ['TRANG_TÍNH1', 'TRANG_TINH1', 'SHEET1', 'SHEET'];
  for (const ws of workbook.getWorksheets()) {
    const sNameUpper = ws.getName().toUpperCase().replace(/[\s\-_()\[\]]/g, '');
    if (defaultSheetNames.includes(sNameUpper) && workbook.getWorksheets().length > 1) {
      try {
        ws.delete();
        console.log(`Deleted default blank worksheet: ${ws.getName()}`);
      } catch (e) {
        // Ignore deletion errors if protected
      }
    }
  }

  console.log(`Successfully generated ${generatedCount} site report sheets!`);
}

// ══════════════════════════════════════════════════════════════════════
//  ROW FILTERING & SITE EXTRACTION
// ══════════════════════════════════════════════════════════════════════

function filterRowsForSite(rows: RowRecord[], site: string): RowRecord[] {
  if (rows.length === 0) return [];
  const siteUpper = site.toUpperCase();

  // Check if rows have a Site / Factory column
  const first = rows[0];
  const siteColKey = Object.keys(first).find(k => {
    const norm = k.trim().toLowerCase();
    return norm === 'site' || norm === 'factory' || norm === 'factory site' || norm === 'facility' || norm === 'plant';
  });

  if (!siteColKey) {
    // If no site column exists, assume the passed rows are already dedicated to this site
    return rows;
  }

  return rows.filter(r => {
    const val = String(r[siteColKey] ?? '').toUpperCase().trim();
    if (siteUpper === 'VH2') return val === 'VH2' || val.includes('VH2');
    if (siteUpper === 'VH') return (val === 'VH' || val.includes('VH')) && !val.includes('VH2');
    if (siteUpper === 'JV2') return val === 'JV2' || val.includes('JV2');
    if (siteUpper === 'JV') return (val === 'JV' || val.includes('JV')) && !val.includes('JV2');
    return val === siteUpper;
  });
}

function extractSiteData(
  reRows: RowRecord[],
  fttRows: RowRecord[],
  hfpaRows: RowRecord[],
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

  // Top 3 defects per model — FTT
  const fttTop3: { model: string; defects: [string, number][] }[] = [];
  for (const m of top5ModelNames) {
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

  // Top 3 defects per model — HFPA (Fail only)
  const hfpaDefectCols = defectHeadersAfter(hfpaRows, "Defect Q'ty (Total Defect Q'ty)");
  const hfpaFail = hfpaRows.filter(r => String(r['Pass/Fail'] ?? '').trim().toLowerCase() === 'fail');
  const hfpaTop3: { model: string; defects: [string, number][] }[] = [];
  for (const m of top5ModelNames) {
    let mRows = hfpaFail.filter(r => String(r['Model 2'] ?? '').trim() === m);
    if (mRows.length === 0 && hfpaFail.length > 0 && 'Model' in hfpaFail[0]) {
      mRows = hfpaFail.filter(r => String(r['Model'] ?? '').trim() === m);
    }
    const sums = sumColumns(mRows, hfpaDefectCols);
    hfpaTop3.push({ model: m, defects: topNMap(sums, 3) });
  }

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

  // 1. Top 5 countries (QC)
  styleSectionBand("Top 5 countries by fail qty (Re-Inspection, QC)", 2);
  styleHeaderRow(["Country", "Fail Qty"]);
  writeRows(
    data.top5_countries.map(([c, q]) => [c, q]),
    [ExcelScript.HorizontalAlignment.left, ExcelScript.HorizontalAlignment.right],
    [null, "#,##0"]
  );
  row += 1;

  // 2. Top 5 models (QC)
  styleSectionBand("Top 5 models by fail qty (Re-Inspection, QC)", 2);
  styleHeaderRow(["Model", "Fail Qty"]);
  writeRows(
    data.top5_models.map(([m, q]) => [m, q]),
    [ExcelScript.HorizontalAlignment.left, ExcelScript.HorizontalAlignment.right],
    [null, "#,##0"]
  );
  row += 1;

  // 3. Top 5 defects across top 5 models (QC + MA)
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

  // 4. Top 3 defects per model — Re-Inspection QC
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

  // 5. Top 3 defects per model — Re-Inspection MA
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

  // 6. Top 3 defects per model — FTT
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

  // 7. Top 3 defects per model — HFPA
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

