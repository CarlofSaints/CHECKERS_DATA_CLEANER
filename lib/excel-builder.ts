/**
 * excel-builder.ts
 * Generates the clean "Checkers B2B" Excel file using exceljs.
 * Preserves all 5 sheets expected by the upload system.
 * Only "Separate View" is fully populated; others exist as placeholders.
 */

import ExcelJS from 'exceljs';
import { ParsedDirtyFile } from './excel-reader';

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function formatDate(d: Date): string {
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
}

// Columns before the dynamic date columns (positions A–R = 1–18)
const PRE_DATE_COLS = [
  'Vendor Subrange', 'Site', 'Department', 'Catergory Group', 'Category',
  'Sub-category', 'Article Key', 'Article', 'PL Label 4', 'Sell UOM',
  'Orderable UOM', 'RP Type', 'New', 'Lstd', 'Stk', 'Ord Blk',
  'Active Stores', 'Sell Price (Incl)',
];

// Columns after the dynamic date columns (positions Y–AH = 25–34 with 6 date cols)
const POST_DATE_COLS = [
  'Total Sales Units',
  'Curr ROS',
  'Latest Period Sales (Excl)',
  'Stock Value @MAC',         // blank — not in dirty file
  'Stock Qty',                // blank — not in dirty file
  'WOS',
  'Date Last Sold',           // blank — not in dirty file
  'Date Last Ordered',        // blank — not in dirty file
  'Date Last Receipted',      // blank — not in dirty file
  'DC Stock',                 // blank — not in dirty file
];

// Columns that have no source in the dirty file → always blank
const BLANK_COLS = new Set([
  'Stock Value @MAC', 'Stock Qty',
  'Date Last Sold', 'Date Last Ordered', 'Date Last Receipted', 'DC Stock',
]);

// Shared cell styles
const HEADER_FILL: ExcelJS.Fill = {
  type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' },
};
const GROUP_FILL: ExcelJS.Fill = {
  type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' },
};
const CENTER_ALIGN: Partial<ExcelJS.Alignment> = { horizontal: 'center', vertical: 'middle', wrapText: true };
const BOLD_FONT: Partial<ExcelJS.Font> = { bold: true };

// Sheets fully built by this tool — never overwritten from dirty file
const BUILT_SHEETS = new Set(['separate view', 'vendor view']);

export async function buildCleanFile(data: ParsedDirtyFile, clientName: string): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();

  // Index extra sheets from dirty file by lowercase name for O(1) lookup
  const extraByName = new Map<string, (string | number | null)[][]>();
  for (const es of data.extraSheets) {
    extraByName.set(es.name.trim().toLowerCase(), es.data);
  }

  // Helper: create a sheet and populate from dirty file if a matching name exists
  function addSheet(name: string): ExcelJS.Worksheet {
    const ws = wb.addWorksheet(name);
    const srcData = extraByName.get(name.trim().toLowerCase());
    if (srcData && !BUILT_SHEETS.has(name.trim().toLowerCase())) {
      srcData.forEach(rowArr => {
        ws.addRow(rowArr);
      });
    }
    return ws;
  }

  // Create all 5 sheets in the expected order
  addSheet('Parameter Selection');
  addSheet('Consolidated View');
  addSheet('Consolidated by P.Org');
  const separateWs = wb.addWorksheet('Separate View');
  addSheet('Vendor and VSR Total View');

  buildSeparateView(separateWs, data, clientName);

  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

function buildSeparateView(ws: ExcelJS.Worksheet, data: ParsedDirtyFile, clientName: string) {
  const allCols   = [...PRE_DATE_COLS, ...data.dateColumns, ...POST_DATE_COLS];
  const totalCols = allCols.length; // 34 (A–AH)

  // ── Rows 1–5: header area ────────────────────────────────────────────────

  // Row 1: Title
  ws.addRow(['Vendor Article Sales and Stock']);
  ws.mergeCells(1, 1, 1, totalCols);
  const r1c1 = ws.getRow(1).getCell(1);
  r1c1.font      = { bold: true, size: 14 };
  r1c1.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

  // Row 2: Period
  ws.addRow([`For Last 6 Weeks, Ending ${formatDate(data.latestDate)}`]);
  ws.mergeCells(2, 1, 2, totalCols);
  const r2c1 = ws.getRow(2).getCell(1);
  r2c1.font      = { size: 11 };
  r2c1.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

  // Row 3: empty
  ws.addRow([]);

  // Row 4: Vendor / client name
  const vendorDisplay = `Vendor: ${data.vendorName || clientName}`;
  ws.addRow([vendorDisplay]);
  ws.mergeCells(4, 1, 4, 5);
  ws.getRow(4).getCell(1).font = { bold: true, size: 11 };

  // Row 5: empty
  ws.addRow([]);

  // ── Row 6: Group headers ─────────────────────────────────────────────────
  ws.addRow([]);  // Row 6 placeholder — cells set individually below

  // # Sites  → N6:Q6 = cols 14–17
  ws.mergeCells(6, 14, 6, 17);
  applyGroupCell(ws.getRow(6).getCell(14), '# Sites');

  // Sell Price (Incl) → R6:R7 = col 18, merged down into row 7
  ws.mergeCells(6, 18, 7, 18);
  applyGroupCell(ws.getRow(6).getCell(18), 'Sell Price (Incl)');

  // POS Sales → S6:AA6 = cols 19–27 (6 date cols + Total Sales + Curr ROS + Latest Period Sales)
  ws.mergeCells(6, 19, 6, 27);
  applyGroupCell(ws.getRow(6).getCell(19), 'POS Sales');

  // All Store Stock → AB6:AG6 = cols 28–33
  ws.mergeCells(6, 28, 6, 33);
  applyGroupCell(ws.getRow(6).getCell(28), 'All Store Stock');

  // DC Stock → AH6:AH7 = col 34, merged down into row 7
  ws.mergeCells(6, 34, 7, 34);
  applyGroupCell(ws.getRow(6).getCell(34), 'DC Stock');

  // ── Row 7: Column headers ────────────────────────────────────────────────
  // Use getRow(7) directly — addRow() would append as row 8 because the
  // Sell Price / DC Stock merges (6:7) have already registered row 7 internally.
  const headerRow = ws.getRow(7);
  allCols.forEach((col, idx) => {
    const cell     = headerRow.getCell(idx + 1);
    cell.value     = col;
    cell.font      = BOLD_FONT;
    cell.fill      = HEADER_FILL;
    cell.alignment = CENTER_ALIGN;
    cell.border    = { bottom: { style: 'thin', color: { argb: 'FF9DC3E6' } } };
  });

  // ── Rows 8+: Data ────────────────────────────────────────────────────────
  data.rows.forEach(row => {
    const rowData = allCols.map(col => BLANK_COLS.has(col) ? null : (row[col] ?? null));
    ws.addRow(rowData);
  });

  // ── Column widths ────────────────────────────────────────────────────────
  const WIDTHS: Record<string, number> = {
    'Vendor Subrange': 22, 'Site': 30, 'Department': 18,
    'Catergory Group': 18, 'Category': 18, 'Sub-category': 18,
    'Article Key': 14, 'Article': 30, 'PL Label 4': 18,
    'Sell UOM': 10, 'Orderable UOM': 14, 'RP Type': 10,
    'New': 7, 'Lstd': 7, 'Stk': 7, 'Ord Blk': 9, 'Active Stores': 12,
    'Sell Price (Incl)': 14,
    'Total Sales Units': 14, 'Curr ROS': 10, 'Latest Period Sales (Excl)': 20,
    'Stock Value @MAC': 16, 'Stock Qty': 10, 'WOS': 8,
    'Date Last Sold': 14, 'Date Last Ordered': 14, 'Date Last Receipted': 16,
    'DC Stock': 10,
  };
  allCols.forEach((col, idx) => {
    const colNum = idx + 1;
    ws.getColumn(colNum).width = WIDTHS[col] ?? 14;
  });

  // Freeze panes at row 8 (keep header rows visible when scrolling)
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 7 }];
}

function applyGroupCell(cell: ExcelJS.Cell, value: string) {
  cell.value     = value;
  cell.font      = { bold: true, size: 10 };
  cell.fill      = GROUP_FILL;
  cell.alignment = CENTER_ALIGN;
  cell.border    = {
    bottom: { style: 'thin', color: { argb: 'FF9DC3E6' } },
  };
}
