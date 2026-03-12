/**
 * excel-reader.ts
 * Reads a "dirty" Checkers vnd-art-sales Excel file from the VENDOR VIEW sheet
 * and returns a structured, normalised representation ready for clean file generation.
 */

import * as XLSX from 'xlsx';

export interface ExtraSheet {
  name: string;
  data: (string | number | null)[][];
}

export interface ParsedDirtyFile {
  vendorName: string;       // from Vendor column first data row
  dateColumns: string[];    // e.g. ["21.12.2025 units", "28.12.2025 units", ...]
  latestDate: Date;         // most recent date across all date cols
  rows: Record<string, string | number | null>[];  // keyed by clean column name
  extraSheets: ExtraSheet[]; // all sheets from dirty file except Vendor View & Separate View
}

// Columns to explicitly skip (by lowercase header)
const SKIP_HEADERS = new Set([
  'vendor', 'purchase org', 'country', 'site banner', 'division',
]);

// Pattern to detect weekly VALUE columns (skip — clean file only uses units)
const DATE_VALUE_PATTERN = /^\d{2}\.\d{2}\.\d{4}\s+value$/i;
const DATE_UNIT_PATTERN  = /^\d{2}\.\d{2}\.\d{4}\s+units$/i;
const GROWTH_PATTERN     = /growth/i;

// Dirty header (lowercase) → clean column name
const HEADER_MAP: Record<string, string> = {
  'vendor sub-range':    'Vendor Subrange',
  'vendor sub range':    'Vendor Subrange',
  'site':                'Site',
  'department':          'Department',
  'category group':      'Catergory Group',   // preserve typo from clean template
  'category':            'Category',
  'sub-category':        'Sub-category',
  'sub category':        'Sub-category',
  'article key':         'Article Key',
  'article':             'Article',
  'pl label 4':          'PL Label 4',
  'sell uom':            'Sell UOM',
  'orderable uom':       'Orderable UOM',
  'rp type':             'RP Type',
  'new':                 'New',
  'lstd':                'Lstd',
  'stk':                 'Stk',
  'blk':                 'Ord Blk',
  'ord blk':             'Ord Blk',
  'active stores':       'Active Stores',
  'sell price (incl)':   'Sell Price (Incl)',
  'total sales units':   'Total Sales Units',
  'curr ros':            'Curr ROS',
  'total sales value':   'Latest Period Sales (Excl)',
  'wos':                 'WOS',
};

export function parseDirtyFile(buffer: Buffer): ParsedDirtyFile {
  const wb = XLSX.read(buffer, { type: 'buffer', cellFormula: true, cellDates: false });

  // Find VENDOR VIEW sheet (case-insensitive)
  const sheetName = wb.SheetNames.find(n => n.trim().toLowerCase() === 'vendor view');
  if (!sheetName) {
    const found = wb.SheetNames.join(', ');
    throw new Error(`No "Vendor view" sheet found. Sheets in file: ${found}`);
  }

  const ws = wb.Sheets[sheetName];
  const allRows = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
    header: 1,
    defval: null,
    raw: true,
  });

  // Find header row — scan rows 0-9 for one containing known headers
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(10, allRows.length); i++) {
    const rowLower = allRows[i].map(c => String(c ?? '').trim().toLowerCase());
    if (
      rowLower.includes('vendor sub-range') ||
      rowLower.includes('vendor sub range') ||
      rowLower.includes('article key')
    ) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx === -1) throw new Error('Could not locate header row in VENDOR VIEW sheet');

  const rawHeaders = allRows[headerRowIdx].map(c => String(c ?? '').trim());
  const dataRows   = allRows.slice(headerRowIdx + 1).filter(r =>
    r.some(c => c !== null && c !== '')
  );

  // Detect formula-heavy columns (user-inserted scratch columns)
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1:A1');
  const formulaCols = new Set<number>();
  for (let c = range.s.c; c <= range.e.c; c++) {
    let formulaCount = 0;
    let totalCount   = 0;
    for (let r = headerRowIdx + 1; r <= Math.min(headerRowIdx + 30, range.e.r); r++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== undefined) {
        totalCount++;
        if (cell.f) formulaCount++;
      }
    }
    if (totalCount > 0 && formulaCount / totalCount > 0.4) formulaCols.add(c);
  }

  // Build column map: index → clean name (null = skip)
  const colMap: (string | null)[] = rawHeaders.map((h, idx) => {
    const hLower = h.toLowerCase();
    if (!h)                         return null;
    if (formulaCols.has(idx))       return null;
    if (SKIP_HEADERS.has(hLower))   return null;
    if (DATE_VALUE_PATTERN.test(h)) return null;
    if (GROWTH_PATTERN.test(h))     return null;
    if (DATE_UNIT_PATTERN.test(h))  return h;       // keep date cols as-is
    return HEADER_MAP[hLower] ?? null;              // null = unknown, skip
  });

  // Collect date unit columns in order
  const dateColumns: string[] = [];
  rawHeaders.forEach((h, idx) => {
    if (DATE_UNIT_PATTERN.test(h) && !formulaCols.has(idx)) {
      dateColumns.push(h);
    }
  });

  // Parse latest date from date column headers
  let latestDate = new Date(0);
  dateColumns.forEach(col => {
    const parts = col.split(' ')[0].split('.');
    if (parts.length === 3) {
      const [d, m, y] = parts.map(Number);
      const date = new Date(y, m - 1, d);
      if (date > latestDate) latestDate = date;
    }
  });

  // Get vendor name from first data row, Vendor column
  const vendorIdx = rawHeaders.findIndex(h => h.toLowerCase() === 'vendor');
  let vendorName = '';
  if (vendorIdx >= 0 && dataRows.length > 0) {
    vendorName = String(dataRows[0][vendorIdx] ?? '');
  }

  // Build output rows
  const rows: Record<string, string | number | null>[] = dataRows.map(row => {
    const out: Record<string, string | number | null> = {};
    colMap.forEach((cleanName, idx) => {
      if (cleanName) {
        const v = row[idx];
        out[cleanName] = v === undefined ? null : (v as string | number | null);
      }
    });
    return out;
  });

  // Collect all other sheets (pass-through to clean file)
  const SKIP_SHEETS = new Set(['vendor view', 'separate view']);
  const extraSheets: ExtraSheet[] = wb.SheetNames
    .filter(n => !SKIP_SHEETS.has(n.trim().toLowerCase()))
    .map(n => ({
      name: n,
      data: XLSX.utils.sheet_to_json<(string | number | null)[]>(wb.Sheets[n], {
        header: 1,
        defval: null,
        raw: true,
      }),
    }));

  return { vendorName, dateColumns, latestDate, rows, extraSheets };
}
