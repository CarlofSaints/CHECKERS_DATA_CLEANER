import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { parseDirtyFile } from '@/lib/excel-reader';

export const maxDuration = 30;

/**
 * Extract all recognisable dates from a filename string.
 * Handles: YYYYMMDD, YYYY-MM-DD, YYYY.MM.DD, DD-MM-YYYY, DD.MM.YYYY
 */
function extractDatesFromFilename(str: string): Date[] {
  const dates: Date[] = [];

  // YYYYMMDD  e.g. 20250309
  for (const m of str.matchAll(/\b(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\b/g)) {
    dates.push(new Date(+m[1], +m[2] - 1, +m[3]));
  }

  // YYYY-MM-DD or YYYY.MM.DD  e.g. 2025-03-09
  for (const m of str.matchAll(/\b(20\d{2})[-./](0[1-9]|1[0-2])[-./](0[1-9]|[12]\d|3[01])\b/g)) {
    dates.push(new Date(+m[1], +m[2] - 1, +m[3]));
  }

  // DD-MM-YYYY or DD.MM.YYYY  e.g. 09.03.2025
  for (const m of str.matchAll(/\b(0[1-9]|[12]\d|3[01])[-./](0[1-9]|1[0-2])[-./](20\d{2})\b/g)) {
    dates.push(new Date(+m[3], +m[2] - 1, +m[1]));
  }

  // Deduplicate by timestamp
  const seen = new Set<number>();
  return dates.filter(d => {
    const t = d.getTime();
    if (seen.has(t)) return false;
    seen.add(t);
    return true;
  });
}

export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: 'Could not parse form data' }, { status: 400 });
  }

  const file       = formData.get('file') as File | null;
  const clientName = (formData.get('clientName') as string | null)?.trim() ?? '';

  if (!file || !clientName) {
    return NextResponse.json({ error: 'Missing file or client name' }, { status: 400 });
  }

  try {
    const buffer   = Buffer.from(await file.arrayBuffer());
    const warnings: string[] = [];

    // ── 1. Parse dirty file (needed for date columns) ───────────────────────
    let parsed: ReturnType<typeof parseDirtyFile> | null = null;
    try {
      parsed = parseDirtyFile(buffer);
    } catch {
      // If we can't parse the dirty file the process route will give a proper error.
      return NextResponse.json({ warnings: [] });
    }

    // ── 2. Check B6 of Parameter Selection for client name ──────────────────
    const wb          = XLSX.read(buffer, { type: 'buffer' });
    const paramName   = wb.SheetNames.find(n => n.trim().toLowerCase() === 'parameter selection');

    if (!paramName) {
      warnings.push(
        `No "Parameter Selection" sheet found — cannot verify the client name.`
      );
    } else {
      const ws     = wb.Sheets[paramName];
      const cell   = ws['B6'];
      const b6Val  = cell ? String(cell.v ?? cell.w ?? '').trim() : '';

      if (!b6Val) {
        warnings.push(
          `Cell B6 of the Parameter Selection sheet is empty — cannot verify the client name.`
        );
      } else if (!b6Val.toLowerCase().includes(clientName.toLowerCase())) {
        warnings.push(
          `Client name "${clientName}" was not found in cell B6 of the Parameter Selection sheet` +
          ` (B6 contains: "${b6Val}"). You may have the wrong file or client name.`
        );
      }
    }

    // ── 3. Check date in filename against data date columns ─────────────────
    const filenameBase    = file.name.replace(/\.[^.]+$/, '');
    const datesInFilename = extractDatesFromFilename(filenameBase);

    if (datesInFilename.length > 0 && parsed.dateColumns.length > 0) {
      // Build a set of timestamps from the data's date columns (DD.MM.YYYY units)
      const dataTimestamps = new Set<number>(
        parsed.dateColumns.flatMap(col => {
          const parts = col.split(' ')[0].split('.');
          if (parts.length !== 3) return [];
          const [d, m, y] = parts.map(Number);
          return [new Date(y, m - 1, d).getTime()];
        })
      );

      const anyMatch = datesInFilename.some(d => dataTimestamps.has(d.getTime()));

      if (!anyMatch) {
        const filenameDatesStr = datesInFilename
          .map(d => d.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }))
          .join(', ');
        const first = parsed.dateColumns[0].split(' ')[0];
        const last  = parsed.dateColumns[parsed.dateColumns.length - 1].split(' ')[0];
        warnings.push(
          `Date(s) found in the filename (${filenameDatesStr}) do not appear in the data` +
          ` (data covers ${first} – ${last}). The file may be from the wrong period.`
        );
      }
    }

    return NextResponse.json({ warnings });

  } catch (err) {
    console.error('[/api/validate]', err);
    // Validation errors should never block processing — return empty warnings
    return NextResponse.json({ warnings: [] });
  }
}
