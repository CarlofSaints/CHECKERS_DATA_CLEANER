import { NextRequest, NextResponse } from 'next/server';
import { parseDirtyFile } from '@/lib/excel-reader';
import { buildCleanFile } from '@/lib/excel-builder';
import { uploadToSharePoint } from '@/lib/graph-oj';

export const maxDuration = 60;

function padDate(n: number) { return String(n).padStart(2, '0'); }

function buildFilename(clientName: string, latestDate: Date): string {
  const y = latestDate.getFullYear();
  const m = padDate(latestDate.getMonth() + 1);
  const d = padDate(latestDate.getDate());
  const safe = clientName.trim().toUpperCase().replace(/[^A-Z0-9 _-]/g, '').trim();
  return `CHECKERS B2B ${safe} ${y}-${m}-${d}.xlsx`;
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

  if (!file)       return NextResponse.json({ error: 'No file provided' },         { status: 400 });
  if (!clientName) return NextResponse.json({ error: 'Client name is required' },  { status: 400 });

  try {
    // 1. Read dirty file
    const buffer = Buffer.from(await file.arrayBuffer());
    const parsed = parseDirtyFile(buffer);

    if (parsed.dateColumns.length === 0) {
      return NextResponse.json({ error: 'No weekly date columns found in VENDOR VIEW sheet' }, { status: 422 });
    }

    // 2. Build clean file
    const cleanBuffer = await buildCleanFile(parsed, clientName);

    // 3. Generate filename
    const filename = buildFilename(clientName, parsed.latestDate);

    // 4. Upload to SharePoint
    const webUrl = await uploadToSharePoint(filename, cleanBuffer);

    return NextResponse.json({
      success:   true,
      filename,
      rows:      parsed.rows.length,
      dateRange: `${parsed.dateColumns[0]} → ${parsed.dateColumns[parsed.dateColumns.length - 1]}`,
      webUrl,
    });

  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error('[/api/process]', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
