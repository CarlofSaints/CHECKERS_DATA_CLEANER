/**
 * graph-oj.ts
 * OuterJoin tenant — SharePoint file upload via Microsoft Graph API.
 */

const TENANT_ID   = process.env.OJ_TENANT_ID!;
const CLIENT_ID   = process.env.OJ_CLIENT_ID!;
const CLIENT_SECRET = process.env.OJ_CLIENT_SECRET!;
const SP_HOST     = (process.env.OJ_SP_HOST ?? 'exceler8xl.sharepoint.com').trim();
const LIBRARY     = (process.env.OJ_SP_LIBRARY ?? 'Clients').trim();
const UPLOAD_PATH = (process.env.OJ_SP_CLEANER_PATH ?? 'WAHL/04_Operations/Projects/SRC data cleaner/CLEANED FILES').trim();

async function getToken(): Promise<string> {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type:    'client_credentials',
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope:         'https://graph.microsoft.com/.default',
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) {
    throw new Error(`OJ auth failed: ${data.error_description ?? JSON.stringify(data)}`);
  }
  return data.access_token as string;
}

function encodePath(path: string): string {
  return path.split('/').map(seg => encodeURIComponent(seg)).join('/');
}

async function getDriveId(token: string): Promise<string> {
  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SP_HOST}:/`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!siteRes.ok) throw new Error(`OJ: could not get site: ${await siteRes.text()}`);
  const site = await siteRes.json();

  const drivesRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!drivesRes.ok) throw new Error(`OJ: could not list drives: ${await drivesRes.text()}`);
  const drives = await drivesRes.json();

  const drive = drives.value?.find((d: { name: string }) => d.name === LIBRARY);
  if (!drive) {
    const names = drives.value?.map((d: { name: string }) => d.name).join(', ');
    throw new Error(`OJ: library "${LIBRARY}" not found. Available: ${names}`);
  }
  return drive.id as string;
}

export async function uploadToSharePoint(filename: string, content: Buffer): Promise<string> {
  const token   = await getToken();
  const driveId = await getDriveId(token);
  const filePath = encodePath(`${UPLOAD_PATH}/${filename}`);

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: content as unknown as BodyInit,
    }
  );

  if (!res.ok) throw new Error(`OJ: upload failed (${res.status}): ${await res.text()}`);
  const item = await res.json();
  return (item.webUrl as string) ?? '';
}
