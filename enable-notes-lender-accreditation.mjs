/**
 * Enables notes (annotations) on nc_LenderAccreditationIssue entity.
 * Run once: node enable-notes-lender-accreditation.mjs
 */

import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, '.env') });

const { CLIENT_ID, CLIENT_SECRET, TENANT_ID, D365_URL } = process.env;
const BASE_URL = D365_URL.endsWith('/') ? D365_URL : `${D365_URL}/`;
const ENTITY = 'nc_lenderaccreditationissue';

async function getToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: `${D365_URL}/.default`,
        grant_type: 'client_credentials',
      }).toString(),
    }
  );
  const data = await res.json();
  if (!res.ok) throw new Error(`Auth failed: ${JSON.stringify(data)}`);
  return data.access_token;
}

async function main() {
  console.log('Authenticating...');
  const token = await getToken();
  console.log('Authenticated\n');

  console.log(`Enabling notes on ${ENTITY}...`);

  const res = await fetch(`${BASE_URL}api/data/v9.2/EntityDefinitions(LogicalName='${ENTITY}')`, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'Content-Type': 'application/json; charset=utf-8',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      'MSCRM.MergeLabels': 'true',
    },
    body: JSON.stringify({ HasNotes: true }),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Failed: ${res.status} ${text}`);
  }

  console.log('Notes enabled.\n');

  console.log('Publishing...');
  const pub = await fetch(`${BASE_URL}api/data/v9.2/PublishAllXml`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'Content-Type': 'application/json; charset=utf-8',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
    },
    body: JSON.stringify({}),
  });

  if (!pub.ok) {
    const text = await pub.text();
    throw new Error(`Publish failed: ${pub.status} ${text}`);
  }

  console.log('Published.\n');
  console.log('='.repeat(50));
  console.log('Done. One manual step remaining:');
  console.log('');
  console.log('Add Timeline to the form:');
  console.log('  Power Apps > Tables > Lender Accreditation Issue');
  console.log('  > Forms > Main form > + Component > Timeline');
  console.log('  > Drop onto canvas > Save and publish');
  console.log('='.repeat(50));
}

main().catch(e => {
  console.error('\nFailed:', e.message);
  process.exit(1);
});
