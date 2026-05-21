import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, '.env') });

const { CLIENT_ID, CLIENT_SECRET, TENANT_ID, D365_URL } = process.env;

const token = await (async () => {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: `${D365_URL}/.default`, grant_type: 'client_credentials' }).toString(),
  });
  const d = await res.json();
  return d.access_token;
})();

const base = D365_URL.endsWith('/') ? D365_URL : `${D365_URL}/`;
const url = `${base}api/data/v9.2/systemforms?$filter=objecttypecode eq 'contact' and type eq 2&$select=name,formid,isdefault,description&$orderby=isdefault desc`;

const res = await fetch(url, {
  headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
});
const data = await res.json();
data.value.forEach(f => console.log(`${f.isdefault ? '[DEFAULT]' : '         '} ${f.name} — ${f.formid}`));
