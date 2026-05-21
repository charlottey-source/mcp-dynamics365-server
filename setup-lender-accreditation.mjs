/**
 * Setup script: Lender Accreditation Issues entity for Dynamics 365
 * Creates entity, fields, relationship to Contact, then publishes.
 * Run once: node setup-lender-accreditation.mjs
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
const ATTR_ENDPOINT = `api/data/v9.2/EntityDefinitions(LogicalName='${ENTITY}')/Attributes`;

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

async function api(token, endpoint, method = 'GET', body) {
  const res = await fetch(`${BASE_URL}${endpoint}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'Content-Type': 'application/json; charset=utf-8',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`${method} ${endpoint} → ${res.status}: ${text}`);
  }

  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

function label(text) {
  return { LocalizedLabels: [{ Label: text, LanguageCode: 1033 }] };
}

async function step(name, fn) {
  process.stdout.write(`  ${name}... `);
  try {
    await fn();
    console.log('done');
  } catch (e) {
    if (e.message.includes('already exists') || e.message.includes('0x80044230') || e.message.includes('0x80060892')) {
      console.log('already exists, skipping');
    } else {
      console.log('FAILED');
      throw e;
    }
  }
}

async function main() {
  console.log('\nAuthenticating...');
  const token = await getToken();
  console.log('Authenticated\n');

  // ── 1. Create entity ────────────────────────────────────────────────────────
  console.log('1. Creating entity');
  await step('nc_LenderAccreditationIssue', () =>
    api(token, 'api/data/v9.2/EntityDefinitions', 'POST', {
      '@odata.type': 'Microsoft.Dynamics.CRM.EntityMetadata',
      SchemaName: 'nc_LenderAccreditationIssue',
      LogicalName: ENTITY,
      DisplayName: label('Lender Accreditation Issue'),
      DisplayCollectionName: label('Lender Accreditation Issues'),
      Description: label('Tracks lender accreditation issues for broker contacts'),
      OwnershipType: 'UserOwned',
      IsActivity: false,
      HasActivities: false,
      HasNotes: false,
      PrimaryNameAttribute: 'nc_name',
      Attributes: [
        {
          '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
          SchemaName: 'nc_Name',
          LogicalName: 'nc_name',
          IsPrimaryName: true,
          DisplayName: label('Name'),
          RequiredLevel: { Value: 'None' },
          MaxLength: 200,
          Format: 'Text',
        },
      ],
    })
  );

  // Brief pause for entity to register
  await new Promise(r => setTimeout(r, 4000));

  // ── 2. Create fields ────────────────────────────────────────────────────────
  console.log('\n2. Creating fields');

  await step('nc_lender (Lender option set)', () =>
    api(token, ATTR_ENDPOINT, 'POST', {
      '@odata.type': 'Microsoft.Dynamics.CRM.PicklistAttributeMetadata',
      SchemaName: 'nc_Lender',
      LogicalName: 'nc_lender',
      DisplayName: label('Lender'),
      RequiredLevel: { Value: 'ApplicationRequired' },
      OptionSet: {
        '@odata.type': 'Microsoft.Dynamics.CRM.OptionSetMetadata',
        IsGlobal: false,
        OptionSetType: 'Picklist',
        Options: [
          { Value: 100000000, Label: label('ANZ') },
          { Value: 100000001, Label: label('CBA') },
          { Value: 100000002, Label: label('NAB') },
          { Value: 100000003, Label: label('Westpac') },
          { Value: 100000004, Label: label('AMP') },
          { Value: 100000005, Label: label('Macquarie') },
        ],
      },
    })
  );

  await step('nc_dateceased (Date Ceased)', () =>
    api(token, ATTR_ENDPOINT, 'POST', {
      '@odata.type': 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
      SchemaName: 'nc_DateCeased',
      LogicalName: 'nc_dateceased',
      DisplayName: label('Date Ceased'),
      RequiredLevel: { Value: 'None' },
      Format: 'DateOnly',
      DateTimeBehavior: { Value: 'DateOnly' },
    })
  );

  await step('nc_reason (Reason)', () =>
    api(token, ATTR_ENDPOINT, 'POST', {
      '@odata.type': 'Microsoft.Dynamics.CRM.MemoAttributeMetadata',
      SchemaName: 'nc_Reason',
      LogicalName: 'nc_reason',
      DisplayName: label('Reason'),
      RequiredLevel: { Value: 'None' },
      MaxLength: 2000,
      Format: 'TextArea',
    })
  );

  await step('nc_reaccredited (Re-accredited checkbox)', () =>
    api(token, ATTR_ENDPOINT, 'POST', {
      '@odata.type': 'Microsoft.Dynamics.CRM.BooleanAttributeMetadata',
      SchemaName: 'nc_ReAccredited',
      LogicalName: 'nc_reaccredited',
      DisplayName: label('Re-accredited'),
      RequiredLevel: { Value: 'None' },
      OptionSet: {
        '@odata.type': 'Microsoft.Dynamics.CRM.BooleanOptionSetMetadata',
        TrueOption: { Value: 1, Label: label('Yes') },
        FalseOption: { Value: 0, Label: label('No') },
      },
    })
  );

  await step('nc_datereaccredited (Date Re-accredited)', () =>
    api(token, ATTR_ENDPOINT, 'POST', {
      '@odata.type': 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
      SchemaName: 'nc_DateReAccredited',
      LogicalName: 'nc_datereaccredited',
      DisplayName: label('Date Re-accredited'),
      RequiredLevel: { Value: 'None' },
      Format: 'DateOnly',
      DateTimeBehavior: { Value: 'DateOnly' },
    })
  );

  // ── 3. Create 1:N relationship Contact → Issue ───────────────────────────
  console.log('\n3. Creating relationship (Contact → Lender Accreditation Issues)');
  await step('nc_contact_nc_lenderaccreditationissue', () =>
    api(token, 'api/data/v9.2/RelationshipDefinitions', 'POST', {
      '@odata.type': 'Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata',
      SchemaName: 'nc_contact_nc_lenderaccreditationissue',
      ReferencedEntity: 'contact',
      ReferencingEntity: ENTITY,
      ReferencedAttribute: 'contactid',
      Lookup: {
        '@odata.type': 'Microsoft.Dynamics.CRM.LookupAttributeMetadata',
        SchemaName: 'nc_ContactId',
        LogicalName: 'nc_contactid',
        DisplayName: label('Contact'),
        RequiredLevel: { Value: 'ApplicationRequired' },
        Description: label('The broker contact this accreditation issue relates to'),
      },
      AssociatedMenuConfiguration: {
        Behavior: 'UseLabel',
        Group: 'Details',
        Label: label('Lender Accreditation Issues'),
        Order: 10000,
        IsCustomizable: true,
      },
      CascadeConfiguration: {
        Assign: 'NoCascade',
        Delete: 'Cascade',
        Merge: 'Cascade',
        Reparent: 'NoCascade',
        Share: 'NoCascade',
        Unshare: 'NoCascade',
      },
    })
  );

  // ── 4. Publish ──────────────────────────────────────────────────────────────
  console.log('\n4. Publishing customizations (may take 30-60s)...');
  await api(token, 'api/data/v9.2/PublishAllXml', 'POST', {});
  console.log('  Published\n');

  console.log('='.repeat(60));
  console.log('Setup complete. Manual steps remaining in Dynamics UI:');
  console.log('');
  console.log('A. Add subgrid to Contact form:');
  console.log('   Settings > Customizations > Customize the System');
  console.log('   > Entities > Contact > Forms > Main form');
  console.log('   > Compliance tab > after Education section');
  console.log('   > Insert > Sub-Grid');
  console.log('   > Entity: Lender Accreditation Issues');
  console.log('   > Default View: Active Lender Accreditation Issues');
  console.log('   > Label: Lender Accreditation Issues');
  console.log('');
  console.log('B. Add Business Rule on nc_lenderaccreditationissue form:');
  console.log('   > When nc_reaccredited = No → hide nc_datereaccredited');
  console.log('   > When nc_reaccredited = Yes → show nc_datereaccredited');
  console.log('='.repeat(60));
}

main().catch(e => {
  console.error('\nFailed:', e.message);
  process.exit(1);
});
