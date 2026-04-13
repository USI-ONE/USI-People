const https = require('https');

// ══════════════════════════════════════════════════════════
// ██  ORG CHART (mirrored from shared/org.js)
// ══════════════════════════════════════════════════════════

const ORG_MEMBERS = [
  { email: 'sammy@usicomputer.com', name: 'Sammy Wong', managerEmail: null, role: 'executive' },
  { email: 'gary@usicomputer.com', name: 'Gary Gonzalez', managerEmail: 'sammy@usicomputer.com', role: 'executive' },
  { email: 'johnathank@usicomputer.com', name: 'Johnathan Kendrick', managerEmail: 'sammy@usicomputer.com', role: 'executive' },
  { email: 'RachelW@usicomputer.com', name: 'Rachel Wong', managerEmail: 'sammy@usicomputer.com', role: 'executive' },
  { email: 'rachel@usicomputer.com', name: 'Rachel Connett', managerEmail: 'gary@usicomputer.com', role: 'executive' },
  { email: 'CWall@usicomputer.com', name: 'Chris Wall', managerEmail: 'johnathank@usicomputer.com', role: 'executive' },
  { email: 'Benson@usicomputer.com', name: 'Benson Chiu', managerEmail: 'RachelW@usicomputer.com', role: 'executive' },
  { email: 'Aaron@usicomputer.com', name: 'Aaron Kendrick', managerEmail: 'RachelW@usicomputer.com', role: 'executive' },
  { email: 'chriss@usicomputer.com', name: 'Chris Sudweeks', managerEmail: 'CWall@usicomputer.com', role: 'manager' },
  { email: 'Ray@usicomputer.com', name: 'Ray Leota', managerEmail: 'CWall@usicomputer.com', role: 'manager' },
  { email: 'Mitch@usicomputer.com', name: 'Mitchell Lindsey', managerEmail: 'rachel@usicomputer.com', role: 'employee' },
  { email: 'Matt.Dasher@usicomputer.com', name: 'Matt Dasher', managerEmail: 'rachel@usicomputer.com', role: 'employee' },
  { email: 'Dave@usicomputer.com', name: 'Dave Hancey', managerEmail: 'rachel@usicomputer.com', role: 'employee' },
  { email: 'mark@usicomputer.com', name: 'Mark Foy', managerEmail: 'chriss@usicomputer.com', role: 'employee' },
  { email: 'MJones@usicomputer.com', name: 'Marc Jones', managerEmail: 'chriss@usicomputer.com', role: 'employee' },
  { email: 'colton@usicomputer.com', name: 'Colton Parker', managerEmail: 'chriss@usicomputer.com', role: 'employee' },
  { email: 'Justin@usicomputer.com', name: 'Justin Silotti', managerEmail: 'Ray@usicomputer.com', role: 'employee' },
  { email: 'ethan@usicomputer.com', name: 'Ethan Sudweeks', managerEmail: 'Ray@usicomputer.com', role: 'employee' },
  { email: 'jwan@usicomputer.com', name: 'Jonathan Wan', managerEmail: 'Benson@usicomputer.com', role: 'employee' },
  { email: 'PeterSmith@usicomputer.com', name: 'Peter Smith', managerEmail: 'Benson@usicomputer.com', role: 'employee' },
  { email: 'Fernando@usicomputer.com', name: 'Fernando Cardenas', managerEmail: 'Benson@usicomputer.com', role: 'employee' },
  { email: 'lsudweeks@usicomputer.com', name: 'Lee Sudweeks', managerEmail: 'Benson@usicomputer.com', role: 'employee' },
  { email: 'Ron@usicomputer.com', name: 'Ron Dickson', managerEmail: 'Aaron@usicomputer.com', role: 'employee' },
  { email: 'Tui@usicomputer.com', name: 'Tui', managerEmail: 'Aaron@usicomputer.com', role: 'employee' },
  { email: 'lee@usicomputer.com', name: 'Lee Allred', managerEmail: 'Aaron@usicomputer.com', role: 'employee' }
];

const HR_ADMINS = ['CWall@usicomputer.com', 'sammy@usicomputer.com'];

function norm(e) { return (e || '').toLowerCase().trim(); }
function isHRAdmin(email) { return HR_ADMINS.some(a => norm(a) === norm(email)); }
function isExecutive(email) { const m = ORG_MEMBERS.find(o => norm(o.email) === norm(email)); return m && m.role === 'executive'; }
function isManager(email) { return ORG_MEMBERS.some(m => norm(m.managerEmail) === norm(email)); }

function getAllReports(managerEmail) {
  const result = [];
  const queue = [norm(managerEmail)];
  const visited = new Set();
  while (queue.length > 0) {
    const current = queue.shift();
    if (visited.has(current)) continue;
    visited.add(current);
    ORG_MEMBERS.filter(m => norm(m.managerEmail) === current).forEach(d => {
      result.push(d);
      queue.push(norm(d.email));
    });
  }
  return result;
}

function getViewableEmails(email) {
  if (isHRAdmin(email) || isExecutive(email)) return ORG_MEMBERS.map(m => m.email);
  if (isManager(email)) {
    const reports = getAllReports(email);
    return [email, ...reports.map(r => r.email)];
  }
  return [email];
}

function canViewItem(userEmail, item, emailField) {
  const viewable = new Set(getViewableEmails(userEmail).map(e => norm(e)));
  const target = item[emailField] || '';
  // If no email field on this item (e.g., Rocks), allow all
  if (!target) return true;
  return viewable.has(norm(target));
}

function canWriteItem(userEmail, item, emailField) {
  if (isHRAdmin(userEmail) || isExecutive(userEmail)) return true;
  const target = item[emailField] || '';
  if (!target) return isManager(userEmail); // No owner = manager/exec only
  if (norm(target) === norm(userEmail)) return true; // Own item
  if (isManager(userEmail)) {
    const reports = getAllReports(userEmail);
    return reports.some(r => norm(r.email) === norm(target));
  }
  return false;
}

// Lists that don't need per-item filtering (visible to all authenticated users)
const PUBLIC_LISTS = ['USI_Rocks', 'USI_OrgChart'];

// Map list names to the email field used for access control
const LIST_EMAIL_FIELD = {
  'USI_Goals': 'EmployeeEmail',
  'USI_PerformanceReviews': 'EmployeeEmail',
  'USI_ExitInterviews': 'EmployeeEmail',
  'USI_OneOnOneItems': 'EmployeeEmail',
  'USI_OneOnOneMeetings': 'EmployeeEmail',
  'USI_Onboarding': 'EmployeeEmail',
  'USI_Offboarding': 'EmployeeEmail'
};

// ══════════════════════════════════════════════════════════
// ██  APP-ONLY TOKEN ACQUISITION
// ══════════════════════════════════════════════════════════

let _appToken = null;
let _appTokenExp = 0;

async function getAppToken() {
  if (_appToken && Date.now() < _appTokenExp) return _appToken;

  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  }).toString();

  const data = await httpPost(
    `login.microsoftonline.com`,
    `/${tenantId}/oauth2/v2.0/token`,
    body,
    { 'Content-Type': 'application/x-www-form-urlencoded' }
  );

  _appToken = data.access_token;
  _appTokenExp = Date.now() + (data.expires_in - 60) * 1000;
  return _appToken;
}

// ══════════════════════════════════════════════════════════
// ██  GRAPH API HELPERS
// ══════════════════════════════════════════════════════════

function httpRequest(method, host, path, body, headers) {
  return new Promise((resolve, reject) => {
    const opts = {
      hostname: host,
      path: path,
      method: method,
      headers: { ...headers }
    };
    if (body) {
      const bodyStr = typeof body === 'string' ? body : JSON.stringify(body);
      opts.headers['Content-Length'] = Buffer.byteLength(bodyStr);
    }
    const req = https.request(opts, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode === 204) return resolve(null);
        try {
          const json = JSON.parse(data);
          if (res.statusCode >= 400) {
            reject(new Error(`Graph ${res.statusCode}: ${JSON.stringify(json.error || json)}`));
          } else {
            resolve(json);
          }
        } catch (e) {
          if (res.statusCode >= 400) reject(new Error(`Graph ${res.statusCode}: ${data}`));
          else resolve(data);
        }
      });
    });
    req.on('error', reject);
    if (body) req.write(typeof body === 'string' ? body : JSON.stringify(body));
    req.end();
  });
}

function httpPost(host, path, body, headers) {
  return httpRequest('POST', host, path, body, headers);
}

async function graphApi(method, path, body) {
  const token = await getAppToken();
  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json'
  };
  return httpRequest(method, 'graph.microsoft.com', `/v1.0${path}`, body, headers);
}

// ══════════════════════════════════════════════════════════
// ██  SHAREPOINT HELPERS
// ══════════════════════════════════════════════════════════

let _siteId = null;
let _listIds = {};

async function getSiteId() {
  if (_siteId) return _siteId;
  const host = process.env.SHAREPOINT_HOST;
  const site = process.env.SHAREPOINT_SITE;
  const result = await graphApi('GET', `/sites/${host}:${site}`);
  _siteId = result.id;
  return _siteId;
}

async function getListId(listName) {
  if (_listIds[listName]) return _listIds[listName];
  const siteId = await getSiteId();
  const result = await graphApi('GET', `/sites/${siteId}/lists?$select=id,displayName`);
  (result.value || []).forEach(l => { _listIds[l.displayName] = l.id; });
  if (!_listIds[listName]) throw new Error(`List "${listName}" not found`);
  return _listIds[listName];
}

function packFields(fields) {
  const title = fields.Title || fields.EmployeeName || fields.EmployeeEmail || 'Item';
  return { Title: String(title).substring(0, 255), Data: JSON.stringify(fields) };
}

function unpackItem(item) {
  const raw = item.fields || {};
  let data = {};
  try { data = raw.Data ? JSON.parse(raw.Data) : {}; } catch (e) {}
  return { id: item.id, fields: { ...data, _spTitle: raw.Title }, ...data, _itemId: item.id };
}

// ══════════════════════════════════════════════════════════
// ██  USER TOKEN VALIDATION
// ══════════════════════════════════════════════════════════

function extractUserEmail(req) {
  const authHeader = req.headers.authorization || req.headers.Authorization || '';
  const token = authHeader.replace('Bearer ', '');
  if (!token) return null;

  // Decode JWT payload (base64url → JSON)
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
    // Azure AD tokens use 'preferred_username', 'upn', or 'email'
    return payload.preferred_username || payload.upn || payload.email || null;
  } catch (e) {
    return null;
  }
}

// ══════════════════════════════════════════════════════════
// ██  MAIN HANDLER
// ══════════════════════════════════════════════════════════

module.exports = async function (context, req) {
  // CORS preflight
  if (req.method === 'OPTIONS') {
    context.res = {
      status: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Authorization, Content-Type'
      }
    };
    return;
  }

  // CORS headers for all responses
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };

  try {
    // Validate user
    const userEmail = extractUserEmail(req);
    if (!userEmail) {
      context.res = { status: 401, headers: corsHeaders, body: JSON.stringify({ error: 'Unauthorized' }) };
      return;
    }

    const action = context.bindingData.action;
    const body = req.body || {};

    let result;

    switch (action) {
      case 'getSiteId': {
        const siteId = await getSiteId();
        result = { siteId };
        break;
      }

      case 'getLists': {
        const siteId = await getSiteId();
        const lists = await graphApi('GET', `/sites/${siteId}/lists?$select=id,displayName`);
        result = { lists: lists.value || [] };
        break;
      }

      case 'listItems': {
        const { listName } = body;
        if (!listName) throw new Error('listName required');

        const siteId = await getSiteId();
        const listId = await getListId(listName);
        const url = `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;

        // Fetch all items with pagination
        let allItems = [];
        let response = await graphApi('GET', url);
        allItems = allItems.concat(response.value || []);
        while (response['@odata.nextLink']) {
          const nextUrl = response['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '');
          response = await graphApi('GET', nextUrl);
          allItems = allItems.concat(response.value || []);
        }

        // Unpack JSON data
        let items = allItems.map(unpackItem);

        // Apply access control filtering
        if (!PUBLIC_LISTS.includes(listName)) {
          const emailField = LIST_EMAIL_FIELD[listName] || 'EmployeeEmail';
          items = items.filter(item => canViewItem(userEmail, item, emailField));
        }

        result = { items };
        break;
      }

      case 'createItem': {
        const { listName: createList, fields } = body;
        if (!createList || !fields) throw new Error('listName and fields required');

        // Access check for writes
        if (!PUBLIC_LISTS.includes(createList)) {
          const emailField = LIST_EMAIL_FIELD[createList] || 'EmployeeEmail';
          if (!canWriteItem(userEmail, fields, emailField)) {
            context.res = { status: 403, headers: corsHeaders, body: JSON.stringify({ error: 'Access denied' }) };
            return;
          }
        }

        const siteId = await getSiteId();
        const listId = await getListId(createList);
        const packed = packFields(fields);
        const created = await graphApi('POST', `/sites/${siteId}/lists/${listId}/items`, { fields: packed });
        result = unpackItem(created);
        break;
      }

      case 'updateItem': {
        const { listName: updateList, itemId, fields: updateFields } = body;
        if (!updateList || !itemId || !updateFields) throw new Error('listName, itemId, and fields required');

        const siteId = await getSiteId();
        const listId = await getListId(updateList);

        // Fetch current item and check access
        const current = await graphApi('GET', `/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`);
        const currentUnpacked = unpackItem(current);

        if (!PUBLIC_LISTS.includes(updateList)) {
          const emailField = LIST_EMAIL_FIELD[updateList] || 'EmployeeEmail';
          if (!canWriteItem(userEmail, currentUnpacked, emailField)) {
            context.res = { status: 403, headers: corsHeaders, body: JSON.stringify({ error: 'Access denied' }) };
            return;
          }
        }

        // Merge and save
        let existing = {};
        try { existing = current.fields.Data ? JSON.parse(current.fields.Data) : {}; } catch (e) {}
        const merged = { ...existing, ...updateFields };
        const packed = packFields(merged);
        await graphApi('PATCH', `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, packed);
        result = { success: true, id: itemId };
        break;
      }

      case 'deleteItem': {
        const { listName: deleteList, itemId: deleteId } = body;
        if (!deleteList || !deleteId) throw new Error('listName and itemId required');

        const siteId = await getSiteId();
        const listId = await getListId(deleteList);

        // Fetch and check access
        const toDelete = await graphApi('GET', `/sites/${siteId}/lists/${listId}/items/${deleteId}?$expand=fields`);
        const toDeleteUnpacked = unpackItem(toDelete);

        if (!PUBLIC_LISTS.includes(deleteList)) {
          const emailField = LIST_EMAIL_FIELD[deleteList] || 'EmployeeEmail';
          if (!canWriteItem(userEmail, toDeleteUnpacked, emailField)) {
            context.res = { status: 403, headers: corsHeaders, body: JSON.stringify({ error: 'Access denied' }) };
            return;
          }
        }

        await graphApi('DELETE', `/sites/${siteId}/lists/${listId}/items/${deleteId}`);
        result = { success: true };
        break;
      }

      default:
        context.res = { status: 400, headers: corsHeaders, body: JSON.stringify({ error: `Unknown action: ${action}` }) };
        return;
    }

    context.res = { status: 200, headers: corsHeaders, body: JSON.stringify(result) };

  } catch (err) {
    context.log.error('API Error:', err.message);
    context.res = {
      status: 500,
      headers: corsHeaders,
      body: JSON.stringify({ error: err.message })
    };
  }
};
