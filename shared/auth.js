/* USI People — Auth & Graph API Infrastructure */

const USI = (() => {
  // ── Config ──
  let _config = {
    clientId: '345c6c12-5f4a-45a6-8e15-191d9051fad8',
    tenantId: '53e2c3bd-d5d4-443d-bd75-e93b8d6a32c1',
    scopes: 'https://graph.microsoft.com/User.Read https://graph.microsoft.com/Files.Read.All https://graph.microsoft.com/Sites.Read.All https://graph.microsoft.com/Sites.ReadWrite.All',
    redirectUri: window.location.href.split('?')[0].split('#')[0],
    sharePointHost: 'universalsystemsinc100.sharepoint.com',
    sharePointSite: '/sites/HR',
    managerEmails: []
  };

  let _siteId = null;
  let _me = null;
  let _listIds = {};

  // ── Helpers ──
  function b64url(buf) {
    const bytes = buf instanceof ArrayBuffer ? new Uint8Array(buf) : buf;
    let s = '';
    bytes.forEach(b => s += String.fromCharCode(b));
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  // ── PKCE OAuth2 ──
  async function pkceChallenge() {
    const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
    const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
    return { verifier, challenge: b64url(digest) };
  }

  async function login() {
    const { verifier, challenge } = await pkceChallenge();
    sessionStorage.setItem('pkce_verifier', verifier);
    const params = new URLSearchParams({
      client_id: _config.clientId,
      response_type: 'code',
      redirect_uri: _config.redirectUri,
      scope: _config.scopes,
      code_challenge: challenge,
      code_challenge_method: 'S256',
      response_mode: 'query'
    });
    window.location = `https://login.microsoftonline.com/${_config.tenantId}/oauth2/v2.0/authorize?${params}`;
  }

  async function handleRedirect() {
    const code = new URLSearchParams(window.location.search).get('code');
    if (!code) { console.log('[USI Auth] No code in URL'); return false; }
    const verifier = sessionStorage.getItem('pkce_verifier');
    if (!verifier) { console.error('[USI Auth] No PKCE verifier found in sessionStorage'); return false; }
    console.log('[USI Auth] Exchanging code for token, redirect_uri:', _config.redirectUri);
    try {
      const body = new URLSearchParams({
        client_id: _config.clientId,
        code,
        redirect_uri: _config.redirectUri,
        grant_type: 'authorization_code',
        code_verifier: verifier,
        scope: _config.scopes
      });
      const res = await fetch(`https://login.microsoftonline.com/${_config.tenantId}/oauth2/v2.0/token`, {
        method: 'POST', body,
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
      });
      const data = await res.json();
      if (data.access_token) {
        console.log('[USI Auth] Token obtained successfully');
        sessionStorage.setItem('usi_token', data.access_token);
        sessionStorage.setItem('usi_token_exp', String(Date.now() + (data.expires_in - 60) * 1000));
        sessionStorage.removeItem('pkce_verifier');
        window.history.replaceState({}, '', _config.redirectUri);
        return true;
      }
      console.error('[USI Auth] Token exchange error:', JSON.stringify(data));
    } catch (e) {
      console.error('[USI Auth] Token exchange failed:', e);
    }
    return false;
  }

  function getToken() {
    const token = sessionStorage.getItem('usi_token');
    const exp = parseInt(sessionStorage.getItem('usi_token_exp') || '0');
    if (token && Date.now() < exp) return token;
    return null;
  }

  function logout() {
    sessionStorage.clear();
    window.location = `https://login.microsoftonline.com/${_config.tenantId}/oauth2/v2.0/logout?post_logout_redirect_uri=${encodeURIComponent(_config.redirectUri)}`;
  }

  // ── Microsoft Graph API ──
  async function graph(method, path, body) {
    const token = getToken();
    if (!token) { login(); throw new Error('Redirecting to login...'); }
    const opts = {
      method,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };
    if (body) opts.body = JSON.stringify(body);
    const r = await fetch(`https://graph.microsoft.com/v1.0${path}`, opts);
    if (r.status === 401) { login(); throw new Error('Session expired'); }
    if (r.status === 429) {
      const retryAfter = parseInt(r.headers.get('Retry-After') || '5');
      await new Promise(ok => setTimeout(ok, retryAfter * 1000));
      return graph(method, path, body);
    }
    if (r.status === 204) return null;
    if (!r.ok) {
      const t = await r.text();
      throw new Error(`Graph API (${r.status}) ${method} ${path}: ${t}`);
    }
    return r.json();
  }

  async function getMe() {
    if (_me) return _me;
    const cached = sessionStorage.getItem('usi_me');
    if (cached) { _me = JSON.parse(cached); return _me; }
    _me = await graph('GET', '/me');
    sessionStorage.setItem('usi_me', JSON.stringify(_me));
    return _me;
  }

  function isManager(email) {
    const e = (email || '').toLowerCase();
    return _config.managerEmails.some(m => m.toLowerCase() === e);
  }

  async function sendMail(to, cc, subject, htmlBody) {
    const message = {
      subject,
      body: { contentType: 'HTML', content: htmlBody },
      toRecipients: (Array.isArray(to) ? to : [to]).map(a => ({ emailAddress: { address: a } }))
    };
    if (cc) {
      message.ccRecipients = (Array.isArray(cc) ? cc : [cc]).map(a => ({ emailAddress: { address: a } }));
    }
    return graph('POST', '/me/sendMail', { message, saveToSentItems: true });
  }

  // ── SharePoint Helpers ──
  async function getSiteId() {
    if (_siteId) return _siteId;
    const cached = sessionStorage.getItem('usi_siteId');
    if (cached) { _siteId = cached; return _siteId; }
    const site = await graph('GET', `/sites/${_config.sharePointHost}:${_config.sharePointSite}`);
    _siteId = site.id;
    sessionStorage.setItem('usi_siteId', _siteId);
    return _siteId;
  }

  let _allListsCache = null;
  async function _fetchAllLists() {
    if (_allListsCache) return _allListsCache;
    const siteId = await getSiteId();
    const result = await graph('GET', `/sites/${siteId}/lists?$select=id,displayName`);
    _allListsCache = result.value || [];
    // Cache all list IDs at once
    _allListsCache.forEach(l => {
      _listIds[l.displayName] = l.id;
      sessionStorage.setItem(`usi_list_${l.displayName}`, l.id);
    });
    console.log(`[USI] Loaded ${_allListsCache.length} lists from SharePoint`);
    return _allListsCache;
  }

  async function ensureList(listName, columns) {
    const cacheKey = `usi_list_${listName}`;
    const cached = sessionStorage.getItem(cacheKey);
    if (cached) { _listIds[listName] = cached; return cached; }

    // Fetch all lists once and find by display name
    try {
      const allLists = await _fetchAllLists();
      const match = allLists.find(l => l.displayName === listName);
      if (match) return match.id;
    } catch (e) { console.warn(`Could not query lists:`, e.message); }

    // List not found
    console.error(`[USI] List "${listName}" not found in SharePoint HR site`);
    throw new Error(`List "${listName}" not found. Please create it manually in SharePoint at the HR site.`);
  }

  // ── JSON Storage Layer ──
  // SharePoint lists have only Title (text) + Data (multi-line text) columns.
  // All custom fields are serialized as JSON in the Data column.
  // Title is used for a human-readable identifier.

  function _packFields(fields) {
    // Store a short identifier in Title, everything in Data as JSON
    const title = fields.Title || fields.EmployeeName || fields.EmployeeEmail || 'Item';
    return { Title: String(title).substring(0, 255), Data: JSON.stringify(fields) };
  }

  function _unpackItem(item) {
    // Merge the stored JSON data back into fields
    const raw = item.fields || {};
    let data = {};
    try { data = raw.Data ? JSON.parse(raw.Data) : {}; } catch (e) { /* not JSON */ }
    return { id: item.id, fields: { ...data, _spTitle: raw.Title }, ...data, _itemId: item.id };
  }

  async function getListItems(listName) {
    const siteId = await getSiteId();
    const listId = _listIds[listName] || await ensureList(listName, []);
    const url = `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;
    console.log(`[USI] Loading items from "${listName}"...`);
    try {
      let allItems = [];
      let response = await graph('GET', url);
      allItems = allItems.concat(response.value || []);
      while (response['@odata.nextLink']) {
        const nextUrl = response['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '');
        response = await graph('GET', nextUrl);
        allItems = allItems.concat(response.value || []);
      }
      // Unpack JSON data from each item
      const unpacked = allItems.map(_unpackItem);
      console.log(`[USI] Loaded ${unpacked.length} items from "${listName}"`);
      return unpacked;
    } catch (e) {
      console.error(`[USI] Failed to load items from "${listName}":`, e.message);
      return [];
    }
  }

  async function createListItem(listName, fields) {
    const siteId = await getSiteId();
    const listId = _listIds[listName] || await ensureList(listName, []);
    const packed = _packFields(fields);
    const result = await graph('POST', `/sites/${siteId}/lists/${listId}/items`, { fields: packed });
    return _unpackItem(result);
  }

  async function updateListItem(listName, itemId, fields) {
    const siteId = await getSiteId();
    const listId = _listIds[listName] || await ensureList(listName, []);
    // Fetch current item to merge data
    const current = await graph('GET', `/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`);
    let existing = {};
    try { existing = current.fields.Data ? JSON.parse(current.fields.Data) : {}; } catch (e) {}
    const merged = { ...existing, ...fields };
    const packed = _packFields(merged);
    return graph('PATCH', `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, packed);
  }

  async function deleteListItem(listName, itemId) {
    const siteId = await getSiteId();
    const listId = _listIds[listName] || await ensureList(listName, []);
    return graph('DELETE', `/sites/${siteId}/lists/${listId}/items/${itemId}`);
  }

  // ── Column Definition Helpers ──
  function textCol(name) { return { name, text: {} }; }
  function multilineCol(name) { return { name, text: { allowMultipleLines: true, textType: 'plain' } }; }
  function numberCol(name) { return { name, number: {} }; }
  function dateCol(name) { return { name, dateTime: { format: 'dateOnly' } }; }
  function dateTimeCol(name) { return { name, dateTime: { format: 'dateTime' } }; }
  function choiceCol(name, choices) { return { name, choice: { choices } }; }
  function boolCol(name) { return { name, boolean: {} }; }

  // ── Init ──
  function init(config) {
    Object.assign(_config, config);
    _config.redirectUri = config.redirectUri || window.location.href.split('?')[0].split('#')[0];
  }

  // ── Public API ──
  return {
    init, login, handleRedirect, getToken, logout,
    graph, getMe, isManager, sendMail,
    getSiteId, ensureList, getListItems, createListItem, updateListItem, deleteListItem,
    textCol, multilineCol, numberCol, dateCol, dateTimeCol, choiceCol, boolCol,
    get config() { return _config; }
  };
})();
