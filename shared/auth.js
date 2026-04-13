/* USI People — Auth & API Proxy Infrastructure */

const USI = (() => {
  // ── Config ──
  let _config = {
    clientId: '345c6c12-5f4a-45a6-8e15-191d9051fad8',
    tenantId: '53e2c3bd-d5d4-443d-bd75-e93b8d6a32c1',
    scopes: 'User.Read',
    redirectUri: window.location.href.split('?')[0].split('#')[0],
    apiBase: 'https://usi-timecard-api-b5gzaqhpdte9fpcj.westus2-01.azurewebsites.net',
    managerEmails: []
  };

  let _me = null;

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

  // ── Microsoft Graph API (for user profile only) ──
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

  // ══════════════════════════════════════════════════════════
  // ██  API PROXY — all data operations go through Azure Function
  // ══════════════════════════════════════════════════════════

  async function apiCall(action, payload) {
    const token = getToken();
    if (!token) { login(); throw new Error('Redirecting to login...'); }
    const res = await fetch(`${_config.apiBase}/api/data/${action}`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload || {})
    });
    if (res.status === 401) { login(); throw new Error('Session expired'); }
    if (res.status === 403) throw new Error('Access denied');
    if (!res.ok) {
      const t = await res.text();
      throw new Error(`API ${action}: ${t}`);
    }
    return res.json();
  }

  // ── Data Operations (via API proxy) ──

  async function getSiteId() {
    // Not needed client-side anymore, but kept for API compatibility
    return 'proxy';
  }

  async function ensureList(listName, columns) {
    // Lists are managed server-side; this is now a no-op
    console.log(`[USI] List "${listName}" — managed by API proxy`);
    return listName;
  }

  async function getListItems(listName) {
    console.log(`[USI] Loading items from "${listName}" via API...`);
    try {
      const result = await apiCall('listItems', { listName });
      const items = result.items || [];
      console.log(`[USI] Loaded ${items.length} items from "${listName}"`);
      return items;
    } catch (e) {
      console.error(`[USI] Failed to load items from "${listName}":`, e.message);
      return [];
    }
  }

  // Fetch multiple lists in one API call — much faster than individual calls
  async function batchGetListItems(listNames) {
    console.log(`[USI] Batch loading ${listNames.length} lists via API...`);
    try {
      const result = await apiCall('batchListItems', { listNames });
      listNames.forEach(name => {
        console.log(`[USI] Loaded ${(result[name] || []).length} items from "${name}"`);
      });
      return result;
    } catch (e) {
      console.error('[USI] Batch load failed:', e.message);
      // Fallback: load individually
      const result = {};
      for (const name of listNames) {
        result[name] = await getListItems(name);
      }
      return result;
    }
  }

  // Pre-warm the API proxy (triggers app token acquisition + site ID lookup)
  async function warmUp() {
    try { await apiCall('getSiteId', {}); } catch(e) { /* ignore */ }
  }

  async function createListItem(listName, fields) {
    return apiCall('createItem', { listName, fields });
  }

  async function updateListItem(listName, itemId, fields) {
    return apiCall('updateItem', { listName, itemId, fields });
  }

  async function deleteListItem(listName, itemId) {
    return apiCall('deleteItem', { listName, itemId });
  }

  // ── Column Definition Helpers (kept for API compatibility, unused) ──
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
    getSiteId, ensureList, getListItems, batchGetListItems, warmUp, createListItem, updateListItem, deleteListItem,
    textCol, multilineCol, numberCol, dateCol, dateTimeCol, choiceCol, boolCol,
    get config() { return _config; }
  };
})();
