const params = new URLSearchParams(location.search);
const statusEl = document.getElementById('status');
const folderInput = document.getElementById('folderInput');
const folderPathEl = document.getElementById('folderPath');
const clientLabel = document.getElementById('clientLabel');
const tableHead = document.querySelector('#dataTable thead');
const tableBody = document.querySelector('#dataTable tbody');
const tableMeta = document.getElementById('tableMeta');
const rowDetail = document.getElementById('rowDetail');
const graphMeta = document.getElementById('graphMeta');
const accessMeta = document.getElementById('accessMeta');

const state = {
  path: params.get('path') || '',
  client: params.get('client') || '',
  sessionId: params.get('session') || '',
  demo: params.get('demo') === '1',
  source: 'MessageTrace',
  offset: 0,
  limit: 50,
  total: 0,
  sortBy: '',
  ascending: false,
  columns: [],
  rows: [],
  sources: [],
  packs: [],
  network: null,
  accessNetwork: null,
};

const DEMO = {
  path: 'C:\\Reports\\Contoso\\Client1_Ticket1839334',
  sources: [
    { name: 'MessageTrace', present: true, rowCount: 13, columns: ['Received', 'SenderAddress', 'RecipientAddress', 'Subject', 'Status', 'FromIP', 'Size'], timeCol: 'Received', preset: [] },
    { name: 'SignInLogs', present: true, rowCount: 7, columns: ['CreatedDateTime', 'UserPrincipalName', 'IPAddress', 'AppDisplayName', 'Status', 'CountryOrRegion', 'ClientAppUsed'], timeCol: 'CreatedDateTime', preset: [] },
    { name: 'UnifiedAuditLogs', present: true, rowCount: 6, columns: ['CreationDate', 'UserIds', 'Operations', 'RecordType', 'ClientIP', 'ResultStatus', '_Operation', '_ClientIP', '_MailboxOwnerUPN'], timeCol: 'CreationDate', preset: [] },
  ],
  graph: {
    matched: 13,
    nodes: [
      { id: 'alice@contoso.com', label: 'alice@contoso.com', count: 9 },
      { id: 'mallory@fabrikam.com', label: 'mallory@fabrikam.com', count: 4 },
      { id: 'bob@contoso.com', label: 'bob@contoso.com', count: 3 },
      { id: 'carol@contoso.com', label: 'carol@contoso.com', count: 3 },
      { id: 'eve@adventure-works.com', label: 'eve@adventure-works.com', count: 2 },
      { id: 'ap@fabrikam.com', label: 'ap@fabrikam.com', count: 1 },
      { id: 'dave@tailspintoys.com', label: 'dave@tailspintoys.com', count: 1 },
      { id: 'finance@contoso.com', label: 'finance@contoso.com', count: 1 },
    ],
    edges: [
      { from: 'alice@contoso.com', to: 'mallory@fabrikam.com', count: 2 },
      { from: 'alice@contoso.com', to: 'bob@contoso.com', count: 2 },
      { from: 'alice@contoso.com', to: 'carol@contoso.com', count: 2 },
      { from: 'alice@contoso.com', to: 'ap@fabrikam.com', count: 1 },
      { from: 'alice@contoso.com', to: 'finance@contoso.com', count: 1 },
      { from: 'bob@contoso.com', to: 'alice@contoso.com', count: 1 },
      { from: 'mallory@fabrikam.com', to: 'alice@contoso.com', count: 1 },
      { from: 'eve@adventure-works.com', to: 'alice@contoso.com', count: 2 },
      { from: 'carol@contoso.com', to: 'dave@tailspintoys.com', count: 1 },
    ],
    timeline: [
      { t: '2026-01-15T09:00', count: 3 },
      { t: '2026-01-15T10:00', count: 3 },
      { t: '2026-01-15T11:00', count: 2 },
      { t: '2026-01-15T12:00', count: 0 },
      { t: '2026-01-15T13:00', count: 1 },
      { t: '2026-01-15T14:00', count: 1 },
      { t: '2026-01-15T15:00', count: 1 },
      { t: '2026-01-15T16:00', count: 0 },
      { t: '2026-01-15T17:00', count: 0 },
      { t: '2026-01-15T18:00', count: 0 },
      { t: '2026-01-15T19:00', count: 0 },
      { t: '2026-01-15T20:00', count: 0 },
      { t: '2026-01-15T21:00', count: 0 },
      { t: '2026-01-15T22:00', count: 0 },
      { t: '2026-01-15T23:00', count: 0 },
      { t: '2026-01-16T00:00', count: 0 },
      { t: '2026-01-16T01:00', count: 0 },
      { t: '2026-01-16T02:00', count: 0 },
      { t: '2026-01-16T03:00', count: 0 },
      { t: '2026-01-16T04:00', count: 0 },
      { t: '2026-01-16T05:00', count: 0 },
      { t: '2026-01-16T06:00', count: 0 },
      { t: '2026-01-16T07:00', count: 0 },
      { t: '2026-01-16T08:00', count: 2 },
    ],
  },
  tables: {
    MessageTrace: {
      columns: ['Received', 'SenderAddress', 'RecipientAddress', 'Subject', 'Status', 'FromIP'],
      rows: [
        { Received: '2026-01-15T09:05:00Z', SenderAddress: 'alice@contoso.com', RecipientAddress: 'bob@contoso.com', Subject: 'Q1 planning', Status: 'Delivered', FromIP: '203.0.113.10' },
        { Received: '2026-01-15T09:12:00Z', SenderAddress: 'alice@contoso.com', RecipientAddress: 'finance@contoso.com', Subject: 'Q1 planning', Status: 'Delivered', FromIP: '203.0.113.10' },
        { Received: '2026-01-15T09:40:00Z', SenderAddress: 'bob@contoso.com', RecipientAddress: 'alice@contoso.com', Subject: 'Re: Q1 planning', Status: 'Delivered', FromIP: '203.0.113.11' },
        { Received: '2026-01-15T10:02:00Z', SenderAddress: 'alice@contoso.com', RecipientAddress: 'mallory@fabrikam.com', Subject: 'Invoice 1044', Status: 'Delivered', FromIP: '198.51.100.44' },
        { Received: '2026-01-15T10:03:00Z', SenderAddress: 'alice@contoso.com', RecipientAddress: 'ap@fabrikam.com', Subject: 'Invoice 1044', Status: 'Delivered', FromIP: '198.51.100.44' },
        { Received: '2026-01-15T10:18:00Z', SenderAddress: 'mallory@fabrikam.com', RecipientAddress: 'alice@contoso.com', Subject: 'Re: Invoice 1044', Status: 'Delivered', FromIP: '198.51.100.44' },
        { Received: '2026-01-15T11:05:00Z', SenderAddress: 'eve@adventure-works.com', RecipientAddress: 'alice@contoso.com', Subject: 'Password reset', Status: 'Failed', FromIP: '203.0.113.88' },
        { Received: '2026-01-15T11:07:00Z', SenderAddress: 'eve@adventure-works.com', RecipientAddress: 'alice@contoso.com', Subject: 'Password reset', Status: 'Delivered', FromIP: '203.0.113.88' },
        { Received: '2026-01-15T13:22:00Z', SenderAddress: 'alice@contoso.com', RecipientAddress: 'carol@contoso.com', Subject: 'Fwd: Invoice 1044', Status: 'Delivered', FromIP: '203.0.113.10' },
        { Received: '2026-01-15T14:10:00Z', SenderAddress: 'carol@contoso.com', RecipientAddress: 'dave@tailspintoys.com', Subject: 'Vendor onboarding', Status: 'Delivered', FromIP: '203.0.113.12' },
        { Received: '2026-01-15T15:45:00Z', SenderAddress: 'alice@contoso.com', RecipientAddress: 'bob@contoso.com;carol@contoso.com', Subject: 'Standup notes', Status: 'Delivered', FromIP: '203.0.113.10' },
        { Received: '2026-01-16T08:15:00Z', SenderAddress: 'alice@contoso.com', RecipientAddress: 'mallory@fabrikam.com', Subject: 'Wire instructions', Status: 'Delivered', FromIP: '198.51.100.44' },
        { Received: '2026-01-16T08:40:00Z', SenderAddress: 'olivia.t@example.org', RecipientAddress: 'alice@contoso.com', Subject: 'Re: Wire instructions', Status: 'Delivered', FromIP: '192.0.2.55' },
      ],
    },
    SignInLogs: {
      columns: ['CreatedDateTime', 'UserPrincipalName', 'IPAddress', 'AppDisplayName', 'Status', 'CountryOrRegion', 'ClientAppUsed'],
      rows: [
        { CreatedDateTime: '2026-01-15T08:55:00Z', UserPrincipalName: 'alice@contoso.com', IPAddress: '203.0.113.10', AppDisplayName: 'Office 365', Status: 'Success', CountryOrRegion: 'US', ClientAppUsed: 'Browser' },
        { CreatedDateTime: '2026-01-15T09:01:00Z', UserPrincipalName: 'bob@contoso.com', IPAddress: '203.0.113.11', AppDisplayName: 'Outlook Mobile', Status: 'Success', CountryOrRegion: 'US', ClientAppUsed: 'Mobile Apps and Desktop clients' },
        { CreatedDateTime: '2026-01-15T10:01:00Z', UserPrincipalName: 'alice@contoso.com', IPAddress: '198.51.100.44', AppDisplayName: 'Office 365', Status: 'Success', CountryOrRegion: 'DE', ClientAppUsed: 'Browser' },
        { CreatedDateTime: '2026-01-15T10:04:00Z', UserPrincipalName: 'alice@contoso.com', IPAddress: '198.51.100.44', AppDisplayName: 'Azure Portal', Status: 'Success', CountryOrRegion: 'DE', ClientAppUsed: 'Browser' },
        { CreatedDateTime: '2026-01-15T11:02:00Z', UserPrincipalName: 'eve@adventure-works.com', IPAddress: '203.0.113.88', AppDisplayName: 'Office 365', Status: 'Failure', CountryOrRegion: 'RU', ClientAppUsed: 'Browser' },
        { CreatedDateTime: '2026-01-16T08:12:00Z', UserPrincipalName: 'alice@contoso.com', IPAddress: '192.0.2.55', AppDisplayName: 'Office 365', Status: 'Success', CountryOrRegion: 'NG', ClientAppUsed: 'Browser' },
        { CreatedDateTime: '2026-01-16T08:13:00Z', UserPrincipalName: 'alice@contoso.com', IPAddress: '192.0.2.55', AppDisplayName: 'Microsoft Authenticator', Status: 'Failure', CountryOrRegion: 'NG', ClientAppUsed: 'Mobile Apps and Desktop clients' },
      ],
    },
    UnifiedAuditLogs: {
      columns: ['CreationDate', 'UserIds', 'Operations', 'RecordType', 'ClientIP', 'ResultStatus', '_MailboxOwnerUPN'],
      rows: [
        { CreationDate: '2026-01-15T09:10:00Z', UserIds: 'alice@contoso.com', Operations: 'MailItemsAccessed', RecordType: 'ExchangeItem', ClientIP: '203.0.113.10', ResultStatus: 'Success', _MailboxOwnerUPN: 'alice@contoso.com' },
        { CreationDate: '2026-01-15T10:06:00Z', UserIds: 'alice@contoso.com', Operations: 'MailItemsAccessed', RecordType: 'ExchangeItem', ClientIP: '198.51.100.44', ResultStatus: 'Success', _MailboxOwnerUPN: 'alice@contoso.com' },
        { CreationDate: '2026-01-15T10:08:00Z', UserIds: 'alice@contoso.com', Operations: 'UpdateInboxRules', RecordType: 'ExchangeAdmin', ClientIP: '198.51.100.44', ResultStatus: 'Success', _MailboxOwnerUPN: 'alice@contoso.com' },
        { CreationDate: '2026-01-15T10:09:00Z', UserIds: 'alice@contoso.com', Operations: 'Set-Mailbox', RecordType: 'ExchangeAdmin', ClientIP: '198.51.100.44', ResultStatus: 'Success', _MailboxOwnerUPN: 'alice@contoso.com' },
        { CreationDate: '2026-01-15T13:30:00Z', UserIds: 'carol@contoso.com', Operations: 'MailItemsAccessed', RecordType: 'ExchangeItem', ClientIP: '203.0.113.12', ResultStatus: 'Success', _MailboxOwnerUPN: 'carol@contoso.com' },
        { CreationDate: '2026-01-16T08:20:00Z', UserIds: 'alice@contoso.com', Operations: 'UserLoggedIn', RecordType: 'AzureActiveDirectory', ClientIP: '192.0.2.55', ResultStatus: 'Success', _MailboxOwnerUPN: '' },
      ],
    },
  },
};

if (state.demo) {
  state.path = DEMO.path;
  state.client = state.client || 'demo';
}
if (state.client) clientLabel.textContent = state.demo ? 'Contoso demo (sample data)' : `Client ${state.client}`;
if (state.path) {
  folderInput.value = state.path;
  folderPathEl.textContent = state.demo ? `${state.path}  —  synthetic Contoso pack` : state.path;
}

function matchDemoRow(row, filters) {
  return (filters || []).every((f) => {
    const col = f.column === '_address'
      ? `${row.SenderAddress || ''} ${row.RecipientAddress || ''} ${row.UserPrincipalName || ''} ${row.UserIds || ''}`
      : (f.column === '_ip'
        ? `${row.FromIP || ''} ${row.IPAddress || ''} ${row.ClientIP || ''}`
        : (f.column === '_geo'
          ? `${row.CountryOrRegion || ''}`
          : (f.column === '_time'
            ? (row.Received || row.CreatedDateTime || row.CreationDate || '')
            : String(row[f.column] ?? ''))));
    const val = String(f.value ?? '');
    if (f.op === 'icontains') return col.toLowerCase().includes(val.toLowerCase());
    if (f.op === 'gte') return col >= val;
    if (f.op === 'lte') return col <= val;
    return true;
  });
}

function demoHourKey(received) {
  const d = new Date(received);
  if (Number.isNaN(d.getTime())) return '';
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}T${pad(d.getUTCHours())}:00`;
}

function demoDomain(addr) {
  const at = String(addr || '').lastIndexOf('@');
  return at < 0 ? '' : String(addr).slice(at + 1).toLowerCase();
}

function buildDemoMailGraph(filters, opts = {}) {
  const hideInternal = Boolean(opts.hideInternal);
  const minCount = Math.max(1, Number(opts.minCount) || 1);
  const rows = DEMO.tables.MessageTrace.rows.filter((row) => matchDemoRow(row, filters));
  const nodeCounts = {};
  const edgeCounts = {};
  const buckets = {};
  rows.forEach((row) => {
    const sender = String(row.SenderAddress || '').toLowerCase();
    const recips = String(row.RecipientAddress || '').split(/[;,\s]+/).filter((s) => s.includes('@')).map((s) => s.toLowerCase());
    const hk = demoHourKey(row.Received);
    if (hk) buckets[hk] = (buckets[hk] || 0) + 1;
    if (!sender) return;
    nodeCounts[sender] = (nodeCounts[sender] || 0) + 1;
    recips.forEach((to) => {
      if (hideInternal && demoDomain(sender) && demoDomain(sender) === demoDomain(to)) return;
      nodeCounts[to] = (nodeCounts[to] || 0) + 1;
      const ek = `${sender}|${to}`;
      edgeCounts[ek] = (edgeCounts[ek] || 0) + 1;
    });
  });
  const edges = Object.entries(edgeCounts)
    .filter(([, count]) => count >= minCount)
    .map(([key, count]) => {
      const [from, to] = key.split('|');
      return { from, to, count };
    });
  const nodes = Object.entries(nodeCounts)
    .map(([id, count]) => ({ id, label: id, count }))
    .sort((a, b) => b.count - a.count);
  const timeline = Object.entries(buckets)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([t, count]) => ({ t, count }));
  return { matched: rows.length, nodes, edges, timeline };
}

function isPrivateDemoIp(ip) {
  const p = String(ip || '');
  return p.startsWith('10.') || p.startsWith('192.168.') || p.startsWith('127.') || /^172\.(1[6-9]|2\d|3[0-1])\./.test(p);
}

function buildDemoAccessGraph(filters, opts = {}) {
  const includeSignIn = opts.includeSignIn !== false;
  const includeUal = opts.includeUal !== false;
  const hidePrivate = opts.hidePrivateIp !== false;
  const minCount = Math.max(1, Number(opts.minCount) || 1);
  const ipCountry = {};
  const events = [];
  if (includeSignIn) {
    DEMO.tables.SignInLogs.rows.filter((row) => matchDemoRow(row, filters.filter((f) => f.column !== '_geo'))).forEach((row) => {
      const ip = String(row.IPAddress || '');
      if (row.CountryOrRegion) ipCountry[ip] = row.CountryOrRegion;
      events.push({
        user: String(row.UserPrincipalName || '').toLowerCase(),
        ip,
        country: row.CountryOrRegion || '',
        failed: /fail/i.test(String(row.Status || '')),
        t: row.CreatedDateTime,
        source: 'signin',
      });
    });
  }
  if (includeUal) {
    DEMO.tables.UnifiedAuditLogs.rows.filter((row) => matchDemoRow(row, filters.filter((f) => f.column !== '_geo'))).forEach((row) => {
      events.push({
        user: String(row.UserIds || '').toLowerCase(),
        ip: String(row.ClientIP || ''),
        country: '',
        failed: /fail/i.test(String(row.ResultStatus || '')),
        t: row.CreationDate,
        source: 'ual',
      });
    });
  }
  const geoFilters = filters.filter((f) => f.column === '_geo');
  const edgeCounts = {};
  const edgeFails = {};
  const userCounts = {};
  const userFails = {};
  const ipCounts = {};
  const ipFails = {};
  const ipSources = {};
  const countryCounts = {};
  const countryFails = {};
  const buckets = {};
  const bucketFails = {};
  let signInMatched = 0;
  let ualMatched = 0;
  events.forEach((ev) => {
    if (!ev.user || !ev.ip) return;
    if (hidePrivate && isPrivateDemoIp(ev.ip)) return;
    let country = ev.country || ipCountry[ev.ip] || 'Unknown';
    if (ev.country) ipCountry[ev.ip] = ev.country;
    if (geoFilters.some((f) => !String(country).toLowerCase().includes(String(f.value || '').toLowerCase()))) return;
    if (ev.source === 'signin') signInMatched += 1;
    else ualMatched += 1;
    const hk = demoHourKey(ev.t);
    if (hk) {
      buckets[hk] = (buckets[hk] || 0) + 1;
      if (ev.failed) bucketFails[hk] = (bucketFails[hk] || 0) + 1;
    }
    userCounts[ev.user] = (userCounts[ev.user] || 0) + 1;
    if (ev.failed) userFails[ev.user] = (userFails[ev.user] || 0) + 1;
    ipCounts[ev.ip] = (ipCounts[ev.ip] || 0) + 1;
    if (ev.failed) ipFails[ev.ip] = (ipFails[ev.ip] || 0) + 1;
    ipSources[ev.ip] = ipSources[ev.ip] || {};
    ipSources[ev.ip][ev.source] = true;
    const ek = `${ev.user}|${ev.ip}`;
    edgeCounts[ek] = (edgeCounts[ek] || 0) + 1;
    if (ev.failed) edgeFails[ek] = (edgeFails[ek] || 0) + 1;
    countryCounts[country] = (countryCounts[country] || 0) + 1;
    if (ev.failed) countryFails[country] = (countryFails[country] || 0) + 1;
  });
  const edges = Object.entries(edgeCounts)
    .filter(([, count]) => count >= minCount)
    .map(([key, count]) => {
      const [from, to] = key.split('|');
      return { from, to, count, failCount: edgeFails[key] || 0, country: ipCountry[to] || 'Unknown' };
    });
  const nodes = [
    ...Object.keys(userCounts).map((id) => ({
      id: `u:${id}`, label: id, kind: 'user', count: userCounts[id], failCount: userFails[id] || 0, country: '',
    })),
    ...Object.keys(ipCounts).map((id) => ({
      id: `i:${id}`, label: id, kind: 'ip', count: ipCounts[id], failCount: ipFails[id] || 0, country: ipCountry[id] || 'Unknown',
      sources: Object.keys(ipSources[id] || {}),
    })),
  ];
  const countries = Object.entries(countryCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([country, count]) => ({ country, count, failCount: countryFails[country] || 0, users: 0, ips: 0 }));
  const timeline = Object.entries(buckets)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([t, count]) => ({ t, count, failCount: bucketFails[t] || 0 }));
  return { matched: signInMatched + ualMatched, signInMatched, ualMatched, nodes, edges, countries, timeline };
}

function demoApi(path, options) {
  const route = String(path || '').split('?')[0];
  let body = {};
  try { body = options.body ? JSON.parse(options.body) : {}; } catch { body = {}; }
  const filters = Array.isArray(body.filters) ? body.filters : [];
  if (route === '/api/investigate/packs') {
    return {
      ok: true,
      result: {
        packs: [{
          path: DEMO.path,
          name: 'Client1_Ticket1839334',
          company: 'Contoso',
          sources: ['MessageTrace', 'SignInLogs', 'UnifiedAuditLogs'],
          exists: true,
          curated: false,
          writtenAt: '2026-01-16T08:40:00Z',
          organization: 'Contoso',
          ticket: '1839334',
          origin: 'demo',
        }],
        roots: [],
        count: 1,
      },
    };
  }
  if (route === '/api/investigate/sources') {
    return { ok: true, result: { path: DEMO.path, sources: DEMO.sources } };
  }
  if (route === '/api/investigate/mail-graph') {
    return {
      ok: true,
      result: {
        path: DEMO.path,
        ...buildDemoMailGraph(filters, body),
      },
    };
  }
  if (route === '/api/investigate/access-graph') {
    return {
      ok: true,
      result: {
        path: DEMO.path,
        ...buildDemoAccessGraph(filters, body),
      },
    };
  }
  if (route === '/api/investigate/query') {
    const table = DEMO.tables[body.source] || { columns: [], rows: [] };
    let rows = table.rows.filter((row) => matchDemoRow(row, filters));
    if (body.sortBy) {
      const key = body.sortBy;
      rows = rows.slice().sort((a, b) => String(a[key] || '').localeCompare(String(b[key] || '')));
      if (body.ascending === false) rows.reverse();
    }
    const offset = Number(body.offset) || 0;
    const limit = Number(body.limit) || 50;
    return {
      ok: true,
      result: {
        path: DEMO.path,
        source: body.source,
        total: rows.length,
        offset,
        limit,
        columns: table.columns,
        rows: rows.slice(offset, offset + limit),
      },
    };
  }
  if (route === '/api/investigate/export') {
    return { ok: true, result: { file: '(demo mode — export skipped)', rowCount: 0 } };
  }
  throw new Error(`Demo has no handler for ${route}`);
}

async function api(path, options = {}, timeoutMs = 180000) {
  if (state.demo) return demoApi(path, options);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(path, {
      headers: { 'Content-Type': 'application/json', ...(options.headers || {}) },
      ...options,
      signal: controller.signal,
    });
    const text = await res.text();
    let body = text;
    try { body = text ? JSON.parse(text) : null; } catch { /* plain */ }
    if (!res.ok) {
      const err = typeof body === 'object' && body?.error ? body.error : text;
      throw new Error(err || res.statusText);
    }
    return body;
  } catch (e) {
    if (e.name === 'AbortError') throw new Error(`Request timed out after ${timeoutMs / 1000}s`);
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

function setStatus(msg) {
  statusEl.textContent = msg || '';
}

function toLocalInput(isoOrKey) {
  if (!isoOrKey) return '';
  const raw = String(isoOrKey);
  const d = new Date(raw.includes('T') && raw.length === 16 ? `${raw}:00Z` : raw);
  if (Number.isNaN(d.getTime())) return '';
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function fromLocalInput(value, endOfBucket) {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return '';
  if (endOfBucket) d.setMinutes(d.getMinutes() + 59, 59, 999);
  return d.toISOString();
}

function currentFilters(mode = 'table') {
  const filters = [];
  const start = document.getElementById('timeStart').value;
  const end = document.getElementById('timeEnd').value;
  const address = document.getElementById('addressQ').value.trim();
  const ip = document.getElementById('ipQ').value.trim();
  const status = document.getElementById('statusQ').value.trim();
  if (start) filters.push({ column: '_time', op: 'gte', value: fromLocalInput(start, false) });
  if (end) filters.push({ column: '_time', op: 'lte', value: fromLocalInput(end, true) });
  if (address) filters.push({ column: '_address', op: 'icontains', value: address });
  if (ip) filters.push({ column: '_ip', op: 'icontains', value: ip });
  const geo = document.getElementById('geoQ').value.trim();
  if (geo) filters.push({ column: '_geo', op: 'icontains', value: geo });
  if (status) {
    if (mode === 'mail') {
      filters.push({ column: 'Status', op: 'icontains', value: status });
    } else if (mode === 'access') {
      if (state.source === 'SignInLogs') {
        filters.push({ column: 'Status', op: 'icontains', value: status });
      } else if (state.source === 'UnifiedAuditLogs') {
        filters.push({ column: '_Operation', op: 'icontains', value: status });
      }
    } else {
      const col = state.source === 'MessageTrace' ? 'Status'
        : (state.source === 'SignInLogs' ? 'Status' : 'Operations');
      filters.push({ column: col, op: 'icontains', value: status });
      if (state.source === 'UnifiedAuditLogs') {
        filters[filters.length - 1] = { column: '_Operation', op: 'icontains', value: status };
      }
    }
  }
  return filters;
}

function payload(extra = {}) {
  const mode = extra.filterMode || 'table';
  const rest = { ...extra };
  delete rest.filterMode;
  return JSON.stringify({
    path: state.path,
    filters: currentFilters(mode),
    ...rest,
  });
}

function present(name) {
  return state.sources.some((s) => s.name === name && s.present);
}

async function loadSources() {
  if (!state.path) return;
  setStatus('Loading folder inventory…');
  const data = await api('/api/investigate/sources', { method: 'POST', body: payload() });
  state.sources = data.result?.sources || [];
  const bits = state.sources.map((s) => s.present ? `${s.name} ${s.rowCount}` : `${s.name} (missing)`);
  setStatus(bits.join(' · '));
  document.querySelectorAll('.tabBtn').forEach((btn) => {
    btn.disabled = !present(btn.dataset.source);
  });
  if (!present(state.source)) {
    const first = state.sources.find((s) => s.present);
    if (first) state.source = first.name;
  }
  document.querySelectorAll('.tabBtn').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.source === state.source);
  });
}

function renderTable(result) {
  state.total = result.total || 0;
  state.offset = result.offset || 0;
  state.columns = result.columns || [];
  state.rows = result.rows || [];
  tableHead.innerHTML = `<tr>${state.columns.map((c) => {
    const ind = state.sortBy === c ? (state.ascending ? ' ▲' : ' ▼') : '';
    return `<th class="sortable" data-col="${c}">${c}${ind}</th>`;
  }).join('')}</tr>`;
  tableBody.innerHTML = state.rows.map((row, i) => {
    const cells = state.columns.map((c) => {
      const v = row[c];
      const text = v == null ? '' : String(v);
      return `<td>${escapeHtml(text)}</td>`;
    }).join('');
    return `<tr data-idx="${i}">${cells}</tr>`;
  }).join('');
  const start = state.total ? state.offset + 1 : 0;
  const end = Math.min(state.offset + state.rows.length, state.total);
  tableMeta.textContent = `${state.source}: ${start}-${end} of ${state.total}`;
  document.getElementById('prevBtn').disabled = state.offset <= 0;
  document.getElementById('nextBtn').disabled = state.offset + state.rows.length >= state.total;
  rowDetail.classList.remove('show');
  rowDetail.textContent = '';
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

async function loadTable() {
  if (!state.path || !present(state.source)) {
    tableMeta.textContent = `No ${state.source} CSV in this folder.`;
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    return;
  }
  tableMeta.textContent = 'Loading…';
  const data = await api('/api/investigate/query', {
    method: 'POST',
    body: payload({
      source: state.source,
      sortBy: state.sortBy,
      ascending: state.ascending,
      limit: state.limit,
      offset: state.offset,
    }),
  });
  renderTable(data.result || {});
}

function renderNetwork(graph) {
  const host = document.getElementById('network');
  const nodes = (graph.nodes || []).map((n) => ({
    id: n.id,
    label: n.label,
    value: Math.max(1, n.count || 1),
    title: `${n.label}\n${n.count} messages`,
  }));
  const edges = (graph.edges || []).map((e, i) => ({
    id: `e${i}`,
    from: e.from,
    to: e.to,
    value: e.count,
    title: `${e.from} → ${e.to} (${e.count})`,
    arrows: 'to',
  }));
  graphMeta.textContent = `${graph.matched || 0} filtered messages · ${nodes.length} nodes · ${edges.length} edges`;

  if (typeof vis === 'undefined' || !vis.Network) {
    host.innerHTML = `<pre style="padding:0.5rem;white-space:pre-wrap">${escapeHtml(
      edges.slice(0, 40).map((e) => `${e.from} -> ${e.to} (${e.value})`).join('\n') || 'No edges (or vis-network CDN blocked).'
    )}</pre>`;
    return;
  }
  const data = { nodes: new vis.DataSet(nodes), edges: new vis.DataSet(edges) };
  const options = {
    nodes: { shape: 'dot', scaling: { min: 8, max: 36 }, font: { size: 12 } },
    edges: { color: { color: '#8c959f' }, smooth: { type: 'continuous' } },
    physics: { enabled: true, stabilization: { iterations: 40, fit: true } },
    interaction: { hover: true, tooltipDelay: 80 },
  };
  if (state.network) {
    try { state.network.destroy(); } catch { /* already gone */ }
    state.network = null;
    host.innerHTML = '';
  }
  state.network = new vis.Network(host, data, options);
  state.network.once('stabilized', () => {
    try { state.network.setOptions({ physics: false }); } catch { /* ignore */ }
  });
  state.network.on('click', (ev) => {
    if (!ev.nodes.length) return;
    document.getElementById('addressQ').value = ev.nodes[0];
    state.offset = 0;
    setTimeout(() => {
      refreshVisualsAndTable().catch((e) => setStatus(e.message));
    }, 0);
  });
}

function fillTimelineGaps(points, bucket) {
  const src = (points || []).filter((p) => p && p.t);
  if (src.length < 2) return src.slice();
  const parseKey = (t) => {
    const raw = String(t);
    return new Date(raw.includes('T') && raw.length <= 16 ? `${raw}:00Z` : raw);
  };
  const fmtKey = (d) => {
    const pad = (n) => String(n).padStart(2, '0');
    if (bucket === 'day') return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}T${pad(d.getUTCHours())}:00`;
  };
  const map = new Map();
  src.forEach((p) => map.set(p.t, Number(p.count) || 0));
  const times = src.map((p) => parseKey(p.t)).filter((d) => !Number.isNaN(d.getTime())).sort((a, b) => a - b);
  if (times.length < 2) return src.slice();
  const stepMs = bucket === 'day' ? 86400000 : 3600000;
  if (!stepMs) return src.slice();
  const maxSteps = bucket === 'day' ? 90 : 72;
  if ((times[times.length - 1].getTime() - times[0].getTime()) / stepMs > maxSteps) return src.slice();
  const out = [];
  for (let t = times[0].getTime(), n = 0; t <= times[times.length - 1].getTime() && n < maxSteps + 1; t += stepMs, n++) {
    const key = fmtKey(new Date(t));
    out.push({ t: key, count: map.has(key) ? map.get(key) : 0 });
  }
  return out;
}

function formatTick(t) {
  const raw = String(t || '');
  const d = new Date(raw.includes('T') && raw.length <= 16 ? `${raw}:00Z` : raw);
  if (Number.isNaN(d.getTime())) return raw;
  const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getUTCMonth()];
  if (!raw.includes('T')) return `${d.getUTCDate()} ${mon}`;
  const hh = String(d.getUTCHours()).padStart(2, '0');
  return `${d.getUTCDate()} ${mon} ${hh}:00`;
}

function renderTimeline(graph, hostId = 'timeline', unit = 'message') {
  const host = document.getElementById(hostId);
  if (!host) return;
  const bucketEl = hostId === 'accessTimeline' ? document.getElementById('accessBucket') : document.getElementById('bucket');
  const bucket = bucketEl?.value || 'hour';
  const series = fillTimelineGaps(graph.timeline || [], bucket);
  const max = Math.max(1, ...series.map((x) => Number(x.count) || 0));
  host.replaceChildren();
  series.forEach((p) => {
    const col = document.createElement('div');
    col.className = 'timelineCol';
    const bar = document.createElement('button');
    bar.type = 'button';
    bar.className = 'timelineBar' + (p.count ? '' : ' zero');
    bar.style.height = `${Math.max(4, Math.round((p.count / max) * 120))}px`;
    const fail = Number(p.failCount) || 0;
    const failBit = fail ? `, ${fail} failed` : '';
    bar.title = `${formatTick(p.t)}: ${p.count} ${unit}${p.count === 1 ? '' : 's'}${failBit}`;
    bar.disabled = !p.count;
    bar.addEventListener('click', () => {
      const start = toLocalInput(p.t);
      document.getElementById('timeStart').value = start;
      const d = new Date(start);
      if (bucket === 'day') d.setDate(d.getDate() + 1);
      else d.setHours(d.getHours() + 1);
      document.getElementById('timeEnd').value = toLocalInput(d.toISOString());
      state.offset = 0;
      refreshVisualsAndTable().catch((e) => setStatus(e.message));
    });
    const lbl = document.createElement('div');
    lbl.className = 'timelineLbl';
    lbl.textContent = formatTick(p.t);
    col.append(bar, lbl);
    host.append(col);
  });
}

function geoColor(country) {
  const palette = ['#0969da', '#1a7f37', '#9a6700', '#8250df', '#0550ae', '#6639ba'];
  const s = String(country || 'Unknown');
  let h = 0;
  for (let i = 0; i < s.length; i += 1) h = (h * 31 + s.charCodeAt(i)) >>> 0;
  return palette[h % palette.length];
}

function renderCountries(graph) {
  const host = document.getElementById('countryBars');
  if (!host) return;
  const rows = graph.countries || [];
  const max = Math.max(1, ...rows.map((c) => Number(c.count) || 0));
  host.replaceChildren();
  rows.forEach((c) => {
    const ok = Math.max(0, (Number(c.count) || 0) - (Number(c.failCount) || 0));
    const fail = Number(c.failCount) || 0;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'countryRow';
    btn.title = `${c.country}: ${c.count} events (${fail} failed)`;
    const code = document.createElement('span');
    code.className = 'countryCode';
    code.textContent = c.country;
    const track = document.createElement('span');
    track.className = 'countryTrack';
    if (ok) {
      const okEl = document.createElement('span');
      okEl.className = 'countryOk';
      okEl.style.width = `${(ok / max) * 100}%`;
      track.append(okEl);
    }
    if (fail) {
      const failEl = document.createElement('span');
      failEl.className = 'countryFail';
      failEl.style.width = `${(fail / max) * 100}%`;
      track.append(failEl);
    }
    const meta = document.createElement('span');
    meta.className = 'countryMeta';
    meta.textContent = fail ? `${c.count} (${fail} fail)` : String(c.count);
    btn.append(code, track, meta);
    btn.addEventListener('click', () => {
      document.getElementById('geoQ').value = c.country === 'Unknown' ? '' : c.country;
      state.offset = 0;
      if (present('SignInLogs')) {
        state.source = 'SignInLogs';
        document.querySelectorAll('.tabBtn').forEach((b) => b.classList.toggle('active', b.dataset.source === 'SignInLogs'));
      }
      refreshVisualsAndTable().catch((e) => setStatus(e.message));
    });
    host.append(btn);
  });
}

function renderAccessNetwork(graph) {
  const host = document.getElementById('accessNetwork');
  const nodes = (graph.nodes || []).map((n) => {
    const fail = Number(n.failCount) || 0;
    const isUser = n.kind === 'user';
    const country = n.country || '';
    const title = isUser
      ? `${n.label}\n${n.count} events${fail ? `\n${fail} failed` : ''}`
      : `${n.label}\n${country || 'Unknown'}\n${n.count} events${fail ? `\n${fail} failed` : ''}`;
    return {
      id: n.id,
      label: isUser ? n.label : `${n.label}${country && country !== 'Unknown' ? `\n${country}` : ''}`,
      value: Math.max(1, n.count || 1),
      title,
      group: isUser ? 'user' : (fail ? 'ipFail' : 'ip'),
      color: isUser ? undefined : { background: geoColor(country), border: fail ? '#cf222e' : geoColor(country) },
    };
  });
  const edges = (graph.edges || []).map((e, i) => ({
    id: `a${i}`,
    from: `u:${e.from}`,
    to: `i:${e.to}`,
    value: e.count,
    title: `${e.from} → ${e.to} (${e.count}${e.failCount ? `, ${e.failCount} failed` : ''})`,
    color: { color: e.failCount ? '#cf222e' : '#8c959f' },
  }));
  accessMeta.textContent = `${graph.matched || 0} access events · ${graph.signInMatched || 0} sign-in · ${graph.ualMatched || 0} UAL · ${nodes.length} nodes · ${(graph.countries || []).length} countries`;
  if (!graph.matched && document.getElementById('hidePrivateIp')?.checked) {
    accessMeta.textContent += ' — 0 public IPs. Uncheck Hide private IPs if this pack only has RFC1918 addresses.';
  }

  if (typeof vis === 'undefined' || !vis.Network) {
    host.innerHTML = `<pre style="padding:0.5rem;white-space:pre-wrap">${escapeHtml(
      edges.slice(0, 40).map((e) => `${e.from} -> ${e.to} (${e.value})`).join('\n') || 'No edges (or vis-network CDN blocked).'
    )}</pre>`;
    return;
  }
  const data = { nodes: new vis.DataSet(nodes), edges: new vis.DataSet(edges) };
  const options = {
    nodes: { scaling: { min: 8, max: 36 }, font: { size: 12 } },
    groups: {
      user: { shape: 'box', color: { background: '#ddf4ff', border: '#0969da' } },
      ip: { shape: 'dot' },
      ipFail: { shape: 'dot' },
    },
    edges: { smooth: { type: 'continuous' } },
    physics: { enabled: true, stabilization: { iterations: 40, fit: true } },
    interaction: { hover: true, tooltipDelay: 80 },
  };
  if (state.accessNetwork) {
    try { state.accessNetwork.destroy(); } catch { /* already gone */ }
    state.accessNetwork = null;
    host.innerHTML = '';
  }
  state.accessNetwork = new vis.Network(host, data, options);
  state.accessNetwork.once('stabilized', () => {
    try { state.accessNetwork.setOptions({ physics: false }); } catch { /* ignore */ }
  });
  state.accessNetwork.on('click', (ev) => {
    if (!ev.nodes.length) return;
    const id = String(ev.nodes[0]);
    if (id.startsWith('u:')) document.getElementById('addressQ').value = id.slice(2);
    else if (id.startsWith('i:')) document.getElementById('ipQ').value = id.slice(2);
    state.offset = 0;
    setTimeout(() => {
      refreshVisualsAndTable().catch((e) => setStatus(e.message));
    }, 0);
  });
}

async function loadAccessGraph() {
  if (!state.path || (!present('SignInLogs') && !present('UnifiedAuditLogs'))) {
    accessMeta.textContent = 'No Sign-in or UAL CSV in this folder.';
    return;
  }
  accessMeta.textContent = 'Building login graph…';
  try {
    const data = await api('/api/investigate/access-graph', {
      method: 'POST',
      body: payload({
        filterMode: 'access',
        minCount: parseInt(document.getElementById('accessMinCount').value, 10) || 1,
        maxNodes: parseInt(document.getElementById('accessMaxNodes').value, 10) || 80,
        includeSignIn: document.getElementById('includeSignIn').checked,
        includeUal: document.getElementById('includeUal').checked,
        hidePrivateIp: document.getElementById('hidePrivateIp').checked,
        bucket: document.getElementById('accessBucket').value,
      }),
    });
    const graph = data.result || {};
    renderCountries(graph);
    renderAccessNetwork(graph);
    renderTimeline(graph, 'accessTimeline', 'event');
  } catch (e) {
    const msg = String(e.message || e);
    accessMeta.textContent = /not found|unknown investigate/i.test(msg)
      ? 'Restart the web runner to enable login activity (the UI loaded, but this process is still on the old API).'
      : msg;
    throw e;
  }
}

async function loadGraph() {
  if (!state.path || !present('MessageTrace')) {
    graphMeta.textContent = 'No MessageTrace CSV in this folder.';
    return;
  }
  graphMeta.textContent = 'Building graph…';
  const data = await api('/api/investigate/mail-graph', {
    method: 'POST',
    body: payload({
      filterMode: 'mail',
      minCount: parseInt(document.getElementById('minCount').value, 10) || 1,
      maxNodes: parseInt(document.getElementById('maxNodes').value, 10) || 80,
      hideInternal: document.getElementById('hideInternal').checked,
      bucket: document.getElementById('bucket').value,
    }),
  });
  const graph = data.result || {};
  renderNetwork(graph);
  renderTimeline(graph);
}

async function refreshVisualsAndTable() {
  const results = await Promise.allSettled([loadGraph(), loadAccessGraph(), loadTable()]);
  const failed = results.find((r) => r.status === 'rejected');
  if (failed) {
    const msg = failed.reason?.message || String(failed.reason || '');
    if (!/not found|unknown investigate/i.test(msg)) setStatus(msg);
  }
}

async function refreshAll() {
  if (!state.path) {
    setStatus('Choose a report pack, or paste a folder path.');
    return;
  }
  try {
    await loadSources();
    await refreshVisualsAndTable();
  } catch (e) {
    setStatus(e.message);
  }
}

function applyFolderPath(path, extra = {}) {
  state.path = String(path || '').trim();
  folderInput.value = state.path;
  folderPathEl.textContent = state.path;
  const q = new URLSearchParams();
  if (state.path) q.set('path', state.path);
  if (state.client) q.set('client', state.client);
  if (state.sessionId) q.set('session', state.sessionId);
  if (state.demo) q.set('demo', '1');
  history.replaceState(null, '', `investigate.html?${q}`);
  state.offset = 0;
  if (!extra.skipRefresh) refreshAll();
}

function formatPackLabel(p) {
  const bits = [];
  if (p.company) bits.push(p.company);
  bits.push(p.name || p.path);
  if (p.ticket) bits.push(`#${p.ticket}`);
  if (p.organization && p.organization !== p.company) bits.push(p.organization);
  if (p.curated) bits.push('curated');
  if (p.sources?.length) {
    bits.push(p.sources.map((s) => (s === 'MessageTrace' ? 'MT' : (s === 'SignInLogs' ? 'SI' : 'UAL'))).join('+'));
  }
  if (!p.exists) bits.push('missing');
  return bits.join(' · ');
}

function renderPackSelect(packs) {
  const sel = document.getElementById('packSelect');
  if (!sel) return;
  const current = state.path;
  const rows = packs || [];
  if (!rows.length) {
    sel.innerHTML = '<option value="">No report packs found</option>';
    return;
  }
  sel.innerHTML = rows.map((p) => {
    const value = String(p.path || '').replace(/"/g, '&quot;');
    const disabled = p.exists === false ? ' disabled' : '';
    const selected = current && String(p.path).toLowerCase() === current.toLowerCase() ? ' selected' : '';
    return `<option value="${value}"${disabled}${selected}>${formatPackLabel(p)}</option>`;
  }).join('');
  if (current && ![...sel.options].some((o) => o.value.toLowerCase() === current.toLowerCase())) {
    const opt = document.createElement('option');
    opt.value = current;
    opt.textContent = current;
    opt.selected = true;
    sel.append(opt);
  }
}

async function loadPackList(opts = {}) {
  const sel = document.getElementById('packSelect');
  const q = (document.getElementById('packSearch')?.value || '').trim();
  const params = new URLSearchParams();
  if (q) params.set('q', q);
  if (state.sessionId) params.set('session', state.sessionId);
  try {
    const data = await api(`/api/investigate/packs?${params}`);
    state.packs = data.result?.packs || [];
    renderPackSelect(state.packs);
    if (opts.autoOpen && !state.path) {
      const first = state.packs.find((p) => p.exists && p.path);
      if (first) applyFolderPath(first.path);
    }
  } catch (e) {
    if (sel) sel.innerHTML = '<option value="">Could not list packs — paste a path below</option>';
    if (!state.path) setStatus(`Pack list unavailable (${e.message}). Restart the web runner, or paste a folder path.`);
  }
}

document.getElementById('reloadBtn').addEventListener('click', () => {
  applyFolderPath(folderInput.value.trim());
});
document.getElementById('openPackBtn')?.addEventListener('click', () => {
  const sel = document.getElementById('packSelect');
  applyFolderPath(sel?.value || '');
});
document.getElementById('packSelect')?.addEventListener('change', () => {
  const sel = document.getElementById('packSelect');
  if (sel?.value) applyFolderPath(sel.value);
});
let packSearchTimer = null;
document.getElementById('packSearch')?.addEventListener('input', () => {
  clearTimeout(packSearchTimer);
  packSearchTimer = setTimeout(() => loadPackList().catch((e) => setStatus(e.message)), 250);
});

document.getElementById('applyBtn').addEventListener('click', () => {
  state.offset = 0;
  refreshAll();
});
document.getElementById('clearBtn').addEventListener('click', () => {
  document.getElementById('timeStart').value = '';
  document.getElementById('timeEnd').value = '';
  document.getElementById('addressQ').value = '';
  document.getElementById('ipQ').value = '';
  document.getElementById('geoQ').value = '';
  document.getElementById('statusQ').value = '';
  state.offset = 0;
  refreshAll();
});
document.getElementById('graphBtn').addEventListener('click', () => loadGraph().catch((e) => setStatus(e.message)));
document.getElementById('accessBtn').addEventListener('click', () => loadAccessGraph().catch((e) => setStatus(e.message)));

document.querySelectorAll('.tabBtn').forEach((btn) => {
  btn.addEventListener('click', () => {
    state.source = btn.dataset.source;
    state.offset = 0;
    state.sortBy = '';
    document.querySelectorAll('.tabBtn').forEach((b) => b.classList.toggle('active', b === btn));
    loadTable().catch((e) => setStatus(e.message));
  });
});

document.getElementById('prevBtn').addEventListener('click', () => {
  state.offset = Math.max(0, state.offset - state.limit);
  loadTable().catch((e) => setStatus(e.message));
});
document.getElementById('nextBtn').addEventListener('click', () => {
  state.offset += state.limit;
  loadTable().catch((e) => setStatus(e.message));
});

tableHead.addEventListener('click', (e) => {
  const th = e.target.closest('th[data-col]');
  if (!th) return;
  const col = th.dataset.col;
  if (state.sortBy === col) state.ascending = !state.ascending;
  else {
    state.sortBy = col;
    state.ascending = col.toLowerCase().includes('date') || col.toLowerCase().includes('time') || col === 'Received';
    if (state.ascending && (col.toLowerCase().includes('date') || col === 'Received' || col === 'CreatedDateTime' || col === 'CreationDate')) {
      state.ascending = false;
    }
  }
  state.offset = 0;
  loadTable().catch((err) => setStatus(err.message));
});

tableBody.addEventListener('click', (e) => {
  const tr = e.target.closest('tr[data-idx]');
  if (!tr) return;
  tableBody.querySelectorAll('tr').forEach((r) => r.classList.remove('selected'));
  tr.classList.add('selected');
  const row = state.rows[Number(tr.dataset.idx)];
  rowDetail.textContent = JSON.stringify(row, null, 2);
  rowDetail.classList.add('show');
});

document.getElementById('exportBtn').addEventListener('click', async () => {
  try {
    setStatus('Exporting filtered CSV…');
    const data = await api('/api/investigate/export', {
      method: 'POST',
      body: payload({
        source: state.source,
        sortBy: state.sortBy,
        ascending: state.ascending,
      }),
    });
    setStatus(`Wrote ${data.result?.rowCount || 0} rows to ${data.result?.file || ''}`);
  } catch (e) {
    setStatus(e.message);
  }
});

if (state.path || state.demo) refreshAll();
loadPackList({ autoOpen: Boolean(state.sessionId) && !state.path }).catch((e) => setStatus(e.message));
