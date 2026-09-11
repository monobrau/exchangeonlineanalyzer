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

const state = {
  path: params.get('path') || '',
  client: params.get('client') || '',
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
  network: null,
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
        : (f.column === '_time'
          ? (row.Received || row.CreatedDateTime || row.CreationDate || '')
          : String(row[f.column] ?? '')));
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

function demoApi(path, options) {
  let body = {};
  try { body = options.body ? JSON.parse(options.body) : {}; } catch { body = {}; }
  const filters = Array.isArray(body.filters) ? body.filters : [];
  if (path === '/api/investigate/sources') {
    return { ok: true, result: { path: DEMO.path, sources: DEMO.sources } };
  }
  if (path === '/api/investigate/mail-graph') {
    return {
      ok: true,
      result: {
        path: DEMO.path,
        ...buildDemoMailGraph(filters, body),
      },
    };
  }
  if (path === '/api/investigate/query') {
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
  if (path === '/api/investigate/export') {
    return { ok: true, result: { file: '(demo mode — export skipped)', rowCount: 0 } };
  }
  throw new Error(`Demo has no handler for ${path}`);
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
  const d = new Date(raw.includes('T') && raw.length === 16 ? `${raw}:00` : raw);
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

function currentFilters() {
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
  if (status) {
    const col = state.source === 'MessageTrace' ? 'Status'
      : (state.source === 'SignInLogs' ? 'Status' : 'Operations');
    filters.push({ column: col, op: 'icontains', value: status });
    if (state.source === 'UnifiedAuditLogs') {
      filters[filters.length - 1] = { column: '_Operation', op: 'icontains', value: status };
    }
  }
  return filters;
}

function payload(extra = {}) {
  return JSON.stringify({
    path: state.path,
    filters: currentFilters(),
    ...extra,
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
      Promise.all([loadGraph(), loadTable()]).catch((e) => setStatus(e.message));
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

function renderTimeline(graph) {
  const host = document.getElementById('timeline');
  const bucket = document.getElementById('bucket')?.value || 'hour';
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
    bar.title = `${formatTick(p.t)}: ${p.count} message${p.count === 1 ? '' : 's'}`;
    bar.disabled = !p.count;
    bar.addEventListener('click', () => {
      const start = toLocalInput(p.t);
      document.getElementById('timeStart').value = start;
      const d = new Date(start);
      if (bucket === 'day') d.setDate(d.getDate() + 1);
      else d.setHours(d.getHours() + 1);
      document.getElementById('timeEnd').value = toLocalInput(d.toISOString());
      state.offset = 0;
      loadTable().catch((e) => setStatus(e.message));
    });
    const lbl = document.createElement('div');
    lbl.className = 'timelineLbl';
    lbl.textContent = formatTick(p.t);
    col.append(bar, lbl);
    host.append(col);
  });
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

async function refreshAll() {
  if (!state.path) {
    setStatus('Set a report folder to begin.');
    return;
  }
  try {
    await loadSources();
    await Promise.all([loadGraph(), loadTable()]);
  } catch (e) {
    setStatus(e.message);
  }
}

document.getElementById('reloadBtn').addEventListener('click', () => {
  state.path = folderInput.value.trim();
  folderPathEl.textContent = state.path;
  const q = new URLSearchParams();
  if (state.path) q.set('path', state.path);
  if (state.client) q.set('client', state.client);
  history.replaceState(null, '', `investigate.html?${q}`);
  state.offset = 0;
  refreshAll();
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
  document.getElementById('statusQ').value = '';
  state.offset = 0;
  refreshAll();
});
document.getElementById('graphBtn').addEventListener('click', () => loadGraph().catch((e) => setStatus(e.message)));

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
