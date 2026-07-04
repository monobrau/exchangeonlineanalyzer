const logEl = document.getElementById('log');
const sessionInfo = document.getElementById('sessionInfo');
const tenantsEl = document.getElementById('tenants');
const appRegsEl = document.getElementById('appRegs');

let appRegistrations = [];

function log(msg) {
  const line = `[${new Date().toLocaleTimeString()}] ${msg}\n`;
  logEl.textContent += line;
  logEl.scrollTop = logEl.scrollHeight;
}

async function api(path, options = {}) {
  const res = await fetch(path, {
    headers: { 'Content-Type': 'application/json', ...(options.headers || {}) },
    ...options,
  });
  const text = await res.text();
  let body = text;
  try { body = text ? JSON.parse(text) : null; } catch { /* plain text */ }
  if (!res.ok) {
    const err = typeof body === 'object' && body?.error ? body.error : text;
    throw new Error(err || res.statusText);
  }
  return body;
}

function tenantIdFromDisplay(displayText) {
  if (!displayText) return null;
  const m = displayText.match(/\(([a-fA-F0-9-]{36})\)/);
  if (m) return m[1];
  const stripped = displayText.replace(/\s*\(ESR\)\s*$/, '').trim();
  return /^[a-fA-F0-9-]{36}$/.test(stripped) ? stripped : null;
}

function renderTenants(session) {
  tenantsEl.innerHTML = '';
  if (!session?.tenants?.length) {
    tenantsEl.innerHTML = '<p class="muted">No tenants yet.</p>';
    return;
  }

  for (const t of session.tenants) {
    const div = document.createElement('div');
    div.className = 'tenant';
    div.dataset.client = t.clientNumber;

    const appRegOptions = appRegistrations.map(a =>
      `<option value="${a.tenantId || ''}">${a.displayText}</option>`
    ).join('');

    div.innerHTML = `
      <strong>Client ${t.clientNumber}</strong>
      <span class="muted"> — Graph: ${t.graphAuthenticated ? 'yes' : 'no'}, EXO: ${t.exchangeAuthenticated ? 'yes' : 'no'}</span>
      <div>
        <label>App reg tenant
          <select class="appRegSelect"><option value="">(interactive / auto WCM)</option>${appRegOptions}</select>
        </label>
        <label><input type="checkbox" class="interactiveCheck" /> Force interactive Graph</label>
      </div>
      <div>
        <button class="graphAuth">Graph Auth</button>
        <button class="exoAuth" ${!t.graphAuthenticated ? 'disabled' : ''}>Exchange Auth</button>
        <button class="statusBtn">Tail status</button>
      </div>
      <pre class="tenantStatus muted"></pre>
    `;

    div.querySelector('.graphAuth').addEventListener('click', () => graphAuth(t.clientNumber, div));
    div.querySelector('.exoAuth').addEventListener('click', () => exoAuth(t.clientNumber));
    div.querySelector('.statusBtn').addEventListener('click', () => tailStatus(t.clientNumber, div));

    tenantsEl.appendChild(div);
  }
}

async function refreshSession() {
  const session = await api('/api/session');
  sessionInfo.textContent = session ? `Session ${session.sessionId} (${session.tenantCount} tenant(s))` : 'No session';
  renderTenants(session);
  return session;
}

async function pollAuth(clientNumber, startedToken, successPrefix, failPrefix, waitSeconds = 120) {
  const deadline = Date.now() + waitSeconds * 1000;
  while (Date.now() < deadline) {
    await new Promise(r => setTimeout(r, 2000));
    const session = await api('/api/session');
    const t = session.tenants.find(x => x.clientNumber === clientNumber);
    const resp = t?.lastResponse || '';
    if (resp && resp !== startedToken && !resp.startsWith(startedToken)) {
      return resp;
    }
    if (resp.startsWith(successPrefix)) return resp;
    if (resp.startsWith(failPrefix)) throw new Error(resp);
  }
  throw new Error('Auth timed out');
}

async function graphAuth(clientNumber, div) {
  const tenantId = div.querySelector('.appRegSelect').value || null;
  const interactive = div.querySelector('.interactiveCheck').checked;
  let cmd = 'GRAPH_AUTH';
  if (tenantId) cmd += `|TENANT_ID:${tenantId}`;
  if (interactive) cmd += '|INTERACTIVE:1';

  log(`Client ${clientNumber}: Graph auth…`);
  const initial = await api(`/api/tenants/${clientNumber}/command`, {
    method: 'POST',
    body: JSON.stringify({ command: cmd, waitSeconds: 30 }),
  });

  let final = initial.response;
  if (initial.response === 'GRAPH_AUTH_STARTED') {
    final = await pollAuth(clientNumber, 'GRAPH_AUTH_STARTED', 'GRAPH_AUTH_SUCCESS', 'GRAPH_AUTH_FAILED');
  }
  log(`Client ${clientNumber}: ${final}`);
  await refreshSession();
}

async function exoAuth(clientNumber) {
  log(`Client ${clientNumber}: Exchange auth…`);
  const initial = await api(`/api/tenants/${clientNumber}/command`, {
    method: 'POST',
    body: JSON.stringify({ command: 'EXCHANGE_AUTH', waitSeconds: 30 }),
  });

  let final = initial.response;
  if (initial.response === 'EXCHANGE_AUTH_STARTED') {
    final = await pollAuth(clientNumber, 'EXCHANGE_AUTH_STARTED', 'EXCHANGE_AUTH_SUCCESS', 'EXCHANGE_AUTH_FAILED', 180);
  }
  log(`Client ${clientNumber}: ${final}`);
  await refreshSession();
}

async function tailStatus(clientNumber, div) {
  const data = await api(`/api/tenants/${clientNumber}/status`);
  div.querySelector('.tenantStatus').textContent = data.status || '(empty)';
}

async function runAction(label, fn) {
  log(label);
  try {
    await fn();
  } catch (e) {
    log(`Error: ${e.message}`);
  }
}

document.getElementById('btnNewSession').addEventListener('click', () => runAction('Creating session…', async () => {
  await api('/api/session', {
    method: 'POST',
    body: JSON.stringify({
      investigatorName: '',
      companyName: '',
      daysBack: 7,
      reportSelections: {
        IncludeMessageTrace: true,
        IncludeInboxRules: true,
        IncludeTransportRules: true,
        IncludeSignInLogs: true,
        IncludeAuditLogs: true,
        SignInLogsDaysBack: 7,
        MessageTraceDaysBack: 7,
      },
    }),
  });
  log('Session ready.');
  await refreshSession();
}));

document.getElementById('btnAddTenant').addEventListener('click', () => runAction('Adding tenant worker…', async () => {
  const t = await api('/api/tenants', { method: 'POST' });
  log(`Started Client ${t.clientNumber} (PID ${t.processId}). A PowerShell worker window should open on this PC.`);
  await refreshSession();
}));

document.getElementById('btnRefresh').addEventListener('click', () => refreshSession().catch(e => log(`Error: ${e.message}`)));
document.getElementById('btnLoadAppRegs').addEventListener('click', () => runAction('Loading app registrations…', async () => {
  appRegistrations = await api('/api/app-registrations');
  appRegsEl.textContent = appRegistrations.map(a => a.displayText).join('\n') || '(none in WCM)';
  await refreshSession();
}));

refreshSession().catch(e => log(`Error: ${e.message}`));
