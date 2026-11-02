const params = new URLSearchParams(location.search);
const statusEl = document.getElementById('status');
const state = {
  folder: params.get('folder') || '',
  client: params.get('client') || '',
  demo: params.get('demo') === '1',
  view: null,
  selectedKey: '',
};

const AI_DEFAULTS = { chatgpt: 'gpt-4.1-mini', claude: 'claude-sonnet-5', hatz: 'gpt-5.2' };
const AI_NAMES = { chatgpt: 'ChatGPT', claude: 'Claude', hatz: 'Hatz' };
const AI_STATIC = {
  chatgpt: [
    { name: 'gpt-4.1-mini', displayName: 'GPT-4.1 mini', developer: 'OpenAI' },
    { name: 'gpt-4.1', displayName: 'GPT-4.1', developer: 'OpenAI' },
  ],
  claude: [
    { name: 'claude-sonnet-5', displayName: 'Claude Sonnet 5', developer: 'Anthropic' },
    { name: 'claude-haiku-4-5', displayName: 'Claude Haiku 4.5', developer: 'Anthropic' },
  ],
  hatz: [
    { name: 'auto', displayName: 'Auto', developer: 'Hatz' },
    { name: 'gpt-5.2', displayName: 'GPT-5.2', developer: 'OpenAI' },
    { name: 'gpt-4o', displayName: 'GPT-4o', developer: 'OpenAI' },
    { name: 'claude-sonnet-5', displayName: 'Claude Sonnet', developer: 'Anthropic' },
    { name: 'claude-haiku-4-5', displayName: 'Claude Haiku 4.5', developer: 'Anthropic' },
  ],
};
function aiModelChoices(provider) {
  const saved = state.ai || {};
  const live = provider === 'hatz' && saved.provider === 'hatz' ? asArray(saved.models) : [];
  const rows = live.length ? live : (AI_STATIC[provider] || []);
  return rows.map((row) => ({
    name: String(row.name || '').trim(),
    displayName: String(row.displayName || row.name || '').trim(),
    developer: String(row.developer || 'Other').trim() || 'Other',
  })).filter((row) => row.name);
}

function fillAiModelSelect(provider, selected) {
  const select = document.getElementById('aiModel');
  const chosen = String(selected || AI_DEFAULTS[provider] || '').trim();
  const models = aiModelChoices(provider);
  const names = new Set(models.map((row) => row.name));
  const groups = [];
  const byDeveloper = new Map();
  models.forEach((row) => {
    if (!byDeveloper.has(row.developer)) {
      const bucket = [];
      byDeveloper.set(row.developer, bucket);
      groups.push([row.developer, bucket]);
    }
    byDeveloper.get(row.developer).push(row);
  });
  if (chosen && !names.has(chosen)) groups.push(['Saved', [{ name: chosen, displayName: chosen, developer: 'Saved' }]]);
  select.replaceChildren();
  groups.forEach(([developer, rows]) => {
    const group = document.createElement('optgroup');
    group.label = developer;
    rows.forEach((row) => {
      const option = document.createElement('option');
      option.value = row.name;
      option.textContent = row.displayName && row.displayName !== row.name ? row.displayName : row.name;
      group.appendChild(option);
    });
    select.appendChild(group);
  });
  if (chosen && [...select.options].some((option) => option.value === chosen)) select.value = chosen;
  else if (select.options.length) select.selectedIndex = 0;
}

const demoAi = { provider: '', model: '', hasKey: false, keyHint: '', apiKey: '' };

function demoAiResult() {
  const models = (AI_STATIC[demoAi.provider] || []).map((row) => ({ ...row }));
  return { provider: demoAi.provider, model: demoAi.model, hasKey: demoAi.hasKey, keyHint: demoAi.keyHint, models };
}

const DEMO = {
  Folder: 'C:\\Reports\\Contoso\\Triage_demo',
  Warning: 'Message trace was limited to 10 days. Older mail needs a historical trace, which this pull does not run.',
  Rows: [
    {
      RowKey: 'row-phish',
      SenderAddress: 'jane.doe.exec@gmail.com',
      RecipientAddress: 'alex@contoso.com',
      Subject: 'Jane Doe password expires',
      Verdict: 'Malicious',
      SuggestedVerdict: 'Malicious',
      Layer: 'anti-phish',
      Evidence: ['Exec name in subject from a freemail sender', 'Forefront CAT: PHSH', 'compauth=fail reason 001'],
      ReachedMailbox: false,
      Impersonation: 'External',
      Outcome: 'Quarantine',
      RawHeader: 'Authentication-Results: spf=fail smtp.mailfrom=example.com; dkim=none; dmarc=fail header.from=example.com; compauth=fail reason=001\r\nX-Forefront-Antispam-Report: CIP:203.0.113.10;CTRY:US;SCL:9;SFV:SPM;CAT:PHSH;SFTY:9.20\r\nX-Microsoft-Antispam: BCL:8',
    },
    {
      RowKey: 'row-delivered',
      SenderAddress: 'billing@evil.example',
      RecipientAddress: 'ap@contoso.com',
      Subject: 'Updated accounts payable bank details',
      Verdict: 'Malicious',
      SuggestedVerdict: 'Malicious',
      Layer: '',
      Evidence: ['Invoice or payment-fraud wording in the subject'],
      ReachedMailbox: true,
      Impersonation: '',
      Outcome: 'ReachedMailbox',
      RawHeader: '',
    },
    {
      RowKey: 'row-spoof',
      SenderAddress: 'ceo.spoof@contoso.com',
      RecipientAddress: 'cfo@contoso.com',
      Subject: 'Jane Doe wire',
      Verdict: 'Malicious',
      SuggestedVerdict: 'Malicious',
      Layer: 'anti-phish',
      Evidence: ['Exec name in subject from an internal-domain sender'],
      ReachedMailbox: false,
      Impersonation: 'SpoofedInternal',
      Outcome: 'Quarantine',
      RawHeader: '',
    },
    {
      RowKey: 'row-noise',
      SenderAddress: 'news@newsletter.example',
      RecipientAddress: 'alex@contoso.com',
      Subject: 'Weekly product newsletter',
      Verdict: 'Noise',
      SuggestedVerdict: 'Noise',
      Layer: '',
      Evidence: ['Bulk newsletter'],
      ReachedMailbox: true,
      Impersonation: '',
      Outcome: 'Delivered',
      RawHeader: '',
    },
    {
      RowKey: 'row-quote',
      SenderAddress: 'vendor@parts.example',
      RecipientAddress: 'alex@contoso.com',
      Subject: 'Quote for the dock repair',
      Verdict: 'Unknown',
      SuggestedVerdict: 'Unknown',
      Layer: '',
      Evidence: ['No rule matched'],
      ReachedMailbox: true,
      Impersonation: '',
      Outcome: 'Delivered',
      RawHeader: '',
    },
    {
      RowKey: 'row-ship-1',
      SenderAddress: 'notices@shop.example',
      RecipientAddress: 'alex@contoso.com',
      Subject: 'Your order has shipped',
      Verdict: 'Unknown',
      SuggestedVerdict: 'Unknown',
      Layer: '',
      Evidence: ['No rule matched'],
      ReachedMailbox: true,
      Impersonation: '',
      Outcome: 'Delivered',
      RawHeader: '',
    },
    {
      RowKey: 'row-ship-2',
      SenderAddress: 'notices@shop.example',
      RecipientAddress: 'sam@contoso.com',
      Subject: 'Your order has shipped',
      Verdict: 'Unknown',
      SuggestedVerdict: 'Unknown',
      Layer: '',
      Evidence: ['No rule matched'],
      ReachedMailbox: true,
      Impersonation: '',
      Outcome: 'Delivered',
      RawHeader: '',
    },
    {
      RowKey: 'row-ship-3',
      SenderAddress: 'notices@shop.example',
      RecipientAddress: 'pat@contoso.com',
      Subject: 'Your order has shipped',
      Verdict: 'Unknown',
      SuggestedVerdict: 'Unknown',
      Layer: '',
      Evidence: ['No rule matched'],
      ReachedMailbox: true,
      Impersonation: '',
      Outcome: 'Delivered',
      RawHeader: '',
    },
  ],
  Findings: [
    { Check: 'InboundConnectorEnhancedFiltering', Severity: 'High', Detail: 'Inbound connector INKY has no EFSkipIPs or Enhanced Filtering, and MX does not point at Exchange Online Protection.' },
    { Check: 'DmarcPolicyNone', Severity: 'High', Detail: 'DMARC for contoso.com is p=none.' },
  ],
  Orca: {
    DefenderForOffice365: true,
    ServiceLabel: 'Defender for Office 365',
    FailCount: 1,
    PassCount: 4,
    Groups: [
      { Severity: 'High', Items: [{ Name: 'Safe Links', Area: 'Microsoft Defender for Office 365 Policies', Result: 'Fail', FailRecommendation: 'Enable Safe Links for the domains you protect.' }] },
      { Severity: 'Low', Items: [{ Name: 'Spam action', Area: 'EOP', Result: 'Informational', FailRecommendation: '' }] },
    ],
  },
  TicketNote: 'Email security triage - Contoso\nMalicious: 3\nReached a mailbox: 1',
  ClientEmail: 'Hello,\n\nWe reviewed recent email for Contoso.\n\nWhat your staff should do\n- If an email says a password is expiring, open the service directly.\n- Confirm payment changes with a phone call you already have on file.',
  Actions: [
    { Id: 'block:bad@gmail.com', Kind: 'Block', Title: 'bad@gmail.com', Detail: '1 delivered message. Credential lure.', License: 'EOP', Apply: 'Worker', Script: "New-TenantAllowBlockListItems -ListType Sender -Block -Entries 'bad@gmail.com' -NoExpiration -Notes 'Email triage'" },
    { Id: 'allow:ap@vendor.example', Kind: 'Allow', Title: 'ap@vendor.example', Detail: '1 held message looks legitimate.', License: 'EOP', Apply: 'Worker', Script: "New-TenantAllowBlockListItems -ListType Sender -Allow -Entries 'ap@vendor.example' -ExpirationDate (Get-Date).ToUniversalTime().AddDays(30) -Notes 'Email triage'\r\nRelease-QuarantineMessage -Identity 'qid-9'" },
    { Id: 'config:dmarc:contoso.com', Kind: 'Config', Title: 'DMARC for contoso.com', Detail: 'DMARC for contoso.com is p=none.', License: 'DNS', Apply: 'Manual', Script: '# DNS TXT record. The worker cannot publish this.\r\n# Host: _dmarc.contoso.com\r\n# Value: v=DMARC1; p=quarantine; rua=mailto:dmarc@contoso.com' },
  ],
};

function setStatus(text) { statusEl.textContent = text || ''; }

function asArray(value) {
  if (value == null) return [];
  return Array.isArray(value) ? value : [value];
}

function esc(text) {
  return String(text ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

async function api(path, options = {}) {
  if (state.demo) return demoApi(path, options);
  const res = await fetch(path, {
    headers: { 'Content-Type': 'application/json' },
    ...options,
  });
  const text = await res.text();
  let body = text;
  try { body = text ? JSON.parse(text) : null; } catch { /* plain */ }
  if (!res.ok) {
    const err = body && body.error ? body.error : text;
    throw new Error(err || res.statusText);
  }
  return body;
}

function demoApi(path, options) {
  const method = (options && options.method) || 'GET';
  let body = {};
  try { body = options && options.body ? JSON.parse(options.body) : {}; } catch { body = {}; }
  if (path === '/api/triage/ai-settings' && method === 'GET') {
    return { ok: true, result: demoAiResult() };
  }
  if (path === '/api/triage/ai-settings' && method === 'POST') {
    const provider = String(body.provider || '').toLowerCase();
    if (provider && provider !== 'chatgpt' && provider !== 'claude' && provider !== 'hatz') throw new Error('Provider must be ChatGPT, Claude, or Hatz AI.');
    demoAi.provider = provider === 'off' ? '' : provider;
    demoAi.model = String(body.model || '').trim() || AI_DEFAULTS[demoAi.provider] || '';
    if (body.apiKey) {
      demoAi.apiKey = String(body.apiKey);
      demoAi.hasKey = true;
      demoAi.keyHint = '····' + demoAi.apiKey.slice(-4);
    }
    return { ok: true, result: demoAiResult() };
  }
  if (path === '/api/triage/memberberry-note') {
    const review = String(body.review || '').trim();
    if (!review) throw new Error('Finish a review before writing the ticket note.');
    const ticketNumber = String(body.ticketNumber || '').trim();
    const subject = ticketNumber
      ? `[Resolved]: Email triage - False Positive - Ticket #${ticketNumber}`
      : '[Resolved]: Email triage - False Positive';
    const text = [
      `Subject: ${subject}`,
      'Classification: False Positive',
      '',
      'Good afternoon,',
      '',
      'We have classified this as FALSE POSITIVE.',
      '',
      'The sample review did not show a malicious message that needs action.',
      '',
      'Thanks,',
      '',
      'Demo Technician',
    ].join('\n');
    return {
      ok: true,
      result: {
        subject,
        classification: 'False Positive',
        text,
        references: 'Demo note. Memberberry and memory vault are not called in demo mode.',
      },
    };
  }
  if (path === '/api/triage/ai-review') {
    if (!demoAi.provider || !demoAi.hasKey) throw new Error('Choose ChatGPT, Claude, or Hatz AI and save an API key first.');
    const rows = asArray(DEMO.Rows);
    const included = rows.filter((row) => rowSentToAi(row, aiBatchSkipKeys(rows)));
    const leftOut = rows.length - included.length;
    const names = included.map((r) => r.SenderAddress).join(', ') || 'none';
    return {
      ok: true,
      result: {
        provider: demoAi.provider,
        model: demoAi.model,
        text: `Demo review (${demoAi.provider} / ${demoAi.model}). Messages in this review: ${names}. ${leftOut} messages were left out of this review. The DMARC p=none finding is the config item to fix first. This is sample text, not a live model.`,
      },
    };
  }
  if (path.startsWith('/api/triage/packs')) {
    return { ok: true, result: [{ Company: 'Contoso', Name: 'Triage_demo', Folder: DEMO.Folder }] };
  }
  if (path.startsWith('/api/triage/result') || path === '/api/triage/override') {
    return { ok: true, result: DEMO };
  }
  if (path.startsWith('/api/triage/notes') && (!options.method || options.method === 'GET')) {
    return { ok: true, result: { vips: ['Jane Doe'], filterVendor: 'INKY', contacts: ['it@contoso.com'], dnsHost: 'Cloudflare', extraFreemail: [] } };
  }
  if (path === '/api/triage/notes') return { ok: true, result: {} };
  if (path.startsWith('/api/triage/file')) {
    return { ok: true, result: { FileName: 'demo.txt', Content: 'demo' } };
  }
  return { ok: true, result: DEMO };
}

function selectedRow() {
  const rows = asArray(state.view && state.view.Rows);
  return rows.find((r) => r.RowKey === state.selectedKey) || null;
}

function aiBatchSkipKeys(rows) {
  const list = asArray(rows).map((row, index) => ({
    key: row.RowKey || `row${index}`,
    sender: String(row.SenderAddress || '').trim().toLowerCase(),
    subject: String(row.Subject || '').trim().toLowerCase(),
    interesting: row.Verdict === 'Malicious' || !!row.Impersonation,
  }));
  const bySender = {};
  list.forEach((item) => {
    if (!item.sender) return;
    if (!bySender[item.sender]) bySender[item.sender] = [];
    bySender[item.sender].push(item);
  });
  const skip = {};
  Object.keys(bySender).forEach((sender) => {
    const items = bySender[sender];
    const boring = items.filter((item) => !item.interesting);
    if (boring.length >= 5 && boring.length === items.length) {
      boring.forEach((item) => { skip[item.key] = true; });
      return;
    }
    const bySubject = {};
    items.forEach((item) => {
      if (!item.subject) return;
      if (!bySubject[item.subject]) bySubject[item.subject] = [];
      bySubject[item.subject].push(item);
    });
    Object.keys(bySubject).forEach((subject) => {
      const group = bySubject[subject];
      if (group.length < 3 || group.some((item) => item.interesting)) return;
      group.forEach((item) => { skip[item.key] = true; });
    });
  });
  return skip;
}

function rowSentToAi(row, skip) {
  if (!row) return false;
  if (row.AiReview === true || row.AiReview === 'True') return true;
  if (row.AiReview === false || row.AiReview === 'False') return false;
  if (row.Verdict === 'Malicious' || row.Impersonation) return true;
  const key = row.RowKey || '';
  if (key && skip && skip[key]) return false;
  return row.Verdict === 'Unknown';
}

function updateAiSendCount() {
  const el = document.getElementById('aiSendCount');
  if (!el) return;
  const rows = asArray(state.view && state.view.Rows);
  if (!rows.length) {
    el.textContent = 'Load a pack, then uncheck rows you do not want sent.';
    return;
  }
  const sent = rows.filter((row) => rowSentToAi(row, aiBatchSkipKeys(rows))).length;
  el.textContent = `${sent} of ${rows.length} loaded messages will be sent. Noise, Legit, and repeated mail that is not marked malicious start unchecked.`;
}

function renderTable(tbody, rows, cells, colspan) {
  const span = colspan || 6;
  tbody.innerHTML = rows.map((row) => `<tr data-key="${esc(row.RowKey || '')}">${cells(row)}</tr>`).join('')
    || `<tr><td class="muted" colspan="${span}">No rows</td></tr>`;
}

function renderTriage() {
  const verdict = document.getElementById('verdictFilter').value;
  const q = document.getElementById('textFilter').value.trim().toLowerCase();
  const allRows = asArray(state.view && state.view.Rows);
  const skip = aiBatchSkipKeys(allRows);
  let rows = allRows;
  if (verdict) rows = rows.filter((r) => r.Verdict === verdict);
  if (q) {
    rows = rows.filter((r) => `${r.SenderAddress} ${r.Subject} ${asArray(r.Evidence).join(' ')}`.toLowerCase().includes(q));
  }
  const tbody = document.querySelector('#triageTable tbody');
  renderTable(tbody, rows, (r) => `
    <td class="check"><input type="checkbox" data-ai="${esc(r.RowKey)}" ${rowSentToAi(r, skip) ? 'checked' : ''} aria-label="Send ${esc(r.SenderAddress)} to AI" /></td>
    <td>${esc(r.Verdict)}</td>
    <td>${esc(r.SenderAddress)}</td>
    <td>${esc(r.Subject)}</td>
    <td>${esc(r.Layer)}</td>
    <td>${esc(asArray(r.Evidence).join('; '))}</td>
    <td><select data-override="${esc(r.RowKey)}">
      ${['Malicious', 'Legit', 'Noise', 'Unknown'].map((v) => `<option value="${v}" ${v === r.Verdict ? 'selected' : ''}>${esc(v)}</option>`).join('')}
    </select></td>`, 7);
  updateAiSendCount();
  tbody.querySelectorAll('tr').forEach((tr) => {
    tr.addEventListener('click', () => {
      state.selectedKey = tr.dataset.key || '';
      tbody.querySelectorAll('tr').forEach((el) => el.classList.toggle('selected', el === tr));
      const row = selectedRow();
      document.getElementById('btnCopyHeaders').disabled = !(row && row.RawHeader);
    });
  });
  tbody.querySelectorAll('select[data-override]').forEach((sel) => {
    sel.addEventListener('click', (ev) => ev.stopPropagation());
    sel.addEventListener('change', () => applyOverride(sel.dataset.override, sel.value));
  });
  tbody.querySelectorAll('input[data-ai]').forEach((box) => {
    box.addEventListener('click', (ev) => ev.stopPropagation());
    box.addEventListener('change', () => {
      setAiReview(box.dataset.ai, box.checked).catch((err) => setStatus(err.message));
    });
  });
}

function renderDelivery() {
  const rows = state.view && state.view.DeliveryRows
    ? asArray(state.view.DeliveryRows)
    : asArray(state.view && state.view.Rows).filter((r) => r.ReachedMailbox === true || r.ReachedMailbox === 'True');
  const tbody = document.querySelector('#deliveryTable tbody');
  renderTable(tbody, rows, (r) => `<td>${esc(r.SenderAddress)}</td><td>${esc(r.RecipientAddress)}</td><td>${esc(r.Subject)}</td><td>${esc(r.Outcome)}</td>`);
}

function renderImpersonation() {
  const rows = state.view && state.view.ImpersonationRows
    ? asArray(state.view.ImpersonationRows)
    : asArray(state.view && state.view.Rows).filter((r) => r.Impersonation);
  const tbody = document.querySelector('#impersonationTable tbody');
  renderTable(tbody, rows, (r) => `<td>${esc(r.Impersonation)}</td><td>${esc(r.SenderAddress)}</td><td>${esc(r.Subject)}</td><td>${esc(r.Verdict)}</td>`);
}

function renderConfig() {
  const orca = (state.view && state.view.Orca) || {};
  document.getElementById('orcaService').textContent = orca.ServiceLabel || 'ORCA results are not in this pack yet.';
  const groups = asArray(orca.Groups);
  document.getElementById('orcaGroups').innerHTML = groups.map((g) => {
    const items = asArray(g.Items).map((item) => `<li><strong>${esc(item.Name)}</strong> <span class="muted">${esc(item.Result)} · ${esc(item.Area)}</span><div>${esc(item.FailRecommendation || '')}</div></li>`).join('');
    return `<h2 class="sev-${esc(g.Severity)}">${esc(g.Severity)}</h2><ul>${items}</ul>`;
  }).join('') || '<p class="muted">No ORCA failures or informational results.</p>';
  const findings = asArray(state.view && state.view.Findings);
  document.getElementById('customFindings').innerHTML = findings.map((f) => `<p class="sev-${esc(f.Severity)}"><strong>${esc(f.Severity)}</strong> ${esc(f.Detail)}</p>`).join('')
    || '<p class="muted">No custom findings.</p>';
}

function renderOutputs() {
  document.getElementById('ticketNote').textContent = (state.view && state.view.TicketNote) || '';
  document.getElementById('clientEmail').textContent = (state.view && state.view.ClientEmail) || '';
}

const actionPick = new Set();
const actionSort = {
  Block: { key: 'Title', dir: 1 },
  Allow: { key: 'Title', dir: 1 },
  Config: { key: 'Title', dir: 1 },
};

function actionItems(kind) {
  return asArray(state.view && state.view.Actions).filter((item) => item.Kind === kind);
}

function visibleActionItems(kind) {
  const box = document.querySelector(`#tab-actions [data-kind="${kind}"] [data-filter]`);
  const q = (box && box.value || '').trim().toLowerCase();
  let items = actionItems(kind);
  if (q) {
    items = items.filter((item) => `${item.Title} ${item.Detail} ${item.License} ${item.Apply}`.toLowerCase().includes(q));
  }
  const sort = actionSort[kind];
  items = items.slice().sort((a, b) => {
    const av = String(a[sort.key] || '');
    const bv = String(b[sort.key] || '');
    return av.localeCompare(bv) * sort.dir;
  });
  return items;
}

function selectedActionItems(kind) {
  return actionItems(kind).filter((item) => actionPick.has(item.Id));
}

function howLabel(item) {
  return item.Apply === 'Worker' ? 'Apply' : 'Copy only';
}

function renderActions() {
  const help = document.getElementById('actionHelp');
  if (help) {
    help.textContent = state.client
      ? 'Select the rows you want. Apply runs them in this client\'s Exchange session. Copy PowerShell is always available for the same commands. DNS, and Defender for Office 365 changes when that license was not detected, stay copy-only.'
      : 'Select the rows you want. Open this page from the tenant card to apply them in Exchange. Copy PowerShell works either way.';
  }
  document.querySelectorAll('#tab-actions [data-kind]').forEach((section) => {
    const kind = section.dataset.kind;
    const items = visibleActionItems(kind);
    const workerCount = actionItems(kind).filter((item) => item.Apply === 'Worker').length;
    const applyBtn = section.querySelector('[data-act="apply"]');
    const copyBtn = section.querySelector('[data-act="copy"]');
    if (applyBtn) {
      applyBtn.disabled = !state.client || state.demo || workerCount === 0;
      applyBtn.classList.toggle('primary', Boolean(state.client) && !state.demo && workerCount > 0);
    }
    if (copyBtn) copyBtn.classList.toggle('primary', !state.client || state.demo || workerCount === 0);
    const showLicense = kind === 'Config';
    const tbody = section.querySelector('tbody');
    const cols = showLicense ? 5 : 4;
    if (!items.length) {
      const hint = actionItems(kind).length
        ? 'Nothing matches the filter.'
        : kind === 'Allow'
          ? 'No held message looks clearly legitimate. On the Triage tab, set a row to Legit and it will show up here.'
          : kind === 'Block'
            ? 'No delivered message cleared the high-confidence bar.'
            : 'No configuration changes for this pack.';
      tbody.innerHTML = `<tr><td class="muted" colspan="${cols}">${esc(hint)}</td></tr>`;
    } else {
      tbody.innerHTML = items.map((item) => `<tr>
        <td class="check"><input type="checkbox" data-id="${esc(item.Id)}" ${actionPick.has(item.Id) ? 'checked' : ''} aria-label="Select ${esc(item.Title)}" /></td>
        <td>${esc(item.Title)}</td>
        <td>${esc(item.Detail)}</td>
        ${showLicense ? `<td>${esc(item.License)}</td>` : ''}
        <td>${esc(howLabel(item))}</td>
      </tr>`).join('');
    }
    const visibleBox = section.querySelector('[data-act="visible"]');
    if (visibleBox) {
      const ids = items.map((item) => item.Id);
      visibleBox.checked = ids.length > 0 && ids.every((id) => actionPick.has(id));
      visibleBox.indeterminate = !visibleBox.checked && ids.some((id) => actionPick.has(id));
    }
    const sort = actionSort[kind];
    section.querySelectorAll('[data-sort]').forEach((th) => {
      const mark = th.dataset.sort === sort.key ? (sort.dir > 0 ? ' \u25B2' : ' \u25BC') : '';
      const base = th.dataset.label || th.textContent.replace(/ [\u25B2\u25BC]$/, '');
      th.dataset.label = base;
      th.textContent = base + mark;
    });
  });
}

function scriptFor(items) {
  const body = items.map((item) => item.Script || '').filter(Boolean).join('\r\n\r\n');
  return `# Email triage\r\n# Review each command. Apply uses the same commands when the worker can run them.\r\n\r\n${body}\r\n`;
}

function extractToken(value, prefixes) {
  const text = String(value || '');
  for (const prefix of prefixes) {
    const idx = text.indexOf(prefix);
    if (idx >= 0) return text.slice(idx);
  }
  return text.trim();
}

async function pollApply(client) {
  const deadline = Date.now() + 120000;
  while (Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 2000));
    const data = await api(`/api/tenants/${client}/response`);
    const text = extractToken(data && data.response, ['TRIAGE_APPLY:', 'TRIAGE_FAILED:']);
    if (text.startsWith('TRIAGE_APPLY:') || text.startsWith('TRIAGE_FAILED:')) return text;
  }
  throw new Error('Apply timed out. The worker may still be running it.');
}

async function copyActions(kind) {
  const items = selectedActionItems(kind);
  if (!items.length) {
    setStatus('Select at least one row.');
    return;
  }
  const manual = items.filter((item) => item.Apply !== 'Worker').length;
  const note = manual ? ` ${manual} of them are copy-only.` : '';
  try {
    await navigator.clipboard.writeText(scriptFor(items));
  } catch (err) {
    setStatus('Could not reach the clipboard. Click the page once, then copy again.');
    return;
  }
  setStatus(`Copied PowerShell for ${items.length} item(s).${note}`);
}

async function applyActions(kind) {
  const chosen = selectedActionItems(kind);
  const items = chosen.filter((item) => item.Apply === 'Worker');
  const manual = chosen.length - items.length;
  if (!state.client || state.demo) {
    setStatus('Open this page from the tenant card to apply. Copy PowerShell works either way.');
    return;
  }
  if (!items.length) {
    setStatus('Nothing selected can be applied from the worker. Copy PowerShell for those steps.');
    return;
  }
  const label = kind === 'Block' ? 'block' : kind === 'Allow' ? 'allow' : 'change';
  if (!window.confirm(`Apply ${items.length} ${label}(s) in Exchange Online for client ${state.client}?`)) return;
  setStatus('Checking the worker…');
  const worker = await api(`/api/tenants/${state.client}/ensure-worker`, {
    method: 'POST',
    body: JSON.stringify({ restartIfDead: false, showConsole: false }),
  });
  if (!worker || !worker.alive) {
    setStatus('The worker is not running. Restart it from the tenant card, sign in to Exchange, then apply again. Copy PowerShell is ready either way.');
    return;
  }
  const command = `TRIAGE_APPLY|FOLDER:${state.folder}|IDS:${items.map((item) => item.Id).join(';')}`;
  setStatus(`Applying ${items.length}…`);
  const initial = await api(`/api/tenants/${state.client}/command`, {
    method: 'POST',
    body: JSON.stringify({ command, noWait: true, waitSeconds: 0 }),
  });
  let text = extractToken(initial && initial.response, ['TRIAGE_APPLY:', 'TRIAGE_FAILED:', 'TRIAGE_STARTED']);
  if (!text.startsWith('TRIAGE_APPLY:') && !text.startsWith('TRIAGE_FAILED:')) {
    text = await pollApply(state.client);
  }
  if (text.startsWith('TRIAGE_FAILED:')) {
    const detail = text.slice('TRIAGE_FAILED:'.length);
    if (detail.indexOf('Unknown triage command') >= 0) {
      setStatus('This worker was started before Apply existed. Restart it from the tenant card, sign in to Exchange, and try again. Copy PowerShell works now.');
      return;
    }
    throw new Error(detail);
  }
  let results = [];
  try { results = asArray(JSON.parse(text.slice('TRIAGE_APPLY:'.length)).results); } catch { results = []; }
  const applied = results.filter((r) => r.Status === 'Applied').length;
  const failed = results.filter((r) => r.Status === 'Failed');
  const parts = [`Applied ${applied} of ${items.length}.`];
  if (failed.length) parts.push(failed.map((r) => `${r.Id}: ${r.Detail}`).join(' '));
  if (manual) parts.push(`${manual} selected item(s) stay as PowerShell. Use Copy for those.`);
  setStatus(parts.join(' '));
}

function renderAll() {
  const warn = document.getElementById('warning');
  const text = (state.view && state.view.Warning) || '';
  warn.hidden = !text;
  warn.textContent = text;
  renderTriage();
  renderActions();
  renderDelivery();
  renderImpersonation();
  renderConfig();
  renderOutputs();
}

async function loadPacks() {
  const data = await api(`/api/triage/packs${state.client ? `?client=${encodeURIComponent(state.client)}` : ''}`);
  const packs = asArray(data.result);
  const sel = document.getElementById('packSelect');
  sel.innerHTML = packs.map((p) => `<option value="${esc(p.Folder)}">${esc(p.Company)} / ${esc(p.Name)}</option>`).join('');
  if (state.folder) sel.value = state.folder;
  else if (packs[0]) state.folder = packs[0].Folder;
}

async function loadResult() {
  if (!state.folder && !state.demo) {
    setStatus('Choose a pack.');
    return;
  }
  setStatus('Loading triage pack…');
  const q = new URLSearchParams();
  if (state.folder) q.set('folder', state.folder);
  if (state.client) q.set('client', state.client);
  const verdict = document.getElementById('verdictFilter').value;
  q.set('limit', '500');
  if (verdict) q.set('verdict', verdict);
  const data = await api(`/api/triage/result?${q}`);
  state.view = data.result;
  renderAll();
  const counts = (state.view && state.view.Counts) || {};
  const shown = asArray(state.view && state.view.Rows).length;
  const slice = (state.view && state.view.Slice) || 'rows';
  const sliceTotal = counts[slice] != null ? counts[slice] : counts.Total;
  const packTotal = counts.Total != null ? ` Pack has ${counts.Total} rows.` : '';
  setStatus(`Showing ${shown} of ${sliceTotal == null ? shown : sliceTotal} ${slice}.${packTotal}`);
}

async function setAiReview(rowKey, include) {
  const keys = String(rowKey || '').split('|').map((key) => key.trim()).filter(Boolean);
  if (!keys.length) return;
  if (state.demo) {
    keys.forEach((key) => {
      const row = asArray(DEMO.Rows).find((r) => r.RowKey === key);
      if (row) row.AiReview = include;
    });
    renderTriage();
    setStatus(include ? 'Included in the AI review.' : 'Left out of the AI review.');
    return;
  }
  setStatus('Updating the AI review selection…');
  const data = await api('/api/triage/override', {
    method: 'POST',
    body: JSON.stringify({
      folder: state.folder,
      rowKey: keys.join('|'),
      verdict: include ? 'Include' : 'Omit',
      client: state.client,
      limit: 500,
      verdictFilter: document.getElementById('verdictFilter').value,
    }),
  });
  state.view = data.result;
  renderAll();
  setStatus(include ? 'Included in the AI review.' : 'Left out of the AI review.');
}

function visibleTriageKeys() {
  return [...document.querySelectorAll('#triageTable tbody tr')]
    .map((tr) => tr.dataset.key || '')
    .filter(Boolean);
}

async function applyOverride(rowKey, verdict) {
  if (verdict === 'Include' || verdict === 'Omit') {
    await setAiReview(rowKey, verdict === 'Include');
    return;
  }
  if (!verdict || verdict === 'suggested' || state.demo) {
    if (state.demo && verdict && verdict !== 'suggested') {
      const row = asArray(DEMO.Rows).find((r) => r.RowKey === rowKey);
      if (row) row.Verdict = verdict;
      renderTriage();
    }
    return;
  }
  setStatus('Saving override…');
  const data = await api('/api/triage/override', {
    method: 'POST',
    body: JSON.stringify({
      folder: state.folder,
      rowKey,
      verdict,
      client: state.client,
      limit: 500,
      verdictFilter: document.getElementById('verdictFilter').value,
    }),
  });
  state.view = data.result;
  renderAll();
  setStatus('Override saved. Original CSVs were not changed.');
}

async function loadNotes() {
  if (!state.client && !state.demo) return;
  const data = await api(`/api/triage/notes?client=${encodeURIComponent(state.client || 'demo')}`);
  const n = data.result || {};
  document.getElementById('noteVips').value = asArray(n.vips).join(', ');
  document.getElementById('noteVendor').value = n.filterVendor || '';
  document.getElementById('noteContacts').value = asArray(n.contacts).join(', ');
  document.getElementById('noteDns').value = n.dnsHost || '';
  document.getElementById('noteFree').value = asArray(n.extraFreemail).join(', ');
}

function splitList(text) {
  return text.split(/[,;\n]/).map((s) => s.trim()).filter(Boolean);
}

async function saveNotes() {
  const body = {
    client: state.client || 'demo',
    vips: splitList(document.getElementById('noteVips').value),
    filterVendor: document.getElementById('noteVendor').value.trim(),
    contacts: splitList(document.getElementById('noteContacts').value),
    dnsHost: document.getElementById('noteDns').value.trim(),
    extraFreemail: splitList(document.getElementById('noteFree').value),
  };
  await api('/api/triage/notes', { method: 'POST', body: JSON.stringify(body) });
  setStatus(state.demo ? 'Demo notes are not written to disk.' : 'Client notes saved.');
}

async function downloadFile(name) {
  const q = new URLSearchParams({ folder: state.folder || DEMO.Folder, name });
  const data = await api(`/api/triage/file?${q}`);
  const file = data.result || {};
  const blob = new Blob([file.Content || ''], { type: 'text/plain' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = file.FileName || name;
  a.click();
  URL.revokeObjectURL(a.href);
}

document.getElementById('tab-actions').addEventListener('click', (ev) => {
  const section = ev.target.closest('[data-kind]');
  if (!section) return;
  const kind = section.dataset.kind;
  const sortTh = ev.target.closest('[data-sort]');
  if (sortTh) {
    const key = sortTh.dataset.sort;
    if (actionSort[kind].key === key) actionSort[kind].dir *= -1;
    else { actionSort[kind].key = key; actionSort[kind].dir = 1; }
    renderActions();
    return;
  }
  const act = ev.target.closest('[data-act]');
  if (!act) return;
  if (act.dataset.act === 'all') {
    visibleActionItems(kind).forEach((item) => actionPick.add(item.Id));
    renderActions();
  } else if (act.dataset.act === 'none') {
    actionItems(kind).forEach((item) => actionPick.delete(item.Id));
    renderActions();
  } else if (act.dataset.act === 'copy') {
    copyActions(kind).catch((err) => setStatus(err.message));
  } else if (act.dataset.act === 'apply') {
    applyActions(kind).catch((err) => setStatus(err.message));
  }
});
document.getElementById('tab-actions').addEventListener('change', (ev) => {
  const section = ev.target.closest('[data-kind]');
  if (!section) return;
  const kind = section.dataset.kind;
  if (ev.target.dataset.act === 'visible') {
    const ids = visibleActionItems(kind).map((item) => item.Id);
    ids.forEach((id) => (ev.target.checked ? actionPick.add(id) : actionPick.delete(id)));
    renderActions();
    return;
  }
  if (ev.target.dataset.id) {
    if (ev.target.checked) actionPick.add(ev.target.dataset.id);
    else actionPick.delete(ev.target.dataset.id);
    renderActions();
  }
});
document.getElementById('tab-actions').addEventListener('input', (ev) => {
  if (ev.target.dataset.filter != null) renderActions();
});
document.getElementById('tabs').addEventListener('click', (ev) => {
  const btn = ev.target.closest('button[data-tab]');
  if (!btn) return;
  document.querySelectorAll('#tabs button').forEach((el) => el.classList.toggle('active', el === btn));
  document.querySelectorAll('.panel').forEach((el) => el.classList.toggle('active', el.id === `tab-${btn.dataset.tab}`));
});
document.getElementById('btnLoad').addEventListener('click', () => {
  state.folder = document.getElementById('packSelect').value;
  loadResult().catch((err) => setStatus(err.message));
});
document.getElementById('verdictFilter').addEventListener('change', () => {
  loadResult().catch((err) => setStatus(err.message));
});
document.getElementById('textFilter').addEventListener('input', renderTriage);
document.getElementById('btnAiDropVisible').addEventListener('click', () => {
  const keys = visibleTriageKeys();
  if (!keys.length) {
    setStatus('No visible messages to update.');
    return;
  }
  setAiReview(keys.join('|'), false).catch((err) => setStatus(err.message));
});
document.getElementById('btnAiSendVisible').addEventListener('click', () => {
  const keys = visibleTriageKeys();
  if (!keys.length) {
    setStatus('No visible messages to update.');
    return;
  }
  setAiReview(keys.join('|'), true).catch((err) => setStatus(err.message));
});
document.getElementById('btnCopyHeaders').addEventListener('click', async () => {
  const row = selectedRow();
  if (!row || !row.RawHeader) return;
  await navigator.clipboard.writeText(row.RawHeader);
  setStatus('Raw headers copied. Paste them into Message Header Analyzer for a full analysis.');
});
document.getElementById('btnCopyTicket').addEventListener('click', async () => {
  let text = document.getElementById('ticketNote').textContent || '';
  if (text.indexOf('Use the copy button') >= 0 && !state.demo) {
    const q = new URLSearchParams({ folder: state.folder, name: 'TicketNote.txt' });
    const data = await api(`/api/triage/file?${q}`);
    text = (data.result && data.result.Content) || text;
  }
  await navigator.clipboard.writeText(text);
  setStatus('Ticket note copied.');
});
document.getElementById('btnCopyEmail').addEventListener('click', async () => {
  await navigator.clipboard.writeText(document.getElementById('clientEmail').textContent || '');
  setStatus('Client email copied.');
});
document.getElementById('btnSaveNotes').addEventListener('click', () => saveNotes().catch((err) => setStatus(err.message)));
document.querySelectorAll('[data-file]').forEach((btn) => {
  btn.addEventListener('click', () => downloadFile(btn.dataset.file).catch((err) => setStatus(err.message)));
});

function aiSummaryText(settings) {
  if (!settings || !settings.provider) return 'off';
  const name = AI_NAMES[settings.provider] || settings.provider;
  return settings.model ? `${name} · ${settings.model}` : name;
}

function applyAiSettings(settings) {
  const provider = (settings && settings.provider) || '';
  document.getElementById('aiProvider').value = provider;
  state.ai = settings || {};
  fillAiModelSelect(provider, (settings && settings.model) || AI_DEFAULTS[provider] || '');
  document.getElementById('aiKey').value = '';
  document.getElementById('aiKeyHint').textContent = settings && settings.hasKey ? `Saved key ${settings.keyHint}` : 'No key saved';
  document.getElementById('aiSummary').textContent = aiSummaryText(settings);
  document.getElementById('btnAiReview').disabled = !(provider && settings && settings.hasKey && state.view);
}

function onAiProviderChange() {
  const provider = document.getElementById('aiProvider').value;
  const current = document.getElementById('aiModel').value;
  const choices = aiModelChoices(provider);
  const keep = choices.some((row) => row.name === current);
  fillAiModelSelect(provider, keep ? current : (AI_DEFAULTS[provider] || ''));
  document.getElementById('aiSummary').textContent = aiSummaryText({
    provider,
    model: document.getElementById('aiModel').value,
  });
}

async function loadAiSettings() {
  const data = await api('/api/triage/ai-settings');
  applyAiSettings(data.result || {});
}

async function saveAiSettings() {
  const data = await api('/api/triage/ai-settings', {
    method: 'POST',
    body: JSON.stringify({
      provider: document.getElementById('aiProvider').value,
      model: document.getElementById('aiModel').value,
      apiKey: document.getElementById('aiKey').value.trim(),
    }),
  });
  applyAiSettings(data.result || {});
  setStatus(state.demo ? 'Demo AI settings stay in this tab only.' : 'AI review settings saved on this PC.');
}

async function runAiReview() {
  if (!state.view) {
    setStatus('Load a triage pack first.');
    return;
  }
  const name = AI_NAMES[state.ai && state.ai.provider] || 'the model';
  setStatus(`Asking ${name}…`);
  document.getElementById('btnAiReview').disabled = true;
  try {
    const data = await api('/api/triage/ai-review', {
      method: 'POST',
      body: JSON.stringify({
        folder: state.folder || (state.demo ? DEMO.Folder : ''),
        client: state.client || '',
      }),
    });
    const out = document.getElementById('aiReview');
    out.textContent = (data.result && data.result.text) || '';
    document.getElementById('aiReviewWrap').hidden = !out.textContent;
    document.getElementById('aiCard').open = true;
    setStatus('AI review finished.');
  } finally {
    const settings = state.ai || {};
    document.getElementById('btnAiReview').disabled = !(settings.provider && settings.hasKey && state.view);
  }
}

document.getElementById('aiProvider').addEventListener('change', onAiProviderChange);
document.getElementById('aiModel').addEventListener('change', () => {
  document.getElementById('aiSummary').textContent = aiSummaryText({
    provider: document.getElementById('aiProvider').value,
    model: document.getElementById('aiModel').value,
  });
});
document.getElementById('btnSaveAi').addEventListener('click', () => saveAiSettings().catch((err) => setStatus(err.message)));
document.getElementById('btnAiReview').addEventListener('click', () => runAiReview().catch((err) => setStatus(err.message)));
async function writeMemberberryNote() {
  const review = document.getElementById('aiReview').textContent || '';
  if (!review.trim()) {
    setStatus('Finish a review before writing the ticket note.');
    return;
  }
  setStatus('Writing the Memberberry note…');
  const data = await api('/api/triage/memberberry-note', {
    method: 'POST',
    body: JSON.stringify({
      review,
      company: (state.view && state.view.Company) || '',
      ticketNumber: (document.getElementById('memberberryTicket').value || '').trim(),
    }),
  });
  const note = document.getElementById('memberberryNote');
  note.textContent = (data.result && data.result.text) || '';
  document.getElementById('memberberryRefs').textContent = (data.result && data.result.references) || '';
  document.getElementById('memberberryWrap').hidden = !note.textContent;
  setStatus(note.textContent ? 'Memberberry note ready.' : 'The note was empty.');
}

document.getElementById('btnMemberberry').addEventListener('click', () => writeMemberberryNote().catch((err) => setStatus(err.message)));
document.getElementById('btnCopyMemberberry').addEventListener('click', async () => {
  const text = document.getElementById('memberberryNote').textContent || '';
  if (!text) {
    setStatus('Write the note before copying it.');
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    setStatus('Memberberry note copied.');
  } catch (err) {
    setStatus('Could not reach the clipboard. Click the page once, then copy again.');
  }
});
document.getElementById('btnCopyAi').addEventListener('click', async () => {
  const text = document.getElementById('aiReview').textContent || '';
  if (!text) {
    setStatus('Run a review before copying.');
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    setStatus('AI review copied.');
  } catch (err) {
    setStatus('Could not reach the clipboard. Click the page once, then copy again.');
  }
});

if (state.client) document.getElementById('clientLabel').textContent = `Client ${state.client}`;
if (state.demo) document.getElementById('clientLabel').textContent = 'Demo pack';

loadPacks()
  .then(() => loadResult())
  .then(() => loadNotes())
  .then(() => loadAiSettings())
  .catch((err) => setStatus(err.message));
