const state = {
  catalog: null,
  tickets: [],
  ticket: null,
  tenants: [],
  review: '',
  audit: null,
  memory: {},
  windowStart: '',
  pendingRepull: null,
  busy: false,
  queue: null,
  boardStatuses: [],
  statusChoices: [],
  thread: [],
  threadKind: 'discussion',
  manageWrites: [],
  picked: new Set(),
  lastPickId: '',
};

const REMEDIATE_OK = [
  'REMEDIATE_SUCCESS:', 'REMEDIATE_RESTRICTED:', 'REMEDIATE_RULES:', 'REMEDIATE_MAILBOX:',
  'REMEDIATE_TRANSPORT:', 'REMEDIATE_CONNECTORS:', 'REMEDIATE_AUTHMETHODS:', 'REMEDIATE_DEVICES:',
  'REMEDIATE_APPS:', 'REMEDIATE_OAUTH:', 'REMEDIATE_MOBILE:', 'REMEDIATE_FOLDERS:',
  'REMEDIATE_AUTOREPLY:', 'REMEDIATE_ORGFWD:', 'REMEDIATE_JUNK:', 'REMEDIATE_JOURNAL:',
  'REMEDIATE_HOLD:', 'REMEDIATE_ELSEWHERE:', 'REMEDIATE_ROLES:', 'REMEDIATE_INTUNE:',
  'REMEDIATE_APPCREDS:', 'REMEDIATE_FLOWS:', 'REMEDIATE_BEC:', 'REMEDIATE_HUNTER_GRAPH:',
  'REMEDIATE_HUNTER_LISTS:',
];
const REPULL_LOGS = ['trace', 'ual', 'sharing', 'files', 'dlp', 'alerts', 'signins', 'apps'];

const $ = (id) => document.getElementById(id);

function showOutput(name, options) {
  const card = $('outputCard');
  if (card) card.hidden = false;
  document.querySelectorAll('.outputTab').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.output === name);
  });
  const map = {
    ticket: 'panelTicket',
    review: 'panelReview',
    client: 'panelClient',
    internal: 'panelInternal',
    incident: 'panelIncident',
  };
  if (name === 'internal') {
    const wrap = $('internalWrap');
    if (wrap) wrap.open = true;
  }
  if (name === 'incident') {
    const wrap = $('incidentWrap');
    if (wrap) wrap.open = true;
  }
  const panel = $(map[name] || 'panelTicket');
  if (panel && panel.scrollIntoView) panel.scrollIntoView({ block: 'nearest' });
  const setup = $('ticketSetup');
  if (setup && !(options && options.keepSetup)) setup.open = false;
}

function backToQueue() {
  $('ticketCard').hidden = true;
  $('outputCard').hidden = true;
  $('workBar').hidden = true;
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value == null || value === '') return [];
  return [value];
}

const CLIENT_INSERTS = {
  authorized: [
    "I've reviewed this alert and the related activity.",
    '',
    'After looking into the details, here are my findings.',
    '',
    'I can confirm this is authorized activity. No further action is required on this alert.',
  ].join('\n'),
  'false-positive': [
    "I've reviewed this alert and the related activity.",
    '',
    'After looking into the details, here are my findings.',
    '',
    'This appears to be benign activity that triggered an incorrect alert. After thorough investigation, I can confirm this is a false positive.',
  ].join('\n'),
  'true-positive': [
    "I've reviewed this alert and the related activity.",
    '',
    'After looking into the details, here are my findings.',
    '',
    'This represents a legitimate security concern that required remediation. Appropriate steps have been taken to address the issue. Please mark this as a true positive.',
  ].join('\n'),
  'unsecured-credentials': [
    "I've reviewed the Huntress alert regarding potential unsecured credentials on this endpoint.",
    '',
    'Huntress identified one or more files on this endpoint (or a network location accessed from it) that appear to store passwords in plain text. This is a security best-practice concern rather than confirmation of active compromise.',
    '',
    'We recommend working with the affected user(s) to move any needed credentials into an approved password manager and remove plain-text password files from workstations and shared drives. Before deleting anything, make sure required passwords are captured in a secure, encrypted location.',
    '',
    'This is a low-severity advisory, but it is worth addressing promptly since unsecured credential files are commonly targeted during post-compromise activity.',
  ].join('\n'),
};

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[ch]));
}

async function api(path, options = {}, timeoutMs = 20000) {
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
    try { body = text ? JSON.parse(text) : null; } catch { /* plain text */ }
    if (!res.ok) {
      const err = typeof body === 'object' && body?.error ? body.error : text;
      throw new Error(err || res.statusText);
    }
    return body;
  } catch (e) {
    if (e.name === 'AbortError') throw new Error(`Request timed out after ${timeoutMs / 1000}s: ${path}`);
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

function setStatus(id, text, tone) {
  const el = $(id);
  if (!el) return;
  el.textContent = text || '';
  el.className = 'status' + (tone ? ` ${tone}` : '');
}

function setBusy(busy) {
  state.busyDepth = Math.max(0, (state.busyDepth || 0) + (busy ? 1 : -1));
  state.busy = state.busyDepth > 0;
  ['btnRefresh', 'btnScan', 'btnTriage', 'btnEnrich', 'btnNotes', 'btnReviseReview', 'btnReviseNotes', 'btnPost', 'btnRepull', 'btnReviewAgain', 'btnTenants', 'btnFindUser', 'btnReturn', 'btnSetStatus', 'btnAddTenant', 'btnExoAuth', 'btnGraphAuth'].forEach((id) => {
    const el = $(id);
    if (el) el.disabled = state.busy;
  });
}

function normalize(value) {
  return (value == null ? '' : String(value)).trim();
}

function extractToken(value, prefixes) {
  const text = normalize(value);
  if (!text || !prefixes?.length) return text;
  for (const prefix of prefixes) {
    const idx = text.indexOf(prefix);
    if (idx >= 0) return text.slice(idx);
  }
  return text;
}

function parsePayload(final) {
  const prefixes = ['REMEDIATE_FAILED:', ...REMEDIATE_OK];
  const text = extractToken(final || '', prefixes);
  for (const prefix of prefixes) {
    if (text.startsWith(prefix)) {
      const rest = text.slice(prefix.length);
      try { return { prefix, data: JSON.parse(rest), raw: rest }; }
      catch { return { prefix, data: null, raw: rest }; }
    }
  }
  return { prefix: '', data: null, raw: text };
}

async function ensureWorker(clientNumber) {
  const worker = await api(`/api/tenants/${clientNumber}/ensure-worker`, {
    method: 'POST',
    body: JSON.stringify({ restartIfDead: true }),
  }, 60000);
  if (worker.restarted || worker.requiresReauth) {
    throw new Error('The worker restarted. Sign in again with Exchange Auth and Graph Auth in this window, then run the check.');
  }
  if (!worker.alive) {
    throw new Error('The PowerShell worker is not running. Open the main window and sign in, then try again.');
  }
}

async function sendCommand(clientNumber, command, label, waitSeconds) {
  const before = await api(`/api/tenants/${clientNumber}/response`);
  let prior = normalize(before?.response);
  const initial = await api(`/api/tenants/${clientNumber}/command`, {
    method: 'POST',
    body: JSON.stringify({ command, noWait: true, waitSeconds: 0 }),
  });
  let text = extractToken(initial.response, ['REMEDIATE_FAILED:', ...REMEDIATE_OK, 'REMEDIATE_STARTED']);
  if (text && text !== 'REMEDIATE_STARTED' && text !== prior && (text.startsWith('REMEDIATE_FAILED:') || REMEDIATE_OK.some((p) => text.startsWith(p)))) {
    return text;
  }
  const deadline = Date.now() + waitSeconds * 1000;
  let lastNote = 0;
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, 2000));
    const data = await api(`/api/tenants/${clientNumber}/response`);
    const resp = normalize(data?.response);
    const elapsed = Math.round((Date.now() - (deadline - waitSeconds * 1000)) / 1000);
    if (elapsed - lastNote >= 15) {
      setStatus('scanStatus', `${label}… (${elapsed}s)`);
      lastNote = elapsed;
    }
    if (!resp || resp === prior) continue;
    const fail = extractToken(resp, ['REMEDIATE_FAILED:']);
    const ok = extractToken(resp, REMEDIATE_OK);
    if (fail.startsWith('REMEDIATE_FAILED:')) throw new Error(fail);
    if (REMEDIATE_OK.some((p) => ok.startsWith(p))) return ok;
    if (resp.includes('REMEDIATE_STARTED')) prior = resp;
  }
  throw new Error(`${label} timed out`);
}

function isSocEmail(addr) {
  const domain = String(addr || '').toLowerCase().split('@')[1] || '';
  return /(^|\.)(perchsecurity\.com|connectwise\.com|connectwise\.net|myconnectwise\.net|barracuda\.com|barracudanetworks\.com|barracudamsp\.com|huntress\.io|huntress\.com)$/.test(domain);
}

function emailsIn(text) {
  const found = String(text || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [];
  const seen = new Set();
  const out = [];
  found.forEach((raw) => {
    const addr = raw.toLowerCase();
    if (seen.has(addr) || isSocEmail(addr)) return;
    seen.add(addr);
    out.push(addr);
  });
  return out.slice(0, 40);
}

function mailboxesInAlert(text) {
  const source = String(text || '');
  const seen = new Set();
  const out = [];
  const add = (raw) => {
    const addr = String(raw || '').trim().toLowerCase();
    if (!addr || seen.has(addr) || isSocEmail(addr)) return;
    if (!/^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$/.test(addr)) return;
    seen.add(addr);
    out.push(addr);
  };
  const labeled = /(?:identity|user(?:\s*principal\s*name)?|upn|mailbox|affected\s+user|username)\**\s*[:=]\s*\**\s*([a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,})/gi;
  let match = labeled.exec(source);
  while (match) {
    add(match[1]);
    match = labeled.exec(source);
  }
  const personal = /\/personal\/([a-z0-9._-]+)(?:\/|\?|#|$)/gi;
  const tlds = new Set(['com', 'net', 'org', 'biz', 'us', 'io', 'co', 'email', 'info', 'edu', 'gov', 'ca', 'uk']);
  match = personal.exec(source);
  while (match) {
    const parts = match[1].toLowerCase().split('_').filter(Boolean);
    const tld = parts[parts.length - 1];
    if (parts.length >= 3 && tlds.has(tld)) {
      const domainLabels = parts[parts.length - 2] === 'onmicrosoft' ? 2 : 1;
      if (parts.length > domainLabels + 1) {
        const domain = parts.slice(parts.length - domainLabels - 1).join('.');
        const local = parts.slice(0, parts.length - domainLabels - 1).join('.');
        add(`${local}@${domain}`);
      }
    }
    match = personal.exec(source);
  }
  return out;
}

function reportedLead(mailbox) {
  const text = alertSourceText();
  const want = String(mailbox || '').trim().toLowerCase();
  const home = want.split('@')[1] || '';
  const found = [];
  const re = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  let match = re.exec(text);
  while (match) {
    const addr = match[0].toLowerCase();
    const domain = addr.split('@')[1] || '';
    if (addr && addr !== want && domain !== home && !isSocEmail(addr)) {
      const before = text.slice(Math.max(0, match.index - 48), match.index).toLowerCase();
      const score = /from|sender|phish|suspicious|reported|reply-to|return-path/.test(before) ? 2 : 1;
      const existing = found.find((row) => row.addr === addr);
      if (existing) existing.score += score;
      else found.push({ addr, score });
    }
    match = re.exec(text);
  }
  found.sort((a, b) => b.score - a.score);
  const subjectMatch = text.match(/subject\**\s*[:=]\s*\**\s*([^\n|]{1,80})/i);
  const subject = subjectMatch ? subjectMatch[1].replace(/[|"\\]/g, '').trim().slice(0, 60) : '';
  return { senders: found.slice(0, 3).map((row) => row.addr), subject };
}

function alertSourceText() {
  const notes = asArray(state.thread).filter((note) => note && note.text);
  const early = notes.slice(0, 2).map((note) => note.text).join('\n');
  if (early.trim()) return early;
  return String((state.ticket && state.ticket.ticketContent) || '').slice(0, 6000);
}

function normName(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function tenantName(tenant) {
  return tenant.exoOrganizationName || tenant.graphTenantName || `Client ${tenant.clientNumber}`;
}

function pickTenant(company) {
  const want = normName(company);
  if (!want) return '';
  const ranked = state.tenants.map((tenant) => {
    const name = normName(tenantName(tenant));
    let score = 0;
    if (name && (name === want || name.includes(want) || want.includes(name))) score = 2;
    if (tenant.exchangeAuthenticated) score += 1;
    if (tenant.graphAuthenticated) score += 1;
    return { tenant, score };
  }).filter((row) => row.score >= 2);
  ranked.sort((a, b) => b.score - a.score);
  return ranked[0] ? String(ranked[0].tenant.clientNumber) : '';
}

function selectedTenant() {
  const id = $('tenantPick').value;
  return state.tenants.find((tenant) => String(tenant.clientNumber) === id) || null;
}

function showTicketAssignment(statusName, ownerMode) {
  if (!state.ticket) return;
  const memberId = (state.catalog && state.catalog.memberIdentifier) || 'cknospe';
  const memberName = (state.catalog && state.catalog.memberName) || memberId;
  state.ticket.statusName = statusName;
  const listed = state.tickets.find((ticket) => String(ticket.id) === String(state.ticket.ticketId));
  const bits = [state.ticket.companyName, statusName];
  if (ownerMode === 'member') {
    bits.push(`Owner ${memberName}`);
    if (listed) listed.owner = memberId;
  } else if (ownerMode === 'clear') {
    if (listed) listed.owner = '';
  } else if (listed && listed.owner) {
    bits.push(`Owner ${listed.owner}`);
  }
  if (state.ticket.contactName) bits.push(state.ticket.contactName);
  $('ticketMeta').textContent = bits.filter(Boolean).join(' · ');
  if (listed) listed.status = statusName;
  renderTickets();
}

const QUEUE_SETTINGS_KEY = 'eoaHunterQueue';
const QUEUE_SORTS = ['id', 'priority', 'company', 'summary', 'triage', 'status', 'owner', 'entered', 'age'];
const QUEUE_OWNERS = ['none', 'any', 'mine', 'owned'];
const QUEUE_STATUSES = ['New - Save 1st to send user Ticket Information'];

const HIDDEN_SUMMARIES = ['r-vulnerability scan & report', 'security foundations review'];

function isRoutineSummary(summary) {
  const text = String(summary || '').trim().toLowerCase();
  return HIDDEN_SUMMARIES.some((name) => text === name);
}

function defaultQueueSettings() {
  return { sortBy: 'id', sortOrder: 'desc', owner: 'none', status: '', keyword: '', hideRoutine: true, saved: [] };
}

function loadQueueSettings() {
  const base = defaultQueueSettings();
  try {
    const raw = JSON.parse(localStorage.getItem(QUEUE_SETTINGS_KEY) || '');
    if (!raw || typeof raw !== 'object') return base;
    const saved = asArray(raw.saved).filter((item) => item && item.name).map((item) => ({
      name: String(item.name),
      owner: QUEUE_OWNERS.includes(item.owner) ? item.owner : 'none',
      status: String(item.status || ''),
      keyword: String(item.keyword || ''),
      hideRoutine: item.hideRoutine !== false,
      sortBy: QUEUE_SORTS.includes(item.sortBy) ? item.sortBy : 'id',
      sortOrder: item.sortOrder === 'asc' ? 'asc' : 'desc',
    }));
    return {
      sortBy: QUEUE_SORTS.includes(raw.sortBy) ? raw.sortBy : 'id',
      sortOrder: raw.sortOrder === 'asc' ? 'asc' : 'desc',
      owner: QUEUE_OWNERS.includes(raw.owner) ? raw.owner : 'none',
      status: String(raw.status || ''),
      keyword: String(raw.keyword || ''),
      hideRoutine: raw.hideRoutine !== false,
      saved,
    };
  } catch {
    return base;
  }
}

function saveQueueSettings() {
  localStorage.setItem(QUEUE_SETTINGS_KEY, JSON.stringify(state.queue));
}

function applyQueueForm() {
  $('filterOwner').value = state.queue.owner;
  $('filterKeyword').value = state.queue.keyword;
  $('hideRoutine').checked = state.queue.hideRoutine !== false;
  fillStatusFilter();
  fillSavedFilters();
}

function fillStatusFilter() {
  const pick = $('filterStatus');
  const current = state.queue.status;
  const names = state.boardStatuses.length ? [...state.boardStatuses] : [...QUEUE_STATUSES];
  state.tickets.forEach((ticket) => {
    const status = String(ticket.status || '').trim();
    if (status && !names.some((name) => name.toLowerCase() === status.toLowerCase())) names.push(status);
  });
  if (current && !names.includes(current)) names.push(current);
  pick.innerHTML = '<option value="">Any status</option>' + names.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join('');
  pick.value = names.includes(current) ? current : '';
}

function fillSavedFilters(selected) {
  const pick = $('savedFilters');
  const name = selected == null ? pick.value : selected;
  pick.innerHTML = '<option value="">Choose a saved filter</option>' + state.queue.saved.map((item) => `<option value="${escapeHtml(item.name)}">${escapeHtml(item.name)}</option>`).join('');
  if ([...pick.options].some((opt) => opt.value === name)) pick.value = name;
}

function readQueueForm() {
  state.queue.owner = $('filterOwner').value;
  state.queue.status = $('filterStatus').value;
  state.queue.keyword = $('filterKeyword').value;
  state.queue.hideRoutine = $('hideRoutine').checked;
  saveQueueSettings();
  renderTickets();
}

function visibleTickets() {
  const settings = state.queue;
  const member = ((state.catalog && state.catalog.memberIdentifier) || 'cknospe').toLowerCase();
  const keyword = settings.keyword.trim().toLowerCase();
  const rows = state.tickets.filter((ticket) => {
    const owner = String(ticket.owner || '').trim();
    if (settings.owner === 'none' && owner) return false;
    if (settings.owner === 'owned' && !owner) return false;
    if (settings.owner === 'mine' && owner.toLowerCase() !== member) return false;
    if (settings.status && ticket.status !== settings.status) return false;
    if (settings.hideRoutine !== false && isRoutineSummary(ticket.summary)) return false;
    if (keyword) {
      const hay = [ticket.id, ticket.company, ticket.summary, ticket.status, ticket.owner, ticket.contact, ticket.priority].join(' ').toLowerCase();
      if (!hay.includes(keyword)) return false;
    }
    return true;
  });
  const dir = settings.sortOrder === 'asc' ? 1 : -1;
  rows.sort((a, b) => {
    if (settings.sortBy === 'id') return ((Number(a.id) || 0) - (Number(b.id) || 0)) * dir;
    if (settings.sortBy === 'age') return ((Number(a.ageDays) || 0) - (Number(b.ageDays) || 0)) * dir;
    if (settings.sortBy === 'priority') return ((Number(a.prioritySort) || 0) - (Number(b.prioritySort) || 0)) * dir;
    if (settings.sortBy === 'triage') return ((a.triaged ? 1 : 0) - (b.triaged ? 1 : 0)) * dir;
    const key = settings.sortBy === 'entered' ? 'entered' : settings.sortBy;
    return String(a[key] || '').localeCompare(String(b[key] || '')) * dir;
  });
  return rows;
}

function paintSortHeaders() {
  document.querySelectorAll('button.sort').forEach((btn) => {
    const active = btn.dataset.sort === state.queue.sortBy;
    btn.classList.toggle('active', active);
    const mark = active ? (state.queue.sortOrder === 'asc' ? ' ▲' : ' ▼') : '';
    btn.textContent = `${btn.dataset.label}${mark}`;
  });
}

function updateQueueCount() {
  const board = (state.catalog && state.catalog.boardName) || 'Security';
  const visible = visibleTickets();
  const shown = visible.length;
  const triaged = visible.filter((ticket) => ticket.triaged).length;
  const bits = [`${shown} of ${state.tickets.length} open ticket(s) on ${board}`];
  if (triaged) bits.push(`${triaged} triaged`);
  if (state.queue.owner === 'none') bits.push('no owner');
  else if (state.queue.owner === 'mine') bits.push('mine');
  else if (state.queue.owner === 'owned') bits.push('has an owner');
  if (state.queue.status) bits.push(state.queue.status);
  if (state.queue.keyword.trim()) bits.push(`“${state.queue.keyword.trim()}”`);
  if (state.queue.hideRoutine !== false) bits.push('scan and foundations tickets hidden');
  setStatus('queueStatus', bits.join(' · '));
}

function paintPickVisible() {
  const box = $('pickVisible');
  if (!box) return;
  const visible = visibleTickets();
  const picked = visible.filter((ticket) => state.picked.has(String(ticket.id))).length;
  box.checked = visible.length > 0 && picked === visible.length;
  box.indeterminate = picked > 0 && picked < visible.length;
}

function pickTicketRange(id, on) {
  const ids = visibleTickets().map((ticket) => String(ticket.id));
  const current = ids.indexOf(String(id));
  const anchor = ids.indexOf(String(state.lastPickId || ''));
  if (current < 0) return;
  const start = anchor < 0 ? current : Math.min(current, anchor);
  const end = anchor < 0 ? current : Math.max(current, anchor);
  for (let index = start; index <= end; index += 1) {
    if (on) state.picked.add(ids[index]);
    else state.picked.delete(ids[index]);
  }
  state.lastPickId = String(id);
  document.querySelectorAll('.pickTicket').forEach((box) => {
    box.checked = state.picked.has(box.dataset.id);
  });
  paintPickVisible();
  paintEnrichButton();
}

function isIncidentTicket(ticket) {
  return String((ticket && ticket.type) || '').trim().toLowerCase() === 'incident';
}

function parentIdOf(ticket) {
  const id = parseInt(ticket && ticket.parentId, 10);
  return id > 0 ? String(id) : '';
}

function enrichTargets() {
  const ready = [];
  const readyIds = new Set();
  let triaged = 0;
  let other = 0;
  let children = 0;
  const childParents = [];
  const addReady = (ticket) => {
    const id = String(ticket.id);
    if (readyIds.has(id)) return;
    ready.push(ticket);
    readyIds.add(id);
  };
  for (const id of state.picked) {
    const ticket = state.tickets.find((row) => String(row.id) === String(id));
    if (ticket && ticket.triaged) {
      triaged += 1;
      continue;
    }
    if (!ticket || !isIncidentTicket(ticket)) {
      other += 1;
      continue;
    }
    const parentId = parentIdOf(ticket);
    if (parentId) {
      children += 1;
      childParents.push(parentId);
      continue;
    }
    addReady(ticket);
  }
  let parentMissing = 0;
  const seenParents = new Set();
  for (const parentId of childParents) {
    if (seenParents.has(parentId) || readyIds.has(parentId)) {
      seenParents.add(parentId);
      continue;
    }
    seenParents.add(parentId);
    const parent = state.tickets.find((row) => String(row.id) === parentId);
    if (parent && isIncidentTicket(parent) && !parent.triaged) addReady(parent);
    else if (!parent || !isIncidentTicket(parent)) parentMissing += 1;
  }
  return { ready, triaged, other, children, parentMissing };
}

function paintEnrichButton() {
  const button = $('btnEnrich');
  if (!button) return;
  const { ready } = enrichTargets();
  button.textContent = ready.length ? `Add alert enrichment (${ready.length})` : 'Add alert enrichment';
  button.title = 'Posts an alert review on Incident tickets that are not already triaged. Child tickets are left to the parent.';
}

function alertTextFrom(content, thread) {
  const parts = [];
  const body = String(content || '').trim();
  if (body) parts.push(body);
  for (const note of asArray(thread)) {
    if (!note || !note.text) continue;
    if (note.kind === 'internal' || note.kind === 'resolution') continue;
    if (isComposedClientEmail(note) || isIncidentText(note.text)) continue;
    const text = String(note.text).trim();
    if (!text || parts.some((part) => part === text || part.includes(text))) continue;
    parts.push(text);
  }
  return parts.join('\n\n').trim();
}

async function loadAlertSource(ticketId) {
  const ticket = await api('/api/manage/ticket', {
    method: 'POST',
    body: JSON.stringify({ ticketId: String(ticketId) }),
  }, 90000);
  let thread = [];
  try {
    const data = await api('/api/manage/hunter-post', {
      method: 'POST',
      body: JSON.stringify({ ticketId: String(ticketId), status: 'ticket-thread' }),
    }, 60000);
    thread = asArray(data && data.notes && data.notes.thread);
  } catch {
    thread = [];
  }
  return { ticket, alertText: alertTextFrom(ticket && ticket.ticketContent, thread) };
}

async function draftAlertNote(ticket, alertText) {
  const data = await api('/api/triage/memberberry-note', {
    method: 'POST',
    body: JSON.stringify({
      review: alertText,
      ticket: alertText,
      ticketNumber: String((ticket && (ticket.ticketId || ticket.id)) || ''),
      company: (ticket && (ticket.companyName || ticket.company)) || '',
      contact: (ticket && (ticket.contactName || ticket.contact)) || '',
      audience: 'alert',
      actions: [],
    }),
  }, 180000);
  const note = (data.result && (data.result.internalText || data.result.text)) || '';
  if (!note.trim()) throw new Error('The alert review came back empty.');
  return { note, references: (data.result && data.result.references) || '' };
}

async function postAlertNote(ticketId, note) {
  const data = await api('/api/manage/hunter-post', {
    method: 'POST',
    body: JSON.stringify({
      ticketId: String(ticketId),
      status: 'keep',
      includeRecipients: false,
      time: 'none',
      notes: { client: '', internal: note, incident: '' },
    }),
  }, 120000);
  const errors = asArray(data && data.errors).filter(Boolean);
  if ((data && data.ok === false) || errors.length) {
    throw new Error(errors.join(' ') || 'Manage did not post the note.');
  }
  return data;
}

function showEnrichLog(lines) {
  const box = $('enrichLog');
  if (!box) return;
  box.hidden = !lines.length;
  box.textContent = lines.join('\n');
}

async function enrichSelected() {
  if (!state.picked.size) throw new Error('Select one or more tickets in the list.');
  const { ready, triaged, other, children, parentMissing } = enrichTargets();
  const skipped = [];
  if (triaged) skipped.push(`${triaged} already triaged`);
  if (other) skipped.push(`${other} not an Incident`);
  if (children) skipped.push(`${children} child ticket(s)`);
  if (parentMissing) skipped.push(`${parentMissing} child ticket(s) whose parent is not an open Incident here`);
  if (!ready.length) {
    throw new Error(skipped.length ? `Nothing to triage. ${skipped.join(', ')}.` : 'Nothing to triage.');
  }
  const ids = ready.map((ticket) => String(ticket.id));
  const listed = ids.length > 20 ? `${ids.slice(0, 20).join(', ')}, and ${ids.length - 20} more` : ids.join(', ');
  const prompt = [
    `Add a Hatz-AI alert review as an internal note on ${ids.length} Incident ticket(s)?`,
    '',
    listed,
    '',
  ];
  if (skipped.length) prompt.push(`Skipped: ${skipped.join(', ')}.`, '');
  prompt.push(
    'Each note says it is an alert-only review and that you still need to investigate.',
    'This does not pull logs, email the client, or change the ticket status.',
  );
  const ok = window.confirm(prompt.join('\n'));
  if (!ok) return;
  const lines = [];
  if (skipped.length) lines.push(`Skipped: ${skipped.join(', ')}.`);
  let posted = 0;
  for (let index = 0; index < ids.length; index += 1) {
    const id = ids[index];
    setStatus('queueStatus', `Alert enrichment ${index + 1} of ${ids.length}: ticket ${id}…`);
    try {
      const source = await loadAlertSource(id);
      if (!source.alertText) throw new Error('no alert text on the ticket.');
      const drafted = await draftAlertNote(source.ticket, source.alertText);
      await postAlertNote(id, drafted.note);
      const row = state.tickets.find((ticket) => String(ticket.id) === String(id));
      if (row) row.triaged = true;
      posted += 1;
      state.picked.delete(String(id));
      lines.push(`#${id}: internal note posted. Status unchanged.`);
    } catch (error) {
      lines.push(`#${id}: ${error.message}`);
    }
    showEnrichLog(lines);
    renderTickets();
  }
  const failed = ids.length - posted;
  setStatus('queueStatus', `Alert enrichment finished. ${posted} posted, ${failed} not posted.`, failed ? 'warn' : '');
}

function renderTickets() {
  const rows = $('ticketRows');
  const selected = state.ticket && String(state.ticket.ticketId);
  const visible = visibleTickets();
  rows.innerHTML = visible.map((ticket) => `<tr class="ticket${selected === String(ticket.id) ? ' selected' : ''}" data-id="${escapeHtml(ticket.id)}">
    <td><input type="checkbox" class="pickTicket" data-id="${escapeHtml(ticket.id)}"${state.picked.has(String(ticket.id)) ? ' checked' : ''} /></td>
    <td>${escapeHtml(ticket.id)}</td>
    <td><span class="pri"><i style="background:${escapeHtml(ticket.priorityColor || '#6c757d')}"></i>${escapeHtml(shortPriority(ticket.priority))}</span></td>
    <td title="${escapeHtml(ticket.company || '')}">${escapeHtml(ticket.company || '')}</td>
    <td title="${escapeHtml(ticket.summary || '')}">${escapeHtml(ticket.summary || '')}</td>
    <td>${ticket.triaged ? '<span class="triageMark" title="An alert triage note is already on this ticket">Triaged</span>' : ''}</td>
    <td>${escapeHtml(ticket.status || '')}</td>
    <td>${escapeHtml(ticket.owner || '')}</td>
    <td>${escapeHtml(formatEntered(ticket.entered))}</td>
    <td>${ticket.ageDays == null || ticket.ageDays === '' ? '' : escapeHtml(`${ticket.ageDays} d`)}</td>
  </tr>`).join('');
  rows.querySelectorAll('tr.ticket').forEach((row) => {
    row.addEventListener('mousedown', (event) => {
      if (event.shiftKey && !event.target.closest('.pickTicket')) event.preventDefault();
    });
    row.addEventListener('click', (event) => {
      if (event.shiftKey) {
        pickTicketRange(row.dataset.id, true);
        return;
      }
      state.lastPickId = row.dataset.id;
      loadTicket(row.dataset.id).catch((e) => setStatus('queueStatus', e.message, 'bad'));
    });
  });
  rows.querySelectorAll('.pickTicket').forEach((box) => {
    box.addEventListener('click', (event) => {
      event.stopPropagation();
      if (event.shiftKey) {
        pickTicketRange(box.dataset.id, box.checked);
        return;
      }
      if (box.checked) state.picked.add(box.dataset.id);
      else state.picked.delete(box.dataset.id);
      state.lastPickId = box.dataset.id;
      paintPickVisible();
      paintEnrichButton();
    });
  });
  paintPickVisible();
  paintEnrichButton();
  paintSortHeaders();
  fillStatusFilter();
  updateQueueCount();
}

function sortQueue(key) {
  if (state.queue.sortBy === key) {
    state.queue.sortOrder = state.queue.sortOrder === 'asc' ? 'desc' : 'asc';
  } else {
    state.queue.sortBy = key;
    state.queue.sortOrder = (key === 'id' || key === 'age' || key === 'entered' || key === 'triage') ? 'desc' : 'asc';
  }
  saveQueueSettings();
  renderTickets();
}

function saveNamedFilter() {
  const name = $('filterName').value.trim();
  if (!name) throw new Error('Enter a name for this filter.');
  const entry = {
    name,
    owner: state.queue.owner,
    status: state.queue.status,
    keyword: state.queue.keyword,
    hideRoutine: state.queue.hideRoutine !== false,
    sortBy: state.queue.sortBy,
    sortOrder: state.queue.sortOrder,
  };
  state.queue.saved = state.queue.saved.filter((item) => item.name.toLowerCase() !== name.toLowerCase());
  state.queue.saved.push(entry);
  state.queue.saved.sort((a, b) => a.name.localeCompare(b.name));
  saveQueueSettings();
  fillSavedFilters(name);
  setStatus('queueStatus', `Saved filter “${name}”.`);
}

function applySavedFilter(name) {
  const entry = state.queue.saved.find((item) => item.name === name);
  if (!entry) return;
  state.queue.owner = entry.owner;
  state.queue.status = entry.status;
  state.queue.keyword = entry.keyword;
  state.queue.hideRoutine = entry.hideRoutine !== false;
  state.queue.sortBy = entry.sortBy;
  state.queue.sortOrder = entry.sortOrder;
  $('filterName').value = entry.name;
  applyQueueForm();
  $('savedFilters').value = entry.name;
  saveQueueSettings();
  renderTickets();
}

function deleteSavedFilter() {
  const name = $('savedFilters').value;
  if (!name) throw new Error('Choose a saved filter to delete.');
  state.queue.saved = state.queue.saved.filter((item) => item.name !== name);
  saveQueueSettings();
  fillSavedFilters('');
  if ($('filterName').value === name) $('filterName').value = '';
  setStatus('queueStatus', `Deleted saved filter “${name}”.`);
}

function shortPriority(name) {
  const text = String(name || '').trim();
  const match = text.match(/(\d+)\s*$/);
  return match ? `P${match[1]}` : text;
}

function parseManageTime(value) {
  const text = String(value || '').trim();
  if (!text) return null;
  if (/[zZ]$|[+-]\d{2}:?\d{2}$/.test(text)) {
    const dated = new Date(text);
    return Number.isNaN(dated.getTime()) ? null : dated;
  }
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?/);
  if (iso) {
    const dated = new Date(Date.UTC(+iso[1], +iso[2] - 1, +iso[3], +iso[4], +iso[5], +(iso[6] || 0)));
    return Number.isNaN(dated.getTime()) ? null : dated;
  }
  const us = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*([AaPp][Mm]))?)?$/);
  if (us) {
    let hour = us[4] ? parseInt(us[4], 10) : 0;
    const minute = us[5] ? parseInt(us[5], 10) : 0;
    const second = us[6] ? parseInt(us[6], 10) : 0;
    const ampm = us[7] || '';
    if (/p/i.test(ampm) && hour < 12) hour += 12;
    if (/a/i.test(ampm) && hour === 12) hour = 0;
    const dated = new Date(Date.UTC(+us[3], +us[1] - 1, +us[2], hour, minute, second));
    return Number.isNaN(dated.getTime()) ? null : dated;
  }
  const dated = new Date(text);
  return Number.isNaN(dated.getTime()) ? null : dated;
}

function formatClock(date) {
  return date.toLocaleString(undefined, {
    month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit', timeZoneName: 'short',
  });
}

function formatEntered(value) {
  const date = parseManageTime(value);
  if (!date) return value ? String(value) : '';
  return formatClock(date);
}

const MONTHS = {
  january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
  july: 6, august: 7, september: 8, october: 9, november: 10, december: 11,
};

function localTimeText(text) {
  return String(text || '')
    .replace(/\b(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2}:\d{2}(?:\.\d+)?)(Z)?\b/g, (all, day, clock, zone) => {
      const date = new Date(`${day}T${clock}${zone || 'Z'}`);
      return Number.isNaN(date.getTime()) ? all : formatClock(date);
    })
    .replace(/\b(?:(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+)?([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4}),?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])\s+UTC\b/g, (all, month, day, year, hour, minute, second, ampm) => {
      const monthIndex = MONTHS[month.toLowerCase()];
      if (monthIndex == null) return all;
      let h = parseInt(hour, 10);
      if (/p/i.test(ampm) && h < 12) h += 12;
      if (/a/i.test(ampm) && h === 12) h = 0;
      const date = new Date(Date.UTC(+year, monthIndex, +day, h, +minute, +(second || 0)));
      return Number.isNaN(date.getTime()) ? all : formatClock(date);
    });
}

async function loadBoardStatuses() {
  try {
    const data = await api('/api/manage/hunter-post', {
      method: 'POST',
      body: JSON.stringify({ ticketId: '0', status: 'board-statuses' }),
    }, 60000);
    state.boardStatuses = asArray(data && data.notes && data.notes.statuses).map((name) => String(name || '').trim()).filter(Boolean);
    state.statusChoices = asArray(data && data.notes && data.notes.choices).map((choice) => ({
      id: Number(choice.id),
      name: String(choice.name || '').trim(),
      closed: choice.closed === true || choice.closed === 'True',
    })).filter((choice) => choice.id && choice.name);
    fillTicketStatus();
  } catch {
    state.boardStatuses = [];
  }
}

async function refreshQueue() {
  setStatus('queueStatus', 'Loading Security tickets…');
  const data = await api('/api/manage/security-tickets', {}, 90000);
  state.catalog = data.catalog || {};
  state.tickets = asArray(data.tickets);
  await loadBoardStatuses();
  const member = state.catalog.memberIdentifier || 'cknospe';
  $('memberLabel').textContent = member;
  renderTickets();
}

async function refreshTenants() {
  const data = await api('/api/session');
  state.tenants = asArray(data && data.tenants);
  const pick = $('tenantPick');
  const previous = pick.value;
  if (!state.tenants.length) {
    pick.innerHTML = '<option value="">No tenant yet</option>';
    syncTenantAuth();
    return;
  }
  pick.innerHTML = state.tenants.map((tenant) => {
    const auth = `${tenant.graphAuthenticated ? 'Graph' : 'no Graph'}, ${tenant.exchangeAuthenticated ? 'Exchange' : 'no Exchange'}`;
    return `<option value="${escapeHtml(tenant.clientNumber)}">${escapeHtml(tenantName(tenant))} (${escapeHtml(auth)})</option>`;
  }).join('');
  const company = state.ticket && state.ticket.companyName;
  const matched = pickTenant(company);
  if (previous && [...pick.options].some((opt) => opt.value === previous)) pick.value = previous;
  else if (matched) pick.value = matched;
  syncTenantAuth();
}

function syncTenantAuth() {
  const empty = state.tenants.length === 0;
  if ($('tenantEmpty')) $('tenantEmpty').hidden = !empty;
  if ($('btnAddTenant')) $('btnAddTenant').hidden = !empty;
  const tenant = selectedTenant();
  if ($('btnExoAuth')) $('btnExoAuth').hidden = !(tenant && !tenant.exchangeAuthenticated);
  if ($('btnGraphAuth')) $('btnGraphAuth').hidden = !(tenant && tenant.exchangeAuthenticated && !tenant.graphAuthenticated);
}

function fillMailboxPick(addresses, preferred) {
  const pick = $('mailboxPick');
  pick.innerHTML = '<option value="">Choose an address</option>' + addresses.map((addr) => `<option value="${escapeHtml(addr)}">${escapeHtml(addr)}</option>`).join('');
  if (preferred && addresses.includes(preferred)) pick.value = preferred;
}

function recipientLine(ticket) {
  const parts = [];
  if (ticket.contactEmail) {
    parts.push(`Contact ${ticket.contactEmail} (${ticket.emailContact ? 'email on' : 'email off'})`);
  } else if (ticket.contactName) {
    parts.push(`Contact ${ticket.contactName} (no email address)`);
  }
  if (ticket.automaticEmailCc) {
    parts.push(`CC ${ticket.automaticEmailCc} (${ticket.emailCc ? 'email on' : 'email off'})`);
  }
  if (!parts.length) return 'This ticket has no client email recipients.';
  const canSend = Boolean((ticket.emailContact && ticket.contactEmail) || (ticket.emailCc && ticket.automaticEmailCc));
  if (!canSend) parts.push('Manage will not send mail until the contact or CC email flag is on.');
  return parts.join(' · ');
}

async function claimTicket() {
  if (!state.ticket) return;
  setStatus('scanStatus', 'Setting the ticket to In Progress and assigning you as owner and resource…');
  try {
    const posted = await api('/api/manage/hunter-post', {
      method: 'POST',
      body: JSON.stringify({ ticketId: state.ticket.ticketId, status: 'progress' }),
    }, 60000);
    if (posted && posted.statusUpdated) showTicketAssignment(posted.statusName || 'In Progress', 'member');
    if (posted && posted.ok === false) {
      const detail = asArray(posted.errors).filter(Boolean).join(' ');
      setStatus('scanStatus', detail || 'Manage did not update the ticket.', 'warn');
      return;
    }
    setStatus('scanStatus', 'Ticket is In Progress and assigned to you.');
  } catch (e) {
    setStatus('scanStatus', `Could not update the ticket. ${e.message}`, 'warn');
  }
  await refreshManageLog(state.ticket.ticketId);
}

async function loadTicket(ticketId) {
  setStatus('queueStatus', `Loading ticket ${ticketId}…`);
  const data = await api('/api/manage/ticket', {
    method: 'POST',
    body: JSON.stringify({ ticketId }),
  }, 90000);
  state.ticket = data;
  state.review = '';
  state.audit = null;
  state.span = null;
  state.pendingRepull = null;
  $('ticketCard').hidden = false;
  $('workBar').hidden = false;
  $('notesCard').hidden = true;
  $('review').value = '';
  $('useClient').checked = false;
  $('useInternal').checked = false;
  $('useIncident').checked = false;
  $('btnNotes').hidden = true;
  $('btnReviseReview').hidden = true;
  $('reviewRevision').value = '';
  $('notesRevision').value = '';
  if ($('ticketSetup')) $('ticketSetup').open = true;
  showOutput('ticket', { keepSetup: true });
  $('btnRepull').hidden = true;
  $('btnReviewAgain').hidden = true;
  const ticketLabel = `Ticket #${data.ticketId}${data.summary ? ` — ${data.summary}` : ''}`;
  $('ticketTitle').textContent = 'Mailbox and status';
  $('workBarTitle').textContent = ticketLabel;
  fillTicketStatus();
  const bits = [data.companyName, data.statusName, data.contactName].filter(Boolean);
  $('ticketMeta').textContent = bits.join(' · ');
  state.thread = [];
  state.threadKind = 'discussion';
  renderThread();
  await loadThread(data.ticketId);
  $('userMatches').innerHTML = '';
  const fromAlert = mailboxesInAlert(alertSourceText());
  const rest = emailsIn(data.ticketContent || '').filter((addr) => !fromAlert.includes(addr));
  const contact = String(data.contactEmail || '').trim().toLowerCase();
  if (contact && !isSocEmail(contact) && !fromAlert.includes(contact) && !rest.includes(contact)) rest.push(contact);
  const preferred = fromAlert[0] || '';
  fillMailboxPick([...fromAlert, ...rest].slice(0, 40), preferred);
  $('mailbox').value = preferred;
  $('userSearch').value = preferred || data.contactName || '';
  setStatus('userStatus', preferred ? `Mailbox from the alert: ${preferred}` : 'No mailbox address in the alert.');
  const canEmail = Boolean((data.emailContact && data.contactEmail) || (data.emailCc && data.automaticEmailCc));
  $('includeRecipients').checked = canEmail;
  $('recipientList').textContent = recipientLine(data);
  await refreshTenants();
  renderTickets();
  setStatus('queueStatus', `Loaded ticket #${data.ticketId}.`);
  await claimTicket();
}

function linkify(text) {
  return String(text || '').split(/(https?:\/\/[^\s]+)/g).map((part) => {
    if (!/^https?:\/\//.test(part)) return escapeHtml(localTimeText(part));
    const clean = part.replace(/[),.;]+$/, '');
    const trail = part.slice(clean.length);
    return `<a href="${escapeHtml(clean)}" target="_blank" rel="noopener noreferrer">${escapeHtml(clean)}</a>${escapeHtml(trail)}`;
  }).join('');
}

function isIncidentText(text) {
  const body = String(text || '').trim();
  return /^subject:\s*incident report\b/i.test(body) || /\bINCIDENT SUMMARY\b/.test(body);
}

function isComposedClientEmail(note) {
  if (!note || note.kind !== 'discussion') return false;
  const body = String(note.text || '').trim();
  if (isIncidentText(body)) return false;
  if (/^subject:/i.test(body)) return true;
  return /^(good morning|good afternoon|good evening|hello)\b/i.test(body);
}

function splitThread() {
  const notes = asArray(state.thread);
  const incident = notes.filter((note) => note.kind === 'discussion' && isIncidentText(note.text));
  const client = notes.filter((note) => isComposedClientEmail(note));
  return {
    ticket: notes.filter((note) => note.kind === 'description' || note.kind === 'resolution' || note.kind === 'other' || (note.kind === 'discussion' && !incident.includes(note) && !client.includes(note))),
    client,
    internal: notes.filter((note) => note.kind === 'internal'),
    incident,
  };
}

function renderNoteHtml(note, index) {
  const when = note.when ? formatEntered(note.when) : '';
  const who = note.kind === 'description' ? (note.who || 'Initial description') : (note.who || 'Unknown');
  const kindLabel = { description: 'Description', discussion: 'Discussion', internal: 'Internal', resolution: 'Resolution', other: 'Note' }[note.kind] || '';
  const meta = [kindLabel, who, when].filter(Boolean).join(' · ');
  const time = note.kind === 'discussion' && note.id
    ? `<button type="button" class="noteTimeBtn">Add time</button>`
    : '';
  const remove = note.id ? `<button type="button" class="noteDeleteBtn">Delete</button>` : '';
  const actions = (time || remove) ? `<div class="noteActions">${time}${remove}</div>` : '';
  const timePanel = note.kind === 'discussion' && note.id
    ? `<div class="noteTime" hidden>
        <label>Time
          <select class="noteTimeType">
            <option value="remote">Remote Consulting (billable)</option>
            <option value="communication">Communication (not billable)</option>
          </select>
        </label>
        <label>Start <input type="datetime-local" class="noteTimeStart" /></label>
        <label>Hours <input type="number" class="noteTimeHours" min="0.25" max="24" step="0.25" value="0.25" style="min-width:5rem" /></label>
        <button type="button" class="noteTimeAdd">Add time to this note</button>
        <span class="noteTimeStatus status"></span>
      </div>`
    : '';
  const kindClass = note.kind ? ` kind-${note.kind}` : '';
  return `<div class="threadNote${kindClass}" data-note-id="${escapeHtml(note.id || '')}" data-note-index="${index}"><div class="meta">${escapeHtml(meta)}</div><div class="body">${linkify(note.text)}</div>${actions}${timePanel}</div>`;
}

function renderNoteList(box, notes, emptyText) {
  if (!box) return;
  box.classList.toggle('empty', notes.length === 0);
  box.innerHTML = notes.length
    ? notes.map((note, index) => renderNoteHtml(note, index)).join('')
    : `<p class="muted">${emptyText}</p>`;
}

function renderThread() {
  const split = splitThread();
  const ticketNotes = split.ticket.slice();
  const body = state.ticket ? String(state.ticket.ticketContent || '').trim() : '';
  const hasDescription = ticketNotes.some((note) => note.kind === 'description');
  const already = ticketNotes.some((note) => String(note.text || '').trim() === body);
  if (body && !hasDescription && !already) {
    ticketNotes.unshift({ kind: 'description', who: 'Initial description', when: '', text: body, id: '' });
  }
  renderNoteList($('ticketThread'), ticketNotes, 'No original ticket text.');
  renderNoteList($('clientNotes'), split.client, 'No client email on this ticket yet.');
  renderNoteList($('internalNotes'), split.internal, 'No internal note on this ticket yet.');
  renderNoteList($('incidentNotes'), split.incident, 'No incident report on this ticket yet.');
  const counts = {
    ticket: ticketNotes.length,
    review: (($('review') && $('review').value) || '').trim() ? 1 : 0,
    client: split.client.length,
    internal: split.internal.length,
    incident: split.incident.length,
  };
  const labels = {
    ticket: 'Ticket',
    review: 'Review',
    client: 'Client email',
    internal: 'Internal note',
    incident: 'Incident report',
  };
  document.querySelectorAll('.outputTab').forEach((btn) => {
    const name = btn.dataset.output;
    btn.textContent = `${labels[name] || name} (${counts[name] || 0})`;
  });
}

async function addTimeToNote(button) {
  if (!state.ticket) throw new Error('Choose a ticket first.');
  const card = button.closest('.threadNote');
  const noteId = card && card.dataset.noteId;
  if (!noteId) throw new Error('Choose a discussion note.');
  const type = card.querySelector('.noteTimeType').value;
  const hours = card.querySelector('.noteTimeHours').value;
  const start = card.querySelector('.noteTimeStart').value;
  const typeLabel = card.querySelector('.noteTimeType').selectedOptions[0].textContent;
  const who = (card.querySelector('.meta').textContent || '').trim();
  const ok = window.confirm([
    `Add ${hours} h of ${typeLabel} to this discussion note?`,
    who,
    `Starting ${start.replace('T', ' ')}.`,
    '',
    'This does not post a new note and does not change the status.',
  ].join('\n'));
  if (!ok) return;
  const status = card.querySelector('.noteTimeStatus');
  status.textContent = 'Adding time…';
  status.className = 'noteTimeStatus status';
  const data = await api('/api/manage/hunter-post', {
    method: 'POST',
    body: JSON.stringify({
      ticketId: state.ticket.ticketId,
      status: 'add-time',
      noteId,
      time: type,
      hours: Number(hours),
      timeStart: start,
    }),
  }, 60000);
  const done = asArray(data.completed).filter(Boolean);
  const errors = asArray(data.errors).filter(Boolean);
  status.textContent = [...done, ...errors].join(' ') || 'Time added.';
  if (errors.length) status.classList.add('bad');
  if (errors.length) throw new Error(status.textContent);
  await refreshManageLog(state.ticket.ticketId);
}

async function deleteNote(button) {
  if (!state.ticket) throw new Error('Choose a ticket first.');
  const card = button.closest('.threadNote');
  const noteId = card && card.dataset.noteId;
  if (!noteId) throw new Error('That entry is not a note that can be deleted.');
  const who = (card.querySelector('.meta').textContent || '').trim();
  const preview = (card.querySelector('.body').textContent || '').trim().split('\n').find((line) => line.trim()) || '';
  const ok = window.confirm([
    `Delete this note from ticket #${state.ticket.ticketId}?`,
    who,
    preview.length > 140 ? `${preview.slice(0, 140)}…` : preview,
    '',
    'This removes the note from ConnectWise. It does not change the status or remove a time entry.',
    'If the note was already emailed, deleting it does not unsend that email.',
  ].join('\n'));
  if (!ok) return;
  const data = await api('/api/manage/hunter-post', {
    method: 'POST',
    body: JSON.stringify({
      ticketId: state.ticket.ticketId,
      status: 'delete-note',
      noteId,
    }),
  }, 60000);
  const errors = asArray(data && data.errors).filter(Boolean);
  if ((data && data.ok === false) || errors.length) {
    throw new Error(errors.join(' ') || 'Manage did not delete the note.');
  }
  await loadThread(state.ticket.ticketId);
  const stillTriaged = asArray(state.thread).some((note) => String(note.text || '').includes('HATZ-AI POWERED REVIEW'));
  const listed = state.tickets.find((ticket) => String(ticket.id) === String(state.ticket.ticketId));
  if (listed) listed.triaged = stillTriaged;
  renderTickets();
  setStatus('queueStatus', (asArray(data.completed).filter(Boolean).join(' ')) || `Deleted note ${noteId}.`);
}

function renderManageLog() {
  const box = $('manageLogBody');
  if (!box) return;
  const rows = asArray(state.manageWrites);
  if (!rows.length) {
    box.textContent = 'No Manage writes from this hunter yet.';
    return;
  }
  box.textContent = rows.map((row) => {
    const when = row.at ? formatEntered(row.at) : '';
    const outcome = row.outcome === 'failed' ? 'failed' : 'ok';
    return [when, outcome, row.summary].filter(Boolean).join(' · ');
  }).join('\n');
}

async function loadThread(ticketId) {
  const data = await api('/api/manage/hunter-post', {
    method: 'POST',
    body: JSON.stringify({ ticketId: String(ticketId), status: 'ticket-thread' }),
  }, 60000);
  state.thread = asArray(data && data.notes && data.notes.thread);
  state.manageWrites = asArray(data && data.notes && data.notes.audit);
  renderThread();
  renderManageLog();
  const errors = asArray(data && data.errors).filter(Boolean);
  if (errors.length && !state.thread.length) {
    setStatus('queueStatus', `Ticket notes: ${errors.join(' ')}`, 'warn');
  }
}

async function refreshManageLog(ticketId) {
  try {
    const data = await api('/api/manage/hunter-post', {
      method: 'POST',
      body: JSON.stringify({ ticketId: String(ticketId), status: 'manage-log' }),
    }, 20000);
    state.manageWrites = asArray(data && data.notes && data.notes.audit);
    renderManageLog();
  } catch {
    /* The Manage write already finished. A missing local log line is not a failed check. */
  }
}

function insertClientLine() {
  const id = $('clientInsert').value;
  $('clientInsert').value = '';
  const text = CLIENT_INSERTS[id];
  if (!text) return;
  const box = $('clientText');
  if (!box.value.includes(text)) {
    box.value = box.value.trim() ? `${box.value.trim()}\n\n${text}` : text;
  }
  $('useClient').checked = true;
}

function windowFromTicket(text) {
  const raw = String(text || '');
  const iso = raw.match(/\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(?::\d{2})?(?:Z|[+-]\d{2}:?\d{2})?/);
  const us = raw.match(/\b(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})\s*([AaPp][Mm])?)?/);
  let event = null;
  if (iso) {
    const parsed = new Date(iso[0].includes('T') || iso[0].endsWith('Z') ? iso[0] : iso[0].replace(' ', 'T'));
    if (!Number.isNaN(parsed.getTime())) event = parsed;
  } else if (us) {
    let hour = us[4] ? parseInt(us[4], 10) : 0;
    const minute = us[5] ? parseInt(us[5], 10) : 0;
    const ampm = us[6] || '';
    if (/p/i.test(ampm) && hour < 12) hour += 12;
    if (/a/i.test(ampm) && hour === 12) hour = 0;
    event = new Date(parseInt(us[3], 10), parseInt(us[1], 10) - 1, parseInt(us[2], 10), hour, minute, 0);
  }
  const now = new Date();
  let days = 2;
  if (event && !Number.isNaN(event.getTime())) {
    const start = new Date(event.getTime() - 24 * 60 * 60 * 1000);
    days = Math.max(2, Math.min(90, Math.ceil((now.getTime() - start.getTime()) / (24 * 60 * 60 * 1000))));
  }
  const startUtc = new Date(now.getTime() - days * 24 * 60 * 60 * 1000);
  return { days, start: startUtc.toISOString(), end: now.toISOString(), hadTime: Boolean(event && !Number.isNaN(event.getTime())) };
}

function formatLocal(value) {
  const parsed = Date.parse(String(value || ''));
  if (!Number.isFinite(parsed)) return String(value || '');
  const d = new Date(parsed);
  const pad = (n) => String(n).padStart(2, '0');
  let hour = d.getHours();
  const ampm = hour >= 12 ? 'PM' : 'AM';
  hour = hour % 12 || 12;
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${hour}:${pad(d.getMinutes())} ${ampm}`;
}

function localZone() {
  try {
    const parts = new Intl.DateTimeFormat('en-US', { timeZoneName: 'long' }).formatToParts(new Date());
    const name = parts.find((p) => p.type === 'timeZoneName');
    if (name && name.value && !/^(?:UTC|GMT)\b/i.test(name.value)) return name.value;
  } catch { /* keep the fallback */ }
  return 'local time';
}

function inWindow(value, startIso) {
  const text = String(value || '').trim();
  if (!text || !startIso) return null;
  const when = Date.parse(text);
  const start = Date.parse(startIso);
  if (Number.isNaN(when) || Number.isNaN(start)) return null;
  return when >= start;
}

function flagOn(value) {
  return value === true || value === 'True';
}

function factsFor(mailbox, audit, windowNote) {
  const mem = state.memory || {};
  const want = String(mailbox || '').trim().toLowerCase();
  const rules = asArray(mem.rules).map((rule) => ({
    name: rule.Name || rule.Identity || '',
    hidden: flagOn(rule.Hidden),
    redirectTo: rule.RedirectTo || '',
    forwardTo: rule.ForwardTo || '',
    forwardAsAttachmentTo: rule.ForwardAsAttachmentTo || '',
    deleteMessage: flagOn(rule.DeleteMessage),
    description: String(rule.Description || '').slice(0, 180),
  }));
  const forwards = [];
  const delegates = [];
  asArray(mem.mailbox).forEach((mbx) => {
    const upn = String(mbx.UserPrincipalName || '').trim().toLowerCase();
    if (want && upn && upn !== want) return;
    asArray(mbx.Forwards).forEach((row) => { if (row.Address) forwards.push({ address: row.Address, field: row.Field || '' }); });
    if (!asArray(mbx.Forwards).length) {
      if (mbx.ForwardingSmtpAddress) forwards.push({ address: mbx.ForwardingSmtpAddress, field: 'Smtp' });
      if (mbx.ForwardingAddress) forwards.push({ address: mbx.ForwardingAddress, field: 'Recipient' });
    }
    asArray(mbx.Delegates).forEach((row) => { if (row.User) delegates.push({ user: row.User, right: row.Right || '' }); });
  });
  const facts = {
    mailbox,
    rules,
    forwards,
    delegates,
    ticketNumber: state.ticket ? String(state.ticket.ticketId || '') : '',
    ticket: (state.ticket && state.ticket.ticketContent) || '',
    company: (state.ticket && state.ticket.companyName) || '',
    authMethods: asArray(mem.authMethods),
    devices: asArray(mem.devices),
    intune: asArray(mem.intune),
    oauth: asArray(mem.oauth),
    roles: asArray(mem.roles),
    mobile: asArray(mem.mobile),
  };
  if (audit) {
    facts.mfaPosture = audit.MfaPosture || audit.mfaPosture || null;
    facts.messageTrace = asArray(audit.MessageTrace || audit.messageTrace);
    if (Object.prototype.hasOwnProperty.call(audit, 'ReportedMessages') || Object.prototype.hasOwnProperty.call(audit, 'reportedMessages')) {
      facts.reportedMessages = asArray(audit.ReportedMessages || audit.reportedMessages);
    }
    facts.ual = asArray(audit.Ual || audit.ual);
    if (Object.prototype.hasOwnProperty.call(audit, 'Dlp') || Object.prototype.hasOwnProperty.call(audit, 'dlp')) facts.dlp = asArray(audit.Dlp || audit.dlp);
    if (Object.prototype.hasOwnProperty.call(audit, 'Files') || Object.prototype.hasOwnProperty.call(audit, 'files')) facts.files = asArray(audit.Files || audit.files);
    if (Object.prototype.hasOwnProperty.call(audit, 'Sharing') || Object.prototype.hasOwnProperty.call(audit, 'sharing')) facts.sharing = asArray(audit.Sharing || audit.sharing);
    if (Object.prototype.hasOwnProperty.call(audit, 'Alerts') || Object.prototype.hasOwnProperty.call(audit, 'alerts')) facts.alerts = asArray(audit.Alerts || audit.alerts);
    facts.signIns = asArray(audit.SignIns || audit.signIns);
    facts.apps = asArray(audit.Apps || audit.apps);
    if (Object.prototype.hasOwnProperty.call(audit, 'Blasts') || Object.prototype.hasOwnProperty.call(audit, 'blasts')) facts.blasts = asArray(audit.Blasts || audit.blasts);
    facts.pullNotes = asArray(audit.PullNotes || audit.pullNotes);
    if (audit.Days || audit.days) facts.windowDays = audit.Days || audit.days;
  }
  const startIso = state.windowStart || '';
  const apps = [];
  const seen = new Set();
  const addApp = (app, fromAudit) => {
    if (!app) return;
    const name = String(app.DisplayName || app.displayName || '').trim();
    const created = String(app.Created || app.created || app.createdDateTime || '').trim();
    if (!name) return;
    const key = `${name}|${created}`.toLowerCase();
    if (seen.has(key)) return;
    const inside = inWindow(created, startIso);
    if (inside === false) return;
    if (inside !== true && !fromAudit) return;
    seen.add(key);
    apps.push({ displayName: name, created });
  };
  asArray(facts.apps).forEach((app) => addApp(app, true));
  asArray(mem.apps).forEach((app) => addApp(app, false));
  facts.apps = apps.slice(0, 25);
  const risky = /redirect|forward|delete|drop|bcc|blind|connector|route|reject|quarantine/i;
  facts.transport = asArray(mem.transport).filter((rule) => {
    const when = inWindow(rule.WhenChanged || rule.whenChanged || '', startIso);
    const text = `${rule.Name || ''} ${rule.Description || ''}`;
    return when !== false || risky.test(text);
  }).slice(0, 40).map((rule) => ({
    name: rule.Name || rule.Identity || '',
    state: rule.State || '',
    priority: rule.Priority,
    whenChanged: rule.WhenChanged || rule.whenChanged || '',
    description: String(rule.Description || '').slice(0, 180),
  }));
  facts.connectors = asArray(mem.connectors).filter((row) => {
    const when = inWindow(row.WhenCreated || row.whenCreated || '', startIso);
    const outbound = String(row.Direction || '').toLowerCase() === 'outbound';
    const enabled = row.Enabled === true || row.Enabled === 'True';
    return when !== false || (outbound && enabled);
  }).slice(0, 30).map((row) => ({
    direction: row.Direction || '',
    name: row.Name || row.Identity || '',
    enabled: (row.Enabled === true || row.Enabled === 'True') ? 'enabled' : 'disabled',
    connectorType: row.ConnectorType || '',
    hosts: row.SmartHosts || '',
    domains: row.Domains || '',
    whenCreated: row.WhenCreated || row.whenCreated || '',
  }));
  if (windowNote) facts.windowNote = windowNote;
  return facts;
}

function repullLogs(review) {
  return asArray(review && (review.repullLogs || review.RepullLogs))
    .map((name) => String(name || '').trim().toLowerCase())
    .filter((name) => REPULL_LOGS.includes(name));
}

function repullLabel(logs) {
  return Array.isArray(logs) && logs.length ? logs.join(', ') : 'message trace and audit';
}

function followUpNote(again, logs) {
  const span = `${formatLocal(again.start)} through ${formatLocal(again.end)} ${localZone()}`;
  const cap = again.days >= 90
    ? 'This is the longest lookback. Do not ask for another pull. If the record is not here, say it was not in the logs.'
    : 'If this window still cannot answer the investigation, ask for a longer look up to 90 days.';
  if (Array.isArray(logs) && logs.length) {
    return `Follow-up pull of ${logs.join(', ')} only, ${again.days} days, ${span}. The other logs stay on their earlier window. ${cap}`;
  }
  return `Follow-up pull, ${again.days} days, ${span}. Message trace and unified audit stop at tenant retention if that is shorter. ${cap}`;
}

function mergeAudit(prior, next, logs) {
  if (!prior || !next || !Array.isArray(logs) || !logs.length) return next || prior;
  const copy = { ...prior };
  const fields = {
    trace: ['MessageTrace', 'messageTrace', 'Blasts', 'blasts'],
    ual: ['Ual', 'ual'],
    sharing: ['Sharing', 'sharing'],
    files: ['Files', 'files'],
    dlp: ['Dlp', 'dlp'],
    alerts: ['Alerts', 'alerts'],
    signins: ['SignIns', 'signIns'],
    apps: ['Apps', 'apps'],
  };
  logs.forEach((log) => {
    (fields[log] || []).forEach((key) => {
      if (Object.prototype.hasOwnProperty.call(next, key)) copy[key] = next[key];
    });
  });
  copy.Days = next.Days || next.days || prior.Days;
  copy.PullNotes = [...asArray(prior.PullNotes || prior.pullNotes), ...asArray(next.PullNotes || next.pullNotes)];
  return copy;
}

function auditFailed(audit) {
  if (!audit) return true;
  return !['Ual', 'ual', 'Sharing', 'sharing', 'Files', 'files', 'SignIns', 'signIns', 'Dlp', 'dlp']
    .some((key) => Object.prototype.hasOwnProperty.call(audit, key));
}

async function pullAudit(clientNumber, mailbox, span, logs, lead) {
  const named = Array.isArray(logs) ? logs.filter((name) => REPULL_LOGS.includes(name)) : [];
  const extra = named.length ? `|LOGS:${named.join(',')}` : '';
  const from = (!named.length && lead && lead.senders && lead.senders.length) ? `|FROM:${lead.senders.join(',')}` : '';
  const subject = (!named.length && from && lead.subject) ? `|SUBJECT:${lead.subject}` : '';
  const cmd = `REMEDIATE_BEC_AUDIT|USER:${mailbox}|START:${span.start}|END:${span.end}${extra}${from}${subject}`;
  const label = named.length ? `Pulling ${named.join(', ')}` : 'Pulling message trace, audit, DLP, alerts, and sign-ins';
  const final = await cachedCommand(clientNumber, cmd, label, 300);
  const parsed = parsePayload(final || '');
  if (parsed.prefix !== 'REMEDIATE_BEC:' || !parsed.data) {
    const detail = (parsed.data && (parsed.data.message || parsed.data.Message)) || parsed.raw || 'audit pull failed';
    return { PullNotes: [String(detail)], Days: span.days };
  }
  state.blasts = asArray(parsed.data.Blasts || parsed.data.blasts);
  return parsed.data;
}

const pullCache = new Map();

function takeListRows(data, warn) {
  const errors = data.Errors || data.errors || {};
  const take = (key, field, assign) => {
    const err = errors[key] || '';
    if (err) warn.push(`${key}: ${err}`);
    else assign(asArray(data[field]));
  };
  take('rules', 'Rules', (rows) => { state.memory.rules = rows; });
  take('mailbox', 'Users', (rows) => { state.memory.mailbox = rows; });
  take('mobile', 'Devices', (rows) => { state.memory.mobile = rows; });
  take('transport', 'Transport', (rows) => { state.memory.transport = rows; });
  take('connectors', 'Connectors', (rows) => { state.memory.connectors = rows; });
}

function takeGraphRows(data, warn) {
  const errors = data.Errors || data.errors || {};
  const take = (key, field, assign) => {
    const err = errors[key] || '';
    if (err) warn.push(`${field}: ${err}`);
    else assign(asArray(data[field]));
  };
  take('auth', 'Methods', (rows) => { state.memory.authMethods = rows; });
  take('devices', 'Devices', (rows) => { state.memory.devices = rows; });
  take('oauth', 'Grants', (rows) => { state.memory.oauth = rows; });
  take('intune', 'Intune', (rows) => { state.memory.intune = rows; });
  take('roles', 'Roles', (rows) => { state.memory.roles = rows; });
  take('apps', 'Apps', (rows) => { state.memory.apps = rows; });
}

async function cachedCommand(clientNumber, command, label, waitSeconds) {
  const key = `${clientNumber}|${command}`;
  if (pullCache.has(key)) return pullCache.get(key);
  const final = await sendCommand(clientNumber, command, label, waitSeconds);
  const parsed = parsePayload(final || '');
  if (parsed.data && !String(parsed.prefix || '').startsWith('REMEDIATE_FAILED')) pullCache.set(key, final);
  return final;
}

async function pullLists(clientNumber, mailbox, tenant) {
  state.memory = {};
  const both = Boolean(tenant.exchangeAuthenticated && tenant.graphAuthenticated);
  if (both) {
    setStatus('scanStatus', 'Pulling mailbox settings and Graph state together…');
    try {
      const final = await cachedCommand(clientNumber, `REMEDIATE_HUNTER_LISTS|USER:${mailbox}|WITHGRAPH`, 'Pulling mailbox and Graph state', 300);
      const parsed = parsePayload(final || '');
      if (parsed.prefix !== 'REMEDIATE_HUNTER_PACK:' || !parsed.data) {
        throw new Error((parsed.data && (parsed.data.message || parsed.data.Message)) || parsed.raw || 'Combined pull failed');
      }
      const warn = [];
      const lists = parsed.data.Lists || parsed.data.lists || {};
      const graph = parsed.data.Graph || parsed.data.graph || {};
      takeListRows(lists, warn);
      takeGraphRows(graph, warn);
      if (warn.length) setStatus('scanStatus', warn.join(' '), 'warn');
      return;
    } catch (e) {
      setStatus('scanStatus', `Combined pull failed, continuing with separate pulls. ${e.message}`, 'warn');
    }
  }
  if (tenant.exchangeAuthenticated) {
    setStatus('scanStatus', 'Listing inbox rules, forwarding, mobile devices, transport rules, and connectors…');
    try {
      const final = await cachedCommand(clientNumber, `REMEDIATE_HUNTER_LISTS|USER:${mailbox}`, 'Listing mailbox and tenant mail settings', 300);
      const parsed = parsePayload(final || '');
      if (parsed.prefix !== 'REMEDIATE_HUNTER_LISTS:' || !parsed.data) {
        throw new Error((parsed.data && (parsed.data.message || parsed.data.Message)) || parsed.raw || 'Mailbox list pull failed');
      }
      const warn = [];
      takeListRows(parsed.data, warn);
      if (warn.length) setStatus('scanStatus', warn.join(' '), 'warn');
    } catch (e) {
      setStatus('scanStatus', `Mailbox list pull failed, continuing. ${e.message}`, 'warn');
    }
  }
  if (!tenant.graphAuthenticated) return;
  setStatus('scanStatus', 'Pulling MFA methods, devices, OAuth, Intune, roles, and apps…');
  try {
    const final = await cachedCommand(clientNumber, `REMEDIATE_HUNTER_GRAPH|USER:${mailbox}`, 'Pulling Graph mailbox state', 300);
    const parsed = parsePayload(final || '');
    if (parsed.prefix !== 'REMEDIATE_HUNTER_GRAPH:' || !parsed.data) {
      throw new Error((parsed.data && (parsed.data.message || parsed.data.Message)) || parsed.raw || 'Graph mailbox pull failed');
    }
    const warn = [];
    takeGraphRows(parsed.data, warn);
    if (warn.length) setStatus('scanStatus', warn.join(' '), 'warn');
  } catch (e) {
    setStatus('scanStatus', `Graph mailbox pull failed, continuing. ${e.message}`, 'warn');
  }
}

async function requestReview(mailbox, audit, windowNote, revision) {
  const facts = factsFor(mailbox, audit, windowNote);
  const note = String(revision || '').trim();
  if (note) {
    facts.revision = note;
    facts.priorReview = ($('review').value || state.review || '').trim();
  }
  const started = await api('/api/triage/bec-review-start', {
    method: 'POST',
    body: JSON.stringify(facts),
  }, 30000);
  const jobId = started && started.job && started.job.id;
  if (!jobId) throw new Error('The review did not start.');
  const began = Date.now();
  let last = '';
  while (Date.now() - began < 360000) {
    const data = await api(`/api/triage/bec-review-job?id=${encodeURIComponent(jobId)}`, {}, 20000);
    const job = data.job || {};
    if (job.text && job.text !== last) {
      last = job.text;
      $('review').value = job.text;
      showOutput('review', { keepSetup: true });
    }
    if (job.done) {
      if (job.error) throw new Error(job.error);
      return job.result || {};
    }
    await new Promise((resolve) => setTimeout(resolve, 400));
  }
  throw new Error('The review timed out.');
}

function showReview(review) {
  state.review = (review && review.text) || '';
  const verdict = String((review && (review.verdict || review.Verdict)) || '').trim();
  const badge = $('verdictBadge');
  if (badge) {
    badge.hidden = !verdict;
    badge.textContent = verdict ? `Verdict: ${verdict}` : '';
  }
  const box = $('review');
  box.value = state.review;
  $('btnNotes').hidden = !state.review;
  $('btnReviseReview').hidden = !state.review;
  if (state.review) $('btnNotes').disabled = false;
  $('btnReviewAgain').hidden = !(state.review && state.audit && state.span);
  if (state.review) {
    renderThread();
    showOutput('review');
  }
}

function offerRepull(days, logs) {
  state.pendingRepull = { days, logs };
  $('btnRepull').hidden = false;
  $('btnRepull').textContent = `Pull ${days} days of ${repullLabel(logs)} and review again`;
}

async function reviewWithFollowUps(clientNumber, mailbox, audit, span) {
  const tz = localZone();
  const windowNote = span.hadTime
    ? `Ticket time was read as ${tz}. This pull is ${span.days} days, ${formatLocal(span.start)} through ${formatLocal(span.end)} ${tz}. Log times below are ${tz}.`
    : `No event time was found. This pull is the last ${span.days} days, ${formatLocal(span.start)} through ${formatLocal(span.end)} ${tz}. Log times below are ${tz}.`;
  setStatus('scanStatus', 'Writing the threat review…');
  let review = await requestReview(mailbox, audit, windowNote);
  showReview(review);
  const seen = new Set();
  let followUps = 0;
  while (followUps < 2) {
    const asked = parseInt(review.repullDays, 10) || 0;
    const logs = repullLogs(review);
    const pulled = parseInt((audit && (audit.Days || audit.days)) || span.days, 10) || span.days;
    if (!(asked > pulled)) break;
    const longer = Math.min(90, asked);
    const key = `${longer}:${logs.join(',')}`;
    if (seen.has(key)) break;
    seen.add(key);
    if (!$('autoRepull').checked) {
      offerRepull(longer, logs);
      setStatus('scanStatus', `The investigation still needs ${longer} days of ${repullLabel(logs)}.`);
      return;
    }
    const again = {
      days: longer,
      start: new Date(Date.now() - longer * 24 * 60 * 60 * 1000).toISOString(),
      end: new Date().toISOString(),
    };
    setStatus('scanStatus', `The investigation needs ${repullLabel(logs)} going back ${longer} days. Pulling those logs…`);
    let next = null;
    try { next = await pullAudit(clientNumber, mailbox, again, logs); }
    catch (e) { next = { PullNotes: [e.message], Days: longer }; }
    audit = logs.length ? mergeAudit(audit, next, logs) : next;
    state.audit = audit;
    if (!logs.length) state.windowStart = again.start;
    followUps += 1;
    review = await requestReview(mailbox, audit, followUpNote(again, logs));
    showReview(review);
  }
  state.pendingRepull = null;
  $('btnRepull').hidden = true;
  setStatus('scanStatus', 'Threat review finished.');
}

async function runScan() {
  const mailbox = $('mailbox').value.trim();
  const tenant = selectedTenant();
  if (!state.ticket) throw new Error('Choose a ticket first.');
  if (!mailbox) throw new Error('Enter the mailbox to check.');
  if (!tenant) throw new Error('Add a tenant below and sign in to Exchange and Graph.');
  if (!tenant.graphAuthenticated && !tenant.exchangeAuthenticated) {
    throw new Error('Sign in with Exchange Auth, then Graph Auth, before the check.');
  }
  state.pendingRepull = null;
  state.audit = null;
  state.span = null;
  $('btnRepull').hidden = true;
  $('btnReviewAgain').hidden = true;
  $('notesCard').hidden = true;
  $('review').value = '';
  $('reviewRevision').value = '';
  $('notesRevision').value = '';
  $('btnReviseReview').hidden = true;
  $('useClient').checked = false;
  $('useInternal').checked = false;
  $('useIncident').checked = false;
  $('btnNotes').hidden = true;
  showOutput('ticket', { keepSetup: true });
  const clientNumber = tenant.clientNumber;
  await claimTicket();
  await ensureWorker(clientNumber);
  await pullLists(clientNumber, mailbox, tenant);
  const span = windowFromTicket(state.ticket.ticketContent || '');
  state.windowStart = span.start;
  setStatus('scanStatus', `Threat check for ${mailbox}, last ${span.days} day(s)…`);
  let audit = null;
  const lead = reportedLead(mailbox);
  try { audit = await pullAudit(clientNumber, mailbox, span, null, lead); }
  catch (e) { audit = { PullNotes: [e.message], Days: span.days }; }
  if (auditFailed(audit)) {
    setStatus('scanStatus', 'Audit logs were not collected. Pulling them again…');
    try { audit = await pullAudit(clientNumber, mailbox, span, null, lead); }
    catch (e) { audit = { PullNotes: [e.message], Days: span.days }; }
  }
  state.audit = audit;
  state.span = span;
  await reviewWithFollowUps(clientNumber, mailbox, audit, span);
}

async function reviewAgain() {
  const mailbox = $('mailbox').value.trim();
  const tenant = selectedTenant();
  if (!state.audit || !state.span || !mailbox) throw new Error('Run the threat check before reviewing again.');
  if (!tenant) throw new Error('Sign in to the company on the main window, then refresh tenants.');
  $('btnReviewAgain').hidden = true;
  await reviewWithFollowUps(tenant.clientNumber, mailbox, state.audit, state.span);
}

async function runRepull() {
  const pending = state.pendingRepull;
  const mailbox = $('mailbox').value.trim();
  const tenant = selectedTenant();
  if (!pending || !mailbox || !tenant) return;
  $('btnRepull').hidden = true;
  const again = {
    days: pending.days,
    start: new Date(Date.now() - pending.days * 24 * 60 * 60 * 1000).toISOString(),
    end: new Date().toISOString(),
    hadTime: true,
  };
  await ensureWorker(tenant.clientNumber);
  let next = null;
  try { next = await pullAudit(tenant.clientNumber, mailbox, again, pending.logs); }
  catch (e) { next = { PullNotes: [e.message], Days: pending.days }; }
  const audit = pending.logs.length ? mergeAudit(state.audit, next, pending.logs) : next;
  state.audit = audit;
  if (!pending.logs.length) state.windowStart = again.start;
  setStatus('scanStatus', 'Writing the threat review…');
  const review = await requestReview(mailbox, audit, followUpNote(again, pending.logs));
  showReview(review);
  const more = parseInt(review.repullDays, 10) || 0;
  if (more > pending.days) {
    offerRepull(Math.min(90, more), repullLogs(review));
    setStatus('scanStatus', `The investigation still needs ${repullLabel(repullLogs(review))} for ${Math.min(90, more)} days.`);
    return;
  }
  state.pendingRepull = null;
  setStatus('scanStatus', 'Threat review finished.');
}

function blastSection() {
  const blocks = [];
  asArray(state.blasts).forEach((blast) => {
    const recips = asArray(blast.Recipients || blast.recipients).map((r) => String(r || '').trim()).filter(Boolean);
    if (recips.length < 5) return;
    const subject = String(blast.Subject || blast.subject || '(no subject)').trim() || '(no subject)';
    const when = formatLocal(blast.Received || blast.received || '');
    const lines = [subject];
    if (when) lines.push(when);
    recips.forEach((addr) => lines.push(addr));
    blocks.push(lines.join('\n'));
  });
  if (!blocks.length) return '';
  return `DELIVERED BLAST RECIPIENTS\n\n${blocks.join('\n\n')}`;
}

async function reviseReview() {
  const mailbox = $('mailbox').value.trim();
  const revision = $('reviewRevision').value.trim();
  if (!state.audit || !state.span || !mailbox) throw new Error('Run the threat check before revising the review.');
  if (!revision) throw new Error('Write what the next review should change.');
  const span = state.span;
  const tz = localZone();
  const windowNote = span.hadTime
    ? `Ticket time was read as ${tz}. This pull is ${span.days} days, ${formatLocal(span.start)} through ${formatLocal(span.end)} ${tz}. Log times below are ${tz}.`
    : `No event time was found. This pull is the last ${span.days} days, ${formatLocal(span.start)} through ${formatLocal(span.end)} ${tz}. Log times below are ${tz}.`;
  setStatus('scanStatus', 'Rewriting the review with your revision…');
  const review = await requestReview(mailbox, state.audit, windowNote, revision);
  showReview(review);
  const asked = parseInt(review.repullDays, 10) || 0;
  const pulled = parseInt((state.audit && (state.audit.Days || state.audit.days)) || span.days, 10) || span.days;
  if (asked > pulled) {
    offerRepull(Math.min(90, asked), repullLogs(review));
    setStatus('scanStatus', 'Review revised from the logs already pulled. It still asks for a longer window.');
    return;
  }
  state.pendingRepull = null;
  $('btnRepull').hidden = true;
  setStatus('scanStatus', 'Review revised from the logs already pulled.');
}

async function writeNotes(applyRevision) {
  const review = ($('review').value || state.review || '').trim();
  if (!review) throw new Error('Finish the threat check before writing the notes.');
  const revision = applyRevision ? $('notesRevision').value.trim() : '';
  if (applyRevision && !revision) throw new Error('Write what the notes should change.');
  state.review = review;
  setStatus('scanStatus', applyRevision
    ? 'Rewriting the notes with your revision…'
    : 'Writing the internal note, client email, and incident report…');
  const currentDrafts = revision ? [
    'Client email:',
    $('clientText').value.trim(),
    '',
    'Internal note:',
    $('internalText').value.trim(),
    '',
    'Incident report:',
    $('incidentText').value.trim(),
  ].join('\n') : '';
  let reviewPayload = review;
  if (revision) {
    reviewPayload += `\n\n<<<TECHNICIAN_REVISION>>>\n${revision}`;
    if (currentDrafts.trim()) reviewPayload += `\n\n<<<CURRENT_DRAFTS>>>\n${currentDrafts}`;
  }
  const data = await api('/api/triage/memberberry-note', {
    method: 'POST',
    body: JSON.stringify({
      review: reviewPayload,
      ticket: (state.ticket && state.ticket.ticketContent) || '',
      ticketNumber: state.ticket ? String(state.ticket.ticketId || '') : '',
      company: (state.ticket && state.ticket.companyName) || '',
      contact: (state.ticket && state.ticket.contactName) || '',
      audience: 'all',
      actions: [],
      revision,
      currentDrafts,
    }),
  }, 360000);
  const result = data.result || {};
  $('clientText').value = result.clientText || '';
  $('internalText').value = result.internalText || result.text || '';
  let incident = result.incidentText || '';
  const blasts = blastSection();
  if (blasts && !/DELIVERED BLAST RECIPIENTS/i.test(incident)) incident = `${incident.trim()}\n\n${blasts}`.trim();
  $('incidentText').value = incident;
  $('useClient').checked = Boolean($('clientText').value.trim());
  $('useInternal').checked = Boolean($('internalText').value.trim());
  const confirmedIncident = Boolean(result.confirmedIncident || result.ConfirmedIncident);
  $('useIncident').checked = confirmedIncident && Boolean(incident.trim());
  $('useEscalation').checked = false;
  $('clientInsert').value = '';
  if (!$('escFrom').value.trim()) {
    $('escFrom').value = (state.catalog && state.catalog.memberName) || 'Chris Knospe';
  }
  if (!$('escSummary').value.trim()) {
    $('escSummary').value = review.split(/\n\s*\n/)[0].trim().slice(0, 700);
  }
  $('notesCard').hidden = false;
  const internalWrap = $('internalWrap');
  const incidentWrap = $('incidentWrap');
  if (internalWrap) internalWrap.open = Boolean($('internalText').value.trim());
  if (incidentWrap) incidentWrap.open = Boolean(incident.trim());
  showOutput(confirmedIncident ? 'incident' : 'client');
  const draftCol = document.querySelector('.draftCol');
  if (draftCol) draftCol.scrollTop = 0;
  setStatus('scanStatus', confirmedIncident
    ? 'Notes ready. The ticket confirms an incident, so the incident report is selected.'
    : 'Notes ready. Choose what to post.');
}

function composeEscalation() {
  const kind = $('escKind').value === 'hierarchical' ? 'Hierarchical (to a lead)' : 'Functional (to another team)';
  const from = $('escFrom').value.trim();
  const to = $('escTo').value.trim();
  const reasons = [];
  if ($('escExpertise').checked) reasons.push('Requires network / server / security / application expertise');
  if ($('escAccess').checked) reasons.push('Access limitation');
  if ($('escVendor').checked) reasons.push('Vendor-specific issue');
  if ($('escSla').checked) reasons.push('Time-sensitive / SLA risk');
  if ($('escOther').checked) {
    const other = $('escOtherText').value.trim();
    if (!other) throw new Error('Fill in the other escalation reason.');
    reasons.push(`Other: ${other}`);
  }
  const summary = $('escSummary').value.trim();
  if (!from) throw new Error('Say who the escalation is from.');
  if (!to) throw new Error('Say who the escalation is to.');
  if (!reasons.length) throw new Error('Pick at least one escalation reason.');
  if (!summary) throw new Error('Add a short escalation summary.');
  return [
    'ESCALATION TYPE',
    kind,
    '',
    'ROUTING',
    `Escalating From: ${from}`,
    `Escalating To: ${to}`,
    '',
    'REASON FOR ESCALATION',
    ...reasons.map((reason) => `- ${reason}`),
    '',
    'SUMMARY OF ISSUE',
    summary,
  ].join('\n');
}

function fillTicketStatus() {
  const pick = $('ticketStatus');
  if (!pick) return;
  const choices = state.statusChoices || [];
  pick.innerHTML = choices.map((choice) => `<option value="${choice.id}">${escapeHtml(choice.name)}${choice.closed ? ' — closes the ticket' : ''}</option>`).join('');
  const name = state.ticket && state.ticket.statusName;
  const match = choices.find((choice) => choice.name === name);
  if (match) pick.value = String(match.id);
  syncCloseTime();
}

function selectedStatusChoice() {
  const id = Number($('ticketStatus').value);
  return (state.statusChoices || []).find((choice) => choice.id === id) || null;
}

function datetimeLocalValue(date) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function syncCloseTime() {
  const closing = Boolean(selectedStatusChoice() && selectedStatusChoice().closed);
  const row = $('closeTime');
  if (row) row.hidden = !closing;
  const none = !closing || $('closeTimePick').value === 'none';
  $('closeStart').disabled = none;
  $('closeHours').disabled = none;
  if (closing && !$('closeStart').value) {
    $('closeStart').value = datetimeLocalValue(new Date(Date.now() - 15 * 60 * 1000));
  }
}

async function setTicketStatus() {
  if (!state.ticket) throw new Error('Choose a ticket first.');
  const choice = selectedStatusChoice();
  if (!choice) throw new Error('Choose a status.');
  const time = choice.closed ? $('closeTimePick').value : 'none';
  const hours = $('closeHours').value;
  const start = $('closeStart').value;
  const lines = [
    `Set ticket #${state.ticket.ticketId} to ${choice.name}?`,
    '',
  ];
  if (choice.closed) {
    lines.push(time === 'none' ? 'Time: no time entry.' : `Time: ${$('closeTimePick').selectedOptions[0].textContent}, ${hours} h, starting ${start.replace('T', ' ')}.`);
    lines.push('This closes the ticket.');
  } else {
    lines.push('This does not close the ticket.');
  }
  if (!window.confirm(lines.join('\n'))) return;
  setStatus('statusChange', 'Updating the ticket status…');
  const data = await api('/api/manage/hunter-post', {
    method: 'POST',
    body: JSON.stringify({
      ticketId: state.ticket.ticketId,
      status: 'set-status',
      statusId: choice.id,
      time,
      hours: Number(hours),
      timeStart: start,
    }),
  }, 60000);
  const done = asArray(data.completed).filter(Boolean);
  const errors = asArray(data.errors).filter(Boolean);
  if (data.statusUpdated && data.statusName) showTicketAssignment(data.statusName, 'keep');
  const summary = [...done, ...errors].join(' ');
  setStatus('statusChange', summary || 'Status updated.', errors.length ? 'bad' : '');
  if (errors.length) throw new Error(summary);
  await refreshManageLog(state.ticket.ticketId);
}

function syncHours() {
  const none = $('timePick').value === 'none';
  $('hours').disabled = none;
}

async function postNotes() {
  if (!state.ticket) throw new Error('Choose a ticket first.');
  const notes = {
    client: $('useClient').checked ? $('clientText').value.trim() : '',
    internal: $('useInternal').checked ? $('internalText').value.trim() : '',
    incident: $('useIncident').checked ? $('incidentText').value.trim() : '',
  };
  const escalation = $('useEscalation').checked ? composeEscalation() : '';
  if (escalation) notes.internal = [notes.internal, escalation].filter(Boolean).join('\n\n');
  if (!notes.client && !notes.internal && !notes.incident) throw new Error('Choose at least one note to post.');
  const time = $('timePick').value;
  if (time !== 'none' && !notes.client) throw new Error('Time is added only with the client email.');
  const member = (state.catalog && state.catalog.memberIdentifier) || 'cknospe';
  const statusLabel = $('statusPick').selectedOptions[0].textContent;
  const timeLabel = $('timePick').selectedOptions[0].textContent;
  const email = $('includeRecipients').checked;
  const lines = [
    `Post to ticket #${state.ticket.ticketId} as ${member}?`,
    '',
    `Status: ${statusLabel}`,
    `Client email: ${notes.client ? (email ? 'yes, email existing recipients' : 'yes, no email') : 'no'}`,
    `Time: ${notes.client && time !== 'none' ? `${timeLabel}, ${$('hours').value} h` : 'no'}`,
    `Internal note: ${notes.internal ? 'yes' : 'no'}`,
    `Escalation: ${escalation ? 'yes, included in the internal note' : 'no'}`,
    `Incident report: ${notes.incident ? (email ? 'yes, email existing recipients' : 'yes, no email') : 'no'}`,
    '',
    'This does not close the ticket.',
  ];
  if (!window.confirm(lines.join('\n'))) return;
  setStatus('postStatus', 'Posting to Manage…');
  const data = await api('/api/manage/hunter-post', {
    method: 'POST',
    body: JSON.stringify({
      ticketId: state.ticket.ticketId,
      status: $('statusPick').value,
      includeRecipients: email,
      time,
      hours: Number($('hours').value),
      memberIdentifier: member,
      notes,
    }),
  }, 120000);
  const done = asArray(data.completed).filter(Boolean);
  const errors = asArray(data.errors).filter(Boolean);
  if (data.statusUpdated && data.statusName) showTicketAssignment(data.statusName, 'keep');
  const summary = [...done, ...errors].join(' ');
  setStatus('postStatus', summary || 'Posted.', errors.length ? 'bad' : '');
  if (notes.internal.includes('HATZ-AI POWERED REVIEW') && !errors.some((line) => /internal note/i.test(line))) {
    const listed = state.tickets.find((ticket) => String(ticket.id) === String(state.ticket.ticketId));
    if (listed) listed.triaged = true;
  }
  await loadThread(state.ticket.ticketId).catch(() => {});
}

function useMailbox(upn) {
  const value = String(upn || '').trim();
  if (!value) return;
  $('mailbox').value = value;
  const pick = $('mailboxPick');
  pick.value = [...pick.options].some((opt) => opt.value === value) ? value : '';
  setStatus('userStatus', `Mailbox to check: ${value}`);
}

function renderUserMatches(users) {
  const box = $('userMatches');
  const rows = asArray(users).map((user) => {
    if (typeof user === 'string') return { upn: user.trim(), name: '' };
    const upn = String(user.UserPrincipalName || user.userPrincipalName || '').trim();
    const name = String(user.DisplayName || user.displayName || '').trim();
    return { upn, name };
  }).filter((user) => user.upn);
  if (!rows.length) {
    box.innerHTML = '';
    setStatus('userStatus', 'No user in that tenant matched.', 'warn');
    return;
  }
  box.innerHTML = rows.map((user) => {
    const label = user.name ? `${user.name} (${user.upn})` : user.upn;
    return `<button type="button" class="userMatch" data-upn="${escapeHtml(user.upn)}">Use ${escapeHtml(label)}</button>`;
  }).join('');
  box.querySelectorAll('.userMatch').forEach((btn) => {
    btn.addEventListener('click', () => useMailbox(btn.dataset.upn));
  });
  setStatus('userStatus', rows.length === 1 ? 'One user matched. Choose that mailbox if it is the right person.' : `${rows.length} users matched. Choose the mailbox to check.`);
}

async function findUser() {
  const tenant = selectedTenant();
  if (!tenant) throw new Error('Add a tenant and sign in with Graph Auth before looking up a user.');
  if (!tenant.graphAuthenticated) throw new Error('Graph sign-in is required to look up users. Use Graph Auth in this window.');
  const terms = $('userSearch').value.split(',').map((part) => part.trim()).filter(Boolean);
  if (!terms.length) throw new Error('Enter a name or email to look up.');
  await ensureWorker(tenant.clientNumber);
  setStatus('userStatus', 'Looking up users…');
  $('userMatches').innerHTML = '';
  const before = await api(`/api/tenants/${tenant.clientNumber}/response`);
  let prior = normalize(before?.response);
  const command = `VALIDATE_USERS|SEARCH_TERMS:${JSON.stringify(terms)}`;
  const initial = await api(`/api/tenants/${tenant.clientNumber}/command`, {
    method: 'POST',
    body: JSON.stringify({ command, noWait: true, waitSeconds: 0 }),
  });
  const prefixes = ['VALIDATE_USERS_SUCCESS:', 'VALIDATE_USERS_FAILED:', 'VALIDATE_USERS_STARTED'];
  let text = extractToken(initial.response, prefixes);
  const deadline = Date.now() + 180000;
  while (text === 'VALIDATE_USERS_STARTED' || !text || text === prior) {
    if (Date.now() > deadline) throw new Error('User lookup timed out');
    await new Promise((resolve) => setTimeout(resolve, 2000));
    const data = await api(`/api/tenants/${tenant.clientNumber}/response`);
    const resp = normalize(data?.response);
    if (!resp || resp === prior) continue;
    text = extractToken(resp, prefixes);
    if (resp.includes('VALIDATE_USERS_STARTED')) prior = resp;
  }
  if (text.startsWith('VALIDATE_USERS_FAILED:')) throw new Error(text.slice('VALIDATE_USERS_FAILED:'.length) || 'User lookup failed.');
  if (!text.startsWith('VALIDATE_USERS_SUCCESS:')) throw new Error(text || 'User lookup failed.');
  let parsed = {};
  try { parsed = JSON.parse(text.slice('VALIDATE_USERS_SUCCESS:'.length)); } catch { parsed = {}; }
  renderUserMatches(parsed.Users || parsed.users || []);
}

async function ensureSession() {
  const session = await api('/api/session');
  if (session && session.active) return session;
  return api('/api/session', { method: 'POST', body: JSON.stringify({ daysBack: 7 }) });
}

async function addTenant() {
  setStatus('authStatus', 'Starting a tenant…');
  await ensureSession();
  const created = await api('/api/tenants', { method: 'POST' });
  await refreshTenants();
  if (created && created.clientNumber) $('tenantPick').value = String(created.clientNumber);
  syncTenantAuth();
  setStatus('authStatus', 'Tenant added. Run Exchange Auth, then Graph Auth. Sign-in opens on this PC.');
}

async function waitForWorkerReady(clientNumber) {
  const deadline = Date.now() + 45000;
  while (Date.now() < deadline) {
    const data = await api(`/api/tenants/${clientNumber}/status?tailLines=25`);
    const status = data?.status || '';
    if (/ready to receive commands/i.test(status) && /Ready! Waiting for/i.test(status)) return;
    await new Promise((resolve) => setTimeout(resolve, 800));
  }
  throw new Error('The sign-in window did not become ready. Try again.');
}

async function runAuthCommand(clientNumber, command, successPrefix, failPrefix, label) {
  const before = await api(`/api/tenants/${clientNumber}/response`);
  const prior = normalize(before?.response);
  const initial = await api(`/api/tenants/${clientNumber}/command`, {
    method: 'POST',
    body: JSON.stringify({ command, noWait: true, waitSeconds: 0 }),
  });
  const read = (text) => {
    const value = normalize(text);
    if (!value || value === prior) return null;
    if (value.includes(failPrefix)) {
      const at = value.indexOf(failPrefix);
      throw new Error(value.slice(at) || `${label} failed.`);
    }
    if (value.includes(successPrefix)) return value.slice(value.indexOf(successPrefix));
    return null;
  };
  const immediate = read(initial.response);
  if (immediate) return immediate;
  const waitSeconds = 300;
  const deadline = Date.now() + waitSeconds * 1000;
  let lastNote = 0;
  while (Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 2000));
    const data = await api(`/api/tenants/${clientNumber}/response`);
    const elapsed = Math.round((Date.now() - (deadline - waitSeconds * 1000)) / 1000);
    if (elapsed - lastNote >= 15) {
      setStatus('authStatus', `${label}… (${elapsed}s). Finish sign-in in the window on this PC.`);
      lastNote = elapsed;
    }
    const found = read(data?.response);
    if (found) return found;
  }
  throw new Error(`${label} timed out. Finish the sign-in window, then try again.`);
}

async function exchangeAuth() {
  const tenant = selectedTenant();
  if (!tenant) throw new Error('Add a tenant first.');
  setStatus('authStatus', 'Opening a sign-in window for Exchange…');
  await api(`/api/tenants/${tenant.clientNumber}/restart`, {
    method: 'POST',
    body: JSON.stringify({ showConsole: true }),
  }, 60000);
  await waitForWorkerReady(tenant.clientNumber);
  setStatus('authStatus', 'Complete Exchange sign-in in the window on this PC…');
  await runAuthCommand(tenant.clientNumber, 'EXCHANGE_AUTH', 'EXCHANGE_AUTH_SUCCESS', 'EXCHANGE_AUTH_FAILED', 'Exchange sign-in');
  await refreshTenants();
  setStatus('authStatus', 'Exchange sign-in finished. Run Graph Auth next.');
}

async function graphAuth() {
  await refreshTenants();
  const tenant = selectedTenant();
  if (!tenant) throw new Error('Add a tenant first.');
  if (!tenant.exchangeAuthenticated) throw new Error('Run Exchange Auth first.');
  setStatus('authStatus', 'Starting Graph sign-in…');
  const worker = await api(`/api/tenants/${tenant.clientNumber}/ensure-worker`, {
    method: 'POST',
    body: JSON.stringify({ restartIfDead: true, showConsole: true }),
  }, 60000);
  if (worker.restarted) await waitForWorkerReady(tenant.clientNumber);
  let command = 'GRAPH_AUTH';
  if (tenant.exoTenantId) command += `|TENANT_ID:${tenant.exoTenantId}`;
  else command += '|INTERACTIVE:1';
  setStatus('authStatus', 'Complete Graph sign-in in the window on this PC…');
  await runAuthCommand(tenant.clientNumber, command, 'GRAPH_AUTH_SUCCESS', 'GRAPH_AUTH_FAILED', 'Graph sign-in');
  await refreshTenants();
  setStatus('authStatus', 'Graph sign-in finished. This tenant can run a threat check.');
}

async function returnToQueue() {
  if (!state.ticket) throw new Error('Choose a ticket first.');
  const id = state.ticket.ticketId;
  const ok = window.confirm(`Send ticket #${id} back to the Security queue?\n\nStatus becomes Dispatch. You are removed as the owner and resource. The ticket is not closed.`);
  if (!ok) return;
  setStatus('scanStatus', 'Sending the ticket back to Dispatch…');
  const posted = await api('/api/manage/hunter-post', {
    method: 'POST',
    body: JSON.stringify({ ticketId: id, status: 'dispatch' }),
  }, 60000);
  if (posted && posted.statusUpdated) showTicketAssignment('Dispatch', 'clear');
  const done = asArray(posted && posted.completed).filter(Boolean);
  const errors = asArray(posted && posted.errors).filter(Boolean);
  if (posted && posted.ok === false) {
    throw new Error([...done, ...errors].join(' ') || 'Manage did not update the ticket.');
  }
  setStatus('scanStatus', done.join(' ') || 'Ticket is back on the Security queue.');
  await refreshManageLog(id);
}

$('outputCard').addEventListener('click', (event) => {
  const remove = event.target.closest('.noteDeleteBtn');
  if (remove) {
    setBusy(true);
    deleteNote(remove).catch((error) => setStatus('queueStatus', error.message, 'bad')).finally(() => setBusy(false));
    return;
  }
  const toggle = event.target.closest('.noteTimeBtn');
  if (toggle) {
    const card = toggle.closest('.threadNote');
    const panel = card && card.querySelector('.noteTime');
    if (!panel) return;
    panel.hidden = !panel.hidden;
    const start = panel.querySelector('.noteTimeStart');
    if (!panel.hidden && start && !start.value) start.value = datetimeLocalValue(new Date(Date.now() - 15 * 60 * 1000));
    return;
  }
  const add = event.target.closest('.noteTimeAdd');
  if (!add) return;
  setBusy(true);
  addTimeToNote(add).catch((error) => {
    const status = add.parentElement.querySelector('.noteTimeStatus');
    if (status) {
      status.textContent = error.message;
      status.classList.add('bad');
    }
  }).finally(() => setBusy(false));
});
$('btnBackToQueue').addEventListener('click', backToQueue);
document.querySelectorAll('.outputTab').forEach((btn) => {
  btn.addEventListener('click', () => showOutput(btn.dataset.output));
});
$('clientInsert').addEventListener('change', insertClientLine);
$('btnRefresh').addEventListener('click', () => refreshQueue().catch((e) => setStatus('queueStatus', e.message, 'bad')));
$('pickVisible').addEventListener('change', () => {
  const on = $('pickVisible').checked;
  for (const ticket of visibleTickets()) {
    if (on) state.picked.add(String(ticket.id));
    else state.picked.delete(String(ticket.id));
  }
  renderTickets();
});
$('btnEnrich').addEventListener('click', () => {
  setBusy(true);
  enrichSelected().catch((e) => setStatus('queueStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnFindUser').addEventListener('click', () => {
  setBusy(true);
  findUser().catch((e) => setStatus('userStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnTenants').addEventListener('click', () => refreshTenants().catch((e) => setStatus('authStatus', e.message, 'bad')));
$('tenantPick').addEventListener('change', syncTenantAuth);
$('btnAddTenant').addEventListener('click', () => {
  setBusy(true);
  addTenant().catch((e) => setStatus('authStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnExoAuth').addEventListener('click', () => {
  setBusy(true);
  exchangeAuth().catch((e) => setStatus('authStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnGraphAuth').addEventListener('click', () => {
  setBusy(true);
  graphAuth().catch((e) => setStatus('authStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnReturn').addEventListener('click', () => {
  setBusy(true);
  returnToQueue().catch((e) => setStatus('scanStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('mailboxPick').addEventListener('change', () => {
  if ($('mailboxPick').value) $('mailbox').value = $('mailboxPick').value;
});
$('timePick').addEventListener('change', syncHours);
$('ticketStatus').addEventListener('change', syncCloseTime);
$('closeTimePick').addEventListener('change', syncCloseTime);
$('btnSetStatus').addEventListener('click', () => {
  setBusy(true);
  setTicketStatus().catch((e) => setStatus('statusChange', e.message, 'bad')).finally(() => setBusy(false));
});
async function triageAlert() {
  if (!state.ticket) throw new Error('Choose a ticket first.');
  const alertText = alertTextFrom(state.ticket.ticketContent, state.thread);
  if (!alertText) throw new Error('This ticket has no alert text to triage.');
  setStatus('scanStatus', 'Reading the alert and writing the internal note…');
  const drafted = await draftAlertNote(state.ticket, alertText);
  const note = drafted.note;
  $('review').value = note;
  $('internalText').value = note;
  $('clientText').value = '';
  $('incidentText').value = '';
  $('useClient').checked = false;
  $('useInternal').checked = true;
  $('useIncident').checked = false;
  $('useEscalation').checked = false;
  $('statusPick').value = 'keep';
  $('timePick').value = 'none';
  syncHours();
  $('btnNotes').hidden = true;
  $('notesCard').hidden = false;
  renderThread();
  showOutput('internal');
  const refs = drafted.references || '';
  setStatus('scanStatus', refs ? `Alert review ready. ${refs}` : 'Alert review ready. Edit the internal note, then post it.');
}

$('btnTriage').addEventListener('click', () => {
  setBusy(true);
  triageAlert().catch((e) => setStatus('scanStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnScan').addEventListener('click', () => {
  setBusy(true);
  runScan().catch((e) => {
    setStatus('scanStatus', e.message, 'bad');
    if (state.audit && state.span) $('btnReviewAgain').hidden = false;
  }).finally(() => setBusy(false));
});
$('btnReviewAgain').addEventListener('click', () => {
  setBusy(true);
  reviewAgain().catch((e) => {
    setStatus('scanStatus', e.message, 'bad');
    if (state.audit && state.span) $('btnReviewAgain').hidden = false;
  }).finally(() => setBusy(false));
});
$('btnRepull').addEventListener('click', () => {
  setBusy(true);
  runRepull().catch((e) => setStatus('scanStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnReviseReview').addEventListener('click', () => {
  setBusy(true);
  reviseReview().catch((e) => setStatus('scanStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnNotes').addEventListener('click', () => {
  setBusy(true);
  writeNotes(false).catch((e) => setStatus('scanStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnReviseNotes').addEventListener('click', () => {
  setBusy(true);
  writeNotes(true).catch((e) => setStatus('scanStatus', e.message, 'bad')).finally(() => setBusy(false));
});
$('btnPost').addEventListener('click', () => {
  setBusy(true);
  postNotes().catch((e) => setStatus('postStatus', e.message, 'bad')).finally(() => setBusy(false));
});

state.queue = loadQueueSettings();
applyQueueForm();
document.querySelectorAll('button.sort').forEach((btn) => {
  btn.addEventListener('click', () => sortQueue(btn.dataset.sort));
});
$('filterOwner').addEventListener('change', readQueueForm);
$('filterStatus').addEventListener('change', readQueueForm);
$('filterKeyword').addEventListener('input', readQueueForm);
$('hideRoutine').addEventListener('change', readQueueForm);
$('btnSaveFilter').addEventListener('click', () => {
  try { saveNamedFilter(); } catch (e) { setStatus('queueStatus', e.message, 'bad'); }
});
$('savedFilters').addEventListener('change', () => {
  if ($('savedFilters').value) applySavedFilter($('savedFilters').value);
});
$('btnDeleteFilter').addEventListener('click', () => {
  try { deleteSavedFilter(); } catch (e) { setStatus('queueStatus', e.message, 'bad'); }
});
syncHours();
refreshQueue().catch((e) => setStatus('queueStatus', e.message, 'bad'));
refreshTenants().catch(() => {});
