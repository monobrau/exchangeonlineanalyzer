const logEl = document.getElementById('log');

const sessionInfo = document.getElementById('sessionInfo');

const tenantsEl = document.getElementById('tenants');

const appRegsEl = document.getElementById('appRegs');

let appRegistrations = [];

const clientBusy = new Set();

const tenantUiState = new Map();

let currentSessionId = null;
let serverFeatures = { sessionHistory: false, sessionHistoryActions: false, reportSelections: false, exportPresets: false, noWaitCommands: false, hiddenWorkers: true, wcmManagement: false, workerLogTabs: false, liongardIntegration: false, huntressIntegration: false, sentinelOneIntegration: false };
const uiStateSyncTimers = new Map();
const sessionReportSyncTimer = { id: null };
const HISTORY_SORT_STORAGE = { saved: 'bulkRunnerHistorySort.saved', archived: 'bulkRunnerHistorySort.archived' };

function parseHistorySort(value, fallback = 'updatedAt:desc') {
  const raw = value || fallback;
  const [sortBy, sortOrder] = raw.split(':');
  return {
    sortBy: sortBy || 'updatedAt',
    sortOrder: sortOrder === 'asc' ? 'asc' : 'desc',
  };
}

let savedHistoryRows = [];
let archivedHistoryRows = [];
const historySort = {
  saved: parseHistorySort(localStorage.getItem(HISTORY_SORT_STORAGE.saved)),
  archived: parseHistorySort(localStorage.getItem(HISTORY_SORT_STORAGE.archived), 'archivedAt:desc'),
};
const historySearch = { saved: '', archived: '' };
const workerLogOffsets = new Map();
const workerLogPollers = new Map();
let activeLogTab = 'activity';
let exportPresetsFromServer = [];

const HUNTRESS_PULL_OPTIONS = [
  { key: 'signalsFootholds', label: 'Signals: Footholds' },
  { key: 'signalsAntivirus', label: 'Signals: Antivirus' },
  { key: 'signalsProcessInsights', label: 'Signals: Process Insights' },
  { key: 'signalsManagedItdr', label: 'Signals: Managed ITDR' },
  { key: 'signalsSiem', label: 'Signals: SIEM' },
  { key: 'incidents', label: 'Incident reports' },
  { key: 'agents', label: 'Agents' },
  { key: 'identities', label: 'Identities' },
  { key: 'escalations', label: 'Escalations' },
];

const S1_PULL_OPTIONS = [
  { key: 'threats', label: 'Threats' },
  { key: 'agents', label: 'Agents' },
  { key: 'activities', label: 'Activities' },
];

function buildSecurityCheckboxGroup(options, prefix, disabledKeys = []) {
  return options.map((opt) => {
    const off = disabledKeys.includes(opt.key) ? ' disabled' : '';
    return `<label><input type="checkbox" class="${prefix}Pull" data-pull="${opt.key}"${off} /> ${opt.label}</label>`;
  }).join('');
}

function buildSecurityIntegrationsPanelHtml() {
  return `
      <details class="securityIntegrations collapsible">
        <summary>Security integrations <span class="securityIntegrationsHint muted">(Liongard · Huntress · SentinelOne)</span></summary>
        <div class="collapsible-body securityPanel">
          <div class="row">
            <button type="button" class="resolveSecurity small">Resolve client</button>
            <button type="button" class="previewSecurity small">Preview counts</button>
            <button type="button" class="exportSecurity primary small">Export selected pulls</button>
          </div>
          <div class="securityResolveStatus muted" style="font-size:0.85rem"></div>
          <div class="securityChips"></div>
          <div class="sectionLabel">Huntress pulls</div>
          <div class="securityGrid huntressPulls">${buildSecurityCheckboxGroup(HUNTRESS_PULL_OPTIONS, 'huntress')}</div>
          <div class="sectionLabel">SentinelOne</div>
          <div class="row">
            <label>Console profile
              <select class="s1ProfileSelect">
                <option value="connectwise">ConnectWise S1 (full read)</option>
                <option value="barracuda_xdr">Barracuda XDR S1 (read-only)</option>
              </select>
            </label>
            <label>Site ID <input type="text" class="s1SiteIdInput" placeholder="from Liongard or ticket" style="width:12rem" /></label>
          </div>
          <div class="securityNotice s1FallbackNotice" style="display:none"></div>
          <div class="securityGrid s1Pulls">${buildSecurityCheckboxGroup(S1_PULL_OPTIONS, 's1')}</div>
          <label><input type="checkbox" class="includeLiongardContext" /> Include Liongard baseline (systems + detections)</label>
          <div class="securityPreview muted" style="font-size:0.82rem;margin-top:0.35rem"></div>
        </div>
      </details>`;
}

function getTenantTicketContent(div) {
  return (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || '';
}

function getTenantCompanyName(div, clientNumber) {
  const prior = tenantUiState.get(String(clientNumber)) || {};
  return prior.organizationHint || div.dataset.organizationHint || '';
}

function applySecurityResolveToPanel(div, data) {
  const statusEl = div.querySelector('.securityResolveStatus');
  const chipsEl = div.querySelector('.securityChips');
  const huntressSection = div.querySelector('.huntressPulls');
  const s1Section = div.querySelector('.s1Pulls');
  const s1Select = div.querySelector('.s1ProfileSelect');
  const s1Site = div.querySelector('.s1SiteIdInput');
  const fallback = div.querySelector('.s1FallbackNotice');

  div.dataset.liongardEnvironmentId = data.matched ? String(data.environmentId) : '';
  div.dataset.huntressOrganizationId = data.huntressOrgId ? String(data.huntressOrgId) : (div.dataset.huntressOrganizationId || '');
  div.dataset.securityResolved = data.matched ? '1' : '0';

  if (statusEl) {
    if (data.matched) {
      statusEl.textContent = `Liongard: ${data.environmentName} (#${data.environmentId}, score ${data.matchScore}) · SOC: ${data.socSource || 'unknown'}`;
    } else if (data.error) {
      statusEl.textContent = data.error;
    } else {
      statusEl.textContent = 'No Liongard environment match — manual org/site selection still available.';
    }
  }

  const chips = [];
  const stack = data.securityStack || {};
  if (stack.alertTitle) {
    chips.push(`<span class="securityChip on" title="Alert from Manage ticket">${escapeHtml(stack.alertTitle)}</span>`);
  }
  (stack.labels || []).forEach((label) => {
    if (label && label !== stack.alertTitle) {
      chips.push(`<span class="securityChip on">${escapeHtml(label)}</span>`);
    }
  });
  if (data.huntress?.hasHuntress && stack.useHuntress !== false) {
    if (data.huntress.edr) chips.push('<span class="securityChip on">Huntress EDR</span>');
    if (data.huntress.itdr) chips.push('<span class="securityChip on">Huntress ITDR</span>');
    if (data.huntress.siem) chips.push('<span class="securityChip on">Huntress SIEM</span>');
  }
  if (data.sentinelOne?.hasSentinelOne && (stack.useS1ConnectWise || stack.useS1Barracuda)) {
    chips.push(`<span class="securityChip on">SentinelOne${data.sentinelOne.siteName ? `: ${escapeHtml(data.sentinelOne.siteName)}` : ''}</span>`);
  }
  if (stack.useS1Barracuda || data.socSource === 'barracuda_xdr') chips.push('<span class="securityChip warn">Barracuda XDR stack</span>');
  if (stack.useS1ConnectWise || data.socSource === 'connectwise') chips.push('<span class="securityChip on">ConnectWise MDR stack</span>');
  if (stack.identitySource === 'microsoft' || stack.isIdentityAlert) chips.push('<span class="securityChip on">M365 / Graph investigation</span>');
  if (chipsEl) chipsEl.innerHTML = chips.join('');

  const huntressWanted = stack.useHuntress !== false && (stack.useHuntress || data.huntress?.hasHuntress);
  const s1Wanted = stack.useS1ConnectWise || stack.useS1Barracuda;
  const recH = data.recommendedPulls?.huntress || {};
  const recS = data.recommendedPulls?.sentinelOne || {};
  div.querySelectorAll('.huntressPull').forEach((cb) => {
    const key = cb.dataset.pull;
    cb.checked = Boolean(recH[key]);
    cb.disabled = !huntressWanted || !data.huntress?.hasHuntress;
  });
  div.querySelectorAll('.s1Pull').forEach((cb) => {
    const key = cb.dataset.pull;
    cb.checked = Boolean(recS[key]);
    cb.disabled = !s1Wanted;
  });

  if (statusEl && stack.alertTitle) {
    const base = statusEl.textContent || '';
    if (!base.includes('Alert:')) {
      statusEl.textContent = `Alert: ${stack.alertTitle}${base ? ` · ${base}` : ''}`;
    }
  }

  if (huntressSection) {
    huntressSection.style.opacity = huntressWanted && data.huntress?.hasHuntress ? '1' : '0.45';
    huntressSection.title = huntressWanted ? '' : 'Ticket alert does not indicate Huntress — expand if client uses Huntress anyway';
  }
  if (s1Section) {
    s1Section.style.opacity = s1Wanted ? '1' : '0.45';
    s1Section.title = s1Wanted ? '' : 'Ticket alert does not indicate SentinelOne endpoint stack';
  }

  if (s1Select) {
    if (stack.useS1Barracuda) s1Select.value = 'barracuda_xdr';
    else if (stack.useS1ConnectWise) s1Select.value = 'connectwise';
    else {
      const hint = data.sentinelOne?.consoleHint || data.socSource;
      if (hint && hint !== 'unknown') s1Select.value = hint === 'barracuda_xdr' ? 'barracuda_xdr' : 'connectwise';
    }
  }
  if (s1Site && data.sentinelOne?.siteId && !s1Site.value) s1Site.value = data.sentinelOne.siteId;

  if (fallback) {
    if (data.s1Resolve?.barracudaFallback) {
      fallback.textContent = data.s1Resolve.barracudaFallback;
      fallback.style.display = 'block';
    } else {
      fallback.style.display = 'none';
    }
  }
}

async function resolveSecurityIntegrations(clientNumber, div) {
  const companyName = getTenantCompanyName(div, clientNumber);
  if (!companyName) {
    log(`Client ${clientNumber}: resolve security — set company via Manage ticket fetch first.`);
    return;
  }
  const btn = div.querySelector('.resolveSecurity');
  if (btn) btn.disabled = true;
  try {
    const ticketContent = getTenantTicketContent(div);
    const data = await api('/api/liongard/resolve-client', {
      method: 'POST',
      body: JSON.stringify({ companyName, ticketContent }),
    }, 45000);

    if (data.huntress?.hasHuntress) {
      try {
        const orgMatch = await api('/api/huntress/organizations', { method: 'GET' }, 45000);
        const orgs = orgMatch.organizations || [];
        const norm = companyName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
        const hit = orgs.find((o) => {
          const n = (o.name || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
          return n === norm || n.includes(norm) || norm.includes(n);
        });
        if (hit) data.huntressOrgId = hit.id;
      } catch (_) {}
    }

    try {
      data.s1Resolve = await api('/api/sentinelone/resolve-site', {
        method: 'POST',
        body: JSON.stringify({
          companyName,
          ticketContent,
          liongardEnvironmentId: data.environmentId || 0,
          profileName: div.querySelector('.s1ProfileSelect')?.value || '',
          siteId: div.querySelector('.s1SiteIdInput')?.value || data.sentinelOne?.siteId || '',
        }),
      }, 30000);
    } catch (e) {
      data.s1Resolve = { barracudaFallback: e.message };
    }

    applySecurityResolveToPanel(div, data);
    div.dataset.securityResolveJson = JSON.stringify({
      environmentId: data.environmentId,
      huntressOrgId: data.huntressOrgId,
      socSource: data.socSource,
    });
    saveTenantUiState(clientNumber, div);
    log(`Client ${clientNumber}: security resolve complete.`);
  } catch (e) {
    applySecurityResolveToPanel(div, { matched: false, error: e.message, huntress: {}, sentinelOne: {}, recommendedPulls: {} });
    log(`Client ${clientNumber}: security resolve failed: ${e.message}`);
  } finally {
    if (btn) btn.disabled = false;
  }
}

function readHuntressSelections(div) {
  const out = {};
  div.querySelectorAll('.huntressPull:checked').forEach((cb) => { out[cb.dataset.pull] = true; });
  return out;
}

function readS1Selections(div) {
  const out = {};
  div.querySelectorAll('.s1Pull:checked').forEach((cb) => { out[cb.dataset.pull] = true; });
  return out;
}

async function previewSecurityIntegrations(clientNumber, div) {
  const previewEl = div.querySelector('.securityPreview');
  const resolved = JSON.parse(div.dataset.securityResolveJson || '{}');
  const dateStart = div.querySelector('.dateStart')?.value;
  const dateEnd = div.querySelector('.dateEnd')?.value;
  const parts = [];

  if (resolved.huntressOrgId) {
    try {
      const body = { organizationId: resolved.huntressOrgId };
      if (dateStart) body.updatedSince = new Date(dateStart).toISOString();
      const p = await api('/api/huntress/preview', { method: 'POST', body: JSON.stringify(body) }, 120000);
      parts.push(`Huntress: ${JSON.stringify(p.counts)}`);
    } catch (e) {
      parts.push(`Huntress preview: ${e.message}`);
    }
  }

  const profileName = div.querySelector('.s1ProfileSelect')?.value;
  if (profileName) {
    try {
      const body = { profileName, siteId: div.querySelector('.s1SiteIdInput')?.value || '' };
      if (dateStart) body.createdAfter = new Date(dateStart).toISOString();
      if (dateEnd) body.createdBefore = new Date(dateEnd).toISOString();
      const p = await api('/api/sentinelone/preview', { method: 'POST', body: JSON.stringify(body) }, 120000);
      parts.push(`S1 (${profileName}): ${JSON.stringify(p.counts)}`);
    } catch (e) {
      parts.push(`S1 preview: ${e.message}`);
    }
  }

  if (previewEl) previewEl.textContent = parts.join(' · ') || 'Resolve client first.';
}

async function exportSecurityIntegrations(clientNumber, div) {
  const companyName = getTenantCompanyName(div, clientNumber);
  const ticketNumber = (div.querySelector('.ticketInput')?.value || '').trim();
  const outputFolder = div.dataset.outputFolder || '';
  const resolved = JSON.parse(div.dataset.securityResolveJson || '{}');
  const dateStart = div.querySelector('.dateStart')?.value;
  const dateEnd = div.querySelector('.dateEnd')?.value;
  const exported = [];

  const btn = div.querySelector('.exportSecurity');
  if (btn) btn.disabled = true;
  try {
    if (div.querySelector('.includeLiongardContext')?.checked && resolved.environmentId) {
      const body = { environmentId: resolved.environmentId, companyName, outputFolder, ticketNumber };
      if (dateStart) body.startDate = new Date(dateStart).toISOString();
      if (dateEnd) body.endDate = new Date(dateEnd).toISOString();
      const r = await api('/api/liongard/export-context', { method: 'POST', body: JSON.stringify(body) }, 120000);
      exported.push(...(r.files || []));
    }

    const huntressSelections = readHuntressSelections(div);
    if (resolved.huntressOrgId && Object.keys(huntressSelections).length) {
      const body = {
        organizationId: resolved.huntressOrgId,
        companyName,
        outputFolder,
        ticketNumber,
        selections: huntressSelections,
      };
      if (dateStart) body.updatedSince = new Date(dateStart).toISOString();
      const r = await api('/api/huntress/export', { method: 'POST', body: JSON.stringify(body) }, 300000);
      exported.push(...(r.files || []));
    }

    const s1Selections = readS1Selections(div);
    const profileName = div.querySelector('.s1ProfileSelect')?.value;
    if (profileName && Object.keys(s1Selections).length) {
      const body = {
        profileName,
        companyName,
        outputFolder,
        ticketNumber,
        siteId: div.querySelector('.s1SiteIdInput')?.value || '',
        selections: s1Selections,
      };
      if (dateStart) body.createdAfter = new Date(dateStart).toISOString();
      if (dateEnd) body.createdBefore = new Date(dateEnd).toISOString();
      const r = await api('/api/sentinelone/export', { method: 'POST', body: JSON.stringify(body) }, 300000);
      exported.push(...(r.files || []));
    }

    log(`Client ${clientNumber}: exported ${exported.length} security integration file(s).`);
    if (exported.length) log(exported.join('\n'));
  } catch (e) {
    log(`Client ${clientNumber}: security export failed: ${e.message}`);
  } finally {
    if (btn) btn.disabled = false;
  }
}

function wireSecurityIntegrationsPanel(div, clientNumber) {
  div.querySelector('.resolveSecurity')?.addEventListener('click', () => resolveSecurityIntegrations(clientNumber, div));
  div.querySelector('.previewSecurity')?.addEventListener('click', () => previewSecurityIntegrations(clientNumber, div));
  div.querySelector('.exportSecurity')?.addEventListener('click', () => exportSecurityIntegrations(clientNumber, div));
  div.querySelectorAll('.huntressPull, .s1Pull, .includeLiongardContext, .s1ProfileSelect, .s1SiteIdInput').forEach((el) => {
    el.addEventListener('change', () => saveTenantUiState(clientNumber, div));
  });
}

function restoreSecurityUiState(div, saved) {
  if (!saved?.security) return;
  const s = saved.security;
  if (s.resolveJson) div.dataset.securityResolveJson = s.resolveJson;
  if (s.s1Profile && div.querySelector('.s1ProfileSelect')) div.querySelector('.s1ProfileSelect').value = s.s1Profile;
  if (s.s1SiteId && div.querySelector('.s1SiteIdInput')) div.querySelector('.s1SiteIdInput').value = s.s1SiteId;
  if (div.querySelector('.includeLiongardContext')) div.querySelector('.includeLiongardContext').checked = Boolean(s.includeLiongardContext);
  if (s.huntressPulls) {
    div.querySelectorAll('.huntressPull').forEach((cb) => { cb.checked = Boolean(s.huntressPulls[cb.dataset.pull]); });
  }
  if (s.s1Pulls) {
    div.querySelectorAll('.s1Pull').forEach((cb) => { cb.checked = Boolean(s.s1Pulls[cb.dataset.pull]); });
  }
  if (s.resolveJson) {
    try {
      const partial = JSON.parse(s.resolveJson);
      applySecurityResolveToPanel(div, {
        matched: Boolean(partial.environmentId),
        environmentId: partial.environmentId,
        environmentName: saved.organizationHint || '',
        matchScore: 100,
        socSource: partial.socSource,
        huntress: { hasHuntress: Boolean(partial.huntressOrgId), edr: true },
        sentinelOne: { hasSentinelOne: true, siteId: s.s1SiteId },
        recommendedPulls: {},
        huntressOrgId: partial.huntressOrgId,
      });
    } catch (_) {}
  }
}

const REPORT_EXPORT_GROUPS = [
  {
    title: 'Exchange Online',
    items: [
      { key: 'IncludeMessageTrace', label: 'Message trace' },
      { key: 'IncludeInboxRules', label: 'Inbox rules' },
      { key: 'IncludeTransportRules', label: 'Transport rules' },
      { key: 'IncludeMailFlowConnectors', label: 'Mail flow connectors' },
      { key: 'IncludeMailboxForwarding', label: 'Mailbox forwarding' },
      { key: 'IncludeUnifiedAuditLogs', label: 'Unified audit logs (EXO)' },
    ],
  },
  {
    title: 'Microsoft Graph — identity',
    items: [
      { key: 'IncludeSignInLogs', label: 'Sign-in logs' },
      { key: 'IncludeAuditLogs', label: 'Directory audit logs' },
      { key: 'IncludeMfaCoverage', label: 'MFA coverage' },
      { key: 'IncludeConditionalAccessPolicies', label: 'Conditional Access policies' },
      { key: 'IncludeAppRegistrations', label: 'App registrations' },
      { key: 'IncludeIntuneDevices', label: 'Intune devices' },
    ],
  },
  {
    title: 'Microsoft Graph — security & activity',
    items: [
      { key: 'IncludeSecurityAlerts', label: 'Security alerts' },
      { key: 'IncludeSecurityIncidents', label: 'Security incidents' },
      { key: 'IncludeSharePointActivity', label: 'SharePoint activity' },
      { key: 'IncludeOneDriveActivity', label: 'OneDrive activity' },
      { key: 'IncludeTeamsActivity', label: 'Teams activity' },
      { key: 'IncludeSharePointSharing', label: 'SharePoint sharing links' },
      { key: 'IncludeAnonymousSharePointSharing', label: 'Anonymous SharePoint sharing' },
      { key: 'IncludeSharePointFileSharingLinks', label: 'SharePoint file sharing links' },
      { key: 'IncludeDLPViolations', label: 'DLP violations' },
      { key: 'IncludeSharePointOneDriveFileActions', label: 'SharePoint/OneDrive file actions' },
    ],
  },
  {
    title: 'Third-party SOC (use Security integrations panel to export)',
    items: [
      { key: 'IncludeHuntressSignals', label: 'Huntress signals (preset flag)' },
      { key: 'IncludeHuntressIncidents', label: 'Huntress incidents (preset flag)' },
      { key: 'IncludeHuntressAgents', label: 'Huntress agents (preset flag)' },
      { key: 'IncludeS1Threats', label: 'SentinelOne threats (preset flag)' },
      { key: 'IncludeS1Agents', label: 'SentinelOne agents (preset flag)' },
      { key: 'IncludeS1Activities', label: 'SentinelOne activities (preset flag)' },
      { key: 'IncludeLiongardContext', label: 'Liongard baseline (preset flag)' },
    ],
  },
];

/** Fallback if /api/export-presets is unavailable — matches Settings Get-BecExportPresetSelections. */
const DEFAULT_BEC_PRESET_NAME = 'BEC / Business Email Compromise';

const REPORT_EXPORT_PRESETS = {
  [DEFAULT_BEC_PRESET_NAME]: {
    IncludeMessageTrace: true,
    IncludeUnifiedAuditLogs: true,
    IncludeInboxRules: true,
    IncludeTransportRules: true,
    IncludeMailFlowConnectors: false,
    IncludeMailboxForwarding: true,
    IncludeAuditLogs: true,
    IncludeSignInLogs: true,
    IncludeMfaCoverage: true,
    IncludeConditionalAccessPolicies: true,
    IncludeAppRegistrations: true,
    IncludeSecurityAlerts: true,
    IncludeSecurityIncidents: true,
    IncludeIntuneDevices: true,
    IncludeSharePointActivity: false,
    IncludeOneDriveActivity: false,
    IncludeTeamsActivity: false,
    IncludeSharePointSharing: false,
    IncludeAnonymousSharePointSharing: false,
    IncludeSharePointFileSharingLinks: false,
    IncludeDLPViolations: false,
    IncludeSharePointOneDriveFileActions: false,
  },
};

function defaultReportSelections() {
  return {
    ...REPORT_EXPORT_PRESETS[DEFAULT_BEC_PRESET_NAME],
    SignInLogsDaysBack: 7,
    MessageTraceDaysBack: 7,
  };
}

function reportExportCheckboxClass(key) {
  return `rs_${key}`;
}

function buildReportExportsPanelHtml(scope) {
  const html = [];
  for (const group of REPORT_EXPORT_GROUPS) {
    html.push(`<div class="reportExportGroup"><div class="reportExportGroupTitle">${group.title}</div><div class="reportExportGrid">`);
    for (const item of group.items) {
      html.push(`<label><input type="checkbox" class="${reportExportCheckboxClass(item.key)}" data-rs-key="${item.key}" /> ${item.label}</label>`);
    }
    html.push('</div></div>');
  }
  return html.join('');
}

function readReportSelectionsFromContainer(container) {
  const rs = defaultReportSelections();
  if (!container) return rs;
  for (const group of REPORT_EXPORT_GROUPS) {
    for (const item of group.items) {
      const el = container.querySelector(`.${reportExportCheckboxClass(item.key)}`);
      if (el) rs[item.key] = Boolean(el.checked);
    }
  }
  const mt = document.getElementById('messageTraceDays');
  const si = document.getElementById('signInLogsDays');
  if (mt) rs.MessageTraceDaysBack = parseInt(mt.value, 10) || 7;
  if (si) rs.SignInLogsDaysBack = parseInt(si.value, 10) || 7;
  return rs;
}

function applyReportSelectionsToContainer(container, selections) {
  if (!container || !selections) return;
  const rs = { ...defaultReportSelections(), ...selections };
  for (const group of REPORT_EXPORT_GROUPS) {
    for (const item of group.items) {
      const el = container.querySelector(`.${reportExportCheckboxClass(item.key)}`);
      if (el) el.checked = Boolean(rs[item.key]);
    }
  }
}

function countEnabledReportExports(selections) {
  if (!selections) return 0;
  let n = 0;
  for (const group of REPORT_EXPORT_GROUPS) {
    for (const item of group.items) {
      if (selections[item.key]) n += 1;
    }
  }
  return n;
}

function getRequiredAuthFromReportSelections(selections) {
  const graphReports = [
    'IncludeAuditLogs', 'IncludeSignInLogs', 'IncludeMfaCoverage', 'IncludeConditionalAccessPolicies',
    'IncludeAppRegistrations', 'IncludeSecurityAlerts', 'IncludeSecurityIncidents', 'IncludeIntuneDevices',
    'IncludeSharePointActivity', 'IncludeOneDriveActivity', 'IncludeTeamsActivity', 'IncludeSharePointSharing',
  ];
  const exchangeOnlyReports = [
    'IncludeMessageTrace', 'IncludeUnifiedAuditLogs', 'IncludeTransportRules',
    'IncludeMailFlowConnectors', 'IncludeMailboxForwarding',
  ];
  let needsGraph = graphReports.some((k) => selections?.[k] === true);
  let needsExchange = exchangeOnlyReports.some((k) => selections?.[k] === true);
  if (selections?.IncludeInboxRules === true) {
    if (needsExchange) needsExchange = true;
    else needsGraph = true;
  }
  return { needsGraph, needsExchange };
}

function getEffectiveReportSelectionsForTenant(t, session) {
  const ui = t?.uiState;
  const useDefaults = ui?.useSessionReportDefaults !== false;
  if (!useDefaults && ui?.reportSelections) {
    return { ...defaultReportSelections(), ...ui.reportSelections };
  }
  return { ...defaultReportSelections(), ...(session?.reportSelections || {}) };
}

function tenantHasRequiredAuth(t, session) {
  const req = getRequiredAuthFromReportSelections(getEffectiveReportSelectionsForTenant(t, session));
  if (req.needsGraph && !t.graphAuthenticated) return false;
  if (req.needsExchange && !t.exchangeAuthenticated) return false;
  return true;
}

function updateSessionReportExportsSummary(selections) {
  const el = document.getElementById('sessionReportExportsSummary');
  if (!el) return;
  const n = countEnabledReportExports(selections || readReportSelectionsFromContainer(document.getElementById('sessionReportExportsBody')));
  el.textContent = n ? `(${n} enabled)` : '(none enabled)';
}

function updateTenantReportExportsHint(div, useDefaults) {
  const hint = div?.querySelector('.reportExportsHint');
  if (!hint) return;
  if (useDefaults) {
    hint.textContent = '(session defaults)';
    hint.classList.remove('customized');
  } else {
    hint.textContent = '(custom for this client)';
    hint.classList.add('customized');
  }
}

function scheduleSessionReportSelectionsSync() {
  if (!serverFeatures.reportSelections) return;
  if (sessionReportSyncTimer.id) clearTimeout(sessionReportSyncTimer.id);
  sessionReportSyncTimer.id = setTimeout(async () => {
    sessionReportSyncTimer.id = null;
    try {
      const session = await api('/api/session');
      if (!hasActiveSession(session)) return;
      const body = {
        reportSelections: readReportSelectionsFromContainer(document.getElementById('sessionReportExportsBody')),
        daysBack: parseInt(document.getElementById('daysBack')?.value, 10) || 7,
      };
      const updated = await api('/api/session/report-selections', {
        method: 'POST',
        body: JSON.stringify(body),
      });
      updateSessionReportExportsSummary(updated.reportSelections);
    } catch (e) {
      if (String(e.message || e).includes('Not found')) {
        serverFeatures.reportSelections = false;
        return;
      }
      log(`Report defaults sync: ${e.message}`);
    }
  }, 700);
}

function setAllReportSelections(container, enabled) {
  if (!container) return;
  for (const group of REPORT_EXPORT_GROUPS) {
    for (const item of group.items) {
      const el = container.querySelector(`.${reportExportCheckboxClass(item.key)}`);
      if (el) el.checked = enabled;
    }
  }
}

async function loadExportPresetsFromServer(options = {}) {
  const applyDefaultBec = options.applyDefaultBec !== false;
  if (!serverFeatures.exportPresets) {
    if (applyDefaultBec) {
      const body = document.getElementById('sessionReportExportsBody');
      if (body) {
        applyReportSelectionsToContainer(body, defaultReportSelections());
        updateSessionReportExportsSummary(defaultReportSelections());
      }
    }
    return;
  }
  try {
    const data = await api('/api/export-presets');
    exportPresetsFromServer = data.presets || [];
    const sel = document.getElementById('sessionReportPreset');
    if (!sel) return;
    const previous = sel.value;
    sel.innerHTML = '';
    for (const p of exportPresetsFromServer) {
      const opt = document.createElement('option');
      opt.value = p.name;
      opt.textContent = p.name;
      sel.appendChild(opt);
    }
    const becName =
      exportPresetsFromServer.find(p => p.name === DEFAULT_BEC_PRESET_NAME)?.name ||
      exportPresetsFromServer.find(p => p.selections)?.name ||
      '';
    const preferPrevious = previous && [...sel.options].some(o => o.value === previous);
    const chosen = preferPrevious ? previous : becName;
    if (chosen) sel.value = chosen;
    if (applyDefaultBec && chosen && !String(chosen).startsWith('Custom')) {
      const body = document.getElementById('sessionReportExportsBody');
      if (body && applyExportPresetByName(chosen, body)) {
        updateSessionReportExportsSummary(readReportSelectionsFromContainer(body));
      }
    }
  } catch {
    exportPresetsFromServer = [];
  }
}

function applyExportPresetByName(name, container) {
  const preset = exportPresetsFromServer.find(p => p.name === name);
  if (!preset || !preset.selections) return false;
  applyReportSelectionsToContainer(container, { ...defaultReportSelections(), ...preset.selections });
  return true;
}

function initSessionReportExportsPanel() {
  const body = document.getElementById('sessionReportExportsBody');
  if (!body || body.dataset.initialized === '1') return;
  body.innerHTML = buildReportExportsPanelHtml('session');
  body.dataset.initialized = '1';
  const markPresetCustom = () => {
    const sel = document.getElementById('sessionReportPreset');
    if (!sel) return;
    const custom = [...sel.options].find(o => String(o.value).startsWith('Custom'));
    if (custom) sel.value = custom.value;
  };
  body.addEventListener('change', () => {
    markPresetCustom();
    updateSessionReportExportsSummary(readReportSelectionsFromContainer(body));
    scheduleSessionReportSelectionsSync();
  });
  document.getElementById('sessionReportPreset')?.addEventListener('change', (e) => {
    const name = e.target.value;
    if (!name || name.startsWith('Custom')) return;
    if (applyExportPresetByName(name, body)) {
      updateSessionReportExportsSummary(readReportSelectionsFromContainer(body));
      scheduleSessionReportSelectionsSync();
      return;
    }
    const legacy = REPORT_EXPORT_PRESETS[e.target.value];
    if (legacy) {
      applyReportSelectionsToContainer(body, { ...defaultReportSelections(), ...legacy });
      updateSessionReportExportsSummary(readReportSelectionsFromContainer(body));
      scheduleSessionReportSelectionsSync();
    }
  });
  document.getElementById('btnReportSelectAll')?.addEventListener('click', () => {
    setAllReportSelections(body, true);
    markPresetCustom();
    updateSessionReportExportsSummary(readReportSelectionsFromContainer(body));
    scheduleSessionReportSelectionsSync();
  });
  document.getElementById('btnReportSelectNone')?.addEventListener('click', () => {
    setAllReportSelections(body, false);
    markPresetCustom();
    updateSessionReportExportsSummary(readReportSelectionsFromContainer(body));
    scheduleSessionReportSelectionsSync();
  });
  ['messageTraceDays', 'signInLogsDays', 'daysBack'].forEach(id => {
    document.getElementById(id)?.addEventListener('change', () => scheduleSessionReportSelectionsSync());
  });
}

function wireTenantReportExportsPanel(div, clientNumber) {
  const useDefaultsCheck = div.querySelector('.useSessionReportDefaults');
  const customBody = div.querySelector('.tenantReportExportsCustom');
  const applyVisibility = (skipSave = false) => {
    const useDefaults = Boolean(useDefaultsCheck?.checked);
    if (customBody) customBody.style.display = useDefaults ? 'none' : 'block';
    updateTenantReportExportsHint(div, useDefaults);
    if (!skipSave) saveTenantUiState(clientNumber, div);
  };
  useDefaultsCheck?.addEventListener('change', () => {
    if (!useDefaultsCheck.checked && customBody) {
      const saved = tenantUiState.get(String(clientNumber));
      if (!saved?.reportSelections) {
        applyReportSelectionsToContainer(customBody, readReportSelectionsFromContainer(document.getElementById('sessionReportExportsBody')));
      }
    }
    applyVisibility(false);
  });
  customBody?.addEventListener('change', () => saveTenantUiState(clientNumber, div));
  if (customBody && !customBody.querySelector('.tenantReportSelectAll')) {
    const bar = document.createElement('div');
    bar.className = 'row';
    bar.innerHTML = '<button type="button" class="small tenantReportSelectAll">Select all</button><button type="button" class="small tenantReportSelectNone">Deselect all</button>';
    customBody.prepend(bar);
    bar.querySelector('.tenantReportSelectAll')?.addEventListener('click', () => {
      setAllReportSelections(customBody, true);
      saveTenantUiState(clientNumber, div);
    });
    bar.querySelector('.tenantReportSelectNone')?.addEventListener('click', () => {
      setAllReportSelections(customBody, false);
      saveTenantUiState(clientNumber, div);
    });
  }
  applyVisibility(true);
}

function workerCommandBody(command) {
  return JSON.stringify({ command, noWait: true, waitSeconds: 0 });
}

async function ensureWorkerForCommands(clientNumber, { restartIfDead = true, showConsole = false } = {}) {
  return api(`/api/tenants/${clientNumber}/ensure-worker`, {
    method: 'POST',
    body: JSON.stringify({ restartIfDead, showConsole }),
  });
}

async function requireLiveWorker(clientNumber, { restartIfDead = true, showConsole = false, actionLabel = 'continue' } = {}) {
  const worker = await ensureWorkerForCommands(clientNumber, { restartIfDead, showConsole });
  if (worker.restarted) {
    log(`Client ${clientNumber}: PowerShell worker was not running — restarted (PID ${worker.processId}). Re-run Exchange Auth and Graph Auth, then ${actionLabel}.`);
    await refreshSession();
    return null;
  }
  if (!worker.alive) {
    log(`Client ${clientNumber}: PowerShell worker is not running. Click Restart worker, re-authenticate, then ${actionLabel}.`);
    return null;
  }
  return worker;
}


function tenantCollapseStorageKey(clientNumber) {
  const sid = currentSessionId || 'none';
  return `eoa_bulk_tenant_open_${sid}_${clientNumber}`;
}

function isTenantCollapsed(clientNumber) {
  return localStorage.getItem(tenantCollapseStorageKey(clientNumber)) === '0';
}

function setTenantCollapsed(clientNumber, collapsed) {
  localStorage.setItem(tenantCollapseStorageKey(clientNumber), collapsed ? '0' : '1');
}

function formatHistoryWhen(iso) {
  if (!iso) return '—';
  try {
    return new Date(iso).toLocaleString();
  } catch {
    return iso;
  }
}

function tenantSummaryTitle(t, div) {
  const org = t.exoOrganizationName || t.graphTenantName
    || div?.querySelector('.appRegSelect')?.selectedOptions?.[0]?.textContent?.trim()
    || t.uiState?.organizationHint || '';
  const ticket = div?.querySelector('.ticketInput')?.value?.trim() || t.uiState?.ticket || '';
  const parts = [`Client ${t.clientNumber}`];
  if (org && !org.startsWith('(')) parts.push(org);
  if (ticket) parts.push(`#${ticket}`);
  return parts.join(' · ');
}

function resolveTenantOutputFolder(t, div) {
  return t.outputFolder || div?.dataset?.outputFolder
    || tenantUiState.get(String(t.clientNumber))?.priorOutputFolder || '';
}

function getTenantStatusState(t, div) {
  const lastResp = normalizeResponse(t.lastResponse || div?.dataset?.lastResponse || '');
  const outputFolder = resolveTenantOutputFolder(t, div);
  const clientKey = String(t.clientNumber);
  const busy = t.reportInProgress || clientBusy.has(clientKey);
  const inGraph = lastResp === 'GRAPH_AUTH_STARTED';
  const inExo = lastResp === 'EXCHANGE_AUTH_STARTED';
  const failed = lastResp.includes('_FAILED') || lastResp.includes('ERROR');
  const workerAlive = t.workerAlive !== false;

  if (!workerAlive) {
    return { code: 'failed', label: 'Worker stopped — restart & re-auth', openable: false };
  }
  if (outputFolder && !busy) {
    return { code: 'complete', label: 'Reports ready — open folder', outputFolder, openable: true };
  }
  if (busy || t.reportInProgress) {
    return { code: 'generating', label: 'Generating…', openable: false };
  }
  if (inGraph) return { code: 'auth', label: 'Graph auth…', openable: false };
  if (inExo) return { code: 'auth', label: 'EXO auth…', openable: false };
  if (failed) return { code: 'failed', label: 'Failed — expand for details', openable: false };

  const g = Boolean(t.graphAuthenticated || div?.dataset?.graphAuthenticated === '1');
  const e = Boolean(t.exchangeAuthenticated || div?.dataset?.exchangeAuthenticated === '1');
  const authShort = `G${g ? '✓' : '○'} E${e ? '✓' : '○'}`;
  if (g && e) return { code: 'ready', label: `${authShort} · Ready`, openable: false };
  return { code: 'auth-needed', label: `${authShort} · Need auth`, openable: false };
}

function applyTenantSummary(details, t, div) {
  if (!details) return;
  const titleEl = details.querySelector('.tenantSummaryTitle');
  const statusSlot = details.querySelector('.tenantSummaryStatus');
  if (titleEl) titleEl.textContent = tenantSummaryTitle(t, div);

  const status = getTenantStatusState(t, div);
  if (!statusSlot) return;

  if (status.openable && status.outputFolder) {
    statusSlot.innerHTML = `<button type="button" class="tenantStatusBtn success">${escapeHtml(status.label)}</button>`;
    const btn = statusSlot.querySelector('.tenantStatusBtn');
    if (btn) {
      btn.dataset.outputFolder = status.outputFolder;
      btn.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); });
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        openReports(btn.dataset.outputFolder);
      });
    }
  } else {
    statusSlot.innerHTML = `<span class="tenantStatusBadge status-${status.code}">${escapeHtml(status.label)}</span>`;
  }
}

function getTenantBodyEl(clientNumber) {

  const details = tenantsEl.querySelector(`details.tenant[data-client="${clientNumber}"]`);

  return details?.querySelector('.tenantBody') || null;

}

function refreshTenantSummaryUI(clientNumber, partial = {}) {
  const details = tenantsEl.querySelector(`details.tenant[data-client="${clientNumber}"]`);
  if (!details) return;
  const div = details.querySelector('.tenantBody');
  if (!div) return;
  const t = {
    clientNumber,
    exoOrganizationName: partial.exoOrganizationName,
    graphTenantName: partial.graphTenantName,
    uiState: tenantUiState.get(String(clientNumber)),
    graphAuthenticated: partial.graphAuthenticated ?? div.dataset.graphAuthenticated === '1',
    exchangeAuthenticated: partial.exchangeAuthenticated ?? div.dataset.exchangeAuthenticated === '1',
    workerAlive: partial.workerAlive ?? div.dataset.workerAlive !== '0',
    reportInProgress: partial.reportInProgress ?? false,
    outputFolder: partial.outputFolder ?? div.dataset.outputFolder,
    lastResponse: partial.lastResponse ?? div.dataset.lastResponse ?? '',
  };
  applyTenantSummary(details, t, div);
}

function tenantSummaryLabel(t, div) {
  return tenantSummaryTitle(t, div);
}

function scheduleTenantUiStateSync(clientNumber, div, immediate = false) {
  const key = String(clientNumber);
  if (uiStateSyncTimers.has(key)) clearTimeout(uiStateSyncTimers.get(key));
  if (immediate) {
    syncTenantUiStateToServer(clientNumber, div).catch(() => {});
    return;
  }
  uiStateSyncTimers.set(key, setTimeout(() => {
    uiStateSyncTimers.delete(key);
    syncTenantUiStateToServer(clientNumber, div).catch(() => {});
  }, 600));
}

async function syncTenantUiStateToServer(clientNumber, div) {
  saveTenantUiState(clientNumber, div);
  const state = tenantUiState.get(String(clientNumber));
  if (!state) return;
  const payload = {
    ...state,
    ticketContentLength: (state.ticketContent || '').length,
  };
  delete payload.ticketContent;
  await api(`/api/tenants/${clientNumber}/ui-state`, {
    method: 'POST',
    body: JSON.stringify(payload),
  });
}

async function removeTenant(clientNumber, div) {
  try {
    const session = await api('/api/session');
    const t = session.tenants?.find(x => Number(x.clientNumber) === Number(clientNumber));
    if (t?.reportInProgress) {
      if (!window.confirm(`Client ${clientNumber} is generating reports. Remove anyway?`)) return;
    } else if (!window.confirm(`Remove Client ${clientNumber} from this session? The PowerShell worker window will be stopped.`)) {
      return;
    }
    saveTenantUiState(clientNumber, div);
    clientBusy.delete(String(clientNumber));
    await api(`/api/tenants/${clientNumber}`, { method: 'DELETE' });
    tenantUiState.delete(String(clientNumber));
    log(`Client ${clientNumber} removed.`);
    await refreshSession();
  } catch (e) {
    log(`Remove Client ${clientNumber} failed: ${e.message}`);
  }
}

function escapeHtml(text) {
  return String(text ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function historySortIndicator(sortBy, sortOrder) {
  return sortOrder === 'asc' ? '▲' : '▼';
}

function formatHistoryClients(row) {
  if (row.clients?.length) {
    return row.clients.map(c => {
      const parts = [`Client ${c.clientNumber}`];
      if (c.organization) parts.push(c.organization);
      if (c.ticket) parts.push(`#${c.ticket}`);
      return escapeHtml(parts.join(' · '));
    }).join('<br>');
  }
  const orgs = row.organizations || [];
  if (orgs.length) {
    return orgs.map((o, i) => escapeHtml(`Client ${i + 1} · ${o}`)).join('<br>');
  }
  return '—';
}

function filterHistoryRows(rows, query) {
  const q = (query || '').trim().toLowerCase();
  if (!q) return rows;
  return rows.filter(row => {
    const clientBits = (row.clients || []).flatMap(c => [
      String(c.clientNumber),
      c.organization,
      c.ticket,
    ]);
    const hay = [
      row.sessionId,
      row.clientsLabel,
      (row.ticketNumbers || []).join(' '),
      (row.organizations || []).join(' '),
      ...clientBits,
    ].filter(Boolean).join(' ').toLowerCase();
    return hay.includes(q);
  });
}

function renderHistorySortHeader(label, sortKey, mode, sortState) {
  const active = sortState.sortBy === sortKey;
  const cls = active ? 'sortable sortActive' : 'sortable';
  const ind = active ? `<span class="sortInd">${historySortIndicator(sortKey, sortState.sortOrder)}</span>` : '<span class="sortInd"></span>';
  return `<th class="${cls}" data-sort="${sortKey}" data-mode="${mode}">${label}${ind}</th>`;
}

function renderHistoryTable(rows, mode, sortState, searchQuery) {
  const hasSearch = Boolean((searchQuery || '').trim());
  if (!rows.length) {
    if (hasSearch) return '<p class="muted">No sessions match your search.</p>';
    return mode === 'archived'
      ? '<p class="muted">No archived sessions.</p>'
      : '<p class="muted">No saved sessions yet. History appears after you add a tenant, fetch a ticket, or authenticate — not on every page load.</p>';
  }

  const whenKey = mode === 'archived' ? 'archivedAt' : 'updatedAt';
  const whenLabel = mode === 'archived' ? 'Archived' : 'Updated';
  const html = [`<table class="historyTable"><thead><tr>
    ${renderHistorySortHeader(whenLabel, whenKey, mode, sortState)}
    ${renderHistorySortHeader('Session', 'sessionId', mode, sortState)}
    ${renderHistorySortHeader('Clients', 'clients', mode, sortState)}
    ${renderHistorySortHeader('Tickets', 'ticket', mode, sortState)}
    ${renderHistorySortHeader('Tenants', 'tenants', mode, sortState)}
    <th></th>
  </tr></thead><tbody>`];

  for (const row of rows) {
    const tickets = escapeHtml((row.ticketNumbers || []).join(', ') || '—');
    const clients = formatHistoryClients(row);
    const whenSource = mode === 'archived'
      ? (row.archivedAt || row.updatedAt || row.createdAt)
      : (row.updatedAt || row.createdAt);
    const when = formatHistoryWhen(whenSource);
    const sid = escapeHtml(row.sessionId || '');
    const actions = [`<button type="button" class="small restoreSession" data-session-id="${sid}">Restore</button>`];

    if (serverFeatures.sessionHistoryActions) {
      if (mode === 'saved') {
        actions.push(`<button type="button" class="small archiveSession" data-session-id="${sid}">Archive</button>`);
      } else {
        actions.push(`<button type="button" class="small unarchiveSession" data-session-id="${sid}">Unarchive</button>`);
      }
      actions.push(`<button type="button" class="small danger deleteSession" data-session-id="${sid}">Delete</button>`);
    }

    html.push(`<tr>
      <td>${when}</td>
      <td><code>${sid}</code></td>
      <td class="historyClients">${clients}</td>
      <td>${tickets}</td>
      <td>${row.tenantCount ?? 0}</td>
      <td class="historyActions">${actions.join('')}</td>
    </tr>`);
  }

  html.push('</tbody></table>');
  return html.join('');
}

function wireHistoryTableSort(listEl, mode) {
  if (!listEl) return;
  listEl.querySelectorAll('th.sortable').forEach(th => {
    th.addEventListener('click', () => {
      const sortKey = th.dataset.sort;
      if (!sortKey) return;
      const current = historySort[mode];
      if (current.sortBy === sortKey) {
        current.sortOrder = current.sortOrder === 'asc' ? 'desc' : 'asc';
      } else {
        current.sortBy = sortKey;
        const descByDefault = ['updatedAt', 'createdAt', 'archivedAt', 'tenants'].includes(sortKey);
        current.sortOrder = descByDefault ? 'desc' : 'asc';
      }
      localStorage.setItem(HISTORY_SORT_STORAGE[mode], `${current.sortBy}:${current.sortOrder}`);
      loadSessionHistory().catch(err => log(`History sort failed: ${err.message}`));
    });
  });
}

function renderHistorySection(mode) {
  const rows = mode === 'saved' ? savedHistoryRows : archivedHistoryRows;
  const listEl = document.getElementById(mode === 'saved' ? 'sessionHistoryList' : 'sessionArchiveList');
  const summaryEl = document.getElementById(mode === 'saved' ? 'savedHistorySummary' : 'sessionArchiveSummary');
  const filtered = filterHistoryRows(rows, historySearch[mode]);

  if (summaryEl) {
    const total = rows.length;
    const shown = filtered.length;
    if (!total) {
      summaryEl.textContent = mode === 'saved' ? '(none yet)' : '(empty)';
    } else if (historySearch[mode]?.trim() && shown !== total) {
      summaryEl.textContent = `(${shown} of ${total})`;
    } else {
      summaryEl.textContent = mode === 'saved' ? `(${total} saved)` : `(${total} archived)`;
    }
  }

  if (listEl) {
    listEl.innerHTML = renderHistoryTable(filtered, mode, historySort[mode], historySearch[mode]);
    wireHistoryTableActions(listEl);
    wireHistoryTableSort(listEl, mode);
  }
}

async function fetchSessionHistoryList(archived, sortBy, sortOrder) {
  const params = new URLSearchParams({
    archived: archived ? '1' : '0',
    sort: sortBy,
    order: sortOrder,
    limit: '100',
  });
  const data = await api(`/api/sessions/history?${params.toString()}`);
  return data.sessions || [];
}

function wireHistoryTableActions(listEl) {
  if (!listEl) return;
  listEl.querySelectorAll('.restoreSession').forEach(btn => {
    btn.addEventListener('click', () => restoreSessionFromHistory(btn.dataset.sessionId));
  });
  listEl.querySelectorAll('.archiveSession').forEach(btn => {
    btn.addEventListener('click', () => archiveSessionFromHistory(btn.dataset.sessionId));
  });
  listEl.querySelectorAll('.unarchiveSession').forEach(btn => {
    btn.addEventListener('click', () => unarchiveSessionFromHistory(btn.dataset.sessionId));
  });
  listEl.querySelectorAll('.deleteSession').forEach(btn => {
    btn.addEventListener('click', () => deleteSessionFromHistory(btn.dataset.sessionId));
  });
}

async function archiveSessionFromHistory(sessionId) {
  if (!sessionId) return;
  if (!window.confirm(`Archive session ${sessionId}?\n\nIt will move to the Archive list.`)) return;
  await api(`/api/sessions/history/${encodeURIComponent(sessionId)}/archive`, { method: 'POST' });
  log(`Archived session ${sessionId}.`);
  await loadSessionHistory();
}

async function unarchiveSessionFromHistory(sessionId) {
  if (!sessionId) return;
  if (!window.confirm(`Restore session ${sessionId} from archive to Saved sessions?`)) return;
  await api(`/api/sessions/history/${encodeURIComponent(sessionId)}/unarchive`, { method: 'POST' });
  log(`Unarchived session ${sessionId}.`);
  await loadSessionHistory();
}

async function deleteSessionFromHistory(sessionId) {
  if (!sessionId) return;
  if (!window.confirm(`Permanently delete session ${sessionId}?\n\nThis cannot be undone.`)) return;
  await api(`/api/sessions/history/${encodeURIComponent(sessionId)}`, { method: 'DELETE' });
  log(`Deleted session ${sessionId} from history.`);
  await loadSessionHistory();
}

async function loadSessionHistory() {
  const listEl = document.getElementById('sessionHistoryList');
  const archiveListEl = document.getElementById('sessionArchiveList');
  const summaryEl = document.getElementById('sessionHistorySummary');
  const savedSummaryEl = document.getElementById('savedHistorySummary');
  const archiveSummaryEl = document.getElementById('sessionArchiveSummary');

  if (!serverFeatures.sessionHistory) {
    const disabledMsg = '(restart web runner to enable)';
    if (summaryEl) summaryEl.textContent = disabledMsg;
    if (savedSummaryEl) savedSummaryEl.textContent = disabledMsg;
    if (archiveSummaryEl) archiveSummaryEl.textContent = '';
    if (listEl) listEl.innerHTML = '<p class="muted">Session history requires a web runner restart to load the latest server code.</p>';
    if (archiveListEl) archiveListEl.innerHTML = '';
    return;
  }

  try {
    const [savedRows, archivedRows] = await Promise.all([
      fetchSessionHistoryList(false, historySort.saved.sortBy, historySort.saved.sortOrder),
      fetchSessionHistoryList(true, historySort.archived.sortBy, historySort.archived.sortOrder),
    ]);

    savedHistoryRows = savedRows;
    archivedHistoryRows = archivedRows;

    if (summaryEl) {
      summaryEl.textContent = savedRows.length ? `(${savedRows.length} saved)` : '(none yet)';
    }

    renderHistorySection('saved');
    renderHistorySection('archived');
  } catch (e) {
    if (summaryEl) summaryEl.textContent = '(load failed)';
    if (savedSummaryEl) savedSummaryEl.textContent = '(load failed)';
    if (archiveSummaryEl) archiveSummaryEl.textContent = '';
    if (listEl) listEl.textContent = String(e.message || e);
    if (archiveListEl) archiveListEl.textContent = '';
    if (String(e.message || e).includes('Not found')) serverFeatures.sessionHistory = false;
  }
}

async function restoreSessionFromHistory(sessionId) {
  if (!sessionId) return;
  const ok = window.confirm(
    `Restore session ${sessionId}?\n\nThis replaces the current in-memory session with a new one using saved settings. Existing PowerShell worker windows are NOT closed automatically — close or remove them manually if needed.`
  );
  if (!ok) return;
  log(`Restoring session from ${sessionId}…`);
  const data = await api('/api/session/restore', {
    method: 'POST',
    body: JSON.stringify({ sessionId, force: true }),
  });
  tenantUiState.clear();
  clientBusy.clear();
  if (data.session) {
    currentSessionId = data.session.sessionId;
    applySessionSettingsToUi(data.session);
  }
  const snapshots = data.tenantSnapshots || [];
  const restoredTenantCount = data.session?.tenantCount || 0;
  if (restoredTenantCount > 0) {
    for (const snap of snapshots) {
      const ui = snap.uiState && typeof snap.uiState === 'object' ? { ...snap.uiState } : {};
      if (snap.exoOrganizationName && !ui.orgHint) ui.orgHint = snap.exoOrganizationName;
      if (snap.graphTenantName && !ui.orgHint) ui.orgHint = snap.graphTenantName;
      if (ui.organizationHint && !ui.orgHint) ui.orgHint = ui.organizationHint;
      if (snap.outputFolder) ui.priorOutputFolder = snap.outputFolder;
      const key = String(snap.clientNumber || snapshots.indexOf(snap) + 1);
      tenantUiState.set(key, { ...(tenantUiState.get(key) || {}), ...ui });
    }
    log(`Restored from ${data.restoredFrom || sessionId}. Reconnected ${restoredTenantCount} tenant(s) to saved workers where still running.`);
  } else {
    for (const snap of snapshots) {
    const t = await api('/api/tenants', { method: 'POST' });
    const ui = snap.uiState && typeof snap.uiState === 'object' ? { ...snap.uiState } : {};
    if (snap.exoOrganizationName && !ui.orgHint) ui.orgHint = snap.exoOrganizationName;
    if (snap.graphTenantName && !ui.orgHint) ui.orgHint = snap.graphTenantName;
    if (ui.organizationHint && !ui.orgHint) ui.orgHint = ui.organizationHint;
    if (snap.outputFolder) ui.priorOutputFolder = snap.outputFolder;
    tenantUiState.set(String(t.clientNumber), ui);
    await api(`/api/tenants/${t.clientNumber}/ui-state`, {
      method: 'POST',
      body: JSON.stringify({ ...ui, ticketContentLength: ui.ticketContentLength || 0 }),
    });
  }
    log(`Restored from ${data.restoredFrom || sessionId}. Added ${snapshots.length} tenant slot(s). Re-run Exchange/Graph auth and Fetch from Manage.`);
  }
  await refreshSession();
  await loadSessionHistory();
}


function pad2(n) {

  return String(n).padStart(2, '0');

}

function toDateTimeLocalValue(date) {

  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}T${pad2(date.getHours())}:${pad2(date.getMinutes())}`;

}

function defaultDateStartValue() {

  const days = Math.max(1, parseInt(document.getElementById('daysBack')?.value, 10) || 10);

  const d = new Date();

  d.setDate(d.getDate() - days);

  return toDateTimeLocalValue(d);

}

function defaultDateEndValue() {

  return toDateTimeLocalValue(new Date());

}

function parseSearchTerms(text) {

  return (text || '')

    .split(',')

    .map(s => s.trim())

    .filter(Boolean);

}

function saveTenantUiState(clientNumber, div) {

  if (!div) return;

  const useSessionReportDefaults = Boolean(div.querySelector('.useSessionReportDefaults')?.checked ?? true);
  const customPanel = div.querySelector('.tenantReportExportsCustom');
  const reportSelections = useSessionReportDefaults ? null : readReportSelectionsFromContainer(customPanel);

  const prior = tenantUiState.get(String(clientNumber)) || {};

  tenantUiState.set(String(clientNumber), {

    ticket: (div.querySelector('.ticketInput')?.value || '').trim(),

    ticketContent: (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || prior.ticketContent || '',

    organizationHint: prior.organizationHint || '',

    tenantId: div.querySelector('.appRegSelect')?.value || '',

    interactive: Boolean(div.querySelector('.interactiveCheck')?.checked),

    filterUsers: Boolean(div.querySelector('.filterUsersCheck')?.checked),

    userSearch: div.querySelector('.userSearchInput')?.value || '',

    validatedUsers: JSON.parse(div.dataset.validatedUsers || '[]'),

    dateStart: div.querySelector('.dateStart')?.value || '',

    dateEnd: div.querySelector('.dateEnd')?.value || '',

    useSessionReportDefaults,

    reportSelections,

    security: {
      resolveJson: div.dataset.securityResolveJson || '',
      s1Profile: div.querySelector('.s1ProfileSelect')?.value || 'connectwise',
      s1SiteId: div.querySelector('.s1SiteIdInput')?.value || '',
      includeLiongardContext: Boolean(div.querySelector('.includeLiongardContext')?.checked),
      huntressPulls: Object.fromEntries([...div.querySelectorAll('.huntressPull')].map((cb) => [cb.dataset.pull, cb.checked])),
      s1Pulls: Object.fromEntries([...div.querySelectorAll('.s1Pull')].map((cb) => [cb.dataset.pull, cb.checked])),
    },

  });

  scheduleTenantUiStateSync(clientNumber, div);

}

function restoreTenantUiState(clientNumber, div) {

  const saved = tenantUiState.get(String(clientNumber));

  if (!saved || !div) return;

  const ticket = div.querySelector('.ticketInput');

  const appReg = div.querySelector('.appRegSelect');

  const interactive = div.querySelector('.interactiveCheck');

  const preview = div.querySelector('.ticketPreview');

  const filterUsers = div.querySelector('.filterUsersCheck');

  const userSearch = div.querySelector('.userSearchInput');

  const validateBtn = div.querySelector('.validateUsers');

  const dateStart = div.querySelector('.dateStart');

  const dateEnd = div.querySelector('.dateEnd');

  if (ticket) ticket.value = saved.ticket;

  if (appReg && saved.tenantId) appReg.value = saved.tenantId;

  if (interactive) interactive.checked = saved.interactive;

  if (filterUsers) filterUsers.checked = saved.filterUsers;

  if (userSearch) userSearch.value = saved.userSearch;

  if (dateStart && saved.dateStart) dateStart.value = saved.dateStart;

  if (dateEnd && saved.dateEnd) dateEnd.value = saved.dateEnd;

  restoreSecurityUiState(div, saved);

  if (saved.validatedUsers?.length) {

    div.dataset.validatedUsers = JSON.stringify(saved.validatedUsers);

    updateValidatedUsersDisplay(div, saved.validatedUsers);

  }

  if (filterUsers) {

    const enabled = filterUsers.checked;

    if (userSearch) userSearch.disabled = !enabled;

    if (validateBtn) validateBtn.disabled = !enabled;

  }

  if (saved.ticketContent) {

    div.dataset.ticketContent = saved.ticketContent;

    const paste = div.querySelector('.ticketPaste');
    if (paste && !paste.value) paste.value = saved.ticketContent;

    if (preview) {

      preview.textContent = `Manage ticket loaded (${saved.ticketContent.length} chars)`;

      preview.style.display = 'block';

    }

  }

  const useDefaultsCheck = div.querySelector('.useSessionReportDefaults');
  const customPanel = div.querySelector('.tenantReportExportsCustom');
  const useDefaults = saved.useSessionReportDefaults !== false;
  if (useDefaultsCheck) useDefaultsCheck.checked = useDefaults;
  if (customPanel && saved.reportSelections) {
    applyReportSelectionsToContainer(customPanel, saved.reportSelections);
  }
  updateTenantReportExportsHint(div, useDefaults);
  if (customPanel) customPanel.style.display = useDefaults ? 'none' : 'block';

}

function updateValidatedUsersDisplay(div, users) {

  const status = div.querySelector('.userValidationStatus');

  const list = div.querySelector('.validatedUsersList');

  if (!users?.length) {

    if (status) status.textContent = '';

    if (list) { list.style.display = 'none'; list.textContent = ''; }

    return;

  }

  if (status) status.textContent = `${users.length} user(s) validated`;

  if (list) {

    list.textContent = users.join('\n');

    list.style.display = 'block';

  }

}

function normalizeResponse(value) {

  return (value == null ? '' : String(value)).trim();

}

async function withClientLock(clientNumber, fn) {

  const key = String(clientNumber);

  if (clientBusy.has(key)) {

    log(`Client ${clientNumber}: busy — wait for the current operation to finish.`);

    return;

  }

  clientBusy.add(key);

  try {

    await fn();

  } finally {

    clientBusy.delete(key);

  }

}

function setTenantButtonsDisabled(div, disabled) {

  if (!div) return;

  div.querySelectorAll('button.graphAuth, button.exoAuth, button.generateReports, button.validateUsers').forEach(btn => {

    btn.disabled = disabled;

  });

}

function defaultSessionBody() {

  return {

    investigatorName: '',

    companyName: '',

    daysBack: 7,

    reportSelections: defaultReportSelections(),

  };

}

function sessionBodyFromUi() {

  const body = defaultSessionBody();

  body.daysBack = parseInt(document.getElementById('daysBack')?.value, 10) || 7;

  body.reportSelections = readReportSelectionsFromContainer(document.getElementById('sessionReportExportsBody'));

  return body;

}

function applySessionSettingsToUi(session) {

  initSessionReportExportsPanel();

  if (session?.reportSelections) {

    applyReportSelectionsToContainer(document.getElementById('sessionReportExportsBody'), session.reportSelections);

    updateSessionReportExportsSummary(session.reportSelections);

    const rs = session.reportSelections;

    const mt = document.getElementById('messageTraceDays');

    const si = document.getElementById('signInLogsDays');

    const db = document.getElementById('daysBack');

    if (mt && rs.MessageTraceDaysBack != null) mt.value = String(rs.MessageTraceDaysBack);

    if (si && rs.SignInLogsDaysBack != null) si.value = String(rs.SignInLogsDaysBack);

    if (db && session.daysBack != null) db.value = String(session.daysBack);

  }

}

function hasActiveSession(session) {

  return Boolean(session?.sessionId);

}

async function ensureSession(options = {}) {

  let session = await api('/api/session');

  if (hasActiveSession(session)) {

    applySessionSettingsToUi(session);

    return session;

  }

  session = await api('/api/session', {

    method: 'POST',

    body: JSON.stringify(sessionBodyFromUi()),

  });

  applySessionSettingsToUi(session);

  if (!options.quiet) {

    log('Session ready.');

  }

  return session;

}

function log(msg) {

  const line = `[${new Date().toLocaleTimeString()}] ${msg}\n`;

  if (logEl) {
    logEl.textContent += line;
    logEl.scrollTop = logEl.scrollHeight;
  }

}

function getWorkerLogPanel(clientNumber) {
  const id = `logPanelClient${clientNumber}`;
  let panel = document.getElementById(id);
  if (!panel) {
    panel = document.createElement('div');
    panel.id = id;
    panel.className = 'logPanel';
    panel.innerHTML = `<pre class="workerLogPre"></pre>`;
    document.querySelector('.card:has(#logTabs)')?.appendChild(panel);
  }
  return panel.querySelector('.workerLogPre');
}

function syncLogTabs(tenantNumbers) {
  const tabsEl = document.getElementById('logTabs');
  if (!tabsEl) return;
  const nums = new Set((tenantNumbers || []).map(n => String(n)));
  tabsEl.querySelectorAll('[data-log-tab]').forEach(btn => {
    const tab = btn.dataset.logTab;
    if (tab !== 'activity' && !nums.has(tab.replace('client', ''))) {
      btn.remove();
      const panel = document.getElementById(`logPanelClient${tab.replace('client', '')}`);
      panel?.remove();
      workerLogOffsets.delete(tab.replace('client', ''));
      stopWorkerLogPoll(tab.replace('client', ''));
    }
  });
  for (const n of nums) {
    if (!tabsEl.querySelector(`[data-log-tab="client${n}"]`)) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.dataset.logTab = `client${n}`;
      btn.textContent = `Client ${n}`;
      btn.addEventListener('click', () => switchLogTab(`client${n}`));
      tabsEl.appendChild(btn);
    }
  }
}

function switchLogTab(tabId) {
  activeLogTab = tabId;
  document.querySelectorAll('#logTabs [data-log-tab]').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.logTab === tabId);
  });
  document.querySelectorAll('.logPanel').forEach(p => p.classList.remove('active'));
  if (tabId === 'activity') {
    document.getElementById('logPanelActivity')?.classList.add('active');
  } else {
    document.getElementById(`logPanel${tabId.charAt(0).toUpperCase()}${tabId.slice(1)}`)?.classList.add('active');
  }
}

function startWorkerLogPoll(clientNumber) {
  const key = String(clientNumber);
  if (workerLogPollers.has(key)) return;
  const timer = setInterval(() => pollWorkerLog(clientNumber).catch(() => {}), 2500);
  workerLogPollers.set(key, timer);
  void pollWorkerLog(clientNumber);
}

function stopWorkerLogPoll(clientNumber) {
  const key = String(clientNumber);
  if (workerLogPollers.has(key)) {
    clearInterval(workerLogPollers.get(key));
    workerLogPollers.delete(key);
  }
}

async function pollWorkerLog(clientNumber) {
  const key = String(clientNumber);
  const offset = workerLogOffsets.get(key) || 0;
  const data = await api(`/api/tenants/${clientNumber}/status?tailLines=300&sinceOffset=${offset}`);
  const pre = getWorkerLogPanel(clientNumber);
  if (data.offset != null) workerLogOffsets.set(key, data.offset);
  const chunk = data.status || '';
  if (chunk && pre) {
    pre.textContent += (pre.textContent && !pre.textContent.endsWith('\n') ? '\n' : '') + chunk;
    pre.scrollTop = pre.scrollHeight;
  }
}

function focusClientLogTab(clientNumber) {
  syncLogTabs([clientNumber]);
  switchLogTab(`client${clientNumber}`);
  startWorkerLogPoll(clientNumber);
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

async function detectServerFeatures() {

  try {

    const health = await api('/api/health', {}, 5000);

    if (health?.features) serverFeatures = { ...serverFeatures, ...health.features };

    else if (health?.version === '0.3.0' || health?.features?.exportPresets) {
      serverFeatures = {
        sessionHistory: true,
        sessionHistoryActions: true,
        reportSelections: true,
        exportPresets: true,
        noWaitCommands: true,
        hiddenWorkers: health.hiddenWorkers !== false,
        wcmManagement: true,
        workerLogTabs: true,
      };
    } else if (health?.version === '0.2.1') serverFeatures = { sessionHistory: true, sessionHistoryActions: true, reportSelections: true, noWaitCommands: true };

    else if (health?.version === '0.2.0') serverFeatures = { sessionHistory: true, sessionHistoryActions: false, reportSelections: true, noWaitCommands: true };

  } catch {

    /* keep defaults */

  }

}

function authInfoHtml(t) {

  const parts = [];

  if (t.exoOrganizationName || t.exoTenantId) {

    parts.push(`EXO: ${t.exoOrganizationName || t.exoTenantId}`);

  }

  if (t.graphTenantName || t.graphTenantId) {

    parts.push(`Graph: ${t.graphTenantName || t.graphTenantId}`);

  }

  if (t.exoTenantId && t.graphTenantId &&

      t.exoTenantId.toLowerCase() !== t.graphTenantId.toLowerCase()) {

    return '<div class="outputPath" style="color:#cf222e">Tenant mismatch — pick the matching App reg tenant and re-run Graph Auth.</div>';

  }

  return parts.length ? `<div class="tenantAuthInfo muted">${parts.join(' · ')}</div>` : '';

}

function autoSelectAppRegForExo(div, exoTenantId) {

  if (!exoTenantId || !div) return;

  const sel = div.querySelector('.appRegSelect');

  if (!sel || sel.value) return;

  const match = appRegistrations.find(a => (a.tenantId || '').toLowerCase() === exoTenantId.toLowerCase());

  if (match) sel.value = match.tenantId;

}

function renderTenants(session) {

  tenantsEl.innerHTML = '';

  if (!session?.tenants?.length) {

    tenantsEl.innerHTML = '<p class="muted">No tenants yet.</p>';

    return;

  }

  const autoWcmLabel = appRegistrations.length > 1

    ? '(auto WCM — requires Exchange Auth first)'

    : '(auto WCM)';

  for (const t of session.tenants) {

    const details = document.createElement('details');

    details.className = 'tenant collapsible';

    details.dataset.client = t.clientNumber;

    if (!isTenantCollapsed(t.clientNumber)) details.open = true;

    const div = document.createElement('div');

    div.className = 'tenantBody';

    details.dataset.client = t.clientNumber;

    div.dataset.graphAuthenticated = t.graphAuthenticated ? '1' : '0';
    div.dataset.exchangeAuthenticated = t.exchangeAuthenticated ? '1' : '0';
    div.dataset.workerAlive = t.workerAlive === false ? '0' : '1';
    if (t.outputFolder) div.dataset.outputFolder = t.outputFolder;
    if (t.lastResponse) div.dataset.lastResponse = normalizeResponse(t.lastResponse);
    const appRegOptions = appRegistrations.map(a =>

      `<option value="${a.tenantId || ''}">${a.displayText}</option>`

    ).join('');

    const canGenerate = tenantHasRequiredAuth(t, session) && !t.reportInProgress;

    const outputHtml = t.outputFolder

      ? `<div class="outputPath muted">Reports: ${t.outputFolder}</div>`

      : '';

    const lastResp = normalizeResponse(t.lastResponse);

    const errorHtml = (lastResp.includes('_FAILED') || lastResp.includes('ERROR'))

      ? `<div class="outputPath" style="color:#cf222e">Last: ${lastResp.substring(0, 220)}${lastResp.length > 220 ? '…' : ''}</div>`

      : '';

    const inProgressGraph = lastResp === 'GRAPH_AUTH_STARTED';

    const inProgressExo = lastResp === 'EXCHANGE_AUTH_STARTED';

    if (t.uiState && typeof t.uiState === 'object') {
      const existing = tenantUiState.get(String(t.clientNumber)) || {};
      const merged = { ...existing, ...t.uiState };
      const localValidated = existing.validatedUsers;
      const serverValidated = t.uiState.validatedUsers;
      if (Array.isArray(localValidated) && localValidated.length
          && (!Array.isArray(serverValidated) || !serverValidated.length)) {
        merged.validatedUsers = localValidated;
      }
      tenantUiState.set(String(t.clientNumber), merged);
    }

    const summary = document.createElement('summary');

    summary.className = 'tenantSummary';

    summary.innerHTML = `

      <span class="tenantSummaryTitle">${escapeHtml(tenantSummaryTitle(t, null))}</span>

      <span class="tenantSummaryStatus"></span>

      <span class="tenantSummaryActions"><button type="button" class="removeTenant danger small">Remove</button></span>

    `;

    div.innerHTML = `

      <span class="muted">Graph: ${t.graphAuthenticated ? 'yes' : 'no'}, EXO: ${t.exchangeAuthenticated ? 'yes' : 'no'}${t.workerAlive === false ? ', worker stopped' : ''}${t.reportInProgress ? ', generating…' : ''}${inProgressGraph ? ', graph auth…' : ''}${inProgressExo ? ', exo auth…' : ''}</span>

      ${authInfoHtml(t)}

      <div class="row">

        <label>App reg tenant

          <select class="appRegSelect"><option value="">${autoWcmLabel}</option>${appRegOptions}</select>

        </label>

        <label><input type="checkbox" class="interactiveCheck" /> Force interactive Graph</label>

      </div>

      <div class="sectionLabel">Timeframe</div>

      <div class="row">

        <label>Start <input type="datetime-local" class="dateStart" value="${defaultDateStartValue()}" /></label>

        <label>End <input type="datetime-local" class="dateEnd" value="${defaultDateEndValue()}" /></label>

      </div>

      <details class="tenantReportExports collapsible">

        <summary>Report exports <span class="reportExportsHint muted">(session defaults)</span></summary>

        <div class="collapsible-body">

          <label><input type="checkbox" class="useSessionReportDefaults" checked /> Use session defaults</label>

          <div class="tenantReportExportsCustom" style="display:none;margin-top:0.5rem">

            <p class="muted" style="margin:0.35rem 0">Override exports for this client only.</p>

            ${buildReportExportsPanelHtml('tenant')}

          </div>

        </div>

      </details>

      <div class="sectionLabel">Users (optional)</div>

      <div class="row">

        <label><input type="checkbox" class="filterUsersCheck" /> Filter by users</label>

      </div>

      <div class="row">

        <input type="text" class="userSearchInput" placeholder="name, upn@domain.com — comma-separated" disabled />

        <button type="button" class="validateUsers" disabled>Validate users</button>

        <span class="userValidationStatus muted"></span>

      </div>

      <div class="validatedUsersList muted" style="display:none"></div>

      <div class="row">

        <label>Ticket # (optional)

          <input type="text" class="ticketInput" placeholder="1839334" />

        </label>

        <button type="button" class="fetchTicket">Fetch from Manage</button>

        <button type="button" class="extractEmails">Extract emails from ticket</button>

      </div>

      <div class="ticketPreview muted" style="display:none;font-size:0.85rem;margin:0.35rem 0;"></div>

      <label class="muted" style="font-size:0.85rem;display:block;margin:0.35rem 0 0.15rem">Or paste ticket text (when Manage unavailable)</label>
      <textarea class="ticketPaste" rows="3" placeholder="Paste ticket body…" style="width:100%;font:inherit;margin-bottom:0.35rem"></textarea>

      ${buildSecurityIntegrationsPanelHtml()}

      <div>

        <button class="exoAuth">Exchange Auth</button>

        <button class="graphAuth">Graph Auth</button>

        <button class="generateReports primary" ${canGenerate ? '' : 'disabled'}>${t.reportInProgress ? 'Generating…' : 'Generate Reports'}</button>

        <button class="openReports success" ${t.outputFolder ? '' : 'disabled'}>Open Reports</button>

        <button class="restartWorker">Restart worker</button>

        <button class="showConsoleRestart">Show console</button>

        <button class="statusBtn">Tail status</button>

        <button class="openStatusFile">Open status file</button>

        <button class="graphDisconnect">Log out Graph</button>

        <button class="resetAuth">Reset auth</button>

        <button class="analyzeReports" ${t.outputFolder ? '' : 'disabled'}>Analyze reports</button>

      </div>

      ${errorHtml}

      ${outputHtml}

    `;

    div.querySelector('.filterUsersCheck').addEventListener('change', (e) => {

      const enabled = e.target.checked;

      div.querySelector('.userSearchInput').disabled = !enabled;

      div.querySelector('.validateUsers').disabled = !enabled || div.dataset.graphAuthenticated !== '1';

      if (enabled && div.dataset.graphAuthenticated !== '1') {
        const msg = `Client ${t.clientNumber}: Filter by users requires Graph Auth. Complete Graph Auth, then Validate users, before Generate — otherwise user-scoped exports (UAL, message trace, etc.) run for the whole tenant.`;
        log(msg);
        window.alert(msg);
      }

      if (!enabled) {

        div.dataset.validatedUsers = '[]';

        updateValidatedUsersDisplay(div, []);

      }

      saveTenantUiState(t.clientNumber, div);

    });

    div.querySelector('.graphAuth').addEventListener('click', () => graphAuth(t.clientNumber, div));

    div.querySelector('.exoAuth').addEventListener('click', () => exoAuth(t.clientNumber, div));

    div.querySelector('.validateUsers').addEventListener('click', () => validateUsers(t.clientNumber, div));

    div.querySelector('.fetchTicket')?.addEventListener('click', () => fetchManageTicket(t.clientNumber, div));

    div.querySelector('.ticketPaste')?.addEventListener('input', () => {
      div.dataset.ticketContent = div.querySelector('.ticketPaste')?.value || '';
      saveTenantUiState(t.clientNumber, div);
    });

    div.querySelector('.generateReports').addEventListener('click', () => generateReports(t.clientNumber, div));

    div.querySelector('.openReports').addEventListener('click', () => openReports(t.outputFolder));

    div.querySelector('.restartWorker').addEventListener('click', () => restartWorker(t.clientNumber, div));

    div.querySelector('.showConsoleRestart')?.addEventListener('click', () => restartWorker(t.clientNumber, div, { showConsole: true }));

    div.querySelector('.statusBtn').addEventListener('click', () => tailStatus(t.clientNumber, div));

    div.querySelector('.openStatusFile')?.addEventListener('click', () => openTenantStatusFile(t.clientNumber));

    div.querySelector('.graphDisconnect')?.addEventListener('click', () => graphDisconnect(t.clientNumber, div));

    div.querySelector('.resetAuth')?.addEventListener('click', () => resetAuth(t.clientNumber, div));

    div.querySelectorAll('.extractEmails').forEach(btn => btn.addEventListener('click', () => extractEmailsFromTicket(t.clientNumber, div)));

    div.querySelector('.analyzeReports')?.addEventListener('click', () => analyzeTenantReports(t.clientNumber, div, t.outputFolder));

    restoreTenantUiState(t.clientNumber, div);
    wireTenantReportExportsPanel(div, t.clientNumber);
    wireSecurityIntegrationsPanel(div, t.clientNumber);
    applyTenantSummary(details, t, div);
    details.addEventListener('toggle', () => setTenantCollapsed(t.clientNumber, !details.open));
    const removeBtn = summary.querySelector('.removeTenant');
    removeBtn?.addEventListener('mousedown', (e) => {
      e.preventDefault();
      e.stopPropagation();
    });
    removeBtn?.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      removeTenant(t.clientNumber, div);
    });
    autoSelectAppRegForExo(div, t.exoTenantId);
    details.appendChild(summary);
    details.appendChild(div);
    tenantsEl.appendChild(details);

  }

}

async function refreshSession() {

  const session = await api('/api/session');

  if (hasActiveSession(session)) {

    applySessionSettingsToUi(session);

    const count = session.tenantCount ?? session.tenants?.length ?? 0;

    currentSessionId = session.sessionId;

    sessionInfo.textContent = `Session ${session.sessionId} (${count} tenant(s))`;

    renderTenants(session);

    if (session?.tenants?.length) {
      syncLogTabs(session.tenants.map(t => t.clientNumber));
      for (const t of session.tenants) {
        const busy = t.reportInProgress || clientBusy.has(String(t.clientNumber));
        if (busy) startWorkerLogPoll(t.clientNumber);
        else stopWorkerLogPoll(t.clientNumber);
      }
    } else {
      syncLogTabs([]);
    }

  } else {

    sessionInfo.textContent = 'No session';

    renderTenants(null);

  }

  return session;

}

async function pollWorkerResponse(clientNumber, startedToken, successPrefixes, failPrefix, waitSeconds = 300, progressLabel = 'working', onProgress = null) {

  const deadline = Date.now() + waitSeconds * 1000;

  let lastLogAt = 0;
  let sawStarted = !startedToken;

  while (Date.now() < deadline) {

    await new Promise(r => setTimeout(r, 2000));

    const data = await api(`/api/tenants/${clientNumber}/response`);

    const resp = normalizeResponse(data?.response);

    const elapsed = Math.round((Date.now() - (deadline - waitSeconds * 1000)) / 1000);

    if (elapsed - lastLogAt >= 15) {

      log(`Client ${clientNumber}: ${progressLabel}… (${elapsed}s)`);

      lastLogAt = elapsed;

      if (onProgress) {

        try { await onProgress(elapsed); } catch { /* ignore */ }

      }

    }

    if (!resp) {

      continue;

    }

    if (startedToken && (resp === startedToken || resp.startsWith(startedToken))) {

      sawStarted = true;
      continue;

    }

    if (!sawStarted) {

      const isTerminal = (failPrefix && resp.startsWith(failPrefix))
        || successPrefixes.some((prefix) => resp.startsWith(prefix));
      if (!isTerminal) {
        continue;
      }

    }

    if (failPrefix && resp.startsWith(failPrefix)) {

      throw new Error(resp);

    }

    for (const prefix of successPrefixes) {

      if (resp.startsWith(prefix)) {

        return resp;

      }

    }

  }

  throw new Error(`${progressLabel} timed out`);

}

async function pollAuth(clientNumber, startedToken, successPrefix, failPrefix, waitSeconds = 300) {

  return pollWorkerResponse(clientNumber, startedToken, [successPrefix], failPrefix, waitSeconds, 'waiting for auth');

}

async function waitForWorkerReady(clientNumber, timeoutSec = 45) {

  const deadline = Date.now() + timeoutSec * 1000;

  while (Date.now() < deadline) {

    const data = await api(`/api/tenants/${clientNumber}/status?tailLines=25`);

    const status = data?.status || '';

    const lines = status.split(/\r?\n/).filter(Boolean);

    const tail = lines.slice(-6).join('\n');

    if (/ready to receive commands/i.test(tail) && /Ready! Waiting for/i.test(tail)) {

      return;

    }

    await new Promise(r => setTimeout(r, 800));

  }

  throw new Error('Worker did not become ready in time');

}

async function ensureVisibleWorkerForAuth(clientNumber, div) {

  if (!serverFeatures.hiddenWorkers) return div;

  const uiKey = String(clientNumber);

  const ui = tenantUiState.get(uiKey) || {};

  const session = await api('/api/session');

  const tenant = session.tenants?.find(x => Number(x.clientNumber) === Number(clientNumber));

  if (ui.workerVisibleConsole) {

    try {

      if (tenant?.processId) {

        await waitForWorkerReady(clientNumber, 12);

        return getTenantBodyEl(clientNumber) || div;

      }

    } catch { /* fall through to restart */ }

  }

  if (tenant?.exchangeAuthenticated && tenant?.processId) {

    return getTenantBodyEl(clientNumber) || div;

  }

  if (tenant?.lastResponse?.includes('_AUTH_STARTED')) {

    log(`Client ${clientNumber}: prior auth looks stuck — restarting with visible console…`);

  } else {

    log(`Client ${clientNumber}: opening visible PowerShell window for sign-in…`);

  }

  if (div) {

    tenantUiState.set(uiKey, { ...ui, workerVisibleConsole: true });

    saveTenantUiState(clientNumber, div);

  }

  await restartWorker(clientNumber, div, { showConsole: true, skipReauthHint: true, skipRefresh: true });

  log(`Client ${clientNumber}: waiting for worker to start…`);

  await waitForWorkerReady(clientNumber);

  return getTenantBodyEl(clientNumber) || div;

}

async function graphAuth(clientNumber, div) {

  await withClientLock(clientNumber, async () => {

    saveTenantUiState(clientNumber, div);

    setTenantButtonsDisabled(div, true);

    try {

      if (!await requireLiveWorker(clientNumber, { actionLabel: 'Graph Auth' })) {
        return;
      }

      let tenantId = null;
      try {
        const session = await api('/api/session');
        const tenant = session.tenants?.find(x => Number(x.clientNumber) === Number(clientNumber));
        if (tenant?.exoTenantId) {
          tenantId = tenant.exoTenantId;
          autoSelectAppRegForExo(div, tenant.exoTenantId);
          log(`Client ${clientNumber}: auto WCM using EXO tenant ${tenantId}.`);
        }
      } catch { /* ignore */ }
      if (!tenantId) {
        tenantId = div?.querySelector('.appRegSelect')?.value || null;
      }
      const ui = tenantUiState.get(String(clientNumber)) || {};
      const orgHint = (ui.organizationHint || '').trim();
      if (tenantId && orgHint) {
        const selected = appRegistrations.find(a => (a.tenantId || '').toLowerCase() === tenantId.toLowerCase());
        const hintKey = orgHint.split(',')[0].trim().toLowerCase();
        if (selected && hintKey && !selected.displayText.toLowerCase().includes(hintKey)) {
          log(`Client ${clientNumber}: App reg "${selected.displayText}" may not match ticket org "${orgHint}" — pick the correct tenant in App reg before Graph Auth.`);
        }
      }
      const interactive = div?.querySelector('.interactiveCheck')?.checked;

      let cmd = 'GRAPH_AUTH';

      if (tenantId) cmd += `|TENANT_ID:${tenantId}`;

      if (interactive) cmd += '|INTERACTIVE:1';

      log(`Client ${clientNumber}: Graph auth…`);
      focusClientLogTab(clientNumber);
      div.dataset.lastResponse = 'GRAPH_AUTH_STARTED';
      refreshTenantSummaryUI(clientNumber, { lastResponse: 'GRAPH_AUTH_STARTED' });

      const initial = await api(`/api/tenants/${clientNumber}/command`, {

        method: 'POST',

        body: workerCommandBody(cmd),

      });

      let final = normalizeResponse(initial.response);

      if (!final || final === 'GRAPH_AUTH_STARTED' || final.startsWith('GRAPH_AUTH_SUCCESS') || final.startsWith('GRAPH_AUTH_FAILED')) {

        if (!final) {

          log(`Client ${clientNumber}: waiting for worker to acknowledge auth command…`);

        } else if (final === 'GRAPH_AUTH_STARTED') {

          log(`Client ${clientNumber}: complete sign-in in the browser popup on this PC…`);

        }

        final = await pollAuth(clientNumber, 'GRAPH_AUTH_STARTED', 'GRAPH_AUTH_SUCCESS', 'GRAPH_AUTH_FAILED', 300);

      }

      log(`Client ${clientNumber}: ${final || '(no response)'}`);

    } catch (e) {

      log(`Client ${clientNumber}: Error: ${e.message}`);

    } finally {

      setTenantButtonsDisabled(div, false);

      saveTenantUiState(clientNumber, div);

      await refreshSession();

    }

  });

}

async function exoAuth(clientNumber, div) {

  await withClientLock(clientNumber, async () => {

    saveTenantUiState(clientNumber, div);

    setTenantButtonsDisabled(div, true);

    try {

      if (!await requireLiveWorker(clientNumber, { actionLabel: 'Exchange Auth' })) {
        return;
      }

      div = await ensureVisibleWorkerForAuth(clientNumber, div) || div;

      log(`Client ${clientNumber}: Exchange auth…`);
      focusClientLogTab(clientNumber);
      div.dataset.lastResponse = 'EXCHANGE_AUTH_STARTED';
      refreshTenantSummaryUI(clientNumber, { lastResponse: 'EXCHANGE_AUTH_STARTED' });

      const initial = await api(`/api/tenants/${clientNumber}/command`, {

        method: 'POST',

        body: workerCommandBody('EXCHANGE_AUTH'),

      });

      let final = normalizeResponse(initial.response);

      if (!final || final === 'EXCHANGE_AUTH_STARTED') {

        if (!final) {

          log(`Client ${clientNumber}: waiting for worker to acknowledge auth command…`);

        } else {

          log(`Client ${clientNumber}: complete Exchange sign-in in the popup on this PC…`);

        }

        final = await pollAuth(clientNumber, 'EXCHANGE_AUTH_STARTED', 'EXCHANGE_AUTH_SUCCESS', 'EXCHANGE_AUTH_FAILED', 300);

      }

      log(`Client ${clientNumber}: ${final || '(no response)'}`);

      if (final.startsWith('EXCHANGE_AUTH_SUCCESS') && !final.includes('TENANT_ID:')) {
        log(`Client ${clientNumber}: EXO tenant ID not detected — select App reg tenant before Graph Auth if auto WCM fails.`);
      }
    } catch (e) {

      log(`Client ${clientNumber}: Error: ${e.message}`);

    } finally {

      setTenantButtonsDisabled(div, false);

      saveTenantUiState(clientNumber, div);

      await refreshSession();

    }

  });

}

async function validateUsers(clientNumber, div) {

  await withClientLock(clientNumber, async () => {

    saveTenantUiState(clientNumber, div);

    if (!await requireLiveWorker(clientNumber, { actionLabel: 'validate users' })) {
      return;
    }

    const session = await api('/api/session');
    const tenant = session.tenants?.find(x => Number(x.clientNumber) === Number(clientNumber));
    if (!tenant?.graphAuthenticated) {
      const msg = `Client ${clientNumber}: Graph Auth is required to validate users. Complete Graph Auth first, then Validate users.`;
      log(msg);
      window.alert(msg);
      return;
    }
    const terms = parseSearchTerms(div.querySelector('.userSearchInput')?.value);

    if (!terms.length) {

      log(`Client ${clientNumber}: enter one or more search terms (comma-separated).`);

      return;

    }

    const btn = div.querySelector('.validateUsers');

    if (btn) btn.disabled = true;

    try {

      const cmd = `VALIDATE_USERS|SEARCH_TERMS:${JSON.stringify(terms)}`;

      log(`Client ${clientNumber}: validating ${terms.length} search term(s)…`);

      const initial = await api(`/api/tenants/${clientNumber}/command`, {

        method: 'POST',

        body: workerCommandBody(cmd),

      });

      let final = normalizeResponse(initial.response);

      if (!final || final === 'VALIDATE_USERS_STARTED') {

        final = await pollWorkerResponse(

          clientNumber,

          'VALIDATE_USERS_STARTED',

          ['VALIDATE_USERS_SUCCESS:'],

          'VALIDATE_USERS_FAILED:',

          300,

          'validating users'

        );

      }

      if (final.startsWith('VALIDATE_USERS_SUCCESS:')) {

        const json = final.replace('VALIDATE_USERS_SUCCESS:', '');

        const result = JSON.parse(json);

        const users = Array.isArray(result.Users) ? result.Users : [];

        div.dataset.validatedUsers = JSON.stringify(users);

        updateValidatedUsersDisplay(div, users);

        saveTenantUiState(clientNumber, div);
        await syncTenantUiStateToServer(clientNumber, div);

        if (users.length) {

          log(`Client ${clientNumber}: validated ${users.length} user(s).`);

        } else {

          log(`Client ${clientNumber}: no users matched search terms.`);

        }

      } else {

        log(`Client ${clientNumber}: ${final}`);

      }

    } catch (e) {

      log(`Client ${clientNumber}: Validate failed: ${e.message}`);

    } finally {

      if (btn) btn.disabled = !div.querySelector('.filterUsersCheck')?.checked;

      await refreshSession();

    }

  });

}

async function tailStatus(clientNumber, div) {
  focusClientLogTab(clientNumber);
  await pollWorkerLog(clientNumber);
}

async function openTenantStatusFile(clientNumber) {
  const data = await api(`/api/tenants/${clientNumber}/status?tailLines=1`);
  if (data.statusFile) {
    const dir = String(data.statusFile).replace(/[/\\][^/\\]+$/, '');
    await api('/api/open-folder', { method: 'POST', body: JSON.stringify({ path: dir }) });
    log(`Opened session folder for Client ${clientNumber} status file.`);
  }
}

async function graphDisconnect(clientNumber, div) {
  await withClientLock(clientNumber, async () => {
    focusClientLogTab(clientNumber);
    log(`Client ${clientNumber}: logging out Graph…`);
    await api(`/api/tenants/${clientNumber}/command`, { method: 'POST', body: workerCommandBody('GRAPH_DISCONNECT') });
    await refreshSession();
  });
}

async function resetAuth(clientNumber, div) {
  await withClientLock(clientNumber, async () => {
    focusClientLogTab(clientNumber);
    log(`Client ${clientNumber}: resetting auth state…`);
    await api(`/api/tenants/${clientNumber}/command`, { method: 'POST', body: workerCommandBody('CANCEL_AUTH') });
    try {
      await pollWorkerResponse(clientNumber, null, ['CANCEL_AUTH_SUCCESS'], null, 60, 'resetting auth');
    } catch {
      log(`Client ${clientNumber}: reset auth did not confirm within 60s — check worker log.`);
    }
    refreshTenantSummaryUI(clientNumber, {
      graphAuthenticated: false,
      exchangeAuthenticated: false,
      lastResponse: 'CANCEL_AUTH_SUCCESS',
    });
    await refreshSession();
    log(`Client ${clientNumber}: auth reset. Re-run Exchange Auth, then Graph Auth.`);
  });
}

async function extractEmailsFromTicket(clientNumber, div) {
  const content = (div.querySelector('.ticketPaste')?.value || '').trim()
    || div.dataset.ticketContent
    || tenantUiState.get(String(clientNumber))?.ticketContent
    || '';
  if (!content.trim()) {
    log(`Client ${clientNumber}: fetch or paste ticket content first.`);
    return;
  }
  try {
    const data = await api('/api/ticket/extract-emails', { method: 'POST', body: JSON.stringify({ ticketContent: content }) });
    const emails = data.emails || [];
    if (!emails.length) {
      log(`Client ${clientNumber}: no emails found in ticket.`);
      return;
    }
    const filter = div.querySelector('.filterUsersCheck');
    const search = div.querySelector('.userSearchInput');
    if (filter) filter.checked = true;
    if (search) {
      search.disabled = false;
      search.value = emails.join(', ');
    }
    div.querySelector('.validateUsers').disabled = false;
    saveTenantUiState(clientNumber, div);
    log(`Client ${clientNumber}: extracted ${emails.length} email(s) into user search.`);
  } catch (e) {
    log(`Client ${clientNumber}: Extract emails failed: ${e.message}`);
  }
}

async function analyzeTenantReports(clientNumber, div, outputFolder) {
  const folder = outputFolder || div.dataset.outputFolder;
  if (!folder) {
    log(`Client ${clientNumber}: no report folder yet.`);
    return;
  }
  try {
    log(`Client ${clientNumber}: analyzing reports in ${folder}…`);
    const data = await api('/api/analyze-reports', { method: 'POST', body: JSON.stringify({ path: folder }) });
    log(`Client ${clientNumber}: analysis complete.`);
    if (data.result?.Summary) log(String(data.result.Summary));
  } catch (e) {
    log(`Client ${clientNumber}: Analyze failed: ${e.message}`);
  }
}

function appendTicketAndDateRange(cmd, ticketNumber, ticketContent, dateStart, dateEnd) {

  let out = cmd;

  const ticket = (ticketNumber || '').trim();

  const content = ticketContent || '';

  if (ticket || content) {

    const ticketData = {

      TicketNumbers: ticket ? [ticket] : [],

      TicketContent: content,

    };

    out += `|TICKET_DATA:${JSON.stringify(ticketData)}`;

  }

  let startVal = dateStart;
  let endVal = dateEnd;
  if (!startVal || !endVal) {
    const daysBackEl = document.getElementById('daysBack');
    const days = Math.max(1, parseInt(daysBackEl?.value, 10) || 10);
    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - days);
    if (!startVal) startVal = toDateTimeLocalValue(start);
    if (!endVal) endVal = toDateTimeLocalValue(end);
  }

  {
    const start = new Date(startVal);
    const end = new Date(endVal);
    if (!Number.isNaN(start.getTime()) && !Number.isNaN(end.getTime()) && end >= start) {
      const fmt = (d) => {
        const pad = (n) => String(n).padStart(2, '0');
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
      };
      out += `|DATE_RANGE:${JSON.stringify({ StartDate: fmt(start), EndDate: fmt(end) })}`;
    }
  }

  return out;

}

function getValidatedUsersForTenant(clientNumber, div) {
  let fromDataset = [];
  try {
    fromDataset = JSON.parse(div?.dataset?.validatedUsers || '[]');
  } catch { /* ignore */ }
  if (Array.isArray(fromDataset) && fromDataset.length) return fromDataset;

  const fromState = tenantUiState.get(String(clientNumber))?.validatedUsers;
  if (Array.isArray(fromState) && fromState.length) {
    if (div) div.dataset.validatedUsers = JSON.stringify(fromState);
    return fromState;
  }
  return [];
}

function buildGenerateReportsCommand(clientNumber, div, ticketNumber, ticketContent) {

  const filterOn = Boolean(div.querySelector('.filterUsersCheck')?.checked);

  const validatedUsers = getValidatedUsersForTenant(clientNumber, div);

  const searchTerms = parseSearchTerms(div.querySelector('.userSearchInput')?.value);

  const dateStart = div.querySelector('.dateStart')?.value;

  const dateEnd = div.querySelector('.dateEnd')?.value;

  let cmd;

  if (filterOn && validatedUsers.length) {

    cmd = `GENERATE_REPORTS|SelectedUsers:${JSON.stringify(validatedUsers)}`;

  } else if (filterOn && searchTerms.length) {

    cmd = `GENERATE_REPORTS_SEARCH:${JSON.stringify(searchTerms)}`;

  } else {

    cmd = 'GENERATE_REPORTS';

  }

  return appendTicketAndDateRange(cmd, ticketNumber, ticketContent, dateStart, dateEnd);

}

async function fetchManageTicket(clientNumber, div) {

  const ticketInput = div.querySelector('.ticketInput');

  const ticketId = (ticketInput?.value || '').trim();

  if (!ticketId) {

    log(`Client ${clientNumber}: enter a ticket # first.`);

    return;

  }

  const btn = div.querySelector('.fetchTicket');

  if (btn) btn.disabled = true;

  try {

    log(`Client ${clientNumber}: fetching Manage ticket ${ticketId}…`);

    const data = await api('/api/manage/ticket', {

      method: 'POST',

      body: JSON.stringify({ ticketId }),

    });

    div.dataset.ticketContent = data.ticketContent || '';
    const paste = div.querySelector('.ticketPaste');
    if (paste) paste.value = data.ticketContent || '';

    const preview = div.querySelector('.ticketPreview');

    if (preview) {

      const summary = data.summary ? ` — ${data.summary}` : '';
      const stackHint = data.securityStack?.labels?.length
        ? ` · Stack: ${data.securityStack.labels.join(', ')}`
        : '';

      preview.textContent = `Loaded ticket #${data.ticketId}${summary}${stackHint} (${data.contentLength || 0} chars)`;

      preview.style.display = 'block';

    }

    if (ticketInput && data.ticketId) ticketInput.value = String(data.ticketId).trim();

    if (data.companyName) {
      const existing = tenantUiState.get(String(clientNumber)) || {};
      existing.organizationHint = String(data.companyName).trim();
      tenantUiState.set(String(clientNumber), existing);
      div.dataset.organizationHint = String(data.companyName).trim();
    }

    saveTenantUiState(clientNumber, div);

    await syncTenantUiStateToServer(clientNumber, div);

    refreshTenantSummaryUI(clientNumber, data.companyName ? { exoOrganizationName: data.companyName } : {});

    await resolveSecurityIntegrations(clientNumber, div);

    log(`Client ${clientNumber}: Manage ticket loaded (${data.contentLength || 0} chars).`);

  } catch (e) {

    log(`Client ${clientNumber}: Fetch ticket failed: ${e.message}`);

  } finally {

    if (btn) btn.disabled = false;

  }

}

async function generateReports(clientNumber, div) {

  await withClientLock(clientNumber, async () => {

    const btn = div.querySelector('.generateReports');

    try {

    await syncTenantUiStateToServer(clientNumber, div);

    if (!await requireLiveWorker(clientNumber, { actionLabel: 'Generate Reports' })) {
      return;
    }

    const session = await api('/api/session');
    const tenant = session.tenants?.find(x => Number(x.clientNumber) === Number(clientNumber));
    const rsInfo = await api(`/api/tenants/${clientNumber}/report-selections`);
    const req = getRequiredAuthFromReportSelections(rsInfo.effective || {});
    const missingAuth = [];
    if (req.needsGraph && !tenant?.graphAuthenticated) missingAuth.push('Graph');
    if (req.needsExchange && !tenant?.exchangeAuthenticated) missingAuth.push('Exchange Online');
    if (missingAuth.length) {
      log(`Client ${clientNumber}: ${missingAuth.join(' and ')} must be authenticated for the selected reports.`);
      return;
    }

    const filterOn = Boolean(div.querySelector('.filterUsersCheck')?.checked);
    if (filterOn) {
      if (!tenant?.graphAuthenticated) {
        const msg = `Client ${clientNumber}: Filter by users is enabled but Graph is not authenticated.\n\nComplete Graph Auth, then Validate users, before Generate.\n\nWithout Graph, user filtering cannot be applied and exports such as Unified Audit Logs would pull the whole tenant.`;
        log(msg.replace(/\n+/g, ' '));
        window.alert(msg);
        return;
      }
      const validatedUsers = getValidatedUsersForTenant(clientNumber, div);
      const searchTerms = parseSearchTerms(div.querySelector('.userSearchInput')?.value);
      if (!validatedUsers.length && !searchTerms.length) {
        const msg = `Client ${clientNumber}: Filter by users is enabled, but no users are validated and the search box is empty.\n\nEnter a UPN/name, click Validate users, then Generate — otherwise the export would run for all users.`;
        log(msg.replace(/\n+/g, ' '));
        window.alert(msg);
        return;
      }
    }

    focusClientLogTab(clientNumber);

    const logPre = getWorkerLogPanel(clientNumber);
    if (logPre) {
      logPre.textContent += (logPre.textContent ? '\n\n' : '') + `--- Generate reports ${new Date().toLocaleString()} ---\n`;
      logPre.scrollTop = logPre.scrollHeight;
    }
    stopWorkerLogPoll(clientNumber);
    startWorkerLogPoll(clientNumber);

    btn.disabled = true;

    btn.textContent = 'Generating…';

    setTenantButtonsDisabled(div, true);

    const dateStartEl = div.querySelector('.dateStart');
    const dateEndEl = div.querySelector('.dateEnd');
    if (dateStartEl && !dateStartEl.value) dateStartEl.value = defaultDateStartValue();
    if (dateEndEl && !dateEndEl.value) dateEndEl.value = defaultDateEndValue();

    const dateStart = dateStartEl?.value;

    const dateEnd = dateEndEl?.value;

    if (dateStart && dateEnd && new Date(dateEnd) < new Date(dateStart)) {

      log(`Client ${clientNumber}: End date must be on or after start date.`);

      btn.disabled = false;

      btn.textContent = 'Generate Reports';

      setTenantButtonsDisabled(div, false);

      return;

    }

    log(`Client ${clientNumber}: Generate reports…`);
    div.dataset.lastResponse = 'GENERATE_REPORTS_STARTED';
    refreshTenantSummaryUI(clientNumber, { reportInProgress: true, lastResponse: 'GENERATE_REPORTS_STARTED' });

      const ticketNumber = div.querySelector('.ticketInput')?.value || '';

      const ticketContent = (div.querySelector('.ticketPaste')?.value || '').trim()
        || div.dataset.ticketContent
        || tenantUiState.get(String(clientNumber))?.ticketContent
        || '';

      const cmd = buildGenerateReportsCommand(clientNumber, div, ticketNumber, ticketContent);

      const initial = await api(`/api/tenants/${clientNumber}/command`, {

        method: 'POST',

        body: workerCommandBody(cmd),

      });

      let final = normalizeResponse(initial.response);

      const generateTerminal = /^(GENERATE_REPORTS_SUCCESS:|GENERATE_REPORTS_NO_DATA:|GENERATE_REPORTS_FAILED:)/;
      if (!generateTerminal.test(final)) {

        final = await pollWorkerResponse(

          clientNumber,

          'GENERATE_REPORTS_STARTED',

          ['GENERATE_REPORTS_SUCCESS:', 'GENERATE_REPORTS_NO_DATA:'],

          'GENERATE_REPORTS_FAILED:',

          1800,

          'generating reports',

          async () => {

            refreshTenantSummaryUI(clientNumber, { reportInProgress: true, lastResponse: 'GENERATE_REPORTS_STARTED' });

            const worker = await api(`/api/tenants/${clientNumber}/worker`);
            if (!worker?.alive) {
              throw new Error('PowerShell worker stopped during report generation. Restart worker, re-authenticate, and Generate again.');
            }

          }

        );

      }

      if (final.startsWith('GENERATE_REPORTS_SUCCESS:')) {

        const path = final.replace('GENERATE_REPORTS_SUCCESS:', '').trim();

        log(`Client ${clientNumber}: Reports saved to ${path}`);

        div.dataset.outputFolder = path;

        const analyzeBtn = div.querySelector('.analyzeReports');
        const openBtn = div.querySelector('.openReports');
        if (analyzeBtn) analyzeBtn.disabled = false;
        if (openBtn) openBtn.disabled = false;

        refreshTenantSummaryUI(clientNumber, { outputFolder: path, reportInProgress: false, lastResponse: final });

      } else if (final.startsWith('GENERATE_REPORTS_NO_DATA:')) {

        log(`Client ${clientNumber}: No report data (${final.replace('GENERATE_REPORTS_NO_DATA:', '').trim()})`);

      } else {

        log(`Client ${clientNumber}: ${final}`);

      }

    } catch (e) {

      log(`Client ${clientNumber}: Error: ${e.message}`);
      refreshTenantSummaryUI(clientNumber, { reportInProgress: false });

    } finally {

      if (btn) {
        btn.disabled = false;
        btn.textContent = 'Generate Reports';
      }
      setTenantButtonsDisabled(div, false);
      saveTenantUiState(clientNumber, div);

      await refreshSession();

    }

  });

}

async function openReports(folderPath) {

  if (!folderPath) return;

  log(`Opening ${folderPath}`);

  await api('/api/open-folder', {

    method: 'POST',

    body: JSON.stringify({ path: folderPath }),

  });

}

async function restartWorker(clientNumber, div, options = {}) {

  clientBusy.delete(String(clientNumber));

  saveTenantUiState(clientNumber, div);

  log(`Client ${clientNumber}: restarting worker${options.showConsole ? ' (visible console)' : ''}…`);

  try {

    const t = await api(`/api/tenants/${clientNumber}/restart`, {
      method: 'POST',
      body: JSON.stringify({ showConsole: Boolean(options.showConsole) }),
    });

    if (!options.skipReauthHint) {

      log(`Client ${clientNumber}: new worker PID ${t.processId}. Re-run Exchange Auth, then Graph Auth.`);

    } else {

      log(`Client ${clientNumber}: new worker PID ${t.processId}.`);

    }

    workerLogOffsets.set(String(clientNumber), 0);

    if (!options.skipRefresh) {

      await refreshSession();

    }

  } catch (e) {

    log(`Client ${clientNumber}: Error: ${e.message}`);

  }

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

  const current = await api('/api/session');

  if (current?.tenantCount > 0) {

    const ok = window.confirm(

      'Start a new session? This clears the in-memory session (tenant workers keep running in their PowerShell windows until you close them).'

    );

    if (!ok) return;

  }

  await api('/api/session', {

    method: 'POST',

    body: JSON.stringify(sessionBodyFromUi()),

  });

  log('Session ready.');

  await refreshSession();

}));

document.getElementById('btnAddTenant').addEventListener('click', () => runAction('Adding tenant worker…', async () => {

  await ensureSession({ quiet: true });

  const t = await api('/api/tenants', { method: 'POST' });

  log(`Started Client ${t.clientNumber} (PID ${t.processId}). A PowerShell worker window should open on this PC.`);

  await refreshSession();

}));

document.getElementById('btnRefresh').addEventListener('click', () => refreshSession().catch(e => log(`Error: ${e.message}`)));

async function loadAppRegistrations(options = {}) {
  const { quiet = false, refreshTenants = true, forceRefreshFromGraph = false } = options;
  const summaryEl = document.getElementById('appRegsSummary');
  if (summaryEl) summaryEl.textContent = '(loading…)';
  if (!quiet) log(forceRefreshFromGraph ? 'Refreshing app registrations from Graph…' : 'Loading app registrations…');
  try {
    const url = forceRefreshFromGraph ? '/api/app-registrations?refresh=1' : '/api/app-registrations';
    appRegistrations = await api(url);
    const count = appRegistrations.length;
    appRegsEl.textContent = appRegistrations.map(a => a.displayText).join('\n') || '(none in WCM)';
    if (summaryEl) {
      summaryEl.textContent = count ? `(${count} in WCM)` : '(none in WCM)';
    }
    if (!quiet) log(`Loaded ${count} app registration(s).`);
    if (refreshTenants) await refreshSession();
  } catch (e) {
    if (summaryEl) summaryEl.textContent = '(load failed)';
    appRegsEl.textContent = String(e.message || e);
    if (!quiet) log(`App registrations: ${e.message}`);
    throw e;
  }
}

document.getElementById('btnLoadAppRegs')?.addEventListener('click', () => runAction('Refreshing app registrations from Graph (may take a minute)…', () => loadAppRegistrations({ quiet: true, forceRefreshFromGraph: true })));

document.getElementById('btnOpenSessionTemp')?.addEventListener('click', () => runAction('Opening session temp folder…', async () => {
  await api('/api/session/open-temp', { method: 'POST', body: '{}' });
  log('Opened session temp folder.');
}));

document.getElementById('btnCreateGraphApp')?.addEventListener('click', () => runAction('Create Graph App (browser sign-in on this PC)…', async () => {
  const data = await api('/api/wcm/create-graph-app', { method: 'POST', body: '{}' });
  if (data.result?.WcmSaved) {
    log(`Graph app created for ${data.result.TenantDisplayName || data.result.TenantId}. Select it in App reg tenant, then Graph Auth.`);
  } else {
    log(`Create Graph App finished (exit ${data.exitCode}). See ${data.logPath || 'temp log'} if needed.`);
  }
  await loadAppRegistrations({ quiet: true, forceRefreshFromGraph: true });
}));

document.getElementById('btnDeleteGraphApp')?.addEventListener('click', async () => {
  if (!appRegistrations.length) {
    log('No app registrations loaded.');
    return;
  }
  const picks = window.prompt('Enter tenant ID(s) to delete (comma-separated):');
  if (!picks) return;
  const tenantIds = picks.split(/[,;\s]+/).map(s => s.trim()).filter(Boolean);
  if (!tenantIds.length) return;
  if (!window.confirm(`Delete Graph app registration(s) for ${tenantIds.join(', ')}? This removes Entra apps and WCM creds.`)) return;
  await runAction('Deleting Graph app(s)…', async () => {
    const data = await api('/api/wcm/delete-graph-app', { method: 'POST', body: JSON.stringify({ tenantIds }) });
    log(`Delete completed for ${(data.removed || []).length} tenant(s).`);
    await loadAppRegistrations({ quiet: true, forceRefreshFromGraph: true });
  });
});

document.getElementById('btnExportWcm')?.addEventListener('click', async () => {
  const path = window.prompt('Export file path (.eoa-creds):', `${(await api('/api/health')).projectRoot || ''}\\graph-apps.eoa-creds`);
  const password = window.prompt('Encryption password:');
  if (!path || !password) return;
  await runAction('Exporting WCM credentials…', async () => {
    await api('/api/wcm/export', { method: 'POST', body: JSON.stringify({ path, password }) });
    log(`Exported credentials to ${path}`);
  });
});

document.getElementById('btnImportWcm')?.addEventListener('click', async () => {
  const path = window.prompt('Import file path (.eoa-creds):');
  const password = window.prompt('Encryption password:');
  if (!path || !password) return;
  await runAction('Importing WCM credentials…', async () => {
    const data = await api('/api/wcm/import', { method: 'POST', body: JSON.stringify({ path, password }) });
    log(`Imported ${data.imported ?? 0} credential(s).`);
    await loadAppRegistrations({ quiet: true });
  });
});

document.getElementById('btnClearWcm')?.addEventListener('click', async () => {
  try {
    const data = await api('/api/wcm/entries');
    const entries = data.entries || [];
    if (!entries.length) {
      log('No WCM entries to clear.');
      return;
    }
    const list = entries.map((e, i) => `${i + 1}. ${e.displayText}`).join('\n');
    const picks = window.prompt(`Enter row number(s) to remove from WCM only (comma-separated):\n\n${list}`);
    if (!picks) return;
    const indices = picks.split(/[,;\s]+/).map(s => parseInt(s, 10) - 1).filter(i => i >= 0 && i < entries.length);
    if (!indices.length) return;
    const items = indices.map(i => entries[i]);
    if (!window.confirm(`Remove ${items.length} WCM entr(y/ies) from this PC only?`)) return;
    await runAction('Clearing local WCM…', async () => {
      const result = await api('/api/wcm/clear-local', { method: 'POST', body: JSON.stringify({ items }) });
      log(`Removed ${result.removed ?? 0} WCM entr(y/ies).`);
      await loadAppRegistrations({ quiet: true });
    });
  } catch (e) {
    log(`Clear WCM failed: ${e.message}`);
  }
});

document.getElementById('btnRefreshHistory')?.addEventListener('click', () => runAction('Refreshing session history…', () => loadSessionHistory()));

document.getElementById('savedHistorySearch')?.addEventListener('input', (e) => {
  historySearch.saved = e.target.value;
  renderHistorySection('saved');
});

document.getElementById('archivedHistorySearch')?.addEventListener('input', (e) => {
  historySearch.archived = e.target.value;
  renderHistorySection('archived');
});

(async () => {
  document.querySelector('#logTabs [data-log-tab="activity"]')?.addEventListener('click', () => switchLogTab('activity'));
  initSessionReportExportsPanel();
  applyReportSelectionsToContainer(document.getElementById('sessionReportExportsBody'), defaultReportSelections());
  updateSessionReportExportsSummary(defaultReportSelections());
  log('Loading…');
  try {
    await detectServerFeatures();
    // Populate presets + select/apply BEC before session hydrate (session may override selections).
    await loadExportPresetsFromServer({ applyDefaultBec: true });
    await ensureSession({ quiet: true });
    log('Session connected.');
    void loadAppRegistrations({ quiet: true }).catch(e => log(`App registrations: ${e.message}`));
    void loadSessionHistory().catch(() => {});
  } catch (e) {
    log(`Startup error: ${e.message}`);
  }
})();

