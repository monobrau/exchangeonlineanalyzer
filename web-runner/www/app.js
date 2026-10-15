const logEl = document.getElementById('log');

const sessionInfo = document.getElementById('sessionInfo');

const tenantsEl = document.getElementById('tenants');

const appRegsEl = document.getElementById('appRegs');

let appRegistrations = [];

const clientBusy = new Set();

const tenantUiState = new Map();

// Client numbers whose containment <details> was open before the last re-render.
// Saved results must not reopen it — see restoreContainmentOutput.
const containmentPanelOpen = new Set();

let currentSessionId = null;

function containmentStorageKey(clientNumber) {
  return `eoa.containment.${currentSessionId || 'session'}.${clientNumber}`;
}

function hasContainmentPayload(c) {
  if (!c || typeof c !== 'object') return false;
  return Boolean(
    c.status || c.capabilities || c.restrictedEmail || c.authMethods || c.devices
    || c.apps || c.rules || c.mailbox || c.transport || c.connectors || c.oauth
    || c.mobile || c.intune || c.folders || c.autoreply || c.orgfwd || c.junk
    || c.journal || c.hold || c.elsewhere || c.roles || c.appcreds || c.flows
    || (Array.isArray(c.actions) && c.actions.length)
  );
}

function persistContainmentToSessionStorage(clientNumber, containment) {
  if (!clientNumber || !hasContainmentPayload(containment)) return;
  try {
    sessionStorage.setItem(containmentStorageKey(clientNumber), JSON.stringify(containment));
  } catch {
    /* quota or private mode */
  }
}

function readContainmentFromSessionStorage(clientNumber) {
  if (!clientNumber) return null;
  try {
    const raw = sessionStorage.getItem(containmentStorageKey(clientNumber));
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return hasContainmentPayload(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

function resolveContainmentState(clientNumber, existing = {}) {
  if (hasContainmentPayload(existing.containment)) return existing.containment;
  return readContainmentFromSessionStorage(clientNumber);
}
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

function buildCurateLogsPanelHtml(hasFolder) {
  const disabled = hasFolder ? '' : 'disabled';
  return `
      <details class="curateLogs collapsible">
        <summary>Curate logs <span class="muted">(include/exclude related activity → curated CSV set)</span></summary>
        <div class="collapsible-body curatePanel">
          <p class="muted" style="font-size:0.85rem;margin:0 0 0.45rem">
            Load facets from the report folder, select values to exclude (noise) or include (focus), preview counts, then export a <code>Curated_*</code> folder. Originals are never modified.
          </p>
          <div class="row">
            <label>Mode
              <select class="curateModeSelect">
                <option value="exclude" selected>Exclude selected (drop noise)</option>
                <option value="include">Include selected only (focus)</option>
              </select>
            </label>
            <button type="button" class="loadCurateFacets small" ${disabled}>Load facets</button>
            <button type="button" class="previewCurate small" ${disabled}>Preview counts</button>
            <button type="button" class="exportCurate primary small" ${disabled}>Export curated set</button>
            <button type="button" class="openCuratedFolder success small" disabled>Open curated folder</button>
          </div>
          <div class="sectionLabel">Likely tenant WAN / office IPs</div>
          <p class="muted" style="font-size:0.8rem;margin:0 0 0.35rem">
            Ranked from successful public SignInLogs IPs (plus UAL/MessageTrace overlap). Verify before excluding — not a firewall source of truth.
          </p>
          <div class="curateWanSuggestions muted" style="font-size:0.85rem">Load facets to suggest WAN IPs.</div>
          <div class="row" style="margin-top:0.35rem">
            <button type="button" class="selectSuggestedWan small" ${disabled}>Select suggested for exclude</button>
            <button type="button" class="clearWanSelection small" ${disabled}>Clear WAN selection</button>
          </div>
          <label style="display:block;margin-top:0.45rem;font-size:0.85rem">Paste known WAN IPs (one per line or comma-separated)
            <textarea class="curateWanPaste" rows="2" placeholder="203.0.113.10&#10;198.51.100.20" style="width:100%;max-width:40rem;display:block;margin-top:0.2rem;font:inherit"></textarea>
          </label>
          <div class="row">
            <button type="button" class="applyWanPaste small" ${disabled}>Add pasted IPs to selection</button>
          </div>
          <div class="curateStatus muted" style="font-size:0.85rem"></div>
          <div class="curateFacets"></div>
          <div class="curatePreview muted"></div>
        </div>
      </details>`;
}

function buildBecInvestigatePanelHtml() {
  return `
      <section class="becInvestigate">
        <div class="sectionLabel">Threat check</div>
        <p class="muted" style="font-size:0.85rem;margin:0 0 0.45rem">One mailbox. Covers business email compromise, data exfiltration, DLP, and other malicious activity, including new app registrations, mailbox forwarding, delegates, transport rules, and connectors. If Filter by users has exactly one address, that mailbox is used. A ticket on this card is included when you filled one in. Open Containment when you want the lists and the remediation buttons.</p>
        <div class="row">
          <label>Mailbox <input type="text" class="becMailbox" placeholder="uses the scoped user when there is one" /></label>
          <button type="button" class="becInvestigateRun primary" disabled>Threat check</button>
          <label title="When the review needs an older event, pull a longer window for only the logs that can answer it, up to 90 days or tenant retention."><input type="checkbox" class="becAutoRepull" checked /> Pull a longer window automatically</label>
        </div>
        <p class="becInvestigateStatus muted" style="font-size:0.85rem;margin:0.2rem 0"></p>
        <div class="becReviewWrap" style="display:none">
          <div class="row">
            <button type="button" class="becReviewCopy small">Copy review</button>
            <button type="button" class="becManualRepull small" style="display:none">Pull longer window and review again</button>
          </div>
          <pre class="becReviewText copyBlock"></pre>
        </div>
        <div class="row">
          <button type="button" class="writeMemberberry small">Write notes</button>
        </div>
        <div class="becNotes" style="display:none">
          <p class="memberberryRefs muted" style="font-size:0.85rem;margin:0.2rem 0"></p>
          <div class="sectionLabel">Internal ticket note</div>
          <pre class="memberberryText copyBlock"></pre>
          <button type="button" class="copyMemberberry small">Copy internal note</button>
          <div class="sectionLabel">Client email</div>
          <pre class="becClientText copyBlock"></pre>
          <button type="button" class="copyBecClient small">Copy client email</button>
          <div class="sectionLabel">Incident report</div>
          <pre class="becIncidentText copyBlock"></pre>
          <button type="button" class="copyBecIncident small">Copy incident report</button>
        </div>
        <div class="storeSuggestions" style="display:none">
          <div class="sectionLabel">Save for next time</div>
          <p class="muted" style="font-size:0.85rem;margin:0 0 0.35rem">The last AI pass can suggest one Memberberry note and one memory-vault note. Nothing is written until you press a button. Edit the text first if it is not quite right.</p>
          <label>Memberberry client <input type="text" class="storeMbClient" placeholder="Client name" /></label>
          <textarea class="storeMbNote" rows="2" placeholder="Note to append for this client" style="width:100%;font:inherit;margin:0.25rem 0"></textarea>
          <button type="button" class="storeMbSave small">Add to Memberberry</button>
          <textarea class="storeMvNote" rows="2" placeholder="Soc memory for the next review" style="width:100%;font:inherit;margin:0.35rem 0 0.25rem"></textarea>
          <button type="button" class="storeMvSave small">Add to memory vault</button>
          <p class="storeSuggestStatus muted" style="font-size:0.85rem;margin:0.35rem 0 0"></p>
        </div>
      </section>`;
}

function buildContainmentPanelHtml() {
  return `
      <details class="containmentPanel collapsible">
        <summary>Containment <span class="muted">(BEC playbook — this tenant, validated users)</span></summary>
        <div class="collapsible-body containmentBody">
          <p class="muted" style="font-size:0.85rem;margin:0 0 0.45rem">
            Work top to bottom after <strong>Validate users</strong>. Most list/status buttons run immediately; long tenant-wide or mailbox-wide lists ask first. Writes ask for a confirm popup.
            Transport rules, connectors, apps, org auto-forward, journaling, and roles are tenant-wide.
          </p>
          <div class="containmentGraphHint muted"></div>
          <button type="button" class="containmentUpdateGraphScopes small" style="display:none" title="Patch the existing River Run Graph app with missing application permissions. Does not rotate the client secret.">Update Graph App scopes</button>
          <div class="containmentUserList muted">Validate users first.</div>
          <div class="containmentStatus" role="status" aria-live="polite"></div>

          <section class="containmentPhase">
            <div class="containmentPhaseTitle">Quick actions</div>
            <p class="containmentPhaseHint">First move on a confirmed compromise, for the users selected above. Each step is logged individually, so a partial failure is visible.</p>
            <div class="row">
              <button type="button" class="containmentAction containmentQuickLock dangerFill small" disabled title="Blocks sign-in, then revokes all sessions, refresh tokens, and MFA sessions. Registered MFA methods are kept (no re-registration) and the password is not changed.">Block access + revoke sessions</button>
            </div>
            <p class="muted" style="font-size:0.78rem;margin:0.35rem 0 0.6rem">Runs <strong>Block sign-in</strong> then <strong>Revoke sessions</strong> — blocking first so a valid password cannot mint a new token in the gap. Revoking covers MFA sessions too, so the user must satisfy MFA again, but their registered methods stay and no re-registration is forced. The password is left alone; reset it below if the attacker knows it.</p>
            <div class="row">
              <button type="button" class="containmentAction containmentQuickRestore success small" disabled title="Resets to a random password nobody is told, then unblocks sign-in. The user gets back in only at passwordreset.microsoftonline.com. If the reset fails, sign-in stays blocked.">Reset password + unblock access</button>
            </div>
            <p class="muted" style="font-size:0.78rem;margin:0.35rem 0 0">Handback after containment: <strong>Reset with random password</strong>, then <strong>Unblock sign-in</strong>. Nobody is shown the new password — send the user <a href="https://passwordreset.microsoftonline.com/" target="_blank" rel="noopener">passwordreset.microsoftonline.com</a>. If the reset fails, the account stays blocked.</p>
            <div class="row" style="margin-top:0.6rem">
              <button type="button" class="containmentBecCheck small" disabled title="One selected user. Pulls mailbox persistence, forwarding, delegates, transport rules, connectors, new app registrations, MFA, devices, Intune, OAuth, roles, sign-ins, message trace, unified audit, DLP, and security alerts. A ticket is optional.">Threat check</button>
            </div>
            <p class="muted" style="font-size:0.78rem;margin:0.35rem 0 0">Same check as Threat check at the top of this tenant. It uses the mailbox already filled there, or the one selected user here. A clock time on the ticket is this PC's local time. The review, ticket note, client email, and incident report use that same local time. The first pull is only as long as that time needs, or the last 2 days when the ticket has no time. When Pull a longer window automatically is checked, the check keeps pulling a longer window for the logs it still needs, up to 90 days or retention, until the investigation can be finished. Otherwise use Pull longer window and review again.</p>
          </section>

          <div class="row">
            <button type="button" class="containmentSavePacks small" title="Write per-user zips into the current report folder">Save containment zips</button>
            <button type="button" class="containmentClearUserPulls small" title="Clear per-user list results so you can pull the next user. Tenant-wide lists and the account-change log stay.">Clear user pulls</button>
          </div>
          <p class="muted" style="font-size:0.78rem;margin:0 0 0.45rem">Save writes <code>Containment_&lt;user&gt;.zip</code> (with <code>actions.csv</code>) and <code>Remediation.csv</code> into the current report folder. Clear user pulls removes MFA/mailbox/device rows for the next user; the account-change log, tenant-wide lists, and saved zips stay.</p>

          <section class="containmentPhase">
            <div class="containmentPhaseTitle">1. Lock the account</div>
            <p class="containmentPhaseHint">Stop the attacker’s current session, then invalidate the password. Send the user <a href="https://passwordreset.microsoftonline.com/" target="_blank" rel="noopener">passwordreset.microsoftonline.com</a> — not a password. Block sign-in if you need the account frozen while you hunt.</p>
            <div class="row">
              <button type="button" class="containmentAction containmentSigninStatus small" disabled>Check sign-in status</button>
              <button type="button" class="containmentAction containmentRevoke small" disabled>Revoke sessions</button>
              <button type="button" class="containmentAction containmentResetPassword small" disabled>Reset with random password</button>
              <button type="button" class="containmentAction containmentBlock small" disabled>Block sign-in</button>
            </div>
            <p class="muted" style="font-size:0.8rem;margin:0.35rem 0 0.25rem">
              After reset, send <a href="https://passwordreset.microsoftonline.com/" target="_blank" rel="noopener">passwordreset.microsoftonline.com</a>. Assign a password only if you must.
            </p>
            <div class="row">
              <label>Assign password (optional)
                <input type="password" class="containmentAssignPasswordInput" autocomplete="new-password" placeholder="only if you must set a specific password" />
              </label>
              <button type="button" class="containmentAction containmentAssignPasswordBtn small" disabled>Set this password</button>
            </div>
            <div class="containmentSubLabel">Preserve evidence</div>
            <p class="muted" style="font-size:0.78rem;margin:0 0 0.25rem">Turn on litigation hold, 30-day deleted-item retention, and mailbox audit before you delete rules or mail.</p>
            <div class="row">
              <button type="button" class="containmentAction containmentHoldStatus small" disabled>Check hold / audit</button>
              <button type="button" class="containmentAction containmentEnableHold small" disabled>Enable hold + audit</button>
            </div>
            <div class="containmentHoldWrap" style="display:none">
              <table class="historyTable containmentHoldTable">
                <thead>
                  <tr>
                    <th>Mailbox</th>
                    <th>Litigation hold</th>
                    <th>Hold date</th>
                    <th>Retain deleted</th>
                    <th>Audit</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
          </section>

          <section class="containmentPhase">
            <div class="containmentPhaseTitle">2. Identity footholds</div>
            <p class="containmentPhaseHint">List first, then remove attacker MFA methods, OAuth consents, Entra devices, Exchange ActiveSync partnerships, and Intune-managed devices.</p>
            <div class="containmentSubLabel">MFA methods</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListMfa small" disabled>List MFA methods</button>
              <button type="button" class="containmentAction containmentRevokeMfa small" disabled>Revoke MFA sessions</button>
              <button type="button" class="containmentAction containmentDeleteMfa danger small" disabled>Remove selected MFA methods</button>
              <button type="button" class="containmentAction containmentReregisterMfa danger small" disabled>Wipe MFA + require re-register</button>
            </div>
            <div class="containmentMfaWrap" style="display:none">
              <table class="historyTable containmentMfaTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>User</th>
                    <th>Type</th>
                    <th>Details</th>
                    <th>Created</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Registered devices</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListDevices small" disabled>List registered devices</button>
              <button type="button" class="containmentAction containmentDeleteDevices danger small" disabled>Remove selected devices</button>
            </div>
            <div class="containmentDevicesWrap" style="display:none">
              <table class="historyTable containmentDevicesTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>User</th>
                    <th>Device</th>
                    <th>OS</th>
                    <th>Trust</th>
                    <th>Relation</th>
                    <th>Last sign-in</th>
                  </tr>
                </thead>
                <tbody>                </tbody>
              </table>
            </div>
            <div class="containmentSubLabel">OAuth consents</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListOauth small" disabled>List OAuth consents</button>
              <button type="button" class="containmentAction containmentDeleteOauth danger small" disabled>Revoke selected consents</button>
            </div>
            <div class="containmentOauthWrap" style="display:none">
              <table class="historyTable containmentOauthTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>User</th>
                    <th>App</th>
                    <th>Scopes</th>
                    <th>Type</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Exchange ActiveSync</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListMobile small" disabled>List mobile partnerships</button>
              <button type="button" class="containmentAction containmentDeleteMobile danger small" disabled>Remove selected partnerships</button>
            </div>
            <div class="containmentMobileWrap" style="display:none">
              <table class="historyTable containmentMobileTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>User</th>
                    <th>Device</th>
                    <th>Type</th>
                    <th>First sync</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Intune managed devices</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListIntune small" disabled>List Intune devices</button>
              <button type="button" class="containmentAction containmentRetireIntune small" disabled>Retire selected</button>
              <button type="button" class="containmentAction containmentWipeIntune danger small" disabled>Wipe selected</button>
            </div>
            <div class="containmentIntuneWrap" style="display:none">
              <table class="historyTable containmentIntuneTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>User</th>
                    <th>Device</th>
                    <th>OS</th>
                    <th>Compliance</th>
                    <th>Last sync</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
          </section>

          <section class="containmentPhase">
            <div class="containmentPhaseTitle">3. Mailbox persistence</div>
            <p class="containmentPhaseHint">Classic BEC: hidden inbox rules, mailbox forwarding (SMTP and/or internal recipient), unexpected delegates, folder ACL, auto-reply, junk allow-lists, and rights this user has on other mailboxes. Check Restricted Users if outbound mail was blocked.</p>
            <div class="containmentSubLabel">AI review for this mailbox</div>
            <p class="muted" style="font-size:0.78rem;margin:0 0 0.35rem">Same check as Threat check under Quick actions. A ticket on this card is included when you filled one in. Remediate runs only the boxes you leave checked, and it uses the same containment actions as the buttons below.</p>
            <button type="button" class="containmentBecReview small" disabled>Review this mailbox</button>
            <div class="containmentBecWrap" style="display:none">
              <button type="button" class="containmentBecCopy small">Copy review</button>
              <button type="button" class="becManualRepull small" style="display:none">Pull longer window and review again</button>
              <pre class="containmentBecText copyBlock"></pre>
              <div class="containmentBecActions" style="display:none">
                <div class="row">
                  <button type="button" class="containmentBecAll small">All</button>
                  <button type="button" class="containmentBecNone small">None</button>
                  <button type="button" class="containmentBecRun danger small">Remediate</button>
                </div>
                <div class="containmentBecList"></div>
              </div>
            </div>
            <div class="containmentSubLabel">Inbox rules</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListRules small" disabled>List inbox rules</button>
              <button type="button" class="containmentAction containmentDeleteRules danger small" disabled>Delete selected rules</button>
            </div>
            <div class="containmentRulesWrap" style="display:none">
              <table class="historyTable containmentRulesTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Name</th>
                    <th>On</th>
                    <th>Priority</th>
                    <th>Hidden</th>
                    <th>Details</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Forwarding and delegation</div>
            <div class="row">
              <button type="button" class="containmentAction containmentMailboxStatus small" disabled>Check mailbox access</button>
              <button type="button" class="containmentAction containmentRemoveForward danger small" disabled>Remove selected forwarding</button>
              <button type="button" class="containmentAction containmentRemoveDelegate danger small" disabled>Remove selected delegates</button>
              <button type="button" class="containmentAction containmentClearForward small" disabled>Clear all forwarding</button>
            </div>
            <div class="containmentMailboxWrap" style="display:none">
              <table class="historyTable containmentMailboxTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Mailbox</th>
                    <th>Type</th>
                    <th>Target</th>
                    <th>Keep copy</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Add forwarding or a delegate (only if needed after cleanup)</div>
            <div class="row">
              <label>Forward to
                <input type="text" class="containmentForwardTo" placeholder="user@domain.com" />
              </label>
              <label><input type="checkbox" class="containmentForwardKeep" checked /> Keep a copy</label>
              <button type="button" class="containmentAction containmentSetForward small" disabled>Set forwarding</button>
            </div>
            <div class="row">
              <label>Delegate
                <input type="text" class="containmentDelegateUser" placeholder="delegate@domain.com" />
              </label>
              <label>Right
                <select class="containmentDelegateRight">
                  <option value="FullAccess">Full Access</option>
                  <option value="SendAs">Send As</option>
                  <option value="SendOnBehalf">Send on Behalf</option>
                </select>
              </label>
              <button type="button" class="containmentAction containmentAddDelegate small" disabled>Add delegate</button>
            </div>
            <div class="containmentSubLabel">Folder permissions</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListFolders small" disabled title="Walks every folder on the selected mailbox(es). Can take a minute or more.">List folder permissions</button>
              <button type="button" class="containmentAction containmentDeleteFolders danger small" disabled>Remove selected folder permissions</button>
            </div>
            <div class="containmentFoldersWrap" style="display:none">
              <table class="historyTable containmentFoldersTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Mailbox</th>
                    <th>Folder</th>
                    <th>User</th>
                    <th>Rights</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Automatic replies</div>
            <div class="row">
              <button type="button" class="containmentAction containmentAutoreplyStatus small" disabled>Check auto-reply</button>
              <button type="button" class="containmentAction containmentDisableAutoreply small" disabled>Disable auto-reply</button>
            </div>
            <div class="containmentAutoreplyWrap" style="display:none">
              <table class="historyTable containmentAutoreplyTable">
                <thead>
                  <tr>
                    <th>Mailbox</th>
                    <th>State</th>
                    <th>External</th>
                    <th>Start</th>
                    <th>End</th>
                    <th>Message</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Junk trusted senders</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListJunk small" disabled>List trusted senders</button>
              <button type="button" class="containmentAction containmentDeleteJunk danger small" disabled>Remove selected trusted entries</button>
            </div>
            <div class="containmentJunkWrap" style="display:none">
              <table class="historyTable containmentJunkTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Mailbox</th>
                    <th>List</th>
                    <th>Address</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Rights on other mailboxes</div>
            <p class="muted" style="font-size:0.78rem;margin:0 0 0.25rem">Send As and Send on Behalf are fast. Full Access scans every mailbox and can take several minutes.</p>
            <div class="row">
              <button type="button" class="containmentAction containmentListElsewhere small" disabled title="Send As / Send on Behalf are quick. Full Access scans every mailbox and can take several minutes.">List rights elsewhere</button>
              <button type="button" class="containmentAction containmentDeleteElsewhere danger small" disabled>Remove selected grants</button>
            </div>
            <div class="containmentElsewhereWrap" style="display:none">
              <table class="historyTable containmentElsewhereTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>User</th>
                    <th>Mailbox</th>
                    <th>Right</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Restricted from sending</div>
            <div class="row">
              <button type="button" class="containmentAction containmentRestrictedStatus small" disabled>Check restricted status</button>
              <button type="button" class="containmentAction containmentUnrestrict small" style="display:none" disabled>Unrestrict</button>
            </div>
          </section>

          <section class="containmentPhase">
            <div class="containmentPhaseTitle">4. Tenant-wide persistence</div>
            <p class="containmentPhaseHint">Malicious mail-flow rules, connectors, org auto-forward, journaling, app registrations, secrets/owners, directory roles / groups / Exchange RBAC, and Power Automate if the admin module is loaded.</p>
            <div class="containmentSubLabel">Transport rules</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListTransport small" disabled title="Tenant-wide. Large tenants can take a minute.">List transport rules</button>
              <button type="button" class="containmentAction containmentDeleteTransport danger small" disabled>Delete selected transport rules</button>
            </div>
            <div class="containmentTransportWrap" style="display:none">
              <table class="historyTable containmentTransportTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Name</th>
                    <th>On</th>
                    <th>Priority</th>
                    <th>Details</th>
                    <th>Changed</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Connectors</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListConnectors small" disabled>List connectors</button>
              <button type="button" class="containmentAction containmentDeleteConnectors danger small" disabled>Delete selected connectors</button>
            </div>
            <div class="containmentConnectorsWrap" style="display:none">
              <table class="historyTable containmentConnectorsTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Direction</th>
                    <th>Name</th>
                    <th>On</th>
                    <th>Details</th>
                    <th>Created</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">App registrations and other-tenant apps</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListApps small" disabled title="Walks the whole tenant. Large directories can take a minute or more.">List app registrations</button>
              <button type="button" class="containmentAction containmentDeleteApps danger small" disabled>Remove selected apps</button>
            </div>
            <div class="containmentAppsWrap" style="display:none">
              <table class="historyTable containmentAppsTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Kind</th>
                    <th>Name</th>
                    <th>App ID</th>
                    <th>Created</th>
                    <th>Publisher</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">App secrets, certificates, and owners</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListAppcreds small" disabled title="Walks every app registration. Can take several minutes.">List secrets / owners</button>
              <button type="button" class="containmentAction containmentDeleteAppcreds danger small" disabled>Remove selected secrets / owners</button>
            </div>
            <div class="containmentAppcredsWrap" style="display:none">
              <table class="historyTable containmentAppcredsTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Kind</th>
                    <th>App</th>
                    <th>Name</th>
                    <th>Created</th>
                    <th>Expires</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Org auto-forward</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListOrgfwd small" disabled>List org auto-forward</button>
              <button type="button" class="containmentAction containmentDisableOrgfwd danger small" disabled>Disable selected auto-forward</button>
            </div>
            <div class="containmentOrgfwdWrap" style="display:none">
              <table class="historyTable containmentOrgfwdTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Kind</th>
                    <th>Name</th>
                    <th>Auto-forward</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Journaling rules</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListJournal small" disabled>List journal rules</button>
              <button type="button" class="containmentAction containmentDeleteJournal danger small" disabled>Delete selected journal rules</button>
            </div>
            <div class="containmentJournalWrap" style="display:none">
              <table class="historyTable containmentJournalTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Name</th>
                    <th>Recipient</th>
                    <th>Journal to</th>
                    <th>On</th>
                    <th>Changed</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Directory roles, groups, Exchange RBAC</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListRoles small" disabled title="Directory roles, group memberships, and Exchange RBAC. Can take a minute or more.">List roles and groups</button>
              <button type="button" class="containmentAction containmentDeleteRoles danger small" disabled>Remove selected roles / groups</button>
            </div>
            <div class="containmentRolesWrap" style="display:none">
              <table class="historyTable containmentRolesTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>User</th>
                    <th>Kind</th>
                    <th>Name</th>
                    <th>Details</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
            <div class="containmentSubLabel">Power Automate</div>
            <div class="row">
              <button type="button" class="containmentAction containmentListFlows small" disabled title="Tenant-wide Power Automate list. Can take a minute if the admin module is loaded.">List flows</button>
              <button type="button" class="containmentAction containmentDeleteFlows danger small" disabled>Delete selected flows</button>
            </div>
            <div class="containmentFlowsWrap" style="display:none">
              <table class="historyTable containmentFlowsTable">
                <thead>
                  <tr>
                    <th></th>
                    <th>Name</th>
                    <th>Environment</th>
                    <th>Enabled</th>
                    <th>Created</th>
                  </tr>
                </thead>
                <tbody></tbody>
              </table>
            </div>
          </section>

          <section class="containmentPhase">
            <div class="containmentPhaseTitle">5. Restore access</div>
            <p class="containmentPhaseHint">After persistence is gone and the user has reset at <a href="https://passwordreset.microsoftonline.com/" target="_blank" rel="noopener">passwordreset.microsoftonline.com</a>. Unrestrict only if Check restricted status found them on Restricted entities.</p>
            <div class="row">
              <button type="button" class="containmentAction containmentUnblock small" disabled>Unblock sign-in</button>
            </div>
          </section>
        </div>
      </details>`;
}

function buildCaCatalogPanelHtml() {
  return `
      <details class="caCatalogPanel collapsible">
        <summary>Conditional Access <span class="muted">(catalog + user exceptions — Graph Auth, then Refresh)</span></summary>
        <div class="caCatalogBackdrop" hidden></div>
        <div class="collapsible-body caCatalogBody">
          <div class="caCatalogPopupBar">
            <h3 class="caCatalogPopupTitle">Conditional Access</h3>
            <button type="button" class="caCatalogClosePopup small" title="Return the catalog to the tenant card">Close popup</button>
          </div>
          <p class="muted" style="font-size:0.85rem;margin:0 0 0.45rem">
            Current policies and named locations after <strong>Graph Auth</strong>. Open a policy to add or remove included users and exclusions. <strong>Refresh</strong> uses the worker cache; <strong>Hard refresh</strong> reloads from Graph.
          </p>
          <div class="caCatalogHint muted">Complete Graph Auth to list current policies.</div>
          <div class="row">
            <button type="button" class="caCatalogRefresh small" disabled title="List policies and named locations from the worker session cache.">Refresh</button>
            <button type="button" class="caCatalogHardRefresh small" disabled title="Bypass the session cache and reload policies and named locations from Graph.">Hard refresh</button>
            <button type="button" class="caCatalogOpenPopup small" title="Open this catalog in a wider popup">Open popup</button>
          </div>
          <div class="caCatalogStatus" role="status" aria-live="polite"></div>
          <div class="caCatalogSubLabel">Policies</div>
          <div class="caCatalogPoliciesWrap">
            <table class="historyTable caCatalogPoliciesTable">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>State</th>
                  <th>Users</th>
                  <th>Exclusions</th>
                  <th>Locations</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
          <div class="caCatalogSubLabel">Named locations</div>
          <div class="caCatalogLocationsWrap">
            <table class="historyTable caCatalogLocationsTable">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>Type</th>
                  <th>IP ranges / countries</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
          <div class="caCatalogDrawer" hidden></div>
        </div>
      </details>`;
}

const TRIAGE_SUCCESS_PREFIXES = ['TRIAGE_QUARANTINE:', 'TRIAGE_DELIVERY:', 'TRIAGE_IMPERSONATION:', 'TRIAGE_CONFIG:', 'TRIAGE_ALL:'];

function buildEmailTriagePanelHtml() {
  return `
      <details class="emailTriagePanel collapsible">
        <summary>Email triage <span class="muted">(pulls are read-only — choose blocks, allows, and config fixes on the triage page)</span></summary>
        <div class="collapsible-body">
          <p class="muted" style="font-size:0.85rem;margin:0 0 0.45rem">
            Quarantine, delivery, impersonation, and config pulls stay read-only. On the triage page, select phishing to block, legitimate mail to release, or config fixes. Apply runs the Exchange commands in this worker. DNS and unlicensed Defender changes are copy-only. The downloaded actions.ps1 file stays commented.
          </p>
          <div class="row">
            <label>Days <input type="number" class="triageDays" min="1" max="30" value="7" /></label>
            <label>Sender domain <input type="text" class="triageSenderDomain" placeholder="optional" /></label>
            <label>Header top N <input type="number" class="triageHeaderTopN" min="0" max="25" value="5" /></label>
            <label>Leg merge seconds <input type="number" class="triageWindow" min="1" max="600" value="120" /></label>
            <label><input type="checkbox" class="triageFreemail" checked /> Freemail senders only</label>
          </div>
          <div class="row">
            <button type="button" class="triagePull small" data-kind="QUARANTINE">Quarantine</button>
            <button type="button" class="triagePull small" data-kind="DELIVERY">Delivery</button>
            <button type="button" class="triagePull small" data-kind="IMPERSONATION">Impersonation</button>
            <button type="button" class="triagePull small" data-kind="CONFIG">Config</button>
            <button type="button" class="triagePull small" data-kind="ALL">All</button>
            <a class="emailTriageOpen" href="triage.html" target="_blank">Open triage</a>
          </div>
          <div class="emailTriageStatus muted" role="status"></div>
        </div>
      </details>`;
}

function parseTriagePayload(final) {
  const prefixes = ['TRIAGE_FAILED:', ...TRIAGE_SUCCESS_PREFIXES];
  const text = extractWorkerTokenResponse(final || '', prefixes);
  for (const prefix of prefixes) {
    if (text.startsWith(prefix)) {
      const rest = text.slice(prefix.length);
      try { return { prefix, data: JSON.parse(rest), raw: rest }; }
      catch { return { prefix, data: null, raw: rest }; }
    }
  }
  return { prefix: '', data: null, raw: text };
}

async function wireEmailTriagePanel(div, clientNumber) {
  const statusEl = div.querySelector('.emailTriageStatus');
  const openLink = div.querySelector('.emailTriageOpen');
  if (openLink) openLink.href = `triage.html?client=${encodeURIComponent(clientNumber)}`;
  div.querySelectorAll('.triagePull').forEach((btn) => {
    btn.addEventListener('click', () => runEmailTriagePull(div, clientNumber, btn.dataset.kind).catch((err) => {
      if (statusEl) statusEl.textContent = err.message || String(err);
      log(`Client ${clientNumber}: ${err.message || err}`);
    }));
  });
}

async function runEmailTriagePull(div, clientNumber, kind) {
  if (div.dataset.exchangeAuthenticated !== '1') {
    const msg = `Client ${clientNumber}: Exchange Auth is required before an email triage pull.`;
    log(msg);
    window.alert(msg);
    return;
  }
  if (!await requireLiveWorker(clientNumber, { actionLabel: 'email triage' })) return;
  const days = div.querySelector('.triageDays')?.value || '7';
  const domain = (div.querySelector('.triageSenderDomain')?.value || '').trim();
  const topN = div.querySelector('.triageHeaderTopN')?.value || '5';
  const windowSec = div.querySelector('.triageWindow')?.value || '120';
  const freemail = div.querySelector('.triageFreemail')?.checked ? '1' : '0';
  const statusEl = div.querySelector('.emailTriageStatus');
  if (Number(days) > 10 && (kind === 'DELIVERY' || kind === 'ALL' || kind === 'IMPERSONATION')) {
    if (statusEl) statusEl.textContent = 'Message trace only covers about 10 days for some statuses. This pull will warn instead of running a historical trace.';
  }
  const parts = [`TRIAGE_${kind}`, `DAYS:${days}`, `FREEMAIL:${freemail}`, `HEADER_TOPN:${topN}`, `WINDOW:${windowSec}`];
  if (domain) parts.push(`SENDER_DOMAIN:${domain}`);
  const priorFolder = div.dataset.triageFolder || '';
  if (priorFolder) parts.push(`FOLDER:${priorFolder}`);
  if (statusEl) statusEl.textContent = `Running ${kind}…`;
  const initial = await api(`/api/tenants/${clientNumber}/command`, {
    method: 'POST',
    body: workerCommandBody(parts.join('|')),
  });
  let final = extractWorkerTokenResponse(initial.response, ['TRIAGE_FAILED:', ...TRIAGE_SUCCESS_PREFIXES, 'TRIAGE_STARTED']);
  if (!final || final === 'TRIAGE_STARTED') {
    final = await pollWorkerResponse(
      clientNumber,
      'TRIAGE_STARTED',
      TRIAGE_SUCCESS_PREFIXES,
      'TRIAGE_FAILED:',
      900,
      `email triage ${kind}`
    );
  }
  const parsed = parseTriagePayload(final || '');
  if (parsed.prefix === 'TRIAGE_FAILED:' || !parsed.prefix) {
    throw new Error(parsed.raw || 'Email triage failed');
  }
  const folder = parsed.data && parsed.data.folder;
  if (folder) {
    div.dataset.triageFolder = folder;
    const openLink = div.querySelector('.emailTriageOpen');
    if (openLink) openLink.href = `triage.html?client=${encodeURIComponent(clientNumber)}&folder=${encodeURIComponent(folder)}`;
  }
  const warning = parsed.data && parsed.data.warning;
  if (statusEl) statusEl.textContent = warning || `Pack ready (${parsed.data && parsed.data.rows != null ? parsed.data.rows : '?'} rows). Use Open triage.`;
}

function upsertCurateRule(rules, mode, source, facet, value) {
  if (!source || !facet || value == null || value === '') return;
  let rule = rules.find((r) => r.source === source && r.facet === facet);
  if (!rule) {
    rule = { source, facet, op: mode === 'include' ? 'include' : 'exclude', values: [] };
    rules.push(rule);
  }
  if (!rule.values.includes(value)) rule.values.push(value);
}

function parseWanIpList(text) {
  if (!text) return [];
  return [...new Set(
    String(text)
      .split(/[\s,;]+/)
      .map((s) => s.trim())
      .filter((s) => s && /[:.]/.test(s))
  )];
}

function collectCurateRules(div) {
  const rules = [];
  const mode = div.querySelector('.curateModeSelect')?.value || 'exclude';
  div.querySelectorAll('.curateValueCheck:checked').forEach((cb) => {
    const source = cb.dataset.source;
    const facet = cb.dataset.facet;
    let value = cb.dataset.value;
    if (!source || !facet || value == null) return;
    try { value = decodeURIComponent(value); } catch { /* keep raw */ }
    upsertCurateRule(rules, mode, source, facet, value);
  });
  div.querySelectorAll('.curateWanCheck:checked').forEach((cb) => {
    let value = cb.dataset.value;
    try { value = decodeURIComponent(value); } catch { /* keep raw */ }
    upsertCurateRule(rules, mode, 'SignInLogs', 'IPAddress', value);
    upsertCurateRule(rules, mode, 'UnifiedAuditLogs', 'ClientIP', value);
  });
  const pasted = parseWanIpList(div.querySelector('.curateWanPaste')?.value || '');
  const extra = Array.isArray(div._curateExtraWanIps) ? div._curateExtraWanIps : [];
  [...pasted, ...extra].forEach((ip) => {
    upsertCurateRule(rules, mode, 'SignInLogs', 'IPAddress', ip);
    upsertCurateRule(rules, mode, 'UnifiedAuditLogs', 'ClientIP', ip);
  });
  return { mode, rules };
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function renderCurateWanSuggestions(div, wan) {
  const host = div.querySelector('.curateWanSuggestions');
  if (!host) return;
  const list = wan?.suggestions || [];
  if (!list.length) {
    host.innerHTML = '<span class="muted">No public successful sign-in IPs found to suggest as WAN.</span>';
    return;
  }
  host.innerHTML = `
    <div class="curateValues">
      ${list.map((s) => {
        const enc = encodeURIComponent(String(s.ip));
        const meta = `${s.successCount} ok / ${s.failureCount} fail · ${s.userCount} user(s)` +
          (s.countries?.length ? ` · ${escapeHtml(s.countries.slice(0, 2).join(', '))}` : '');
        const reason = escapeHtml(s.reason || '');
        const checked = s.suggested ? '' : '';
        return `<label title="${reason}"><input type="checkbox" class="curateWanCheck" data-value="${enc}" data-suggested="${s.suggested ? '1' : '0'}" ${checked}/> <strong>${escapeHtml(s.ip)}</strong> <span class="muted">(${meta})</span></label>`;
      }).join('')}
    </div>
    <p class="muted" style="font-size:0.78rem;margin:0.35rem 0 0">${escapeHtml(wan.note || '')}</p>`;
}

function setCurateIpFacetChecks(div, ips, checked) {
  const want = new Set((ips || []).map(String));
  div.querySelectorAll('.curateValueCheck').forEach((cb) => {
    if (cb.dataset.source !== 'SignInLogs' && cb.dataset.source !== 'UnifiedAuditLogs') return;
    if (cb.dataset.facet !== 'IPAddress' && cb.dataset.facet !== 'ClientIP') return;
    let value = cb.dataset.value;
    try { value = decodeURIComponent(value); } catch { /* keep */ }
    if (want.has(value)) cb.checked = checked;
  });
  div.querySelectorAll('.curateWanCheck').forEach((cb) => {
    let value = cb.dataset.value;
    try { value = decodeURIComponent(value); } catch { /* keep */ }
    if (want.has(value)) cb.checked = checked;
  });
}

function renderCurateFacets(div, sources) {
  const host = div.querySelector('.curateFacets');
  const status = div.querySelector('.curateStatus');
  if (!host) return;
  const present = (sources || []).filter((s) => s.present);
  if (!present.length) {
    host.innerHTML = '<p class="muted">No curatable CSVs found in this report folder.</p>';
    if (status) status.textContent = 'No SignInLogs / audit / mail CSVs detected.';
    return;
  }
  if (status) {
    status.textContent = present.map((s) => `${s.name}: ${s.rowCount} rows`).join(' · ');
  }
  host.innerHTML = present.map((src) => {
    const facets = (src.facets || []).map((f) => {
      const values = (f.values || []).map((v) => {
        const label = `${escapeHtml(v.value)} (${v.count})`;
        const encVal = encodeURIComponent(String(v.value));
        return `<label><input type="checkbox" class="curateValueCheck" data-source="${src.name}" data-facet="${f.name}" data-value="${encVal}" /> ${label}</label>`;
      }).join('');
      if (!values) return '';
      return `<div class="curateFacet"><div class="curateFacetName">${escapeHtml(f.name)} <span class="muted">via ${escapeHtml(f.column)}</span></div><div class="curateValues">${values}</div></div>`;
    }).join('');
    return `<div class="curateSource"><div class="curateSourceTitle">${escapeHtml(src.name)} <span class="muted">(${src.rowCount})</span></div>${facets || '<span class="muted">No facet columns matched.</span>'}</div>`;
  }).join('');
}

function formatCuratePreview(files) {
  if (!files || !files.length) return 'No files to curate.';
  return files.map((f) => `${f.source}: ${f.beforeCount} → ${f.afterCount} (dropped ${f.dropped})`).join('\n');
}

async function loadCurateFacets(clientNumber, div) {
  const folder = resolveTenantOutputFolder({ clientNumber, outputFolder: div.dataset.outputFolder }, div);
  if (!folder) {
    log(`Client ${clientNumber}: no report folder yet.`);
    return;
  }
  try {
    log(`Client ${clientNumber}: loading curation facets + WAN suggestions from ${folder}…`);
    const data = await api('/api/curate/facets', {
      method: 'POST',
      body: JSON.stringify({ path: folder, topValues: 40, wanTop: 12 }),
    }, 300000);
    div._curateFacets = data.result;
    renderCurateFacets(div, data.result?.sources || []);
    renderCurateWanSuggestions(div, data.result?.wanSuggestions);
    const wanCount = data.result?.wanSuggestions?.count || 0;
    log(`Client ${clientNumber}: curation facets loaded (${wanCount} WAN suggestion(s)).`);
  } catch (e) {
    log(`Client ${clientNumber}: Load facets failed: ${e.message}`);
  }
}

async function previewCurate(clientNumber, div) {
  const folder = resolveTenantOutputFolder({ clientNumber, outputFolder: div.dataset.outputFolder }, div);
  if (!folder) {
    log(`Client ${clientNumber}: no report folder yet.`);
    return;
  }
  const { mode, rules } = collectCurateRules(div);
  if (!rules.length) {
    log(`Client ${clientNumber}: select at least one facet value to preview.`);
    return;
  }
  try {
    log(`Client ${clientNumber}: previewing curation (${mode}, ${rules.length} rule group(s))…`);
    const data = await api('/api/curate/preview', {
      method: 'POST',
      body: JSON.stringify({ path: folder, mode, rules }),
    }, 300000);
    const previewEl = div.querySelector('.curatePreview');
    if (previewEl) previewEl.textContent = formatCuratePreview(data.result?.files || []);
    log(`Client ${clientNumber}: curation preview ready.`);
  } catch (e) {
    log(`Client ${clientNumber}: Curate preview failed: ${e.message}`);
  }
}

async function exportCurate(clientNumber, div) {
  const folder = resolveTenantOutputFolder({ clientNumber, outputFolder: div.dataset.outputFolder }, div);
  if (!folder) {
    log(`Client ${clientNumber}: no report folder yet.`);
    return;
  }
  const { mode, rules } = collectCurateRules(div);
  if (!rules.length) {
    log(`Client ${clientNumber}: select at least one facet value to export.`);
    return;
  }
  try {
    log(`Client ${clientNumber}: exporting curated set (${mode})…`);
    const data = await api('/api/curate/export', {
      method: 'POST',
      body: JSON.stringify({ path: folder, mode, rules }),
    }, 300000);
    const out = data.result?.outputFolder || '';
    div.dataset.curatedFolder = out;
    const openBtn = div.querySelector('.openCuratedFolder');
    if (openBtn) openBtn.disabled = !out;
    const previewEl = div.querySelector('.curatePreview');
    if (previewEl) {
      previewEl.textContent = `${formatCuratePreview(data.result?.files || [])}\n\nWrote: ${out}`;
    }
    log(`Client ${clientNumber}: curated set written to ${out}`);
    fillInvestigateFolderSelect(div, { clientNumber, outputFolder: div.dataset.outputFolder });
  } catch (e) {
    log(`Client ${clientNumber}: Curate export failed: ${e.message}`);
  }
}

function wireCurateLogsPanel(div, clientNumber) {
  div.querySelector('.loadCurateFacets')?.addEventListener('click', () => loadCurateFacets(clientNumber, div));
  div.querySelector('.previewCurate')?.addEventListener('click', () => previewCurate(clientNumber, div));
  div.querySelector('.exportCurate')?.addEventListener('click', () => exportCurate(clientNumber, div));
  div.querySelector('.openCuratedFolder')?.addEventListener('click', () => {
    const path = div.dataset.curatedFolder;
    if (path) openReports(path);
  });
  div.querySelector('.curateModeSelect')?.addEventListener('change', () => {
    const previewEl = div.querySelector('.curatePreview');
    if (previewEl) previewEl.textContent = 'Mode changed — run Preview counts again.';
  });
  div.querySelector('.selectSuggestedWan')?.addEventListener('click', () => {
    const mode = div.querySelector('.curateModeSelect');
    if (mode) mode.value = 'exclude';
    const ips = [];
    div.querySelectorAll('.curateWanCheck').forEach((cb) => {
      if (cb.dataset.suggested !== '1') return;
      cb.checked = true;
      let value = cb.dataset.value;
      try { value = decodeURIComponent(value); } catch { /* keep */ }
      ips.push(value);
    });
    setCurateIpFacetChecks(div, ips, true);
    log(`Client ${clientNumber}: selected ${ips.length} suggested WAN IP(s) for exclude.`);
  });
  div.querySelector('.clearWanSelection')?.addEventListener('click', () => {
    div.querySelectorAll('.curateWanCheck').forEach((cb) => { cb.checked = false; });
    div._curateExtraWanIps = [];
    const paste = div.querySelector('.curateWanPaste');
    if (paste) paste.value = '';
    log(`Client ${clientNumber}: cleared WAN IP selection.`);
  });
  div.querySelector('.applyWanPaste')?.addEventListener('click', () => {
    const ips = parseWanIpList(div.querySelector('.curateWanPaste')?.value || '');
    if (!ips.length) {
      log(`Client ${clientNumber}: no IPs found in paste box.`);
      return;
    }
    div._curateExtraWanIps = ips;
    setCurateIpFacetChecks(div, ips, true);
    // Also check any matching WAN suggestion rows
    div.querySelectorAll('.curateWanCheck').forEach((cb) => {
      let value = cb.dataset.value;
      try { value = decodeURIComponent(value); } catch { /* keep */ }
      if (ips.includes(value)) cb.checked = true;
    });
    const mode = div.querySelector('.curateModeSelect');
    if (mode) mode.value = 'exclude';
    log(`Client ${clientNumber}: added ${ips.length} pasted WAN IP(s) to curation selection.`);
  });
}

function getTenantTicketContent(div) {
  return (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || '';
}

function getTenantCompanyName(div, clientNumber) {
  const prior = tenantUiState.get(String(clientNumber)) || {};
  return (prior.organizationHint || div.dataset.organizationHint || div.dataset.exoOrganizationName || div.dataset.graphTenantName || '').trim();
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
      { key: 'IncludeExchangeItemAggregated', label: 'UAL: ExchangeItemAggregated (opt-in, high volume)' },
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
    IncludeExchangeItemAggregated: false,
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
    MessageTraceStatuses: [],
  };
}

const MESSAGE_TRACE_STATUSES = [
  { value: 'Delivered', label: 'Delivered', hint: 'Mailbox delivery (junk/quarantine can still show as Delivered)' },
  { value: 'Failed', label: 'Failed', hint: 'Not delivered' },
  { value: 'Pending', label: 'Pending' },
  { value: 'Expanded', label: 'Expanded', hint: 'Distribution-group expansion' },
  { value: 'Quarantined', label: 'Quarantined', hint: 'Last 10 days only' },
  { value: 'FilteredAsSpam', label: 'Filtered as spam', hint: 'Last 10 days only' },
  { value: 'GettingStatus', label: 'Getting status' },
];

const MT_STATUS_TEN_DAY = new Set(['Quarantined', 'FilteredAsSpam']);
const MT_STATUS_ISSUE_SET = ['Failed', 'Quarantined', 'FilteredAsSpam'];

let mtStatusDialogTarget = null;
let mtStatusDialogWired = false;

function reportExportCheckboxClass(key) {
  return `rs_${key}`;
}

function messageTraceStatusesNeedTenDayWindow(statuses) {
  return (statuses || []).some((s) => MT_STATUS_TEN_DAY.has(s));
}

function parseMessageTraceStatuses(raw) {
  if (!raw) return [];
  if (Array.isArray(raw)) return raw.map((s) => String(s || '').trim()).filter(Boolean);
  return String(raw).split(/[,;]/).map((s) => s.trim()).filter(Boolean);
}

function buildReportExportsPanelHtml(scope) {
  const html = [];
  for (const group of REPORT_EXPORT_GROUPS) {
    html.push(`<div class="reportExportGroup"><div class="reportExportGroupTitle">${group.title}</div><div class="reportExportGrid">`);
    for (const item of group.items) {
      if (item.key === 'IncludeMessageTrace') {
        html.push(`<div class="reportExportMt"><label><input type="checkbox" class="${reportExportCheckboxClass(item.key)}" data-rs-key="${item.key}" /> ${item.label}</label><button type="button" class="small mtStatusOpen">Statuses…</button></div>`);
      } else {
        html.push(`<label><input type="checkbox" class="${reportExportCheckboxClass(item.key)}" data-rs-key="${item.key}" /> ${item.label}</label>`);
      }
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
  if (container.dataset.ualRecordTypes) {
    try {
      const parsed = JSON.parse(container.dataset.ualRecordTypes);
      if (Array.isArray(parsed) && parsed.length) rs.UnifiedAuditLogRecordTypes = parsed;
    } catch (_) {}
  }
  if (container.dataset.messageTraceStatuses) {
    try {
      const parsed = JSON.parse(container.dataset.messageTraceStatuses);
      rs.MessageTraceStatuses = parseMessageTraceStatuses(parsed);
    } catch (_) {
      rs.MessageTraceStatuses = [];
    }
  } else {
    rs.MessageTraceStatuses = [];
  }
  if (container.dataset.exportPresetName) {
    rs.ExportPresetName = container.dataset.exportPresetName;
  }
  if (rs.IncludeExchangeItemAggregated && Array.isArray(rs.UnifiedAuditLogRecordTypes)
      && !rs.UnifiedAuditLogRecordTypes.includes('ExchangeItemAggregated')) {
    rs.UnifiedAuditLogRecordTypes = [...rs.UnifiedAuditLogRecordTypes, 'ExchangeItemAggregated'];
  }
  return rs;
}

function applyMessageTraceDays(days) {
  const mt = document.getElementById('messageTraceDays');
  if (!mt || days == null) return;
  const want = String(days);
  if (![...mt.options].some((o) => o.value === want)) {
    const opt = document.createElement('option');
    opt.value = want;
    opt.textContent = want;
    mt.appendChild(opt);
  }
  mt.value = want;
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
  if (Array.isArray(rs.UnifiedAuditLogRecordTypes) && rs.UnifiedAuditLogRecordTypes.length) {
    container.dataset.ualRecordTypes = JSON.stringify(rs.UnifiedAuditLogRecordTypes);
  } else {
    delete container.dataset.ualRecordTypes;
  }
  const mtStatuses = parseMessageTraceStatuses(rs.MessageTraceStatuses);
  if (mtStatuses.length) container.dataset.messageTraceStatuses = JSON.stringify(mtStatuses);
  else delete container.dataset.messageTraceStatuses;
  if (rs.ExportPresetName) container.dataset.exportPresetName = rs.ExportPresetName;
  else delete container.dataset.exportPresetName;
  if (rs.MessageTraceDaysBack != null) applyMessageTraceDays(rs.MessageTraceDaysBack);
  updateReportScopeHints(container, rs);
}

function formatUalScopeHint(selections) {
  if (!selections?.IncludeUnifiedAuditLogs) return 'UAL: off';
  const types = selections.UnifiedAuditLogRecordTypes;
  if (Array.isArray(types) && types.length) {
    return `UAL RecordTypes (${types.length}): ${types.join(', ')}`;
  }
  return 'UAL RecordTypes: default set (no Aggregated)';
}

function formatMessageTraceScopeHint(selections) {
  if (!selections?.IncludeMessageTrace) return 'Message trace: off';
  const st = parseMessageTraceStatuses(selections.MessageTraceStatuses);
  if (st.length) {
    const days = messageTraceStatusesNeedTenDayWindow(st) ? ' (clamped to 10 days)' : '';
    return `Message trace statuses: ${st.join(', ')}${days}`;
  }
  return 'Message trace statuses: all';
}

function updateUalScopeHint(container, selections) {
  updateReportScopeHints(container, selections);
}

function updateReportScopeHints(container, selections) {
  if (!container) return;
  const rs = selections || readReportSelectionsFromContainer(container);
  let ual = container.querySelector('.ualScopeHint');
  if (!ual) {
    ual = document.createElement('p');
    ual.className = 'ualScopeHint muted';
    container.appendChild(ual);
  }
  ual.textContent = formatUalScopeHint(rs);
  let mt = container.querySelector('.mtScopeHint');
  if (!mt) {
    mt = document.createElement('p');
    mt.className = 'mtScopeHint muted';
    container.appendChild(mt);
  }
  mt.textContent = formatMessageTraceScopeHint(rs);
}

function renderMessageTraceStatusList(filterText, selected) {
  const list = document.getElementById('mtStatusList');
  if (!list) return;
  const q = String(filterText || '').trim().toLowerCase();
  const chosen = new Set(selected || []);
  const html = [];
  for (const item of MESSAGE_TRACE_STATUSES) {
    const hay = `${item.value} ${item.label} ${item.hint || ''}`.toLowerCase();
    if (q && !hay.includes(q)) continue;
    const hint = item.hint ? `<span class="muted"> — ${item.hint}</span>` : '';
    html.push(`<label><input type="checkbox" data-mt-status="${item.value}" ${chosen.has(item.value) ? 'checked' : ''} /> ${item.label}${hint}</label>`);
  }
  list.innerHTML = html.join('') || '<p class="muted">No statuses match that search.</p>';
  updateMessageTraceStatusWarn();
}

function readMessageTraceStatusDialogSelection() {
  return [...document.querySelectorAll('#mtStatusList input[data-mt-status]:checked')].map((el) => el.dataset.mtStatus);
}

function updateMessageTraceStatusWarn() {
  const el = document.getElementById('mtStatusWarn');
  if (!el) return;
  const selected = readMessageTraceStatusDialogSelection();
  if (!selected.length) {
    el.textContent = 'No statuses selected — Generate pulls all delivery statuses.';
    return;
  }
  if (messageTraceStatusesNeedTenDayWindow(selected)) {
    el.textContent = 'Quarantined / Filtered as spam require a search shorter than 10 days. Generate will clamp the message-trace window.';
    return;
  }
  el.textContent = `Filter: ${selected.join(', ')}.`;
}

function applyMessageTraceStatusChip(chip) {
  let next = [];
  if (chip === 'all') next = [];
  else if (chip === 'issues') next = [...MT_STATUS_ISSUE_SET];
  else if (chip) next = [chip];
  renderMessageTraceStatusList(document.getElementById('mtStatusSearch')?.value, next);
}

function closeMessageTraceStatusDialog() {
  const dlg = document.getElementById('mtStatusDialog');
  if (dlg) dlg.hidden = true;
  mtStatusDialogTarget = null;
}

function openMessageTraceStatusDialog(container) {
  if (!container) return;
  ensureMessageTraceStatusDialog();
  mtStatusDialogTarget = container;
  const rs = readReportSelectionsFromContainer(container);
  const search = document.getElementById('mtStatusSearch');
  if (search) search.value = '';
  renderMessageTraceStatusList('', parseMessageTraceStatuses(rs.MessageTraceStatuses));
  const dlg = document.getElementById('mtStatusDialog');
  if (dlg) dlg.hidden = false;
  search?.focus();
}

function applyMessageTraceStatusDialog() {
  if (!mtStatusDialogTarget) return;
  const selected = readMessageTraceStatusDialogSelection();
  if (selected.length) mtStatusDialogTarget.dataset.messageTraceStatuses = JSON.stringify(selected);
  else delete mtStatusDialogTarget.dataset.messageTraceStatuses;
  const include = mtStatusDialogTarget.querySelector(`.${reportExportCheckboxClass('IncludeMessageTrace')}`);
  if (include && selected.length) include.checked = true;
  if (messageTraceStatusesNeedTenDayWindow(selected)) {
    const mt = document.getElementById('messageTraceDays');
    const current = parseInt(mt?.value, 10) || 7;
    if (current > 10) applyMessageTraceDays(10);
  }
  mtStatusDialogTarget.dataset.exportPresetName = 'Custom (manual selection)';
  const sessionPreset = document.getElementById('sessionReportPreset');
  if (sessionPreset && mtStatusDialogTarget.id === 'sessionReportExportsBody') {
    const custom = [...sessionPreset.options].find((o) => String(o.value).startsWith('Custom'));
    if (custom) sessionPreset.value = custom.value;
  }
  updateReportScopeHints(mtStatusDialogTarget);
  updateSessionReportExportsSummary(readReportSelectionsFromContainer(mtStatusDialogTarget));
  scheduleSessionReportSelectionsSync();
  const tenantDetails = mtStatusDialogTarget.closest('.tenant');
  const tenantBody = mtStatusDialogTarget.closest('.tenantBody');
  if (tenantDetails?.dataset.client && tenantBody) {
    saveTenantUiState(tenantDetails.dataset.client, tenantBody);
  }
  closeMessageTraceStatusDialog();
}

function ensureMessageTraceStatusDialog() {
  if (mtStatusDialogWired) return;
  const dlg = document.getElementById('mtStatusDialog');
  if (!dlg) return;
  mtStatusDialogWired = true;
  dlg.addEventListener('click', (e) => {
    if (e.target?.dataset?.mtDismiss) closeMessageTraceStatusDialog();
  });
  document.getElementById('mtStatusCancel')?.addEventListener('click', closeMessageTraceStatusDialog);
  document.getElementById('mtStatusApply')?.addEventListener('click', applyMessageTraceStatusDialog);
  document.getElementById('mtStatusSearch')?.addEventListener('input', (e) => {
    renderMessageTraceStatusList(e.target.value, readMessageTraceStatusDialogSelection());
  });
  dlg.querySelectorAll('[data-mt-chip]').forEach((btn) => {
    btn.addEventListener('click', () => applyMessageTraceStatusChip(btn.dataset.mtChip));
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && dlg && !dlg.hidden) closeMessageTraceStatusDialog();
  });
}

function wireMessageTraceStatusUi(container) {
  if (!container || container.dataset.mtStatusWired === '1') return;
  container.dataset.mtStatusWired = '1';
  container.querySelector('.mtStatusOpen')?.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    openMessageTraceStatusDialog(container);
  });
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
  const mode = ui?.reportExportMode || (ui?.useSessionReportDefaults === false ? 'custom' : 'session');
  if (mode !== 'session' && ui?.reportSelections) {
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

function updateTenantReportExportsHint(div, mode, presetName) {
  const hint = div?.querySelector('.reportExportsHint');
  if (!hint) return;
  const m = mode || 'session';
  if (m === 'session') {
    hint.textContent = '(session defaults)';
    hint.classList.remove('customized');
  } else if (m === 'preset') {
    hint.textContent = presetName ? `(preset: ${presetName})` : '(tenant preset)';
    hint.classList.add('customized');
  } else {
    hint.textContent = '(custom for this client)';
    hint.classList.add('customized');
  }
}

function getTenantReportExportMode(div) {
  const sel = div?.querySelector('.tenantReportExportMode');
  if (sel?.value) return sel.value;
  // Legacy: checkbox useSessionReportDefaults
  const legacy = div?.querySelector('.useSessionReportDefaults');
  if (legacy && !legacy.checked) return 'custom';
  return 'session';
}

function populateTenantPresetSelect(selectEl) {
  if (!selectEl) return;
  const previous = selectEl.value;
  selectEl.innerHTML = '';
  const presets = exportPresetsFromServer.length
    ? exportPresetsFromServer
    : Object.keys(REPORT_EXPORT_PRESETS).map((name) => ({ name, selections: REPORT_EXPORT_PRESETS[name] }));
  for (const p of presets) {
    if (!p?.name || String(p.name).startsWith('Custom')) continue;
    const opt = document.createElement('option');
    opt.value = p.name;
    opt.textContent = p.name;
    selectEl.appendChild(opt);
  }
  if (previous && [...selectEl.options].some((o) => o.value === previous)) {
    selectEl.value = previous;
  } else if (selectEl.options.length) {
    const bec = [...selectEl.options].find((o) => o.value === DEFAULT_BEC_PRESET_NAME);
    selectEl.value = bec ? bec.value : selectEl.options[0].value;
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
        daysBack: sessionRelativeDays(),
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
    document.querySelectorAll('.tenantReportPreset').forEach((sel) => populateTenantPresetSelect(sel));
  } catch {
    exportPresetsFromServer = [];
  }
}

function applyExportPresetByName(name, container) {
  const preset = exportPresetsFromServer.find(p => p.name === name)
    || (REPORT_EXPORT_PRESETS[name] ? { name, selections: REPORT_EXPORT_PRESETS[name] } : null);
  if (!preset || !preset.selections) return false;
  const merged = {
    ...defaultReportSelections(),
    ...preset.selections,
    ExportPresetName: name,
  };
  if (!Array.isArray(merged.UnifiedAuditLogRecordTypes) || !merged.UnifiedAuditLogRecordTypes.length) {
    // Fallback if server omitted UAL types (older runner)
    if (name.match(/BEC|Business Email/i)) {
      merged.UnifiedAuditLogRecordTypes = ['ExchangeItem', 'ExchangeItemGroup', 'ExchangeAdmin'];
    }
  }
  applyReportSelectionsToContainer(container, merged);
  return true;
}

function initSessionReportExportsPanel() {
  const body = document.getElementById('sessionReportExportsBody');
  if (!body || body.dataset.initialized === '1') return;
  body.innerHTML = buildReportExportsPanelHtml('session');
  body.dataset.initialized = '1';
  wireMessageTraceStatusUi(body);
  const markPresetCustom = () => {
    const sel = document.getElementById('sessionReportPreset');
    if (!sel) return;
    const custom = [...sel.options].find(o => String(o.value).startsWith('Custom'));
    if (custom) sel.value = custom.value;
  };
  body.addEventListener('change', (e) => {
    if (e.target?.closest?.('#mtStatusDialog')) return;
    markPresetCustom();
    delete body.dataset.ualRecordTypes;
    delete body.dataset.exportPresetName;
    body.dataset.exportPresetName = 'Custom (manual selection)';
    updateReportScopeHints(body);
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
    updateReportScopeHints(body);
    updateSessionReportExportsSummary(readReportSelectionsFromContainer(body));
    scheduleSessionReportSelectionsSync();
  });
  document.getElementById('btnReportSelectNone')?.addEventListener('click', () => {
    setAllReportSelections(body, false);
    markPresetCustom();
    updateReportScopeHints(body);
    updateSessionReportExportsSummary(readReportSelectionsFromContainer(body));
    scheduleSessionReportSelectionsSync();
  });
  ['messageTraceDays', 'signInLogsDays', 'relAmount', 'relUnit'].forEach(id => {
    document.getElementById(id)?.addEventListener('change', () => {
      updateSessionTimeframeUi();
      scheduleSessionReportSelectionsSync();
    });
  });
  updateSessionTimeframeUi();
}

function wireTenantTimeframePanel(div, clientNumber) {
  ['.relAmount', '.relUnit'].forEach((sel) => {
    div.querySelector(sel)?.addEventListener('change', () => {
      updateTenantTimeframeUi(div);
      saveTenantUiState(clientNumber, div);
    });
  });
  updateTenantTimeframeUi(div);
}

function wireTenantReportExportsPanel(div, clientNumber) {
  const modeSel = div.querySelector('.tenantReportExportMode');
  const presetWrap = div.querySelector('.tenantReportPresetWrap');
  const presetSel = div.querySelector('.tenantReportPreset');
  const customBody = div.querySelector('.tenantReportExportsCustom');
  const legacyCheck = div.querySelector('.useSessionReportDefaults');

  populateTenantPresetSelect(presetSel);
  if (customBody) wireMessageTraceStatusUi(customBody);

  const applyVisibility = (skipSave = false) => {
    const mode = getTenantReportExportMode(div);
    if (legacyCheck) legacyCheck.checked = mode === 'session';
    if (presetWrap) presetWrap.style.display = mode === 'preset' ? '' : 'none';
    if (customBody) customBody.style.display = mode === 'custom' ? 'block' : 'none';
    if (mode === 'preset' && presetSel && customBody) {
      applyExportPresetByName(presetSel.value, customBody);
    }
    updateTenantReportExportsHint(div, mode, presetSel?.value);
    if (!skipSave) saveTenantUiState(clientNumber, div);
  };

  modeSel?.addEventListener('change', () => {
    const mode = modeSel.value;
    if (mode === 'custom' && customBody) {
      const saved = tenantUiState.get(String(clientNumber));
      if (!saved?.reportSelections || saved.reportExportMode === 'session') {
        applyReportSelectionsToContainer(
          customBody,
          readReportSelectionsFromContainer(document.getElementById('sessionReportExportsBody'))
        );
      } else if (saved.reportSelections) {
        applyReportSelectionsToContainer(customBody, saved.reportSelections);
      }
    }
    applyVisibility(false);
  });
  presetSel?.addEventListener('change', () => {
    if (getTenantReportExportMode(div) === 'preset' && customBody && presetSel.value) {
      applyExportPresetByName(presetSel.value, customBody);
      updateTenantReportExportsHint(div, 'preset', presetSel.value);
      saveTenantUiState(clientNumber, div);
    }
  });
  customBody?.addEventListener('change', () => {
    delete customBody.dataset.ualRecordTypes;
    customBody.dataset.exportPresetName = 'Custom (manual selection)';
    updateReportScopeHints(customBody);
    saveTenantUiState(clientNumber, div);
  });
  if (customBody && !customBody.querySelector('.tenantReportSelectAll')) {
    const bar = document.createElement('div');
    bar.className = 'row';
    bar.innerHTML = '<button type="button" class="small tenantReportSelectAll">Select all</button><button type="button" class="small tenantReportSelectNone">Deselect all</button>';
    customBody.prepend(bar);
    bar.querySelector('.tenantReportSelectAll')?.addEventListener('click', () => {
      setAllReportSelections(customBody, true);
      delete customBody.dataset.ualRecordTypes;
      customBody.dataset.exportPresetName = 'Custom (manual selection)';
      updateReportScopeHints(customBody);
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
  }, 60000);
}

async function waitForGraphAppJob(started) {
  if (!started?.started) return started;
  log(`Graph app console opened (PID ${started.processId}). Finish sign-in there — other runner actions stay available.`);
  const deadline = Date.now() + 10 * 60 * 1000;
  while (Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 2000));
    const st = await api('/api/wcm/graph-app-job', {}, 10000);
    if (!st?.running) return st;
  }
  throw new Error('Graph app console is still open after 10 minutes. Finish or close that window, then retry.');
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
    const actions = [];
    if (row.hasOutputs) {
      actions.push(`<button type="button" class="small investigateSession" data-session-id="${sid}">Investigate</button>`);
    }
    actions.push(`<button type="button" class="small restoreSession" data-session-id="${sid}">Restore</button>`);

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
  listEl.querySelectorAll('.investigateSession').forEach(btn => {
    btn.addEventListener('click', () => {
      const sid = btn.dataset.sessionId;
      if (sid) window.location.href = `investigate.html?session=${encodeURIComponent(sid)}`;
    });
  });
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

  const days = sessionRelativeDays();

  const d = new Date();

  d.setDate(d.getDate() - days);

  return toDateTimeLocalValue(d);

}

function defaultDateEndValue() {

  return toDateTimeLocalValue(new Date());

}

const RELATIVE_MAX_DAYS = 90;

const RELATIVE_UNITS = [
  ['days', 'Days'],
  ['weeks', 'Weeks'],
  ['months', 'Months'],
  ['max', `Max (${RELATIVE_MAX_DAYS} days)`],
];

function relativeUnitOptionsHtml(selectedUnit, includeCustom) {
  const units = includeCustom ? [...RELATIVE_UNITS, ['custom', 'Custom range']] : RELATIVE_UNITS;
  return units
    .map(([value, label]) => `<option value="${value}"${value === selectedUnit ? ' selected' : ''}>${label}</option>`)
    .join('');
}

function resolveRelativeWindow(amount, unit) {
  const end = new Date();
  const start = new Date(end.getTime());
  const n = Math.max(1, parseInt(amount, 10) || 1);
  if (unit === 'weeks') start.setDate(start.getDate() - (n * 7));
  else if (unit === 'months') start.setMonth(start.getMonth() - n);
  else if (unit === 'max') start.setDate(start.getDate() - RELATIVE_MAX_DAYS);
  else start.setDate(start.getDate() - n);

  // Message trace and Graph both stop at ~90 days, so a longer window buys nothing.
  const earliest = new Date(end.getTime());
  earliest.setDate(earliest.getDate() - RELATIVE_MAX_DAYS);
  const capped = start < earliest;
  return { start: capped ? earliest : start, end, capped };
}

function relativeWindowDays(amount, unit) {
  const { start, end } = resolveRelativeWindow(amount, unit);
  return Math.max(1, Math.round((end - start) / 86400000));
}

function relativeWindowHint(amount, unit) {
  if (unit === 'custom') return 'Using the start and end dates below.';
  const { start, end, capped } = resolveRelativeWindow(amount, unit);
  const d = (x) => `${x.getFullYear()}-${pad2(x.getMonth() + 1)}-${pad2(x.getDate())}`;
  return `${d(start)} → ${d(end)}${capped ? ` (capped at ${RELATIVE_MAX_DAYS} days)` : ''}`;
}

function sessionRelativeUnit() {
  return document.getElementById('relUnit')?.value || 'days';
}

function sessionRelativeAmount() {
  return Math.max(1, parseInt(document.getElementById('relAmount')?.value, 10) || 7);
}

function sessionRelativeDays() {
  return relativeWindowDays(sessionRelativeAmount(), sessionRelativeUnit());
}

function updateSessionTimeframeUi() {
  const unit = sessionRelativeUnit();
  const amountWrap = document.getElementById('relAmountWrap');
  const hint = document.getElementById('relHint');
  if (amountWrap) amountWrap.style.display = unit === 'max' ? 'none' : '';
  if (hint) hint.textContent = relativeWindowHint(sessionRelativeAmount(), unit);
}

function getTenantRelativeUnit(div) {
  return div?.querySelector('.relUnit')?.value || 'days';
}

function getTenantRelativeAmount(div) {
  return Math.max(1, parseInt(div?.querySelector('.relAmount')?.value, 10) || 7);
}

function updateTenantTimeframeUi(div) {
  if (!div) return;
  const unit = getTenantRelativeUnit(div);
  const amountWrap = div.querySelector('.relAmountWrap');
  const customRow = div.querySelector('.customRangeRow');
  const hint = div.querySelector('.relHint');
  if (amountWrap) amountWrap.style.display = (unit === 'max' || unit === 'custom') ? 'none' : '';
  if (customRow) customRow.style.display = unit === 'custom' ? '' : 'none';
  if (hint) hint.textContent = relativeWindowHint(getTenantRelativeAmount(div), unit);
}

// Resolved when Generate runs so a relative window always counts back from the click,
// not from whenever the tenant card was rendered.
function applyTenantRelativeWindow(div) {
  if (!div) return;
  const startEl = div.querySelector('.dateStart');
  const endEl = div.querySelector('.dateEnd');
  if (getTenantRelativeUnit(div) === 'custom') {
    if (startEl && !startEl.value) startEl.value = defaultDateStartValue();
    if (endEl && !endEl.value) endEl.value = defaultDateEndValue();
    return;
  }
  const { start, end } = resolveRelativeWindow(getTenantRelativeAmount(div), getTenantRelativeUnit(div));
  if (startEl) startEl.value = toDateTimeLocalValue(start);
  if (endEl) endEl.value = toDateTimeLocalValue(end);
}

function parseSearchTerms(text) {

  return (text || '')

    .split(',')

    .map(s => s.trim())

    .filter(Boolean);

}

function saveTenantUiState(clientNumber, div) {

  if (!div) return;

  const reportExportMode = getTenantReportExportMode(div);
  const useSessionReportDefaults = reportExportMode === 'session';
  const customPanel = div.querySelector('.tenantReportExportsCustom');
  const presetSel = div.querySelector('.tenantReportPreset');
  let reportSelections = null;
  if (reportExportMode === 'preset' && customPanel && presetSel?.value) {
    applyExportPresetByName(presetSel.value, customPanel);
    reportSelections = readReportSelectionsFromContainer(customPanel);
  } else if (reportExportMode === 'custom' && customPanel) {
    reportSelections = readReportSelectionsFromContainer(customPanel);
  }

  const prior = tenantUiState.get(String(clientNumber)) || {};

  let validatedUsers = [];
  try { validatedUsers = JSON.parse(div.dataset.validatedUsers || '[]'); } catch { validatedUsers = []; }
  validatedUsers = normalizeUserList(validatedUsers);
  if (!validatedUsers.length && Array.isArray(prior.validatedUsers) && prior.validatedUsers.length) {
    validatedUsers = normalizeUserList(prior.validatedUsers);
  }
  if (validatedUsers.length) div.dataset.validatedUsers = JSON.stringify(validatedUsers);

  tenantUiState.set(String(clientNumber), {

    ...prior,

    ticket: (div.querySelector('.ticketInput')?.value || '').trim(),

    ticketContent: (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || prior.ticketContent || '',

    organizationHint: prior.organizationHint || '',

    tenantId: div.querySelector('.appRegSelect')?.value || '',

    appRegManual: isAppRegManual(div),

    interactive: Boolean(div.querySelector('.interactiveCheck')?.checked),

    filterUsers: Boolean(div.querySelector('.filterUsersCheck')?.checked),

    userSearch: div.querySelector('.userSearchInput')?.value || '',

    validatedUsers,

    dateStart: div.querySelector('.dateStart')?.value || '',

    dateEnd: div.querySelector('.dateEnd')?.value || '',

    relAmount: getTenantRelativeAmount(div),

    relUnit: getTenantRelativeUnit(div),

    useSessionReportDefaults,

    reportExportMode,

    exportPresetName: reportExportMode === 'preset' ? (presetSel?.value || '') : (reportSelections?.ExportPresetName || ''),

    reportSelections,

    security: {
      resolveJson: div.dataset.securityResolveJson || '',
      s1Profile: div.querySelector('.s1ProfileSelect')?.value || 'connectwise',
      s1SiteId: div.querySelector('.s1SiteIdInput')?.value || '',
      includeLiongardContext: Boolean(div.querySelector('.includeLiongardContext')?.checked),
      huntressPulls: Object.fromEntries([...div.querySelectorAll('.huntressPull')].map((cb) => [cb.dataset.pull, cb.checked])),
      s1Pulls: Object.fromEntries([...div.querySelectorAll('.s1Pull')].map((cb) => [cb.dataset.pull, cb.checked])),
    },

    auditGap: readAuditGap(div) || prior.auditGap || null,

    becReview: div.querySelector('.becReviewText')?.textContent || prior.becReview || '',

  });

  scheduleTenantUiStateSync(clientNumber, div);

}

function restoreTenantUiState(clientNumber, div) {

  let saved = tenantUiState.get(String(clientNumber));

  if (!saved || !div) return;

  const storedContainment = resolveContainmentState(clientNumber, saved);
  if (storedContainment && saved.containment !== storedContainment) {
    saved = { ...saved, containment: storedContainment };
    tenantUiState.set(String(clientNumber), saved);
  }

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

  setAppRegManual(div, Boolean(saved.appRegManual && saved.tenantId));
  if (appReg && isAppRegManual(div)) appReg.value = saved.tenantId;

  if (interactive) interactive.checked = saved.interactive;

  if (filterUsers) filterUsers.checked = saved.filterUsers;

  if (userSearch) userSearch.value = saved.userSearch;

  if (dateStart && saved.dateStart) dateStart.value = saved.dateStart;

  if (dateEnd && saved.dateEnd) dateEnd.value = saved.dateEnd;

  const relAmount = div.querySelector('.relAmount');

  const relUnit = div.querySelector('.relUnit');

  if (relAmount && saved.relAmount != null) relAmount.value = String(saved.relAmount);

  if (relUnit && saved.relUnit) relUnit.value = saved.relUnit;

  restoreSecurityUiState(div, saved);

  const restoredUsers = normalizeUserList(saved.validatedUsers);
  if (restoredUsers.length) {

    div.dataset.validatedUsers = JSON.stringify(restoredUsers);

    updateValidatedUsersDisplay(div, restoredUsers);

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

  if (saved.becReview) {
    const reviewPre = div.querySelector('.becReviewText');
    const reviewWrap = div.querySelector('.becReviewWrap');
    if (reviewPre && !reviewPre.textContent) reviewPre.textContent = saved.becReview;
    if (reviewWrap && reviewPre && reviewPre.textContent) reviewWrap.style.display = 'block';
  }

  if (saved.auditGap && typeof saved.auditGap === 'object') {
    div.dataset.auditGap = JSON.stringify(saved.auditGap);
  }

  const modeSel = div.querySelector('.tenantReportExportMode');
  const presetWrap = div.querySelector('.tenantReportPresetWrap');
  const presetSel = div.querySelector('.tenantReportPreset');
  const useDefaultsCheck = div.querySelector('.useSessionReportDefaults');
  const customPanel = div.querySelector('.tenantReportExportsCustom');
  let mode = saved.reportExportMode;
  if (!mode) mode = saved.useSessionReportDefaults === false ? 'custom' : 'session';
  populateTenantPresetSelect(presetSel);
  if (modeSel) modeSel.value = mode;
  if (useDefaultsCheck) useDefaultsCheck.checked = mode === 'session';
  if (presetSel && saved.exportPresetName) {
    if ([...presetSel.options].some((o) => o.value === saved.exportPresetName)) {
      presetSel.value = saved.exportPresetName;
    }
  }
  if (customPanel && saved.reportSelections) {
    applyReportSelectionsToContainer(customPanel, saved.reportSelections);
  } else if (mode === 'preset' && customPanel && presetSel?.value) {
    applyExportPresetByName(presetSel.value, customPanel);
  }
  if (presetWrap) presetWrap.style.display = mode === 'preset' ? '' : 'none';
  if (customPanel) customPanel.style.display = mode === 'custom' ? 'block' : 'none';
  updateTenantReportExportsHint(div, mode, presetSel?.value);

  refreshContainmentUsers(div);
  updateContainmentButtons(div);
  restoreContainmentOutput(div, saved, clientNumber);
  updateCaCatalogButtons(div);
  restoreCaCatalogOutput(div, saved);

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

  refreshContainmentUsers(div);

}

function normalizeUserList(value) {
  if (value == null || value === '') return [];
  const raw = Array.isArray(value) ? value : [value];
  return raw.map((u) => {
    if (u == null || u === '') return '';
    if (typeof u === 'string') return u.trim();
    return String(u.UserPrincipalName || u.userPrincipalName || u || '').trim();
  }).filter(Boolean);
}

function refreshContainmentUsers(div) {
  const list = div?.querySelector('.containmentUserList');
  if (!list) return;
  const clientNumber = div.closest('details.tenant')?.dataset?.client;
  let users = clientNumber
    ? getValidatedUsersForTenant(clientNumber, div)
    : normalizeUserList((() => { try { return JSON.parse(div.dataset.validatedUsers || '[]'); } catch { return []; } })());
  if (!users.length) {
    list.innerHTML = 'Validate users first. Containment acts only on those UPNs.';
    list.classList.add('muted');
    delete div.dataset.restrictedEmailJson;
    updateContainmentButtons(div);
    return;
  }
  list.classList.remove('muted');
  list.innerHTML = users.map((u) => {
    const safe = escapeHtml(String(u));
    return `<label><input type="checkbox" class="containmentUser" value="${safe}" checked /> ${safe}</label>`;
  }).join('');
  if (users.length === 1) fillThreatMailboxIfEmpty(div, users[0]);
  updateContainmentButtons(div);
}

function getSelectedContainmentUsers(div) {
  return [...(div?.querySelectorAll('.containmentUser:checked') || [])].map((cb) => cb.value).filter(Boolean);
}

// Commercial SSPR portal — the page that is actually in place. aka.ms/sspr only redirects here.
const CONTAINMENT_PASSWORD_RESET_URL = 'https://passwordreset.microsoftonline.com/';

function containmentPasswordResetUrl(preferred, users) {
  const fromServer = String(preferred || '').trim();
  if (/^https:\/\/passwordreset\.microsoftonline\.(com|us)\/?/i.test(fromServer)) return fromServer;
  const upn = Array.isArray(users) ? users[0] : users;
  const domain = String(upn || '').split('@')[1];
  if (domain) return `${CONTAINMENT_PASSWORD_RESET_URL}?whr=${encodeURIComponent(domain)}`;
  return CONTAINMENT_PASSWORD_RESET_URL;
}

function containmentPasswordResetInstruction(url) {
  return `Do not send a password. Send the user ${url}`;
}

function getRestrictedContainmentHits(div, selectedUsers) {
  let rows = [];
  try { rows = JSON.parse(div?.dataset?.restrictedEmailJson || '[]'); } catch { rows = []; }
  const selected = new Set((selectedUsers || []).map((u) => String(u).toLowerCase()));
  return rows.filter((r) => r.Restricted && selected.has(String(r.UserPrincipalName || '').toLowerCase()));
}

function applyContainmentCapabilities(div, caps) {
  if (!div || !caps) return;
  if (caps.canRevoke === false) div.dataset.graphCanRevoke = '0';
  else if (caps.canRevoke === true) div.dataset.graphCanRevoke = '1';
  if (caps.canBlock === false) div.dataset.graphCanBlock = '0';
  else if (caps.canBlock === true) div.dataset.graphCanBlock = '1';
  if (caps.canAuthWrite === false) div.dataset.graphCanAuthWrite = '0';
  else if (caps.canAuthWrite === true) div.dataset.graphCanAuthWrite = '1';
  if (caps.canDeviceDelete === false) div.dataset.graphCanDeviceDelete = '0';
  else if (caps.canDeviceDelete === true) div.dataset.graphCanDeviceDelete = '1';
  if (caps.canAppWrite === false) div.dataset.graphCanAppWrite = '0';
  else if (caps.canAppWrite === true) div.dataset.graphCanAppWrite = '1';
  if (caps.canPasswordReset === false) div.dataset.graphCanPasswordReset = '0';
  else if (caps.canPasswordReset === true) div.dataset.graphCanPasswordReset = '1';
  if (caps.canOauthWrite === false) div.dataset.graphCanOauthWrite = '0';
  else if (caps.canOauthWrite === true) div.dataset.graphCanOauthWrite = '1';
  if (caps.canIntune === false) div.dataset.graphCanIntune = '0';
  else if (caps.canIntune === true) div.dataset.graphCanIntune = '1';
  if (caps.canIntuneWipe === false) div.dataset.graphCanIntuneWipe = '0';
  else if (caps.canIntuneWipe === true) div.dataset.graphCanIntuneWipe = '1';
  if (caps.canRoles === false) div.dataset.graphCanRoles = '0';
  else if (caps.canRoles === true) div.dataset.graphCanRoles = '1';
  if (caps.canGroupWrite === false) div.dataset.graphCanGroupWrite = '0';
  else if (caps.canGroupWrite === true) div.dataset.graphCanGroupWrite = '1';
  if (caps.reason) div.dataset.graphCapabilityReason = caps.reason;
  rememberContainment(div, { capabilities: caps });
  const details = div.closest('details.tenant');
  const clientNumber = details?.dataset?.client;
  if (clientNumber) {
    const st = tenantUiState.get(String(clientNumber)) || {};
    st.graphCapabilities = caps;
    tenantUiState.set(String(clientNumber), st);
  }
  const hint = div.querySelector('.containmentGraphHint');
  if (hint) {
    hint.textContent = caps.reason || '';
    hint.style.display = caps.reason ? 'block' : 'none';
  }
  const updateScopesBtn = div.querySelector('.containmentUpdateGraphScopes');
  if (updateScopesBtn) updateScopesBtn.style.display = caps.reason ? '' : 'none';
  updateContainmentButtons(div);
}

function updateContainmentButtons(div) {
  if (!div) return;
  const graph = div.dataset.graphAuthenticated === '1';
  const exo = div.dataset.exchangeAuthenticated === '1';
  const users = getSelectedContainmentUsers(div);
  const hasUsers = users.length > 0;
  const canRevoke = graph && hasUsers && div.dataset.graphCanRevoke !== '0';
  const canBlock = graph && hasUsers && div.dataset.graphCanBlock !== '0';
  const set = (sel, on) => {
    const el = div.querySelector(sel);
    if (el) el.disabled = !on;
  };
  set('.containmentQuickLock', canBlock && canRevoke);
  set('.containmentSigninStatus', graph && hasUsers);
  set('.containmentRevoke', canRevoke);
  set('.containmentRevokeMfa', canRevoke);
  set('.containmentBlock', canBlock);
  set('.containmentUnblock', canBlock);
  const canPassword = graph && hasUsers && div.dataset.graphCanPasswordReset !== '0';
  set('.containmentQuickRestore', canPassword && canBlock);
  set('.containmentResetPassword', canPassword);
  set('.containmentAssignPasswordBtn', canPassword);
  set('.containmentListMfa', graph && hasUsers);
  const hasSelectedMfa = [...(div.querySelectorAll('.containmentMfaPick:checked') || [])].length > 0;
  set('.containmentDeleteMfa', graph && hasSelectedMfa && div.dataset.graphCanAuthWrite !== '0');
  set('.containmentListDevices', graph && hasUsers);
  const hasSelectedDevices = [...(div.querySelectorAll('.containmentDevicePick:checked') || [])].length > 0;
  set('.containmentDeleteDevices', graph && hasSelectedDevices && div.dataset.graphCanDeviceDelete !== '0');
  set('.containmentListApps', graph);
  const hasSelectedApps = [...(div.querySelectorAll('.containmentAppPick:checked') || [])].length > 0;
  set('.containmentDeleteApps', graph && hasSelectedApps && div.dataset.graphCanAppWrite !== '0');
  set('.containmentRestrictedStatus', exo && hasUsers);
  const restrictedHits = getRestrictedContainmentHits(div, users);
  const unrestrictBtn = div.querySelector('.containmentUnrestrict');
  if (unrestrictBtn) {
    unrestrictBtn.style.display = restrictedHits.length ? '' : 'none';
    unrestrictBtn.disabled = !(exo && restrictedHits.length);
  }
  set('.containmentListRules', exo && hasUsers);
  set('.containmentBecReview', (exo || graph) && users.length === 1);
  set('.containmentBecCheck', (exo || graph) && users.length === 1);
  const becRun = div.querySelector('.becInvestigateRun');
  if (becRun) becRun.disabled = !(exo || graph);
  const hasSelectedRules = [...(div.querySelectorAll('.containmentRulePick:checked') || [])].length > 0;
  const rulesMailbox = div.dataset.containmentRulesUser || (users.length === 1 ? users[0] : '');
  set('.containmentDeleteRules', exo && Boolean(rulesMailbox) && hasSelectedRules);
  set('.containmentMailboxStatus', exo && hasUsers);
  set('.containmentSetForward', exo && hasUsers);
  set('.containmentClearForward', exo && hasUsers);
  set('.containmentAddDelegate', exo && hasUsers);
  const hasSelectedForwards = [...(div.querySelectorAll('.containmentForwardPick:checked') || [])].length > 0;
  set('.containmentRemoveForward', exo && hasSelectedForwards);
  const hasSelectedDelegates = [...(div.querySelectorAll('.containmentDelegatePick:checked') || [])].length > 0;
  set('.containmentRemoveDelegate', exo && hasSelectedDelegates);
  set('.containmentListTransport', exo);
  set('.containmentListConnectors', exo);
  const hasSelectedTransport = [...(div.querySelectorAll('.containmentTransportPick:checked') || [])].length > 0;
  set('.containmentDeleteTransport', exo && hasSelectedTransport);
  const hasSelectedConnectors = [...(div.querySelectorAll('.containmentConnectorPick:checked') || [])].length > 0;
  set('.containmentDeleteConnectors', exo && hasSelectedConnectors);
  set('.containmentReregisterMfa', graph && hasUsers && div.dataset.graphCanAuthWrite !== '0');
  set('.containmentListOauth', graph && hasUsers);
  const hasSelectedOauth = [...(div.querySelectorAll('.containmentOauthPick:checked') || [])].length > 0;
  set('.containmentDeleteOauth', graph && hasSelectedOauth && div.dataset.graphCanOauthWrite !== '0');
  set('.containmentListMobile', exo && hasUsers);
  const hasSelectedMobile = [...(div.querySelectorAll('.containmentMobilePick:checked') || [])].length > 0;
  set('.containmentDeleteMobile', exo && hasSelectedMobile);
  set('.containmentListIntune', graph && hasUsers && div.dataset.graphCanIntune !== '0');
  const hasSelectedIntune = [...(div.querySelectorAll('.containmentIntunePick:checked') || [])].length > 0;
  set('.containmentRetireIntune', graph && hasSelectedIntune && div.dataset.graphCanIntuneWipe !== '0');
  set('.containmentWipeIntune', graph && hasSelectedIntune && div.dataset.graphCanIntuneWipe !== '0');
  set('.containmentListFolders', exo && hasUsers);
  const hasSelectedFolders = [...(div.querySelectorAll('.containmentFolderPick:checked') || [])].length > 0;
  set('.containmentDeleteFolders', exo && hasSelectedFolders);
  set('.containmentAutoreplyStatus', exo && hasUsers);
  set('.containmentDisableAutoreply', exo && hasUsers);
  set('.containmentListJunk', exo && hasUsers);
  const hasSelectedJunk = [...(div.querySelectorAll('.containmentJunkPick:checked') || [])].length > 0;
  set('.containmentDeleteJunk', exo && hasSelectedJunk);
  set('.containmentListElsewhere', exo && hasUsers);
  const hasSelectedElsewhere = [...(div.querySelectorAll('.containmentElsewherePick:checked') || [])].length > 0;
  set('.containmentDeleteElsewhere', exo && hasSelectedElsewhere);
  set('.containmentHoldStatus', exo && hasUsers);
  set('.containmentEnableHold', exo && hasUsers);
  set('.containmentListOrgfwd', exo);
  const hasSelectedOrgfwd = [...(div.querySelectorAll('.containmentOrgfwdPick:checked') || [])].length > 0;
  set('.containmentDisableOrgfwd', exo && hasSelectedOrgfwd);
  set('.containmentListJournal', exo);
  const hasSelectedJournal = [...(div.querySelectorAll('.containmentJournalPick:checked') || [])].length > 0;
  set('.containmentDeleteJournal', exo && hasSelectedJournal);
  set('.containmentListRoles', graph && hasUsers);
  const hasSelectedRoles = [...(div.querySelectorAll('.containmentRolePick:checked') || [])].length > 0;
  set('.containmentDeleteRoles', graph && hasSelectedRoles && (div.dataset.graphCanRoles !== '0' || div.dataset.graphCanGroupWrite !== '0'));
  set('.containmentListAppcreds', graph);
  const hasSelectedAppcreds = [...(div.querySelectorAll('.containmentAppcredPick:checked') || [])].length > 0;
  set('.containmentDeleteAppcreds', graph && hasSelectedAppcreds && div.dataset.graphCanAppWrite !== '0');
  set('.containmentListFlows', true);
  const hasSelectedFlows = [...(div.querySelectorAll('.containmentFlowPick:checked') || [])].length > 0;
  set('.containmentDeleteFlows', hasSelectedFlows);
}

let containmentBatchConfirmed = false;
function confirmContainmentPopup(message) {
  if (containmentBatchConfirmed) return true;
  return window.confirm(message);
}

const CONTAINMENT_SLOW_WARNINGS = {
  'list-elsewhere': 'List rights elsewhere checks Send As and Send on Behalf quickly, then scans every mailbox for Full Access. That can take several minutes and the worker stays busy until it finishes.\n\nContinue?',
  'list-apps': 'Listing app registrations walks the whole tenant and can take a minute or more on large directories. The worker stays busy until it finishes.\n\nContinue?',
  'list-appcreds': 'Listing secrets and owners walks every app registration and can take several minutes. The worker stays busy until it finishes.\n\nContinue?',
  'list-roles': 'Listing directory roles, group memberships, and Exchange RBAC can take a minute or more. The worker stays busy until it finishes.\n\nContinue?',
  'list-folders': 'Listing folder permissions walks each folder on the selected mailbox(es) and can take a minute or more.\n\nContinue?',
  'list-flows': 'Listing Power Automate flows is tenant-wide and can take a minute if the admin module is loaded.\n\nContinue?',
  'list-transport': 'Listing transport rules is tenant-wide. Large tenants can take a minute.\n\nContinue?',
};

function confirmSlowContainment(kind) {
  const msg = CONTAINMENT_SLOW_WARNINGS[kind];
  if (!msg) return true;
  return confirmContainmentPopup(msg);
}

function classifyContainmentStatusTone(text) {
  const t = String(text || '');
  if (!t) return '';
  if (/[.…]$/.test(t) || /^Busy —/i.test(t) || /^(Checking|Listing|Resetting|Setting|Removing|Deleting|Clearing|Adding|Blocking|Unblocking|Revoking|Quick actions:)/i.test(t)) {
    return 'statusBusy';
  }
  const blocked = /SIGN-IN BLOCKED|accountEnabled=False|: RESTRICTED/i.test(t);
  const allowed = /SIGN-IN ALLOWED|accountEnabled=True|NOT RESTRICTED|not restricted/i.test(t);
  if (/FAILED|failed\.|Error:/i.test(t) && !blocked && !allowed) return 'statusBad';
  if (blocked && allowed) return 'statusWarn';
  if (blocked) return 'statusBad';
  if (allowed || /Finished\. Success|password reset/i.test(t)) return 'statusOk';
  if (/FAILED|failed\./i.test(t)) return 'statusBad';
  return 'statusWarn';
}

function formatContainmentUserStatus(users) {
  return (users || []).map((u) => {
    const upn = u.UserPrincipalName || u.userPrincipalName || '';
    if (u.Error) return `${upn}: ERROR — ${u.Error}`;
    if (u.Restricted === true) {
      return `${upn}: RESTRICTED${u.Reason ? ` (${u.Reason})` : ''}${u.CreatedDateTime ? ` since ${u.CreatedDateTime}` : ''}`;
    }
    if (u.Restricted === false) return `${upn}: NOT RESTRICTED (not on Restricted entities)`;
    if (typeof u.AccountEnabled === 'boolean') {
      const label = u.AccountEnabled ? 'SIGN-IN ALLOWED' : 'SIGN-IN BLOCKED';
      return `${upn}: ${label}${u.DisplayName ? ` (${u.DisplayName})` : ''}`;
    }
    return JSON.stringify(u);
  }).join('\n');
}

function formatContainmentUserStatusHtml(users) {
  return (users || []).map((u) => {
    const upn = escapeHtml(String(u.UserPrincipalName || u.userPrincipalName || ''));
    if (u.Error) {
      return `<div class="signinChip error"><span class="chipLabel">ERROR</span><span class="chipUpn">${upn}</span><span>${escapeHtml(String(u.Error))}</span></div>`;
    }
    if (u.Restricted === true) {
      const extra = [u.Reason, u.CreatedDateTime].filter(Boolean).map((x) => escapeHtml(String(x))).join(' · ');
      return `<div class="signinChip restricted"><span class="chipLabel">RESTRICTED</span><span class="chipUpn">${upn}</span>${extra ? `<span>${extra}</span>` : ''}</div>`;
    }
    if (u.Restricted === false) {
      return `<div class="signinChip ok"><span class="chipLabel">NOT RESTRICTED</span><span class="chipUpn">${upn}</span><span>not on Restricted entities</span></div>`;
    }
    if (typeof u.AccountEnabled === 'boolean') {
      const blocked = u.AccountEnabled === false;
      const name = u.DisplayName ? `<span>${escapeHtml(String(u.DisplayName))}</span>` : '';
      return `<div class="signinChip ${blocked ? 'blocked' : 'allowed'}"><span class="chipLabel">${blocked ? 'SIGN-IN BLOCKED' : 'SIGN-IN ALLOWED'}</span><span class="chipUpn">${upn}</span>${name}</div>`;
    }
    return `<div class="signinChip"><span class="chipUpn">${upn}</span><span>${escapeHtml(JSON.stringify(u))}</span></div>`;
  }).join('');
}

function setContainmentStatus(div, text, opts = {}) {
  const el = div.querySelector('.containmentStatus');
  if (el) {
    el.classList.remove('muted', 'hasStatus', 'statusOk', 'statusWarn', 'statusBad', 'statusBusy');
    // Restoring a previous tenant's status must not announce it. A live region
    // inside <details> makes Chromium open that panel and scroll it into view.
    if (opts.quietLive) el.setAttribute('aria-live', 'off');
    if (opts.html) el.innerHTML = opts.html;
    else el.textContent = text || '';
    if (text) {
      el.classList.add('hasStatus');
      const tone = opts.tone || classifyContainmentStatusTone(text);
      if (tone) el.classList.add(tone);
      // Only jump to this banner when it is the result surface (sign-in chips,
      // password/SSPR text, failures). List/delete results live in the table
      // next to the button — do not yank the viewport back to the top.
      if (opts.reveal && !opts.quietLive) {
        try { el.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } catch { /* ignore */ }
      }
    }
    if (opts.quietLive) {
      setTimeout(() => {
        if (el.getAttribute('aria-live') === 'off') el.setAttribute('aria-live', 'polite');
      }, 200);
    }
  }
  if (!opts.skipSave && text && !/[.…]$/.test(text) && !/^Busy —/i.test(text)) {
    rememberContainment(div, { status: text });
  }
}

function setContainmentUserStatus(div, users) {
  setContainmentStatus(div, formatContainmentUserStatus(users), { html: formatContainmentUserStatusHtml(users), reveal: true });
}

function rememberContainment(div, patch, immediate = false) {
  const clientNumber = div?.closest('details.tenant')?.dataset?.client;
  if (!clientNumber || !patch) return;
  const prior = tenantUiState.get(String(clientNumber)) || {};
  const containment = { ...(prior.containment || {}), ...patch };
  tenantUiState.set(String(clientNumber), { ...prior, containment });
  persistContainmentToSessionStorage(clientNumber, containment);
  scheduleTenantUiStateSync(clientNumber, div, immediate);
}

function rememberContainmentResult(div, key, data) {
  if (!div || !key) return;
  rememberContainment(div, { [key]: data }, true);
}

function snapshotOpenContainmentPanels() {
  containmentPanelOpen.clear();
  tenantsEl?.querySelectorAll('details.tenant').forEach((details) => {
    const clientNumber = details.dataset.client;
    const panel = details.querySelector('details.containmentPanel');
    if (clientNumber && panel?.open) containmentPanelOpen.add(String(clientNumber));
  });
}

let viewportPinGeneration = 0;

function pinViewport(x, y) {
  const generation = ++viewportPinGeneration;
  let ticks = 0;
  const pin = () => {
    if (generation !== viewportPinGeneration) return;
    if (window.scrollX !== x || window.scrollY !== y) window.scrollTo(x, y);
    const active = document.activeElement;
    if (active && active.tagName === 'SUMMARY' && active.closest('#tenants')) {
      active.blur();
    }
    ticks += 1;
    if (ticks === 1) requestAnimationFrame(pin);
    else if (ticks === 2) requestAnimationFrame(pin);
    else if (ticks === 3) setTimeout(pin, 0);
    else if (ticks === 4) setTimeout(pin, 50);
    else if (ticks === 5) setTimeout(pin, 250);
  };
  pin();
}

function restoreContainmentOutput(div, state, clientNumber) {
  const c = state?.containment;
  if (!div || !c || typeof c !== 'object') return;
  if (c.capabilities) applyContainmentCapabilities(div, c.capabilities);
  if (c.restrictedEmail) {
    try { div.dataset.restrictedEmailJson = JSON.stringify(asArray(c.restrictedEmail)); } catch { /* ignore */ }
  }
  if (c.authMethods) applyAuthMethodsResult(div, asArray(c.authMethods));
  if (c.devices) applyDevicesResult(div, asArray(c.devices));
  if (c.apps) applyAppsResult(div, asArray(c.apps));
  if (c.connectors) applyConnectorsResult(div, asArray(c.connectors));
  if (c.oauth) applyOauthResult(div, asArray(c.oauth));
  if (c.mobile) applyMobileResult(div, asArray(c.mobile));
  if (c.intune) applyIntuneResult(div, asArray(c.intune));
  if (c.folders) applyFoldersResult(div, asArray(c.folders));
  if (c.autoreply) applyAutoreplyResult(div, asArray(c.autoreply));
  if (c.orgfwd) applyOrgfwdResult(div, asArray(c.orgfwd));
  if (c.junk) applyJunkResult(div, asArray(c.junk));
  if (c.journal) applyJournalResult(div, asArray(c.journal));
  if (c.hold) applyHoldResult(div, asArray(c.hold));
  if (c.elsewhere) applyElsewhereResult(div, asArray(c.elsewhere));
  if (c.roles) applyRolesResult(div, asArray(c.roles));
  if (c.appcreds) applyAppcredsResult(div, asArray(c.appcreds));
  if (c.flows) applyFlowsResult(div, asArray(c.flows));
  if (c.transport) applyTransportRulesResult(div, asArray(c.transport));
  if (c.mailbox) applyMailboxAccessResult(div, asArray(c.mailbox));
  if (c.rules) renderContainmentRules(div, c.rulesUser || '', asArray(c.rules));
  if (c.status) setContainmentStatus(div, c.status, { skipSave: true, quietLive: true });
  // Saved containment output used to force this panel open. Chromium then focuses
  // the summary and scrolls it into view, so a refresh or a new tenant jumps to
  // an earlier client that still has results. Reopen only if it was already open.
  const panel = div.querySelector('details.containmentPanel');
  if (panel && clientNumber && containmentPanelOpen.has(String(clientNumber))) panel.open = true;
}

const USER_CONTAINMENT_KEYS = [
  'authMethods', 'devices', 'oauth', 'mobile', 'intune', 'folders', 'autoreply',
  'junk', 'mailbox', 'rules', 'rulesUser', 'hold', 'elsewhere', 'roles', 'restrictedEmail',
];
const TENANT_CONTAINMENT_KEYS = ['transport', 'connectors', 'apps', 'orgfwd', 'journal', 'appcreds', 'flows'];

function rowsForContainmentUser(rows, upn) {
  const needle = String(upn || '').toLowerCase();
  return asArray(rows).filter((r) => {
    if (!r || typeof r !== 'object') return false;
    const vals = [r.UserPrincipalName, r.userPrincipalName, r.Mailbox, r.User, r.mailbox];
    return vals.some((v) => String(v || '').toLowerCase() === needle);
  });
}

function containmentActionsToCsv(rows) {
  const header = 'Timestamp,UPN,Action,Result,Detail';
  const esc = (value) => {
    const s = String(value ?? '');
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return [header, ...asArray(rows).map((r) => [
    r.Timestamp || r.timestamp || '',
    r.UPN || r.upn || '',
    r.Action || r.action || '',
    r.Result || r.result || '',
    r.Detail || r.detail || '',
  ].map(esc).join(','))].join('\n');
}

function rememberContainmentAction(div, entries) {
  const rows = asArray(entries).filter((e) => e && (e.action || e.Action));
  if (!rows.length) return;
  const clientNumber = div?.closest('details.tenant')?.dataset?.client;
  if (!clientNumber) return;
  const prior = tenantUiState.get(String(clientNumber)) || {};
  const existing = asArray(prior.containment?.actions);
  const stamped = rows.map((e) => ({
    Timestamp: e.Timestamp || e.timestamp || new Date().toISOString(),
    UPN: String(e.UPN || e.upn || ''),
    Action: String(e.Action || e.action || ''),
    Result: String(e.Result || e.result || 'success'),
    Detail: String(e.Detail || e.detail || '').slice(0, 500),
  }));
  rememberContainment(div, { actions: existing.concat(stamped) }, true);
}

function noteContainmentAction(div, kind, targets, parsed) {
  const logged = new Set([
    'revoke', 'block', 'unblock', 'unrestrict', 'reset-password', 'assign-password',
    'delete-mfa', 'reregister-mfa', 'delete-devices', 'delete-apps', 'delete-oauth',
    'delete-mobile', 'wipe-intune', 'retire-intune', 'delete-rule', 'set-forward',
    'remove-forward', 'clear-forward', 'add-delegate', 'remove-delegate',
    'delete-folders', 'disable-autoreply', 'delete-junk', 'delete-elsewhere',
    'enable-hold', 'disable-orgfwd', 'delete-journal', 'delete-roles',
    'delete-appcreds', 'delete-flows', 'delete-transport', 'delete-connectors',
  ]);
  if (!logged.has(kind)) return;
  const action = kind === 'assign-password' ? 'reset-password' : kind;
  const details = asArray(parsed?.data?.Details);
  if (details.length) {
    rememberContainmentAction(div, details.map((line) => {
      const text = String(line);
      const split = text.indexOf(':');
      const upn = split > 0 ? text.slice(0, split).trim() : '';
      const rest = split > 0 ? text.slice(split + 1).trim() : text;
      const failed = /\bfail(ed|ure)?\b|\berror\b|\bdenied\b/i.test(rest);
      const safeDetail = (action === 'reset-password')
        ? (failed ? rest : (kind === 'assign-password' ? 'mode=assign' : 'mode=random'))
        : rest;
      return { upn, action, result: failed ? 'failed' : 'success', detail: safeDetail };
    }));
    return;
  }
  const users = asArray(targets).map((t) => {
    if (t == null) return '';
    if (typeof t === 'string') return t;
    return t.UserPrincipalName || t.userPrincipalName || t.Mailbox || t.user || '';
  }).filter(Boolean);
  const ok = parsed?.prefix && !String(parsed.prefix).includes('FAILED');
  const result = ok ? 'success' : 'failed';
  let detail = '';
  if (kind === 'reset-password' || kind === 'assign-password') {
    detail = kind === 'assign-password' ? 'mode=assign' : 'mode=random';
  } else if (parsed?.data?.SuccessCount != null || parsed?.data?.FailCount != null) {
    detail = `success=${parsed.data.SuccessCount || 0}; failed=${parsed.data.FailCount || 0}`;
  } else if (parsed?.raw) {
    detail = String(parsed.raw).slice(0, 240);
  }
  rememberContainmentAction(div, (users.length ? users : ['']).map((upn) => ({ upn, action, result, detail })));
}

function rememberReportFolder(clientNumber, path) {
  if (!clientNumber || !path) return;
  const prior = tenantUiState.get(String(clientNumber)) || {};
  const folders = Array.isArray(prior.reportFolders) ? prior.reportFolders.filter(Boolean) : [];
  if (!folders.includes(path)) folders.push(path);
  tenantUiState.set(String(clientNumber), { ...prior, reportFolders: folders });
}

function buildContainmentExportPacks(div, clientNumber) {
  const c = tenantUiState.get(String(clientNumber))?.containment || {};
  const actions = asArray(c.actions);
  const users = getSelectedContainmentUsers(div);
  const allUsers = new Set(users.map((u) => String(u)));
  USER_CONTAINMENT_KEYS.forEach((key) => {
    asArray(c[key]).forEach((row) => {
      const upn = row?.UserPrincipalName || row?.userPrincipalName || row?.Mailbox;
      if (upn) allUsers.add(String(upn));
    });
  });
  if (c.rulesUser) allUsers.add(String(c.rulesUser));
  actions.forEach((row) => {
    if (row?.UPN || row?.upn) allUsers.add(String(row.UPN || row.upn));
  });
  const packs = [];
  for (const upn of allUsers) {
    if (!upn || upn === '_tenant') continue;
    const files = {
      'readme.txt': `Containment pull for ${upn}\nSaved ${new Date().toISOString()}\nIncludes actions.csv (password reset, revoke, block, and other account changes).\n`,
    };
    if (c.status) files['status.txt'] = String(c.status);
    const userFiles = {
      authMethods: 'mfa.json',
      devices: 'entra-devices.json',
      oauth: 'oauth-consents.json',
      mobile: 'activesync.json',
      intune: 'intune.json',
      folders: 'folder-permissions.json',
      autoreply: 'auto-reply.json',
      junk: 'junk-trusted.json',
      mailbox: 'mailbox-access.json',
      hold: 'hold-audit.json',
      elsewhere: 'rights-elsewhere.json',
      roles: 'roles-groups.json',
      restrictedEmail: 'restricted-users.json',
    };
    Object.entries(userFiles).forEach(([key, name]) => {
      const rows = rowsForContainmentUser(c[key], upn);
      if (rows.length) files[name] = rows;
    });
    if (c.rules && String(c.rulesUser || '').toLowerCase() === upn.toLowerCase()) {
      files['inbox-rules.json'] = asArray(c.rules);
    }
    const userActions = actions.filter((row) => String(row.UPN || row.upn || '').toLowerCase() === upn.toLowerCase());
    if (userActions.length) files['actions.csv'] = containmentActionsToCsv(userActions);
    if (Object.keys(files).length > 2 || files['mfa.json'] || files['mailbox-access.json'] || files['inbox-rules.json'] || files['actions.csv']) {
      packs.push({ user: upn, files });
    } else if (c.status && String(c.status).toLowerCase().includes(upn.toLowerCase())) {
      packs.push({ user: upn, files });
    }
  }
  const tenantFiles = {};
  TENANT_CONTAINMENT_KEYS.forEach((key) => {
    if (c[key] && asArray(c[key]).length) tenantFiles[`${key}.json`] = c[key];
  });
  const tenantActions = actions.filter((row) => !String(row.UPN || row.upn || '').trim());
  if (tenantActions.length) tenantFiles['actions.csv'] = containmentActionsToCsv(tenantActions);
  if (Object.keys(tenantFiles).length) {
    tenantFiles['readme.txt'] = `Tenant-wide containment lists and account-change log\nSaved ${new Date().toISOString()}\n`;
    packs.push({ user: '_tenant', files: tenantFiles });
  }
  return packs;
}

async function saveContainmentPacks(clientNumber, div) {
  const packs = buildContainmentExportPacks(div, clientNumber);
  const actions = asArray(tenantUiState.get(String(clientNumber))?.containment?.actions);
  if (!packs.length && !actions.length) {
    log(`Client ${clientNumber}: nothing to save — list containment data or perform an account change first.`);
    return;
  }
  let data;
  try {
    data = await api(`/api/tenants/${clientNumber}/containment-pack`, {
      method: 'POST',
      body: JSON.stringify({
        outputFolder: div.dataset.outputFolder || '',
        companyName: getTenantCompanyName(div, clientNumber),
        packs,
        actions,
      }),
    });
  } catch (e) {
    if (String(e.message || e).includes('Not found')) {
      throw new Error('Save containment zips needs a web-runner restart (new /api/tenants/.../containment-pack route).');
    }
    throw e;
  }
  if (data.folder) {
    div.dataset.outputFolder = data.folder;
    rememberReportFolder(clientNumber, data.folder);
    const openBtn = div.querySelector('.openReports');
    if (openBtn) openBtn.disabled = false;
    refreshTenantSummaryUI(clientNumber, { outputFolder: data.folder });
  }
  const names = (data.files || []).map((f) => String(f).split(/[/\\]/).pop()).join(', ');
  const audit = data.auditCsv ? `\nAccount changes: ${String(data.auditCsv).split(/[/\\]/).pop()}` : '';
  setContainmentStatus(div, `Saved ${data.files?.length || 0} zip(s) to ${data.folder}${names ? `\n${names}` : ''}${audit}`, { reveal: true });
  log(`Client ${clientNumber}: saved containment zips to ${data.folder}`);
}

function clearContainmentUserPulls(div) {
  const clientNumber = div?.closest('details.tenant')?.dataset?.client;
  if (!clientNumber) return;
  const prior = tenantUiState.get(String(clientNumber)) || {};
  const next = { ...(prior.containment || {}) };
  USER_CONTAINMENT_KEYS.forEach((key) => { delete next[key]; });
  next.status = 'User containment pulls cleared. Account-change log and tenant-wide lists kept. Saved zips were not deleted.';
  tenantUiState.set(String(clientNumber), { ...prior, containment: next });
  persistContainmentToSessionStorage(clientNumber, next);
  scheduleTenantUiStateSync(clientNumber, div, true);
  delete div.dataset.restrictedEmailJson;
  delete div.dataset.containmentRulesUser;
  applyAuthMethodsResult(div, []);
  applyDevicesResult(div, []);
  applyOauthResult(div, []);
  applyMobileResult(div, []);
  applyIntuneResult(div, []);
  applyFoldersResult(div, []);
  applyAutoreplyResult(div, []);
  applyJunkResult(div, []);
  applyMailboxAccessResult(div, []);
  applyHoldResult(div, []);
  applyElsewhereResult(div, []);
  applyRolesResult(div, []);
  renderContainmentRules(div, '', []);
  const after = tenantUiState.get(String(clientNumber)) || {};
  const cleaned = { ...(after.containment || {}) };
  USER_CONTAINMENT_KEYS.forEach((key) => { delete cleaned[key]; });
  cleaned.status = next.status;
  tenantUiState.set(String(clientNumber), { ...after, containment: cleaned });
  persistContainmentToSessionStorage(clientNumber, cleaned);
  setContainmentStatus(div, next.status, { skipSave: true });
}

const REMEDIATE_SUCCESS_PREFIXES = [
  'REMEDIATE_SUCCESS:', 'REMEDIATE_RESTRICTED:', 'REMEDIATE_RULES:', 'REMEDIATE_MAILBOX:',
  'REMEDIATE_TRANSPORT:', 'REMEDIATE_CONNECTORS:', 'REMEDIATE_AUTHMETHODS:', 'REMEDIATE_DEVICES:',
  'REMEDIATE_APPS:', 'REMEDIATE_OAUTH:', 'REMEDIATE_MOBILE:', 'REMEDIATE_FOLDERS:',
  'REMEDIATE_AUTOREPLY:', 'REMEDIATE_ORGFWD:', 'REMEDIATE_JUNK:', 'REMEDIATE_JOURNAL:',
  'REMEDIATE_HOLD:', 'REMEDIATE_ELSEWHERE:', 'REMEDIATE_ROLES:', 'REMEDIATE_INTUNE:',
  'REMEDIATE_APPCREDS:', 'REMEDIATE_FLOWS:', 'REMEDIATE_BEC:',
];

function extractWorkerTokenResponse(value, prefixes) {
  const text = normalizeResponse(value);
  if (!text || !prefixes?.length) return text;
  for (const prefix of prefixes) {
    const idx = text.indexOf(prefix);
    if (idx >= 0) return text.slice(idx);
  }
  return text;
}

function parseRemediatePayload(final) {
  const prefixes = ['REMEDIATE_FAILED:', ...REMEDIATE_SUCCESS_PREFIXES];
  const text = extractWorkerTokenResponse(final || '', prefixes);
  for (const prefix of prefixes) {
    if (text.startsWith(prefix)) {
      const rest = text.slice(prefix.length);
      try {
        return { prefix, data: JSON.parse(rest), raw: rest };
      } catch {
        return { prefix, data: null, raw: rest };
      }
    }
  }
  return { prefix: '', data: null, raw: text };
}

function shortenGraphError(text) {
  const t = String(text || '').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  if (/not applicable to target tenant/i.test(t)) return 'Intune is not available in this tenant.';
  if (/Request_ResourceNotFound|does not exist or one of its queried reference-property/i.test(t)) {
    return 'Already gone (not found).';
  }
  const m = t.match(/"message"\s*:\s*"((?:\\.|[^"\\])*)"/);
  if (m) {
    return m[1].replace(/\\u0027/g, "'").replace(/\\"/g, '"').replace(/\\n/g, ' ');
  }
  if (t.startsWith('{') || t.length > 280) {
    return `${t.slice(0, 240)}${t.length > 240 ? '…' : ''}`;
  }
  return t;
}

function isGraphAlreadyGone(text) {
  return /Request_ResourceNotFound|does not exist or one of its queried reference-property|Already gone/i.test(String(text || ''));
}

function summarizeRemediateFailure(parsed, fallback) {
  const data = parsed && parsed.data && typeof parsed.data === 'object' ? parsed.data : null;
  if (data) {
    const deleted = asArray(data.Deleted).concat(asArray(data.Results));
    const failed = deleted.filter((x) => x && x.Success === false);
    const ok = Number(data.SuccessCount);
    const fail = Number(data.FailCount);
    const hasCounts = Number.isFinite(ok) || Number.isFinite(fail);
    if (failed.length || (hasCounts && (ok || fail))) {
      const gone = failed.filter((x) => isGraphAlreadyGone(x.Error)).length;
      if (fail > 0 && !(ok > 0) && gone === failed.length && failed.length) {
        return { text: `Already gone (${gone}). Reloaded remaining rows.`, tone: 'statusOk' };
      }
      const first = failed[0]?.Error || data.reason || '';
      const short = shortenGraphError(first);
      const counts = hasCounts ? `Success ${ok || 0}; failed ${fail || 0}.` : (fallback || 'Action failed.');
      return { text: short ? `${counts} ${short}` : counts, tone: fail > 0 ? 'statusBad' : 'statusWarn' };
    }
    if (data.reason) return { text: String(data.reason), tone: 'statusBad' };
    const rowErr = asArray(data.Devices).concat(asArray(data.Users), asArray(data.Methods), asArray(data.Grants))
      .map((x) => x && x.Error)
      .find(Boolean);
    if (rowErr) return { text: shortenGraphError(rowErr), tone: 'statusWarn' };
  }
  const raw = (parsed && parsed.raw) || '';
  if (raw) {
    return { text: shortenGraphError(raw) || fallback || 'Action failed.', tone: 'statusBad' };
  }
  return { text: fallback || 'Action failed.', tone: 'statusBad' };
}

function showContainmentFailure(div, clientNumber, parsed, fallback) {
  const s = summarizeRemediateFailure(parsed, fallback);
  setContainmentStatus(div, s.text, { tone: s.tone, reveal: true });
  log(`Client ${clientNumber}: ${s.text}`);
  return s;
}

function containmentDateText(...values) {
  for (const value of values) {
    if (value == null || value === '') continue;
    if (typeof value === 'object') {
      const inner = containmentDateText(
        value.DateTime,
        value.dateTime,
        value.UtcDateTime,
        value.utcDateTime,
        value.Value,
        value.value,
        value.startDateTime,
        value.StartDateTime,
        value.WhenChanged,
        value.WhenCreated,
        value.createdDateTime
      );
      if (inner) return inner;
      continue;
    }
    const s = String(value).trim();
    if (!s || s === '[object Object]') continue;
    const msDotNet = s.match(/^\/Date\((-?\d+)(?:[+-]\d+)?\)\/$/);
    if (msDotNet) {
      const d = new Date(Number(msDotNet[1]));
      if (!Number.isNaN(d.getTime())) return formatPresentedLocal(d.toISOString());
    }
    if (/(?:Z|[+-]\d{2}:?\d{2})$/i.test(s) || /^\d{4}-\d{2}-\d{2}T/.test(s)) {
      const presented = formatPresentedLocal(s);
      if (presented) return presented;
    }
    return s;
  }
  return '';
}

function parsePresentedLocalMs(text) {
  const m = String(text || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])\b/);
  if (!m) return null;
  let hour = Number(m[4]);
  const ampm = m[7] || '';
  if (/p/i.test(ampm) && hour < 12) hour += 12;
  if (/a/i.test(ampm) && hour === 12) hour = 0;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), hour, Number(m[5]), Number(m[6] || 0));
  return Number.isNaN(d.getTime()) ? null : d.getTime();
}

function parseContainmentDateMs(text) {
  const raw = containmentDateText(text).replace(/\s+/g, ' ').trim();
  if (!raw) return null;
  const localMs = parsePresentedLocalMs(raw);
  if (localMs != null) return localMs;
  const created = raw.match(/^created\s+(.+)$/i);
  const src = created ? created[1].trim() : raw;
  const iso = src.match(/^(\d{4}-\d{2}-\d{2}(?:[T\s]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?)/);
  if (iso) {
    let stamp = iso[1];
    if (!/[TZ+-]/.test(stamp) && /^\d{4}-\d{2}-\d{2} /.test(stamp)) stamp = stamp.replace(' ', 'T');
    const d = Date.parse(stamp);
    return Number.isFinite(d) ? d : null;
  }
  const us = src.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (us) {
    const d = new Date(Number(us[3]), Number(us[1]) - 1, Number(us[2]), Number(us[4] || 0), Number(us[5] || 0), Number(us[6] || 0));
    return Number.isNaN(d.getTime()) ? null : d.getTime();
  }
  return null;
}

function parseContainmentCell(text) {
  const s = String(text || '').replace(/\s+/g, ' ').trim();
  if (!s || s === '(none)' || s === '—') return { empty: true, kind: 'empty', n: 0, t: '' };
  const ms = parseContainmentDateMs(s);
  if (ms != null) return { empty: false, kind: 'date', n: ms, t: s.toLowerCase() };
  if (/^-?\d+(\.\d+)?$/.test(s)) return { empty: false, kind: 'num', n: Number(s), t: s };
  return { empty: false, kind: 'text', n: 0, t: s.toLowerCase() };
}

function containmentHeaderLabel(th) {
  if (!th) return '';
  const clone = th.cloneNode(true);
  clone.querySelectorAll('.sortInd').forEach((n) => n.remove());
  return String(clone.textContent || '').replace(/\s+/g, ' ').trim();
}

const CONTAINMENT_DATE_HEADER_RES = [
  /^last sign-in$/i,
  /^last sync$/i,
  /^first sync$/i,
  /^created$/i,
  /^changed$/i,
  /^modified$/i,
  /^hold date$/i,
  /^start$/i,
  /^expires$/i,
  /^end$/i,
  /^date$/i,
];

function isContainmentDateHeader(label) {
  return CONTAINMENT_DATE_HEADER_RES.some((re) => re.test(label));
}

function findDefaultContainmentDateColumn(table) {
  const ths = [...table.querySelectorAll('thead th')];
  let best = -1;
  let bestRank = CONTAINMENT_DATE_HEADER_RES.length;
  ths.forEach((th, idx) => {
    if (th.classList.contains('colEmpty')) return;
    const label = containmentHeaderLabel(th);
    const rank = CONTAINMENT_DATE_HEADER_RES.findIndex((re) => re.test(label));
    if (rank >= 0 && rank < bestRank) {
      best = idx;
      bestRank = rank;
    }
  });
  return best;
}

function isContainmentEmptyCellText(text) {
  const s = String(text || '').replace(/\s+/g, ' ').trim();
  return !s || s === '—' || s === '-';
}

function hideEmptyContainmentColumns(table) {
  if (!table) return;
  const ths = [...table.querySelectorAll('thead tr:first-child th')];
  const rows = [...table.querySelectorAll('tbody tr')];
  ths.forEach((th) => th.classList.remove('colEmpty'));
  rows.forEach((row) => {
    [...row.children].forEach((cell) => cell.classList.remove('colEmpty'));
  });
  const isPlaceholder = !rows.length || (rows.length === 1 && rows[0].querySelector('td.muted, td[colspan]'));
  if (isPlaceholder) return;
  ths.forEach((th, i) => {
    const label = containmentHeaderLabel(th);
    if (i === 0 && !label) return;
    const allEmpty = rows.every((row) => isContainmentEmptyCellText(row.children[i]?.textContent));
    if (!allEmpty) return;
    th.classList.add('colEmpty');
    rows.forEach((row) => {
      if (row.children[i]) row.children[i].classList.add('colEmpty');
    });
  });
  if (table.dataset.sortCol !== undefined && table.dataset.sortCol !== '') {
    const th = ths[Number(table.dataset.sortCol)];
    if (th && th.classList.contains('colEmpty')) {
      delete table.dataset.sortCol;
      delete table.dataset.sortDir;
      delete table.dataset.sortUser;
    }
  }
}

function updateContainmentSortIndicators(table) {
  const col = table.dataset.sortCol;
  const dir = table.dataset.sortDir || 'desc';
  table.querySelectorAll('thead th').forEach((th, idx) => {
    const ind = th.querySelector('.sortInd');
    if (!ind) return;
    const active = col !== undefined && col !== '' && Number(col) === idx;
    th.classList.toggle('sortActive', active);
    ind.textContent = active ? (dir === 'asc' ? '▲' : '▼') : '';
  });
}

function applyContainmentTableSort(table) {
  const tbody = table.querySelector('tbody');
  if (!tbody) return;
  const rows = [...tbody.querySelectorAll('tr')];
  const isEmpty = !rows.length || (rows.length === 1 && rows[0].querySelector('td.muted, td[colspan]'));
  if (isEmpty) {
    updateContainmentSortIndicators(table);
    return;
  }
  let col = table.dataset.sortCol !== undefined && table.dataset.sortCol !== '' ? Number(table.dataset.sortCol) : -1;
  if (!(col >= 0) && table.dataset.sortUser !== '1') {
    col = findDefaultContainmentDateColumn(table);
    if (col >= 0) {
      table.dataset.sortCol = String(col);
      if (!table.dataset.sortDir) table.dataset.sortDir = 'desc';
    }
  }
  if (!(col >= 0)) {
    updateContainmentSortIndicators(table);
    return;
  }
  const dir = table.dataset.sortDir === 'asc' ? 1 : -1;
  const ths = [...table.querySelectorAll('thead th')];
  const preferDate = isContainmentDateHeader(containmentHeaderLabel(ths[col]));
  rows.sort((a, b) => {
    const va = parseContainmentCell(a.children[col]?.textContent);
    const vb = parseContainmentCell(b.children[col]?.textContent);
    if (va.empty && vb.empty) return 0;
    if (va.empty) return 1;
    if (vb.empty) return -1;
    if (preferDate) {
      const da = va.kind === 'date' ? va.n : null;
      const db = vb.kind === 'date' ? vb.n : null;
      if (da == null && db == null) return va.t.localeCompare(vb.t, undefined, { numeric: true, sensitivity: 'base' }) * dir;
      if (da == null) return 1;
      if (db == null) return -1;
      return (da - db) * dir;
    }
    if (va.kind === 'num' && vb.kind === 'num') return (va.n - vb.n) * dir;
    if (va.kind === 'date' && vb.kind === 'date') return (va.n - vb.n) * dir;
    return va.t.localeCompare(vb.t, undefined, { numeric: true, sensitivity: 'base' }) * dir;
  });
  rows.forEach((row) => tbody.appendChild(row));
  updateContainmentSortIndicators(table);
}

function wireContainmentTableSort(table) {
  if (!table) return;
  hideEmptyContainmentColumns(table);
  if (table.dataset.sortWired !== '1') {
    table.dataset.sortWired = '1';
    [...table.querySelectorAll('thead th')].forEach((th, idx) => {
      const label = containmentHeaderLabel(th);
      if (idx === 0 && !label) return;
      th.classList.add('sortable');
      if (!th.querySelector('.sortInd')) {
        const ind = document.createElement('span');
        ind.className = 'sortInd';
        th.appendChild(ind);
      }
      th.addEventListener('click', () => {
        const cur = table.dataset.sortCol;
        if (cur === String(idx)) {
          table.dataset.sortDir = table.dataset.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          table.dataset.sortCol = String(idx);
          table.dataset.sortDir = isContainmentDateHeader(label) ? 'desc' : 'asc';
        }
        table.dataset.sortUser = '1';
        applyContainmentTableSort(table);
      });
    });
  }
  applyContainmentTableSort(table);
}

function containmentCreatedText(created, details) {
  const direct = containmentDateText(created);
  if (direct) return direct;
  const m = String(details || '').match(/created\s+(.+?)(?:\s*·|$)/i);
  return m ? m[1].trim() : '';
}

function renderContainmentRules(div, mailbox, rules) {
  const rows = asArray(rules);
  rememberContainment(div, { rules: rows, rulesUser: mailbox || '' }, true);
  const wrap = div.querySelector('.containmentRulesWrap');
  const tbody = div.querySelector('.containmentRulesTable tbody');
  if (!wrap || !tbody) return;
  div.dataset.containmentRulesUser = mailbox || '';
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="muted">No inbox rules (including hidden).</td></tr>';
    wrap.style.display = 'block';
    updateContainmentButtons(div);
    wireContainmentTableSort(div.querySelector('.containmentRulesTable'));
    return;
  }
  tbody.innerHTML = rows.map((rule, idx) => {
    // RuleIdentity first: mailboxes often hold several rules with the same name (".", "..."),
    // and Remove-InboxRule rejects an ambiguous name with "matches multiple entries".
    const identity = escapeHtml(String(rule.RuleIdentity || rule.Identity || rule.Name || ''));
    const name = escapeHtml(String(rule.Name || rule.Identity || ''));
    const details = [
      rule.Description,
      rule.RedirectTo ? `Redirect: ${rule.RedirectTo}` : '',
      rule.ForwardTo ? `Forward: ${rule.ForwardTo}` : '',
      rule.ForwardAsAttachmentTo ? `Fwd attach: ${rule.ForwardAsAttachmentTo}` : '',
      rule.DeleteMessage ? 'Deletes messages' : '',
      rule.From ? `From: ${rule.From}` : '',
    ].filter(Boolean).join(' · ');
    return `<tr>
      <td><input type="checkbox" class="containmentRulePick" data-identity="${identity}" data-label="${name}" data-idx="${idx}" /></td>
      <td>${name}</td>
      <td>${rule.Enabled ? 'yes' : 'no'}</td>
      <td>${rule.Priority != null ? escapeHtml(String(rule.Priority)) : ''}</td>
      <td>${rule.Hidden ? 'yes' : 'no'}</td>
      <td>${escapeHtml(details)}</td>
    </tr>`;
  }).join('');
  wrap.style.display = 'block';
  tbody.querySelectorAll('.containmentRulePick').forEach((cb) => {
    cb.addEventListener('change', () => updateContainmentButtons(div));
  });
  updateContainmentButtons(div);
  wireContainmentTableSort(div.querySelector('.containmentRulesTable'));
}

function applyMailboxAccessResult(div, users) {
  rememberContainmentResult(div, 'mailbox', users);
  const wrap = div.querySelector('.containmentMailboxWrap');
  const tbody = div.querySelector('.containmentMailboxTable tbody');
  if (!wrap || !tbody) return;
  const rows = [];
  const mailboxList = asArray(users);
  mailboxList.forEach((mbx) => {
    const upn = mbx.UserPrincipalName || '';
    if (mbx.Error) {
      rows.push({ kind: 'error', mailbox: upn, type: 'Error', target: mbx.Error, keep: '', field: '' });
      return;
    }
    const listed = asArray(mbx.Forwards);
    if (listed.length) {
      listed.forEach((f) => {
        const field = f.Field === 'Recipient' ? 'Recipient' : 'Smtp';
        rows.push({
          kind: 'forward',
          mailbox: upn,
          type: field === 'Recipient' ? 'Forward recipient' : 'Forward SMTP',
          target: f.Address || '',
          keep: mbx.DeliverToMailboxAndForward ? 'yes' : 'no',
          field,
        });
      });
    } else {
      const smtp = mbx.ForwardingSmtpAddress || '';
      const recip = mbx.ForwardingAddress || '';
      if (smtp) {
        rows.push({ kind: 'forward', mailbox: upn, type: 'Forward SMTP', target: smtp, keep: mbx.DeliverToMailboxAndForward ? 'yes' : 'no', field: 'Smtp' });
      }
      if (recip) {
        rows.push({ kind: 'forward', mailbox: upn, type: 'Forward recipient', target: recip, keep: mbx.DeliverToMailboxAndForward ? 'yes' : 'no', field: 'Recipient' });
      }
      if (!smtp && !recip) {
        rows.push({ kind: 'none', mailbox: upn, type: 'Forward', target: '(none)', keep: '', field: '' });
      }
    }
    asArray(mbx.Delegates).forEach((d) => {
      rows.push({
        kind: 'delegate',
        mailbox: upn,
        type: d.Right || '',
        target: d.User || '',
        keep: '',
        field: '',
      });
    });
  });
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="muted">No forwarding or delegates.</td></tr>';
    wrap.style.display = 'block';
    updateContainmentButtons(div);
    wireContainmentTableSort(div.querySelector('.containmentMailboxTable'));
    return;
  }
  tbody.innerHTML = rows.map((row) => {
    let pick = '';
    if (row.kind === 'delegate') {
      pick = `<input type="checkbox" class="containmentDelegatePick" data-mailbox="${escapeHtml(row.mailbox)}" data-right="${escapeHtml(row.type)}" data-user="${escapeHtml(row.target)}" />`;
    } else if (row.kind === 'forward' && row.field) {
      pick = `<input type="checkbox" class="containmentForwardPick" data-mailbox="${escapeHtml(row.mailbox)}" data-field="${escapeHtml(row.field)}" data-target="${escapeHtml(row.target)}" />`;
    }
    return `<tr>
      <td>${pick}</td>
      <td>${escapeHtml(row.mailbox)}</td>
      <td>${escapeHtml(row.type)}</td>
      <td>${escapeHtml(row.target)}</td>
      <td>${escapeHtml(row.keep)}</td>
    </tr>`;
  }).join('');
  wrap.style.display = 'block';
  tbody.querySelectorAll('.containmentForwardPick, .containmentDelegatePick').forEach((cb) => {
    cb.addEventListener('change', () => updateContainmentButtons(div));
  });
  updateContainmentButtons(div);
  wireContainmentTableSort(div.querySelector('.containmentMailboxTable'));
}

function renderGenericPickTable(div, wrapSel, tbodySel, pickClass, emptyCols, emptyText, htmlRows) {
  const wrap = div.querySelector(wrapSel);
  const tbody = div.querySelector(tbodySel);
  if (!wrap || !tbody) return;
  if (!htmlRows.length) {
    tbody.innerHTML = `<tr><td colspan="${emptyCols}" class="muted">${emptyText}</td></tr>`;
    wrap.style.display = 'block';
    updateContainmentButtons(div);
    wireContainmentTableSort(tbody.closest('table'));
    return;
  }
  tbody.innerHTML = htmlRows.join('');
  wrap.style.display = 'block';
  tbody.querySelectorAll(`.${pickClass}`).forEach((cb) => {
    cb.addEventListener('change', () => updateContainmentButtons(div));
  });
  updateContainmentButtons(div);
  wireContainmentTableSort(tbody.closest('table'));
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value == null || value === '') return [];
  return [value];
}

function applyTransportRulesResult(div, rules) {
  const list = asArray(rules).map((r) => ({ ...r, Description: String(r.Description || '').slice(0, 240) }));
  rememberContainmentResult(div, 'transport', list);
  const html = list.map((rule) => {
    const identity = escapeHtml(String(rule.Identity || rule.Name || ''));
    return `<tr>
      <td><input type="checkbox" class="containmentTransportPick" data-identity="${identity}" /></td>
      <td>${escapeHtml(String(rule.Name || ''))}</td>
      <td>${escapeHtml(String(rule.State || ''))}</td>
      <td>${rule.Priority != null ? escapeHtml(String(rule.Priority)) : ''}</td>
      <td>${escapeHtml(String(rule.Description || ''))}</td>
      <td>${escapeHtml(containmentDateText(rule.WhenChanged, rule.WhenChangedUTC, rule.LastModifiedTime))}</td>
    </tr>`;
  });
  renderGenericPickTable(div, '.containmentTransportWrap', '.containmentTransportTable tbody', 'containmentTransportPick', 6, 'No transport rules.', html);
}

function applyAuthMethodsResult(div, methods) {
  rememberContainmentResult(div, 'authMethods', methods);
  const list = asArray(methods).filter((m) => m && (m.Id || m.Error || m.Type));
  const html = list.map((m) => {
    const canDelete = m.CanDelete !== false && m.Id;
    const pick = canDelete
      ? `<input type="checkbox" class="containmentMfaPick" data-user="${escapeHtml(m.UserPrincipalName || '')}" data-id="${escapeHtml(m.Id || '')}" data-odata="${escapeHtml(m.ODataType || '')}" />`
      : '';
    return `<tr>
      <td>${pick}</td>
      <td>${escapeHtml(m.UserPrincipalName || '')}</td>
      <td>${escapeHtml(m.Type || '')}</td>
      <td>${escapeHtml(m.Details || m.Error || '')}</td>
      <td>${escapeHtml(containmentCreatedText(m.Created, m.Details))}</td>
    </tr>`;
  });
  renderGenericPickTable(div, '.containmentMfaWrap', '.containmentMfaTable tbody', 'containmentMfaPick', 5, 'No MFA methods.', html);
}

function applyDevicesResult(div, devices) {
  rememberContainmentResult(div, 'devices', devices);
  const list = asArray(devices).filter((d) => d && (d.Id || d.Error || d.DisplayName));
  const html = list.filter((d) => d.Id || d.Error).map((d) => {
    const pick = d.Id
      ? `<input type="checkbox" class="containmentDevicePick" data-id="${escapeHtml(d.Id || '')}" data-user="${escapeHtml(d.UserPrincipalName || '')}" />`
      : '';
    return `<tr>
      <td>${pick}</td>
      <td>${escapeHtml(d.UserPrincipalName || '')}</td>
      <td>${escapeHtml(d.DisplayName || '')}</td>
      <td>${escapeHtml(d.OperatingSystem || '')}</td>
      <td>${escapeHtml(d.TrustType || '')}</td>
      <td>${escapeHtml(d.Relation || '')}</td>
      <td>${escapeHtml(containmentDateText(d.LastSignIn, d.approximateLastSignInDateTime) || d.Error || '')}</td>
    </tr>`;
  });
  renderGenericPickTable(div, '.containmentDevicesWrap', '.containmentDevicesTable tbody', 'containmentDevicePick', 7, 'No registered or owned devices.', html);
}

function applyAppsResult(div, apps) {
  rememberContainmentResult(div, 'apps', apps);
  const list = asArray(apps);
  const html = list.map((a) => {
    const id = escapeHtml(String(a.Id || ''));
    const kind = escapeHtml(String(a.Kind || ''));
    return `<tr>
      <td><input type="checkbox" class="containmentAppPick" data-id="${id}" data-kind="${kind}" data-name="${escapeHtml(String(a.DisplayName || ''))}" /></td>
      <td>${kind}</td>
      <td>${escapeHtml(String(a.DisplayName || ''))}</td>
      <td>${escapeHtml(String(a.AppId || ''))}</td>
      <td>${escapeHtml(containmentDateText(a.Created, a.createdDateTime))}</td>
      <td>${escapeHtml(String(a.Publisher || ''))}</td>
    </tr>`;
  });
  renderGenericPickTable(div, '.containmentAppsWrap', '.containmentAppsTable tbody', 'containmentAppPick', 6, 'No app registrations or other-tenant enterprise apps.', html);
}

function applyConnectorsResult(div, connectors) {
  rememberContainmentResult(div, 'connectors', connectors);
  const list = asArray(connectors);
  const html = list.map((c) => {
    const name = escapeHtml(String(c.Name || ''));
    const direction = escapeHtml(String(c.Direction || ''));
    const details = [c.ConnectorType, c.SmartHosts ? `SmartHosts: ${c.SmartHosts}` : '', c.SenderIPAddresses ? `IPs: ${c.SenderIPAddresses}` : '', c.Domains ? `Domains: ${c.Domains}` : ''].filter(Boolean).join(' · ');
    return `<tr>
      <td><input type="checkbox" class="containmentConnectorPick" data-name="${name}" data-direction="${direction}" /></td>
      <td>${direction}</td>
      <td>${name}</td>
      <td>${c.Enabled ? 'yes' : 'no'}</td>
      <td>${escapeHtml(details)}</td>
      <td>${escapeHtml(containmentDateText(c.WhenCreated, c.WhenChanged, c.WhenCreatedUTC))}</td>
    </tr>`;
  });
  renderGenericPickTable(div, '.containmentConnectorsWrap', '.containmentConnectorsTable tbody', 'containmentConnectorPick', 6, 'No inbound or outbound connectors.', html);
}

function applyOauthResult(div, grants) {
  rememberContainmentResult(div, 'oauth', grants);
  const html = asArray(grants).filter((g) => g.Id || g.Error).map((g) => `<tr>
      <td>${g.Id ? `<input type="checkbox" class="containmentOauthPick" data-id="${escapeHtml(String(g.Id))}" data-user="${escapeHtml(String(g.UserPrincipalName || ''))}" data-app="${escapeHtml(String(g.App || ''))}" />` : ''}</td>
      <td>${escapeHtml(String(g.UserPrincipalName || ''))}</td>
      <td>${escapeHtml(String(g.App || g.Error || ''))}</td>
      <td>${escapeHtml(String(g.Scope || ''))}</td>
      <td>${escapeHtml(String(g.ConsentType || ''))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentOauthWrap', '.containmentOauthTable tbody', 'containmentOauthPick', 5, 'No OAuth consents.', html);
}

function applyMobileResult(div, devices) {
  rememberContainmentResult(div, 'mobile', devices);
  const html = asArray(devices).filter((d) => d.Identity || d.Error).map((d) => `<tr>
      <td>${d.Identity ? `<input type="checkbox" class="containmentMobilePick" data-identity="${escapeHtml(String(d.Identity))}" data-user="${escapeHtml(String(d.UserPrincipalName || ''))}" />` : ''}</td>
      <td>${escapeHtml(String(d.UserPrincipalName || ''))}</td>
      <td>${escapeHtml(String(d.FriendlyName || d.Error || ''))}</td>
      <td>${escapeHtml(String(d.DeviceType || ''))}</td>
      <td>${escapeHtml(containmentDateText(d.FirstSync, d.FirstSyncTime))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentMobileWrap', '.containmentMobileTable tbody', 'containmentMobilePick', 5, 'No ActiveSync partnerships.', html);
}

function applyIntuneResult(div, devices) {
  rememberContainmentResult(div, 'intune', devices);
  const list = asArray(devices);
  const errOnly = list.filter((d) => d.Error && !d.Id);
  if (errOnly.length && !list.some((d) => d.Id)) {
    const msg = shortenGraphError(errOnly[0].Error) || errOnly[0].Error;
    renderGenericPickTable(div, '.containmentIntuneWrap', '.containmentIntuneTable tbody', 'containmentIntunePick', 6, msg, []);
    return;
  }
  const html = list.filter((d) => d.Id || d.Error).map((d) => `<tr>
      <td>${d.Id ? `<input type="checkbox" class="containmentIntunePick" data-id="${escapeHtml(String(d.Id))}" data-name="${escapeHtml(String(d.DeviceName || ''))}" />` : ''}</td>
      <td>${escapeHtml(String(d.UserPrincipalName || ''))}</td>
      <td>${escapeHtml(String(d.DeviceName || d.Error || ''))}</td>
      <td>${escapeHtml(String(d.Os || ''))}</td>
      <td>${escapeHtml(String(d.Compliance || ''))}</td>
      <td>${escapeHtml(containmentDateText(d.LastSync, d.lastSyncDateTime))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentIntuneWrap', '.containmentIntuneTable tbody', 'containmentIntunePick', 6, 'No Intune-managed devices.', html);
}

function applyFoldersResult(div, perms) {
  rememberContainmentResult(div, 'folders', perms);
  const html = asArray(perms).filter((p) => (p.User && p.User !== '(none)') || p.Error).map((p) => `<tr>
      <td>${p.User && p.User !== '(none)' ? `<input type="checkbox" class="containmentFolderPick" data-user="${escapeHtml(String(p.UserPrincipalName || ''))}" data-folder="${escapeHtml(String(p.Folder || ''))}" data-trustee="${escapeHtml(String(p.User || ''))}" />` : ''}</td>
      <td>${escapeHtml(String(p.UserPrincipalName || ''))}</td>
      <td>${escapeHtml(String(p.Folder || ''))}</td>
      <td>${escapeHtml(String(p.User || p.Error || ''))}</td>
      <td>${escapeHtml(String(p.AccessRights || ''))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentFoldersWrap', '.containmentFoldersTable tbody', 'containmentFolderPick', 5, 'No folder permissions besides empty Default/Anonymous.', html);
}

function applyAutoreplyResult(div, users) {
  rememberContainmentResult(div, 'autoreply', users);
  const wrap = div.querySelector('.containmentAutoreplyWrap');
  const tbody = div.querySelector('.containmentAutoreplyTable tbody');
  if (!wrap || !tbody) return;
  const list = asArray(users);
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="6" class="muted">No auto-reply data.</td></tr>';
  } else {
    tbody.innerHTML = list.map((u) => `<tr>
      <td>${escapeHtml(String(u.UserPrincipalName || ''))}</td>
      <td>${escapeHtml(String(u.AutoReplyState || u.Error || ''))}</td>
      <td>${escapeHtml(String(u.ExternalAudience || ''))}</td>
      <td>${escapeHtml(containmentDateText(u.StartTime))}</td>
      <td>${escapeHtml(containmentDateText(u.EndTime))}</td>
      <td>${escapeHtml(String((u.InternalMessage || u.ExternalMessage || '').replace(/<[^>]+>/g, ' ').slice(0, 180)))}</td>
    </tr>`).join('');
  }
  wrap.style.display = 'block';
  wireContainmentTableSort(div.querySelector('.containmentAutoreplyTable'));
}

function applyOrgfwdResult(div, policies) {
  rememberContainmentResult(div, 'orgfwd', policies);
  const html = asArray(policies).filter((p) => p.Identity || p.Error).map((p) => {
    const value = p.Kind === 'OutboundSpam' ? (p.AutoForwardingMode || '') : String(p.AutoForward);
    return `<tr>
      <td>${p.Identity ? `<input type="checkbox" class="containmentOrgfwdPick" data-kind="${escapeHtml(String(p.Kind || ''))}" data-identity="${escapeHtml(String(p.Identity))}" />` : ''}</td>
      <td>${escapeHtml(String(p.Kind || ''))}</td>
      <td>${escapeHtml(String(p.Name || p.Error || ''))}</td>
      <td>${escapeHtml(value)}</td>
    </tr>`;
  });
  renderGenericPickTable(div, '.containmentOrgfwdWrap', '.containmentOrgfwdTable tbody', 'containmentOrgfwdPick', 4, 'No remote domains or outbound spam policies.', html);
}

function applyJunkResult(div, entries) {
  rememberContainmentResult(div, 'junk', entries);
  const html = asArray(entries).filter((e) => (e.Address && e.Address !== '(none)') || e.Error).map((e) => `<tr>
      <td>${e.Address && e.Address !== '(none)' ? `<input type="checkbox" class="containmentJunkPick" data-user="${escapeHtml(String(e.UserPrincipalName || ''))}" data-list="${escapeHtml(String(e.List || ''))}" data-address="${escapeHtml(String(e.Address || ''))}" />` : ''}</td>
      <td>${escapeHtml(String(e.UserPrincipalName || ''))}</td>
      <td>${escapeHtml(String(e.List || ''))}</td>
      <td>${escapeHtml(String(e.Address || e.Error || ''))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentJunkWrap', '.containmentJunkTable tbody', 'containmentJunkPick', 4, 'No trusted senders or recipients.', html);
}

function applyJournalResult(div, rules) {
  rememberContainmentResult(div, 'journal', rules);
  const html = asArray(rules).filter((r) => r.Identity || r.Error).map((r) => `<tr>
      <td>${r.Identity ? `<input type="checkbox" class="containmentJournalPick" data-identity="${escapeHtml(String(r.Identity))}" />` : ''}</td>
      <td>${escapeHtml(String(r.Name || r.Error || ''))}</td>
      <td>${escapeHtml(String(r.Recipient || ''))}</td>
      <td>${escapeHtml(String(r.JournalEmailAddress || ''))}</td>
      <td>${r.Enabled ? 'yes' : 'no'}</td>
      <td>${escapeHtml(containmentDateText(r.WhenChanged, r.WhenChangedUTC))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentJournalWrap', '.containmentJournalTable tbody', 'containmentJournalPick', 6, 'No journal rules.', html);
}

function applyHoldResult(div, users) {
  rememberContainmentResult(div, 'hold', users);
  const wrap = div.querySelector('.containmentHoldWrap');
  const tbody = div.querySelector('.containmentHoldTable tbody');
  if (!wrap || !tbody) return;
  const list = asArray(users);
  tbody.innerHTML = list.length
    ? list.map((u) => `<tr>
      <td>${escapeHtml(String(u.UserPrincipalName || ''))}</td>
      <td>${u.Error ? escapeHtml(String(u.Error)) : (u.LitigationHoldEnabled ? 'yes' : 'no')}</td>
      <td>${escapeHtml(containmentDateText(u.LitigationHoldDate))}</td>
      <td>${escapeHtml(String(u.RetainDeletedItemsFor || ''))}</td>
      <td>${u.AuditEnabled ? 'yes' : 'no'}</td>
    </tr>`).join('')
    : '<tr><td colspan="5" class="muted">No mailbox hold data.</td></tr>';
  wrap.style.display = 'block';
  wireContainmentTableSort(div.querySelector('.containmentHoldTable'));
}

function applyElsewhereResult(div, grants) {
  rememberContainmentResult(div, 'elsewhere', grants);
  const html = asArray(grants).filter((g) => (g.Mailbox && g.Mailbox !== '(none)') || g.Error).map((g) => `<tr>
      <td>${g.Mailbox && g.Mailbox !== '(none)' && !g.Error ? `<input type="checkbox" class="containmentElsewherePick" data-user="${escapeHtml(String(g.UserPrincipalName || ''))}" data-mailbox="${escapeHtml(String(g.Mailbox || ''))}" data-right="${escapeHtml(String(g.Right || ''))}" data-trustee="${escapeHtml(String(g.Trustee || g.UserPrincipalName || ''))}" />` : ''}</td>
      <td>${escapeHtml(String(g.UserPrincipalName || ''))}</td>
      <td>${escapeHtml(String(g.Mailbox || g.Error || ''))}</td>
      <td>${escapeHtml(String(g.Right || ''))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentElsewhereWrap', '.containmentElsewhereTable tbody', 'containmentElsewherePick', 4, 'No Send As, Send on Behalf, or Full Access grants on other mailboxes.', html);
}

function applyRolesResult(div, roles) {
  rememberContainmentResult(div, 'roles', roles);
  const html = asArray(roles).filter((r) => r.Id || r.Error).map((r) => `<tr>
      <td>${r.Id && r.CanRemove !== false ? `<input type="checkbox" class="containmentRolePick" data-id="${escapeHtml(String(r.Id))}" data-kind="${escapeHtml(String(r.Kind || ''))}" data-user="${escapeHtml(String(r.UserPrincipalName || ''))}" data-name="${escapeHtml(String(r.Name || ''))}" />` : ''}</td>
      <td>${escapeHtml(String(r.UserPrincipalName || ''))}</td>
      <td>${escapeHtml(String(r.Kind || ''))}</td>
      <td>${escapeHtml(String(r.Name || r.Error || ''))}</td>
      <td>${escapeHtml(String(r.Details || ''))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentRolesWrap', '.containmentRolesTable tbody', 'containmentRolePick', 5, 'No directory roles, groups, or Exchange RBAC assignments.', html);
}

function applyAppcredsResult(div, creds) {
  rememberContainmentResult(div, 'appcreds', creds);
  const html = asArray(creds).filter((c) => c.AppId || c.Error).map((c) => `<tr>
      <td>${c.AppId && c.Kind && c.Kind !== 'Certificate' ? `<input type="checkbox" class="containmentAppcredPick" data-kind="${escapeHtml(String(c.Kind))}" data-appid="${escapeHtml(String(c.AppId))}" data-keyid="${escapeHtml(String(c.KeyId || ''))}" data-ownerid="${escapeHtml(String(c.OwnerId || ''))}" />` : ''}</td>
      <td>${escapeHtml(String(c.Kind || ''))}</td>
      <td>${escapeHtml(String(c.AppName || ''))}</td>
      <td>${escapeHtml(String(c.DisplayName || c.Error || ''))}</td>
      <td>${escapeHtml(containmentDateText(c.Created, c.startDateTime, c.StartDateTime))}</td>
      <td>${escapeHtml(containmentDateText(c.End, c.endDateTime, c.EndDateTime))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentAppcredsWrap', '.containmentAppcredsTable tbody', 'containmentAppcredPick', 6, 'No app secrets or owners (certificates are listed but must be removed in Entra).', html);
}

function applyFlowsResult(div, flows) {
  rememberContainmentResult(div, 'flows', flows);
  const html = asArray(flows).filter((f) => f.Id || f.Error).map((f) => `<tr>
      <td>${f.Id ? `<input type="checkbox" class="containmentFlowPick" data-id="${escapeHtml(String(f.Id))}" data-env="${escapeHtml(String(f.Environment || ''))}" />` : ''}</td>
      <td>${escapeHtml(String(f.Name || f.Error || ''))}</td>
      <td>${escapeHtml(String(f.Environment || ''))}</td>
      <td>${f.Enabled === true ? 'yes' : (f.Enabled === false ? 'no' : '')}</td>
      <td>${escapeHtml(containmentDateText(f.Created, f.CreatedTime, f.lastModifiedTime))}</td>
    </tr>`);
  renderGenericPickTable(div, '.containmentFlowsWrap', '.containmentFlowsTable tbody', 'containmentFlowPick', 5, 'No flows returned.', html);
}

async function sendRemediateCommand(clientNumber, div, command, progressLabel, waitSeconds = 180) {
  if (!await requireLiveWorker(clientNumber, { actionLabel: progressLabel || 'containment' })) {
    return null;
  }
  const initial = await api(`/api/tenants/${clientNumber}/command`, {
    method: 'POST',
    body: workerCommandBody(command),
  });
  let final = extractWorkerTokenResponse(initial.response, ['REMEDIATE_FAILED:', ...REMEDIATE_SUCCESS_PREFIXES, 'REMEDIATE_STARTED']);
  if (!final || final === 'REMEDIATE_STARTED') {
    final = await pollWorkerResponse(
      clientNumber,
      'REMEDIATE_STARTED',
      REMEDIATE_SUCCESS_PREFIXES,
      'REMEDIATE_FAILED:',
      waitSeconds,
      progressLabel || 'containment'
    );
  }
  return final;
}

const CA_SUCCESS_PREFIXES = ['CA_POLICIES:', 'CA_LOCATIONS:', 'CA_POLICY:', 'CA_USERS:'];

function parseCaPayload(final) {
  const prefixes = ['CA_FAILED:', ...CA_SUCCESS_PREFIXES];
  const text = extractWorkerTokenResponse(final || '', prefixes);
  for (const prefix of prefixes) {
    if (text.startsWith(prefix)) {
      const rest = text.slice(prefix.length);
      try {
        return { prefix, data: JSON.parse(rest), raw: rest };
      } catch {
        return { prefix, data: null, raw: rest };
      }
    }
  }
  return { prefix: '', data: null, raw: text };
}

function rememberCaCatalog(div, patch) {
  const clientNumber = div?.closest('details.tenant')?.dataset?.client;
  if (!clientNumber || !patch) return;
  const prior = tenantUiState.get(String(clientNumber)) || {};
  tenantUiState.set(String(clientNumber), { ...prior, caCatalog: { ...(prior.caCatalog || {}), ...patch } });
}

function setCaCatalogStatus(div, text, tone) {
  const el = div.querySelector('.caCatalogStatus');
  if (!el) return;
  el.textContent = text || '';
  el.className = 'caCatalogStatus' + (text ? ` hasStatus ${tone || ''}` : '');
}

function caStateChip(state) {
  const s = String(state || '');
  let cls = 'muted';
  let label = s || 'unknown';
  if (/^enabled$/i.test(s)) { cls = 'ok'; label = 'enabled'; }
  else if (/report/i.test(s)) { cls = 'report'; label = 'report-only'; }
  else if (/disabled/i.test(s)) { cls = 'muted'; label = 'disabled'; }
  return `<span class="caStateChip ${cls}">${escapeHtml(label)}</span>`;
}

const CA_ISO_COUNTRIES = 'AD,AE,AF,AG,AI,AL,AM,AO,AQ,AR,AS,AT,AU,AW,AX,AZ,BA,BB,BD,BE,BF,BG,BH,BI,BJ,BL,BM,BN,BO,BQ,BR,BS,BT,BV,BW,BY,BZ,CA,CC,CD,CF,CG,CH,CI,CK,CL,CM,CN,CO,CR,CU,CV,CW,CX,CY,CZ,DE,DJ,DK,DM,DO,DZ,EC,EE,EG,EH,ER,ES,ET,FI,FJ,FK,FM,FO,FR,GA,GB,GD,GE,GF,GG,GH,GI,GL,GM,GN,GP,GQ,GR,GS,GT,GU,GW,GY,HK,HM,HN,HR,HT,HU,ID,IE,IL,IM,IN,IO,IQ,IR,IS,IT,JE,JM,JO,JP,KE,KG,KH,KI,KM,KN,KP,KR,KW,KY,KZ,LA,LB,LC,LI,LK,LR,LS,LT,LU,LV,LY,MA,MC,MD,ME,MF,MG,MH,MK,ML,MM,MN,MO,MP,MQ,MR,MS,MT,MU,MV,MW,MX,MY,MZ,NA,NC,NE,NF,NG,NI,NL,NO,NP,NR,NU,NZ,OM,PA,PE,PF,PG,PH,PK,PL,PM,PN,PR,PS,PT,PW,PY,QA,RE,RO,RS,RU,RW,SA,SB,SC,SD,SE,SG,SH,SI,SJ,SK,SL,SM,SN,SO,SR,SS,ST,SV,SX,SY,SZ,TC,TD,TF,TG,TH,TJ,TK,TL,TM,TN,TO,TR,TT,TV,TW,TZ,UA,UG,UM,US,UY,UZ,VA,VC,VE,VG,VI,VN,VU,WF,WS,YE,YT,ZA,ZM,ZW'.split(',');

function isCaAllowlistName(name) {
  return /allow|trust|permit|safe|whitelist/i.test(String(name || ''));
}

function parseCaCountryCodes(loc) {
  const fromField = asArray(loc?.countries).map((c) => String(c || '').trim().toUpperCase()).filter((c) => /^[A-Z]{2}$/.test(c));
  if (fromField.length) return fromField;
  const detail = String(loc?.detail || '');
  if (/^Allowed:/i.test(detail) || /range|\//i.test(detail)) return [];
  return (detail.match(/\b[A-Z]{2}\b/g) || []).filter((c) => CA_ISO_COUNTRIES.includes(c));
}

function invertCaCountryCodes(blocked) {
  const set = new Set(asArray(blocked).map((c) => String(c || '').trim().toUpperCase()).filter(Boolean));
  if (!set.size) return [];
  return CA_ISO_COUNTRIES.filter((c) => !set.has(c));
}

function withCaAllowedCountries(loc) {
  if (!loc || typeof loc !== 'object') return loc;
  const existing = asArray(loc.allowedCountries).filter(Boolean);
  if (existing.length) return loc;
  if (/IP range/i.test(loc.type || '')) return loc;
  if (isCaAllowlistName(loc.displayName || loc.label)) return loc;
  const blocked = parseCaCountryCodes(loc);
  if (!blocked.length) return loc;
  const allowed = invertCaCountryCodes(blocked);
  return {
    ...loc,
    countries: blocked,
    allowedCountries: allowed,
    countryMode: 'blocklist',
    type: loc.type && loc.type !== 'Location' ? loc.type : 'Country blocklist',
    detail: allowed.length ? `Allowed: ${allowed.join(', ')}` : 'Allowed: (none)',
  };
}

function formatCaLocationRef(item) {
  if (item == null || item === '') return '';
  if (typeof item !== 'object') {
    const raw = String(item);
    if (raw === 'AllTrusted') return 'All trusted IPs';
    if (raw === 'All') return 'All';
    return raw;
  }
  const loc = withCaAllowedCountries(item);
  const name = loc.label || loc.displayName || loc.id || '';
  const allowed = asArray(loc.allowedCountries).filter(Boolean);
  if (allowed.length) return `${name} (allowed: ${allowed.join(', ')})`;
  const cidrs = asArray(loc.ipRanges).filter(Boolean);
  const extra = cidrs.length ? cidrs.join(', ') : (loc.detail && /IP range|blocklist/i.test(loc.type || '') ? loc.detail : '');
  return extra ? `${name} (${extra})` : (loc.detail ? `${name} (${loc.detail})` : name);
}

function formatCaLocationList(items) {
  const labels = asArray(items).map(formatCaLocationRef).filter(Boolean);
  return labels.length ? labels.join(', ') : '—';
}

function formatCaLocationSummary(policy) {
  const inc = asArray(policy?.includeLocations || policy?.includeLocationIds);
  const exc = asArray(policy?.excludeLocations || policy?.excludeLocationIds);
  const incIps = asArray(policy?.includeIpRanges);
  const excIps = asArray(policy?.excludeIpRanges);
  const parts = [];
  if (inc.length) parts.push(`Include: ${formatCaLocationList(inc)}`);
  else if (incIps.length) parts.push(`Include IPs: ${incIps.join(', ')}`);
  if (exc.length) parts.push(`Exclude: ${formatCaLocationList(exc)}`);
  else if (excIps.length) parts.push(`Exclude IPs: ${excIps.join(', ')}`);
  return parts.join(' | ') || '—';
}

const CA_GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const CA_WELL_KNOWN_ROLES = {
  '62e90394-69f5-4237-9190-012177145e10': 'Global Administrator',
  '9b895d77-7237-7278-86d2-56d3bd9720bd': 'Application Administrator',
  '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3': 'Application Administrator',
  'c4e39bd9-1100-46d3-8c65-fb160da0071f': 'Authentication Administrator',
  '0526716b-113d-4c15-b2c8-68e3c22b9f80': 'Authentication Policy Administrator',
  'b0f54661-2d74-4c50-afa3-1ec803f12efe': 'Billing Administrator',
  '158c047a-c907-4556-b7ef-446551a6b5f7': 'Cloud Application Administrator',
  'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9': 'Conditional Access Administrator',
  '44367163-eba1-44c3-98af-f5787879f96a': 'Dynamics 365 Administrator',
  '29232cdf-9323-42fd-ade2-1d097af3e4de': 'Exchange Administrator',
  'f2ef992c-3afb-46b9-b7cf-a126ee74c451': 'Global Reader',
  '729827e3-9c14-49f7-bb1b-9608f156bbb8': 'Helpdesk Administrator',
  '3a2c62dc-778b-46ae-8129-6b6e3c87330d': 'Intune Administrator',
  '3a2c62db-5318-420d-8d74-23affee5d9d5': 'Intune Administrator',
  '966707d0-3269-4727-9be2-8c5a1d0f0f90': 'Password Administrator',
  '966707d0-3269-4727-9be2-8c3a10f19b9d': 'Password Administrator',
  '1169501e-1cda-4adb-99d3-e6c394ecb116': 'Power Platform Administrator',
  '7be44c8a-adaf-4e2a-884d-2d2b0dec86ba': 'Privileged Authentication Administrator',
  '7be44c8a-adaf-4e2a-84d6-ab2649e08a13': 'Privileged Authentication Administrator',
  'e8611ab8-c189-46e8-94e1-8229d6d9673b': 'Privileged Role Administrator',
  'e8611ab8-c189-46e8-94e1-60213ab1f814': 'Privileged Role Administrator',
  '194ae4cb-b126-40b2-bd5b-6091b380977d': 'Security Administrator',
  '5d6b6bb7-de71-4623-b4af-96380a352509': 'Security Reader',
  'f28a1f50-f6e7-4571-818b-6a12f2af6b6c': 'SharePoint Administrator',
  '69091246-20e8-4a56-aa4d-066075b2a7a8': 'Teams Administrator',
  'fe930be7-5e62-47db-91af-98c3a49a38b1': 'User Administrator',
};

function caWellKnownRoleName(id) {
  if (!id) return '';
  return CA_WELL_KNOWN_ROLES[String(id)] || CA_WELL_KNOWN_ROLES[String(id).toLowerCase()] || '';
}

function formatCaPrincipalLabel(item) {
  if (item == null || item === '') return '';
  const id = (item && typeof item === 'object') ? String(item.id || '') : String(item);
  let label = (item && typeof item === 'object')
    ? String(item.label || item.displayName || item.id || '')
    : String(item);
  const role = caWellKnownRoleName(label) || caWellKnownRoleName(id);
  if (role) return role;
  if (id && new RegExp(`\\s*\\(${id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\)\\s*$`, 'i').test(label)) {
    label = label.replace(new RegExp(`\\s*\\(${id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\)\\s*$`, 'i'), '').trim();
  }
  return label;
}

function formatCaUserScope(policy) {
  const bits = [];
  const incU = formatCaNamedList(policy?.includeUsers);
  const excU = formatCaNamedList(policy?.excludeUsers);
  const incG = formatCaNamedList(policy?.includeGroups);
  const excG = formatCaNamedList(policy?.excludeGroups);
  const incR = formatCaNamedList(policy?.includeRoles);
  const excR = formatCaNamedList(policy?.excludeRoles);
  if (incU !== '—') bits.push(`Include users: ${incU}`);
  if (incG !== '—') bits.push(`Include groups: ${incG}`);
  if (incR !== '—') bits.push(`Include roles: ${incR}`);
  if (excU !== '—') bits.push(`Exclude users: ${excU}`);
  if (excG !== '—') bits.push(`Exclude groups: ${excG}`);
  if (excR !== '—') bits.push(`Exclude roles: ${excR}`);
  if (bits.length) return bits.join(' | ');
  return String(policy?.userScope || '') || '—';
}

function formatCaExclusionSummary(policy) {
  const bits = [];
  const users = formatCaNamedList(policy?.excludeUsers);
  const groups = formatCaNamedList(policy?.excludeGroups);
  const roles = formatCaNamedList(policy?.excludeRoles);
  if (users !== '—') bits.push(`Users: ${users}`);
  if (groups !== '—') bits.push(`Groups: ${groups}`);
  if (roles !== '—') bits.push(`Roles: ${roles}`);
  if (!bits.length && policy?.exclusions && policy.exclusions !== '—') bits.push(policy.exclusions);
  const excLoc = asArray(policy?.excludeLocations);
  if (excLoc.length) bits.push(`Locations: ${formatCaLocationList(excLoc)}`);
  return bits.length ? [...new Set(bits)].join(' | ') : '—';
}

function formatCaNamedList(items) {
  const labels = asArray(items).map((x) => formatCaPrincipalLabel(x)).filter(Boolean);
  return labels.length ? labels.join(', ') : '—';
}

function mergeCaNamedList(preferred, fallback) {
  const byId = new Map();
  asArray(fallback).forEach((item) => {
    const id = String((item && typeof item === 'object') ? (item.id || '') : item).toLowerCase();
    if (id) byId.set(id, item);
  });
  const source = asArray(preferred).length ? preferred : fallback;
  return asArray(source).map((item) => {
    const id = String((item && typeof item === 'object') ? (item.id || '') : item);
    const cached = byId.get(id.toLowerCase());
    const label = formatCaPrincipalLabel(item);
    const cachedLabel = formatCaPrincipalLabel(cached);
    const useCached = cachedLabel && !CA_GUID_RE.test(cachedLabel) && (CA_GUID_RE.test(label) || !label);
    return {
      id: id || (cached && cached.id) || '',
      label: useCached ? cachedLabel : (label || cachedLabel || id),
    };
  });
}

function mergeCaPolicyNames(detail, listed) {
  if (!detail) return detail;
  if (!listed) return detail;
  return {
    ...listed,
    ...detail,
    includeUsers: mergeCaNamedList(detail.includeUsers, listed.includeUsers),
    excludeUsers: mergeCaNamedList(detail.excludeUsers, listed.excludeUsers),
    includeGroups: mergeCaNamedList(detail.includeGroups, listed.includeGroups),
    excludeGroups: mergeCaNamedList(detail.excludeGroups, listed.excludeGroups),
    includeRoles: mergeCaNamedList(detail.includeRoles, listed.includeRoles),
    excludeRoles: mergeCaNamedList(detail.excludeRoles, listed.excludeRoles),
  };
}

function caJsonCommand(name, payload) {
  return `${name}|${JSON.stringify(payload || {})}`;
}

function confirmCaWrite(message) {
  return window.confirm(message);
}

function confirmCaEnabledPolicy(policy, action) {
  const state = String(policy?.state || '');
  if (!/^enabled$/i.test(state)) return true;
  return confirmCaWrite(`Policy '${policy.displayName || policy.id}' is ENABLED.\n\n${action}\n\nContinue?`);
}

function caUserItems(policy, key) {
  return asArray(policy?.[key]).map((item) => {
    const id = String((item && typeof item === 'object') ? (item.id || '') : item);
    return { id, label: formatCaPrincipalLabel(item) || id };
  }).filter((x) => x.id);
}

function renderCaUserCheckList(items, className) {
  const rows = asArray(items).filter((x) => x.id && x.id !== 'All' && x.id !== 'GuestsOrExternalUsers' && x.id !== 'None');
  if (!rows.length) return '<div class="muted">None</div>';
  return `<div class="caCatalogUserList">${rows.map((x) => `
    <label class="caCatalogUserRow">
      <input type="checkbox" class="${className}" value="${escapeHtml(x.id)}" />
      <span>${escapeHtml(x.label)}</span>
    </label>`).join('')}</div>`;
}

function rememberCaAudit(div, entry) {
  if (!div || !entry) return;
  const clientNumber = div.closest('details.tenant')?.dataset?.client;
  if (!clientNumber) return;
  const prior = tenantUiState.get(String(clientNumber)) || {};
  const catalog = { ...(prior.caCatalog || {}) };
  catalog.audit = asArray(catalog.audit).concat([{
    Timestamp: entry.Timestamp || new Date().toISOString(),
    UPN: String(entry.UPN || entry.Upn || ''),
    Action: String(entry.Action || ''),
    Result: String(entry.Result || 'success'),
    Detail: String(entry.Detail || '').slice(0, 500),
  }]);
  tenantUiState.set(String(clientNumber), { ...prior, caCatalog: catalog });
}

function selectedCaUserIds(root, selector) {
  return [...(root?.querySelectorAll(`${selector}:checked`) || [])].map((el) => el.value).filter(Boolean);
}

function grantControlsText(grant) {
  if (!grant) return '—';
  const builtIn = asArray(grant.builtInControls || grant.BuiltInControls);
  const custom = asArray(grant.customAuthenticationFactors || grant.CustomAuthenticationFactors);
  const tou = asArray(grant.termsOfUse || grant.TermsOfUse);
  const op = grant.operator || grant.Operator || '';
  const parts = [];
  if (op) parts.push(`Operator: ${op}`);
  if (builtIn.length) parts.push(builtIn.join(', '));
  if (custom.length) parts.push(`Custom: ${custom.join(', ')}`);
  if (tou.length) parts.push(`Terms of use: ${tou.length}`);
  return parts.join(' | ') || '—';
}

function sessionControlsText(session) {
  if (!session || typeof session !== 'object') return '—';
  const keys = Object.keys(session).filter((k) => {
    if (k.startsWith('@')) return false;
    const v = session[k];
    return v != null && v !== false && v !== '';
  });
  return keys.length ? keys.join(', ') : '—';
}

function updateCaCatalogButtons(div) {
  if (!div) return;
  const graph = div.dataset.graphAuthenticated === '1';
  div.querySelectorAll('.caCatalogRefresh, .caCatalogHardRefresh').forEach((btn) => {
    btn.disabled = !graph;
  });
  const hint = div.querySelector('.caCatalogHint');
  if (hint) {
    hint.textContent = graph
      ? 'Graph Auth connected. Refresh uses the worker cache. Hard refresh reloads from Graph.'
      : 'Complete Graph Auth to list current policies.';
  }
}

function renderCaPolicies(div, policies) {
  const tbody = div.querySelector('.caCatalogPoliciesTable tbody');
  if (!tbody) return;
  const rows = asArray(policies);
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="muted">No Conditional Access policies.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map((p) => {
    const id = String(p.id || '');
    return `<tr class="caCatalogPolicyRow" data-policy-id="${escapeHtml(id)}" title="Open policy detail">
      <td>${escapeHtml(p.displayName || id)}</td>
      <td>${caStateChip(p.state)}</td>
      <td>${escapeHtml(formatCaUserScope(p))}</td>
      <td>${escapeHtml(formatCaExclusionSummary(p))}</td>
      <td>${escapeHtml(formatCaLocationSummary(p))}</td>
    </tr>`;
  }).join('');
}

function renderCaLocations(div, locations) {
  const tbody = div.querySelector('.caCatalogLocationsTable tbody');
  if (!tbody) return;
  const rows = asArray(locations);
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="3" class="muted">No named locations.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map((raw) => {
    const l = withCaAllowedCountries(raw);
    const allowed = asArray(l.allowedCountries).filter(Boolean);
    const cidrs = asArray(l.ipRanges).filter(Boolean);
    const detail = allowed.length
      ? `Allowed: ${allowed.join(', ')}`
      : (cidrs.length ? cidrs.join(', ') : (l.detail || ''));
    return `<tr>
    <td>${escapeHtml(l.displayName || l.id || '')}</td>
    <td>${escapeHtml(l.type || '')}</td>
    <td>${escapeHtml(detail)}</td>
  </tr>`;
  }).join('');
}

function renderCaPolicyDetail(div, policy) {
  const drawer = div.querySelector('.caCatalogDrawer');
  if (!drawer) return;
  if (!policy) {
    drawer.hidden = true;
    drawer.innerHTML = '';
    return;
  }
  drawer.hidden = false;
  const includeUsers = caUserItems(policy, 'includeUsers');
  const excludeUsers = caUserItems(policy, 'excludeUsers');
  const includesAll = includeUsers.some((x) => x.id === 'All');
  drawer.innerHTML = `
    <h3>${escapeHtml(policy.displayName || policy.id || 'Policy')}</h3>
    <p class="muted" style="margin:0 0 0.35rem">${caStateChip(policy.state)} <span>${escapeHtml(formatCaUserScope(policy))}</span></p>
    <ul>
      <li>Include users: ${escapeHtml(formatCaNamedList(policy.includeUsers))}</li>
      <li>Exclude users: ${escapeHtml(formatCaNamedList(policy.excludeUsers))}</li>
      <li>Include groups: ${escapeHtml(formatCaNamedList(policy.includeGroups))}</li>
      <li>Exclude groups: ${escapeHtml(formatCaNamedList(policy.excludeGroups))}</li>
      <li>Include roles: ${escapeHtml(formatCaNamedList(policy.includeRoles))}</li>
      <li>Exclude roles: ${escapeHtml(formatCaNamedList(policy.excludeRoles))}</li>
      <li>Include locations: ${escapeHtml(formatCaLocationList(policy.includeLocations))}</li>
      <li>Exclude locations: ${escapeHtml(formatCaLocationList(policy.excludeLocations))}</li>
      <li>Grant: ${escapeHtml(grantControlsText(policy.grantControls))}</li>
      <li>Session: ${escapeHtml(sessionControlsText(policy.sessionControls))}</li>
    </ul>
    <div class="caCatalogEdit" data-policy-id="${escapeHtml(policy.id || '')}">
      <h4>User exceptions</h4>
      <p class="caCatalogEditHint">Search does not page the whole tenant. Groups and roles stay read-only this increment. Enabled policies ask for confirm.</p>
      <div class="caCatalogSubLabel">Included users${includesAll ? ' (All users)' : ''}</div>
      ${includesAll ? '<div class="muted">All users. Adding specific accounts replaces All.</div>' : renderCaUserCheckList(includeUsers, 'caIncludeUser')}
      <div class="row">
        <button type="button" class="caRemoveInclude small" ${includesAll ? 'disabled' : ''}>Remove selected include</button>
      </div>
      <div class="caCatalogSubLabel">Excluded users</div>
      ${renderCaUserCheckList(excludeUsers, 'caExcludeUser')}
      <div class="row">
        <button type="button" class="caRemoveExclude small">Remove selected exclusions</button>
      </div>
      <div class="caCatalogSubLabel">Find users</div>
      <div class="caCatalogSearchRow">
        <input type="search" class="caUserSearch" placeholder="Name or email (2+ characters)" autocomplete="off" />
        <button type="button" class="caUserSearchBtn small">Search</button>
      </div>
      <div class="caCatalogSearchHits muted">Search to add people to include or exclude.</div>
      <div class="row">
        <button type="button" class="caAddInclude small">Add selected to include</button>
        <button type="button" class="caAddExclude small">Add selected to exclusions</button>
      </div>
    </div>`;
}

function restoreCaCatalogOutput(div, state) {
  const c = state?.caCatalog;
  if (!div || !c || typeof c !== 'object') return;
  if (c.policies) renderCaPolicies(div, c.policies);
  if (c.locations) renderCaLocations(div, c.locations);
  if (c.detail) renderCaPolicyDetail(div, c.detail);
  if (c.status) setCaCatalogStatus(div, c.status, c.statusTone || '');
}

async function sendCaCommand(clientNumber, div, command, progressLabel, waitSeconds = 180) {
  if (!await requireLiveWorker(clientNumber, { actionLabel: progressLabel || 'Conditional Access' })) {
    return null;
  }
  const initial = await api(`/api/tenants/${clientNumber}/command`, {
    method: 'POST',
    body: workerCommandBody(command),
  });
  let final = extractWorkerTokenResponse(initial.response, ['CA_FAILED:', ...CA_SUCCESS_PREFIXES, 'CA_STARTED']);
  if (!final || final === 'CA_STARTED') {
    final = await pollWorkerResponse(
      clientNumber,
      'CA_STARTED',
      CA_SUCCESS_PREFIXES,
      'CA_FAILED:',
      waitSeconds,
      progressLabel || 'Conditional Access'
    );
  }
  return final;
}

async function refreshCaCatalog(clientNumber, div, force) {
  if (div.dataset.graphAuthenticated !== '1') {
    const msg = `Client ${clientNumber}: Graph Auth is required to list Conditional Access policies.`;
    log(msg);
    window.alert(msg);
    return;
  }
  const suffix = force ? '|REFRESH:1' : '';
  if (force) rememberCaCatalog(div, { policies: [], locations: [], detail: null });
  setCaCatalogStatus(div, force ? 'Reloading from Graph…' : 'Loading Conditional Access catalog…', 'statusBusy');
  const panel = div.querySelector('details.caCatalogPanel');
  if (panel) panel.open = true;
  const policiesFinal = await sendCaCommand(clientNumber, div, `CA_LIST_POLICIES${suffix}`, 'listing CA policies');
  const policiesParsed = parseCaPayload(policiesFinal || '');
  if (policiesParsed.prefix !== 'CA_POLICIES:') {
    throw new Error(policiesParsed.raw || 'CA_LIST_POLICIES failed');
  }
  const policies = asArray(policiesParsed.data?.Policies);
  renderCaPolicies(div, policies);

  const locationsFinal = await sendCaCommand(clientNumber, div, `CA_LIST_LOCATIONS${suffix}`, 'listing CA locations');
  const locationsParsed = parseCaPayload(locationsFinal || '');
  if (locationsParsed.prefix !== 'CA_LOCATIONS:') {
    throw new Error(locationsParsed.raw || 'CA_LIST_LOCATIONS failed');
  }
  const locations = asArray(locationsParsed.data?.Locations);
  renderCaLocations(div, locations);
  const status = `${policies.length} polic${policies.length === 1 ? 'y' : 'ies'}, ${locations.length} location${locations.length === 1 ? '' : 's'}${force ? '' : ' (cached after first load)'}`;
  setCaCatalogStatus(div, status, 'statusOk');
  rememberCaCatalog(div, { policies, locations, status, statusTone: 'statusOk' });
}

function currentCaDetail(div) {
  const clientNumber = div?.closest('details.tenant')?.dataset?.client;
  return tenantUiState.get(String(clientNumber))?.caCatalog?.detail || null;
}

function applyCaPolicyResult(clientNumber, div, parsed, statusText) {
  const policy = parsed?.data?.Policy || parsed?.data;
  if (!policy) throw new Error(parsed?.raw || 'CA policy payload missing');
  const listed = asArray(tenantUiState.get(String(clientNumber))?.caCatalog?.policies)
    .find((p) => String(p?.id || '') === String(policy.id || ''));
  const merged = mergeCaPolicyNames(policy, listed);
  renderCaPolicyDetail(div, merged);
  rememberCaCatalog(div, { detail: merged });
  if (parsed?.data?.Audit) rememberCaAudit(div, parsed.data.Audit);
  if (statusText) setCaCatalogStatus(div, statusText, 'statusOk');
  return merged;
}

async function loadCaPolicyDetail(clientNumber, div, policyId) {
  if (!policyId) return;
  setCaCatalogStatus(div, 'Loading policy detail…', 'statusBusy');
  const final = await sendCaCommand(clientNumber, div, `CA_GET_POLICY|POLICY_ID:${policyId}`, 'loading CA policy', 120);
  const parsed = parseCaPayload(final || '');
  if (parsed.prefix !== 'CA_POLICY:') {
    throw new Error(parsed.raw || 'CA_GET_POLICY failed');
  }
  applyCaPolicyResult(clientNumber, div, parsed, `Opened ${parsed.data?.Policy?.displayName || policyId}`);
}

async function patchCaPolicyUsers(clientNumber, div, scope, op, ids, allowAllFallback) {
  const policy = currentCaDetail(div);
  const policyId = policy?.id;
  if (!policyId) throw new Error('Open a policy first.');
  const wanted = asArray(ids).filter(Boolean);
  if (!wanted.length) throw new Error('Select at least one user.');
  const replacingAll = scope === 'include' && op === 'add' && caUserItems(policy, 'includeUsers').some((x) => x.id === 'All');
  if (replacingAll && !confirmCaWrite(`Policy '${policy.displayName}' currently includes All users.\n\nAdding specific accounts replaces All with only the selected people.\n\nContinue?`)) {
    return;
  }
  const emptyingInclude = scope === 'include' && op === 'remove' && !allowAllFallback;
  const remaining = caUserItems(policy, 'includeUsers').filter((x) => x.id !== 'All' && !wanted.includes(x.id));
  const hasOther = caUserItems(policy, 'includeGroups').length || caUserItems(policy, 'includeRoles').length;
  if (emptyingInclude && remaining.length === 0 && !hasOther) {
    if (!confirmCaWrite(`Removing the last included user leaves no include target.\n\nFallback to All users?\n\nContinue?`)) return;
    allowAllFallback = true;
  }
  const action = `${op === 'add' ? 'Add' : 'Remove'} ${wanted.length} user(s) on ${scope}.`;
  if (!confirmCaEnabledPolicy(policy, action)) return;
  setCaCatalogStatus(div, 'Updating policy users…', 'statusBusy');
  const final = await sendCaCommand(
    clientNumber,
    div,
    caJsonCommand('CA_PATCH_USERS', { policyId, scope, op, ids: wanted, allowAllFallback: Boolean(allowAllFallback) }),
    'updating CA users',
    180
  );
  const parsed = parseCaPayload(final || '');
  if (parsed.prefix !== 'CA_POLICY:') {
    throw new Error(parsed.raw || 'CA_PATCH_USERS failed');
  }
  applyCaPolicyResult(clientNumber, div, parsed, `Updated users on ${parsed.data?.Policy?.displayName || policyId}`);
  log(`Client ${clientNumber}: CA ${op} ${scope} ${wanted.length} user(s) on ${policy.displayName || policyId}.`);
  try { await refreshCaCatalog(clientNumber, div, false); } catch { /* list refresh is best-effort */ }
}

async function searchCaCatalogUsers(clientNumber, div) {
  const input = div.querySelector('.caUserSearch');
  const hits = div.querySelector('.caCatalogSearchHits');
  const q = String(input?.value || '').trim();
  if (q.length < 2) {
    if (hits) hits.textContent = 'Enter at least 2 characters.';
    return;
  }
  setCaCatalogStatus(div, 'Searching users…', 'statusBusy');
  const final = await sendCaCommand(clientNumber, div, caJsonCommand('CA_SEARCH_USERS', { q }), 'searching CA users', 90);
  const parsed = parseCaPayload(final || '');
  if (parsed.prefix !== 'CA_USERS:') {
    throw new Error(parsed.raw || 'CA_SEARCH_USERS failed');
  }
  const users = asArray(parsed.data?.Users);
  if (hits) {
    hits.classList.remove('muted');
    hits.innerHTML = users.length
      ? users.map((u) => {
          const id = String(u.id || '');
          const label = formatCaPrincipalLabel(u) || u.userPrincipalName || id;
          const dept = u.department ? ` — ${u.department}` : '';
          return `<label class="caCatalogUserRow">
            <input type="checkbox" class="caSearchUser" value="${escapeHtml(id)}" />
            <span>${escapeHtml(label)}${escapeHtml(dept)}</span>
          </label>`;
        }).join('')
      : '<div class="muted">No users matched.</div>';
  }
  setCaCatalogStatus(div, `${users.length} user${users.length === 1 ? '' : 's'} matched`, users.length ? 'statusOk' : 'statusWarn');
}

let caCatalogPopupEscBound = false;

function setCaCatalogPopup(div, open) {
  const panel = div?.querySelector('details.caCatalogPanel');
  const body = div?.querySelector('.caCatalogBody');
  const backdrop = div?.querySelector('.caCatalogBackdrop');
  if (!body) return;
  if (open) {
    if (panel) panel.open = true;
    body.classList.add('caCatalogPopped');
    if (backdrop) backdrop.hidden = false;
  } else {
    body.classList.remove('caCatalogPopped');
    if (backdrop) backdrop.hidden = true;
  }
}

function runCaCatalogRefresh(clientNumber, div, force) {
  withClientLock(clientNumber, async () => {
    focusClientLogTab(clientNumber);
    try {
      await refreshCaCatalog(clientNumber, div, force);
      log(`Client ${clientNumber}: Conditional Access catalog ${force ? 'reloaded from Graph' : 'loaded'}.`);
    } catch (err) {
      const msg = err?.message || String(err);
      log(`Client ${clientNumber}: CA catalog error: ${msg}`);
      setCaCatalogStatus(div, msg, 'statusBad');
      rememberCaCatalog(div, { status: msg, statusTone: 'statusBad' });
    }
  });
}

function wireCaCatalogPanel(div, clientNumber) {
  updateCaCatalogButtons(div);
  restoreCaCatalogOutput(div, tenantUiState.get(String(clientNumber)));
  div.querySelector('.caCatalogRefresh')?.addEventListener('click', (e) => {
    runCaCatalogRefresh(clientNumber, div, Boolean(e.shiftKey));
  });
  div.querySelector('.caCatalogHardRefresh')?.addEventListener('click', () => {
    runCaCatalogRefresh(clientNumber, div, true);
  });
  div.querySelectorAll('.caCatalogOpenPopup').forEach((btn) => {
    btn.addEventListener('click', () => setCaCatalogPopup(div, true));
  });
  div.querySelector('.caCatalogClosePopup')?.addEventListener('click', () => setCaCatalogPopup(div, false));
  div.querySelector('.caCatalogBackdrop')?.addEventListener('click', () => setCaCatalogPopup(div, false));
  if (!caCatalogPopupEscBound) {
    caCatalogPopupEscBound = true;
    document.addEventListener('keydown', (e) => {
      if (e.key !== 'Escape') return;
      document.querySelectorAll('.caCatalogBody.caCatalogPopped').forEach((body) => {
        const host = body.closest('.tenantBody');
        if (host) setCaCatalogPopup(host, false);
      });
    });
  }
  div.querySelector('.caCatalogPoliciesTable')?.addEventListener('click', (e) => {
    const row = e.target.closest('.caCatalogPolicyRow');
    if (!row) return;
    const policyId = row.dataset.policyId;
    withClientLock(clientNumber, async () => {
      focusClientLogTab(clientNumber);
      try {
        await loadCaPolicyDetail(clientNumber, div, policyId);
      } catch (err) {
        const msg = err?.message || String(err);
        log(`Client ${clientNumber}: CA policy detail error: ${msg}`);
        setCaCatalogStatus(div, msg, 'statusBad');
      }
    });
  });
  const runCaEdit = (work) => {
    withClientLock(clientNumber, async () => {
      focusClientLogTab(clientNumber);
      try {
        await work();
      } catch (err) {
        const msg = err?.message || String(err);
        log(`Client ${clientNumber}: CA edit error: ${msg}`);
        setCaCatalogStatus(div, msg, 'statusBad');
      }
    });
  };
  const drawer = div.querySelector('.caCatalogDrawer');
  drawer?.addEventListener('click', (e) => {
    if (e.target.closest('.caUserSearchBtn')) {
      runCaEdit(() => searchCaCatalogUsers(clientNumber, div));
      return;
    }
    if (e.target.closest('.caAddExclude')) {
      runCaEdit(() => patchCaPolicyUsers(clientNumber, div, 'exclude', 'add', selectedCaUserIds(drawer, '.caSearchUser')));
      return;
    }
    if (e.target.closest('.caAddInclude')) {
      runCaEdit(() => patchCaPolicyUsers(clientNumber, div, 'include', 'add', selectedCaUserIds(drawer, '.caSearchUser')));
      return;
    }
    if (e.target.closest('.caRemoveExclude')) {
      runCaEdit(() => patchCaPolicyUsers(clientNumber, div, 'exclude', 'remove', selectedCaUserIds(drawer, '.caExcludeUser')));
      return;
    }
    if (e.target.closest('.caRemoveInclude')) {
      runCaEdit(() => patchCaPolicyUsers(clientNumber, div, 'include', 'remove', selectedCaUserIds(drawer, '.caIncludeUser')));
    }
  });
  drawer?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && e.target.classList?.contains('caUserSearch')) {
      e.preventDefault();
      runCaEdit(() => searchCaCatalogUsers(clientNumber, div));
    }
  });
}

async function runContainmentQuickSteps(clientNumber, div, label, steps, options = {}) {
  const results = [];
  let halted = false;
  for (const step of steps) {
    if (halted) {
      results.push({ step, ok: false, skipped: true, detail: options.skipReason || '' });
      continue;
    }
    setContainmentStatus(div, `${label}: ${step.label}…`);
    const final = await sendRemediateCommand(clientNumber, div, step.cmd, step.label);
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    noteContainmentAction(div, step.kind, step.targets || [], parsed);
    const ok = parsed.prefix === 'REMEDIATE_SUCCESS:';
    const details = asArray(parsed.data?.Details);
    results.push({ step, parsed, ok, detail: details.length ? details.join('\n') : summarizeRemediateFailure(parsed, final || '').text });
    log(`Client ${clientNumber}: ${label} — ${step.label} ${ok ? 'finished' : 'FAILED'}.`);
    if (!ok && options.stopOnFailure) halted = true;
  }
  return results;
}

function formatContainmentQuickSummary(results) {
  return results.map((r) => {
    if (r.skipped) return `SKIPPED — ${r.step.label}${r.detail ? ` (${r.detail})` : ''}`;
    return `${r.ok ? 'OK' : 'FAILED'} — ${r.step.label}${r.detail ? `:\n${r.detail}` : ''}`;
  }).join('\n\n');
}

async function runContainmentAction(clientNumber, div, kind) {
  const users = getSelectedContainmentUsers(div);
  const graph = div.dataset.graphAuthenticated === '1';
  const exo = div.dataset.exchangeAuthenticated === '1';
  const usersJson = JSON.stringify(users);
  const writes = {
    revoke: { cmd: `REMEDIATE_REVOKE_SESSIONS|USERS:${usersJson}`, action: 'revoke', need: 'graph', label: 'revoking sessions' },
    block: { cmd: `REMEDIATE_BLOCK|USERS:${usersJson}`, action: 'block', need: 'graph', label: 'blocking sign-in' },
    unblock: { cmd: `REMEDIATE_UNBLOCK|USERS:${usersJson}`, action: 'unblock', need: 'graph', label: 'unblocking sign-in' },
    unrestrict: { cmd: `REMEDIATE_UNRESTRICT_EMAIL|USERS:${usersJson}`, action: 'unrestrict', need: 'exo', label: 'unrestricting email' },
  };

  if (kind === 'quick-lock') {
    if (!graph) {
      log(`Client ${clientNumber}: Graph Auth is required to block sign-in and revoke sessions.`);
      return;
    }
    if (!users.length) {
      log(`Client ${clientNumber}: select one or more validated users.`);
      return;
    }
    const confirmMsg = `Block sign-in AND revoke all sessions for:\n${users.join('\n')}\n\n`
      + 'Sign-in is blocked first, then every session, refresh token, and MFA session is revoked.\n\n'
      + 'Registered MFA methods are kept (no re-registration) and the password is NOT changed — handle those separately if needed.\n\nContinue?';
    if (!confirmContainmentPopup(confirmMsg)) return;

    // Block before revoke: revoking first leaves a window where a known password can mint a new token.
    const results = await runContainmentQuickSteps(clientNumber, div, 'Quick actions', [
      { kind: 'block', label: 'blocking sign-in', cmd: `REMEDIATE_BLOCK|USERS:${usersJson}`, targets: users },
      { kind: 'revoke', label: 'revoking sessions', cmd: `REMEDIATE_REVOKE_SESSIONS|USERS:${usersJson}`, targets: users },
    ]);
    setContainmentStatus(div, formatContainmentQuickSummary(results), { reveal: true });
    await runContainmentAction(clientNumber, div, 'signin-status');
    return;
  }

  if (kind === 'quick-restore') {
    if (!graph) {
      log(`Client ${clientNumber}: Graph Auth is required to reset the password and unblock sign-in.`);
      return;
    }
    if (!users.length) {
      log(`Client ${clientNumber}: select one or more validated users.`);
      return;
    }
    const resetUrl = containmentPasswordResetUrl(null, users);
    const confirmMsg = `Reset to a random password AND unblock sign-in for:\n${users.join('\n')}\n\n`
      + `The random password is never shown to anyone. Send the user ${resetUrl} so they can set their own.\n\n`
      + 'If the reset fails, the account stays blocked.\n\nContinue?';
    if (!confirmContainmentPopup(confirmMsg)) return;

    // Reset before unblock, and only unblock if the reset worked: restoring access while the
    // attacker still knows the password undoes the containment.
    const results = await runContainmentQuickSteps(clientNumber, div, 'Quick actions', [
      {
        kind: 'reset-password',
        label: 'resetting password (random)',
        cmd: `REMEDIATE_RESET_PASSWORD|USERS:${usersJson}|OPTIONS:${JSON.stringify({ Mode: 'random' })}`,
        targets: users,
      },
      { kind: 'unblock', label: 'unblocking sign-in', cmd: `REMEDIATE_UNBLOCK|USERS:${usersJson}`, targets: users },
    ], { stopOnFailure: true, skipReason: 'password reset failed, so sign-in stays blocked' });

    const resetOk = Boolean(results[0]?.ok);
    const lines = [formatContainmentQuickSummary(results)];
    if (resetOk) {
      const resetUrl = containmentPasswordResetUrl(results[0]?.parsed?.data?.SsprUrl, users);
      lines.push(containmentPasswordResetInstruction(resetUrl));
      log(`Client ${clientNumber}: password reset — send the user ${resetUrl}`);
    }
    setContainmentStatus(div, lines.join('\n\n'), { reveal: true });
    await runContainmentAction(clientNumber, div, 'signin-status');
    return;
  }

  if (kind === 'signin-status') {
    if (!graph || !users.length) return;
    setContainmentStatus(div, 'Checking sign-in status…');
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_SIGNIN_STATUS|USERS:${usersJson}`, 'checking sign-in status');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    if (parsed.prefix === 'REMEDIATE_SUCCESS:') {
      setContainmentUserStatus(div, parsed.data?.Users);
      log(`Client ${clientNumber}: sign-in status updated.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Sign-in status failed.');
    }
    return;
  }

  if (kind === 'reset-password' || kind === 'assign-password') {
    if (!graph || !users.length) return;
    const assigned = (div.querySelector('.containmentAssignPasswordInput')?.value || '');
    const mode = kind === 'assign-password' ? 'assign' : 'random';
    if (mode === 'assign') {
      if (!assigned.trim()) {
        log(`Client ${clientNumber}: enter a password to assign, or use Reset with random password.`);
        return;
      }
      if (assigned.length < 8) {
        log(`Client ${clientNumber}: assigned password must be at least 8 characters.`);
        return;
      }
    }
    const resetUrl = containmentPasswordResetUrl(null, users);
    const confirmMsg = mode === 'assign'
      ? `Set a specific password for:\n${users.join('\n')}\n\nPrefer sending ${resetUrl} instead of this password. Continue?`
      : `Reset password for:\n${users.join('\n')}\n\nThis invalidates the current password so an attacker cannot use it. Do not send a password — send ${resetUrl}.\n\nContinue?`;
    if (!confirmContainmentPopup(confirmMsg)) return;
    const options = mode === 'assign' ? { Mode: 'assign', Password: assigned } : { Mode: 'random' };
    setContainmentStatus(div, mode === 'assign' ? 'Setting assigned password…' : 'Resetting password (random)…');
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_RESET_PASSWORD|USERS:${usersJson}|OPTIONS:${JSON.stringify(options)}`, 'resetting password');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    if (parsed.prefix === 'REMEDIATE_SUCCESS:') {
      const resetUrl = containmentPasswordResetUrl(parsed.data?.SsprUrl, users);
      const lines = [
        ...(Array.isArray(parsed.data?.Details) ? parsed.data.Details : []),
        '',
        containmentPasswordResetInstruction(resetUrl),
      ];
      setContainmentStatus(div, lines.filter((x, i, a) => x !== '' || a[i - 1] !== '').join('\n'), { reveal: true });
      const field = div.querySelector('.containmentAssignPasswordInput');
      if (field) field.value = '';
      log(`Client ${clientNumber}: password reset finished. Send the user ${resetUrl}`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Password reset failed.');
    }
    noteContainmentAction(div, kind, users, parsed);
    return;
  }

  if (kind === 'list-mfa') {
    if (!graph || !users.length) return;
    setContainmentStatus(div, 'Listing MFA methods…');
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_LIST_AUTH_METHODS|USERS:${usersJson}`, 'listing MFA methods');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    if (parsed.prefix === 'REMEDIATE_AUTHMETHODS:') {
      applyAuthMethodsResult(div, parsed.data?.Methods);
      setContainmentStatus(div, `${asArray(parsed.data?.Methods).length} MFA method row(s).`);
      log(`Client ${clientNumber}: MFA methods loaded.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'List MFA methods failed.');
    }
    return;
  }

  if (kind === 'delete-mfa') {
    const selected = [...div.querySelectorAll('.containmentMfaPick:checked')].map((cb) => ({
      UserPrincipalName: cb.dataset.user,
      Id: cb.dataset.id,
      ODataType: cb.dataset.odata,
    })).filter((x) => x.UserPrincipalName && x.Id);
    if (!selected.length) {
      log(`Client ${clientNumber}: select one or more MFA methods to remove.`);
      return;
    }
    if (!confirmContainmentPopup(`Remove ${selected.length} MFA method(s)?\n\n${selected.map((s) => `${s.UserPrincipalName} ${s.ODataType || s.Id}`).join('\n')}`)) return;
    setContainmentStatus(div, `Removing ${selected.length} MFA method(s)…`);
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_DELETE_AUTH_METHODS|ITEMS:${JSON.stringify(selected)}`, 'removing MFA methods');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    if (parsed.data?.Methods) applyAuthMethodsResult(div, parsed.data.Methods);
    if (parsed.prefix === 'REMEDIATE_AUTHMETHODS:') {
      setContainmentStatus(div, `Removed ${parsed.data?.SuccessCount || 0}; failed ${parsed.data?.FailCount || 0}.`);
      log(`Client ${clientNumber}: MFA method delete finished.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Remove MFA methods failed.');
    }
    noteContainmentAction(div, kind, selected, parsed);
    return;
  }

  if (kind === 'list-devices') {
    if (!graph || !users.length) return;
    setContainmentStatus(div, 'Listing registered devices…');
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_LIST_DEVICES|USERS:${usersJson}`, 'listing devices');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    if (parsed.prefix === 'REMEDIATE_DEVICES:') {
      applyDevicesResult(div, parsed.data?.Devices);
      setContainmentStatus(div, `${asArray(parsed.data?.Devices).filter((d) => d.Id).length} device(s).`);
      log(`Client ${clientNumber}: devices loaded.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'List devices failed.');
    }
    return;
  }

  if (kind === 'delete-devices') {
    const selected = [...div.querySelectorAll('.containmentDevicePick:checked')].map((cb) => ({
      Id: cb.dataset.id,
      UserPrincipalName: cb.dataset.user,
    })).filter((x) => x.Id);
    if (!selected.length) {
      log(`Client ${clientNumber}: select one or more devices to remove.`);
      return;
    }
    if (!confirmContainmentPopup(`Remove ${selected.length} Entra device object(s)?\n\n${selected.map((s) => s.Id).join('\n')}`)) return;
    setContainmentStatus(div, `Removing ${selected.length} device(s)…`);
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_DELETE_DEVICES|DEVICES:${JSON.stringify(selected)}`, 'removing devices');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    if (parsed.data?.Devices) applyDevicesResult(div, parsed.data.Devices);
    if (parsed.prefix === 'REMEDIATE_DEVICES:') {
      setContainmentStatus(div, `Removed ${parsed.data?.SuccessCount || 0}; failed ${parsed.data?.FailCount || 0}.`);
      log(`Client ${clientNumber}: device delete finished.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Remove devices failed.');
    }
    noteContainmentAction(div, kind, selected, parsed);
    return;
  }

  if (kind === 'list-apps') {
    if (!graph) return;
    if (!confirmSlowContainment(kind)) return;
    setContainmentStatus(div, 'Listing app registrations…');
    const final = await sendRemediateCommand(clientNumber, div, 'REMEDIATE_LIST_APPS', 'listing apps');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    if (parsed.prefix === 'REMEDIATE_APPS:') {
      applyAppsResult(div, parsed.data?.Apps);
      setContainmentStatus(div, `${asArray(parsed.data?.Apps).length} app(s).`);
      log(`Client ${clientNumber}: app registrations loaded.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'List apps failed.');
    }
    return;
  }

  if (kind === 'delete-apps') {
    const selected = [...div.querySelectorAll('.containmentAppPick:checked')].map((cb) => ({
      Id: cb.dataset.id,
      Kind: cb.dataset.kind,
      DisplayName: cb.dataset.name,
    })).filter((x) => x.Id);
    if (!selected.length) {
      log(`Client ${clientNumber}: select one or more apps to remove.`);
      return;
    }
    if (!confirmContainmentPopup(`Delete ${selected.length} app(s)? This removes the Entra object.\n\n${selected.map((s) => `${s.Kind}: ${s.DisplayName || s.Id}`).join('\n')}`)) return;
    setContainmentStatus(div, `Deleting ${selected.length} app(s)…`);
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_DELETE_APPS|APPS:${JSON.stringify(selected)}`, 'deleting apps');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    if (parsed.data?.Apps) applyAppsResult(div, parsed.data.Apps);
    if (parsed.prefix === 'REMEDIATE_APPS:') {
      setContainmentStatus(div, `Deleted ${parsed.data?.SuccessCount || 0}; failed ${parsed.data?.FailCount || 0}.`);
      log(`Client ${clientNumber}: app delete finished.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Delete apps failed.');
    }
    noteContainmentAction(div, kind, selected, parsed);
    return;
  }

  if (kind === 'restricted-status') {
    if (!exo || !users.length) return;
    setContainmentStatus(div, 'Checking Restricted Users list…');
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_RESTRICTED_EMAIL_STATUS|USERS:${usersJson}`, 'checking restricted email');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_RESTRICTED:') {
      div.dataset.restrictedEmailJson = JSON.stringify(parsed.data?.Users || []);
      rememberContainmentResult(div, 'restrictedEmail', parsed.data?.Users || []);
      setContainmentUserStatus(div, parsed.data?.Users);
      updateContainmentButtons(div);
      const hits = getRestrictedContainmentHits(div, users);
      log(hits.length
        ? `Client ${clientNumber}: ${hits.length} selected user(s) restricted from sending — Unrestrict is available.`
        : `Client ${clientNumber}: selected user(s) are not on Restricted entities.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Restricted status check failed.');
    }
    return;
  }

  if (kind === 'list-rules') {
    if (!exo || !users.length) return;
    const mailbox = users[0];
    setContainmentStatus(div, `Listing inbox rules for ${mailbox}…`);
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_LIST_INBOX_RULES|USER:${mailbox}`, 'listing inbox rules');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_RULES:') {
      renderContainmentRules(div, parsed.data?.User || mailbox, parsed.data?.Rules || []);
      setContainmentStatus(div, `${(parsed.data?.Rules || []).length} inbox rule(s) for ${mailbox}.`);
      log(`Client ${clientNumber}: loaded inbox rules for ${mailbox}.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'List inbox rules failed.');
    }
    return;
  }

  if (kind === 'delete-rule') {
    const mailbox = div.dataset.containmentRulesUser || users[0];
    if (!exo || !mailbox) {
      log(`Client ${clientNumber}: list inbox rules first, then select rows to delete.`);
      return;
    }
    const picked = [...div.querySelectorAll('.containmentRulePick:checked')];
    const selected = picked.map((cb) => cb.dataset.identity).filter(Boolean);
    if (!selected.length) {
      log(`Client ${clientNumber}: select one or more inbox rules to delete.`);
      return;
    }
    const labels = picked.map((cb) => cb.dataset.label || cb.dataset.identity).filter(Boolean);
    if (!confirmContainmentPopup(`Delete ${selected.length} inbox rule(s) for ${mailbox}?\n\n${labels.join('\n')}`)) {
      return;
    }
    const cmd = `REMEDIATE_DELETE_INBOX_RULES|USER:${mailbox}|RULES:${JSON.stringify(selected)}`;
    setContainmentStatus(div, `Deleting ${selected.length} rule(s)…`);
    const final = await sendRemediateCommand(clientNumber, div, cmd, 'deleting inbox rules');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_SUCCESS:') {
      renderContainmentRules(div, parsed.data?.User || mailbox, parsed.data?.Rules || []);
      setContainmentStatus(div, `Deleted ${parsed.data?.SuccessCount || 0}; failed ${parsed.data?.FailCount || 0}.`);
      log(`Client ${clientNumber}: inbox rule delete finished.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Delete inbox rules failed.');
    }
    noteContainmentAction(div, kind, mailbox, parsed);
    return;
  }

  if (kind === 'mailbox-status') {
    if (!exo || !users.length) return;
    setContainmentStatus(div, 'Checking mailbox forwarding and delegates…');
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_MAILBOX_STATUS|USERS:${usersJson}`, 'checking mailbox access');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_MAILBOX:') {
      const mailboxUsers = asArray(parsed.data?.Users);
      applyMailboxAccessResult(div, mailboxUsers);
      setContainmentStatus(div, `Mailbox access loaded for ${mailboxUsers.length} user(s).`);
      log(`Client ${clientNumber}: mailbox access updated.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Mailbox access check failed.');
    }
    return;
  }

  if (kind === 'set-forward') {
    if (!exo || !users.length) return;
    const smtp = (div.querySelector('.containmentForwardTo')?.value || '').trim();
    if (!smtp) {
      log(`Client ${clientNumber}: enter a Forward to address.`);
      return;
    }
    const deliver = div.querySelector('.containmentForwardKeep')?.checked ? '1' : '0';
    const mailbox = users[0];
    if (!confirmContainmentPopup(`Set forwarding for ${mailbox} to ${smtp} (keep copy: ${deliver === '1' ? 'yes' : 'no'})?`)) return;
    setContainmentStatus(div, `Setting forwarding for ${mailbox}…`);
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_SET_FORWARDING|USER:${mailbox}|SMTP:${smtp}|DELIVER:${deliver}`, 'setting forwarding');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_MAILBOX:') {
      applyMailboxAccessResult(div, parsed.data?.Users || []);
      setContainmentStatus(div, `Forwarding set for ${mailbox}.`);
      log(`Client ${clientNumber}: forwarding set for ${mailbox}.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Set forwarding failed.');
    }
    noteContainmentAction(div, kind, mailbox, parsed);
    return;
  }

  if (kind === 'remove-forward') {
    const selected = [...div.querySelectorAll('.containmentForwardPick:checked')].map((cb) => ({
      UserPrincipalName: cb.dataset.mailbox,
      Field: cb.dataset.field,
      Address: cb.dataset.target,
    })).filter((x) => x.UserPrincipalName && x.Field);
    if (!selected.length) {
      log(`Client ${clientNumber}: select one or more forwarding rows to remove.`);
      return;
    }
    const summary = selected.map((s) => `${s.Field} ${s.Address} on ${s.UserPrincipalName}`).join('\n');
    if (!confirmContainmentPopup(`Remove ${selected.length} forwarding entry(ies)?\n\n${summary}`)) return;
    setContainmentStatus(div, `Removing ${selected.length} forwarding entry(ies)…`);
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_REMOVE_FORWARDING|ITEMS:${JSON.stringify(selected)}`, 'removing forwarding');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_MAILBOX:') {
      applyMailboxAccessResult(div, parsed.data?.Users || []);
      setContainmentStatus(div, Array.isArray(parsed.data?.Details) ? parsed.data.Details.join('\n') : `Removed ${parsed.data?.SuccessCount || 0} forwarding entry(ies).`);
      log(`Client ${clientNumber}: selected forwarding removed.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Remove forwarding failed.');
    }
    noteContainmentAction(div, kind, selected, parsed);
    return;
  }

  if (kind === 'clear-forward') {
    if (!exo || !users.length) return;
    if (!confirmContainmentPopup(`Clear all mailbox forwarding (SMTP and recipient) for:\n${users.join('\n')}\n\nContinue?`)) return;
    setContainmentStatus(div, 'Clearing forwarding…');
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_CLEAR_FORWARDING|USERS:${usersJson}`, 'clearing forwarding');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_MAILBOX:') {
      applyMailboxAccessResult(div, parsed.data?.Users || []);
      setContainmentStatus(div, Array.isArray(parsed.data?.Details) ? parsed.data.Details.join('\n') : 'Forwarding cleared.');
      log(`Client ${clientNumber}: forwarding cleared.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Clear forwarding failed.');
    }
    noteContainmentAction(div, kind, users, parsed);
    return;
  }

  if (kind === 'add-delegate') {
    if (!exo || !users.length) return;
    const delegate = (div.querySelector('.containmentDelegateUser')?.value || '').trim();
    const right = div.querySelector('.containmentDelegateRight')?.value || 'FullAccess';
    if (!delegate) {
      log(`Client ${clientNumber}: enter a delegate UPN.`);
      return;
    }
    const mailbox = users[0];
    if (!confirmContainmentPopup(`Add ${right} for ${delegate} on ${mailbox}?`)) return;
    setContainmentStatus(div, `Adding ${right} for ${delegate}…`);
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_ADD_DELEGATION|USER:${mailbox}|DELEGATE:${delegate}|RIGHT:${right}`, 'adding delegate');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_MAILBOX:') {
      applyMailboxAccessResult(div, parsed.data?.Users || []);
      setContainmentStatus(div, `Added ${right} for ${delegate} on ${mailbox}.`);
      log(`Client ${clientNumber}: delegate added.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Add delegate failed.');
    }
    noteContainmentAction(div, kind, mailbox, parsed);
    return;
  }

  if (kind === 'remove-delegate') {
    const selected = [...div.querySelectorAll('.containmentDelegatePick:checked')].map((cb) => ({
      mailbox: cb.dataset.mailbox,
      user: cb.dataset.user,
      right: cb.dataset.right,
    })).filter((x) => x.mailbox && x.user && x.right);
    if (!selected.length) {
      log(`Client ${clientNumber}: select one or more delegates to remove.`);
      return;
    }
    const summary = selected.map((s) => `${s.right} ${s.user} on ${s.mailbox}`).join('\n');
    if (!confirmContainmentPopup(`Remove selected delegates?\n\n${summary}`)) return;
    let lastUsers = [];
    for (const item of selected) {
      setContainmentStatus(div, `Removing ${item.right} for ${item.user}…`);
      const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_REMOVE_DELEGATION|USER:${item.mailbox}|DELEGATE:${item.user}|RIGHT:${item.right}`, 'removing delegate');
      const parsed = parseRemediatePayload(final || '');
      if (parsed.prefix === 'REMEDIATE_MAILBOX:') {
        lastUsers = asArray(parsed.data?.Users);
      } else {
        showContainmentFailure(div, clientNumber, parsed, 'Remove delegate failed.');
        return;
      }
    }
    if (lastUsers.length) applyMailboxAccessResult(div, lastUsers);
    setContainmentStatus(div, `Removed ${selected.length} delegate grant(s).`);
    log(`Client ${clientNumber}: delegates removed.`);
    noteContainmentAction(div, kind, selected, { prefix: 'REMEDIATE_MAILBOX:', data: { SuccessCount: selected.length, FailCount: 0 } });
    return;
  }

  if (kind === 'list-transport') {
    if (!exo) return;
    if (!confirmSlowContainment(kind)) return;
    setContainmentStatus(div, 'Listing transport rules…');
    const final = await sendRemediateCommand(clientNumber, div, 'REMEDIATE_LIST_TRANSPORT_RULES', 'listing transport rules');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_TRANSPORT:') {
      const transportRules = asArray(parsed.data?.Rules);
      applyTransportRulesResult(div, transportRules);
      setContainmentStatus(div, `${transportRules.length} transport rule(s).`);
      log(`Client ${clientNumber}: transport rules loaded.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'List transport rules failed.');
    }
    return;
  }

  if (kind === 'delete-transport') {
    const ids = [...div.querySelectorAll('.containmentTransportPick:checked')].map((cb) => cb.dataset.identity).filter(Boolean);
    if (!ids.length) {
      log(`Client ${clientNumber}: select one or more transport rules to delete.`);
      return;
    }
    if (!confirmContainmentPopup(`Delete ${ids.length} transport rule(s)?\n\n${ids.join('\n')}`)) return;
    setContainmentStatus(div, `Deleting ${ids.length} transport rule(s)…`);
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_DELETE_TRANSPORT_RULES|RULES:${JSON.stringify(ids)}`, 'deleting transport rules');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_TRANSPORT:') {
      applyTransportRulesResult(div, parsed.data?.Rules || []);
      setContainmentStatus(div, `Deleted ${parsed.data?.SuccessCount || 0}; failed ${parsed.data?.FailCount || 0}.`);
      log(`Client ${clientNumber}: transport rule delete finished.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Delete transport rules failed.');
    }
    noteContainmentAction(div, kind, '', parsed);
    return;
  }

  if (kind === 'list-connectors') {
    if (!exo) return;
    setContainmentStatus(div, 'Listing connectors…');
    const final = await sendRemediateCommand(clientNumber, div, 'REMEDIATE_LIST_CONNECTORS', 'listing connectors');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_CONNECTORS:') {
      const connectors = asArray(parsed.data?.Connectors);
      applyConnectorsResult(div, connectors);
      setContainmentStatus(div, `${connectors.length} connector(s).`);
      log(`Client ${clientNumber}: connectors loaded.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'List connectors failed.');
    }
    return;
  }

  if (kind === 'delete-connectors') {
    const items = [...div.querySelectorAll('.containmentConnectorPick:checked')].map((cb) => ({
      Name: cb.dataset.name,
      Direction: cb.dataset.direction,
    })).filter((x) => x.Name);
    if (!items.length) {
      log(`Client ${clientNumber}: select one or more connectors to delete.`);
      return;
    }
    if (!confirmContainmentPopup(`Delete ${items.length} connector(s)?\n\n${items.map((i) => `${i.Direction}: ${i.Name}`).join('\n')}`)) return;
    setContainmentStatus(div, `Deleting ${items.length} connector(s)…`);
    const final = await sendRemediateCommand(clientNumber, div, `REMEDIATE_DELETE_CONNECTORS|CONNECTORS:${JSON.stringify(items)}`, 'deleting connectors');
    const parsed = parseRemediatePayload(final || '');
    if (parsed.prefix === 'REMEDIATE_CONNECTORS:') {
      applyConnectorsResult(div, parsed.data?.Connectors || []);
      setContainmentStatus(div, `Deleted ${parsed.data?.SuccessCount || 0}; failed ${parsed.data?.FailCount || 0}.`);
      log(`Client ${clientNumber}: connector delete finished.`);
    } else {
      showContainmentFailure(div, clientNumber, parsed, 'Delete connectors failed.');
    }
    noteContainmentAction(div, kind, '', parsed);
    return;
  }

  const extra = {
    'list-oauth': { need: 'graph', users: true, cmd: () => `REMEDIATE_LIST_OAUTH_GRANTS|USERS:${usersJson}`, label: 'listing OAuth consents', prefix: 'REMEDIATE_OAUTH:', apply: (d, data) => applyOauthResult(d, data?.Grants), count: (data) => asArray(data?.Grants).filter((g) => g.Id).length },
    'delete-oauth': { need: 'graph', confirm: true, prefix: 'REMEDIATE_OAUTH:', apply: (d, data) => applyOauthResult(d, data?.Grants), picks: '.containmentOauthPick:checked', map: (cb) => ({ Id: cb.dataset.id, UserPrincipalName: cb.dataset.user }), token: 'ITEMS', cmdName: 'REMEDIATE_DELETE_OAUTH_GRANTS', label: 'revoking OAuth consents', confirmText: (items) => `Revoke ${items.length} OAuth consent(s)?` },
    'reregister-mfa': { need: 'graph', users: true, write: true, cmd: () => `REMEDIATE_REREGISTER_MFA|USERS:${usersJson}`, label: 'wiping MFA methods', prefix: 'REMEDIATE_SUCCESS:', confirmText: () => `Delete all removable MFA methods and revoke sessions for:\n${users.join('\n')}\n\nThe user must re-register MFA. Continue?` },
    'list-mobile': { need: 'exo', users: true, cmd: () => `REMEDIATE_LIST_MOBILE_DEVICES|USERS:${usersJson}`, label: 'listing mobile partnerships', prefix: 'REMEDIATE_MOBILE:', apply: (d, data) => applyMobileResult(d, data?.Devices), count: (data) => asArray(data?.Devices).filter((x) => x.Identity).length },
    'delete-mobile': { need: 'exo', confirm: true, prefix: 'REMEDIATE_MOBILE:', apply: (d, data) => applyMobileResult(d, data?.Devices), picks: '.containmentMobilePick:checked', map: (cb) => ({ Identity: cb.dataset.identity, UserPrincipalName: cb.dataset.user }), token: 'ITEMS', cmdName: 'REMEDIATE_DELETE_MOBILE_DEVICES', label: 'removing mobile partnerships', confirmText: (items) => `Remove ${items.length} ActiveSync partnership(s)?` },
    'list-intune': { need: 'graph', users: true, cmd: () => `REMEDIATE_LIST_INTUNE|USERS:${usersJson}`, label: 'listing Intune devices', prefix: 'REMEDIATE_INTUNE:', apply: (d, data) => applyIntuneResult(d, data?.Devices), count: (data) => asArray(data?.Devices).filter((x) => x.Id).length, after: (data) => { const err = asArray(data?.Devices).find((x) => x.Error && !x.Id)?.Error; if (err) setContainmentStatus(div, shortenGraphError(err), { tone: 'statusWarn' }); } },
    'wipe-intune': { need: 'graph', confirm: true, prefix: 'REMEDIATE_INTUNE:', picks: '.containmentIntunePick:checked', map: (cb) => ({ Id: cb.dataset.id }), token: 'DEVICES', cmdName: 'REMEDIATE_WIPE_INTUNE', label: 'wiping Intune devices', confirmText: (items) => `WIPE ${items.length} Intune device(s)? This is destructive.` },
    'retire-intune': { need: 'graph', confirm: true, prefix: 'REMEDIATE_INTUNE:', picks: '.containmentIntunePick:checked', map: (cb) => ({ Id: cb.dataset.id }), token: 'DEVICES', cmdName: 'REMEDIATE_RETIRE_INTUNE', label: 'retiring Intune devices', confirmText: (items) => `Retire ${items.length} Intune device(s)?` },
    'list-folders': { need: 'exo', users: true, cmd: () => `REMEDIATE_LIST_FOLDER_PERMS|USERS:${usersJson}`, label: 'listing folder permissions', prefix: 'REMEDIATE_FOLDERS:', apply: (d, data) => applyFoldersResult(d, data?.Permissions), count: (data) => asArray(data?.Permissions).filter((x) => x.User && x.User !== '(none)').length },
    'delete-folders': { need: 'exo', confirm: true, prefix: 'REMEDIATE_FOLDERS:', apply: (d, data) => applyFoldersResult(d, data?.Permissions), picks: '.containmentFolderPick:checked', map: (cb) => ({ UserPrincipalName: cb.dataset.user, Folder: cb.dataset.folder, User: cb.dataset.trustee }), token: 'ITEMS', cmdName: 'REMEDIATE_DELETE_FOLDER_PERMS', label: 'removing folder permissions', confirmText: (items) => `Remove ${items.length} folder permission(s)?` },
    'autoreply-status': { need: 'exo', users: true, cmd: () => `REMEDIATE_GET_AUTOREPLY|USERS:${usersJson}`, label: 'checking auto-reply', prefix: 'REMEDIATE_AUTOREPLY:', apply: (d, data) => applyAutoreplyResult(d, data?.Users) },
    'disable-autoreply': { need: 'exo', users: true, write: true, cmd: () => `REMEDIATE_DISABLE_AUTOREPLY|USERS:${usersJson}`, label: 'disabling auto-reply', prefix: 'REMEDIATE_AUTOREPLY:', apply: (d, data) => applyAutoreplyResult(d, data?.Users), confirmText: () => `Disable automatic replies for:\n${users.join('\n')}` },
    'list-junk': { need: 'exo', users: true, cmd: () => `REMEDIATE_LIST_JUNK|USERS:${usersJson}`, label: 'listing trusted senders', prefix: 'REMEDIATE_JUNK:', apply: (d, data) => applyJunkResult(d, data?.Entries), count: (data) => asArray(data?.Entries).filter((x) => x.Address && x.Address !== '(none)').length },
    'delete-junk': { need: 'exo', confirm: true, prefix: 'REMEDIATE_JUNK:', apply: (d, data) => applyJunkResult(d, data?.Entries), picks: '.containmentJunkPick:checked', map: (cb) => ({ UserPrincipalName: cb.dataset.user, List: cb.dataset.list, Address: cb.dataset.address }), token: 'ITEMS', cmdName: 'REMEDIATE_DELETE_JUNK', label: 'removing trusted senders', confirmText: (items) => `Remove ${items.length} trusted sender/recipient(s)?` },
    'list-elsewhere': { need: 'exo', users: true, wait: 300, cmd: () => `REMEDIATE_LIST_ELSEWHERE|USERS:${usersJson}`, label: 'listing rights on other mailboxes', prefix: 'REMEDIATE_ELSEWHERE:', apply: (d, data) => applyElsewhereResult(d, data?.Grants), count: (data) => asArray(data?.Grants).filter((x) => x.Mailbox && x.Mailbox !== '(none)').length },
    'delete-elsewhere': { need: 'exo', confirm: true, wait: 300, prefix: 'REMEDIATE_ELSEWHERE:', apply: (d, data) => applyElsewhereResult(d, data?.Grants), picks: '.containmentElsewherePick:checked', map: (cb) => ({ UserPrincipalName: cb.dataset.user, Mailbox: cb.dataset.mailbox, Right: cb.dataset.right, Trustee: cb.dataset.trustee }), token: 'ITEMS', cmdName: 'REMEDIATE_DELETE_ELSEWHERE', label: 'removing grants on other mailboxes', confirmText: (items) => `Remove ${items.length} mailbox grant(s) this user has elsewhere?` },
    'hold-status': { need: 'exo', users: true, cmd: () => `REMEDIATE_GET_MAILBOX_HOLD|USERS:${usersJson}`, label: 'checking hold / audit', prefix: 'REMEDIATE_HOLD:', apply: (d, data) => applyHoldResult(d, data?.Users) },
    'enable-hold': { need: 'exo', users: true, write: true, cmd: () => `REMEDIATE_SET_MAILBOX_HOLD|USERS:${usersJson}`, label: 'enabling hold + audit', prefix: 'REMEDIATE_HOLD:', apply: (d, data) => applyHoldResult(d, data?.Users), confirmText: () => `Enable litigation hold, 30-day deleted-item retention, and mailbox audit for:\n${users.join('\n')}` },
    'list-orgfwd': { need: 'exo', cmd: () => 'REMEDIATE_LIST_ORG_FORWARD', label: 'listing org auto-forward', prefix: 'REMEDIATE_ORGFWD:', apply: (d, data) => applyOrgfwdResult(d, data?.Policies), count: (data) => asArray(data?.Policies).length },
    'disable-orgfwd': { need: 'exo', confirm: true, prefix: 'REMEDIATE_ORGFWD:', apply: (d, data) => applyOrgfwdResult(d, data?.Policies), picks: '.containmentOrgfwdPick:checked', map: (cb) => ({ Kind: cb.dataset.kind, Identity: cb.dataset.identity }), token: 'ITEMS', cmdName: 'REMEDIATE_SET_ORG_FORWARD', label: 'disabling org auto-forward', confirmText: (items) => `Disable auto-forward on ${items.length} remote domain / outbound spam policy(ies)?` },
    'list-journal': { need: 'exo', cmd: () => 'REMEDIATE_LIST_JOURNAL', label: 'listing journal rules', prefix: 'REMEDIATE_JOURNAL:', apply: (d, data) => applyJournalResult(d, data?.Rules), count: (data) => asArray(data?.Rules).filter((x) => x.Identity).length },
    'delete-journal': { need: 'exo', confirm: true, prefix: 'REMEDIATE_JOURNAL:', apply: (d, data) => applyJournalResult(d, data?.Rules), picks: '.containmentJournalPick:checked', map: (cb) => ({ Identity: cb.dataset.identity }), token: 'ITEMS', cmdName: 'REMEDIATE_DELETE_JOURNAL', label: 'deleting journal rules', confirmText: (items) => `Delete ${items.length} journal rule(s)?` },
    'list-roles': { need: 'graph', users: true, cmd: () => `REMEDIATE_LIST_ROLES|USERS:${usersJson}`, label: 'listing roles and groups', prefix: 'REMEDIATE_ROLES:', apply: (d, data) => applyRolesResult(d, data?.Roles), count: (data) => asArray(data?.Roles).filter((x) => x.Id).length },
    'delete-roles': { need: 'graph', confirm: true, prefix: 'REMEDIATE_ROLES:', apply: (d, data) => applyRolesResult(d, data?.Roles), picks: '.containmentRolePick:checked', map: (cb) => ({ Id: cb.dataset.id, Kind: cb.dataset.kind, UserPrincipalName: cb.dataset.user }), token: 'ITEMS', cmdName: 'REMEDIATE_DELETE_ROLES', label: 'removing roles / groups', confirmText: (items) => `Remove ${items.length} role assignment / group membership(s)?` },
    'list-appcreds': { need: 'graph', cmd: () => 'REMEDIATE_LIST_APP_CREDS', label: 'listing app secrets / owners', prefix: 'REMEDIATE_APPCREDS:', apply: (d, data) => applyAppcredsResult(d, data?.Credentials), count: (data) => asArray(data?.Credentials).filter((x) => x.AppId).length },
    'delete-appcreds': { need: 'graph', confirm: true, prefix: 'REMEDIATE_APPCREDS:', apply: (d, data) => applyAppcredsResult(d, data?.Credentials), picks: '.containmentAppcredPick:checked', map: (cb) => ({ Kind: cb.dataset.kind, AppId: cb.dataset.appid, KeyId: cb.dataset.keyid, OwnerId: cb.dataset.ownerid }), token: 'ITEMS', cmdName: 'REMEDIATE_DELETE_APP_CREDS', label: 'removing secrets / owners', confirmText: (items) => `Remove ${items.length} app secret(s) or owner(s)?` },
    'list-flows': { cmd: () => 'REMEDIATE_LIST_FLOWS', label: 'listing Power Automate flows', prefix: 'REMEDIATE_FLOWS:', apply: (d, data) => applyFlowsResult(d, data?.Flows), after: (data) => { if (data?.Message) setContainmentStatus(div, data.Message, { reveal: true }); } },
    'delete-flows': { confirm: true, prefix: 'REMEDIATE_FLOWS:', apply: (d, data) => applyFlowsResult(d, data?.Flows), picks: '.containmentFlowPick:checked', map: (cb) => ({ Id: cb.dataset.id, Environment: cb.dataset.env }), token: 'ITEMS', cmdName: 'REMEDIATE_DELETE_FLOWS', label: 'deleting flows', confirmText: (items) => `Delete ${items.length} flow(s)?` },
  };
  const extraSpec = extra[kind];
  if (extraSpec) {
    if (extraSpec.need === 'graph' && !graph) {
      log(`Client ${clientNumber}: Graph Auth is required.`);
      return;
    }
    if (extraSpec.need === 'exo' && !exo) {
      log(`Client ${clientNumber}: Exchange Auth is required.`);
      return;
    }
    if (extraSpec.users && !users.length) {
      log(`Client ${clientNumber}: select one or more validated users.`);
      return;
    }
    let cmd = extraSpec.cmd ? extraSpec.cmd() : '';
    let selected = [];
    if (extraSpec.picks) {
      selected = [...div.querySelectorAll(extraSpec.picks)].map(extraSpec.map).filter((x) => Object.values(x).some(Boolean));
      if (!selected.length) {
        log(`Client ${clientNumber}: select one or more rows first.`);
        return;
      }
      cmd = `${extraSpec.cmdName}|${extraSpec.token}:${JSON.stringify(selected)}`;
    }
    if (!confirmSlowContainment(kind)) return;
    if ((extraSpec.confirm || extraSpec.write) && !confirmContainmentPopup(extraSpec.confirmText(selected))) return;
    setContainmentStatus(div, `${extraSpec.label}…`);
    const final = await sendRemediateCommand(clientNumber, div, cmd, extraSpec.label, extraSpec.wait || 180);
    const parsed = parseRemediatePayload(final || '');
    if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
    if (parsed.prefix === extraSpec.prefix) {
      if (extraSpec.apply) extraSpec.apply(div, parsed.data);
      const statusOpts = extraSpec.apply ? {} : { reveal: true };
      if (parsed.data?.SuccessCount != null) setContainmentStatus(div, `Finished. Success ${parsed.data.SuccessCount || 0}; failed ${parsed.data.FailCount || 0}.`, statusOpts);
      else if (extraSpec.count) setContainmentStatus(div, `${extraSpec.count(parsed.data)} row(s).`, statusOpts);
      else if (Array.isArray(parsed.data?.Details)) setContainmentStatus(div, parsed.data.Details.join('\n'), statusOpts);
      else setContainmentStatus(div, 'Done.', statusOpts);
      if (extraSpec.after) extraSpec.after(parsed.data);
      log(`Client ${clientNumber}: ${extraSpec.label} finished.`);
    } else {
      if (extraSpec.apply && parsed.data) extraSpec.apply(div, parsed.data);
      showContainmentFailure(div, clientNumber, parsed, `${extraSpec.label} failed.`);
    }
    if (extraSpec.confirm || extraSpec.write) {
      noteContainmentAction(div, kind, selected.length ? selected : users, parsed);
    }
    return;
  }

  const spec = writes[kind];
  if (!spec) return;
  if (spec.need === 'graph' && !graph) {
    log(`Client ${clientNumber}: Graph Auth is required for ${spec.action}.`);
    return;
  }
  if (spec.need === 'exo' && !exo) {
    log(`Client ${clientNumber}: Exchange Auth is required for ${spec.action}.`);
    return;
  }
  if (!users.length) {
    log(`Client ${clientNumber}: select one or more validated users.`);
    return;
  }
  if (kind === 'block' || kind === 'unblock') {
    const statusFinal = await sendRemediateCommand(clientNumber, div, `REMEDIATE_SIGNIN_STATUS|USERS:${usersJson}`, 'checking sign-in status');
    const statusParsed = parseRemediatePayload(statusFinal || '');
    if (statusParsed.data?.Capabilities) applyContainmentCapabilities(div, statusParsed.data.Capabilities);
    if (statusParsed.prefix === 'REMEDIATE_SUCCESS:') {
      setContainmentUserStatus(div, statusParsed.data?.Users);
    }
  }
  if (kind === 'unrestrict') {
    const hits = getRestrictedContainmentHits(div, users);
    if (!hits.length) {
      log(`Client ${clientNumber}: Check restricted status first. Unrestrict is only available when a selected user is on Restricted entities.`);
      return;
    }
    spec.cmd = `REMEDIATE_UNRESTRICT_EMAIL|USERS:${JSON.stringify(hits.map((h) => h.UserPrincipalName))}`;
    setContainmentUserStatus(div, hits);
  }
  const confirmTargets = kind === 'unrestrict'
    ? getRestrictedContainmentHits(div, users).map((h) => h.UserPrincipalName)
    : users;
  if (!confirmContainmentPopup(`${spec.label} for:\n${confirmTargets.join('\n')}\n\nContinue?`)) {
    return;
  }
  setContainmentStatus(div, `${spec.label}…`);
  const final = await sendRemediateCommand(clientNumber, div, spec.cmd, spec.label);
  const parsed = parseRemediatePayload(final || '');
  if (parsed.data?.Capabilities) applyContainmentCapabilities(div, parsed.data.Capabilities);
  noteContainmentAction(div, kind, confirmTargets, parsed);
  if (parsed.prefix === 'REMEDIATE_SUCCESS:') {
    const details = parsed.data?.Details;
    setContainmentStatus(div, Array.isArray(details) ? details.join('\n') : (parsed.raw || 'Done.'), { reveal: true });
    log(`Client ${clientNumber}: ${spec.action} finished.`);
    if (kind === 'block' || kind === 'unblock') {
      await runContainmentAction(clientNumber, div, 'signin-status');
    }
    if (kind === 'unrestrict') {
      await runContainmentAction(clientNumber, div, 'restricted-status');
    }
  } else {
    showContainmentFailure(div, clientNumber, parsed, `${spec.action} failed.`);
  }
}

const BEC_RUN_ORDER = ['enable-hold', 'delete-rules', 'remove-forward', 'remove-delegates', 'revoke', 'block', 'reset-password'];

function becRuleIsSuspicious(rule) {
  if (!rule) return false;
  if (rule.Hidden === true || rule.hidden === true || rule.Hidden === 'True' || rule.hidden === 'True') return true;
  if (String(rule.RedirectTo || rule.redirectTo || '').trim()) return true;
  if (String(rule.ForwardTo || rule.forwardTo || '').trim()) return true;
  if (String(rule.ForwardAsAttachmentTo || rule.forwardAsAttachmentTo || '').trim()) return true;
  if (rule.DeleteMessage === true || rule.deleteMessage === true || rule.DeleteMessage === 'True') return true;
  return false;
}

function containmentMemory(div) {
  const clientNumber = div?.closest('details.tenant')?.dataset?.client;
  return tenantUiState.get(String(clientNumber))?.containment || {};
}

function buildBecFacts(div, mailbox) {
  const mem = containmentMemory(div);
  const want = String(mailbox || '').trim().toLowerCase();
  const flagOn = (value) => value === true || value === 'True';
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
    if (want && upn !== want) return;
    const listed = asArray(mbx.Forwards);
    if (listed.length) {
      listed.forEach((row) => {
        if (row.Address) forwards.push({ address: row.Address, field: row.Field || '' });
      });
    } else {
      if (mbx.ForwardingSmtpAddress) forwards.push({ address: mbx.ForwardingSmtpAddress, field: 'Smtp' });
      if (mbx.ForwardingAddress) forwards.push({ address: mbx.ForwardingAddress, field: 'Recipient' });
    }
    asArray(mbx.Delegates).forEach((row) => {
      if (row.User) delegates.push({ user: row.User, right: row.Right || '' });
    });
  });
  const ticketNumber = (div.querySelector('.ticketInput')?.value || '').trim();
  const ticket = (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || '';
  const clientNumber = div.closest('details.tenant')?.dataset?.client;
  return { mailbox, rules, forwards, delegates, ticketNumber, ticket, company: getTenantCompanyName(div, clientNumber) };
}

function renderBecReview(div, result) {
  const text = (result && result.text) || '';
  const mainWrap = div.querySelector('.becReviewWrap');
  const mainPre = div.querySelector('.becReviewText');
  if (mainPre) mainPre.textContent = text;
  if (mainWrap) mainWrap.style.display = text ? 'block' : 'none';
  const wrap = div.querySelector('.containmentBecWrap');
  const pre = div.querySelector('.containmentBecText');
  const actions = div.querySelector('.containmentBecActions');
  const list = div.querySelector('.containmentBecList');
  if (!wrap || !pre || !actions || !list) return;
  pre.textContent = text;
  wrap.style.display = 'block';
  const items = asArray(result && result.remediations);
  if (!items.length) {
    actions.style.display = 'none';
    list.innerHTML = '';
    setContainmentStatus(div, 'Threat review finished. No remediation was suggested.', { reveal: true });
    return;
  }
  list.innerHTML = items.map((item) => {
    const id = item.Id || item.id || '';
    const checked = item.Checked === true || item.checked === true;
    return `<label class="row" style="align-items:flex-start">
      <input type="checkbox" class="containmentBecPick" value="${escapeHtml(id)}" ${checked ? 'checked' : ''} />
      <span><strong>${escapeHtml(item.Title || item.title || id)}</strong><br><span class="muted">${escapeHtml(item.Detail || item.detail || '')}</span></span>
    </label>`;
  }).join('');
  actions.style.display = 'block';
  setContainmentStatus(div, `Threat review finished. ${items.length} remediation option(s). Nothing has been changed yet.`, { reveal: true });
}

function selectBecRuleBoxes(div) {
  const rules = asArray(containmentMemory(div).rules);
  div.querySelectorAll('.containmentRulePick').forEach((box) => {
    const idx = Number(box.dataset.idx);
    box.checked = becRuleIsSuspicious(rules[idx]);
  });
  updateContainmentButtons(div);
}

function becWindowFromTicket(text) {
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

function formatPresentedLocal(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  const parsed = Date.parse(s);
  if (!Number.isFinite(parsed)) return s;
  const d = new Date(parsed);
  const pad = (n) => String(n).padStart(2, '0');
  let hour = d.getHours();
  const ampm = hour >= 12 ? 'PM' : 'AM';
  hour = hour % 12 || 12;
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${hour}:${pad(d.getMinutes())} ${ampm}`;
}

function presentedLocalZone() {
  try {
    const parts = new Intl.DateTimeFormat('en-US', { timeZoneName: 'long' }).formatToParts(new Date());
    const name = parts.find((p) => p.type === 'timeZoneName');
    if (name && name.value && !/^(?:UTC|GMT)\b/i.test(name.value)) return name.value;
  } catch { /* keep the fallback */ }
  return 'local time';
}

function becFactsFromMemory(div, mailbox, audit, windowNote) {
  const facts = buildBecFacts(div, mailbox);
  const mem = containmentMemory(div);
  facts.authMethods = asArray(mem.authMethods);
  facts.devices = asArray(mem.devices);
  facts.intune = asArray(mem.intune);
  facts.oauth = asArray(mem.oauth);
  facts.roles = asArray(mem.roles);
  facts.mobile = asArray(mem.mobile);
  if (audit) {
    facts.mfaPosture = audit.MfaPosture || audit.mfaPosture || null;
    facts.messageTrace = asArray(audit.MessageTrace || audit.messageTrace);
    facts.ual = asArray(audit.Ual || audit.ual);
    if (audit && (Object.prototype.hasOwnProperty.call(audit, 'Dlp') || Object.prototype.hasOwnProperty.call(audit, 'dlp'))) {
      facts.dlp = asArray(audit.Dlp || audit.dlp);
    }
    if (audit && (Object.prototype.hasOwnProperty.call(audit, 'Files') || Object.prototype.hasOwnProperty.call(audit, 'files'))) {
      facts.files = asArray(audit.Files || audit.files);
    }
    if (audit && (Object.prototype.hasOwnProperty.call(audit, 'Sharing') || Object.prototype.hasOwnProperty.call(audit, 'sharing'))) {
      facts.sharing = asArray(audit.Sharing || audit.sharing);
    }
    if (audit && (Object.prototype.hasOwnProperty.call(audit, 'Alerts') || Object.prototype.hasOwnProperty.call(audit, 'alerts'))) {
      facts.alerts = asArray(audit.Alerts || audit.alerts);
    }
    facts.signIns = asArray(audit.SignIns || audit.signIns);
    facts.apps = asArray(audit.Apps || audit.apps);
    if (Object.prototype.hasOwnProperty.call(audit, 'Blasts') || Object.prototype.hasOwnProperty.call(audit, 'blasts')) {
      facts.blasts = asArray(audit.Blasts || audit.blasts);
    }
    facts.pullNotes = asArray(audit.PullNotes || audit.pullNotes);
    if (audit.Days || audit.days) facts.windowDays = audit.Days || audit.days;
  }
  const startIso = div.dataset.becWindowStart || '';
  facts.apps = threatAppsInWindow(mem.apps, facts.apps, startIso);
  facts.transport = threatTransportRows(mem.transport, startIso);
  facts.connectors = threatConnectorRows(mem.connectors, startIso);
  const gap = readAuditGap(div);
  if (gap && (asArray(gap.unchecked).length || gap.lead)) facts.collectionGap = gap;
  if (windowNote) facts.windowNote = windowNote;
  return facts;
}

function readAuditGap(div) {
  try {
    const parsed = JSON.parse(div?.dataset?.auditGap || '');
    if (parsed && typeof parsed === 'object') return parsed;
  } catch { /* stored gap is not json */ }
  return null;
}

function rememberAuditGap(div, gap) {
  const clientNumber = div.closest('details.tenant')?.dataset?.client;
  const prior = tenantUiState.get(String(clientNumber)) || {};
  const existing = prior.auditGap && typeof prior.auditGap === 'object' ? prior.auditGap : (readAuditGap(div) || {});
  const next = {
    mailbox: gap.mailbox || existing.mailbox || '',
    unchecked: Array.isArray(gap.unchecked) ? gap.unchecked : (existing.unchecked || []),
    lead: gap.lead || existing.lead || '',
  };
  div.dataset.auditGap = JSON.stringify(next);
  tenantUiState.set(String(clientNumber), { ...prior, auditGap: next });
  scheduleTenantUiStateSync(clientNumber, div, true);
}

function noteAuditCollected(div, mailbox) {
  const existing = readAuditGap(div);
  if (!existing) return;
  if (!existing.lead && !asArray(existing.unchecked).length) return;
  rememberAuditGap(div, { mailbox, unchecked: [], lead: existing.lead || '' });
}

function threatInWindow(value, startIso) {
  const text = String(value || '').trim();
  if (!text || !startIso) return null;
  const when = Date.parse(text);
  const start = Date.parse(startIso);
  if (Number.isNaN(when) || Number.isNaN(start)) return null;
  return when >= start;
}

function threatAppsInWindow(listed, audited, startIso) {
  const rows = [];
  const seen = new Set();
  const add = (app, fromAudit) => {
    if (!app) return;
    const name = String(app.DisplayName || app.displayName || '').trim();
    const created = String(app.Created || app.created || app.createdDateTime || '').trim();
    if (!name) return;
    const key = `${name}|${created}`.toLowerCase();
    if (seen.has(key)) return;
    const inWindow = threatInWindow(created, startIso);
    if (inWindow === false) return;
    if (inWindow !== true && !fromAudit) return;
    seen.add(key);
    rows.push({ displayName: name, created });
  };
  asArray(audited).forEach((app) => add(app, true));
  asArray(listed).forEach((app) => add(app, false));
  return rows.slice(0, 25);
}

function threatTransportRows(rules, startIso) {
  const risky = /redirect|forward|delete|drop|bcc|blind|connector|route|reject|quarantine/i;
  const picked = [];
  asArray(rules).forEach((rule) => {
    const when = String(rule.WhenChanged || rule.whenChanged || '').trim();
    const inWindow = threatInWindow(when, startIso);
    const text = `${rule.Name || ''} ${rule.Description || ''}`;
    if (inWindow === false && !risky.test(text)) return;
    picked.push({
      name: rule.Name || rule.Identity || '',
      state: rule.State || '',
      priority: rule.Priority,
      whenChanged: when,
      description: String(rule.Description || '').slice(0, 180),
    });
  });
  return picked.slice(0, 40);
}

function threatConnectorRows(connectors, startIso) {
  const picked = [];
  asArray(connectors).forEach((row) => {
    const when = String(row.WhenCreated || row.whenCreated || '').trim();
    const inWindow = threatInWindow(when, startIso);
    const outbound = String(row.Direction || '').toLowerCase() === 'outbound';
    const enabled = row.Enabled === true || row.Enabled === 'True';
    if (inWindow === false && !(outbound && enabled)) return;
    picked.push({
      direction: row.Direction || '',
      name: row.Name || row.Identity || '',
      enabled: enabled ? 'enabled' : 'disabled',
      connectorType: row.ConnectorType || '',
      hosts: row.SmartHosts || '',
      domains: row.Domains || '',
      whenCreated: when,
    });
  });
  return picked.slice(0, 30);
}

function auditPullFailed(audit) {
  if (!audit) return true;
  const keys = ['Ual', 'ual', 'Sharing', 'sharing', 'Files', 'files', 'SignIns', 'signIns', 'Dlp', 'dlp'];
  return !keys.some((key) => Object.prototype.hasOwnProperty.call(audit, key));
}

const BEC_REPULL_LOGS = ['trace', 'ual', 'sharing', 'files', 'dlp', 'alerts', 'signins', 'apps'];

function becRepullLogs(review) {
  return asArray(review && (review.repullLogs || review.RepullLogs))
    .map((name) => String(name || '').trim().toLowerCase())
    .filter((name) => BEC_REPULL_LOGS.includes(name));
}

function becRepullLabel(logs) {
  return Array.isArray(logs) && logs.length ? logs.join(', ') : 'message trace and audit';
}

function becFollowUpNote(again, logs) {
  const span = `${formatPresentedLocal(again.start)} through ${formatPresentedLocal(again.end)} ${presentedLocalZone()}`;
  const cap = again.days >= 90
    ? 'This is the longest lookback. Do not ask for another pull. If the record is not here, say it was not in the logs.'
    : 'If this window still cannot answer the investigation, ask for a longer look up to 90 days.';
  if (Array.isArray(logs) && logs.length) {
    return `Follow-up pull of ${logs.join(', ')} only, ${again.days} days, ${span}. The other logs stay on their earlier window. ${cap}`;
  }
  return `Follow-up pull, ${again.days} days, ${span}. Message trace and unified audit stop at tenant retention if that is shorter. ${cap}`;
}

function mergeBecAudit(prior, next, logs) {
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

async function pullBecAudit(clientNumber, div, mailbox, window, logs) {
  const named = Array.isArray(logs) ? logs.filter((name) => BEC_REPULL_LOGS.includes(name)) : [];
  const extra = named.length ? `|LOGS:${named.join(',')}` : '';
  const cmd = `REMEDIATE_BEC_AUDIT|USER:${mailbox}|START:${window.start}|END:${window.end}${extra}`;
  const pulling = named.length ? `pulling ${named.join(', ')}` : 'pulling message trace, audit, DLP, alerts, and sign-ins';
  const final = await sendRemediateCommand(clientNumber, div, cmd, pulling, 300);
  const parsed = parseRemediatePayload(final || '');
  if (parsed.prefix !== 'REMEDIATE_BEC:' || !parsed.data) {
    const detail = (parsed.data && (parsed.data.message || parsed.data.Message)) || parsed.raw || 'audit pull failed';
    rememberAuditGap(div, {
      mailbox,
      unchecked: ['unified audit', 'file sharing', 'DLP', 'security alerts', 'sign-ins'],
    });
    return { PullNotes: [String(detail)], Days: window.days };
  }
  noteAuditCollected(div, mailbox);
  const blasts = asArray(parsed.data.Blasts || parsed.data.blasts);
  if (blasts.length) div.dataset.becBlasts = JSON.stringify(blasts);
  else delete div.dataset.becBlasts;
  return parsed.data;
}

function blastRecipientSection(div) {
  let blasts = [];
  try { blasts = JSON.parse(div?.dataset?.becBlasts || '[]'); } catch { blasts = []; }
  if (!Array.isArray(blasts)) return '';
  const blocks = [];
  blasts.forEach((blast) => {
    const recips = asArray(blast.Recipients || blast.recipients).map((r) => String(r || '').trim()).filter(Boolean);
    if (recips.length < 5) return;
    const subject = String(blast.Subject || blast.subject || '(no subject)').trim() || '(no subject)';
    const when = formatPresentedLocal(blast.Received || blast.received || '');
    const lines = [subject];
    if (when) lines.push(when);
    recips.forEach((addr) => lines.push(addr));
    blocks.push(lines.join('\n'));
  });
  if (!blocks.length) return '';
  return `DELIVERED BLAST RECIPIENTS\n\n${blocks.join('\n\n')}`;
}

function setBecStatus(div, text) {
  const el = div.querySelector('.becInvestigateStatus');
  if (el) el.textContent = text || '';
}

function hideBecRepull(div) {
  div.querySelectorAll('.becManualRepull').forEach((btn) => { btn.style.display = 'none'; });
  delete div.dataset.becRepull;
}

function offerBecRepull(div, days, logs) {
  const which = Array.isArray(logs) && logs.length ? `${logs.join(', ')} for ` : '';
  div.querySelectorAll('.becManualRepull').forEach((btn) => {
    btn.style.display = '';
    btn.disabled = false;
    btn.textContent = `Pull ${which}${days} days and review again`;
  });
}

function focusContainmentUser(div, mailbox) {
  const want = String(mailbox || '').trim().toLowerCase();
  const boxes = [...div.querySelectorAll('.containmentUser')];
  if (!want || !boxes.length) return;
  const match = boxes.find((box) => String(box.value || '').toLowerCase() === want);
  boxes.forEach((box) => { box.checked = box === match; });
  updateContainmentButtons(div);
}

function becActionsTaken(div, mailbox) {
  const clientNumber = div.closest('details.tenant')?.dataset?.client;
  const prior = tenantUiState.get(String(clientNumber)) || {};
  const want = String(mailbox || '').trim().toLowerCase();
  return asArray(prior.containment?.actions).filter((row) => {
    const upn = String(row.UPN || row.upn || '').trim().toLowerCase();
    return !upn || !want || upn === want;
  }).slice(-40);
}

async function requestBecReview(div, mailbox, audit, windowNote) {
  const data = await api('/api/triage/bec-review', {
    method: 'POST',
    body: JSON.stringify(becFactsFromMemory(div, mailbox, audit, windowNote)),
  }, 360000);
  return data.result || {};
}

async function runBecRepull(clientNumber, div) {
  let pending = null;
  try { pending = JSON.parse(div.dataset.becRepull || ''); } catch { pending = null; }
  if (!pending || !pending.mailbox) return;
  const longer = Math.min(90, parseInt(pending.days, 10) || 0);
  if (!longer) return;
  const logs = becRepullLogs({ repullLogs: pending.logs });
  const prior = div._becAudit || null;
  hideBecRepull(div);
  const again = {
    days: longer,
    start: new Date(Date.now() - longer * 24 * 60 * 60 * 1000).toISOString(),
    end: new Date().toISOString(),
  };
  const label = becRepullLabel(logs);
  const status = logs.length ? `Pulling ${longer} days of ${label}…` : `Pulling ${longer} days of message trace and audit…`;
  setBecStatus(div, status);
  setContainmentStatus(div, status);
  log(`Client ${clientNumber}: Threat check expanding to ${longer} days of ${label}.`);
  let next = null;
  try { next = await pullBecAudit(clientNumber, div, pending.mailbox, again, logs); } catch (e) {
    next = { PullNotes: [e.message], Days: longer };
  }
  const audit = logs.length ? mergeBecAudit(prior, next, logs) : next;
  div._becAudit = audit;
  if (!logs.length) div.dataset.becWindowStart = again.start;
  const review = await requestBecReview(div, pending.mailbox, audit, becFollowUpNote(again, logs));
  renderBecReview(div, review);
  const more = parseInt(review.repullDays, 10) || 0;
  if (more > longer) {
    const moreLogs = becRepullLogs(review);
    div._becAudit = audit;
    div.dataset.becRepull = JSON.stringify({ mailbox: pending.mailbox, days: Math.min(90, more), logs: moreLogs });
    offerBecRepull(div, Math.min(90, more), moreLogs);
    setBecStatus(div, `The investigation still needs ${becRepullLabel(moreLogs)} for ${Math.min(90, more)} days.`);
    return;
  }
  setBecStatus(div, 'Threat review finished.');
}

function scopedThreatMailbox(div) {
  const clientNumber = div.closest('details.tenant')?.dataset?.client;
  const validated = clientNumber ? getValidatedUsersForTenant(clientNumber, div) : [];
  if (validated.length === 1) return validated[0];
  const picked = getSelectedContainmentUsers(div);
  if (picked.length === 1) return picked[0];
  const terms = parseSearchTerms(div.querySelector('.userSearchInput')?.value);
  if (terms.length === 1 && terms[0].includes('@')) return terms[0];
  return '';
}

function fillThreatMailboxIfEmpty(div, mailbox) {
  const box = div.querySelector('.becMailbox');
  const value = String(mailbox || '').trim();
  if (box && value && !box.value.trim()) box.value = value;
}

async function runBecReview(clientNumber, div, opts = {}) {
  const fromMain = opts.from === 'main';
  const users = getSelectedContainmentUsers(div);
  let mailbox = '';
  if (fromMain) {
    mailbox = (div.querySelector('.becMailbox')?.value || '').trim();
    if (!mailbox) mailbox = scopedThreatMailbox(div);
  } else if (users.length === 1) {
    mailbox = users[0];
  } else {
    mailbox = (div.querySelector('.becMailbox')?.value || '').trim();
  }
  if (!mailbox) {
    const message = fromMain
      ? 'Enter one mailbox, or scope one user with Filter by users.'
      : 'Select one mailbox to review.';
    setBecStatus(div, message);
    setContainmentStatus(div, message, { reveal: !fromMain });
    return;
  }
  const graph = div.dataset.graphAuthenticated === '1';
  const exo = div.dataset.exchangeAuthenticated === '1';
  if (!graph && !exo) {
    const message = 'Graph or Exchange auth is required for a threat check.';
    setBecStatus(div, message);
    setContainmentStatus(div, message, { reveal: !fromMain });
    return;
  }
  if (fromMain) focusContainmentUser(div, mailbox);
  const mailboxBox = div.querySelector('.becMailbox');
  if (mailboxBox) mailboxBox.value = mailbox;
  div.dataset.becMailbox = mailbox;
  const ticketId = (div.querySelector('.ticketInput')?.value || '').trim();
  const pasted = (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || '';
  if (ticketId && !pasted) {
    try { await fetchManageTicket(clientNumber, div); } catch (e) {
      log(`Client ${clientNumber}: ticket fetch failed, checking the mailbox without it. ${e.message}`);
    }
  }
  const ticket = (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || '';
  const becWindow = becWindowFromTicket(ticket);
  div.dataset.becWindowStart = becWindow.start;
  containmentBatchConfirmed = true;
  try {
    const pulls = [];
    if (exo) pulls.push('list-rules', 'mailbox-status', 'list-mobile', 'list-transport', 'list-connectors');
    if (graph) pulls.push('list-mfa', 'list-devices', 'list-oauth', 'list-intune', 'list-roles', 'list-apps');
    for (const kind of pulls) {
      try { await runContainmentAction(clientNumber, div, kind); } catch (e) {
        log(`Client ${clientNumber}: ${kind} failed, continuing the threat check. ${e.message}`);
      }
    }
    hideBecRepull(div);
    const progress = `Threat check for ${mailbox}, last ${becWindow.days} day(s)…`;
    setBecStatus(div, progress);
    setContainmentStatus(div, progress, { reveal: !fromMain });
    let audit = null;
    try {
      audit = await pullBecAudit(clientNumber, div, mailbox, becWindow);
    } catch (e) {
      audit = { PullNotes: [e.message], Days: becWindow.days };
    }
    if (auditPullFailed(audit)) {
      const retryText = `Audit logs were not collected. Pulling unified audit, file sharing, DLP, and sign-ins again…`;
      setBecStatus(div, retryText);
      setContainmentStatus(div, retryText, { reveal: !fromMain });
      log(`Client ${clientNumber}: threat check is downloading the missing audit logs.`);
      try {
        audit = await pullBecAudit(clientNumber, div, mailbox, becWindow);
      } catch (e) {
        audit = { PullNotes: [e.message], Days: becWindow.days };
      }
    }
    const tz = presentedLocalZone();
    let windowNote = becWindow.hadTime
      ? `Ticket time was read as ${tz}. This pull is ${becWindow.days} days, ${formatPresentedLocal(becWindow.start)} through ${formatPresentedLocal(becWindow.end)} ${tz}. Log times below are ${tz}.`
      : `No event time was found. This pull is the last ${becWindow.days} days, ${formatPresentedLocal(becWindow.start)} through ${formatPresentedLocal(becWindow.end)} ${tz}. Log times below are ${tz}.`;
    let review = await requestBecReview(div, mailbox, audit, windowNote);
    renderBecReview(div, review);
    renderStoreSuggestions(div, review);
    div._becAudit = audit;
    const seenPulls = new Set();
    let followUps = 0;
    div.dataset.becExpanding = '1';
    try {
      while (followUps < 2) {
        const asked = parseInt(review.repullDays, 10) || 0;
        const askedLogs = becRepullLogs(review);
        const pulled = parseInt((audit && (audit.Days || audit.days)) || becWindow.days, 10) || becWindow.days;
        if (!(asked > pulled)) break;
        const longer = Math.min(90, asked);
        const pullKey = `${longer}:${askedLogs.join(',')}`;
        if (seenPulls.has(pullKey)) break;
        seenPulls.add(pullKey);
        const label = becRepullLabel(askedLogs);
        const auto = div.querySelector('.becAutoRepull')?.checked !== false;
        if (!auto) {
          div.dataset.becRepull = JSON.stringify({ mailbox, days: longer, logs: askedLogs });
          offerBecRepull(div, longer, askedLogs);
          const waitText = `The investigation still needs ${longer} days of ${label}. Pull that window when you are ready.`;
          setBecStatus(div, waitText);
          setContainmentStatus(div, waitText, { reveal: !fromMain });
          break;
        }
        const again = {
          days: longer,
          start: new Date(Date.now() - longer * 24 * 60 * 60 * 1000).toISOString(),
          end: new Date().toISOString(),
        };
        const againText = `The investigation needs ${label} going back ${longer} days. Pulling those logs…`;
        setBecStatus(div, againText);
        setContainmentStatus(div, againText, { reveal: !fromMain });
        log(`Client ${clientNumber}: Threat check expanding to ${longer} days of ${label}.`);
        const prior = audit;
        let next = null;
        try { next = await pullBecAudit(clientNumber, div, mailbox, again, askedLogs); } catch (e) {
          next = { PullNotes: [e.message], Days: longer };
        }
        audit = askedLogs.length ? mergeBecAudit(prior, next, askedLogs) : next;
        div._becAudit = audit;
        if (!askedLogs.length) div.dataset.becWindowStart = again.start;
        followUps += 1;
        review = await requestBecReview(div, mailbox, audit, becFollowUpNote(again, askedLogs));
        renderBecReview(div, review);
        renderStoreSuggestions(div, review);
      }
    } finally {
      delete div.dataset.becExpanding;
    }
    renderBecReview(div, review);
    renderStoreSuggestions(div, review);
  if (!div.dataset.becRepull) setBecStatus(div, 'Threat review finished.');
    if (fromMain) {
      div.querySelector('.becReviewWrap')?.scrollIntoView({ block: 'nearest' });
    } else {
      const panel = div.querySelector('details.containmentPanel');
      const wrap = div.querySelector('.containmentBecWrap');
      if (panel) panel.open = true;
      if (wrap) wrap.scrollIntoView({ block: 'nearest' });
    }
  } finally {
    containmentBatchConfirmed = false;
  }
}

async function runBecRemediations(clientNumber, div) {
  const selected = [...div.querySelectorAll('.containmentBecPick:checked')].map((box) => box.value).filter(Boolean);
  if (!selected.length) {
    setContainmentStatus(div, 'Select at least one remediation.', { reveal: true });
    return;
  }
  const users = getSelectedContainmentUsers(div);
  const mailbox = users[0] || div.dataset.containmentRulesUser || 'the selected mailbox';
  const labels = selected.map((id) => {
    const box = [...div.querySelectorAll('.containmentBecPick')].find((item) => item.value === id);
    const title = box?.parentElement?.querySelector('strong')?.textContent || id;
    return title;
  });
  if (!window.confirm(`Run ${selected.length} remediation(s) for ${mailbox}?\n\n${labels.join('\n')}\n\nThis uses the containment actions and changes the mailbox.`)) return;
  const kindFor = {
    'delete-rules': 'delete-rule',
    'remove-forward': 'remove-forward',
    'remove-delegates': 'remove-delegate',
    'enable-hold': 'enable-hold',
    revoke: 'revoke',
    block: 'block',
    'reset-password': 'reset-password',
  };
  containmentBatchConfirmed = true;
  try {
    for (const id of BEC_RUN_ORDER) {
      if (!selected.includes(id) || !kindFor[id]) continue;
      if (id === 'delete-rules') selectBecRuleBoxes(div);
      if (id === 'remove-forward') {
        div.querySelectorAll('.containmentForwardPick').forEach((box) => { box.checked = true; });
        updateContainmentButtons(div);
      }
      if (id === 'remove-delegates') {
        div.querySelectorAll('.containmentDelegatePick').forEach((box) => { box.checked = true; });
        updateContainmentButtons(div);
      }
      await runContainmentAction(clientNumber, div, kindFor[id]);
    }
  } finally {
    containmentBatchConfirmed = false;
  }
  setContainmentStatus(div, 'Selected BEC remediations finished. Check the status above for each step.', { reveal: true });
}

function wireContainmentPanel(div, clientNumber) {
  const savedCaps = tenantUiState.get(String(clientNumber))?.graphCapabilities;
  if (savedCaps) applyContainmentCapabilities(div, savedCaps);
  refreshContainmentUsers(div);
  updateContainmentButtons(div);
  div.querySelector('.containmentUserList')?.addEventListener('change', () => {
    updateContainmentButtons(div);
    const picked = getSelectedContainmentUsers(div);
    const box = div.querySelector('.becMailbox');
    if (box && picked.length === 1) box.value = picked[0];
  });
  const bind = (sel, kind) => {
    div.querySelector(sel)?.addEventListener('click', () => {
      withClientLock(clientNumber, async () => {
        focusClientLogTab(clientNumber);
        try {
          await runContainmentAction(clientNumber, div, kind);
        } catch (e) {
          log(`Client ${clientNumber}: containment error: ${e.message}`);
          setContainmentStatus(div, e.message, { reveal: true });
        }
      });
    });
  };
  bind('.containmentQuickLock', 'quick-lock');
  bind('.containmentQuickRestore', 'quick-restore');
  bind('.containmentSigninStatus', 'signin-status');
  bind('.containmentRevoke', 'revoke');
  bind('.containmentBlock', 'block');
  bind('.containmentUnblock', 'unblock');
  bind('.containmentResetPassword', 'reset-password');
  bind('.containmentAssignPasswordBtn', 'assign-password');
  bind('.containmentListMfa', 'list-mfa');
  bind('.containmentRevokeMfa', 'revoke');
  bind('.containmentDeleteMfa', 'delete-mfa');
  bind('.containmentListDevices', 'list-devices');
  bind('.containmentDeleteDevices', 'delete-devices');
  bind('.containmentListApps', 'list-apps');
  bind('.containmentDeleteApps', 'delete-apps');
  bind('.containmentRestrictedStatus', 'restricted-status');
  bind('.containmentUnrestrict', 'unrestrict');
  bind('.containmentListRules', 'list-rules');
  bind('.containmentDeleteRules', 'delete-rule');
  bind('.containmentMailboxStatus', 'mailbox-status');
  bind('.containmentSetForward', 'set-forward');
  bind('.containmentRemoveForward', 'remove-forward');
  bind('.containmentClearForward', 'clear-forward');
  bind('.containmentAddDelegate', 'add-delegate');
  bind('.containmentRemoveDelegate', 'remove-delegate');
  bind('.containmentListTransport', 'list-transport');
  bind('.containmentDeleteTransport', 'delete-transport');
  bind('.containmentListConnectors', 'list-connectors');
  bind('.containmentDeleteConnectors', 'delete-connectors');
  bind('.containmentReregisterMfa', 'reregister-mfa');
  bind('.containmentListOauth', 'list-oauth');
  bind('.containmentDeleteOauth', 'delete-oauth');
  bind('.containmentListMobile', 'list-mobile');
  bind('.containmentDeleteMobile', 'delete-mobile');
  bind('.containmentListIntune', 'list-intune');
  bind('.containmentRetireIntune', 'retire-intune');
  bind('.containmentWipeIntune', 'wipe-intune');
  bind('.containmentListFolders', 'list-folders');
  bind('.containmentDeleteFolders', 'delete-folders');
  bind('.containmentAutoreplyStatus', 'autoreply-status');
  bind('.containmentDisableAutoreply', 'disable-autoreply');
  bind('.containmentListJunk', 'list-junk');
  bind('.containmentDeleteJunk', 'delete-junk');
  bind('.containmentListElsewhere', 'list-elsewhere');
  bind('.containmentDeleteElsewhere', 'delete-elsewhere');
  bind('.containmentHoldStatus', 'hold-status');
  bind('.containmentEnableHold', 'enable-hold');
  bind('.containmentListOrgfwd', 'list-orgfwd');
  bind('.containmentDisableOrgfwd', 'disable-orgfwd');
  bind('.containmentListJournal', 'list-journal');
  bind('.containmentDeleteJournal', 'delete-journal');
  bind('.containmentListRoles', 'list-roles');
  bind('.containmentDeleteRoles', 'delete-roles');
  bind('.containmentListAppcreds', 'list-appcreds');
  bind('.containmentDeleteAppcreds', 'delete-appcreds');
  bind('.containmentListFlows', 'list-flows');
  bind('.containmentDeleteFlows', 'delete-flows');
  div.querySelector('.containmentUpdateGraphScopes')?.addEventListener('click', () => {
    const tid = div.querySelector('.appRegSelect')?.value || '';
    runUpdateGraphAppScopes(tid).catch((e) => log(`Update Graph App scopes: ${e.message}`));
  });
  div.querySelector('.containmentSavePacks')?.addEventListener('click', () => {
    saveContainmentPacks(clientNumber, div).catch((e) => {
      log(`Client ${clientNumber}: save containment zips failed: ${e.message}`);
      setContainmentStatus(div, e.message, { reveal: true });
    });
  });
  div.querySelector('.containmentClearUserPulls')?.addEventListener('click', () => {
    if (!confirmContainmentPopup('Clear per-user containment lists (MFA, mailbox, devices, rules, etc.) so you can pull the next user?\n\nTenant-wide lists stay. Zips already saved are not deleted.')) return;
    clearContainmentUserPulls(div);
    log(`Client ${clientNumber}: cleared per-user containment pulls.`);
  });
  const startBecCheck = () => {
    withClientLock(clientNumber, async () => {
      focusClientLogTab(clientNumber);
      try {
        await runBecReview(clientNumber, div);
      } catch (e) {
        log(`Client ${clientNumber}: Threat check error: ${e.message}`);
        setContainmentStatus(div, e.message, { reveal: true });
      }
    });
  };
  div.querySelector('.containmentBecCheck')?.addEventListener('click', startBecCheck);
  div.querySelector('.containmentBecReview')?.addEventListener('click', startBecCheck);
  div.querySelector('.becInvestigateRun')?.addEventListener('click', () => {
    withClientLock(clientNumber, async () => {
      focusClientLogTab(clientNumber);
      try {
        await runBecReview(clientNumber, div, { from: 'main' });
      } catch (e) {
        log(`Client ${clientNumber}: Threat check error: ${e.message}`);
        setBecStatus(div, e.message);
        setContainmentStatus(div, e.message);
        const mainPre = div.querySelector('.becReviewText');
        const mainWrap = div.querySelector('.becReviewWrap');
        if (mainPre && !mainPre.textContent) mainPre.textContent = e.message;
        if (mainWrap && mainPre && mainPre.textContent) mainWrap.style.display = 'block';
      }
    });
  });
  div.querySelectorAll('.becManualRepull').forEach((btn) => {
    btn.addEventListener('click', () => {
      withClientLock(clientNumber, async () => {
        focusClientLogTab(clientNumber);
        try {
          await runBecRepull(clientNumber, div);
        } catch (e) {
          log(`Client ${clientNumber}: Threat check error: ${e.message}`);
          setBecStatus(div, e.message);
        }
      });
    });
  });
  div.querySelector('.becReviewCopy')?.addEventListener('click', async () => {
    const text = div.querySelector('.becReviewText')?.textContent || '';
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      log(`Client ${clientNumber}: BEC review copied.`);
    } catch (err) {
      log(`Client ${clientNumber}: could not reach the clipboard. Click the page once, then copy again.`);
    }
  });
  div.querySelector('.containmentBecAll')?.addEventListener('click', () => {
    div.querySelectorAll('.containmentBecPick').forEach((box) => { box.checked = true; });
  });
  div.querySelector('.containmentBecNone')?.addEventListener('click', () => {
    div.querySelectorAll('.containmentBecPick').forEach((box) => { box.checked = false; });
  });
  div.querySelector('.containmentBecRun')?.addEventListener('click', () => {
    withClientLock(clientNumber, async () => {
      focusClientLogTab(clientNumber);
      try {
        await runBecRemediations(clientNumber, div);
      } catch (e) {
        log(`Client ${clientNumber}: BEC remediation error: ${e.message}`);
        setContainmentStatus(div, e.message, { reveal: true });
      }
    });
  });
  div.querySelector('.containmentBecCopy')?.addEventListener('click', async () => {
    const text = div.querySelector('.containmentBecText')?.textContent || '';
    if (!text) {
      setContainmentStatus(div, 'Run a review before copying.', { reveal: true });
      return;
    }
    try {
      await navigator.clipboard.writeText(text);
      setContainmentStatus(div, 'BEC review copied.', { reveal: true });
    } catch (e) {
      setContainmentStatus(div, 'Could not reach the clipboard. Click the page once, then copy again.', { reveal: true });
    }
  });
  restoreContainmentOutput(div, tenantUiState.get(String(clientNumber)), clientNumber);
}

function normalizeResponse(value) {

  return (value == null ? '' : String(value)).trim();

}

async function withClientLock(clientNumber, fn) {

  const key = String(clientNumber);

  if (clientBusy.has(key)) {

    log(`Client ${clientNumber}: busy — wait for the current operation to finish.`);
    const body = getTenantBodyEl(clientNumber);
    if (body) setContainmentStatus(body, 'Busy — wait for the current containment command to finish.', { skipSave: true });

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

  div.querySelectorAll('button.graphAuth, button.exoAuth, button.generateReports, button.validateUsers, button.collectHistorical').forEach(btn => {

    btn.disabled = disabled;

  });

  updateContainmentButtons(div);
  updateCaCatalogButtons(div);

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

  body.daysBack = sessionRelativeDays();

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

    const relAmount = document.getElementById('relAmount');

    const relUnit = document.getElementById('relUnit');

    if (mt && rs.MessageTraceDaysBack != null) mt.value = String(rs.MessageTraceDaysBack);

    if (si && rs.SignInLogsDaysBack != null) si.value = String(rs.SignInLogsDaysBack);

    // The session stores a resolved day count, so restore as days (or Max at the 90-day cap).
    if (session.daysBack != null) {

      const days = Math.max(1, parseInt(session.daysBack, 10) || 7);

      if (relUnit) relUnit.value = days >= RELATIVE_MAX_DAYS ? 'max' : 'days';

      if (relAmount) relAmount.value = String(Math.min(days, RELATIVE_MAX_DAYS));

    }

    updateSessionTimeframeUi();

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

function isAppRegManual(div) {
  return div?.dataset.appRegManual === '1';
}

function setAppRegManual(div, manual) {
  if (div) div.dataset.appRegManual = manual ? '1' : '0';
}

function autoSelectAppRegForExo(div, exoTenantId) {

  if (!exoTenantId || !div) return;

  const sel = div.querySelector('.appRegSelect');

  if (!sel || isAppRegManual(div)) return;

  const match = appRegistrations.find(a => (a.tenantId || '').toLowerCase() === exoTenantId.toLowerCase());

  if (match) sel.value = match.tenantId;

}

function renderTenants(session) {

  const scrollX = window.scrollX;
  const scrollY = window.scrollY;
  snapshotOpenContainmentPanels();

  tenantsEl.innerHTML = '';

  if (!session?.tenants?.length) {

    tenantsEl.innerHTML = '<p class="muted">No tenants yet.</p>';
    pinViewport(scrollX, scrollY);

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
    if (t.exoOrganizationName) div.dataset.exoOrganizationName = t.exoOrganizationName;
    if (t.graphTenantName) div.dataset.graphTenantName = t.graphTenantName;
    if (t.outputFolder) div.dataset.outputFolder = t.outputFolder;
    if (t.lastResponse) div.dataset.lastResponse = normalizeResponse(t.lastResponse);
    const appRegOptions = appRegistrations.map(a =>

      `<option value="${a.tenantId || ''}">${a.displayText}</option>`

    ).join('');

    const canGenerate = tenantHasRequiredAuth(t, session) && !t.reportInProgress;

    const lastResp = normalizeResponse(t.lastResponse);

    const errorHtml = (lastResp.includes('_FAILED') || lastResp.includes('ERROR'))

      ? `<div class="outputPath" style="color:#cf222e">Last: ${lastResp.substring(0, 220)}${lastResp.length > 220 ? '…' : ''}</div>`

      : '';

    const inProgressGraph = lastResp === 'GRAPH_AUTH_STARTED';

    const inProgressExo = lastResp === 'EXCHANGE_AUTH_STARTED';

    const existing = tenantUiState.get(String(t.clientNumber)) || {};
    if (t.uiState && typeof t.uiState === 'object') {
      const merged = { ...existing, ...t.uiState };
      const localValidated = existing.validatedUsers;
      const serverValidated = t.uiState.validatedUsers;
      if (Array.isArray(localValidated) && localValidated.length
          && (!Array.isArray(serverValidated) || !serverValidated.length)) {
        merged.validatedUsers = localValidated;
      }
      const localContainment = resolveContainmentState(t.clientNumber, existing);
      const serverContainment = t.uiState.containment;
      if (hasContainmentPayload(localContainment)) {
        merged.containment = {
          ...(serverContainment && typeof serverContainment === 'object' ? serverContainment : {}),
          ...localContainment,
        };
      } else if (hasContainmentPayload(serverContainment)) {
        merged.containment = serverContainment;
      }
      tenantUiState.set(String(t.clientNumber), merged);
    } else {
      const stored = resolveContainmentState(t.clientNumber, existing);
      if (stored) {
        tenantUiState.set(String(t.clientNumber), { ...existing, containment: stored });
      }
    }

    const uiAfterMerge = tenantUiState.get(String(t.clientNumber)) || existing;
    const folders = [];
    if (Array.isArray(uiAfterMerge.reportFolders)) folders.push(...uiAfterMerge.reportFolders.filter(Boolean));
    if (t.outputFolder && !folders.includes(t.outputFolder)) folders.push(t.outputFolder);
    const outputHtml = folders.length
      ? `<div class="outputPath muted">${folders.length > 1 ? `${folders.length} report packs. Latest: ` : 'Reports: '}${escapeHtml(folders[folders.length - 1])}</div>`
      : '';

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

      ${buildBecInvestigatePanelHtml()}

      <div class="row">

        <label>App reg tenant

          <select class="appRegSelect"><option value="">${autoWcmLabel}</option>${appRegOptions}</select>

        </label>

        <label><input type="checkbox" class="interactiveCheck" /> Force interactive Graph</label>

      </div>

      <div class="sectionLabel">Timeframe</div>

      <div class="row">

        <label class="relAmountWrap">Last <input type="number" class="relAmount" min="1" max="90" value="${sessionRelativeAmount()}" style="min-width:4rem" /></label>

        <label>Unit <select class="relUnit">${relativeUnitOptionsHtml(sessionRelativeUnit(), true)}</select></label>

        <span class="relHint muted"></span>

      </div>

      <div class="row customRangeRow" style="display:none">

        <label>Start <input type="datetime-local" class="dateStart" value="${defaultDateStartValue()}" /></label>

        <label>End <input type="datetime-local" class="dateEnd" value="${defaultDateEndValue()}" /></label>

      </div>

      <details class="tenantReportExports collapsible">

        <summary>Report exports <span class="reportExportsHint muted">(session defaults)</span></summary>

        <div class="collapsible-body">

          <div class="row" style="flex-wrap:wrap;gap:0.5rem;align-items:end">
            <label>Export scope
              <select class="tenantReportExportMode">
                <option value="session" selected>Session defaults</option>
                <option value="preset">Investigation preset</option>
                <option value="custom">Custom selection</option>
              </select>
            </label>
            <label class="tenantReportPresetWrap" style="display:none">Preset
              <select class="tenantReportPreset"></select>
            </label>
            <input type="checkbox" class="useSessionReportDefaults" checked style="display:none" aria-hidden="true" />
          </div>

          <div class="tenantReportExportsCustom" style="display:none;margin-top:0.5rem">

            <p class="muted" style="margin:0.35rem 0">Custom exports for this client only. Changing checkboxes clears preset-scoped UAL types (uses default UAL set).</p>

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

      ${buildContainmentPanelHtml()}

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
      <div class="row">
        <button type="button" class="suggestLogs small">Suggest logs from this ticket</button>
        <button type="button" class="suggestMoreLogs small" ${t.outputFolder ? '' : 'disabled'}>Ask if more logs are needed</button>
        <button type="button" class="writeMemberberry small">Write notes</button>
      </div>
      <div class="logPlan" style="display:none;margin:0.35rem 0 0.6rem">
        <pre class="logPlanText" style="white-space:pre-wrap;background:#f6f8fa;padding:0.6rem;border-radius:6px;margin:0.35rem 0"></pre>
        <button type="button" class="applyLogPlan primary small" disabled>Apply this plan</button>
      </div>

      ${buildSecurityIntegrationsPanelHtml()}

      ${buildCurateLogsPanelHtml(Boolean(t.outputFolder))}

      ${buildCaCatalogPanelHtml()}

      ${buildEmailTriagePanelHtml()}

      <div>

        <button class="exoAuth">Exchange Auth</button>

        <button class="graphAuth">Graph Auth</button>

        <button type="button" class="caCatalogJump small" title="Scroll to the read-only Conditional Access catalog on this tenant">CA catalog</button>

        <button class="generateReports primary" ${canGenerate ? '' : 'disabled'}>${t.reportInProgress ? 'Generating…' : (t.outputFolder ? 'Generate another report pack' : 'Generate Reports')}</button>

        <button class="collectHistorical" title="Download completed 90-day historical message trace searches queued by earlier runs">Collect historical trace</button>

        <button class="openReports success" ${t.outputFolder ? '' : 'disabled'}>Open Reports</button>

        <button class="restartWorker">Restart worker</button>

        <button class="showConsoleRestart">Show console</button>

        <button class="statusBtn">Tail status</button>

        <button class="openStatusFile">Open status file</button>

        <button class="graphDisconnect">Log out Graph</button>

        <button class="resetAuth">Reset auth</button>

        <button class="investigateLogs" ${t.outputFolder ? '' : 'disabled'}>Investigate logs</button>

        <select class="investigateFolderSelect" ${t.outputFolder ? '' : 'disabled'} title="Pack to investigate"></select>

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

    div.querySelector('.caCatalogJump')?.addEventListener('click', () => {
      const panel = div.querySelector('details.caCatalogPanel');
      if (panel) {
        panel.open = true;
        panel.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      }
    });

    div.querySelector('.exoAuth').addEventListener('click', () => exoAuth(t.clientNumber, div));

    div.querySelector('.validateUsers').addEventListener('click', () => validateUsers(t.clientNumber, div));

    div.querySelector('.collectHistorical').addEventListener('click', () => collectHistoricalTrace(t.clientNumber, div));

    wireContainmentPanel(div, t.clientNumber);
    wireCaCatalogPanel(div, t.clientNumber);
    wireEmailTriagePanel(div, t.clientNumber);

    div.querySelector('.fetchTicket')?.addEventListener('click', () => fetchManageTicket(t.clientNumber, div));
    div.querySelector('.suggestLogs')?.addEventListener('click', () => suggestLogsFromTicket(t.clientNumber, div).catch((e) => log(`Client ${t.clientNumber}: log plan error: ${e.message}`)));
    div.querySelector('.suggestMoreLogs')?.addEventListener('click', () => suggestMoreLogs(t.clientNumber, div).catch((e) => log(`Client ${t.clientNumber}: log plan error: ${e.message}`)));
    div.querySelector('.applyLogPlan')?.addEventListener('click', () => applyLogPlan(t.clientNumber, div));
    div.querySelectorAll('.writeMemberberry').forEach((btn) => {
      btn.addEventListener('click', () => writeMemberberryNote(t.clientNumber, div).catch((e) => log(`Client ${t.clientNumber}: Memberberry note error: ${e.message}`)));
    });
    div.querySelector('.copyMemberberry')?.addEventListener('click', () => copyBecNote(t.clientNumber, div, '.memberberryText', 'internal note'));
    div.querySelector('.copyBecClient')?.addEventListener('click', () => copyBecNote(t.clientNumber, div, '.becClientText', 'client email'));
    div.querySelector('.copyBecIncident')?.addEventListener('click', () => copyBecNote(t.clientNumber, div, '.becIncidentText', 'incident report'));
    div.querySelector('.storeMbSave')?.addEventListener('click', () => saveMemberberrySuggestion(t.clientNumber, div).catch((e) => log(`Client ${t.clientNumber}: Memberberry save error: ${e.message}`)));
    div.querySelector('.storeMvSave')?.addEventListener('click', () => saveMemorySuggestion(t.clientNumber, div).catch((e) => log(`Client ${t.clientNumber}: memory vault save error: ${e.message}`)));

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

    fillInvestigateFolderSelect(div, t);
    div.querySelector('.investigateLogs')?.addEventListener('click', () => openInvestigateLogs(t.clientNumber, div));

    restoreTenantUiState(t.clientNumber, div);
    div.querySelector('.appRegSelect')?.addEventListener('change', () => {
      const picked = div.querySelector('.appRegSelect')?.value || '';
      setAppRegManual(div, Boolean(picked));
      saveTenantUiState(t.clientNumber, div);
    });
    wireTenantTimeframePanel(div, t.clientNumber);
    wireTenantReportExportsPanel(div, t.clientNumber);
    wireSecurityIntegrationsPanel(div, t.clientNumber);
    wireCurateLogsPanel(div, t.clientNumber);
    applyTenantSummary(details, t, div);
    details.addEventListener('toggle', (e) => {
      if (e.target !== details) return;
      setTenantCollapsed(t.clientNumber, !details.open);
    });
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

  pinViewport(scrollX, scrollY);

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

    const extractedFail = failPrefix ? extractWorkerTokenResponse(resp, [failPrefix]) : '';
    const extractedOk = extractWorkerTokenResponse(resp, successPrefixes);
    const matchedFail = Boolean(failPrefix && extractedFail.startsWith(failPrefix));
    const matchedOk = successPrefixes.some((prefix) => extractedOk.startsWith(prefix));

    if (matchedFail) {

      throw new Error(extractedFail);

    }

    if (matchedOk) {

      return extractedOk;

    }

    if (startedToken && (resp === startedToken || resp.startsWith(`${startedToken}`))) {

      sawStarted = true;
      continue;

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

      const sel = div?.querySelector('.appRegSelect');
      const interactive = Boolean(div?.querySelector('.interactiveCheck')?.checked);
      let tenantId = null;
      if (isAppRegManual(div) && sel?.value) {
        tenantId = sel.value;
        log(`Client ${clientNumber}: using selected App reg ${tenantId}.`);
      } else {
        try {
          const session = await api('/api/session');
          const tenant = session.tenants?.find(x => Number(x.clientNumber) === Number(clientNumber));
          if (tenant?.exoTenantId) {
            tenantId = tenant.exoTenantId;
            autoSelectAppRegForExo(div, tenant.exoTenantId);
            const label = appRegTenantLabel(tenantId);
            log(`Client ${clientNumber}: auto WCM using EXO tenant ${label} (${tenantId}).`);
          }
        } catch { /* ignore */ }
      }
      if (!tenantId && !interactive) {
        log(`Client ${clientNumber}: auto WCM has no Exchange tenant ID. Run Exchange Auth first, or pick App reg tenant. Not guessing among WCM entries.`);
        return;
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

        const users = normalizeUserList(result.Users);

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

async function collectHistoricalTrace(clientNumber, div) {

  await withClientLock(clientNumber, async () => {

    if (!await requireLiveWorker(clientNumber, { actionLabel: 'collect historical trace' })) {
      return;
    }

    const session = await api('/api/session');
    const tenant = session.tenants?.find(x => Number(x.clientNumber) === Number(clientNumber));
    if (!tenant?.exchangeAuthenticated) {
      const msg = `Client ${clientNumber}: Exchange Auth is required to collect historical message trace reports.`;
      log(msg);
      window.alert(msg);
      return;
    }

    const btn = div.querySelector('.collectHistorical');
    if (btn) btn.disabled = true;

    try {

      log(`Client ${clientNumber}: checking for completed historical message trace reports…`);

      const initial = await api(`/api/tenants/${clientNumber}/command`, {
        method: 'POST',
        body: workerCommandBody('COLLECT_HISTORICAL_TRACE'),
      });

      let final = normalizeResponse(initial.response);

      if (!final || final === 'COLLECT_HISTORICAL_TRACE_STARTED') {
        final = await pollWorkerResponse(
          clientNumber,
          'COLLECT_HISTORICAL_TRACE_STARTED',
          ['COLLECT_HISTORICAL_TRACE_SUCCESS:'],
          'COLLECT_HISTORICAL_TRACE_FAILED:',
          600,
          'collecting historical message trace'
        );
      }

      if (final.startsWith('COLLECT_HISTORICAL_TRACE_SUCCESS:')) {

        const result = JSON.parse(final.replace('COLLECT_HISTORICAL_TRACE_SUCCESS:', ''));

        if (result.RowCount > 0) {
          log(`Client ${clientNumber}: collected ${result.RowCount} historical message trace row(s) -> ${result.CsvPath}`);
          if (result.PendingJobs > 0) {
            log(`Client ${clientNumber}: ${result.PendingJobs} historical search(es) still running - collect again later for the rest.`);
          }
        } else if (result.PendingJobs > 0) {
          log(`Client ${clientNumber}: ${result.PendingJobs} historical search(es) still running on Microsoft's queue. These take hours - try again later.`);
        } else {
          log(`Client ${clientNumber}: no historical searches queued. Generate with a window longer than the live API covers to request them.`);
        }

      } else {

        log(`Client ${clientNumber}: ${final}`);

      }

    } catch (e) {

      log(`Client ${clientNumber}: Collect historical trace failed: ${e.message}`);

    } finally {

      if (btn) btn.disabled = false;

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

function listInvestigateFolders(div, t) {
  const set = [];
  const add = (p) => {
    if (p && !set.includes(p)) set.push(p);
  };
  add(t?.outputFolder);
  add(div?.dataset?.outputFolder);
  add(div?.dataset?.curatedFolder);
  const ui = tenantUiState.get(String(t?.clientNumber || ''));
  (ui?.reportFolders || []).forEach(add);
  return set;
}

function fillInvestigateFolderSelect(div, t) {
  const sel = div?.querySelector('.investigateFolderSelect');
  if (!sel) return;
  const folders = listInvestigateFolders(div, t);
  const current = sel.value;
  sel.innerHTML = folders.map((p) => {
    const name = p.replace(/\\/g, '/').split('/').slice(-2).join('/') || p;
    return `<option value="${p.replace(/"/g, '&quot;')}">${name}</option>`;
  }).join('');
  if (current && folders.includes(current)) sel.value = current;
  else if (t?.outputFolder && folders.includes(t.outputFolder)) sel.value = t.outputFolder;
  sel.disabled = folders.length === 0;
  const btn = div.querySelector('.investigateLogs');
  if (btn) btn.disabled = folders.length === 0;
}

function tenantInvestigateWindow(div) {
  if (!div) return null;
  if (getTenantRelativeUnit(div) === 'custom') {
    const start = div.querySelector('.dateStart')?.value || '';
    const end = div.querySelector('.dateEnd')?.value || '';
    if (!start || !end) return null;
    return { start, end };
  }
  const { start, end } = resolveRelativeWindow(getTenantRelativeAmount(div), getTenantRelativeUnit(div));
  return { start: toDateTimeLocalValue(start), end: toDateTimeLocalValue(end) };
}

function investigateLogsSearchParams(clientNumber, div, path) {
  const q = new URLSearchParams({ path, client: String(clientNumber) });
  const users = getSelectedContainmentUsers(div);
  if (users.length === 1) q.set('address', users[0]);
  const range = tenantInvestigateWindow(div);
  if (range?.start) q.set('from', range.start);
  if (range?.end) q.set('to', range.end);
  return q;
}

function openInvestigateLogs(clientNumber, div) {
  const sel = div.querySelector('.investigateFolderSelect');
  const path = (sel?.value || div.dataset.outputFolder || '').trim();
  if (!path) {
    log(`Client ${clientNumber}: no report folder to investigate.`);
    return;
  }
  const q = investigateLogsSearchParams(clientNumber, div, path);
  window.location.href = `investigate.html?${q}`;
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
    const days = sessionRelativeDays();
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
  fromDataset = normalizeUserList(fromDataset);
  if (fromDataset.length) {
    if (div) div.dataset.validatedUsers = JSON.stringify(fromDataset);
    return fromDataset;
  }

  const fromState = normalizeUserList(tenantUiState.get(String(clientNumber))?.validatedUsers);
  if (fromState.length) {
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

function logPlanLabel(key) {
  for (const group of REPORT_EXPORT_GROUPS) {
    const item = group.items.find((row) => row.key === key);
    if (item) return item.label;
  }
  return key;
}

function renderLogPlan(div, plan) {
  const wrap = div.querySelector('.logPlan');
  const pre = div.querySelector('.logPlanText');
  const apply = div.querySelector('.applyLogPlan');
  if (!wrap || !pre || !apply) return;
  const logs = asArray(plan && (plan.logs || plan.Logs));
  const days = plan && (plan.days || plan.Days);
  const note = (plan && (plan.note || plan.Note)) || '';
  const more = plan && (plan.more === true || plan.More === true);
  const lines = [];
  if (!logs.length || (plan.phase === 'followup' && !more)) {
    lines.push(note || 'No extra logs suggested.');
    apply.disabled = true;
  } else {
    lines.push(`${days} day${days === 1 ? '' : 's'}`);
    logs.forEach((key) => lines.push(logPlanLabel(key)));
    if (note) lines.push(note);
    apply.disabled = false;
  }
  pre.textContent = lines.join('\n');
  wrap.style.display = 'block';
  div.dataset.logPlan = JSON.stringify({ days, logs, phase: plan.phase || 'initial' });
  renderStoreSuggestions(div, plan);
}

async function requestLogPlan(clientNumber, div, phase) {
  let ticket = (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || '';
  const ticketId = (div.querySelector('.ticketInput')?.value || '').trim();
  if (!ticket && ticketId) {
    await fetchManageTicket(clientNumber, div);
    ticket = (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || '';
  }
  if (!ticket) {
    log(`Client ${clientNumber}: paste a ticket or enter a ticket # first.`);
    return null;
  }
  const body = { ticket, phase };
  if (phase === 'followup') {
    const folder = div.dataset.outputFolder || tenantUiState.get(String(clientNumber))?.outputFolder || '';
    if (!folder) {
      log(`Client ${clientNumber}: generate a report pack before asking for more logs.`);
      return null;
    }
    body.folder = folder;
  }
  log(`Client ${clientNumber}: asking for a log plan…`);
  const data = await api('/api/triage/log-plan', {
    method: 'POST',
    body: JSON.stringify(body),
  }, 120000);
  return data.result || null;
}

async function suggestLogsFromTicket(clientNumber, div) {
  const plan = await requestLogPlan(clientNumber, div, 'initial');
  if (!plan) return;
  renderLogPlan(div, plan);
  log(`Client ${clientNumber}: log plan ready (${asArray(plan.logs).length} log(s), ${plan.days} day(s)).`);
}

async function suggestMoreLogs(clientNumber, div) {
  const plan = await requestLogPlan(clientNumber, div, 'followup');
  if (!plan) return;
  renderLogPlan(div, plan);
  const extra = asArray(plan.logs).length;
  log(`Client ${clientNumber}: ${plan.more && extra ? `${extra} more log(s) suggested.` : 'No extra logs suggested.'}`);
}

function renderStoreSuggestions(div, result) {
  const wrap = div.querySelector('.storeSuggestions');
  if (!wrap || !result) return;
  const client = String(result.storeClient || result.StoreClient || '').trim();
  const note = String(result.storeNote || result.StoreNote || '').trim();
  const memory = String(result.storeMemory || result.StoreMemory || '').trim();
  if (!client && !note && !memory) return;
  const clientBox = wrap.querySelector('.storeMbClient');
  const noteBox = wrap.querySelector('.storeMbNote');
  const memoryBox = wrap.querySelector('.storeMvNote');
  if (client && clientBox) clientBox.value = client;
  else if (!clientBox?.value && note) {
    const guessed = getTenantCompanyName(div, div.closest('details.tenant')?.dataset?.client);
    if (clientBox && guessed) clientBox.value = guessed;
  }
  if (note && noteBox) noteBox.value = note;
  if (memory && memoryBox) memoryBox.value = memory;
  wrap.style.display = 'block';
}

function renderMemberberryNote(div, result) {
  const wrap = div.querySelector('.becNotes');
  const internal = div.querySelector('.memberberryText');
  const client = div.querySelector('.becClientText');
  const incident = div.querySelector('.becIncidentText');
  const refs = div.querySelector('.memberberryRefs');
  if (internal) internal.textContent = (result && (result.internalText || result.text)) || '';
  if (client) client.textContent = (result && result.clientText) || '';
  let incidentText = (result && result.incidentText) || '';
  const blastText = blastRecipientSection(div);
  if (blastText && !/DELIVERED BLAST RECIPIENTS/i.test(incidentText)) {
    incidentText = `${incidentText.trim()}\n\n${blastText}`.trim();
  }
  if (incident) incident.textContent = incidentText;
  if (refs) refs.textContent = (result && result.references) || '';
  const any = Boolean(internal?.textContent || client?.textContent || incident?.textContent);
  if (wrap) wrap.style.display = any ? 'block' : 'none';
}

function ticketContactName(div) {
  const ticket = (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || '';
  const match = ticket.match(/^\s*Contact:\s*(\S.+)$/im);
  return match ? match[1].trim() : '';
}

async function writeMemberberryNote(clientNumber, div) {
  if (div.dataset.becExpanding === '1') {
    setBecStatus(div, 'Still pulling the logs this investigation needs.');
    return;
  }
  const review = (div.querySelector('.becReviewText')?.textContent || div.querySelector('.containmentBecText')?.textContent || '').trim();
  if (!review) {
    log(`Client ${clientNumber}: finish the mailbox review before writing the notes.`);
    setBecStatus(div, 'Finish the threat check before writing the notes.');
    return;
  }
  const ticket = (div.querySelector('.ticketPaste')?.value || '').trim() || div.dataset.ticketContent || '';
  const ticketNumber = (div.querySelector('.ticketInput')?.value || '').trim();
  const company = getTenantCompanyName(div, clientNumber);
  const contact = ticketContactName(div);
  const mailbox = div.dataset.becMailbox || (div.querySelector('.becMailbox')?.value || '').trim();
  setBecStatus(div, 'Writing the internal note, client email, and incident report…');
  log(`Client ${clientNumber}: writing the internal note, client email, and incident report…`);
  const data = await api('/api/triage/memberberry-note', {
    method: 'POST',
    body: JSON.stringify({
      review,
      ticket,
      ticketNumber,
      company,
      contact,
      audience: 'all',
      actions: becActionsTaken(div, mailbox),
    }),
  }, 360000);
  renderMemberberryNote(div, data.result || {});
  div.querySelector('.becNotes')?.scrollIntoView({ block: 'nearest' });
  renderStoreSuggestions(div, data.result || {});
  setBecStatus(div, 'Notes ready.');
  log(`Client ${clientNumber}: internal note, client email, and incident report ready.`);
}

function setStoreStatus(div, text) {
  const el = div.querySelector('.storeSuggestStatus');
  if (el) el.textContent = text || '';
}

async function saveMemberberrySuggestion(clientNumber, div) {
  const client = (div.querySelector('.storeMbClient')?.value || '').trim();
  const note = (div.querySelector('.storeMbNote')?.value || '').trim();
  if (!client || !note) {
    setStoreStatus(div, 'Enter the client name and the note before adding it.');
    return;
  }
  setStoreStatus(div, 'Adding the note to Memberberry…');
  const data = await api('/api/triage/memberberry-save', {
    method: 'POST',
    body: JSON.stringify({ client, note }),
  }, 30000);
  const savedName = (data.result && (data.result.client || data.result.Client)) || client;
  setStoreStatus(div, `Added to Memberberry for ${savedName}.`);
  log(`Client ${clientNumber}: Memberberry note added for ${savedName}.`);
}

async function saveMemorySuggestion(clientNumber, div) {
  const text = (div.querySelector('.storeMvNote')?.value || '').trim();
  if (!text) {
    setStoreStatus(div, 'Enter the memory before adding it.');
    return;
  }
  setStoreStatus(div, 'Adding the note to memory vault…');
  const data = await api('/api/triage/memory-save', {
    method: 'POST',
    body: JSON.stringify({ text }),
  }, 60000);
  const duplicate = data.result && (data.result.duplicate === true || data.result.Duplicate === true);
  setStoreStatus(div, duplicate ? 'That memory was already in the soc space.' : 'Added to the soc space in memory vault.');
  log(`Client ${clientNumber}: ${duplicate ? 'memory vault already had that note.' : 'memory vault note added.'}`);
}

async function copyBecNote(clientNumber, div, selector, label) {
  const text = div.querySelector(selector)?.textContent || '';
  if (!text) {
    log(`Client ${clientNumber}: write the ${label} before copying it.`);
    return;
  }
  try {
    await navigator.clipboard.writeText(text);
    log(`Client ${clientNumber}: ${label} copied.`);
  } catch (err) {
    log(`Client ${clientNumber}: could not reach the clipboard. Click the page once, then copy again.`);
  }
}

function applyLogPlan(clientNumber, div) {
  let plan = null;
  try { plan = JSON.parse(div.dataset.logPlan || ''); } catch { plan = null; }
  const logs = asArray(plan && plan.logs);
  const days = Math.max(1, Math.min(90, parseInt(plan && plan.days, 10) || 1));
  if (!logs.length) {
    log(`Client ${clientNumber}: there is no log plan to apply.`);
    return;
  }
  const modeSel = div.querySelector('.tenantReportExportMode');
  const customBody = div.querySelector('.tenantReportExportsCustom');
  if (modeSel) {
    modeSel.value = 'custom';
    modeSel.dispatchEvent(new Event('change'));
  }
  const selections = {};
  for (const group of REPORT_EXPORT_GROUPS) {
    for (const item of group.items) selections[item.key] = logs.includes(item.key);
  }
  selections.MessageTraceDaysBack = days;
  selections.SignInLogsDaysBack = days;
  applyReportSelectionsToContainer(customBody, selections);
  const unit = div.querySelector('.relUnit');
  const amount = div.querySelector('.relAmount');
  if (unit) unit.value = 'days';
  if (amount) amount.value = String(days);
  updateTenantTimeframeUi(div);
  applyTenantRelativeWindow(div);
  const details = div.querySelector('details.tenantReportExports');
  if (details) details.open = true;
  saveTenantUiState(clientNumber, div);
  log(`Client ${clientNumber}: applied ${logs.length} log(s) for ${days} day(s). Generate Reports uses this selection.`);
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

    }, 90000);

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

    applyTenantRelativeWindow(div);

    const dateStartEl = div.querySelector('.dateStart');
    const dateEndEl = div.querySelector('.dateEnd');

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
        const priorFolder = div.dataset.outputFolder || '';

        log(`Client ${clientNumber}: Reports saved to ${path}. Graph and EXO stay connected — generate again anytime.`);

        if (priorFolder && path && priorFolder !== path) {
          try {
            const copied = await api(`/api/tenants/${clientNumber}/copy-containment-zips`, {
              method: 'POST',
              body: JSON.stringify({ from: priorFolder, to: path }),
            });
            if (copied.files?.length) log(`Client ${clientNumber}: copied ${copied.files.length} containment zip(s) into the new pack.`);
          } catch (copyErr) {
            log(`Client ${clientNumber}: could not copy containment zips: ${copyErr.message}`);
          }
        }

        div.dataset.outputFolder = path;
        rememberReportFolder(clientNumber, path);

        const openBtn = div.querySelector('.openReports');
        const investigateBtn = div.querySelector('.investigateLogs');
        const investigateSel = div.querySelector('.investigateFolderSelect');
        if (openBtn) openBtn.disabled = false;
        if (investigateBtn) investigateBtn.disabled = false;
        if (investigateSel) investigateSel.disabled = false;
        fillInvestigateFolderSelect(div, { clientNumber, outputFolder: path });
        div.querySelectorAll('.loadCurateFacets, .previewCurate, .exportCurate, .selectSuggestedWan, .clearWanSelection, .applyWanPaste').forEach((btn) => {
          btn.disabled = false;
        });

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
        btn.textContent = div.dataset.outputFolder ? 'Generate another report pack' : 'Generate Reports';
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
    const logPanel = getWorkerLogPanel(clientNumber);
    if (logPanel) logPanel.textContent = '';

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
    appRegistrations = await api(url, {}, forceRefreshFromGraph ? 120000 : 45000);
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

document.getElementById('btnInvestigatePacks')?.addEventListener('click', () => {
  window.location.href = 'investigate.html';
});

async function runUpdateGraphAppScopes(preferredTenantId) {
  const tid = String(preferredTenantId || '').trim();
  const target = tid ? appRegTenantLabel(tid) : 'a tenant you choose during sign-in';
  const msg = `Add missing Graph application permissions to the existing River Run Security Investigator app for ${target}? The client secret and WCM credentials stay the same. Sign in as a tenant admin in the console, then run Graph Auth again on that tenant.`;
  if (!window.confirm(msg)) return;
  await runAction('Update Graph App scopes (browser sign-in on this PC)…', async () => {
    const body = tid ? { tenantId: tid } : {};
    let data;
    try {
      data = await waitForGraphAppJob(await api('/api/wcm/update-graph-app-scopes', { method: 'POST', body: JSON.stringify(body) }, 30000));
    } catch (e) {
      if (String(e.message || e).includes('Not found')) {
        throw new Error('Update Graph App scopes needs a web-runner restart (new /api/wcm/update-graph-app-scopes route).');
      }
      throw e;
    }
    const r = data.result || {};
    if (data.exitCode === 0 && r.UpdatedExisting) {
      log(`Graph app scopes updated for ${r.TenantDisplayName || r.TenantId || tid || 'tenant'} (granted ${r.RolesGranted || 0}, already ${r.RolesAlready || 0}, failed ${r.RolesFailed || 0}). Run Graph Auth again.`);
    } else {
      log(`Update Graph App scopes finished (exit ${data.exitCode}). See ${data.logPath || 'temp log'} if needed.`);
    }
  });
}

document.getElementById('btnCreateGraphApp')?.addEventListener('click', () => runAction('Create Graph App (browser sign-in on this PC)…', async () => {
  const data = await waitForGraphAppJob(await api('/api/wcm/create-graph-app', { method: 'POST', body: '{}' }, 30000));
  if (data.result?.WcmSaved) {
    log(`Graph app created for ${data.result.TenantDisplayName || data.result.TenantId}. Select it in App reg tenant, then Graph Auth.`);
  } else {
    log(`Create Graph App finished (exit ${data.exitCode}). See ${data.logPath || 'temp log'} if needed.`);
  }
  await loadAppRegistrations({ quiet: true, forceRefreshFromGraph: true });
}));

function appRegTenantLabel(tenantId) {
  if (!tenantId) return '(none selected)';
  const match = appRegistrations.find(a => (a.tenantId || '').toLowerCase() === String(tenantId).toLowerCase());
  return match?.displayText || tenantId;
}

document.getElementById('btnUpdateGraphAppScopes')?.addEventListener('click', async () => {
  // Prefill from the tenant the user is actually looking at. Defaulting to the first card
  // in the page silently offers an unrelated client's tenant, and the raw GUID in the prompt
  // gives no hint that it is the wrong one.
  const selected = Array.from(document.querySelectorAll('details.tenant'))
    .map(card => ({
      open: card.open,
      client: card.dataset.client || '?',
      tenantId: card.querySelector('.appRegSelect')?.value || '',
    }))
    .filter(entry => entry.tenantId);

  const openEntries = selected.filter(entry => entry.open);
  const candidates = openEntries.length ? openEntries : selected;
  const distinct = Array.from(new Set(candidates.map(entry => entry.tenantId)));
  const prefill = distinct.length === 1 ? distinct[0] : '';

  const listing = selected.length
    ? `\n\nApp reg tenant per client:\n${selected.map(e => `  Client ${e.client}${e.open ? ' (open)' : ''}: ${appRegTenantLabel(e.tenantId)}`).join('\n')}`
    : '';
  const chosen = prefill
    ? `\n\nDefault below: ${appRegTenantLabel(prefill)}`
    : '\n\nNo single tenant is in focus, so no default is filled in.';

  const typed = window.prompt(`Tenant ID to update (blank = choose during sign-in).${listing}${chosen}`, prefill);
  if (typed === null) return;
  const tid = typed.trim();
  if (tid) log(`Update Graph App scopes target: ${appRegTenantLabel(tid)}`);
  await runUpdateGraphAppScopes(tid);
});

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

