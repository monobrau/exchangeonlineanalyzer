/**
 * Microsoft sign-in (MSAL) + delegated Graph: tenant context + app registration CRUD.
 * MSAL is loaded dynamically so a CDN failure does not prevent ms-graph.js from loading.
 *
 * Client ID resolution: server (EOA_MS_GRAPH_SPA_CLIENT_ID or bundled in app) OR localStorage
 * key eoa_ms_graph_spa_client_id — no server env required if the user pastes a GUID once.
 */
let pca = null;

const LOCAL_STORAGE_MS_CLIENT_ID = "eoa_ms_graph_spa_client_id";

/** Must match app/ms_graph_spa.py DELEGATED_GRAPH_SCOPES */
const DELEGATED_GRAPH_SCOPES = [
  "User.Read",
  "Organization.Read.All",
  "Application.ReadWrite.All",
];

async function importMsalBrowser() {
  const urls = [
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.26.1/+esm",
    "https://esm.sh/@azure/msal-browser@3.26.1",
  ];
  let last;
  for (const u of urls) {
    try {
      return await import(/* webpackIgnore: true */ u);
    } catch (e) {
      last = e;
    }
  }
  throw last || new Error("MSAL import failed");
}

function redirectUriForPage() {
  const origin = window.location.origin;
  const p = window.location.pathname || "/";
  if (p === "/app" || p.startsWith("/app/")) return `${origin}/app`;
  return `${origin}/`;
}

async function fetchMsalConfig() {
  const r = await fetch("/api/v1/auth/msal-config", { credentials: "same-origin" });
  if (!r.ok) return { enabled: false, scopes: DELEGATED_GRAPH_SCOPES };
  const cfg = await r.json();
  try {
    const stored =
      typeof localStorage !== "undefined" ? localStorage.getItem(LOCAL_STORAGE_MS_CLIENT_ID) : null;
    if (stored && /^[0-9a-f-]{36}$/i.test(stored.trim())) {
      const cid = stored.trim();
      if (!cfg.enabled || !(cfg.clientId || "").toString().trim()) {
        return {
          enabled: true,
          clientId: cid,
          authority: "https://login.microsoftonline.com/organizations",
          scopes: Array.isArray(cfg.scopes) && cfg.scopes.length ? cfg.scopes : DELEGATED_GRAPH_SCOPES,
          redirectPath: "/",
          clientIdSource: "localStorage",
        };
      }
    }
  } catch {
    /* private mode / no localStorage */
  }
  return cfg;
}

async function graphJson(path, token, opts = {}) {
  const url = path.startsWith("http") ? path : `https://graph.microsoft.com/v1.0${path}`;
  const headers = {
    ...(opts.headers || {}),
    Authorization: `Bearer ${token}`,
  };
  if (opts.body) headers["Content-Type"] = "application/json";
  const r = await fetch(url, {
    ...opts,
    headers,
  });
  const text = await r.text();
  if (!r.ok) throw new Error(text.slice(0, 800) || r.statusText);
  if (r.status === 204 || !text) return null;
  return JSON.parse(text);
}

async function graphListAll(firstPath, token) {
  let url = firstPath.startsWith("http") ? firstPath : `https://graph.microsoft.com/v1.0${firstPath}`;
  const items = [];
  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    const text = await r.text();
    if (!r.ok) throw new Error(text.slice(0, 800) || r.statusText);
    const data = JSON.parse(text);
    const chunk = data.value;
    if (Array.isArray(chunk)) items.push(...chunk);
    url = data["@odata.nextLink"] || null;
  }
  return items;
}

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s;
  return d.innerHTML;
}

/** Turn Graph/HTTP error body into a short user-facing string. */
function formatGraphError(err) {
  const raw = String(err && err.message ? err.message : err);
  try {
    const j = JSON.parse(raw);
    const msg = j.error && j.error.message;
    const code = j.error && j.error.code;
    if (msg) return code ? `${code}: ${msg}` : msg;
  } catch {
    /* not JSON */
  }
  return raw.slice(0, 1200);
}

export async function initMicrosoftGraphUI() {
  const mount = document.getElementById("ms-graph-mount");
  if (!mount) return;

  const cfg = await fetchMsalConfig();
  if (!cfg.enabled) {
    mount.innerHTML = `
      <p class="hint">Paste your Entra <strong>Application (client) ID</strong> for a <strong>single-page application</strong> registration (saved in this browser only). Then use <strong>Sign in with Microsoft</strong>. Or set <code>EOA_MS_GRAPH_SPA_CLIENT_ID</code> on the server.</p>
      <div class="ms-row">
        <input type="text" id="ms-client-id-input" class="input-grow" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" spellcheck="false" autocomplete="off" />
        <button type="button" class="primary" id="ms-save-client-id">Save & continue</button>
      </div>
      <p id="ms-client-id-msg" class="msg" role="status"></p>
      <p class="hint">Redirect URI in Entra must include <code>${escapeHtml(redirectUriForPage())}</code> · Delegated: User.Read, Organization.Read.All, Application.ReadWrite.All (admin consent).</p>
    `;
    const inp = document.getElementById("ms-client-id-input");
    const btn = document.getElementById("ms-save-client-id");
    const msg = document.getElementById("ms-client-id-msg");
    btn.addEventListener("click", () => {
      const v = (inp.value || "").trim();
      if (!/^[0-9a-f-]{36}$/i.test(v)) {
        msg.textContent = "Enter a valid GUID (client ID).";
        return;
      }
      try {
        localStorage.setItem(LOCAL_STORAGE_MS_CLIENT_ID, v);
      } catch (e) {
        msg.textContent = String(e.message || e);
        return;
      }
      location.reload();
    });
    return;
  }

  let PublicClientApplication;
  let InteractionRequiredAuthError;
  try {
    const msal = await importMsalBrowser();
    PublicClientApplication = msal.PublicClientApplication;
    InteractionRequiredAuthError = msal.InteractionRequiredAuthError;
  } catch (e) {
    mount.innerHTML =
      `<p class="msg">Could not load the MSAL script (CDN blocked or offline).</p>` +
      `<p class="hint">${escapeHtml(String(e.message || e))}</p>` +
      `<p class="hint">Try another network, allow <code>cdn.jsdelivr.net</code> / <code>esm.sh</code>, or host MSAL locally.</p>`;
    return;
  }

  async function acquireToken(scopes) {
    const acc = pca.getActiveAccount() || pca.getAllAccounts()[0];
    if (!acc) throw new Error("Sign in with Microsoft first.");
    const req = { scopes, account: acc };
    try {
      const r = await pca.acquireTokenSilent(req);
      return r.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        const r = await pca.acquireTokenPopup({ ...req, prompt: "consent" });
        return r.accessToken;
      }
      throw e;
    }
  }

  pca = new PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: cfg.authority,
      redirectUri: redirectUriForPage(),
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: true },
  });
  await pca.initialize();
  await pca.handleRedirectPromise();

  const loginRequest = { scopes: cfg.scopes || DELEGATED_GRAPH_SCOPES };

  mount.innerHTML = `
    <div class="ms-row">
      <button type="button" class="btn-ms" id="ms-signin">Sign in with Microsoft</button>
      <button type="button" class="ghost" id="ms-reconsent" hidden title="Use after admin grants Application.ReadWrite.All">Re-consent permissions</button>
      <button type="button" class="ghost" id="ms-signout" hidden>Sign out (Microsoft)</button>
    </div>
    <p id="ms-tenant-msg" class="msg" role="status"></p>
    <div id="ms-tenant-box" class="ms-tenant-box" hidden>
      <p><strong>Tenant</strong> <span id="ms-tenant-name">—</span></p>
      <p class="mono small"><span id="ms-tenant-id">—</span></p>
      <button type="button" class="primary" id="ms-use-tenant">Use this tenant for bulk job</button>
    </div>
    <p class="hint small" id="ms-clear-client-id-wrap" hidden>
      <button type="button" class="linklike" id="ms-clear-saved-client-id">Clear browser-stored client ID</button>
    </p>
    <hr class="sep" />
    <h3 class="h3">App registrations (Graph)</h3>
    <p class="hint">Creates and lists apps via <strong>delegated</strong> Graph from your browser (interactive login). Requires API permission <code>Application.ReadWrite.All</code> (Delegated) on the EOA SPA app — usually <strong>admin consent</strong> in Entra. If create/list fails with forbidden or consent errors, ask a Global Administrator to grant consent, then use <strong>Re-consent permissions</strong> and try again.</p>
    <div class="ms-row">
      <input type="text" id="ms-new-app-name" class="input-grow" placeholder="New app display name" maxlength="200" />
      <button type="button" class="primary" id="ms-create-app">Create</button>
      <button type="button" class="ghost" id="ms-refresh-apps">Refresh list</button>
    </div>
    <p id="ms-apps-msg" class="msg" role="status"></p>
    <div class="table-wrap">
      <table class="ms-apps-table">
        <thead>
          <tr><th>Display name</th><th>Application (client) ID</th><th>Object ID</th><th></th></tr>
        </thead>
        <tbody id="ms-apps-body"><tr><td colspan="4" class="empty">Sign in to load applications.</td></tr></tbody>
      </table>
    </div>
  `;

  const el = (id) => document.getElementById(id);
  const tenantMsg = el("ms-tenant-msg");
  const appsMsg = el("ms-apps-msg");

  if (cfg.clientIdSource === "localStorage") {
    const w = el("ms-clear-client-id-wrap");
    if (w) w.hidden = false;
    el("ms-clear-saved-client-id")?.addEventListener("click", () => {
      try {
        localStorage.removeItem(LOCAL_STORAGE_MS_CLIENT_ID);
      } catch {
        /* ignore */
      }
      location.reload();
    });
  }

  async function refreshTenantUi() {
    const acc = pca.getActiveAccount() || pca.getAllAccounts()[0];
    el("ms-signout").hidden = !acc;
    el("ms-reconsent").hidden = !acc;
    if (!acc) {
      el("ms-tenant-box").hidden = true;
      return;
    }
    pca.setActiveAccount(acc);
    const tid = acc.tenantId;
    el("ms-tenant-id").textContent = tid;
    tenantMsg.textContent = "";
    try {
      const token = await acquireToken(loginRequest.scopes);
      const org = await graphJson("/organization", token);
      const v = org && org.value && org.value[0];
      el("ms-tenant-name").textContent = (v && v.displayName) || "(organization)";
    } catch (e) {
      tenantMsg.textContent = formatGraphError(e);
      el("ms-tenant-name").textContent = "—";
    }
    el("ms-tenant-box").hidden = false;
  }

  async function refreshAppsTable() {
    const tbody = el("ms-apps-body");
    const acc = pca.getActiveAccount() || pca.getAllAccounts()[0];
    if (!acc) {
      tbody.innerHTML =
        '<tr><td colspan="4" class="empty">Sign in with Microsoft to manage app registrations.</td></tr>';
      return;
    }
    appsMsg.textContent = "Loading…";
    try {
      const token = await acquireToken(loginRequest.scopes);
      const apps = await graphListAll(
        "/applications?$select=id,appId,displayName,signInAudience&$top=200",
        token
      );
      appsMsg.textContent = apps.length ? `${apps.length} application(s).` : "No applications returned.";
      if (apps.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" class="empty">No applications (or insufficient permission).</td></tr>';
        return;
      }
      tbody.innerHTML = apps
        .map((a) => {
          const oid = escapeHtml(a.id || "");
          const dn = escapeHtml(a.displayName || "");
          const cid = escapeHtml(a.appId || "");
          return `<tr>
            <td>${dn}</td>
            <td class="mono">${cid}</td>
            <td class="mono">${oid.slice(0, 8)}…</td>
            <td class="ms-actions">
              <button type="button" class="linklike" data-act="rename" data-id="${oid}">Rename</button>
              <button type="button" class="linklike danger" data-act="delete" data-id="${oid}">Delete</button>
            </td>
          </tr>`;
        })
        .join("");
      tbody.querySelectorAll("button[data-act]").forEach((btn) => {
        btn.addEventListener("click", async () => {
          const id = btn.getAttribute("data-id");
          const act = btn.getAttribute("data-act");
          if (act === "delete") {
            if (!confirm("Delete this app registration? This cannot be undone.")) return;
            appsMsg.textContent = "Deleting…";
            try {
              const token = await acquireToken(loginRequest.scopes);
              const r = await fetch(`https://graph.microsoft.com/v1.0/applications/${id}`, {
                method: "DELETE",
                headers: { Authorization: `Bearer ${token}` },
              });
              if (!r.ok) throw new Error((await r.text()).slice(0, 400));
              appsMsg.textContent = "Deleted.";
              await refreshAppsTable();
            } catch (e) {
              appsMsg.textContent = String(e.message || e);
            }
            return;
          }
          if (act === "rename") {
            const name = window.prompt("New display name:");
            if (name == null || !String(name).trim()) return;
            appsMsg.textContent = "Saving…";
            try {
              const token = await acquireToken(loginRequest.scopes);
              await graphJson(
                `/applications/${id}`,
                token,
                {
                  method: "PATCH",
                  body: JSON.stringify({ displayName: String(name).trim() }),
                }
              );
              appsMsg.textContent = "Updated.";
              await refreshAppsTable();
            } catch (e) {
              appsMsg.textContent = formatGraphError(e);
            }
          }
        });
      });
    } catch (e) {
      appsMsg.textContent = formatGraphError(e);
      tbody.innerHTML = '<tr><td colspan="4" class="empty">Could not load applications.</td></tr>';
    }
  }

  el("ms-signin").addEventListener("click", async () => {
    tenantMsg.textContent = "";
    try {
      await pca.loginPopup(loginRequest);
      await refreshTenantUi();
      await refreshAppsTable();
    } catch (e) {
      tenantMsg.textContent = formatGraphError(e);
    }
  });

  el("ms-reconsent").addEventListener("click", async () => {
    tenantMsg.textContent = "";
    try {
      await pca.loginPopup({ ...loginRequest, prompt: "consent" });
      await refreshTenantUi();
      await refreshAppsTable();
      tenantMsg.textContent = "Permissions refreshed.";
    } catch (e) {
      tenantMsg.textContent = formatGraphError(e);
    }
  });

  el("ms-signout").addEventListener("click", () => {
    const acc = pca.getActiveAccount();
    if (acc) pca.logoutPopup({ account: acc });
    el("ms-tenant-box").hidden = true;
    el("ms-apps-body").innerHTML =
      '<tr><td colspan="4" class="empty">Sign in with Microsoft to manage app registrations.</td></tr>';
    appsMsg.textContent = "";
    tenantMsg.textContent = "Signed out.";
  });

  el("ms-use-tenant").addEventListener("click", () => {
    const tid = el("ms-tenant-id").textContent.trim();
    const ta = document.getElementById("tenant-ids");
    if (ta && tid) {
      ta.value = tid;
      tenantMsg.textContent = "Tenant ID copied into the bulk job field.";
    }
  });

  el("ms-create-app").addEventListener("click", async () => {
    const name = (el("ms-new-app-name").value || "").trim();
    if (!name) {
      appsMsg.textContent = "Enter a display name.";
      return;
    }
    appsMsg.textContent = "Creating…";
    try {
      const token = await acquireToken(loginRequest.scopes);
      await graphJson("/applications", token, {
        method: "POST",
        body: JSON.stringify({
          displayName: name,
          signInAudience: "AzureADMyOrg",
        }),
      });
      el("ms-new-app-name").value = "";
      appsMsg.textContent = "Created.";
      await refreshAppsTable();
    } catch (e) {
      appsMsg.textContent = formatGraphError(e);
    }
  });

  el("ms-refresh-apps").addEventListener("click", () => refreshAppsTable());

  const existing = pca.getAllAccounts();
  if (existing.length) {
    pca.setActiveAccount(existing[0]);
    await refreshTenantUi();
    await refreshAppsTable();
  }
}

export function getMsGraphTenantIdForJob() {
  if (!pca) return null;
  const acc = pca.getActiveAccount() || pca.getAllAccounts()[0];
  return acc ? acc.tenantId : null;
}
