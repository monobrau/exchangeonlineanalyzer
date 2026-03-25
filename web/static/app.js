const $ = (sel) => document.querySelector(sel);

/** Set after dynamic import of ./ms-graph.js (avoids aborting the whole app if MSAL CDN is blocked). */
let getMsGraphTenantIdForJob = () => null;

/** When OIDC is on and there is no bearer token, main UI stays hidden (see initAuth). */
let appUnlocked = true;

/** @type {null | { graph_app_configured: boolean; exo: string; job_default_tenant_id: string | null; use_python_graph_worker?: boolean; use_pwsh_worker?: boolean; exo_organization?: string | null }} */
let connectionsStatus = null;

window.addEventListener("eoa-ms-tenant", () => {
  updateJobTenantHint();
});

function collectJobOptionsFromDom() {
  const daysEl = document.getElementById("job-days-back");
  const signEl = document.getElementById("job-signin-days");
  let days = parseInt(daysEl?.value || "10", 10);
  if (!Number.isFinite(days)) days = 10;
  days = Math.min(365, Math.max(1, days));
  let signDays = parseInt(signEl?.value || "7", 10);
  if (![1, 7, 30].includes(signDays)) signDays = 7;
  const minimal = document.getElementById("job-minimal-graph")?.checked;
  if (minimal) {
    return {
      minimal_graph_test: true,
      reports: ["organization"],
      days_back: days,
      message_trace_days_back: days,
      sign_in_logs_days_back: signDays,
    };
  }
  const opts = {
    days_back: days,
    message_trace_days_back: days,
    sign_in_logs_days_back: signDays,
  };
  document.querySelectorAll("#job-form input[type=checkbox][data-opt]").forEach((cb) => {
    const k = cb.getAttribute("data-opt");
    if (k && cb.checked) opts[k] = true;
  });
  return opts;
}

function anyReportCheckboxChecked() {
  return [...document.querySelectorAll("#job-form input[type=checkbox][data-opt]")].some((cb) => cb.checked);
}

async function api(path, opts = {}) {
  const headers = { ...opts.headers };
  const token = sessionStorage.getItem("eoa_bearer");
  if (token) headers.Authorization = `Bearer ${token}`;
  const r = await fetch(path, { ...opts, headers, credentials: "same-origin" });
  if (r.status === 401) {
    const raw = await r.text();
    let detail = "";
    try {
      const j = JSON.parse(raw);
      if (j && j.detail != null) {
        if (Array.isArray(j.detail)) {
          detail = j.detail
            .map((x) => (x && typeof x === "object" && x.msg != null ? String(x.msg) : JSON.stringify(x)))
            .join("; ");
        } else {
          detail = String(j.detail);
        }
      }
    } catch {
      if (raw && raw.trim()) detail = raw.trim().slice(0, 400);
    }
    throw new Error(
      detail ||
        "Unauthorized (401). Set OIDC or paste a Bearer token in sessionStorage key eoa_bearer)."
    );
  }
  if (!r.ok) {
    const t = await r.text();
    throw new Error(t || r.statusText);
  }
  const ct = r.headers.get("content-type");
  if (ct && ct.includes("application/json")) return r.json();
  return r.text();
}

function badgeClass(status) {
  const s = (status || "").toLowerCase();
  if (s === "queued") return "queued";
  if (s === "running") return "running";
  if (s === "succeeded") return "succeeded";
  if (s === "failed") return "failed";
  return "queued";
}

function fmtTime(iso) {
  if (!iso) return "—";
  try {
    const d = new Date(iso);
    return d.toLocaleString();
  } catch {
    return iso;
  }
}

async function downloadArtifact(jobId, filename) {
  const token = sessionStorage.getItem("eoa_bearer");
  const headers = {};
  if (token) headers.Authorization = `Bearer ${token}`;
  const q = new URLSearchParams({ file: filename });
  const r = await fetch(`/api/v1/jobs/${jobId}/artifact?${q}`, {
    headers,
    credentials: "same-origin",
  });
  if (r.status === 401) {
    throw new Error("Unauthorized (401). Set eoa_bearer in sessionStorage if OIDC is on.");
  }
  if (!r.ok) {
    const t = await r.text();
    throw new Error(t || r.statusText);
  }
  const blob = await r.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${filename}`;
  a.click();
  URL.revokeObjectURL(url);
}

function artifactCell(j) {
  if (j.status !== "succeeded") return "—";
  const files = Array.isArray(j.artifact_files) ? j.artifact_files : [];
  if (files.length === 0) return "—";
  return files
    .map(
      (f) =>
        `<button type="button" class="linklike" data-dl="${j.id}" data-file="${escapeHtml(f)}">${escapeHtml(f)}</button>`
    )
    .join(" ");
}

async function loadJobs() {
  const tbody = $("#jobs-body");
  const msg = $("#list-msg");
  msg.textContent = "";
  try {
    const data = await api("/api/v1/jobs?limit=30");
    if (!data.jobs || data.jobs.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" class="empty">No jobs yet.</td></tr>';
      renderActivityStrip([]);
      return;
    }
    renderActivityStrip(data.jobs);
    tbody.innerHTML = data.jobs
      .map(
        (j) => `
      <tr>
        <td class="mono" title="${j.id}">${j.id.slice(0, 8)}…</td>
        <td class="mono">${escapeHtml(tenantFromJobPayload(j))}</td>
        <td><span class="badge ${badgeClass(j.status)}">${escapeHtml(j.status)}</span></td>
        <td>${fmtTime(j.created_at)}</td>
        <td class="mono">${j.artifact_uri ? escapeHtml(String(j.artifact_uri).slice(0, 36)) + (String(j.artifact_uri).length > 36 ? "…" : "") : "—"}</td>
        <td>${artifactCell(j)}</td>
        <td class="mono err-cell">${j.error_message ? escapeHtml(String(j.error_message).slice(0, 64)) + (String(j.error_message).length > 64 ? "…" : "") : "—"}</td>
        <td>${j.request_payload ? `<button type="button" class="linklike" data-rerun="${j.id}">Run again</button>` : "—"}</td>
      </tr>`
      )
      .join("");
    tbody.querySelectorAll("button[data-dl]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.getAttribute("data-dl");
        const file = btn.getAttribute("data-file") || "summary.json";
        try {
          await downloadArtifact(id, file);
        } catch (e) {
          msg.textContent = String(e.message || e);
        }
      });
    });
    tbody.querySelectorAll("button[data-rerun]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.getAttribute("data-rerun");
        if (id) await runAgainJob(id);
      });
    });
  } catch (e) {
    msg.textContent = String(e.message || e);
    tbody.innerHTML = '<tr><td colspan="8" class="empty">Could not load jobs.</td></tr>';
    renderActivityStrip([]);
  }
}

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s;
  return d.innerHTML;
}

async function loadConnectionsStatus() {
  try {
    connectionsStatus = await api("/api/v1/connections/status");
    renderConnectionsStatus();
    updateJobTenantHint();
  } catch (e) {
    const el = document.getElementById("conn-status-mount");
    if (el) el.innerHTML = `<p class="msg">${escapeHtml(String(e.message || e))}</p>`;
  }
}

function renderConnectionsStatus() {
  const el = document.getElementById("conn-status-mount");
  if (!el || !connectionsStatus) return;
  const g = connectionsStatus.graph_app_configured;
  const exo = connectionsStatus.exo;
  const exoOrg = connectionsStatus.exo_organization;
  const def = connectionsStatus.job_default_tenant_id;
  const exoLabel =
    exo === "ready" ? "Ready" : exo === "skipped" ? "Skipped (Graph-only)" : "Not configured";
  const exoBadge = exo === "ready" ? "succeeded" : exo === "skipped" ? "queued" : "failed";
  el.innerHTML = `
    <div class="conn-grid">
      <div class="conn-card">
        <div class="conn-card-head">
          <strong>Microsoft Graph (server)</strong>
          <span class="badge ${g ? "succeeded" : "failed"}">${g ? "Configured" : "Not configured"}</span>
        </div>
        <p class="hint small">App ID + secret for the Python Graph worker (app-only).</p>
      </div>
      <div class="conn-card">
        <div class="conn-card-head">
          <strong>Exchange Online (PowerShell)</strong>
          <span class="badge ${exoBadge}">${escapeHtml(exoLabel)}</span>
        </div>
        <p class="hint small">${
          exoOrg
            ? `Organization: <span class="mono">${escapeHtml(exoOrg)}</span>`
            : "Certificate + app for the worker, or enable skip in Settings."
        }</p>
      </div>
    </div>
    ${def ? `<p class="hint small">Default job tenant: <span class="mono">${escapeHtml(def)}</span></p>` : ""}
  `;
}

function updateJobTenantHint() {
  const p = document.getElementById("job-tenant-status");
  if (!p) return;
  const def = connectionsStatus && connectionsStatus.job_default_tenant_id;
  if (def) {
    p.textContent = `Default tenant from Settings: ${def.slice(0, 8)}… — leave queue empty for one job for that directory, or add more GUIDs.`;
    p.classList.remove("job-tenant-warn");
  } else {
    p.textContent = "Add at least one tenant GUID to the queue, or set EOA_JOB_DEFAULT_TENANT_ID in Settings.";
    p.classList.add("job-tenant-warn");
  }
}

const TENANT_QUEUE_KEY = "eoa_tenant_queue";
const SK_ACTIVITY_DISMISSED = "eoa_activity_dismissed";
const SK_ACTIVITY_MIN = "eoa_activity_minimized";

function loadStrSet(key) {
  try {
    const raw = sessionStorage.getItem(key);
    if (!raw) return new Set();
    const arr = JSON.parse(raw);
    return new Set(Array.isArray(arr) ? arr : []);
  } catch {
    return new Set();
  }
}

function saveStrSet(key, set) {
  sessionStorage.setItem(key, JSON.stringify([...set]));
}

function getTenantQueue() {
  try {
    const raw = sessionStorage.getItem(TENANT_QUEUE_KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return [];
    return arr.filter((x) => x && x.tenantId && /^[0-9a-f-]{36}$/i.test(String(x.tenantId).trim()));
  } catch {
    return [];
  }
}

function setTenantQueue(q) {
  const clean = q
    .map((x) => ({
      tenantId: String(x.tenantId).trim(),
      label: x.label ? String(x.label) : "",
    }))
    .filter((x) => /^[0-9a-f-]{36}$/i.test(x.tenantId));
  const seen = new Set();
  const uniq = [];
  for (const x of clean) {
    const k = x.tenantId.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    uniq.push(x);
  }
  sessionStorage.setItem(TENANT_QUEUE_KEY, JSON.stringify(uniq));
}

function updateSubmitButtonLabel() {
  const btn = document.getElementById("job-submit-btn");
  if (!btn) return;
  const n = getTenantQueue().length;
  if (n > 1) btn.textContent = `Create ${n} bulk jobs`;
  else btn.textContent = "Create bulk job";
}

function renderTenantQueue() {
  const ul = document.getElementById("tenant-queue-list");
  if (!ul) return;
  const q = getTenantQueue();
  if (q.length === 0) {
    ul.innerHTML =
      '<li class="tenant-queue-empty hint small">Queue empty — your Microsoft sign-in tenant is used on submit, or add GUIDs here.</li>';
    updateSubmitButtonLabel();
    return;
  }
  ul.innerHTML = q
    .map(
      (x, i) =>
        `<li class="tenant-queue-item"><span class="mono">${escapeHtml(x.tenantId)}</span> ` +
        `<button type="button" class="linklike" data-queue-remove="${i}" aria-label="Remove">Remove</button></li>`
    )
    .join("");
  ul.querySelectorAll("[data-queue-remove]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const idx = parseInt(btn.getAttribute("data-queue-remove"), 10);
      const q2 = getTenantQueue();
      if (!Number.isFinite(idx)) return;
      q2.splice(idx, 1);
      setTenantQueue(q2);
      renderTenantQueue();
    });
  });
  updateSubmitButtonLabel();
}

function tenantFromJobPayload(j) {
  const p = j.request_payload;
  if (!p || typeof p !== "object") return "—";
  const t = p.tenant_ids;
  if (Array.isArray(t) && t[0]) return `${String(t[0]).slice(0, 8)}…`;
  return "—";
}

function applyRequestPayloadToForm(payload) {
  if (!payload || typeof payload !== "object") return;
  const opts = payload.options || {};
  const daysEl = document.getElementById("job-days-back");
  const signEl = document.getElementById("job-signin-days");
  const minimalEl = document.getElementById("job-minimal-graph");
  const d = parseInt(opts.days_back, 10);
  if (Number.isFinite(d) && daysEl) daysEl.value = String(Math.min(365, Math.max(1, d)));
  const sd = parseInt(opts.sign_in_logs_days_back, 10);
  if ([1, 7, 30].includes(sd) && signEl) signEl.value = String(sd);
  if (minimalEl) minimalEl.checked = !!opts.minimal_graph_test;
  document.querySelectorAll("#job-form input[type=checkbox][data-opt]").forEach((cb) => {
    const k = cb.getAttribute("data-opt");
    if (k) cb.checked = !!opts[k];
  });
  const tids = Array.isArray(payload.tenant_ids) ? payload.tenant_ids : [];
  if (tids.length) {
    setTenantQueue(
      tids.map((id) => ({
        tenantId: String(id).trim(),
        label: "",
      }))
    );
  }
  renderTenantQueue();
}

async function runAgainJob(jobId) {
  const formMsg = $("#form-msg");
  if (formMsg) formMsg.textContent = "";
  try {
    const j = await api(`/api/v1/jobs/${jobId}`);
    if (!j.request_payload) {
      if (formMsg) formMsg.textContent = "This job has no stored request (cannot pre-fill).";
      return;
    }
    applyRequestPayloadToForm(j.request_payload);
    if (formMsg) formMsg.textContent = "Form updated from job — adjust options and submit.";
    document.getElementById("new-job-title")?.scrollIntoView({ behavior: "smooth" });
  } catch (e) {
    if (formMsg) formMsg.textContent = String(e.message || e);
  }
}

function renderActivityStrip(jobs) {
  const mount = document.getElementById("job-activity-strip");
  if (!mount) return;
  const dismissed = loadStrSet(SK_ACTIVITY_DISMISSED);
  const minimized = loadStrSet(SK_ACTIVITY_MIN);
  const slice = (jobs || []).filter((j) => !dismissed.has(j.id)).slice(0, 12);
  if (slice.length === 0) {
    mount.innerHTML = "";
    return;
  }
  mount.innerHTML = slice
    .map((j) => {
      const isMin = minimized.has(j.id);
      const files = Array.isArray(j.artifact_files) ? j.artifact_files : [];
      const dl =
        j.status === "succeeded" && files.length
          ? files
              .map(
                (f) =>
                  `<button type="button" class="linklike" data-act-dl="${escapeHtml(j.id)}" data-file="${escapeHtml(f)}">${escapeHtml(f)}</button>`
              )
              .join(" ")
          : "—";
      const body = isMin
        ? ""
        : `<div class="job-activity-body">
          <span class="mono">${escapeHtml(tenantFromJobPayload(j))}</span>
          <span class="job-activity-time">${fmtTime(j.created_at)}</span>
          <div class="job-activity-dl">${dl}</div>
        </div>`;
      return `<div class="job-activity-card${isMin ? " job-activity-card--min" : ""}" data-job-id="${escapeHtml(j.id)}">
        <div class="job-activity-card-head">
          <span class="mono">${j.id.slice(0, 8)}…</span>
          <span class="badge ${badgeClass(j.status)}">${escapeHtml(j.status)}</span>
          <span class="job-activity-spacer"></span>
          <button type="button" class="ghost sm" data-act="min" data-id="${escapeHtml(j.id)}">${isMin ? "Expand" : "Minimize"}</button>
          <button type="button" class="ghost sm" data-act="rerun" data-id="${escapeHtml(j.id)}">Run again</button>
          <button type="button" class="ghost sm" data-act="dismiss" data-id="${escapeHtml(j.id)}">Dismiss</button>
        </div>
        ${body}
      </div>`;
    })
    .join("");

  const listMsg = $("#list-msg");
  mount.querySelectorAll("button[data-act]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const id = btn.getAttribute("data-id");
      const act = btn.getAttribute("data-act");
      if (!id) return;
      if (act === "dismiss") {
        const s = loadStrSet(SK_ACTIVITY_DISMISSED);
        s.add(id);
        saveStrSet(SK_ACTIVITY_DISMISSED, s);
        await loadJobs();
        return;
      }
      if (act === "min") {
        const s = loadStrSet(SK_ACTIVITY_MIN);
        if (s.has(id)) s.delete(id);
        else s.add(id);
        saveStrSet(SK_ACTIVITY_MIN, s);
        await loadJobs();
        return;
      }
      if (act === "rerun") await runAgainJob(id);
    });
  });
  mount.querySelectorAll("button[data-act-dl]").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const id = btn.getAttribute("data-act-dl");
      const file = btn.getAttribute("data-file") || "summary.json";
      try {
        await downloadArtifact(id, file);
      } catch (e) {
        if (listMsg) listMsg.textContent = String(e.message || e);
      }
    });
  });
}

const SETTINGS_SECTION_TITLES = {
  general: "General & CORS",
  oidc: "API sign-in (OIDC / Authentik)",
  workers: "Workers (PowerShell & Graph)",
  graph_server: "Microsoft Graph — server (app credentials)",
  exo_pwsh: "Exchange Online — PowerShell worker (Linux)",
  microsoft_spa: "Microsoft 365 — browser sign-in (MSAL)",
};

/** @type {Record<string, { kind: string, value?: unknown, has_value?: boolean }> | null} */
let settingsInitial = null;

function closeSettingsModal() {
  const m = document.getElementById("settings-modal");
  if (m) m.hidden = true;
  settingsInitial = null;
}

function renderSettingsForm(data) {
  const form = document.getElementById("settings-form");
  const note = document.getElementById("settings-note");
  if (!form || !note) return;
  note.textContent = data.note || "";
  const items = Array.isArray(data.items) ? data.items : [];
  settingsInitial = {};
  for (const it of items) {
    settingsInitial[it.env_key] =
      it.kind === "secret"
        ? { kind: "secret", has_value: !!it.has_value }
        : { kind: it.kind, value: it.value };
  }
  const bySection = {};
  for (const it of items) {
    const sec = it.section || "general";
    if (!bySection[sec]) bySection[sec] = [];
    bySection[sec].push(it);
  }
  const order = ["general", "oidc", "workers", "graph_server", "exo_pwsh", "microsoft_spa"];
  const parts = [];
  for (const sec of order) {
    const group = bySection[sec];
    if (!group || !group.length) continue;
    const title = SETTINGS_SECTION_TITLES[sec] || sec;
    parts.push(`<div class="settings-section"><h3>${escapeHtml(title)}</h3>`);
    for (const it of group) {
      const rk = it.restart_required ? ' <span class="warn-tag">restart</span>' : "";
      const desc = it.description ? `<p class="field-hint">${escapeHtml(it.description)}</p>` : "";
      if (it.kind === "bool") {
        parts.push(
          `<div class="settings-field settings-row-inline" data-env-key="${escapeHtml(it.env_key)}" data-kind="bool">` +
            `<label><input type="checkbox" data-bool ${it.value ? "checked" : ""} />` +
            `<strong>${escapeHtml(it.label)}</strong>${rk}</label>${desc}</div>`
        );
      } else if (it.kind === "secret") {
        const hv = it.has_value ? "<p class=\"field-hint\">A value is stored (hidden).</p>" : "";
        parts.push(
          `<div class="settings-field" data-env-key="${escapeHtml(it.env_key)}" data-kind="secret">` +
            `<label>${escapeHtml(it.label)}${rk}</label>${hv}${desc}` +
            `<input type="password" autocomplete="new-password" data-secret-input placeholder="New value (optional)" />` +
            `<label class="settings-secret-clear settings-row-inline">` +
            `<input type="checkbox" data-secret-clear /> Remove override (fall back to web/.env only)</label></div>`
        );
      } else if (it.kind === "int") {
        parts.push(
          `<div class="settings-field" data-env-key="${escapeHtml(it.env_key)}" data-kind="int">` +
            `<label>${escapeHtml(it.label)}${rk}</label>${desc}` +
            `<input type="number" data-int value="${escapeHtml(String(it.value ?? ""))}" /></div>`
        );
      } else {
        parts.push(
          `<div class="settings-field" data-env-key="${escapeHtml(it.env_key)}" data-kind="${escapeHtml(it.kind)}">` +
            `<label>${escapeHtml(it.label)}${rk}</label>${desc}` +
            `<input type="text" data-text value="${escapeHtml(String(it.value ?? ""))}" spellcheck="false" /></div>`
        );
      }
    }
    parts.push("</div>");
  }
  form.innerHTML = parts.join("");
}

function collectSettingsPatch() {
  if (!settingsInitial) return {};
  const patch = {};
  document.querySelectorAll("#settings-form .settings-field").forEach((wrap) => {
    const key = wrap.getAttribute("data-env-key");
    const kind = wrap.getAttribute("data-kind");
    const ini = settingsInitial[key];
    if (!key || !ini) return;
    if (kind === "bool") {
      const cb = wrap.querySelector("[data-bool]");
      const now = !!(cb && cb.checked);
      if (ini.value !== now) patch[key] = now;
      return;
    }
    if (kind === "secret") {
      const clr = wrap.querySelector("[data-secret-clear]");
      const inp = wrap.querySelector("[data-secret-input]");
      if (clr && clr.checked) {
        patch[key] = "";
        return;
      }
      const v = (inp && inp.value) ? String(inp.value).trim() : "";
      if (v) patch[key] = v;
      return;
    }
    if (kind === "int") {
      const inp = wrap.querySelector("[data-int]");
      const n = parseInt(inp && inp.value, 10);
      if (!Number.isFinite(n)) return;
      if (ini.value !== n) patch[key] = n;
      return;
    }
    const inp = wrap.querySelector("[data-text]");
    const v = inp ? String(inp.value) : "";
    if (v !== String(ini.value ?? "")) patch[key] = v;
  });
  return patch;
}

async function openSettingsModal() {
  const modal = document.getElementById("settings-modal");
  const msg = document.getElementById("settings-msg");
  if (!modal) return;
  modal.hidden = false;
  if (msg) msg.textContent = "Loading…";
  try {
    const data = await api("/api/v1/settings/runtime-env");
    renderSettingsForm(data);
    if (msg) msg.textContent = "";
  } catch (e) {
    if (msg) msg.textContent = String(e.message || e);
    renderSettingsForm({ items: [], note: "" });
  }
}

const jobForm = $("#job-form");
if (jobForm) {
  jobForm.addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const formMsg = $("#form-msg");
    formMsg.textContent = "";
    const minimal = document.getElementById("job-minimal-graph")?.checked;
    if (!minimal && !anyReportCheckboxChecked()) {
      formMsg.textContent = "Select at least one report, or enable minimal server job.";
      return;
    }
    const options = collectJobOptionsFromDom();
    const q = getTenantQueue();
    let tenantIds = q.map((x) => x.tenantId);
    if (tenantIds.length === 0) {
      let tid = "";
      try {
        tid = (typeof getMsGraphTenantIdForJob === "function" && getMsGraphTenantIdForJob()) || "";
      } catch {
        tid = "";
      }
      if (tid && /^[0-9a-f-]{36}$/i.test(tid)) tenantIds = [tid];
    }
    if (tenantIds.length === 0 && connectionsStatus?.job_default_tenant_id) {
      tenantIds = [connectionsStatus.job_default_tenant_id];
    }
    if (tenantIds.length === 0) {
      formMsg.textContent =
        "Sign in with Microsoft above, add tenant GUID(s) to the queue, or set EOA_JOB_DEFAULT_TENANT_ID in Settings.";
      return;
    }
    try {
      for (let i = 0; i < tenantIds.length; i++) {
        await api("/api/v1/jobs/bulk", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ tenant_ids: [tenantIds[i]], options }),
        });
        if (i < tenantIds.length - 1) await new Promise((r) => setTimeout(r, 80));
      }
      formMsg.textContent =
        tenantIds.length > 1 ? `Created ${tenantIds.length} jobs (one per directory).` : "Job created.";
      await loadJobs();
    } catch (e) {
      formMsg.textContent = String(e.message || e);
    }
  });
}

document.getElementById("btn-queue-add")?.addEventListener("click", () => {
  const formMsg = $("#form-msg");
  if (formMsg) formMsg.textContent = "";
  let tid = "";
  try {
    tid = (typeof getMsGraphTenantIdForJob === "function" && getMsGraphTenantIdForJob()) || "";
  } catch {
    tid = "";
  }
  if (!tid || !/^[0-9a-f-]{36}$/i.test(tid)) {
    if (formMsg) formMsg.textContent = "Sign in with Microsoft in the section above first.";
    return;
  }
  const q = getTenantQueue();
  if (q.some((x) => x.tenantId.toLowerCase() === tid.toLowerCase())) return;
  q.push({ tenantId: tid, label: "Signed-in" });
  setTenantQueue(q);
  renderTenantQueue();
  updateJobTenantHint();
});

document.getElementById("btn-queue-add-manual")?.addEventListener("click", () => {
  const formMsg = $("#form-msg");
  if (formMsg) formMsg.textContent = "";
  const inp = document.getElementById("queue-tenant-input");
  const raw = (inp && inp.value ? String(inp.value) : "").trim();
  if (!/^[0-9a-f-]{36}$/i.test(raw)) {
    if (formMsg) formMsg.textContent = "Enter a valid tenant GUID to add to the queue.";
    return;
  }
  const q = getTenantQueue();
  if (q.some((x) => x.tenantId.toLowerCase() === raw.toLowerCase())) {
    if (inp) inp.value = "";
    return;
  }
  q.push({ tenantId: raw, label: "" });
  setTenantQueue(q);
  if (inp) inp.value = "";
  renderTenantQueue();
});

document.getElementById("btn-queue-clear")?.addEventListener("click", () => {
  setTenantQueue([]);
  renderTenantQueue();
});

$("#refresh").addEventListener("click", () => loadJobs());

document.getElementById("btn-settings")?.addEventListener("click", () => openSettingsModal());
document.getElementById("settings-close")?.addEventListener("click", () => closeSettingsModal());
document.getElementById("settings-modal")?.addEventListener("click", (ev) => {
  if (ev.target && ev.target.id === "settings-modal") closeSettingsModal();
});
document.getElementById("settings-save")?.addEventListener("click", async () => {
  const msg = document.getElementById("settings-msg");
  if (msg) msg.textContent = "";
  const patch = collectSettingsPatch();
  if (!Object.keys(patch).length) {
    if (msg) msg.textContent = "No changes to save.";
    return;
  }
  try {
    const out = await api("/api/v1/settings/runtime-env", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ patch }),
    });
    const rr = out && out.restart_recommended;
    if (msg) {
      msg.textContent = rr
        ? "Saved. Restart the API process for CORS / session secret changes to apply."
        : "Saved. New worker flags apply on the next job.";
    }
    const data = await api("/api/v1/settings/runtime-env");
    renderSettingsForm(data);
    await loadConnectionsStatus();
  } catch (e) {
    if (msg) msg.textContent = String(e.message || e);
  }
});

function unlockAppUi() {
  appUnlocked = true;
  document.body.classList.remove("auth-locked");
  const gate = $("#auth-gate");
  const main = $("#app-main");
  if (gate) gate.hidden = true;
  if (main) main.hidden = false;
}

function lockAppUi() {
  appUnlocked = false;
  document.body.classList.add("auth-locked");
  const gate = $("#auth-gate");
  const main = $("#app-main");
  if (gate) gate.hidden = false;
  if (main) main.hidden = true;
}

function setJobsPlaceholderWhenLocked() {
  const tbody = $("#jobs-body");
  const msg = $("#list-msg");
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="8" class="empty">Sign in to load jobs.</td></tr>';
  if (msg) msg.textContent = "";
  renderActivityStrip([]);
}

async function initAuth() {
  const login = $("#auth-login");
  const out = $("#auth-logout");
  const gateBtn = $("#auth-gate-btn");
  const loginHref = "/api/v1/auth/oidc/login";

  let oidcEnabled = false;
  try {
    const sr = await fetch("/api/v1/auth/status", { credentials: "same-origin" });
    if (sr.ok) {
      const st = await sr.json();
      oidcEnabled = !!st.oidc_login_enabled;
    }
  } catch {
    /* ignore */
  }

  let me = null;
  let meHttp = 0;
  try {
    const mr = await fetch("/api/v1/me", { credentials: "same-origin" });
    meHttp = mr.status;
    if (mr.ok) me = await mr.json();
  } catch {
    me = null;
  }

  function wireOidcLoginLinks() {
    login.style.display = "";
    login.setAttribute("href", loginHref);
    if (gateBtn) gateBtn.setAttribute("href", loginHref);
  }

  if (me && me.auth === "disabled") {
    unlockAppUi();
  } else if (me && me.auth === "oidc" && me.sub) {
    wireOidcLoginLinks();
    out.style.display = "inline-block";
    unlockAppUi();
  } else if (oidcEnabled || meHttp === 401) {
    wireOidcLoginLinks();
    out.style.display = "none";
    lockAppUi();
  } else {
    unlockAppUi();
  }

  out.addEventListener("click", async (e) => {
    e.preventDefault();
    try {
      await fetch("/api/v1/auth/logout", { method: "POST", credentials: "same-origin" });
    } catch {
      /* ignore */
    }
    sessionStorage.removeItem("eoa_bearer");
    out.style.display = "none";
    location.href = "/";
  });
}

async function loadMsGraphModule() {
  if (window.__eoaMsGraphLoaded) return;
  window.__eoaMsGraphLoaded = true;
  const msMount = document.getElementById("ms-graph-mount");
  try {
    const msUrl = new URL("./ms-graph.js", import.meta.url).href;
    const m = await import(msUrl);
    getMsGraphTenantIdForJob = m.getMsGraphTenantIdForJob;
    await m.initMicrosoftGraphUI();
    updateJobTenantHint();
  } catch (e) {
    console.error("EOA: Microsoft 365 panel failed to load", e);
    window.__eoaMsGraphLoaded = false;
    if (msMount) {
      msMount.innerHTML = `<p class="msg">Microsoft 365 controls could not load: ${escapeHtml(
        String(e.message || e)
      )}</p><p class="hint">Often: network blocked the MSAL script (CDN), or <code>ms-graph.js</code> failed to load. Open DevTools (F12) → Console / Network.</p>`;
    }
  }
}

(async () => {
  await initAuth();
  if (!appUnlocked) {
    setJobsPlaceholderWhenLocked();
    return;
  }
  await loadConnectionsStatus();
  await loadMsGraphModule();
  renderTenantQueue();
  updateJobTenantHint();
  await loadJobs();
})();

setInterval(() => {
  if (!appUnlocked) return;
  loadJobs();
}, 8000);
