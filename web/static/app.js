const $ = (sel) => document.querySelector(sel);

/** When OIDC is on and there is no bearer token, main UI stays hidden (see initAuth). */
let appUnlocked = true;

function parseTenantIds(raw) {
  return raw
    .split(/[\n,]+/)
    .map((s) => s.trim())
    .filter(Boolean);
}

async function api(path, opts = {}) {
  const headers = { ...opts.headers };
  const token = sessionStorage.getItem("eoa_bearer");
  if (token) headers.Authorization = `Bearer ${token}`;
  const r = await fetch(path, { ...opts, headers });
  if (r.status === 401) {
    let detail = "";
    try {
      const j = JSON.parse(await r.text());
      if (j && j.detail != null) detail = String(j.detail);
    } catch {
      /* ignore */
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
  const r = await fetch(`/api/v1/jobs/${jobId}/artifact?${q}`, { headers });
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
      tbody.innerHTML = '<tr><td colspan="6" class="empty">No jobs yet.</td></tr>';
      return;
    }
    tbody.innerHTML = data.jobs
      .map(
        (j) => `
      <tr>
        <td class="mono" title="${j.id}">${j.id.slice(0, 8)}…</td>
        <td><span class="badge ${badgeClass(j.status)}">${escapeHtml(j.status)}</span></td>
        <td>${fmtTime(j.created_at)}</td>
        <td class="mono">${j.artifact_uri ? escapeHtml(String(j.artifact_uri).slice(0, 36)) + (String(j.artifact_uri).length > 36 ? "…" : "") : "—"}</td>
        <td>${artifactCell(j)}</td>
        <td class="mono err-cell">${j.error_message ? escapeHtml(String(j.error_message).slice(0, 64)) + (String(j.error_message).length > 64 ? "…" : "") : "—"}</td>
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
  } catch (e) {
    msg.textContent = String(e.message || e);
    tbody.innerHTML = '<tr><td colspan="6" class="empty">Could not load jobs.</td></tr>';
  }
}

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s;
  return d.innerHTML;
}

$("#job-form").addEventListener("submit", async (ev) => {
  ev.preventDefault();
  const formMsg = $("#form-msg");
  formMsg.textContent = "";
  const tenant_ids = parseTenantIds($("#tenant-ids").value);
  let options = {};
  try {
    options = JSON.parse($("#options-json").value || "{}");
  } catch {
    formMsg.textContent = "Options must be valid JSON.";
    return;
  }
  if (tenant_ids.length === 0) {
    formMsg.textContent = "Add at least one tenant ID.";
    return;
  }
  try {
    await api("/api/v1/jobs/bulk", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ tenant_ids, options }),
    });
    formMsg.textContent = "Job created.";
    $("#tenant-ids").value = "";
    await loadJobs();
  } catch (e) {
    formMsg.textContent = String(e.message || e);
  }
});

$("#refresh").addEventListener("click", () => loadJobs());

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

async function initAuth() {
  const login = $("#auth-login");
  const out = $("#auth-logout");
  const gateBtn = $("#auth-gate-btn");

  try {
    const r = await fetch("/api/v1/auth/status");
    const s = await r.json();
    if (s.oidc_login_enabled) {
      login.style.display = "";
      const href = "/api/v1/auth/oidc/login";
      login.setAttribute("href", href);
      if (gateBtn) gateBtn.setAttribute("href", href);
      const hasToken = !!sessionStorage.getItem("eoa_bearer");
      out.style.display = hasToken ? "inline-block" : "none";
      if (!hasToken) {
        lockAppUi();
      } else {
        unlockAppUi();
      }
    } else {
      unlockAppUi();
    }
  } catch {
    unlockAppUi();
  }

  out.addEventListener("click", (e) => {
    e.preventDefault();
    sessionStorage.removeItem("eoa_bearer");
    out.style.display = "none";
    location.href = "/";
  });
}

(async () => {
  await initAuth();
  if (!appUnlocked) return;
  await loadJobs();
})();

setInterval(() => {
  if (!appUnlocked) return;
  loadJobs();
}, 8000);
