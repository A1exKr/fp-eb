"use strict";

// --------------------------------------------------------------------------- //
// Helpers
// --------------------------------------------------------------------------- //
async function api(method, path, body, isForm = false) {
    const opts = { method, headers: {} };
    if (body && !isForm) {
        opts.headers["Content-Type"] = "application/json";
        opts.body = JSON.stringify(body);
    } else if (body && isForm) {
        opts.body = body;
    }
    const res = await fetch(path, opts);
    if (!res.ok) {
        let msg = `${res.status} ${res.statusText}`;
        try {
            const j = await res.json();
            if (j && j.detail) msg = typeof j.detail === "string" ? j.detail : JSON.stringify(j.detail);
        } catch (_) { }
        throw new Error(msg);
    }
    if (res.status === 204) return null;
    const ct = res.headers.get("content-type") || "";
    return ct.includes("application/json") ? res.json() : res.text();
}

let toastTimer;
function toast(msg, isErr = false) {
    const t = document.getElementById("toast");
    t.textContent = msg;
    t.hidden = false;
    t.classList.toggle("err", isErr);
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (t.hidden = true), 3500);
}

const toArr = (s) => (s || "").split(",").map((x) => x.trim()).filter(Boolean);
const toCsv = (a) => (a || []).join(", ");
const escapeHtml = (s) =>
    String(s == null ? "" : s).replace(/[&<>"']/g, (c) =>
        ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );
const escapeAttr = (s) => escapeHtml(s).replace(/"/g, "&quot;");

function renderTable(tableEl, columns, rows, rowActions) {
    const head = "<thead><tr>" + columns.map((c) => `<th>${c.label}</th>`).join("") + "<th></th></tr></thead>";
    const body = rows
        .map((r) => {
            const tds = columns.map((c) => `<td>${escapeHtml(c.get(r))}</td>`).join("");
            return `<tr>${tds}<td><div class="act">${rowActions(r)}</div></td></tr>`;
        })
        .join("");
    tableEl.innerHTML = head + "<tbody>" + (body || `<tr><td colspan="${columns.length + 1}" style="color:#6b7280">No records yet.</td></tr>`) + "</tbody>";
}

function fillForm(form, values) {
    Object.entries(values).forEach(([k, v]) => {
        const el = form.elements[k];
        if (el) el.value = v == null ? "" : v;
    });
}

// --------------------------------------------------------------------------- //
// Caches + datalists
// --------------------------------------------------------------------------- //
let ratesCache = [];
let personnelCache = [];
let presetsCache = [];
let referenceCache = [];
let assetsCache = [];

function setDatalist(id, values) {
    const dl = document.getElementById(id);
    const uniq = [...new Set(values.filter(Boolean))];
    dl.innerHTML = uniq.map((v) => `<option value="${escapeAttr(v)}"></option>`).join("");
}
function refreshDatalists() {
    const roles = [...ratesCache.map((r) => r.role), ...personnelCache.flatMap((p) => p.roles || [])];
    setDatalist("roles-list", roles);
    setDatalist("assets-list", assetsCache.map((a) => a.id));
    setDatalist("personnel-list", personnelCache.map((p) => p.id));
}

// --------------------------------------------------------------------------- //
// Tabs
// --------------------------------------------------------------------------- //
document.getElementById("tabs").addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    const tab = b.dataset.tab;
    document.querySelectorAll(".tabs button").forEach((x) => x.classList.toggle("active", x === b));
    document.querySelectorAll(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === "tab-" + tab));
});

// --------------------------------------------------------------------------- //
// Unit rates
// --------------------------------------------------------------------------- //
const formRates = document.getElementById("form-rates");
formRates.addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = new FormData(formRates);
    const body = {
        role: f.get("role"),
        rate: parseFloat(f.get("rate")),
        currency: f.get("currency") || "USD",
        effective_from: f.get("effective_from") || null,
    };
    try {
        await api("POST", "/v1/admin/rates", body);
        formRates.reset();
        formRates.currency.value = "USD";
        await loadRates();
        toast("Rate saved");
    } catch (err) {
        toast(err.message, true);
    }
});
async function loadRates() {
    ratesCache = await api("GET", "/v1/admin/rates");
    renderTable(
        document.getElementById("table-rates"),
        [
            { label: "Role", get: (r) => r.role },
            { label: "Rate", get: (r) => r.rate },
            { label: "Currency", get: (r) => r.currency },
            { label: "Effective from", get: (r) => r.effective_from || "" },
        ],
        ratesCache,
        (r) => `<button data-act="edit-rate" data-id="${r.id}">Edit</button><button class="del" data-act="del-rate" data-id="${r.id}">Delete</button>`
    );
    refreshDatalists();
}

// --------------------------------------------------------------------------- //
// Personnel
// --------------------------------------------------------------------------- //
const formPersonnel = document.getElementById("form-personnel");
formPersonnel.addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = new FormData(formPersonnel);
    const body = {
        id: f.get("id"),
        name: f.get("name"),
        title: f.get("title") || "",
        roles: toArr(f.get("roles")),
        cv_asset_id: f.get("cv_asset_id") || null,
    };
    try {
        await api("POST", "/v1/admin/personnel", body);
        formPersonnel.reset();
        await loadPersonnel();
        toast("Team member saved");
    } catch (err) {
        toast(err.message, true);
    }
});
async function loadPersonnel() {
    personnelCache = await api("GET", "/v1/admin/personnel");
    renderTable(
        document.getElementById("table-personnel"),
        [
            { label: "ID", get: (r) => r.id },
            { label: "Name", get: (r) => r.name },
            { label: "Title", get: (r) => r.title },
            { label: "Roles", get: (r) => toCsv(r.roles) },
            { label: "CV asset", get: (r) => r.cv_asset_id || "" },
        ],
        personnelCache,
        (r) => `<button data-act="edit-person" data-id="${escapeAttr(r.id)}">Edit</button><button class="del" data-act="del-person" data-id="${escapeAttr(r.id)}">Delete</button>`
    );
    refreshDatalists();
}

// --------------------------------------------------------------------------- //
// Team presets (with assignments sub-editor)
// --------------------------------------------------------------------------- //
const formPresets = document.getElementById("form-presets");
const assignmentsTable = document.getElementById("assignments-table");

function renderAssignments(rows) {
    assignmentsTable.innerHTML =
        "<thead><tr><th>Role</th><th>Person id</th><th>Rate (blank = inherit)</th><th></th></tr></thead><tbody>" +
        rows
            .map(
                (a, i) => `<tr>
      <td><input data-f="role" list="roles-list" value="${escapeAttr(a.role || "")}"></td>
      <td><input data-f="person_id" list="personnel-list" value="${escapeAttr(a.person_id || "")}"></td>
      <td><input data-f="rate" type="number" step="0.01" value="${a.rate == null ? "" : a.rate}"></td>
      <td><button type="button" class="del" data-act="rm-assign" data-i="${i}">✕</button></td>
    </tr>`
            )
            .join("") +
        "</tbody>";
}
function gatherAssignments() {
    return [...assignmentsTable.querySelectorAll("tbody tr")]
        .map((tr) => {
            const g = (f) => tr.querySelector(`[data-f="${f}"]`).value;
            const rate = g("rate");
            return { role: g("role").trim(), person_id: g("person_id").trim() || null, rate: rate === "" ? null : parseFloat(rate) };
        })
        .filter((a) => a.role);
}
document.getElementById("add-assignment").addEventListener("click", () => {
    const rows = gatherAssignments();
    rows.push({ role: "", person_id: "", rate: null });
    renderAssignments(rows);
});
formPresets.addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = new FormData(formPresets);
    const body = {
        id: f.get("id"),
        name: f.get("name"),
        types: toArr(f.get("types")),
        assignments: gatherAssignments(),
    };
    try {
        await api("POST", "/v1/admin/presets", body);
        formPresets.reset();
        renderAssignments([]);
        await loadPresets();
        toast("Preset saved");
    } catch (err) {
        toast(err.message, true);
    }
});
async function loadPresets() {
    presetsCache = await api("GET", "/v1/admin/presets");
    renderTable(
        document.getElementById("table-presets"),
        [
            { label: "ID", get: (r) => r.id },
            { label: "Name", get: (r) => r.name },
            { label: "Types", get: (r) => toCsv(r.types) },
            { label: "Roles", get: (r) => (r.assignments || []).map((a) => `${a.role}${a.rate != null ? " @" + a.rate : ""}`).join(", ") },
        ],
        presetsCache,
        (r) => `<button data-act="edit-preset" data-id="${escapeAttr(r.id)}">Edit</button><button class="del" data-act="del-preset" data-id="${escapeAttr(r.id)}">Delete</button>`
    );
}

// --------------------------------------------------------------------------- //
// Reference projects
// --------------------------------------------------------------------------- //
const formReference = document.getElementById("form-reference");
formReference.addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = new FormData(formReference);
    const body = {
        id: f.get("id"),
        name: f.get("name"),
        project_type: f.get("project_type") || "",
        location: f.get("location") || "",
        keywords: toArr(f.get("keywords")),
        summary: f.get("summary") || "",
    };
    try {
        await api("POST", "/v1/admin/reference-projects", body);
        formReference.reset();
        await loadReference();
        toast("Reference project saved");
    } catch (err) {
        toast(err.message, true);
    }
});
async function loadReference() {
    referenceCache = await api("GET", "/v1/admin/reference-projects");
    renderTable(
        document.getElementById("table-reference"),
        [
            { label: "ID", get: (r) => r.id },
            { label: "Name", get: (r) => r.name },
            { label: "Type", get: (r) => r.project_type },
            { label: "Location", get: (r) => r.location },
            { label: "Keywords", get: (r) => toCsv(r.keywords) },
        ],
        referenceCache,
        (r) => `<button data-act="edit-ref" data-id="${escapeAttr(r.id)}">Edit</button><button class="del" data-act="del-ref" data-id="${escapeAttr(r.id)}">Delete</button>`
    );
}

// --------------------------------------------------------------------------- //
// Assets (upload + list)
// --------------------------------------------------------------------------- //
const formAsset = document.getElementById("form-asset-upload");
formAsset.addEventListener("submit", async (e) => {
    e.preventDefault();
    const fd = new FormData(formAsset);
    try {
        await api("POST", "/v1/admin/assets/upload", fd, true);
        formAsset.reset();
        await loadAssets();
        toast("Asset uploaded");
    } catch (err) {
        toast(err.message, true);
    }
});
async function loadAssets() {
    assetsCache = await api("GET", "/v1/admin/assets");
    renderTable(
        document.getElementById("table-assets"),
        [
            { label: "ID", get: (r) => r.id },
            { label: "Kind", get: (r) => r.kind },
            { label: "Role", get: (r) => r.role || "" },
            { label: "Ref project", get: (r) => r.reference_project_id || "" },
            { label: "Filename", get: (r) => r.filename || "" },
            { label: "Size", get: (r) => (r.size_bytes != null ? r.size_bytes : "") },
        ],
        assetsCache,
        (r) => `<button class="del" data-act="del-asset" data-id="${escapeAttr(r.id)}">Delete</button>`
    );
    refreshDatalists();
}

// --------------------------------------------------------------------------- //
// Row actions (delete / edit) via delegation
// --------------------------------------------------------------------------- //
document.addEventListener("click", async (e) => {
    const b = e.target.closest("button[data-act]");
    if (!b) return;
    const { act, id } = b.dataset;
    try {
        switch (act) {
            case "del-rate":
                await api("DELETE", "/v1/admin/rates/" + id);
                await loadRates();
                toast("Deleted");
                break;
            case "edit-rate": {
                const r = ratesCache.find((x) => String(x.id) === String(id));
                if (r) fillForm(formRates, { role: r.role, rate: r.rate, currency: r.currency, effective_from: r.effective_from || "" });
                break;
            }
            case "del-person":
                await api("DELETE", "/v1/admin/personnel/" + encodeURIComponent(id));
                await loadPersonnel();
                toast("Deleted");
                break;
            case "edit-person": {
                const p = personnelCache.find((x) => x.id === id);
                if (p) fillForm(formPersonnel, { id: p.id, name: p.name, title: p.title, roles: toCsv(p.roles), cv_asset_id: p.cv_asset_id || "" });
                break;
            }
            case "del-preset":
                await api("DELETE", "/v1/admin/presets/" + encodeURIComponent(id));
                await loadPresets();
                toast("Deleted");
                break;
            case "edit-preset": {
                const p = presetsCache.find((x) => x.id === id);
                if (p) {
                    fillForm(formPresets, { id: p.id, name: p.name, types: toCsv(p.types) });
                    renderAssignments(p.assignments || []);
                }
                break;
            }
            case "del-ref":
                await api("DELETE", "/v1/admin/reference-projects/" + encodeURIComponent(id));
                await loadReference();
                toast("Deleted");
                break;
            case "edit-ref": {
                const r = referenceCache.find((x) => x.id === id);
                if (r) fillForm(formReference, { id: r.id, name: r.name, project_type: r.project_type, location: r.location, keywords: toCsv(r.keywords), summary: r.summary });
                break;
            }
            case "del-asset":
                await api("DELETE", "/v1/admin/assets/" + encodeURIComponent(id));
                await loadAssets();
                toast("Deleted");
                break;
            case "rm-assign": {
                const rows = gatherAssignments();
                rows.splice(parseInt(b.dataset.i, 10), 1);
                renderAssignments(rows);
                break;
            }
        }
    } catch (err) {
        toast(err.message, true);
    }
});

// --------------------------------------------------------------------------- //
// Init
// --------------------------------------------------------------------------- //
async function loadWho() {
    try {
        const me = await api("GET", "/v1/me");
        document.getElementById("who").textContent = `${me.user}${me.groups && me.groups.length ? " · " + me.groups.join(", ") : ""}`;
    } catch (_) {
        document.getElementById("who").textContent = "not signed in";
    }
}

(async function init() {
    renderAssignments([]);
    await loadWho();
    for (const fn of [loadRates, loadPersonnel, loadPresets, loadReference, loadAssets]) {
        try {
            await fn();
        } catch (err) {
            toast(err.message, true);
        }
    }
})();
