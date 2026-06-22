const { useState, useEffect, useCallback, useRef, useMemo, createContext, useContext } = React;

// ── FAVICON ──
(() => { const l = document.querySelector("link[rel='icon']") || document.createElement("link"); l.rel = "icon"; l.href = "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='6' fill='%23CDA04B'/><text x='16' y='23' font-family='Georgia,serif' font-size='22' font-weight='bold' fill='%2328434C' text-anchor='middle'>L</text></svg>"; document.head.appendChild(l); })();

// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
  clientId: "32e75ffa-747a-4cf0-8209-6a19150c4547",
  tenantId: "33575d04-ca7b-4396-8011-9eaea4030b46",
  siteId: "vanrockre.sharepoint.com,a02c1cd8-9f1f-4827-8286-7b6b7ce74232,01202419-6625-4499-b0d5-8ceb1cffdba3",
  appName: "EMPLOYEE LIFECYCLE",
  version: "0.3.0",
  lists: {
    employees:           "Employees",            // shared
    journeys:            "ELC_Journeys",
    templateTasks:       "ELC_TemplateTasks",
    journeyTasks:        "ELC_JourneyTasks",
    config:              "ELC_Config",
    apps:                "ELC_Apps",              // registry of NewShire apps + their per-app role columns
    notes:               "ELC_EmployeeNotes",     // coaching / discipline / PIP / 1:1 / praise
    audit:               "ELC_PermissionAudit",   // log of every per-app role change
    files:               "ELC_EmployeeFiles",     // SharePoint document library
    reviews:             "ELC_QuarterlyReviews",  // quarterly performance reviews
    payChanges:          "ELC_PayChanges",        // compensation history
    emailTemplates:      "ELC_EmailTemplates",    // reusable email templates with {{var}} substitution
    // Cross-app reads — NewShire University LMS lists. Read-only here.
    luCourses:           "TrainingCourses",
    luAssignments:       "TrainingAssignments",
    luCompletions:       "TrainingCompletions",
  },
  // NewShire University deep-link base — used in Training Compliance cards
  universityUrl:       "https://newshirepm.github.io/newshire-university/",
  reviewIntervalDays: 90,         // expected cadence between reviews (once a baseline exists)
  reviewOverdueDays: 100,         // flag as overdue once gap exceeds this
  // First-cycle anchor: until each employee has a baseline review, this is the
  // org-wide due date that's used in place of an "X days since last" overdue count.
  // After June 30, 2026 passes with no review, the normal overdue flagging kicks in.
  firstReviewDueDate: "2026-06-30",
  // Employees who don't need a quarterly review tracked here — typically the
  // highest-ranking people whose performance is handled directly by ownership.
  reviewExemptEmails: ["jwhite@vanrockre.com"],
  adminEmails: ["bturner@newshirepm.com"],
  filesLibraryUrl: "https://vanrockre.sharepoint.com/ELC_EmployeeFiles",
  // Predefined groups shown in dropdowns. Users can also create custom groups by typing.
  templateGroups: {
    Onboarding: [
      "On Site Team Member",
      "Off Site Team Member",
      "Virtual Assistant",
      "Maintenance",
      "Corporate Office Team Member",
    ],
    Offboarding: [
      "Voluntary Resignation",
      "Involuntary Termination",
    ],
  },
};

const GRAPH = "https://graph.microsoft.com/v1.0";
const SITE = `${GRAPH}/sites/${CONFIG.siteId}`;
const SCOPES = ["Sites.ReadWrite.All", "User.Read", "Mail.Send"];

// ============================================================
// PALETTE — NewShire light theme
// ============================================================
const C = {
  hdr: "#1C3740", hdrH: "#213F4A",
  t7: "#28434C", t6: "#2F5260", t5: "#3A6577", t4: "#4A7E91", t3: "#6FA0B0", t1: "#D6E7EC", t0: "#EDF4F7",
  g7: "#9E7B2F", g6: "#B8922E", g5: "#CDA04B", g4: "#D4AF61", g1: "#F8F0DB", g0: "#FFFBF0",
  wh: "#FFFFFF", pg: "#F7F8F7",
  b1: "#E8EAEA", b2: "#D0D8DC", b3: "#A8B0B0", b4: "#7A8585", b6: "#3E4A4A",
  ok: "#2D8A5A", okb: "rgba(45,138,90,0.08)",
  er: "#C44B3B", erb: "rgba(196,75,59,0.06)",
  wn: "#D4960A", wnb: "rgba(212,150,10,0.08)",
  inf: "#4A78B0", infb: "rgba(74,120,176,0.08)",
  pu: "#5B3FA8", pub: "rgba(91,63,168,0.08)",
};
const font = "'Source Sans 3','Segoe UI',sans-serif";
const mono = "'Source Code Pro',Consolas,monospace";

// ============================================================
// STYLES
// ============================================================
const S = {
  page: { fontFamily: font, background: C.pg, minHeight: "100vh", color: C.t7 },
  hdr: { background: C.hdr, borderBottom: `2.5px solid ${C.g5}`, padding: "0 20px", display: "flex", alignItems: "center", justifyContent: "space-between", minHeight: 56, position: "sticky", top: 0, zIndex: 100 },
  hdrL: { display: "flex", alignItems: "center", gap: 12 },
  logo: { width: 32, height: 32, background: C.g5, borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, fontWeight: 700, color: C.t7, fontFamily: "Georgia,serif" },
  hdrTitle: { color: "#fff", fontSize: 15, fontWeight: 700, letterSpacing: ".08em" },
  hdrSub: { color: C.g4, fontSize: 10, letterSpacing: ".1em", textTransform: "uppercase" },
  hdrUser: { color: C.t1, fontSize: 13, fontWeight: 600 },
  hdrRole: { fontSize: 10, color: C.g4, textTransform: "uppercase", letterSpacing: ".07em" },
  tabBar: { background: C.wh, borderBottom: `1px solid ${C.b2}`, display: "flex", padding: "0 16px", overflowX: "auto", WebkitOverflowScrolling: "touch", position: "sticky", top: 56, zIndex: 90 },
  tab: a => ({ padding: "11px 15px", fontSize: 13, fontWeight: a ? 600 : 400, color: a ? C.t7 : C.b4, borderBottom: `2.5px solid ${a ? C.g5 : "transparent"}`, cursor: "pointer", whiteSpace: "nowrap", background: "none", border: "none", borderTop: "none", borderLeft: "none", borderRight: "none", fontFamily: font, minHeight: 44 }),
  content: { maxWidth: 1240, margin: "0 auto", padding: "20px 18px 60px" },
  card: { background: C.wh, border: `1px solid ${C.b2}`, borderRadius: 6, boxShadow: "0 1px 3px rgba(28,55,64,.06)", padding: 16, marginBottom: 14 },
  cardT: { fontSize: 15, fontWeight: 600, color: C.t7, paddingBottom: 10, borderBottom: `1px solid ${C.b1}`, marginBottom: 12, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 },
  cardS: { fontSize: 12, color: C.b4 },
  sec: { fontSize: 10, fontWeight: 700, color: C.b4, textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 8 },
  label: { display: "block", fontSize: 12, fontWeight: 600, color: C.t7, marginBottom: 4 },
  input: { width: "100%", padding: "8px 10px", fontSize: 13, fontFamily: font, color: C.t7, background: C.wh, border: `1px solid ${C.b2}`, borderRadius: 4, outline: "none", boxSizing: "border-box" },
  select: { width: "100%", padding: "8px 10px", fontSize: 13, fontFamily: font, color: C.t7, background: C.wh, border: `1px solid ${C.b2}`, borderRadius: 4, cursor: "pointer", boxSizing: "border-box" },
  textarea: { width: "100%", padding: "8px 10px", fontSize: 13, fontFamily: font, color: C.t7, background: C.wh, border: `1px solid ${C.b2}`, borderRadius: 4, outline: "none", boxSizing: "border-box", resize: "vertical", minHeight: 64 },
  btn: (bg, fg) => ({ display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 6, padding: "8px 14px", fontSize: 12, fontWeight: 600, fontFamily: font, color: fg || "#fff", background: bg || C.hdr, border: "none", borderRadius: 4, cursor: "pointer", minHeight: 38, whiteSpace: "nowrap" }),
  btnO: (fg, bdr) => ({ display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 6, padding: "8px 14px", fontSize: 12, fontWeight: 600, fontFamily: font, color: fg || C.t7, background: C.wh, border: `1px solid ${bdr || C.b2}`, borderRadius: 4, cursor: "pointer", minHeight: 38, whiteSpace: "nowrap" }),
  xs: { padding: "4px 10px", fontSize: 11, minHeight: 28 },
  sm: { padding: "6px 12px", fontSize: 12, minHeight: 32 },
  th: { textAlign: "left", padding: "9px 11px", fontSize: 11, fontWeight: 700, color: C.t7, background: C.t0, borderBottom: `2px solid ${C.t1}`, whiteSpace: "nowrap" },
  td: { padding: "9px 11px", fontSize: 13, color: C.b6, borderBottom: `1px solid ${C.b1}`, verticalAlign: "top" },
  kpi: { background: C.wh, border: `1px solid ${C.b2}`, borderRadius: 6, padding: "12px 14px", flex: "1 1 130px", minWidth: 120 },
  kpiL: { fontSize: 10, fontWeight: 700, color: C.b4, textTransform: "uppercase", letterSpacing: ".07em" },
  kpiV: { fontSize: 26, fontWeight: 700, fontFamily: mono, color: C.t7, lineHeight: 1.1, marginTop: 4 },
  modalBg: { position: "fixed", inset: 0, background: "rgba(28,55,64,.45)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "60px 16px 20px", zIndex: 1000, overflowY: "auto" },
  modal: { background: C.wh, borderRadius: 8, maxWidth: 720, width: "100%", boxShadow: "0 10px 40px rgba(28,55,64,.25)", overflow: "hidden" },
  modalH: { padding: "14px 18px", borderBottom: `1px solid ${C.b1}`, display: "flex", alignItems: "center", justifyContent: "space-between", background: C.t0 },
  modalB: { padding: 18 },
  modalF: { padding: "12px 18px", borderTop: `1px solid ${C.b1}`, display: "flex", justifyContent: "flex-end", gap: 8, background: C.pg },
};

// ============================================================
// META TEMPLATE GROUPS — auto-apply based on rules; not user-selectable
// at journey start. Empty TemplateGroup = applies to every group of that
// JourneyType (truly universal). A named meta-group has an appliesTo rule
// that decides whether it runs given the user's chosen group.
// ============================================================
const META_GROUP_RULES = {
  "Common - W-2": {
    journeyType: "Onboarding",
    label: "Common — W-2 onboarding (excludes Virtual Assistant & 1099 contractors)",
    // Runs for every onboarding group EXCEPT Virtual Assistant (1099 contractor) and
    // any hire flagged as a 1099 contractor — those get the Common - 1099 track instead.
    appliesTo: (selectedGroup, opts) => selectedGroup !== "Virtual Assistant" && !opts?.isContractor,
  },
  "Common - 1099": {
    journeyType: "Onboarding",
    label: "Common — 1099 contractor onboarding",
    // Runs only when the hire is flagged as a 1099 contractor. Virtual Assistant has its
    // own dedicated contractor track, so this meta-group does not also pile on for VA.
    appliesTo: (selectedGroup, opts) => !!opts?.isContractor && selectedGroup !== "Virtual Assistant",
  },
};
function isMetaGroup(name) { return Object.prototype.hasOwnProperty.call(META_GROUP_RULES, name); }
// A template can apply to multiple groups by storing a comma- or semicolon-
// separated list in TemplateGroup. Empty = universal (applies to every group).
function templateGroups(tpl) {
  if (!tpl || !tpl.TemplateGroup) return [];
  return String(tpl.TemplateGroup).split(/[;,]/).map(s => s.trim()).filter(Boolean);
}
function templateAppliesToGroup(tpl, selectedGroup, opts) {
  const groups = templateGroups(tpl);
  if (groups.length === 0) return true;                                // universal
  if (groups.includes(selectedGroup)) return true;                     // direct hit
  for (const g of groups) {
    const meta = META_GROUP_RULES[g];
    if (meta && meta.appliesTo(selectedGroup, opts)) return true;      // meta hit
  }
  return false;
}

// ============================================================
// HELPERS
// ============================================================
// These are calendar dates (start day, due day) — NOT instants. Parse them as LOCAL
// dates so they never shift across the UTC boundary (e.g. "2026-05-27" must display as
// May 27 in EST, not May 26). localDate() reads the leading YYYY-MM-DD from a date-only
// string OR a full ISO timestamp and builds a local-midnight Date.
function localDate(value) {
  if (!value) return null;
  const m = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}
function ymd(d) { const p = n => String(n).padStart(2, "0"); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`; }
function todayIso() { return ymd(new Date()); }
function addDays(iso, n) { const d = localDate(iso) || new Date(); d.setDate(d.getDate() + (Number(n) || 0)); return ymd(d); }
function fmtDate(iso) { const d = localDate(iso); return d ? d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) : "—"; }
function daysFromNow(iso) {
  const d = localDate(iso); if (!d) return null;
  const t = new Date();
  const target = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const today = new Date(t.getFullYear(), t.getMonth(), t.getDate());
  return Math.round((target - today) / 86400000);
}
function initials(name) { if (!name) return "?"; const p = name.trim().split(/\s+/); return p.length > 1 ? (p[0][0] + p[p.length - 1][0]).toUpperCase() : name.slice(0, 2).toUpperCase(); }
function classifyDue(due, status) { if (status === "Done") return "done"; if (!due) return "none"; const d = daysFromNow(due); if (d == null) return "none"; if (d < 0) return "overdue"; if (d <= 3) return "soon"; return "ok"; }
function uniq(arr) { return Array.from(new Set(arr.filter(Boolean))); }
// Resolve the manager hierarchy from the Employees list so onboarding fills the same
// ManagerName / Level2 / Level3 columns the rest of the list uses. Walks up to 3 levels
// by following each manager's own ManagerEmail. Unresolved levels stay "" (cleanFields
// drops them, leaving the column blank rather than erroring).
function managerChain(managerEmail, employees) {
  const out = { ManagerName: "", Level2ManagerEmail: "", Level2ManagerName: "", Level3ManagerEmail: "", Level3ManagerName: "" };
  const find = email => employees.find(e => (e.Email || "").toLowerCase() === (email || "").toLowerCase());
  const nameOf = e => (e && (e.Title || e.Name)) || "";
  const m1 = find(managerEmail);
  if (!m1) return out;
  out.ManagerName = nameOf(m1);
  out.Level2ManagerEmail = (m1.ManagerEmail || "").toLowerCase();
  const m2 = find(out.Level2ManagerEmail);
  if (!m2) return out;
  out.Level2ManagerName = nameOf(m2);
  out.Level3ManagerEmail = (m2.ManagerEmail || "").toLowerCase();
  const m3 = find(out.Level3ManagerEmail);
  if (m3) out.Level3ManagerName = nameOf(m3);
  return out;
}

// ============================================================
// GRAPH API
// ============================================================
const lUrl = n => `${SITE}/lists/${encodeURIComponent(n)}/items`;
const iUrl = (n, id) => `${SITE}/lists/${encodeURIComponent(n)}/items/${id}/fields`;

async function gGet(token, url) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw new Error(`GET ${r.status} ${url.split('/').slice(-2).join('/')}`);
  return r.json();
}
async function gAll(token, url) {
  let out = [], next = url;
  while (next) {
    const d = await gGet(token, next);
    out = out.concat(d.value || []);
    next = d["@odata.nextLink"] || null;
  }
  return out;
}
// Strip empty/undefined values + SharePoint system fields — SharePoint rejects
// "" for Date/Number/Boolean columns with "badArgument", and rejects PATCHing
// the system-owned fields (id, Created, Modified, etc.) with "Field 'id' is not
// recognized". Keeps false (Yes/No) and 0 (Number) intentionally.
const SYSTEM_FIELDS = new Set([
  // Item identity / audit
  "id", "Created", "Modified", "_UIVersionString",
  "ContentType", "ContentTypeId", "Attachments", "Order",
  "AuthorLookupId", "EditorLookupId", "Author", "Editor",
  "LinkTitle", "LinkTitleNoMenu",
  "AppAuthor", "AppEditor", "AppAuthorLookupId", "AppEditorLookupId",
  "AppCreated", "AppModified",
  // Counts and folder/file metadata that come back from Graph but are
  // never user-writable on a list item.
  "ItemChildCount", "FolderChildCount",
  "FileSystemObjectType", "FSObjType",
  "FileRef", "FileDirRef", "FileLeafRef", "ParentLeafRef",
  "Path", "UniqueId", "DocIcon",
  "ServerRedirectedEmbedUri", "ServerRedirectedEmbedUrl",
  "ProgId", "ScopeId", "HTML_x0020_File_x0020_Type",
  // Default view internal columns
  "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime",
  "_ComplianceTagUserId", "_IsRecord",
]);
function cleanFields(fields) {
  const out = {};
  for (const [k, v] of Object.entries(fields || {})) {
    if (v === undefined || v === null) continue;
    if (typeof v === "string" && v.trim() === "") continue;
    if (SYSTEM_FIELDS.has(k)) continue;
    if (k.startsWith("@") || k.startsWith("_")) continue;
    out[k] = v;
  }
  return out;
}
async function gPost(token, url, fields) {
  const r = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ fields: cleanFields(fields) }),
  });
  if (!r.ok) throw new Error(`POST ${r.status} ${await r.text().catch(() => '')}`);
  return r.json();
}
async function gPatch(token, url, fields) {
  const r = await fetch(url, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(cleanFields(fields)),
  });
  if (!r.ok) throw new Error(`PATCH ${r.status} ${await r.text().catch(() => '')}`);
  return r.json();
}
async function gDelete(token, url) {
  const r = await fetch(url, { method: "DELETE", headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok && r.status !== 204) throw new Error(`DELETE ${r.status}`);
}
async function sendEmail(token, to, subject, body) {
  try {
    await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: { subject, body: { contentType: "HTML", content: body }, toRecipients: [{ emailAddress: { address: to } }] },
        saveToSentItems: false,
      }),
    });
  } catch (e) { console.warn("[ELC] sendEmail failed:", e.message); }
}

async function safeGet(token, name, url) { try { return await gAll(token, url); } catch (e) { console.warn(`[ELC] ${name} load failed:`, e.message); return []; } }

// ============================================================
// MSAL HOOK
// ============================================================
function useMsal() {
  const [inst, setInst] = useState(null);
  const [acct, setAcct] = useState(null);
  const [token, setToken] = useState(null);
  const [err, setErr] = useState(null);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    const init = async () => {
      try {
        if (!window.msal) await new Promise((res, rej) => {
          const existing = document.querySelector("script[src*='msal-browser']");
          if (existing) { existing.addEventListener('load', res); existing.addEventListener('error', () => rej(new Error('MSAL load failed'))); setTimeout(() => window.msal ? res() : null, 500); }
          else {
            const s = document.createElement("script");
            s.src = "https://unpkg.com/@azure/msal-browser@2.38.3/lib/msal-browser.min.js";
            s.onload = res; s.onerror = () => rej(new Error("MSAL load failed"));
            document.head.appendChild(s);
          }
        });
        const i = new window.msal.PublicClientApplication({
          auth: { clientId: CONFIG.clientId, authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`, redirectUri: window.location.origin + window.location.pathname },
          cache: { cacheLocation: "sessionStorage" },
        });
        await i.initialize();
        try { const rr = await i.handleRedirectPromise(); if (rr) setAcct(rr.account); } catch (e) { /* ignore */ }
        setInst(i);
        const accounts = i.getAllAccounts();
        if (accounts.length > 0) {
          setAcct(accounts[0]);
          try { const r = await i.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] }); setToken(r.accessToken); } catch (e) { /* user will need to click sign in */ }
        }
        setReady(true);
      } catch (e) { setErr(e.message); setReady(true); }
    };
    init();
  }, []);

  const login = useCallback(async () => {
    if (!inst) return;
    try {
      const r = await inst.loginPopup({ scopes: SCOPES });
      setAcct(r.account);
      const t = await inst.acquireTokenSilent({ scopes: SCOPES, account: r.account });
      setToken(t.accessToken); setErr(null);
    } catch (e) {
      if (e.errorCode === "user_cancelled") return;
      if (e.errorCode === "popup_window_error" || e.errorCode === "empty_window_error") {
        try { await inst.loginRedirect({ scopes: SCOPES }); } catch (e2) { setErr(e2.message); }
        return;
      }
      setErr(e.message);
    }
  }, [inst]);

  const refresh = useCallback(async () => {
    if (!inst || !acct) return null;
    try { const r = await inst.acquireTokenSilent({ scopes: SCOPES, account: acct }); setToken(r.accessToken); return r.accessToken; }
    catch { try { const r = await inst.acquireTokenPopup({ scopes: SCOPES }); setToken(r.accessToken); return r.accessToken; } catch (e) { setErr(e.message); return null; } }
  }, [inst, acct]);

  const logout = useCallback(() => { if (!inst) return; inst.logoutPopup().catch(() => {}); setAcct(null); setToken(null); }, [inst]);

  return { acct, token, login, logout, refresh, err, ready };
}

// ============================================================
// DEFAULT TEMPLATE TASKS — seed when ELC_TemplateTasks is empty
// ============================================================
// Convention: TemplateGroup === "" means the task applies to EVERY group of that JourneyType
// (truly universal). A named TemplateGroup limits the task to that group. A meta-group name
// (see META_GROUP_RULES, e.g. "Common - W-2") applies via a rule — Common-W-2 runs for every
// onboarding group EXCEPT Virtual Assistant.
//
// Source: NewShire_Onboarding_Offboarding_Checklists.xlsx (authored by Brandy Turner).
const DEFAULT_TEMPLATES = [
  // ═════════════════════════════════════════════════════════════════════
  // ONBOARDING — Common - W-2  (auto-applies to all W-2 onboarding groups,
  // i.e. Corporate / On-Site / Off-Site / Maintenance. Excluded for VA.)
  // ═════════════════════════════════════════════════════════════════════
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Pre-Offer", Title:"Background check authorization signed", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Pre-Offer", Title:"Reference checks completed", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Pre-Offer", Title:"Compensation approved (Cara / John)", AssigneeRole:"Admin", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"T-7 Days", Title:"Offer letter generated and sent", AssigneeRole:"Admin", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"T-7 Days", Title:"Offer letter signed and returned", AssigneeRole:"Employee", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"T-7 Days", Title:"Send employment paperwork packet (I-9, W-4, SC W-4, direct deposit, emergency contact, beneficiary)", AssigneeRole:"HR", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"T-7 Days", Title:"Complete I-9 Section 1, W-4, SC W-4, direct deposit forms", AssigneeRole:"Employee", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"T-7 Days", Title:"Submit emergency contact and beneficiary forms", AssigneeRole:"Employee", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"T-7 Days", Title:"Confirm Day 1 logistics with new hire", AssigneeRole:"Manager", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"T-5 Days", Title:"Create M365 account and assign license", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Day 1",    Title:"Bring acceptable I-9 ID documents", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"Passport OR DL + SSN card/birth cert." },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Day 1",    Title:"Complete I-9 Section 2 with ID verification", AssigneeRole:"HR", OffsetDays:0, Required:true, Notes:"Federal deadline: Day 3." },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Day 1",    Title:"Set M365 password, enable MFA", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Day 1",    Title:"Sign handbook, NDA, harassment policy, IT acceptable use", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Day 1",    Title:"Take photo for directory and ID badge", AssigneeRole:"HR", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Day 30",   Title:"30-day check-in: what's working, what's unclear, where stuck", AssigneeRole:"Manager", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Day 60",   Title:"Role competency review", AssigneeRole:"Manager", OffsetDays:60, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Day 90",   Title:"Probation review and benefits enrollment", AssigneeRole:"HR", OffsetDays:90, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - W-2", Phase:"Day 90",   Title:"Complete benefits enrollment (election deadline is firm)", AssigneeRole:"Employee", OffsetDays:90, Required:true, Notes:"" },

  // ═════════════════════════════════════════════════════════════════════
  // ONBOARDING — Corporate Office Team Member
  // Office-based roles at 333 Wade Hampton Blvd. Operations, accounting, admin.
  // ═════════════════════════════════════════════════════════════════════
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-5 Days", Title:"Add to M365 security groups (Corporate Office)", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-5 Days", Title:"Create AppFolio user with role-based permissions", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-5 Days", Title:"Assign phone extension", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-3 Days", Title:"Add to SharePoint sites (NewShire main + role-specific)", AssigneeRole:"Admin", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-3 Days", Title:"Add to Teams channels (General + role-specific)", AssigneeRole:"Admin", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-3 Days", Title:"Order laptop, monitor, keyboard, mouse, headset", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-3 Days", Title:"Order business cards and name plate", AssigneeRole:"Manager", OffsetDays:-3, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-1 Day", Title:"Prepare workspace (desk, chair, supplies)", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-1 Day", Title:"Notify office staff of new hire arrival", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-1 Day", Title:"Confirm Day 1 logistics with new hire", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"T-1 Day", Title:"Confirm start time, location, dress code, what to bring", AssigneeRole:"Employee", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 1", Title:"Welcome meeting and building tour", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 1", Title:"Issue equipment, building keys, badge", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 1", Title:"Install Teams/Outlook on workstation", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 1", Title:"Set up email signature using NewShire template", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 1", Title:"Add profile photo to Teams and AppFolio", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 1", Title:"Complete NewShire University orientation module", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"When LMS live." },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 1", Title:"Team lunch / intro meetings", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Week 1", Title:"Role-specific training: AppFolio modules and SOPs", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Week 1", Title:"Portfolio overview (properties, owners, key residents)", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Week 1", Title:"Read SOPs for role; save SOP library to SharePoint Favorites", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Week 1", Title:"Schedule intro 1:1s with key team members and contacts", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Week 1", Title:"Shadow direct manager", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Weeks 2-4", Title:"Paired work with experienced team member", AssigneeRole:"Manager", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Weeks 2-4", Title:"Begin solo low-risk tasks", AssigneeRole:"Employee", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Weeks 2-4", Title:"Complete required NewShire University modules", AssigneeRole:"Employee", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 30", Title:"Self-assessment prepared for check-in", AssigneeRole:"Employee", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 60", Title:"Adjust training plan if gaps identified", AssigneeRole:"Manager", OffsetDays:60, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Corporate Office Team Member", Phase:"Day 90", Title:"Confirm permanent status", AssigneeRole:"Manager", OffsetDays:90, Required:true, Notes:"" },

  // ═════════════════════════════════════════════════════════════════════
  // ONBOARDING — On Site Team Member
  // Property manager / on-site leasing at a single property.
  // ═════════════════════════════════════════════════════════════════════
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Pre-Offer", Title:"Background check — ENHANCED (unit entry, resident interaction)", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"Supplements the Common-W-2 background check." },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Pre-Offer", Title:"Drug screen completed (industry standard for on-site)", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Pre-Offer", Title:"If employee unit: verify availability and lease/license terms", AssigneeRole:"Admin", OffsetDays:-21, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-7 Days", Title:"If employee unit: lease/license agreement prepared", AssigneeRole:"Admin", OffsetDays:-7, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-7 Days", Title:"Confirm Day 1 logistics, dress code, and parking", AssigneeRole:"Manager", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-5 Days", Title:"Create AppFolio user — SCOPED to specific property only", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-5 Days", Title:"Add to on-call rotation calendar", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-3 Days", Title:"Add to SharePoint sites and property-specific Teams channel", AssigneeRole:"Admin", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-3 Days", Title:"Order mobile equipment (laptop or tablet for AppFolio mobile)", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-3 Days", Title:"Order business cards with property branding", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-3 Days", Title:"Prepare property master key set (key control log)", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-3 Days", Title:"Prepare on-site office key", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-1 Day", Title:"Pull vendor contact list for the property", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-1 Day", Title:"Pull current resident roster and notable accounts", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-1 Day", Title:"If employee unit: complete pre-move-in inspection", AssigneeRole:"Manager", OffsetDays:-1, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"T-1 Day", Title:"Confirm Day 1 logistics", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 1", Title:"Sign master key control log", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 1", Title:"Sign on-site office key receipt", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 1", Title:"Equipment handoff: laptop/tablet, mobile, AppFolio mobile app installed", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 1", Title:"Set up mobile apps (AppFolio mobile, Teams)", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 1", Title:"Sign on-call policy (in addition to common handbook/NDA/etc.)", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 1", Title:"CONDUCT FULL PROPERTY TOUR: every unit, common areas, mechanical, problem areas", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 1", Title:"Review property-specific SOPs and emergency protocols", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 1", Title:"Review on-call rotation and after-hours expectations", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 1", Title:"If employee unit: lease/license signed, move-in inspection, keys issued", AssigneeRole:"Manager", OffsetDays:0, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Week 1", Title:"Open and close property procedures walked through", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Week 1", Title:"Common area inspection routine reviewed", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Week 1", Title:"Vendor introductions (those who work the property)", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Week 1", Title:"Review property-specific lease terms and special provisions", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Week 1", Title:"Review pending leasing prospects and open work orders", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Week 1", Title:"Resident introductions (or planned announcement)", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Week 1", Title:"Schedule intro 1:1s with key contacts", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Weeks 2-4", Title:"Shadow + paired leasing tours / move-ins / move-outs", AssigneeRole:"Manager", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Weeks 2-4", Title:"Begin solo low-complexity resident interactions", AssigneeRole:"Employee", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Weeks 2-4", Title:"Complete required NewShire University modules including fair housing", AssigneeRole:"Employee", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 30", Title:"Verify all property systems familiarity", AssigneeRole:"Manager", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Day 60", Title:"Leasing performance review (supplements common competency review)", AssigneeRole:"Manager", OffsetDays:60, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Compliance", Title:"Confirm SC Property Manager licensing", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"Required for property management functions in SC." },
  { JourneyType:"Onboarding", TemplateGroup:"On Site Team Member", Phase:"Compliance", Title:"Fair housing training completed before any leasing decisions", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },

  // ═════════════════════════════════════════════════════════════════════
  // ONBOARDING — Off Site Team Member
  // Regional/floating/multi-property. Vehicle + multi-property key control.
  // ═════════════════════════════════════════════════════════════════════
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Pre-Offer", Title:"Background check — ENHANCED (multi-property unit entry, resident interaction)", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"Supplements the Common-W-2 background check." },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Pre-Offer", Title:"MVR check completed (driving multiple properties)", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Pre-Offer", Title:"Drug screen completed", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Pre-Offer", Title:"If personal vehicle: verify personal auto insurance with business use rider", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-7 Days", Title:"Decide: company vehicle vs personal vehicle + mileage reimbursement", AssigneeRole:"Manager", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-5 Days", Title:"Create AppFolio user — scoped to assigned portfolio", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-5 Days", Title:"Mobile phone setup (work expected on the go)", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-5 Days", Title:"If company vehicle: schedule assignment, registration, insurance, fuel card", AssigneeRole:"Manager", OffsetDays:-5, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-3 Days", Title:"Add to SharePoint sites and Teams channels (multi-property)", AssigneeRole:"Admin", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-3 Days", Title:"Order laptop, mobile, peripherals", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-3 Days", Title:"Prepare master key sets for each assigned property (key control log)", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-3 Days", Title:"Compile property access lists and vendor contacts", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-3 Days", Title:"If mileage reimbursement: mileage tracking app selected", AssigneeRole:"Manager", OffsetDays:-3, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-1 Day", Title:"Build Week 1 property tour schedule", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"T-1 Day", Title:"Confirm Day 1 logistics", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Day 1", Title:"Bring driver's license (for vehicle assignment)", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Day 1", Title:"Sign master key control log (all assigned properties)", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Day 1", Title:"Equipment handoff: laptop, mobile, AppFolio mobile app", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Day 1", Title:"Vehicle handoff (if company vehicle): inspection, registration, insurance, fuel card", AssigneeRole:"Manager", OffsetDays:0, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Day 1", Title:"Install mileage tracking app if applicable", AssigneeRole:"Employee", OffsetDays:0, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Day 1", Title:"Sign vehicle use policy + on-call (in addition to common handbook/NDA/etc.)", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Day 1", Title:"Review assigned portfolio and coverage expectations", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Week 1", Title:"FULL TOUR EACH ASSIGNED PROPERTY: units, common areas, mechanical, problem areas", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Week 1", Title:"Meet on-site staff at each property", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Week 1", Title:"Vendor introductions at each property", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Week 1", Title:"Review property-specific SOPs and emergency protocols", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Week 1", Title:"Schedule intro 1:1s with key contacts", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Weeks 2-4", Title:"Paired work covering multiple properties", AssigneeRole:"Manager", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Weeks 2-4", Title:"Begin solo low-risk tasks across portfolio", AssigneeRole:"Employee", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Off Site Team Member", Phase:"Weeks 2-4", Title:"Complete required NewShire University modules including fair housing", AssigneeRole:"Employee", OffsetDays:14, Required:true, Notes:"" },

  // ═════════════════════════════════════════════════════════════════════
  // ONBOARDING — Common - 1099  (independent contractor; replaces the W-2
  // paperwork track when the "1099 contractor" box is checked at journey start.
  // Group-neutral so it layers onto On Site / Off Site / Maintenance / Corporate.
  // Excluded for Virtual Assistant, which has its own dedicated contractor track.)
  // ═════════════════════════════════════════════════════════════════════
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"Pre-Engagement", Title:"Independent Contractor Agreement signed (NOT an employment offer)", AssigneeRole:"Admin", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"Pre-Engagement", Title:"Scope of Work / deliverables defined and signed", AssigneeRole:"Admin", OffsetDays:-21, Required:true, Notes:"Define deliverables, not hours/schedule — preserves contractor status." },
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"Pre-Engagement", Title:"Confidentiality / NDA signed", AssigneeRole:"Admin", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"Pre-Engagement", Title:"Acknowledge contractor status: no tax withholding, no employee benefits or wage protections", AssigneeRole:"Employee", OffsetDays:-21, Required:true, Notes:"Replaces the W-2 I-9 / W-4 / SC W-4 / direct-deposit packet." },
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"Pre-Engagement", Title:"W-9 collected (US) OR W-8BEN (foreign contractor)", AssigneeRole:"Accounting", OffsetDays:-21, Required:true, Notes:"Needed for 1099-NEC reporting." },
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"Pre-Engagement", Title:"If applicable: collect Certificate of Insurance (GL / workers' comp waiver)", AssigneeRole:"Admin", OffsetDays:-14, Required:false, Notes:"Common for on-site or trade contractors." },
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"T-7 Days", Title:"Confirm pay rate, invoicing process, and payment frequency", AssigneeRole:"Accounting", OffsetDays:-7, Required:true, Notes:"Contractor invoices; paid outside payroll." },
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"T-5 Days", Title:"Provision only the system access the engagement requires", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"Scope access to the deliverable; avoid employee-style full access." },
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"Compliance", Title:"Misclassification risk review: behavioral control, financial control, relationship (IRS/DOL)", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"Worker should control means & methods. Re-evaluate if scope drifts toward employee-like control." },
  { JourneyType:"Onboarding", TemplateGroup:"Common - 1099", Phase:"Annual", Title:"1099-NEC issued at year end (US contractors paid ≥ $600)", AssigneeRole:"Accounting", OffsetDays:365, Required:true, Notes:"" },

  // ═════════════════════════════════════════════════════════════════════
  // ONBOARDING — Virtual Assistant  (1099 contractor — distinct from W-2)
  // No I-9, no benefits, no UI eligibility. Misclassification risk applies.
  // ═════════════════════════════════════════════════════════════════════
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Pre-Engagement", Title:"Independent Contractor Agreement signed (NOT employment offer)", AssigneeRole:"Admin", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Pre-Engagement", Title:"Confidentiality Agreement / NDA signed", AssigneeRole:"Admin", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Pre-Engagement", Title:"IP Assignment Agreement signed", AssigneeRole:"Admin", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Pre-Engagement", Title:"Scope of Work document signed", AssigneeRole:"Admin", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Pre-Engagement", Title:"Acknowledge contractor status; no employee benefits or wage protections", AssigneeRole:"Employee", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Pre-Engagement", Title:"Sign ICA, NDA, IP Assignment, SOW", AssigneeRole:"Employee", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Pre-Engagement", Title:"W-9 collected (US) OR W-8BEN (foreign contractor)", AssigneeRole:"Accounting", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Pre-Engagement", Title:"Background check if budget allows (recommended)", AssigneeRole:"Admin", OffsetDays:-21, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Pre-Engagement", Title:"Reference checks completed", AssigneeRole:"Admin", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-7 Days", Title:"Schedule kickoff call (time zone aligned)", AssigneeRole:"Admin", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-7 Days", Title:"Confirm payment method and frequency (typically bi-weekly to contractors)", AssigneeRole:"Accounting", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-7 Days", Title:"Provide BYOD security requirements document", AssigneeRole:"Admin", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-5 Days", Title:"Create M365 guest account OR limited Business Basic license", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-5 Days", Title:"Add to SPECIFIC Teams channels only (not all NewShire channels)", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-5 Days", Title:"Create AppFolio user with SCOPED task-specific permissions", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-5 Days", Title:"Set up time tracking software (Hubstaff / Time Doctor / etc.)", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-3 Days", Title:"Add to SPECIFIC SharePoint folders only (not full site access)", AssigneeRole:"Admin", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-3 Days", Title:"Set up password manager access for shared business credentials only", AssigneeRole:"Admin", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"T-3 Days", Title:"Confirm BYOD requirements understood (MFA, secure browser, password manager)", AssigneeRole:"Employee", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 1 (Kickoff)", Title:"Kickoff call: team intro, workflow, expectations", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 1 (Kickoff)", Title:"Communication protocol: Teams hours, response time, ticketing", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 1 (Kickoff)", Title:"Review scope of work and KPIs", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 1 (Kickoff)", Title:"Walk through tools, SOPs, and templates", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 1 (Kickoff)", Title:"Time tracking software setup verified working", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 1 (Kickoff)", Title:"BYOD security verified: MFA enabled, secure browser, password manager", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 1 (Kickoff)", Title:"Confirm reporting cadence (daily standup or weekly summary)", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Week 1", Title:"Task-specific training", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Week 1", Title:"Shadow + paired work with assigning manager", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Week 1", Title:"Review SOP library", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Week 1", Title:"Begin assigned tasks under close oversight", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Week 1", Title:"Add to VA Activity Tracker (VATrackerRole = VA)", AssigneeRole:"Admin", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Weeks 2-4", Title:"Performance review against scope", AssigneeRole:"Manager", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Weeks 2-4", Title:"Refine workflows", AssigneeRole:"Manager", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 30", Title:"30-day performance review against KPIs", AssigneeRole:"Manager", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 30", Title:"Determine ongoing engagement level", AssigneeRole:"Manager", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 60", Title:"Role/scope adjustment if needed", AssigneeRole:"Manager", OffsetDays:60, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Day 90", Title:"Long-term engagement decision", AssigneeRole:"Manager", OffsetDays:90, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Annual", Title:"1099-NEC issued at year end (US contractors)", AssigneeRole:"Accounting", OffsetDays:365, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Ongoing", Title:"Track hours invoiced vs paid", AssigneeRole:"Accounting", OffsetDays:365, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Annual", Title:"Annual ICA renewal review", AssigneeRole:"Admin", OffsetDays:365, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Compliance", Title:"Misclassification risk review: ensure true contractor relationship", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"Control, integration, exclusivity, equipment — watch DOL/IRS factors." },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Compliance", Title:"Confirm: no I-9 required, no benefits offered, no UI eligibility on separation", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Virtual Assistant", Phase:"Compliance", Title:"If foreign contractor: confirm tax treaty + W-8BEN renewals (every 3 yrs)", AssigneeRole:"Accounting", OffsetDays:0, Required:false, Notes:"" },

  // ═════════════════════════════════════════════════════════════════════
  // ONBOARDING — Maintenance (WG Maintenance LLC)
  // Field tech. Vehicle, tools, uniforms, master keys, trade certs.
  // ═════════════════════════════════════════════════════════════════════
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Pre-Offer", Title:"MVR check completed", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Pre-Offer", Title:"Drug screen completed", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Pre-Offer", Title:"Verify trade certifications: EPA 608 (HVAC refrigerant), state licenses", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Pre-Offer", Title:"Verify lead-safe RRP certification (required for pre-1978 properties)", AssigneeRole:"HR", OffsetDays:-21, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Pre-Offer", Title:"Personal auto insurance verification if using personal vehicle", AssigneeRole:"HR", OffsetDays:-21, Required:false, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-7 Days", Title:"Confirm Day 1 location (shop or office)", AssigneeRole:"Manager", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-5 Days", Title:"Note: M365 license is typically Business Basic (email/Teams focus)", AssigneeRole:"Admin", OffsetDays:-5, Required:false, Notes:"Common-W-2 already creates the account; this is a license-tier note for Maintenance." },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-5 Days", Title:"Create AppFolio user — Mobile / Work Order focused permissions", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-5 Days", Title:"Add to Maintenance Teams channel", AssigneeRole:"Admin", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-5 Days", Title:"Coordinate vehicle assignment OR personal vehicle approval", AssigneeRole:"Manager", OffsetDays:-5, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-3 Days", Title:"Prepare tool inventory list (signed for at issue)", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-3 Days", Title:"Order uniforms (shirts, jacket)", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-3 Days", Title:"Prepare master key set + key control log entry", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-3 Days", Title:"Prepare fuel card or company credit card", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-3 Days", Title:"Prepare photo ID badge (residents verify identity)", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-3 Days", Title:"Vehicle finalized: assignment, registration, insurance card, fuel card", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-1 Day", Title:"Property list and access compiled", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"T-1 Day", Title:"Confirm Day 1 logistics", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"Bring driver's license + trade certifications (EPA 608, lead RRP, state licenses)", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"Sign tool inventory receipt", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"Sign master key control log", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"Receive uniforms, fuel card, photo ID badge, mobile phone if company-provided", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"Vehicle handoff: inspection, registration, insurance card, fuel card", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"Install AppFolio mobile app and verify login", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"Sign vehicle use, drug policy, safety, tool use/return (in addition to common handbook/NDA/etc.)", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"Safety training: ladder, PPE, OSHA basics, lockout/tagout, electrical safety", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"Conducted by Lead Tech." },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"Review property list and access", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"AppFolio work order walkthrough", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"Conducted by Lead Tech." },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 1", Title:"After-hours / on-call rotation explained", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Week 1", Title:"Ride-along with senior tech (multiple properties)", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"Lead Tech." },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Week 1", Title:"Lead-safe work practices training if working pre-1978 properties", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Week 1", Title:"Vendor introductions (parts suppliers, specialty trades)", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Week 1", Title:"Vehicle maintenance and inspection schedule reviewed", AssigneeRole:"Manager", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Week 1", Title:"Read SOPs for role; save to phone/tablet", AssigneeRole:"Employee", OffsetDays:3, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Weeks 2-4", Title:"Begin solo work orders (low complexity)", AssigneeRole:"Employee", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Weeks 2-4", Title:"Quality check on completed WOs by Lead Tech", AssigneeRole:"Manager", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Weeks 2-4", Title:"Complete required NewShire University modules including fair housing", AssigneeRole:"Employee", OffsetDays:14, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 30", Title:"Maintenance-specific: WO quality, time per WO, comeback rate review", AssigneeRole:"Manager", OffsetDays:30, Required:true, Notes:"Supplements the common 30-day check-in." },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Day 60", Title:"Vehicle and tool inventory check", AssigneeRole:"Manager", OffsetDays:60, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Annual", Title:"Trade certification renewal tracking (EPA 608, lead RRP)", AssigneeRole:"Admin", OffsetDays:365, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Annual", Title:"Annual MVR re-pull for drivers", AssigneeRole:"HR", OffsetDays:365, Required:true, Notes:"" },
  { JourneyType:"Onboarding", TemplateGroup:"Maintenance", Phase:"Annual", Title:"Tool inventory annual audit", AssigneeRole:"Manager", OffsetDays:365, Required:true, Notes:"" },

  // ═════════════════════════════════════════════════════════════════════
  // OFFBOARDING — Voluntary Resignation
  // Resignation with notice. Role-specific equipment recovery on Last Day.
  // ═════════════════════════════════════════════════════════════════════
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Notice Received", Title:"Submit resignation in writing with last day worked", AssigneeRole:"Employee", OffsetDays:-14, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Notice Received", Title:"Acknowledge resignation in writing", AssigneeRole:"HR", OffsetDays:-14, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Notice Received", Title:"Notify Cara, John, Brandy, and Accounting", AssigneeRole:"Manager", OffsetDays:-14, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Notice Received", Title:"Set last working day on calendar", AssigneeRole:"Manager", OffsetDays:-14, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-14 to T-10", Title:"Define knowledge transfer scope: active files, recurring tasks, contacts", AssigneeRole:"Manager", OffsetDays:-12, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-14 to T-10", Title:"Identify successor or coverage plan", AssigneeRole:"Manager", OffsetDays:-12, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-14 to T-10", Title:"Meet with manager to align on transition plan", AssigneeRole:"Employee", OffsetDays:-12, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-14 to T-10", Title:"Build transition document (active items, recurring tasks, contacts, deadlines)", AssigneeRole:"Employee", OffsetDays:-12, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-14 to T-10", Title:"Schedule communications to owners / vendors / residents", AssigneeRole:"Admin", OffsetDays:-12, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-7", Title:"Begin documented handoff: AppFolio notes, open WOs, leasing prospects, owner comms", AssigneeRole:"Employee", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-7", Title:"Reassign open AppFolio work orders, leasing prospects, tasks", AssigneeRole:"Employee", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-7", Title:"Document business logins, vendor portals, account ownership", AssigneeRole:"Employee", OffsetDays:-7, Required:true, Notes:"Business credentials only." },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-7", Title:"Schedule exit interview", AssigneeRole:"HR", OffsetDays:-7, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-3", Title:"Save and turn over work product stored locally — nothing business stays on personal devices", AssigneeRole:"Employee", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-3", Title:"Confirm forwarding address for W-2 and benefits correspondence", AssigneeRole:"Employee", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-3", Title:"Confirm equipment return list with employee", AssigneeRole:"Manager", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T-3", Title:"Schedule final access cutoff time", AssigneeRole:"Admin", OffsetDays:-3, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"Submit final timesheet", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"Submit outstanding expense reports", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"Conduct exit interview", AssigneeRole:"HR", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"Sign return-of-property acknowledgment", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"Sign NDA reaffirmation reminder", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"CORPORATE: Collect laptop, phone, building key, badge, credit cards", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"If departing employee was Corporate." },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"ON-SITE: Collect property keys, master keys, on-site office key, badge, vehicle, employee unit move-out", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"If departing employee was On-Site." },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"OFF-SITE: Collect master keys for ALL assigned properties, vehicle, laptop, phone, fuel/credit cards", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"If departing employee was Off-Site." },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"VA: No physical property to return — confirm access termination instead", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"If departing contractor was a VA." },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"MAINTENANCE: Collect tools (inventory), vehicle (inspection), master keys, uniforms, fuel card, photo ID", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"If departing employee was Maintenance." },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"Issue SC UI separation notice (W-2 only; not for VA 1099)", AssigneeRole:"HR", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"Disable M365 sign-in, force password reset, revoke MFA tokens", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"Remove from AppFolio, SharePoint, Teams, all vendor portals", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"Convert mailbox to shared, set auto-responder, forward 30 days to successor", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Last Day", Title:"VA only: Immediate access termination, no shared mailbox conversion", AssigneeRole:"Admin", OffsetDays:0, Required:false, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T+1", Title:"Notify residents, owners, vendors of new point of contact", AssigneeRole:"Manager", OffsetDays:1, Required:true, Notes:"Successor." },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T+1", Title:"Verify all accounts disabled", AssigneeRole:"Admin", OffsetDays:1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T+1", Title:"ON-SITE/OFF-SITE/MAINTENANCE: Assess rekey need for properties with master access", AssigneeRole:"Admin", OffsetDays:1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Next Payday", Title:"Issue final paycheck (SC § 41-10-50: next regular payday, max 30 days)", AssigneeRole:"Accounting", OffsetDays:7, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Next Payday", Title:"Pay out unused PTO if handbook requires (review policy)", AssigneeRole:"Accounting", OffsetDays:7, Required:true, Notes:"SC does not require payout unless company policy says so." },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T+30", Title:"Archive personnel file", AssigneeRole:"Admin", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T+30", Title:"Export M365 data and store per retention policy", AssigneeRole:"Admin", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"T+30", Title:"Delete shared mailbox", AssigneeRole:"Admin", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Retention", Title:"I-9 retained: 3 yrs from hire OR 1 yr from termination, whichever is later", AssigneeRole:"HR", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Retention", Title:"Personnel file retained per company policy and SC requirements", AssigneeRole:"HR", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Retention", Title:"Payroll records retained 4 years (IRS)", AssigneeRole:"Accounting", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Voluntary Resignation", Phase:"Post-Separation", Title:"NDA and non-solicit obligations remain in effect", AssigneeRole:"Employee", OffsetDays:1, Required:true, Notes:"No resident, owner, vendor, or financial data leaves with employee." },

  // ═════════════════════════════════════════════════════════════════════
  // OFFBOARDING — Involuntary Termination
  // Two-person rule. Access cuts at end of meeting (or before, if security risk).
  // ═════════════════════════════════════════════════════════════════════
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Pre-Meeting", Title:"Coordinate with relevant team members (Brandy + HR)", AssigneeRole:"HR", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Pre-Meeting", Title:"Prepare termination letter", AssigneeRole:"Admin", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Pre-Meeting", Title:"Calculate final pay: regular wages + accrued PTO per policy", AssigneeRole:"Accounting", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Pre-Meeting", Title:"Stage account disable to execute at meeting end", AssigneeRole:"Admin", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Pre-Meeting", Title:"Prepare role-specific equipment/property recovery list", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Pre-Meeting", Title:"Prepare SC UI separation notice", AssigneeRole:"HR", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Pre-Meeting", Title:"Prepare NDA reaffirmation reminder", AssigneeRole:"HR", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Pre-Meeting", Title:"If benefits-enrolled: prepare COBRA notice", AssigneeRole:"HR", OffsetDays:-1, Required:false, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Pre-Meeting", Title:"Confirm two-person meeting plan (manager + HR witness)", AssigneeRole:"Manager", OffsetDays:-1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Termination Meeting", Title:"Conduct termination with witness present (ALWAYS 2 people)", AssigneeRole:"HR", OffsetDays:0, Required:true, Notes:"Two-person rule." },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Termination Meeting", Title:"Communicate final pay info, last day, benefits status", AssigneeRole:"HR", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Termination Meeting", Title:"CORPORATE: Collect building keys, badge, laptop, phone, credit/fuel cards", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Termination Meeting", Title:"ON-SITE: Collect property keys, master keys, on-site office key, badge, mobile, employee unit move-out", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Termination Meeting", Title:"OFF-SITE: Collect master keys for ALL properties, vehicle keys, laptop, phone, credit/fuel cards", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Termination Meeting", Title:"VA: Terminate access immediately (often delivered by email); no physical recovery", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Termination Meeting", Title:"MAINTENANCE: Collect tools (inventory), vehicle keys (inspection), master keys for ALL properties, uniforms, photo ID, fuel card", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Termination Meeting", Title:"Provide forwarding address for W-2 and COBRA correspondence", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Termination Meeting", Title:"Acknowledge receipt of separation paperwork", AssigneeRole:"Employee", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"End of Meeting", Title:"Disable M365 sign-in", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"Execute at end of meeting (or before if security risk)." },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"End of Meeting", Title:"Remove from AppFolio, SharePoint, Teams", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"End of Meeting", Title:"Revoke MFA tokens, force password reset", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"End of Meeting", Title:"Remote-wipe any company device not recovered", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Same Day", Title:"Escort to retrieve personal items", AssigneeRole:"Manager", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Same Day", Title:"Process final paycheck (SC § 41-10-50: next regular payday, max 30 days)", AssigneeRole:"Accounting", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Same Day", Title:"Issue SC UI separation notice", AssigneeRole:"HR", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Same Day", Title:"Set mailbox auto-responder and forward to successor", AssigneeRole:"Admin", OffsetDays:0, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"T+1", Title:"Notify residents, owners, vendors of new point of contact", AssigneeRole:"Manager", OffsetDays:1, Required:true, Notes:"Successor." },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"T+1", Title:"Verify all accounts disabled and remote wipe complete", AssigneeRole:"Admin", OffsetDays:1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"T+5", Title:"ON-SITE/OFF-SITE/MAINTENANCE: Make rekey decision and execute if required", AssigneeRole:"Admin", OffsetDays:5, Required:true, Notes:"Properties with master access." },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"T+14", Title:"Mail COBRA notice if benefits-enrolled", AssigneeRole:"HR", OffsetDays:14, Required:false, Notes:"Employer: within 14 days; plan admin: 60 days from qualifying event." },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"T+30", Title:"Archive personnel file with termination documentation", AssigneeRole:"Admin", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"T+30", Title:"Export M365 data per retention policy", AssigneeRole:"Admin", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"T+30", Title:"Delete shared mailbox", AssigneeRole:"Admin", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Retention", Title:"I-9 retained: 3 yrs from hire OR 1 yr from termination, whichever is later", AssigneeRole:"HR", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Retention", Title:"Termination documentation retained per policy (recommended 7 years)", AssigneeRole:"HR", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Retention", Title:"Payroll records retained 4 years (IRS)", AssigneeRole:"Accounting", OffsetDays:30, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Post-Separation", Title:"NDA, confidentiality, non-solicit obligations remain in effect", AssigneeRole:"Employee", OffsetDays:1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Post-Separation", Title:"No resident, owner, vendor, or financial data may be retained or used", AssigneeRole:"Employee", OffsetDays:1, Required:true, Notes:"" },
  { JourneyType:"Offboarding", TemplateGroup:"Involuntary Termination", Phase:"Post-Separation", Title:"Return any remote-location property within 5 business days", AssigneeRole:"Employee", OffsetDays:5, Required:true, Notes:"" },
];

// ============================================================
// DATA CONTEXT
// ============================================================
const DataCtx = createContext(null);
const useData = () => useContext(DataCtx);

// ============================================================
// LOADER
// ============================================================
// ============================================================
// DEFAULT APPS REGISTRY — fallback if ELC_Apps list is empty
// ============================================================
const DEFAULT_APPS = [
  { Title: "VA Tracker",          AppKey: "vatracker",  ColumnName: "VATrackerRole",  Roles: ["VA","Manager","Regional","Admin"],        IconLetter: "V", Color: "#3A6577", Active: true, OrderIdx: 10, OnboardingDefault: "None",     Description: "Productivity tracker for virtual assistants" },
  { Title: "PM Hub",              AppKey: "pmhub",      ColumnName: "PMHubRole",       Roles: ["PM","Regional","Owner","Admin"],          IconLetter: "P", Color: "#1C3740", Active: true, OrderIdx: 20, OnboardingDefault: "None",     Description: "Property manager activity & help queue" },
  { Title: "NewShire University", AppKey: "university", ColumnName: "UniversityRole",  Roles: ["Employee","Manager","Admin"],             IconLetter: "U", Color: "#CDA04B", Active: true, OrderIdx: 30, OnboardingDefault: "Employee", Description: "Training, courses, compliance learning paths" },
  { Title: "Expense Manager",     AppKey: "expense",    ColumnName: "ExpenseRole",     Roles: ["Employee","Accounting","Admin"],          IconLetter: "E", Color: "#2D8A5A", Active: true, OrderIdx: 40, OnboardingDefault: "Employee", Description: "Mileage & expense reimbursement" },
  { Title: "CAHP Compliance Hub", AppKey: "cahp",       ColumnName: "CAHPRole",        Roles: ["Viewer","Editor","Admin"],                IconLetter: "C", Color: "#4A78B0", Active: true, OrderIdx: 50, OnboardingDefault: "None",     Description: "Affordable housing compliance tracker" },
  { Title: "AppFolio Dashboard",  AppKey: "appfolio",   ColumnName: "AppFolioRole",    Roles: ["Viewer","Editor","Admin"],                IconLetter: "A", Color: "#B8922E", Active: true, OrderIdx: 60, OnboardingDefault: "None",     Description: "NewShire AppFolio analytics dashboard" },
  { Title: "ShowMojo Sync",       AppKey: "showmojo",   ColumnName: "ShowMojoRole",    Roles: ["Viewer","Editor","Admin"],                IconLetter: "S", Color: "#5B3FA8", Active: true, OrderIdx: 70, OnboardingDefault: "None",     Description: "ShowMojo listing/showings sync" },
  { Title: "Renewal Manager",     AppKey: "renewal",    ColumnName: "RenewalRole",     Roles: ["Viewer","Editor","Admin"],                IconLetter: "R", Color: "#C44B3B", Active: true, OrderIdx: 80, OnboardingDefault: "None",     Description: "Lease renewal workflow & tracking" },
  { Title: "Employee Lifecycle",  AppKey: "elc",        ColumnName: "ELCRole",         Roles: ["Employee","Manager","HR","IT","Admin"],   IconLetter: "L", Color: "#CDA04B", Active: true, OrderIdx: 90, OnboardingDefault: "Employee", Description: "Onboarding / offboarding / HR file system" },
];
function normalizeApp(raw) {
  let roles = raw.Roles;
  if (typeof roles === "string") { try { roles = JSON.parse(roles); } catch { roles = roles.split(",").map(s => s.trim()).filter(Boolean); } }
  if (!Array.isArray(roles)) roles = [];
  return { ...raw, Roles: roles };
}

async function loadAll(token) {
  const [emp, jrn, tpl, jt, cfg, apps, notes, audit, reviews, pay, emailTpl, luCourses, luAssign, luComp] = await Promise.all([
    safeGet(token, "Employees",      `${lUrl(CONFIG.lists.employees)}?expand=fields&$top=500`),
    safeGet(token, "Journeys",       `${lUrl(CONFIG.lists.journeys)}?expand=fields&$top=500`),
    safeGet(token, "TemplateTasks",  `${lUrl(CONFIG.lists.templateTasks)}?expand=fields&$top=500`),
    safeGet(token, "JourneyTasks",   `${lUrl(CONFIG.lists.journeyTasks)}?expand=fields&$top=2000`),
    safeGet(token, "Config",         `${lUrl(CONFIG.lists.config)}?expand=fields&$top=5`),
    safeGet(token, "Apps",           `${lUrl(CONFIG.lists.apps)}?expand=fields&$top=100`),
    safeGet(token, "Notes",          `${lUrl(CONFIG.lists.notes)}?expand=fields&$top=2000`),
    safeGet(token, "Audit",          `${lUrl(CONFIG.lists.audit)}?expand=fields&$top=2000`),
    safeGet(token, "Reviews",        `${lUrl(CONFIG.lists.reviews)}?expand=fields&$top=2000`),
    safeGet(token, "PayChanges",     `${lUrl(CONFIG.lists.payChanges)}?expand=fields&$top=2000`),
    safeGet(token, "EmailTemplates", `${lUrl(CONFIG.lists.emailTemplates)}?expand=fields&$top=200`),
    safeGet(token, "LU Courses",     `${lUrl(CONFIG.lists.luCourses)}?expand=fields&$top=500`),
    safeGet(token, "LU Assignments", `${lUrl(CONFIG.lists.luAssignments)}?expand=fields&$top=4000`),
    safeGet(token, "LU Completions", `${lUrl(CONFIG.lists.luCompletions)}?expand=fields&$top=4000`),
  ]);
  const employees = emp.map(e => ({ id: e.id, ...e.fields }));
  const journeys = jrn.map(j => ({ id: j.id, ...j.fields }));
  const templates = tpl.map(t => ({ id: t.id, ...t.fields }));
  const journeyTasks = jt.map(t => ({ id: t.id, ...t.fields }));
  let appsList = apps.map(a => normalizeApp({ id: a.id, ...a.fields })).filter(a => a.Active !== false).sort((a, b) => (a.OrderIdx || 0) - (b.OrderIdx || 0));
  if (appsList.length === 0) appsList = DEFAULT_APPS.map((a, i) => ({ ...a, id: `default-${i}` }));
  const notesList = notes.map(n => ({ id: n.id, ...n.fields }));
  const auditList = audit.map(a => ({ id: a.id, ...a.fields }));
  const reviewsList = reviews.map(r => ({ id: r.id, ...r.fields }));
  const payChangesList = pay.map(p => ({ id: p.id, ...p.fields }));
  const emailTemplatesList = emailTpl.map(t => ({ id: t.id, ...t.fields })).filter(t => t.Active !== false).sort((a, b) => (a.Category || "").localeCompare(b.Category || "") || (a.Title || "").localeCompare(b.Title || ""));
  // NewShire University — read-only cross-app data for training compliance.
  // Lookup-shape fields surface as `<base>LookupId` via Graph (per memory note).
  const luCoursesList = luCourses.map(c => ({ id: String(c.id), ...c.fields }));
  const luAssignmentsList = luAssign.map(a => ({
    id: String(a.id),
    EmployeeEmail: (a.fields?.AssignEmployeeEmail || "").toLowerCase(),
    CourseId: String(a.fields?.AssignCourseIDLookupId || a.fields?.AssignCourseID || ""),
    DueDate: a.fields?.AssignDueDate ? String(a.fields.AssignDueDate).split("T")[0] : "",
    AssignedDate: a.fields?.AssignDate || a.fields?.Created || "",
    Status: a.fields?.AssignStatus || "",
  }));
  const luCompletionsList = luComp.map(c => ({
    id: String(c.id),
    EmployeeEmail: (c.fields?.EmployeeEmail || "").toLowerCase(),
    CourseId: String(c.fields?.CompCourseIDLookupId || c.fields?.CompCourseID || ""),
    CompletedDate: c.fields?.CompletedDate ? String(c.fields.CompletedDate).split("T")[0] : "",
    Score: Number(c.fields?.Score || 0),
    Status: c.fields?.CompStatus || "",
  }));
  const config = cfg.length > 0 ? (() => { try { return JSON.parse(cfg[0].fields.ConfigJSON || "{}"); } catch { return {}; } })() : {};
  return {
    employees, journeys, templates, journeyTasks, config,
    apps: appsList, notes: notesList, audit: auditList,
    reviews: reviewsList, payChanges: payChangesList, emailTemplates: emailTemplatesList,
    luCourses: luCoursesList, luAssignments: luAssignmentsList, luCompletions: luCompletionsList,
  };
}

// ============================================================
// SMALL COMPONENTS
// ============================================================
function Avatar({ name, size = 30, color }) {
  const palette = [
    { bg: C.t0, fg: C.t7 }, { bg: C.g1, fg: C.g6 }, { bg: C.erb, fg: C.er },
    { bg: C.pub, fg: C.pu }, { bg: C.okb, fg: C.ok }, { bg: C.infb, fg: C.inf },
  ];
  const idx = name ? (name.charCodeAt(0) + (name.charCodeAt(1) || 0)) % palette.length : 0;
  const c = color || palette[idx];
  return <div style={{ width: size, height: size, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: size * 0.34, background: c.bg, color: c.fg, flexShrink: 0 }}>{initials(name)}</div>;
}
function Badge({ type = "neutral", children, dot = true }) {
  const m = { ok: { c: C.ok, b: C.okb }, er: { c: C.er, b: C.erb }, wn: { c: C.wn, b: C.wnb }, inf: { c: C.inf, b: C.infb }, pu: { c: C.pu, b: C.pub }, neutral: { c: C.b4, b: C.b1 } }[type] || { c: C.b4, b: C.b1 };
  return <span style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "2px 9px", fontSize: 11, fontWeight: 600, borderRadius: 99, color: m.c, background: m.b, whiteSpace: "nowrap" }}>{dot && <span style={{ width: 5, height: 5, borderRadius: "50%", background: "currentColor" }} />}{children}</span>;
}
function ProgressBar({ value, max = 100, color = C.g5, height = 6 }) {
  const pct = Math.min(100, Math.round((value / Math.max(1, max)) * 100));
  return (
    <div style={{ width: "100%", background: C.b1, borderRadius: 99, height, overflow: "hidden" }}>
      <div style={{ width: `${pct}%`, height: "100%", background: color, borderRadius: 99, transition: "width .25s" }} />
    </div>
  );
}
function Modal({ title, onClose, children, footer, width }) {
  useEffect(() => { const h = e => e.key === "Escape" && onClose(); window.addEventListener("keydown", h); return () => window.removeEventListener("keydown", h); }, [onClose]);
  return (
    <div style={S.modalBg} onClick={onClose}>
      <div style={{ ...S.modal, maxWidth: width || S.modal.maxWidth }} onClick={e => e.stopPropagation()}>
        <div style={S.modalH}>
          <div style={{ fontSize: 14, fontWeight: 700 }}>{title}</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, color: C.b4, cursor: "pointer", lineHeight: 1 }}>×</button>
        </div>
        <div style={S.modalB}>{children}</div>
        {footer && <div style={S.modalF}>{footer}</div>}
      </div>
    </div>
  );
}
function Empty({ title, sub }) { return <div style={{ textAlign: "center", padding: "40px 16px", color: C.b4 }}><div style={{ fontSize: 14, fontWeight: 600, color: C.t7, marginBottom: 4 }}>{title}</div><div style={{ fontSize: 12 }}>{sub}</div></div>; }

// ============================================================
// ROLE DETECTION
// ============================================================
// Role precedence:
//   1. CONFIG.adminEmails  (super-admin failsafe — can't be locked out)
//   2. ELCRole column on Employees (set via App Permissions UI) ← canonical
//   3. Legacy AccessLevel / JobTitle heuristics (for first-run before ELCRole is set)
//
// VAs (1099 contractors) are identified by Employees.PersonaType = "Virtual Assistant",
// not by a separate ELCRole. In this app they're treated the same as a regular Employee:
// they only see their own My Tasks and their own onboarding/offboarding journey.
// VA-specific behavior (template-group selection, 1099 paperwork) is driven by PersonaType.
const VALID_ELC_ROLES = new Set(["Admin", "HR", "IT", "Manager", "Employee"]);
// Roles that can be freely combined on one person (small-company reality — one person
// may wear several hats). Employee is the baseline; VA is tracking-only and exclusive.
const COMBINABLE_ROLES = ["Admin", "HR", "IT", "Manager"];
const ROLE_RANK = { Admin: 5, HR: 4, IT: 3, Manager: 2, Employee: 1, VA: 0 };
// ELCRole can now hold several roles as a delimited list (e.g. "Admin; HR; IT").
// parseRoles normalizes a raw cell (string, delimited string, or array) into a clean array.
function parseRoles(raw) {
  if (raw == null) return [];
  const arr = Array.isArray(raw) ? raw : String(raw).split(/[;,]/);
  return arr.map(s => String(s).trim()).filter(s => s && s !== "None" && VALID_ELC_ROLES.has(s));
}
// Highest-ranked role — used for single-role displays and the VA access gate.
function primaryRole(roles) {
  return (roles || []).slice().sort((a, b) => (ROLE_RANK[b] ?? -1) - (ROLE_RANK[a] ?? -1))[0] || "Employee";
}
// Full set of roles a person holds (array). Super-admins are always Admin.
function detectRoles(emp, email) {
  if (CONFIG.adminEmails.includes((email || "").toLowerCase())) return ["Admin"];
  if (!emp) return ["Employee"];
  const explicit = parseRoles(emp.ELCRole);
  if (explicit.length) return explicit;
  // Legacy fallback heuristics (only used before ELCRole is set) → a single role.
  const al = (emp.AccessLevel || "").toLowerCase();
  if (al.includes("admin") || al.includes("owner")) return ["Admin"];
  const title = (emp.JobTitle || "").toLowerCase();
  if (title.includes("hr") || al.includes("hr")) return ["HR"];
  if (title.includes("it") || al.includes("it")) return ["IT"];
  if (title.includes("manager") || title.includes("regional") || title.includes("owner")) return ["Manager"];
  return ["Employee"];
}
// Primary (single) role — kept for backward-compatible call sites and badges.
function detectRole(emp, email) { return primaryRole(detectRoles(emp, email)); }
// A task belongs to me if it's assigned to my email, OR (when it has no specific
// assignee email) it's a role-queue task whose AssigneeRole is one of my roles.
// This is why selecting HR/IT/Admin surfaces those onboarding tasks in My Tasks.
function taskIsMine(t, myEmail, hasRole) {
  const email = (t.AssigneeEmail || "").toLowerCase();
  if (email) return email === myEmail;
  return !!t.AssigneeRole && hasRole(t.AssigneeRole);
}

// ============================================================
// LOGIN SCREEN
// ============================================================
function LoginScreen({ onLogin, err }) {
  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: `linear-gradient(135deg, ${C.hdr} 0%, ${C.t6} 100%)`, padding: 20 }}>
      <div style={{ background: C.wh, borderRadius: 12, padding: 36, maxWidth: 420, width: "100%", boxShadow: "0 20px 60px rgba(0,0,0,.25)", textAlign: "center" }}>
        <div style={{ width: 64, height: 64, background: C.g5, borderRadius: 12, margin: "0 auto 18px", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 32, fontWeight: 700, color: C.t7, fontFamily: "Georgia,serif" }}>L</div>
        <div style={{ fontSize: 20, fontWeight: 700, color: C.t7, marginBottom: 6, letterSpacing: ".06em" }}>EMPLOYEE LIFECYCLE</div>
        <div style={{ fontSize: 13, color: C.b4, marginBottom: 24 }}>NewShire onboarding & offboarding</div>
        <button onClick={onLogin} style={{ ...S.btn(C.hdr), width: "100%", padding: "12px 14px", fontSize: 14 }}>Sign in with Microsoft</button>
        {err && <div style={{ marginTop: 14, padding: 10, background: C.erb, color: C.er, borderRadius: 4, fontSize: 12 }}>{err}</div>}
        <div style={{ marginTop: 20, fontSize: 11, color: C.b3 }}>v{CONFIG.version}</div>
      </div>
    </div>
  );
}

// ============================================================
// HEADER + TAB BAR
// ============================================================
function Header({ user, role, roles, onLogout }) {
  return (
    <div style={S.hdr}>
      <div style={S.hdrL}>
        <div style={S.logo}>L</div>
        <div>
          <div style={S.hdrTitle}>EMPLOYEE LIFECYCLE</div>
          <div style={S.hdrSub}>NewShire · Onboarding & Offboarding</div>
        </div>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ textAlign: "right" }}>
          <div style={S.hdrUser}>{user?.name || user?.email || "—"}</div>
          <div style={S.hdrRole}>{roles && roles.length ? roles.join(" · ") : role}</div>
        </div>
        <button onClick={onLogout} style={{ ...S.btnO(C.t1, "rgba(255,255,255,.18)"), background: "transparent", color: C.t1 }}>Sign out</button>
      </div>
    </div>
  );
}
function TabBar({ tabs, active, onChange }) {
  return (
    <div style={S.tabBar}>
      {tabs.map(t => (
        <button key={t.key} style={S.tab(active === t.key)} onClick={() => onChange(t.key)}>
          {t.label}{typeof t.count === "number" && t.count > 0 && (
            <span style={{ marginLeft: 6, background: active === t.key ? C.g1 : C.b1, color: active === t.key ? C.g7 : C.b4, fontSize: 10, fontWeight: 700, padding: "1px 7px", borderRadius: 99 }}>{t.count}</span>
          )}
        </button>
      ))}
    </div>
  );
}

// ============================================================
// JOURNEY HELPERS
// ============================================================
function journeyProgress(journey, journeyTasks) {
  const tasks = journeyTasks.filter(t => String(t.JourneyId) === String(journey.id));
  const done = tasks.filter(t => t.Status === "Done").length;
  return { done, total: tasks.length, pct: tasks.length ? Math.round((done / tasks.length) * 100) : 0, tasks };
}
function journeyAnchorDate(j) { return j.JourneyType === "Offboarding" ? j.EndDate : j.StartDate; }

// Manager-line check: returns true if `managerEmail` is any tier (1, 2, or 3) up the
// chain of command for this employee — i.e., manager, grandmanager, or great-grandmanager.
function isInManagerLine(emp, managerEmail) {
  if (!emp || !managerEmail) return false;
  const me = (managerEmail || "").toLowerCase();
  return [emp.ManagerEmail, emp.Level2ManagerEmail, emp.Level3ManagerEmail]
    .some(m => (m || "").toLowerCase() === me);
}

// Visibility rule for a single journey card. Same rule used everywhere the
// Onboarding/Offboarding lists are rendered or counted.
//   Admin / HR             — see all (incl. unshared offboarding)
//   Manager                — see anyone in their line of succession (1/2/3 tiers up);
//                            for offboarding, ALSO requires SharedWithEmployee
//   Everyone else          — see only their own; offboarding additionally requires
//                            SharedWithEmployee
function canSeeJourney(journey, hasRole, currentEmail, employees) {
  if (hasRole("Admin", "HR")) return true;
  const isOff = journey.JourneyType === "Offboarding";
  if (isOff && !journey.SharedWithEmployee) return false;
  const empRow = employees.find(e => (e.Email || "").toLowerCase() === (journey.EmployeeEmail || "").toLowerCase());
  if (hasRole("Manager") && isInManagerLine(empRow, currentEmail)) return true;
  return (journey.EmployeeEmail || "").toLowerCase() === (currentEmail || "").toLowerCase();
}

// ============================================================
// JOURNEYS LIST (Onboarding / Offboarding)
// ============================================================
function JourneysTab({ type }) {
  const { state, actions, role, hasRole, currentEmail } = useData();
  const [openId, setOpenId] = useState(null);
  const [startOpen, setStartOpen] = useState(false);
  const canStart = hasRole("Admin", "HR");

  const list = state.journeys
    .filter(j => j.JourneyType === type)
    .filter(j => canSeeJourney(j, hasRole, currentEmail, state.employees))
    .map(j => ({ ...j, _emp: state.employees.find(e => (e.Email || "").toLowerCase() === (j.EmployeeEmail || "").toLowerCase()), _p: journeyProgress(j, state.journeyTasks) }))
    // Oldest → newest by anchor date (StartDate for onboarding, EndDate for offboarding).
    // Journeys with no date sort to the end.
    .sort((a, b) => (journeyAnchorDate(a) || "9999-12-31").localeCompare(journeyAnchorDate(b) || "9999-12-31"));

  const active = list.filter(j => j.Status !== "Complete" && j.Status !== "Cancelled");
  const completed = list.filter(j => j.Status === "Complete" || j.Status === "Cancelled");

  const k = {
    total: active.length,
    overdueTasks: active.reduce((s, j) => s + j._p.tasks.filter(t => classifyDue(t.DueDate, t.Status) === "overdue").length, 0),
    pctAvg: active.length ? Math.round(active.reduce((s, j) => s + j._p.pct, 0) / active.length) : 0,
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
        <div style={S.kpi}><div style={S.kpiL}>Active</div><div style={S.kpiV}>{k.total}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Avg Progress</div><div style={S.kpiV}>{k.pctAvg}%</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Overdue Tasks</div><div style={{ ...S.kpiV, color: k.overdueTasks > 0 ? C.er : C.t7 }}>{k.overdueTasks}</div></div>
      </div>

      <div style={S.card}>
        <div style={S.cardT}>
          <span>Active {type}</span>
          {canStart && <button style={S.btn(C.hdr)} onClick={() => setStartOpen(true)}>+ Start {type}</button>}
        </div>
        {active.length === 0 ? <Empty title={`No active ${type.toLowerCase()} journeys`} sub={canStart ? `Click "Start ${type}" to begin.` : "Check back later."} /> : (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))", gap: 12 }}>
            {active.map(j => <JourneyCard key={j.id} j={j} onClick={() => setOpenId(j.id)} />)}
          </div>
        )}
      </div>

      {completed.length > 0 && (
        <div style={S.card}>
          <div style={S.cardT}>Completed {type} ({completed.length})</div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th style={S.th}>Employee</th><th style={S.th}>Anchor Date</th><th style={S.th}>Status</th><th style={S.th}>Tasks</th><th style={S.th}></th></tr></thead>
              <tbody>
                {completed.map(j => (
                  <tr key={j.id}>
                    <td style={S.td}><div style={{ display: "flex", alignItems: "center", gap: 8 }}><Avatar name={j.EmployeeName || j._emp?.Title} size={26} /><div><div style={{ fontWeight: 600, color: C.t7 }}>{j.EmployeeName || j._emp?.Title || j.EmployeeEmail}</div><div style={{ fontSize: 11, color: C.b4 }}>{j.JobTitle || j._emp?.JobTitle || ""}</div></div></div></td>
                    <td style={S.td}>{fmtDate(journeyAnchorDate(j))}</td>
                    <td style={S.td}><Badge type={j.Status === "Complete" ? "ok" : "neutral"}>{j.Status}</Badge></td>
                    <td style={S.td}>{j._p.done}/{j._p.total}</td>
                    <td style={S.td}><button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => setOpenId(j.id)}>Open</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {openId && <JourneyDetailModal journeyId={openId} onClose={() => setOpenId(null)} />}
      {startOpen && <StartJourneyModal type={type} onClose={() => setStartOpen(false)} />}
    </div>
  );
}

function JourneyCard({ j, onClick }) {
  const overdue = j._p.tasks.filter(t => classifyDue(t.DueDate, t.Status) === "overdue").length;
  const anchor = journeyAnchorDate(j);
  const dayLabel = j.JourneyType === "Offboarding" ? (daysFromNow(anchor) >= 0 ? `${daysFromNow(anchor)} days to last day` : `${-daysFromNow(anchor)} days past last day`) : (daysFromNow(anchor) > 0 ? `Starts in ${daysFromNow(anchor)} days` : `Day ${-daysFromNow(anchor)}`);
  return (
    <div onClick={onClick} style={{ background: C.wh, border: `1px solid ${C.b2}`, borderRadius: 6, padding: 14, cursor: "pointer", transition: ".15s", boxShadow: "0 1px 2px rgba(28,55,64,.05)" }}
      onMouseEnter={e => { e.currentTarget.style.borderColor = C.g5; e.currentTarget.style.boxShadow = "0 2px 8px rgba(205,160,75,.18)"; }}
      onMouseLeave={e => { e.currentTarget.style.borderColor = C.b2; e.currentTarget.style.boxShadow = "0 1px 2px rgba(28,55,64,.05)"; }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
        <Avatar name={j.EmployeeName || j._emp?.Title} size={38} />
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontWeight: 700, color: C.t7, fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{j.EmployeeName || j._emp?.Title || j.EmployeeEmail}</div>
          <div style={{ fontSize: 11, color: C.b4 }}>{j.JobTitle || j._emp?.JobTitle || ""}</div>
          {j.TemplateGroup && <div style={{ marginTop: 4 }}><Badge type="pu">{j.TemplateGroup}</Badge></div>}
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <div style={{ fontSize: 11, color: C.b4 }}>{dayLabel}</div>
        <Badge type={overdue > 0 ? "er" : j._p.pct === 100 ? "ok" : "inf"}>{j._p.done}/{j._p.total}</Badge>
      </div>
      <ProgressBar value={j._p.pct} color={overdue > 0 ? C.er : j._p.pct === 100 ? C.ok : C.g5} />
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8, fontSize: 11, color: C.b4 }}>
        <span>{fmtDate(anchor)}</span>
        <span>{j._p.pct}%</span>
      </div>
      {overdue > 0 && <div style={{ marginTop: 8, fontSize: 11, color: C.er, fontWeight: 600 }}>⚠ {overdue} overdue task{overdue > 1 ? "s" : ""}</div>}
    </div>
  );
}

// ============================================================
// JOURNEY DETAIL — task list w/ inline edit
// ============================================================
function JourneyDetailModal({ journeyId, onClose }) {
  const { state, actions, role, hasRole, currentEmail } = useData();
  const j = state.journeys.find(x => String(x.id) === String(journeyId));
  if (!j) return null;
  const emp = state.employees.find(e => (e.Email || "").toLowerCase() === (j.EmployeeEmail || "").toLowerCase());
  const tasks = state.journeyTasks.filter(t => String(t.JourneyId) === String(j.id))
    .sort((a, b) => (a.OffsetDays || 0) - (b.OffsetDays || 0) || (a.DueDate || "").localeCompare(b.DueDate || "") || (a.OrderIdx || 0) - (b.OrderIdx || 0));
  const phases = uniq(tasks.map(t => t.Phase || "General"));
  const canEdit = hasRole("Admin", "HR") || (emp && (emp.ManagerEmail || "").toLowerCase() === (currentEmail || "").toLowerCase());

  const [saving, setSaving] = useState(false);
  const setTask = async (task, patch) => {
    setSaving(true); try { await actions.updateTask(task.id, patch); } catch (e) { alert("Save failed: " + e.message); } finally { setSaving(false); }
  };

  // ── Edit the anchor date (Start Date / Last Day) if circumstances change ──
  const anchorField = j.JourneyType === "Offboarding" ? "EndDate" : "StartDate";
  const anchorLabel = j.JourneyType === "Offboarding" ? "Last Day" : "Start Date";
  const [editDate, setEditDate] = useState(false);
  const [dateDraft, setDateDraft] = useState("");
  const beginEditDate = () => { setDateDraft(reviewDateOnly(journeyAnchorDate(j))); setEditDate(true); };
  const saveDate = async () => {
    if (!dateDraft) return alert(`${anchorLabel} is required.`);
    setSaving(true);
    try {
      // Tasks are generated as OffsetDays from the anchor — offer to re-align open ones.
      const shiftable = tasks.filter(t => t.Status !== "Done" && t.OffsetDays != null && t.OffsetDays !== "");
      const shift = shiftable.length > 0 && confirm(`Move the ${shiftable.length} open task due date(s) to match the new ${anchorLabel.toLowerCase()}? (Completed tasks are left as-is.)`);
      if (shift) for (const t of shiftable) await actions.updateTask(t.id, { DueDate: addDays(dateDraft, Number(t.OffsetDays) || 0) });
      await actions.updateJourney(j.id, { [anchorField]: dateDraft });
      // Onboarding start date also lives on the employee record (drives review cadence) — keep them in sync.
      if (anchorField === "StartDate" && emp) await actions.upsertEmployee({ Email: j.EmployeeEmail, StartDate: dateDraft });
      setEditDate(false);
    } catch (e) { alert("Save failed: " + e.message); } finally { setSaving(false); }
  };

  const cancelJourney = async () => { if (!confirm(`Cancel this ${j.JourneyType.toLowerCase()} journey?`)) return; await actions.updateJourney(j.id, { Status: "Cancelled" }); onClose(); };
  const reopenJourney = async () => { await actions.updateJourney(j.id, { Status: "In Progress" }); };
  const completeJourney = async () => { await actions.updateJourney(j.id, { Status: "Complete" }); };

  const p = journeyProgress(j, state.journeyTasks);

  return (
    <Modal title={`${j.JourneyType} — ${j.EmployeeName || emp?.Title || j.EmployeeEmail}`} width={860} onClose={onClose} footer={
      <>
        {hasRole("Admin") && j.Status !== "Cancelled" && <button style={S.btnO(C.er, C.er)} onClick={cancelJourney}>Cancel Journey</button>}
        {hasRole("Admin", "HR") && j.Status === "Cancelled" && <button style={S.btnO(C.t5)} onClick={reopenJourney}>Reopen</button>}
        {hasRole("Admin", "HR") && j.Status !== "Complete" && p.pct === 100 && <button style={S.btn(C.ok)} onClick={completeJourney}>Mark Complete</button>}
        <button style={S.btnO()} onClick={onClose}>Close</button>
      </>
    }>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>
        <div><div style={S.sec}>Status</div><Badge type={j.Status === "Complete" ? "ok" : j.Status === "Cancelled" ? "neutral" : "inf"}>{j.Status}</Badge></div>
        <div><div style={S.sec}>{anchorLabel}</div>
          {editDate ? (
            <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
              <input type="date" style={{ ...S.input, width: 150 }} value={dateDraft} disabled={saving} onChange={e => setDateDraft(e.target.value)} />
              <button style={{ ...S.btn(C.hdr), ...S.xs }} disabled={saving} onClick={saveDate}>{saving ? "Saving…" : "Save"}</button>
              <button style={{ ...S.btnO(), ...S.xs }} disabled={saving} onClick={() => setEditDate(false)}>Cancel</button>
            </div>
          ) : (
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <span style={{ fontWeight: 600, color: C.t7 }}>{fmtDate(journeyAnchorDate(j))}</span>
              {canEdit && j.Status !== "Cancelled" && <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={beginEditDate}>Edit</button>}
            </div>
          )}
        </div>
        <div><div style={S.sec}>Progress</div><div style={{ fontWeight: 600 }}>{p.done}/{p.total} ({p.pct}%)</div><ProgressBar value={p.pct} /></div>
      </div>
      <div style={{ background: C.t0, border: `1px solid ${C.t1}`, borderRadius: 6, padding: 10, marginBottom: 14, fontSize: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "auto 1fr", gap: "4px 12px" }}>
          <div style={{ color: C.b4, fontWeight: 600 }}>Email</div><div>{j.EmployeeEmail}</div>
          <div style={{ color: C.b4, fontWeight: 600 }}>Title</div><div>{j.JobTitle || emp?.JobTitle || "—"}</div>
          <div style={{ color: C.b4, fontWeight: 600 }}>Manager</div><div>{j.ManagerEmail || emp?.ManagerEmail || "—"}</div>
          {j.TemplateGroup && <><div style={{ color: C.b4, fontWeight: 600 }}>Group</div><div><Badge type="pu">{j.TemplateGroup}</Badge></div></>}
          {j.OffboardReason && <><div style={{ color: C.b4, fontWeight: 600 }}>Reason</div><div>{j.OffboardReason}</div></>}
          {j.Notes && <><div style={{ color: C.b4, fontWeight: 600 }}>Notes</div><div style={{ whiteSpace: "pre-wrap" }}>{j.Notes}</div></>}
        </div>
      </div>

      {/* Offboarding visibility gate — confidential to HR/Admin until shared with employee + their manager line */}
      {j.JourneyType === "Offboarding" && (
        <div style={{ marginBottom: 14, padding: 10, background: j.SharedWithEmployee ? C.okb : C.wnb, border: `1px solid ${j.SharedWithEmployee ? C.ok + "33" : C.wn + "33"}`, borderRadius: 6, fontSize: 12, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
          <div>
            {j.SharedWithEmployee ? (
              <><strong style={{ color: C.ok }}>👁 Shared</strong> — the employee and their manager line can see this journey.</>
            ) : (
              <><strong style={{ color: C.wn }}>🔒 Confidential</strong> — only Admin / HR can see this journey. The employee will <em>not</em> be informed via this app until you share it.</>
            )}
          </div>
          {hasRole("Admin", "HR") && j.Status !== "Cancelled" && (
            <button
              style={j.SharedWithEmployee ? S.btnO(C.wn, C.wn) : S.btn(C.ok)}
              disabled={saving}
              onClick={async () => {
                const next = !j.SharedWithEmployee;
                if (!next && !confirm("Unshare this journey? The employee and their manager line will lose visibility immediately.")) return;
                if (next && j.JourneyType === "Offboarding" && (j.OffboardReason || "").startsWith("Termination") && !confirm("This is an involuntary termination. Sharing will make the journey visible to the employee. Continue?")) return;
                setSaving(true);
                try { await actions.updateJourney(j.id, { SharedWithEmployee: next }); }
                catch (e) { alert("Failed to update: " + e.message); }
                finally { setSaving(false); }
              }}
            >
              {j.SharedWithEmployee ? "Unshare" : "Share with employee"}
            </button>
          )}
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <div style={{ ...S.sec, marginBottom: 0 }}>Tasks</div>
        {canEdit && j.Status !== "Cancelled" && j.Status !== "Complete" && (
          <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => actions.openAddTask(j.id)}>+ Add Task</button>
        )}
      </div>

      {phases.map(ph => (
        <div key={ph} style={{ marginBottom: 16 }}>
          <div style={{ ...S.sec, marginBottom: 6 }}>{ph}</div>
          <div style={{ border: `1px solid ${C.b1}`, borderRadius: 6, overflow: "hidden" }}>
            {tasks.filter(t => (t.Phase || "General") === ph).map((t, idx, arr) => {
              const cls = classifyDue(t.DueDate, t.Status);
              const dueColor = { overdue: C.er, soon: C.wn, ok: C.b6, done: C.b4, none: C.b4 }[cls];
              return (
                <div key={t.id} style={{ display: "grid", gridTemplateColumns: "auto 1fr auto", alignItems: "center", gap: 10, padding: "10px 12px", borderBottom: idx < arr.length - 1 ? `1px solid ${C.b1}` : "none", background: t.Status === "Done" ? C.okb : C.wh }}>
                  <input type="checkbox" checked={t.Status === "Done"} disabled={!canEdit || saving} onChange={e => setTask(t, { Status: e.target.checked ? "Done" : "Pending", CompletedDate: e.target.checked ? todayIso() : "", CompletedBy: e.target.checked ? currentEmail : "" })} style={{ width: 18, height: 18, accentColor: C.ok, cursor: canEdit ? "pointer" : "default" }} />
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: 13, color: C.t7, textDecoration: t.Status === "Done" ? "line-through" : "none" }}>{t.Title}{t.Required === false && <span style={{ marginLeft: 6, fontSize: 10, color: C.b4, fontWeight: 500 }}>(optional)</span>}</div>
                    <div style={{ display: "flex", gap: 12, fontSize: 11, marginTop: 3, flexWrap: "wrap" }}>
                      <span style={{ color: C.b4 }}>Assignee: <strong style={{ color: C.t6 }}>{t.AssigneeRole || "—"}{t.AssigneeEmail ? ` · ${t.AssigneeEmail}` : ""}</strong></span>
                      <span style={{ color: dueColor, fontWeight: 600 }}>Due {fmtDate(t.DueDate)}</span>
                      {t.Status === "Done" && t.CompletedDate && <span style={{ color: C.ok }}>✓ {fmtDate(t.CompletedDate)}</span>}
                    </div>
                    {t.Notes && <div style={{ fontSize: 11, color: C.b4, marginTop: 4, whiteSpace: "pre-wrap" }}>{t.Notes}</div>}
                  </div>
                  {canEdit && <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => actions.openTaskEdit(t.id)}>Edit</button>}
                </div>
              );
            })}
          </div>
        </div>
      ))}
    </Modal>
  );
}

// ============================================================
// TASK EDIT MODAL
// ============================================================
// ============================================================
// AD-HOC TASK MODAL — add a one-off task to an existing journey
// (lives outside the template — won't recur for future journeys)
// ============================================================
function AdHocTaskModal({ journeyId, onClose }) {
  const { state, actions, currentEmail } = useData();
  const j = state.journeys.find(x => String(x.id) === String(journeyId));
  const anchor = j ? journeyAnchorDate(j) : todayIso();
  const [f, setF] = useState({
    Title: "",
    Phase: "Ad-hoc",
    AssigneeRole: "HR",
    AssigneeEmail: "",
    DueDate: anchor || todayIso(),
    Required: true,
    Notes: "",
    Status: "Pending",
  });
  const [saving, setSaving] = useState(false);
  if (!j) return null;

  // Auto-compute OffsetDays from the chosen due date so re-anchoring (Start/End-date edits)
  // moves this task along with the rest.
  const offset = anchor && f.DueDate ? Math.round((new Date(f.DueDate) - new Date(anchor)) / 86400000) : 0;

  const save = async () => {
    if (!f.Title.trim()) return alert("Title is required.");
    setSaving(true);
    try {
      await actions.createTask({
        JourneyId: String(j.id),
        EmployeeEmail: (j.EmployeeEmail || "").toLowerCase(),
        Phase: f.Phase || "Ad-hoc",
        Title: f.Title.trim(),
        AssigneeRole: f.AssigneeRole || "",
        AssigneeEmail: f.AssigneeRole === "Manager"
          ? (j.ManagerEmail || "").toLowerCase()
          : f.AssigneeRole === "Employee"
          ? (j.EmployeeEmail || "").toLowerCase()
          : (f.AssigneeEmail || "").toLowerCase(),
        DueDate: f.DueDate || "",
        OffsetDays: offset,
        Required: !!f.Required,
        Status: "Pending",
        Notes: f.Notes || "",
        OrderIdx: 9999, // sort to the bottom within its phase
        TemplateId: "",
      });
      await actions.reload();
      onClose();
    } catch (e) { alert("Failed to add task: " + e.message); } finally { setSaving(false); }
  };

  return (
    <Modal title={`Add Task — ${j.JourneyType}: ${j.EmployeeName || j.EmployeeEmail}`} width={620} onClose={onClose} footer={
      <>
        <button style={S.btnO()} onClick={onClose} disabled={saving}>Cancel</button>
        <button style={S.btn(C.hdr)} onClick={save} disabled={saving}>{saving ? "Adding…" : "Add Task"}</button>
      </>
    }>
      <div style={{ display: "grid", gap: 12 }}>
        <div><label style={S.label}>Title *</label><input style={S.input} value={f.Title} placeholder="What needs to happen?" onChange={e => setF({ ...f, Title: e.target.value })} autoFocus /></div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Phase</label><input style={S.input} list="phase-suggestions" value={f.Phase} onChange={e => setF({ ...f, Phase: e.target.value })} />
            <datalist id="phase-suggestions">
              {uniq(state.journeyTasks.filter(t => String(t.JourneyId) === String(j.id)).map(t => t.Phase)).filter(Boolean).map(p => <option key={p} value={p} />)}
              <option value="Ad-hoc" />
            </datalist>
          </div>
          <div><label style={S.label}>Due Date</label><input style={S.input} type="date" value={f.DueDate || ""} onChange={e => setF({ ...f, DueDate: e.target.value })} /></div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Assignee Role</label>
            <select style={S.select} value={f.AssigneeRole} onChange={e => setF({ ...f, AssigneeRole: e.target.value })}>
              <option value="">—</option><option>HR</option><option>IT</option><option>Manager</option><option>Employee</option><option>Admin</option><option>Accounting</option>
            </select>
          </div>
          <div><label style={S.label}>Assignee Email (override)</label>
            <input style={S.input} value={f.AssigneeEmail} disabled={f.AssigneeRole === "Manager" || f.AssigneeRole === "Employee"}
              placeholder={f.AssigneeRole === "Manager" ? (j.ManagerEmail || "—") : f.AssigneeRole === "Employee" ? (j.EmployeeEmail || "—") : "optional"}
              onChange={e => setF({ ...f, AssigneeEmail: e.target.value.toLowerCase() })} />
          </div>
        </div>
        <div><label style={S.label}>Notes</label><textarea style={S.textarea} value={f.Notes} onChange={e => setF({ ...f, Notes: e.target.value })} /></div>
        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}><input type="checkbox" checked={f.Required} onChange={e => setF({ ...f, Required: e.target.checked })} /> Required</label>
        <div style={{ fontSize: 11, color: C.b4 }}>
          Anchor offset: <strong style={{ fontFamily: mono }}>{offset > 0 ? "+" : ""}{offset}d</strong> from {j.JourneyType === "Offboarding" ? "last day" : "start date"} — task moves with the anchor if it's edited later.
        </div>
      </div>
    </Modal>
  );
}

function TaskEditModal({ taskId, onClose }) {
  const { state, actions } = useData();
  const t = state.journeyTasks.find(x => String(x.id) === String(taskId));
  const [f, setF] = useState(() => ({ Title: t?.Title || "", DueDate: t?.DueDate || "", AssigneeEmail: t?.AssigneeEmail || "", AssigneeRole: t?.AssigneeRole || "", Notes: t?.Notes || "", Status: t?.Status || "Pending", Required: t?.Required !== false }));
  const [saving, setSaving] = useState(false);
  if (!t) return null;
  const save = async () => { setSaving(true); try { await actions.updateTask(t.id, f); onClose(); } catch (e) { alert("Save failed: " + e.message); } finally { setSaving(false); } };
  const del = async () => { if (!confirm("Delete this task?")) return; setSaving(true); try { await actions.deleteTask(t.id); onClose(); } catch (e) { alert("Delete failed: " + e.message); } finally { setSaving(false); } };

  return (
    <Modal title="Edit Task" onClose={onClose} footer={
      <>
        <button style={S.btnO(C.er, C.er)} onClick={del} disabled={saving}>Delete</button>
        <button style={S.btnO()} onClick={onClose} disabled={saving}>Cancel</button>
        <button style={S.btn(C.hdr)} onClick={save} disabled={saving}>{saving ? "Saving…" : "Save"}</button>
      </>
    }>
      <div style={{ display: "grid", gap: 12 }}>
        <div><label style={S.label}>Title</label><input style={S.input} value={f.Title} onChange={e => setF({ ...f, Title: e.target.value })} /></div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Due Date</label><input style={S.input} type="date" value={f.DueDate || ""} onChange={e => setF({ ...f, DueDate: e.target.value })} /></div>
          <div><label style={S.label}>Status</label><select style={S.select} value={f.Status} onChange={e => setF({ ...f, Status: e.target.value })}><option>Pending</option><option>In Progress</option><option>Done</option><option>Blocked</option></select></div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Assignee Role</label><select style={S.select} value={f.AssigneeRole} onChange={e => setF({ ...f, AssigneeRole: e.target.value })}><option value="">—</option><option>HR</option><option>IT</option><option>Manager</option><option>Employee</option><option>Admin</option></select></div>
          <div><label style={S.label}>Assignee Email (override)</label><input style={S.input} value={f.AssigneeEmail} placeholder="optional" onChange={e => setF({ ...f, AssigneeEmail: e.target.value })} /></div>
        </div>
        <div><label style={S.label}>Notes</label><textarea style={S.textarea} value={f.Notes} onChange={e => setF({ ...f, Notes: e.target.value })} /></div>
        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}><input type="checkbox" checked={f.Required} onChange={e => setF({ ...f, Required: e.target.checked })} /> Required</label>
      </div>
    </Modal>
  );
}

// ============================================================
// START JOURNEY MODAL
// ============================================================
function StartJourneyModal({ type, onClose }) {
  const { state, actions, currentEmail } = useData();
  const isOnboarding = type === "Onboarding";
  // Multi-step wizard for onboarding (Profile → Permissions → Confirm); offboarding is single-step
  const [step, setStep] = useState(1);
  const [f, setF] = useState({
    EmployeeEmail: "", EmployeeName: "", JobTitle: "", ManagerEmail: "",
    Department: "", StartDate: isOnboarding ? todayIso() : "", EndDate: !isOnboarding ? todayIso() : "",
    Notes: "", existingEmpId: "", OffboardReason: "", IsContractor: false,
    TemplateGroup: (CONFIG.templateGroups[type] || [])[0] || "",
  });
  // Permission selections: { ColumnName: roleValue }
  const [perms, setPerms] = useState(() => {
    const o = {};
    for (const a of state.apps) o[a.ColumnName] = a.OnboardingDefault && a.OnboardingDefault !== "None" ? (a.OnboardingDefault === "Lowest" ? (a.Roles?.[0] || "") : a.OnboardingDefault === "Standard" ? (a.Roles?.[0] || "") : a.OnboardingDefault) : "";
    return o;
  });
  const [saving, setSaving] = useState(false);
  const [warn, setWarn] = useState("");

  // Generation rule: include templates that are universal (empty group), exact-group
  // match, or a meta-group whose rule applies (e.g. Common - W-2). Sorted chronologically.
  const tplOpts = { isContractor: isOnboarding && f.IsContractor };
  const usableTemplates = state.templates
    .filter(t => t.JourneyType === type && t.Active !== false)
    .filter(t => templateAppliesToGroup(t, f.TemplateGroup, tplOpts))
    .sort((a, b) => (a.OffsetDays || 0) - (b.OffsetDays || 0) || (a.Title || "").localeCompare(b.Title || ""));
  const availableGroups = getAvailableGroups(state, type);

  const onPickEmployee = (id) => {
    if (!id) { setF({ ...f, existingEmpId: "", EmployeeEmail: "", EmployeeName: "", JobTitle: "", ManagerEmail: "" }); return; }
    const e = state.employees.find(x => String(x.id) === String(id));
    if (!e) return;
    setF({ ...f, existingEmpId: id, EmployeeEmail: (e.Email || "").toLowerCase(), EmployeeName: e.Title || "", JobTitle: e.JobTitle || "", ManagerEmail: (e.ManagerEmail || "").toLowerCase(), Department: e.Department || f.Department });
    // Pre-populate permissions from existing record (for re-onboarding/transfer)
    if (isOnboarding) {
      const o = { ...perms };
      for (const a of state.apps) if (e[a.ColumnName]) o[a.ColumnName] = e[a.ColumnName];
      setPerms(o);
    }
  };

  const validateStep1 = () => {
    if (!f.EmployeeEmail) return "Employee email is required.";
    if (isOnboarding && !f.EmployeeName) return "Full name is required for new hires.";
    if (isOnboarding && !f.StartDate) return "Start date is required.";
    if (!isOnboarding && !f.EndDate) return "Last day is required.";
    return null;
  };

  const submit = async () => {
    setWarn("");
    const v = validateStep1(); if (v) { setStep(1); return setWarn(v); }
    const dup = state.journeys.find(j => j.JourneyType === type && (j.EmployeeEmail || "").toLowerCase() === f.EmployeeEmail.toLowerCase() && j.Status !== "Complete" && j.Status !== "Cancelled");
    if (dup) { if (!confirm(`An active ${type.toLowerCase()} journey already exists for this employee. Create another anyway?`)) return; }

    setSaving(true);
    try {
      const anchor = !isOnboarding ? f.EndDate : f.StartDate;
      // 1. Upsert employee on onboarding (creates Employees row if new)
      let empRecord = null;
      if (isOnboarding) {
        const chain = managerChain((f.ManagerEmail || "").toLowerCase(), state.employees);
        empRecord = await actions.upsertEmployee({
          Title: f.EmployeeName,
          Name: f.EmployeeName,
          Email: f.EmployeeEmail.toLowerCase(),
          JobTitle: f.JobTitle,
          Department: f.Department,
          StartDate: f.StartDate || "",
          PersonaType: f.TemplateGroup === "Virtual Assistant" ? "Virtual Assistant" : "Employee",
          OnboardingStatus: "Pending",
          ManagerEmail: (f.ManagerEmail || "").toLowerCase(),
          ...chain,
          EmployeeActive: true,
        });
      }
      // 2. Create the journey
      const jrn = await actions.createJourney({
        JourneyType: type,
        EmployeeEmail: f.EmployeeEmail.toLowerCase(),
        EmployeeName: f.EmployeeName,
        JobTitle: f.JobTitle,
        ManagerEmail: (f.ManagerEmail || "").toLowerCase(),
        Department: f.Department,
        StartDate: f.StartDate || "",
        EndDate: f.EndDate || "",
        Status: "In Progress",
        Notes: f.Notes,
        OffboardReason: !isOnboarding ? f.OffboardReason : "",
        TemplateGroup: f.TemplateGroup || "",
        CreatedBy: currentEmail,
      });
      // 3. Set permissions (onboarding) with audit log against the new journey
      if (isOnboarding) {
        const cleanPerms = {};
        for (const [k, v2] of Object.entries(perms)) if (v2) cleanPerms[k] = v2;
        if (Object.keys(cleanPerms).length > 0) {
          await actions.setEmployeePermissions(f.EmployeeEmail.toLowerCase(), cleanPerms, "Initial onboarding permissions", String(jrn.id), empRecord);
        }
      }
      // 4. Create tasks from templates
      let idx = 0;
      for (const tpl of usableTemplates) {
        idx++;
        const due = addDays(anchor, tpl.OffsetDays || 0);
        await actions.createTask({
          JourneyId: String(jrn.id),
          EmployeeEmail: f.EmployeeEmail.toLowerCase(),
          Phase: tpl.Phase || "General",
          Title: tpl.Title,
          AssigneeRole: tpl.AssigneeRole || "",
          AssigneeEmail: tpl.AssigneeRole === "Manager" ? (f.ManagerEmail || "") : tpl.AssigneeRole === "Employee" ? f.EmployeeEmail.toLowerCase() : "",
          DueDate: due,
          OffsetDays: tpl.OffsetDays || 0,
          Required: tpl.Required !== false,
          Status: "Pending",
          Notes: tpl.Notes || "",
          OrderIdx: idx,
          TemplateId: String(tpl.id),
        });
      }
      await actions.reload();
      onClose();
    } catch (e) {
      setWarn("Failed to start journey: " + e.message);
    } finally { setSaving(false); }
  };

  const steps = isOnboarding ? [
    { n: 1, label: "Employee" },
    { n: 2, label: "App Permissions" },
    { n: 3, label: "Confirm" },
  ] : [{ n: 1, label: "Details" }];

  return (
    <Modal title={`Start ${type}`} onClose={onClose} width={720} footer={
      <>
        {step > 1 && <button style={S.btnO()} onClick={() => setStep(step - 1)} disabled={saving}>Back</button>}
        <button style={S.btnO()} onClick={onClose} disabled={saving}>Cancel</button>
        {isOnboarding && step < 3
          ? <button style={S.btn(C.hdr)} onClick={() => { const v = step === 1 ? validateStep1() : null; if (v) return setWarn(v); setWarn(""); setStep(step + 1); }} disabled={saving}>Next →</button>
          : <button style={S.btn(C.hdr)} onClick={submit} disabled={saving}>{saving ? "Creating…" : `Start ${type}`}</button>}
      </>
    }>
      {/* Stepper */}
      {isOnboarding && (
        <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
          {steps.map(s => (
            <div key={s.n} style={{ flex: 1, padding: "8px 10px", borderRadius: 4, background: s.n === step ? C.hdr : s.n < step ? C.t1 : C.b1, color: s.n === step ? "#fff" : s.n < step ? C.t7 : C.b4, fontSize: 12, fontWeight: 600, textAlign: "center" }}>
              <span style={{ opacity: 0.6, marginRight: 6 }}>{s.n}</span>{s.label}
            </div>
          ))}
        </div>
      )}

      {/* STEP 1 — Employee details */}
      {step === 1 && (
        <div style={{ display: "grid", gap: 12 }}>
          <div>
            <label style={S.label}>{isOnboarding ? "Type of hire *" : "Type of departure *"}</label>
            <select style={S.select} value={f.TemplateGroup} onChange={e => setF({ ...f, TemplateGroup: e.target.value })}>
              {availableGroups.length === 0 && <option value="">(no groups defined — uses common tasks only)</option>}
              {availableGroups.map(g => (<option key={g} value={g}>{g}</option>))}
            </select>
            <div style={{ fontSize: 11, color: C.b4, marginTop: 4 }}>
              Determines which template group runs. <strong>{usableTemplates.length}</strong> task{usableTemplates.length === 1 ? "" : "s"} will generate
              {(() => {
                const all = state.templates.filter(t => t.JourneyType === type && t.Active !== false);
                const universal = all.filter(t => templateGroups(t).length === 0).length;
                const groupSpecific = all.filter(t => {
                  const tg = templateGroups(t);
                  return tg.length > 0 && tg.includes(f.TemplateGroup) && !tg.some(g => isMetaGroup(g));
                }).length;
                const metaBreakdown = {};
                all.filter(t => templateGroups(t).some(g => isMetaGroup(g)) && templateAppliesToGroup(t, f.TemplateGroup, tplOpts))
                  .forEach(t => {
                    const meta = templateGroups(t).find(g => isMetaGroup(g));
                    if (meta) metaBreakdown[meta] = (metaBreakdown[meta] || 0) + 1;
                  });
                const parts = [];
                if (universal) parts.push(`${universal} universal`);
                Object.entries(metaBreakdown).forEach(([k, v]) => parts.push(`${v} ${k}`));
                if (groupSpecific) parts.push(`${groupSpecific} group-specific`);
                return parts.length ? " (" + parts.join(" + ") + ")" : "";
              })()}.
            </div>
          </div>
          {isOnboarding && f.TemplateGroup !== "Virtual Assistant" && (
            <label style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 13, background: f.IsContractor ? C.g0 : C.b1, border: `1px solid ${f.IsContractor ? C.g4 : C.b2}`, borderRadius: 6, padding: "10px 12px" }}>
              <input type="checkbox" style={{ marginTop: 2 }} checked={f.IsContractor} onChange={e => setF({ ...f, IsContractor: e.target.checked })} />
              <span>
                <strong>1099 independent contractor</strong> (not W-2 payroll)
                <div style={{ fontSize: 11, color: C.b4, marginTop: 2 }}>
                  Swaps the W-2 paperwork track (I-9, W-4, direct deposit, benefits) for the contractor track: ICA, W-9/W-8BEN, scope of work, and a misclassification risk review.
                </div>
              </span>
            </label>
          )}
          <div>
            <label style={S.label}>{isOnboarding ? "Existing employee (rehire/transfer)" : "Select active employee *"}</label>
            <select style={S.select} value={f.existingEmpId} onChange={e => onPickEmployee(e.target.value)}>
              <option value="">{isOnboarding ? "— New hire —" : "— Pick an employee —"}</option>
              {state.employees
                .filter(e => e.Email)
                .filter(e => isOnboarding || e.EmployeeActive !== false)
                .sort((a, b) => (a.Title || "").localeCompare(b.Title || ""))
                .map(e => (<option key={e.id} value={e.id}>{e.Title} — {e.Email}</option>))}
            </select>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <div><label style={S.label}>Full Name {isOnboarding && "*"}</label><input style={S.input} value={f.EmployeeName} onChange={e => setF({ ...f, EmployeeName: e.target.value })} /></div>
            <div><label style={S.label}>Email *</label><input style={S.input} value={f.EmployeeEmail} disabled={!isOnboarding && !!f.existingEmpId} onChange={e => setF({ ...f, EmployeeEmail: e.target.value.toLowerCase() })} placeholder="name@newshirepm.com" /></div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <div><label style={S.label}>Job Title</label><input style={S.input} value={f.JobTitle} onChange={e => setF({ ...f, JobTitle: e.target.value })} /></div>
            <div><label style={S.label}>Department</label><input style={S.input} value={f.Department} onChange={e => setF({ ...f, Department: e.target.value })} /></div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <div>
              <label style={S.label}>Direct Supervisor Email</label>
              <input style={S.input} value={f.ManagerEmail} placeholder="who this person reports to" onChange={e => setF({ ...f, ManagerEmail: e.target.value.toLowerCase() })} />
              <div style={{ fontSize: 11, color: C.b4, marginTop: 4 }}>The person this hire reports to day-to-day — used to route Manager onboarding tasks.</div>
            </div>
            {isOnboarding ? (
              <div><label style={S.label}>Start Date *</label><input style={S.input} type="date" value={f.StartDate} onChange={e => setF({ ...f, StartDate: e.target.value })} /></div>
            ) : (
              <div><label style={S.label}>Last Day *</label><input style={S.input} type="date" value={f.EndDate} onChange={e => setF({ ...f, EndDate: e.target.value })} /></div>
            )}
          </div>
          {!isOnboarding && (
            <div><label style={S.label}>Reason for offboarding</label>
              <select style={S.select} value={f.OffboardReason} onChange={e => setF({ ...f, OffboardReason: e.target.value })}>
                <option value="">—</option>
                <option>Resignation</option><option>Retirement</option><option>End of Contract</option>
                <option>Termination — Performance</option><option>Termination — Misconduct</option>
                <option>Layoff / Position Eliminated</option><option>Other</option>
              </select>
            </div>
          )}
          <div><label style={S.label}>Notes</label><textarea style={S.textarea} value={f.Notes} onChange={e => setF({ ...f, Notes: e.target.value })} placeholder="Anything the team needs to know" /></div>
          {!isOnboarding && (
            <div style={{ background: C.infb, border: `1px solid ${C.inf}33`, borderRadius: 4, padding: 10, fontSize: 12, color: C.inf }}>
              This will create <strong>{usableTemplates.length}</strong> offboarding task{usableTemplates.length === 1 ? "" : "s"} (offsets from the last day). The employee's app permissions will be reviewable on the Last Day tasks.
            </div>
          )}
        </div>
      )}

      {/* STEP 2 — Permissions (onboarding only) */}
      {isOnboarding && step === 2 && (
        <div style={{ display: "grid", gap: 12 }}>
          <div style={{ background: C.infb, border: `1px solid ${C.inf}33`, borderRadius: 4, padding: 10, fontSize: 12, color: C.inf }}>
            Choose which NewShire apps this employee needs access to and what role each gets. <strong>None</strong> = no access. You can change any of these later from the employee's profile.
          </div>
          {state.apps.length === 0
            ? <Empty title="No apps registered" sub="Run the provisioning script (or add rows to ELC_Apps) to populate the registry." />
            : <PermissionMatrix apps={state.apps} values={perms} onChange={(col, v) => setPerms({ ...perms, [col]: v })} />}
        </div>
      )}

      {/* STEP 3 — Confirm (onboarding only) */}
      {isOnboarding && step === 3 && (
        <div style={{ display: "grid", gap: 10 }}>
          <div style={S.card}>
            <div style={S.sec}>New Hire</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: C.t7 }}>{f.EmployeeName} <span style={{ fontWeight: 400, color: C.b4 }}>· {f.EmployeeEmail}</span></div>
            <div style={{ fontSize: 12, color: C.b4, marginTop: 2 }}>{f.JobTitle || "—"}{f.Department ? ` · ${f.Department}` : ""} · Manager: {f.ManagerEmail || "—"} · Starts {fmtDate(f.StartDate)}</div>
            <div style={{ marginTop: 6, display: "flex", gap: 6, flexWrap: "wrap" }}><Badge type="pu">{f.TemplateGroup || "Common"}</Badge>{f.IsContractor && <Badge type="wn">1099 Contractor</Badge>}</div>
          </div>
          <div style={S.card}>
            <div style={S.sec}>App Permissions</div>
            {Object.entries(perms).filter(([, v]) => v).length === 0 ? <div style={{ fontSize: 12, color: C.b4 }}>No app permissions selected.</div> : (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {Object.entries(perms).filter(([, v]) => v).map(([col, role]) => {
                  const app = state.apps.find(a => a.ColumnName === col);
                  return <Badge key={col} type="inf">{app?.Title || col}: {role}</Badge>;
                })}
              </div>
            )}
          </div>
          <div style={{ background: C.g0, border: `1px dashed ${C.g4}`, borderRadius: 4, padding: 10, fontSize: 12, color: C.g7 }}>
            On submit: ① {f.existingEmpId ? "Update" : "Add"} <strong>{f.EmployeeName}</strong> in the Employees list. ② Set the {Object.values(perms).filter(Boolean).length} app permission(s) above (audit-logged). ③ Generate {usableTemplates.length} onboarding tasks from the active template.
          </div>
        </div>
      )}

      {warn && <div style={{ background: C.erb, color: C.er, padding: 10, borderRadius: 4, fontSize: 12, marginTop: 12 }}>{warn}</div>}
    </Modal>
  );
}

// ============================================================
// TEMPLATES TAB
// ============================================================
// All non-meta groups that exist for a given JourneyType — these are the ones
// users actually select at journey start. Meta-groups (Common - W-2 etc.) are
// shown only in Templates editor UI, not in the journey-start picker.
function getAvailableGroups(state, journeyType) {
  const predefined = CONFIG.templateGroups[journeyType] || [];
  // Multi-group templates contribute each of their groups individually.
  const custom = uniq(state.templates
    .filter(t => t.JourneyType === journeyType)
    .flatMap(t => templateGroups(t))
    .filter(g => g && !isMetaGroup(g)));
  return uniq([...predefined, ...custom]);
}
// All Phase strings seen in the templates of a given JourneyType — used to
// power the Phase autocomplete in the template editor.
function getAvailablePhases(state, journeyType) {
  return uniq(state.templates
    .filter(t => t.JourneyType === journeyType && t.Phase)
    .map(t => String(t.Phase).trim())
    .filter(Boolean));
}
// Meta-groups defined for a given JourneyType — used to populate the Templates editor filter
function getMetaGroupsForType(journeyType) {
  return Object.entries(META_GROUP_RULES)
    .filter(([, rule]) => rule.journeyType === journeyType)
    .map(([name, rule]) => ({ name, label: rule.label }));
}

function TemplatesTab() {
  const { state, actions, hasRole } = useData();
  const [seed, setSeed] = useState(false);
  const [editing, setEditing] = useState(null);
  const [filter, setFilter] = useState("Onboarding");
  const [groupFilter, setGroupFilter] = useState("__all__"); // __all__, __common__, or group name
  const canEdit = hasRole("Admin", "HR");

  const groups = getAvailableGroups(state, filter);
  const metaGroups = getMetaGroupsForType(filter);
  const visible = state.templates
    .filter(t => t.JourneyType === filter)
    .filter(t => {
      if (groupFilter === "__all__") return true;
      const tg = templateGroups(t);
      if (groupFilter === "__common__") return tg.length === 0;
      return tg.includes(groupFilter);
    })
    .sort((a, b) => (a.OffsetDays || 0) - (b.OffsetDays || 0) || (a.Title || "").localeCompare(b.Title || ""));

  const seedDefaults = async () => {
    // Idempotent: skip any default that already exists (matched on type + group + title),
    // so re-seeding after new defaults are added only creates the genuinely-new tasks.
    const keyOf = t => `${t.JourneyType}|${t.TemplateGroup || ""}|${(t.Title || "").trim()}`;
    const have = new Set(state.templates.map(keyOf));
    const toSeed = DEFAULT_TEMPLATES.filter(t => !have.has(keyOf(t)));
    if (toSeed.length === 0) { alert("All default templates already exist — nothing to seed."); return; }
    const skipped = DEFAULT_TEMPLATES.length - toSeed.length;
    if (!confirm(`Seed ${toSeed.length} new default template task(s)?${skipped ? ` ${skipped} already present will be skipped.` : ""}`)) return;
    setSeed(true);
    try {
      for (const tpl of toSeed) {
        await actions.createTemplate({ ...tpl, Active: true });
      }
      await actions.reload();
    } catch (e) { alert("Seed failed: " + e.message); } finally { setSeed(false); }
  };

  // New default GROUPS for this journey type that the user has none of yet.
  // We only ever add wholesale-missing groups (e.g. the new Common - 1099 track),
  // never individual tasks — so customized/edited templates are never duplicated
  // or overwritten. A group the user already has even one task in is left alone.
  const newGroups = useMemo(() => {
    const present = new Set(state.templates.filter(t => t.JourneyType === filter).map(t => t.TemplateGroup || ""));
    const defGroups = new Set(DEFAULT_TEMPLATES.filter(t => t.JourneyType === filter).map(t => t.TemplateGroup || ""));
    return [...defGroups].filter(g => g && !present.has(g)); // skip the "" universal bucket
  }, [state.templates, filter]);
  const newGroupTasks = useMemo(
    () => DEFAULT_TEMPLATES.filter(t => t.JourneyType === filter && newGroups.includes(t.TemplateGroup || "")),
    [filter, newGroups]
  );

  const seedNewGroups = async () => {
    if (newGroupTasks.length === 0) return;
    const label = newGroups.join(", ");
    if (!confirm(`Add ${newGroupTasks.length} task(s) from ${newGroups.length} new group(s) you don't have yet (${label})?\n\nYour existing and customized templates are left completely untouched.`)) return;
    setSeed(true);
    try {
      for (const tpl of newGroupTasks) await actions.createTemplate({ ...tpl, Active: true });
      await actions.reload();
    } catch (e) { alert("Seed failed: " + e.message); } finally { setSeed(false); }
  };

  // Group counts for the filter dropdown
  const groupCounts = useMemo(() => {
    const m = { __all__: 0, __common__: 0 };
    for (const g of groups) m[g] = 0;
    for (const t of state.templates.filter(x => x.JourneyType === filter)) {
      m.__all__++;
      const tg = templateGroups(t);
      if (tg.length === 0) { m.__common__++; continue; }
      // Multi-group templates count in every group they belong to.
      for (const g of tg) m[g] = (m[g] || 0) + 1;
    }
    return m;
  }, [state.templates, filter, groups.join("|")]);

  return (
    <div>
      <div style={S.card}>
        <div style={S.cardT}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            <span>Task Templates</span>
            <select style={{ ...S.select, width: "auto" }} value={filter} onChange={e => { setFilter(e.target.value); setGroupFilter("__all__"); }}>
              <option>Onboarding</option><option>Offboarding</option>
            </select>
            <select style={{ ...S.select, width: "auto" }} value={groupFilter} onChange={e => setGroupFilter(e.target.value)}>
              <option value="__all__">All groups ({groupCounts.__all__})</option>
              <option value="__common__">Universal — applies to every group ({groupCounts.__common__ || 0})</option>
              {metaGroups.length > 0 && (
                <optgroup label="Meta-groups (rule-based)">
                  {metaGroups.map(m => <option key={m.name} value={m.name}>{m.label} ({groupCounts[m.name] || 0})</option>)}
                </optgroup>
              )}
              <optgroup label="Group-specific">
                {groups.map(g => <option key={g} value={g}>{g} ({groupCounts[g] || 0})</option>)}
              </optgroup>
            </select>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            {canEdit && state.templates.length === 0 && <button style={S.btnO(C.t5)} onClick={seedDefaults} disabled={seed}>{seed ? "Seeding…" : "Seed defaults"}</button>}
            {canEdit && state.templates.length > 0 && newGroups.length > 0 && <button style={S.btnO(C.t5)} onClick={seedNewGroups} disabled={seed}>{seed ? "Seeding…" : `Add new group${newGroups.length === 1 ? "" : "s"}: ${newGroups.join(", ")}`}</button>}
            {canEdit && <button style={S.btn(C.hdr)} onClick={() => setEditing({ JourneyType: filter, TemplateGroup: groupFilter === "__all__" || groupFilter === "__common__" ? "" : groupFilter, Phase: "", Title: "", AssigneeRole: "HR", OffsetDays: 0, Required: true, Notes: "", Active: true })}>+ New Template</button>}
          </div>
        </div>
        <div style={{ fontSize: 11, color: C.b4, marginBottom: 10 }}>
          <strong>Common</strong> tasks apply to every group of this type. <strong>Group-specific</strong> tasks only generate when that group is chosen at journey start. So a "Maintenance" onboarding gets Common + Maintenance tasks, never On-Site or VA tasks.
        </div>
        {visible.length === 0 ? (
          <Empty title={`No ${filter} templates in this view`} sub={canEdit ? "Click 'Seed defaults' to load NewShire's starter set, or '+ New Template' to add your own." : "Ask HR/Admin to set this up."} />
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr>
                <th style={S.th}>Group</th><th style={S.th}>Phase</th><th style={S.th}>Title</th><th style={S.th}>Assignee</th><th style={S.th}>Offset</th><th style={S.th}>Required</th><th style={S.th}></th>
              </tr></thead>
              <tbody>
                {visible.map(t => {
                  const tg = templateGroups(t);
                  return (
                  <tr key={t.id}>
                    <td style={S.td}>{tg.length === 0
                      ? <Badge type="neutral">Universal</Badge>
                      : <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
                          {tg.map(g => <Badge key={g} type={isMetaGroup(g) ? "inf" : "pu"}>{g}</Badge>)}
                        </div>}</td>
                    <td style={S.td}><Badge type="neutral">{t.Phase || "—"}</Badge></td>
                    <td style={S.td}><div style={{ fontWeight: 600, color: C.t7 }}>{t.Title}</div>{t.Notes && <div style={{ fontSize: 11, color: C.b4 }}>{t.Notes}</div>}</td>
                    <td style={S.td}>{t.AssigneeRole}</td>
                    <td style={S.td}><span style={{ fontFamily: mono, fontSize: 12 }}>{t.OffsetDays > 0 ? `+${t.OffsetDays}` : t.OffsetDays}d</span></td>
                    <td style={S.td}>{t.Required === false ? <Badge type="neutral">Optional</Badge> : <Badge type="ok">Required</Badge>}</td>
                    <td style={S.td}>{canEdit && <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => setEditing(t)}>Edit</button>}</td>
                  </tr>
                );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
      {editing && <TemplateEditModal tpl={editing} onClose={() => setEditing(null)} />}
    </div>
  );
}

function TemplateEditModal({ tpl, onClose }) {
  const { state, actions } = useData();
  const isNew = !tpl.id;
  const [f, setF] = useState({ JourneyType: tpl.JourneyType, TemplateGroup: tpl.TemplateGroup || "", Phase: tpl.Phase || "", Title: tpl.Title || "", AssigneeRole: tpl.AssigneeRole || "HR", OffsetDays: tpl.OffsetDays ?? 0, Required: tpl.Required !== false, Notes: tpl.Notes || "", Active: tpl.Active !== false });
  const [saving, setSaving] = useState(false);
  // Working list of group names; persisted back as a comma-separated string.
  const selectedGroups = useMemo(() => templateGroups({ TemplateGroup: f.TemplateGroup }), [f.TemplateGroup]);
  const groupSuggestions = getAvailableGroups(state, f.JourneyType);
  const phaseSuggestions = getAvailablePhases(state, f.JourneyType);
  const groupDatalistId = `tpl-groups-${f.JourneyType}`;
  const phaseDatalistId = `tpl-phases-${f.JourneyType}`;
  const [groupAdd, setGroupAdd] = useState("");

  const setGroups = (next) => setF({ ...f, TemplateGroup: uniq(next).join(", ") });
  const addGroup = (name) => {
    const v = (name || "").trim();
    if (!v) return;
    if (selectedGroups.includes(v)) return;
    setGroups([...selectedGroups, v]);
  };
  const removeGroup = (name) => setGroups(selectedGroups.filter(g => g !== name));
  const toggleGroup = (name) => selectedGroups.includes(name) ? removeGroup(name) : addGroup(name);

  const save = async () => {
    if (!f.Title) return alert("Title is required.");
    setSaving(true);
    try {
      if (isNew) await actions.createTemplate(f);
      else await actions.updateTemplate(tpl.id, f);
      await actions.reload();
      onClose();
    } catch (e) { alert("Save failed: " + e.message); } finally { setSaving(false); }
  };
  const del = async () => {
    if (!confirm("Delete this template? Existing journeys will not be affected.")) return;
    setSaving(true);
    try { await actions.deleteTemplate(tpl.id); await actions.reload(); onClose(); } catch (e) { alert("Delete failed: " + e.message); } finally { setSaving(false); }
  };

  // Suggestions split into "already added" vs "available to add" so the
  // chip area below the input shows toggleable shortcuts for known groups.
  const availableShortcuts = groupSuggestions.filter(g => !selectedGroups.includes(g));

  return (
    <Modal title={isNew ? "New Template Task" : "Edit Template Task"} onClose={onClose} footer={
      <>
        {!isNew && <button style={S.btnO(C.er, C.er)} onClick={del} disabled={saving}>Delete</button>}
        <button style={S.btnO()} onClick={onClose} disabled={saving}>Cancel</button>
        <button style={S.btn(C.hdr)} onClick={save} disabled={saving}>{saving ? "Saving…" : "Save"}</button>
      </>
    }>
      <div style={{ display: "grid", gap: 12 }}>
        <div><label style={S.label}>Journey Type</label><select style={{ ...S.select, width: 200 }} value={f.JourneyType} onChange={e => setF({ ...f, JourneyType: e.target.value })}><option>Onboarding</option><option>Offboarding</option></select></div>
        <div>
          <label style={S.label}>Groups <span style={{ fontWeight: 400, color: C.b4 }}>(none = applies to every group of this type)</span></label>
          {/* Selected groups as removable chips */}
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, padding: "6px 8px", border: `1px solid ${C.b2}`, borderRadius: 4, background: C.wh, minHeight: 38, alignItems: "center" }}>
            {selectedGroups.length === 0 && <span style={{ color: C.b4, fontSize: 12 }}>No groups selected — task will apply to every group of this type</span>}
            {selectedGroups.map(g => (
              <span key={g} style={{ display: "inline-flex", alignItems: "center", gap: 4, background: isMetaGroup(g) ? C.infb : C.pub, color: isMetaGroup(g) ? C.inf : C.pu, fontSize: 11, fontWeight: 600, padding: "3px 4px 3px 8px", borderRadius: 99 }}>
                {g}
                <button onClick={() => removeGroup(g)} style={{ border: "none", background: "transparent", color: "inherit", fontSize: 14, cursor: "pointer", padding: "0 4px", lineHeight: 1, fontWeight: 700 }} title="Remove">×</button>
              </span>
            ))}
          </div>
          {/* Free-form add input with datalist autocomplete from existing groups */}
          <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
            <input
              style={{ ...S.input, flex: 1 }}
              list={groupDatalistId}
              value={groupAdd}
              placeholder="Type a group name and press Enter (or pick from list below)"
              onChange={e => setGroupAdd(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); addGroup(groupAdd); setGroupAdd(""); } }}
            />
            <datalist id={groupDatalistId}>{groupSuggestions.map(g => <option key={g} value={g} />)}</datalist>
            <button style={{ ...S.btnO(C.t5), ...S.sm }} onClick={() => { addGroup(groupAdd); setGroupAdd(""); }}>Add</button>
          </div>
          {/* Toggleable shortcuts for any predefined / existing group not yet selected */}
          {availableShortcuts.length > 0 && (
            <div style={{ marginTop: 6, fontSize: 11, color: C.b4 }}>
              Quick add:{" "}
              {availableShortcuts.map(g => (
                <button key={g} type="button" onClick={() => toggleGroup(g)} style={{ display: "inline-flex", alignItems: "center", gap: 3, fontSize: 11, fontWeight: 500, padding: "2px 8px", marginRight: 4, marginTop: 4, background: "transparent", color: C.t6, border: `1px dashed ${C.b2}`, borderRadius: 99, cursor: "pointer" }}>+ {g}</button>
              ))}
            </div>
          )}
        </div>
        <div><label style={S.label}>Title</label><input style={S.input} value={f.Title} onChange={e => setF({ ...f, Title: e.target.value })} /></div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10 }}>
          <div>
            <label style={S.label}>Phase</label>
            <input style={S.input} list={phaseDatalistId} value={f.Phase} placeholder="Day 1, Week 1, …" onChange={e => setF({ ...f, Phase: e.target.value })} />
            <datalist id={phaseDatalistId}>{phaseSuggestions.map(p => <option key={p} value={p} />)}</datalist>
          </div>
          <div><label style={S.label}>Assignee Role</label><select style={S.select} value={f.AssigneeRole} onChange={e => setF({ ...f, AssigneeRole: e.target.value })}><option>HR</option><option>IT</option><option>Manager</option><option>Employee</option><option>Admin</option><option>Accounting</option></select></div>
          <div><label style={S.label}>Offset Days</label><input style={S.input} type="number" value={f.OffsetDays} onChange={e => setF({ ...f, OffsetDays: parseInt(e.target.value) || 0 })} /></div>
          <div style={{ alignSelf: "end" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, paddingTop: 8 }}><input type="checkbox" checked={f.Required} onChange={e => setF({ ...f, Required: e.target.checked })} /> Required</label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13 }}><input type="checkbox" checked={f.Active} onChange={e => setF({ ...f, Active: e.target.checked })} /> Active</label>
          </div>
        </div>
        <div><label style={S.label}>Notes</label><textarea style={S.textarea} value={f.Notes} onChange={e => setF({ ...f, Notes: e.target.value })} /></div>
        <div style={{ fontSize: 11, color: C.b4 }}>
          <strong>Groups</strong> control which onboarding/offboarding flows this task appears in. Multi-select: the task generates when ANY of the listed groups is chosen at journey start. Leave empty to apply to every group of this type. Type a new name and press Enter (or click Add) to create a new group on the fly.<br />
          <strong>Phase</strong> picks from any phase you already use, or type a new one.<br />
          <strong>Offset</strong> is days from the anchor: positive = after start (onboarding) or after last day (offboarding); negative = before. E.g. <code style={{ fontFamily: mono }}>-7</code> = 7 days before.
        </div>
      </div>
    </Modal>
  );
}

// ============================================================
// MY TASKS
// ============================================================
function MyTasksTab() {
  const { state, currentEmail, hasRole } = useData();
  const myEmail = (currentEmail || "").toLowerCase();
  // Drop tasks whose parent journey is Cancelled or Complete — those tasks
  // are no longer actionable even if individually still "Pending".
  const mine = state.journeyTasks.filter(t => t.Status !== "Done" && taskIsMine(t, myEmail, hasRole))
    .map(t => ({ ...t, _j: state.journeys.find(j => String(j.id) === String(t.JourneyId)) }))
    .filter(t => t._j && t._j.Status !== "Cancelled" && t._j.Status !== "Complete")
    .sort((a, b) => (a.DueDate || "9999").localeCompare(b.DueDate || "9999"));
  const overdue = mine.filter(t => classifyDue(t.DueDate, t.Status) === "overdue");
  const soon = mine.filter(t => classifyDue(t.DueDate, t.Status) === "soon");
  const later = mine.filter(t => !["overdue", "soon"].includes(classifyDue(t.DueDate, t.Status)));

  const Section = ({ title, items, type }) => items.length === 0 ? null : (
    <div style={S.card}>
      <div style={S.cardT}><span>{title}</span><Badge type={type}>{items.length}</Badge></div>
      {items.map(t => <TaskRow key={t.id} t={t} />)}
    </div>
  );
  if (mine.length === 0) return <div style={S.card}><Empty title="Nothing assigned to you" sub="You're all caught up." /></div>;
  return (
    <div>
      <Section title="Overdue" items={overdue} type="er" />
      <Section title="Due Soon" items={soon} type="wn" />
      <Section title="Upcoming" items={later} type="inf" />
    </div>
  );
}
function TaskRow({ t }) {
  const { actions, currentEmail } = useData();
  const cls = classifyDue(t.DueDate, t.Status);
  const color = { overdue: C.er, soon: C.wn, ok: C.b6 }[cls] || C.b6;
  const markDone = () => actions.updateTask(t.id, { Status: "Done", CompletedDate: todayIso(), CompletedBy: currentEmail });
  return (
    <div style={{ display: "grid", gridTemplateColumns: "auto 1fr auto", gap: 10, padding: "10px 0", borderBottom: `1px solid ${C.b1}`, alignItems: "center" }}>
      <input type="checkbox" onChange={markDone} style={{ width: 18, height: 18, accentColor: C.ok, cursor: "pointer" }} />
      <div>
        <div style={{ fontWeight: 600, color: C.t7 }}>{t.Title}</div>
        <div style={{ fontSize: 11, color: C.b4, marginTop: 2 }}>
          {t._j?.JourneyType} · {t._j?.EmployeeName || t.EmployeeEmail} · <span style={{ color, fontWeight: 600 }}>Due {fmtDate(t.DueDate)}</span>
          {!t.AssigneeEmail && t.AssigneeRole && <span style={{ marginLeft: 6 }}><Badge type="neutral" dot={false}>{t.AssigneeRole} queue</Badge></span>}
        </div>
        {t.Notes && <div style={{ fontSize: 11, color: C.b4, marginTop: 4 }}>{t.Notes}</div>}
      </div>
      <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => actions.openJourney(t.JourneyId)}>View Journey</button>
    </div>
  );
}

// ============================================================
// PERMISSION MATRIX — reusable: onboarding wizard + employee detail
// ============================================================
function PermissionMatrix({ apps, values, onChange, disabled, dense }) {
  return (
    <div style={{ display: "grid", gap: dense ? 6 : 8 }}>
      {apps.map(a => {
        const cur = values[a.ColumnName] || "None";
        const choices = ["None", ...(a.Roles || [])];
        return (
          <div key={a.AppKey} style={{ display: "grid", gridTemplateColumns: "32px 1fr 160px", alignItems: "center", gap: 10, padding: dense ? "6px 8px" : "8px 10px", border: `1px solid ${cur === "None" ? C.b1 : C.t1}`, background: cur === "None" ? C.wh : C.t0, borderRadius: 6 }}>
            <div style={{ width: 28, height: 28, background: a.Color || C.t5, color: "#fff", borderRadius: 5, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13, fontFamily: "Georgia,serif" }}>{a.IconLetter || a.Title?.[0] || "?"}</div>
            <div style={{ minWidth: 0 }}>
              <div style={{ fontWeight: 600, color: C.t7, fontSize: 13 }}>{a.Title}</div>
              {a.Description && !dense && <div style={{ fontSize: 11, color: C.b4 }}>{a.Description}</div>}
            </div>
            <select style={S.select} disabled={disabled} value={cur} onChange={e => onChange(a.ColumnName, e.target.value)}>
              {choices.map(c => <option key={c} value={c === "None" ? "" : c}>{c}</option>)}
            </select>
          </div>
        );
      })}
    </div>
  );
}

// ============================================================
// NOTE EDIT MODAL — coaching, discipline, PIP, 1:1, praise, …
// ============================================================
const NOTE_TYPE_COLORS = {
  Coaching: "inf", Discipline: "er", Praise: "ok", PIP: "wn",
  "1:1": "neutral", Termination: "er", "Policy Acknowledgement": "neutral", General: "neutral",
};
function NoteEditModal({ noteId, forEmail, onClose }) {
  const { state, actions, currentEmail } = useData();
  const existing = noteId ? state.notes.find(n => String(n.id) === String(noteId)) : null;
  const [f, setF] = useState(() => existing ? { ...existing } : {
    EmployeeEmail: forEmail || "",
    NoteType: "Coaching",
    NoteDate: todayIso(),
    Body: "",
    Confidential: true,
    VisibleToManager: true,
    VisibleToEmployee: false,
    Status: "Open",
    Severity: "Info",
    FollowUpDate: "",
    AttachmentLinks: "",
    Title: "",
  });
  const [saving, setSaving] = useState(false);
  const isEdit = !!existing;

  const save = async () => {
    if (!f.EmployeeEmail) return alert("Employee email is required.");
    if (!f.Body) return alert("Note body is required.");
    setSaving(true);
    try {
      const payload = { ...f, Title: f.Title || `${f.NoteType} — ${fmtDate(f.NoteDate || todayIso())}` };
      if (isEdit) await actions.updateNote(noteId, payload);
      else await actions.createNote(payload);
      onClose();
    } catch (e) { alert("Save failed: " + e.message); } finally { setSaving(false); }
  };
  const del = async () => {
    if (!confirm("Delete this note? This cannot be undone.")) return;
    setSaving(true);
    try { await actions.deleteNote(noteId); onClose(); } catch (e) { alert("Delete failed: " + e.message); } finally { setSaving(false); }
  };

  return (
    <Modal title={isEdit ? "Edit Note" : "New Note"} width={680} onClose={onClose} footer={
      <>
        {isEdit && <button style={S.btnO(C.er, C.er)} onClick={del} disabled={saving}>Delete</button>}
        <button style={S.btnO()} onClick={onClose} disabled={saving}>Cancel</button>
        <button style={S.btn(C.hdr)} onClick={save} disabled={saving}>{saving ? "Saving…" : "Save"}</button>
      </>
    }>
      <div style={{ display: "grid", gap: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Type</label>
            <select style={S.select} value={f.NoteType} onChange={e => setF({ ...f, NoteType: e.target.value })}>
              {["Coaching","Discipline","Praise","PIP","1:1","Termination","Policy Acknowledgement","General"].map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
          <div><label style={S.label}>Date</label><input style={S.input} type="date" value={f.NoteDate || ""} onChange={e => setF({ ...f, NoteDate: e.target.value })} /></div>
          <div><label style={S.label}>Severity</label>
            <select style={S.select} value={f.Severity || "Info"} onChange={e => setF({ ...f, Severity: e.target.value })}>
              {["Info","Low","Medium","High","Critical"].map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
        </div>
        <div><label style={S.label}>Title (optional)</label><input style={S.input} value={f.Title || ""} placeholder="auto-generated if blank" onChange={e => setF({ ...f, Title: e.target.value })} /></div>
        <div><label style={S.label}>Body *</label><textarea style={{ ...S.textarea, minHeight: 140 }} value={f.Body || ""} onChange={e => setF({ ...f, Body: e.target.value })} placeholder="What happened, observations, conversation, agreed next steps…" /></div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Status</label>
            <select style={S.select} value={f.Status || "Open"} onChange={e => setF({ ...f, Status: e.target.value })}>
              {["Open","In Progress","Resolved","Escalated"].map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
          <div><label style={S.label}>Follow-up Date</label><input style={S.input} type="date" value={f.FollowUpDate || ""} onChange={e => setF({ ...f, FollowUpDate: e.target.value })} /></div>
        </div>
        <div><label style={S.label}>Attachment Links (one per line)</label><textarea style={S.textarea} value={f.AttachmentLinks || ""} onChange={e => setF({ ...f, AttachmentLinks: e.target.value })} placeholder="https://vanrockre.sharepoint.com/.../filename.pdf" /></div>
        <div style={{ display: "flex", gap: 18, flexWrap: "wrap", fontSize: 13 }}>
          <label style={{ display: "flex", alignItems: "center", gap: 6 }}><input type="checkbox" checked={!!f.Confidential} onChange={e => setF({ ...f, Confidential: e.target.checked })} /> Confidential</label>
          <label style={{ display: "flex", alignItems: "center", gap: 6 }}><input type="checkbox" checked={!!f.VisibleToManager} onChange={e => setF({ ...f, VisibleToManager: e.target.checked })} /> Visible to manager</label>
          <label style={{ display: "flex", alignItems: "center", gap: 6 }}><input type="checkbox" checked={!!f.VisibleToEmployee} onChange={e => setF({ ...f, VisibleToEmployee: e.target.checked })} /> Visible to employee</label>
        </div>
        <div style={{ fontSize: 11, color: C.b4 }}>Author: <strong>{existing?.AuthorEmail || currentEmail}</strong></div>
      </div>
    </Modal>
  );
}

// ============================================================
// REVIEWS — helpers
// ============================================================
const REVIEW_RATINGS = ["Outstanding", "Exceeds Expectations", "Meets Expectations", "Needs Improvement", "Below Expectations"];
const REVIEW_RATING_TYPE = { "Outstanding": "ok", "Exceeds Expectations": "ok", "Meets Expectations": "inf", "Needs Improvement": "wn", "Below Expectations": "er" };
const REVIEW_STATUSES = ["Scheduled", "In Progress", "Conducted", "Acknowledged", "Cancelled"];
function currentQuarterLabel(date = new Date()) {
  const q = Math.floor(date.getMonth() / 3) + 1;
  return `${date.getFullYear()}-Q${q}`;
}
const REVIEW_NEW_HIRE_MILESTONES = [30, 60, 90]; // days after StartDate for a new hire's first three reviews
const REVIEW_QUARTER_DAYS = 91;                  // cadence once the 90-day review is done
function reviewDateOnly(s) { const d = localDate(s); return d ? ymd(d) : ""; }
function conductedReviews(employeeEmail, reviews) {
  return reviews
    .filter(r => (r.EmployeeEmail || "").toLowerCase() === (employeeEmail || "").toLowerCase() && r.ConductedDate && r.Status !== "Cancelled")
    .sort((a, b) => (b.ConductedDate || "").localeCompare(a.ConductedDate || ""));
}
function lastConductedReview(employeeEmail, reviews) {
  return conductedReviews(employeeEmail, reviews)[0] || null;
}
// Returns { state, daysSinceLast, daysUntilDue, lastReview, nextDueDate }
// State values:
//   "ok"          — within cadence
//   "due-soon"    — within 14 days of due
//   "overdue"     — past due (negative daysUntilDue)
//   "not-started" — StartDate is in the future
//   "onboarding"  — actively onboarding; review tracking is handled by the
//                   onboarding template's Day-90 task, so this tab skips them
//   "never"       — no StartDate AND no reviews, can't schedule (data cleanup)
//
// First-cycle behaviour: until an employee has one review on file AND they're past
// the 90-day onboarding window, the next-due date snaps to CONFIG.firstReviewDueDate
// (the org-wide kickoff date). This avoids retroactively flagging long-tenured
// employees as "X years overdue" when reviews are first rolled out.
function reviewStatusFor(employee, reviews, journeys = []) {
  const email = (employee?.Email || "").toLowerCase();
  const startIso = reviewDateOnly(employee?.StartDate);
  const done = conductedReviews(email, reviews);
  const last = done[0] || null;
  const n = done.length;
  const daysSinceLast = last ? Math.floor((Date.now() - new Date(last.ConductedDate).getTime()) / 86400000) : null;

  // Exempt from in-app review tracking (e.g. ownership-level roles).
  if ((CONFIG.reviewExemptEmails || []).map(e => e.toLowerCase()).includes(email)) {
    return { state: "exempt", lastReview: last, daysSinceLast, nextDueDate: null, daysUntilDue: null };
  }

  // Actively onboarding — reviews tracked inside their onboarding journey, not here.
  const inOnboarding = (journeys || []).some(j =>
    (j.EmployeeEmail || "").toLowerCase() === email
    && j.JourneyType === "Onboarding"
    && j.Status !== "Complete" && j.Status !== "Cancelled");
  if (inOnboarding) return { state: "onboarding", lastReview: last, daysSinceLast, nextDueDate: null, daysUntilDue: null };

  // Hasn't started yet → nothing is due; never overdue.
  if (startIso && daysFromNow(startIso) > 0) {
    const firstDue = addDays(startIso, REVIEW_NEW_HIRE_MILESTONES[0]);
    return { state: "not-started", lastReview: last, daysSinceLast, nextDueDate: firstDue, daysUntilDue: daysFromNow(firstDue) };
  }

  // Default schedule: 30/60/90 from start for first three, then quarterly from last review.
  let nextDue = null;
  if (n < REVIEW_NEW_HIRE_MILESTONES.length) {
    if (startIso) nextDue = addDays(startIso, REVIEW_NEW_HIRE_MILESTONES[n]);
    else if (last) nextDue = addDays(reviewDateOnly(last.ConductedDate), REVIEW_QUARTER_DAYS);
  } else if (last) {
    nextDue = addDays(reviewDateOnly(last.ConductedDate), REVIEW_QUARTER_DAYS);
  }

  // First-cycle override: established employee (past 90-day window) with no reviews →
  // use the org-wide first-cycle date instead of a retroactive 30-day milestone.
  if (!last && CONFIG.firstReviewDueDate) {
    const ninetyDay = startIso ? addDays(startIso, REVIEW_NEW_HIRE_MILESTONES[REVIEW_NEW_HIRE_MILESTONES.length - 1]) : null;
    const past90 = ninetyDay ? daysFromNow(ninetyDay) < 0 : true; // no startIso = treat as established
    if (past90) nextDue = CONFIG.firstReviewDueDate;
  }

  if (!nextDue) return { state: "never", lastReview: last, daysSinceLast, nextDueDate: null, daysUntilDue: null };

  const daysUntilDue = daysFromNow(nextDue);
  let st = "ok";
  if (daysUntilDue < 0) st = "overdue";
  else if (daysUntilDue <= 14) st = "due-soon";
  return { state: st, lastReview: last, daysSinceLast, nextDueDate: nextDue, daysUntilDue };
}
function fmtMoney(n) { if (n == null || isNaN(n)) return "—"; return "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

// ============================================================
// NewShire University training-compliance helper
// Returns { complete, inProgress, overdue, dueSoon, upcoming, total, overall,
//           overdueItems[], dueSoonItems[], recentPasses[] }
// where `overall` is one of "ok" | "due-soon" | "overdue" | "none".
// `none` means the employee has no assignments on file in NewShire University.
// ============================================================
function getTrainingComplianceFor(employeeEmail, assignments, completions, courses) {
  const email = (employeeEmail || "").toLowerCase();
  const myAssignments = (assignments || []).filter(a => a.EmployeeEmail === email);
  if (myAssignments.length === 0) {
    return { complete: 0, inProgress: 0, overdue: 0, dueSoon: 0, upcoming: 0, total: 0, overall: "none", overdueItems: [], dueSoonItems: [], recentPasses: [] };
  }
  const myCompletions = (completions || []).filter(c => c.EmployeeEmail === email);
  const passedByCourse = {};
  for (const c of myCompletions) {
    if (c.Status === "Passed") {
      const prev = passedByCourse[c.CourseId];
      if (!prev || c.CompletedDate > prev.CompletedDate) passedByCourse[c.CourseId] = c;
    }
  }
  const courseNameById = {};
  for (const c of courses || []) courseNameById[String(c.id)] = c.Title || c.CourseName || c.CourseCode || `Course ${c.id}`;

  const overdueItems = [], dueSoonItems = [];
  let complete = 0, inProgress = 0, overdue = 0, dueSoon = 0, upcoming = 0;
  for (const a of myAssignments) {
    const passed = passedByCourse[a.CourseId];
    if (passed) { complete++; continue; }
    inProgress++;
    const du = daysFromNow(a.DueDate);
    const nameAndDate = { courseName: courseNameById[a.CourseId] || `Course ${a.CourseId}`, dueDate: a.DueDate, daysUntilDue: du };
    if (du !== null && du < 0) { overdue++; overdueItems.push(nameAndDate); }
    else if (du !== null && du <= 14) { dueSoon++; dueSoonItems.push(nameAndDate); }
    else { upcoming++; }
  }
  overdueItems.sort((x, y) => (x.dueDate || "").localeCompare(y.dueDate || ""));
  dueSoonItems.sort((x, y) => (x.dueDate || "").localeCompare(y.dueDate || ""));
  const recentPasses = myCompletions.filter(c => c.Status === "Passed").sort((x, y) => (y.CompletedDate || "").localeCompare(x.CompletedDate || "")).slice(0, 3)
    .map(c => ({ courseName: courseNameById[c.CourseId] || `Course ${c.CourseId}`, completedDate: c.CompletedDate, score: c.Score }));
  const overall = overdue > 0 ? "overdue" : dueSoon > 0 ? "due-soon" : "ok";
  return { complete, inProgress, overdue, dueSoon, upcoming, total: myAssignments.length, overall, overdueItems, dueSoonItems, recentPasses };
}

// ============================================================
// REVIEW EDIT MODAL
// ============================================================
function ReviewEditModal({ reviewId, forEmail, onClose }) {
  const { state, actions, currentEmail } = useData();
  const existing = reviewId ? state.reviews.find(r => String(r.id) === String(reviewId)) : null;
  const [f, setF] = useState(() => existing ? { ...existing } : {
    EmployeeEmail: forEmail || "",
    ReviewPeriod: currentQuarterLabel(),
    DueDate: todayIso(),
    ConductedDate: "",
    ReviewerEmail: currentEmail,
    Rating: "Meets Expectations",
    Status: "Scheduled",
    Strengths: "",
    GrowthAreas: "",
    GoalsNextQuarter: "",
    EmployeeComments: "",
    AttachmentLinks: "",
    Confidential: true,
    AcknowledgedDate: "",
    Title: "",
  });
  const [saving, setSaving] = useState(false);
  const isEdit = !!existing;

  const save = async () => {
    if (!f.EmployeeEmail) return alert("Employee email is required.");
    if (!f.ReviewPeriod) return alert("Review period is required.");
    setSaving(true);
    try {
      const payload = { ...f, Title: f.Title || `${f.ReviewPeriod} — ${f.EmployeeEmail}` };
      if (isEdit) await actions.updateReview(reviewId, payload);
      else await actions.createReview(payload);
      onClose();
    } catch (e) { alert("Save failed: " + e.message); } finally { setSaving(false); }
  };
  const del = async () => {
    if (!confirm("Delete this review? Cannot be undone.")) return;
    setSaving(true);
    try { await actions.deleteReview(reviewId); onClose(); } catch (e) { alert("Delete failed: " + e.message); } finally { setSaving(false); }
  };

  return (
    <Modal title={isEdit ? "Edit Quarterly Review" : "New Quarterly Review"} width={780} onClose={onClose} footer={
      <>
        {isEdit && <button style={S.btnO(C.er, C.er)} onClick={del} disabled={saving}>Delete</button>}
        <button style={S.btnO()} onClick={onClose} disabled={saving}>Cancel</button>
        <button style={S.btn(C.hdr)} onClick={save} disabled={saving}>{saving ? "Saving…" : "Save"}</button>
      </>
    }>
      <div style={{ display: "grid", gap: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Review Period *</label><input style={S.input} value={f.ReviewPeriod} placeholder="2026-Q2" onChange={e => setF({ ...f, ReviewPeriod: e.target.value })} /></div>
          <div><label style={S.label}>Status</label><select style={S.select} value={f.Status} onChange={e => setF({ ...f, Status: e.target.value })}>{REVIEW_STATUSES.map(s => <option key={s}>{s}</option>)}</select></div>
          <div><label style={S.label}>Rating</label><select style={S.select} value={f.Rating} onChange={e => setF({ ...f, Rating: e.target.value })}>{REVIEW_RATINGS.map(s => <option key={s}>{s}</option>)}</select></div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Due Date</label><input style={S.input} type="date" value={f.DueDate || ""} onChange={e => setF({ ...f, DueDate: e.target.value })} /></div>
          <div><label style={S.label}>Conducted Date</label><input style={S.input} type="date" value={f.ConductedDate || ""} onChange={e => setF({ ...f, ConductedDate: e.target.value })} /></div>
          <div><label style={S.label}>Reviewer Email</label><input style={S.input} value={f.ReviewerEmail || ""} onChange={e => setF({ ...f, ReviewerEmail: e.target.value.toLowerCase() })} /></div>
        </div>
        <div><label style={S.label}>Strengths</label><textarea style={S.textarea} value={f.Strengths || ""} onChange={e => setF({ ...f, Strengths: e.target.value })} placeholder="What this employee is doing well…" /></div>
        <div><label style={S.label}>Growth Areas</label><textarea style={S.textarea} value={f.GrowthAreas || ""} onChange={e => setF({ ...f, GrowthAreas: e.target.value })} placeholder="Where to focus development…" /></div>
        <div><label style={S.label}>Goals for Next Quarter</label><textarea style={S.textarea} value={f.GoalsNextQuarter || ""} onChange={e => setF({ ...f, GoalsNextQuarter: e.target.value })} placeholder="Concrete, measurable objectives…" /></div>
        <div><label style={S.label}>Employee Comments</label><textarea style={S.textarea} value={f.EmployeeComments || ""} onChange={e => setF({ ...f, EmployeeComments: e.target.value })} placeholder="Employee's response / acknowledgement notes…" /></div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Acknowledged Date</label><input style={S.input} type="date" value={f.AcknowledgedDate || ""} onChange={e => setF({ ...f, AcknowledgedDate: e.target.value })} /></div>
          <div style={{ alignSelf: "end", display: "flex", alignItems: "center", gap: 6, fontSize: 13, paddingBottom: 8 }}>
            <input type="checkbox" checked={!!f.Confidential} onChange={e => setF({ ...f, Confidential: e.target.checked })} /> Confidential
          </div>
        </div>
        <div><label style={S.label}>Attachment Links (one per line)</label><textarea style={S.textarea} value={f.AttachmentLinks || ""} onChange={e => setF({ ...f, AttachmentLinks: e.target.value })} placeholder="https://vanrockre.sharepoint.com/.../signed-review.pdf" /></div>
      </div>
    </Modal>
  );
}

// ============================================================
// PAY CHANGE EDIT MODAL
// ============================================================
const PAY_TYPES = ["Hourly", "Salary", "Salary + Commission", "Commission Only", "1099 Contract", "Other"];
const PAY_CHANGE_TYPES = ["Initial Hire", "Merit Increase", "Promotion", "Cost of Living", "Market Adjustment", "Schedule Change", "Demotion", "Other"];
function PayChangeEditModal({ payId, forEmail, onClose }) {
  const { state, actions, currentEmail } = useData();
  const existing = payId ? state.payChanges.find(p => String(p.id) === String(payId)) : null;
  // Pre-fill PreviousPay with the most recent NewPay for this employee
  const priorPay = useMemo(() => {
    const email = (existing?.EmployeeEmail || forEmail || "").toLowerCase();
    if (!email) return null;
    const history = state.payChanges
      .filter(p => (p.EmployeeEmail || "").toLowerCase() === email && (!existing || String(p.id) !== String(existing.id)))
      .sort((a, b) => (b.EffectiveDate || "").localeCompare(a.EffectiveDate || ""));
    return history[0] ? history[0].NewPay : null;
  }, [state.payChanges, forEmail, existing]);

  const [f, setF] = useState(() => existing ? { ...existing } : {
    EmployeeEmail: forEmail || "",
    EffectiveDate: todayIso(),
    PreviousPay: priorPay ?? "",
    NewPay: "",
    PayType: "Hourly",
    ChangeType: "Merit Increase",
    ApprovedBy: currentEmail,
    Reason: "",
    RelatedReviewId: "",
    Confidential: true,
    Title: "",
  });
  const [saving, setSaving] = useState(false);
  const isEdit = !!existing;

  const prev = parseFloat(f.PreviousPay);
  const next = parseFloat(f.NewPay);
  const pct = !isNaN(prev) && !isNaN(next) && prev > 0 ? ((next - prev) / prev) * 100 : null;
  const delta = !isNaN(prev) && !isNaN(next) ? next - prev : null;

  const save = async () => {
    if (!f.EmployeeEmail) return alert("Employee email is required.");
    if (!f.EffectiveDate) return alert("Effective date is required.");
    if (f.NewPay === "" || isNaN(parseFloat(f.NewPay))) return alert("New pay is required.");
    setSaving(true);
    try {
      const payload = {
        ...f,
        PreviousPay: f.PreviousPay === "" ? 0 : parseFloat(f.PreviousPay),
        NewPay: parseFloat(f.NewPay),
        Title: f.Title || `${f.ChangeType} — ${f.EmployeeEmail} (${fmtDate(f.EffectiveDate)})`,
      };
      if (isEdit) await actions.updatePayChange(payId, payload);
      else await actions.createPayChange(payload);
      onClose();
    } catch (e) { alert("Save failed: " + e.message); } finally { setSaving(false); }
  };
  const del = async () => {
    if (!confirm("Delete this pay change record? Cannot be undone.")) return;
    setSaving(true);
    try { await actions.deletePayChange(payId); onClose(); } catch (e) { alert("Delete failed: " + e.message); } finally { setSaving(false); }
  };

  // Reviews for this employee, for the optional review link
  const empReviews = state.reviews
    .filter(r => (r.EmployeeEmail || "").toLowerCase() === (f.EmployeeEmail || "").toLowerCase())
    .sort((a, b) => (b.ConductedDate || b.DueDate || "").localeCompare(a.ConductedDate || a.DueDate || ""));

  return (
    <Modal title={isEdit ? "Edit Pay Change" : "New Pay Change"} width={720} onClose={onClose} footer={
      <>
        {isEdit && <button style={S.btnO(C.er, C.er)} onClick={del} disabled={saving}>Delete</button>}
        <button style={S.btnO()} onClick={onClose} disabled={saving}>Cancel</button>
        <button style={S.btn(C.hdr)} onClick={save} disabled={saving}>{saving ? "Saving…" : "Save"}</button>
      </>
    }>
      <div style={{ display: "grid", gap: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Effective Date *</label><input style={S.input} type="date" value={f.EffectiveDate || ""} onChange={e => setF({ ...f, EffectiveDate: e.target.value })} /></div>
          <div><label style={S.label}>Pay Type</label><select style={S.select} value={f.PayType} onChange={e => setF({ ...f, PayType: e.target.value })}>{PAY_TYPES.map(t => <option key={t}>{t}</option>)}</select></div>
          <div><label style={S.label}>Change Type</label><select style={S.select} value={f.ChangeType} onChange={e => setF({ ...f, ChangeType: e.target.value })}>{PAY_CHANGE_TYPES.map(t => <option key={t}>{t}</option>)}</select></div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Previous Pay</label><input style={S.input} type="number" step="0.01" value={f.PreviousPay} onChange={e => setF({ ...f, PreviousPay: e.target.value })} /></div>
          <div><label style={S.label}>New Pay *</label><input style={S.input} type="number" step="0.01" value={f.NewPay} onChange={e => setF({ ...f, NewPay: e.target.value })} /></div>
          <div style={{ alignSelf: "end", padding: "8px 10px", border: `1px solid ${C.b1}`, borderRadius: 4, background: C.t0, fontSize: 12 }}>
            {pct !== null ? (
              <div>
                <div style={{ fontWeight: 700, color: pct > 0 ? C.ok : pct < 0 ? C.er : C.b4, fontFamily: mono }}>{pct > 0 ? "+" : ""}{pct.toFixed(1)}%</div>
                <div style={{ color: C.b4, fontSize: 11 }}>{delta > 0 ? "+" : ""}{fmtMoney(Math.abs(delta))}{delta < 0 ? " ↓" : delta > 0 ? " ↑" : ""}</div>
              </div>
            ) : <span style={{ color: C.b4 }}>—</span>}
          </div>
        </div>
        <div><label style={S.label}>Approved By</label><input style={S.input} value={f.ApprovedBy || ""} onChange={e => setF({ ...f, ApprovedBy: e.target.value.toLowerCase() })} /></div>
        <div><label style={S.label}>Reason / Justification</label><textarea style={S.textarea} value={f.Reason || ""} onChange={e => setF({ ...f, Reason: e.target.value })} placeholder="Performance, role change, market correction…" /></div>
        {empReviews.length > 0 && (
          <div><label style={S.label}>Linked Quarterly Review (optional)</label>
            <select style={S.select} value={f.RelatedReviewId || ""} onChange={e => setF({ ...f, RelatedReviewId: e.target.value })}>
              <option value="">— None —</option>
              {empReviews.map(r => <option key={r.id} value={r.id}>{r.ReviewPeriod} · {r.Rating} · {fmtDate(r.ConductedDate || r.DueDate)}</option>)}
            </select>
          </div>
        )}
        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13 }}><input type="checkbox" checked={!!f.Confidential} onChange={e => setF({ ...f, Confidential: e.target.checked })} /> Confidential</label>
      </div>
    </Modal>
  );
}

// ============================================================
// EMAIL TEMPLATES — variable substitution, default library, send modal
// ============================================================
const EMAIL_TEMPLATE_CATEGORIES = ["Welcome", "Onboarding", "Reminder", "Performance", "Offboarding", "Celebration", "Other"];

// Resolves {{Vars}} using the employee record + manager + a few fixed values.
// Unrecognised tokens are left as-is so authors can see what didn't match.
function emailVars({ emp, manager, sender, journey }) {
  const today = new Date();
  const fullName = emp?.Title || emp?.Name || "";
  const firstName = (fullName || "").split(/\s+/)[0] || "";
  const lastName  = (fullName || "").split(/\s+/).slice(1).join(" ") || "";
  return {
    FirstName: firstName,
    LastName: lastName,
    FullName: fullName,
    PreferredName: emp?.PreferredName || firstName,
    WorkEmail: emp?.Email || "",
    PersonalEmail: emp?.PersonalEmail || "",
    JobTitle: emp?.JobTitle || "",
    Department: emp?.Department || "",
    StartDate: emp?.StartDate ? fmtDate(emp.StartDate) : "",
    EndDate: emp?.EndDate ? fmtDate(emp.EndDate) : "",
    ManagerName: manager?.Title || manager?.Name || "",
    ManagerEmail: manager?.Email || emp?.ManagerEmail || "",
    Level2ManagerName: emp?.Level2ManagerName || "",
    Level2ManagerEmail: emp?.Level2ManagerEmail || "",
    Level3ManagerName: emp?.Level3ManagerName || "",
    Level3ManagerEmail: emp?.Level3ManagerEmail || "",
    SenderName: sender?.Title || sender?.Name || sender?.Email || "",
    SenderEmail: sender?.Email || "",
    Today: fmtDate(today.toISOString().slice(0, 10)),
    JourneyType: journey?.JourneyType || "",
  };
}
function applyVars(text, vars) {
  if (!text) return "";
  return String(text).replace(/\{\{\s*([A-Za-z][A-Za-z0-9_]*)\s*\}\}/g, (m, k) => (vars[k] != null && vars[k] !== "" ? vars[k] : m));
}
// Detect any unsubstituted {{Vars}} so the SendEmailModal can warn before sending.
function unresolvedVars(text) {
  if (!text) return [];
  return Array.from(new Set((String(text).match(/\{\{\s*[A-Za-z][A-Za-z0-9_]*\s*\}\}/g) || []).map(s => s.replace(/[{} ]/g, ""))));
}

// Compact starter library — seeded by an admin from Setup -> Email Templates.
const DEFAULT_EMAIL_TEMPLATES = [
  { Title: "Welcome — Day 1", Category: "Welcome", DefaultTo: "Work,Personal", DefaultCc: "{{ManagerEmail}}", Subject: "Welcome to NewShire, {{FirstName}}!", Body: "Hi {{FirstName}},\n\nWelcome to NewShire! Your first day is {{StartDate}}. Your manager, {{ManagerName}}, will greet you and walk you through Day 1 logistics.\n\nA few links to keep handy:\n• Employee Lifecycle app (where you'll find your onboarding tasks)\n• NewShire University (training)\n\nReach out to {{SenderName}} or {{ManagerName}} if anything comes up before then.\n\nWelcome aboard,\n{{SenderName}}" },
  { Title: "Pre-Start — confirm Day 1 logistics", Category: "Onboarding", DefaultTo: "Personal", DefaultCc: "", Subject: "Day 1 logistics — your start at NewShire on {{StartDate}}", Body: "Hi {{FirstName}},\n\nA quick note to confirm logistics for your first day:\n\n• Start date: {{StartDate}}\n• Start time: 9:00 AM\n• Location: 333 Wade Hampton Blvd\n• Who to ask for: {{ManagerName}}\n• Parking: visitor lot, we'll get you a badge\n• Bring: I-9 ID documents (passport OR driver's license + SSN card/birth certificate)\n• Dress: business casual\n\nReply if you have any questions.\n\nLooking forward to it,\n{{SenderName}}" },
  { Title: "30-day check-in reminder", Category: "Reminder", DefaultTo: "Work", DefaultCc: "{{ManagerEmail}}", Subject: "30-day check-in coming up — {{FullName}}", Body: "{{ManagerName}},\n\nThis is a reminder that {{FullName}}'s 30-day check-in is coming up. Please schedule a 30-minute conversation covering:\n\n• What's working / what's unclear\n• Where they're stuck\n• Adjustments to their training plan\n\nThanks,\n{{SenderName}}" },
  { Title: "90-day review notice", Category: "Performance", DefaultTo: "Work", DefaultCc: "{{ManagerEmail}}", Subject: "90-day review due — {{FullName}}", Body: "{{ManagerName}},\n\nIt's time for {{FullName}}'s 90-day review and probation decision. Please complete the review in the Employee Lifecycle app and confirm whether they pass probation.\n\nBenefits enrollment is also due now — HR will reach out separately.\n\nThanks,\n{{SenderName}}" },
  { Title: "Resignation acknowledgement", Category: "Offboarding", DefaultTo: "Work,Personal", DefaultCc: "{{ManagerEmail}}", Subject: "Confirming your resignation — last day {{EndDate}}", Body: "Hi {{FirstName}},\n\nWe've received your resignation and confirmed your last working day as {{EndDate}}.\n\nIn the coming weeks we'll walk through:\n• Knowledge transfer and successor coverage\n• Final timesheet + expense reports\n• Equipment return\n• Exit interview\n\nYou'll see these as tasks in the Employee Lifecycle app. Reach out anytime.\n\nThanks for your contributions, {{FirstName}}.\n\n{{SenderName}}" },
  { Title: "Birthday", Category: "Celebration", DefaultTo: "Work", DefaultCc: "", Subject: "Happy birthday, {{FirstName}}! 🎂", Body: "Hi {{FirstName}},\n\nThe whole NewShire team is wishing you a happy birthday today. Hope you get to celebrate with the people you love.\n\nCheers,\n{{SenderName}}" },
  { Title: "Anniversary — one year", Category: "Celebration", DefaultTo: "Work", DefaultCc: "{{ManagerEmail}}", Subject: "One year at NewShire — congrats, {{FirstName}}!", Body: "Hi {{FirstName}},\n\nIt's been a year since you joined NewShire on {{StartDate}}. Thank you for everything you've done in your time here — we're glad you're with us.\n\nLooking forward to many more,\n{{SenderName}}" },
];

// ============================================================
// SEND EMAIL MODAL — compose from a template, pick recipients (work / personal / manager) + CC
// ============================================================
function SendEmailModal({ forEmail, onClose }) {
  const { state, actions, currentEmail } = useData();
  const emp = state.employees.find(e => (e.Email || "").toLowerCase() === (forEmail || "").toLowerCase());
  const manager = emp ? state.employees.find(e => (e.Email || "").toLowerCase() === (emp.ManagerEmail || "").toLowerCase()) : null;
  const sender = state.employees.find(e => (e.Email || "").toLowerCase() === (currentEmail || "").toLowerCase()) || { Email: currentEmail };

  const [tplId, setTplId] = useState("");
  const [toWork, setToWork] = useState(true);
  const [toPersonal, setToPersonal] = useState(false);
  const [toManager, setToManager] = useState(false);
  const [toL2, setToL2] = useState(false);
  const [toL3, setToL3] = useState(false);
  const [extraTo, setExtraTo] = useState("");
  const [cc, setCc] = useState("");
  const [subject, setSubject] = useState("");
  const [body, setBody] = useState("");
  const [sending, setSending] = useState(false);
  const [warn, setWarn] = useState("");

  const vars = useMemo(() => emailVars({ emp, manager, sender }), [emp, manager, sender]);
  const renderedSubject = applyVars(subject, vars);
  const renderedBody = applyVars(body, vars);
  const unresolved = uniq([...unresolvedVars(renderedSubject), ...unresolvedVars(renderedBody)]);

  // Load a template by id — uses defaults for To/Cc the first time.
  const loadTemplate = (id) => {
    setTplId(id);
    setWarn("");
    if (!id) { setSubject(""); setBody(""); return; }
    const t = state.emailTemplates.find(x => String(x.id) === String(id));
    if (!t) return;
    setSubject(t.Subject || "");
    setBody(t.Body || "");
    const dt = (t.DefaultTo || "").split(",").map(s => s.trim()).filter(Boolean);
    setToWork(dt.includes("Work"));
    setToPersonal(dt.includes("Personal"));
    setToManager(dt.includes("Manager") || /\{\{\s*ManagerEmail\s*\}\}/.test(t.DefaultCc || ""));
    setToL2(dt.includes("Level2Manager") || /\{\{\s*Level2ManagerEmail\s*\}\}/.test(t.DefaultCc || ""));
    setToL3(dt.includes("Level3Manager") || /\{\{\s*Level3ManagerEmail\s*\}\}/.test(t.DefaultCc || ""));
    if (t.DefaultCc) setCc(applyVars(t.DefaultCc, vars));
  };

  if (!emp) return <Modal title="Send Email" onClose={onClose}><Empty title="Employee not found" sub={forEmail} /></Modal>;

  const recipients = useMemo(() => {
    const out = new Set();
    if (toWork && emp.Email) out.add(emp.Email.toLowerCase());
    if (toPersonal && emp.PersonalEmail) out.add(emp.PersonalEmail.toLowerCase());
    if (toManager) {
      const m = (manager?.Email || emp.ManagerEmail || "").toLowerCase();
      if (m) out.add(m);
    }
    if (toL2 && emp.Level2ManagerEmail) out.add(emp.Level2ManagerEmail.toLowerCase());
    if (toL3 && emp.Level3ManagerEmail) out.add(emp.Level3ManagerEmail.toLowerCase());
    extraTo.split(/[,;]/).map(s => s.trim().toLowerCase()).filter(Boolean).forEach(a => out.add(a));
    return Array.from(out);
  }, [toWork, toPersonal, toManager, toL2, toL3, extraTo, emp, manager]);
  const ccList = useMemo(() => {
    return Array.from(new Set(applyVars(cc, vars).split(/[,;\n]/).map(s => s.trim().toLowerCase()).filter(Boolean)));
  }, [cc, vars]);

  const send = async () => {
    setWarn("");
    if (recipients.length === 0) return setWarn("Pick at least one recipient.");
    if (!renderedSubject.trim()) return setWarn("Subject is required.");
    if (toPersonal && !emp.PersonalEmail) return setWarn(`No personal email on file for ${emp.Title || emp.Email}. Add one on the Profile tab first, or uncheck "Personal email".`);
    if (toManager && !(manager?.Email || emp.ManagerEmail)) return setWarn("No manager email on file for this employee.");
    if (toL2 && !emp.Level2ManagerEmail) return setWarn(`No 2nd-level manager email on file for ${emp.Title || emp.Email}.`);
    if (toL3 && !emp.Level3ManagerEmail) return setWarn(`No 3rd-level manager email on file for ${emp.Title || emp.Email}.`);
    if (unresolved.length > 0 && !confirm(`These variables didn't resolve and will send as-is: ${unresolved.join(", ")}. Send anyway?`)) return;

    setSending(true);
    try {
      const bodyHtml = renderedBody.split("\n").map(line => line.length ? `<div>${line.replace(/</g, "&lt;")}</div>` : "<div>&nbsp;</div>").join("");
      await actions.sendEmail({ to: recipients, cc: ccList, subject: renderedSubject, bodyHtml });
      // Stamp the employee record with a "last contacted" timestamp for visibility.
      try { await actions.upsertEmployee({ Email: emp.Email, LastContactedAt: new Date().toISOString(), LastContactedBy: currentEmail }); } catch { /* non-fatal */ }
      onClose();
    } catch (e) { setWarn("Send failed: " + e.message); } finally { setSending(false); }
  };

  return (
    <Modal title={`Send Email — ${emp.Title || emp.Email}`} width={760} onClose={onClose} footer={
      <>
        <button style={S.btnO()} onClick={onClose} disabled={sending}>Cancel</button>
        <button style={S.btn(C.hdr)} onClick={send} disabled={sending}>{sending ? "Sending…" : "Send"}</button>
      </>
    }>
      <div style={{ display: "grid", gap: 12 }}>
        <div>
          <label style={S.label}>Template</label>
          <select style={S.select} value={tplId} onChange={e => loadTemplate(e.target.value)}>
            <option value="">— Blank message —</option>
            {EMAIL_TEMPLATE_CATEGORIES.map(cat => {
              const items = state.emailTemplates.filter(t => (t.Category || "Other") === cat);
              if (items.length === 0) return null;
              return <optgroup key={cat} label={cat}>{items.map(t => <option key={t.id} value={t.id}>{t.Title}</option>)}</optgroup>;
            })}
          </select>
        </div>
        <div>
          <div style={S.label}>To</div>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", fontSize: 13 }}>
            <label style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <input type="checkbox" checked={toWork} onChange={e => setToWork(e.target.checked)} />
              Work — <code style={{ fontFamily: mono, fontSize: 12 }}>{emp.Email || "—"}</code>
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, color: emp.PersonalEmail ? C.t7 : C.b3 }}>
              <input type="checkbox" checked={toPersonal} disabled={!emp.PersonalEmail} onChange={e => setToPersonal(e.target.checked)} />
              Personal — <code style={{ fontFamily: mono, fontSize: 12 }}>{emp.PersonalEmail || "(none on file)"}</code>
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, color: (manager?.Email || emp.ManagerEmail) ? C.t7 : C.b3 }}>
              <input type="checkbox" checked={toManager} disabled={!(manager?.Email || emp.ManagerEmail)} onChange={e => setToManager(e.target.checked)} />
              Manager — <code style={{ fontFamily: mono, fontSize: 12 }}>{manager?.Email || emp.ManagerEmail || "(no manager on file)"}</code>
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, color: emp.Level2ManagerEmail ? C.t7 : C.b3 }}>
              <input type="checkbox" checked={toL2} disabled={!emp.Level2ManagerEmail} onChange={e => setToL2(e.target.checked)} />
              2nd-Level — <code style={{ fontFamily: mono, fontSize: 12 }}>{emp.Level2ManagerEmail || "(none on file)"}</code>{emp.Level2ManagerName ? <span style={{ marginLeft: 4, color: C.b4 }}>({emp.Level2ManagerName})</span> : null}
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, color: emp.Level3ManagerEmail ? C.t7 : C.b3 }}>
              <input type="checkbox" checked={toL3} disabled={!emp.Level3ManagerEmail} onChange={e => setToL3(e.target.checked)} />
              3rd-Level — <code style={{ fontFamily: mono, fontSize: 12 }}>{emp.Level3ManagerEmail || "(none on file)"}</code>{emp.Level3ManagerName ? <span style={{ marginLeft: 4, color: C.b4 }}>({emp.Level3ManagerName})</span> : null}
            </label>
          </div>
          <input style={{ ...S.input, marginTop: 6 }} value={extraTo} placeholder="Other To addresses (comma-separated)" onChange={e => setExtraTo(e.target.value)} />
        </div>
        <div><label style={S.label}>CC (comma-separated; supports {`{{Vars}}`})</label>
          <input style={S.input} list="elc-emp-emails" value={cc} placeholder="cc1@example.com, cc2@example.com" onChange={e => setCc(e.target.value)} />
          <datalist id="elc-emp-emails">{state.employees.filter(e => e.Email).map(e => <option key={e.id} value={e.Email}>{e.Title}</option>)}</datalist>
        </div>
        <div><label style={S.label}>Subject</label><input style={S.input} value={subject} onChange={e => setSubject(e.target.value)} placeholder="Subject line — {{Vars}} are substituted" /></div>
        <div><label style={S.label}>Body</label><textarea style={{ ...S.textarea, minHeight: 220, fontFamily: mono, fontSize: 13 }} value={body} placeholder="Hi {{FirstName}}, …" onChange={e => setBody(e.target.value)} /></div>
        <div style={{ background: C.t0, border: `1px solid ${C.t1}`, borderRadius: 4, padding: 10, fontSize: 12 }}>
          <div style={S.sec}>Preview</div>
          <div style={{ fontSize: 12, color: C.b4 }}>Subject:</div>
          <div style={{ fontWeight: 600, color: C.t7, marginBottom: 6 }}>{renderedSubject || <em style={{ color: C.b4 }}>(empty)</em>}</div>
          <div style={{ fontSize: 12, color: C.b4 }}>Body:</div>
          <div style={{ whiteSpace: "pre-wrap", color: C.t7 }}>{renderedBody || <em style={{ color: C.b4 }}>(empty)</em>}</div>
          {unresolved.length > 0 && <div style={{ marginTop: 8, padding: 6, background: C.wnb, border: `1px solid ${C.wn}33`, borderRadius: 4, fontSize: 11, color: C.wn }}><strong>Unresolved:</strong> {unresolved.join(", ")}</div>}
        </div>
        {warn && <div style={{ background: C.erb, color: C.er, padding: 10, borderRadius: 4, fontSize: 12 }}>{warn}</div>}
        <div style={{ fontSize: 11, color: C.b4 }}>
          Sent from <strong>{currentEmail}</strong> via Outlook. Recipients see this as a regular email from your account; replies come back to you.
        </div>
      </div>
    </Modal>
  );
}

// ============================================================
// EMAIL TEMPLATE EDIT MODAL — admin-edits to the reusable library
// ============================================================
function EmailTemplateEditModal({ tplId, onClose }) {
  const { state, actions } = useData();
  const existing = tplId ? state.emailTemplates.find(t => String(t.id) === String(tplId)) : null;
  const [f, setF] = useState(() => existing ? { ...existing } : {
    Title: "", Category: "Other", Subject: "", Body: "", DefaultTo: "Work", DefaultCc: "", Active: true, Notes: "",
  });
  const [saving, setSaving] = useState(false);
  const isEdit = !!existing;

  const save = async () => {
    if (!f.Title.trim()) return alert("Name is required.");
    if (!f.Subject.trim()) return alert("Subject is required.");
    setSaving(true);
    try {
      if (isEdit) await actions.updateEmailTemplate(tplId, f);
      else await actions.createEmailTemplate(f);
      onClose();
    } catch (e) { alert("Save failed: " + e.message); } finally { setSaving(false); }
  };
  const del = async () => {
    if (!confirm("Delete this template?")) return;
    setSaving(true);
    try { await actions.deleteEmailTemplate(tplId); onClose(); } catch (e) { alert("Delete failed: " + e.message); } finally { setSaving(false); }
  };
  const toFlags = (f.DefaultTo || "").split(",").map(s => s.trim()).filter(Boolean);
  const flipTo = v => {
    const has = toFlags.includes(v);
    const next = has ? toFlags.filter(x => x !== v) : [...toFlags, v];
    setF({ ...f, DefaultTo: next.join(",") });
  };

  return (
    <Modal title={isEdit ? "Edit Email Template" : "New Email Template"} width={720} onClose={onClose} footer={
      <>
        {isEdit && <button style={S.btnO(C.er, C.er)} onClick={del} disabled={saving}>Delete</button>}
        <button style={S.btnO()} onClick={onClose} disabled={saving}>Cancel</button>
        <button style={S.btn(C.hdr)} onClick={save} disabled={saving}>{saving ? "Saving…" : "Save"}</button>
      </>
    }>
      <div style={{ display: "grid", gap: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <div><label style={S.label}>Name *</label><input style={S.input} value={f.Title} onChange={e => setF({ ...f, Title: e.target.value })} /></div>
          <div><label style={S.label}>Category</label><select style={S.select} value={f.Category} onChange={e => setF({ ...f, Category: e.target.value })}>{EMAIL_TEMPLATE_CATEGORIES.map(c => <option key={c}>{c}</option>)}</select></div>
        </div>
        <div>
          <div style={S.label}>Default recipients (checkboxes pre-fill the send modal)</div>
          <div style={{ display: "flex", gap: 14, fontSize: 13, flexWrap: "wrap" }}>
            <label style={{ display: "flex", gap: 6 }}><input type="checkbox" checked={toFlags.includes("Work")} onChange={() => flipTo("Work")} /> Work email</label>
            <label style={{ display: "flex", gap: 6 }}><input type="checkbox" checked={toFlags.includes("Personal")} onChange={() => flipTo("Personal")} /> Personal email</label>
            <label style={{ display: "flex", gap: 6 }}><input type="checkbox" checked={toFlags.includes("Manager")} onChange={() => flipTo("Manager")} /> Manager</label>
            <label style={{ display: "flex", gap: 6 }}><input type="checkbox" checked={toFlags.includes("Level2Manager")} onChange={() => flipTo("Level2Manager")} /> 2nd-Level Manager</label>
            <label style={{ display: "flex", gap: 6 }}><input type="checkbox" checked={toFlags.includes("Level3Manager")} onChange={() => flipTo("Level3Manager")} /> 3rd-Level Manager</label>
          </div>
        </div>
        <div><label style={S.label}>Default CC (literal or {`{{Vars}}`})</label><input style={S.input} value={f.DefaultCc} placeholder="hr@newshirepm.com, {{ManagerEmail}}" onChange={e => setF({ ...f, DefaultCc: e.target.value })} /></div>
        <div><label style={S.label}>Subject *</label><input style={S.input} value={f.Subject} onChange={e => setF({ ...f, Subject: e.target.value })} /></div>
        <div><label style={S.label}>Body</label><textarea style={{ ...S.textarea, minHeight: 220, fontFamily: mono, fontSize: 13 }} value={f.Body} onChange={e => setF({ ...f, Body: e.target.value })} /></div>
        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}><input type="checkbox" checked={f.Active !== false} onChange={e => setF({ ...f, Active: e.target.checked })} /> Active (appears in the Send modal)</label>
        <div style={{ background: C.t0, border: `1px solid ${C.t1}`, borderRadius: 4, padding: 10, fontSize: 11, color: C.b6 }}>
          <strong>Variables:</strong> {`{{FirstName}}`}, {`{{LastName}}`}, {`{{FullName}}`}, {`{{PreferredName}}`}, {`{{WorkEmail}}`}, {`{{PersonalEmail}}`}, {`{{JobTitle}}`}, {`{{Department}}`}, {`{{StartDate}}`}, {`{{EndDate}}`}, {`{{ManagerName}}`}, {`{{ManagerEmail}}`}, {`{{Level2ManagerName}}`}, {`{{Level2ManagerEmail}}`}, {`{{Level3ManagerName}}`}, {`{{Level3ManagerEmail}}`}, {`{{SenderName}}`}, {`{{SenderEmail}}`}, {`{{Today}}`}, {`{{JourneyType}}`}.
        </div>
      </div>
    </Modal>
  );
}

// ============================================================
// EMPLOYEE DETAIL MODAL — Profile, Permissions, Notes, Coaching, Files, Journeys
// ============================================================
function EmployeeDetailModal({ email, onClose }) {
  const { state, actions, role, hasRole, currentEmail } = useData();
  const emp = state.employees.find(e => (e.Email || "").toLowerCase() === email);
  const [sub, setSub] = useState("profile");
  const [saving, setSaving] = useState(false);
  const [permDraft, setPermDraft] = useState({});
  const [permReason, setPermReason] = useState("");
  const [profileDraft, setProfileDraft] = useState(() => emp ? {
    Title: emp.Title || "", JobTitle: emp.JobTitle || "", Email: emp.Email || "",
    PersonalEmail: emp.PersonalEmail || "", ManagerEmail: emp.ManagerEmail || "",
    Department: emp.Department || "", EmployeeActive: emp.EmployeeActive !== false,
    EC1Name: emp.EC1Name || "", EC1Relationship: emp.EC1Relationship || "", EC1Phone: emp.EC1Phone || "", EC1Email: emp.EC1Email || "",
    EC2Name: emp.EC2Name || "", EC2Relationship: emp.EC2Relationship || "", EC2Phone: emp.EC2Phone || "", EC2Email: emp.EC2Email || "",
  } : { EmployeeActive: true });
  const canEdit = hasRole("Admin", "HR");

  const journeys = useMemo(() => emp ? state.journeys.filter(j => (j.EmployeeEmail || "").toLowerCase() === email).sort((a, b) => (b.Modified || "").localeCompare(a.Modified || "")) : [], [emp, state.journeys, email]);
  const notes    = useMemo(() => emp ? state.notes.filter(n => (n.EmployeeEmail || "").toLowerCase() === email).sort((a, b) => (b.NoteDate || "").localeCompare(a.NoteDate || "")) : [], [emp, state.notes, email]);
  const audit    = useMemo(() => emp ? state.audit.filter(a => (a.EmployeeEmail || "").toLowerCase() === email).sort((a, b) => (b.ChangedAt || "").localeCompare(a.ChangedAt || "")) : [], [emp, state.audit, email]);
  const reviews  = useMemo(() => emp ? state.reviews.filter(r => (r.EmployeeEmail || "").toLowerCase() === email).sort((a, b) => (b.ConductedDate || b.DueDate || "").localeCompare(a.ConductedDate || a.DueDate || "")) : [], [emp, state.reviews, email]);
  const pay      = useMemo(() => emp ? state.payChanges.filter(p => (p.EmployeeEmail || "").toLowerCase() === email).sort((a, b) => (b.EffectiveDate || "").localeCompare(a.EffectiveDate || "")) : [], [emp, state.payChanges, email]);
  const reviewStatus = useMemo(() => reviewStatusFor(emp, state.reviews, state.journeys), [state.reviews, state.journeys, emp]);

  // Coaching aggregation — must run unconditionally to satisfy rules-of-hooks
  const coaching = useMemo(() => {
    if (!emp) return { trainingDone: 0, expenseSubmits: 0, vaActivities: 0, overdueTasks: 0 };
    const trainingDone   = 0; // future: pull from TrainingCompletions
    const expenseSubmits = 0; // future: pull from Expense Log
    const vaActivities   = 0; // future: pull from VA_Activity
    const overdueTasks   = state.journeyTasks.filter(t => journeys.some(j => String(j.id) === String(t.JourneyId)) && classifyDue(t.DueDate, t.Status) === "overdue").length;
    return { trainingDone, expenseSubmits, vaActivities, overdueTasks };
  }, [emp, state.journeyTasks, journeys]);

  // Initialise perm draft on email change
  useEffect(() => {
    if (!emp) return;
    const initial = {};
    for (const a of state.apps) initial[a.ColumnName] = emp[a.ColumnName] || "";
    setPermDraft(initial);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [email, emp ? emp.id : null]);

  if (!emp) return (
    <Modal title="Employee" onClose={onClose}><Empty title="Not found" sub={`No employee in the Employees list matches ${email}`} /></Modal>
  );

  const permChanged = Object.entries(permDraft).some(([k, v]) => (emp[k] || "") !== (v || ""));
  const savePerms = async () => {
    setSaving(true);
    try { await actions.setEmployeePermissions(email, permDraft, permReason, null); setPermReason(""); }
    catch (e) { alert("Save failed: " + e.message); }
    finally { setSaving(false); }
  };
  const saveProfile = async () => {
    setSaving(true);
    try {
      await actions.upsertEmployee({ ...profileDraft, Email: emp.Email });
      onClose();
    } catch (e) { alert("Save failed: " + e.message); }
    finally { setSaving(false); }
  };

  const SubTab = ({ k, label, count }) => (
    <button onClick={() => setSub(k)} style={{ ...S.tab(sub === k), padding: "9px 13px", fontSize: 12, minHeight: 38 }}>
      {label}{typeof count === "number" && count > 0 && <span style={{ marginLeft: 6, background: sub === k ? C.g1 : C.b1, color: sub === k ? C.g7 : C.b4, fontSize: 10, fontWeight: 700, padding: "1px 7px", borderRadius: 99 }}>{count}</span>}
    </button>
  );

  return (
    <Modal width={920} title={null} onClose={onClose} footer={null}>
      <div style={{ marginTop: -18, marginLeft: -18, marginRight: -18, padding: "14px 18px", background: `linear-gradient(135deg, ${C.hdr} 0%, ${C.t6} 100%)`, color: "#fff", display: "flex", alignItems: "center", gap: 14 }}>
        <Avatar name={emp.Title} size={48} color={{ bg: C.g5, fg: C.t7 }} />
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 18, fontWeight: 700 }}>{emp.Title || emp.Email}</div>
          <div style={{ fontSize: 12, color: C.t1 }}>{emp.JobTitle || ""} {emp.Department ? `· ${emp.Department}` : ""}</div>
          <div style={{ fontSize: 11, color: C.g4, marginTop: 2 }}>{emp.Email} {emp.ManagerEmail ? `· Manager: ${emp.ManagerEmail}` : ""}</div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 6, alignItems: "flex-end" }}>
          <button style={{ ...S.btnO("#fff", "rgba(255,255,255,.3)"), background: "rgba(255,255,255,.08)", color: "#fff" }} onClick={() => actions.openSendEmail(email)}>✉ Send Email</button>
          <div style={{ display: "flex", gap: 4 }}>
            {emp.EmployeeActive === false ? <Badge type="neutral">Inactive</Badge> : <Badge type="ok">Active</Badge>}
            {journeys[0] && journeys[0].Status !== "Complete" && journeys[0].Status !== "Cancelled" && <Badge type={journeys[0].JourneyType === "Onboarding" ? "inf" : "wn"}>{journeys[0].JourneyType}</Badge>}
          </div>
        </div>
      </div>
      <div style={{ marginLeft: -18, marginRight: -18, borderBottom: `1px solid ${C.b1}`, padding: "0 18px", display: "flex", overflowX: "auto" }}>
        <SubTab k="profile"      label="Profile" />
        <SubTab k="permissions"  label="App Permissions" count={state.apps.filter(a => (emp[a.ColumnName] || "None") !== "None" && emp[a.ColumnName]).length} />
        <SubTab k="reviews"      label="Reviews" count={reviews.length} />
        <SubTab k="compensation" label="Compensation" count={pay.length} />
        <SubTab k="notes"        label="Notes & Discipline" count={notes.length} />
        <SubTab k="coaching"     label="Coaching" />
        <SubTab k="journeys"     label="Journeys" count={journeys.length} />
        <SubTab k="files"        label="Files" />
        {canEdit && <SubTab k="audit" label="Audit" count={audit.length} />}
      </div>
      <div style={{ paddingTop: 16, maxHeight: "calc(100vh - 280px)", overflowY: "auto" }}>
        {sub === "profile" && (
          <div style={{ display: "grid", gap: 12 }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <div><label style={S.label}>Full Name</label><input style={S.input} disabled={!canEdit} value={profileDraft.Title || ""} onChange={e => setProfileDraft({ ...profileDraft, Title: e.target.value })} /></div>
              <div><label style={S.label}>Work Email</label><input style={S.input} disabled value={profileDraft.Email || ""} /></div>
              <div><label style={S.label}>Personal Email</label><input style={S.input} disabled={!canEdit} value={profileDraft.PersonalEmail || ""} placeholder="optional" onChange={e => setProfileDraft({ ...profileDraft, PersonalEmail: e.target.value.toLowerCase() })} /></div>
              <div><label style={S.label}>Job Title</label><input style={S.input} disabled={!canEdit} value={profileDraft.JobTitle || ""} onChange={e => setProfileDraft({ ...profileDraft, JobTitle: e.target.value })} /></div>
              <div><label style={S.label}>Department</label><input style={S.input} disabled={!canEdit} value={profileDraft.Department || ""} onChange={e => setProfileDraft({ ...profileDraft, Department: e.target.value })} /></div>
              <div><label style={S.label}>Manager Email</label><input style={S.input} disabled={!canEdit} value={profileDraft.ManagerEmail || ""} onChange={e => setProfileDraft({ ...profileDraft, ManagerEmail: e.target.value.toLowerCase() })} /></div>
              <div><label style={S.label}>Status</label>
                <select style={S.select} disabled={!canEdit} value={profileDraft.EmployeeActive ? "active" : "inactive"} onChange={e => setProfileDraft({ ...profileDraft, EmployeeActive: e.target.value === "active" })}>
                  <option value="active">Active</option><option value="inactive">Inactive</option>
                </select>
              </div>
            </div>
            <div style={{ marginTop: 8 }}>
              <div style={S.sec}>Emergency Contacts</div>
              <div style={{ fontSize: 11, color: C.b4, marginBottom: 10 }}>Used for HR + manager only in the event of emergency. Confidential.</div>
              {[1, 2].map(n => {
                const k = (suf) => `EC${n}${suf}`;
                const label = n === 1 ? "Primary" : "Secondary";
                return (
                  <div key={n} style={{ background: C.t0, border: `1px solid ${C.t1}`, borderRadius: 6, padding: 12, marginBottom: 10 }}>
                    <div style={{ fontWeight: 700, fontSize: 12, color: C.t7, marginBottom: 8, textTransform: "uppercase", letterSpacing: ".06em" }}>{label}{n === 2 && " (optional)"}</div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                      <div><label style={S.label}>Name</label><input style={S.input} disabled={!canEdit} value={profileDraft[k("Name")] || ""} onChange={e => setProfileDraft({ ...profileDraft, [k("Name")]: e.target.value })} /></div>
                      <div><label style={S.label}>Relationship</label><input style={S.input} list={`ec-relationships-${n}`} disabled={!canEdit} value={profileDraft[k("Relationship")] || ""} placeholder="Spouse, Parent, Sibling…" onChange={e => setProfileDraft({ ...profileDraft, [k("Relationship")]: e.target.value })} />
                        <datalist id={`ec-relationships-${n}`}>
                          {["Spouse","Partner","Parent","Child","Sibling","Friend","Guardian","Other"].map(r => <option key={r} value={r} />)}
                        </datalist>
                      </div>
                      <div><label style={S.label}>Phone</label><input style={S.input} type="tel" disabled={!canEdit} value={profileDraft[k("Phone")] || ""} placeholder="(555) 555-5555" onChange={e => setProfileDraft({ ...profileDraft, [k("Phone")]: e.target.value })} /></div>
                      <div><label style={S.label}>Email <span style={{ color: C.b4 }}>(optional)</span></label><input style={S.input} type="email" disabled={!canEdit} value={profileDraft[k("Email")] || ""} onChange={e => setProfileDraft({ ...profileDraft, [k("Email")]: e.target.value.toLowerCase() })} /></div>
                    </div>
                  </div>
                );
              })}
            </div>
            {canEdit && <div style={{ display: "flex", justifyContent: "flex-end" }}><button style={S.btn(C.hdr)} disabled={saving} onClick={saveProfile}>{saving ? "Saving…" : "Save Profile"}</button></div>}
          </div>
        )}
        {sub === "permissions" && (
          <div style={{ display: "grid", gap: 12 }}>
            {/* One-click separation: revoke all app roles + mark inactive.
                Wraps setEmployeePermissions (so every change is audited) and
                then upsertEmployee for the Active flag. Re-activating only
                flips Active back on — does NOT restore prior roles. */}
            {canEdit && (
              <div style={{ background: emp.EmployeeActive === false ? C.t0 : C.erb, border: `1px solid ${emp.EmployeeActive === false ? C.t1 : C.er + "33"}`, borderRadius: 6, padding: 12, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                <div>
                  <div style={{ fontWeight: 700, color: emp.EmployeeActive === false ? C.t7 : C.er, fontSize: 13 }}>{emp.EmployeeActive === false ? "Employee is INACTIVE" : "Active employee"}</div>
                  <div style={{ fontSize: 11, color: C.b4, marginTop: 2 }}>
                    {emp.EmployeeActive === false
                      ? "All app roles are zeroed out. Re-activate only flips the Active flag — you'll need to re-set permissions manually."
                      : "One-click revoke: sets every app role to None (audit logged per app) and marks the employee inactive."}
                  </div>
                </div>
                {emp.EmployeeActive === false ? (
                  <button style={S.btnO(C.t5)} disabled={saving} onClick={async () => {
                    if (!confirm(`Re-activate ${emp.Title || emp.Email}? This only flips the Active flag — app permissions stay at None and must be re-set manually.`)) return;
                    setSaving(true);
                    try { await actions.upsertEmployee({ Email: emp.Email, EmployeeActive: true }); }
                    catch (e) { alert("Failed: " + e.message); }
                    finally { setSaving(false); }
                  }}>Re-activate</button>
                ) : (
                  <button style={S.btn(C.er)} disabled={saving} onClick={async () => {
                    const live = state.apps.filter(a => (emp[a.ColumnName] || "None") !== "None" && emp[a.ColumnName]);
                    if (!confirm(`Revoke all app access for ${emp.Title || emp.Email} and mark them inactive?\n\nThis will set ${live.length} active role${live.length === 1 ? "" : "s"} to None and flip the employee inactive. Each role change is logged.`)) return;
                    setSaving(true);
                    try {
                      // 1. Zero out every app role through the audited path
                      const allNone = {};
                      for (const a of state.apps) allNone[a.ColumnName] = "";
                      await actions.setEmployeePermissions(email, allNone, "Bulk deactivation", null);
                      // 2. Flip Active off
                      await actions.upsertEmployee({ Email: emp.Email, EmployeeActive: false });
                      // 3. Mirror the cleared draft locally so the matrix below reflects reality
                      const cleared = {}; for (const a of state.apps) cleared[a.ColumnName] = "";
                      setPermDraft(cleared);
                    } catch (e) { alert("Failed: " + e.message); }
                    finally { setSaving(false); }
                  }}>Deactivate & revoke all access</button>
                )}
              </div>
            )}
            <div style={{ background: C.infb, border: `1px solid ${C.inf}33`, borderRadius: 4, padding: 10, fontSize: 12, color: C.inf }}>
              Set the role this employee has in each NewShire app. <strong>None</strong> means they have no access. Changes are saved with an audit entry.
            </div>
            <PermissionMatrix apps={state.apps} values={permDraft} disabled={!canEdit || emp.EmployeeActive === false} onChange={(col, v) => setPermDraft({ ...permDraft, [col]: v })} />
            {canEdit && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 10, alignItems: "end" }}>
                <div><label style={S.label}>Reason for change (optional)</label><input style={S.input} value={permReason} placeholder="e.g. promotion, new responsibility, etc." onChange={e => setPermReason(e.target.value)} /></div>
                <button style={S.btn(C.hdr)} disabled={!permChanged || saving} onClick={savePerms}>{saving ? "Saving…" : "Save Permissions"}</button>
              </div>
            )}
            {audit.length > 0 && (
              <div style={{ marginTop: 8 }}>
                <div style={S.sec}>Recent changes</div>
                {audit.slice(0, 5).map(a => (
                  <div key={a.id} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: `1px solid ${C.b1}`, fontSize: 12 }}>
                    <span><strong>{a.AppKey}</strong>: <span style={{ color: C.b4 }}>{a.OldRole}</span> → <strong style={{ color: C.t7 }}>{a.NewRole}</strong></span>
                    <span style={{ color: C.b4 }}>{a.ChangedBy} · {fmtDate(a.ChangedAt)}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
        {sub === "reviews" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                {reviewStatus.state === "overdue" && <Badge type="er">⚠ Past due {fmtDate(reviewStatus.nextDueDate)} ({Math.abs(reviewStatus.daysUntilDue)}d)</Badge>}
                {reviewStatus.state === "due-soon" && <Badge type="wn">Due {fmtDate(reviewStatus.nextDueDate)}{reviewStatus.lastReview ? "" : " (first review)"}</Badge>}
                {reviewStatus.state === "ok" && <Badge type="ok">{reviewStatus.lastReview ? `Last: ${fmtDate(reviewStatus.lastReview.ConductedDate)} · next ` : "First review "}due {fmtDate(reviewStatus.nextDueDate)}</Badge>}
                {reviewStatus.state === "onboarding" && <Badge type="inf">In onboarding — review tracked by Day-90 task</Badge>}
                {reviewStatus.state === "exempt" && <Badge type="neutral">Exempt — reviewed directly by ownership</Badge>}
                {reviewStatus.state === "never" && <Badge type="neutral">No reviews — add a start date to schedule</Badge>}
                {reviewStatus.state === "not-started" && <Badge type="neutral">Starts {fmtDate(emp.StartDate)} · first review due {fmtDate(reviewStatus.nextDueDate)}</Badge>}
              </div>
              {canEdit && <button style={S.btn(C.hdr)} onClick={() => actions.openReviewNew(email)}>+ New Review</button>}
            </div>
            {(() => {
              const tc = getTrainingComplianceFor(email, state.luAssignments, state.luCompletions, state.luCourses);
              if (tc.overall === "none") return (
                <div style={{ background: C.t0, border: `1px solid ${C.t1}`, borderRadius: 6, padding: 12, marginBottom: 12, fontSize: 12, color: C.b4 }}>
                  <strong style={{ color: C.t7 }}>Training Compliance</strong> — no assignments on file in NewShire University.
                </div>
              );
              const headerColor = tc.overdue > 0 ? C.er : tc.dueSoon > 0 ? C.wn : C.ok;
              return (
                <div style={{ background: C.wh, border: `1px solid ${C.b1}`, borderRadius: 6, padding: 12, marginBottom: 12 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 6 }}>
                    <div style={{ fontWeight: 700, color: C.t7 }}>Training Compliance (NewShire University)</div>
                    <a href={CONFIG.universityUrl} target="_blank" rel="noreferrer" style={{ fontSize: 11, color: C.inf, textDecoration: "none" }}>Open NewShire University ↗</a>
                  </div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
                    <Badge type="ok">{tc.complete} complete</Badge>
                    {tc.inProgress > 0 && <Badge type="inf">{tc.inProgress} in progress</Badge>}
                    {tc.overdue > 0 && <Badge type="er">{tc.overdue} overdue</Badge>}
                    {tc.dueSoon > 0 && <Badge type="wn">{tc.dueSoon} due soon</Badge>}
                    {tc.upcoming > 0 && <Badge type="neutral">{tc.upcoming} upcoming</Badge>}
                  </div>
                  <ProgressBar value={Math.round((tc.complete / Math.max(1, tc.total)) * 100)} color={headerColor} />
                  <div style={{ fontSize: 11, color: C.b4, marginTop: 4 }}>{tc.complete} of {tc.total} assignments passed</div>
                  {tc.overdueItems.length > 0 && (
                    <div style={{ marginTop: 10 }}>
                      <div style={{ ...S.sec, color: C.er, marginBottom: 4 }}>Overdue</div>
                      {tc.overdueItems.map((it, i) => (
                        <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, padding: "3px 0", borderBottom: i < tc.overdueItems.length - 1 ? `1px solid ${C.b1}` : "none" }}>
                          <span>{it.courseName}</span>
                          <span style={{ color: C.er, fontWeight: 600 }}>due {fmtDate(it.dueDate)} · {Math.abs(it.daysUntilDue)}d past</span>
                        </div>
                      ))}
                    </div>
                  )}
                  {tc.dueSoonItems.length > 0 && (
                    <div style={{ marginTop: 10 }}>
                      <div style={{ ...S.sec, color: C.wn, marginBottom: 4 }}>Due Soon</div>
                      {tc.dueSoonItems.map((it, i) => (
                        <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, padding: "3px 0", borderBottom: i < tc.dueSoonItems.length - 1 ? `1px solid ${C.b1}` : "none" }}>
                          <span>{it.courseName}</span>
                          <span style={{ color: C.wn, fontWeight: 600 }}>due {fmtDate(it.dueDate)} · {it.daysUntilDue}d</span>
                        </div>
                      ))}
                    </div>
                  )}
                  {tc.recentPasses.length > 0 && (
                    <div style={{ marginTop: 10 }}>
                      <div style={{ ...S.sec, marginBottom: 4 }}>Recent Passes</div>
                      {tc.recentPasses.map((it, i) => (
                        <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, padding: "3px 0", borderBottom: i < tc.recentPasses.length - 1 ? `1px solid ${C.b1}` : "none" }}>
                          <span>{it.courseName}</span>
                          <span style={{ color: C.b4 }}>{fmtDate(it.completedDate)} · <strong style={{ color: C.ok }}>{it.score}%</strong></span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })()}
            {reviews.length === 0 ? <Empty title="No reviews on file" sub={canEdit ? "Click '+ New Review' to record the first one." : "Nothing on file."} /> : reviews.map(r => (
              <div key={r.id} style={{ background: C.wh, border: `1px solid ${C.b1}`, borderRadius: 6, padding: 12, marginBottom: 10 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    <span style={{ fontWeight: 700, color: C.t7, fontSize: 14 }}>{r.ReviewPeriod || "—"}</span>
                    <Badge type={REVIEW_RATING_TYPE[r.Rating] || "neutral"}>{r.Rating || "Unrated"}</Badge>
                    <Badge type={r.Status === "Acknowledged" || r.Status === "Conducted" ? "ok" : r.Status === "Cancelled" ? "neutral" : "inf"}>{r.Status}</Badge>
                    {r.Confidential && <Badge type="neutral">🔒 Confidential</Badge>}
                  </div>
                  {canEdit && <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => actions.openReviewEdit(r.id)}>Edit</button>}
                </div>
                <div style={{ fontSize: 11, color: C.b4, marginBottom: 8 }}>
                  Conducted {fmtDate(r.ConductedDate)} by {r.ReviewerEmail || "—"} · Due {fmtDate(r.DueDate)}{r.AcknowledgedDate ? ` · Acknowledged ${fmtDate(r.AcknowledgedDate)}` : ""}
                </div>
                {r.Strengths && <div style={{ marginBottom: 6 }}><span style={{ fontSize: 11, fontWeight: 700, color: C.ok, textTransform: "uppercase" }}>Strengths</span><div style={{ fontSize: 13, whiteSpace: "pre-wrap" }}>{r.Strengths}</div></div>}
                {r.GrowthAreas && <div style={{ marginBottom: 6 }}><span style={{ fontSize: 11, fontWeight: 700, color: C.wn, textTransform: "uppercase" }}>Growth Areas</span><div style={{ fontSize: 13, whiteSpace: "pre-wrap" }}>{r.GrowthAreas}</div></div>}
                {r.GoalsNextQuarter && <div style={{ marginBottom: 6 }}><span style={{ fontSize: 11, fontWeight: 700, color: C.inf, textTransform: "uppercase" }}>Goals Next Quarter</span><div style={{ fontSize: 13, whiteSpace: "pre-wrap" }}>{r.GoalsNextQuarter}</div></div>}
                {r.EmployeeComments && <div style={{ marginBottom: 6 }}><span style={{ fontSize: 11, fontWeight: 700, color: C.pu, textTransform: "uppercase" }}>Employee Comments</span><div style={{ fontSize: 13, whiteSpace: "pre-wrap" }}>{r.EmployeeComments}</div></div>}
                {r.AttachmentLinks && (
                  <div style={{ marginTop: 6, fontSize: 11 }}>
                    {r.AttachmentLinks.split(/\r?\n/).filter(Boolean).map((u, i) => <div key={i}><a href={u} target="_blank" rel="noreferrer" style={{ color: C.inf }}>{u}</a></div>)}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
        {sub === "compensation" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
              {(() => {
                const current = pay[0];
                const initial = pay[pay.length - 1];
                const totalChange = current && initial && current !== initial && initial.PreviousPay
                  ? ((current.NewPay - (initial.PreviousPay || initial.NewPay)) / (initial.PreviousPay || initial.NewPay)) * 100
                  : null;
                return (
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {current && <Badge type="inf">Current: {fmtMoney(current.NewPay)} · {current.PayType}</Badge>}
                    {current && <Badge type="neutral">Effective {fmtDate(current.EffectiveDate)}</Badge>}
                    {totalChange !== null && <Badge type={totalChange >= 0 ? "ok" : "er"}>Cumulative {totalChange >= 0 ? "+" : ""}{totalChange.toFixed(1)}%</Badge>}
                  </div>
                );
              })()}
              {canEdit && <button style={S.btn(C.hdr)} onClick={() => actions.openPayNew(email)}>+ Record Pay Change</button>}
            </div>
            {pay.length === 0 ? <Empty title="No compensation history yet" sub={canEdit ? "Click '+ Record Pay Change' to log the initial pay." : "Nothing on file."} /> : (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead><tr><th style={S.th}>Effective</th><th style={S.th}>Change Type</th><th style={S.th}>From</th><th style={S.th}>To</th><th style={S.th}>%</th><th style={S.th}>Type</th><th style={S.th}>Approved By</th><th style={S.th}></th></tr></thead>
                  <tbody>
                    {pay.map(p => {
                      const pct = p.PreviousPay > 0 ? ((p.NewPay - p.PreviousPay) / p.PreviousPay) * 100 : null;
                      return (
                        <tr key={p.id}>
                          <td style={S.td}><div style={{ fontWeight: 600 }}>{fmtDate(p.EffectiveDate)}</div></td>
                          <td style={S.td}><Badge type={p.ChangeType === "Demotion" ? "er" : p.ChangeType === "Promotion" ? "ok" : "neutral"}>{p.ChangeType}</Badge></td>
                          <td style={S.td}>{p.PreviousPay ? fmtMoney(p.PreviousPay) : <span style={{ color: C.b4 }}>—</span>}</td>
                          <td style={S.td}><strong>{fmtMoney(p.NewPay)}</strong></td>
                          <td style={S.td}>{pct !== null ? <span style={{ fontFamily: mono, color: pct > 0 ? C.ok : pct < 0 ? C.er : C.b4 }}>{pct > 0 ? "+" : ""}{pct.toFixed(1)}%</span> : <span style={{ color: C.b4 }}>—</span>}</td>
                          <td style={S.td}>{p.PayType}</td>
                          <td style={S.td}>{p.ApprovedBy || "—"}</td>
                          <td style={S.td}>{canEdit && <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => actions.openPayEdit(p.id)}>Edit</button>}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
            {pay.some(p => p.Reason) && (
              <div style={{ marginTop: 14, fontSize: 12 }}>
                <div style={S.sec}>Reasons / Justifications</div>
                {pay.filter(p => p.Reason).map(p => (
                  <div key={p.id} style={{ padding: "6px 0", borderBottom: `1px solid ${C.b1}` }}>
                    <strong>{fmtDate(p.EffectiveDate)}</strong> · {p.ChangeType} — <span style={{ color: C.b6 }}>{p.Reason}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
        {sub === "notes" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {Object.entries(notes.reduce((m, n) => { m[n.NoteType || "General"] = (m[n.NoteType || "General"] || 0) + 1; return m; }, {})).map(([t, n]) => (
                  <Badge key={t} type={NOTE_TYPE_COLORS[t] || "neutral"}>{t} · {n}</Badge>
                ))}
              </div>
              {canEdit && <button style={S.btn(C.hdr)} onClick={() => actions.openNoteNew(email)}>+ New Note</button>}
            </div>
            {notes.length === 0 ? <Empty title="No notes yet" sub={canEdit ? "Add the first one." : "Nothing on file."} /> : notes.map(n => (
              <div key={n.id} style={{ background: C.wh, border: `1px solid ${C.b1}`, borderRadius: 6, padding: 12, marginBottom: 10 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    <Badge type={NOTE_TYPE_COLORS[n.NoteType] || "neutral"}>{n.NoteType}</Badge>
                    {n.Severity && n.Severity !== "Info" && <Badge type={n.Severity === "Critical" || n.Severity === "High" ? "er" : "wn"}>{n.Severity}</Badge>}
                    {n.Status && n.Status !== "Open" && <Badge type={n.Status === "Resolved" ? "ok" : "neutral"}>{n.Status}</Badge>}
                    {n.Confidential && <Badge type="neutral">🔒 Confidential</Badge>}
                    <span style={{ fontSize: 12, color: C.b4 }}>{fmtDate(n.NoteDate)} · by {n.AuthorEmail}</span>
                  </div>
                  {canEdit && <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => actions.openNoteEdit(n.id)}>Edit</button>}
                </div>
                {n.Title && <div style={{ fontWeight: 600, color: C.t7, marginBottom: 4 }}>{n.Title}</div>}
                <div style={{ whiteSpace: "pre-wrap", fontSize: 13, color: C.t7 }}>{n.Body}</div>
                {n.FollowUpDate && <div style={{ marginTop: 8, fontSize: 11, color: C.wn }}>Follow-up: {fmtDate(n.FollowUpDate)}</div>}
                {n.AttachmentLinks && (
                  <div style={{ marginTop: 8, fontSize: 11 }}>
                    {n.AttachmentLinks.split(/\r?\n/).filter(Boolean).map((u, i) => <div key={i}><a href={u} target="_blank" rel="noreferrer" style={{ color: C.inf }}>{u}</a></div>)}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
        {sub === "coaching" && (
          <div>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
              <div style={S.kpi}><div style={S.kpiL}>Notes (90d)</div><div style={S.kpiV}>{notes.filter(n => n.NoteDate && (Date.now() - new Date(n.NoteDate).getTime()) / 86400000 <= 90).length}</div></div>
              <div style={S.kpi}><div style={S.kpiL}>Open Disciplinary</div><div style={{ ...S.kpiV, color: notes.filter(n => n.NoteType === "Discipline" && n.Status !== "Resolved").length > 0 ? C.er : C.t7 }}>{notes.filter(n => n.NoteType === "Discipline" && n.Status !== "Resolved").length}</div></div>
              <div style={S.kpi}><div style={S.kpiL}>Praise (12mo)</div><div style={{ ...S.kpiV, color: C.ok }}>{notes.filter(n => n.NoteType === "Praise" && n.NoteDate && (Date.now() - new Date(n.NoteDate).getTime()) / 86400000 <= 365).length}</div></div>
              <div style={S.kpi}><div style={S.kpiL}>Overdue Tasks</div><div style={{ ...S.kpiV, color: coaching.overdueTasks > 0 ? C.wn : C.t7 }}>{coaching.overdueTasks}</div></div>
            </div>
            <div style={S.card}>
              <div style={S.cardT}>Cross-app activity</div>
              <div style={{ fontSize: 12, color: C.b4, marginBottom: 8 }}>Aggregated coaching data from other NewShire apps. Each row reads from the source app's SharePoint lists.</div>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr><th style={S.th}>Source</th><th style={S.th}>Metric</th><th style={S.th}>Value</th><th style={S.th}>Last 30d</th></tr></thead>
                <tbody>
                  <tr><td style={S.td}>NewShire University</td><td style={S.td}>Trainings completed</td><td style={S.td}>{coaching.trainingDone}</td><td style={S.td}><span style={{ color: C.b4 }}>—</span></td></tr>
                  <tr><td style={S.td}>VA Tracker</td><td style={S.td}>Activity logged (hrs)</td><td style={S.td}>{coaching.vaActivities}</td><td style={S.td}><span style={{ color: C.b4 }}>—</span></td></tr>
                  <tr><td style={S.td}>Expense Manager</td><td style={S.td}>Reports submitted</td><td style={S.td}>{coaching.expenseSubmits}</td><td style={S.td}><span style={{ color: C.b4 }}>—</span></td></tr>
                  <tr><td style={S.td}>PM Hub</td><td style={S.td}>Help requests opened</td><td style={S.td}><span style={{ color: C.b4 }}>—</span></td><td style={S.td}><span style={{ color: C.b4 }}>—</span></td></tr>
                </tbody>
              </table>
              <div style={{ marginTop: 10, padding: 10, background: C.g0, border: `1px dashed ${C.g4}`, borderRadius: 4, fontSize: 11, color: C.g7 }}>
                <strong>Wiring TODO:</strong> Add per-app list-read functions in <code style={{ fontFamily: mono }}>loadAll()</code> and aggregate by EmployeeEmail. Each app already keys its lists by email, so this is mostly mapping work.
              </div>
            </div>
          </div>
        )}
        {sub === "journeys" && (
          <div>
            {journeys.length === 0 ? <Empty title="No journeys" sub="This employee has no onboarding/offboarding history yet." /> : journeys.map(j => {
              const p = journeyProgress(j, state.journeyTasks);
              return (
                <div key={j.id} style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 12, padding: 12, border: `1px solid ${C.b1}`, borderRadius: 6, marginBottom: 8, alignItems: "center" }}>
                  <div>
                    <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 4 }}>
                      <Badge type={j.JourneyType === "Onboarding" ? "inf" : "wn"}>{j.JourneyType}</Badge>
                      <Badge type={j.Status === "Complete" ? "ok" : j.Status === "Cancelled" ? "neutral" : "inf"}>{j.Status}</Badge>
                    </div>
                    <div style={{ fontSize: 12, color: C.b4 }}>{fmtDate(journeyAnchorDate(j))} · {p.done}/{p.total} tasks ({p.pct}%)</div>
                    <ProgressBar value={p.pct} />
                  </div>
                  <button style={S.btnO(C.t5)} onClick={() => { onClose(); actions.openJourney(j.id); }}>Open</button>
                </div>
              );
            })}
          </div>
        )}
        {sub === "files" && (
          <div>
            <div style={S.card}>
              <div style={S.cardT}>Employee Documents</div>
              <div style={{ fontSize: 13, marginBottom: 10 }}>Files are stored in SharePoint at <code style={{ fontFamily: mono, fontSize: 12 }}>ELC_EmployeeFiles/{email}/</code></div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <a href={`${CONFIG.filesLibraryUrl}/Forms/AllItems.aspx?viewpath=%2FELC%5FEmployeeFiles%2FForms%2FAllItems%2Easpx&FilterField1=EmployeeEmail&FilterValue1=${encodeURIComponent(email)}`} target="_blank" rel="noreferrer" style={S.btn(C.hdr)}>Open in SharePoint</a>
                <a href={`${CONFIG.filesLibraryUrl}/${encodeURIComponent(email)}`} target="_blank" rel="noreferrer" style={S.btnO(C.t5)}>Open Employee Folder</a>
              </div>
              <div style={{ marginTop: 12, fontSize: 11, color: C.b4 }}>
                Typical contents: offer letter, signed handbook, W-4 / I-9, performance reviews, disciplinary docs, exit paperwork. Tag each file with <code>DocCategory</code> in SharePoint.
              </div>
            </div>
          </div>
        )}
        {sub === "audit" && canEdit && (
          <div>
            {audit.length === 0 ? <Empty title="No permission changes recorded yet" sub="Audit entries appear here when you change any app role." /> : (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr><th style={S.th}>When</th><th style={S.th}>App</th><th style={S.th}>Old → New</th><th style={S.th}>By</th><th style={S.th}>Reason</th></tr></thead>
                <tbody>{audit.map(a => (
                  <tr key={a.id}><td style={S.td}>{fmtDate(a.ChangedAt)}</td><td style={S.td}>{a.AppKey}</td><td style={S.td}><span style={{ color: C.b4 }}>{a.OldRole}</span> → <strong>{a.NewRole}</strong></td><td style={S.td}>{a.ChangedBy}</td><td style={S.td}>{a.Reason || "—"}</td></tr>
                ))}</tbody>
              </table>
            )}
          </div>
        )}
      </div>
    </Modal>
  );
}

// ============================================================
// EMPLOYEES TAB (Admin / HR) — click row to open detail
// ============================================================
function EmployeesTab() {
  const { state, actions } = useData();
  const [q, setQ] = useState("");
  const [showInactive, setShowInactive] = useState(false);
  const employees = state.employees
    .filter(e => showInactive || e.EmployeeActive !== false)
    .filter(e => !q || (e.Title || "").toLowerCase().includes(q.toLowerCase()) || (e.Email || "").toLowerCase().includes(q.toLowerCase()) || (e.JobTitle || "").toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => (a.Title || "").localeCompare(b.Title || ""));

  const byEmail = (email) => state.journeys.find(j => (j.EmployeeEmail || "").toLowerCase() === (email || "").toLowerCase() && j.Status !== "Complete" && j.Status !== "Cancelled");
  const permCount = (e) => state.apps.filter(a => (e[a.ColumnName] || "None") !== "None" && e[a.ColumnName]).length;
  const noteCount = (e) => state.notes.filter(n => (n.EmployeeEmail || "").toLowerCase() === (e.Email || "").toLowerCase()).length;

  return (
    <div>
      <div style={S.card}>
        <div style={S.cardT}>
          <span>Employees ({employees.length})</span>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <input style={{ ...S.input, width: 220 }} placeholder="Search…" value={q} onChange={e => setQ(e.target.value)} />
            <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12 }}><input type="checkbox" checked={showInactive} onChange={e => setShowInactive(e.target.checked)} /> Show inactive</label>
          </div>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th style={S.th}>Name</th><th style={S.th}>Title</th><th style={S.th}>Manager</th><th style={S.th}>ELC Role</th><th style={S.th}>Apps</th><th style={S.th}>Notes</th><th style={S.th}>Status</th><th style={S.th}>Active Journey</th></tr></thead>
            <tbody>
              {employees.map(e => {
                const j = byEmail(e.Email);
                const elcRoles = parseRoles(e.ELCRole);
                const roleBadgeType = r => r === "Admin" ? "er" : r === "HR" ? "pu" : r === "IT" ? "inf" : r === "Manager" ? "wn" : "neutral";
                return (
                  <tr key={e.id} style={{ cursor: "pointer" }} onClick={() => actions.openEmployee(e.Email)}
                      onMouseEnter={ev => ev.currentTarget.style.background = C.t0}
                      onMouseLeave={ev => ev.currentTarget.style.background = ""}>
                    <td style={S.td}><div style={{ display: "flex", alignItems: "center", gap: 8 }}><Avatar name={e.Title} size={28} /><div style={{ fontWeight: 600, color: C.t7 }}>{e.Title}</div></div></td>
                    <td style={S.td}>{e.JobTitle || "—"}</td>
                    <td style={S.td}>{e.ManagerEmail || "—"}</td>
                    <td style={S.td}>{elcRoles.length ? <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>{elcRoles.map(r => <Badge key={r} type={roleBadgeType(r)}>{r}</Badge>)}</div> : <span style={{ color: C.b4 }}>—</span>}</td>
                    <td style={S.td}><Badge type={permCount(e) > 0 ? "inf" : "neutral"}>{permCount(e)}</Badge></td>
                    <td style={S.td}><Badge type={noteCount(e) > 0 ? "wn" : "neutral"}>{noteCount(e)}</Badge></td>
                    <td style={S.td}>{e.EmployeeActive === false ? <Badge type="neutral">Inactive</Badge> : <Badge type="ok">Active</Badge>}</td>
                    <td style={S.td} onClick={ev => ev.stopPropagation()}>{j ? <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => actions.openJourney(j.id)}><Badge type={j.JourneyType === "Onboarding" ? "inf" : "wn"}>{j.JourneyType}</Badge> · {fmtDate(journeyAnchorDate(j))}</button> : <span style={{ color: C.b4 }}>—</span>}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// REVIEWS TAB — quarterly review dashboard (overdue / due-soon / upcoming)
// ============================================================
function ReviewsTab() {
  const { state, actions, role, hasRole, currentEmail } = useData();
  const isManagerOrUp = hasRole("Admin", "HR", "Manager");
  // Scope: Admin/HR see everyone, Managers see their direct reports, others see themselves
  const visibleEmployees = useMemo(() => {
    if (hasRole("Admin", "HR")) return state.employees.filter(e => e.EmployeeActive !== false);
    if (hasRole("Manager")) return state.employees.filter(e => e.EmployeeActive !== false && (e.ManagerEmail || "").toLowerCase() === currentEmail);
    return state.employees.filter(e => (e.Email || "").toLowerCase() === currentEmail);
  }, [state.employees, role, currentEmail]);

  const rows = useMemo(() => visibleEmployees.map(e => {
    const status = reviewStatusFor(e, state.reviews, state.journeys);
    return { emp: e, ...status };
  }), [visibleEmployees, state.reviews, state.journeys]);

  const overdue    = rows.filter(r => r.state === "overdue");
  const dueSoon    = rows.filter(r => r.state === "due-soon");
  const okSoon     = rows.filter(r => r.state === "ok");
  const never      = rows.filter(r => r.state === "never");
  const notStarted = rows.filter(r => r.state === "not-started");
  const onboarding = rows.filter(r => r.state === "onboarding");
  const exempt     = rows.filter(r => r.state === "exempt");

  const Card = ({ r, badgeType, badge }) => {
    const tc = getTrainingComplianceFor(r.emp.Email, state.luAssignments, state.luCompletions, state.luCourses);
    const tcBadgeType = tc.overall === "overdue" ? "er" : tc.overall === "due-soon" ? "wn" : tc.overall === "ok" ? "ok" : "neutral";
    const tcText = tc.overall === "none" ? "No training assigned"
      : tc.overdue > 0 ? `${tc.complete}/${tc.total} · ${tc.overdue} overdue`
      : tc.dueSoon > 0 ? `${tc.complete}/${tc.total} · ${tc.dueSoon} due soon`
      : `${tc.complete}/${tc.total} complete`;
    return (
      <div onClick={() => actions.openEmployee(r.emp.Email)} style={{ cursor: "pointer", background: C.wh, border: `1px solid ${C.b1}`, borderRadius: 6, padding: 12, boxShadow: "0 1px 2px rgba(28,55,64,.04)" }}
           onMouseEnter={ev => { ev.currentTarget.style.borderColor = C.g5; ev.currentTarget.style.boxShadow = "0 2px 8px rgba(205,160,75,.18)"; }}
           onMouseLeave={ev => { ev.currentTarget.style.borderColor = C.b1; ev.currentTarget.style.boxShadow = "0 1px 2px rgba(28,55,64,.04)"; }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
          <Avatar name={r.emp.Title} size={36} />
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontWeight: 700, color: C.t7, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.emp.Title}</div>
            <div style={{ fontSize: 11, color: C.b4 }}>{r.emp.JobTitle || ""}</div>
          </div>
          <Badge type={badgeType}>{badge}</Badge>
        </div>
        <div style={{ fontSize: 11, color: C.b4 }}>
          {r.lastReview
            ? <>Last: <strong>{fmtDate(r.lastReview.ConductedDate)}</strong> · <Badge type={REVIEW_RATING_TYPE[r.lastReview.Rating] || "neutral"} dot={false}>{r.lastReview.Rating}</Badge></>
            : <>No reviews on file</>}
        </div>
        {r.nextDueDate && <div style={{ fontSize: 11, color: C.b4, marginTop: 4 }}>Next due: <strong>{fmtDate(r.nextDueDate)}</strong></div>}
        <div style={{ marginTop: 6, paddingTop: 6, borderTop: `1px dashed ${C.b1}`, fontSize: 11, color: C.b4, display: "flex", alignItems: "center", gap: 6 }}>
          <span>Training:</span>
          <Badge type={tcBadgeType} dot={false}>{tcText}</Badge>
        </div>
      </div>
    );
  };

  const Section = ({ title, items, type, badge }) => items.length === 0 ? null : (
    <div style={S.card}>
      <div style={S.cardT}><span>{title}</span><Badge type={type}>{items.length}</Badge></div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(290px, 1fr))", gap: 10 }}>
        {items.map(r => <Card key={r.emp.id} r={r} badgeType={type} badge={badge(r)} />)}
      </div>
    </div>
  );

  return (
    <div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
        <div style={S.kpi}><div style={S.kpiL}>Overdue</div><div style={{ ...S.kpiV, color: overdue.length > 0 ? C.er : C.t7 }}>{overdue.length}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Due ≤ 14 days</div><div style={{ ...S.kpiV, color: dueSoon.length > 0 ? C.wn : C.t7 }}>{dueSoon.length}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>On Track</div><div style={S.kpiV}>{okSoon.length}</div></div>
        {onboarding.length > 0 && <div style={S.kpi}><div style={S.kpiL}>In Onboarding</div><div style={S.kpiV}>{onboarding.length}</div></div>}
        {notStarted.length > 0 && <div style={S.kpi}><div style={S.kpiL}>Not Started Yet</div><div style={S.kpiV}>{notStarted.length}</div></div>}
        {never.length > 0 && <div style={S.kpi}><div style={S.kpiL}>Missing Start Date</div><div style={{ ...S.kpiV, color: C.wn }}>{never.length}</div></div>}
      </div>
      <div style={{ fontSize: 12, color: C.b4, marginBottom: 10 }}>
        {hasRole("Admin", "HR") ? "Showing all active employees." : hasRole("Manager") ? "Showing your direct reports." : "Showing your own review status."}
        {" "}First-cycle anchor: <strong>{fmtDate(CONFIG.firstReviewDueDate)}</strong>. After each employee has one review on file, the quarterly cadence (30/60/90 then every {REVIEW_QUARTER_DAYS} days) takes over.
      </div>
      <Section title="Overdue" items={overdue} type="er" badge={r => `${Math.abs(r.daysUntilDue)}d past due`} />
      <Section title="Due Soon" items={dueSoon} type="wn" badge={r => `due ${fmtDate(r.nextDueDate)}`} />
      <Section title="On Track" items={okSoon} type="ok" badge={r => `due ${fmtDate(r.nextDueDate)}`} />
      <Section title="In Onboarding (review handled by Day-90 task)" items={onboarding} type="inf" badge={() => "in onboarding"} />
      <Section title="Missing Start Date" items={never} type="wn" badge={() => "no start date"} />
      <Section title="Not Started Yet" items={notStarted} type="neutral" badge={r => `starts ${fmtDate(r.emp.StartDate)}`} />
      <Section title="Exempt — reviewed by ownership" items={exempt} type="neutral" badge={() => "exempt"} />
    </div>
  );
}

// ============================================================
// REPORTS TAB
// ============================================================
function ReportsTab() {
  const { state } = useData();
  const active = state.journeys.filter(j => j.Status !== "Complete" && j.Status !== "Cancelled");
  const ob = active.filter(j => j.JourneyType === "Onboarding");
  const of = active.filter(j => j.JourneyType === "Offboarding");
  const allTasks = state.journeyTasks.filter(t => active.some(j => String(j.id) === String(t.JourneyId)));
  const overdue = allTasks.filter(t => classifyDue(t.DueDate, t.Status) === "overdue");
  const byRole = {};
  overdue.forEach(t => { const r = t.AssigneeRole || "Unassigned"; byRole[r] = (byRole[r] || 0) + 1; });
  const last30 = state.journeys.filter(j => j.Status === "Complete" && j.Modified && (Date.now() - new Date(j.Modified).getTime()) / 86400000 <= 30);
  const activeEmps = state.employees.filter(e => e.EmployeeActive !== false);
  const overdueReviews = activeEmps.filter(e => reviewStatusFor(e, state.reviews, state.journeys).state === "overdue");
  const neverReviewed  = activeEmps.filter(e => reviewStatusFor(e, state.reviews, state.journeys).state === "never");
  const raisesLast90 = state.payChanges.filter(p => p.EffectiveDate && (Date.now() - new Date(p.EffectiveDate).getTime()) / 86400000 <= 90 && (p.NewPay > (p.PreviousPay || 0)));
  // Training compliance roll-up across active employees.
  const trainingRollup = activeEmps.reduce((acc, e) => {
    const tc = getTrainingComplianceFor(e.Email, state.luAssignments, state.luCompletions, state.luCourses);
    if (tc.overall === "overdue") acc.overdueEmps.push(e);
    acc.overdueCount += tc.overdue;
    acc.dueSoonCount += tc.dueSoon;
    return acc;
  }, { overdueEmps: [], overdueCount: 0, dueSoonCount: 0 });
  return (
    <div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
        <div style={S.kpi}><div style={S.kpiL}>Active Onboarding</div><div style={S.kpiV}>{ob.length}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Active Offboarding</div><div style={S.kpiV}>{of.length}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Overdue Tasks</div><div style={{ ...S.kpiV, color: overdue.length > 0 ? C.er : C.t7 }}>{overdue.length}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Overdue Reviews</div><div style={{ ...S.kpiV, color: overdueReviews.length > 0 ? C.er : C.t7 }}>{overdueReviews.length}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Never Reviewed</div><div style={{ ...S.kpiV, color: neverReviewed.length > 0 ? C.wn : C.t7 }}>{neverReviewed.length}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Raises (90d)</div><div style={S.kpiV}>{raisesLast90.length}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Overdue Training</div><div style={{ ...S.kpiV, color: trainingRollup.overdueCount > 0 ? C.er : C.t7 }}>{trainingRollup.overdueCount}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Training Due Soon</div><div style={{ ...S.kpiV, color: trainingRollup.dueSoonCount > 0 ? C.wn : C.t7 }}>{trainingRollup.dueSoonCount}</div></div>
        <div style={S.kpi}><div style={S.kpiL}>Completed (30d)</div><div style={S.kpiV}>{last30.length}</div></div>
      </div>
      <div style={S.card}>
        <div style={S.cardT}>Overdue Tasks by Role</div>
        {Object.keys(byRole).length === 0 ? <Empty title="No overdue tasks" sub="Everyone is on track." /> : (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th style={S.th}>Role</th><th style={S.th}>Overdue</th></tr></thead>
            <tbody>{Object.entries(byRole).sort((a, b) => b[1] - a[1]).map(([r, n]) => <tr key={r}><td style={S.td}>{r}</td><td style={S.td}><Badge type="er">{n}</Badge></td></tr>)}</tbody>
          </table>
        )}
      </div>
      <div style={S.card}>
        <div style={S.cardT}>Recent Activity</div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr><th style={S.th}>When</th><th style={S.th}>Journey</th><th style={S.th}>Employee</th><th style={S.th}>Status</th></tr></thead>
          <tbody>
            {state.journeys.slice().sort((a, b) => (b.Modified || "").localeCompare(a.Modified || "")).slice(0, 12).map(j => (
              <tr key={j.id}>
                <td style={S.td}>{fmtDate(j.Modified)}</td>
                <td style={S.td}><Badge type={j.JourneyType === "Onboarding" ? "inf" : "wn"}>{j.JourneyType}</Badge></td>
                <td style={S.td}>{j.EmployeeName || j.EmployeeEmail}</td>
                <td style={S.td}><Badge type={j.Status === "Complete" ? "ok" : j.Status === "Cancelled" ? "neutral" : "inf"}>{j.Status}</Badge></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================
// SETUP / DIAGNOSTICS — first-run helper
// ============================================================
function AccessOverviewCard() {
  const { state, actions, hasRole } = useData();
  const [saving, setSaving] = useState({});
  const canEdit = hasRole("Admin");
  const ELC_ROLES = ["Admin", "HR", "IT", "Manager", "Employee"];
  const roleBadgeType = r => r === "Admin" ? "er" : r === "HR" ? "pu" : r === "IT" ? "inf" : r === "Manager" ? "wn" : "neutral";
  const active = state.employees.filter(e => e.EmployeeActive !== false).sort((a, b) => (a.Title || "").localeCompare(b.Title || ""));
  // A person now shows under every role they hold (explicit list, or the legacy fallback).
  const byRole = ELC_ROLES.reduce((m, r) => { m[r] = active.filter(e => detectRoles(e, e.Email).includes(r)); return m; }, {});

  // Write the selected role set as a delimited list (e.g. "Admin; HR"). Empty → "None" (fallback).
  const setRoles = async (email, newRoles) => {
    setSaving(s => ({ ...s, [email]: true }));
    try { await actions.setEmployeePermissions(email, { ELCRole: newRoles.join("; ") }, "Set via Setup → Access overview", null); }
    catch (e) { alert("Save failed: " + e.message); }
    finally { setSaving(s => ({ ...s, [email]: false })); }
  };

  return (
    <div style={S.card}>
      <div style={S.cardT}>Access to this app</div>
      <div style={{ fontSize: 12, color: C.b4, marginBottom: 10 }}>
        Controls what each employee can do <em>inside</em> the Employee Lifecycle app. One person can hold several roles — check all that apply. <strong>Admin</strong> can do anything. <strong>HR</strong> can run journeys, set permissions, and write notes. <strong>IT</strong> can complete IT-assigned tasks. <strong>Manager</strong> can see their reports' onboarding/offboarding. With no roles checked, the person only sees their own tasks (Employee). <strong>VA</strong> is tracking-only and gates the person out of the app. The super-admin <code style={{ fontFamily: mono }}>{CONFIG.adminEmails.join(", ")}</code> is always Admin regardless of this column.
      </div>
      <div style={{ display: "grid", gap: 6, marginBottom: 12 }}>
        {ELC_ROLES.map(r => (
          <div key={r} style={{ display: "grid", gridTemplateColumns: "100px 1fr", gap: 8, alignItems: "center", padding: "6px 0" }}>
            <Badge type={roleBadgeType(r)}>{r}</Badge>
            <div style={{ fontSize: 12, color: C.b6 }}>{byRole[r].length === 0 ? <em style={{ color: C.b4 }}>nobody</em> : byRole[r].map(e => e.Title).join(", ")}</div>
          </div>
        ))}
      </div>
      {canEdit && (
        <div style={{ borderTop: `1px solid ${C.b1}`, paddingTop: 12 }}>
          <div style={S.sec}>Set roles</div>
          <div style={{ maxHeight: 360, overflowY: "auto", border: `1px solid ${C.b1}`, borderRadius: 4 }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th style={S.th}>Employee</th><th style={S.th}>Detected (fallback)</th><th style={S.th}>Roles (check all that apply)</th></tr></thead>
              <tbody>
                {active.map(e => {
                  const cur = parseRoles(e.ELCRole);
                  const detected = detectRole({ ...e, ELCRole: null }, e.Email);
                  const busy = !!saving[e.Email];
                  const toggle = r => {
                    setRoles(e.Email, cur.includes(r) ? cur.filter(x => x !== r) : [...cur, r]);
                  };
                  return (
                    <tr key={e.id}>
                      <td style={S.td}><div style={{ fontWeight: 600 }}>{e.Title}</div><div style={{ fontSize: 11, color: C.b4 }}>{e.Email}</div></td>
                      <td style={S.td}><Badge type="neutral">{detected}</Badge></td>
                      <td style={S.td}>
                        <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center" }}>
                          {COMBINABLE_ROLES.map(r => (
                            <label key={r} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, color: C.b6, cursor: busy ? "default" : "pointer" }}>
                              <input type="checkbox" checked={cur.includes(r)} disabled={busy} onChange={() => toggle(r)} />{r}
                            </label>
                          ))}
                          {busy && <span style={{ fontSize: 11, color: C.b4 }}>saving…</span>}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function EmailTemplatesCard() {
  const { state, actions, hasRole } = useData();
  const canEdit = hasRole("Admin", "HR");
  const [seeding, setSeeding] = useState(false);
  const seedDefaults = async () => {
    if (!confirm(`Seed ${DEFAULT_EMAIL_TEMPLATES.length} default email templates? Adds to existing — does not replace.`)) return;
    setSeeding(true);
    try { for (const t of DEFAULT_EMAIL_TEMPLATES) await actions.createEmailTemplate({ ...t, Active: true }); }
    catch (e) { alert("Seed failed: " + e.message); } finally { setSeeding(false); }
  };
  const grouped = useMemo(() => {
    const m = {};
    for (const t of state.emailTemplates) { const c = t.Category || "Other"; (m[c] = m[c] || []).push(t); }
    return m;
  }, [state.emailTemplates]);

  return (
    <div style={S.card}>
      <div style={S.cardT}>
        <span>Email Templates</span>
        <div style={{ display: "flex", gap: 8 }}>
          {canEdit && state.emailTemplates.length === 0 && <button style={S.btnO(C.t5)} disabled={seeding} onClick={seedDefaults}>{seeding ? "Seeding…" : "Seed defaults"}</button>}
          {canEdit && <button style={S.btn(C.hdr)} onClick={() => actions.openEmailTplNew()}>+ New Template</button>}
        </div>
      </div>
      <div style={{ fontSize: 12, color: C.b4, marginBottom: 10 }}>
        Reusable email bodies for the Send Email button. Body and subject support {`{{Variables}}`} (FirstName, ManagerName, StartDate, etc.) — substituted at send time. Default recipients (Work / Personal / Manager) and CC pre-fill the modal.
      </div>
      {state.emailTemplates.length === 0 ? <Empty title="No email templates yet" sub={canEdit ? "Click 'Seed defaults' to load the starter set, or '+ New Template' to write your own." : "Ask HR/Admin to set this up."} /> : (
        <div style={{ display: "grid", gap: 8 }}>
          {EMAIL_TEMPLATE_CATEGORIES.map(cat => {
            const items = grouped[cat] || [];
            if (items.length === 0) return null;
            return (
              <div key={cat}>
                <div style={{ ...S.sec, marginBottom: 4 }}>{cat}</div>
                <div style={{ border: `1px solid ${C.b1}`, borderRadius: 4, overflow: "hidden" }}>
                  {items.map((t, i) => (
                    <div key={t.id} style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 10, alignItems: "center", padding: "8px 10px", borderBottom: i < items.length - 1 ? `1px solid ${C.b1}` : "none", background: C.wh }}>
                      <div style={{ minWidth: 0 }}>
                        <div style={{ fontWeight: 600, color: C.t7 }}>{t.Title}</div>
                        <div style={{ fontSize: 11, color: C.b4, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.Subject}</div>
                        <div style={{ fontSize: 10, color: C.b4, marginTop: 2 }}>To: {t.DefaultTo || "(none)"} {t.DefaultCc ? `· CC: ${t.DefaultCc}` : ""}</div>
                      </div>
                      {canEdit && <button style={{ ...S.btnO(C.t5), ...S.xs }} onClick={() => actions.openEmailTplEdit(t.id)}>Edit</button>}
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function SetupTab() {
  const { state, actions } = useData();
  const [seedingApps, setSeedingApps] = useState(false);
  const realApps = state.apps.filter(a => !String(a.id).startsWith("default-"));
  const checks = [
    { label: `${CONFIG.lists.employees} list`,      ok: state.employees.length > 0, hint: "Shared list — should already exist." },
    { label: `${CONFIG.lists.journeys} list`,       ok: true, hint: "Provisioned by scripts/provision-lists.ps1." },
    { label: `${CONFIG.lists.templateTasks} list (has templates)`, ok: state.templates.length > 0, hint: "Templates tab → 'Seed defaults' to load NewShire's starter task set." },
    { label: `${CONFIG.lists.journeyTasks} list`,   ok: true, hint: "Provisioned." },
    { label: `${CONFIG.lists.apps} registry (has app rows)`, ok: realApps.length > 0, hint: "Defines which NewShire apps appear in the Permissions matrix." },
    { label: `${CONFIG.lists.notes} list`,          ok: true, hint: "Stores coaching, discipline, PIP, 1:1, praise notes." },
    { label: `${CONFIG.lists.audit} list`,          ok: true, hint: "Audit trail of permission changes." },
    { label: `${CONFIG.lists.files} document library`, ok: true, hint: "Per-employee HR documents." },
    { label: `${CONFIG.lists.reviews} list`,        ok: true, hint: "Quarterly performance reviews. Re-run provision script if missing." },
    { label: `${CONFIG.lists.payChanges} list`,     ok: true, hint: "Compensation history per employee. Re-run provision script if missing." },
    { label: `${CONFIG.lists.emailTemplates} list`, ok: true, hint: "Reusable email templates. Re-run provision script if missing." },
    { label: "Employees.PersonalEmail column",      ok: state.employees.some(e => e.PersonalEmail), hint: "Add personal email addresses on the Profile sub-tab. Falls back to Work email if missing." },
  ];

  const seedApps = async () => {
    if (!confirm(`Seed ${DEFAULT_APPS.length} default apps into the registry?`)) return;
    setSeedingApps(true);
    try {
      for (const a of DEFAULT_APPS) await actions.createApp(a);
      await actions.reload();
    } catch (e) { alert("Seed failed: " + e.message); } finally { setSeedingApps(false); }
  };

  const Row = ({ label, ok, hint }) => (
    <div style={{ display: "grid", gridTemplateColumns: "auto 1fr auto", gap: 10, alignItems: "center", padding: "10px 0", borderBottom: `1px solid ${C.b1}` }}>
      <div style={{ fontSize: 18 }}>{ok ? "✅" : "⚠️"}</div>
      <div><div style={{ fontWeight: 600 }}>{label}</div><div style={{ fontSize: 11, color: C.b4 }}>{hint}</div></div>
      <Badge type={ok ? "ok" : "wn"}>{ok ? "ok" : "set up"}</Badge>
    </div>
  );

  return (
    <div>
      <div style={S.card}>
        <div style={S.cardT}>Setup checklist</div>
        {checks.map(c => <Row key={c.label} {...c} />)}
      </div>

      <AccessOverviewCard />

      <EmailTemplatesCard />

      <div style={S.card}>
        <div style={S.cardT}><span>Apps Registry</span>
          {realApps.length === 0 && <button style={S.btnO(C.t5)} disabled={seedingApps} onClick={seedApps}>{seedingApps ? "Seeding…" : "Seed defaults"}</button>}
        </div>
        <div style={{ fontSize: 12, color: C.b4, marginBottom: 8 }}>
          Each row maps a NewShire app to the column on the <strong>Employees</strong> list that stores per-app role. The provisioning script also creates that column (Choice with the listed roles + 'None').
        </div>
        {state.apps.length === 0 ? <Empty title="No apps registered" sub="Click 'Seed defaults' or add rows to ELC_Apps." /> : (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th style={S.th}></th><th style={S.th}>App</th><th style={S.th}>Employees Column</th><th style={S.th}>Roles</th><th style={S.th}>Default on Hire</th></tr></thead>
            <tbody>
              {state.apps.map(a => (
                <tr key={a.id}>
                  <td style={S.td}><div style={{ width: 26, height: 26, background: a.Color || C.t5, color: "#fff", borderRadius: 4, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontFamily: "Georgia,serif" }}>{a.IconLetter || a.Title?.[0]}</div></td>
                  <td style={S.td}><div style={{ fontWeight: 600 }}>{a.Title}</div><div style={{ fontSize: 11, color: C.b4 }}>{a.AppKey}</div></td>
                  <td style={S.td}><code style={{ fontFamily: mono, fontSize: 12 }}>{a.ColumnName}</code></td>
                  <td style={S.td}><div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>{(a.Roles || []).map(r => <Badge key={r} type="neutral">{r}</Badge>)}</div></td>
                  <td style={S.td}>{a.OnboardingDefault || "None"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      <div style={S.card}>
        <div style={S.cardT}>Required SharePoint Lists</div>
        <div style={{ fontSize: 12, color: C.b4, lineHeight: 1.6 }}>
          Run <code style={{ fontFamily: mono }}>scripts/provision-lists.ps1</code> once from the repo root to create everything:
          <ul style={{ marginLeft: 18, marginTop: 6, lineHeight: 1.8 }}>
            <li><strong>ELC_Journeys</strong> — one row per onboarding/offboarding instance</li>
            <li><strong>ELC_TemplateTasks</strong> — reusable task templates by phase/role</li>
            <li><strong>ELC_JourneyTasks</strong> — materialised tasks per journey</li>
            <li><strong>ELC_Config</strong> — single-row JSON settings</li>
            <li><strong>ELC_Apps</strong> — NewShire app registry (this controls the Permissions matrix)</li>
            <li><strong>ELC_EmployeeNotes</strong> — coaching, discipline, PIP, 1:1, praise</li>
            <li><strong>ELC_PermissionAudit</strong> — audit log of every permission change</li>
            <li><strong>ELC_EmployeeFiles</strong> — SharePoint document library (folder per employee)</li>
          </ul>
          The script also adds a Choice column on the <strong>Employees</strong> list for each registered app (e.g. <code style={{ fontFamily: mono }}>VATrackerRole</code>, <code style={{ fontFamily: mono }}>PMHubRole</code>).
        </div>
      </div>
    </div>
  );
}

// ============================================================
// MAIN APP
// ============================================================
function App() {
  const { acct, token, login, logout, refresh, err: authErr, ready } = useMsal();
  const [state, setState] = useState({ employees: [], journeys: [], templates: [], journeyTasks: [], config: {}, apps: [], notes: [], audit: [], reviews: [], payChanges: [], emailTemplates: [], luCourses: [], luAssignments: [], luCompletions: [] });
  const [loading, setLoading] = useState(false);
  const [hasLoadedOnce, setHasLoadedOnce] = useState(false);
  const [loadErr, setLoadErr] = useState(null);
  const [tab, setTab] = useState("onboarding");
  const [editTaskId, setEditTaskId] = useState(null);
  const [openJourneyId, setOpenJourneyId] = useState(null);
  const [openEmployeeEmail, setOpenEmployeeEmail] = useState(null);
  const [editNoteId, setEditNoteId] = useState(null);
  const [newNoteFor, setNewNoteFor] = useState(null);
  const [editReviewId, setEditReviewId] = useState(null);
  const [newReviewFor, setNewReviewFor] = useState(null);
  const [editPayId, setEditPayId] = useState(null);
  const [newPayFor, setNewPayFor] = useState(null);
  const [adHocTaskForJourneyId, setAdHocTaskForJourneyId] = useState(null);
  const [sendEmailFor, setSendEmailFor] = useState(null);
  const [editEmailTplId, setEditEmailTplId] = useState(null);

  const currentEmail = (acct?.username || "").toLowerCase();
  const me = state.employees.find(e => (e.Email || "").toLowerCase() === currentEmail);
  const roles = detectRoles(me, currentEmail);
  const role = primaryRole(roles);                       // highest role (badges, VA gate)
  const hasRole = (...names) => names.some(n => roles.includes(n)); // any-of capability check

  const reload = useCallback(async () => {
    if (!token) return;
    setLoading(true); setLoadErr(null);
    try { const d = await loadAll(token); setState(d); setHasLoadedOnce(true); }
    catch (e) { setLoadErr(e.message); }
    finally { setLoading(false); }
  }, [token]);

  useEffect(() => { if (token) reload(); }, [token, reload]);

  const withToken = useCallback(async (fn) => {
    let tk = token; try { return await fn(tk); }
    catch (e) {
      if ((e.message || "").includes("401") || (e.message || "").includes("403")) { tk = await refresh(); if (tk) return await fn(tk); }
      throw e;
    }
  }, [token, refresh]);

  const actions = useMemo(() => ({
    reload,
    openTaskEdit: id => setEditTaskId(id),
    openAddTask: journeyId => setAdHocTaskForJourneyId(journeyId),
    openJourney: id => setOpenJourneyId(id),
    openEmployee: email => setOpenEmployeeEmail((email || "").toLowerCase()),
    openNoteEdit: id => setEditNoteId(id),
    openNoteNew: email => setNewNoteFor((email || "").toLowerCase()),
    createJourney: async (fields) => withToken(async tk => {
      const r = await gPost(tk, lUrl(CONFIG.lists.journeys), { Title: `${fields.JourneyType} — ${fields.EmployeeName || fields.EmployeeEmail}`, ...fields });
      return { id: r.id, ...r.fields };
    }),
    updateJourney: async (id, patch) => withToken(async tk => { await gPatch(tk, iUrl(CONFIG.lists.journeys, id), patch); await reload(); }),
    createTask: async (fields) => withToken(async tk => {
      const r = await gPost(tk, lUrl(CONFIG.lists.journeyTasks), { Title: fields.Title, ...fields });
      return { id: r.id, ...r.fields };
    }),
    updateTask: async (id, patch) => withToken(async tk => {
      await gPatch(tk, iUrl(CONFIG.lists.journeyTasks, id), patch);
      setState(s => ({ ...s, journeyTasks: s.journeyTasks.map(t => String(t.id) === String(id) ? { ...t, ...patch } : t) }));
    }),
    deleteTask: async (id) => withToken(async tk => {
      await gDelete(tk, `${SITE}/lists/${CONFIG.lists.journeyTasks}/items/${id}`);
      setState(s => ({ ...s, journeyTasks: s.journeyTasks.filter(t => String(t.id) !== String(id)) }));
    }),
    createTemplate: async (fields) => withToken(async tk => { await gPost(tk, lUrl(CONFIG.lists.templateTasks), { Title: fields.Title, ...fields }); }),
    updateTemplate: async (id, patch) => withToken(async tk => { await gPatch(tk, iUrl(CONFIG.lists.templateTasks, id), patch); }),
    deleteTemplate: async (id) => withToken(async tk => { await gDelete(tk, `${SITE}/lists/${CONFIG.lists.templateTasks}/items/${id}`); }),

    // ── Employees (upsert) ──
    upsertEmployee: async (fields) => withToken(async tk => {
      const email = (fields.Email || "").toLowerCase();
      const existing = state.employees.find(e => (e.Email || "").toLowerCase() === email);
      if (existing) {
        const patch = { ...fields };
        await gPatch(tk, iUrl(CONFIG.lists.employees, existing.id), patch);
        const updated = { ...existing, ...patch };
        setState(s => ({ ...s, employees: s.employees.map(x => String(x.id) === String(existing.id) ? updated : x) }));
        return updated;
      }
      const newFields = { Title: fields.Title || fields.Name || email, ...fields };
      if (newFields.EmployeeActive === undefined) newFields.EmployeeActive = true;
      const r = await gPost(tk, lUrl(CONFIG.lists.employees), newFields);
      const created = { id: r.id, ...r.fields };
      setState(s => ({ ...s, employees: [...s.employees, created] }));
      return created;
    }),
    setEmployeePermissions: async (employeeEmail, perms, reason, journeyId, empRecord) => withToken(async tk => {
      const email = (employeeEmail || "").toLowerCase();
      // Prefer state, but fall back to a record passed in by the caller (e.g. a hire
      // just created via upsertEmployee that isn't in state.employees yet this render).
      const existing = state.employees.find(e => (e.Email || "").toLowerCase() === email)
        || (empRecord && (empRecord.Email || "").toLowerCase() === email ? empRecord : null);
      if (!existing) throw new Error(`Employee not found: ${email}`);
      const patch = {}; const audits = [];
      for (const [col, newRole] of Object.entries(perms)) {
        const oldRole = existing[col] || "None";
        if ((newRole || "None") === oldRole) continue;
        patch[col] = newRole || "None";
        const app = state.apps.find(a => a.ColumnName === col);
        audits.push({ EmployeeEmail: email, AppKey: app?.AppKey || col, OldRole: oldRole, NewRole: newRole || "None", ChangedBy: currentEmail, ChangedAt: new Date().toISOString(), Reason: reason || "", JourneyId: journeyId || "" });
      }
      if (Object.keys(patch).length > 0) {
        await gPatch(tk, iUrl(CONFIG.lists.employees, existing.id), patch);
      }
      for (const a of audits) {
        await gPost(tk, lUrl(CONFIG.lists.audit), { Title: `${a.AppKey}: ${a.OldRole} → ${a.NewRole}`, ...a }).catch(e => console.warn("audit log failed:", e.message));
      }
      setState(s => ({ ...s, employees: s.employees.map(e => String(e.id) === String(existing.id) ? { ...e, ...patch } : e), audit: [...s.audit, ...audits.map((a, i) => ({ id: `local-${Date.now()}-${i}`, ...a }))] }));
    }),

    // ── Notes / Discipline ──
    createNote: async (fields) => withToken(async tk => {
      const r = await gPost(tk, lUrl(CONFIG.lists.notes), { Title: fields.Title || `${fields.NoteType} — ${fields.EmployeeEmail}`, AuthorEmail: currentEmail, NoteDate: fields.NoteDate || todayIso(), ...fields });
      setState(s => ({ ...s, notes: [...s.notes, { id: r.id, ...r.fields }] }));
      return r;
    }),
    updateNote: async (id, patch) => withToken(async tk => {
      await gPatch(tk, iUrl(CONFIG.lists.notes, id), patch);
      setState(s => ({ ...s, notes: s.notes.map(n => String(n.id) === String(id) ? { ...n, ...patch } : n) }));
    }),
    deleteNote: async (id) => withToken(async tk => {
      await gDelete(tk, `${SITE}/lists/${CONFIG.lists.notes}/items/${id}`);
      setState(s => ({ ...s, notes: s.notes.filter(n => String(n.id) !== String(id)) }));
    }),

    // ── Apps registry ──
    updateApp: async (id, patch) => withToken(async tk => { await gPatch(tk, iUrl(CONFIG.lists.apps, id), patch); }),
    createApp:  async (fields) => withToken(async tk => { await gPost(tk, lUrl(CONFIG.lists.apps), { Title: fields.Title, ...fields, Roles: typeof fields.Roles === "string" ? fields.Roles : JSON.stringify(fields.Roles || []) }); }),

    // ── Quarterly Reviews ──
    openReviewEdit: id => setEditReviewId(id),
    openReviewNew:  email => setNewReviewFor((email || "").toLowerCase()),
    createReview: async (fields) => withToken(async tk => {
      const r = await gPost(tk, lUrl(CONFIG.lists.reviews), { Title: fields.Title || `${fields.ReviewPeriod} — ${fields.EmployeeEmail}`, ReviewerEmail: fields.ReviewerEmail || currentEmail, ...fields });
      const created = { id: r.id, ...r.fields };
      setState(s => ({ ...s, reviews: [...s.reviews, created] }));
      return created;
    }),
    updateReview: async (id, patch) => withToken(async tk => {
      await gPatch(tk, iUrl(CONFIG.lists.reviews, id), patch);
      setState(s => ({ ...s, reviews: s.reviews.map(r => String(r.id) === String(id) ? { ...r, ...patch } : r) }));
    }),
    deleteReview: async (id) => withToken(async tk => {
      await gDelete(tk, `${SITE}/lists/${CONFIG.lists.reviews}/items/${id}`);
      setState(s => ({ ...s, reviews: s.reviews.filter(r => String(r.id) !== String(id)) }));
    }),

    // ── Pay Changes ──
    openPayEdit: id => setEditPayId(id),
    openPayNew:  email => setNewPayFor((email || "").toLowerCase()),
    createPayChange: async (fields) => withToken(async tk => {
      const r = await gPost(tk, lUrl(CONFIG.lists.payChanges), { Title: fields.Title || `${fields.ChangeType} — ${fields.EmployeeEmail}`, ApprovedBy: fields.ApprovedBy || currentEmail, ...fields });
      const created = { id: r.id, ...r.fields };
      setState(s => ({ ...s, payChanges: [...s.payChanges, created] }));
      return created;
    }),
    updatePayChange: async (id, patch) => withToken(async tk => {
      await gPatch(tk, iUrl(CONFIG.lists.payChanges, id), patch);
      setState(s => ({ ...s, payChanges: s.payChanges.map(p => String(p.id) === String(id) ? { ...p, ...patch } : p) }));
    }),
    deletePayChange: async (id) => withToken(async tk => {
      await gDelete(tk, `${SITE}/lists/${CONFIG.lists.payChanges}/items/${id}`);
      setState(s => ({ ...s, payChanges: s.payChanges.filter(p => String(p.id) !== String(id)) }));
    }),

    // ── Email Templates (admin-edited library of reusable templates) ──
    openEmailTplEdit: id => setEditEmailTplId(id),
    openEmailTplNew: () => setEditEmailTplId("__new__"),
    createEmailTemplate: async (fields) => withToken(async tk => {
      const r = await gPost(tk, lUrl(CONFIG.lists.emailTemplates), { Title: fields.Title, Active: true, ...fields });
      const created = { id: r.id, ...r.fields };
      setState(s => ({ ...s, emailTemplates: [...s.emailTemplates, created] }));
      return created;
    }),
    updateEmailTemplate: async (id, patch) => withToken(async tk => {
      await gPatch(tk, iUrl(CONFIG.lists.emailTemplates, id), patch);
      setState(s => ({ ...s, emailTemplates: s.emailTemplates.map(t => String(t.id) === String(id) ? { ...t, ...patch } : t) }));
    }),
    deleteEmailTemplate: async (id) => withToken(async tk => {
      await gDelete(tk, `${SITE}/lists/${CONFIG.lists.emailTemplates}/items/${id}`);
      setState(s => ({ ...s, emailTemplates: s.emailTemplates.filter(t => String(t.id) !== String(id)) }));
    }),
    // Compose / send. Returns sent payload for the caller to record if desired.
    openSendEmail: forEmail => setSendEmailFor((forEmail || "").toLowerCase()),
    sendEmail: async ({ to, cc, subject, bodyHtml }) => withToken(async tk => {
      const toRecipients = (to || []).filter(Boolean).map(a => ({ emailAddress: { address: a } }));
      const ccRecipients = (cc || []).filter(Boolean).map(a => ({ emailAddress: { address: a } }));
      const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
        method: "POST",
        headers: { Authorization: `Bearer ${tk}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          message: {
            subject: subject || "",
            body: { contentType: "HTML", content: bodyHtml || "" },
            toRecipients, ccRecipients,
          },
          saveToSentItems: true,
        }),
      });
      if (!res.ok) throw new Error(`sendMail ${res.status} ${await res.text().catch(() => '')}`);
    }),
  }), [withToken, reload, state.employees, state.apps, currentEmail]);

  const ctxValue = { state, actions, role, roles, hasRole, currentEmail, me };

  // Reviews tab count — must be called every render to satisfy rules-of-hooks,
  // so it lives ABOVE the early returns below.
  const reviewCount = useMemo(() => {
    const scope = hasRole("Admin", "HR") ? state.employees.filter(e => e.EmployeeActive !== false)
      : hasRole("Manager") ? state.employees.filter(e => e.EmployeeActive !== false && (e.ManagerEmail || "").toLowerCase() === currentEmail)
      : state.employees.filter(e => (e.Email || "").toLowerCase() === currentEmail);
    return scope.reduce((n, e) => {
      const s = reviewStatusFor(e, state.reviews, state.journeys);
      // Badge counts the action-needed buckets (overdue / due-soon) plus the data-cleanup "never"
      return n + ((s.state === "overdue" || s.state === "due-soon" || s.state === "never") ? 1 : 0);
    }, 0);
  }, [state.employees, state.reviews, state.journeys, role, currentEmail]);

  if (!ready) return <div style={{ ...S.page, display: "flex", alignItems: "center", justifyContent: "center" }}><div style={{ animation: "pulse 1.4s ease-in-out infinite", color: C.b4 }}>Loading…</div></div>;
  if (!acct) return <LoginScreen onLogin={login} err={authErr} />;

  // Header counts respect the same visibility rule as the tabs themselves —
  // a regular Employee shouldn't see "3 offboardings" if they can't see any.
  const obCount = state.journeys.filter(j => j.JourneyType === "Onboarding" && j.Status !== "Complete" && j.Status !== "Cancelled" && canSeeJourney(j, hasRole, currentEmail, state.employees)).length;
  const offCount = state.journeys.filter(j => j.JourneyType === "Offboarding" && j.Status !== "Complete" && j.Status !== "Cancelled" && canSeeJourney(j, hasRole, currentEmail, state.employees)).length;
  const myCount = state.journeyTasks.filter(t => {
    if (t.Status === "Done") return false;
    if (!taskIsMine(t, currentEmail, hasRole)) return false;
    const j = state.journeys.find(jj => String(jj.id) === String(t.JourneyId));
    return j && j.Status !== "Cancelled" && j.Status !== "Complete";
  }).length;

  const isAdmin = hasRole("Admin", "HR");
  const isManagerOrUp = isAdmin || hasRole("Manager");
  const tabs = [
    { key: "onboarding", label: "Onboarding", count: obCount },
    { key: "offboarding", label: "Offboarding", count: offCount },
    { key: "mytasks", label: "My Tasks", count: myCount },
    isManagerOrUp && { key: "reviews", label: "Reviews", count: reviewCount },
    isAdmin && { key: "employees", label: "Employees" },
    isAdmin && { key: "templates", label: "Templates" },
    isAdmin && { key: "reports", label: "Reports" },
    isAdmin && { key: "setup", label: "Setup" },
  ].filter(Boolean);

  return (
    <DataCtx.Provider value={ctxValue}>
      <div style={S.page}>
        <Header user={{ name: me?.Title, email: currentEmail }} role={role} roles={roles} onLogout={logout} />
        <TabBar tabs={tabs} active={tab} onChange={setTab} />
        <div style={S.content}>
          {loadErr && <div style={{ background: C.erb, color: C.er, padding: 10, borderRadius: 4, fontSize: 12, marginBottom: 14 }}>{loadErr} <button onClick={reload} style={{ ...S.btnO(C.er, C.er), ...S.xs, marginLeft: 8 }}>Retry</button></div>}
          {loading && hasLoadedOnce && <div style={{ position: "fixed", top: 60, right: 18, background: C.t0, color: C.t7, fontSize: 11, fontWeight: 600, padding: "5px 10px", borderRadius: 99, border: `1px solid ${C.t1}`, zIndex: 50 }}>Reloading…</div>}
          {!hasLoadedOnce && loading ? <div style={{ ...S.card, textAlign: "center", color: C.b4, padding: 40 }}>Loading data from SharePoint…</div> :
            tab === "onboarding" ? <JourneysTab type="Onboarding" /> :
            tab === "offboarding" ? <JourneysTab type="Offboarding" /> :
            tab === "mytasks" ? <MyTasksTab /> :
            tab === "reviews" ? <ReviewsTab /> :
            tab === "employees" ? <EmployeesTab /> :
            tab === "templates" ? <TemplatesTab /> :
            tab === "reports" ? <ReportsTab /> :
            tab === "setup" ? <SetupTab /> : null}
        </div>
        {editTaskId && <TaskEditModal taskId={editTaskId} onClose={() => setEditTaskId(null)} />}
        {openJourneyId && <JourneyDetailModal journeyId={openJourneyId} onClose={() => setOpenJourneyId(null)} />}
        {openEmployeeEmail && <EmployeeDetailModal email={openEmployeeEmail} onClose={() => setOpenEmployeeEmail(null)} />}
        {(editNoteId || newNoteFor) && <NoteEditModal noteId={editNoteId} forEmail={newNoteFor} onClose={() => { setEditNoteId(null); setNewNoteFor(null); }} />}
        {(editReviewId || newReviewFor) && <ReviewEditModal reviewId={editReviewId} forEmail={newReviewFor} onClose={() => { setEditReviewId(null); setNewReviewFor(null); }} />}
        {(editPayId || newPayFor) && <PayChangeEditModal payId={editPayId} forEmail={newPayFor} onClose={() => { setEditPayId(null); setNewPayFor(null); }} />}
        {adHocTaskForJourneyId && <AdHocTaskModal journeyId={adHocTaskForJourneyId} onClose={() => setAdHocTaskForJourneyId(null)} />}
        {sendEmailFor && <SendEmailModal forEmail={sendEmailFor} onClose={() => setSendEmailFor(null)} />}
        {editEmailTplId && <EmailTemplateEditModal tplId={editEmailTplId === "__new__" ? null : editEmailTplId} onClose={() => setEditEmailTplId(null)} />}
      </div>
    </DataCtx.Provider>
  );
}
