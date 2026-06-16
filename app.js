// Trip Timeline Builder — vanilla JS, persisted to localStorage.
//
// Storage layout (multi-trip):
//   trip-builder-trips           : JSON array of {id, name, start, end} (registry)
//   trip-builder-trip-<id>       : JSON of full trip state for that id
//   trip-builder-v7              : legacy single-trip blob (auto-migrated on first run)

const LEGACY_KEY = "trip-builder-v7";
const REGISTRY_KEY = "trip-builder-trips";
const TRIP_KEY_PREFIX = "trip-builder-trip-";

let CURRENT_TRIP_ID = null;
let IS_OWNER = false;       // signed-in user owns this trip (only they can re-share)
let CAN_EDIT = false;       // owner OR editor-level access
let CAN_SEE_PRICING = false; // owner, editor, viewer-pricing, or token-bearing anon
function STORAGE_KEY_FOR(id) { return TRIP_KEY_PREFIX + id; }
function STORAGE_KEY() { return STORAGE_KEY_FOR(CURRENT_TRIP_ID); }

function safeParse(s) { try { return JSON.parse(s); } catch (e) { return null; } }

// Wait until window.fb (firebase-client.js) has finished loading.
function whenFb() {
  if (window.fb) return Promise.resolve();
  return new Promise((resolve) => {
    const iv = setInterval(() => { if (window.fb) { clearInterval(iv); resolve(); } }, 30);
  });
}

// Fetch a trip from Firestore. Anonymous public viewers use the dedicated
// publicTrips path via window.fb.loadPublicTrip instead.
async function fetchTrip(slug) {
  await whenFb();
  if (!window.fb.user) return { ok: false, status: 401 };
  return window.fb.loadTrip(slug);
}

async function putTrip(slug, blob) {
  await whenFb();
  if (!window.fb.user) return { ok: false, status: 401 };
  return window.fb.saveTrip(slug, blob);
}


function readRegistry() {
  try {
    const raw = localStorage.getItem(REGISTRY_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}
function writeRegistry(list) {
  localStorage.setItem(REGISTRY_KEY, JSON.stringify(list));
}
function upsertRegistry(entry) {
  const list = readRegistry();
  const i = list.findIndex(t => t.id === entry.id);
  if (i >= 0) list[i] = { ...list[i], ...entry };
  else list.push(entry);
  writeRegistry(list);
}
function newTripId() {
  return "t" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}
function migrateLegacy() {
  const legacy = localStorage.getItem(LEGACY_KEY);
  if (!legacy) return null;
  try {
    const parsed = JSON.parse(legacy);
    const id = newTripId();
    localStorage.setItem(STORAGE_KEY_FOR(id), legacy);
    upsertRegistry({
      id,
      name: parsed.name || "Untitled trip",
      start: parsed.start || null,
      end: parsed.end || null,
    });
    localStorage.removeItem(LEGACY_KEY);
    return id;
  } catch (e) {
    return null;
  }
}

const LANES = [
  { key: "location",   label: "Where" },
  { key: "lodging",    label: "Lodging" },
  { key: "flights",    label: "Flights" },
  { key: "rental",     label: "Transportation", optional: true },
  { key: "activities", label: "Activities" },
];

const state = {
  name: "",
  start: null,
  end: null,
  events: [],
  segmentSize: "auto",
  tzAware: true,
  homeTz: "America/Los_Angeles",
  activeView: "main",       // "main" | "options"
  options: [],              // [{ id, name, events: [...] }]
  optionRangeStart: null,
  optionRangeEnd: null,
  shrunkDays: [],           // ISO dates the user has manually shrunk to a thin column
};

// --- date helpers (treat dates as plain calendar days, not instants) ---

function parseDay(iso) {
  if (!iso) return null;
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d);
}

function toISO(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

function dayDiff(a, b) {
  const ms = parseDay(b) - parseDay(a);
  return Math.round(ms / 86400000);
}

function fmtShort(date) {
  return date.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

const DOW = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

// --- timezone helpers ---

// Return UTC ms for a wall-clock time (y, m, d, h, mn) in tz.
function wallToUtc(y, m, d, h, mn, tz) {
  let guess = Date.UTC(y, m - 1, d, h, mn);
  for (let i = 0; i < 2; i++) {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: tz, hour12: false,
      year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit",
    }).formatToParts(new Date(guess));
    const get = t => +parts.find(p => p.type === t).value;
    let gh = get("hour"); if (gh === 24) gh = 0;
    const wallAtGuess = Date.UTC(get("year"), get("month") - 1, get("day"), gh, get("minute"));
    const offset = wallAtGuess - guess;
    guess = Date.UTC(y, m - 1, d, h, mn) - offset;
  }
  return guess;
}

function dayStartUtcInTz(dateStr, tz) {
  const [y, m, d] = dateStr.split("-").map(Number);
  return wallToUtc(y, m, d, 0, 0, tz);
}

function tzShortName(tz, dateStr) {
  const [y, m, d] = dateStr.split("-").map(Number);
  const utc = wallToUtc(y, m, d, 12, 0, tz);
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: tz, timeZoneName: "short",
  }).formatToParts(new Date(utc));
  const tn = parts.find(p => p.type === "timeZoneName");
  return tn ? tn.value : "";
}

// --- persistence ---

function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY());
    if (!raw) return seed();
    const parsed = demojibakeWalk(JSON.parse(raw));
    Object.assign(state, parsed);
  } catch (e) {
    seed();
  }
}

// Walk an object and try to repair any string field that looks like UTF-8
// bytes mistakenly read as Latin-1 (the '→' → 'â', '·' → 'Â·' family of
// corruptions). Idempotent and safe on clean strings — they don't match
// the marker regex and pass through.
const MOJI_MARKERS = /[Ââ-]/;
function tryDemojibake(s) {
  if (typeof s !== "string" || !MOJI_MARKERS.test(s)) return s;
  // Each character must fit in a single byte for re-decoding to make sense.
  const bytes = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) {
    const cp = s.charCodeAt(i);
    if (cp > 0xFF) return s;
    bytes[i] = cp;
  }
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch { return s; }
}
function demojibakeWalk(v) {
  if (typeof v === "string") return tryDemojibake(v);
  if (Array.isArray(v)) return v.map(demojibakeWalk);
  if (v && typeof v === "object") {
    const out = {};
    for (const [k, vv] of Object.entries(v)) out[k] = demojibakeWalk(vv);
    return out;
  }
  return v;
}

// True while bootstrapping is loading the trip — suppresses syncs that would
// fire from save() calls inside the load itself.
let BOOTING = true;

// Auto-add "Where" events between flights so the timeline always shows where
// you'll be. After every save we wipe the existing auto-locations and rebuild
// from the current flight list. User-edited or user-created locations are left
// alone (auto events carry _autoLoc: true; if the user edits one we strip the
// flag elsewhere so we won't clobber it next time).
function reconcileAutoLocations() {
  if (!Array.isArray(state.rejectedAutoLocs)) state.rejectedAutoLocs = [];
  state.events = state.events.filter(e => !e._autoLoc);
  const flights = state.events
    .filter(e => e.lane === "flights"
      && e.start && e.end
      && !(e.title || "").toLowerCase().endsWith("layover"))
    .sort((a, b) => (a.start + (a.startTime || "00:00")).localeCompare(b.start + (b.startTime || "00:00")));
  if (flights.length === 0) return;

  // Extract IATA code-pair from "KQ 427 SEA → GRU" / "SEA → GRU" / "SEA - GRU".
  const extractCodes = (title) => {
    if (!title) return null;
    const m = title.match(/\b([A-Z]{3})\s*(?:→|->|-|–)\s*([A-Z]{3})\b/);
    return m ? { from: m[1], to: m[2] } : null;
  };

  const hasManualLocationCovering = (start, end) => state.events.some(e =>
    e.lane === "location" && !e._autoLoc
    && e.start && e.end && e.start <= end && e.end >= start);
  const isRejected = (start, end) => state.rejectedAutoLocs.some(r =>
    r.start <= end && r.end >= start);

  const palette = ["violet", "amber", "rose", "emerald", "teal", "indigo"];
  let colorIdx = 0;
  const toCreate = [];

  for (let i = 0; i < flights.length; i++) {
    const f = flights[i];
    const codes = extractCodes(f.title);
    if (!codes) continue;
    const next = flights[i + 1];
    // Where they are between this flight's arrival and the next flight's departure
    // (or to trip end if this is the last flight).
    const start = f.end;
    const end = next ? next.start : (state.end || f.end);
    if (!start || !end || end < start) continue;
    if (hasManualLocationCovering(start, end)) continue;
    if (isRejected(start, end)) continue;
    const startTime = f.endTime || null;
    const endTime = next ? (next.startTime || null) : null;
    toCreate.push({
      id: uid(),
      title: codes.to,
      lane: "location",
      start,
      end,
      ...(startTime ? { startTime } : {}),
      ...(endTime ? { endTime } : {}),
      color: palette[colorIdx++ % palette.length],
      _autoLoc: true,
    });
  }
  if (toCreate.length) state.events.push(...toCreate);
}

function save() {
  reconcileAutoLocations();
  state.modifiedAt = Date.now();
  localStorage.setItem(STORAGE_KEY(), JSON.stringify(state));
  upsertRegistry({
    id: CURRENT_TRIP_ID,
    name: state.name || "Untitled trip",
    start: state.start || null,
    end: state.end || null,
  });
  if (!BOOTING) {
    scheduleSync();
  }
  updateSyncIndicator();
}

// Per-device UI state — current tab, collapsed days, segment-size dropdown,
// etc. Persists locally so it survives reloads on the same device, but does
// NOT bump modifiedAt or push to the worker. Otherwise just clicking a tab on
// device B with stale data would overwrite real edits made on device A.
function saveLocal() {
  localStorage.setItem(STORAGE_KEY(), JSON.stringify(state));
}

let SYNC_TIMER = null;
let SYNC_PENDING = false;
let SYNC_LAST_STATUS = null;  // "saved" | "saving" | "error" | null

function scheduleSync() {
  SYNC_PENDING = true;
  updateSyncIndicator();
  clearTimeout(SYNC_TIMER);
  SYNC_TIMER = setTimeout(syncNow, 1500);
}

async function syncNow() {
  SYNC_PENDING = false;
  const slug = tripSlug();
  if (!slug) return;
  // Not signed in (shouldn't happen — auth gate blocks). No-op.
  if (!window.fb?.user) {
    SYNC_LAST_STATUS = null;
    updateSyncIndicator();
    return;
  }
  // Not the owner — can't write.
  if (!CAN_EDIT) {
    SYNC_LAST_STATUS = null;
    updateSyncIndicator();
    return;
  }
  SYNC_LAST_STATUS = "saving";
  updateSyncIndicator();
  const res = await putTrip(slug, state);
  if (res.ok) {
    SYNC_LAST_STATUS = "saved";
  } else {
    SYNC_LAST_STATUS = "error";
  }
  updateSyncIndicator();
}

function updateSyncIndicator() {
  const el = document.getElementById("sync-indicator");
  if (!el) return;
  if (!CAN_EDIT) {
    el.textContent = CAN_SEE_PRICING ? "View only · prices visible" : "View only";
    el.className = "sync-indicator viewer";
    return;
  }
  if (SYNC_PENDING) { el.textContent = "Unsaved…"; el.className = "sync-indicator pending"; return; }
  if (SYNC_LAST_STATUS === "saving") { el.textContent = "Saving…"; el.className = "sync-indicator pending"; return; }
  if (SYNC_LAST_STATUS === "error") { el.textContent = "Sync error"; el.className = "sync-indicator error"; return; }
  if (SYNC_LAST_STATUS === "saved") { el.textContent = "Saved"; el.className = "sync-indicator saved"; return; }
  el.textContent = "";
  el.className = "sync-indicator";
}

function tripSlug() {
  const list = readRegistry();
  const t = list.find(x => x.id === CURRENT_TRIP_ID);
  return t?.slug || null;
}

// Pricing fields stripped from the public export (same fields the Pricing tab
// reads/writes — see line-item / split definitions above).
const PRICING_KEYS = ["lineItems", "priceSplit", "priceToken", "viewerToken"];

function stripPricing(obj) {
  const copy = JSON.parse(JSON.stringify(obj));
  for (const k of PRICING_KEYS) delete copy[k];
  return copy;
}

// Share button: show a dropdown with two URLs — the public trip link and a
// pricing-visible viewer link gated by a per-trip token. The token is
// generated on first request and persisted on the trip blob.
async function shareTrip() {
  const slug = tripSlug();
  if (!slug) { alert("This trip doesn't have a slug yet — give it a name first."); return; }
  showShareDropdown();
}

function showShareDropdown() {
  document.getElementById("share-dropdown")?.remove();
  const panel = document.createElement("div");
  panel.id = "share-dropdown";
  panel.className = "export-links";
  const pricingRow = "";
  // Public live link (anonymous viewers, no account needed). Owner-only.
  const publicLinkRow = IS_OWNER ? `
    <div class="export-links-row" style="margin-top:12px; padding-top:12px; border-top: 1px solid var(--line);">
      <label>Public link (no sign-in needed)</label>
      <div data-public-link-state>
        <button type="button" data-public-create>Create public link</button>
        <span data-public-link-msg style="margin-left:8px;font-size:12px;color:var(--muted);"></span>
      </div>
    </div>
  ` : "";

  // Only the owner can re-share (per rules). Editor can edit content but
  // not change who sees what.
  const shareByEmailRow = IS_OWNER ? `
    <div class="export-links-row" style="margin-top:12px; padding-top:12px; border-top: 1px solid var(--line);">
      <label>Share with a myitin user</label>
      <div class="export-links-input">
        <input type="email" data-which="share-email" placeholder="friend@example.com" />
        <select data-which="share-level" style="padding:6px 8px; border:1px solid var(--line); border-radius:4px; font-size:13px;">
          <option value="viewer" selected>View (no prices)</option>
          <option value="viewer-pricing">View + prices</option>
          <option value="editor">Editor (full access)</option>
        </select>
        <button type="button" data-share-email>Share</button>
      </div>
      <div data-share-email-msg style="font-size:12px;margin-top:6px;min-height:16px;"></div>
      <div data-shared-with-list style="font-size:12px;margin-top:6px;"></div>
    </div>
  ` : "";
  panel.innerHTML = `
    ${publicLinkRow}
    ${shareByEmailRow}
    <button type="button" class="export-links-close" aria-label="Close">×</button>
  `;
  // Public live link create / disable.
  const publicState = panel.querySelector('div[data-public-link-state]');
  const publicMsg = panel.querySelector('span[data-public-link-msg]');
  function publicLinkUrl(token) {
    return `${location.origin}${location.pathname}?p=${encodeURIComponent(token)}`;
  }
  async function renderPublicLink() {
    if (!publicState || !IS_OWNER) return;
    const slug = tripSlug();
    const res = await window.fb.loadTrip(slug);
    const token = res.body?.publicToken;
    if (!token) {
      publicState.innerHTML = `
        <button type="button" data-public-create>Create public link</button>
        <span data-public-link-msg style="margin-left:8px;font-size:12px;color:var(--muted);"></span>
      `;
      publicState.querySelector('[data-public-create]').addEventListener("click", async () => {
        try {
          await window.fb.createPublicLink(slug);
          renderPublicLink();
        } catch (e) {
          publicState.querySelector('[data-public-link-msg]').textContent = e.message;
        }
      });
    } else {
      const url = publicLinkUrl(token);
      publicState.innerHTML = `
        <div class="export-links-input">
          <input type="text" readonly value="${url.replace(/"/g, "&quot;")}" />
          <button type="button" data-public-copy>Copy</button>
          <button type="button" data-public-wa title="Send via WhatsApp">WhatsApp</button>
          <button type="button" data-public-disable class="ghost">Disable</button>
        </div>
        <span data-public-link-msg style="font-size:12px;color:var(--muted);"></span>
      `;
      publicState.querySelector('[data-public-wa]').addEventListener("click", () => {
        const tripName = state.name || "Trip";
        const text = `${tripName}\n${url}`;
        window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, "_blank", "noopener");
      });
      publicState.querySelector('[data-public-copy]').addEventListener("click", async () => {
        const btn = publicState.querySelector('[data-public-copy]');
        try {
          await navigator.clipboard.writeText(url);
          const orig = btn.textContent;
          btn.textContent = "Copied";
          setTimeout(() => { btn.textContent = orig; }, 1200);
        } catch (_) {
          const inp = publicState.querySelector('input');
          inp.select(); document.execCommand("copy");
        }
      });
      publicState.querySelector('[data-public-disable]').addEventListener("click", async () => {
        if (!confirm("Disable the public link? Anyone with the URL will lose access.")) return;
        try {
          await window.fb.removePublicLink(slug);
          renderPublicLink();
        } catch (e) {
          publicState.querySelector('[data-public-link-msg]').textContent = e.message;
        }
      });
    }
  }
  if (publicState) renderPublicLink();

  // Share-by-email handler + current permissions listing.
  const shareInput = panel.querySelector('input[data-which=share-email]');
  const shareLevel = panel.querySelector('select[data-which=share-level]');
  const shareBtn = panel.querySelector('button[data-share-email]');
  const shareMsg = panel.querySelector('div[data-share-email-msg]');
  const sharedList = panel.querySelector('div[data-shared-with-list]');
  const LEVEL_LABELS = {
    "viewer": "View",
    "viewer-pricing": "View + prices",
    "editor": "Editor",
  };
  async function renderSharedWith() {
    if (!sharedList) return;
    const slug = tripSlug();
    const res = await window.fb.loadTrip(slug);
    const perms = (res.body?.permissions) || {};
    const uids = Object.keys(perms);
    if (!uids.length) { sharedList.textContent = "Not shared with anyone yet."; return; }
    sharedList.innerHTML = '<div style="margin-bottom:4px;">Shared with:</div>';
    for (const uid of uids) {
      const prof = await window.fb.lookupProfile(uid);
      const row = document.createElement("div");
      row.style.cssText = "display:flex;align-items:center;gap:6px;margin:3px 0;";
      const who = document.createElement("span");
      who.textContent = prof.displayName || prof.email || uid.slice(0, 6);
      who.style.cssText = "flex:1;background:var(--panel-2);border:1px solid var(--line);border-radius:999px;padding:2px 10px;";
      row.appendChild(who);
      const sel = document.createElement("select");
      sel.style.cssText = "padding:2px 6px; border:1px solid var(--line); border-radius:4px; font-size:12px;";
      for (const lvl of ["viewer", "viewer-pricing", "editor"]) {
        const o = document.createElement("option");
        o.value = lvl; o.textContent = LEVEL_LABELS[lvl];
        if (perms[uid] === lvl) o.selected = true;
        sel.appendChild(o);
      }
      sel.addEventListener("change", async () => {
        try {
          await window.fb.setTripPermissionLevel(slug, uid, sel.value);
          shareMsg.textContent = `Updated ${who.textContent} to ${LEVEL_LABELS[sel.value]}.`;
          shareMsg.style.color = "#2a8b4a";
        } catch (e) {
          shareMsg.textContent = e.message; shareMsg.style.color = "#b12a48";
          renderSharedWith();
        }
      });
      row.appendChild(sel);
      const x = document.createElement("button");
      x.textContent = "×";
      x.title = "Unshare";
      x.style.cssText = "background:none;border:none;cursor:pointer;color:var(--muted);padding:0 4px;font-size:16px;";
      x.addEventListener("click", async () => {
        try { await window.fb.unshareTripWithUid(slug, uid); await renderSharedWith(); }
        catch (e) { shareMsg.textContent = e.message; shareMsg.style.color = "#b12a48"; }
      });
      row.appendChild(x);
      sharedList.appendChild(row);
    }
  }
  if (shareBtn) {
    shareBtn.addEventListener("click", async () => {
      const email = (shareInput.value || "").trim();
      const level = shareLevel?.value || "viewer";
      if (!email) { shareMsg.textContent = "Enter an email."; shareMsg.style.color = "#b12a48"; return; }
      try {
        await window.fb.shareTripWithEmail(tripSlug(), email, level);
        shareMsg.textContent = `Shared with ${email} (${LEVEL_LABELS[level]}).`;
        shareMsg.style.color = "#2a8b4a";
        shareInput.value = "";
        renderSharedWith();
      } catch (e) {
        shareMsg.textContent = e.message;
        shareMsg.style.color = "#b12a48";
      }
    });
    renderSharedWith();
  }

  panel.querySelector(".export-links-close").addEventListener("click", () => panel.remove());
  document.getElementById("share-btn")?.parentElement?.appendChild(panel);
  setTimeout(() => {
    const onDoc = (e) => {
      if (!panel.contains(e.target) && e.target.id !== "share-btn") {
        panel.remove();
        document.removeEventListener("click", onDoc);
      }
    };
    document.addEventListener("click", onDoc);
  }, 0);
}

function seed() {
  state.name = "Zanzibar 2026";
  state.start = "2026-06-17";
  state.end = "2026-07-09";
  state.tzAware = true;
  state.homeTz = "America/Los_Angeles";
  state.segmentSize = "auto";
  state.events = [
    // Where ----------------------------------------------------------
    { id: uid(), title: "Zanzibar", lane: "location", color: "teal",
      start: "2026-06-19", startTime: "07:15", startTz: "Africa/Dar_es_Salaam",
      end:   "2026-06-24", endTime:   "12:10", endTz:   "Africa/Dar_es_Salaam",
      notes: "" },
    { id: uid(), title: "Tanzania safari", lane: "location", color: "amber",
      start: "2026-06-24", startTime: "13:15", startTz: "Africa/Dar_es_Salaam",
      end:   "2026-07-01", endTime:   "23:59", endTz:   "Africa/Dar_es_Salaam",
      notes: "8-day premium safari · Serengeti / Ngorongoro / Tarangire." },

    // Flights & layovers ---------------------------------------------
    { id: uid(), title: "LH493 YVR → FRA", lane: "flights", color: "indigo",
      start: "2026-06-17", startTime: "16:15", startTz: "America/Los_Angeles",
      end:   "2026-06-18", endTime:   "11:00", endTz:   "Europe/Berlin",
      notes: "Lufthansa · Business (P)" },
    { id: uid(), title: "FRA layover", lane: "flights", color: "grey",
      start: "2026-06-18", startTime: "11:00", startTz: "Europe/Berlin",
      end:   "2026-06-18", endTime:   "19:35", endTz:   "Europe/Berlin",
      notes: "Connection between LH493 and 4Y134." },
    { id: uid(), title: "4Y134 FRA → ZNZ", lane: "flights", color: "indigo",
      start: "2026-06-18", startTime: "19:35", startTz: "Europe/Berlin",
      end:   "2026-06-19", endTime:   "07:15", endTz:   "Africa/Dar_es_Salaam",
      notes: "Discover Airlines · Business (P)" },
    { id: uid(), title: "UI 628 ZNZ → ARK", lane: "flights", color: "indigo",
      start: "2026-06-24", startTime: "12:10", startTz: "Africa/Dar_es_Salaam",
      end:   "2026-06-24", endTime:   "13:15", endTz:   "Africa/Dar_es_Salaam",
      notes: "Auric Air · Economy · De Havilland Dash-8 · $242" },
    { id: uid(), title: "4Y131 ZNZ → FRA", lane: "flights", color: "rose",
      start: "2026-07-08", startTime: "08:00", startTz: "Africa/Dar_es_Salaam",
      end:   "2026-07-08", endTime:   "16:05", endTz:   "Europe/Berlin",
      notes: "Discover Airlines · Business (Z)" },
    { id: uid(), title: "FRA overnight", lane: "flights", color: "grey",
      start: "2026-07-08", startTime: "16:05", startTz: "Europe/Berlin",
      end:   "2026-07-09", endTime:   "10:45", endTz:   "Europe/Berlin",
      notes: "Overnight in FRA. Check in at the Lufthansa ticket counter." },
    { id: uid(), title: "UA8717 FRA → SEA", lane: "flights", color: "rose",
      start: "2026-07-09", startTime: "10:45", startTz: "Europe/Berlin",
      end:   "2026-07-09", endTime:   "11:55", endTz:   "America/Los_Angeles",
      notes: "Lufthansa (operating UA8717) · Business (Z)" },

    // Lodging — safari lodges ---------------------------------------
    { id: uid(), title: "Lake Duluti Safari Lodge", lane: "lodging", color: "amber",
      start: "2026-06-24", end: "2026-06-24",
      notes: "Arusha · Day 1 night" },
    { id: uid(), title: "Acacia Farm Lodge", lane: "lodging", color: "amber",
      start: "2026-06-25", end: "2026-06-25",
      notes: "Karatu · Day 2 night" },
    { id: uid(), title: "Conserve Safari Camp", lane: "lodging", color: "amber",
      start: "2026-06-26", end: "2026-06-28",
      notes: "Serengeti · Days 3–5 (3 nights)" },
    { id: uid(), title: "Acacia Farm Lodge", lane: "lodging", color: "amber",
      start: "2026-06-29", end: "2026-06-29",
      notes: "Karatu · Day 6 night (second stay)" },
    { id: uid(), title: "Lake Duluti Safari Lodge", lane: "lodging", color: "amber",
      start: "2026-06-30", end: "2026-06-30",
      notes: "Tarangire · Day 7 night — placeholder using the Arusha lodge (travel agency arranged)." },

    // Activities — Tanzania safari day-by-day ------------------------
    { id: uid(), title: "Day 1 · Arrival in Arusha", lane: "activities", color: "emerald",
      start: "2026-06-24", end: "2026-06-24",
      notes: "Lodging: 4★ hotel in Arusha." },
    { id: uid(), title: "Day 2 · Ngorongoro Crater", lane: "activities", color: "emerald",
      start: "2026-06-25", end: "2026-06-25",
      notes: "Arusha → Ngorongoro Crater. Lodging: 4★ in Karatu." },
    { id: uid(), title: "Day 3 · Drive to Serengeti", lane: "activities", color: "emerald",
      start: "2026-06-26", end: "2026-06-26",
      notes: "Ngorongoro → Serengeti National Park. Lodging: 4★ in Serengeti Plains." },
    { id: uid(), title: "Day 4 · Central Serengeti", lane: "activities", color: "emerald",
      start: "2026-06-27", end: "2026-06-27",
      notes: "Central Serengeti exploration. Lodging: 4★ in Serengeti Plains." },
    { id: uid(), title: "Day 5 · Serengeti game drives", lane: "activities", color: "emerald",
      start: "2026-06-28", end: "2026-06-28",
      notes: "Serengeti game drives. Lodging: 4★ in Serengeti Plains." },
    { id: uid(), title: "Day 6 · Lake Eyasi & Karatu", lane: "activities", color: "emerald",
      start: "2026-06-29", end: "2026-06-29",
      notes: "Serengeti → Lake Eyasi → Karatu. Lodging: 4★ in Karatu." },
    { id: uid(), title: "Day 7 · Tarangire", lane: "activities", color: "emerald",
      start: "2026-06-30", end: "2026-06-30",
      notes: "Karatu → Tarangire National Park. Lodging: 4★ in Tarangire." },
    { id: uid(), title: "Day 8 · Departure from Arusha", lane: "activities", color: "emerald",
      start: "2026-07-01", end: "2026-07-01",
      notes: "Tarangire → Arusha → Departure." },
  ];
}

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

// --- TZ map per day: which timezone is the user "in" on each day ---

// IATA airport code → IANA timezone. Covers the airports we're most likely to
// fly through; unknown codes fall back to the home tz.
const AIRPORT_TZ = {
  // US Eastern
  MCO:"America/New_York", JFK:"America/New_York", LGA:"America/New_York", EWR:"America/New_York",
  BOS:"America/New_York", IAD:"America/New_York", DCA:"America/New_York", BWI:"America/New_York",
  PHL:"America/New_York", ATL:"America/New_York", MIA:"America/New_York", FLL:"America/New_York",
  PBI:"America/New_York", RDU:"America/New_York", CLT:"America/New_York", TPA:"America/New_York",
  RSW:"America/New_York", JAX:"America/New_York", SAV:"America/New_York", CMH:"America/New_York",
  CLE:"America/New_York", PIT:"America/New_York", BUF:"America/New_York", ROC:"America/New_York",
  SYR:"America/New_York", BTV:"America/New_York", DTW:"America/Detroit",
  // US Central
  ORD:"America/Chicago", MDW:"America/Chicago", DFW:"America/Chicago", DAL:"America/Chicago",
  IAH:"America/Chicago", HOU:"America/Chicago", AUS:"America/Chicago", SAT:"America/Chicago",
  MSY:"America/Chicago", MSP:"America/Chicago", MCI:"America/Chicago", STL:"America/Chicago",
  BNA:"America/Chicago", MEM:"America/Chicago", OMA:"America/Chicago", OKC:"America/Chicago",
  TUL:"America/Chicago", LIT:"America/Chicago", MKE:"America/Chicago", IND:"America/Indianapolis",
  // US Mountain
  DEN:"America/Denver", SLC:"America/Denver", ABQ:"America/Denver", BIL:"America/Denver",
  COS:"America/Denver", BOI:"America/Boise",
  // US Mountain (Arizona — no DST)
  PHX:"America/Phoenix", TUS:"America/Phoenix",
  // US Pacific
  LAX:"America/Los_Angeles", SFO:"America/Los_Angeles", SAN:"America/Los_Angeles",
  LAS:"America/Los_Angeles", OAK:"America/Los_Angeles", SJC:"America/Los_Angeles",
  SMF:"America/Los_Angeles", BUR:"America/Los_Angeles", ONT:"America/Los_Angeles",
  SNA:"America/Los_Angeles", PDX:"America/Los_Angeles", SEA:"America/Los_Angeles",
  RNO:"America/Los_Angeles",
  // US Alaska / Hawaii
  ANC:"America/Anchorage", FAI:"America/Anchorage", JNU:"America/Juneau",
  HNL:"Pacific/Honolulu", OGG:"Pacific/Honolulu", LIH:"Pacific/Honolulu", KOA:"Pacific/Honolulu",
  // Canada
  YVR:"America/Los_Angeles", YYZ:"America/Toronto", YUL:"America/Toronto",
  YOW:"America/Toronto", YYC:"America/Edmonton", YEG:"America/Edmonton",
  YHZ:"America/Halifax", YWG:"America/Winnipeg",
  // Mexico / Latin America
  MEX:"America/Mexico_City", CUN:"America/Cancun", GDL:"America/Mexico_City",
  MTY:"America/Monterrey", SJD:"America/Mazatlan",
  GRU:"America/Sao_Paulo", GIG:"America/Sao_Paulo", BSB:"America/Sao_Paulo",
  EZE:"America/Argentina/Buenos_Aires", AEP:"America/Argentina/Buenos_Aires",
  SCL:"America/Santiago", LIM:"America/Lima", BOG:"America/Bogota",
  UIO:"America/Guayaquil", PTY:"America/Panama", SJO:"America/Costa_Rica",
  // Europe
  LHR:"Europe/London", LGW:"Europe/London", STN:"Europe/London", LTN:"Europe/London",
  MAN:"Europe/London", BHX:"Europe/London", EDI:"Europe/London", GLA:"Europe/London",
  DUB:"Europe/Dublin",
  CDG:"Europe/Paris", ORY:"Europe/Paris", NCE:"Europe/Paris", LYS:"Europe/Paris",
  FRA:"Europe/Berlin", MUC:"Europe/Berlin", HAM:"Europe/Berlin", DUS:"Europe/Berlin",
  BER:"Europe/Berlin", TXL:"Europe/Berlin",
  AMS:"Europe/Amsterdam", BRU:"Europe/Brussels",
  MAD:"Europe/Madrid", BCN:"Europe/Madrid",
  FCO:"Europe/Rome", MXP:"Europe/Rome", VCE:"Europe/Rome", NAP:"Europe/Rome",
  ZRH:"Europe/Zurich", GVA:"Europe/Zurich",
  VIE:"Europe/Vienna", PRG:"Europe/Prague", WAW:"Europe/Warsaw", BUD:"Europe/Budapest",
  CPH:"Europe/Copenhagen", OSL:"Europe/Oslo", ARN:"Europe/Stockholm", HEL:"Europe/Helsinki",
  ATH:"Europe/Athens", IST:"Europe/Istanbul",
  LIS:"Europe/Lisbon", OPO:"Europe/Lisbon",
  // Middle East
  DXB:"Asia/Dubai", AUH:"Asia/Dubai", DOH:"Asia/Qatar", KWI:"Asia/Kuwait",
  AMM:"Asia/Amman", TLV:"Asia/Jerusalem", RUH:"Asia/Riyadh", JED:"Asia/Riyadh",
  // Asia
  NRT:"Asia/Tokyo", HND:"Asia/Tokyo", KIX:"Asia/Tokyo", NGO:"Asia/Tokyo",
  ICN:"Asia/Seoul", GMP:"Asia/Seoul",
  PEK:"Asia/Shanghai", PVG:"Asia/Shanghai", CAN:"Asia/Shanghai", SZX:"Asia/Shanghai",
  HKG:"Asia/Hong_Kong", MFM:"Asia/Macau", TPE:"Asia/Taipei",
  SIN:"Asia/Singapore", KUL:"Asia/Kuala_Lumpur", BKK:"Asia/Bangkok",
  HAN:"Asia/Ho_Chi_Minh", SGN:"Asia/Ho_Chi_Minh", MNL:"Asia/Manila",
  CGK:"Asia/Jakarta", DPS:"Asia/Makassar",
  DEL:"Asia/Kolkata", BOM:"Asia/Kolkata", BLR:"Asia/Kolkata", MAA:"Asia/Kolkata", HYD:"Asia/Kolkata",
  // Africa
  JNB:"Africa/Johannesburg", CPT:"Africa/Johannesburg", DUR:"Africa/Johannesburg",
  NBO:"Africa/Nairobi", ADD:"Africa/Addis_Ababa",
  JRO:"Africa/Dar_es_Salaam", ARK:"Africa/Dar_es_Salaam",
  ZNZ:"Africa/Dar_es_Salaam", DAR:"Africa/Dar_es_Salaam",
  CAI:"Africa/Cairo", LOS:"Africa/Lagos", ABV:"Africa/Lagos",
  CMN:"Africa/Casablanca", RAK:"Africa/Casablanca",
  SEZ:"Indian/Mahe", MRU:"Indian/Mauritius",
  // Oceania
  SYD:"Australia/Sydney", MEL:"Australia/Melbourne", BNE:"Australia/Brisbane",
  PER:"Australia/Perth", ADL:"Australia/Adelaide",
  AKL:"Pacific/Auckland", WLG:"Pacific/Auckland", CHC:"Pacific/Auckland",
  NAN:"Pacific/Fiji", PPT:"Pacific/Tahiti",
};

// IATA airport code → [lat, lon, city name] for the flight map. Covers the
// same set as AIRPORT_TZ; unknown codes get skipped on the map.
const AIRPORT_COORDS = {
  MCO:[28.43,-81.31,"Orlando"], JFK:[40.64,-73.78,"New York JFK"], LGA:[40.78,-73.87,"New York LGA"],
  EWR:[40.69,-74.17,"Newark"], BOS:[42.36,-71.01,"Boston"], IAD:[38.95,-77.46,"Washington IAD"],
  DCA:[38.85,-77.04,"Washington DCA"], BWI:[39.18,-76.67,"Baltimore"], PHL:[39.87,-75.24,"Philadelphia"],
  ATL:[33.64,-84.43,"Atlanta"], MIA:[25.79,-80.29,"Miami"], FLL:[26.07,-80.15,"Fort Lauderdale"],
  PBI:[26.68,-80.10,"West Palm Beach"], RDU:[35.88,-78.79,"Raleigh"], CLT:[35.21,-80.94,"Charlotte"],
  TPA:[27.98,-82.53,"Tampa"], RSW:[26.54,-81.76,"Fort Myers"], JAX:[30.49,-81.69,"Jacksonville"],
  CMH:[39.99,-82.89,"Columbus"], CLE:[41.41,-81.85,"Cleveland"], PIT:[40.49,-80.23,"Pittsburgh"],
  BUF:[42.94,-78.73,"Buffalo"], DTW:[42.21,-83.35,"Detroit"],
  ORD:[41.98,-87.91,"Chicago ORD"], MDW:[41.79,-87.75,"Chicago MDW"], DFW:[32.90,-97.04,"Dallas DFW"],
  DAL:[32.85,-96.85,"Dallas DAL"], IAH:[29.99,-95.34,"Houston IAH"], HOU:[29.65,-95.28,"Houston HOU"],
  AUS:[30.19,-97.67,"Austin"], SAT:[29.53,-98.47,"San Antonio"], MSY:[29.99,-90.26,"New Orleans"],
  MSP:[44.88,-93.22,"Minneapolis"], MCI:[39.30,-94.71,"Kansas City"], STL:[38.75,-90.37,"St. Louis"],
  BNA:[36.12,-86.68,"Nashville"], MEM:[35.04,-89.98,"Memphis"], OMA:[41.30,-95.89,"Omaha"],
  OKC:[35.39,-97.60,"Oklahoma City"], MKE:[42.95,-87.90,"Milwaukee"], IND:[39.72,-86.29,"Indianapolis"],
  DEN:[39.86,-104.67,"Denver"], SLC:[40.79,-111.98,"Salt Lake City"], ABQ:[35.04,-106.61,"Albuquerque"],
  BIL:[45.81,-108.54,"Billings"], COS:[38.81,-104.70,"Colorado Springs"], BOI:[43.56,-116.22,"Boise"],
  PHX:[33.43,-112.01,"Phoenix"], TUS:[32.12,-110.94,"Tucson"],
  LAX:[33.94,-118.41,"Los Angeles"], SFO:[37.62,-122.38,"San Francisco"], SAN:[32.73,-117.19,"San Diego"],
  LAS:[36.08,-115.15,"Las Vegas"], OAK:[37.72,-122.22,"Oakland"], SJC:[37.36,-121.93,"San Jose"],
  SMF:[38.69,-121.59,"Sacramento"], BUR:[34.20,-118.36,"Burbank"], ONT:[34.06,-117.60,"Ontario CA"],
  SNA:[33.68,-117.87,"Santa Ana"], PDX:[45.59,-122.60,"Portland"], SEA:[47.45,-122.31,"Seattle"],
  RNO:[39.50,-119.77,"Reno"],
  ANC:[61.17,-149.99,"Anchorage"], FAI:[64.81,-147.86,"Fairbanks"],
  HNL:[21.32,-157.92,"Honolulu"], OGG:[20.90,-156.43,"Maui"], LIH:[21.98,-159.34,"Kauai"], KOA:[19.74,-156.05,"Kona"],
  YVR:[49.19,-123.18,"Vancouver"], YYZ:[43.68,-79.63,"Toronto"], YUL:[45.47,-73.74,"Montreal"],
  YOW:[45.32,-75.67,"Ottawa"], YYC:[51.11,-114.02,"Calgary"], YEG:[53.31,-113.58,"Edmonton"],
  YHZ:[44.88,-63.51,"Halifax"], YWG:[49.91,-97.24,"Winnipeg"],
  MEX:[19.44,-99.07,"Mexico City"], CUN:[21.04,-86.87,"Cancun"], GDL:[20.52,-103.31,"Guadalajara"],
  MTY:[25.78,-100.11,"Monterrey"], SJD:[23.15,-109.72,"San Jose del Cabo"],
  GRU:[-23.43,-46.48,"São Paulo"], GIG:[-22.81,-43.25,"Rio de Janeiro"], BSB:[-15.87,-47.92,"Brasília"],
  EZE:[-34.82,-58.54,"Buenos Aires EZE"], AEP:[-34.56,-58.42,"Buenos Aires AEP"],
  SCL:[-33.39,-70.79,"Santiago"], LIM:[-12.02,-77.11,"Lima"], BOG:[4.70,-74.14,"Bogotá"],
  UIO:[-0.13,-78.36,"Quito"], PTY:[9.07,-79.38,"Panama City"], SJO:[9.99,-84.21,"San José CR"],
  LHR:[51.47,-0.45,"London Heathrow"], LGW:[51.15,-0.19,"London Gatwick"], STN:[51.89,0.24,"London Stansted"],
  LTN:[51.87,-0.37,"London Luton"], MAN:[53.36,-2.27,"Manchester"], EDI:[55.95,-3.37,"Edinburgh"],
  DUB:[53.43,-6.27,"Dublin"],
  CDG:[49.00,2.55,"Paris CDG"], ORY:[48.72,2.38,"Paris Orly"], NCE:[43.66,7.21,"Nice"], LYS:[45.73,5.08,"Lyon"],
  FRA:[50.04,8.56,"Frankfurt"], MUC:[48.35,11.79,"Munich"], HAM:[53.63,10.00,"Hamburg"],
  DUS:[51.29,6.77,"Düsseldorf"], BER:[52.36,13.51,"Berlin"], TXL:[52.55,13.29,"Berlin TXL"],
  AMS:[52.31,4.76,"Amsterdam"], BRU:[50.90,4.48,"Brussels"],
  MAD:[40.49,-3.57,"Madrid"], BCN:[41.30,2.08,"Barcelona"],
  FCO:[41.80,12.25,"Rome"], MXP:[45.63,8.72,"Milan"], VCE:[45.51,12.35,"Venice"], NAP:[40.89,14.29,"Naples"],
  ZRH:[47.46,8.55,"Zurich"], GVA:[46.24,6.11,"Geneva"],
  VIE:[48.11,16.57,"Vienna"], PRG:[50.10,14.26,"Prague"], WAW:[52.17,20.97,"Warsaw"], BUD:[47.44,19.26,"Budapest"],
  CPH:[55.62,12.66,"Copenhagen"], OSL:[60.19,11.10,"Oslo"], ARN:[59.65,17.92,"Stockholm"], HEL:[60.32,24.96,"Helsinki"],
  ATH:[37.94,23.95,"Athens"], IST:[40.98,28.81,"Istanbul"],
  LIS:[38.77,-9.13,"Lisbon"], OPO:[41.24,-8.68,"Porto"],
  DXB:[25.25,55.36,"Dubai"], AUH:[24.43,54.65,"Abu Dhabi"], DOH:[25.27,51.62,"Doha"],
  KWI:[29.23,47.97,"Kuwait"], AMM:[31.72,35.99,"Amman"], TLV:[32.01,34.89,"Tel Aviv"],
  RUH:[24.96,46.70,"Riyadh"], JED:[21.68,39.15,"Jeddah"],
  NRT:[35.77,140.39,"Tokyo Narita"], HND:[35.55,139.78,"Tokyo Haneda"], KIX:[34.43,135.24,"Osaka"], NGO:[34.86,136.81,"Nagoya"],
  ICN:[37.46,126.44,"Seoul ICN"], GMP:[37.56,126.79,"Seoul GMP"],
  PEK:[40.08,116.58,"Beijing"], PVG:[31.14,121.81,"Shanghai"], CAN:[23.39,113.31,"Guangzhou"], SZX:[22.64,113.81,"Shenzhen"],
  HKG:[22.31,113.91,"Hong Kong"], MFM:[22.15,113.59,"Macau"], TPE:[25.08,121.23,"Taipei"],
  SIN:[1.36,103.99,"Singapore"], KUL:[2.74,101.71,"Kuala Lumpur"], BKK:[13.69,100.75,"Bangkok"],
  HAN:[21.22,105.81,"Hanoi"], SGN:[10.82,106.66,"Ho Chi Minh"], MNL:[14.51,121.02,"Manila"],
  CGK:[-6.13,106.66,"Jakarta"], DPS:[-8.75,115.17,"Bali"],
  DEL:[28.56,77.10,"Delhi"], BOM:[19.09,72.87,"Mumbai"], BLR:[13.20,77.71,"Bangalore"],
  MAA:[12.99,80.17,"Chennai"], HYD:[17.24,78.43,"Hyderabad"],
  JNB:[-26.13,28.24,"Johannesburg"], CPT:[-33.97,18.60,"Cape Town"], DUR:[-29.61,31.12,"Durban"],
  NBO:[-1.32,36.93,"Nairobi"], JRO:[-3.43,37.07,"Kilimanjaro"], ZNZ:[-6.22,39.22,"Zanzibar"],
  DAR:[-6.88,39.20,"Dar es Salaam"], ARK:[-3.37,36.63,"Arusha"], ADD:[8.98,38.80,"Addis Ababa"],
  CAI:[30.11,31.41,"Cairo"], LOS:[6.58,3.32,"Lagos"], ABV:[9.01,7.27,"Abuja"],
  CMN:[33.37,-7.59,"Casablanca"], RAK:[31.61,-8.04,"Marrakech"],
  SEZ:[-4.67,55.52,"Mahé"], MRU:[-20.43,57.68,"Mauritius"],
  SYD:[-33.95,151.18,"Sydney"], MEL:[-37.67,144.84,"Melbourne"], BNE:[-27.38,153.12,"Brisbane"],
  PER:[-31.94,115.97,"Perth"], ADL:[-34.95,138.53,"Adelaide"],
  AKL:[-37.01,174.79,"Auckland"], WLG:[-41.33,174.81,"Wellington"], CHC:[-43.49,172.53,"Christchurch"],
  NAN:[-17.76,177.45,"Nadi"], PPT:[-17.55,-149.61,"Tahiti"],
};

// Try to resolve a string ("MCO", "Orlando MCO", "Orlando") to an IANA tz via
// the airport map. Matches the first 3-letter all-caps code, then falls through.
function tzFromIataLike(text) {
  if (!text) return null;
  // Try whole string first (case insensitive), then any standalone 3-letter token.
  const up = text.toUpperCase();
  if (AIRPORT_TZ[up]) return AIRPORT_TZ[up];
  const m = up.match(/\b([A-Z]{3})\b/g);
  if (m) {
    for (const code of m) if (AIRPORT_TZ[code]) return AIRPORT_TZ[code];
  }
  return null;
}

function computeDayTzMap(start, end, events, homeTz, tzAware) {
  const map = {};
  let cur = parseDay(start);
  const endDate = parseDay(end);
  let currentTz = homeTz;

  // Build flight arrivals (with inferred destination tz from IATA codes if
  // smart-parse didn't set endTz). And location events with matching titles
  // override the tz for their full date range.
  const arrByDate = {};
  const locOverrides = []; // [{start, end, tz}]
  if (tzAware) {
    for (const ev of events) {
      if (ev.lane === "flights" && ev.end && !(ev.title || "").toLowerCase().endsWith("layover")) {
        // For flights, the arrival tz = the destination airport. Pull the
        // arrival code out of "SEA → MCO" so we don't accidentally use the
        // origin tz.
        const m = (ev.title || "").match(/\b([A-Z]{3})\s*(?:→|->|-|–)\s*([A-Z]{3})\b/);
        const destTz = m ? AIRPORT_TZ[m[2]] : null;
        const tz = ev.endTz || destTz || tzFromIataLike(ev.notes);
        if (tz) (arrByDate[ev.end] ||= []).push({ ...ev, _resolvedTz: tz });
      }
      if (ev.lane === "location" && ev.start && ev.end) {
        const tz = ev.endTz || ev.startTz || tzFromIataLike(ev.title) || tzFromIataLike(ev.notes);
        if (tz) locOverrides.push({ start: ev.start, end: ev.end, tz });
      }
    }
  }

  while (cur <= endDate) {
    const ds = toISO(cur);
    if (tzAware && arrByDate[ds] && arrByDate[ds].length) {
      const latest = arrByDate[ds].slice().sort((a, b) =>
        (a.endTime || "00:00").localeCompare(b.endTime || "00:00")).pop();
      currentTz = latest._resolvedTz || latest.endTz;
    }
    // Location event overlapping this day takes priority — that's where the
    // user *is* on that day, regardless of when the last flight landed.
    let dayTz = currentTz;
    for (const o of locOverrides) {
      if (o.start <= ds && o.end >= ds) { dayTz = o.tz; break; }
    }
    map[ds] = dayTz;
    cur = addDays(cur, 1);
  }
  return map;
}

// --- UTC bounds for an event ---

function eventBounds(ev, dayTzMap, homeTz, tzAware) {
  let tzS = ev.startTz;
  let tzE = ev.endTz;
  // For flights without explicit per-leg tz, infer from the IATA codes in
  // the title ("SEA → MCO") so the bar's UTC duration is the real flight
  // duration regardless of which view mode is active.
  if (ev.lane === "flights" && (!tzS || !tzE)) {
    const m = (ev.title || "").match(/\b([A-Z]{3})\s*(?:→|->|-|–)\s*([A-Z]{3})\b/);
    if (m) {
      if (!tzS && AIRPORT_TZ[m[1]]) tzS = AIRPORT_TZ[m[1]];
      if (!tzE && AIRPORT_TZ[m[2]]) tzE = AIRPORT_TZ[m[2]];
    }
  }
  tzS = tzS || dayTzMap[ev.start] || homeTz;
  tzE = tzE || dayTzMap[ev.end]   || homeTz;
  // When Follow timezone is off, events that have no explicit clock times
  // (locations, all-day activities, lodging without times) should align to
  // the calendar grid in home tz — otherwise their UTC midnight in
  // destination tz lands at random PST hours and pulls the bar off the day.
  if (!tzAware && !ev.startTime && !ev.endTime) {
    tzS = homeTz;
    tzE = homeTz;
  }
  const [sy, sm, sd] = ev.start.split("-").map(Number);
  const [ey, em, ed] = ev.end.split("-").map(Number);
  // Lodging without explicit times: pack against 3pm check-in / 11am
  // check-out so consecutive hotels on the same changeover day don't get
  // treated as overlapping and stacked on separate rows.
  if (ev.lane === "lodging" && !ev.startTime && !ev.endTime) {
    const sUtc = wallToUtc(sy, sm, sd, 15, 0, tzS);
    let eUtc;
    if (ev.start === ev.end) {
      const next = toISO(addDays(parseDay(ev.end), 1));
      const [ny, nm, nd] = next.split("-").map(Number);
      eUtc = wallToUtc(ny, nm, nd, 11, 0, tzE);
    } else {
      eUtc = wallToUtc(ey, em, ed, 11, 0, tzE);
    }
    return { sUtc, eUtc };
  }
  let sUtc, eUtc;
  if (ev.startTime) {
    const [h, mn] = ev.startTime.split(":").map(Number);
    sUtc = wallToUtc(sy, sm, sd, h, mn, tzS);
  } else {
    sUtc = wallToUtc(sy, sm, sd, 0, 0, tzS);
  }
  if (ev.endTime) {
    const [h, mn] = ev.endTime.split(":").map(Number);
    eUtc = wallToUtc(ey, em, ed, h, mn, tzE);
  } else {
    const next = toISO(addDays(parseDay(ev.end), 1));
    const [eyN, emN, edN] = next.split("-").map(Number);
    eUtc = wallToUtc(eyN, emN, edN, 0, 0, tzE);
  }
  return { sUtc, eUtc };
}

// --- pack into sub-rows so overlapping events stack ---

function packRowsByUtc(items) {
  items.sort((a, b) => a.sUtc - b.sUtc);
  const rows = [];
  return items.map(it => {
    let row = 0;
    while (row < rows.length && rows[row] > it.sUtc + 1) row++;
    rows[row] = it.eUtc;
    return { ...it, row };
  });
}

// --- DOM helpers ---

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}

function makeTitle(ev) {
  const startD = fmtShort(parseDay(ev.start));
  const endD = fmtShort(parseDay(ev.end));
  const range = startD === endD ? startD : `${startD} – ${endD}`;
  let s = `${ev.title}\n${range}`;
  if (ev.startTime || ev.endTime) {
    const stz = ev.startTz ? ` ${tzShortName(ev.startTz, ev.start)}` : "";
    const etz = ev.endTz   ? ` ${tzShortName(ev.endTz, ev.end)}`     : "";
    s += `\n${ev.startTime || ""}${stz} → ${ev.endTime || ""}${etz}`;
  }
  if (ev.notes) s += `\n\n${ev.notes}`;
  return s;
}

// --- Map a UTC time to a fractional day position in our timeline ---

function toggleShrinkDay(ds) {
  if (!state.shrunkDays) state.shrunkDays = [];
  const i = state.shrunkDays.indexOf(ds);
  if (i >= 0) state.shrunkDays.splice(i, 1);
  else state.shrunkDays.push(ds);
  saveLocal();   // collapsed-day choice is per-device UI
  renderApp();
}

function fracToPx(frac, dayWidths, dayOffsets) {
  if (frac <= 0) return 0;
  const idx = Math.min(Math.floor(frac), dayWidths.length - 1);
  const within = Math.max(0, Math.min(1, frac - idx));
  return dayOffsets[idx] + dayWidths[idx] * within;
}

function utcToFrac(utc, dayUtcBounds) {
  for (let i = 0; i < dayUtcBounds.length; i++) {
    const d = dayUtcBounds[i];
    if (utc <= d.endUtc) {
      const dur = d.endUtc - d.startUtc;
      const within = Math.max(0, Math.min(1, (utc - d.startUtc) / dur));
      return i + within;
    }
  }
  return dayUtcBounds.length;
}

// Position an event on the timeline using its declared local clock times
// within each day cell — bypasses UTC/TZ math so flights show up at the
// times shown on their boarding pass regardless of TZ shifts mid-trip.
function eventLocalFracs(ev, rangeStart) {
  const startDayIdx = dayDiff(rangeStart, ev.start);
  const endDayIdx = dayDiff(rangeStart, ev.end);
  function timeFrac(t) {
    if (!t) return 0;
    const [h, m] = t.split(":").map(Number);
    return (h + (m || 0) / 60) / 24;
  }
  const leftFrac = startDayIdx + (ev.startTime ? timeFrac(ev.startTime) : 0);
  const rightFrac = ev.endTime ? endDayIdx + timeFrac(ev.endTime) : endDayIdx + 1;
  return { leftFrac, rightFrac };
}

// --- Render a timeline grid (axis + lanes) for a date range ---

// Resolve an event id to the root of its merge chain (so picking a child as a
// merge target redirects to the parent it's already under).
function resolveMergeRoot(id, eventsList) {
  const byId = new Map(eventsList.map(e => [e.id, e]));
  let cur = byId.get(id);
  const seen = new Set();
  while (cur && cur.mergedInto && byId.has(cur.mergedInto) && !seen.has(cur.id)) {
    seen.add(cur.id);
    cur = byId.get(cur.mergedInto);
  }
  return cur ? cur.id : id;
}

// Given a flat event list, hide merged children and expand each parent's date
// range to the union of itself + all its children. The original events stay in
// state.events untouched — this is render-time only.
function mergeAwareEvents(events) {
  const byId = new Map(events.map(e => [e.id, e]));
  const childrenOf = new Map();
  for (const ev of events) {
    if (ev.mergedInto && byId.has(ev.mergedInto)) {
      const root = resolveMergeRoot(ev.id, events);
      if (root === ev.id) continue;
      if (!childrenOf.has(root)) childrenOf.set(root, []);
      childrenOf.get(root).push(ev);
    }
  }
  return events
    .filter(ev => !(ev.mergedInto && byId.has(ev.mergedInto)))
    .map(ev => {
      const kids = childrenOf.get(ev.id);
      if (!kids?.length) return ev;
      let start = ev.start, end = ev.end;
      for (const k of kids) {
        if (k.start && k.start < start) start = k.start;
        if (k.end && k.end > end) end = k.end;
      }
      return { ...ev, start, end, _mergeCount: kids.length };
    });
}

function renderTimeline(container, rangeStart, rangeEnd, opts) {
  const { dayTzMap, homeTz, dayPx, compact, tzAware } = opts;
  const events = mergeAwareEvents(opts.events || state.events);
  // Axis labels always show destination TZ when it differs from home, so the
  // user sees "where they'll be" even when 'Follow timezone' is off (which
  // only affects event positioning, not labelling).
  const displayTzMap = computeDayTzMap(rangeStart, rangeEnd, events, homeTz, true);
  const totalDays = dayDiff(rangeStart, rangeEnd) + 1;

  // Lane area's pixel width = container - lane label column. Used to compute
  // bar widths in pixels for stretch mode so very short events don't collapse
  // to flex/intrinsic size when the calc result would be ~0.
  const LANE_LABEL_PX = 110;
  const stretchLanePx = Math.max(200, (container.clientWidth || 600) - LANE_LABEL_PX);

  const dayUtcBounds = [];
  for (let i = 0; i < totalDays; i++) {
    const ds = toISO(addDays(parseDay(rangeStart), i));
    const tz = dayTzMap[ds] || homeTz;
    const startUtc = dayStartUtcInTz(ds, tz);
    const nextDs = toISO(addDays(parseDay(ds), 1));
    const nextTz = dayTzMap[nextDs] || tz;
    const endUtc = dayStartUtcInTz(nextDs, nextTz);
    dayUtcBounds.push({ ds, tz, startUtc, endUtc });
  }

  const visible = events.filter(ev => {
    // 1-night lodging entries (start === end, no times) visually extend into
    // the next morning for check-out — make sure they're still included in a
    // breakdown segment that starts the morning after.
    let effEnd = ev.end;
    if (ev.lane === "lodging" && !ev.startTime && !ev.endTime && ev.start === ev.end) {
      effEnd = toISO(addDays(parseDay(ev.end), 1));
    }
    return !(effEnd < rangeStart || ev.start > rangeEnd);
  });

  // Compute density: a day is "active" if any non-location event touches it.
  // Location bars (long destination stays) don't count toward density.
  const dayActive = new Array(totalDays).fill(false);
  for (const ev of visible) {
    if (ev.lane === "location") continue;
    for (let i = 0; i < totalDays; i++) {
      const ds = dayUtcBounds[i].ds;
      if (ev.start <= ds && ev.end >= ds) dayActive[i] = true;
    }
  }

  // Per-day pixel widths (only when in fixed-day mode).
  // Default: every day full-width. User can click to "shrink" a day; the
  // freed space gets redistributed across the remaining full-width days
  // so the timeline still fills the panel.
  const SHRUNK_PX = 22;
  const shrunkSet = new Set(state.shrunkDays || []);
  let dayWidths = null;
  if (dayPx) {
    const shrunkCount = dayUtcBounds.filter(b => shrunkSet.has(b.ds)).length;
    const nonShrunkCount = totalDays - shrunkCount;
    const totalAvailable = totalDays * dayPx;
    const fullPx = nonShrunkCount > 0
      ? Math.max(60, Math.floor((totalAvailable - shrunkCount * SHRUNK_PX) / nonShrunkCount))
      : dayPx;
    dayWidths = dayUtcBounds.map(b => shrunkSet.has(b.ds) ? SHRUNK_PX : fullPx);
  }
  const dayOffsetsPx = [];
  let cumPx = 0;
  if (dayWidths) {
    for (const w of dayWidths) { dayOffsetsPx.push(cumPx); cumPx += w; }
  }
  const totalPx = cumPx;

  const grid = el("div", "timeline-grid" + (dayPx ? " fixed-day-width" : ""));
  if (dayPx) grid.style.setProperty("--day-px", `${dayPx}px`);

  // Row 1: lane spacer + axis
  grid.appendChild(el("div", "lane-spacer"));

  const axis = el("div", "axis" + (compact ? " compact" : ""));
  const today = toISO(new Date());
  for (let i = 0; i < totalDays; i++) {
    const { ds, tz } = dayUtcBounds[i];
    const day = parseDay(ds);
    const cell = el("div", "day");
    if (ds === today) cell.classList.add("today");
    if (dayWidths) {
      cell.style.flex = `0 0 ${dayWidths[i]}px`;
      cell.style.cursor = "pointer";
      const isShrunk = shrunkSet.has(ds);
      if (isShrunk) {
        cell.classList.add("compact-day");
        cell.title = `${DOW[day.getDay()]} ${day.toLocaleDateString(undefined, { month: "short", day: "numeric" })} — click to expand`;
      } else {
        cell.title = "Click to shrink this day";
      }
      cell.addEventListener("click", () => toggleShrinkDay(ds));
    }
    cell.appendChild(el("span", "dow", DOW[day.getDay()]));
    cell.appendChild(el("span", "num", compact
      ? String(day.getDate())
      : `${day.toLocaleDateString(undefined, { month: "short" })} ${day.getDate()}`));
    // When Follow timezone is on, show every day's destination tz on the
    // axis — including days you're at home (so you can see "PST" until you
    // fly out and "EAT" after). When off, the whole trip is in home tz so
    // the label adds no information; omit it.
    if (tzAware) {
      const displayTz = displayTzMap[ds] || tz || homeTz;
      cell.appendChild(el("span", "tz", tzShortName(displayTz, ds)));
    }
    axis.appendChild(cell);
  }
  grid.appendChild(axis);

  // Lanes
  for (const lane of LANES) {
    const laneEvents = visible.filter(ev => (ev.lane || "activities") === lane.key);
    // Optional lanes (rental car, etc.) are hidden entirely when there are
    // no events for them, so trips without one don't show an empty row.
    if (lane.optional && laneEvents.length === 0) continue;
    grid.appendChild(el("div", "lane-label", lane.label));

    const laneArea = el("div", "lane-events");

    if (laneEvents.length === 0) {
      laneArea.classList.add("empty");
      grid.appendChild(laneArea);
      continue;
    }

    const items = laneEvents.map(ev => {
      // displayTzMap always holds the *real* destination tz per day (independent
      // of state.tzAware) — pass it so event times convert from their true tz
      // to UTC even in PST-axis mode, instead of being reinterpreted as PST.
      const { sUtc, eUtc } = eventBounds(ev, displayTzMap, homeTz, tzAware);
      // Compute the *visual* fracs upfront so the packer can use them. UTC
      // bounds aren't enough when an event's lodging-packing override carves
      // it down to a tiny morning sliver in this segment — without using the
      // visual fracs, the packer would treat the full UTC range as occupied
      // and bump a non-overlapping later event onto a second row.
      let leftFrac, rightFrac;
      if (tzAware) {
        ({ leftFrac, rightFrac } = eventLocalFracs(ev, rangeStart));
      } else {
        leftFrac = utcToFrac(sUtc, dayUtcBounds);
        rightFrac = utcToFrac(eUtc, dayUtcBounds);
      }
      if (ev.lane === "lodging" && !ev.startTime && !ev.endTime) {
        const startDayIdx = dayDiff(rangeStart, ev.start);
        const endDayIdx = dayDiff(rangeStart, ev.end);
        leftFrac = startDayIdx + 15 / 24;
        if (ev.start === ev.end) {
          rightFrac = startDayIdx + 1 + 11 / 24;
        } else {
          rightFrac = endDayIdx + 11 / 24;
        }
      }
      // Clip to segment so the packer compares only what's actually drawn.
      const clippedLeft = Math.max(0, leftFrac);
      const clippedRight = Math.min(totalDays, rightFrac);
      return { ev, sUtc, eUtc, leftFrac, rightFrac, clippedLeft, clippedRight };
    });
    // Pack by visual fracs — events whose clipped visual ranges don't overlap
    // share a row, regardless of their full UTC range.
    items.sort((a, b) => a.clippedLeft - b.clippedLeft);
    const rowEnds = [];
    const packed = items.map(it => {
      let row = 0;
      while (row < rowEnds.length && rowEnds[row] > it.clippedLeft + 0.0001) row++;
      rowEnds[row] = it.clippedRight;
      return { ...it, row };
    });
    const rowCount = Math.max(...packed.map(p => p.row + 1));
    laneArea.style.setProperty("--rows", rowCount);

    // Night shading (8pm–8am) per day — sits behind the event bars.
    appendNightShades(laneArea, totalDays, dayWidths, dayOffsetsPx, shrunkSet, dayUtcBounds);

    for (const { ev, sUtc, eUtc, row, leftFrac: lf0, rightFrac: rf0 } of packed) {
      // Visual fracs were already computed (with the lodging override) before
      // packing — reuse them so packing and rendering can't disagree.
      let leftFrac = Math.max(0, lf0);
      let rightFrac = Math.min(totalDays, rf0);
      if (rightFrac <= leftFrac) continue;

      const colorVal = ev.color || "indigo";
      const isHex = colorVal.startsWith("#");
      const bar = el("div", `event ${isHex ? "" : colorVal}` +
        (ev._isOption ? " is-option" : "") +
        (ev.tentative ? " tentative" : ""));
      if (isHex) {
        bar.style.background = colorVal;
        bar.style.borderColor = colorVal;
      }
      let widthPx;
      if (dayWidths) {
        const leftPx = fracToPx(leftFrac, dayWidths, dayOffsetsPx);
        const rightPx = fracToPx(rightFrac, dayWidths, dayOffsetsPx);
        widthPx = Math.max(2, rightPx - leftPx - 2);
        bar.style.left = `${leftPx + 1}px`;
        bar.style.width = `${widthPx}px`;
      } else {
        const leftPx = (leftFrac / totalDays) * stretchLanePx;
        widthPx = Math.max(2, ((rightFrac - leftFrac) / totalDays) * stretchLanePx - 2);
        bar.style.left = `${leftPx + 1}px`;
        bar.style.width = `${widthPx}px`;
      }
      // Drop text + padding when the bar is too narrow — keeps the bar's
      // pixel width honest so back-to-back short events (a 1-hour flight
      // followed by a layover) don't overlap. Applies in both stretch and
      // fixed-day-width modes.
      if (widthPx < 28) {
        bar.style.padding = "0";
        bar.dataset.narrow = "1";
      }
      bar.style.top = `calc(${row} * (var(--row-h) + 4px) + 4px)`;
      bar.textContent = bar.dataset.narrow ? "" : ev.title;
      bar.title = makeTitle(ev);
      bar.addEventListener("click", () => openEventDialog(ev.id, ev._optionId || null));
      laneArea.appendChild(bar);
    }

    grid.appendChild(laneArea);
  }

  container.innerHTML = "";
  container.appendChild(grid);
}

// Render translucent grey strips for the night portion of each day
// (00:00–08:00 and 20:00–24:00) behind the event bars in a lane area.
function appendNightShades(laneArea, totalDays, dayWidths, dayOffsetsPx, shrunkSet, dayUtcBounds) {
  const layer = el("div", "night-shades");
  for (let i = 0; i < totalDays; i++) {
    if (shrunkSet && shrunkSet.has(dayUtcBounds[i].ds)) continue;
    const dayLeftFrac = i;
    const morningEndFrac = i + 8 / 24;
    const eveningStartFrac = i + 20 / 24;
    const dayRightFrac = i + 1;

    function place(el2, leftFrac, rightFrac) {
      if (dayWidths) {
        const leftPx = fracToPx(leftFrac, dayWidths, dayOffsetsPx);
        const rightPx = fracToPx(rightFrac, dayWidths, dayOffsetsPx);
        el2.style.left = `${leftPx}px`;
        el2.style.width = `${Math.max(0, rightPx - leftPx)}px`;
      } else {
        el2.style.left = `${(leftFrac / totalDays) * 100}%`;
        el2.style.width = `${((rightFrac - leftFrac) / totalDays) * 100}%`;
      }
    }
    const morning = el("div", "night-shade");
    place(morning, dayLeftFrac, morningEndFrac);
    const evening = el("div", "night-shade");
    place(evening, eveningStartFrac, dayRightFrac);
    layer.appendChild(morning);
    layer.appendChild(evening);
  }
  laneArea.appendChild(layer);
}

// --- Pricing tab ---
//
// state.lineItems: [{ id, eventIds: [...], cost: number, label?: string, tentative: bool }]
// Each event can belong to at most one line item. Pricing tab lets you click
// event pills to bundle them into a single-cost line item.

const pricingSelection = new Set();
const LANE_LABEL = { location: "Where", lodging: "Lodging", flights: "Flights", rental: "Transportation", activities: "Activities" };
const LANE_ORDER = ["flights", "lodging", "rental", "activities", "location"];

function fmtMoney(n) {
  if (typeof n !== "number" || isNaN(n)) return "$0";
  const sign = n < 0 ? "-" : "";
  // Round to 2 decimals, then drop trailing zeros (and the dot if whole).
  const abs = Math.round(Math.abs(n) * 100) / 100;
  const [intPart, decPart] = abs.toFixed(2).split(".");
  const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const trimmedDec = decPart.replace(/0+$/, "");
  return sign + "$" + withCommas + (trimmedDec ? "." + trimmedDec : "");
}

function ensureLineItems() {
  if (!Array.isArray(state.lineItems)) state.lineItems = [];
}

function eventToLineItem(eventId) {
  ensureLineItems();
  // While editing a line item, that line item's events are treated as
  // unbundled so they can be re-selected/deselected freely.
  return state.lineItems.find(li =>
    li.id !== editingLineItemId && li.eventIds.includes(eventId));
}

function ensurePricingHasOthers() {
  if (typeof state.pricingHasOthers !== "boolean") {
    // Heuristic for older trips: if any line item has per-group overrides
    // or the priceSplit has more than one group, assume others are involved.
    const hasOverrides = (state.lineItems || []).some(li =>
      (li.pricing?.parties || []).some(p => p.name && p.name !== "Mine"));
    const multiGroup = state.priceSplit?.groups?.length > 1;
    state.pricingHasOthers = !!(hasOverrides || multiGroup);
  }
}

function applyPricingHasOthersUI() {
  const tab = document.getElementById("tab-pricing");
  if (!tab) return;
  tab.classList.toggle("solo", !state.pricingHasOthers);
  const cb = document.getElementById("pricing-has-others-cb");
  if (cb) cb.checked = !!state.pricingHasOthers;
}

function renderPricing() {
  ensureLineItems();
  ensurePriceSplit();
  ensurePricingHasOthers();
  applyPricingHasOthersUI();
  // Auto-prune deleted events from every line item — stops "(deleted)"
  // labels lingering after an event was removed and re-added (re-add gets
  // a new id, so the old id stays orphaned in the bundle).
  const validIds = new Set(state.events.map(e => e.id));
  for (const li of state.lineItems) {
    li.eventIds = li.eventIds.filter(eid => validIds.has(eid));
  }
  renderSplitEditor();
  renderPricingPills();
  renderPricingLineItems();
  renderPricingSummary();
  renderSettleUp();
  renderLineItemForm();
}

// Settle-up: for each other group, net = (their fair share of all items) minus
// (what they've already paid). Positive → they owe you; negative → you owe them.
// A per-person checkbox marks them paid-up; clicking the pill opens a breakdown.
function renderSettleUp() {
  const container = document.getElementById("pricing-settle");
  if (!container) return;
  container.innerHTML = "";
  ensurePriceSplit();
  if (!state.pricingHasOthers || state.priceSplit.groups.length < 2) return;
  if (!state.priceSettled) state.priceSettled = {};

  const rows = [];
  state.priceSplit.groups.slice(1).forEach((g, i) => {
    const idx = i + 1;
    let share = 0, paid = 0;
    for (const li of state.lineItems) {
      share += lineItemGroupAmount(li, g.id, idx);
      paid += linePaidBy(li, g.id);
    }
    const net = Math.round((share - paid) * 100) / 100;
    const settled = !!state.priceSettled[g.id];
    if (Math.abs(net) < 0.005 && paid === 0 && !settled) return; // nothing to settle
    rows.push({ g, idx, net, settled });
  });
  if (!rows.length) return;

  const wrap = el("div", "settle-row");
  wrap.appendChild(el("span", "settle-label", "Settle up:"));
  for (const { g, idx, net, settled } of rows) {
    const item = el("span", "settle-item");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "settle-check";
    cb.checked = settled;
    cb.title = `Mark ${g.name} as paid up`;
    cb.addEventListener("change", () => {
      if (cb.checked) state.priceSettled[g.id] = true;
      else delete state.priceSettled[g.id];
      save();
      renderPricing();
    });
    item.appendChild(cb);

    const pill = el("span", "settle-pill " + (settled ? "settle-done" : net < -0.005 ? "settle-neg" : "settle-pos"));
    if (settled) pill.textContent = `${g.name}: paid up ✓`;
    else if (net > 0.005) pill.textContent = `${g.name} owes you ${fmtMoney(net)}`;
    else if (net < -0.005) pill.textContent = `You owe ${g.name} ${fmtMoney(-net)}`;
    else pill.textContent = `${g.name}: settled`;
    pill.title = "Click for a breakdown";
    pill.addEventListener("click", () => showSettleBreakdown(idx));
    item.appendChild(pill);

    wrap.appendChild(item);
  }
  container.appendChild(wrap);
}

// Modal: itemized breakdown for one group — each line item's total, their fair
// share, and what they paid, so you can see exactly how the balance is made up.
function showSettleBreakdown(groupIdx) {
  const g = state.priceSplit.groups[groupIdx];
  if (!g) return;
  document.getElementById("settle-dialog")?.remove();
  const dlg = document.createElement("dialog");
  dlg.id = "settle-dialog";
  dlg.className = "settle-dialog";

  const head = el("div", "settle-dialog-head");
  head.appendChild(el("h3", null, `Settle up · ${g.name}`));
  const closeBtn = el("button", "settle-close", "✕");
  closeBtn.type = "button";
  head.appendChild(closeBtn);
  dlg.appendChild(head);

  const table = el("table", "settle-table");
  const thead = el("thead");
  const htr = el("tr");
  htr.appendChild(el("th", null, "Item"));
  htr.appendChild(el("th", "num", "Total"));
  htr.appendChild(el("th", "num", `${g.name} share`));
  htr.appendChild(el("th", "num", `${g.name} paid`));
  thead.appendChild(htr);
  table.appendChild(thead);

  const tbody = el("tbody");
  let totShare = 0, totPaid = 0, totTotal = 0, anyRow = false;
  for (const li of state.lineItems) {
    const share = lineItemGroupAmount(li, g.id, groupIdx);
    const paid = linePaidBy(li, g.id);
    if (!share && !paid) continue;
    anyRow = true;
    const total = lineItemTotal(li);
    totShare += share; totPaid += paid; totTotal += total;
    const titles = li.eventIds.map(id => state.events.find(e => e.id === id)?.title).filter(Boolean);
    const tr = el("tr");
    tr.appendChild(el("td", null, li.label || titles.join(" + ") || "(item)"));
    tr.appendChild(el("td", "num", fmtMoney(total)));
    tr.appendChild(el("td", "num", fmtMoney(share)));
    tr.appendChild(el("td", "num " + (paid ? "settle-paid-cell" : ""), paid ? fmtMoney(paid) : "—"));
    tbody.appendChild(tr);
  }
  if (!anyRow) {
    const tr = el("tr"); const td = el("td", null, "No shared items yet."); td.colSpan = 4; tr.appendChild(td); tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  const tfoot = el("tfoot");
  const ftr = el("tr");
  ftr.appendChild(el("td", null, "Total"));
  ftr.appendChild(el("td", "num", fmtMoney(totTotal)));
  ftr.appendChild(el("td", "num", fmtMoney(totShare)));
  ftr.appendChild(el("td", "num", fmtMoney(totPaid)));
  tfoot.appendChild(ftr);
  table.appendChild(tfoot);
  dlg.appendChild(table);

  const net = Math.round((totShare - totPaid) * 100) / 100;
  dlg.appendChild(el("div", "settle-verdict " + (net < -0.005 ? "settle-neg" : "settle-pos"),
    Math.abs(net) < 0.005 ? `Settled up with ${g.name}.`
      : net > 0 ? `${g.name} owes you ${fmtMoney(net)}`
      : `You owe ${g.name} ${fmtMoney(-net)}`));
  dlg.appendChild(el("div", "settle-note",
    `The "${g.name} paid" column is what they fronted; their share of everything else is what you covered.`));

  document.body.appendChild(dlg);
  closeBtn.addEventListener("click", () => dlg.close());
  dlg.addEventListener("click", e => { if (e.target === dlg) dlg.close(); });
  dlg.addEventListener("close", () => dlg.remove());
  dlg.showModal();
}

function lineItemEarliestDate(li) {
  let best = null;
  for (const id of li.eventIds) {
    const ev = state.events.find(e => e.id === id);
    if (ev && ev.start && (!best || ev.start < best)) best = ev.start;
  }
  return best;
}

function renderPricingPills() {
  const container = document.getElementById("pricing-pills");
  if (!container) return;
  container.innerHTML = "";
  const events = [...state.events].sort((a, b) => (a.start || "").localeCompare(b.start || ""));
  if (events.length === 0) {
    container.innerHTML = `<div class="empty-state" style="padding:0">No events yet — add some on the My itinerary tab.</div>`;
    return;
  }
  for (const ev of events) {
    const inLI = eventToLineItem(ev.id);
    const pill = document.createElement("span");
    pill.className = "pricing-pill";
    if (inLI) pill.classList.add("in-line-item");
    if (pricingSelection.has(ev.id)) pill.classList.add("selected");
    const colorVal = ev.color || "indigo";
    const dot = document.createElement("span");
    dot.className = "pill-dot";
    if (colorVal.startsWith("#")) dot.style.background = colorVal;
    else dot.style.backgroundColor = `var(--${colorVal})`;
    pill.appendChild(dot);
    pill.appendChild(document.createTextNode(ev.title));
    if (!inLI) {
      pill.addEventListener("click", () => {
        if (pricingSelection.has(ev.id)) pricingSelection.delete(ev.id);
        else pricingSelection.add(ev.id);
        renderPricingPills();
        updatePricingSelectionSummary();
      });
    } else {
      pill.title = `Already in: ${inLI.label || "(unlabeled)"} — ${fmtMoney(inLI.cost)}`;
    }
    container.appendChild(pill);
  }
  updatePricingSelectionSummary();
}

function updatePricingSelectionSummary() {
  const sum = document.getElementById("pricing-selection-summary");
  if (!sum) return;
  sum.textContent = pricingSelection.size === 0 ? "" : `${pricingSelection.size} selected`;
}

function renderPricingLineItems() {
  const container = document.getElementById("pricing-line-items");
  if (!container) return;
  container.innerHTML = "";
  ensureLineItems();
  // Group line items by the dominant lane of their events.
  const byLane = {};
  for (const li of state.lineItems) {
    let lane = li.lane;
    if (!lane) {
      const lanes = li.eventIds.map(id => state.events.find(e => e.id === id)?.lane).filter(Boolean);
      const counts = {};
      lanes.forEach(l => { counts[l] = (counts[l] || 0) + 1; });
      lane = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0] || "activities";
    }
    (byLane[lane] ||= []).push(li);
  }
  for (const lane of LANE_ORDER) {
    const items = byLane[lane];
    if (!items || items.length === 0) continue;
    // <details> so the user can click the lane header to collapse the rows.
    const group = document.createElement("details");
    group.className = "pricing-group";
    group.open = true;
    const subtotal = items.reduce((s, li) => s + lineItemTotal(li), 0);
    // Per-group subtotal for this lane (matches the right-side pills layout).
    const lanePerGroup = state.priceSplit.groups.map((g, gi) =>
      items.reduce((s, li) => s + lineItemGroupAmount(li, g.id, gi), 0));
    const h = document.createElement("summary");
    h.className = "pricing-group-head";
    const labelSpan = document.createElement("span");
    labelSpan.className = "pricing-group-label";
    labelSpan.textContent = LANE_LABEL[lane] || lane;
    h.appendChild(labelSpan);
    // Right-side area uses the SAME grid template as line items so pills
    // + total + (no actions) line up vertically across header and rows.
    const right = document.createElement("span");
    right.className = "lane-subtotal-row li-right";
    right.style.setProperty("--group-count", state.pricingHasOthers ? state.priceSplit.groups.length : 0);
    if (state.pricingHasOthers) {
      state.priceSplit.groups.forEach((g, gi) => {
        const amt = lanePerGroup[gi];
        const pill = document.createElement("span");
        pill.className = `group-pill outline-${g.color || "indigo"}`;
        pill.style.gridColumn = gi + 1;
        if (amt) pill.textContent = `${g.name}: ${fmtMoney(amt)}`;
        else pill.classList.add("empty");
        right.appendChild(pill);
      });
    }
    const subTotGc = (state.pricingHasOthers ? state.priceSplit.groups.length : 0) + 1;
    const subTot = document.createElement("span");
    subTot.className = "group-total li-cost";
    subTot.style.gridColumn = subTotGc;
    subTot.textContent = fmtMoney(subtotal);
    right.appendChild(subTot);
    // Empty placeholder for the actions column so totals line up with rows.
    const actionsSpacer = document.createElement("span");
    actionsSpacer.className = "li-actions actions-spacer";
    actionsSpacer.style.gridColumn = subTotGc + 1;
    right.appendChild(actionsSpacer);
    h.appendChild(right);
    group.appendChild(h);
    // Sort items chronologically by their earliest event start date.
    items.sort((a, b) => {
      const da = lineItemEarliestDate(a) || "9999-12-31";
      const db = lineItemEarliestDate(b) || "9999-12-31";
      return da.localeCompare(db);
    });
    const ul = document.createElement("ul");
    for (const li of items) {
      const liEl = document.createElement("li");
      const main = document.createElement("div");
      main.className = "li-main";
      const labelRow = document.createElement("div");
      labelRow.className = "li-label-row";
      const label = document.createElement("span");
      label.className = "li-label";
      const eventTitles = li.eventIds.map(id => state.events.find(e => e.id === id)?.title).filter(Boolean);
      label.textContent = li.label || eventTitles.join(" + ") || "(empty bundle)";
      labelRow.appendChild(label);
      const date = lineItemEarliestDate(li);
      if (date) {
        const dateEl = document.createElement("span");
        dateEl.className = "li-date";
        dateEl.textContent = date;
        labelRow.appendChild(dateEl);
      }
      main.appendChild(labelRow);
      const evList = document.createElement("div");
      evList.className = "li-events";
      if (li.label) evList.textContent = eventTitles.join(" · ");
      if (li.label) main.appendChild(evList);
      // Per-group amount pills + total cost laid out in a grid so pills align
      // vertically across line items (one column per default-split group plus
      // a column for the total).
      ensurePriceSplit();
      const right = document.createElement("div");
      right.className = "li-right";
      const liGc = state.pricingHasOthers ? state.priceSplit.groups.length : 0;
      right.style.setProperty("--group-count", liGc);
      if (state.pricingHasOthers) {
        state.priceSplit.groups.forEach((g, idx) => {
          const amt = lineItemGroupAmount(li, g.id, idx);
          const isOverridden = li.overrides && li.overrides[g.id] != null;
          const pill = document.createElement("span");
          pill.className = `group-pill outline-${g.color || "indigo"}` + (isOverridden ? " overridden" : "");
          pill.style.gridColumn = idx + 1;
          if (amt) {
            pill.textContent = `${g.name}: ${fmtMoney(amt)}`;
            if (isOverridden) pill.title = "Overridden (default would be " +
              fmtMoney((li.total || 0) * (g.share || 0)) + ")";
          } else {
            pill.classList.add("empty");
          }
          right.appendChild(pill);
        });
      }
      const cost = document.createElement("span");
      cost.className = "li-cost";
      cost.style.gridColumn = liGc + 1;
      cost.textContent = fmtMoney(lineItemTotal(li));
      right.appendChild(cost);
      // Actions also live in the grid so they align across rows.
      const actionsCol = document.createElement("span");
      actionsCol.className = "li-actions";
      actionsCol.style.gridColumn = liGc + 2;
      const editBtn2 = document.createElement("button");
      editBtn2.textContent = "Edit";
      editBtn2.addEventListener("click", () => editLineItem(li.id));
      const delBtn2 = document.createElement("button");
      delBtn2.textContent = "×";
      delBtn2.className = "danger";
      delBtn2.title = "Delete line item";
      delBtn2.addEventListener("click", () => {
        state.lineItems = state.lineItems.filter(x => x.id !== li.id);
        save();
        renderPricing();
      });
      actionsCol.appendChild(editBtn2);
      actionsCol.appendChild(delBtn2);
      right.appendChild(actionsCol);
      liEl.appendChild(main);
      liEl.appendChild(right);
      ul.appendChild(liEl);
    }
    group.appendChild(ul);
    container.appendChild(group);
  }
  if (state.lineItems.length === 0) {
    container.innerHTML = `<div class="empty-state">No line items yet — select pills above and click + Add line item.</div>`;
  }
}

function renderPricingSummary() {
  const container = document.getElementById("pricing-summary");
  if (!container) return;
  ensureLineItems();
  ensurePriceSplit();
  let booked = 0, tentativeTotal = 0;
  const perGroup = state.priceSplit.groups.map(() => 0);
  for (const li of state.lineItems) {
    const total = lineItemTotal(li);
    const isTent = li.eventIds.some(id => state.events.find(e => e.id === id)?.tentative);
    if (isTent) tentativeTotal += total;
    else booked += total;
    state.priceSplit.groups.forEach((g, idx) => {
      perGroup[idx] += lineItemGroupAmount(li, g.id, idx);
    });
  }
  container.innerHTML = "";
  // Single right-aligned row: per-group pills + grand total inline together.
  const row = document.createElement("div");
  row.className = "pricing-summary-row";
  if (state.pricingHasOthers) {
    state.priceSplit.groups.forEach((g, idx) => {
      if (!perGroup[idx]) return;
      const pill = document.createElement("span");
      pill.className = `group-pill summary-pill outline-${g.color || "indigo"}`;
      pill.textContent = `${g.name}: ${fmtMoney(perGroup[idx])}`;
      row.appendChild(pill);
    });
  }
  const total = document.createElement("span");
  total.className = "pricing-summary-total";
  total.textContent = `Total: ${fmtMoney(booked + tentativeTotal)}`;
  row.appendChild(total);
  container.appendChild(row);
}

// ID of the line item currently being edited (null = adding new).
let editingLineItemId = null;

function editLineItem(id) {
  const li = state.lineItems.find(x => x.id === id);
  if (!li) return;
  ensurePriceSplit();
  editingLineItemId = id;
  // Prune deleted events from the bundle while we're here.
  const validIds = new Set(state.events.map(e => e.id));
  li.eventIds = li.eventIds.filter(eid => validIds.has(eid));
  pricingSelection.clear();
  li.eventIds.forEach(eid => pricingSelection.add(eid));
  document.getElementById("pricing-label").value = li.label || "";
  const laneSel = document.getElementById("pricing-lane");
  if (laneSel) laneSel.value = li.lane || "";
  // Load total — convert from legacy shapes if needed.
  formTotalInput = String(lineItemTotal(li) || "");
  // Load overrides.
  for (const k of Object.keys(formOverrides)) delete formOverrides[k];
  if (li.overrides) {
    for (const [k, v] of Object.entries(li.overrides)) {
      if (v && typeof v === "object" && "pct" in v) formOverrides[k] = `${v.pct}%`;
      else formOverrides[k] = String(v);
    }
  }
  for (const k of Object.keys(formPaid)) delete formPaid[k];
  if (li.paid) {
    for (const [k, v] of Object.entries(li.paid)) formPaid[k] = String(v);
  }
  const addBtn = document.getElementById("pricing-add-btn");
  if (addBtn) addBtn.textContent = "Save changes";
  const heading = document.getElementById("pricing-add-heading");
  if (heading) heading.textContent = "Edit line item";
  const cancelBtn = document.getElementById("pricing-cancel-edit");
  if (cancelBtn) cancelBtn.hidden = false;
  // Show the bundle pills only if this item actually bundles itinerary items.
  const bundleDetails = document.getElementById("pricing-bundle-details");
  if (bundleDetails) bundleDetails.open = li.eventIds.length > 0;
  renderPricing();
  document.querySelector(".pricing-builder")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function cancelLineItemEdit() {
  editingLineItemId = null;
  pricingSelection.clear();
  document.getElementById("pricing-label").value = "";
  formTotalInput = "";
  for (const k of Object.keys(formOverrides)) delete formOverrides[k];
  for (const k of Object.keys(formPaid)) delete formPaid[k];
  const addBtn = document.getElementById("pricing-add-btn");
  if (addBtn) addBtn.textContent = "+ Add line item";
  const heading = document.getElementById("pricing-add-heading");
  if (heading) heading.textContent = "Add line item";
  const cancelBtn = document.getElementById("pricing-cancel-edit");
  if (cancelBtn) cancelBtn.hidden = true;
  renderPricing();
}

// Parse a math-y amount string. Accepts:
//   "250"          → 250
//   "1/4*1000"     → 250
//   "25%*800"      → 200
//   "2*150"        → 300
//   ""             → 0
// Returns NaN for unsafe/invalid input.
function parseAmount(str) {
  if (str == null || str === "") return 0;
  const s = String(str).trim();
  if (!s) return 0;
  // Whitelist characters to avoid arbitrary code execution via Function.
  if (!/^[\d.+\-*/() %]+$/.test(s)) return NaN;
  // Replace % with /100 (after a number, this means percent).
  const normalized = s.replace(/%/g, "/100");
  try {
    const v = Function(`"use strict"; return (${normalized});`)();
    return typeof v === "number" && isFinite(v) ? v : NaN;
  } catch (e) { return NaN; }
}

// --- price split (trip-level default) ---
// Compute each group's normalized share from its raw shareInput. If any
// value exceeds 1 it's treated as a quantity and the whole set normalizes
// against the sum, so the user can type "2, 1, 1" and get 50%/25%/25%.
// Pure fractions/percents (≤1) are used as-is.
function recomputeShares() {
  if (!state.priceSplit || !state.priceSplit.groups) return;
  const parsed = state.priceSplit.groups.map(g => {
    const v = parseAmount(g.shareInput);
    return isNaN(v) ? 0 : v;
  });
  const anyOver1 = parsed.some(v => v > 1);
  if (anyOver1) {
    const sum = parsed.reduce((a, b) => a + b, 0);
    state.priceSplit.groups.forEach((g, i) => {
      g.share = sum > 0 ? parsed[i] / sum : 0;
    });
  } else {
    state.priceSplit.groups.forEach((g, i) => { g.share = parsed[i]; });
  }
}

function ensurePriceSplit() {
  if (!state.priceSplit || !Array.isArray(state.priceSplit.groups) || state.priceSplit.groups.length === 0) {
    // Seed from any group names already used in legacy line items so the
    // user's previous parties become the default split automatically.
    const names = new Set();
    for (const li of (state.lineItems || [])) {
      const parties = li.pricing?.parties || [];
      parties.forEach(p => p.name && names.add(p.name));
    }
    if (!names.has("Mine")) names.add("Mine");
    const list = ["Mine", ...[...names].filter(n => n !== "Mine")];
    const equal = (1 / list.length);
    const palette = ["indigo", "rose", "emerald", "amber"];
    state.priceSplit = {
      groups: list.map((name, i) => ({
        id: "g" + (i + 1),
        name,
        color: palette[i % palette.length],
        shareInput: list.length === 1 ? "100%" : `1/${list.length}`,
        share: equal,
      })),
    };
  } else {
    const palette = ["indigo", "rose", "emerald", "amber"];
    for (let i = 0; i < state.priceSplit.groups.length; i++) {
      const g = state.priceSplit.groups[i];
      if (!g.color) g.color = palette[i % palette.length];
    }
  }
  recomputeShares();
}

function partyAmount(p) {
  return parseAmount(p.input) || 0;
}
function lineItemTotal(li) {
  // New shape: explicit total. Legacy: parties summed, or cost.
  if (li.total != null) return Number(li.total) || 0;
  if (li.pricing && Array.isArray(li.pricing.parties)) {
    return li.pricing.parties.reduce((s, p) => s + partyAmount(p), 0);
  }
  return Number(li.cost) || 0;
}
// Resolve a single override entry to a $ amount given the line-item total.
// Returns NaN for invalid entries; null if no override set.
function resolveOverrideAmount(entry, total) {
  if (entry == null) return null;
  if (entry && typeof entry === "object" && "pct" in entry) {
    const pct = Number(entry.pct);
    return isNaN(pct) ? NaN : total * (pct / 100);
  }
  return parseAmount(entry);
}

// Compute per-group $ amounts for a total + overrides map. Overridden groups
// take their explicit amount; remaining groups split the leftover by their
// default shares (normalized among themselves).
function splitWithOverrides(total, overrides, groups) {
  let overSum = 0;
  let unsharedShare = 0;
  const overrideAmts = {};
  for (const g of groups) {
    const v = resolveOverrideAmount(overrides?.[g.id], total);
    if (v != null && !isNaN(v)) {
      overrideAmts[g.id] = v;
      overSum += v;
    } else {
      unsharedShare += g.share || 0;
    }
  }
  const remainder = total - overSum;
  const out = {};
  for (const g of groups) {
    if (g.id in overrideAmts) out[g.id] = overrideAmts[g.id];
    else if (unsharedShare > 0) out[g.id] = remainder * ((g.share || 0) / unsharedShare);
    else out[g.id] = 0;
  }
  return { amounts: out, remainder, overSum };
}

function lineItemGroupAmount(li, groupId, groupIdx) {
  if (li.total != null) {
    const { amounts } = splitWithOverrides(
      Number(li.total) || 0, li.overrides || {}, state.priceSplit.groups
    );
    return amounts[groupId] || 0;
  }
  // Legacy parties model: amount stored per party; map by index.
  if (li.pricing && Array.isArray(li.pricing.parties)) {
    return partyAmount(li.pricing.parties[groupIdx] || {});
  }
  // Legacy cost — count entirely toward the first group.
  if (li.cost != null) return groupIdx === 0 ? (Number(li.cost) || 0) : 0;
  return 0;
}
function lineItemMineTotal(li) {
  ensurePriceSplit();
  return lineItemGroupAmount(li, state.priceSplit.groups[0].id, 0);
}
function lineItemOthersTotal(li) {
  ensurePriceSplit();
  return state.priceSplit.groups.slice(1).reduce((s, g, i) =>
    s + lineItemGroupAmount(li, g.id, i + 1), 0);
}

// Working state for the add/edit line-item form.
let formTotalInput = "";
const formOverrides = {}; // { groupId: input string }
const formPaid = {};      // { groupId: input string } — amount each group already paid

// How much a group actually paid (fronted) toward a line item. You (the first
// group) are the settle-up hub, so only other groups' payments are recorded;
// anything not covered by them is treated as paid by you.
function linePaidBy(li, gid) {
  return (li.paid && Number(li.paid[gid])) || 0;
}

// Parse a single override input. Trailing % → percent-of-total; otherwise
// a flat dollar amount. Returns { kind: "pct" | "amt", value } or null.
function parseOverrideInput(raw) {
  if (raw == null) return null;
  const s = String(raw).trim();
  if (!s) return null;
  if (/%\s*$/.test(s)) {
    const num = parseAmount(s.replace(/%\s*$/, ""));
    if (isNaN(num)) return null;
    return { kind: "pct", value: num };
  }
  const num = parseAmount(s);
  if (isNaN(num)) return null;
  return { kind: "amt", value: num };
}

function refreshSplitDisplay() {
  // Update the % spans + warn + downstream pricing without rebuilding rows.
  // Caller is responsible for already having updated state.priceSplit.
  recomputeShares();
  const rows = document.querySelectorAll("#pricing-split-rows .pricing-split-row");
  rows.forEach((row, idx) => {
    const g = state.priceSplit.groups[idx];
    if (!g) return;
    const pct = row.querySelector(".group-pct");
    if (pct) {
      pct.textContent = `${(g.share * 100).toFixed(1)}%`;
      pct.classList.remove("invalid");
    }
  });
  const sumShares = state.priceSplit.groups.reduce((s, g) => s + (g.share || 0), 0);
  const warn = document.getElementById("pricing-split-warn");
  if (warn) {
    warn.textContent = Math.abs(sumShares - 1) > 0.001
      ? `Shares total ${(sumShares * 100).toFixed(1)}% (expected 100%)`
      : "";
  }
  // Refresh just the bits that depend on shares — leave the split editor
  // and its inputs alone so focus is preserved while typing.
  renderPricingLineItems();
  renderPricingSummary();
  updateLineItemFormBreakdown();
}

function renderSplitEditor() {
  ensurePriceSplit();
  const container = document.getElementById("pricing-split-rows");
  if (!container) return;
  container.innerHTML = "";
  state.priceSplit.groups.forEach((g, idx) => {
    const row = document.createElement("div");
    row.className = "pricing-split-row";
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "group-name";
    nameInput.value = g.name;
    nameInput.placeholder = idx === 0 ? "Mine" : `Group ${idx + 1}`;
    nameInput.addEventListener("input", () => {
      g.name = nameInput.value;
      // Update displayed names without rebuilding rows.
      renderPricingLineItems();
      renderPricingSummary();
      renderLineItemForm();
      save();
    });
    const shareInput = document.createElement("input");
    shareInput.type = "text";
    shareInput.className = "group-share";
    shareInput.value = g.shareInput;
    shareInput.placeholder = "1/2, 50%, or 2";
    const pct = document.createElement("span");
    pct.className = "group-pct";
    shareInput.addEventListener("input", () => {
      g.shareInput = shareInput.value;
      refreshSplitDisplay();
      save();
    });
    // Inline color swatches — click to assign a color to this group.
    const swatchRow = document.createElement("span");
    swatchRow.className = "group-color-swatches";
    const palette = ["indigo","rose","emerald","amber","teal","violet","sky","pink","lime","orange","cyan","grey"];
    palette.forEach(c => {
      const sw = document.createElement("button");
      sw.type = "button";
      sw.className = `color-swatch outline-${c}` + (g.color === c ? " selected" : "");
      sw.style.background = `var(--${c})`;
      sw.title = c;
      sw.addEventListener("click", () => {
        g.color = c;
        save();
        renderSplitEditor();
        renderPricingLineItems();
        renderPricingSummary();
      });
      swatchRow.appendChild(sw);
    });
    row.appendChild(nameInput);
    row.appendChild(shareInput);
    row.appendChild(pct);
    row.appendChild(swatchRow);
    if (state.priceSplit.groups.length > 1) {
      const x = document.createElement("button");
      x.type = "button";
      x.className = "group-remove";
      x.textContent = "×";
      x.title = "Remove this group";
      x.addEventListener("click", () => {
        state.priceSplit.groups.splice(idx, 1);
        save(); renderPricing();
      });
      row.appendChild(x);
    }
    container.appendChild(row);
  });
  const addBtn = document.getElementById("pricing-add-split-group");
  if (addBtn) addBtn.style.display = state.priceSplit.groups.length >= 4 ? "none" : "";
  // Initial display update.
  refreshSplitDisplay();
}

function renderLineItemForm() {
  ensurePriceSplit();
  const totalEl = document.getElementById("pricing-total");
  if (totalEl && totalEl.value !== formTotalInput) totalEl.value = formTotalInput;
  if (totalEl && !totalEl.dataset.bound) {
    totalEl.dataset.bound = "1";
    totalEl.addEventListener("input", () => {
      formTotalInput = totalEl.value;
      updateLineItemFormBreakdown();
    });
  }
  // Render override rows (one per default-split group).
  const orContainer = document.getElementById("pricing-overrides-rows");
  if (orContainer) {
    orContainer.innerHTML = "";
    state.priceSplit.groups.forEach(g => {
      const row = document.createElement("div");
      row.className = "pricing-override-row";
      const lab = document.createElement("span");
      lab.className = "or-label";
      lab.textContent = g.name;
      const input = document.createElement("input");
      input.type = "text";
      input.className = "or-input";
      input.placeholder = `$ or % (e.g. 200 or 25%)`;
      input.value = formOverrides[g.id] || "";
      // Track value as typed (so Save uses the current text) but only
      // refresh the hint display when the field loses focus / commits.
      input.addEventListener("input", () => {
        if (input.value === "") delete formOverrides[g.id];
        else formOverrides[g.id] = input.value;
      });
      input.addEventListener("change", updateLineItemFormBreakdown);
      input.addEventListener("blur", updateLineItemFormBreakdown);
      const def = document.createElement("span");
      def.className = "or-default";
      row.appendChild(lab);
      row.appendChild(input);
      row.appendChild(def);
      orContainer.appendChild(row);
    });
  }
  // "Who paid" rows — one per group except you (the first group / hub).
  const paidContainer = document.getElementById("pricing-paid-rows");
  if (paidContainer) {
    paidContainer.innerHTML = "";
    // Current per-group shares (so "share" can prefill the right amount).
    const total = parseAmount(formTotalInput) || 0;
    const ovMap = {};
    for (const g of state.priceSplit.groups) {
      const parsed = parseOverrideInput(formOverrides[g.id]);
      if (parsed) ovMap[g.id] = parsed.kind === "pct" ? { pct: parsed.value } : parsed.value;
    }
    const { amounts } = splitWithOverrides(total, ovMap, state.priceSplit.groups);

    // Header row (aligned columns, no visible table).
    ["", "Full", "Share", "Amount"].forEach((h, i) =>
      paidContainer.appendChild(el("span", "pp-h" + (i ? " pp-center" : ""), h)));

    state.priceSplit.groups.slice(1).forEach(g => {
      const share = Math.round((amounts[g.id] || 0) * 100) / 100;
      const name = el("span", "or-label", g.name);
      const full = document.createElement("input");
      full.type = "checkbox"; full.className = "settle-check pp-center";
      full.title = `${g.name} paid the full amount`;
      const shareCb = document.createElement("input");
      shareCb.type = "checkbox"; shareCb.className = "settle-check pp-center";
      shareCb.title = `${g.name} paid their share (${fmtMoney(share)})`;
      const input = document.createElement("input");
      input.type = "text"; input.className = "or-input pp-amount";
      input.placeholder = "$";
      input.value = formPaid[g.id] || "";

      const sync = () => {
        const t = parseAmount(formTotalInput) || 0;
        const c = parseAmount(formPaid[g.id]);
        full.checked = t > 0 && !isNaN(c) && Math.abs(c - t) < 0.005;
        shareCb.checked = !full.checked && share > 0 && !isNaN(c) && Math.abs(c - share) < 0.005;
      };
      const setAmt = (v) => {
        if (v == null || v === "" || !(Number(v) > 0)) delete formPaid[g.id];
        else formPaid[g.id] = String(v);
        input.value = formPaid[g.id] || "";
        sync();
      };
      sync();
      full.addEventListener("change", () => setAmt(full.checked ? (parseAmount(formTotalInput) || 0) : ""));
      shareCb.addEventListener("change", () => setAmt(shareCb.checked ? share : ""));
      input.addEventListener("input", () => {
        if (input.value === "") delete formPaid[g.id]; else formPaid[g.id] = input.value;
        sync();
      });

      paidContainer.appendChild(name);
      paidContainer.appendChild(full);
      paidContainer.appendChild(shareCb);
      paidContainer.appendChild(input);
    });
  }
  updateLineItemFormBreakdown();
}

function updateLineItemFormBreakdown() {
  const total = parseAmount(formTotalInput) || 0;
  // Build an overrides map matching the saved shape so we can reuse the helper.
  const overrides = {};
  for (const g of state.priceSplit.groups) {
    const parsed = parseOverrideInput(formOverrides[g.id]);
    if (!parsed) continue;
    overrides[g.id] = parsed.kind === "pct" ? { pct: parsed.value } : parsed.value;
  }
  const { amounts } = splitWithOverrides(total, overrides, state.priceSplit.groups);
  const out = state.priceSplit.groups.map(g => `${g.name}: ${fmtMoney(amounts[g.id] || 0)}`);
  const breakdown = document.getElementById("pricing-default-breakdown");
  if (breakdown) breakdown.textContent = out.join(" · ");
  // Normalize shares so defaults reflect actual proportions even if shares
  // weren't recomputed in state.
  const shareSum = state.priceSplit.groups.reduce((s, g) => s + (g.share || 0), 0) || 1;
  const defaultAmt = (g) => total * ((g.share || 0) / shareSum);
  // Hint block inside the override panel: total + default split.
  const hint = document.getElementById("pricing-overrides-hint");
  if (hint) {
    const defaults = state.priceSplit.groups
      .map(g => `${g.name} ${fmtMoney(defaultAmt(g))}`)
      .join(" · ");
    hint.innerHTML = `<span class="ph-total">Total ${fmtMoney(total)}</span> · Defaults: ${defaults}`;
  }
  // Update each row's computed-amount display.
  const rows = document.querySelectorAll("#pricing-overrides-rows .pricing-override-row");
  rows.forEach((row, idx) => {
    const g = state.priceSplit.groups[idx];
    if (!g) return;
    const def = row.querySelector(".or-default");
    if (!def) return;
    const isOverridden = g.id in overrides;
    if (isOverridden) {
      def.textContent = `= ${fmtMoney(amounts[g.id] || 0)} (default ${fmtMoney(defaultAmt(g))})`;
    } else {
      def.textContent = `= ${fmtMoney(amounts[g.id] || 0)} (remainder)`;
    }
  });
}

function addPricingLineItem() {
  const labelInput = document.getElementById("pricing-label");
  const total = parseAmount(formTotalInput);
  // A line item needs either a name (a manual cost not tied to itinerary items)
  // or one or more bundled items from the list — either is fine in any mode.
  if (!labelInput.value.trim() && pricingSelection.size === 0) {
    alert("Enter a name for this line item, or select items from the list to bundle.");
    return;
  }
  if (isNaN(total) || total <= 0) { alert("Enter a total amount."); return; }
  ensureLineItems();
  // Snapshot overrides into a plain {groupId: amount} object (parsed numbers).
  const overrides = {};
  for (const [gid, raw] of Object.entries(formOverrides)) {
    const parsed = parseOverrideInput(raw);
    if (!parsed) continue;
    overrides[gid] = parsed.kind === "pct" ? { pct: parsed.value } : parsed.value;
  }
  const laneSel = document.getElementById("pricing-lane");
  const lane = laneSel?.value || null;
  // Snapshot who-paid amounts (parsed dollars) keyed by group id.
  const paid = {};
  for (const [gid, raw] of Object.entries(formPaid)) {
    const amt = parseAmount(raw);
    if (!isNaN(amt) && amt > 0) paid[gid] = amt;
  }
  const hasPaid = Object.keys(paid).length > 0;
  if (editingLineItemId) {
    const li = state.lineItems.find(x => x.id === editingLineItemId);
    if (li) {
      li.eventIds = [...pricingSelection];
      li.label = labelInput.value.trim() || null;
      li.total = total;
      li.overrides = overrides;
      if (hasPaid) li.paid = paid; else delete li.paid;
      if (lane) li.lane = lane; else delete li.lane;
      delete li.cost;
      delete li.pricing;
    }
    editingLineItemId = null;
    const addBtn = document.getElementById("pricing-add-btn");
    if (addBtn) addBtn.textContent = "+ Add line item";
    const heading = document.getElementById("pricing-add-heading");
    if (heading) heading.textContent = "Add line item";
    const cancelBtn = document.getElementById("pricing-cancel-edit");
    if (cancelBtn) cancelBtn.hidden = true;
  } else {
    state.lineItems.push({
      id: "li" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5),
      eventIds: [...pricingSelection],
      label: labelInput.value.trim() || null,
      total,
      overrides,
      ...(hasPaid ? { paid } : {}),
      ...(lane ? { lane } : {}),
    });
  }
  pricingSelection.clear();
  labelInput.value = "";
  if (laneSel) laneSel.value = "";
  formTotalInput = "";
  for (const k of Object.keys(formOverrides)) delete formOverrides[k];
  for (const k of Object.keys(formPaid)) delete formPaid[k];
  save();
  renderPricing();
}

// "still need to book" panel. Lists tentative events; lets you bundle a
// few together with a label and book them as a group.
const todoSelection = new Set();

function ensureTodoBundles() {
  if (!Array.isArray(state.todoBundles)) state.todoBundles = [];
}
function findBundleForEvent(eventId) {
  ensureTodoBundles();
  return state.todoBundles.find(b => b.eventIds.includes(eventId));
}

function todoSwatch(ev) {
  const swatch = document.createElement("span");
  swatch.className = "todo-swatch";
  const colorVal = ev.color || "indigo";
  if (colorVal.startsWith("#")) swatch.style.background = colorVal;
  else swatch.style.backgroundColor = `var(--${colorVal})`;
  return swatch;
}

function renderTodoList() {
  const panel = document.getElementById("todo-panel");
  const list = document.getElementById("todo-list");
  if (!panel || !list) return;
  ensureTodoBundles();

  const tentative = state.events.filter(e => e.tentative)
    .sort((a, b) => (a.start || "").localeCompare(b.start || ""));
  // Panel stays visible even with zero tentative items, so the "+ Add to-do
  // item" input at the bottom is always reachable. Bundle/list cleanup still
  // runs.
  panel.hidden = false;
  if (tentative.length === 0) {
    list.innerHTML = "";
    todoSelection.clear();
    state.todoBundles = state.todoBundles.filter(b =>
      b.eventIds.some(id => state.events.find(e => e.id === id)?.tentative));
    return;
  }
  list.innerHTML = "";

  // Drop bundles that no longer have any tentative members.
  const tentIds = new Set(tentative.map(e => e.id));
  state.todoBundles = state.todoBundles
    .map(b => ({ ...b, eventIds: b.eventIds.filter(id => tentIds.has(id)) }))
    .filter(b => b.eventIds.length > 0);

  // Render bundles first.
  for (const bundle of state.todoBundles) {
    const li = document.createElement("li");
    li.className = "bundle";
    const head = document.createElement("div");
    head.className = "bundle-head";
    const label = document.createElement("span");
    label.className = "bundle-label";
    label.textContent = bundle.label || "(unlabeled bundle)";
    const meta = document.createElement("span");
    meta.className = "bundle-meta";
    meta.textContent = `${bundle.eventIds.length} items`;
    const actions = document.createElement("span");
    actions.className = "bundle-actions";
    const bookBtn = document.createElement("button");
    bookBtn.textContent = "Mark booked";
    bookBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      bundle.eventIds.forEach(id => {
        const ev = state.events.find(x => x.id === id);
        if (ev) ev.tentative = false;
      });
      state.todoBundles = state.todoBundles.filter(b => b.id !== bundle.id);
      save();
      renderApp();
    });
    const renameBtn = document.createElement("button");
    renameBtn.textContent = "Rename";
    renameBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const next = prompt("Rename bundle:", bundle.label || "");
      if (next == null) return;
      bundle.label = next.trim() || bundle.label;
      save();
      renderTodoList();
    });
    const unBtn = document.createElement("button");
    unBtn.textContent = "Unbundle";
    unBtn.className = "danger";
    unBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      state.todoBundles = state.todoBundles.filter(b => b.id !== bundle.id);
      save();
      renderTodoList();
    });
    actions.appendChild(bookBtn);
    actions.appendChild(renameBtn);
    actions.appendChild(unBtn);
    head.appendChild(label);
    head.appendChild(meta);
    head.appendChild(actions);

    const children = document.createElement("div");
    children.className = "bundle-children";
    for (const id of bundle.eventIds) {
      const ev = state.events.find(x => x.id === id);
      if (!ev) continue;
      const child = document.createElement("div");
      child.className = "bundle-child";
      child.appendChild(todoSwatch(ev));
      const t = document.createElement("span");
      t.textContent = ev.title;
      child.appendChild(t);
      const date = document.createElement("span");
      date.textContent = ev.start === ev.end ? ev.start : `${ev.start} → ${ev.end}`;
      child.appendChild(date);
      child.addEventListener("click", () => openEventDialog(ev.id, null));
      child.style.cursor = "pointer";
      children.appendChild(child);
    }
    li.appendChild(head);
    li.appendChild(children);
    list.appendChild(li);
  }

  // Then unbundled tentative events.
  for (const ev of tentative) {
    if (findBundleForEvent(ev.id)) continue;
    const li = document.createElement("li");
    const title = document.createElement("span");
    title.className = "todo-title";
    title.textContent = ev.title;
    const meta = document.createElement("span");
    meta.className = "todo-meta";
    meta.textContent = ev.start === ev.end ? ev.start : `${ev.start} → ${ev.end}`;
    li.appendChild(todoSwatch(ev));
    li.appendChild(title);
    li.appendChild(meta);
    li.addEventListener("click", () => openEventDialog(ev.id, null));
    list.appendChild(li);
  }
}

function addTodoBundle() {
  ensureTodoBundles();
  if (todoSelection.size < 2) {
    alert("Select at least 2 items to bundle.");
    return;
  }
  const labelInput = document.getElementById("todo-bundle-label");
  state.todoBundles.push({
    id: "tb" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5),
    eventIds: [...todoSelection],
    label: labelInput.value.trim() || null,
  });
  todoSelection.clear();
  labelInput.value = "";
  save();
  renderTodoList();
}

// --- flight map (Leaflet) ---

let _flightMap = null;          // Leaflet map instance
const _flightMapLayers = [];    // markers + polylines we own, to clear on re-render

function renderFlightMap() {
  const panel = document.getElementById("map-panel");
  const container = document.getElementById("flight-map");
  if (!panel || !container || typeof L === "undefined") return;

  // Pull the destination chain out of the user's flights. Skip layover events
  // and require a "AAA → BBB" code pair in the title to look up coordinates.
  const flights = state.events
    .filter(e => e.lane === "flights" && e.start && e.title
      && !(e.title || "").toLowerCase().endsWith("layover"))
    .sort((a, b) => (a.start + (a.startTime || "00:00"))
      .localeCompare(b.start + (b.startTime || "00:00")));
  const legs = [];
  for (const f of flights) {
    const m = (f.title || "").match(/\b([A-Z]{3})\s*(?:→|->|-|–)\s*([A-Z]{3})\b/);
    if (!m) continue;
    const a = AIRPORT_COORDS[m[1]];
    const b = AIRPORT_COORDS[m[2]];
    if (!a || !b) continue;
    legs.push({ from: m[1], to: m[2], a, b });
  }
  if (legs.length === 0) {
    panel.hidden = true;
    if (_flightMap) { _flightMap.remove(); _flightMap = null; _flightMapLayers.length = 0; }
    return;
  }
  panel.hidden = false;

  if (!_flightMap) {
    _flightMap = L.map(container, { worldCopyJump: true, scrollWheelZoom: false });
    L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; OpenStreetMap",
      maxZoom: 12,
    }).addTo(_flightMap);
  }
  // Clear previous markers/lines.
  for (const layer of _flightMapLayers) _flightMap.removeLayer(layer);
  _flightMapLayers.length = 0;

  // Distinct colors per airport — cycles if you visit more than the palette.
  const PALETTE = [
    "#4f46e5", "#f59e0b", "#10b981", "#ef4444", "#06b6d4",
    "#a855f7", "#ec4899", "#14b8a6", "#f97316", "#8b5cf6",
    "#22c55e", "#0ea5e9", "#eab308", "#d946ef", "#84cc16",
  ];
  const colorFor = new Map();
  const points = [];
  const legendEntries = []; // [{code, name, color}]
  const addAirport = (code, coords) => {
    if (colorFor.has(code)) return;
    const color = PALETTE[colorFor.size % PALETTE.length];
    colorFor.set(code, color);
    const [lat, lon, name] = coords;
    points.push([lat, lon]);
    legendEntries.push({ code, name, color });
    const marker = L.circleMarker([lat, lon], {
      radius: 7, color: "#fff", fillColor: color, fillOpacity: 1, weight: 2,
    }).bindTooltip(`${code} — ${name}`, { direction: "top", offset: [0, -8] })
      .addTo(_flightMap);
    _flightMapLayers.push(marker);
  };
  for (const { from, to, a, b } of legs) {
    addAirport(from, a);
    addAirport(to, b);
  }
  _flightMap.fitBounds(points, { padding: [30, 30] });
  // Leaflet needs a size invalidation after the container becomes visible.
  setTimeout(() => { _flightMap?.invalidateSize(); }, 0);

  // Render legend at bottom of the map panel.
  const legendEl = document.getElementById("flight-map-legend");
  if (legendEl) {
    legendEl.innerHTML = "";
    for (const { code, name, color } of legendEntries) {
      const item = document.createElement("span");
      item.className = "flight-map-legend-item";
      const dot = document.createElement("span");
      dot.className = "flight-map-legend-dot";
      dot.style.background = color;
      const codeEl = document.createElement("span");
      codeEl.className = "flight-map-legend-code";
      codeEl.textContent = code;
      const nameEl = document.createElement("span");
      nameEl.className = "flight-map-legend-name";
      nameEl.textContent = name;
      item.appendChild(dot);
      item.appendChild(codeEl);
      item.appendChild(nameEl);
      legendEl.appendChild(item);
    }
  }
}

// --- breakdown segment sizing ---

function chooseSegmentSize(totalDays) {
  if (state.segmentSize !== "auto") return Number(state.segmentSize);
  if (totalDays <= 7) return totalDays;
  if (totalDays <= 60) return 7;
  return 14;
}

// --- top-level render ---

// Expand the trip's start/end so they always cover every real (non-todo) event.
// Expand-only — a range set wider than the events is left as-is. Fixes trips
// whose stored dates are narrower than the itinerary (e.g. a flight that departs
// before the "Where" range) so the total-days count and timeline include it.
function reconcileTripDates() {
  for (const e of state.events) {
    if (!e.start || !e.end || isTypedTodo(e)) continue;
    if (!state.start || e.start < state.start) state.start = e.start;
    if (!state.end || e.end > state.end) state.end = e.end;
  }
}

function render() {
  reconcileTripDates();
  document.getElementById("trip-name").value = state.name;
  document.getElementById("trip-start").value = state.start || "";
  document.getElementById("trip-end").value = state.end || "";
  document.getElementById("segment-size").value = state.segmentSize;
  document.getElementById("tz-aware").checked = state.tzAware !== false;

  const overview = document.getElementById("overview");
  const breakdown = document.getElementById("breakdown");

  if (!state.start || !state.end || state.end < state.start) {
    overview.innerHTML = `<div class="empty-state">Set start and end dates to build your timeline.</div>`;
    breakdown.innerHTML = "";
    document.getElementById("trip-length").textContent = "";
    const bp = document.getElementById("breakdown-panel");
    if (bp) bp.hidden = true;
    return;
  }

  const totalDays = dayDiff(state.start, state.end) + 1;
  document.getElementById("trip-length").textContent = `${totalDays} day${totalDays === 1 ? "" : "s"}`;

  // Short trips: hide the Breakdown section entirely (overview covers it) — but
  // the list views (day-by-day / by-type) are useful at any length, so keep the
  // panel for those.
  const breakdownView = ["outline", "bytype"].includes(state.breakdownView) ? state.breakdownView : "timeline";
  const breakdownPanel = document.getElementById("breakdown-panel");
  if (breakdownPanel) breakdownPanel.hidden = totalDays < 5 && breakdownView === "timeline";

  const homeTz = state.homeTz || "America/Los_Angeles";
  const tzAware = state.tzAware !== false;
  const dayTzMap = computeDayTzMap(state.start, state.end, state.events, homeTz, tzAware);

  // Overview: on desktop, stretch to fill the panel (everything fits, no
  // scroll). On a phone, use a fixed readable day width so the timeline scrolls
  // horizontally instead of cramming the whole trip into the narrow screen.
  // Typed to-do items are excluded (To do tab); tentative events still show.
  const isPhone = window.matchMedia("(max-width: 720px)").matches;
  renderTimeline(overview, state.start, state.end, {
    dayTzMap, homeTz, dayPx: isPhone ? 48 : null, compact: totalDays > 14, tzAware,
    events: timelineEvents(),
  });

  renderFlightMap();

  renderTodoList();

  // Breakdown: fixed day width — short segments stay physically short.
  breakdown.innerHTML = "";

  // Reflect the active view on the toggle and show/hide the "Group by" control.
  const bToggle = document.getElementById("breakdown-view-toggle");
  if (bToggle) bToggle.querySelectorAll("button").forEach(b =>
    b.classList.toggle("active", b.dataset.bview === breakdownView));
  const segCtrls = document.getElementById("segment-controls");
  if (segCtrls) segCtrls.hidden = breakdownView !== "timeline";

  // Day-by-day outline view: chronological list instead of segmented timelines.
  if (breakdownView === "outline") {
    renderItineraryOutline(breakdown);
    return;
  }
  // By-type view: all flights, then all hotels, etc., each in date order.
  if (breakdownView === "bytype") {
    renderByType(breakdown);
    return;
  }

  // Region day-count summary: total days at each "Where" location, so users
  // see "Zanzibar 6d · Safari 5d · Seychelles 7d" at a glance.
  const regionRow = renderRegionSummary();
  if (regionRow) breakdown.appendChild(regionRow);

  const segSize = chooseSegmentSize(totalDays);

  // If the user toggled region chips, render one segment per matching
  // location event (using its own date range) instead of auto-segmenting.
  if (selectedRegions.size > 0) {
    const matches = state.events
      .filter(ev => ev.lane === "location" && ev.start && ev.end
        && selectedRegions.has((ev.title || "Untitled").trim()))
      .sort((a, b) => a.start.localeCompare(b.start));
    if (matches.length === 0) {
      breakdown.appendChild(Object.assign(document.createElement("div"), { className: "empty-state", textContent: "No matching regions." }));
      return;
    }
    const breakdownWidth = breakdown.clientWidth || 1200;
    const laneLabelW = parseInt(getComputedStyle(document.documentElement).getPropertyValue("--lane-label-w")) || 110;
    const longest = matches.reduce((m, ev) => Math.max(m, dayDiff(ev.start, ev.end) + 1), 1);
    const dayPx = Math.max(60, Math.floor((breakdownWidth - laneLabelW - 4) / longest));
    for (const ev of matches) {
      const segStart = ev.start;
      const segEnd = ev.end;
      const seg = el("div", "segment");
      const head = el("div", "segment-head");
      head.appendChild(el("span", "segment-title", (ev.title || "Untitled").trim()));
      head.appendChild(el("span", "segment-range",
        `${fmtShort(parseDay(segStart))} – ${fmtShort(parseDay(segEnd))} · ${dayDiff(segStart, segEnd) + 1} days`));
      seg.appendChild(head);
      const tl = el("div", "timeline");
      seg.appendChild(tl);
      breakdown.appendChild(seg);
      renderTimeline(tl, segStart, segEnd, {
        dayTzMap, homeTz, dayPx, compact: false, tzAware, events: timelineEvents(),
      });
    }
    return;
  }

  if (totalDays <= 7 || segSize >= totalDays) {
    breakdown.appendChild(Object.assign(document.createElement("div"), { className: "empty-state", textContent: "Trip is short — overview shows the full breakdown." }));
    return;
  }

  // Pick a day width so the longest segment fills the breakdown panel,
  // and shorter segments stay proportionally narrower.
  const breakdownWidth = breakdown.clientWidth || 1200;
  const laneLabelW = parseInt(getComputedStyle(document.documentElement).getPropertyValue("--lane-label-w")) || 110;
  const dayPx = Math.max(60, Math.floor((breakdownWidth - laneLabelW - 4) / segSize));

  // The archive trip can span years and have huge gaps — skip segments
  // with zero events so the breakdown stays useful instead of scrolling
  // through pages of empty rows.
  const isArchive = tripSlug() === "archive";
  const hasEventsInRange = (a, b) => state.events.some(ev =>
    ev.start && ev.end && ev.start <= b && ev.end >= a);

  let cursor = parseDay(state.start);
  let idx = 1;
  while (toISO(cursor) <= state.end) {
    const segStart = toISO(cursor);
    const segEndDate = addDays(cursor, segSize - 1);
    const segEnd = toISO(segEndDate) > state.end ? state.end : toISO(segEndDate);

    if (isArchive && !hasEventsInRange(segStart, segEnd)) {
      cursor = addDays(cursor, segSize);
      idx++;
      continue;
    }

    const seg = el("div", "segment");
    const head = el("div", "segment-head");
    head.appendChild(el("span", "segment-title", `Segment ${idx}`));
    head.appendChild(el("span", "segment-range",
      `${fmtShort(parseDay(segStart))} – ${fmtShort(parseDay(segEnd))} · ${dayDiff(segStart, segEnd) + 1} days`));
    seg.appendChild(head);

    const tl = el("div", "timeline");
    seg.appendChild(tl);
    breakdown.appendChild(seg);

    renderTimeline(tl, segStart, segEnd, {
      dayTzMap, homeTz, dayPx, compact: false, tzAware, events: timelineEvents(),
    });

    cursor = addDays(cursor, segSize);
    idx++;
  }
}

// Sum days at each location ("Where" lane). Returns a chip row element, or
// null if there are no location events.
// Per-device UI state — which region chips are toggled on to filter the
// breakdown. Not persisted; resets on reload.
const selectedRegions = new Set();

function renderRegionSummary() {
  const locs = state.events.filter(ev => ev.lane === "location" && ev.start && ev.end);
  if (locs.length === 0) return null;
  const totals = new Map();
  for (const ev of locs) {
    const days = dayDiff(ev.start, ev.end) + 1;
    const key = (ev.title || "Untitled").trim();
    totals.set(key, (totals.get(key) || 0) + days);
  }
  // Drop any selected regions that no longer exist (renamed/deleted).
  for (const name of [...selectedRegions]) {
    if (!totals.has(name)) selectedRegions.delete(name);
  }
  const row = document.createElement("div");
  row.className = "region-summary";
  for (const [name, days] of totals) {
    const chip = document.createElement("span");
    chip.className = "region-chip region-clickable";
    if (selectedRegions.has(name)) chip.classList.add("region-selected");
    chip.title = "Click to filter breakdown to this region";
    const lbl = document.createElement("span");
    lbl.className = "region-label";
    lbl.textContent = name;
    const val = document.createElement("span");
    val.className = "region-days";
    val.textContent = `${days}d`;
    chip.appendChild(lbl);
    chip.appendChild(val);
    chip.addEventListener("click", () => {
      if (selectedRegions.has(name)) selectedRegions.delete(name);
      else selectedRegions.add(name);
      render();
    });
    row.appendChild(chip);
  }
  const total = [...totals.values()].reduce((s, n) => s + n, 0);
  const tot = document.createElement("span");
  tot.className = "region-chip region-total";
  if (selectedRegions.size > 0) {
    tot.classList.add("region-clickable");
    tot.title = "Clear region filter";
    tot.textContent = `Clear filter (${selectedRegions.size})`;
    tot.addEventListener("click", () => {
      selectedRegions.clear();
      render();
    });
  } else {
    tot.textContent = `Total ${total}d`;
  }
  row.appendChild(tot);
  return row;
}

// --- event dialog ---

const dialog = document.getElementById("event-dialog");
const form = document.getElementById("event-form");

function findEvent(id) {
  // Look in main first, then in any option's events.
  const main = state.events.find(e => e.id === id);
  if (main) return { ev: main, optionId: null };
  for (const opt of state.options) {
    const ev = opt.events.find(e => e.id === id);
    if (ev) return { ev, optionId: opt.id };
  }
  return null;
}

const NAMED_COLORS = ["indigo","teal","sky","cyan","emerald","lime","amber","orange","rose","pink","violet","grey"];

function buildColorGrid(selected) {
  const grid = document.getElementById("color-grid");
  if (!grid) return;
  grid.innerHTML = "";
  const isHex = selected && selected.startsWith("#");
  for (const c of NAMED_COLORS) {
    const label = document.createElement("label");
    label.title = c;
    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = "color";
    radio.value = c;
    if (!isHex && (selected === c || (!selected && c === "indigo"))) radio.checked = true;
    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.style.background = `var(--${c})`;
    label.appendChild(radio);
    label.appendChild(swatch);
    grid.appendChild(label);
  }
  // Custom color picker — uses a radio so it participates in the form, plus an
  // <input type=color> that updates the radio's value when changed.
  const customLabel = document.createElement("label");
  customLabel.title = "Custom color";
  const customRadio = document.createElement("input");
  customRadio.type = "radio";
  customRadio.name = "color";
  customRadio.value = isHex ? selected : "#888888";
  if (isHex) customRadio.checked = true;
  const customSwatch = document.createElement("span");
  customSwatch.className = "swatch custom-swatch";
  customSwatch.textContent = "+";
  if (isHex) customSwatch.style.background = selected;
  const colorPicker = document.createElement("input");
  colorPicker.type = "color";
  colorPicker.value = isHex ? selected : "#888888";
  colorPicker.addEventListener("input", (e) => {
    customRadio.value = e.target.value;
    customRadio.checked = true;
    customSwatch.style.background = e.target.value;
    customSwatch.textContent = "";
  });
  // Click anywhere on the rainbow swatch opens the native color picker.
  // (Clicking the wrapping <label> would normally just select the radio.)
  customLabel.addEventListener("click", (e) => {
    e.preventDefault();
    colorPicker.click();
  });
  customSwatch.appendChild(colorPicker);
  customLabel.appendChild(customRadio);
  customLabel.appendChild(customSwatch);
  grid.appendChild(customLabel);
}

function openEventDialog(id, optionId) {
  form.reset();
  const titleEl = document.getElementById("event-dialog-title");
  const deleteBtn = document.getElementById("event-delete");
  const mergeBtn = document.getElementById("event-merge-btn");
  const mergePanel = document.getElementById("event-merge-panel");
  if (mergePanel) { mergePanel.hidden = true; mergePanel.innerHTML = ""; }

  if (id) {
    const found = findEvent(id);
    if (!found) return;
    const { ev, optionId: foundOpt } = found;
    titleEl.textContent = foundOpt ? "Edit option event" : "Edit event";
    form.elements.id.value = ev.id;
    form.elements.optionId.value = foundOpt || "";
    form.elements.title.value = ev.title;
    form.elements.lane.value = ev.lane || "activities";
    form.elements.start.value = ev.start;
    form.elements.startTime.value = ev.startTime || "";
    form.elements.end.value = ev.end;
    form.elements.endTime.value = ev.endTime || "";
    buildColorGrid(ev.color || "indigo");
    if (form.elements.tentative) form.elements.tentative.checked = !!ev.tentative;
    if (form.elements.confirmation) form.elements.confirmation.value = ev.confirmation || "";
    form.elements.notes.value = ev.notes || "";
    deleteBtn.hidden = false;
    // Top-level events (not inside a stand-alone Option) get the merge button.
    if (mergeBtn) {
      mergeBtn.hidden = !!foundOpt;
      const hasKids = state.events.some(e => e.mergedInto === ev.id);
      mergeBtn.textContent = hasKids ? "Unmerge" : "Merge";
    }
  } else {
    const opt = optionId ? state.options.find(o => o.id === optionId) : null;
    titleEl.textContent = opt ? `Add event to "${opt.name}"` : "Add event";
    form.elements.id.value = "";
    form.elements.optionId.value = optionId || "";
    form.elements.lane.value = "flights";
    const optGroup = opt ? getOptionGroupForOption(opt) : null;
    const defaultStart = optionId ? (optGroup?.start || state.optionRangeStart || state.start) : state.start;
    form.elements.start.value = defaultStart || "";
    form.elements.end.value = defaultStart || "";
    buildColorGrid(optionId ? "indigo" : "emerald");
    deleteBtn.hidden = true;
    if (mergeBtn) mergeBtn.hidden = true;
  }
  dialog.showModal();
}

// Sort key: higher = more likely merge target. Same lane heavily preferred,
// then same start+end, then overlap, then ±1 day adjacency, then same week.
function mergeLikelihood(target, candidate) {
  let s = 0;
  if (target.lane === candidate.lane) s += 1000;
  if (target.start === candidate.start && target.end === candidate.end) s += 500;
  // overlap
  if (target.start <= candidate.end && target.end >= candidate.start) s += 300;
  // ±1 day adjacency on either end
  const d = (a, b) => Math.abs(dayDiff(a, b));
  const adj = Math.min(d(target.end, candidate.start), d(target.start, candidate.end));
  if (adj <= 1) s += 200;
  else if (adj <= 7) s += 50;
  else s -= adj;
  return s;
}

function renderMergePanel(eventId) {
  const panel = document.getElementById("event-merge-panel");
  if (!panel) return;
  const me = state.events.find(e => e.id === eventId);
  if (!me) return;
  panel.innerHTML = "";

  // Show children-with-unmerge if this event is a parent.
  const children = state.events.filter(e => e.mergedInto === me.id);
  if (children.length) {
    const head = document.createElement("div");
    head.className = "merge-panel-head";
    head.textContent = `Merged with ${children.length} other event${children.length === 1 ? "" : "s"}:`;
    panel.appendChild(head);
    for (const c of children) {
      const row = document.createElement("div");
      row.className = "merge-row";
      const label = document.createElement("span");
      label.textContent = `${c.title} — ${c.start}${c.end !== c.start ? ` → ${c.end}` : ""} (${c.lane})`;
      row.appendChild(label);
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "ghost";
      btn.textContent = "Unmerge";
      btn.addEventListener("click", () => {
        delete c.mergedInto;
        save();
        renderMergePanel(eventId);
        render();
      });
      row.appendChild(btn);
      panel.appendChild(row);
    }
  }

  // Always also offer to merge another event into this one.
  const candidates = state.events
    .filter(e => e.id !== me.id && e.mergedInto !== me.id && !e.mergedInto)
    .map(e => ({ ev: e, score: mergeLikelihood(me, e) }))
    .sort((a, b) => b.score - a.score);
  if (candidates.length) {
    const head = document.createElement("div");
    head.className = "merge-panel-head";
    head.textContent = "Merge another event into this one:";
    panel.appendChild(head);
    for (const { ev: c } of candidates) {
      const row = document.createElement("div");
      row.className = "merge-row";
      const label = document.createElement("span");
      label.textContent = `${c.title} — ${c.start}${c.end !== c.start ? ` → ${c.end}` : ""} (${c.lane})`;
      row.appendChild(label);
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = "Merge";
      btn.addEventListener("click", () => {
        c.mergedInto = me.id;
        save();
        renderMergePanel(eventId);
        render();
      });
      row.appendChild(btn);
      panel.appendChild(row);
    }
  } else if (!children.length) {
    const empty = document.createElement("div");
    empty.className = "merge-panel-head";
    empty.textContent = "No other events available to merge.";
    panel.appendChild(empty);
  }
}

document.getElementById("event-merge-btn")?.addEventListener("click", () => {
  const id = document.getElementById("event-form").elements.id.value;
  if (!id) return;
  const panel = document.getElementById("event-merge-panel");
  if (panel.hidden) {
    renderMergePanel(id);
    panel.hidden = false;
  } else {
    panel.hidden = true;
  }
});

form.addEventListener("submit", (e) => {
  e.preventDefault();
  const data = Object.fromEntries(new FormData(form));
  if (!data.title || !data.start || !data.end) return;
  if (data.end < data.start) {
    alert("End day must be on or after start day.");
    return;
  }

  const updates = {
    title: data.title,
    lane: data.lane || "activities",
    start: data.start,
    end: data.end,
    color: data.color,
    notes: data.notes,
    tentative: !!data.tentative,
  };
  if (data.startTime) updates.startTime = data.startTime;
  if (data.endTime) updates.endTime = data.endTime;
  const conf = (data.confirmation || "").trim();
  if (conf) updates.confirmation = conf;

  const optionId = data.optionId || null;
  const targetList = optionId
    ? (state.options.find(o => o.id === optionId)?.events || state.events)
    : state.events;

  if (data.id) {
    const found = findEvent(data.id);
    if (found) {
      Object.assign(found.ev, updates);
      if (!data.startTime) delete found.ev.startTime;
      if (!data.endTime) delete found.ev.endTime;
      if (!conf) delete found.ev.confirmation;
      // If the user edits an auto-generated location, "claim" it so we don't
      // overwrite their changes the next time flights change.
      if (found.ev._autoLoc) delete found.ev._autoLoc;
    }
  } else {
    targetList.push({ id: uid(), ...updates });
  }
  save();
  dialog.close();
  renderApp();
});

document.getElementById("event-cancel").addEventListener("click", () => dialog.close());
document.getElementById("event-delete").addEventListener("click", () => {
  const id = form.elements.id.value;
  if (!id) return;
  const found = findEvent(id);
  if (found) {
    // Remember dismissed Where events so the auto-location reconcile doesn't
    // immediately recreate them on the next save.
    if (!found.optionId && found.ev.lane === "location" && found.ev.start && found.ev.end) {
      if (!Array.isArray(state.rejectedAutoLocs)) state.rejectedAutoLocs = [];
      state.rejectedAutoLocs.push({ start: found.ev.start, end: found.ev.end });
    }
    if (found.optionId) {
      const opt = state.options.find(o => o.id === found.optionId);
      opt.events = opt.events.filter(e => e.id !== id);
    } else {
      state.events = state.events.filter(e => e.id !== id);
    }
  }
  save();
  dialog.close();
  renderApp();
});

// --- top bar wiring ---

// Click-to-edit toggle for the trip name + dates header.
document.getElementById("trip-display")?.addEventListener("click", () => {
  document.getElementById("trip-display").hidden = true;
  document.getElementById("trip-edit-row").hidden = false;
  document.getElementById("trip-name").focus();
  document.getElementById("trip-name").select();
});
document.getElementById("trip-edit-done")?.addEventListener("click", () => {
  document.getElementById("trip-edit-row").hidden = true;
  document.getElementById("trip-display").hidden = false;
  renderApp();
});

function toggleEnabledTab(key, on) {
  if (!state.enabledTabs) state.enabledTabs = {};
  state.enabledTabs[key] = on;
  // Bounce off a disabled tab if it's the current view.
  if (!on && state.activeView === key) state.activeView = "main";
  save();
  renderApp();
}
document.getElementById("trip-menu-btn")?.addEventListener("click", (e) => {
  e.stopPropagation();
  const menu = document.getElementById("trip-menu");
  const btn = document.getElementById("trip-menu-btn");
  const open = menu.hidden;
  menu.hidden = !open;
  btn.setAttribute("aria-expanded", String(open));
});
document.addEventListener("click", (e) => {
  const menu = document.getElementById("trip-menu");
  if (!menu || menu.hidden) return;
  if (e.target.closest("#trip-menu") || e.target.closest("#trip-menu-btn")) return;
  menu.hidden = true;
  document.getElementById("trip-menu-btn")?.setAttribute("aria-expanded", "false");
});
document.getElementById("enable-tab-todo")?.addEventListener("change", (e) => {
  toggleEnabledTab("todo", e.target.checked);
});
document.getElementById("enable-tab-options")?.addEventListener("change", (e) => {
  toggleEnabledTab("options", e.target.checked);
});
document.getElementById("enable-tab-compare")?.addEventListener("change", (e) => {
  toggleEnabledTab("compare", e.target.checked);
});

document.getElementById("trip-name").addEventListener("input", (e) => {
  state.name = e.target.value;
  save();
});
document.getElementById("trip-start").addEventListener("change", (e) => {
  state.start = e.target.value || null;
  save();
  renderApp();
});
document.getElementById("trip-end").addEventListener("change", (e) => {
  state.end = e.target.value || null;
  save();
  renderApp();
});
document.getElementById("segment-size").addEventListener("change", (e) => {
  state.segmentSize = e.target.value;
  saveLocal();   // breakdown grouping is per-device UI
  renderApp();
});
document.getElementById("breakdown-view-toggle")?.addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-bview]");
  if (!btn) return;
  state.breakdownView = btn.dataset.bview;
  saveLocal();   // breakdown view choice is per-device UI
  renderApp();
});
document.getElementById("tz-aware").addEventListener("change", (e) => {
  state.tzAware = e.target.checked;
  save();
  renderApp();
});
document.getElementById("add-event-btn").addEventListener("click", () => openEventDialog(null));
document.getElementById("share-btn")?.addEventListener("click", shareTrip);

// --- Printable itinerary ---
//
// The on-screen views are pixel-positioned timelines that print poorly, so for
// printing we render a plain chronological list grouped by day. Multi-day items
// (lodging, region, rental) become a one-line "ongoing" banner repeated at the
// top of each day they span; single-day events and flights list below by time.

const PRINT_LANE_ORDER = { location: 0, lodging: 1, rental: 2, activities: 3, flights: 4 };

function fmtClock12(hhmm) {
  const [h, m] = hhmm.split(":").map(Number);
  const ap = h < 12 ? "am" : "pm";
  let hr = h % 12;
  if (hr === 0) hr = 12;
  return m ? `${hr}:${String(m).padStart(2, "0")}${ap}` : `${hr}${ap}`;
}

function printTimeLabel(ev) {
  if (ev.startTime && ev.endTime && ev.start === ev.end)
    return `${fmtClock12(ev.startTime)}–${fmtClock12(ev.endTime)}`;
  if (ev.startTime) return fmtClock12(ev.startTime);
  return "";   // no time → leave the time column blank
}

// Spell out a flight's departure and arrival on one line, with timezone
// labels (flights cross zones) and the arrival date when it lands next day.
function flightTimesLine(ev) {
  if (!ev.startTime && !ev.endTime) return "";
  const tz = (zone, dateStr) => zone ? " " + tzShortName(zone, dateStr) : "";
  const parts = [];
  if (ev.startTime) parts.push(`Departs ${fmtClock12(ev.startTime)}${tz(ev.startTz, ev.start)}`);
  if (ev.endTime) {
    let a = `arrives ${fmtClock12(ev.endTime)}${tz(ev.endTz, ev.end)}`;
    if (ev.end && ev.end !== ev.start)
      a += ` (${parseDay(ev.end).toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric" })})`;
    parts.push(a);
  }
  return parts.join(" · ");
}

// Clock change a traveler experiences on a flight: destination UTC offset minus
// origin UTC offset, in minutes (positive = destination is ahead). Null when we
// can't compute it (missing timezone or times).
function flightTzChangeMinutes(ev) {
  if (ev.lane !== "flights" || !ev.startTz || !ev.endTz || !ev.startTime || !ev.endTime) return null;
  const offMin = (zone, dateStr, timeStr) => {
    const [h, mn] = timeStr.split(":").map(Number);
    const [y, m, d] = dateStr.split("-").map(Number);
    const naive = Date.UTC(y, m - 1, d, h, mn);
    return Math.round((naive - wallToUtc(y, m, d, h, mn, zone)) / 60000);
  };
  return offMin(ev.endTz, ev.end, ev.endTime) - offMin(ev.startTz, ev.start, ev.startTime);
}

function fmtTzChange(min) {
  if (min === 0) return "no time change";
  const a = Math.abs(min), h = Math.floor(a / 60), m = a % 60;
  const hm = m ? `${h}h ${m}m` : `${h}h`;
  return `${min > 0 ? "+" : "−"}${hm} time change`;
}

// Multi-day stays/regions/rentals become the ongoing banner; flights always
// read as timed events even when they cross midnight.
function isOngoingEvent(ev) {
  return ev.lane !== "flights" && dayDiff(ev.start, ev.end) >= 1;
}

// Typed to-do items (the "+ Add to-do item" box) are flagged todo:true; legacy
// ones predating the flag match the signature the form used — a single-day
// violet activity with no times. These are the To-do list and are kept out of
// the timeline/breakdown/printout. Dated "still need to book" (tentative)
// events are NOT to-dos: they stay visible (shown as tentative).
function isTypedTodo(ev) {
  return ev.todo === true
    || (ev.tentative && ev.lane === "activities" && ev.color === "violet"
        && ev.start === ev.end && !ev.startTime && !ev.endTime);
}
function timelineEvents() {
  return state.events.filter(e => !isTypedTodo(e));
}

// Event length in minutes (handles crossing midnight via the dates).
function eventDurationMin(ev) {
  if (!ev.startTime || !ev.endTime) return null;
  const [sh, sm] = ev.startTime.split(":").map(Number);
  const [eh, em] = ev.endTime.split(":").map(Number);
  return (dayDiff(ev.start, ev.end) * 1440 + eh * 60 + em) - (sh * 60 + sm);
}
function fmtDurMin(min) {
  if (min == null || min < 0) return "";
  const h = Math.floor(min / 60), m = min % 60;
  return h ? (m ? `${h}h ${m}m` : `${h}h`) : `${m}m`;
}

// A layover sits in the flights lane but is a wait at the airport, not a
// flight — so it doesn't get Depart/Arrive treatment. (Same title convention
// as reconcileAutoLocations.)
function isLayover(ev) {
  return ev.lane === "flights" && (ev.title || "").trim().toLowerCase().endsWith("layover");
}

// Ground/water transport people sometimes file under the flights lane — these
// are not flights, so no Depart/Arrive tag or departs→arrives line.
const NON_FLIGHT_RX = /\b(drive|driving|transfer|train|bus|ferry|taxi|shuttle|boat)\b/i;
function isFlightEvent(ev) {
  return ev.lane === "flights" && !isLayover(ev) && !NON_FLIGHT_RX.test(ev.title || "");
}

// Trip name + date range header used at the top of every printout.
function printHeader(root) {
  const header = el("div", "print-header");
  header.appendChild(el("h1", "print-title", state.name || "Untitled trip"));
  if (state.start && state.end && state.end >= state.start) {
    const fmtLong = iso => parseDay(iso).toLocaleDateString(undefined, { month: "long", day: "numeric", year: "numeric" });
    const days = dayDiff(state.start, state.end) + 1;
    header.appendChild(el("div", "print-subtitle",
      `${fmtLong(state.start)} – ${fmtLong(state.end)} · ${days} day${days === 1 ? "" : "s"}`));
  }
  root.appendChild(header);
}

function buildPrintItinerary() {
  const root = document.getElementById("print-root");
  if (!root) return;
  root.innerHTML = "";
  printHeader(root);
  renderItineraryOutline(root);
}

function buildPrintByType() {
  const root = document.getElementById("print-root");
  if (!root) return;
  root.innerHTML = "";
  printHeader(root);
  renderByType(root);
}

// Swap a day's row with its neighbor and persist the new manual order for that
// day (state.dayOrder[isoDate] = [eventId, ...]). orderedIds is the day's
// current visual order; i is the row's index, dir is -1 (up) or +1 (down).
function moveDayRow(D, orderedIds, i, dir) {
  const j = i + dir;
  if (j < 0 || j >= orderedIds.length) return;
  const ids = orderedIds.slice();
  [ids[i], ids[j]] = [ids[j], ids[i]];
  if (!state.dayOrder) state.dayOrder = {};
  state.dayOrder[D] = ids;
  save();
  renderApp();
}

// "By type" view: group items by lane (Flights, Lodging, Transportation,
// Activities, Where), each list in chronological order. Layovers and typed
// to-dos are omitted; rows open the editor on click.
function renderByType(container) {
  if (!state.start || !state.end || state.end < state.start) {
    container.appendChild(el("div", "outline-empty", "Set trip start and end dates to see items by type."));
    return;
  }
  const events = state.events.filter(ev =>
    ev.start && ev.end && !isTypedTodo(ev) && !isLayover(ev));
  let any = false;
  for (const lane of LANE_ORDER) {
    const items = events.filter(ev => ev.lane === lane)
      .sort((a, b) => (a.start + (a.startTime || "00:00")).localeCompare(b.start + (b.startTime || "00:00")));
    if (!items.length) continue;
    any = true;
    const group = el("div", "bytype-group");
    group.appendChild(el("div", "bytype-head", LANE_LABEL[lane] || lane));
    for (const ev of items) {
      const row = el("div", "outline-event");
      const dateText = ev.start === ev.end
        ? fmtShort(parseDay(ev.start))
        : `${fmtShort(parseDay(ev.start))} – ${fmtShort(parseDay(ev.end))}`;
      row.appendChild(el("span", "outline-time bytype-date", dateText));
      const body = el("div", "outline-body");
      const titleLine = el("div", "outline-title", ev.title || "Untitled");
      if (ev.tentative) titleLine.appendChild(el("span", "outline-tent", " — tentative"));
      body.appendChild(titleLine);
      if (isFlightEvent(ev)) {
        const fl = flightTimesLine(ev);
        if (fl) body.appendChild(el("div", "outline-flight", fl));
        const tc = flightTzChangeMinutes(ev);
        const meta = [];
        if (tc !== null) meta.push(fmtTzChange(tc));
        if (ev.confirmation) meta.push(`Conf# ${ev.confirmation}`);
        if (meta.length) body.appendChild(el("div", "outline-meta", meta.join(" · ")));
      } else {
        if (ev.confirmation) body.appendChild(el("div", "outline-meta", `Conf# ${ev.confirmation}`));
        if (ev.notes) body.appendChild(el("div", "outline-notes", ev.notes));
      }
      row.appendChild(body);
      row.addEventListener("click", () => openEventDialog(ev.id));
      group.appendChild(row);
    }
    container.appendChild(group);
  }
  if (!any) container.appendChild(el("div", "outline-empty", "No items yet."));
}

// Build the chronological day-by-day outline into `container`. Shared by the
// printable itinerary and the on-screen Breakdown "Day by day" view. Multi-day
// stays/regions become an "ongoing" banner repeated atop each day they span;
// single-day events and flights list below by time, with flight detail,
// time-change, and confirmation lines. Rows open the editor on click.
function renderItineraryOutline(container) {
  if (!state.start || !state.end || state.end < state.start) {
    container.appendChild(el("div", "outline-empty", "Set trip start and end dates to see the day-by-day outline."));
    return;
  }

  // Typed to-do items stay in the To do tab only; tentative events are shown.
  const events = state.events.filter(ev => ev.start && ev.end && !isTypedTodo(ev));
  let cursor = parseDay(state.start);
  let any = false;

  while (toISO(cursor) <= state.end) {
    const D = toISO(cursor);
    // Regions / transportation spanning today (lodging handled separately so
    // single-night stays show too).
    const ongoing = events
      .filter(ev => ev.lane !== "lodging" && isOngoingEvent(ev) && ev.start <= D && D <= ev.end);
    // Lodging you sleep at the night of D: a single-night stay shows on its day;
    // a multi-night stay spans check-in through the night before checkout (so it
    // drops off the checkout day). On a transition night, the most recently
    // checked-into hotel wins.
    const lodgingTonight = events.filter(ev => ev.lane === "lodging"
      && (ev.start === ev.end ? ev.start === D : (ev.start <= D && D < ev.end)));
    if (lodgingTonight.length) {
      ongoing.push(lodgingTonight.reduce((a, b) => (b.start > a.start ? b : a)));
    }
    ongoing.sort((a, b) => (PRINT_LANE_ORDER[a.lane] ?? 9) - (PRINT_LANE_ORDER[b.lane] ?? 9)
      || (a.title || "").localeCompare(b.title || ""));
    // Events happening today: those that start today, plus overnight flights
    // that *land* today. A red-eye departing the 19th and arriving 6:20am on
    // the 20th gets an "Arrive" entry on the 20th so that day isn't blank.
    const starting = events.filter(ev =>
      !isOngoingEvent(ev) && ev.start === D && ev.lane !== "lodging");
    const arrivals = events.filter(ev =>
      isFlightEvent(ev) && ev.endTime && ev.end === D && ev.end !== ev.start);
    // Hotel check-ins: any lodging that begins today. Sorted to mid-afternoon
    // (≈3pm) but shown without an exact time, since check-in is often 3–4pm.
    const checkins = events.filter(ev => ev.lane === "lodging" && ev.start === D);
    // At the same time, an arrival comes before everything else (you land, then
    // the layover starts), and check-ins come last.
    const kindRank = { arrival: 0, event: 1, checkin: 2 };
    const dayEntries = [
      ...starting.map(ev => ({ ev, kind: "event", t: ev.startTime || "00:00" })),
      ...arrivals.map(ev => ({ ev, kind: "arrival", t: ev.endTime })),
      ...checkins.map(ev => ({ ev, kind: "checkin", t: ev.startTime || "15:00" })),
    ].sort((a, b) => a.t.localeCompare(b.t) || (kindRank[a.kind] - kindRank[b.kind]));

    // Manual per-day ordering (set via the up/down arrows) overrides the
    // time sort. Items not in the saved order keep their time position at the end.
    const savedOrder = state.dayOrder?.[D];
    if (savedOrder) {
      dayEntries.sort((a, b) => {
        const ia = savedOrder.indexOf(a.ev.id), ib = savedOrder.indexOf(b.ev.id);
        if (ia === -1 && ib === -1) return 0;
        if (ia === -1) return 1;
        if (ib === -1) return -1;
        return ia - ib;
      });
    }

    if (ongoing.length || dayEntries.length) {
      any = true;
      const dayBlock = el("div", "outline-day");
      dayBlock.appendChild(el("div", "outline-day-head",
        parseDay(D).toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" })));

      const banner = el("div", "outline-ongoing");
      for (const ev of ongoing) {
        const item = el("span", "outline-ongoing-item");
        item.appendChild(el("span", "outline-ongoing-label", LANE_LABEL[ev.lane] || ev.lane));
        item.appendChild(document.createTextNode(" " + (ev.title || "Untitled")));
        item.addEventListener("click", () => openEventDialog(ev.id));
        banner.appendChild(item);
      }
      if (banner.children.length) dayBlock.appendChild(banner);

      const orderedIds = dayEntries.map(e => e.ev.id);
      for (let index = 0; index < dayEntries.length; index++) {
        const { ev, kind } = dayEntries[index];
        // Only real flights get the Depart tag, departs→arrives line, and time
        // change. Layovers and ground transport (drives, etc.) read as a plain
        // timed row.
        const isFlight = isFlightEvent(ev);
        const row = el("div", "outline-event");
        // Flights anchor on departure time; arrival rows on arrival time;
        // check-ins read "Afternoon" (≈3pm); layovers show their duration (no
        // timestamp); everything else uses its own time/range.
        let timeText;
        if (kind === "arrival") timeText = fmtClock12(ev.endTime);
        else if (kind === "checkin") timeText = ev.startTime ? fmtClock12(ev.startTime) : "";
        else if (isLayover(ev)) timeText = "";   // duration shown under the title instead
        else timeText = isFlight && ev.startTime ? fmtClock12(ev.startTime) : printTimeLabel(ev);
        row.appendChild(el("span", "outline-time", timeText));

        const body = el("div", "outline-body");
        const titleLine = el("div", "outline-title");
        if (kind === "arrival") titleLine.appendChild(el("span", "outline-arrive-tag", "Arrive"));
        else if (kind === "checkin") titleLine.appendChild(el("span", "outline-checkin-tag", "Check-in"));
        else if (isFlight) titleLine.appendChild(el("span", "outline-depart-tag", "Depart"));
        const tagged = kind === "arrival" || kind === "checkin" || isFlight;
        titleLine.appendChild(document.createTextNode((tagged ? " " : "") + (ev.title || "Untitled")));
        if (ev.tentative) titleLine.appendChild(el("span", "outline-tent", " — tentative"));
        body.appendChild(titleLine);

        if (kind === "arrival") {
          // Where it came from, so the arrival makes sense without flipping back.
          const depTz = ev.startTz ? " " + tzShortName(ev.startTz, ev.start) : "";
          body.appendChild(el("div", "outline-flight",
            `Overnight — departed ${parseDay(ev.start).toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric" })}`
            + (ev.startTime ? ` ${fmtClock12(ev.startTime)}${depTz}` : "")));
          // Time change lands on the arrival — that's when the clock actually shifts.
          const tcArr = flightTzChangeMinutes(ev);
          if (tcArr !== null) body.appendChild(el("div", "outline-meta", fmtTzChange(tcArr)));
        } else if (kind === "checkin") {
          if (ev.confirmation) body.appendChild(el("div", "outline-meta", `Conf# ${ev.confirmation}`));
          if (ev.notes) body.appendChild(el("div", "outline-notes", ev.notes));
        } else {
          if (isLayover(ev)) {
            // Auto-created layovers already carry the duration in their notes
            // ("9hr 20min layover in …"), so only add a computed one if the
            // title/notes don't already mention a duration — avoids doubling it.
            const dur = fmtDurMin(eventDurationMin(ev));
            const hasDur = /\d\s*h(r|rs|our|ours)?\b/i.test(`${ev.title || ""} ${ev.notes || ""}`);
            if (dur && !hasDur) body.appendChild(el("div", "outline-flight", dur));
          } else if (isFlight) {
            const fl = flightTimesLine(ev);
            if (fl) body.appendChild(el("div", "outline-flight", fl));
          }
          // Meta line: confirmation number, plus the time change for same-day
          // flights (overnight flights show it on their Arrive row instead).
          const meta = [];
          const tc = (isFlight && ev.start === ev.end) ? flightTzChangeMinutes(ev) : null;
          if (tc !== null) meta.push(fmtTzChange(tc));
          if (ev.confirmation) meta.push(`Conf# ${ev.confirmation}`);
          if (meta.length) body.appendChild(el("div", "outline-meta", meta.join(" · ")));
          if (ev.notes) body.appendChild(el("div", "outline-notes", ev.notes));
        }

        row.appendChild(body);
        row.addEventListener("click", () => openEventDialog(ev.id));

        // Up/down arrows to reorder items within the day. Hidden on the printout
        // (see CSS) and only shown when there's more than one item to reorder.
        if (dayEntries.length > 1) {
          const reorder = el("div", "outline-reorder");
          const mkBtn = (glyph, dir, disabled, label) => {
            const b = el("button", "outline-reorder-btn", glyph);
            b.type = "button";
            b.title = label;
            b.setAttribute("aria-label", label);
            b.disabled = disabled;
            b.addEventListener("click", (e) => {
              e.stopPropagation();
              moveDayRow(D, orderedIds, index, dir);
            });
            return b;
          };
          reorder.appendChild(mkBtn("▲", -1, index === 0, "Move up"));
          reorder.appendChild(mkBtn("▼", 1, index === dayEntries.length - 1, "Move down"));
          row.appendChild(reorder);
        }

        dayBlock.appendChild(row);
      }

      container.appendChild(dayBlock);
    }
    cursor = addDays(cursor, 1);
  }

  if (!any) container.appendChild(el("div", "outline-empty", "No events yet."));
}

// Render the visual timeline (Whole-trip overview) into the print container.
// The timeline positions bars in pixels from the container width at render
// time, so we measure at a fixed landscape page width — re-rendering at the
// 552px screen width would leave it tiny in a corner of the page.
function buildPrintTimeline() {
  const root = document.getElementById("print-root");
  if (!root) return;
  root.innerHTML = "";
  printHeader(root);

  if (!state.start || !state.end || state.end < state.start) {
    root.appendChild(el("div", "outline-empty", "Set trip start and end dates to print a timeline."));
    return;
  }

  // Width that fills a landscape page (Letter content ≈ 980px at 96dpi); a bit
  // under to stay clear of margin variance across browsers/paper sizes.
  const PRINT_W = 940;
  const homeTz = state.homeTz || "America/Los_Angeles";
  const tzAware = state.tzAware !== false;
  const totalDays = dayDiff(state.start, state.end) + 1;
  const dayTzMap = computeDayTzMap(state.start, state.end, state.events, homeTz, tzAware);

  // Split a long trip into stacked rows (~2 weeks each) so the days stay wide
  // enough to read and the whole timeline fits on one landscape page.
  const MAX_DAYS_PER_ROW = 14;
  const rowCount = Math.ceil(totalDays / MAX_DAYS_PER_ROW);
  const daysPerRow = Math.ceil(totalDays / rowCount);

  // Make the container measurable at the target width for the (synchronous)
  // render; reset afterwards so @media print controls on-screen visibility.
  root.style.display = "block";
  root.style.width = PRINT_W + "px";

  for (let r = 0; r < rowCount; r++) {
    const rowStart = toISO(addDays(parseDay(state.start), r * daysPerRow));
    if (rowStart > state.end) break;
    const rowEndIso = toISO(addDays(parseDay(state.start), (r + 1) * daysPerRow - 1));
    const rowEnd = rowEndIso > state.end ? state.end : rowEndIso;

    const rowWrap = el("div", "print-tl-row");
    if (rowCount > 1) {
      rowWrap.appendChild(el("div", "print-tl-row-label",
        `${fmtShort(parseDay(rowStart))} – ${fmtShort(parseDay(rowEnd))}`));
    }
    const tl = el("div", "timeline overview");
    rowWrap.appendChild(tl);
    root.appendChild(rowWrap);
    renderTimeline(tl, rowStart, rowEnd, {
      dayTzMap, homeTz, dayPx: null, compact: false, tzAware, events: timelineEvents(),
    });
  }

  root.style.display = "";
  root.style.width = "";
}

// Set the page orientation/margins for the next print. Done in JS via the
// unnamed @page (named @page sizes aren't reliably honored by Chrome), so the
// timeline prints landscape and the text layouts print portrait.
function setPrintPage(orientation, marginCm) {
  let s = document.getElementById("print-page-style");
  if (!s) { s = document.createElement("style"); s.id = "print-page-style"; document.head.appendChild(s); }
  s.textContent = `@page { size: ${orientation}; margin: ${marginCm}cm; }`;
}

// Print layouts: "list" is the chronological day-by-day list; "bytype" groups
// items by category; both use the list (text) print styles in portrait.
// "timeline" prints the visual Whole-trip overview in landscape.
function printItinerary(mode) {
  document.body.classList.remove("print-mode-list", "print-mode-timeline");
  if (mode === "timeline") {
    buildPrintTimeline();
    document.body.classList.add("print-mode-timeline");
    setPrintPage("landscape", 1);
  } else if (mode === "bytype") {
    buildPrintByType();
    document.body.classList.add("print-mode-list");
    setPrintPage("portrait", 1.5);
  } else {
    buildPrintItinerary();
    document.body.classList.add("print-mode-list");
    setPrintPage("portrait", 1.5);
  }
  window.print();
}

// Ctrl+P / browser print (no mode chosen) defaults to the chronological list.
window.addEventListener("beforeprint", () => {
  if (!document.body.classList.contains("print-mode-list")
    && !document.body.classList.contains("print-mode-timeline")) {
    buildPrintItinerary();
    document.body.classList.add("print-mode-list");
    setPrintPage("portrait", 1.5);
  }
});
window.addEventListener("afterprint", () => {
  document.body.classList.remove("print-mode-list", "print-mode-timeline");
  // Drop the rendered timeline (with its live click handlers) once printed.
  const root = document.getElementById("print-root");
  if (root) { root.innerHTML = ""; root.style.display = ""; root.style.width = ""; }
});

(function wirePrintMenu() {
  const btn = document.getElementById("print-btn");
  const menu = document.getElementById("print-menu");
  if (!btn || !menu) return;
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const open = menu.hidden;
    menu.hidden = !open;
    btn.setAttribute("aria-expanded", String(open));
  });
  document.addEventListener("click", (e) => {
    if (menu.hidden) return;
    if (e.target.closest("#print-menu") || e.target.closest("#print-btn")) return;
    menu.hidden = true;
    btn.setAttribute("aria-expanded", "false");
  });
  const choose = (mode) => {
    menu.hidden = true;
    btn.setAttribute("aria-expanded", "false");
    printItinerary(mode);
  };
  document.getElementById("print-list")?.addEventListener("click", () => choose("list"));
  document.getElementById("print-bytype")?.addEventListener("click", () => choose("bytype"));
  document.getElementById("print-timeline")?.addEventListener("click", () => choose("timeline"));
})();

// --- Hotel comparison tab ---

function ensureHotelCompare() {
  if (!Array.isArray(state.hotelCompare)) state.hotelCompare = [];
  if (!Array.isArray(state.hotelGroups)) state.hotelGroups = [];
  // Backfill ratings/comments arrays + group + url fields on hotels added
  // before this feature.
  for (const h of state.hotelCompare) {
    if (!Array.isArray(h.ratings)) h.ratings = [];
    if (!Array.isArray(h.comments)) h.comments = [];
    if (h.compareOpen == null) h.compareOpen = false;
    if (!h.sourceUrl && h.url) h.sourceUrl = h.url;
  }
  // Make sure there's always at least one group, and every hotel belongs.
  const orphans = state.hotelCompare.filter(h => !h.groupId);
  if (orphans.length || (state.hotelCompare.length === 0 && state.hotelGroups.length === 0)) {
    let g = state.hotelGroups[0];
    if (!g) {
      g = { id: uid(), name: "Hotels" };
      state.hotelGroups.push(g);
    }
    orphans.forEach(h => { h.groupId = g.id; });
  }
}

function createHotelGroup(name) {
  ensureHotelCompare();
  const taken = new Set(state.hotelGroups.map(g => g.name));
  let n = 1;
  while (taken.has(`Group ${n}`)) n++;
  const g = { id: uid(), name: name || `Group ${n}` };
  state.hotelGroups.push(g);
  return g;
}

const DISPLAY_NAME_KEY = "trip-builder-display-name";
function getDisplayName() {
  let n = localStorage.getItem(DISPLAY_NAME_KEY);
  if (n && n.trim()) return n.trim();
  const entered = prompt("Your name (for comments + ratings on this device):");
  if (!entered || !entered.trim()) return null;
  n = entered.trim();
  localStorage.setItem(DISPLAY_NAME_KEY, n);
  return n;
}
function fmtStars(r) {
  if (r == null) return "—";
  const n = Math.round(Number(r));
  return "★".repeat(n) + "☆".repeat(Math.max(0, 5 - n));
}
function avgRating(ratings) {
  if (!ratings || ratings.length === 0) return null;
  const sum = ratings.reduce((s, r) => s + Number(r.rating || 0), 0);
  return sum / ratings.length;
}

const fmtHotelPrice = (n, ccy) => {
  if (n == null) return "—";
  const sym = ccy === "USD" ? "$" : ccy ? `${ccy} ` : "$";
  return `${sym}${Number(n).toLocaleString(undefined, { maximumFractionDigits: 2 })}`;
};

let activeHotelGroupId = null;

function hotelGroupTopAncestor(g) {
  let cur = g;
  while (cur && cur.parentId) cur = state.hotelGroups.find(x => x.id === cur.parentId);
  return cur || null;
}
function hotelGroupDescendants(rootId) {
  const out = new Set([rootId]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const g of state.hotelGroups) {
      if (g.parentId && out.has(g.parentId) && !out.has(g.id)) { out.add(g.id); changed = true; }
    }
  }
  return out;
}

function ensureRentalCompare() {
  if (!Array.isArray(state.rentalCompare)) state.rentalCompare = [];
}

const RENTAL_FIELDS = [
  { key: "company", label: "Company", placeholder: "Hertz / Enterprise / ..." },
  { key: "vehicle", label: "Vehicle", placeholder: "Compact / SUV / ..." },
  { key: "pickup", label: "Pickup", placeholder: "MCO airport" },
  { key: "pickupDate", label: "Pickup date", placeholder: "YYYY-MM-DD" },
  { key: "dropoff", label: "Return", placeholder: "MCO airport" },
  { key: "dropoffDate", label: "Return date", placeholder: "YYYY-MM-DD" },
  { key: "dailyRate", label: "Daily $", placeholder: "45" },
  { key: "total", label: "Total $", placeholder: "315" },
  { key: "notes", label: "Notes", placeholder: "" },
];

function renderRentalCompare() {
  ensureRentalCompare();
  const wrap = document.getElementById("rental-table");
  if (!wrap) return;
  wrap.innerHTML = "";
  if (state.rentalCompare.length === 0) {
    wrap.innerHTML = `<div class="empty-state">No rentals yet — click "+ Add rental" above to start comparing.</div>`;
    return;
  }
  const tbl = document.createElement("table");
  tbl.className = "compare-table";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  for (const f of RENTAL_FIELDS) {
    const th = document.createElement("th");
    th.textContent = f.label;
    headRow.appendChild(th);
  }
  headRow.appendChild(document.createElement("th"));
  thead.appendChild(headRow);
  tbl.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const r of state.rentalCompare) {
    const tr = document.createElement("tr");
    for (const f of RENTAL_FIELDS) {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.type = (f.key.endsWith("Date")) ? "date" : "text";
      input.value = r[f.key] || "";
      input.placeholder = f.placeholder;
      input.addEventListener("change", () => {
        r[f.key] = input.value;
        save();
      });
      td.appendChild(input);
      tr.appendChild(td);
    }
    const td = document.createElement("td");
    const del = document.createElement("button");
    del.type = "button";
    del.className = "ghost";
    del.textContent = "×";
    del.title = "Remove";
    del.addEventListener("click", () => {
      if (!confirm("Remove this rental from the comparison?")) return;
      state.rentalCompare = state.rentalCompare.filter(x => x.id !== r.id);
      save();
      renderRentalCompare();
    });
    td.appendChild(del);
    tr.appendChild(td);
    tbody.appendChild(tr);
  }
  tbl.appendChild(tbody);
  wrap.appendChild(tbl);
}

document.getElementById("rental-add-btn")?.addEventListener("click", () => {
  ensureRentalCompare();
  state.rentalCompare.push({ id: uid() });
  save();
  renderRentalCompare();
});

document.getElementById("enable-tab-rental")?.addEventListener("change", (e) => {
  toggleEnabledTab("rental", e.target.checked);
});

function renderHotelCompare() {
  ensureHotelCompare();
  const wrap = document.getElementById("compare-table");
  if (!wrap) return;
  wrap.innerHTML = "";

  const topGroups = state.hotelGroups.filter(g => !g.parentId);
  let activeGroup = state.hotelGroups.find(g => g.id === activeHotelGroupId);
  if (!activeGroup) {
    activeGroup = topGroups[0] || null;
    activeHotelGroupId = activeGroup?.id || null;
  }
  const activeTop = activeGroup ? hotelGroupTopAncestor(activeGroup) : null;

  // Top-level sub-nav (one tab per top-level group).
  const subnav = document.createElement("div");
  subnav.className = "hotel-subnav";
  for (const g of topGroups) {
    const tab = document.createElement("button");
    tab.type = "button";
    tab.className = "hotel-subnav-tab" + (g.id === activeTop?.id ? " active" : "");
    tab.textContent = g.name || "(unnamed)";
    tab.addEventListener("click", () => {
      activeHotelGroupId = g.id;
      renderHotelCompare();
    });
    subnav.appendChild(tab);
  }
  // "+ Add group" with an inline menu containing top-level and subgroup options.
  const addWrap = document.createElement("div");
  addWrap.className = "hotel-subnav-add-wrap";
  const addTab = document.createElement("button");
  addTab.type = "button";
  addTab.className = "hotel-subnav-add";
  addTab.textContent = "+ Add group ▾";
  const menu = document.createElement("div");
  menu.className = "hotel-add-menu";
  menu.hidden = true;
  const optTop = document.createElement("button");
  optTop.type = "button";
  optTop.className = "hotel-add-menu-item";
  optTop.textContent = "Add group";
  optTop.addEventListener("click", () => {
    const g = createHotelGroup();
    activeHotelGroupId = g.id;
    save();
    renderHotelCompare();
  });
  menu.appendChild(optTop);
  if (activeTop) {
    const optSub = document.createElement("button");
    optSub.type = "button";
    optSub.className = "hotel-add-menu-item";
    optSub.textContent = `Add subgroup under "${activeTop.name || "group"}"`;
    optSub.addEventListener("click", () => {
      const g = createHotelGroup();
      g.parentId = activeTop.id;
      activeHotelGroupId = g.id;
      save();
      renderHotelCompare();
    });
    menu.appendChild(optSub);
  }
  addTab.addEventListener("click", (e) => {
    e.stopPropagation();
    menu.hidden = !menu.hidden;
    addWrap.classList.toggle("menu-open", !menu.hidden);
    if (!menu.hidden) {
      const close = (ev) => {
        if (!addWrap.contains(ev.target)) {
          menu.hidden = true;
          addWrap.classList.remove("menu-open");
          document.removeEventListener("click", close);
        }
      };
      setTimeout(() => document.addEventListener("click", close));
    }
  });
  addWrap.appendChild(addTab);
  addWrap.appendChild(menu);
  subnav.appendChild(addWrap);
  const hotSlot = document.getElementById("compare-subnav-slot");
  if (hotSlot) { hotSlot.innerHTML = ""; hotSlot.appendChild(subnav); }
  else wrap.appendChild(subnav);

  // Sub-sub-nav for the active top group's children (if any) — also shown
  // when a child is active so the parent's siblings stay reachable.
  // Sub-sub-nav: only render when the active top group already has subgroups.
  if (activeTop) {
    const children = state.hotelGroups.filter(g => g.parentId === activeTop.id);
    if (children.length) {
      const subsub = document.createElement("div");
      subsub.className = "hotel-subsubnav";
      const allTab = document.createElement("button");
      allTab.type = "button";
      allTab.className = "hotel-subsubnav-tab" + (activeGroup?.id === activeTop.id ? " active" : "");
      allTab.textContent = "All";
      allTab.addEventListener("click", () => {
        activeHotelGroupId = activeTop.id;
        renderHotelCompare();
      });
      subsub.appendChild(allTab);
      for (const c of children) {
        const tab = document.createElement("button");
        tab.type = "button";
        tab.className = "hotel-subsubnav-tab" + (c.id === activeGroup?.id ? " active" : "");
        tab.textContent = c.name || "(unnamed)";
        tab.addEventListener("click", () => {
          activeHotelGroupId = c.id;
          renderHotelCompare();
        });
        subsub.appendChild(tab);
      }
      if (hotSlot) hotSlot.appendChild(subsub);
      else wrap.appendChild(subsub);
    }
  }

  // Render the active group's content. If the active group is a top-level
  // group with children, include hotels in any descendant.
  const targetIds = activeGroup ? hotelGroupDescendants(activeGroup.id) : new Set();
  for (const group of state.hotelGroups.filter(g => g.id === activeHotelGroupId)) {
    const groupEl = document.createElement("div");
    groupEl.className = "hotel-group";

    const headRow = document.createElement("div");
    headRow.className = "hotel-group-head";
    const nameInput = document.createElement("input");
    nameInput.className = "hotel-group-name";
    nameInput.value = group.name;
    nameInput.placeholder = "Group name";
    nameInput.addEventListener("input", () => { group.name = nameInput.value; save(); });
    headRow.appendChild(nameInput);

    const headActions = document.createElement("div");
    headActions.className = "hotel-group-actions";
    const addHotelBtn = document.createElement("button");
    addHotelBtn.type = "button";
    addHotelBtn.textContent = "+ Add hotel";
    addHotelBtn.addEventListener("click", () => {
      state.hotelCompare.push({ id: uid(), groupId: group.id, name: "New hotel", ratings: [], comments: [], compareOpen: false });
      save();
      renderHotelCompare();
    });
    headActions.appendChild(addHotelBtn);
    const delGroupBtn = document.createElement("button");
    delGroupBtn.type = "button";
    delGroupBtn.className = "ghost";
    delGroupBtn.textContent = "Delete group";
    delGroupBtn.addEventListener("click", () => {
      const subtree = hotelGroupDescendants(group.id);
      const inGroup = state.hotelCompare.filter(h => subtree.has(h.groupId));
      if ((inGroup.length || subtree.size > 1) && !confirm(`Delete group "${group.name}"${subtree.size > 1 ? " and its subgroups" : ""}${inGroup.length ? ` (${inGroup.length} hotel${inGroup.length === 1 ? "" : "s"})` : ""}?`)) return;
      state.hotelCompare = state.hotelCompare.filter(h => !subtree.has(h.groupId));
      state.hotelGroups = state.hotelGroups.filter(g => !subtree.has(g.id));
      save();
      renderHotelCompare();
    });
    headActions.appendChild(delGroupBtn);
    headRow.appendChild(headActions);
    groupEl.appendChild(headRow);

    const hotelsInGroup = state.hotelCompare.filter(h => targetIds.has(h.groupId));

    // Compact card row at the top of the group.
    const cardRow = document.createElement("div");
    cardRow.className = "hotel-card-row";
    if (hotelsInGroup.length === 0) {
      const empty = document.createElement("div");
      empty.className = "compare-empty";
      empty.textContent = "No hotels yet — paste one above or click + Add hotel.";
      cardRow.appendChild(empty);
    }
    for (const h of hotelsInGroup) {
      cardRow.appendChild(renderHotelCard(h));
    }
    groupEl.appendChild(cardRow);

    // Detailed comparison table — only includes hotels with compareOpen = true.
    const selected = hotelsInGroup.filter(h => h.compareOpen);
    if (selected.length) {
      groupEl.appendChild(renderHotelTable(selected));
    }
    wrap.appendChild(groupEl);
  }

  if (!state.hotelGroups.length) {
    const empty = document.createElement("div");
    empty.className = "compare-empty";
    empty.textContent = "No hotel groups yet — click + Add group above.";
    wrap.appendChild(empty);
  }
}

function renderHotelCard(h) {
  const card = document.createElement("div");
  card.className = "hotel-card"
    + (h.compareOpen ? " compare-on" : "")
    + (h.cardOpen ? " card-open" : "");

  // One-line summary row — click to expand/collapse details below.
  const top = document.createElement("div");
  top.className = "hotel-card-top";
  const chev = document.createElement("span");
  chev.className = "hotel-card-chev";
  chev.textContent = h.cardOpen ? "▾" : "▸";
  top.appendChild(chev);
  const nameLabel = document.createElement("span");
  nameLabel.className = "hotel-card-name-label";
  nameLabel.textContent = h.name || "(unnamed hotel)";
  top.appendChild(nameLabel);
  const priceWrap = document.createElement("span");
  priceWrap.className = "hotel-card-price";
  priceWrap.textContent = fmtHotelPrice(h.totalPrice, h.currency);
  if (h.totalPrice != null) {
    const per = document.createElement("span");
    per.className = "per-night";
    per.textContent = "total";
    priceWrap.appendChild(per);
  }
  top.appendChild(priceWrap);
  top.addEventListener("click", () => {
    h.cardOpen = !h.cardOpen;
    save();
    renderHotelCompare();
  });
  card.appendChild(top);

  if (!h.cardOpen) return card;

  // Expanded details below.
  const nameInput = document.createElement("input");
  nameInput.className = "hotel-card-name";
  nameInput.value = h.name || "";
  nameInput.placeholder = "Hotel name";
  nameInput.addEventListener("input", () => { h.name = nameInput.value; save(); });
  card.appendChild(nameInput);

  const makeField = (label, key, parser) => {
    const row = document.createElement("div");
    row.className = "hotel-card-price-edit";
    const lbl = document.createElement("span");
    lbl.className = "url-label";
    lbl.textContent = label;
    row.appendChild(lbl);
    const input = document.createElement("input");
    input.type = "text";
    input.className = "url-input";
    input.value = h[key] != null ? String(h[key]) : "";
    input.addEventListener("input", () => {
      h[key] = parser(input.value);
      save();
    });
    input.addEventListener("change", renderHotelCompare);
    row.appendChild(input);
    return row;
  };
  const num = v => { const s = v.replace(/[^0-9.]/g, ""); return s === "" ? null : Number(s); };
  const int = v => { const s = v.replace(/[^0-9]/g, ""); return s === "" ? null : parseInt(s, 10); };
  const str = v => v.trim() || null;
  card.appendChild(makeField("$ / night", "pricePerNight", num));
  card.appendChild(makeField("Total $", "totalPrice", num));
  card.appendChild(makeField("# rooms", "roomCount", int));
  card.appendChild(makeField("Room type", "roomType", str));

  // Comment line (single inline note, separate from the multi-author comments table)
  const noteInput = document.createElement("input");
  noteInput.className = "hotel-card-note";
  noteInput.value = h.note || "";
  noteInput.placeholder = "Quick comment / note";
  noteInput.addEventListener("input", () => { h.note = noteInput.value; save(); });
  card.appendChild(noteInput);

  // Editable URL fields.
  const urls = document.createElement("div");
  urls.className = "hotel-card-urls";
  const makeUrlField = (label, key) => {
    const wrap = document.createElement("div");
    wrap.className = "hotel-card-url-row";
    const lbl = document.createElement("span");
    lbl.className = "url-label";
    lbl.textContent = label;
    wrap.appendChild(lbl);
    const input = document.createElement("input");
    input.className = "url-input";
    input.type = "url";
    input.value = h[key] || "";
    input.placeholder = "https://…";
    input.addEventListener("input", () => { h[key] = input.value || null; save(); refreshUrlOpen(); });
    wrap.appendChild(input);
    const open = document.createElement("a");
    open.className = "url-open";
    open.target = "_blank";
    open.rel = "noopener";
    open.textContent = "open";
    const refreshUrlOpen = () => {
      if (h[key]) { open.href = h[key]; open.style.visibility = "visible"; }
      else { open.removeAttribute("href"); open.style.visibility = "hidden"; }
    };
    refreshUrlOpen();
    wrap.appendChild(open);
    return wrap;
  };
  urls.appendChild(makeUrlField("source", "sourceUrl"));
  urls.appendChild(makeUrlField("website", "websiteUrl"));
  card.appendChild(urls);

  const actions = document.createElement("div");
  actions.className = "hotel-card-actions";
  const compareBtn = document.createElement("button");
  compareBtn.type = "button";
  compareBtn.className = h.compareOpen ? "primary" : "";
  compareBtn.textContent = h.compareOpen ? "Hide details" : "Compare";
  compareBtn.addEventListener("click", () => {
    h.compareOpen = !h.compareOpen;
    save();
    renderHotelCompare();
  });
  actions.appendChild(compareBtn);
  // Move to a different group.
  if (state.hotelGroups.length > 1) {
    const moveSel = document.createElement("select");
    moveSel.className = "hotel-card-move";
    moveSel.title = "Move to group";
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Move to…";
    moveSel.appendChild(placeholder);
    for (const g of state.hotelGroups) {
      if (g.id === h.groupId) continue;
      const opt = document.createElement("option");
      opt.value = g.id;
      if (g.parentId) {
        const parent = state.hotelGroups.find(p => p.id === g.parentId);
        opt.textContent = `${parent?.name || "?"} › ${g.name || "(unnamed)"}`;
      } else {
        opt.textContent = g.name || "(unnamed)";
      }
      moveSel.appendChild(opt);
    }
    moveSel.addEventListener("change", () => {
      if (!moveSel.value) return;
      h.groupId = moveSel.value;
      save();
      renderHotelCompare();
    });
    actions.appendChild(moveSel);
  }
  const del = document.createElement("button");
  del.type = "button";
  del.className = "ghost";
  del.textContent = "×";
  del.title = "Remove";
  del.addEventListener("click", () => {
    state.hotelCompare = state.hotelCompare.filter(x => x.id !== h.id);
    save();
    renderHotelCompare();
  });
  actions.appendChild(del);
  card.appendChild(actions);
  return card;
}

function renderHotelTable(hotels) {
  const fmtPrice = fmtHotelPrice;
  const ROWS = [
    ["Platform",     h => h.platform || "—"],
    ["Check in",     h => h.checkIn  || "—"],
    ["Check out",    h => h.checkOut || "—"],
    ["Nights",       h => h.nights != null ? h.nights : "—"],
    ["Per night",    h => fmtPrice(h.pricePerNight, h.currency)],
    ["Total",        h => fmtPrice(h.totalPrice, h.currency)],
    ["# rooms",      h => h.roomCount != null ? h.roomCount : "—"],
    ["Room type",    h => h.roomType || "—"],
    ["Neighborhood", h => h.neighborhood || "—"],
    ["Listed rating",h => h.rating || "—"],
    ["Cancellation", h => h.cancellation || "—"],
    ["Amenities",    h => Array.isArray(h.amenities) && h.amenities.length ? h.amenities.join(", ") : "—"],
    ["Notes",        h => h.notes || "—"],
    ["Source",       h => h.sourceUrl ? `<a href="${escHtml(h.sourceUrl)}" target="_blank" rel="noopener">open</a>` : "—"],
    ["Website",      h => h.websiteUrl ? `<a href="${escHtml(h.websiteUrl)}" target="_blank" rel="noopener">open</a>` : "—"],
  ];

  const tbl = document.createElement("table");
  tbl.className = "compare-table";
  const head = document.createElement("tr");
  head.appendChild(document.createElement("th"));
  for (const h of hotels) {
    const th = document.createElement("th");
    const name = document.createElement("div");
    name.className = "compare-name";
    name.textContent = h.name || "Untitled";
    const del = document.createElement("button");
    del.type = "button";
    del.className = "compare-del";
    del.textContent = "×";
    del.title = "Remove";
    del.addEventListener("click", () => {
      state.hotelCompare = state.hotelCompare.filter(x => x.id !== h.id);
      save();
      renderHotelCompare();
    });
    th.appendChild(name);
    th.appendChild(del);
    head.appendChild(th);
  }
  tbl.appendChild(head);
  for (const [label, getter] of ROWS) {
    const tr = document.createElement("tr");
    const lblTh = document.createElement("th");
    lblTh.className = "compare-label";
    lblTh.textContent = label;
    tr.appendChild(lblTh);
    for (const h of hotels) {
      const td = document.createElement("td");
      const v = getter(h);
      if ((label === "Source" && h.sourceUrl) || (label === "Website" && h.websiteUrl)) td.innerHTML = v;
      else td.textContent = v;
      tr.appendChild(td);
    }
    tbl.appendChild(tr);
  }

  // Group ratings row.
  const ratingsTr = document.createElement("tr");
  const ratingsLabel = document.createElement("th");
  ratingsLabel.className = "compare-label";
  ratingsLabel.textContent = "Group ratings";
  ratingsTr.appendChild(ratingsLabel);
  for (const h of hotels) {
    const td = document.createElement("td");
    td.className = "compare-ratings";
    const avg = avgRating(h.ratings);
    if (avg != null) {
      const summary = document.createElement("div");
      summary.className = "compare-rating-summary";
      summary.textContent = `${fmtStars(avg)}  (${avg.toFixed(1)})`;
      td.appendChild(summary);
    }
    for (const r of h.ratings) {
      const row = document.createElement("div");
      row.className = "compare-rating-row";
      const author = document.createElement("span");
      author.className = "compare-author";
      author.textContent = r.author;
      const stars = document.createElement("span");
      stars.textContent = ` ${fmtStars(r.rating)}`;
      const del = document.createElement("button");
      del.type = "button";
      del.className = "compare-mini-del";
      del.textContent = "×";
      del.title = "Remove";
      del.addEventListener("click", () => {
        h.ratings = h.ratings.filter(x => x.id !== r.id);
        save();
        renderHotelCompare();
      });
      row.appendChild(author);
      row.appendChild(stars);
      row.appendChild(del);
      td.appendChild(row);
    }
    const addBtn = document.createElement("button");
    addBtn.type = "button";
    addBtn.className = "compare-mini-add";
    addBtn.textContent = "+ Add rating";
    addBtn.addEventListener("click", () => {
      const author = getDisplayName();
      if (!author) return;
      const raw = prompt(`Rating from ${author} (1–5):`, "4");
      if (raw == null) return;
      const n = parseInt(raw, 10);
      if (!isFinite(n) || n < 1 || n > 5) { alert("Enter a number 1–5."); return; }
      h.ratings.push({ id: uid(), author, rating: n, when: Date.now() });
      save();
      renderHotelCompare();
    });
    td.appendChild(addBtn);
    ratingsTr.appendChild(td);
  }
  tbl.appendChild(ratingsTr);

  // Comments row.
  const commentsTr = document.createElement("tr");
  const commentsLabel = document.createElement("th");
  commentsLabel.className = "compare-label";
  commentsLabel.textContent = "Comments";
  commentsTr.appendChild(commentsLabel);
  for (const h of hotels) {
    const td = document.createElement("td");
    td.className = "compare-comments";
    for (const c of h.comments) {
      const row = document.createElement("div");
      row.className = "compare-comment-row";
      const author = document.createElement("span");
      author.className = "compare-author";
      author.textContent = c.author + ":";
      const text = document.createElement("span");
      text.textContent = " " + c.text;
      const del = document.createElement("button");
      del.type = "button";
      del.className = "compare-mini-del";
      del.textContent = "×";
      del.title = "Remove";
      del.addEventListener("click", () => {
        h.comments = h.comments.filter(x => x.id !== c.id);
        save();
        renderHotelCompare();
      });
      row.appendChild(author);
      row.appendChild(text);
      row.appendChild(del);
      td.appendChild(row);
    }
    const addBtn = document.createElement("button");
    addBtn.type = "button";
    addBtn.className = "compare-mini-add";
    addBtn.textContent = "+ Add comment";
    addBtn.addEventListener("click", () => {
      const author = getDisplayName();
      if (!author) return;
      const text = prompt(`Comment from ${author}:`);
      if (!text || !text.trim()) return;
      h.comments.push({ id: uid(), author, text: text.trim(), when: Date.now() });
      save();
      renderHotelCompare();
    });
    td.appendChild(addBtn);
    commentsTr.appendChild(td);
  }
  tbl.appendChild(commentsTr);

  return tbl;
}

async function smartParseHotelRemote(input, statusEl) {
  await whenFb();
  if (!window.fb.user) {
    statusEl.textContent = "Sign in to use smart parse.";
    statusEl.className = "paste-status error";
    return null;
  }
  const trimmed = input.trim();
  const urlOnly = /^https?:\/\/\S+$/i.test(trimmed) && !/\s/.test(trimmed);
  const body = urlOnly
    ? { url: trimmed, mode: "hotel-from-url", tripStart: state.start || null, tripEnd: state.end || null }
    : { text: input, mode: "hotel-compare", tripStart: state.start || null, tripEnd: state.end || null };
  statusEl.textContent = urlOnly ? `Fetching ${new URL(trimmed).hostname}…` : "Parsing hotel with Claude…";
  statusEl.className = "paste-status";
  try {
    return await window.fb.smartParse(body);
  } catch (e) {
    statusEl.textContent = `Smart parse failed: ${e.message}`;
    statusEl.className = "paste-status error";
    return null;
  }
}

document.getElementById("compare-smart")?.addEventListener("click", async () => {
  const input = document.getElementById("compare-paste-input");
  const status = document.getElementById("compare-paste-status");
  const text = input.value;
  if (!text.trim()) {
    status.textContent = "Paste a hotel listing or URL first.";
    status.className = "paste-status error";
    return;
  }
  // Detect a list of URLs, one per line — process each as its own hotel.
  // If the input has any non-URL text mixed in, fall through to single-paste
  // mode (Claude treats the whole thing as one listing).
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const allUrls = lines.length > 1 && lines.every(l => /^https?:\/\/\S+$/i.test(l));
  const btn = document.getElementById("compare-smart");
  btn.disabled = true;
  ensureHotelCompare();
  if (allUrls) {
    let added = 0; let failed = 0;
    for (let i = 0; i < lines.length; i++) {
      status.textContent = `Fetching ${i + 1}/${lines.length}: ${new URL(lines[i]).hostname}…`;
      status.className = "paste-status";
      const data = await smartParseHotelRemote(lines[i], status);
      if (data) {
        state.hotelCompare.push({
          id: uid(),
          groupId: activeHotelGroupId || state.hotelGroups[0]?.id,
          ...data,
          sourceUrl: data.sourceUrl || data.url || null,
          websiteUrl: data.websiteUrl || data.website || null,
        });
        save();
        renderHotelCompare();
        added++;
      } else {
        failed++;
      }
    }
    btn.disabled = false;
    status.textContent = failed === 0
      ? `Added ${added} hotels to comparison.`
      : `Added ${added}; ${failed} URL(s) failed (likely bot-blocked).`;
    status.className = failed === 0 ? "paste-status success" : "paste-status";
    input.value = "";
    return;
  }
  const data = await smartParseHotelRemote(text, status);
  btn.disabled = false;
  if (!data) return;
  state.hotelCompare.push({
    id: uid(),
    groupId: activeHotelGroupId || state.hotelGroups[0]?.id,
    ...data,
    sourceUrl: data.sourceUrl || data.url || null,
    websiteUrl: data.websiteUrl || data.website || null,
  });
  save();
  renderHotelCompare();
  status.textContent = `Added ${data.name || "hotel"} to comparison.`;
  status.className = "paste-status success";
  input.value = "";
  renderHotelCompare();
});

document.getElementById("compare-paste-clear")?.addEventListener("click", () => {
  const input = document.getElementById("compare-paste-input");
  const status = document.getElementById("compare-paste-status");
  if (input) input.value = "";
  if (status) { status.textContent = ""; status.className = "paste-status"; }
});

// --- flight paste/parser ---

// City name → IATA code, for parsers that get city names instead of codes.
const CITY_TO_CODE = {
  SEATTLE: "SEA", TAMPA: "TPA", MIAMI: "MIA", ATLANTA: "ATL",
  BOSTON: "BOS", DENVER: "DEN", PORTLAND: "PDX", CHICAGO: "ORD",
  HOUSTON: "IAH", DALLAS: "DFW", PHOENIX: "PHX", DETROIT: "DTW",
  MINNEAPOLIS: "MSP", PHILADELPHIA: "PHL", ORLANDO: "MCO",
  HONOLULU: "HNL", ANCHORAGE: "ANC", VANCOUVER: "YVR",
  NEWYORK: "JFK", LASVEGAS: "LAS", LOSANGELES: "LAX",
  SANFRANCISCO: "SFO", SANDIEGO: "SAN", SALTLAKECITY: "SLC",
  WASHINGTON: "IAD", NEWORLEANS: "MSY", FORTLAUDERDALE: "FLL",
  CHARLOTTE: "CLT", AUSTIN: "AUS", LONDON: "LHR", PARIS: "CDG",
  FRANKFURT: "FRA", AMSTERDAM: "AMS", TOKYO: "NRT", ZURICH: "ZRH",
  DUBLIN: "DUB", DUBAI: "DXB", DOHA: "DOH", ISTANBUL: "IST",
};
const AIRLINE_CODE = {
  DELTA: "DL", UNITED: "UA", AMERICAN: "AA", SOUTHWEST: "WN",
  ALASKA: "AS", JETBLUE: "B6", SPIRIT: "NK", FRONTIER: "F9",
  HAWAIIAN: "HA", AIRCANADA: "AC", LUFTHANSA: "LH",
  BRITISHAIRWAYS: "BA", AIRFRANCE: "AF", KLM: "KL",
};
function cityToCode(name) {
  const k = name.replace(/\s+/g, "").toUpperCase();
  return CITY_TO_CODE[k] || k.slice(0, 3);
}

function to24h(hhmm, ampm) {
  let [h, m] = hhmm.split(":").map(Number);
  if (ampm === "PM" && h < 12) h += 12;
  if (ampm === "AM" && h === 12) h = 0;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

const MONTHS = { Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6, Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12 };

// Parse a natural-language command like:
//   "add hotel placeholder on july 7 for znz hotel"
//   "add activity on jul 4 hiking"
//   "add lodging on july 7 to july 8 at riu palace"
//   "add location on jul 5 cape town"
function parseCommand(text, defaultYear) {
  const trimmed = text.trim();
  if (!trimmed) return null;
  if (!/^add\b/i.test(trimmed)) return null;

  const m = trimmed.match(/^add\s+(?:(hotel|hotels|lodging|flight|flights|activity|activities|location|where|cruise|rental|car)\s+)?(?:placeholder\s+)?(?:on\s+)?(.+)$/i);
  if (!m) return null;

  const laneWord = (m[1] || "").toLowerCase();
  const rest = m[2];

  const laneMap = {
    hotel: "lodging", hotels: "lodging", lodging: "lodging",
    flight: "flights", flights: "flights",
    activity: "activities", activities: "activities",
    location: "location", where: "location",
    cruise: "lodging",
    rental: "rental", car: "rental",
  };
  // Detect lane from keyword OR from words anywhere in the rest of the
  // command (e.g. "add disney cruise from FLL on Dec 27-Jan 3" → lodging).
  let lane = laneMap[laneWord];
  if (!lane) {
    if (/\b(rental car|car rental|rental)\b/i.test(rest)) lane = "rental";
    else if (/\bcruise\b/i.test(rest)) lane = "lodging";
    else if (/\b(hotel|resort|villa|lodge|airbnb|vrbo)\b/i.test(rest)) lane = "lodging";
    else if (/\bflight\b/i.test(rest)) lane = "flights";
    else lane = "activities";
  }

  const range = parseRangeFromText(rest, defaultYear);
  if (!range) return null;

  let title = range.remainder.replace(/^(?:for|at)\s+/i, "").trim();
  if (!title) title = `${lane} placeholder`;

  const colorMap = {
    lodging: "amber",
    flights: "indigo",
    location: "violet",
    activities: "emerald",
    rental: "orange",
  };

  return [{
    id: uid(),
    title,
    lane,
    color: colorMap[lane] || "emerald",
    start: range.start,
    end: range.end || range.start,
    notes: "Added via paste command",
  }];
}

// Parse hotel/reservation confirmations that have labeled date fields:
// "Arrive: Saturday, Dec 19, 2026" / "Depart: Saturday, Dec 26, 2026" / "Check-in" /
// "Check-out". Used for Disney/Marriott/Airbnb/etc. confirmation pastes.
function parseReservation(text, defaultYear) {
  const monthMap = {
    jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12,
    january:1, february:2, march:3, april:4, june:6, july:7, august:8, september:9, october:10, november:11, december:12,
  };
  const dowDateRx = /(?:(?:sun|mon|tue|wed|thu|fri|sat)[a-z]*[,\s]+)?([a-z]+)\.?\s+(\d{1,2})(?:,?\s+(\d{4}))?/i;

  function dateAfter(labels) {
    const lower = text.toLowerCase();
    for (const label of labels) {
      const idx = lower.indexOf(label.toLowerCase());
      if (idx < 0) continue;
      const after = text.slice(idx + label.length);
      const m = after.match(dowDateRx);
      if (!m) continue;
      const mo = monthMap[m[1].toLowerCase()];
      if (!mo) continue;
      const y = m[3] ? +m[3] : defaultYear;
      return `${y}-${String(mo).padStart(2,"0")}-${String(+m[2]).padStart(2,"0")}`;
    }
    return null;
  }

  const start = dateAfter(["Arrive:", "Arrival:", "Check-in", "Check in"]);
  const end   = dateAfter(["Depart:", "Departure:", "Check-out", "Check out"]);
  // Require an Arrive/Check-in label so this doesn't fire on flight pastes
  // or random text that happens to contain a date.
  if (!start) return null;

  const isLodging = /\b(hotel|resort|villa|inn|suite|lodge|bnb|airbnb|vrbo)\b/i.test(text);
  const lane = isLodging ? "lodging" : (end && end !== start ? "lodging" : "activities");

  // Extract a venue title. First preference: the line right after a bare
  // "Hotel" / "Resort" / "Property" header.
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  let title = null;
  for (let i = 0; i < lines.length - 1; i++) {
    if (/^(hotel|resort|property|accommodation)s?\s*$/i.test(lines[i])) {
      title = lines[i + 1];
      break;
    }
  }
  // Second preference: a line that contains a known hotel brand (Marriott
  // confirmations have the venue line — e.g. "Courtyard Mexico City Airport"
  // — buried under nav links like "ENHANCE YOUR STAY"; brand-matching skips
  // past those).
  if (!title) {
    const lowerLines = lines.map(l => l.toLowerCase());
    for (let i = 0; i < lines.length; i++) {
      const ll = lowerLines[i];
      if (SMART_HOTEL_KEYWORDS.some(k => ll.includes(k))
          && lines[i].length < 80
          && !/^(check|confirmation|guest|address|phone|total|room|reservation)/i.test(lines[i])) {
        title = lines[i];
        break;
      }
    }
  }
  // Fallback: first non-label, non-date, non-nav line. We skip short ALL-CAPS
  // lines because hotel emails open with nav links like "ENHANCE YOUR STAY",
  // "SUMMARY OF CHARGES", "CONTACT US" that would otherwise win.
  if (!title) {
    for (const line of lines) {
      if (/^(date|confirmation|arrive|arrival|depart|departure|guests?|hotel|address|check[\s-]?in|check[\s-]?out|reservation)/i.test(line)) continue;
      if (dowDateRx.test(line) && line.length < 40) continue;
      if (/^\d/.test(line)) continue;
      if (line.length > 80) continue;
      // Skip short all-caps nav crumbs.
      if (line === line.toUpperCase() && line.length < 30 && /^[A-Z][A-Z &|]+[A-Z]$/.test(line)) continue;
      title = line;
      break;
    }
  }
  if (!title) title = isLodging ? "Hotel reservation" : "Reservation";

  const colorMap = { lodging: "amber", flights: "indigo", activities: "emerald", location: "violet" };
  return [{
    id: uid(),
    title,
    lane,
    color: colorMap[lane] || "amber",
    start,
    end: end || start,
    notes: "Added via reservation paste",
  }];
}

// Parse cruise-line confirmations (Disney, Royal Caribbean, etc.) where the
// dates are labeled "Departure Date" / "Return Date" or "Embarkation Date" /
// "Disembarkation Date" rather than the hotel "Arrive/Depart" labels.
function parseCruise(text, defaultYear) {
  if (!/\bcruise\b/i.test(text)) return null;
  const monthMap = {
    jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12,
    january:1, february:2, march:3, april:4, june:6, july:7, august:8, september:9, october:10, november:11, december:12,
  };
  const dowDateRx = /(?:(?:sun|mon|tue|wed|thu|fri|sat)[a-z]*[,\s]+)?([a-z]+)\.?\s+(\d{1,2})(?:,?\s+(\d{4}))?/i;

  function dateAfter(labels) {
    const lower = text.toLowerCase();
    for (const label of labels) {
      const idx = lower.indexOf(label.toLowerCase());
      if (idx < 0) continue;
      const after = text.slice(idx + label.length);
      const m = after.match(dowDateRx);
      if (!m) continue;
      const mo = monthMap[m[1].toLowerCase()];
      if (!mo) continue;
      const y = m[3] ? +m[3] : defaultYear;
      return `${y}-${String(mo).padStart(2,"0")}-${String(+m[2]).padStart(2,"0")}`;
    }
    return null;
  }

  const start = dateAfter(["Departure Date", "Embarkation Date", "Sail Date", "Sailing Date"]);
  const end   = dateAfter(["Return Date", "Disembarkation Date", "Arrival Date"]);
  if (!start || !end) return null;

  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  let title = lines.find(l => /^\d+\s+night\s+cruise/i.test(l));
  if (!title) title = lines.find(l => /cruise/i.test(l) && !/:/.test(l) && l.length < 80);
  if (!title) title = "Cruise";

  return [{
    id: uid(),
    title,
    lane: "lodging",
    color: "amber",
    start,
    end,
    notes: "Added via cruise paste",
  }];
}

// --- Smart fallback parser ---
// Scans free-form text for dates, times, and category-defining keywords
// (airlines, hotels, cruise lines, rental car brands, airport codes), and
// builds a best-effort event. Runs as the very last fallback after every
// strict parser fails to match.

const SMART_HOTEL_KEYWORDS = [
  "hotel","motel","resort","villa","villas","inn","suite","suites","lodge","bnb","airbnb","vrbo",
  "polynesian","hilton","marriott","hyatt","sheraton","westin","four seasons","ritz","wyndham",
  "best western","embassy","hampton","courtyard","residence","doubletree","holiday inn","crowne plaza",
  "ramada","quality inn","comfort inn","la quinta","fairfield","candlewood","staybridge","aloft",
  "moxy","ac hotel","element","disney's"
];
const SMART_AIRLINE_KEYWORDS = [
  "delta","united","american","southwest","alaska","jetblue","spirit","frontier","hawaiian",
  "air canada","lufthansa","british airways","air france","klm","emirates","qatar","etihad","singapore",
  "cathay","kenya airways","discover airlines","auric air","ryanair","easyjet","iberia","turkish",
  "saudia","ethiopian","virgin","westjet","aeromexico","copa","avianca","latam","ana","jal","korean",
  "asiana","thai","philippine","air india","aeroflot","sas","finnair","tap","swiss","austrian","brussels"
];
const SMART_CRUISE_KEYWORDS = [
  "cruise","disney cruise","royal caribbean","carnival","norwegian","princess","holland america",
  "msc","celebrity","cunard","viking","oceania","seabourn","silversea","azamara","disney dream",
  "disney magic","disney wish","disney wonder","embarkation","disembarkation"
];
const SMART_RENTAL_KEYWORDS = [
  "rental car","car rental","hertz","avis","enterprise","sixt","budget rent","alamo","national car",
  "thrifty","dollar rent","europcar","fox rent","payless"
];
const SMART_FLIGHT_KEYWORDS = [
  "flight","airline","departure gate","arrival gate","boarding","layover","connection",
  "round trip","one way"
];

function smartContainsAny(lower, keywords) {
  return keywords.some(k => lower.includes(k));
}

// Find every "Month Day" / "ISO" / "M/D" date plus bare days that inherit
// the previous month/year. Returns ISO strings in document order.
function smartFindDates(text, defaultYear) {
  const monthMap = {
    jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, sept:9, oct:10, nov:11, dec:12,
    january:1, february:2, march:3, april:4, june:6, july:7, august:8, september:9, october:10, november:11, december:12,
  };
  const out = [];
  // Month + day (with optional ordinal suffix and year)
  const monthDayRx = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s*(\d{1,2})(?:st|nd|rd|th)?(?:[, ]+(\d{4}))?/gi;
  // Pure ISO YYYY-MM-DD
  const isoRx = /\b(\d{4})-(\d{2})-(\d{2})\b/g;
  // M/D[/Y]
  const slashRx = /\b(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b/g;
  // DDMON form (e.g. 19DEC)
  const ddmonRx = /\b(\d{1,2})(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b/gi;

  function push(idx, iso) { out.push({ idx, iso }); }

  let m;
  while ((m = monthDayRx.exec(text))) {
    const mo = monthMap[m[1].toLowerCase()];
    if (mo) push(m.index, `${m[3] ? +m[3] : defaultYear}-${String(mo).padStart(2,"0")}-${String(+m[2]).padStart(2,"0")}`);
  }
  while ((m = isoRx.exec(text))) {
    push(m.index, `${m[1]}-${m[2]}-${m[3]}`);
  }
  while ((m = slashRx.exec(text))) {
    let y = m[3] ? +m[3] : defaultYear;
    if (y < 100) y += 2000;
    push(m.index, `${y}-${String(+m[1]).padStart(2,"0")}-${String(+m[2]).padStart(2,"0")}`);
  }
  while ((m = ddmonRx.exec(text))) {
    const mo = monthMap[m[2].toLowerCase()];
    if (mo) push(m.index, `${defaultYear}-${String(mo).padStart(2,"0")}-${String(+m[1]).padStart(2,"0")}`);
  }
  out.sort((a, b) => a.idx - b.idx);
  // De-dupe overlapping matches at the same position.
  const seen = new Set();
  return out.filter(d => {
    const key = `${d.idx}|${d.iso}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function smartFindTimes(text) {
  const out = [];
  const rx = /\b(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)?\b/g;
  let m;
  while ((m = rx.exec(text))) {
    let h = +m[1];
    const min = +m[2];
    if (h > 23 || min > 59) continue;
    if (m[3]) {
      const isPM = /pm/i.test(m[3]);
      if (isPM && h < 12) h += 12;
      if (!isPM && h === 12) h = 0;
    }
    out.push({ idx: m.index, iso: `${String(h).padStart(2,"0")}:${String(min).padStart(2,"0")}` });
  }
  return out;
}

function smartFindAirportCodes(text) {
  // Three uppercase letters in parens or surrounded by non-word chars.
  const codes = [];
  const rx = /\(([A-Z]{3})\)|\b([A-Z]{3})\b/g;
  let m;
  while ((m = rx.exec(text))) {
    const code = m[1] || m[2];
    if (code) codes.push(code);
  }
  return codes;
}

function parseSmart(text, defaultYear) {
  const trimmed = text.trim();
  if (trimmed.length < 4) return null;
  const lower = trimmed.toLowerCase();

  const dates = smartFindDates(trimmed, defaultYear);
  if (dates.length === 0) return null;
  const start = dates[0].iso;
  let end = dates[dates.length - 1].iso;
  // If a single date appears twice (start == end == one match), keep them equal.
  // If the second is earlier than the first, swap.
  if (end < start) end = start;

  const times = smartFindTimes(trimmed);
  let startTime = times[0]?.iso || null;
  let endTime = times.length > 1 ? times[times.length - 1].iso : null;

  // Lane detection — most-specific keyword wins.
  let lane = "activities", color = "emerald";
  if (smartContainsAny(lower, SMART_CRUISE_KEYWORDS)) { lane = "lodging"; color = "amber"; }
  else if (smartContainsAny(lower, SMART_RENTAL_KEYWORDS)) { lane = "rental"; color = "orange"; }
  else if (smartContainsAny(lower, SMART_FLIGHT_KEYWORDS) || smartContainsAny(lower, SMART_AIRLINE_KEYWORDS)) { lane = "flights"; color = "indigo"; }
  else if (smartContainsAny(lower, SMART_HOTEL_KEYWORDS)) { lane = "lodging"; color = "amber"; }
  else if (start !== end) { lane = "location"; color = "violet"; }

  // Title — pick the most informative line / brand / route.
  const lines = trimmed.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  let title = null;

  if (lane === "flights") {
    const codes = smartFindAirportCodes(trimmed);
    if (codes.length >= 2) title = `${codes[0]} → ${codes[codes.length - 1]}`;
    if (!title) {
      // First airline name found
      for (const k of SMART_AIRLINE_KEYWORDS) {
        const idx = lower.indexOf(k);
        if (idx >= 0) { title = trimmed.slice(idx, idx + k.length).replace(/\b\w/g, c => c.toUpperCase()); break; }
      }
    }
    if (!title) title = "Flight";
  } else if (lane === "lodging") {
    // Look for a "Hotel" header on its own line, take the next line (cruise
    // confirmations, hotel emails often have this).
    for (let i = 0; i < lines.length - 1; i++) {
      if (/^(hotel|resort|cruise|property|accommodation)s?\s*$/i.test(lines[i])) {
        title = lines[i + 1];
        break;
      }
    }
    if (!title) {
      // Else: the first line that contains a known hotel/cruise brand.
      const allKeys = [...SMART_HOTEL_KEYWORDS, ...SMART_CRUISE_KEYWORDS];
      title = lines.find(l => allKeys.some(k => l.toLowerCase().includes(k)));
    }
    if (!title) title = lines[0] || "Lodging";
  } else if (lane === "rental") {
    title = lines.find(l => SMART_RENTAL_KEYWORDS.some(k => l.toLowerCase().includes(k))) || "Rental car";
  } else {
    // Activity / location: first non-trivial line that isn't a date or time.
    title = lines.find(l =>
      !/^\d/.test(l) && l.length > 2 && !/(^date:|^confirmation|^arrive:|^depart:|^check[- ]in)/i.test(l)
    ) || "Event";
  }
  // Trim long titles to something reasonable.
  if (title.length > 80) title = title.slice(0, 78).trim() + "…";

  // For non-flight events, drop times unless an explicit AM/PM was given —
  // an arbitrary "12:34" in body copy shouldn't pin the bar to that hour.
  if (lane !== "flights") {
    const hasAmPm = /\b\d{1,2}:\d{2}\s*(AM|PM)/i.test(trimmed);
    if (!hasAmPm) { startTime = null; endTime = null; }
  }

  const ev = {
    id: uid(),
    title,
    lane,
    color,
    start,
    end,
    notes: "Auto-detected from paste",
  };
  if (startTime) ev.startTime = startTime;
  if (endTime && endTime !== startTime) ev.endTime = endTime;
  return [ev];
}

// Loose natural-language parser for inputs like "Orlando from Dec 19-26",
// "Hawaii Mar 1 to Mar 8", or "concert on jul 4". Handles cases where the
// stricter parseCommand doesn't fire because there's no "add" prefix.
function parseLooseEvent(text, defaultYear) {
  let trimmed = text.trim();
  if (!trimmed) return null;
  if (trimmed.split(/\r?\n/).filter(l => l.trim()).length > 2) return null;
  trimmed = trimmed.replace(/^add\s+(?:placeholder\s+)?/i, "");

  // Scan the whole string for date references. A token is either "Month Day"
  // (Dec 27, jan1, December 27, 2026) or a bare day number that inherits the
  // most recently seen month/year. Year auto-rolls forward when a later
  // month wraps below an earlier one (Dec→Jan).
  const monthMap = {
    jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, sept:9, oct:10, nov:11, dec:12,
    january:1, february:2, march:3, april:4, june:6, july:7, august:8, september:9, october:10, november:11, december:12,
  };
  // Allow ordinal suffix on day ("July 4th 2026", "1st", "2nd", "3rd").
  // Lookbehind/ahead for ":" excludes digits inside times like "7:10pm".
  const tokenRx = /\b(?:(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s*)?(?<!:)(\d{1,2})(?!:)(?:st|nd|rd|th)?(?:,?\s+(\d{4}))?\b/gi;
  const dates = [];
  let firstDateIdx = -1;
  let curYear = defaultYear;
  let curMonth = null;
  let prevMonth = null;
  let mm;
  while ((mm = tokenRx.exec(trimmed)) !== null) {
    let matchedMonth = false;
    if (mm[1]) {
      const newMonth = monthMap[mm[1].toLowerCase()];
      if (newMonth) {
        if (prevMonth !== null && newMonth < prevMonth) curYear++;
        curMonth = newMonth;
        prevMonth = newMonth;
        matchedMonth = true;
      }
    }
    if (mm[3]) curYear = +mm[3];
    if (curMonth === null) continue; // bare day before any month → skip
    const day = +mm[2];
    if (day < 1 || day > 31) continue;
    if (firstDateIdx < 0) firstDateIdx = mm.index;
    dates.push(`${curYear}-${String(curMonth).padStart(2,"0")}-${String(day).padStart(2,"0")}`);
  }
  if (dates.length === 0) return null;

  const start = dates[0];
  const end = dates[dates.length - 1];

  // Title is everything before the first date, with trailing connector words
  // (on, from, at, in, during) stripped.
  let title = trimmed.slice(0, firstDateIdx).trim();
  title = title.replace(/\s+(?:on|from|at|in|during)\s*$/i, "").trim();
  if (!title) title = "Event";

  // Lane heuristic: a lodging/flight/activity keyword wins; otherwise a
  // multi-day stretch defaults to a location, single day to an activity.
  let lane = "location";
  if (/\b(rental car|car rental|rental|hertz|enterprise|avis|sixt|budget car|alamo|national car|thrifty|dollar rent|europcar)\b/i.test(title)) lane = "rental";
  else if (/\b(hotel|hotels|lodging|inn|resort|villa|airbnb|vrbo|cruise|cruises|hilton|marriott|hyatt|sheraton|westin|wyndham|disney's)\b/i.test(title)) lane = "lodging";
  else if (/\b(flight|flights|airline|delta|united|american|southwest|alaska airlines|jetblue|lufthansa|emirates|qatar)\b/i.test(title)) lane = "flights";
  else if (/\b(boat|charter|tour|excursion|tickets?|game|show|concert|dinner|tasting|reservation|admission|pass|safari)\b/i.test(title)) lane = "activities";
  else if (/\b(activity|activities|tour|excursion|concert|show|game|dinner)\b/i.test(title)) lane = "activities";
  else if (start === end) lane = "activities";

  // Strip a leading "for"/"at"/"in"/"to" if present.
  title = title.replace(/^(?:for|at|in|to)\s+/i, "").trim();

  const colorMap = { lodging: "amber", flights: "indigo", location: "violet", activities: "emerald", rental: "orange" };
  return [{
    id: uid(),
    title,
    lane,
    color: colorMap[lane] || "emerald",
    start,
    end,
    notes: "Added via paste",
  }];
}

// Parse Delta-style itinerary blocks where each flight reads:
//   Sat, 19DEC      DEPART      ARRIVE
//   DELTA 358
//   Delta Comfort Classic (S)\tSEATTLE
//   11:55AM\tTAMPA
//   08:21PM
// City names are spelled out (not codes) and dates have no year.
function parseDeltaItinerary(text, defaultYear) {
  const rawLines = text.split(/\r?\n/);
  // Date line variants: "Sat, 19DEC", "Sat 19 DEC", "Sat, 19 Dec 2026".
  const dateRx = /^([A-Za-z]{3}),?\s+(\d{1,2})\s*([A-Z]{3})/i;
  const monthMap = { JAN:1, FEB:2, MAR:3, APR:4, MAY:5, JUN:6, JUL:7, AUG:8, SEP:9, OCT:10, NOV:11, DEC:12 };
  const blocks = [];
  let cur = null;
  for (const line of rawLines) {
    const dm = line.trim().match(dateRx);
    if (dm) {
      if (cur) blocks.push(cur);
      cur = { dm, lines: [] };
    } else if (cur) {
      cur.lines.push(line);
    }
  }
  if (cur) blocks.push(cur);
  if (blocks.length === 0) return null;

  // Need to see "AIRLINE FLIGHTNUM" (e.g. DELTA 358) somewhere in at least
  // one block — otherwise this isn't a Delta-style paste.
  const flightNumRx = /^\s*([A-Za-z]{2,}(?:\s+[A-Za-z]+)*)\s+(\d{1,4})\s*$/;
  const looksDelta = blocks.some(b =>
    b.lines.some(l => flightNumRx.test(l)));
  if (!looksDelta) return null;

  const events = [];
  let yearOffset = 0;
  let prevMon = null;

  for (const block of blocks) {
    const day = +block.dm[2];
    const mon = monthMap[block.dm[3]];
    if (!mon) continue;
    if (prevMon !== null && mon < prevMon) yearOffset++;
    prevMon = mon;
    const year = defaultYear + yearOffset;

    let flightNum = null, depCity = null, arrCity = null, depTime = null, arrTime = null;
    for (const raw of block.lines) {
      const line = raw.trim();
      if (!line) continue;

      let m;
      // Flight number line: "DELTA 358"
      if (!flightNum && (m = line.match(/^([A-Z]{2,})\s+(\d{1,4})\s*$/))) {
        const code = AIRLINE_CODE[m[1]] || m[1].slice(0, 2);
        flightNum = `${code}${m[2]}`;
        continue;
      }

      // Tab- or wide-space-separated cabin/time + city ("11:55AM\tTAMPA"
      // or "11:55AM   TAMPA"). Some systems strip tabs to spaces on paste.
      const parts = raw.split(/\t+|\s{2,}/).map(p => p.trim()).filter(Boolean);
      const timeAt = (s) => s.match(/^(\d{1,2}:\d{2})\s*(AM|PM)\s*$/i);
      if (parts.length >= 2) {
        const left = parts[0], right = parts[parts.length - 1];
        const lt = timeAt(left);
        if (lt) {
          // Time on the left → arrival city on the right.
          if (!depTime) depTime = to24h(lt[1], lt[2].toUpperCase());
          else if (!arrTime) arrTime = to24h(lt[1], lt[2].toUpperCase());
          if (right && /^[A-Z][A-Z ]+$/.test(right)) {
            if (depCity && !arrCity) arrCity = right;
            else if (!depCity) depCity = right;
          }
          continue;
        }
        if (right && /^[A-Z][A-Z ]+$/.test(right)) {
          if (!depCity) depCity = right;
          else if (!arrCity) arrCity = right;
          continue;
        }
      }

      // Bare time line: "08:21PM"
      const t = timeAt(line);
      if (t) {
        if (!depTime) depTime = to24h(t[1], t[2].toUpperCase());
        else if (!arrTime) arrTime = to24h(t[1], t[2].toUpperCase());
        continue;
      }
      // Bare city line.
      if (/^[A-Z][A-Z ]+$/.test(line)) {
        if (!depCity) depCity = line;
        else if (!arrCity) arrCity = line;
      }
    }

    if (!depCity || !arrCity || !depTime || !arrTime) continue;

    const depCode = cityToCode(depCity);
    const arrCode = cityToCode(arrCity);

    let endY = year, endM = mon, endD = day;
    if (arrTime < depTime) {
      const d2 = new Date(year, mon - 1, day);
      d2.setDate(d2.getDate() + 1);
      endY = d2.getFullYear(); endM = d2.getMonth() + 1; endD = d2.getDate();
    }
    const startISO = `${year}-${String(mon).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
    const endISO   = `${endY}-${String(endM).padStart(2,"0")}-${String(endD).padStart(2,"0")}`;

    events.push({
      id: uid(),
      title: `${flightNum || "Flight"} ${depCode} → ${arrCode}`,
      lane: "flights", color: "indigo",
      start: startISO, startTime: depTime, startTz: AIRPORT_TZ[depCode] || "UTC",
      end:   endISO,   endTime:   arrTime, endTz:   AIRPORT_TZ[arrCode] || "UTC",
      notes: `${depCity.replace(/\s+/g,' ').trim()} → ${arrCity.replace(/\s+/g,' ').trim()}`,
    });
  }

  return events.length ? events : null;
}

function parseRangeFromText(text, defaultYear) {
  const start = consumeDate(text, defaultYear);
  if (!start) return null;
  let remainder = text.slice(start.consumed).trim();
  let end = null;
  // Allow "to/through/until" with required space, OR "-" / en-dash with no
  // space ("Dec 27-Jan 3" or "Dec 27-Jan3").
  const toMatch = remainder.match(/^(?:(?:to|through|until)\s+|[\-–]\s*)/i);
  if (toMatch) {
    const after = remainder.slice(toMatch[0].length);
    const endParse = consumeDate(after, defaultYear);
    if (endParse) {
      end = endParse.iso;
      remainder = after.slice(endParse.consumed).trim();
    }
  }
  return { start: start.iso, end, remainder };
}

function consumeDate(text, defaultYear) {
  const monthMap = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
    january: 1, february: 2, march: 3, april: 4, june: 6, july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
  };
  // "Month Day[, Year]" — also allow no space: "Jan3", "Dec25".
  let m = text.match(/^([A-Za-z]+)\.?\s*(\d{1,2})(?:,?\s+(\d{4}))?\b/);
  if (m && monthMap[m[1].toLowerCase()]) {
    const mo = monthMap[m[1].toLowerCase()];
    const d = +m[2];
    const y = m[3] ? +m[3] : defaultYear;
    return { iso: `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`, consumed: m[0].length };
  }
  // "YYYY-MM-DD"
  m = text.match(/^(\d{4})-(\d{2})-(\d{2})\b/);
  if (m) return { iso: `${m[1]}-${m[2]}-${m[3]}`, consumed: m[0].length };
  // "M/D[/YY[YY]]"
  m = text.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b/);
  if (m) {
    const mo = +m[1];
    const d = +m[2];
    let y = defaultYear;
    if (m[3]) {
      y = +m[3];
      if (y < 100) y += 2000;
    }
    return { iso: `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`, consumed: m[0].length };
  }
  return null;
}

// Pull a total price out of a paste — looks for a "total"/"price"/"amount"
// line followed by a $/USD amount. Returns the numeric value or null. Kept
// conservative so a bare "$5" inside a description doesn't get picked up.
function detectTotalPrice(text) {
  // "Total ... $1,234.56" — currency before amount.
  const rxBefore = /(?:total\s*(?:price|cost|amount|due|charged|for[^\n]*?|stay[^\n]*?)?|grand\s*total|amount\s*due)[^\n]*?(?:US\s*\$|USD\s*\$?|CAD\s*\$?|\$)\s*([\d]{1,3}(?:,\d{3})*(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)/i;
  // "Total ... 1,234.56 USD" — amount before currency.
  const rxAfter = /(?:total\s*(?:price|cost|amount|for[^\n]*?|stay[^\n]*?)?|grand\s*total)[^\n]*?([\d]{1,3}(?:,\d{3})*(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)\s*(?:USD|CAD|US\$)/i;
  const m = text.match(rxBefore) || text.match(rxAfter);
  if (!m) return null;
  const n = parseFloat(m[1].replace(/,/g, ""));
  return isFinite(n) && n > 0 ? n : null;
}

function parseFlights(text, defaultYear) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const events = [];

  const dateRx       = /^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat),?\s+([A-Za-z]{3})\s+(\d{1,2})(?:,\s*(\d{4}))?$/;
  const stopRx       = /^(\d{1,2}:\d{2})\s*(AM|PM)(?:\+(\d))?\s*([A-Za-z'.\-,&/ ]+?)\s*\((\w{3})\)\s*$/;
  const travelRx     = /^Travel time:\s*(.+?)(?:Overnight)?\s*$/i;
  const layoverRx    = /^(\d+\s*hr(?:\s*\d+\s*min)?)\s*layover\s*([A-Za-z' ]+?)\s*\((\w{3})\)(?:Overnight layover)?\s*$/i;
  const flightEndRx  = /[A-Z]{2,3}\s?\d{1,4}$/;
  const flightFullRx = /^(.+?)(Business|Economy|First|Premium Economy)\s*(?:\([^)]+\)\s*)?(.+?)([A-Z]{2,3})\s?(\d{1,4})$/;

  let curDate = null;
  let depStop = null, depOff = 0;
  let arrStop = null, arrOff = 0;
  let travelTime = null;
  let pendingLayover = null;
  let flightLines = [];

  function pushFlight(flightNum, notes) {
    if (!depStop || !arrStop || !curDate) return;
    const startDate = depOff > 0 ? toISO(addDays(parseDay(curDate), depOff)) : curDate;
    const endDate   = arrOff > 0 ? toISO(addDays(parseDay(curDate), arrOff)) : curDate;

    if (pendingLayover) {
      let prev = null;
      for (let i = events.length - 1; i >= 0; i--) {
        if (events[i].lane === "flights" && !events[i].title.endsWith("layover")) {
          prev = events[i];
          break;
        }
      }
      if (prev) {
        events.push({
          id: uid(),
          title: `${pendingLayover.code} layover`,
          lane: "flights",
          color: "grey",
          start: prev.end, startTime: prev.endTime, startTz: prev.endTz,
          end:   startDate, endTime:   depStop.time, endTz:   AIRPORT_TZ[depStop.code] || "UTC",
          notes: `${pendingLayover.duration} layover in ${pendingLayover.city}`,
        });
      }
      pendingLayover = null;
    }

    events.push({
      id: uid(),
      title: `${flightNum} ${depStop.code} → ${arrStop.code}`,
      lane: "flights",
      color: "indigo",
      start: startDate, startTime: depStop.time, startTz: AIRPORT_TZ[depStop.code] || "UTC",
      end:   endDate,   endTime:   arrStop.time, endTz:   AIRPORT_TZ[arrStop.code] || "UTC",
      notes,
    });

    depStop = arrStop = null;
    depOff = arrOff = 0;
    travelTime = null;
    flightLines = [];
  }

  function tryParseFlightInfo() {
    const concat = flightLines.join("");
    const m = concat.match(flightFullRx);
    if (!m) return false;
    const airline  = m[1].trim();
    const cabin    = m[2].trim();
    const aircraft = m[3].trim();
    const flightNo = `${m[4]} ${m[5]}`;
    const notes = `${airline} · ${cabin} · ${aircraft}${travelTime ? " · " + travelTime : ""}`;
    pushFlight(flightNo, notes);
    return true;
  }

  for (const line of lines) {
    let m;
    if ((m = line.match(dateRx))) {
      const month = MONTHS[m[1]];
      if (!month) continue;
      const day = +m[2];
      const year = m[3] ? +m[3] : defaultYear;
      curDate = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      depStop = arrStop = null;
      depOff = arrOff = 0;
      flightLines = [];
      continue;
    }
    if ((m = line.match(layoverRx))) {
      pendingLayover = { duration: m[1].trim(), city: m[2].trim(), code: m[3] };
      continue;
    }
    if ((m = line.match(stopRx))) {
      const off = m[3] ? +m[3] : 0;
      const stop = { time: to24h(m[1], m[2]), name: m[4].trim(), code: m[5] };
      if (!depStop) {
        depStop = stop; depOff = off;
      } else if (!arrStop) {
        arrStop = stop; arrOff = off;
      } else {
        // Already have dep+arr but no flight info found — abandon and start new flight
        depStop = stop; depOff = off;
        arrStop = null; arrOff = 0;
        flightLines = [];
      }
      continue;
    }
    if ((m = line.match(travelRx))) {
      travelTime = m[1].trim();
      continue;
    }
    // Otherwise: accumulate as flight info if we have dep + arr
    if (depStop && arrStop) {
      flightLines.push(line);
      if (flightEndRx.test(line)) {
        tryParseFlightInfo();
      }
    }
  }

  return events;
}

// --- airport display names (for derived locations) ---

const AIRPORT_NAMES = {
  YVR: "Vancouver",
  SEA: "Seattle",
  FRA: "Frankfurt",
  ZNZ: "Zanzibar",
  ARK: "Arusha",
  JRO: "Kilimanjaro",
  DAR: "Dar es Salaam",
  NBO: "Nairobi",
  ADD: "Addis Ababa",
  CPT: "Cape Town",
  JNB: "Johannesburg",
  SEZ: "Seychelles",
  MRU: "Mauritius",
  LAX: "Los Angeles",
  JFK: "New York",
};

function arrCodeFromTitle(title) {
  const m = title.match(/→\s*(\w{3})\s*$/);
  return m ? m[1] : null;
}
function depCodeFromTitle(title) {
  const m = title.match(/(\w{3})\s*→/);
  return m ? m[1] : null;
}

// Given an option's flight events, derive "Where" stays for any
// round-trip pattern (arrive at X, later depart X without a layover bridging).
function fmtDuration(ms) {
  if (ms <= 0) return "0m";
  const totalMin = Math.round(ms / 60000);
  const d = Math.floor(totalMin / (60 * 24));
  const h = Math.floor((totalMin % (60 * 24)) / 60);
  const m = totalMin % 60;
  const parts = [];
  if (d) parts.push(`${d}d`);
  if (h || d) parts.push(`${h}h`);
  if (!d && m) parts.push(`${m}m`);
  return parts.join(" ") || "0m";
}

// Hours-only formatter: "12h" or "12h 30m" (no days rollover).
function fmtHours(ms) {
  if (ms <= 0) return "0h";
  const totalMin = Math.round(ms / 60000);
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return m > 0 ? `${h}h ${m}m` : `${h}h`;
}

function summarizeOption(opt) {
  const homeTz = state.homeTz || "America/Los_Angeles";
  const flights = opt.events.filter(e =>
    e.lane === "flights" && !e.title.endsWith("layover"));
  const layovers = opt.events.filter(e =>
    e.lane === "flights" && e.title.endsWith("layover"));

  const sumMs = arr => arr.reduce((s, ev) => {
    const { sUtc, eUtc } = eventBounds(ev, {}, homeTz);
    return s + Math.max(0, eUtc - sUtc);
  }, 0);

  const flightMs = sumMs(flights);
  const layoverMs = sumMs(layovers);

  // Derived stays + any user-added location events in the option.
  const derived = deriveLocationsForOption(opt);
  const userLocations = opt.events.filter(e => e.lane === "location");
  const allStays = [...userLocations, ...derived];
  const stays = allStays.map(ev => {
    const { sUtc, eUtc } = eventBounds(ev, {}, homeTz);
    return { name: ev.title, ms: Math.max(0, eUtc - sUtc) };
  });

  return { flightMs, layoverMs, stays };
}

function deriveLocationsForOption(opt) {
  const homeTz = state.homeTz || "America/Los_Angeles";
  const flights = opt.events.filter(e =>
    e.lane === "flights" && !e.title.endsWith("layover"));
  const layovers = opt.events.filter(e =>
    e.lane === "flights" && e.title.endsWith("layover"));

  const flightItems = flights.map(ev => {
    const { sUtc, eUtc } = eventBounds(ev, {}, homeTz);
    return { ev, sUtc, eUtc };
  }).sort((a, b) => a.sUtc - b.sUtc);

  const layoverItems = layovers.map(lv => {
    const { sUtc, eUtc } = eventBounds(lv, {}, homeTz);
    return { lv, sUtc, eUtc };
  });

  const out = [];
  for (let i = 0; i < flightItems.length - 1; i++) {
    const cur = flightItems[i];
    const next = flightItems[i + 1];
    const arrCode = arrCodeFromTitle(cur.ev.title);
    const depCode = depCodeFromTitle(next.ev.title);
    if (!arrCode || !depCode || arrCode !== depCode) continue;

    // Skip if a layover already bridges this gap at the same airport.
    const bridged = layoverItems.some(lb =>
      lb.sUtc >= cur.eUtc - 60_000 &&
      lb.eUtc <= next.sUtc + 60_000 &&
      lb.lv.title.includes(arrCode)
    );
    if (bridged) continue;

    out.push({
      id: `derived-${cur.ev.id}-${next.ev.id}`,
      title: AIRPORT_NAMES[arrCode] || arrCode,
      lane: "location",
      color: "teal",
      start:     cur.ev.end,     startTime: cur.ev.endTime,   startTz: cur.ev.endTz,
      end:       next.ev.start,  endTime:   next.ev.startTime, endTz:   next.ev.startTz,
      notes: "Auto-derived from option flights",
    });
  }
  return out;
}

// --- options view ---

function defaultOptionRange() {
  if (!state.start || !state.end) return null;
  const totalDays = dayDiff(state.start, state.end) + 1;
  const segSize = chooseSegmentSize(totalDays);
  const seg3StartIdx = 2 * segSize;
  if (seg3StartIdx >= totalDays) return { start: state.start, end: state.end };
  const segStart = toISO(addDays(parseDay(state.start), seg3StartIdx));
  return { start: segStart, end: state.end };
}

function getOptionRange() {
  const def = defaultOptionRange();
  if (!def) return null;
  return {
    start: state.optionRangeStart || def.start,
    end:   state.optionRangeEnd   || def.end,
  };
}

// Each option belongs to a group; the group carries the date range so users
// can compare several options for "last week" while staging a different set
// for "first week" under a separate header.
function ensureOptionGroups() {
  if (!Array.isArray(state.optionGroups)) state.optionGroups = [];
  if (!Array.isArray(state.options)) state.options = [];
  const orphans = state.options.filter(o => !o.groupId);
  if (orphans.length || (state.options.length === 0 && state.optionGroups.length === 0)) {
    let g = state.optionGroups[0];
    if (!g) {
      const def = defaultOptionRange();
      g = {
        id: uid(),
        name: "Options",
        start: state.optionRangeStart || def?.start || state.start,
        end:   state.optionRangeEnd   || def?.end   || state.end,
      };
      state.optionGroups.push(g);
    }
    orphans.forEach(o => { o.groupId = g.id; });
  }
}

function getOptionGroupForOption(opt) {
  if (!opt || !opt.groupId) return null;
  return state.optionGroups?.find(g => g.id === opt.groupId) || null;
}

function groupRange(group) {
  if (!group) return null;
  const def = defaultOptionRange();
  return { start: group.start || def?.start || state.start, end: group.end || def?.end || state.end };
}

function createOptionGroup(name) {
  ensureOptionGroups();
  const def = defaultOptionRange();
  const taken = new Set(state.optionGroups.map(g => g.name));
  let n = 1;
  while (taken.has(`Group ${n}`)) n++;
  const g = {
    id: uid(),
    name: name || `Group ${n}`,
    start: def?.start || state.start,
    end: def?.end || state.end,
  };
  state.optionGroups.push(g);
  return g;
}

let activeOptionGroupId = null;

function renderOptions() {
  ensureOptionGroups();
  // Natural-numeric sort by name so renames stay in number order
  // ("Option 2" before "Option 10", "1 - cheap" before "2 - flex").
  state.options.sort((a, b) =>
    (a.name || "").localeCompare(b.name || "", undefined, { numeric: true, sensitivity: "base" }));

  const list = document.getElementById("options-list");
  list.innerHTML = "";

  // Sub-nav across groups; selected group renders below.
  if (!state.optionGroups.find(g => g.id === activeOptionGroupId)) {
    activeOptionGroupId = state.optionGroups[0]?.id || null;
  }
  const subnav = document.createElement("div");
  subnav.className = "option-subnav";
  for (const g of state.optionGroups) {
    const tab = document.createElement("button");
    tab.type = "button";
    const applied = state.options.some(o => o.groupId === g.id && isOptionApplied(o.id));
    tab.className = "option-subnav-tab" + (g.id === activeOptionGroupId ? " active" : "");
    tab.textContent = (applied ? "✓ " : "") + (g.name || "(unnamed)");
    tab.addEventListener("click", () => {
      activeOptionGroupId = g.id;
      renderOptions();
    });
    subnav.appendChild(tab);
  }
  const addTab = document.createElement("button");
  addTab.type = "button";
  addTab.className = "option-subnav-add";
  addTab.textContent = "+ Add group";
  addTab.addEventListener("click", () => {
    const g = createOptionGroup();
    activeOptionGroupId = g.id;
    createOption(null, g.id);
    save();
    renderApp();
  });
  subnav.appendChild(addTab);
  const optSlot = document.getElementById("options-subnav-slot");
  if (optSlot) { optSlot.innerHTML = ""; optSlot.appendChild(subnav); }
  else list.appendChild(subnav);

  // Refresh the paste target dropdown — Options tab can only target an
  // existing option, or "+ New option" which creates one at parse time.
  const target = document.getElementById("paste-target");
  if (target) {
    const prev = target.value;
    target.innerHTML = "";
    target.disabled = false;
    for (const opt of state.options) {
      target.appendChild(new Option(opt.name, opt.id));
    }
    target.appendChild(new Option("+ New option", "__new__"));
    const stillExists = state.options.find(o => o.id === prev);
    target.value = stillExists ? prev : "__new__";
  }

  if (!state.start || !state.end) {
    list.appendChild(el("div", "empty-state", "Set your trip dates first."));
    return;
  }

  const homeTz = state.homeTz || "America/Los_Angeles";
  const tzAware = state.tzAware !== false;
  const dayTzMap = computeDayTzMap(state.start, state.end, state.events, homeTz, tzAware);

  for (const group of state.optionGroups.filter(g => g.id === activeOptionGroupId)) {
    const groupEl = el("div", "option-group");
    const optsInGroup = state.options.filter(o => o.groupId === group.id);
    const groupApplied = optsInGroup.some(o => isOptionApplied(o.id));

    // Group header
    const gHead = el("div", "option-group-head");
    const checkEl = el("span", "option-group-check", groupApplied ? "✓" : "");
    gHead.appendChild(checkEl);
    const gName = el("input", "option-group-name");
    gName.type = "text";
    gName.value = group.name;
    gName.placeholder = "Group name";
    gName.addEventListener("input", () => { group.name = gName.value; save(); });
    gHead.appendChild(gName);

    const gStart = el("input", "option-group-date");
    gStart.type = "date";
    gStart.value = group.start || "";
    gStart.addEventListener("change", () => {
      group.start = gStart.value || null;
      save();
      renderOptions();
    });
    gHead.appendChild(gStart);
    gHead.appendChild(el("span", "option-group-dash", "–"));
    const gEnd = el("input", "option-group-date");
    gEnd.type = "date";
    gEnd.value = group.end || "";
    gEnd.addEventListener("change", () => {
      group.end = gEnd.value || null;
      save();
      renderOptions();
    });
    gHead.appendChild(gEnd);

    const gActions = el("div", "option-group-actions");
    const gAddOpt = el("button", null, "+ Add option");
    gAddOpt.type = "button";
    gAddOpt.addEventListener("click", () => { createOption(null, group.id); renderApp(); });
    gActions.appendChild(gAddOpt);
    const gDel = el("button", "ghost", "Delete group");
    gDel.type = "button";
    gDel.addEventListener("click", () => {
      const inGroup = state.options.filter(o => o.groupId === group.id);
      if (inGroup.length && !confirm(`Delete group "${group.name}" and its ${inGroup.length} option${inGroup.length === 1 ? "" : "s"}?`)) return;
      state.options = state.options.filter(o => o.groupId !== group.id);
      state.optionGroups = state.optionGroups.filter(g => g.id !== group.id);
      save();
      renderApp();
    });
    gActions.appendChild(gDel);
    gHead.appendChild(gActions);
    groupEl.appendChild(gHead);

    const range = groupRange(group);
    if (optsInGroup.length === 0) {
      groupEl.appendChild(el("div", "empty-state",
        `No options yet. Click "+ Add option" to stage an alternative for ${fmtShort(parseDay(range.start))} – ${fmtShort(parseDay(range.end))}.`));
      list.appendChild(groupEl);
      continue;
    }

    for (const opt of optsInGroup) {
    const card = el("div", "option-card");

    // Head: name input + price + actions
    const head = el("div", "option-head");
    const nameInput = el("input", "option-name");
    nameInput.type = "text";
    nameInput.value = opt.name;
    nameInput.placeholder = "Option name";
    nameInput.addEventListener("input", () => { opt.name = nameInput.value; save(); });
    head.appendChild(nameInput);

    const priceWrap = el("div", "option-price-wrap");
    priceWrap.appendChild(el("span", "option-price-prefix", "$"));
    const priceInput = el("input", "option-price");
    priceInput.type = "text";
    priceInput.inputMode = "decimal";
    priceInput.placeholder = "0";
    priceInput.value = opt.price != null ? String(opt.price) : "";
    priceInput.addEventListener("input", () => {
      const v = priceInput.value.replace(/[^0-9.]/g, "");
      opt.price = v === "" ? null : Number(v);
      save();
    });
    priceWrap.appendChild(priceInput);
    head.appendChild(priceWrap);

    const actions = el("div", "option-actions");
    const addBtn = el("button", null, "+ Add event");
    addBtn.type = "button";
    addBtn.addEventListener("click", () => openEventDialog(null, opt.id));
    actions.appendChild(addBtn);

    const applied = isOptionApplied(opt.id);
    const applyBtn = el("button", applied ? "danger" : null,
      applied ? "Remove from itinerary" : "Apply to itinerary");
    applyBtn.type = "button";
    applyBtn.addEventListener("click", () => {
      if (applied) removeAppliedOption(opt.id);
      else applyOption(opt.id);
    });
    actions.appendChild(applyBtn);

    const delBtn = el("button", "ghost", "Delete");
    delBtn.type = "button";
    delBtn.addEventListener("click", () => {
      if (!confirm(`Delete option "${opt.name}"?`)) return;
      state.options = state.options.filter(o => o.id !== opt.id);
      save();
      renderApp();
    });
    actions.appendChild(delBtn);

    head.appendChild(actions);
    card.appendChild(head);

    // Combined timeline: confirmed main events (in range) + option's events
    // (dashed) + auto-derived "Where" stays from the option's flight chain.
    const mainInRange = state.events.filter(ev => !(ev.end < range.start || ev.start > range.end));
    const derivedLocations = deriveLocationsForOption(opt);
    const optEvents = [...opt.events, ...derivedLocations]
      .map(ev => ({ ...ev, _isOption: true, _optionId: opt.id }));
    const combined = [...mainInRange, ...optEvents];

    const tlEl = el("div", "timeline");
    card.appendChild(tlEl);

    const breakdownPanel = document.getElementById("tab-options").querySelector(".panel");
    const containerWidth = (breakdownPanel?.clientWidth || 1000) - 32; // panel padding
    const laneLabelW = parseInt(getComputedStyle(document.documentElement).getPropertyValue("--lane-label-w")) || 110;
    const totalDaysInRange = dayDiff(range.start, range.end) + 1;
    const dayPx = Math.max(60, Math.floor((containerWidth - laneLabelW - 30) / totalDaysInRange));

    renderTimeline(tlEl, range.start, range.end, {
      dayTzMap, homeTz, dayPx, compact: false, tzAware, events: combined,
    });

    // Stats row: time at each place + flying + layovers.
    const summary = summarizeOption(opt);
    const stats = el("div", "option-stats");
    if (summary.stays.length === 0 && summary.flightMs === 0 && summary.layoverMs === 0) {
      stats.appendChild(el("span", "stat muted", "Add flights to see time breakdown"));
    } else {
      // Group stays by location name (user might have multiple stops at same place).
      const grouped = new Map();
      for (const s of summary.stays) {
        grouped.set(s.name, (grouped.get(s.name) || 0) + s.ms);
      }
      for (const [name, ms] of grouped) {
        const chip = el("span", "stat stat-place");
        chip.appendChild(el("span", "stat-label", name));
        chip.appendChild(el("span", "stat-value", fmtDuration(ms)));
        stats.appendChild(chip);
      }
      const flying = el("span", "stat stat-flying");
      flying.appendChild(el("span", "stat-label", "Flying"));
      flying.appendChild(el("span", "stat-value", fmtHours(summary.flightMs)));
      stats.appendChild(flying);

      const lay = el("span", "stat stat-layover");
      lay.appendChild(el("span", "stat-label", "Layovers"));
      lay.appendChild(el("span", "stat-value", fmtHours(summary.layoverMs)));
      stats.appendChild(lay);

      const total = summary.flightMs + summary.layoverMs;
      const transit = el("span", "stat stat-transit");
      transit.appendChild(el("span", "stat-label", "Transit total"));
      transit.appendChild(el("span", "stat-value", fmtHours(total)));
      stats.appendChild(transit);
    }
    card.appendChild(stats);

    groupEl.appendChild(card);
    }

    list.appendChild(groupEl);
  }

  renderComparison();
}

function escHtml(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}

function renderComparison() {
  const container = document.getElementById("options-comparison");
  if (!container) return;
  if (state.options.length < 2) { container.innerHTML = ""; return; }

  const rows = state.options.map(opt => {
    const summary = summarizeOption(opt);
    const grouped = new Map();
    for (const s of summary.stays) grouped.set(s.name, (grouped.get(s.name) || 0) + s.ms);
    const stays = [...grouped.entries()]
      .map(([name, ms]) => `${name}: ${fmtDuration(ms)}`)
      .join(", ") || "—";
    const transit = summary.flightMs + summary.layoverMs;
    const stayMs = [...grouped.values()].reduce((a, b) => a + b, 0);
    return {
      opt,
      stays,
      flying: summary.flightMs,
      layovers: summary.layoverMs,
      transit,
      stayMs,
      price: opt.price || 0,
    };
  });

  // Highlight bests in each numeric column.
  const minBy = key => Math.min(...rows.map(r => r[key] || Infinity));
  const maxBy = key => Math.max(...rows.map(r => r[key]));
  const bestStay   = maxBy("stayMs");
  const bestFly    = minBy("flying");
  const bestLay    = minBy("layovers");
  const bestTrans  = minBy("transit");
  const bestPrice  = Math.min(...rows.filter(r => r.price > 0).map(r => r.price));

  const html = `
    <h3>Comparison</h3>
    <table class="comparison-table">
      <thead>
        <tr>
          <th>Option</th>
          <th>Destination time</th>
          <th>Flying</th>
          <th>Layovers</th>
          <th>Transit total</th>
          <th>Price</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map(r => `
          <tr>
            <td class="opt-name">${escHtml(r.opt.name)}</td>
            <td class="${r.stayMs === bestStay ? "best" : ""}">${escHtml(r.stays)}</td>
            <td class="num ${r.flying === bestFly ? "best" : ""}">${fmtHours(r.flying)}</td>
            <td class="num ${r.layovers === bestLay ? "best" : ""}">${fmtHours(r.layovers)}</td>
            <td class="num ${r.transit === bestTrans ? "best" : ""}">${fmtHours(r.transit)}</td>
            <td class="num price ${r.price === bestPrice && r.price > 0 ? "best" : ""}">${r.price > 0 ? "$" + r.price.toLocaleString() : "—"}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
  container.innerHTML = html;
}

function isOptionApplied(optId) {
  return state.events.some(e => e._appliedFrom === optId);
}

function applyOption(optId) {
  const opt = state.options.find(o => o.id === optId);
  if (!opt) return;
  if (opt.events.length === 0) {
    alert("This option has no events yet.");
    return;
  }
  const derived = deriveLocationsForOption(opt);
  const total = opt.events.length + derived.length;
  if (!confirm(`Add ${total} event(s) from "${opt.name}" to your itinerary?\n(${derived.length} auto-derived location event${derived.length === 1 ? "" : "s"} included.)`)) return;
  for (const ev of [...opt.events, ...derived]) {
    const copy = { ...ev, id: uid(), _appliedFrom: optId };
    delete copy._isOption;
    delete copy._optionId;
    state.events.push(copy);
  }
  save();
  renderApp();
}

function removeAppliedOption(optId) {
  const opt = state.options.find(o => o.id === optId);
  const name = opt ? opt.name : "this option";
  const count = state.events.filter(e => e._appliedFrom === optId).length;
  if (!confirm(`Remove ${count} event(s) added from "${name}" from your itinerary?`)) return;
  state.events = state.events.filter(e => e._appliedFrom !== optId);
  save();
  renderApp();
}

// --- view dispatcher / tabs ---

function renderApp() {
  // Preserve scroll position across re-renders so deleting / editing an
  // event doesn't snap the page back to the top.
  const scrollY = window.scrollY;

  // Always sync the topbar fields, regardless of which view is active.
  document.getElementById("trip-name").value = state.name || "";
  document.getElementById("trip-start").value = state.start || "";
  document.getElementById("trip-end").value = state.end || "";
  document.getElementById("tz-aware").checked = state.tzAware !== false;
  const todoCb = document.getElementById("enable-tab-todo");
  const optCb = document.getElementById("enable-tab-options");
  const cmpCb = document.getElementById("enable-tab-compare");
  const rntCb = document.getElementById("enable-tab-rental");
  if (todoCb) todoCb.checked = state.enabledTabs?.todo !== false;
  if (optCb) optCb.checked = state.enabledTabs?.options !== false;
  if (cmpCb) cmpCb.checked = state.enabledTabs?.compare !== false;
  if (rntCb) rntCb.checked = state.enabledTabs?.rental === true;
  // Refresh the click-to-edit display row.
  const displayName = document.getElementById("trip-display-name");
  const displayDates = document.getElementById("trip-display-dates");
  if (displayName) displayName.textContent = state.name || "Untitled trip";
  if (displayDates) {
    if (state.start && state.end) {
      const fmt = d => {
        const [y, m, da] = d.split("-").map(Number);
        return new Date(y, m - 1, da).toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
      };
      displayDates.textContent = `${fmt(state.start)} – ${fmt(state.end)}`;
    } else {
      displayDates.textContent = "Set dates";
    }
  }

  // Tab visibility per role:
  //   owner / editor → all tabs
  //   viewer-pricing → My itin + Pricing
  //   viewer / public link → My itin only
  const pricingAllowed = CAN_SEE_PRICING;
  const editorTabsAllowed = CAN_EDIT;
  const allowedTabs = new Set(["main"]);
  if (pricingAllowed) allowedTabs.add("pricing");
  if (editorTabsAllowed) {
    if (state.enabledTabs?.todo !== false) allowedTabs.add("todo");
    if (state.enabledTabs?.options !== false) allowedTabs.add("options");
    if (state.enabledTabs?.compare !== false) allowedTabs.add("compare");
    if (state.enabledTabs?.rental === true) allowedTabs.add("rental");
  }
  if (!allowedTabs.has(state.activeView)) state.activeView = "main";
  document.querySelectorAll(".tab-btn").forEach(b => {
    b.hidden = !allowedTabs.has(b.dataset.tab);
    b.classList.toggle("active", b.dataset.tab === state.activeView);
  });
  document.getElementById("tab-main").hidden = state.activeView !== "main";
  document.getElementById("tab-todo").hidden = !allowedTabs.has("todo") || state.activeView !== "todo";
  document.getElementById("tab-pricing").hidden = !pricingAllowed || state.activeView !== "pricing";
  document.getElementById("tab-options").hidden = !allowedTabs.has("options") || state.activeView !== "options";
  document.getElementById("tab-compare").hidden = !allowedTabs.has("compare") || state.activeView !== "compare";
  document.getElementById("tab-rental").hidden = !allowedTabs.has("rental") || state.activeView !== "rental";

  updateSyncIndicator();

  if (state.activeView === "main") render();
  else if (state.activeView === "todo") renderTodoList();
  else if (state.activeView === "pricing") renderPricing();
  else if (state.activeView === "compare") renderHotelCompare();
  else if (state.activeView === "rental") renderRentalCompare();
  else renderOptions();

  window.scrollTo({ top: scrollY, behavior: "instant" });
}

document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    state.activeView = btn.dataset.tab;
    saveLocal();   // tab choice is per-device UI, don't push to worker
    renderApp();
  });
});

document.getElementById("todo-add-form")?.addEventListener("submit", (e) => {
  e.preventDefault();
  const input = document.getElementById("todo-add-input");
  const title = input.value.trim();
  if (!title) return;
  // A free-form to-do has no firm date or lane yet — drop it on the activities
  // lane for the trip's start so it sorts to the top and can be edited later.
  const day = state.start || new Date().toISOString().slice(0, 10);
  state.events.push({
    id: uid(),
    title,
    lane: "activities",
    color: "violet",
    start: day,
    end: day,
    tentative: true,
    todo: true,   // typed to-do — keep it out of the timeline/breakdown
  });
  input.value = "";
  save();
  renderApp();
  // Re-focus so the user can keep adding items.
  document.getElementById("todo-add-input")?.focus();
});

document.getElementById("pricing-has-others-cb")?.addEventListener("change", (e) => {
  state.pricingHasOthers = e.target.checked;
  if (!state.pricingHasOthers) pricingSelection.clear();
  save();
  renderPricing();
});
document.getElementById("pricing-add-btn")?.addEventListener("click", addPricingLineItem);
document.getElementById("pricing-clear-sel")?.addEventListener("click", () => {
  pricingSelection.clear();
  renderPricingPills();
});
document.getElementById("pricing-cancel-edit")?.addEventListener("click", cancelLineItemEdit);
document.getElementById("pricing-add-split-group")?.addEventListener("click", () => {
  ensurePriceSplit();
  if (state.priceSplit.groups.length >= 4) return;
  const palette = ["indigo", "rose", "emerald", "amber"];
  const idx = state.priceSplit.groups.length;
  state.priceSplit.groups.push({
    id: "g" + (idx + 1),
    name: `Group ${idx + 1}`,
    color: palette[idx % palette.length],
    shareInput: "0",
    share: 0,
  });
  save();
  renderPricing();
});

function createOption(name, groupId) {
  ensureOptionGroups();
  function nextDefaultName() {
    const taken = new Set(state.options.map(o => o.name));
    let n = 1;
    while (taken.has(`Option ${n}`)) n++;
    return `Option ${n}`;
  }
  // Default to the last group so a fresh "+ Add option" lands where the
  // user is currently working.
  if (!groupId) {
    if (!state.optionGroups.length) createOptionGroup();
    groupId = state.optionGroups[state.optionGroups.length - 1].id;
  }
  const opt = {
    id: uid(),
    name: name || nextDefaultName(),
    groupId,
    events: [],
  };
  state.options.push(opt);
  save();
  return opt;
}

document.getElementById("add-option-btn")?.addEventListener("click", () => {
  if (!getOptionRange()) return;
  createOption();
  renderApp();
});

async function smartParseRemote(text, statusEl, existingEvents) {
  await whenFb();
  if (!window.fb.user) {
    statusEl.textContent = "Sign in to use smart parse.";
    statusEl.className = "paste-status error";
    return null;
  }
  statusEl.textContent = existingEvents ? "Asking Claude to patch existing events…" : "Parsing with Claude…";
  statusEl.className = "paste-status";
  try {
    const body = { text, tripStart: state.start || null, tripEnd: state.end || null };
    if (existingEvents) body.existingEvents = existingEvents;
    return await window.fb.smartParse(body);
  } catch (e) {
    statusEl.textContent = `Smart parse failed: ${e.message}`;
    statusEl.className = "paste-status error";
    return null;
  }
}

function wirePasteBlock({ inputId, parseId, clearId, statusId, targetId, smartId, updateId }) {
  const parseBtn = document.getElementById(parseId);
  if (!parseBtn) return;
  const smartBtn = smartId ? document.getElementById(smartId) : null;
  const updateBtn = updateId ? document.getElementById(updateId) : null;

  const insertEvents = (events, totalPrice, target, status, input) => {
    if (target === "__main__") {
      state.events.push(...events);
      let dateChanged = false;
      for (const e of events) {
        if (e.start && (!state.start || e.start < state.start)) { state.start = e.start; dateChanged = true; }
        if (e.end && (!state.end || e.end > state.end)) { state.end = e.end; dateChanged = true; }
      }
      if (dateChanged) {
        const ts = document.getElementById("trip-start"); if (ts) ts.value = state.start || "";
        const te = document.getElementById("trip-end"); if (te) te.value = state.end || "";
      }
      let pricingNote = "";
      if (totalPrice != null && events.length > 0) {
        if (!Array.isArray(state.lineItems)) state.lineItems = [];
        state.lineItems.push({ id: uid(), eventIds: events.map(e => e.id), label: null, total: totalPrice, overrides: {} });
        pricingNote = ` ($${totalPrice.toLocaleString()} added to Pricing)`;
      }
      save();
      status.textContent = `Added ${events.length} event(s).${pricingNote}`;
      status.className = "paste-status success";
      input.value = "";
      renderApp();
    } else {
      const opt = state.options.find(o => o.id === target);
      if (opt) opt.events.push(...events);
      save();
      status.textContent = `Added ${events.length} event(s).`;
      status.className = "paste-status success";
      input.value = "";
      renderApp();
    }
  };

  if (smartBtn) {
    smartBtn.addEventListener("click", async () => {
      const input = document.getElementById(inputId);
      const status = document.getElementById(statusId);
      const text = input.value;
      if (!text.trim()) {
        status.textContent = "Paste some text first.";
        status.className = "paste-status error";
        return;
      }
      smartBtn.disabled = true;
      const data = await smartParseRemote(text, status);
      smartBtn.disabled = false;
      if (!data) return;
      const events = (data.events || []).map(e => {
        const isLayover = e.lane === "flights" && (e.title || "").toLowerCase().endsWith("layover");
        return {
          id: uid(),
          title: e.title || "Event",
          lane: ["flights","lodging","activities","rental","location"].includes(e.lane) ? e.lane : "activities",
          color: isLayover ? "grey"
            : e.lane === "flights" ? "indigo"
            : e.lane === "lodging" ? "amber"
            : e.lane === "rental" ? "orange"
            : e.lane === "location" ? "violet"
            : "emerald",
          start: e.start, end: e.end || e.start,
          ...(e.startTime ? { startTime: e.startTime } : {}),
          ...(e.endTime ? { endTime: e.endTime } : {}),
          notes: e.notes || "Parsed by Claude",
        };
      }).filter(e => e.start && /^\d{4}-\d{2}-\d{2}$/.test(e.start));
      if (events.length === 0) {
        status.textContent = "Claude didn't find any events in that text.";
        status.className = "paste-status error";
        return;
      }
      let target;
      if (!targetId) target = "__main__";
      else {
        const sel = document.getElementById(targetId);
        target = sel?.value || "__new__";
        if (target === "__new__") { const newOpt = createOption(); target = newOpt.id; }
      }
      insertEvents(events, data.totalPrice, target, status, input);
    });
  }

  if (updateBtn) {
    updateBtn.addEventListener("click", async () => {
      const input = document.getElementById(inputId);
      const status = document.getElementById(statusId);
      const text = input.value;
      if (!text.trim()) {
        status.textContent = "Paste some text first.";
        status.className = "paste-status error";
        return;
      }
      if (!Array.isArray(state.events) || state.events.length === 0) {
        status.textContent = "Nothing to update — use Smart parse to add events first.";
        status.className = "paste-status error";
        return;
      }
      // Send a slim version of existing events (no notes / colors) — Claude only
      // needs the identifying fields to match against.
      const existingSlim = state.events.map(e => ({
        id: e.id, title: e.title, lane: e.lane,
        start: e.start, end: e.end,
        startTime: e.startTime, endTime: e.endTime,
      }));
      updateBtn.disabled = true;
      const data = await smartParseRemote(text, status, existingSlim);
      updateBtn.disabled = false;
      if (!data) return;
      const updates = data.updates || [];
      const newEvents = data.newEvents || [];
      let updatedCount = 0;
      for (const u of updates) {
        const ev = state.events.find(e => e.id === u.id);
        if (!ev || !u.fields) continue;
        Object.assign(ev, u.fields);
        updatedCount++;
      }
      const addedEvents = newEvents.map(e => ({
        id: uid(),
        title: e.title || "Event",
        lane: ["flights","lodging","activities","rental","location"].includes(e.lane) ? e.lane : "activities",
        color: e.lane === "flights" ? "indigo" : e.lane === "lodging" ? "amber" : e.lane === "rental" ? "orange" : e.lane === "location" ? "violet" : "emerald",
        start: e.start, end: e.end || e.start,
        ...(e.startTime ? { startTime: e.startTime } : {}),
        ...(e.endTime ? { endTime: e.endTime } : {}),
        notes: e.notes || "Parsed by Claude",
      })).filter(e => e.start && /^\d{4}-\d{2}-\d{2}$/.test(e.start));
      if (addedEvents.length > 0) state.events.push(...addedEvents);
      if (updatedCount === 0 && addedEvents.length === 0) {
        status.textContent = "Claude didn't find anything to change.";
        status.className = "paste-status error";
        return;
      }
      save();
      const parts = [];
      if (updatedCount > 0) parts.push(`Updated ${updatedCount} event(s)`);
      if (addedEvents.length > 0) parts.push(`added ${addedEvents.length} new`);
      status.textContent = parts.join(", ") + ".";
      status.className = "paste-status success";
      input.value = "";
      renderApp();
    });
  }

  parseBtn.addEventListener("click", () => {
    const input = document.getElementById(inputId);
    const status = document.getElementById(statusId);
    const text = input.value;
    if (!text.trim()) {
      status.textContent = "Paste some flight data first.";
      status.className = "paste-status error";
      return;
    }
    const defaultYear = state.start
      ? Number(state.start.split("-")[0])
      : new Date().getFullYear();

    // Each parser returns either null (no match) or an array of events.
    // parseFlights can return [] when it matches the format but extracts
    // nothing usable — treat that as "no match" so the chain falls through.
    const noMatch = (e) => !e || e.length === 0;
    let events = parseCommand(text, defaultYear);
    if (noMatch(events)) events = parseCruise(text, defaultYear);
    if (noMatch(events)) events = parseReservation(text, defaultYear);
    if (noMatch(events)) events = parseDeltaItinerary(text, defaultYear);
    if (noMatch(events)) events = parseFlights(text, defaultYear);
    // parseLooseEvent handles short single-line natural language; parseSmart
    // is the catch-all for free-form messy pastes when nothing else matched.
    if (noMatch(events)) events = parseLooseEvent(text, defaultYear);
    if (noMatch(events)) events = parseSmart(text, defaultYear);

    if (!events || events.length === 0) {
      status.textContent = "Could not detect anything to add. Try a flight paste or a command like 'add hotel on jul 7 for znz hotel'.";
      status.className = "paste-status error";
      return;
    }
    // targetId may be a static "__main__" (main tab — always the itinerary)
    // or a select element id (options tab — must pick an existing option,
    // or "__new__" to spin up a fresh one at parse time).
    let target;
    if (!targetId) {
      target = "__main__";
    } else {
      const sel = document.getElementById(targetId);
      target = sel?.value || "__new__";
      if (target === "__new__") {
        // Create a new option on demand and route the paste into it.
        const newOpt = createOption();
        target = newOpt.id;
      }
    }
    if (target === "__main__") {
      state.events.push(...events);
      // Expand the trip's date range so newly added events are visible.
      // Important when a fresh trip has no dates yet, or a reservation
      // falls outside the manually-entered range.
      let dateChanged = false;
      for (const e of events) {
        if (e.start && (!state.start || e.start < state.start)) {
          state.start = e.start;
          dateChanged = true;
        }
        if (e.end && (!state.end || e.end > state.end)) {
          state.end = e.end;
          dateChanged = true;
        }
      }
      if (dateChanged) {
        document.getElementById("trip-start").value = state.start || "";
        document.getElementById("trip-end").value = state.end || "";
      }
    } else {
      const opt = state.options.find(o => o.id === target);
      if (opt) opt.events.push(...events);
    }
    // If the paste contained an obvious total price ("Total: $1,234.56" /
    // "Total price for N travelers: US$12,571.00" / etc.), bundle the just-
    // added events into a pricing line item with that total. Only fires for
    // main-itinerary pastes; options pastes are speculative so we skip.
    let pricingNote = "";
    if (target === "__main__") {
      const total = detectTotalPrice(text);
      if (total != null && events.length > 0) {
        if (!Array.isArray(state.lineItems)) state.lineItems = [];
        state.lineItems.push({
          id: uid(),
          eventIds: events.map(e => e.id),
          label: null,
          total,
          overrides: {},
        });
        pricingNote = ` ($${total.toLocaleString()} added to Pricing)`;
      }
    }
    save();
    status.textContent = `Added ${events.length} event(s).${pricingNote}`;
    status.className = "paste-status success";
    input.value = "";
    renderApp();
  });

  document.getElementById(clearId).addEventListener("click", () => {
    document.getElementById(inputId).value = "";
    document.getElementById(statusId).textContent = "";
  });
}

wirePasteBlock({
  inputId: "paste-input",
  parseId: "paste-parse",
  smartId: "paste-smart",
  clearId: "paste-clear",
  statusId: "paste-status",
  targetId: "paste-target",
});
wirePasteBlock({
  inputId: "paste-input-main",
  parseId: "paste-parse-main",
  smartId: "paste-smart-main",
  updateId: "paste-update-main",
  clearId: "paste-clear-main",
  statusId: "paste-status-main",
  targetId: null,
});

document.getElementById("option-range-start").addEventListener("change", (e) => {
  state.optionRangeStart = e.target.value || null;
  saveLocal();   // option-staging picker is per-device UI
  renderOptions();
});
document.getElementById("option-range-end").addEventListener("change", (e) => {
  state.optionRangeEnd = e.target.value || null;
  saveLocal();
  renderOptions();
});

// --- boot ---

async function bootstrap() {
  const params = new URLSearchParams(window.location.search);
  migrateLegacy();

  // ?p=<token> = anonymous public viewer. Load the mirror, render
  // read-only, skip everything trip-registry / sync related.
  const publicToken = params.get("p");
  if (publicToken) {
    await whenFb();
    const blob = await window.fb.loadPublicTrip(publicToken);
    if (!blob) {
      document.body.innerHTML = '<div style="padding:40px;text-align:center;font-family:system-ui">This public link is invalid or has been disabled.</div>';
      return;
    }
    CURRENT_TRIP_ID = blob.slug || "public";
    CAN_EDIT = false;
    CAN_SEE_PRICING = false;
    IS_OWNER = false;
    Object.assign(state, blob);
    if (!state.options) state.options = [];
    ensureOptionGroups();
    state.activeView = "main";
    if (!state.shrunkDays) state.shrunkDays = [];
    delete state.expandedDays;
    BOOTING = false;
    renderApp();
    return;
  }

  const slug = params.get("id");
  let tripId = params.get("trip");

  if (slug && !tripId) {
    const list = readRegistry();
    const hit = list.find(t => t.slug === slug) || list.find(t => t.id === slug);
    if (hit) tripId = hit.id;
  }

  // Wait for Firebase auth state to settle before fetching, so loadTrip()
  // knows whether to read as the owner (with pricing) or as a viewer.
  await whenFb();
  await window.fb.whenAuthReady();

  // Pull the trip — from Firestore when signed in, from the worker for
  // anonymous ?v= viewers. fetchTrip handles both cases.
  let serverState = null;
  let serverIsOwner = false;
  let serverLevel = null; // "owner" | "viewer" | "viewer-pricing" | "editor" | null
  if (slug) {
    const res = await fetchTrip(slug);
    if (res.ok) { serverState = res.body; serverIsOwner = !!res.isOwner; serverLevel = res.level || null; }
    if (!tripId && serverState) {
      tripId = newTripId();
      upsertRegistry({
        id: tripId,
        slug,
        name: serverState.name || "Untitled trip",
        start: serverState.start || null,
        end: serverState.end || null,
      });
    }
  }

  if (!tripId) {
    const list = readRegistry();
    if (list.length > 0) {
      tripId = list[0].id;
    } else {
      window.location.replace("./");
      return;
    }
    const trip = readRegistry().find(t => t.id === tripId);
    const url = new URL(window.location.href);
    url.searchParams.delete("trip");
    url.searchParams.set("id", trip?.slug || tripId);
    window.location.replace(url.toString());
    return;
  }

  CURRENT_TRIP_ID = tripId;

  // Owner mode is now driven entirely by Firestore: loadTrip returns
  // isOwner=true when the signed-in user owns the trip (and the pricing
  // subcollection was readable). For anonymous ?v= viewers the worker
  // returns whatever the viewer token allows, with pricing fields inline.
  const serverHasPricing = serverState && PRICING_KEYS.some(k => serverState[k] !== undefined);
  IS_OWNER = serverIsOwner;
  // Owner OR editor can edit. Editor is a shared-user level that grants
  // full write access (still not owner — can't re-share or change ownership).
  CAN_EDIT = serverIsOwner || serverLevel === "editor";
  CAN_SEE_PRICING = CAN_EDIT || serverLevel === "viewer-pricing";

  // Newest-wins merge: server wins if its modifiedAt is newer-or-equal.
  // Pricing fields, however, are *only* taken from the server when the server
  // actually returned them. Otherwise we preserve whatever local has — the
  // server's silence about pricing means "I'm not authoritative on prices for
  // this request" (viewer mode, or auth dropped), not "prices have been
  // deleted." Without this guard, a single viewer-style read on the owner's
  // device would wipe pricing locally and then write empty pricing back to
  // the server on next sync.
  if (serverState) {
    const localRaw = localStorage.getItem(STORAGE_KEY_FOR(tripId));
    const localMtime = localRaw ? (safeParse(localRaw)?.modifiedAt || 0) : 0;
    const serverMtime = serverState.modifiedAt || 0;
    if (serverMtime >= localMtime) {
      let merged = { ...serverState };
      if (!serverHasPricing) {
        const local = localRaw ? safeParse(localRaw) : null;
        if (local) for (const k of PRICING_KEYS) {
          if (local[k] !== undefined) merged[k] = local[k];
        }
      }
      // Repair any UTF-8-as-Latin-1 mojibake left over from prior PowerShell-
      // mediated server PUTs before persisting.
      merged = demojibakeWalk(merged);
      localStorage.setItem(STORAGE_KEY_FOR(tripId), JSON.stringify(merged));
    }
  }

  load();
  if (!state.options) state.options = [];
  ensureOptionGroups();
  // Always land on the itinerary tab when opening a trip.
  state.activeView = "main";
  if (!state.shrunkDays) state.shrunkDays = [];
  delete state.expandedDays;

  BOOTING = false;
  renderApp();

  // If this device has no password yet but the trip exists on the worker,
  // first interaction (not load) will prompt. Push initial state to the
  // worker if the local copy is newer than what was on the server (covers
  // the case where they had a local trip from before the worker existed).
  if (CAN_EDIT && (!serverState || (state.modifiedAt || 0) > (serverState.modifiedAt || 0))) {
    scheduleSync();
  }
}

bootstrap();

// Re-render on orientation/viewport changes so the breakdown timeline
// (which uses a fixed day-width sized from clientWidth) doesn't stay
// clipped after rotating from portrait to landscape.
let resizeTimer = null;
window.addEventListener("resize", () => {
  if (resizeTimer) clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => {
    if (state.activeView === "main") render();
    else if (state.activeView === "options") renderOptions();
  }, 150);
});
