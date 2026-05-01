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
function STORAGE_KEY_FOR(id) { return TRIP_KEY_PREFIX + id; }
function STORAGE_KEY() { return STORAGE_KEY_FOR(CURRENT_TRIP_ID); }

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
  { key: "rental",     label: "Rental car", optional: true },
  { key: "activities", label: "Activities" },
];

const state = {
  name: "",
  start: null,
  end: null,
  events: [],
  segmentSize: "auto",
  tzAware: true,
  homeTz: "America/Vancouver",
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
    Object.assign(state, JSON.parse(raw));
  } catch (e) {
    seed();
  }
}

function save() {
  localStorage.setItem(STORAGE_KEY(), JSON.stringify(state));
  upsertRegistry({
    id: CURRENT_TRIP_ID,
    name: state.name || "Untitled trip",
    start: state.start || null,
    end: state.end || null,
  });
}

function seed() {
  state.name = "Zanzibar 2026";
  state.start = "2026-06-17";
  state.end = "2026-07-09";
  state.tzAware = true;
  state.homeTz = "America/Vancouver";
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
      start: "2026-06-17", startTime: "16:15", startTz: "America/Vancouver",
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

function computeDayTzMap(start, end, events, homeTz, tzAware) {
  const map = {};
  let cur = parseDay(start);
  const endDate = parseDay(end);
  let currentTz = homeTz;

  // Group flight arrivals by date (only used when tzAware).
  const arrByDate = {};
  if (tzAware) {
    for (const ev of events) {
      if (ev.lane === "flights" && ev.endTz) {
        (arrByDate[ev.end] ||= []).push(ev);
      }
    }
  }

  while (cur <= endDate) {
    const ds = toISO(cur);
    if (tzAware && arrByDate[ds] && arrByDate[ds].length) {
      const latest = arrByDate[ds].slice().sort((a, b) =>
        (a.endTime || "00:00").localeCompare(b.endTime || "00:00")).pop();
      currentTz = latest.endTz;
    }
    map[ds] = currentTz;
    cur = addDays(cur, 1);
  }
  return map;
}

// --- UTC bounds for an event ---

function eventBounds(ev, dayTzMap, homeTz) {
  const tzS = ev.startTz || dayTzMap[ev.start] || homeTz;
  const tzE = ev.endTz   || dayTzMap[ev.end]   || homeTz;
  const [sy, sm, sd] = ev.start.split("-").map(Number);
  const [ey, em, ed] = ev.end.split("-").map(Number);
  let sUtc, eUtc;
  if (ev.startTime) {
    const [h, mn] = ev.startTime.split(":").map(Number);
    sUtc = wallToUtc(sy, sm, sd, h, mn, tzS);
  } else {
    sUtc = wallToUtc(sy, sm, sd, 0, 0, dayTzMap[ev.start] || homeTz);
  }
  if (ev.endTime) {
    const [h, mn] = ev.endTime.split(":").map(Number);
    eUtc = wallToUtc(ey, em, ed, h, mn, tzE);
  } else {
    const next = toISO(addDays(parseDay(ev.end), 1));
    const [eyN, emN, edN] = next.split("-").map(Number);
    eUtc = wallToUtc(eyN, emN, edN, 0, 0, dayTzMap[ev.end] || homeTz);
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
  save();
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

function renderTimeline(container, rangeStart, rangeEnd, opts) {
  const { dayTzMap, homeTz, dayPx, compact, tzAware } = opts;
  const events = opts.events || state.events;
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

  const visible = events.filter(ev => !(ev.end < rangeStart || ev.start > rangeEnd));

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
    if (tzAware) cell.appendChild(el("span", "tz", tzShortName(tz, ds)));
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
      const { sUtc, eUtc } = eventBounds(ev, dayTzMap, homeTz);
      return { ev, sUtc, eUtc };
    });
    const packed = packRowsByUtc(items);
    const rowCount = Math.max(...packed.map(p => p.row + 1));
    laneArea.style.setProperty("--rows", rowCount);

    // Night shading (8pm–8am) per day — sits behind the event bars.
    appendNightShades(laneArea, totalDays, dayWidths, dayOffsetsPx, shrunkSet, dayUtcBounds);

    for (const { ev, sUtc, eUtc, row } of packed) {
      // tzAware: each day cell is its own local TZ — position events by
      // their declared local clock time within those cells (so a flight
      // at 16:15 PDT on day 17 lands at the 16:15 mark, not pushed by the
      // following day's TZ shift).
      // Non-tzAware: every day is in home TZ — convert UTC to home-TZ
      // fraction so a 11:00 Berlin arrival lands at 02:00 in the home grid.
      let leftFrac, rightFrac;
      if (tzAware) {
        ({ leftFrac, rightFrac } = eventLocalFracs(ev, rangeStart));
      } else {
        leftFrac = utcToFrac(sUtc, dayUtcBounds);
        rightFrac = utcToFrac(eUtc, dayUtcBounds);
      }
      leftFrac = Math.max(0, leftFrac);
      rightFrac = Math.min(totalDays, rightFrac);
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
      if (dayWidths) {
        const leftPx = fracToPx(leftFrac, dayWidths, dayOffsetsPx);
        const rightPx = fracToPx(rightFrac, dayWidths, dayOffsetsPx);
        bar.style.left = `${leftPx + 1}px`;
        bar.style.width = `${Math.max(2, rightPx - leftPx - 2)}px`;
      } else {
        // Compute in pixels so a 1-hour bar isn't expanded by the .event's
        // intrinsic min-content (padding+border would otherwise force ~19px
        // and cause sequential events to overlap).
        const leftPx = (leftFrac / totalDays) * stretchLanePx;
        const widthPx = Math.max(2, ((rightFrac - leftFrac) / totalDays) * stretchLanePx - 2);
        bar.style.left = `${leftPx + 1}px`;
        bar.style.width = `${widthPx}px`;
        // Drop text + padding when there isn't room — keeps the bar's actual
        // pixel width honest so adjacent events don't overlap.
        if (widthPx < 28) {
          bar.style.padding = "0";
          bar.dataset.narrow = "1";
        }
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

// --- breakdown segment sizing ---

function chooseSegmentSize(totalDays) {
  if (state.segmentSize !== "auto") return Number(state.segmentSize);
  if (totalDays <= 7) return totalDays;
  if (totalDays <= 60) return 7;
  return 14;
}

// --- top-level render ---

function render() {
  document.getElementById("trip-name").value = state.name;
  document.getElementById("trip-start").value = state.start || "";
  document.getElementById("trip-end").value = state.end || "";
  document.getElementById("segment-size").value = state.segmentSize;
  document.getElementById("tz-aware").checked = !!state.tzAware;

  const overview = document.getElementById("overview");
  const breakdown = document.getElementById("breakdown");

  if (!state.start || !state.end || state.end < state.start) {
    overview.innerHTML = `<div class="empty-state">Set start and end dates to build your timeline.</div>`;
    breakdown.innerHTML = "";
    document.getElementById("trip-length").textContent = "";
    return;
  }

  const totalDays = dayDiff(state.start, state.end) + 1;
  document.getElementById("trip-length").textContent = `${totalDays} day${totalDays === 1 ? "" : "s"}`;

  const homeTz = state.homeTz || "UTC";
  const tzAware = !!state.tzAware;
  const dayTzMap = computeDayTzMap(state.start, state.end, state.events, homeTz, tzAware);

  // Overview: stretch to panel width (no fixed day width).
  renderTimeline(overview, state.start, state.end, {
    dayTzMap, homeTz, dayPx: null, compact: totalDays > 14, tzAware,
  });

  // Breakdown: fixed day width — short segments stay physically short.
  breakdown.innerHTML = "";
  const segSize = chooseSegmentSize(totalDays);

  if (totalDays <= 7 || segSize >= totalDays) {
    breakdown.innerHTML = `<div class="empty-state">Trip is short — overview shows the full breakdown.</div>`;
    return;
  }

  // Pick a day width so the longest segment fills the breakdown panel,
  // and shorter segments stay proportionally narrower.
  const breakdownWidth = breakdown.clientWidth || 1200;
  const laneLabelW = parseInt(getComputedStyle(document.documentElement).getPropertyValue("--lane-label-w")) || 110;
  const dayPx = Math.max(60, Math.floor((breakdownWidth - laneLabelW - 4) / segSize));

  let cursor = parseDay(state.start);
  let idx = 1;
  while (toISO(cursor) <= state.end) {
    const segStart = toISO(cursor);
    const segEndDate = addDays(cursor, segSize - 1);
    const segEnd = toISO(segEndDate) > state.end ? state.end : toISO(segEndDate);

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
      dayTzMap, homeTz, dayPx, compact: false, tzAware,
    });

    cursor = addDays(cursor, segSize);
    idx++;
  }
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
  customSwatch.appendChild(colorPicker);
  customLabel.appendChild(customRadio);
  customLabel.appendChild(customSwatch);
  grid.appendChild(customLabel);
}

function openEventDialog(id, optionId) {
  form.reset();
  const titleEl = document.getElementById("event-dialog-title");
  const deleteBtn = document.getElementById("event-delete");

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
    form.elements.notes.value = ev.notes || "";
    deleteBtn.hidden = false;
  } else {
    const opt = optionId ? state.options.find(o => o.id === optionId) : null;
    titleEl.textContent = opt ? `Add event to "${opt.name}"` : "Add event";
    form.elements.id.value = "";
    form.elements.optionId.value = optionId || "";
    form.elements.lane.value = "flights";
    const defaultStart = optionId ? (state.optionRangeStart || state.start) : state.start;
    form.elements.start.value = defaultStart || "";
    form.elements.end.value = defaultStart || "";
    buildColorGrid(optionId ? "indigo" : "emerald");
    deleteBtn.hidden = true;
  }
  dialog.showModal();
}

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
  save();
  renderApp();
});
document.getElementById("tz-aware").addEventListener("change", (e) => {
  state.tzAware = e.target.checked;
  save();
  renderApp();
});
document.getElementById("add-event-btn").addEventListener("click", () => openEventDialog(null));
document.getElementById("clear-btn").addEventListener("click", () => {
  if (!confirm("Clear the trip and all events?")) return;
  localStorage.removeItem(STORAGE_KEY());
  for (const k of Object.keys(state)) delete state[k];
  seed();
  save();
  renderApp();
});

// --- flight paste/parser ---

const AIRPORT_TZ = {
  YVR: "America/Vancouver",
  SEA: "America/Los_Angeles",
  LAX: "America/Los_Angeles",
  JFK: "America/New_York",
  ORD: "America/Chicago",
  FRA: "Europe/Berlin",
  LHR: "Europe/London",
  CDG: "Europe/Paris",
  AMS: "Europe/Amsterdam",
  ZNZ: "Africa/Dar_es_Salaam",
  ARK: "Africa/Dar_es_Salaam",
  JRO: "Africa/Dar_es_Salaam",
  DAR: "Africa/Dar_es_Salaam",
  NBO: "Africa/Nairobi",
  ADD: "Africa/Addis_Ababa",
  CPT: "Africa/Johannesburg",
  JNB: "Africa/Johannesburg",
  SEZ: "Indian/Mahe",
  MRU: "Indian/Mauritius",
  CAI: "Africa/Cairo",
  DXB: "Asia/Dubai",
  DOH: "Asia/Qatar",
  IST: "Europe/Istanbul",
  // US domestic
  TPA: "America/New_York",
  MIA: "America/New_York",
  ATL: "America/New_York",
  BOS: "America/New_York",
  PHL: "America/New_York",
  IAD: "America/New_York",
  DCA: "America/New_York",
  CLT: "America/New_York",
  MCO: "America/New_York",
  FLL: "America/New_York",
  DTW: "America/Detroit",
  MSP: "America/Chicago",
  DFW: "America/Chicago",
  IAH: "America/Chicago",
  AUS: "America/Chicago",
  MSY: "America/Chicago",
  STL: "America/Chicago",
  MCI: "America/Chicago",
  MEM: "America/Chicago",
  DEN: "America/Denver",
  SLC: "America/Denver",
  PHX: "America/Phoenix",
  LAS: "America/Los_Angeles",
  SFO: "America/Los_Angeles",
  SAN: "America/Los_Angeles",
  PDX: "America/Los_Angeles",
  HNL: "Pacific/Honolulu",
  ANC: "America/Anchorage",
};

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
  // Fallback: first non-label, non-date, non-address-looking line.
  if (!title) {
    for (const line of lines) {
      if (/^(date|confirmation|arrive|arrival|depart|departure|guests?|hotel|address|check[\s-]?in|check[\s-]?out|reservation)/i.test(line)) continue;
      if (dowDateRx.test(line) && line.length < 40) continue;
      if (/^\d/.test(line)) continue;
      if (line.length > 80) continue;
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
  const tokenRx = /\b(?:(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s*)?(\d{1,2})(?:,?\s+(\d{4}))?\b/gi;
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
  if (/\b(rental car|car rental|rental|hertz|enterprise|avis|sixt|budget car)\b/i.test(title)) lane = "rental";
  else if (/\b(hotel|hotels|lodging|inn|resort|villa|airbnb|vrbo|cruise|cruises)\b/i.test(title)) lane = "lodging";
  else if (/\b(flight|flights)\b/i.test(title)) lane = "flights";
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
  const homeTz = state.homeTz || "UTC";
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
  const homeTz = state.homeTz || "UTC";
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

function renderOptions() {
  const range = getOptionRange();
  const list = document.getElementById("options-list");
  list.innerHTML = "";

  // Sync the range inputs.
  if (range) {
    document.getElementById("option-range-start").value = range.start;
    document.getElementById("option-range-end").value = range.end;
  }

  // Refresh the paste target dropdown.
  const target = document.getElementById("paste-target");
  if (target) {
    const prev = target.value;
    target.innerHTML = "";
    target.appendChild(new Option("My itinerary", "__main__"));
    for (const opt of state.options) {
      target.appendChild(new Option(opt.name, opt.id));
    }
    target.value = state.options.find(o => o.id === prev) ? prev : "__main__";
  }

  if (!range) {
    list.appendChild(el("div", "empty-state", "Set your trip dates first."));
    return;
  }

  if (state.options.length === 0) {
    list.appendChild(el("div", "empty-state",
      `No options yet. Click "+ New option" to stage an alternative for ${fmtShort(parseDay(range.start))} – ${fmtShort(parseDay(range.end))}.`));
    return;
  }

  const homeTz = state.homeTz || "UTC";
  const tzAware = !!state.tzAware;
  const dayTzMap = computeDayTzMap(state.start, state.end, state.events, homeTz, tzAware);

  for (const opt of state.options) {
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

    list.appendChild(card);
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
  // Always sync the topbar fields, regardless of which view is active.
  document.getElementById("trip-name").value = state.name || "";
  document.getElementById("trip-start").value = state.start || "";
  document.getElementById("trip-end").value = state.end || "";
  document.getElementById("tz-aware").checked = !!state.tzAware;

  // Tab buttons & panel visibility
  document.querySelectorAll(".tab-btn").forEach(b => {
    b.classList.toggle("active", b.dataset.tab === state.activeView);
  });
  document.getElementById("tab-main").hidden = state.activeView !== "main";
  document.getElementById("tab-options").hidden = state.activeView !== "options";

  if (state.activeView === "main") render();
  else renderOptions();
}

document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    state.activeView = btn.dataset.tab;
    save();
    renderApp();
  });
});

document.getElementById("add-option-btn").addEventListener("click", () => {
  const range = getOptionRange();
  if (!range) return;
  state.options.push({
    id: uid(),
    name: `Option ${state.options.length + 1}`,
    events: [],
  });
  save();
  renderApp();
});

function wirePasteBlock({ inputId, parseId, clearId, statusId, targetId }) {
  const parseBtn = document.getElementById(parseId);
  if (!parseBtn) return;
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
    // parseLooseEvent is intentionally last so it doesn't steal multi-line
    // flight/reservation pastes.
    if (noMatch(events)) events = parseLooseEvent(text, defaultYear);

    if (!events || events.length === 0) {
      status.textContent = "Could not detect anything to add. Try a flight paste or a command like 'add hotel on jul 7 for znz hotel'.";
      status.className = "paste-status error";
      return;
    }
    // targetId may be a static "__main__" (main tab) or a select element id
    // (options tab — user picks which option to add to).
    const target = targetId
      ? document.getElementById(targetId)?.value || "__main__"
      : "__main__";
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
    save();
    status.textContent = `Added ${events.length} event(s).`;
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
  clearId: "paste-clear",
  statusId: "paste-status",
  targetId: "paste-target",
});
wirePasteBlock({
  inputId: "paste-input-main",
  parseId: "paste-parse-main",
  clearId: "paste-clear-main",
  statusId: "paste-status-main",
  targetId: null,
});

document.getElementById("option-range-start").addEventListener("change", (e) => {
  state.optionRangeStart = e.target.value || null;
  save();
  renderOptions();
});
document.getElementById("option-range-end").addEventListener("change", (e) => {
  state.optionRangeEnd = e.target.value || null;
  save();
  renderOptions();
});

// --- boot ---

async function bootstrap() {
  const params = new URLSearchParams(window.location.search);

  // One-time migration of any legacy single-trip blob into the registry.
  const migratedId = migrateLegacy();

  let tripId = params.get("trip");

  // If no ?trip=, try to recover: explicit migrated id, then first registry
  // entry, otherwise punt to the trips landing page.
  if (!tripId) {
    if (migratedId) {
      tripId = migratedId;
    } else {
      const list = readRegistry();
      if (list.length > 0) {
        tripId = list[0].id;
      } else {
        window.location.replace("trips.html");
        return;
      }
    }
    const url = new URL(window.location.href);
    url.searchParams.set("trip", tripId);
    window.location.replace(url.toString());
    return;
  }

  CURRENT_TRIP_ID = tripId;

  const importPath = params.get("import");
  if (importPath) {
    try {
      const res = await fetch(importPath);
      if (!res.ok) throw new Error(`fetch failed ${res.status}`);
      const text = await res.text();
      JSON.parse(text);
      localStorage.setItem(STORAGE_KEY(), text);
      const url = new URL(window.location.href);
      url.searchParams.delete("import");
      window.location.replace(url.toString());
      return;
    } catch (e) {
      console.error("Import failed:", e);
      alert(`Import failed: ${e.message}`);
    }
  }
  // If the page has an embedded snapshot and we have no saved state yet,
  // pre-populate localStorage from it on first visit.
  if (typeof EMBEDDED_STATE !== "undefined" && !localStorage.getItem(STORAGE_KEY())) {
    try {
      localStorage.setItem(STORAGE_KEY(), JSON.stringify(EMBEDDED_STATE));
    } catch (e) {
      console.error("Embedded state load failed:", e);
    }
  }
  load();
  if (!state.options) state.options = [];
  if (!state.activeView) state.activeView = "main";
  if (!state.shrunkDays) state.shrunkDays = [];
  // Migrate stale field from earlier version.
  delete state.expandedDays;
  renderApp();
}

bootstrap();
