// Firebase Cloud Function — smart-parse endpoint.
//
// Replaces the Cloudflare worker's POST /smart-parse. Callable function: the
// Firebase Functions SDK auto-attaches the caller's ID token and verifies it,
// so we just check `request.auth.uid` against the allowlist and call Anthropic.

import { onCall, HttpsError } from "firebase-functions/v2/https";
import { onSchedule } from "firebase-functions/v2/scheduler";
import { defineSecret } from "firebase-functions/params";

const ANTHROPIC_API_KEY = defineSecret("ANTHROPIC_API_KEY");
const ALLOWED_UIDS = defineSecret("ALLOWED_UIDS");

export const smartParse = onCall(
  { secrets: [ANTHROPIC_API_KEY, ALLOWED_UIDS], region: "us-central1", invoker: "public" },
  async (request) => {
    if (!request.auth) throw new HttpsError("unauthenticated", "Sign in required.");
    const allowed = ALLOWED_UIDS.value().split(",").map(s => s.trim()).filter(Boolean);
    if (allowed.length && !allowed.includes(request.auth.uid)) {
      throw new HttpsError("permission-denied", "UID not allowlisted.");
    }
    const apiKey = (ANTHROPIC_API_KEY.value() || "").trim();
    if (!apiKey) throw new HttpsError("failed-precondition", "ANTHROPIC_API_KEY not configured.");

    const body = request.data || {};
    const text = (body.text || "").toString();
    const mode = body.mode || null;
    if (mode !== "hotel-from-url" && !text.trim()) {
      throw new HttpsError("invalid-argument", "empty text");
    }
    if (text.length > 100_000) throw new HttpsError("invalid-argument", "text too long");
    const tripStart = body.tripStart || null;
    const tripEnd = body.tripEnd || null;
    const existingEvents = Array.isArray(body.existingEvents) ? body.existingEvents : null;

    try {
      if (mode === "hotel-compare") {
        return await parseHotelForCompare(text, tripStart, tripEnd, apiKey);
      }
      if (mode === "hotel-from-url") {
        const url = (body.url || "").toString();
        if (!/^https?:\/\//i.test(url)) throw new HttpsError("invalid-argument", "url must start with http(s)://");
        const fetched = await fetchHotelPage(url);
        if (!fetched.ok) throw new HttpsError("unavailable", fetched.error);
        const result = await parseHotelForCompare(fetched.text, tripStart, tripEnd, apiKey);
        if (!result.url) result.url = url;
        return result;
      }
      return await smartParseEvents(text, tripStart, tripEnd, apiKey, existingEvents);
    } catch (e) {
      if (e instanceof HttpsError) throw e;
      throw new HttpsError("internal", `parse failed: ${e.message}`);
    }
  },
);

async function fetchHotelPage(url) {
  let res;
  try {
    res = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
      },
      redirect: "follow",
    });
  } catch (e) {
    return { ok: false, error: `fetch failed: ${e.message}` };
  }
  if (!res.ok) return { ok: false, error: `${url} returned ${res.status}` };
  const ct = res.headers.get("content-type") || "";
  if (!/html|text/i.test(ct)) return { ok: false, error: `unsupported content-type: ${ct}` };
  let html = await res.text();
  if (html.length > 500_000) html = html.slice(0, 500_000);
  let text = html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<\/?[a-z][^>]*>/gi, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ")
    .trim();
  if (text.length < 200) {
    return { ok: false, error: "page returned too little text — likely JS-rendered or bot-blocked. Try pasting the visible text instead." };
  }
  if (text.length > 80_000) text = text.slice(0, 80_000);
  return { ok: true, text };
}

async function parseHotelForCompare(text, tripStart, tripEnd, apiKey) {
  const SCHEMA = `
Return ONLY a JSON object matching this shape — fields you can't determine
should be null (don't make up values):

{
  "name": "string (e.g. 'Courtyard Mexico City Airport' or 'Bela Vista Airbnb')",
  "platform": "Marriott | Hilton | Hyatt | Booking.com | Hotels.com | Airbnb | Vrbo | Expedia | Other | null",
  "checkIn":  "YYYY-MM-DD or null",
  "checkOut": "YYYY-MM-DD or null",
  "nights":   "integer or null",
  "neighborhood": "string or null (e.g. 'Bela Vista, Sao Paulo')",
  "rating":   "string or null (e.g. '4.96', '4.5/5', '4 stars')",
  "pricePerNight": "number or null (USD-equivalent if currency is shown)",
  "totalPrice":   "number or null (this is the headline figure — for the full stay across all rooms)",
  "currency":     "string or null (USD, BRL, EUR, etc.)",
  "roomCount":    "integer or null (number of rooms or units booked)",
  "roomType":     "string or null (e.g. 'Deluxe Ocean View', 'Two-bedroom suite', 'Studio')",
  "amenities":    "array of short strings (e.g. ['Free WiFi','Breakfast','Pool'])",
  "cancellation": "string or null (short summary)",
  "sourceUrl":    "string or null (the URL of the page/listing where the price was found — e.g. the Booking.com or Expedia listing URL)",
  "websiteUrl":   "string or null (the hotel's own website URL if visible, separate from the booking-platform URL)",
  "notes":        "string or null (anything else worth comparing — view, room type, host, etc.)"
}

Output ONLY the JSON object, no commentary, no markdown fencing.`;

  const ctx = (tripStart && tripEnd)
    ? `Trip date range: ${tripStart} to ${tripEnd}. Use it to disambiguate years.`
    : `Trip dates not yet set — use any explicit year, otherwise current year.`;

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 1000,
      system: [{
        type: "text",
        text: "You extract structured hotel/lodging details from listings, emails, and web copy for a side-by-side comparison table. " + SCHEMA,
        cache_control: { type: "ephemeral" },
      }],
      messages: [
        { role: "user", content: `${ctx}\n\nHotel listing text:\n${text}` },
      ],
    }),
  });
  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    throw new Error(`anthropic ${res.status}: ${errText.slice(0, 200)}`);
  }
  const data = await res.json();
  const content = data.content?.[0]?.text || "";
  const clean = content.replace(/^```(?:json)?\s*|\s*```$/g, "").trim();
  try { return JSON.parse(clean); }
  catch { throw new Error(`model returned non-JSON: ${clean.slice(0, 200)}`); }
}

async function smartParseEvents(text, tripStart, tripEnd, apiKey, existingEvents) {
  const updateMode = Array.isArray(existingEvents) && existingEvents.length > 0;

  const SCHEMA_HINT_CREATE = `
Return ONLY a JSON object matching this shape:
{
  "events": [
    {
      "title": "string (concise — for flights include the flight number, e.g. 'KQ 427 SEA → GRU'; for lodging use the hotel name)",
      "lane": "flights | lodging | activities | rental | location",
      "start": "YYYY-MM-DD",
      "end":   "YYYY-MM-DD (same as start for single-day events)",
      "startTime": "HH:MM (24-hour, optional — only for flights and timed events)",
      "endTime":   "HH:MM (24-hour, optional)",
      "notes": "string (optional — airline, cabin, flight numbers, address, etc.)"
    }
  ],
  "totalPrice": number (optional — the total cost across all parsed events, if a total appears in the text)
}

Rules:
- One event per flight leg, including connection segments. If the email shows
  seat assignments like "SEA - MEX: 3E / MEX - GRU: 3D", or two flight
  numbers ("AM495, AM14"), or "1 stop · 20h · ..." — produce TWO events
  (SEA→MEX and MEX→GRU), not one combined SEA→GRU. Same for return.
- BETWEEN connection segments, add a layover event:
    title: "<CODE> layover" (e.g. "MEX layover"),
    lane: "flights",
    start = arrival date of the previous leg, startTime = its arrival time,
    end   = departure date of the next leg, endTime   = its departure time,
    notes: "<duration> layover" if duration is shown.
  Skip the layover if either bracketing leg has no time info — a date-only
  layover bar isn't useful.
- For connection segments where intermediate timing isn't in the email,
  use the same date as the outer leg and leave startTime/endTime off; user
  can adjust later.
- Don't combine round-trips into one event — outbound and return are
  always separate events (and each may itself have multiple legs).
- Hotels: one event spanning check-in date to check-out date. Lane = "lodging".
- Airbnb / Vrbo / short-term-rental pastes: title as "<neighborhood-or-city> Airbnb"
  (e.g. "Bela Vista Airbnb"). If only a host name is available, use "<host>'s Airbnb".
- Use the trip context dates to disambiguate years for ambiguous dates (e.g. "Aug 13" without a year).
- For flight titles ALWAYS include the flight number when one is present: "KQ 427 SEA → GRU" (airline code + number, space, then IATA pair). Never just "Seattle to Sao Paulo" or a city name. If multiple flight numbers appear for the same leg, pick the operating one.
- Skip emails that aren't about travel reservations — return events: [].
- Output ONLY the JSON object, no commentary, no markdown fencing.`;

  const SCHEMA_HINT_UPDATE = `
You are merging new email content into a trip that already has events. Match
each fact in the email against the existing event most likely to be the same
real-world thing (same route + similar date for flights, same hotel name + date
for lodging, etc.) and emit either an update or a new event.

Return ONLY a JSON object matching this shape:
{
  "updates": [
    {
      "id": "string (the existing event id you are updating)",
      "fields": {
        "start": "YYYY-MM-DD (optional — only include fields that changed)",
        "end":   "YYYY-MM-DD (optional)",
        "startTime": "HH:MM (optional, 24-hour)",
        "endTime":   "HH:MM (optional)",
        "title": "string (optional — only if a clearly better title is available)",
        "notes": "string (optional — append or replace, your judgment)"
      }
    }
  ],
  "newEvents": [
    { same shape as the create-mode event objects — used only for legs/items
      that genuinely don't match any existing event }
  ],
  "totalPrice": number (optional — only if the new email contains a total)
}

Rules:
- STRONGLY prefer updates over duplicates. Match by route (same airport pair,
  ignoring direction order) within the same trip date range. If an existing
  flight has the same departure-airport → arrival-airport as something in the
  email and falls within ~3 days of it, treat them as the same flight and
  emit an update — even if start dates appear to differ (the new email may
  have a more accurate date you should overwrite).
- For connection layover events ("MEX layover" etc.) that don't yet exist in
  the existing list, add them in newEvents. Layover events have:
    title: "<CODE> layover", lane: "flights",
    start/startTime = previous leg's arrival, end/endTime = next leg's departure.
- Only include fields in updates.fields that have new/better data — don't echo
  unchanged values.
- Same flight-leg / connection / hotel rules as create mode.
- Output ONLY the JSON object, no commentary, no markdown fencing.`;

  const ctx = (tripStart && tripEnd)
    ? `Trip date range: ${tripStart} to ${tripEnd}.`
    : `Trip dates not yet set — use any explicit year in the text, otherwise current year.`;

  const existingCtx = updateMode
    ? `\n\nExisting events on this trip (id, title, lane, start, end, startTime, endTime):\n${
        existingEvents.map(e => JSON.stringify({
          id: e.id, title: e.title, lane: e.lane,
          start: e.start, end: e.end,
          startTime: e.startTime || null, endTime: e.endTime || null,
        })).join("\n")
      }`
    : "";

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 1500,
      system: [
        {
          type: "text",
          text: "You extract structured travel-event data from confirmation emails and itineraries. " + (updateMode ? SCHEMA_HINT_UPDATE : SCHEMA_HINT_CREATE),
          cache_control: { type: "ephemeral" },
        },
      ],
      messages: [
        { role: "user", content: `${ctx}${existingCtx}\n\nEmail/itinerary text:\n${text}` },
      ],
    }),
  });

  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    throw new Error(`anthropic ${res.status}: ${errText.slice(0, 200)}`);
  }
  const data = await res.json();
  const content = data.content?.[0]?.text || "";
  const clean = content.replace(/^```(?:json)?\s*|\s*```$/g, "").trim();
  let parsed;
  try { parsed = JSON.parse(clean); }
  catch { throw new Error(`model returned non-JSON: ${clean.slice(0, 200)}`); }
  if (updateMode) {
    if (!Array.isArray(parsed.updates)) parsed.updates = [];
    if (!Array.isArray(parsed.newEvents)) parsed.newEvents = [];
  } else {
    if (!Array.isArray(parsed.events)) parsed.events = [];
  }
  return parsed;
}

// ===========================================================================
// QQQ weekly-call "rebound" signal — runs every Friday at 3:00, 3:45, and 3:55
// PM ET. Three checks so you can't miss it and the later ones confirm the
// signal held as prices firm toward the 4pm close. Each run re-reads the live
// prices and fires a phone push (via ntfy.sh) ONLY if still triggered:
// this (near-complete) week's QQQ move < -3% AND VXN >= 25. Silent otherwise.
// ===========================================================================
const NTFY_TOPIC = "danaqqqreboundatm3847482824";

export const qqqReboundSignal = onSchedule(
  { schedule: "0,45,55 15 * * 5", timeZone: "America/New_York", region: "us-central1" },
  async () => {
    const UA = { "User-Agent": "Mozilla/5.0" };
    const getJSON = async (u) => (await fetch(u, { headers: UA })).json();
    const now = Math.floor(Date.now() / 1000);

    // --- QQQ: most recently COMPLETED Monday-open -> Friday-close week ---
    const qd = await getJSON(`https://query1.finance.yahoo.com/v8/finance/chart/QQQ?period1=${now - 30 * 86400}&period2=${now}&interval=1d`);
    const r = qd.chart.result[0], ts = r.timestamp, q = r.indicators.quote[0];
    const days = [];
    for (let i = 0; i < ts.length; i++) {
      if (q.open[i] == null || q.close[i] == null) continue;
      days.push({ t: ts[i], open: q.open[i], close: q.close[i] });
    }
    const monKey = (t) => {
      const d = new Date(t * 1000), dow = d.getUTCDay(), diff = dow === 0 ? 6 : dow - 1;
      return new Date(d.getTime() - diff * 86400000).toISOString().slice(0, 10);
    };
    const weeks = {};
    for (const d of days) (weeks[monKey(d.t)] = weeks[monKey(d.t)] || []).push(d);
    // Friday 3pm: use the CURRENT week (Mon open -> latest/Friday price).
    // At ~3pm the daily endpoint includes today's forming bar, so its "close"
    // is the live price — a near-final read of the week ending today.
    const keys = Object.keys(weeks).sort();
    let wkRet = null, wkLabel = null;
    if (keys.length) {
      const a = weeks[keys[keys.length - 1]].sort((x, y) => x.t - y.t);
      wkRet = (a[a.length - 1].close / a[0].open - 1) * 100;
      wkLabel = keys[keys.length - 1];
    }

    // --- VXN: latest close ---
    const vd = await getJSON(`https://query1.finance.yahoo.com/v8/finance/chart/%5EVXN?period1=${now - 14 * 86400}&period2=${now}&interval=1d`);
    const vc = vd.chart.result[0].indicators.quote[0].close.filter((x) => x != null);
    const vxn = vc[vc.length - 1];

    const etTime = new Intl.DateTimeFormat("en-US", { timeZone: "America/New_York", hour: "numeric", minute: "2-digit" }).format(new Date());
    const fire = wkRet != null && wkRet < -3 && vxn >= 25;
    console.log(`[qqqReboundSignal] ${etTime} ET, week ${wkLabel}: QQQ ${wkRet?.toFixed(2)}%, VXN ${vxn?.toFixed(2)} -> signal ${fire ? "ON" : "off"}`);
    if (!fire) return;

    await fetch("https://ntfy.sh/" + NTFY_TOPIC, {
      method: "POST",
      headers: { "Title": `Rebound signal ON (${etTime} ET) - buy QQQ call`, "Priority": "high", "Tags": "chart_with_upwards_trend" },
      body: `Buy a weekly ATM QQQ call Monday at the open.\nAs of ${etTime} ET: this week QQQ ${wkRet.toFixed(1)}%, VXN ${vxn.toFixed(1)} — both past threshold.`,
    });
    console.log("[qqqReboundSignal] push sent.");
  },
);

// ===========================================================================
// Stock "buy-the-dip / fear" signals — runs each weekday at 4:30 PM ET (after
// close). Pushes only on the DAY a threshold is newly CROSSED (so no daily
// repeats while a condition persists):
//   - VIX crosses above 30 (elevated fear) or 35 (extreme fear)  -> buy signal
//   - SPY crosses >10% or >20% below its 52-week high            -> dip signal
//   - VIX drops below 14 (complacency)                           -> heads-up
// Historical stats are from QQQ/SPY 2010-2026 (a bull-heavy regime).
// ===========================================================================
async function pushNtfy(title, body, tags) {
  await fetch("https://ntfy.sh/" + NTFY_TOPIC, {
    method: "POST",
    headers: { "Title": title, "Priority": "high", "Tags": tags || "chart_with_upwards_trend" },
    body,
  });
}

export const stockDipSignals = onSchedule(
  { schedule: "30 16 * * 1-5", timeZone: "America/New_York", region: "us-central1" },
  async () => {
    const UA = { "User-Agent": "Mozilla/5.0" };
    const getJSON = async (u) => (await fetch(u, { headers: UA })).json();
    const now = Math.floor(Date.now() / 1000);

    // VIX: last two settled daily closes
    const vd = await getJSON(`https://query1.finance.yahoo.com/v8/finance/chart/%5EVIX?period1=${now - 25 * 86400}&period2=${now}&interval=1d`);
    const vC = vd.chart.result[0].indicators.quote[0].close.filter((x) => x != null);
    const vToday = vC[vC.length - 1], vYst = vC[vC.length - 2];

    // SPY: drawdown from trailing 252-day high, today vs yesterday
    const sd = await getJSON(`https://query1.finance.yahoo.com/v8/finance/chart/SPY?period1=${now - 500 * 86400}&period2=${now}&interval=1d`);
    const sC = sd.chart.result[0].indicators.quote[0].close.filter((x) => x != null);
    const ddAt = (idx) => {
      let hi = 0;
      for (let j = Math.max(0, idx - 251); j <= idx; j++) hi = Math.max(hi, sC[j]);
      return (sC[idx] / hi - 1) * 100;
    };
    const ddToday = ddAt(sC.length - 1), ddYst = ddAt(sC.length - 2);

    const alerts = [];
    // --- Fear (VIX crossing UP) ---
    if (vYst < 35 && vToday >= 35)
      alerts.push(["Buy signal: extreme fear (VIX " + vToday.toFixed(0) + ")",
        `VIX crossed above 35 — extreme fear. Since 2010, SPY/VTI averaged ~+14% over the next 3 months from here, positive ~97% of the time. Historically a strong "back up the truck" zone. (Caveat: in 2008-style bears these took longer to pay off.)`, "rotating_light"]);
    else if (vYst < 30 && vToday >= 30)
      alerts.push(["Buy signal: elevated fear (VIX " + vToday.toFixed(0) + ")",
        `VIX crossed above 30 — elevated fear. Since 2010, SPY/VTI averaged ~+9% over the next 3 months from here, ~88% positive. Consider adding to broad-market positions (SPY/VTI).`, "chart_with_upwards_trend"]);
    // --- Dip (SPY drawdown crossing DOWN) ---
    if (ddYst > -20 && ddToday <= -20)
      alerts.push(["Buy signal: SPY " + ddToday.toFixed(0) + "% off its high",
        `SPY closed >20% below its 52-week high (bear territory). Since 2010 that led to ~+9% over the next 3 months, ~87% positive — but in real bear markets (2008) the true bottom came later, so scale in rather than going all at once.`, "chart_with_downwards_trend"]);
    else if (ddYst > -10 && ddToday <= -10)
      alerts.push(["Buy signal: SPY " + ddToday.toFixed(0) + "% off its high",
        `SPY closed >10% below its 52-week high (a correction). Since 2010, 10-20% dips led to ~+6% over the next 3 months, ~76% positive. A solid broad-market dip-buying zone.`, "chart_with_downwards_trend"]);
    // --- Complacency (VIX crossing BELOW 14) ---
    if (vYst >= 14 && vToday < 14)
      alerts.push(["Heads-up: VIX under 14 (complacency)",
        `VIX dropped below 14. WHAT THIS MEANS: VIX is the "fear index" — expected 30-day volatility of the S&P 500. Under 14 is unusually LOW, meaning the market is very calm and investors are pricing in little risk (option premiums are cheap). It is NOT a sell signal, but historically a WEAK time to deploy fresh cash: since 2010, SPY averaged only ~+1.7% over the next 3 months from VIX<14 — the lowest of any regime. Translation: don't chase an all-time-high, complacent tape; be patient with new buys and keep some dry powder.`, "warning"]);

    console.log(`[stockDipSignals] VIX ${vYst?.toFixed(1)}->${vToday?.toFixed(1)}, SPY dd ${ddYst?.toFixed(1)}%->${ddToday?.toFixed(1)}% | ${alerts.length} alert(s)`);
    for (const [title, body, tags] of alerts) await pushNtfy(title, body, tags);
  },
);
