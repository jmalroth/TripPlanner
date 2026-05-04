// Cloudflare Worker — trip storage + read.
//
// Trips live in KV under their slug. The worker handles GET (public read,
// pricing fields stripped unless authorized) and PUT (write, requires the
// OWNER_PASSWORD secret in the Authorization: Bearer <password> header).
//
// Routes:
//   GET  /<slug>             -> 200 JSON trip (pricing stripped) | 404
//   GET  /<slug>  + auth     -> 200 JSON trip (full, with pricing) | 404
//   PUT  /<slug>  + auth     -> 200 { ok: true } | 401 | 413 | 400
//   POST /snap               -> 201 { id } (anonymous one-off snapshot)
//   GET  /snap/<id>          -> 200 JSON snapshot | 404
//   POST /smart-parse + auth -> 200 { events, totalPrice? } | 401 | 502
//
// Secrets:
//   wrangler secret put OWNER_PASSWORD     (write auth)
//   wrangler secret put ANTHROPIC_API_KEY  (smart-parse only)

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
  "Access-Control-Max-Age": "86400",
};

const MAX_BODY_BYTES = 256 * 1024;
const SLUG_RE = /^[a-z0-9][a-z0-9-]{0,80}$/;
const SNAP_ID_RE = /^[0-9a-f]{8,64}$/;
const TRIP_PREFIX = "trip:";
const SNAP_PREFIX = "snap:";

const PRICING_KEYS = ["lineItems", "priceSplit", "priceToken", "viewerToken"];

function newId() {
  const bytes = new Uint8Array(12);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
}

function json(body, init = {}) {
  return new Response(JSON.stringify(body), {
    ...init,
    headers: { "Content-Type": "application/json", ...CORS, ...(init.headers || {}) },
  });
}

function stripPricing(obj) {
  const copy = { ...obj };
  for (const k of PRICING_KEYS) delete copy[k];
  return copy;
}

function isOwner(req, env) {
  if (!env.OWNER_PASSWORD) return false;
  const auth = req.headers.get("Authorization") || "";
  const supplied = auth.replace(/^Bearer\s+/i, "").trim();
  if (!supplied) return false;
  // Constant-time compare so a timing attack can't tease out the password.
  if (supplied.length !== env.OWNER_PASSWORD.length) return false;
  let diff = 0;
  for (let i = 0; i < supplied.length; i++) {
    diff |= supplied.charCodeAt(i) ^ env.OWNER_PASSWORD.charCodeAt(i);
  }
  return diff === 0;
}

export default {
  async fetch(req, env) {
    if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS });

    const url = new URL(req.url);
    const path = url.pathname.replace(/^\/+|\/+$/g, "");
    const parts = path.split("/").filter(Boolean);

    // Snapshot creation.
    if (req.method === "POST" && parts[0] === "snap" && parts.length === 1) {
      const body = await req.text();
      if (body.length > MAX_BODY_BYTES) return json({ error: "snapshot too large" }, { status: 413 });
      try { JSON.parse(body); } catch { return json({ error: "invalid JSON" }, { status: 400 }); }
      const id = newId();
      await env.SNAPSHOTS.put(SNAP_PREFIX + id, body);
      return json({ id }, { status: 201 });
    }

    // Snapshot read.
    if (req.method === "GET" && parts[0] === "snap" && parts.length === 2 && SNAP_ID_RE.test(parts[1])) {
      const value = await env.SNAPSHOTS.get(SNAP_PREFIX + parts[1]);
      if (value == null) return json({ error: "not found" }, { status: 404 });
      return new Response(value, { status: 200, headers: { "Content-Type": "application/json", ...CORS } });
    }

    // Trip routes — single-segment slug.
    if (parts.length === 1 && SLUG_RE.test(parts[0])) {
      const slug = parts[0];
      const key = TRIP_PREFIX + slug;
      const owner = isOwner(req, env);

      if (req.method === "GET") {
        const value = await env.SNAPSHOTS.get(key);
        if (value == null) return json({ error: "not found" }, { status: 404 });
        const parsed = JSON.parse(value);
        if (owner) return json(parsed);
        // ?v=<token> grants pricing-read access (read-only) if it matches the
        // viewerToken stored on the trip.
        const supplied = url.searchParams.get("v");
        if (supplied && parsed.viewerToken && supplied === parsed.viewerToken) {
          return json(parsed);
        }
        return json(stripPricing(parsed));
      }

      if (req.method === "PUT") {
        if (!owner) return json({ error: "unauthorized" }, { status: 401 });
        const body = await req.text();
        if (body.length > MAX_BODY_BYTES) return json({ error: "too large" }, { status: 413 });
        try { JSON.parse(body); } catch { return json({ error: "invalid JSON" }, { status: 400 }); }
        await env.SNAPSHOTS.put(key, body);
        return json({ ok: true });
      }
    }

    if (req.method === "POST" && parts.length === 1 && parts[0] === "smart-parse") {
      if (!isOwner(req, env)) return json({ error: "unauthorized" }, { status: 401 });
      if (!env.ANTHROPIC_API_KEY) return json({ error: "ANTHROPIC_API_KEY not configured" }, { status: 500 });
      let body;
      try { body = await req.json(); } catch { return json({ error: "invalid JSON" }, { status: 400 }); }
      const text = (body.text || "").toString();
      if (!text.trim()) return json({ error: "empty text" }, { status: 400 });
      if (text.length > 100_000) return json({ error: "text too long" }, { status: 413 });
      const tripStart = body.tripStart || null;
      const tripEnd = body.tripEnd || null;
      const existingEvents = Array.isArray(body.existingEvents) ? body.existingEvents : null;
      try {
        const result = await smartParse(text, tripStart, tripEnd, (env.ANTHROPIC_API_KEY || "").trim(), existingEvents);
        return json(result);
      } catch (e) {
        return json({ error: `parse failed: ${e.message}` }, { status: 502 });
      }
    }

    return json({ error: "not found" }, { status: 404 });
  },
};

// Ask Claude to extract structured travel events from a pasted email/web body.
// Returns { events: [...], totalPrice?: number }. Uses Haiku for cost (~$0.002
// per parse) and prompt-cached system prompt to keep tokens cheap.
async function smartParse(text, tripStart, tripEnd, apiKey, existingEvents) {
  const updateMode = Array.isArray(existingEvents) && existingEvents.length > 0;

  const SCHEMA_HINT_CREATE = `
Return ONLY a JSON object matching this shape:
{
  "events": [
    {
      "title": "string (concise — e.g. 'SEA → GRU' for flights, hotel name for lodging)",
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
- For connection segments where intermediate timing isn't in the email,
  use the same date as the outer leg and leave startTime/endTime off; user
  can adjust later.
- Don't combine round-trips into one event — outbound and return are
  always separate events (and each may itself have multiple legs).
- Hotels: one event spanning check-in date to check-out date. Lane = "lodging".
- Use the trip context dates to disambiguate years for ambiguous dates (e.g. "Aug 13" without a year).
- For flight titles use IATA codes: "SEA → GRU", not "Seattle to Sao Paulo".
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
- Prefer updates over duplicates. If the email is about the same flights you
  already have, ALL output should be in "updates" with empty "newEvents".
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
  // Strip optional markdown fencing in case the model adds it despite the rule.
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
