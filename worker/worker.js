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
//
// Set the password once with: wrangler secret put OWNER_PASSWORD

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

    return json({ error: "not found" }, { status: 404 });
  },
};
