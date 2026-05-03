// Cloudflare Worker — anonymous trip snapshot store.
//
// POST /  body=JSON  -> 201 { "id": "<id>" }
// GET  /<id>         -> 200 stored JSON, or 404
//
// The KV binding "SNAPSHOTS" is configured in wrangler.toml.

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Access-Control-Max-Age": "86400",
};

const MAX_BODY_BYTES = 256 * 1024;
const ID_BYTES = 12;

function newId() {
  const bytes = new Uint8Array(ID_BYTES);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
}

function json(body, init = {}) {
  return new Response(JSON.stringify(body), {
    ...init,
    headers: { "Content-Type": "application/json", ...CORS, ...(init.headers || {}) },
  });
}

export default {
  async fetch(req, env) {
    if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS });

    const url = new URL(req.url);
    const path = url.pathname.replace(/^\/+|\/+$/g, "");

    if (req.method === "POST" && path === "") {
      const ct = req.headers.get("content-type") || "";
      if (!ct.includes("application/json")) {
        return json({ error: "expected application/json" }, { status: 415 });
      }
      const body = await req.text();
      if (body.length > MAX_BODY_BYTES) {
        return json({ error: "snapshot too large" }, { status: 413 });
      }
      try { JSON.parse(body); } catch (e) {
        return json({ error: "invalid JSON" }, { status: 400 });
      }
      const id = newId();
      await env.SNAPSHOTS.put(id, body);
      return json({ id }, { status: 201, headers: { "Location": `/${id}` } });
    }

    if (req.method === "GET" && /^[0-9a-f]+$/.test(path)) {
      const value = await env.SNAPSHOTS.get(path);
      if (value == null) return json({ error: "not found" }, { status: 404 });
      return new Response(value, {
        status: 200,
        headers: { "Content-Type": "application/json", ...CORS },
      });
    }

    return json({ error: "not found" }, { status: 404 });
  },
};
