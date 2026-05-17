// Registry sync — keeps the local list of trips (id/slug/name/dates) mirrored
// to the worker so the same list is reachable from any device via dana.html.
//
// Wire-up:
//   - index.html calls window.RegistrySync.schedulePush() after writeRegistry().
//   - app.js calls window.RegistrySync.schedulePush() from save() (after upsert).
//   - index.html calls window.RegistrySync.pullAndMerge() on load if a password
//     is held; merges remote entries into the local registry.

(function () {
  const SNAPSHOT_API = "https://trip-snapshots.daner1231.workers.dev";
  const REGISTRY_KEY = "trip-builder-trips";
  const TRIP_KEY_PREFIX = "trip-builder-trip-";
  const OWNER_PASSWORD_KEY = "trip-builder-owner-password";

  function getPassword() {
    try { return localStorage.getItem(OWNER_PASSWORD_KEY) || null; }
    catch { return null; }
  }
  function readRegistry() {
    try { return JSON.parse(localStorage.getItem(REGISTRY_KEY) || "[]"); }
    catch { return []; }
  }
  function writeRegistry(list) {
    localStorage.setItem(REGISTRY_KEY, JSON.stringify(list));
  }

  let pushTimer = null;
  function schedulePush() {
    if (!getPassword()) return;
    clearTimeout(pushTimer);
    pushTimer = setTimeout(pushNow, 1200);
  }

  // Pull viewerToken out of each trip's local blob if present, so dana.html
  // can render both share URLs (public and pricing) without extra fetches.
  function decorateForPush(list) {
    return list.map(t => {
      const out = { id: t.id, slug: t.slug, name: t.name, start: t.start, end: t.end };
      try {
        const raw = localStorage.getItem(TRIP_KEY_PREFIX + t.id);
        if (raw) {
          const blob = JSON.parse(raw);
          if (blob?.viewerToken) out.viewerToken = blob.viewerToken;
        }
      } catch { /* ignore */ }
      return out;
    });
  }

  async function pushNow() {
    const pw = getPassword();
    if (!pw) return { ok: false, status: 401 };
    const trips = decorateForPush(readRegistry());
    try {
      const res = await fetch(`${SNAPSHOT_API}/registry`, {
        method: "PUT",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${pw}`,
        },
        body: JSON.stringify({ trips, updatedAt: Date.now() }),
      });
      return { ok: res.ok, status: res.status };
    } catch (e) {
      return { ok: false, status: 0, error: e.message };
    }
  }

  async function fetchRemote() {
    const pw = getPassword();
    if (!pw) return { ok: false, status: 401 };
    try {
      const res = await fetch(`${SNAPSHOT_API}/registry`, {
        headers: { "Authorization": `Bearer ${pw}` },
        cache: "no-cache",
      });
      if (res.status === 200) return { ok: true, body: await res.json() };
      return { ok: false, status: res.status };
    } catch (e) {
      return { ok: false, status: 0, error: e.message };
    }
  }

  // Merge the remote registry into the local one. New trips (by slug or id)
  // get appended. Existing trips keep their local id but pick up remote
  // name/dates when remote looks newer (we don't track per-row mtimes, so
  // the simple rule: if local has no dates but remote does, take remote).
  async function pullAndMerge() {
    const remote = await fetchRemote();
    if (!remote.ok) return remote;
    const remoteList = Array.isArray(remote.body?.trips) ? remote.body.trips : [];
    const local = readRegistry();
    const bySlug = new Map(local.filter(t => t.slug).map(t => [t.slug, t]));
    const byId = new Map(local.map(t => [t.id, t]));
    let changed = false;
    for (const r of remoteList) {
      if (!r || (!r.slug && !r.id)) continue;
      const existing = (r.slug && bySlug.get(r.slug)) || (r.id && byId.get(r.id));
      if (!existing) {
        local.push({
          id: r.id || ("t" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)),
          slug: r.slug,
          name: r.name || "Untitled trip",
          start: r.start || null,
          end: r.end || null,
        });
        changed = true;
      } else {
        // Pull in fields the local copy lacks.
        if (!existing.slug && r.slug) { existing.slug = r.slug; changed = true; }
        if (!existing.name && r.name) { existing.name = r.name; changed = true; }
        if (!existing.start && r.start) { existing.start = r.start; changed = true; }
        if (!existing.end && r.end) { existing.end = r.end; changed = true; }
      }
    }
    if (changed) writeRegistry(local);
    return { ok: true, changed };
  }

  window.RegistrySync = { schedulePush, pushNow, fetchRemote, pullAndMerge };
})();
