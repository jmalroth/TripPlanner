// Firebase client — auth + Firestore wrappers for myitin.
//
// Loaded as a module from the HTML pages. Exposes a small surface on
// `window.fb` so the existing non-module app.js can call into it without
// being rewritten as ES modules.
//
// Replace the firebaseConfig values below with the ones from
// console.firebase.google.com → Project settings → Your apps → Web app.

import { initializeApp }
  from "https://www.gstatic.com/firebasejs/10.13.2/firebase-app.js";
import {
  getAuth, onAuthStateChanged, signOut,
  GoogleAuthProvider, signInWithPopup,
  signInWithEmailAndPassword, createUserWithEmailAndPassword,
  sendSignInLinkToEmail, isSignInWithEmailLink, signInWithEmailLink,
  updatePassword,
} from "https://www.gstatic.com/firebasejs/10.13.2/firebase-auth.js";
import {
  getFirestore, doc, getDoc, setDoc, updateDoc, deleteDoc,
  collection, query, where, orderBy, onSnapshot, getDocs,
  serverTimestamp, arrayUnion, arrayRemove, deleteField,
} from "https://www.gstatic.com/firebasejs/10.13.2/firebase-firestore.js";

// --- CONFIG ---------------------------------------------------------------
// Paste the firebaseConfig object from the Firebase console below.
const firebaseConfig = {
  apiKey: "AIzaSyAUUr7qOMtC7gUyaWpbajIxfj0AEoO557k",
  authDomain: "myitin.firebaseapp.com",
  projectId: "myitin",
  storageBucket: "myitin.firebasestorage.app",
  messagingSenderId: "750554614907",
  appId: "1:750554614907:web:5467991759c836ca9866de",
  measurementId: "G-C9Q1138SMY",
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// --- AUTH ----------------------------------------------------------------

let currentUser = null;
let authReady = false;
let _authReadyResolve;
const authReadyPromise = new Promise((r) => { _authReadyResolve = r; });
const authListeners = new Set();

function emailKey(email) {
  return (email || "").trim().toLowerCase();
}

// Write/refresh the profile doc + email → uid mapping so other users can
// look this person up to share trips with them.
async function ensureProfile(user) {
  if (!user) return;
  try {
    await setDoc(doc(db, "users", user.uid), {
      displayName: user.displayName || null,
      email: user.email || null,
      createdAt: serverTimestamp(),
    }, { merge: true });
    const key = emailKey(user.email);
    if (key) {
      await setDoc(doc(db, "userEmails", key), { uid: user.uid }, { merge: true });
    }
  } catch (e) {
    console.warn("ensureProfile failed:", e.message);
  }
}

// If the signed-in user changes (new account on this device, or sign-out),
// purge any trip data left in localStorage so we never show one user's data
// while authenticated as someone else.
function purgeStaleLocalData() {
  try {
    localStorage.removeItem("trip-builder-trips");
    for (const key of Object.keys(localStorage)) {
      if (key.startsWith("trip-builder-trip-")) localStorage.removeItem(key);
    }
  } catch (_) {}
}

const LAST_UID_KEY = "myitin:lastUid";
onAuthStateChanged(auth, async (user) => {
  const prevUid = localStorage.getItem(LAST_UID_KEY);
  const nextUid = user ? user.uid : null;
  if (prevUid !== nextUid) {
    purgeStaleLocalData();
    if (nextUid) localStorage.setItem(LAST_UID_KEY, nextUid);
    else localStorage.removeItem(LAST_UID_KEY);
  }
  currentUser = user;
  if (!authReady) { authReady = true; _authReadyResolve(); }
  let claims = {};
  if (user) {
    try {
      const tokenResult = await user.getIdTokenResult();
      claims = tokenResult.claims || {};
    } catch (_) { /* ignore */ }
    ensureProfile(user); // fire and forget
  }
  for (const fn of authListeners) fn(user, claims);
});

function onAuth(fn) {
  authListeners.add(fn);
  // Fire once with current state so late subscribers don't miss it.
  if (currentUser !== null) {
    currentUser.getIdTokenResult().then(r => fn(currentUser, r.claims || {}));
  } else {
    fn(null, {});
  }
  return () => authListeners.delete(fn);
}

async function signInGoogle() {
  const provider = new GoogleAuthProvider();
  const { user } = await signInWithPopup(auth, provider);
  return user;
}

async function signInPassword(email, password) {
  const { user } = await signInWithEmailAndPassword(auth, email, password);
  return user;
}

async function signUpPassword(email, password) {
  const { user } = await createUserWithEmailAndPassword(auth, email, password);
  return user;
}

const EMAIL_LINK_KEY = "myitin:emailForSignIn";

async function sendMagicLink(email) {
  const actionCodeSettings = {
    url: window.location.origin + window.location.pathname,
    handleCodeInApp: true,
  };
  await sendSignInLinkToEmail(auth, email, actionCodeSettings);
  window.localStorage.setItem(EMAIL_LINK_KEY, email);
}

// Call on page load to complete a magic-link sign in if the URL has one.
async function tryCompleteEmailLink() {
  if (!isSignInWithEmailLink(auth, window.location.href)) return null;
  let email = window.localStorage.getItem(EMAIL_LINK_KEY);
  if (!email) {
    email = window.prompt("Confirm the email address you used to request the sign-in link:");
  }
  if (!email) return null;
  const { user } = await signInWithEmailLink(auth, email, window.location.href);
  window.localStorage.removeItem(EMAIL_LINK_KEY);
  // Strip the sign-in params from the URL.
  const clean = new URL(window.location.href);
  ["apiKey", "oobCode", "mode", "lang"].forEach(p => clean.searchParams.delete(p));
  window.history.replaceState({}, document.title, clean.toString());
  return user;
}

function signOutNow() { return signOut(auth); }

async function setPassword(newPassword) {
  if (!currentUser) throw new Error("not signed in");
  if (!newPassword || newPassword.length < 6) throw new Error("Password must be at least 6 characters.");
  await updatePassword(currentUser, newPassword);
  await setDoc(doc(db, "users", currentUser.uid), { hasPassword: true }, { merge: true });
}

async function hasPassword() {
  if (!currentUser) return false;
  try {
    const snap = await getDoc(doc(db, "users", currentUser.uid));
    return !!(snap.exists() && snap.data().hasPassword);
  } catch (_) { return false; }
}

// --- TRIP DATA -----------------------------------------------------------
//
// Schema:
//   trips/{slug}
//     ownerId: uid
//     sharedWith: [uid]
//     viewerToken: "..."     (kept for link-sharing to logged-out viewers)
//     ...the rest of the trip blob (minus pricing)
//   trips/{slug}/private/pricing
//     lineItems, priceSplit, priceToken

const PRICING_KEYS = ["lineItems", "priceSplit", "priceToken", "viewerToken"];

function splitPricing(blob) {
  const main = { ...blob };
  const pricing = {};
  for (const k of PRICING_KEYS) {
    if (k in main) { pricing[k] = main[k]; delete main[k]; }
  }
  return { main, pricing };
}

function mergePricing(main, pricing) {
  if (!pricing) return main;
  return { ...main, ...pricing };
}

async function loadTrip(slug) {
  const tripRef = doc(db, "trips", slug);
  const snap = await getDoc(tripRef);
  if (!snap.exists()) return { ok: false, status: 404 };
  const main = snap.data();
  const isOwner = !!(currentUser && main.ownerId === currentUser.uid);
  const myLevel = isOwner
    ? "owner"
    : (currentUser && main.permissions && main.permissions[currentUser.uid]) || null;
  let pricing = null;
  // Try to load pricing; rules deny for users without pricing/edit access.
  try {
    const p = await getDoc(doc(db, "trips", slug, "private", "pricing"));
    if (p.exists()) pricing = p.data();
  } catch (_) { /* not authorized for pricing */ }
  return {
    ok: true,
    status: 200,
    body: mergePricing(main, pricing),
    isOwner,
    level: myLevel, // "owner" | "viewer" | "viewer-pricing" | "editor" | null
  };
}

// Sharing levels (account-based — anonymous public links live elsewhere):
//   "viewer"          — sees itinerary, no pricing, no edit
//   "viewer-pricing"  — itinerary + pricing tab, no edit
//   "editor"          — full read/write except can't change ownership
const LEVELS = ["viewer", "viewer-pricing", "editor"];

async function saveTrip(slug, blob) {
  if (!currentUser) return { ok: false, status: 401 };
  const { main, pricing } = splitPricing(blob);
  main.ownerId = currentUser.uid;
  // permissions is a map { uid: level } — source of truth.
  // permissionUids is a flat array of the keys, denormalized so we can
  // query trips with array-contains.
  if (!main.permissions || typeof main.permissions !== "object") main.permissions = {};
  main.permissionUids = Object.keys(main.permissions);
  // Legacy sharedWith is no longer used. Clear it from new writes.
  main.sharedWith = deleteField();
  main.updatedAt = serverTimestamp();
  for (const k of PRICING_KEYS) main[k] = deleteField();
  try {
    await setDoc(doc(db, "trips", slug), main, { merge: true });
    if (Object.keys(pricing).length) {
      await setDoc(doc(db, "trips", slug, "private", "pricing"), pricing, { merge: true });
    }
    // If a public link exists, keep its mirror up to date. Best-effort —
    // a mirror failure shouldn't block the main save.
    try {
      const fresh = await getDoc(doc(db, "trips", slug));
      const token = fresh.exists() && fresh.data().publicToken;
      if (token) {
        const mirror = stripForPublic(fresh.data());
        mirror.ownerId = currentUser.uid;
        mirror.slug = slug;
        await setDoc(doc(db, "publicTrips", token), mirror, { merge: true });
      }
    } catch (_) { /* ignore mirror errors */ }
    return { ok: true, status: 200 };
  } catch (e) {
    return { ok: false, status: e.code === "permission-denied" ? 403 : 500, error: e.message };
  }
}

async function deleteTrip(slug) {
  if (!currentUser) return { ok: false, status: 401 };
  try {
    await deleteDoc(doc(db, "trips", slug, "private", "pricing")).catch(() => {});
    await deleteDoc(doc(db, "trips", slug));
    return { ok: true, status: 200 };
  } catch (e) {
    return { ok: false, status: 500, error: e.message };
  }
}

// Live subscription to the trips the signed-in user owns OR has any
// permission on. Calls `cb(trips[])` whenever anything changes. Returns
// an unsubscribe fn. Each trip carries _role ("owner" | level) so the UI
// can pick gating.
function subscribeMyTrips(cb) {
  if (!currentUser) { cb([]); return () => {}; }
  const uid = currentUser.uid;
  const owned = query(collection(db, "trips"), where("ownerId", "==", uid));
  const shared = query(collection(db, "trips"), where("permissionUids", "array-contains", uid));
  const ownedById = new Map();
  const sharedById = new Map();
  const emit = () => {
    const merged = new Map([...ownedById, ...sharedById]);
    const list = Array.from(merged.values())
      .sort((a, b) => (a.start || "").localeCompare(b.start || ""));
    cb(list);
  };
  const u1 = onSnapshot(owned, s => {
    ownedById.clear();
    s.forEach(d => ownedById.set(d.id, { id: d.id, ...d.data(), _role: "owner" }));
    emit();
  }, (err) => console.error("subscribeMyTrips owned error:", err));
  const u2 = onSnapshot(shared, s => {
    sharedById.clear();
    s.forEach(d => {
      const data = d.data();
      const level = (data.permissions && data.permissions[uid]) || "viewer";
      sharedById.set(d.id, { id: d.id, ...data, _role: level });
    });
    emit();
  }, (err) => console.error("subscribeMyTrips shared error:", err));
  return () => { u1(); u2(); };
}

// Live subscription to a single trip (owner sees pricing too).
function subscribeTrip(slug, cb) {
  const u1 = onSnapshot(doc(db, "trips", slug), async (snap) => {
    if (!snap.exists()) { cb(null); return; }
    const main = snap.data();
    let pricing = null;
    try {
      const p = await getDoc(doc(db, "trips", slug, "private", "pricing"));
      if (p.exists()) pricing = p.data();
    } catch (_) { /* viewer */ }
    const isOwner = currentUser && main.ownerId === currentUser.uid;
    cb({ ...mergePricing(main, pricing), _isOwner: isOwner });
  });
  return u1;
}

// Look up a uid by email. Returns null if no signed-in user has that email
// (they need to have signed in at least once for the mapping to exist).
async function lookupUidByEmail(email) {
  const key = emailKey(email);
  if (!key) return null;
  const snap = await getDoc(doc(db, "userEmails", key));
  return snap.exists() ? (snap.data().uid || null) : null;
}

function assertLevel(level) {
  if (!LEVELS.includes(level)) {
    throw new Error(`Invalid level "${level}". Must be one of: ${LEVELS.join(", ")}`);
  }
}

// Grant access to a user by email. Updates both the permissions map (source
// of truth) and permissionUids array (for queryability) atomically.
async function shareTripWithEmail(slug, email, level = "viewer") {
  if (!currentUser) throw new Error("not signed in");
  assertLevel(level);
  const uid = await lookupUidByEmail(email);
  if (!uid) throw new Error("That email hasn't signed in to myitin yet. Ask them to sign in once, then try again.");
  if (uid === currentUser.uid) throw new Error("That's you.");
  await updateDoc(doc(db, "trips", slug), {
    [`permissions.${uid}`]: level,
    permissionUids: arrayUnion(uid),
  });
  return { uid, level };
}

// Change a recipient's level. uid must already have an entry in permissions.
async function setTripPermissionLevel(slug, uid, level) {
  if (!currentUser) throw new Error("not signed in");
  assertLevel(level);
  await updateDoc(doc(db, "trips", slug), {
    [`permissions.${uid}`]: level,
    permissionUids: arrayUnion(uid),
  });
}

// Revoke a recipient's access entirely.
async function unshareTripWithUid(slug, uid) {
  if (!currentUser) throw new Error("not signed in");
  await updateDoc(doc(db, "trips", slug), {
    [`permissions.${uid}`]: deleteField(),
    permissionUids: arrayRemove(uid),
  });
}

// Read the display name/email for a uid (best-effort — falls back to the uid
// if the profile doc isn't readable for any reason).
async function lookupProfile(uid) {
  try {
    const snap = await getDoc(doc(db, "users", uid));
    if (snap.exists()) return { uid, ...snap.data() };
  } catch (_) {}
  return { uid, displayName: null, email: null };
}

// Create an anonymous one-off snapshot of a trip blob. Returns the snapshot id.
// Anyone with the id can read it (snapshots/{id} rule allows public read).
async function createSnapshot(blob) {
  if (!currentUser) throw new Error("not signed in");
  const id = (crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2)).replace(/-/g, "");
  await setDoc(doc(db, "snapshots", id), {
    blob,
    createdBy: currentUser.uid,
    createdAt: serverTimestamp(),
  });
  return id;
}

async function loadSnapshot(id) {
  const snap = await getDoc(doc(db, "snapshots", id));
  if (!snap.exists()) return null;
  return snap.data().blob;
}

// --- PUBLIC LIVE LINKS ----------------------------------------------------
//
// A trip may have a publicToken on it. When set, a mirror doc lives at
// publicTrips/{token} containing the trip (minus pricing and any other
// secrets). Anyone with the URL can read the mirror without signing in.
// Writes to the mirror are gated on ownerId == auth.uid.
//
// Auto-mirror: saveTrip below pushes the stripped blob to the mirror on
// every save, so the link stays "always live."

function randomToken() {
  const bytes = new Uint8Array(12);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
}

const PUBLIC_STRIP_KEYS = [...PRICING_KEYS, "publicToken"];

function stripForPublic(blob) {
  const copy = JSON.parse(JSON.stringify(blob || {}));
  for (const k of PUBLIC_STRIP_KEYS) delete copy[k];
  return copy;
}

// Owner-only: enable the public link. Generates a token, stores it on the
// trip doc, and seeds the mirror. Returns the token.
async function createPublicLink(slug) {
  if (!currentUser) throw new Error("not signed in");
  const tripSnap = await getDoc(doc(db, "trips", slug));
  if (!tripSnap.exists()) throw new Error("trip not found");
  const trip = tripSnap.data();
  if (trip.ownerId !== currentUser.uid) throw new Error("only the owner can create a public link");
  const token = trip.publicToken || randomToken();
  await updateDoc(doc(db, "trips", slug), { publicToken: token });
  const mirror = stripForPublic(trip);
  mirror.ownerId = currentUser.uid;
  mirror.slug = slug;
  await setDoc(doc(db, "publicTrips", token), mirror, { merge: true });
  return token;
}

// Owner-only: revoke. Removes the mirror and clears the trip's token.
async function removePublicLink(slug) {
  if (!currentUser) throw new Error("not signed in");
  const tripSnap = await getDoc(doc(db, "trips", slug));
  if (!tripSnap.exists()) return;
  const trip = tripSnap.data();
  if (trip.ownerId !== currentUser.uid) throw new Error("only the owner can revoke");
  const token = trip.publicToken;
  if (token) {
    try { await deleteDoc(doc(db, "publicTrips", token)); } catch (_) {}
  }
  await updateDoc(doc(db, "trips", slug), { publicToken: deleteField() });
}

// Anonymous read. Used by trip.html when bootstrapping with ?p=<token>.
async function loadPublicTrip(token) {
  const snap = await getDoc(doc(db, "publicTrips", token));
  if (!snap.exists()) return null;
  return snap.data();
}

// --- USER TRIP LABELS ----------------------------------------------------
//
// Per-user "I'm traveling on this" vs "I'm just watching" choice for trips
// shared with them. Doc shape: { <slug>: "traveling" | "watching" }.
// Absent entries mean uncategorized — the trips page surfaces those for
// the user to choose.

function subscribeMyTripLabels(cb) {
  if (!currentUser) { cb({}); return () => {}; }
  return onSnapshot(doc(db, "userTripLabels", currentUser.uid), (snap) => {
    cb(snap.exists() ? (snap.data() || {}) : {});
  }, (err) => {
    console.error("subscribeMyTripLabels error:", err);
    cb({});
  });
}

// Per-user watching-trip background color. Subscribed to live so a change
// from any tab reflects everywhere instantly.
function subscribeMyProfile(cb) {
  if (!currentUser) { cb({}); return () => {}; }
  return onSnapshot(doc(db, "users", currentUser.uid), (snap) => {
    cb(snap.exists() ? (snap.data() || {}) : {});
  }, (err) => { console.error("subscribeMyProfile error:", err); cb({}); });
}

async function setWatchingColor(color) {
  if (!currentUser) throw new Error("not signed in");
  // Accept null/empty to reset; otherwise sanity-check it's a CSS color-ish string.
  if (color && color.length > 30) throw new Error("color too long");
  await setDoc(doc(db, "users", currentUser.uid),
    { watchingColor: color || deleteField() }, { merge: true });
}

async function setTripLabel(slug, label) {
  if (!currentUser) throw new Error("not signed in");
  if (label !== "traveling" && label !== "watching" && label !== null) {
    throw new Error('label must be "traveling", "watching", or null');
  }
  const update = label === null ? { [slug]: deleteField() } : { [slug]: label };
  await setDoc(doc(db, "userTripLabels", currentUser.uid), update, { merge: true });
}

// --- EXPORT --------------------------------------------------------------

window.fb = {
  // auth
  onAuth, signInGoogle, signInPassword, signUpPassword,
  sendMagicLink, tryCompleteEmailLink, signOut: signOutNow,
  setPassword, hasPassword,
  get user() { return currentUser; },
  whenAuthReady() { return authReadyPromise; },
  // data
  loadTrip, saveTrip, deleteTrip,
  subscribeMyTrips, subscribeTrip,
  // sharing + profile
  lookupUidByEmail, shareTripWithEmail, setTripPermissionLevel,
  unshareTripWithUid, lookupProfile,
  LEVELS,
  // snapshots
  createSnapshot, loadSnapshot,
  // public live links
  createPublicLink, removePublicLink, loadPublicTrip,
  // travel-buddy / watching labels
  subscribeMyTripLabels, setTripLabel,
  // user profile (watching color etc.)
  subscribeMyProfile, setWatchingColor,
};

// Kick off magic-link completion automatically on every page load.
tryCompleteEmailLink().catch(err => console.warn("magic-link completion failed:", err));
