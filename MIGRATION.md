# Firebase migration notes

The app now uses Firebase for auth + trip storage. The old Cloudflare Worker is still around but only handles `/smart-parse` (AI parsing) and the anonymous `?v=` viewer links.

## What changed

- **Auth gate** on `index.html` and `trip.html`. Google + email/password + magic link.
- **Trip storage** moved from Cloudflare KV → Firestore (`trips/{slug}` + `trips/{slug}/private/pricing`).
- **Pricing fields** (`lineItems`, `priceSplit`, `priceToken`) are now in an owner-only subcollection. Shared viewers literally cannot read them — enforced by Firestore rules, not by client-side stripping.
- **Sharing** is now by email — owner types a friend's email into the Share dropdown, they get read-only access after signing in once.
- **The `?v=` anonymous viewer link** still works through the worker for now (for sharing with people who don't want to sign in).

## Files

| File | Role |
|---|---|
| `firebase-client.js` | Auth + Firestore wrappers, exposed on `window.fb`. |
| `firebase.json`, `.firebaserc`, `firestore.rules`, `firestore.indexes.json` | Firebase project config + security rules. |
| `migrate.html` | One-time tool to copy trips from the Cloudflare KV → Firestore. |
| `worker/worker.js` | Retained for `/smart-parse` and `?v=` anonymous reads. |

## Deploy steps (one-time)

1. **Enable in Firebase console** (project `myitin`):
   - Authentication → Sign-in providers: Google, Email/Password (+ Email link toggle).
   - Authentication → Settings → Authorized domains: add wherever you'll host (e.g. `myitin.web.app`, your custom domain).
   - Firestore Database → Create (production mode).
2. **Deploy rules + indexes** — paste `firestore.rules` into the Rules tab in the console and Publish. Add the two composite indexes from `firestore.indexes.json` under the Indexes tab. (Or use the Firebase CLI: `firebase deploy --only firestore:rules,firestore:indexes`.)
3. **Run the migration** at `/migrate.html` once, signed in. Enter the worker URL + owner password, click Dry run, then Run migration.
4. **Hosting deploy** (when ready): `firebase deploy --only hosting`.

## Feature flags (per-user)

Set custom claims via the Admin SDK or `gcloud`:

```bash
gcloud auth login
gcloud config set project myitin

# Example: grant a user the beta-compare feature flag.
firebase functions:shell  # or use a small Node script with firebase-admin
```

In a Node script:

```js
import { initializeApp, applicationDefault } from "firebase-admin/app";
import { getAuth } from "firebase-admin/auth";
initializeApp({ credential: applicationDefault() });
await getAuth().setCustomUserClaims("USER_UID", { betaCompare: true, smartParse: true });
```

The client reads these via `user.getIdTokenResult().claims` — wire UI gating to `claims.betaCompare` etc.

## What still uses the worker

- `/smart-parse` — AI parsing of emails/listings. Still requires `OWNER_PASSWORD` (smart-parse buttons will prompt for it).
- `GET /<slug>?v=<token>` — anonymous viewer reads via shared `?v=` link. Pricing is stripped unless `v=` matches.

When you want to retire the worker entirely, move smart-parse to a Cloud Function that verifies the user's Firebase ID token + checks the `smartParse` custom claim.
