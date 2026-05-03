# Trip snapshot worker

Tiny Cloudflare Worker + KV that backs the **Share** button in the trip
builder. POST a JSON blob, get back an id; GET the id, get the blob.

## One-time setup

```sh
# 1. From this directory:
npm install -g wrangler   # or use `npx wrangler ...` everywhere below
wrangler login            # opens a browser; sign in to Cloudflare

# 2. Create the KV namespace:
wrangler kv namespace create SNAPSHOTS
# Copy the `id = "..."` it prints into wrangler.toml.

# 3. Deploy:
wrangler deploy
# It prints a URL like https://trip-snapshots.<account>.workers.dev
```

## Wire the app to it

Open `app.js` and set:

```js
const SNAPSHOT_API = "https://trip-snapshots.<account>.workers.dev";
```

That's it — Share now POSTs to your worker.

## Free-tier limits

100k reads/day, 1k writes/day, 1k lists/day. Way more than a personal trip
diary needs.
