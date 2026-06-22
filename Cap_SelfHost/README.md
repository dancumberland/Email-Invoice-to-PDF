# Cap Self-Hosted Setup

**URL:** https://hey.dancumberlandlabs.com
**Server:** 159.203.139.119 (DigitalOcean VPS, Tailscale 100.99.136.54, SSH: `claude@100.99.136.54`)
**Storage:** self-hosted **MinIO** on the VPS (migrated off Cloudflare R2 on 2026-03-25)
**Storage public endpoint:** https://s3.dancumberlandlabs.com (Caddy → `localhost:9000`)
**Bucket:** `cap-recordings`

## Architecture

```
Desktop app / browser
   │  (records, then asks cap-web where to upload)
   ▼
Caddy (HTTPS + branding text-replace) ──► cap-web (Next.js, :3000)
   │                                          │ hands back a PRESIGNED upload URL
   │                                          │ built from S3_PUBLIC_ENDPOINT
   ▼                                          ▼
hey.dancumberlandlabs.com            s3.dancumberlandlabs.com ──► MinIO (:9000) ──► cap-recordings
```

The upload is **direct from the client to the storage endpoint** via a presigned URL.
That's why `S3_PUBLIC_ENDPOINT` MUST be a public, client-reachable host
(`https://s3.dancumberlandlabs.com`) — if it's `localhost:9000`, the client tries to
upload to *itself* and you get **"Failed to upload recording."**

## Files on VPS

- `/home/claude/cap/docker-compose.yml` — **the live config** (MinIO). This is the default
  file `docker compose` picks up, so it MUST be the correct one (see the trap below).
- `/home/claude/cap/.env` — secrets (MinIO creds, NextAuth/DB encryption keys, MySQL).
- `/etc/caddy/Caddyfile` — reverse proxy for `hey.` (port 3000) + `s3.` (port 9000) + branding removal.

Local mirrors of `docker-compose.yml` and `Caddyfile` live in this directory. **VPS is the
source of truth** — edit there, then mirror down.

## ⚠️ The two-compose-file trap (caused the 2026-06-22 outage)

There used to be TWO compose files on the VPS with **different env-variable names**:

| | old `docker-compose.yml` (pre-migration) | correct config (MinIO) |
|---|---|---|
| access key | `${CAP_AWS_ACCESS_KEY}` → empty | `${MINIO_ROOT_USER}` → `capadmin` |
| bucket | `${CAP_AWS_BUCKET:-cap}` → `cap` | `${S3_BUCKET}` → `cap-recordings` |
| public endpoint | `${S3_PUBLIC_URL:-localhost:9000}` | `${S3_PUBLIC_ENDPOINT}` → `s3.dancumberlandlabs.com` |

The `.env` was rewritten for MinIO with the **new** names, but a restart from the **old
default** `docker-compose.yml` booted `cap-web` with empty creds + a `localhost` upload
endpoint — uploads failed silently while `hey.` still returned 200.

**Fix applied 2026-06-22:** the correct MinIO config is now the default `docker-compose.yml`
(old one preserved as `docker-compose.yml.bak-jan30-preminiofix`). A bare `docker compose up -d`
is now safe. **Never** start the stack with `-f <some-other-file>` unless you've confirmed its
variable names match `.env`.

## Health / Watchdog

**`cap_healthcheck.sh`** (in this dir; deployed to VPS `~/cap-healthcheck/`, cron `*/15`).
End-to-end probe — proves uploads ACTUALLY work, not just that the box is up:

1. **cap-web config assertion** — `S3_PUBLIC_ENDPOINT == s3.dancumberlandlabs.com`,
   `CAP_AWS_BUCKET == cap-recordings`, access key non-empty. (Catches the trap above.)
2. **MinIO storage round-trip** — write/read/delete a tiny object in `cap-recordings`
   over the public HTTPS endpoint = the exact path the desktop app uses.
3. **App reachability** — `hey.dancumberlandlabs.com` returns 2xx/3xx.

Self-alerts to Slack `#notifications` on 2 consecutive fails (+ recovery notice).
Registered in `Project_Management/fleet-watchdog/registry.py` as **`cap-upload-health`**
(mtime on `~/cap-healthcheck/last_run.log`) — that's the dead-man's-switch for the probe itself.

Manual run: `ssh claude@100.99.136.54 'bash ~/cap-healthcheck/cap_healthcheck.sh; cat ~/cap-healthcheck/last_run.log'`

## Docker Commands

```bash
# Status
ssh claude@100.99.136.54 'cd /home/claude/cap && docker compose ps'
# Logs
ssh claude@100.99.136.54 'cd /home/claude/cap && docker compose logs -f cap-web'
# Restart / re-apply config (safe: default file is the correct MinIO config)
ssh claude@100.99.136.54 'cd /home/claude/cap && docker compose up -d'
```

## Caddy: branding removal + s3 route

`hey.dancumberlandlabs.com` strips Cap branding via response text-replacement (logo footer,
"Recorded with", page title, sidebar). If Cap updates its frontend markup these replacements
break silently — eyeball the page after any `cap-web` image update. `s3.dancumberlandlabs.com`
reverse-proxies to MinIO on `localhost:9000`. Reload after edits: `sudo systemctl reload caddy`.

To re-enable a stripped element, remove its `replace` line in `/etc/caddy/Caddyfile`.

## AI transcription / summaries (optional, currently off)

Add `DEEPGRAM_API_KEY` (transcription) or `GROQ_API_KEY`/`OPENAI_API_KEY` (summaries) to
`/home/claude/cap/.env`, then `docker compose up -d`.

## Database: user marked as Pro (watermark removal)

```bash
ssh claude@100.99.136.54 "docker exec cap-mysql mysql -ucap -p\$(grep '^MYSQL_PASSWORD=' /home/claude/cap/.env | cut -d= -f2-) cap -e 'SELECT email, stripeSubscriptionStatus FROM users;'"
```
Watermark is removed because the user has `stripeSubscriptionStatus = 'active'`.
