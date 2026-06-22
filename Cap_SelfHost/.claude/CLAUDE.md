## Project: Cap Self-Host

Self-hosted Cap (Loom alternative) running at **https://hey.dancumberlandlabs.com**. Dan uses this daily for screen recordings.

---

## Key Facts

- **URL:** https://hey.dancumberlandlabs.com
- **VPS:** `100.99.136.54` (Tailscale) / `159.203.139.119` (public)
- **Storage:** self-hosted **MinIO** on the VPS (migrated off R2 2026-03-25); bucket `cap-recordings`; public endpoint `https://s3.dancumberlandlabs.com` (Caddy → localhost:9000)
- **Architecture:** `Client → Caddy (HTTPS + text replacements) → cap-web :3000`; uploads go **direct via presigned URL** to `s3.dancumberlandlabs.com` → MinIO :9000. `S3_PUBLIC_ENDPOINT` MUST be the public host, never `localhost` (else "Failed to upload recording").
- **Watchdog:** `cap-upload-health` in fleet-watchdog — `*/15` end-to-end probe (config + storage round-trip + reachability), self-alerts Slack. Source `cap_healthcheck.sh`. See README § Health/Watchdog.
- **⚠️ Two-compose-file trap:** the correct MinIO config IS now the default `docker-compose.yml`. A bare `docker compose up -d` is safe; do NOT use `-f <other-file>` without checking its var names match `.env`. See README § the trap (root cause of the 2026-06-22 outage).

## VPS Paths

- `/home/claude/cap/docker-compose.yml` — **live** Docker Compose config (MinIO; correct default)
- `/home/claude/cap/.env` — secrets (MinIO creds, NextAuth/DB keys, MySQL)
- `/etc/caddy/Caddyfile` — reverse proxy (`hey.` :3000 + `s3.` :9000) + Cap branding removal
- `/home/claude/cap-healthcheck/cap_healthcheck.sh` — upload-path watchdog (cron `*/15`)

## Local Files

- `README.md` — full setup + customization notes (source of truth for Caddy `replace` rules, DEEPGRAM/GROQ integration, Pro user flag)
- `docker-compose.r2.yml` — local copy of VPS config
- `Caddyfile` — local copy of VPS Caddy config

## Common Tasks

| Task | Command |
|------|---------|
| Reload Caddy after config edit | `ssh claude@100.99.136.54 'sudo systemctl reload caddy'` |
| Restart Cap containers | `ssh claude@100.99.136.54 'cd /home/claude/cap && docker compose -f docker-compose.r2.yml restart'` |
| View Cap logs | `ssh claude@100.99.136.54 'cd /home/claude/cap && docker compose -f docker-compose.r2.yml logs -f'` |
| Add AI transcription | See README.md — add `DEEPGRAM_API_KEY` to `.env`, restart |
| Add AI summaries | See README.md — add `GROQ_API_KEY` or `OPENAI_API_KEY` to `.env`, restart |
| Re-enable Cap branding / sidebar | See README.md — remove relevant `replace` line in Caddyfile, reload |

## Project-Specific Rules

- Caddy does **response text replacement** to strip Cap branding; if Cap updates its frontend markup, replacements may break silently — check hey.dancumberlandlabs.com visually after Cap container updates
- The "Pro user" watermark removal depends on a manual DB flag (`stripeSubscriptionStatus = 'active'`); see README.md for details
- Always edit the **VPS files** as source of truth, then mirror to local copies (not the other way around)
- This subdirectory is registered in `CoreContext/SYSTEMS.md` under Websites & Hosting
