#!/usr/bin/env bash
# ABOUTME: End-to-end health probe for the self-hosted Cap upload path.
# ABOUTME: Proves uploads ACTUALLY work — not just that the box is up.
#
# Why this exists: on 2026-06-22 cap-web was "up" and hey.dancumberlandlabs.com
# returned 200, but every recording upload failed silently. Root cause: the stack
# had been restarted from the stale default docker-compose.yml (pre-MinIO-migration
# variable names), so cap-web booted with empty S3 creds + S3_PUBLIC_ENDPOINT=
# http://localhost:9000. The server handed the desktop app a presigned URL pointing
# at the user's own machine -> "Failed to upload recording." A root-200 check could
# never have caught this. This probe checks the three layers that can each break the
# upload path independently.
#
# Deployed to VPS at ~/cap-healthcheck/ ; run by cron */15. Source of truth lives
# here in the repo (AI_Tools/Cap_SelfHost/). Liveness (the probe's own death) is
# covered by fleet-watchdog registry key "cap-upload-health" (mtime on last_run.log).
set -uo pipefail

# All config is env-overridable so the FAIL path can be fault-tested in isolation
# (e.g. DRILL=1 FAIL_THRESHOLD=1 EXPECT_BUCKET=bogus DIR=/tmp/captest ./cap_healthcheck.sh).
# DRILL=1 makes post_slack print instead of POST — proves the alert is invoked without spam.
: "${DIR:=$HOME/cap-healthcheck}"
: "${ENV_FILE:=$HOME/cap/.env}"
: "${WEBHOOK_FILE:=$HOME/dancumberlandlabs-content/.slack-webhook}"
: "${EXPECT_ENDPOINT:=https://s3.dancumberlandlabs.com}"
: "${EXPECT_BUCKET:=cap-recordings}"
: "${APP_URL:=https://hey.dancumberlandlabs.com/}"
: "${FAIL_THRESHOLD:=2}"         # alert only after N consecutive fails (mirrors site-uptime)

STAMP="$DIR/last_run.log"        # rewritten EVERY run -> fleet-watchdog mtime dead-man's-switch
FAILCOUNT="$DIR/failcount"       # flap damping across runs

mkdir -p "$DIR"
ts() { date -u +"%Y-%m-%dT%H:%M:%SZ"; }
problems=()

# --- Layer 1: cap-web config assertion (catches the exact 2026-06-22 regression) ---
ENV_DUMP=$(docker inspect cap-web -f '{{range .Config.Env}}{{println .}}{{end}}' 2>/dev/null)
if [ -z "$ENV_DUMP" ]; then
  problems+=("cap-web container not running / not inspectable")
else
  ep=$(printf '%s\n' "$ENV_DUMP" | grep '^S3_PUBLIC_ENDPOINT=' | cut -d= -f2-)
  bk=$(printf '%s\n' "$ENV_DUMP" | grep '^CAP_AWS_BUCKET=' | cut -d= -f2-)
  ak=$(printf '%s\n' "$ENV_DUMP" | grep '^CAP_AWS_ACCESS_KEY=' | cut -d= -f2-)
  [ "$ep" = "$EXPECT_ENDPOINT" ] || problems+=("cap-web S3_PUBLIC_ENDPOINT='$ep' (expected $EXPECT_ENDPOINT)")
  [ "$bk" = "$EXPECT_BUCKET" ]   || problems+=("cap-web CAP_AWS_BUCKET='$bk' (expected $EXPECT_BUCKET)")
  [ -n "$ak" ]                   || problems+=("cap-web CAP_AWS_ACCESS_KEY is EMPTY")
fi

# --- Layer 2: storage round-trip over the PUBLIC endpoint (the path the desktop app uses) ---
# Catches MinIO down, Caddy/DNS/cert breakage on s3 subdomain, missing bucket, bad creds.
U=$(grep '^MINIO_ROOT_USER=' "$ENV_FILE" 2>/dev/null | cut -d= -f2-)
P=$(grep '^MINIO_ROOT_PASSWORD=' "$ENV_FILE" 2>/dev/null | cut -d= -f2-)
if [ -z "$U" ] || [ -z "$P" ]; then
  problems+=("could not read MinIO creds from $ENV_FILE")
else
  KEY="__healthcheck/probe-$(date -u +%s).txt"
  RT=$(docker run --rm --entrypoint sh minio/mc:latest -c "
    mc alias set pub $EXPECT_ENDPOINT '$U' '$P' >/dev/null 2>&1 || { echo ALIAS_FAIL; exit 0; }
    echo cap-probe > /tmp/p.txt
    mc cp /tmp/p.txt pub/$EXPECT_BUCKET/$KEY >/dev/null 2>&1 || { echo WRITE_FAIL; exit 0; }
    out=\$(mc cat pub/$EXPECT_BUCKET/$KEY 2>/dev/null)
    mc rm pub/$EXPECT_BUCKET/$KEY >/dev/null 2>&1
    [ \"\$out\" = cap-probe ] && echo OK || echo READ_FAIL
  " 2>/dev/null | tail -1)
  [ "$RT" = "OK" ] || problems+=("storage round-trip failed: ${RT:-no-response}")
fi

# --- Layer 3: public app reachability ---
code=$(curl -s -o /dev/null -m 15 -w '%{http_code}' "$APP_URL" 2>/dev/null)
case "$code" in 2*|3*) ;; *) problems+=("hey.dancumberlandlabs.com returned HTTP ${code:-000}") ;; esac

post_slack() {
  local text="$1" url
  if [ -n "${DRILL:-}" ]; then
    echo "[DRILL] post_slack invoked — would POST to #notifications:"; echo "$text"; return
  fi
  url=$(head -1 "$WEBHOOK_FILE" 2>/dev/null)
  [ -n "$url" ] || { echo "(no slack webhook)" >&2; return; }
  python3 - "$url" "$text" <<'PY' 2>/dev/null
import json, sys, urllib.request
url, text = sys.argv[1], sys.argv[2]
req = urllib.request.Request(url, data=json.dumps({"text": text}).encode(),
                             headers={"Content-Type": "application/json"}, method="POST")
try: urllib.request.urlopen(req, timeout=10)
except Exception as e: print(f"(slack failed: {e})", file=sys.stderr)
PY
}

prev=$(cat "$FAILCOUNT" 2>/dev/null || echo 0)

if [ ${#problems[@]} -eq 0 ]; then
  echo "$(ts) OK" > "$STAMP"
  echo 0 > "$FAILCOUNT"
  # recovery notice if we had been alerting
  if [ "$prev" -ge "$FAIL_THRESHOLD" ] 2>/dev/null; then
    post_slack "✅ *Cap upload health recovered* — $(ts). Uploads working again (config + storage round-trip pass)."
  fi
else
  n=$(( prev + 1 ))
  echo "$n" > "$FAILCOUNT"
  printf '%s FAIL(%d): %s\n' "$(ts)" "$n" "${problems[*]}" > "$STAMP"
  if [ "$n" -ge "$FAIL_THRESHOLD" ]; then
    body="🔴 *Cap upload health FAIL* (consecutive #$n) — $(ts)"$'\n'
    for p in "${problems[@]}"; do body+="• $p"$'\n'; done
    body+="Recordings will fail to upload. Runbook: AI_Tools/Cap_SelfHost/README.md"
    post_slack "$body"
  fi
fi
