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
# never have caught this. This probe checks the four layers that can each break the
# upload path independently.
#
# Layer 4 (added 2026-07-03) covers a DIFFERENT failure class: not a live regression,
# but a STALE reference. Dan's avatar + org icon were broken-image icons on a public
# share page because both were uploaded 2026-01-30, before the 2026-03-25 R2->MinIO
# migration, and the objects never carried over -- the DB kept pointing at dead keys.
# Layers 1-3 only prove NEW uploads work; nothing checked that EXISTING referenced
# objects were still fetchable. Layer 4 closes that gap.
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

# --- Layer 2: MULTIPART storage round-trip over the PUBLIC endpoint ---
# Runs the EXACT operation the desktop upload failed on (2026-06-22): CreateMultipartUpload,
# which 500'd ("S3 operation failed") while the server had empty creds. A single PUT working
# does NOT prove multipart works (Caddy proxying of uploadId/partNumber, MinIO multipart, the
# server-side initiate op can all break independently) — so we do a full create->upload_part->
# complete->readback->delete cycle with a ~9-byte part. Catches MinIO down, s3-subdomain
# cert/DNS/Caddy breakage, missing bucket, bad creds, AND multipart-specific regressions.
# Needs host python3 + boto3 (present on VPS; if it ever vanishes the probe reports a failure,
# and fleet-watchdog's mtime dead-man's-switch still flags a probe that stops writing).
RT=$(python3 - "$ENV_FILE" "$EXPECT_ENDPOINT" "$EXPECT_BUCKET" <<'PY' 2>/dev/null
import os, sys
env_file, ep, bucket = sys.argv[1], sys.argv[2], sys.argv[3]
try:
    import boto3
    from botocore.config import Config
    env = {}
    for line in open(env_file):
        line = line.strip()
        if line and not line.startswith('#') and '=' in line:
            k, v = line.split('=', 1)
            env[k.replace('export ', '').strip()] = v.strip().strip('"').strip("'")
    s3 = boto3.client('s3', endpoint_url=ep,
                      aws_access_key_id=env['MINIO_ROOT_USER'],
                      aws_secret_access_key=env['MINIO_ROOT_PASSWORD'],
                      region_name='us-east-1',
                      config=Config(s3={'addressing_style': 'path'},
                                    connect_timeout=10, read_timeout=15,
                                    retries={'max_attempts': 1}))
    key = '__healthcheck/mpu-probe.txt'
    uid = s3.create_multipart_upload(Bucket=bucket, Key=key)['UploadId']
    try:
        part = s3.upload_part(Bucket=bucket, Key=key, PartNumber=1, UploadId=uid, Body=b'cap-probe')
        s3.complete_multipart_upload(Bucket=bucket, Key=key, UploadId=uid,
            MultipartUpload={'Parts': [{'PartNumber': 1, 'ETag': part['ETag']}]})
    except Exception:
        s3.abort_multipart_upload(Bucket=bucket, Key=key, UploadId=uid)
        raise
    ok = s3.get_object(Bucket=bucket, Key=key)['Body'].read() == b'cap-probe'
    s3.delete_object(Bucket=bucket, Key=key)
    print('OK' if ok else 'READBACK_MISMATCH')
except Exception as e:
    print(type(e).__name__ + ':' + str(e)[:120])
PY
)
[ "$RT" = "OK" ] || problems+=("multipart round-trip failed: ${RT:-no-response}")

# --- Layer 3: public app reachability ---
code=$(curl -s -o /dev/null -m 15 -w '%{http_code}' "$APP_URL" 2>/dev/null)
case "$code" in 2*|3*) ;; *) problems+=("hey.dancumberlandlabs.com returned HTTP ${code:-000}") ;; esac

# --- Layer 4: referenced-asset integrity (avatars / org+space icons still in the bucket) ---
# Reads every non-null users.image / organizations.iconUrl / spaces.iconUrl and HEADs each
# key against the live bucket. Single-user instance => normally 1-3 keys; cheap even if that
# grows. Fault-tested implicitly by the same EXPECT_BUCKET=bogus drill used for Layer 2 (every
# key HEAD fails against a nonexistent bucket, so this layer reports alongside it).
ASSETS=$(python3 - "$ENV_FILE" "$EXPECT_ENDPOINT" "$EXPECT_BUCKET" <<'PY' 2>/dev/null
import subprocess, sys
env_file, ep, bucket = sys.argv[1], sys.argv[2], sys.argv[3]
try:
    import boto3
    from botocore.config import Config
    env = {}
    for line in open(env_file):
        line = line.strip()
        if line and not line.startswith('#') and '=' in line:
            k, v = line.split('=', 1)
            env[k.replace('export ', '').strip()] = v.strip().strip('"').strip("'")
    s3 = boto3.client('s3', endpoint_url=ep,
                      aws_access_key_id=env['MINIO_ROOT_USER'],
                      aws_secret_access_key=env['MINIO_ROOT_PASSWORD'],
                      region_name='us-east-1',
                      config=Config(s3={'addressing_style': 'path'},
                                    connect_timeout=10, read_timeout=15,
                                    retries={'max_attempts': 1}))

    def query(sql):
        out = subprocess.run(
            ['docker', 'exec', 'cap-mysql', 'mysql', '-ucap', f"-p{env['MYSQL_PASSWORD']}", 'cap', '-N', '-e', sql],
            capture_output=True, text=True, timeout=15)
        return [l for l in out.stdout.splitlines() if l.strip()]

    keys = set()
    keys.update(query("SELECT image FROM users WHERE image IS NOT NULL AND image != '';"))
    keys.update(query("SELECT iconUrl FROM organizations WHERE iconUrl IS NOT NULL AND iconUrl != '';"))
    keys.update(query("SELECT iconUrl FROM spaces WHERE iconUrl IS NOT NULL AND iconUrl != '';"))

    missing = []
    for k in sorted(keys):
        try:
            s3.head_object(Bucket=bucket, Key=k)
        except Exception:
            missing.append(k)
    print('OK' if not missing else 'MISSING:' + ','.join(missing))
except Exception as e:
    print(type(e).__name__ + ':' + str(e)[:200])
PY
)
case "$ASSETS" in
  OK) ;;
  MISSING:*) problems+=("stale image reference(s), object missing from bucket: ${ASSETS#MISSING:}") ;;
  *) problems+=("asset-integrity check failed to run: ${ASSETS:-no-response}") ;;
esac

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
