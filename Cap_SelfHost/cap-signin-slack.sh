#!/usr/bin/env bash
# ABOUTME: Forwards self-hosted Cap's web sign-in code to Slack instead of email.
# ABOUTME: Cap (self-hosted) is Resend-only for auth email; with no RESEND_API_KEY it
#          falls back to "Development Mode" and prints the 6-digit sign-in code to the
#          cap-web container log. This service tails that log and pushes the code to
#          Slack #notifications the instant you click "sign in" at hey.dancumberlandlabs.com.
#
# Why this instead of Resend: Dan has no Resend account and didn't want to create one
# (2026-06-22). Slack is infra he already runs; the code is short-lived (10 min) and lands
# in his own ops channel. No third-party dependency, no domain verification, no patch to
# Cap's bundle (so it survives Cap upgrades — we only read the log Cap already writes).
#
# Run as a systemd USER service (cap-signin-slack.service), Restart=always, lingering on.
# Liveness covered by fleet-watchdog registry key "cap-signin-slack" (process keepalive).
#
# Test in isolation:  DRILL=1 ./cap-signin-slack.sh --selftest   (feeds a sample block, prints)
set -uo pipefail

: "${CONTAINER:=cap-web}"
: "${WEBHOOK_FILE:=$HOME/dancumberlandlabs-content/.slack-webhook}"
: "${APP_URL:=https://hey.dancumberlandlabs.com}"

post_slack() {
  local code="$1" email="$2" text
  text="🔐 *Cap web sign-in code:* \`${code}\`"$'\n'
  text+="📧 ${email:-your account} · expires in ~10 min · *check Slack, not email*"$'\n'
  text+="Enter it at ${APP_URL}"
  if [ -n "${DRILL:-}" ]; then
    echo "[DRILL] would POST to #notifications:"; printf '%s\n' "$text"; return
  fi
  local url; url=$(head -1 "$WEBHOOK_FILE" 2>/dev/null)
  [ -n "$url" ] || { echo "$(date -u +%FT%TZ) ERROR no slack webhook at $WEBHOOK_FILE" >&2; return; }
  python3 - "$url" "$text" <<'PY' 2>/dev/null
import json, sys, urllib.request
url, text = sys.argv[1], sys.argv[2]
req = urllib.request.Request(url, data=json.dumps({"text": text}).encode(),
                             headers={"Content-Type": "application/json"}, method="POST")
try: urllib.request.urlopen(req, timeout=10)
except Exception as e: print(f"(slack post failed: {e})", file=sys.stderr)
PY
  echo "$(date -u +%FT%TZ) forwarded sign-in code for ${email:-?} to Slack"
}

# Parse a stream of cap-web log lines; emit Slack on each verification-code block.
# Block shape (Cap dev-mode):
#   🔐 VERIFICATION CODE (Development Mode)
#   📧 Email: dc@dancumberlandlabs.com
#   🔢 Code: 991911
#   ⏱  Expires in: 10 minutes
parse_stream() {
  local in_block=0 email="" last_code="" code
  while IFS= read -r line; do
    case "$line" in
      *"VERIFICATION CODE"*) in_block=1; email="" ;;
      *"Email:"*) [ "$in_block" = 1 ] && email=$(printf '%s' "$line" | sed -E 's/.*Email:[[:space:]]*//; s/[[:space:]]*$//') ;;
      *"Code:"*)
        [ "$in_block" = 1 ] || continue
        code=$(printf '%s' "$line" | grep -oE '[0-9]{6}' | head -1)
        in_block=0
        [ -n "$code" ] || continue
        [ "$code" = "$last_code" ] && continue   # de-dupe identical re-reads on reconnect
        last_code="$code"
        post_slack "$code" "$email"
        ;;
    esac
  done
}

if [ "${1:-}" = "--selftest" ]; then
  printf '%s\n' \
    "🔐 VERIFICATION CODE (Development Mode)" \
    "📧 Email: dc@dancumberlandlabs.com" \
    "🔢 Code: 424242" \
    "⏱  Expires in: 10 minutes" | parse_stream
  exit 0
fi

echo "$(date -u +%FT%TZ) cap-signin-slack watching '$CONTAINER' logs -> Slack"
# Follow only NEW lines (--tail 0) so we never re-post a code from before startup.
# Outer loop reconnects if cap-web is restarted/recreated.
while true; do
  docker logs -f --tail 0 "$CONTAINER" 2>&1 | parse_stream
  echo "$(date -u +%FT%TZ) log stream ended (container restart?) — reconnecting in 3s" >&2
  sleep 3
done
