#!/usr/bin/env python3
"""
DCL Actions — Slack interactivity listener.

Receives Slack Block Kit button clicks (block_actions) at POST /slack/actions,
verifies the request signature, allowlists Dan's user id, ACKs within Slack's
3-second window, then runs the mapped action in a background thread and posts the
result back into the message thread.

Mirrors the audiopen-webhook pattern (stdlib http.server, no framework, no deps).
Secrets come from /home/claude/slack-actions/.env via systemd EnvironmentFile:
  SLACK_SIGNING_SECRET, SLACK_BOT_TOKEN, DAN_USER_ID, DECISIONS_CHANNEL_ID

Design: CoreContext/Slack_Action_Layer_Spec.md
"""

import hashlib
import hmac
import json
import os
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

import actions as action_registry

PORT = int(os.environ.get("SLACK_ACTIONS_PORT", "9203"))
SIGNING_SECRET = os.environ.get("SLACK_SIGNING_SECRET", "").encode()
DAN_USER_ID = os.environ.get("DAN_USER_ID", "")
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(LOG_DIR, exist_ok=True)


def _now_iso():
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def log(event, **fields):
    rec = {"ts": _now_iso(), "event": event, **fields}
    line = json.dumps(rec)
    sys.stderr.write(f"[slack-actions] {line}\n")
    sys.stderr.flush()
    day = time.strftime("%Y-%m-%d", time.gmtime())
    try:
        with open(os.path.join(LOG_DIR, f"slack-actions-{day}.jsonl"), "a") as f:
            f.write(line + "\n")
    except Exception:
        pass


def verify_signature(headers, raw_body):
    """Slack HMAC v0 scheme. Returns (ok: bool, reason: str)."""
    if not SIGNING_SECRET:
        return False, "no signing secret configured"
    ts = headers.get("X-Slack-Request-Timestamp", "")
    sig = headers.get("X-Slack-Signature", "")
    if not ts or not sig:
        return False, "missing signature headers"
    try:
        if abs(time.time() - int(ts)) > 300:
            return False, "stale timestamp"
    except ValueError:
        return False, "bad timestamp"
    base = b"v0:" + ts.encode() + b":" + raw_body
    expected = "v0=" + hmac.new(SIGNING_SECRET, base, hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected, sig):
        return False, "signature mismatch"
    return True, "ok"


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # we do our own structured logging

    def _send(self, code, body=b"", ctype="text/plain"):
        if isinstance(body, str):
            body = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if body:
            self.wfile.write(body)

    def do_GET(self):
        if urlparse(self.path).path == "/healthz":
            self._send(200, b'{"status":"ok"}', "application/json")
        else:
            self._send(404, b"not found")

    def do_POST(self):
        if urlparse(self.path).path != "/slack/actions":
            self._send(404, b"not found")
            return

        length = int(self.headers.get("Content-Length", 0) or 0)
        raw = self.rfile.read(length) if length else b""

        # Events-API URL-verification handshake (JSON body). Harmless for the
        # interactivity-only setup; kept for forward-compat. Still signed.
        if self.headers.get("Content-Type", "").startswith("application/json"):
            try:
                data = json.loads(raw or b"{}")
            except ValueError:
                data = {}
            if data.get("type") == "url_verification":
                ok, reason = verify_signature(self.headers, raw)
                if not ok:
                    log("url_verify_rejected", reason=reason)
                    self._send(401, b"unauthorized")
                    return
                log("url_verification")
                self._send(200, json.dumps({"challenge": data.get("challenge", "")}).encode(),
                           "application/json")
                return

        ok, reason = verify_signature(self.headers, raw)
        if not ok:
            log("rejected_signature", reason=reason)
            self._send(401, b"unauthorized")
            return

        # Interaction payloads arrive form-encoded as payload=<json>.
        qs = parse_qs(raw.decode("utf-8", "replace"))
        payload_raw = (qs.get("payload") or [""])[0]
        if not payload_raw:
            self._send(400, b"no payload")
            return
        try:
            payload = json.loads(payload_raw)
        except ValueError:
            self._send(400, b"bad payload")
            return

        user_id = (payload.get("user") or {}).get("id", "")
        if DAN_USER_ID and user_id != DAN_USER_ID:
            # Ack so Slack shows no error; silently ignore non-allowlisted actors.
            log("rejected_user", user=user_id)
            self._send(200, b"")
            return

        # ACK within Slack's 3s window, then do the work asynchronously.
        self._send(200, b"")
        threading.Thread(target=self._dispatch, args=(payload,), daemon=True).start()

    def _dispatch(self, payload):
        try:
            ctx = {
                "channel": (payload.get("channel") or {}).get("id"),
                "thread_ts": (payload.get("message") or {}).get("ts"),
                "response_url": payload.get("response_url"),
                "user": (payload.get("user") or {}).get("id"),
            }
            for action in payload.get("actions", []):
                try:
                    spec = json.loads(action.get("value") or "{}")
                except ValueError:
                    spec = {}
                stream = spec.get("stream", "")
                verb = spec.get("verb", "")
                target = spec.get("target", "")
                log("dispatch", key=f"{stream}:{verb}", target=target, user=ctx["user"])
                result = action_registry.run(stream, verb, target, ctx)
                log("dispatch_result", key=f"{stream}:{verb}", result=str(result)[:200])
        except Exception as e:
            log("dispatch_error", error=repr(e))


def main():
    log("starting", port=PORT, has_secret=bool(SIGNING_SECRET), allowlist=bool(DAN_USER_ID))
    ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()


if __name__ == "__main__":
    main()
