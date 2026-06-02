#!/usr/bin/env python3
"""
Post a decision card (Block Kit buttons) to #decisions.

Imported by notifier scripts (e.g. batch_publish.py) to surface an actionable
item with one-tap buttons. Reads its own config so callers never thread secrets:
  SLACK_BOT_TOKEN, DECISIONS_CHANNEL_ID  — from /home/claude/slack-actions/.env
(env vars win over the file, for local testing).

Each button's `value` carries the routing the listener parses:
  {"stream": <s>, "verb": <v>, "target": <t>}

Usage:
    import sys; sys.path.append("/home/claude/slack-actions")
    import slack_decisions
    slack_decisions.post_decision(
        stream="nightly_publish",
        title=":memo: Nightly Publish — 3 failed",
        context="blog-x, blog-y, blog-z",
        actions=[
            {"verb": "retry_failed", "label": ":arrows_counterclockwise: Retry failed",
             "style": "primary", "target": "2026-06-02"},
            {"verb": "mute", "label": ":no_bell: Mute"},
        ],
    )
"""

import json
import os
import urllib.request

_ENV_PATH = "/home/claude/slack-actions/.env"


def _load_cfg():
    cfg = {}
    try:
        with open(_ENV_PATH) as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                cfg[k.strip()] = v.strip().strip('"').strip("'")
    except FileNotFoundError:
        pass
    for k in ("SLACK_BOT_TOKEN", "DECISIONS_CHANNEL_ID"):
        if os.environ.get(k):
            cfg[k] = os.environ[k]
    return cfg


def post_decision(stream, title, context="", actions=None):
    """Returns the Slack API response text, or None if unconfigured/failed."""
    cfg = _load_cfg()
    token = cfg.get("SLACK_BOT_TOKEN", "")
    channel = cfg.get("DECISIONS_CHANNEL_ID", "")
    if not token or not channel:
        return None

    blocks = [{"type": "section", "text": {"type": "mrkdwn", "text": f"*{title}*"}}]
    if context:
        blocks.append({"type": "context", "elements": [{"type": "mrkdwn", "text": context}]})

    elements = []
    for a in (actions or []):
        btn = {
            "type": "button",
            "text": {"type": "plain_text", "text": a["label"], "emoji": True},
            "action_id": f"{stream}_{a['verb']}",
            "value": json.dumps({"stream": stream, "verb": a["verb"], "target": a.get("target", "")}),
        }
        if a.get("style"):
            btn["style"] = a["style"]  # "primary" | "danger"
        elements.append(btn)
    if elements:
        blocks.append({"type": "actions", "elements": elements})

    payload = {"channel": channel, "text": title, "blocks": blocks}
    req = urllib.request.Request(
        "https://slack.com/api/chat.postMessage",
        data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {token}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.read().decode("utf-8", "replace")
    except Exception:
        return None
