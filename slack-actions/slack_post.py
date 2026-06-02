#!/usr/bin/env python3
"""Thin Slack poster (stdlib only) — chat.postMessage + response_url fallback."""

import json
import os
import urllib.request

BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN", "")


def _post_json(url, payload, headers):
    req = urllib.request.Request(
        url, data=json.dumps(payload).encode(), headers=headers, method="POST"
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        return resp.read().decode("utf-8", "replace")


def post_message(channel, text, thread_ts=None, blocks=None):
    if not BOT_TOKEN:
        return None
    payload = {"channel": channel, "text": text}
    if thread_ts:
        payload["thread_ts"] = thread_ts
    if blocks:
        payload["blocks"] = blocks
    return _post_json(
        "https://slack.com/api/chat.postMessage", payload,
        {"Content-Type": "application/json", "Authorization": f"Bearer {BOT_TOKEN}"},
    )


def reply(ctx, text):
    """Reply into the originating message's thread; fall back to response_url."""
    channel = ctx.get("channel")
    if BOT_TOKEN and channel:
        try:
            return post_message(channel, text, thread_ts=ctx.get("thread_ts"))
        except Exception:
            pass
    url = ctx.get("response_url")
    if url:
        try:
            return _post_json(url, {"text": text, "response_type": "in_channel"},
                              {"Content-Type": "application/json"})
        except Exception:
            pass
    return None
