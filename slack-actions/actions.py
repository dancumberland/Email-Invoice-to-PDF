#!/usr/bin/env python3
"""
DCL Actions — dispatch registry. Maps {stream}:{verb} -> handler(target, ctx) -> str.

Handlers run in a background thread (the HTTP request is already ACKed to Slack).
Each returns a short result string and posts user-facing output via slack_post.reply.

Safety tiers (see CoreContext/Slack_Action_Layer_Spec.md):
  green  — idempotent / reversible / no external party → runs immediately
  yellow — touches money/people/public → produces a draft + Confirm button
  red    — never fires the destructive form

Adding a verb: write a function, decorate with @handler("stream:verb").
"""

import os
import subprocess
import time

import slack_post

LOCK_DIR = "/tmp/slack-actions-locks"
os.makedirs(LOCK_DIR, exist_ok=True)

_HANDLERS = {}


def handler(key):
    def deco(fn):
        _HANDLERS[key] = fn
        return fn
    return deco


def run(stream, verb, target, ctx):
    """Look up and run a handler, with a 10-min double-tap lock per target."""
    key = f"{stream}:{verb}"
    fn = _HANDLERS.get(key)
    if not fn:
        slack_post.reply(ctx, f":warning: No handler wired for `{key}`.")
        return f"no handler: {key}"
    lock = os.path.join(LOCK_DIR, (key + "_" + (target or "all")).replace(":", "_").replace("/", "_"))
    if os.path.exists(lock) and (time.time() - os.path.getmtime(lock) < 600):
        slack_post.reply(ctx, f":hourglass_flowing_sand: `{key}` is already running — ignoring the extra tap.")
        return "locked"
    open(lock, "w").close()
    try:
        return fn(target, ctx)
    finally:
        try:
            os.remove(lock)
        except OSError:
            pass


# --- Phase 0: connectivity check -------------------------------------------
@handler("system:ping")
def _ping(target, ctx):
    slack_post.reply(ctx, ":white_check_mark: pong — listener is verified, signed, and allowlisted.")
    return "pong"


# --- Phase 1: nightly publish retry (green / idempotent) -------------------
BATCH_PUBLISH = "/home/claude/Dan-Cumberland-Labs-Content/Operations/Tools/batch_publish.py"
BATCH_PUBLISH_CWD = "/home/claude/Dan-Cumberland-Labs-Content"


@handler("nightly_publish:retry_failed")
def _retry_failed(target, ctx):
    slack_post.reply(ctx, ":arrows_counterclockwise: Retrying the failed slugs…")
    try:
        proc = subprocess.run(
            ["/usr/bin/python3", BATCH_PUBLISH, "--retry-failed"],
            cwd=BATCH_PUBLISH_CWD, capture_output=True, text=True, timeout=900,
        )
        tail = (proc.stdout or proc.stderr or "").strip().splitlines()[-4:]
        summary = "\n".join(tail) if tail else "(no output)"
        emoji = ":white_check_mark:" if proc.returncode == 0 else ":x:"
        slack_post.reply(ctx, f"{emoji} Retry finished (rc={proc.returncode}):\n```{summary}```")
        return f"rc={proc.returncode}"
    except subprocess.TimeoutExpired:
        slack_post.reply(ctx, ":x: Retry timed out after 15 min — check the publisher log.")
        return "timeout"
    except Exception as e:
        slack_post.reply(ctx, f":x: Retry errored: `{e!r}`")
        return f"error: {e!r}"


@handler("nightly_publish:mute")
def _mute(target, ctx):
    slack_post.reply(ctx, ":no_bell: Muted — left the failures as-is for this run.")
    return "muted"


# --- Phase 2: Lane 1 queue refill (green / idempotent) ---------------------
# load_lane1_queue.py is an idempotent, additive fill-to-floor loader (ON CONFLICT
# DO NOTHING, only mutates 'queued' rows). Run WITHOUT --rerank so it never churns.
LANE1_CWD = "/home/claude/Dan-Cumberland-Labs-Content"
LANE1_REFILL = ["/usr/bin/python3", "Operations/Tools/load_lane1_queue.py"]


@handler("lane1:refill")
def _lane1_refill(target, ctx):
    slack_post.reply(ctx, ":inbox_tray: Refilling the Lane 1 queue…")
    try:
        proc = subprocess.run(
            LANE1_REFILL, cwd=LANE1_CWD, capture_output=True, text=True, timeout=300,
        )
        tail = (proc.stdout or proc.stderr or "").strip().splitlines()[-4:]
        summary = "\n".join(tail) if tail else "(no output)"
        emoji = ":white_check_mark:" if proc.returncode == 0 else ":x:"
        slack_post.reply(ctx, f"{emoji} Refill finished (rc={proc.returncode}):\n```{summary}```")
        return f"rc={proc.returncode}"
    except subprocess.TimeoutExpired:
        slack_post.reply(ctx, ":x: Refill timed out after 5 min — check Operations/queue_load.log.")
        return "timeout"
    except Exception as e:
        slack_post.reply(ctx, f":x: Refill errored: `{e!r}`")
        return f"error: {e!r}"


# --- Phase 2: Redwood content loop — unpark / kill (green / reversible) -----
# Both are tiny additive subcommands on content_loop.py. target == the slug.
REDWOOD_CWD = "/home/claude/redwood-content-system/Operations/Tools"


def _redwood_cmd(subcmd, slug, ctx, doing, done):
    if not slug:
        slack_post.reply(ctx, ":warning: No slug on this action — nothing to run.")
        return "no slug"
    slack_post.reply(ctx, f"{doing} `{slug}`…")
    try:
        proc = subprocess.run(
            ["/usr/bin/python3", "content_loop.py", subcmd, "--slug", slug],
            cwd=REDWOOD_CWD, capture_output=True, text=True, timeout=120,
        )
        out = (proc.stdout or proc.stderr or "").strip().splitlines()
        line = out[-1] if out else "(no output)"
        emoji = ":white_check_mark:" if proc.returncode == 0 else ":x:"
        slack_post.reply(ctx, f"{emoji} {done} (rc={proc.returncode}): {line}")
        return f"rc={proc.returncode}"
    except Exception as e:
        slack_post.reply(ctx, f":x: `{subcmd}` errored: `{e!r}`")
        return f"error: {e!r}"


@handler("redwood:unpark")
def _redwood_unpark(target, ctx):
    return _redwood_cmd("unpark", target, ctx, ":arrow_forward: Unparking", "Unparked")


@handler("redwood:kill")
def _redwood_kill(target, ctx):
    return _redwood_cmd("suppress", target, ctx, ":no_bell: Killing re-notify for", "Killed")
