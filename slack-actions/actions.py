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
        # batch_publish.py shells out to the `claude` CLI, which lives in
        # ~/.local/bin — not on systemd's default PATH. Without this prepend the
        # retry dies instantly with "No such file or directory: 'claude'" yet
        # still exits rc=0 (false success). The nightly cron works only because
        # it exports the same PATH inline. (Fixed 2026-06-06.)
        env = {**os.environ, "PATH": "/home/claude/.local/bin:" + os.environ.get("PATH", "")}
        proc = subprocess.run(
            ["/usr/bin/python3", BATCH_PUBLISH, "--retry-failed"],
            cwd=BATCH_PUBLISH_CWD, capture_output=True, text=True, timeout=900,
            env=env,
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


# ---------------------------------------------------------------- fleet watchdog (Tier 2)
# Appended by the fleet-watchdog build. Handlers for #decisions cards posted by
# sweep.py. Trust boundary: `target` is ONLY ever a registry key — never shell data.
import sys as _sys
import json as _json
if "/home/claude/fleet-watchdog" not in _sys.path:
    _sys.path.insert(0, "/home/claude/fleet-watchdog")

_FLEET_STATE_DIR = os.path.expanduser("~/.local/state/fleet-watchdog")


def _fleet_job(target, ctx):
    try:
        import registry
    except Exception as e:
        slack_post.reply(ctx, f":x: fleet registry import failed: `{e!r}`")
        return None
    job = registry.by_name(target)
    if not job:
        slack_post.reply(ctx, f":warning: Unknown job `{target}` — refusing.")
    return job


@handler("fleet_watchdog:restart")
def _fleet_restart(target, ctx):
    job = _fleet_job(target, ctx)
    if not job:
        return f"unknown job: {target}"
    if "restart" not in (job.get("actions") or []) or job.get("signal") != "process":
        slack_post.reply(ctx, f":no_entry: `{target}` is not restart-enabled.")
        return "not restartable"
    cmd = job.get("restart_cmd")
    if not (isinstance(cmd, list) and cmd):
        slack_post.reply(ctx, f":x: `{target}` has no restart_cmd.")
        return "no restart_cmd"
    slack_post.reply(ctx, f":arrows_counterclockwise: Restarting `{target}`…")
    env = dict(os.environ)
    env.setdefault("XDG_RUNTIME_DIR", f"/run/user/{os.getuid()}")
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=60, env=env)
    except Exception as e:
        slack_post.reply(ctx, f":x: Restart errored: `{e!r}`")
        return f"error: {e!r}"
    if proc.returncode != 0:
        tail = (proc.stderr or proc.stdout or "").strip()[-300:]
        slack_post.reply(ctx, f":x: Restart failed (rc={proc.returncode}): ```{tail or '(no output)'}```")
        return f"rc={proc.returncode}"
    # a restart that doesn't STAY up is worse than no restart — verify liveness
    time.sleep(4)
    try:
        import signals
        res = signals.check_process(
            {"service": job.get("service"), "process_kind": "keepalive"},
            time.time(), None, False)
        alive, detail = (res.status == "ok"), res.detail
    except Exception as e:
        alive, detail = True, f"(liveness check error: {e!r})"
    if alive:
        slack_post.reply(ctx, f":white_check_mark: `{target}` restarted and is up ({detail}).")
        return "restarted-ok"
    slack_post.reply(ctx, f":rotating_light: `{target}` restarted but it's *down again* "
                          f"({detail}) — not a transient fault, needs you.")
    return "restarted-died"


@handler("fleet_watchdog:show_diagnosis")
def _fleet_diagnose(target, ctx):
    job = _fleet_job(target, ctx)
    if not job:
        return f"unknown job: {target}"
    slack_post.reply(ctx, f":mag: Diagnosing `{target}`… (advisory)")
    try:
        import signals, diagnose
        from pathlib import Path
        tail, p = "", job.get("path")
        if p and Path(os.path.expanduser(p)).exists():
            tail = signals._read_tail(Path(os.path.expanduser(p)), 6000)
        text = diagnose.diagnose(job, tail, timeout=90)
    except Exception as e:
        slack_post.reply(ctx, f":x: Diagnosis errored: `{e!r}`")
        return f"error: {e!r}"
    slack_post.reply(ctx, f":speech_balloon: *Diagnosis — {target}* (advisory, AI — verify before acting):\n{text}")
    return "diagnosed"


@handler("fleet_watchdog:mute")
def _fleet_mute(target, ctx):
    job = _fleet_job(target, ctx)
    if not job:
        return f"unknown job: {target}"
    host = job.get("host", "vps")
    path = os.path.join(_FLEET_STATE_DIR, f"mutes-{host}.json")
    os.makedirs(_FLEET_STATE_DIR, exist_ok=True)
    try:
        data = _json.load(open(path)) if os.path.exists(path) else {}
    except Exception:
        data = {}
    jobs = data.get("jobs") or {}
    jobs[target] = time.time() + 24 * 3600
    data["jobs"] = jobs
    tmp = path + ".tmp"
    with open(tmp, "w") as fh:
        _json.dump(data, fh, indent=2)
    os.replace(tmp, path)  # atomic; this handler is the sole writer of mutes-*.json
    slack_post.reply(ctx, f":no_bell: Muted `{target}` for 24h "
                          f"(alerts pause; a recovery still posts).")
    return "muted"


# ---------------------------------------------------------------- BLAB pre-call research
# Closes the gap where research-lead never wrote research_status=researched
# back to the sheet. blab_notifier.py posts a Research button to #decisions
# on new AI Strategy discovery bookings (package 279657/284532); this runs
# the unattended pipeline end to end: booking lookup -> headless claude -p
# agent -> sentinel-gated sheet write-back + calendar update.
# research_bridge.py resolves its own claude binary by absolute path
# (~/.local/bin/claude), so it's unaffected by this service's minimal
# systemd PATH — confirmed against the actual unit file before wiring this.
RESEARCH_BRIDGE_DIR = "/home/claude/blab-research-bridge"
RESEARCH_BRIDGE_SCRIPT = os.path.join(RESEARCH_BRIDGE_DIR, "research_bridge.py")
RESEARCH_BRIDGE_TIMEOUT = 1600  # research_bridge.py's own agent timeout is 1500s


@handler("research:run")
def _research_run(target, ctx):
    booking_id = target
    if not booking_id:
        slack_post.reply(ctx, ":warning: No booking id on this action — nothing to research.")
        return "no target"
    slack_post.reply(ctx, f":mag: Running pre-call research for booking `{booking_id}`… (can take up to ~25 min)")
    try:
        proc = subprocess.run(
            ["/usr/bin/python3", RESEARCH_BRIDGE_SCRIPT, "run", "--booking-id", booking_id],
            cwd=RESEARCH_BRIDGE_DIR, capture_output=True, text=True, timeout=RESEARCH_BRIDGE_TIMEOUT,
        )
        # research_bridge.py already prints a Slack-ready mrkdwn summary to
        # stdout on success or failure alike — relay it as-is.
        output = (proc.stdout or "").strip()
        if proc.returncode != 0:
            tail = (proc.stderr or "").strip()[-500:]
            slack_post.reply(ctx, f":x: Research failed for `{booking_id}` (rc={proc.returncode}):\n```{tail or output[-300:] or '(no output)'}```")
            return f"rc={proc.returncode}"
        slack_post.reply(ctx, output or f":warning: Research for `{booking_id}` finished with no output.")
        return "rc=0"
    except subprocess.TimeoutExpired:
        slack_post.reply(ctx, f":x: Research for `{booking_id}` timed out after {RESEARCH_BRIDGE_TIMEOUT}s — check the VPS process (`ps aux | grep research_bridge`).")
        return "timeout"
    except Exception as e:
        slack_post.reply(ctx, f":x: Research handler errored for `{booking_id}`: `{e!r}`")
        return f"error: {e!r}"


@handler("research:skip")
def _research_skip(target, ctx):
    slack_post.reply(ctx, f":no_bell: Skipped — booking `{target}` left un-researched.")
    return "skipped"
