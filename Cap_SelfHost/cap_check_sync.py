#!/usr/bin/env python3
"""Check (and fix) A/V sync on Cap Desktop studio recordings.

Background — the bug this guards against (found 2026-07-24, Cap 0.5.7):

When Cap starts a studio recording, each source stamps the moment its first
frame/sample arrived (`start_time` in recording-meta.json). Usually all three
land on the same instant. When the screen capture is slow to hand over its
first frame, the screen's start_time lands hundreds of milliseconds after the
mic's and the camera's. The editor then anchors the timeline on the LAST
source to start and writes catch-up offsets into project-config.json
(`clips[0].offsets`).

The mic offset is applied once, correctly. The camera offset is applied TWICE
in the render, so the camera window runs ahead of the voice by the size of the
offset. Measured on the 2026-07-24 Pacemark recording: camera offset 776 ms,
camera rendered 1550 ms early, face about three quarters of a second ahead of
the audio. Setting the stored camera offset to 0 renders it correctly at
776 ms, verified by cross-correlating the export against the source tracks.

Usage:
    python3 cap_check_sync.py                 # report on the 5 newest recordings
    python3 cap_check_sync.py --fix           # also zero a bad camera offset on the newest
    python3 cap_check_sync.py --fix PATH.cap  # fix a specific recording

After --fix, re-export:
    /Applications/Cap.app/Contents/MacOS/cap-cli export "PATH.cap" \
        -o ~/Desktop/fixed.mp4 --resolution 1280x720 --quality web
"""
import json
import shutil
import sys
from pathlib import Path

RECORDINGS = Path.home() / "Library/Application Support/so.cap.desktop/recordings"


def load(path, name):
    try:
        return json.loads((path / name).read_text())
    except Exception:
        return None


def report(cap_dir):
    meta = load(cap_dir, "recording-meta.json")
    cfg = load(cap_dir, "project-config.json")
    print(f"\n{cap_dir.name}")
    if not meta or not meta.get("segments"):
        print("  no segment metadata (legacy recording) — skipping")
        return False

    starts = {}
    for key in ("display", "camera", "mic", "system_audio"):
        src = meta["segments"][0].get(key)
        if isinstance(src, dict) and src.get("start_time") is not None:
            starts[key] = src["start_time"]
    for key, val in starts.items():
        print(f"  {key:<13} starts at {val:8.3f}s")
    spread = (max(starts.values()) - min(starts.values())) * 1000 if starts else 0
    print(f"  spread: {spread:.0f} ms" + ("  (sources aligned — nothing to offset)" if spread < 30 else ""))

    offsets = (cfg or {}).get("clips", [{}])[0].get("offsets", {}) if cfg else {}
    if offsets:
        print("  stored offsets: " + ", ".join(f"{k}={v}" for k, v in offsets.items()))

    bad = bool(offsets.get("camera"))
    if bad:
        ms = offsets["camera"] * 1000
        print(f"  ⚠️  camera offset {ms:.0f} ms will be applied twice — the face will run "
              f"~{ms:.0f} ms ahead of the voice. Fix with --fix.")
    elif "camera" in starts:
        print("  ✅ camera offset is zero — export will be in sync")
    return bad


def fix(cap_dir):
    path = cap_dir / "project-config.json"
    cfg = json.loads(path.read_text())
    clip = cfg["clips"][0]
    before = clip["offsets"].get("camera")
    if not before:
        print(f"nothing to fix in {cap_dir.name} (camera offset already {before})")
        return
    backup = path.with_suffix(".json.bak-presyncfix")
    if not backup.exists():
        shutil.copy2(path, backup)
    clip["offsets"]["camera"] = 0.0
    path.write_text(json.dumps(cfg, indent=2))
    print(f"fixed {cap_dir.name}: camera offset {before} -> 0.0  (backup: {backup.name})")
    print("re-export with:\n  /Applications/Cap.app/Contents/MacOS/cap-cli export "
          f'"{cap_dir}" -o ~/Desktop/fixed.mp4 --resolution 1280x720 --quality web')


def main():
    args = sys.argv[1:]
    recent = sorted(RECORDINGS.glob("*.cap"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not recent:
        print(f"no recordings found in {RECORDINGS}")
        return

    if "--fix" in args:
        rest = [a for a in args if a != "--fix"]
        target = Path(rest[0]) if rest else recent[0]
        report(target)
        fix(target)
    else:
        for cap_dir in recent[:5]:
            report(cap_dir)


if __name__ == "__main__":
    main()
