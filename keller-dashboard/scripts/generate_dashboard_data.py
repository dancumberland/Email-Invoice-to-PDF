#!/usr/bin/env python3
"""
Keller Dashboard Data Generator

Parses the Keller Associates project directory and generates data.json
for the dashboard to consume. Uses only Python stdlib.

Usage: python scripts/generate_dashboard_data.py
"""

import json
import os
import re
from datetime import datetime, date
from pathlib import Path

# Paths
KELLER_ROOT = Path.home() / "Documents" / "Work" / "Client_Projects" / "Keller_Associates"
DASHBOARD_ROOT = Path(__file__).parent.parent
DATA_OUT = DASHBOARD_ROOT / "data" / "data.json"
SIGNALS_FILE = DASHBOARD_ROOT / "signals.json"

# Engagement dates
START_DATE = date(2026, 3, 19)
END_DATE = date(2026, 5, 7)

# People data — loaded from Discovery/people.json in the Keller project
PEOPLE_JSON = KELLER_ROOT / "Discovery" / "people.json"

def load_people_data():
    """Load people data from Discovery/people.json."""
    if not PEOPLE_JSON.exists():
        print(f"  WARNING: {PEOPLE_JSON} not found, using empty list")
        return []
    with open(PEOPLE_JSON) as f:
        return json.load(f)

PEOPLE_DATA = load_people_data()

# Readiness assessments (from People.md analysis)
READINESS_DATA = [
    {"name": "Nathan Cleaver", "tier": "champion", "role": "Chief Engineer"},
    {"name": "Brandon Keller", "tier": "champion", "role": "Structural DL"},
    {"name": "Crystal Warner", "tier": "champion", "role": "Proposal Coordinator"},
    {"name": "Andrea Sams", "tier": "early_adopter", "role": "HR Director"},
    {"name": "Kyle Meschko", "tier": "early_adopter", "role": "TI Committee Chair"},
    {"name": "Breck Dalley", "tier": "early_adopter", "role": "Marketing Manager"},
    {"name": "Morgan Cushing", "tier": "early_adopter", "role": "Admin Manager"},
    {"name": "Riley Hoskins", "tier": "pragmatist", "role": "Accounting Manager"},
    {"name": "Tim Trumbo", "tier": "pragmatist", "role": "IT Manager"},
    {"name": "Jason King", "tier": "pragmatist", "role": "W/WW DL"},
    {"name": "James Bledsoe", "tier": "pragmatist", "role": "VP / Board Chair"},
    {"name": "Larry Rupp", "tier": "pragmatist", "role": "CEO"},
    {"name": "Stillman Norton", "tier": "pragmatist", "role": "WA Area Manager"},
    {"name": "Colter Hollingshead", "tier": "skeptic", "role": "Pocatello BOM"},
]

# Deliverables from SOW (+ bonus deliverables for May 7 push)
DELIVERABLES = [
    {
        "id": 1,
        "title": "Technology & Infrastructure Audit",
        "sow_spec": "Written report, 5-10 pages covering each system",
        "status": "drafting",
        "target_date": "2026-04-30",
        "depends_on": [],
        "notes": "Scaffold created 4/20 (Tech_Audit_Scaffold.md). Technology.md + 6 research files populated. IT Glue inventory still pending from Tim.",
        "inputs_needed": [
            {"label": "Tim Trumbo tool audit session #1", "type": "interview", "available": True},
            {"label": "Synoptek IT plan review", "type": "research", "available": True},
            {"label": "Technology.md analysis doc", "type": "analysis", "available": True},
            {"label": "Panzura/LucidLink research", "type": "research", "available": True},
            {"label": "Deltek ecosystem + Vantagepoint scheduling research", "type": "research", "available": True},
            {"label": "Copilot capabilities (Excel, Agents, 2026 features)", "type": "research", "available": True},
            {"label": "Bluebeam batch signing research", "type": "research", "available": True},
            {"label": "Scaffold created (Tech_Audit_Scaffold.md)", "type": "other", "available": True},
            {"label": "IT Glue / full systems inventory", "type": "other", "available": False},
            {"label": "Copilot audit logging confirmation (Tim)", "type": "other", "available": False},
        ],
    },
    {
        "id": 2,
        "title": "AI Readiness Assessment",
        "sow_spec": "Written assessment, 8-12 pages, department by department",
        "status": "drafting",
        "target_date": "2026-04-30",
        "depends_on": [],
        "notes": "Scaffold + all 6 Pacemark dimensions scored + 88 individual tiers. Ready for full draft Apr 21-25.",
        "inputs_needed": [
            {"label": "All Wave 1+2 interviews extracted (89 briefs)", "type": "interview", "available": True},
            {"label": "Wave 3 targeted follow-ups (ElevenLabs + tldv)", "type": "interview", "available": True},
            {"label": "AI Baseline Survey analysis", "type": "research", "available": True},
            {"label": "People.md analysis doc", "type": "analysis", "available": True},
            {"label": "Culture.md analysis doc", "type": "analysis", "available": True},
            {"label": "Workflows.md analysis doc", "type": "analysis", "available": True},
            {"label": "Pacemark dimension scoring (6 dims)", "type": "analysis", "available": True},
            {"label": "Individual tier scoring (88 people)", "type": "analysis", "available": True},
            {"label": "Scaffold (Readiness_Assessment_Scaffold.md)", "type": "other", "available": True},
            {"label": "Non-respondent tier proxy methodology", "type": "analysis", "available": False},
        ],
    },
    {
        "id": 3,
        "title": "Prioritized Workflow Opportunity Map",
        "sow_spec": "Summary page (impact vs effort) + page per department",
        "status": "drafting",
        "target_date": "2026-04-30",
        "depends_on": [],
        "notes": "Scaffold ~80% populated. Eng/Acct/HR/IT strong; Marketing/Admin/Field weaker (data gaps).",
        "inputs_needed": [
            {"label": "Workflows.md analysis doc (92KB)", "type": "analysis", "available": True},
            {"label": "Technology.md for tool capabilities", "type": "analysis", "available": True},
            {"label": "All department interviews extracted", "type": "interview", "available": True},
            {"label": "AEC AI plan review / QC tools research", "type": "research", "available": True},
            {"label": "Deltek Vantagepoint scheduling research", "type": "research", "available": True},
            {"label": "Bluebeam batch signing research", "type": "research", "available": True},
            {"label": "Scaffold (Workflow_Opportunity_Map_Scaffold.md)", "type": "other", "available": True},
            {"label": "Crystal Warner / Breck Dalley interviews (Marketing)", "type": "interview", "available": False},
            {"label": "Field/construction interviews", "type": "interview", "available": False},
        ],
    },
    {
        "id": 4,
        "title": "Implementation Roadmap",
        "sow_spec": "Visual timeline + written companion (4-8 pages)",
        "status": "drafting",
        "target_date": "2026-04-30",
        "depends_on": [1, 2, 3],
        "notes": "Scaffold created 4/20 with four-horizon structure (Foundation/Activation/Integration/Scale). Visual timeline produced at packaging.",
        "inputs_needed": [
            {"label": "Pacemark dimension scoring", "type": "analysis", "available": True},
            {"label": "Scaffold (Implementation_Roadmap_Scaffold.md)", "type": "other", "available": True},
            {"label": "Technology Audit (draft)", "type": "other", "available": False},
            {"label": "Readiness Assessment (draft)", "type": "other", "available": False},
            {"label": "Workflow Opportunity Map (draft)", "type": "other", "available": False},
            {"label": "Budget/resource constraints from Larry", "type": "other", "available": False},
        ],
    },
    {
        "id": 5,
        "title": "Training Program Design",
        "sow_spec": "Workshop outlines, 1-2 pages each, with agenda and exercises",
        "status": "drafting",
        "target_date": "2026-04-30",
        "depends_on": [2],
        "notes": "Scaffold with Workshops A-F outlined. Andrea's Motivosity playbook + Nathan's earned value demo as anchors.",
        "inputs_needed": [
            {"label": "Scaffold (Training_Program_Scaffold.md)", "type": "other", "available": True},
            {"label": "Department-level skill gaps (tier scoring)", "type": "analysis", "available": True},
            {"label": "Copilot capabilities confirmed", "type": "research", "available": True},
            {"label": "Andrea's Motivosity rollout playbook", "type": "analysis", "available": True},
            {"label": "Readiness Assessment (draft)", "type": "other", "available": False},
            {"label": "TI Committee cadence decision", "type": "other", "available": False},
        ],
    },
    {
        "id": 6,
        "title": "Change Management Framework",
        "sow_spec": "Practical playbook, ~10 pages",
        "status": "drafting",
        "target_date": "2026-04-30",
        "depends_on": [2],
        "notes": "Scaffold created 4/20. Champion roster (11 names), 7 comm templates, rollout sequence, resistance handling.",
        "inputs_needed": [
            {"label": "Scaffold (Change_Mgmt_Framework_Scaffold.md)", "type": "other", "available": True},
            {"label": "Culture.md analysis doc", "type": "analysis", "available": True},
            {"label": "11-person champion roster identified", "type": "analysis", "available": True},
            {"label": "Adoption patterns from 89 interviews", "type": "interview", "available": True},
            {"label": "Resistance archetype mapping", "type": "analysis", "available": True},
            {"label": "Readiness Assessment (draft)", "type": "other", "available": False},
            {"label": "Communication templates drafted", "type": "other", "available": False},
        ],
    },
    {
        "id": 7,
        "title": "Interactive Shareholder Workshop",
        "sow_spec": "~2 hours, afternoon session, May 7 at JUMP Boise",
        "status": "not_started",
        "target_date": "2026-05-07",
        "depends_on": [4],
        "notes": "2:15 PM at JUMP Level 4, INSPIRE Studio. Design in week of May 1.",
        "inputs_needed": [
            {"label": "Implementation Roadmap (draft)", "type": "other", "available": False},
            {"label": "TI Committee review (Apr 30)", "type": "other", "available": False},
            {"label": "Keynote deck designed", "type": "other", "available": False},
            {"label": "Workshop exercises designed", "type": "other", "available": False},
        ],
    },
    {
        "id": 8,
        "title": "AI Foundations Workshop Series",
        "sow_spec": "2-3 hands-on sessions, 60-90 min each (complimentary)",
        "status": "not_started",
        "target_date": "2026-05-21",
        "depends_on": [5, 7],
        "notes": "Attendees identified based on shareholder meeting outcomes. Post-May 7.",
        "inputs_needed": [
            {"label": "Training Program Design (final)", "type": "other", "available": False},
            {"label": "Shareholder meeting outcomes", "type": "other", "available": False},
            {"label": "Workshop attendees identified", "type": "other", "available": False},
        ],
    },
    {
        "id": 9,
        "title": "TI Committee Pre-Read (bonus)",
        "sow_spec": "1-page primer before Apr 30 review",
        "status": "drafting",
        "target_date": "2026-04-28",
        "depends_on": [],
        "notes": "Content REFRAMED 4/20 with positive lead (Pacemark Index demoted to context, not centerpiece). Packaging format TBD.",
        "inputs_needed": [
            {"label": "Pacemark dimension scoring", "type": "analysis", "available": True},
            {"label": "Individual tier scoring", "type": "analysis", "available": True},
            {"label": "Champion roster", "type": "analysis", "available": True},
            {"label": "Reframed 1-page content", "type": "other", "available": True},
            {"label": "Cover email (Dan→Larry) drafted", "type": "other", "available": True},
            {"label": "Packaging format decision (PDF/Doc/email)", "type": "other", "available": False},
            {"label": "Gmail draft staged", "type": "other", "available": False},
        ],
    },
    {
        "id": 10,
        "title": "Shareholder Take-Home 1-Pager (bonus)",
        "sow_spec": "Printed handout for May 7 attendees",
        "status": "not_started",
        "target_date": "2026-05-06",
        "depends_on": [4, 7],
        "notes": "Branded with Pacemark. Summarizes roadmap + next steps for shareholders.",
        "inputs_needed": [
            {"label": "Implementation Roadmap (final)", "type": "other", "available": False},
            {"label": "Keynote deck (locked)", "type": "other", "available": False},
            {"label": "Pacemark branding assets", "type": "other", "available": False},
        ],
    },
    {
        "id": 11,
        "title": "Keynote Slide Deck (bonus)",
        "sow_spec": "May 7 presentation, Pacemark spine + top findings",
        "status": "not_started",
        "target_date": "2026-05-05",
        "depends_on": [2, 3, 4],
        "notes": "Built from Readiness + Workflow Map + Roadmap. Interactive workshop structure embedded.",
        "inputs_needed": [
            {"label": "Readiness Assessment (final)", "type": "other", "available": False},
            {"label": "Workflow Opportunity Map (final)", "type": "other", "available": False},
            {"label": "Implementation Roadmap (final)", "type": "other", "available": False},
            {"label": "TI Committee feedback integrated", "type": "other", "available": False},
        ],
    },
    {
        "id": 12,
        "title": "Advisory Sprint Scoping Template (bonus)",
        "sow_spec": "Retainer bridge — aggregates deferred research items",
        "status": "not_started",
        "target_date": "2026-05-05",
        "depends_on": [],
        "notes": "Post-engagement retainer framing. Populates with deferred tool evaluations from 30-min research cap.",
        "inputs_needed": [
            {"label": "Deferred research items catalogued", "type": "analysis", "available": True},
            {"label": "Pacemark dimensional gaps named", "type": "analysis", "available": True},
            {"label": "Retainer scope template", "type": "other", "available": False},
        ],
    },
]

# Research items
RESEARCH_ITEMS = [
    {"topic": "Deltek VantagePoint API / Ecosystem", "priority": "P1", "status": "complete", "file": "Deltek_Ecosystem_Research.md", "feeds": [1, 3]},
    {"topic": "Deltek MCP Server Feasibility", "priority": "P1", "status": "complete", "file": "Deltek_MCP_Server_Feasibility.md", "feeds": [1, 4]},
    {"topic": "Copilot Capabilities Deep Dive", "priority": "P1", "status": "in_progress", "file": "RAG_Architecture_Research.md", "feeds": [1, 2, 5]},
    {"topic": "AI-Powered Plan Review & QC Tools", "priority": "P1", "status": "queued", "file": None, "feeds": [3, 4]},
    {"topic": "AI Security & Data Guidelines", "priority": "P1", "status": "queued", "file": None, "feeds": [2, 6, 7]},
    {"topic": "Joist AI Assessment", "priority": "P1", "status": "complete", "file": "Joist_AI_Research.md", "feeds": [1, 3]},
    {"topic": "Panzura / File Architecture AI-Readiness", "priority": "P1", "status": "complete", "file": "RAG_Architecture_Research.md", "feeds": [1, 4]},
    {"topic": "OpenAsset / DAM Research", "priority": "P2", "status": "complete", "file": "OpenAsset_Research.md", "feeds": [1, 3]},
    {"topic": "LucidLink Deep Research", "priority": "P2", "status": "complete", "file": "LucidLink_Deep_Research.md", "feeds": [1]},
    {"topic": "Synoptek IT Plan Review", "priority": "P2", "status": "complete", "file": "Synoptek_IT_Plan_Review.md", "feeds": [1]},
    {"topic": "AI Baseline Survey Analysis", "priority": "P2", "status": "complete", "file": "AI_Baseline_Survey_Analysis.md", "feeds": [2]},
    {"topic": "Data Pipeline Feasibility", "priority": "P2", "status": "complete", "file": "Data_Pipeline_Feasibility.md", "feeds": [4]},
    {"topic": "Piros Tool Investigation", "priority": "P3", "status": "queued", "file": None, "feeds": [3]},
    {"topic": "Copilot for Excel — AI module capabilities for EVR + spreadsheet workflows", "priority": "P1", "status": "queued", "file": None, "feeds": [1, 3, 4]},
]


def get_brief_names():
    """Get set of names that have extraction briefs."""
    briefs_dir = KELLER_ROOT / "Analysis" / "Briefs"
    if not briefs_dir.exists():
        return set()
    names = set()
    for f in briefs_dir.glob("*.md"):
        # Skip non-person briefs
        if f.stem.startswith("Kickoff") or f.stem.startswith("Larry_Rupp_Pre"):
            continue
        # Convert filename to name: Brandon_Keller -> Brandon Keller
        name = f.stem.replace("_", " ")
        # Handle Larry's discovery brief
        if name == "Larry Rupp Discovery":
            name = "Larry Rupp"
        names.add(name)
    return names


def get_routed_names():
    """Check which people have been routed into analysis docs."""
    analysis_dir = KELLER_ROOT / "Analysis"
    analysis_files = ["Workflows.md", "Technology.md", "Culture.md", "People.md"]
    routed = set()
    for fname in analysis_files:
        fpath = analysis_dir / fname
        if fpath.exists():
            content = fpath.read_text()
            for p in PEOPLE_DATA:
                # Check for last name in analysis docs (more reliable than full name)
                last_name = p["name"].split()[-1]
                if last_name in content:
                    routed.add(p["name"])
    return routed


def get_interview_status(person_name, brief_names, routed_names, today):
    """Determine interview status for a person."""
    person_data = next((p for p in PEOPLE_DATA if p["name"] == person_name), None)
    if not person_data or not person_data["interview_date"]:
        return "scheduled"

    interview_date = date.fromisoformat(person_data["interview_date"])

    if interview_date > today:
        return "scheduled"

    if person_name in brief_names:
        if person_name in routed_names:
            return "routed"
        return "extracted"

    if interview_date <= today:
        return "completed"

    return "scheduled"


def get_pending_pipeline(brief_names, today):
    """Get pipeline items that need extraction or routing."""
    pending = []
    for p in PEOPLE_DATA:
        if not p["interview_date"]:
            continue
        idate = date.fromisoformat(p["interview_date"])
        if idate > today:
            continue
        if p["name"] not in brief_names:
            pending.append({
                "person": p["name"],
                "transcript_date": p["interview_date"],
                "stage": "transcript",
                "brief_path": None,
            })
    return pending


def get_upcoming_interviews(today, brief_names):
    """Get interviews that haven't happened yet (future dates only)."""
    upcoming = []
    for p in PEOPLE_DATA:
        if not p["interview_date"]:
            continue
        idate = date.fromisoformat(p["interview_date"])
        # Only show truly future interviews — if the date has passed or it's today
        # and the person already has a brief/transcript, skip it
        if idate < today:
            continue
        if idate == today and p["name"] in brief_names:
            continue
        upcoming.append({
            "person": p["name"],
            "role": p["role"],
            "date": p["interview_date"],
            "time": p.get("time_mt", ""),
            "time_mt": p.get("time_mt", ""),
            "questions_path": f"Discovery/Interview_Questions/{p['name'].split()[0]}.md",
            "profile_path": f"Discovery/Profiles/{p['name'].replace(' ', '_')}.md",
        })
    # Sort by date then time
    upcoming.sort(key=lambda x: (x["date"], x["time_mt"]))
    return upcoming[:10]


def build_coverage_matrix():
    """Build office x discipline coverage matrix."""
    offices = sorted(set(p["office"] for p in PEOPLE_DATA if p["office"] != "Unknown"))
    disciplines = sorted(set(p["discipline"] for p in PEOPLE_DATA if p["discipline"] != "Unknown"))

    matrix = {}
    for office in offices:
        matrix[office] = {}
        for disc in disciplines:
            names = [p["name"] for p in PEOPLE_DATA
                     if p["office"] == office and p["discipline"] == disc]
            if names:
                matrix[office][disc] = names

    return {"offices": offices, "disciplines": disciplines, "matrix": matrix}


def build_waves(brief_names, today):
    """Build wave progress data."""
    waves = []
    for wave_num in [1, 2, 3]:
        wave_people = [p for p in PEOPLE_DATA if p["wave"] == wave_num]
        if not wave_people:
            if wave_num == 3:
                waves.append({
                    "number": 3,
                    "date_range": "Apr 14-22",
                    "people": [],
                    "completed": 0,
                    "total": 0,
                })
            continue

        completed = sum(1 for p in wave_people
                       if p["interview_date"] and date.fromisoformat(p["interview_date"]) <= today)

        date_ranges = {1: "Mar 19-28", 2: "Mar 31 - Apr 11", 3: "Apr 14-22"}

        waves.append({
            "number": wave_num,
            "date_range": date_ranges.get(wave_num, "TBD"),
            "people": [p["name"] for p in wave_people],
            "completed": completed,
            "total": len(wave_people),
        })

    return waves


def load_agent_status():
    """Load ElevenLabs agent monitor state."""
    state_file = KELLER_ROOT / "scripts" / "output" / "agent_monitor_state.json"
    if not state_file.exists():
        return None
    with open(state_file) as f:
        state = json.load(f)

    thanked = state.get("thanked_people", [])
    completed_people = []
    for t in thanked:
        parts = t.split("|")
        completed_people.append({
            "first_name": parts[0] if len(parts) > 0 else "",
            "role": parts[1] if len(parts) > 1 else "",
            "office": parts[2] if len(parts) > 2 else "",
        })

    return {
        "total_conversations": len(state.get("seen_ids", [])),
        "good_conversations": len(thanked),
        "completed_people": completed_people,
        "person_attempts": state.get("person_attempts", {}),
        "last_check": state.get("last_check", ""),
        "phone_number": "+1 (208) 567-9434",
        "phone_active": True,
    }


def load_clicks():
    """Load click tracker data from Discovery/clicks.json."""
    clicks_file = KELLER_ROOT / "Discovery" / "clicks.json"
    if not clicks_file.exists():
        return {
            "last_updated": "",
            "employees_total": 226,
            "unique_clickers": 0,
            "clickers": [],
        }
    with open(clicks_file) as f:
        data = json.load(f)
    data["unique_clickers"] = len(data.get("clickers", []))
    return data


def build_actions():
    """Build action items — current as of 2026-04-13."""
    return [
        {"description": "Send Dianna announcement email — phone option for agent interviews", "owner": "Dan", "waiting_on": None, "priority": "high"},
        {"description": "Nudge email for non-respondents (Apr 17-18)", "owner": "Dan", "waiting_on": None, "priority": "high"},
        {"description": "Follow up with Tim — VPN resend + Copilot + IT Glue track", "owner": "Dan", "waiting_on": "Tim Trumbo", "priority": "high"},
        {"description": "AI Security & Data Guidelines research (board blocker)", "owner": "Dan", "waiting_on": None, "priority": "high"},
        {"description": "Follow up: Ryan (Pocatello) — abnormal termination at 25s", "owner": "Dan", "waiting_on": None, "priority": "medium"},
        {"description": "Follow up: Sanaz — no audio output, emailed 4/10", "owner": "Dan", "waiting_on": "Sanaz Malaki", "priority": "medium"},
        {"description": "Follow up: Christopher Manos — dropped at 42s, emailed 4/10", "owner": "Dan", "waiting_on": "Christopher Manos", "priority": "medium"},
        {"description": "Begin Copilot capabilities deep dive (blocked on Tim VPN)", "owner": "Dan", "waiting_on": "Tim Trumbo", "priority": "medium"},
        {"description": "Start draft deliverables — target week of 4/20 for Larry review", "owner": "Dan", "waiting_on": None, "priority": "medium"},
        {"description": "James Bledsoe — E&O insurance question reply", "owner": "Dan", "waiting_on": "James Bledsoe", "priority": "low"},
        {"description": "Breck Dalley — win rate on proposals (sent 4/3, nudged 4/7)", "owner": "Dan", "waiting_on": "Breck Dalley", "priority": "low"},
    ]


def build_comms():
    """Build recent communication entries."""
    return [
        {"date": "2026-04-10", "type": "email", "with_person": "Larry Rupp", "summary": "Reimbursement request — flight ($959.13) + hotel ($427.05) = $1,386.18"},
        {"date": "2026-04-10", "type": "email", "with_person": "Dianna Smith", "summary": "Agent issues follow-up — Tim Heald VAD fix, Chris Clark mic issue"},
        {"date": "2026-04-07", "type": "email", "with_person": "James Bledsoe", "summary": "E&O insurance question sent — awaiting reply"},
        {"date": "2026-04-07", "type": "email", "with_person": "Riley Hoskins", "summary": "KPI report data + T&M tracking request — awaiting reply"},
        {"date": "2026-04-01", "type": "email", "with_person": "Crystal Warner", "summary": "RFP site list discussion — forwarding to Breck for follow-up"},
    ]


def build_invoices():
    """Build invoice status."""
    return [
        {"milestone": 1, "amount": 5000, "description": "Project kickoff", "due_description": "Upon signing", "status": "sent"},
        {"milestone": 2, "amount": 5000, "description": "Discovery complete", "due_description": "~Apr 22", "status": "upcoming"},
        {"milestone": 3, "amount": 8000, "description": "Shareholder meeting", "due_description": "May 7", "status": "upcoming"},
        {"milestone": 4, "amount": 10000, "description": "Final deliverables", "due_description": "~May 21", "status": "upcoming"},
    ]


def compute_engagement_meta(today):
    """Compute engagement metadata."""
    days_elapsed = (today - START_DATE).days
    current_week = min(7, max(1, (days_elapsed // 7) + 1))
    days_remaining = (END_DATE - today).days

    if today < date(2026, 4, 23):
        phase = "discovery"
    elif today < date(2026, 5, 7):
        phase = "synthesis"
    else:
        phase = "delivery"

    return {
        "name": "Keller Associates AI Roadmap",
        "value": 28000,
        "start_date": START_DATE.isoformat(),
        "end_date": END_DATE.isoformat(),
        "current_week": current_week,
        "days_remaining": days_remaining,
        "phase": phase,
    }


def main():
    today = date.today()
    brief_names = get_brief_names()
    routed_names = get_routed_names()

    print(f"Generating dashboard data...")
    print(f"  Today: {today}")
    print(f"  Briefs found: {len(brief_names)} — {', '.join(sorted(brief_names))}")
    print(f"  Routed to analysis: {len(routed_names)} — {', '.join(sorted(routed_names))}")

    # Build people with status
    people = []
    for p in PEOPLE_DATA:
        status = get_interview_status(p["name"], brief_names, routed_names, today)
        readiness = next((r for r in READINESS_DATA if r["name"] == p["name"]), None)

        people.append({
            "name": p["name"],
            "role": p["role"],
            "office": p["office"],
            "discipline": p["discipline"],
            "wave": p["wave"],
            "interview_date": p["interview_date"],
            "interview_status": status,
            "brief_path": f"Analysis/Briefs/{p['name'].replace(' ', '_')}.md" if p["name"] in brief_names else None,
            "profile_path": f"Discovery/Profiles/{p['name'].replace(' ', '_')}.md",
            "questions_path": f"Discovery/Interview_Questions/{p['name'].split()[0]}.md",
            "readiness_tier": readiness["tier"] if readiness else None,
            "key_insight": None,
        })

    # Build deliverables with computed input counts
    deliverables = []
    for d in DELIVERABLES:
        inputs = d["inputs_needed"]
        deliverables.append({
            "id": d["id"],
            "title": d["title"],
            "sow_spec": d["sow_spec"],
            "status": d["status"],
            "target_date": d["target_date"],
            "inputs_needed": inputs,
            "inputs_available": sum(1 for i in inputs if i["available"]),
            "inputs_total": len(inputs),
            "depends_on": d["depends_on"],
            "notes": d["notes"],
        })

    # Build research
    research = []
    for r in RESEARCH_ITEMS:
        research.append({
            "topic": r["topic"],
            "priority": r["priority"],
            "status": r["status"],
            "file_path": f"Analysis/Research/{r['file']}" if r["file"] else None,
            "feeds_deliverable": r["feeds"],
        })

    # Load signals
    signals = {"findings": [], "themes": [], "quickWins": [], "quotes": [], "contradictions": [], "gaps": []}
    if SIGNALS_FILE.exists():
        with open(SIGNALS_FILE) as f:
            signals = json.load(f)

    # Load agent + click data
    agent = load_agent_status()
    clicks = load_clicks()

    # Assemble
    data = {
        "generated_at": datetime.now().isoformat(),
        "engagement": compute_engagement_meta(today),
        "people": people,
        "coverage": build_coverage_matrix(),
        "readiness": READINESS_DATA,
        "waves": build_waves(brief_names, today),
        "deliverables": deliverables,
        "research": research,
        "interviews": get_upcoming_interviews(today, brief_names),
        "pipeline": get_pending_pipeline(brief_names, today),
        "actions": build_actions(),
        "comms": build_comms(),
        "invoices": build_invoices(),
        "agent": agent,
        "clicks": clicks,
        **signals,
    }

    # Write output
    DATA_OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(DATA_OUT, "w") as f:
        json.dump(data, f, indent=2)

    print(f"  Output: {DATA_OUT}")
    print(f"  People: {len(people)}")
    print(f"  Interviews done: {sum(1 for p in people if p['interview_status'] != 'scheduled')}")
    print(f"  Pending extractions: {len(data['pipeline'])}")
    print(f"  Deliverables: {len(deliverables)}")
    print(f"  Research items: {len(research)}")
    print(f"  Agent: {agent['total_conversations'] if agent else 0} conversations, {agent['good_conversations'] if agent else 0} completed")
    print(f"  Clicks: {clicks['unique_clickers']} unique clickers of {clicks['employees_total']}")
    print(f"  Signals: {len(signals.get('findings', []))} findings, {len(signals.get('themes', []))} themes")
    print(f"  Done!")


if __name__ == "__main__":
    main()
