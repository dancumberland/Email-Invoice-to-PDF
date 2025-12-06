# Session Documentation Guide

## Session File Naming Convention

All session files must follow this format:

```
YYMMDD.HHMM-Session_Name.md
```

### Format Breakdown

- **YYMMDD**: Year, Month, Day (2-digit year, 2-digit month, 2-digit day)
  - Example: `251126` = November 26, 2025
- **HHMM**: Hour, Minute (24-hour format)
  - Example: `1745` = 5:45 PM
- **Session_Name**: Descriptive name using Title Case and hyphens
  - Use hyphens to separate words (no underscores or spaces)
  - Be descriptive but concise
  - Examples:
    - `Competitor-Analysis`
    - `Reusable-System-Validation`
    - `Content-System-Sync`

### Examples

✅ **CORRECT**
- `251126.1745-Competitor-Analysis.md`
- `251126.1200-Initial-Site-Audit.md`
- `251120.0900-Keyword-Research-Setup.md`

❌ **INCORRECT**
- `251126-SESSION-Reusable-System-Validation.md` (missing time)
- `session_20251126_174541_competitor_analysis.md` (wrong format)
- `251126.1745_Competitor_Analysis.md` (underscores instead of hyphens)

## Session Document Structure

Every session must include:

1. **Header** (first 5 lines)
   ```markdown
   # Session: {Brief Title}
   **Date**: 2025-11-26
   **Time**: 17:45
   **Duration**: {Approx. time spent}
   **Status**: COMPLETE / IN_PROGRESS / BLOCKED
   ```

2. **Quick Summary** (1-2 sentences)
   - What was the main goal
   - What was accomplished

3. **Work Completed** (detailed sections)
   - Major tasks with subsections
   - Code changes with impact
   - Git commits

4. **Key Findings** (if applicable)
   - Important discoveries
   - Decisions made

5. **Files Created/Modified**
   - List all changes
   - Note client-specific vs reusable code

6. **Next Steps**
   - What's pending
   - TODOs for next session
   - Blockers or clarifications needed

7. **Git Status**
   - Commits made
   - Branches affected
   - Remote status

## Session Index

The `SESSIONS.md` file (in the sessions/ folder) serves as master index:

- Lists all sessions chronologically (newest first)
- One-line summary per session
- Quick links for navigation
- Cumulative status of project

**Update the index every time you create a new session.**

## Project Backlog

The `PROJECT_BACKLOG.md` file (at project root) serves as persistent TODO tracking:

- **NEXT**: Current or immediate work (this or next session)
- **SOON**: 1-2 week horizon items
- **BACKLOG**: Lower-priority future work
- **IDEAS**: Exploratory/unvalidated concepts
- **BLOCKED**: Work waiting on external factors
- **Completed**: Items finished this session

The backlog is organized by priority and timeline, allowing you to plan multiple tasks while staying focused on one at a time. Each item includes:
- Priority level (CRITICAL/HIGH/MEDIUM/LOW)
- Estimated effort
- Status and dependencies
- Clear deliverables

**Update the backlog during sessions**: Add emerging TODOs, move items between sections as priorities shift, move completed items to "Completed" section at session end.

## Quick Start for Next Session

When starting a new session:

1. **Check `SESSIONS.md`** (5-minute overview)
   - Read most recent session summary
   - Review "Quick Start" section
   - Check "Next Steps" from last session

2. **Check `PROJECT_BACKLOG.md`** (detailed planning)
   - Review NEXT section for current work
   - Validate all dependencies are ready
   - Check for any blocked items that may be unblocked

3. **Start work** on top NEXT item

This two-part approach provides immediate context (SESSIONS.md) plus detailed planning context (PROJECT_BACKLOG.md) without needing to read full session logs.
