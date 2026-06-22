---
description: Gather a client's current status from Gmail, meetings, and project files
---

# Project Update

Gather current status for a client project from all sources.

## Arguments

- `$ARGUMENTS` - Client name

## What This Does

1. **Check Gmail** - Recent emails to/from client
2. **Check TLDV** - Recent meetings mentioning client
3. **Read PROJECT.md** - Current documented status
4. **Surface signals** - Anything needing attention

## Instructions

Given the client name in `$ARGUMENTS`:

### 1. Search Gmail for recent client communications

Use the Gmail MCP tools to search for recent emails:
- `mcp__gmail-dclabs__search-emails` with query like `from:client OR to:client`
- Look for last 7-14 days of activity
- Note any unanswered emails or action items

### 2. Search TLDV for recent meetings

Use TLDV MCP tools:
- `mcp__tldv__list-meetings` to find recent meetings
- Filter by client name in meeting title or participants
- Get highlights if relevant meetings found

### 3. Read current PROJECT.md

```python
import sys
PM_HOME = '/Users/dancumberland/Documents/Work/Project_Management'
sys.path.insert(0, f'{PM_HOME}/lib')
from project_sync import parse_project_md, get_project_path

filepath = get_project_path("CLIENT_NAME")
if filepath.exists():
    data = parse_project_md(filepath)
    # Show current status, last_contact, next_action
```

### 4. Present unified status

Show the user:
- **Last documented contact**: from PROJECT.md
- **Recent emails**: summary of email activity
- **Recent meetings**: any relevant TLDV meetings
- **Next action**: what's documented as next
- **Signals**: anything needing attention (overdue, unanswered, etc.)

Ask if they want to update any of this information before starting work.
