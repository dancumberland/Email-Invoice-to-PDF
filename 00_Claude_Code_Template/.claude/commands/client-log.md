---
description: Log completed work on a client project and update its status
---

# Log Session

**Use AFTER finishing work on a client project.**

This documents what you accomplished and updates the project status.

## Arguments

- `$ARGUMENTS` - Should contain: client name, and optionally a summary

## What This Does

1. Logs the session to the client's PROJECT.md activity log
2. Updates `last_contact` in both PROJECT.md and the database
3. Syncs to ensure central PM database is current
4. Optionally updates next_action

## Instructions

Parse the arguments to extract:
- **client**: The client name (required)
- **summary**: What was accomplished (ask if not provided)
- **next_steps**: What's the next action (ask if relevant)

Then run:

```python
import sys
PM_HOME = '/Users/dancumberland/Documents/Work/Project_Management'
sys.path.insert(0, f'{PM_HOME}/lib')
from project_sync import end_session, list_projects

# Get client and summary from arguments
# $ARGUMENTS will contain something like "Jess - completed onboarding call"

result = end_session(
    client="CLIENT_NAME",  # Extract from arguments
    summary="SUMMARY",      # Extract from arguments or ask user
    next_steps="NEXT_STEPS" # Optional, can be None
)

print(f"✓ Session logged for {result['client']}")
print(f"  Summary: {result['summary']}")
if result.get('next_steps'):
    print(f"  Next: {result['next_steps']}")
print(f"  Updated: {result['project_file']}")
```

If arguments are unclear, ask the user:
1. Which client?
2. What was accomplished?
3. What's the next step? (optional)

After logging, show the updated project status.
