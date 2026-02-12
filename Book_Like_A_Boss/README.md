# Book Like A Boss Integration

Webhook receiver, booking sync, and Kit email automation for BLAB booking data.

## Files

- **blab-webhook-script.js** - Google Apps Script: webhook receiver + Kit email integration (copy to script.google.com)
- **sync-bookings.js** - Node script to sync Sheet → CoreContext/Calendar/bookings.json
- **PROJECT_HANDOFF_KIT_SEQUENCE.md** - Original planning doc (historical reference only)

## Architecture

```
Book Like A Boss
       │
       ▼ (webhook POST)
Google Apps Script (doPost)
       │
       ▼ (appendRow)
Google Sheet "BLAB Bookings"
       │
       ├──▶ sync-bookings.js → bookings.json → Claude Code analysis
       │
       ▼ (daily 9am trigger: processKitSubscriptions)
Kit API
  - Creates subscriber with meeting-guest-pending tag
  - Sets kit_first_form + Kit_Last_Form = "BLAB Networking Call"
       │
       ▼ (Rule 5298158)
Kit Sequence: "Meeting Guest Newsletter Invite" (3 emails)
  - Email 1 (Day 0): "Following up from our call"
  - Email 2 (Day 3): "Bumping this up"
  - Email 3 (Day 6): "Closing the loop (+ a couple resources)"
       │
       ▼ (guest clicks opt-in link → in-email tagging)
Tag: newsletter-confirmed
       │
       ▼ (Rule 5298362)
  - Subscribe to Welcome Sequence
  - Unsubscribe from invite sequence
```

## Kit Components

| Component | ID | Purpose |
|-----------|------|---------|
| Tag | `meeting-guest-pending` (15850596) | Applied by Apps Script after meeting |
| Tag | `newsletter-confirmed` (15851476) | Added when guest clicks opt-in link |
| Sequence | "Meeting Guest Newsletter Invite" (2648906) | 3-email opt-in invite |
| Rule | 5298158 | `meeting-guest-pending` → subscribe to sequence |
| Rule | 5298362 | `newsletter-confirmed` → Welcome Sequence + unsubscribe from invite |

## Key Functions in blab-webhook-script.js

| Function | Purpose |
|----------|---------|
| `doPost(e)` | Handles BLAB webhook POST, logs to sheet |
| `doGet(e)` | Health check endpoint |
| `processKitSubscriptions()` | Daily trigger — finds past meetings, adds to Kit |
| `addToKit(email, firstName)` | Creates Kit subscriber with tag + source fields |
| `addTagToExistingSubscriber(email)` | Tags existing subscribers, updates Kit_Last_Form only |
| `testKitIntegration()` | Test Kit API connection |
| `dryRunKitProcessing()` | Preview what would be processed |

## Setup

### 1. Google Apps Script (one-time)

1. Go to https://script.google.com
2. Create new project
3. Paste contents of `blab-webhook-script.js`
4. Deploy > New deployment > Web app
5. Execute as: Me
6. Who has access: Anyone
7. Copy the deployment URL

### 2. Book Like A Boss (one-time)

1. Other Settings > Integrations > Manage Webhooks
2. Click "Add"
3. Name: "Calendar Sync"
4. Callback URL: [paste Google Apps Script URL]
5. Select events: New Booking, Canceled Booking, Rescheduled Booking

### 3. Kit API Key (one-time)

1. Get API key from Kit → Settings → Developer → API Keys
2. In Apps Script: Project Settings → Script Properties
3. Add property: `KIT_API_KEY` = your_api_key

### 4. Daily Trigger (one-time)

1. In Apps Script: Triggers → Add Trigger
2. Function: `processKitSubscriptions`
3. Event source: Time-driven
4. Type: Day timer
5. Time: 9am to 10am

### 5. Google Sheet

Sheet: https://docs.google.com/spreadsheets/d/16NrUr0Xkz5TK-437P-2JH2bnywVE7cKjY8wvlwJRSco

- Column K (`kit_processed`): marked "yes" after Apps Script processes a row
- Only "Online Meeting" bookings (appointment ID 35230) are processed for Kit
- Canceled bookings and sales calls are skipped

### 6. Running Booking Sync (separate from Kit)

```bash
node /Users/dancumberland/Documents/Work/AI_Tools/Book_Like_A_Boss/sync-bookings.js
```

## Source Tracking

The Apps Script sets Kit custom fields for attribution:

- **New subscribers**: `kit_first_form` = "BLAB Networking Call", `Kit_Last_Form` = "BLAB Networking Call"
- **Existing subscribers**: Only `Kit_Last_Form` updated (preserves original `kit_first_form`)

## Related Documentation

- Kit setup details: `Dan_Content/_Tools/lead-source-analysis/KIT_SETUP_DOCUMENTATION.md` (Section 11)
- Email content: `Dan_Content/Email_Promotions/260211_Meeting_Guest_Newsletter_Invite_Sequence.md`
