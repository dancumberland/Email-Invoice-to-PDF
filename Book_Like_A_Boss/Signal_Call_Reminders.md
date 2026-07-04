# Signal Call — reminder + no-show system (copy + cadence)

**Status:** ✅ BUILT 2026-06-18 (live, in BLAB admin via Chrome). Optional phone field, 3-reminder schedule, custom-voice Confirmation + Reminder email bodies, and the post-booking Success message are ALL live on the Signal Call appointment. Only the after-a-miss recovery is unbuilt (it's not a BLAB feature — see constraints).

**🟡 PASTE PENDING (2026-07-04):** the on-book Confirmation email (§1) and the Success Message re-sell (§"Confirmation page re-sell") copy were updated to carry the locked seam sentence — *"You'll leave with your first move and the start of a 90-day plan built on it. The plan is yours to keep either way."* — so the deliverable page → booking page → confirmation email all read as one promise. The booking page (dancumberlandlabs.com/signal-call) already shipped. **These two BLAB surfaces are a manual paste in BLAB admin (Customization > Email Messages for the confirmation; General Settings > Success Message) — Dan-side, behind the BLAB login.** Paste the updated §1 body + add the seam line to the Success Message, then flip this back to ✅.
**Goal:** lift show-rate from the ~50–60% of a bare scheduler toward 75–90% with confirmations, reminders, and a recovery path.
**Voice:** peer-to-peer, plain, no hype. SMS kept under ~160 chars.

---

## What's live / what remains (built 2026-06-18, package 295887)

**✅ LIVE:**
- **Optional phone field** — custom field "Mobile phone (optional — for text reminders)", type Telephone Number (BLAB value `phone`), **not required**. SMS goes only to clients who provide a number; everyone still gets the emails. (Dan's decision: "Email + SMS, optional phone.")
- **3 reminders** on the Signal Call appointment, all sent to the client:
  - **Email — 2 days before** (the re-sell / "look ahead").
  - **SMS — 1 day before** (confirm intent).
  - **SMS — 3 hours before** (day-of nudge).
  - "Require phone number for SMS" master toggle is OFF (keeps phone optional).

**✅ ALSO LIVE (custom copy in Dan's voice, set 2026-06-18):**
- **Confirmation email** (on-book) → Customization > Email Messages, "Global" unchecked, custom body set (copy #1, links omitted — BLAB's shell appends the appointment card + reschedule/cancel buttons). Persisted (verified after reload).
- **Reminder email** body (shared by the email reminder) → same tab, "Global" unchecked, custom body set (copy #3). Persisted.
- **Success Message** (post-booking screen re-sell) → General Settings, set + saved (copy adapted from §"Confirmation page re-sell" below).
- Mechanism confirmed: the Email Messages override is the *message text*; BLAB still wraps it with date/time/Google-Meet/reschedule buttons, so no merge tokens were needed in the body.

**🟡 ONLY REMAINING: after-a-miss recovery** — not a BLAB feature (no no-show event). Build as a Kit automation (time-based after the slot, `signal-call-no-show` tag) or handle manually. Copy #6/#7 below.

**⚠️ BLAB MODEL CONSTRAINTS (discovered on build — the drafted cadence doesn't map 1:1):**
- **SMS body is NOT per-service customizable.** BLAB sends a default SMS reminder format (service name + date/time + reschedule link); the per-step SMS copy (#2/#4/#5 — "reply YES", "join link") can't be set per service. Distinct per-SMS wording would require a global SMS template that hits ALL appointments. Left on BLAB's default.
- **One shared Reminder email template** — all email reminders use the same body (fine here: only one email reminder).
- **After-a-miss recovery is NOT a BLAB feature** — there's no no-show event, and a blanket post-appointment "Followup" would email people who *did* show. Do recovery in Kit (time-based) or manually. Copy #6/#7 below are for that path, not BLAB.

BLAB merge fields below are written as `{{field}}` placeholders — map to BLAB's actual tokens when configuring (`{{first_name}}`, `{{date}}`, `{{time}}`, `{{timezone}}`, `{{reschedule_link}}`, `{{join_link}}`).

---

## Phase 0 — confirm in BLAB admin first ✅ ANSWERED 2026-06-12 (Claude agent, live admin walkthrough)

These four answers decide what's buildable. Ship the link-based booking regardless; these unlock the rest.

1. **SMS reminders** — ✅ Available. Configured **per-service** under each appointment's Reminders tab. Up to **5 steps** (email or SMS) per sequence. SMS requires the client to provide a phone number at booking. Plan includes **100 SMS/month**.
2. **Reminder timing** — ✅ Fully flexible custom offsets: independent dropdowns for days (0–30), hours (0–23), minutes (0–55 in 5s). All target offsets (on-book, 48–72h, 24h, 2–3h) are achievable.
3. **Confirmation page / email** — ✅ Customizable at two points: (1) on-page "Success Message" in General Settings (rich text, shown right after booking); (2) per-service Confirmation email under Customization > Email Messages (rich text, with a "Global" override checkbox).
4. **No-show / cancellation webhook** — ⚠️ Partial. Webhooks exist globally (Settings > Integrations > Webhooks) AND per-service, with three triggers: **Created, Canceled, Rescheduled**. There is **no no-show event** (BLAB has no no-show concept) → no-show recovery is manual or time-based, as anticipated.

**Build implications:** the 4-step cadence below fits inside the 5-step per-service limit (on-book confirmation email is separate). Post-miss recovery cannot be webhook-triggered — run it time-based from Kit or manually after the call slot passes.

> The `signal` scheduler was created 2026-06-12 (live at book.dancumberland.com/signal — 30 min, free, Google Meet, 1-day cutoff/60-day window, two required intake questions). Admin: bookme.name appointment 295887. Slug renamed signal-call → signal 2026-06-12.

---

## Calendar tuning (Dan, BLAB admin)

- Booking window: 3–7 days out (cap ≤10–14). Closer dates = higher show-rate.
- Show 2–4 slots/day, not the whole calendar. Scarcity is real, not manufactured.
- Buffers between calls; timezone auto-detect on.
- Self-serve reschedule ON (a reschedule is a save, not a loss).

---

## The reminder cadence

| When | Channel | Purpose |
|---|---|---|
| On booking | Email + SMS | Confirm, set expectations, add to calendar |
| 48–72h before | Email | Re-sell the value; easy reschedule if needed |
| 24h before | SMS | Confirm intent ("reply YES") |
| 2–3h before | SMS | Day-of nudge + join link |
| After a miss | Email (+ SMS if available) | Recover — rebook, no guilt |

---

## Copy

### 1. On-booking confirmation — EMAIL

**Subject:** You're booked — Signal Call, {{date}} at {{time}}

{{first_name}},

You're on the calendar for {{date}} at {{time}} {{timezone}}. Here's what to expect.

We'll walk your Signal Scorer result together and pin down the one constraint capping the rest. You'll leave with your first move and the start of a 90-day plan built on it. The plan is yours to keep either way. Thirty minutes, no pitch.

You don't need to prepare anything — your result is already on file. If you want to get more out of it, come with one sentence on what you're hoping AI does for the firm in the next year.

**Add it to your calendar:** {{calendar_link}}
**Need a different time?** Reschedule here, no problem: {{reschedule_link}}

See you then,
Dan

---

### 2. On-booking confirmation — SMS

You're booked for the Signal Call {{date}} {{time}}. Walk your result + a plan, 30 min, no pitch. Reschedule anytime: {{reschedule_link}} — Dan

---

### 3. 48–72h before — EMAIL

**Subject:** Quick look ahead at your Signal Call

{{first_name}},

Your Signal Call is coming up {{date}} at {{time}} {{timezone}}.

A reminder of what we'll do: walk your result, find the one constraint capping the rest, and map the first move against it. You'll leave with the plan either way.

If something's come up, reschedule in two clicks — I'd rather grab a time that works than have you rush it: {{reschedule_link}}

Talk soon,
Dan

---

### 4. 24h before — SMS

Signal Call tomorrow, {{date}} {{time}}. Reply YES to confirm, or reschedule: {{reschedule_link}} — Dan

---

### 5. 2–3h before — SMS

Signal Call in a couple hours ({{time}}). Join here when it's time: {{join_link}} — see you soon, Dan

---

### 6. After a miss — EMAIL

**Subject:** Missed you — let's grab another time

{{first_name}},

Looks like we missed each other today. No worries, it happens.

The offer stands: thirty minutes, I walk your Signal result with you and leave you a plan. Grab whatever time works: {{reschedule_link}}

If the timing's just not right, that's fine too — you'll keep getting the regular notes.

Dan

---

### 7. After a miss — SMS (if available)

Missed you for the Signal Call today — no worries. Rebook whenever: {{reschedule_link}} — Dan

---

## Confirmation page re-sell (whatever BLAB allows on the post-booking screen)

- One line: "You're booked. Here's what happens on the call."
- The 3-step: walk your result → pinpoint the constraint → map the first move + leave with a plan.
- The seam sentence (same wording as the deliverable page + booking page, so click→call reads as one promise): "You'll leave with your first move and the start of a 90-day plan built on it. The plan is yours to keep either way."
- Add-to-calendar button.
- One verified testimonial (Amanda: 20% margin, 10+ hrs/week per person).
- Soft no-show line: "Can't make it? Reschedule — link's in your email."

---

## Kit wiring (Claude, if BLAB emits a no-show/cancel webhook)

- No-show → apply a `signal-call-no-show` tag → short recovery automation (the miss email above, then back to the regular newsletter).
- Reuse existing: `signal-call-qualified` (20168203), `clicked-booking-link` (17307991), the `signal_scorer_*` fields.
- If BLAB has no webhook: drive recovery off the existing BLAB→Sheet→Kit Apps Script (`AI_Tools/Book_Like_A_Boss/apps-script/`) or handle no-shows manually for now.
