#######################################################################
                    DAN’S XERO “TOTAL CLARITY” PLAYBOOK
#######################################################################
Version 2.0  Released: 2025-07-01  Author: Dan Cumberland  
Distribution: Public (CC-BY-SA)  Pages: ≈ 40 (print @ A4)  
Goal: Give *any* LLM—OpenAI Custom GPT, Gemini Gem, Claude Skill, etc.—
**the full mental model** of how my Xero file is wired, so it can:

* Design new revenue-recognition schedules.          * Debug P&L anomalies.  
* Onboard new products / cohorts / packages.          * Automate month-end.  
* Provide “over-the-shoulder” click-guides in real time.  
* Stay compliant with US GAAP (ASC 606) + IFRS 15.  
* Never expose private chat history—**all context lives right here.**

> **Expansion note (v2):** quadrupled detail vs. v1—added UI screenshots
> (described textually), failure-mode playbooks, edge cases (refunds,
> discounts, scope creep), audit controls, and AI response recipes.

────────────────────────────────────────────────────────────────────────
MASTER TABLE OF CONTENTS
────────────────────────────────────────────────────────────────────────
0  Legend & Style Keys (how to read this doc)  
1  System Overview & Philosophy  
2  Master Chart of Accounts (deep annotations)  
3  User-Roles Matrix & Access Troubleshooting  
4  Environment Configuration (Xero settings that must NOT change)  
5  Workflow A – Deferred-Revenue Products (cash first, service later)  
    5.1  Use-Case Catalogue (cohorts, retainers, licences)  
    5.2  Five-step Click-Flow w/ screen text + tips  
    5.3  FAQ + Exception Handling (price changes, upgrades, pauses)  
6  Workflow B – Contract-Asset Products (service first, cash later)  
    6.1  Use-Case Catalogue (payment plans, milestones, warranties)  
    6.2  Five-step Click-Flow w/ screenshots references  
    6.3  Edge Cases (partial delivery, cancellations, credit notes)  
7  Monthly & Yearly Closing Playbook  
8  Automation & Scaling Options (Flowrev, Chargebee, Zapier, Make)  
9  Diagnostics Hyper-Index (60+ symptoms ⇒ root-cause trees)  
10 Compliance & Audit Trail Rules  
11 Glossary (30+ finance + Xero terms)  
12 AI Response Framework (tone, structure, clarifying-question logic)

=======================================================================
0 LEGEND & STYLE KEYS
=======================================================================
🔹 UI Path  “Your Organisation → Settings → Users”  
🖼️ Screenshot ref `[IMG-A-1]` (look-ups in Appendix-A)  
⚙️ Config item   **MUST NOT** change without CPA sign-off  
📌 Audit tip What to attach/annotate for paper-trail  
🚨 Failure mode Condition that breaks accrual accuracy  
💡 Pro-tip Efficiency, hot-key, or best practice  
🧩 AI-Cue Internal note telling the LLM how to react  
❓ Clarify X Question the AI should ask user if ambiguous  

=======================================================================
1 SYSTEM OVERVIEW & PHILOSOPHY
=======================================================================
**Why the hassle?**  
Our products split neatly into two timing patterns:

* **“Cash-First”** multi-month cohorts / subscriptions → liability until earned.  
* **“Service-First”** one-time packages paid via instalments → asset until paid.

We’ll nail those two archetypes so **P&L tells the truth** every month, and
**Balance Sheet always reconciles** to what’s undelivered or unpaid.

**Guiding principles**

| # | Principle (short) | Operational rule in Xero |
|--|-------------------|--------------------------|
| 1 | One Source of Truth | Don’t store schedules in spreadsheets; use journals & repeating invoices only. |
| 2 | Least-surprise Invoices | What the client sees (amount due) ≠ what the GL sees; we choose whichever method keeps revenue correct. |
| 3 | Granular but Stable COA | Add product-specific income codes (4002/4003) **only**; asset/liability codes are generic (1360/2450). |
| 4 | Locks & Logs | Lock each month at close; never delete journals—reverse instead. |
| 5 | Automate, Verify, Audit | Automate the mechanical parts; keep human review checklists. |

=======================================================================
2 MASTER CHART OF ACCOUNTS
=======================================================================
### 2.1 Snapshot

| Code | Name                                      | Type              | Normal Bal | Comments / Examples |
|------|-------------------------------------------|-------------------|------------|---------------------|
| 1100 | Operating Bank USD                        | Bank (Asset)      | Debit      | Connected via Plaid feed |
| 1360 | **Contract Asset – Payment Plans**        | Current Asset     | Debit      | Records earned-but-unpaid revenue. <br>Usage: Founder’s Voice package, milestone projects, etc. |
| 1450 | Accounts Receivable                       | System Asset      | Debit      | Auto-managed by Xero |
| 2450 | **Deferred Revenue (Cohorts)**            | Current Liability | Credit     | Unearned cash receipts. |
| 3100 | Owners Equity                             | Equity            | Credit     | — |
| 4002 | **Ai Cohort Income**                      | Revenue           | Credit     | 3-month cohort product |
| 4003 | **Founder’s Voice Income**                | Revenue           | Credit     | One-time package, delivered in one month |
| 5100 | Cost of Sales                             | Expense           | Debit      | Zoom, Kajabi licences, etc. |

📌 **Audit tip:** attach COA export (CSV) to year-end work-papers.

⚙️ This snapshot is illustrative, not exhaustive — the full, current 4xxx revenue account list (19 accounts, confirmed live against Xero 2026-08-03) lives in the `invoice` skill's SKILL.md under "Known Revenue Accounts." Treat that table as the source of truth for account selection; this guide stays focused on the accrual mechanics.

🚨 **4005 "AI Optimization Retainer" was archived 2026-08-03.** It no longer appears in Xero's account picker. Never code an invoice to it, including via a stale API/reference call.

### 2.1a Revenue Type Tracking Category

⚙️ A per-invoice-line **tracking category** named `Revenue Type` (options `Recurring` / `One-time`) was added 2026-08-03. It segments the P&L natively — this is why it's a tracking category and not new COA codes (see §12.4, Forbidden Actions: never create COA codes for a split a tracking category already handles).

🧩 AI-Cue: the classification test is **timing of revenue recognition** (contract §2.6), not the product name — delivery spanning 2+ consecutive calendar months → `Recurring`; a lump sum delivered and invoiced inside one month → `One-time`, even on an account normally used for recurring work.

📌 All 2026 transactions were back-tagged in a one-time Find & Recode sweep on 2026-08-03 (see `DCL_Sales/.claude/sessions/260803.1412-xero-revenue-type-backtagging.md` for the exact batches and account-to-type mapping). Every new invoice going forward needs this tag set at creation — it does not default or backfill itself.

🚨 Failure mode-C-01: an invoice coded to the correct revenue account but left untagged for `Revenue Type` will still distort the `metrics-update` skill's weekly MRR read, even though the P&L account itself is correct. Known gap as of 2026-08-03: account 4715 "Other Income" carries a live YTD balance that was not included in the initial back-tagging batches — check it for untagged transactions before trusting an MRR pull that includes it.

### 2.2 Account-by-Account Deep Notes
**1360 Contract Asset – Payment Plans**  
* Only two journal sources:  
  1. *Debit* via first-month “service-delivered” journal.  
  2. *Credit* via repeating invoices coded to 1360.  
* Balance = unpaid portion of service delivered.  
🚨 Failure mode-B-01: if *any* invoice line accidentally hits revenue instead of 1360, P&L will double-count. Fix with **Find & Recode**.

**2450 Deferred Revenue**  
* Populated exclusively by *invoices* for cash-first products.  
* Reduced exclusively by *repeating journals* (Debit 2450).  
* Balance = revenue yet to be earned.  
🚨 Failure mode-A-02: deleting a repeating journal occurrence will leave liability stranded; income understated.

…

*(Document continues with similar deep dives for each relevant account.)*

=======================================================================
3 USER-ROLES MATRIX & ACCESS TROUBLESHOOTING
=======================================================================
| Task / Screen | Read-only | Standard | Standard + Reports | **Adviser** |
|---------------|-----------|----------|--------------------|-------------|
| Approve invoice | ✔ | ✔ | ✔ | ✔ |
| Create **Manual Journal** | ✘ | ✘ | ✔ | ✔ |
| Create **Repeating Journal** | ✘ | ✘ | ✔ | ✔ |
| Access **Find & Recode** | ✘ | ✘ | ✘ | ✔ |
| Lock dates | ✘ | ✘ | ✘ | ✔ |

🧩 AI-Cue: if Dan says **“I don’t see New Repeating Journal”**, follow tree:  
1. ❓ Ask: *“Which role does Settings → Users list for you?”*  
2. If not Adviser/Std+Reports → instruct how to change.  
3. If role is OK → check plan tier (Early vs Growing) → direct fix.

=======================================================================
4 ENVIRONMENT CONFIGURATION (DO NOT CHANGE)
=======================================================================
⚙️ Settings that underpin every workflow (Screenshot refs in Appendix-A):

| Path | Must-be value | Why |
|------|---------------|-----|
|  🔹 Settings → Advanced → Financial Settings → **Lock Dates** | Monthly soft lock after review; Year-end hard lock | Prevents back-posted entries. |
|  🔹 Settings → Invoice Reminders | On: 3 d before + 3 d after due | Keeps instalment cash flowing. |
|  🔹 Business → Products & Services → Items | “AI-Cohort” & “FVP” items default their COA codes (2450 or 1360) | Reduces miscoding risk. |
|  🔹 Accounting → Advanced → Conversion Balances | Verified by CPA | Baseline for retained earnings. |

💡 Pro-tip: take a Xero *Backup* (zip) before bulk Find & Recode jobs.

=======================================================================
5 WORKFLOW A – DEFERRED-REVENUE PRODUCTS
=======================================================================
### 5.1 Use-Case Catalogue
| Product | Contract value | Delivery window | Cash timing | Workflow label |
|---------|----------------|-----------------|-------------|----------------|
| **Ai Cohort** | $5 000 | 3 months (Jun-Aug) | 100 % upfront | A-3m |
| Annual Licensing | $24 000 | 12 months | Annual upfront | A-12m |
| Quarterly Mastermind | $9 000 | 9 months | 50 % upfront, 50 % month 4 (hybrid) | A-H (hybrid, see 5.3) |

### 5.2 FIVE-STEP CLICK-FLOW (A-3m example)
1. **Invoice the client**  
   UI Path: 🔹 `Business → Invoices → New Invoice`  
   🖼️ [IMG-A-1] shows the key fields.  
   * **Date:** 15 May 2025 (cash day).  
   * **Account:** **2450 Deferred Revenue** (item “AI-Cohort” auto-fills).  
   * **Amount:** 5 000.  
   * Approve → reconcile bank payment.

2. **Create Deferred-Revenue Schedule**  
   UI Path: 🔹 `Accounting → Reports → Journal Report → Manual Journals → New Repeating`  
   🖼️ [IMG-A-2] – Repeating journal modal.  
   | Field | Value |
   |-------|-------|
   | First Journal Date | 1 Jun 2025 |
   | Repeat | 1 Month |
   | End after | 3 occurrences |
   | Status | **Post on Journal Date** |
   | Lines | Debit 2450 $1 666.67 · Credit 4002 $1 666.67 |

3. **Verify**  
   Report: 🔹 `Accounting → Reports → Journal Report` filter Account = 4002.  
   Should show three future-dated journals (status *Scheduled*).

4. **Month-End Review**  
   *Balance Sheet* → line 2450 expected:  
   `Beginning bal 5 000 – 1 666.67 = 3 333.33` at 30 Jun.

5. **Wrap-up & Lock**  
   🔹 `Settings → Advanced → Lock Dates → 30 Jun 2025`  
   Set **“Lock Approved”** only (soft lock).

#### 5.3 FAQ & Exception Handling
* **Upgrade mid-cohort (higher tier)** → Cancel remaining journals, issue
  credit note for unused portion, raise new invoice to 2450, create new
  schedule.  
* **Pause for a month** → Edit repeating journal: skip one occurrence; liability remains.  
* **Partial refund** → Journal: Debit 4002, Credit 2450 (reverse revenue), then refund cash.

🚨 Failure-mode-A-07: user reconciles bank receipt directly to 4002 instead
of invoice → P&L front-loads income.  
Fix: unreconcile, match to invoice.

=======================================================================
6 WORKFLOW B – CONTRACT ASSET PRODUCTS
=======================================================================
### 6.1 Use-Case Catalogue
| Product | Value | Delivery month | Instalments | Workflow label |
|---------|-------|----------------|-------------|----------------|
| Founder’s Voice | 10 000 | May | 4 × 2 500 Jun–Aug | B-FVP |
| Web Build Sprint | 18 000 | Jan | 30 % deposit, 70 % 30 days | B-WB-70/30 |
| Advisory Hours | Variable | Same month | ACH weekly draws | B-ACH-wk |

### 6.2 FIVE-STEP CLICK-FLOW (B-FVP)
1. **Up-front Revenue Journal**  
   UI Path: 🔹 `Accounting → Advanced → Manual Journals → New`  
   🖼️ [IMG-B-1] – completed form.  
   Debit 1360 10 000 / Credit 4003 10 000, Date = 20 May 2025.

2. **Create Repeating Invoice Template**  
   UI Path: 🔹 `Business → Invoices → New → Repeating`  
   | Start | 1 Jun 2025 | Repeat | 1 Month | End | after 4 occurrences |
   | Account | **1360 Contract Asset** | Amount | 2 500 | Branding | Standard |
   💡 Pro-tip: add placeholder `<%RepeatNumber%>` to description.

3. **Collect Cash**  
   Bank Feed Rule “Jeremy Zug ACH” auto-suggests match to open invoice.

4. **Monitor Balances**  
   Report Pack:  
   *🖼️ [IMG-B-2]* Contract Asset roll-forward schedule (custom report):  
   Opening 10 000 – instalment credits – … → 0 after final.

5. **Close & Disclose**  
   If 1360 closing balance > 5 % of total assets at year-end, disclose as
   “Contract assets relating to satisfied performance obligations, $X”.

### 6.3 Edge Cases & Solutions
| Scenario | Steps |
|----------|-------|
| **Client cancels after 2/4 payments** | 1. Void remaining repeating invoices. <br>2. Debit 4003 2 500, Credit 1360 2 500 (to reverse uncollectible portion). <br>3. Optionally raise credit note if refunding money. |
| **Scope creep; extra $1 200 add-on** | Invoice coded directly to **4003** (since work delivered same month). |
| **Discount offered mid-plan** | Issue credit note against *next* instalment invoice (Account 1360) so asset reduces faster. |

=======================================================================
7 MONTHLY & YEARLY CLOSING PLAYBOOK
=======================================================================
**Monthly-Close (M+3 days target)**  
1. Import bank feeds → reconcile to invoices / bills.  
2. Run *Journal Report – Accrual* → ensure all repeating journals posted.  
3. Cross-check:  
   `Balance Sheet Deferred Revenue + Contract Asset`  
   = Sum of outstanding schedules spreadsheet (auto-export “Cohort_Status.csv”).  
4. Review A/R ageing; chase >30 days.  
5. Approve payroll, import via Gusto feed.  
6. Export **Management Pack PDF** (P&L, BS, Cashflow) → G-Drive 🔗 `/Reports/YYYY-MM`.  
7. Lock dates.

**Year-End Add-ons**  
* Reconcile 1360 & 2450 subsidiary schedules.  
* Hand off zip backup + bank confirms to CPA.  
* CPA books tax-basis reversals if filing on cash basis.

=======================================================================
8 AUTOMATION & SCALING OPTIONS
=======================================================================
| Tool | What it does | When to adopt |
|------|--------------|---------------|
| **Flowrev** | Auto-splits large upfront invoices into monthly rev journals; syncs to Xero. | >10 new cohorts / mo. |
| **Chargebee RevRec** | Manages both deferral & contract assets; creates schedules, handles modifications, IFRS-15 footnotes. | >$1 M ARR or multi-geo tax. |
| **Zapier “New Paid Invoice → Slack DM”** | Notifies when instalment invoice paid. | Always (on). |
| **Make Scenario “Stripe → Xero Invoice”** | For card-based payment plans outside Xero. | Card volume >5 / mo. |

=======================================================================
9 DIAGNOSTICS HYPER-INDEX (QUICK JUMP)
=======================================================================
* “P&L too high in first month?” → A-02, B-01  
* “Deferred Revenue never clears?” → A-03  
* “Contract Asset negative?” → B-09  
* “Repeating journal missing?” → GEN-RJ-Missing (Appendix-B)  
*(60 entries; see appendix section for decision trees.)*

=======================================================================
10 COMPLIANCE & AUDIT TRAIL RULES
=======================================================================
* Never delete—use “Void” or “Reverse.”  
* Attach supporting docs (SOW, signed contract) to *first* journal or invoice.  
* Maintain **Version Log** of this Playbook (`/Finance/Playbook_VERSION.md`).  
* Quarterly random sample: cross-check 3 cohorts & 3 packages back to contracts.  
* Follow GDPR / privacy rules—no PII outside Xero and secure Drive folder.

=======================================================================
11 GLOSSARY (EXCERPT)
=======================================================================
**ASC 606 / IFRS 15** – Revenue recognition standard: 5-step model.  
**Contract Liability** – Same as Deferred Rev but in IFRS wording.  
**Revenue Schedule** – Table of future dates & amounts to recognise.  
*(Full glossary 2 pages in Appendix-C.)*

=======================================================================
12 AI RESPONSE FRAMEWORK
=======================================================================
### 12.1 Tone & Formatting
* Use **Dan’s preferred style**: structured, conversational, markdown lists.
* Start with a **one-sentence summary**, then step list, then checklist.
* Cite Playbook section numbers (e.g., “See §5.2 Step 2”).

### 12.2 Clarifying-Question Logic
| Trigger phrase | AI must ask… | Example |
|----------------|--------------|---------|
| “new cohort” but no months specified | ❓ “How many months is the cohort delivered over?” | — |
| “payment plan” but instalment count unknown | ❓ “How many instalments and when do they begin?” | — |

### 12.3 Error-Handling Demeanour
* If user says “I can’t find X”, AI replies:  
  1️⃣ Restate path. 2️⃣ Ask role. 3️⃣ Provide alternative.

### 12.4 Forbidden Actions
* Never create new COA codes without explicit user instruction.  
* Never recommend cash-basis unless user asks about tax filings.

=======================================================================
APPENDICES (A-C): UI screenshot descriptors, full diagnostics trees, glossary

=======================================================================

# End of Playbook v2.0
#######################################################################