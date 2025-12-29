---
title: Invoice Email Automation Guide
date: 2025-11-25
purpose: Document the Google Apps Script that watches invoice emails and saves PDFs/attachments to Drive
related_documents:
  - /AI_Tools/Invoice_Email_Automation/invoice_automation.gs
status: Active
key_themes:
  - automation
  - documentation
  - email
  - google_drive
  - apps_script
---
# GOOGLE APPS SCRIPT INVOICE AUTOMATION GUIDE

## What This Script Does

- ✅ Watches the **`invoices`** Gmail label inside `dc@dancumberlandlabs.com`
- ✅ Parses the first non-empty body line for `BUSINESSCODE SenderName` (fallback to subject)
- ✅ Builds filenames like `YYMMDD - BUSINESSCODE - SenderName`
- ✅ Saves **one copy of each distinct attachment** to Drive (deduplicates identical files)
- ✅ When no attachments exist, renders the email body to PDF
- ✅ Labels every processed conversation with **`invoices-processed`** so it never runs twice

**Architecture:** Gmail → Apps Script → Google Drive

**No external services required.** Everything runs inside Google.

---

## Prerequisites

1. **Gmail labels:**
   - `Invoices/New-Invoices` – applied to every invoice email you forward
   - `Invoices/Processed-Invoices` – applied by the script after each run

2. **Drive folder:**
   - `Receipts & Invoices` owned by `Dan.Cumberland@gmail.com`, shared with edit access to `dc@dancumberlandlabs.com`
   - Folder ID: `1-8JhkitHj9Y2iLk-yaabwqtiuIQO21P5` (already in script)

3. **Apps Script project:**
   - Created while logged in as `dc@dancumberlandlabs.com`

**Forwarding & tagging convention:**

When you forward invoices to your invoice inbox (for example `dc+invoices@dancumberlandlabs.com`), type a tagging line as the **first non-empty line of the email body**:

- `BUSINESSCODE SenderName`

Examples:

- `DCL Zoom`
- `TMM ConvertKit`
- `TF Landlord`
- `AIRBNB Cleaner`
- `HSA Blue Cross`

*Note:* `BUSINESSCODE` must be one of `AIRBNB`, `TF`, `TMM`, `DCL`, or `HSA` **in ALL CAPS**. If no valid code is found, the script defaults to `DCL` and extracts the sender name from the forwarded email's `From:` header (e.g., "Cloudflare" from `From: Cloudflare <noreply@cloudflare.com>`).

The script reads this first body line, combines it with the email date, and produces filenames like:

- `251115 - DCL - Zoom.pdf`

**Time to complete:** 10–15 minutes.

---

# STEP 1: PASTE THE SCRIPT

1. Go to https://script.google.com while logged in as `dc@dancumberlandlabs.com`.
2. Click **New project**.
3. Name it: `Invoice Email Automation`.
4. Replace the default `Code.gs` with this complete script:

```javascript
const CONFIG = {
  SOURCE_LABEL: 'Invoices/New-Invoices',
  PROCESSED_LABEL: 'Invoices/Processed-Invoices',
  FOLDER_ID: '1-8JhkitHj9Y2iLk-yaabwqtiuIQO21P5',
};

function processInvoices() {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const processedLabel = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  const query = 'label:' + CONFIG.SOURCE_LABEL + ' -label:' + CONFIG.PROCESSED_LABEL;
  const threads = GmailApp.search(query, 0, 50);

  threads.forEach(thread => {
    handleThread_(thread, folder);
    thread.addLabel(processedLabel);
  });
}

function handleThread_(thread, folder) {
  const messages = thread.getMessages();
  const message = messages[messages.length - 1];
  const subject = message.getSubject() || '';
  const bodyPlain = message.getPlainBody() || '';
  const messageDate = message.getDate();
  const { businessCode, senderName } = extractMeta_(bodyPlain, subject);
  const transactionDate = resolveTransactionDate_(bodyPlain, messageDate);
  const yymmdd = formatDateYYMMDD_(transactionDate);
  const filenameBase = sanitize_(`${yymmdd} - ${businessCode} - ${senderName}` || 'invoice');

  const attachments = uniqueAttachments_(message.getAttachments({
    includeInlineImages: false,
    includeAttachments: true,
  }));

  const inlineImages = message.getAttachments({
    includeInlineImages: true,
    includeAttachments: false,
  });

  if (attachments.length > 0) {
    attachments.forEach(att => {
      const ext = getExtension_(att.getName());
      const fileName = filenameBase + (ext || '');
      folder.createFile(att.copyBlob().setName(fileName));
    });
    return;
  }

  const html = embedInlineImages_(message.getBody() || wrapPlainAsHtml_(bodyPlain), inlineImages);
  const htmlBlob = Utilities.newBlob(html, 'text/html', filenameBase + '.html');
  const pdfBlob = htmlBlob.getAs('application/pdf').setName(filenameBase + '.pdf');
  folder.createFile(pdfBlob);
}

function uniqueAttachments_(attachments) {
  const seen = new Set();
  return (attachments || []).filter(att => {
    const hash = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, att.copyBlob().getBytes())
    );
    if (seen.has(hash)) return false;
    seen.add(hash);
    return true;
  });
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function sanitize_(value) {
  return (value || 'invoice').replace(/[\\/:*?"<>|]+/g, '_').trim();
}

function formatDateYYMMDD_(date) {
  if (!date) return '000000';
  const yy = String(date.getFullYear()).slice(-2);
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return yy + mm + dd;
}

function getExtension_(filename) {
  if (!filename) return '';
  const match = filename.match(/\.([^.]+)$/);
  return match ? '.' + match[1] : '';
}

const KNOWN_CODES = ['AIRBNB', 'TF', 'TMM', 'DCL', 'HSA'];

function parseTagLine_(raw) {
  const cleaned = (raw || '').trim();
  if (!cleaned) return null;
  const parts = cleaned.split(/\s+/);
  if (parts.length < 2) return null;

  const code = (parts[0] || '').toUpperCase();
  if (KNOWN_CODES.indexOf(code) === -1) return null;

  const sender = parts.slice(1).join(' ') || 'Unknown';
  return { businessCode: code, senderName: sender };
}

function extractMeta_(bodyPlain, subject) {
  const lines = (bodyPlain || '').split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const fromBody = parseTagLine_(trimmed);
    if (fromBody) return fromBody;
    break;
  }

  const fromSubject = parseTagLine_(subject);
  if (fromSubject) return fromSubject;

  // Fallback: DCL code with sender name extracted from forwarded email's From header
  const forwardedSender = extractForwardedSenderName_(bodyPlain) || 'Unknown';
  return { businessCode: 'DCL', senderName: forwardedSender };
}

// Extract the domain name from the forwarded email's From: header.
// "From: Cloudflare <noreply@notify.cloudflare.com>" -> "Cloudflare"
function extractForwardedSenderName_(bodyPlain = '') {
  const fromRegex = /^From:\s*(?:(.+?)\s*<[^>]+>|([^<\s]+@([^>\s]+)))\s*$/im;
  const match = bodyPlain.match(fromRegex);
  if (!match) return null;

  if (match[1] && match[1].trim()) {
    return match[1].trim();
  }

  if (match[3]) {
    const domain = match[3].split('.')[0];
    return domain.charAt(0).toUpperCase() + domain.slice(1).toLowerCase();
  }

  return null;
}

function resolveTransactionDate_(bodyPlain, fallbackDate) {
  const inlineDate = extractInlineDate_(bodyPlain);
  if (inlineDate) return inlineDate;

  const quotedDate = extractQuotedForwardDate_(bodyPlain);
  if (quotedDate) return quotedDate;

  return fallbackDate || new Date();
}

function extractInlineDate_(bodyPlain = '') {
  const dateRegex = /\b(20\d{2}-\d{2}-\d{2})\b|\b(\d{6})\b/;
  const match = bodyPlain.match(dateRegex);
  if (!match) return null;

  if (match[1]) {
    const parsed = new Date(match[1]);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  if (match[2]) {
    const yy = match[2].slice(0, 2);
    const mm = match[2].slice(2, 4);
    const dd = match[2].slice(4, 6);
    const iso = `20${yy}-${mm}-${dd}`;
    const parsed = new Date(iso);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  return null;
}

function extractQuotedForwardDate_(bodyPlain = '') {
  const forwardRegex = /On\s+([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4}).+?wrote:/s;
  const match = bodyPlain.match(forwardRegex);
  if (!match) return null;

  const parsed = new Date(match[1]);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function embedInlineImages_(html, inlineAttachments = []) {
  if (!html || !inlineAttachments || inlineAttachments.length === 0) return html;

  const cidMap = inlineAttachments.reduce((map, att) => {
    const contentId = (att.getContentId && att.getContentId()) || '';
    const cid = contentId.replace(/[<>]/g, '').trim();
    if (!cid) return map;

    const blob = att.copyBlob();
    const mimeType = blob.getContentType() || 'application/octet-stream';
    const base64 = Utilities.base64Encode(blob.getBytes());
    map[cid.toLowerCase()] = `data:${mimeType};base64,${base64}`;
    return map;
  }, {});

  if (Object.keys(cidMap).length === 0) return html;

  return html.replace(/src=(['"])cid:([^'"]+)\1/gi, (match, quote, cid) => {
    const normalized = cid.replace(/[<>]/g, '').trim().toLowerCase();
    const dataUri = cidMap[normalized];
    return dataUri ? `src=${quote}${dataUri}${quote}` : match;
  });
}

function wrapPlainAsHtml_(plain) {
  const safe = (plain || 'No body content available')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
  return "<pre style=\"font-family: 'Inter', sans-serif; white-space: pre-wrap;\">" +
         safe +
         '</pre>';
}

// This version:
// - Only processes conversations that have `label:invoices` and **do not** have `label:invoices-processed`.
// - Adds `invoices-processed` to every handled thread, so nothing runs twice unless you remove that label manually.
// - Deduplicates attachments by content hash (`uniqueAttachments_`) so duplicate PDFs don’t clutter Drive.

## STEP 2: AUTHORIZE AND RUN ONCE

1. Click **Run → processInvoices**.
2. Approve the Gmail + Drive permissions when prompted.
3. Check the execution log; you should see each invoice thread handled and labeled `invoices-processed`.

## STEP 3: ADD A TIME-DRIVEN TRIGGER

1. In Apps Script, click the **clock** icon in the left sidebar (Triggers).
2. Click **+ Add Trigger**.
3. Configure:
   - **Function:** `processInvoices`
   - **Deployment:** `Head`
   - **Event source:** Time-driven
   - **Type:** Minutes timer or Hour timer
   - **Interval:** Every 5 minutes (testing) or Every hour (production)
4. Click **Save** and authorize if prompted.

You can still click **Run** manually at any time; the trigger just keeps things on schedule.

---

# STEP 4: TEST THE AUTOMATION

1. Forward a test email to `dc+invoices@dancumberlandlabs.com`.
2. **First line of body:** `DCL Test Vendor` (ALL CAPS code).
3. Attach a sample PDF; optionally include a duplicate second PDF to confirm deduping.
4. Label the email `invoices` in Gmail.
5. Run `processInvoices` manually or wait for the trigger.
6. **Verify:**
   - Drive contains a single attachment (or a PDF of the body) named `YYMMDD - DCL - Test Vendor`.
   - Gmail conversation now shows label `invoices-processed`.

---

# TROUBLESHOOTING

| Symptom | Likely Cause | Fix |
| --- | --- | --- |
| No emails processed | Label mismatch or trigger disabled | Verify the Gmail label is exactly `invoices` and the trigger is active. |
| Duplicate files saved | Deduping removed | Ensure `uniqueAttachments_` is still wrapping `message.getAttachments`. |
| Files owned by workspace account | Script running as `dc@dancumberlandlabs.com` | That's expected; share the folder from your personal account if you need personal ownership. |
| Needs to re-run on an old message | Already has `invoices-processed` label | Remove the label from that conversation and rerun. |

---

# MAINTENANCE

- **Monthly:** Review the Drive folder to confirm naming consistency.
- **As needed:** Update `KNOWN_CODES` array whenever you add a new business code.
- **Volume changes:** Adjust the trigger interval if invoice volume spikes or slows down.
- **Permissions:** Re-authorize the script if Google flags the permissions (visible as trigger failures).

---

# ATTACHMENT DEDUPLICATION

Some vendors email two PDF attachments (one with payment history, one without) but they're byte-for-byte identical. The `uniqueAttachments_` function:

- Computes an MD5 digest of each attachment blob.
- Only keeps the first instance of each hash.
- **Result:** Only one copy of a duplicate attachment is stored, reducing Drive bloat.

If you ever want to keep all versions, remove the `uniqueAttachments_` wrapper and map directly over `message.getAttachments()`.

---

This file reflects the **Google Apps Script–only** implementation. No Make.com components are required.
