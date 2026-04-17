/**
 * Book Like A Boss Webhook Receiver
 *
 * Google Apps Script that receives BLAB webhooks and logs to Google Sheet.
 *
 * SETUP:
 * 1. Go to https://script.google.com and create a new project
 * 2. Paste this entire script
 * 3. Click Deploy > New deployment
 * 4. Choose "Web app"
 * 5. Set "Execute as" to "Me"
 * 6. Set "Who has access" to "Anyone"
 * 7. Click Deploy and copy the URL
 * 8. In BLAB: Other Settings > Integrations > Manage Webhooks > Add
 * 9. Paste the URL as the Callback URL
 * 10. Select desired trigger events (e.g., "New Booking", "Canceled Booking")
 */

// Your Sheet ID - already configured
const SHEET_ID = '16NrUr0Xkz5TK-437P-2JH2bnywVE7cKjY8wvlwJRSco';
const SHEET_NAME = 'Sheet1';

// Slack webhook - Script Properties override, fallback to constant
const SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL')
  || 'REMOVED_SLACK_WEBHOOK_USE_SCRIPT_PROPERTIES';

/**
 * Handles POST requests from BLAB webhooks
 */
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

    // Parse the incoming data
    let data = {};
    let rawPayload = '';

    if (e.postData && e.postData.contents) {
      rawPayload = e.postData.contents;
      try {
        data = JSON.parse(rawPayload);
      } catch (parseError) {
        // If not JSON, might be form-encoded
        data = { raw: rawPayload };
      }
    }

    // Extract common booking fields (flexible - tries multiple possible field names)
    const clientName = extractField(data, ['client_name', 'name', 'customer_name', 'booker_name', 'full_name', 'firstName', 'first_name']) +
                       (extractField(data, ['lastName', 'last_name']) ? ' ' + extractField(data, ['lastName', 'last_name']) : '');

    const clientEmail = extractField(data, ['client_email', 'email', 'customer_email', 'booker_email']);

    const serviceType = extractField(data, ['service', 'service_type', 'event_type_name', 'appointment_type', 'booking_type', 'event_name', 'type']);

    const bookingDatetime = extractField(data, ['datetime', 'start_time', 'start', 'booking_time', 'scheduled_time', 'date', 'startTime', 'start_date']);

    const duration = extractField(data, ['duration', 'duration_minutes', 'length', 'duration_min']);

    const status = extractField(data, ['status', 'booking_status', 'state']) || 'new';

    const eventType = extractField(data, ['type', 'event', 'event_type', 'webhook_event', 'trigger', 'action']) || 'booking';

    const bookingId = extractField(data, ['id', 'booking_id', 'appointment_id', 'confirmation_number', 'reference']);

    const initiatedBy = extractField(data, ['canceled_by', 'rescheduled_by']);

    // Append row to sheet
    sheet.appendRow([
      new Date().toISOString(),           // timestamp
      clientName.trim() || 'Unknown',      // client_name
      clientEmail || '',                   // client_email
      serviceType || '',                   // service_type
      bookingDatetime || '',               // booking_datetime
      duration || '',                      // duration
      status,                              // status
      eventType,                           // event_type
      bookingId || '',                     // booking_id
      rawPayload                           // raw_payload (for debugging)
    ]);

    // Send Slack notification (after sheet append so logging isn't blocked)
    const packageName = data.package_name || serviceType || '';
    sendSlackNotification(clientName.trim(), packageName, bookingDatetime, eventType, bookingId, initiatedBy);

    // Return success
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Booking recorded'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Log error but still return 200 to avoid BLAB retries
    console.error('Webhook error:', error);

    // Try to log the error to sheet
    try {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
      sheet.appendRow([
        new Date().toISOString(),
        'ERROR',
        error.message,
        '', '', '', '', '',
        e.postData ? e.postData.contents : 'No payload'
      ]);
    } catch (logError) {
      console.error('Failed to log error:', logError);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Send booking notification to Slack
 */
function sendSlackNotification(name, service, datetime, eventType, bookingId, initiatedBy) {
  if (!SLACK_WEBHOOK_URL) return;

  // Suppress provider-initiated cancel/reschedule — those are Dan's actions, not client activity
  if (initiatedBy && initiatedBy.toString().toLowerCase() === 'provider') {
    console.log('Skipping Slack notification: provider-initiated ' + eventType);
    return;
  }

  // Determine emoji and label from event type
  const evtLower = (eventType || '').toLowerCase();
  let emoji = '📅', label = 'New booking';
  if (evtLower.includes('cancel')) {
    emoji = '❌'; label = 'Cancelled';
  } else if (evtLower.includes('reschedul')) {
    emoji = '🔄'; label = 'Rescheduled';
  }

  // Format datetime for display (CST)
  let dtDisplay = datetime || 'TBD';
  try {
    const dt = new Date(datetime);
    if (!isNaN(dt.getTime())) {
      dtDisplay = Utilities.formatDate(dt, 'America/Merida', 'EEE, MMM d \'at\' h:mm a');
    }
  } catch (e) { /* keep raw string */ }

  const blabLink = bookingId
    ? 'https://bookme.name/account/0/bookings/' + bookingId
    : 'https://bookme.name/account/0/orders';
  const displayName = name || 'Unknown';
  const displayService = service || 'Booking';

  const message = emoji + ' *' + label + ':* ' + displayName + '\n' +
                  '   ' + displayService + ' — ' + dtDisplay + '\n' +
                  '   <' + blabLink + '|View in BLAB>';

  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true
    });
  } catch (err) {
    console.error('Slack notification failed:', err.message);
  }
}

/**
 * Handles GET requests (for testing the endpoint)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'active',
    message: 'BLAB Webhook Receiver is running. Send POST requests to log bookings.',
    sheet_id: SHEET_ID
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper: Extract a field from data object, trying multiple possible field names
 * Also checks nested objects (e.g., data.booking.name)
 */
function extractField(data, possibleNames) {
  for (const name of possibleNames) {
    // Check top level
    if (data[name] !== undefined && data[name] !== null && data[name] !== '') {
      return String(data[name]);
    }

    // Check common nested objects
    const nestedObjects = ['booking', 'appointment', 'event', 'data', 'payload', 'customer', 'client'];
    for (const nested of nestedObjects) {
      if (data[nested] && typeof data[nested] === 'object') {
        if (data[nested][name] !== undefined && data[nested][name] !== null && data[nested][name] !== '') {
          return String(data[nested][name]);
        }
      }
    }
  }
  return '';
}

/**
 * Test function - simulates a webhook POST
 * Run this to verify the script works before deploying
 */
function testWebhook() {
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        event: 'booking.created',
        booking: {
          id: 'test-123',
          name: 'Test Customer',
          email: 'test@example.com',
          service: 'Discovery Call',
          start_time: '2026-01-30T14:00:00-05:00',
          duration: 30,
          status: 'confirmed'
        }
      })
    }
  };

  const result = doPost(mockEvent);
  console.log('Test result:', result.getContent());
}

// =============================================================================
// KIT EMAIL INTEGRATION
// =============================================================================
// Adds meeting guests to Kit email sequence AFTER their meeting occurs.
// Only processes general networking calls (appointment ID 35230), not sales calls.
// =============================================================================

// Kit API Configuration
// Get your API key from: Kit → Settings → Developer → API Keys
const KIT_API_KEY = PropertiesService.getScriptProperties().getProperty('KIT_API_KEY') || 'YOUR_API_KEY_HERE';
const KIT_TAG = 'meeting-guest-pending';
const KIT_TAG_ID = 15850596;
const TARGET_SERVICE = 'AI Strategy';
const TARGET_APPOINTMENT_ID = '279657';

/**
 * Daily trigger function - processes past meetings for Kit
 *
 * Set up trigger in Apps Script:
 * 1. Edit > Current project's triggers > Add trigger
 * 2. Function: processKitSubscriptions
 * 3. Event source: Time-driven
 * 4. Type: Day timer
 * 5. Time: 9am to 10am
 */
function processKitSubscriptions() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find column indices
  const cols = {
    email: headers.indexOf('client_email'),
    name: headers.indexOf('client_name'),
    serviceType: headers.indexOf('service_type'),
    bookingDatetime: headers.indexOf('booking_datetime'),
    status: headers.indexOf('status'),
    kitProcessed: headers.indexOf('kit_processed'),
    rawPayload: headers.indexOf('raw_payload')
  };

  // If kit_processed column doesn't exist, we can't track processing
  if (cols.kitProcessed === -1) {
    // Check if column K exists and is empty (header row)
    const lastCol = sheet.getLastColumn();
    if (lastCol < 11) {
      // Add kit_processed header
      sheet.getRange(1, 11).setValue('kit_processed');
      cols.kitProcessed = 10; // 0-indexed
    } else {
      cols.kitProcessed = 10; // Assume column K
    }
  }

  const now = new Date();
  let processed = 0;
  let skipped = 0;
  let errors = 0;

  // Process each row (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip if already processed
    if (row[cols.kitProcessed] === 'yes') {
      continue;
    }

    // Skip if no email
    const email = row[cols.email];
    if (!email || email.toString().trim() === '') {
      continue;
    }

    // Skip if this is an error row
    if (row[cols.name] === 'ERROR') {
      continue;
    }

    // Skip if canceled
    const status = row[cols.status];
    if (status && status.toString().toLowerCase() === 'canceled') {
      skipped++;
      continue;
    }

    // Check if this is a general networking meeting (not sales)
    const serviceType = row[cols.serviceType] ? row[cols.serviceType].toString() : '';
    const rawPayload = row[cols.rawPayload] ? row[cols.rawPayload].toString() : '';
    const isTargetMeeting = serviceType.toLowerCase().includes('online meeting') ||
                           rawPayload.includes(TARGET_APPOINTMENT_ID);

    if (!isTargetMeeting) {
      skipped++;
      continue;
    }

    // Check if meeting has already happened
    const bookingDatetime = row[cols.bookingDatetime];
    if (!bookingDatetime) {
      continue;
    }

    const meetingDate = new Date(bookingDatetime);
    if (isNaN(meetingDate.getTime())) {
      // Invalid date
      continue;
    }

    if (meetingDate > now) {
      // Meeting hasn't happened yet - skip for now
      continue;
    }

    // Add to Kit (with UTM attribution parsed from raw payload)
    const firstName = extractFirstName(row[cols.name]);
    const utmFields = parseTrackingSource(row[cols.rawPayload]);
    const success = addToKit(email.toString().trim(), firstName, utmFields);

    if (success) {
      // Mark as processed (column K = index 11 = cols.kitProcessed + 1)
      sheet.getRange(i + 1, cols.kitProcessed + 1).setValue('yes');
      processed++;
    } else {
      errors++;
    }

    // Small delay to avoid rate limiting
    if (processed % 5 === 0) {
      Utilities.sleep(1000);
    }
  }

  console.log(`Kit processing complete. Processed: ${processed}, Skipped: ${skipped}, Errors: ${errors}`);
  return { processed, skipped, errors };
}

/**
 * Add subscriber to Kit with tag
 * @param {string} email - Subscriber email
 * @param {string} firstName - Subscriber first name
 * @param {Object} utmFields - Optional utm_* fields to merge into Kit custom fields
 * @returns {boolean} - Success status
 */
function addToKit(email, firstName, utmFields) {
  if (KIT_API_KEY === 'YOUR_API_KEY_HERE') {
    console.error('Kit API key not configured. Set KIT_API_KEY in Script Properties.');
    return false;
  }

  const url = 'https://api.kit.com/v4/subscribers';

  const fields = {
    kit_first_form: 'BLAB Networking Call',
    kit_last_form: 'BLAB Networking Call'
  };
  if (utmFields && typeof utmFields === 'object') {
    Object.keys(utmFields).forEach(function(k) {
      if (utmFields[k]) fields[k] = utmFields[k];
    });
  }

  const payload = {
    email_address: email,
    first_name: firstName || '',
    tags: [KIT_TAG],
    fields: fields
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: {
      'X-Kit-Api-Key': KIT_API_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code === 200 || code === 201) {
      // Apply tag via dedicated endpoint (tag names don't work in subscriber payload)
      const tagUrl = `https://api.kit.com/v4/tags/${KIT_TAG_ID}/subscribers`;
      UrlFetchApp.fetch(tagUrl, {
        method: 'POST',
        contentType: 'application/json',
        headers: { 'X-Kit-Api-Key': KIT_API_KEY },
        payload: JSON.stringify({ email_address: email }),
        muteHttpExceptions: true
      });
      console.log(`Added to Kit: ${email}`);
      return true;
    } else if (code === 422) {
      // Subscriber may already exist - try to add tag instead
      console.log(`Subscriber may exist, attempting to add tag: ${email}`);
      return addTagToExistingSubscriber(email, utmFields);
    } else {
      console.error(`Kit API error for ${email}: ${code} - ${response.getContentText()}`);
      return false;
    }
  } catch (error) {
    console.error(`Kit API exception for ${email}: ${error.message}`);
    return false;
  }
}

/**
 * Add tag to existing Kit subscriber
 * @param {string} email - Subscriber email
 * @param {Object} utmFields - Optional utm_* fields to merge into Kit custom fields
 * @returns {boolean} - Success status
 */
function addTagToExistingSubscriber(email, utmFields) {
  // First, find the subscriber
  const findUrl = 'https://api.kit.com/v4/subscribers?email_address=' + encodeURIComponent(email);

  const findOptions = {
    method: 'GET',
    headers: {
      'X-Kit-Api-Key': KIT_API_KEY
    },
    muteHttpExceptions: true
  };

  try {
    const findResponse = UrlFetchApp.fetch(findUrl, findOptions);
    const findData = JSON.parse(findResponse.getContentText());

    if (findData.subscribers && findData.subscribers.length > 0) {
      const subscriberId = findData.subscribers[0].id;

      // Update Kit_Last_Form (but NOT kit_first_form - that should never change for existing subscribers).
      // Refresh utm_* so the most recent booking's attribution overwrites stale values.
      const updateUrl = `https://api.kit.com/v4/subscribers/${subscriberId}`;
      const updateFields = { kit_last_form: 'BLAB Networking Call' };
      if (utmFields && typeof utmFields === 'object') {
        Object.keys(utmFields).forEach(function(k) {
          if (utmFields[k]) updateFields[k] = utmFields[k];
        });
      }
      const updatePayload = { fields: updateFields };

      const updateOptions = {
        method: 'PUT',
        contentType: 'application/json',
        headers: {
          'X-Kit-Api-Key': KIT_API_KEY
        },
        payload: JSON.stringify(updatePayload),
        muteHttpExceptions: true
      };

      UrlFetchApp.fetch(updateUrl, updateOptions);

      // Add tag to subscriber (must use numeric tag ID, not name)
      const tagUrl = `https://api.kit.com/v4/tags/${KIT_TAG_ID}/subscribers`;
      const tagPayload = { email_address: email };

      const tagOptions = {
        method: 'POST',
        contentType: 'application/json',
        headers: {
          'X-Kit-Api-Key': KIT_API_KEY
        },
        payload: JSON.stringify(tagPayload),
        muteHttpExceptions: true
      };

      const tagResponse = UrlFetchApp.fetch(tagUrl, tagOptions);
      const tagCode = tagResponse.getResponseCode();

      if (tagCode === 200 || tagCode === 201) {
        console.log(`Added tag to existing subscriber: ${email}`);
        return true;
      }
    }

    return false;
  } catch (error) {
    console.error(`Failed to add tag to existing subscriber ${email}: ${error.message}`);
    return false;
  }
}

/**
 * Extract first name from full name
 * @param {string} fullName - Full name
 * @returns {string} - First name
 */
function extractFirstName(fullName) {
  if (!fullName) return '';
  return fullName.toString().trim().split(' ')[0];
}

/**
 * Parse BLAB raw_payload JSON for the tracking_source custom field and return
 * Kit-ready utm_* fields. Returns {} if nothing usable is present.
 *
 * Expected format in BLAB: custom_field_tracking_source set via Custom Code to
 * a pipe-delimited string: "source=dcl-site|medium=cta|campaign=...|content=...|term=..."
 *
 * Also tolerates a bare URL (takes its query string) or a raw query string.
 *
 * @param {string} rawPayload - JSON string from raw_payload column
 * @returns {Object} - { utm_source, utm_medium, utm_campaign, utm_content, utm_placement }
 */
function parseTrackingSource(rawPayload) {
  if (!rawPayload) return {};
  let data;
  try {
    data = JSON.parse(rawPayload.toString());
  } catch (e) {
    return {};
  }

  // BLAB serializes custom fields as custom_field_<slug>. We look at a small
  // set of likely slugs in case Dan names the field differently.
  const candidates = [
    'custom_field_tracking_source',
    'custom_field_source',
    'custom_field_utm',
    'custom_field_attribution'
  ];
  let raw = '';
  for (let i = 0; i < candidates.length; i++) {
    const v = data[candidates[i]];
    if (v && typeof v === 'string' && v.trim()) { raw = v.trim(); break; }
  }
  if (!raw) return {};

  // If it's a URL or URL fragment, strip to the query string.
  const qIdx = raw.indexOf('?');
  if (qIdx !== -1) raw = raw.slice(qIdx + 1);

  // Split on either `|` (our preferred delimiter) or `&` (url-style).
  const pairs = raw.indexOf('|') !== -1 ? raw.split('|') : raw.split('&');

  const out = {};
  // Map inbound keys to Kit custom field names. BLAB Custom Code can write
  // either short keys (source, medium, ...) or full utm_* keys.
  const keyMap = {
    source: 'utm_source',
    utm_source: 'utm_source',
    medium: 'utm_medium',
    utm_medium: 'utm_medium',
    campaign: 'utm_campaign',
    utm_campaign: 'utm_campaign',
    content: 'utm_content',
    utm_content: 'utm_content',
    term: 'utm_placement',
    utm_term: 'utm_placement',
    placement: 'utm_placement',
    utm_placement: 'utm_placement'
  };

  for (let i = 0; i < pairs.length; i++) {
    const eq = pairs[i].indexOf('=');
    if (eq === -1) continue;
    const k = pairs[i].slice(0, eq).trim().toLowerCase();
    let v = pairs[i].slice(eq + 1).trim();
    if (!v) continue;
    try { v = decodeURIComponent(v); } catch (e) { /* keep raw */ }
    const kitKey = keyMap[k];
    if (kitKey) out[kitKey] = v;
  }

  return out;
}

/**
 * Manual test function for Kit integration
 * Run this to verify Kit API connection works
 */
function testKitIntegration() {
  console.log('Testing Kit API connection...');
  console.log('API Key configured:', KIT_API_KEY !== 'YOUR_API_KEY_HERE');

  if (KIT_API_KEY === 'YOUR_API_KEY_HERE') {
    console.error('ERROR: Kit API key not configured!');
    console.log('To configure:');
    console.log('1. Go to Kit → Settings → Developer → API Keys');
    console.log('2. Create a new API key');
    console.log('3. In Apps Script: File → Project properties → Script properties');
    console.log('4. Add property: KIT_API_KEY = your_api_key');
    return;
  }

  // Test API connection by listing tags
  const url = 'https://api.kit.com/v4/tags';
  const options = {
    method: 'GET',
    headers: {
      'X-Kit-Api-Key': KIT_API_KEY
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      console.log('SUCCESS: Kit API connection working!');
      console.log('Found ' + (data.tags ? data.tags.length : 0) + ' tags');

      // Check if our target tag exists
      const targetTag = data.tags ? data.tags.find(t => t.name === KIT_TAG) : null;
      if (targetTag) {
        console.log('Target tag "' + KIT_TAG + '" exists (ID: ' + targetTag.id + ')');
      } else {
        console.warn('WARNING: Target tag "' + KIT_TAG + '" not found. Create it in Kit first.');
      }
    } else {
      console.error('Kit API error: ' + code + ' - ' + response.getContentText());
    }
  } catch (error) {
    console.error('Kit API exception: ' + error.message);
  }
}

/**
 * Run a dry run of processKitSubscriptions without making changes
 * Shows what would be processed
 */
function dryRunKitProcessing() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const cols = {
    email: headers.indexOf('client_email'),
    name: headers.indexOf('client_name'),
    serviceType: headers.indexOf('service_type'),
    bookingDatetime: headers.indexOf('booking_datetime'),
    status: headers.indexOf('status'),
    kitProcessed: headers.indexOf('kit_processed'),
    rawPayload: headers.indexOf('raw_payload')
  };

  const now = new Date();
  let wouldProcess = [];
  let wouldSkip = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (row[cols.kitProcessed] === 'yes') continue;

    const email = row[cols.email];
    if (!email || email.toString().trim() === '') continue;
    if (row[cols.name] === 'ERROR') continue;

    const status = row[cols.status];
    if (status && status.toString().toLowerCase() === 'canceled') {
      wouldSkip.push({ row: i + 1, email, reason: 'canceled' });
      continue;
    }

    const serviceType = row[cols.serviceType] ? row[cols.serviceType].toString() : '';
    const rawPayload = row[cols.rawPayload] ? row[cols.rawPayload].toString() : '';
    const isTargetMeeting = serviceType.toLowerCase().includes('online meeting') ||
                           rawPayload.includes(TARGET_APPOINTMENT_ID);

    if (!isTargetMeeting) {
      wouldSkip.push({ row: i + 1, email, reason: 'not target meeting type' });
      continue;
    }

    const bookingDatetime = row[cols.bookingDatetime];
    if (!bookingDatetime) continue;

    const meetingDate = new Date(bookingDatetime);
    if (isNaN(meetingDate.getTime())) continue;

    if (meetingDate > now) {
      wouldSkip.push({ row: i + 1, email, reason: 'meeting not yet occurred' });
      continue;
    }

    wouldProcess.push({
      row: i + 1,
      email: email.toString(),
      name: row[cols.name],
      meetingDate: meetingDate.toISOString()
    });
  }

  console.log('=== DRY RUN RESULTS ===');
  console.log('Would process ' + wouldProcess.length + ' subscribers:');
  wouldProcess.forEach(p => console.log('  Row ' + p.row + ': ' + p.email + ' (' + p.name + ')'));

  console.log('\nWould skip ' + wouldSkip.length + ' rows:');
  wouldSkip.slice(0, 10).forEach(s => console.log('  Row ' + s.row + ': ' + s.email + ' - ' + s.reason));
  if (wouldSkip.length > 10) console.log('  ... and ' + (wouldSkip.length - 10) + ' more');

  return { wouldProcess, wouldSkip };
}
