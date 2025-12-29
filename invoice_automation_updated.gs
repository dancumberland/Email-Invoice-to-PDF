const CONFIG = {
  SOURCE_LABEL: 'Invoices/New-Invoices',
  PROCESSED_LABEL: 'Invoices/Processed-Invoices',
  FOLDER_ID: '1-8JhkitHj9Y2iLk-yaabwqtiuIQO21P5',
};

function processInvoices() {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const sourceLabel = GmailApp.getUserLabelByName(CONFIG.SOURCE_LABEL);
  const processedLabel = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  
  if (!sourceLabel) {
    Logger.log('Source label "' + CONFIG.SOURCE_LABEL + '" not found. Please create it in Gmail.');
    return;
  }

  const query = 'label:' + CONFIG.SOURCE_LABEL.replace('/', '-');
  const threads = GmailApp.search(query, 0, 50);

  threads.forEach(thread => {
    handleThread_(thread, folder);
    thread.removeLabel(sourceLabel);
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

  const allAttachments = message.getAttachments({
    includeInlineImages: true,
    includeAttachments: true,
  }) || [];

  const htmlBody = message.getBody() || '';

  // Separate image attachments from document attachments (PDFs, docs, etc.)
  const imageAttachments = [];
  const documentAttachments = [];

  allAttachments.forEach(att => {
    const mimeType = (att.getContentType() || '').toLowerCase();
    if (mimeType.startsWith('image/')) {
      imageAttachments.push(att);
    } else {
      documentAttachments.push(att);
    }
  });

  // Start with email body HTML, embedding any inline images
  let html = embedAllImages_(htmlBody || wrapPlainAsHtml_(bodyPlain), imageAttachments);

  // Convert document attachments (PDFs, etc.) to thumbnail images and append to HTML
  const uniqueDocuments = uniqueAttachments_(documentAttachments);
  if (uniqueDocuments.length > 0) {
    const attachmentHtml = convertAttachmentsToHtml_(uniqueDocuments);
    if (attachmentHtml) {
      // Insert before closing body tag if exists, otherwise append
      if (html.includes('</body>')) {
        html = html.replace('</body>', attachmentHtml + '</body>');
      } else {
        html += attachmentHtml;
      }
    }
  }

  // Create single consolidated PDF
  const htmlBlob = Utilities.newBlob(html, 'text/html', filenameBase + '.html');
  const pdfBlob = htmlBlob.getAs('application/pdf').setName(filenameBase + '.pdf');
  folder.createFile(pdfBlob);
}

/**
 * Convert document attachments to HTML with embedded thumbnail images.
 * Uses Drive thumbnails for PDFs; embeds other documents as download links.
 */
function convertAttachmentsToHtml_(attachments) {
  if (!attachments || attachments.length === 0) return '';

  let html = '<br><hr style="margin: 30px 0; border: 2px solid #333;"><div style="margin-top: 20px;">';
  html += '<h2 style="font-family: Arial, sans-serif; color: #333; margin-bottom: 20px;">Attached Documents</h2>';

  attachments.forEach((att, index) => {
    const name = att.getName() || 'Attachment ' + (index + 1);
    const mimeType = (att.getContentType() || '').toLowerCase();

    html += `<div style="margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px;">`;
    html += `<h3 style="font-family: Arial, sans-serif; color: #555; margin: 0 0 15px 0;">${escapeHtml_(name)}</h3>`;

    if (mimeType === 'application/pdf') {
      // Get thumbnail for PDF
      const thumbnailDataUri = getPdfThumbnail_(att);
      if (thumbnailDataUri) {
        html += `<img src="${thumbnailDataUri}" style="max-width: 100%; height: auto; border: 1px solid #ccc;">`;
      } else {
        html += `<p style="color: #999; font-style: italic;">[PDF preview not available]</p>`;
      }
    } else {
      // For non-PDF documents, just show the filename
      html += `<p style="color: #666;">[Document: ${escapeHtml_(mimeType)}]</p>`;
    }

    html += '</div>';
  });

  html += '</div>';
  return html;
}

/**
 * Get a thumbnail image of a PDF as a data URI.
 * Uploads to Drive temporarily to access the thumbnail, then cleans up.
 * Retries a few times since Drive needs time to generate thumbnails.
 */
function getPdfThumbnail_(attachment) {
  let tempFile = null;
  try {
    // Upload PDF to Drive temporarily
    const blob = attachment.copyBlob();
    tempFile = DriveApp.createFile(blob);
    const fileId = tempFile.getId();

    // Retry up to 5 times with increasing delays (Drive needs time to generate thumbnail)
    const maxRetries = 5;
    const delays = [2000, 3000, 4000, 5000, 5000]; // milliseconds

    for (let attempt = 0; attempt < maxRetries; attempt++) {
      // Wait before checking (Drive needs processing time)
      Utilities.sleep(delays[attempt]);

      // Use Advanced Drive Service to get thumbnail link
      const file = Drive.Files.get(fileId, { fields: 'thumbnailLink' });

      if (file.thumbnailLink) {
        // Fetch the thumbnail image (increase size for better quality)
        const thumbnailUrl = file.thumbnailLink.replace(/=s\d+$/, '=s800');
        const response = UrlFetchApp.fetch(thumbnailUrl, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true,
        });

        if (response.getResponseCode() === 200) {
          const imageBlob = response.getBlob();
          const base64 = Utilities.base64Encode(imageBlob.getBytes());
          const mimeType = imageBlob.getContentType() || 'image/png';
          return `data:${mimeType};base64,${base64}`;
        }
      }

      Logger.log('Thumbnail not ready, attempt ' + (attempt + 1) + ' of ' + maxRetries);
    }

    Logger.log('Failed to get thumbnail after ' + maxRetries + ' attempts');
    return null;
  } catch (e) {
    Logger.log('Error getting PDF thumbnail: ' + e.message);
    return null;
  } finally {
    // Clean up temporary file
    if (tempFile) {
      try {
        tempFile.setTrashed(true);
      } catch (e) {
        Logger.log('Error cleaning up temp file: ' + e.message);
      }
    }
  }
}

function escapeHtml_(text) {
  return (text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
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

function resolveTransactionDate_(bodyPlain, fallbackDate) {
  const inlineDate = extractInlineDate_(bodyPlain);
  if (inlineDate) return inlineDate;

  const quotedDate = extractQuotedForwardDate_(bodyPlain);
  if (quotedDate) return quotedDate;

  const headerDate = extractForwardHeaderDate_(bodyPlain);
  if (headerDate) return headerDate;

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

function extractForwardHeaderDate_(bodyPlain = '') {
  const headerRegex = /^Date:\s+(.+)$/im;
  const match = bodyPlain.match(headerRegex);
  if (!match) return null;

  // Remove trailing "at ..." blocks for stability
  const cleaned = match[1].replace(/\sat\s.+$/i, '').trim();
  const parsed = new Date(cleaned);
  return isNaN(parsed.getTime()) ? null : parsed;
}

// Extract the domain name from the forwarded email's From: header.
// "From: Cloudflare <noreply@notify.cloudflare.com>" -> "Cloudflare"
// Falls back to extracting from domain if no display name present.
function extractForwardedSenderName_(bodyPlain = '') {
  // Match "From: Display Name <email@domain.com>" or "From: email@domain.com"
  const fromRegex = /^From:\s*(?:(.+?)\s*<[^>]+>|([^<\s]+@([^>\s]+)))\s*$/im;
  const match = bodyPlain.match(fromRegex);
  if (!match) return null;

  // If there's a display name (e.g., "Cloudflare"), use it
  if (match[1] && match[1].trim()) {
    return match[1].trim();
  }

  // Otherwise extract domain name from email address
  if (match[3]) {
    const domain = match[3].split('.')[0];
    return domain.charAt(0).toUpperCase() + domain.slice(1).toLowerCase();
  }

  return null;
}

/**
 * Embed ALL image attachments into HTML.
 * Strategy: 
 * 1. Replace cid: references with matching image data URIs
 * 2. For any images that couldn't be matched, append them at the end
 * This ensures all images show up in the PDF regardless of CID matching issues.
 */
function embedAllImages_(html, imageAttachments = []) {
  if (!html) html = '';
  if (!imageAttachments || imageAttachments.length === 0) return html;

  // Build data URIs for all image attachments
  const imageData = imageAttachments.map(att => {
    const name = (att.getName() || '').toLowerCase();
    const blob = att.copyBlob();
    const mimeType = blob.getContentType() || 'image/png';
    const base64 = Utilities.base64Encode(blob.getBytes());
    const dataUri = `data:${mimeType};base64,${base64}`;
    return { name, dataUri, used: false };
  });

  // Replace cid: references with matching image data URIs
  let resultHtml = html.replace(/src=(['"])cid:([^'"]+)\1/gi, (match, quote, cid) => {
    const cidLower = cid.toLowerCase();
    
    // Try to find a matching image
    for (const img of imageData) {
      const nameWithoutExt = img.name.replace(/\.[^.]+$/, '');
      const nameNoDots = img.name.replace(/\./g, '');
      
      if (cidLower.includes(nameWithoutExt) || nameWithoutExt.includes(cidLower) ||
          cidLower.includes(img.name) || img.name.includes(cidLower) ||
          cidLower.includes(nameNoDots) || nameNoDots.includes(cidLower)) {
        img.used = true;
        return `src=${quote}${img.dataUri}${quote}`;
      }
    }
    
    // If no match, use the first unused image as fallback
    for (const img of imageData) {
      if (!img.used) {
        img.used = true;
        return `src=${quote}${img.dataUri}${quote}`;
      }
    }
    
    return match;
  });

  // Append any unused images at the end of the HTML
  // This ensures photos you embed manually still show up even if CID matching fails
  const unusedImages = imageData.filter(img => !img.used);
  if (unusedImages.length > 0) {
    let appendHtml = '<br><hr style="margin: 20px 0;"><div style="margin-top: 20px;">';
    unusedImages.forEach((img, i) => {
      appendHtml += `<div style="margin: 10px 0;"><img src="${img.dataUri}" style="max-width: 100%; height: auto;"></div>`;
    });
    appendHtml += '</div>';
    
    // Insert before closing body tag if exists, otherwise append
    if (resultHtml.includes('</body>')) {
      resultHtml = resultHtml.replace('</body>', appendHtml + '</body>');
    } else {
      resultHtml += appendHtml;
    }
  }

  return resultHtml;
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
