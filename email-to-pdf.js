// Email to PDF/HTML conversion utilities for Paperless-ngx PDF Uploader

/**
 * Get icon for attachment based on file type
 * @param {string} filename - The filename of the attachment
 * @param {string} contentType - The MIME type of the attachment
 * @returns {string} Emoji icon for the file type
 */
function getAttachmentIcon(filename, contentType) {
  const ext = filename.toLowerCase().split('.').pop();
  
  // Check by extension first
  const extensionIcons = {
    'pdf': '📄',
    'doc': '📝',
    'docx': '📝',
    'xls': '📊',
    'xlsx': '📊',
    'ppt': '📽️',
    'pptx': '📽️',
    'jpg': '🖼️',
    'jpeg': '🖼️',
    'png': '🖼️',
    'gif': '🖼️',
    'bmp': '🖼️',
    'webp': '🖼️',
    'svg': '🖼️',
    'zip': '📦',
    'rar': '📦',
    '7z': '📦',
    'tar': '📦',
    'gz': '📦',
    'txt': '📃',
    'csv': '📊',
    'mp3': '🎵',
    'wav': '🎵',
    'mp4': '🎬',
    'avi': '🎬',
    'mov': '🎬',
    'eml': '✉️',
    'msg': '✉️'
  };

  if (extensionIcons[ext]) {
    return extensionIcons[ext];
  }

  // Check by content type
  if (contentType) {
    if (contentType.startsWith('image/')) return '🖼️';
    if (contentType.startsWith('audio/')) return '🎵';
    if (contentType.startsWith('video/')) return '🎬';
    if (contentType.includes('pdf')) return '📄';
    if (contentType.includes('word') || contentType.includes('document')) return '📝';
    if (contentType.includes('excel') || contentType.includes('spreadsheet')) return '📊';
    if (contentType.includes('powerpoint') || contentType.includes('presentation')) return '📽️';
    if (contentType.includes('zip') || contentType.includes('compressed')) return '📦';
  }

  // Default icon
  return '📎';
}

/**
 * Format date for German locale
 * @param {Date|string} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDateGerman(date) {
  const d = new Date(date);
  return d.toLocaleDateString('de-DE', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

/**
 * Escape HTML special characters
 * @param {string} str - String to escape
 * @returns {string} Escaped string
 */
function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

/**
 * Generate HTML content for email to be uploaded as PDF
 * @param {Object} emailData - Email data object
 * @param {string} emailData.subject - Email subject
 * @param {string} emailData.author - Email sender
 * @param {string|Date} emailData.date - Email date
 * @param {string[]} emailData.recipients - Email recipients
 * @param {string} emailData.body - Email body (plain text or HTML)
 * @param {boolean} emailData.isHtml - Whether the body is HTML
 * @param {Array} emailData.attachments - List of attachments
 * @returns {string} HTML content for PDF conversion
 */
function generateEmailHtml(emailData) {
  const {
    subject,
    author,
    date,
    recipients,
    body,
    isHtml,
    attachments
  } = emailData;

  // Format recipients
  const recipientList = Array.isArray(recipients) ? recipients.join(', ') : recipients || '';
  
  // Format attachments with icons
  let attachmentsHtml = '';
  if (attachments && attachments.length > 0) {
    const attachmentItems = attachments.map(att => {
      const icon = getAttachmentIcon(att.name, att.contentType);
      return `<span style="margin-right: 12px; white-space: nowrap;">${icon} ${escapeHtml(att.name)}</span>`;
    }).join('');
    
    attachmentsHtml = `
      <div style="margin-top: 8px;">
        <strong>Anhänge:</strong> ${attachmentItems}
      </div>
    `;
  }

  // Process body content
  let bodyContent = '';
  if (isHtml) {
    // For HTML content, use it directly but sanitize dangerous elements
    bodyContent = body;
  } else {
    // For plain text, convert to HTML with proper line breaks
    bodyContent = `<pre style="white-space: pre-wrap; word-wrap: break-word; font-family: inherit; margin: 0;">${escapeHtml(body)}</pre>`;
  }

  const html = `<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(subject)}</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      max-width: 800px;
      margin: 0 auto;
      padding: 20px;
    }
    .email-header {
      background-color: #f5f5f5;
      border: 1px solid #ddd;
      border-radius: 6px;
      padding: 16px 20px;
      margin-bottom: 20px;
    }
    .email-header-row {
      margin: 6px 0;
      font-size: 14px;
    }
    .email-header-label {
      color: #666;
      min-width: 60px;
      display: inline-block;
    }
    .email-subject {
      font-size: 16px;
      margin: 10px 0 6px 0;
    }
    .email-subject strong {
      font-weight: 700;
    }
    .email-divider {
      border: none;
      border-top: 2px solid #ddd;
      margin: 20px 0;
    }
    .email-body {
      padding: 10px 0;
    }
    .email-attachments {
      font-size: 13px;
      color: #555;
      margin-top: 10px;
      padding-top: 10px;
      border-top: 1px dashed #ddd;
    }
    blockquote {
      border-left: 3px solid #ccc;
      margin: 10px 0;
      padding-left: 15px;
      color: #666;
    }
    pre {
      background-color: #f8f8f8;
      padding: 10px;
      border-radius: 4px;
      overflow-x: auto;
    }
    a {
      color: #0066cc;
    }
  </style>
</head>
<body>
  <div class="email-header">
    <div class="email-header-row">
      <span class="email-header-label">Datum:</span> ${escapeHtml(formatDateGerman(date))}
    </div>
    <div class="email-header-row">
      <span class="email-header-label">Von:</span> ${escapeHtml(author)}
    </div>
    <div class="email-header-row">
      <span class="email-header-label">An:</span> ${escapeHtml(recipientList)}
    </div>
    <div class="email-subject">
      <span class="email-header-label">Betreff:</span> <strong>${escapeHtml(subject)}</strong>
    </div>
    ${attachmentsHtml ? `<div class="email-attachments">${attachmentsHtml}</div>` : ''}
  </div>
  
  <hr class="email-divider">
  
  <div class="email-body">
    ${bodyContent}
  </div>
</body>
</html>`;

  return html;
}

/**
 * Generate a filename from email subject
 * @param {string} subject - Email subject
 * @param {Date|string} date - Email date
 * @returns {string} Sanitized filename
 */
function generateEmailFilename(subject, date) {
  // Format date as YYYY-MM-DD
  const d = new Date(date);
  const dateStr = d.toISOString().split('T')[0];
  
  // Sanitize subject for filename (remove special characters)
  let sanitizedSubject = (subject || 'Email')
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '') // Remove invalid characters
    .replace(/\s+/g, '_')                   // Replace spaces with underscores
    .substring(0, 100);                      // Limit length
  
  return `${dateStr}_${sanitizedSubject}.html`;
}

// Export functions for use in other files
if (typeof window !== 'undefined') {
  window.getAttachmentIcon = getAttachmentIcon;
  window.formatDateGerman = formatDateGerman;
  window.escapeHtml = escapeHtml;
  window.generateEmailHtml = generateEmailHtml;
  window.generateEmailFilename = generateEmailFilename;
}
