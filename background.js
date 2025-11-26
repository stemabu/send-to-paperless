// Background script for Paperless-ngx PDF Uploader
console.log("Paperless PDF Uploader loaded!");

let currentPdfAttachments = [];
let currentMessage = null;

// Constants
const PAPERLESS_TAG_KEY = 'paperless';

// Create context menus for attachments
browser.runtime.onInstalled.addListener(async () => {
  // Remove all existing menus first to avoid conflicts
  await browser.menus.removeAll();

  // Message list context menus
  // Quick upload option
  browser.menus.create({
    id: "quick-upload-pdf-paperless",
    title: "Quick Upload to Paperless-ngx",
    contexts: ["message_list"],
    icons: {
      "32": "icons/icon-32.png",
      "16": "icons/icon-16.png",
      "64": "icons/icon-64.png",
      "128": "icons/icon-128.png"
    }
  });

  // Advanced upload option with dialog
  browser.menus.create({
    id: "advanced-upload-pdf-paperless",
    title: "Upload to Paperless-ngx (with options)...",
    contexts: ["message_list"],
    icons: {
      "32": "icons/icon-32.png",
      "16": "icons/icon-16.png",
      "64": "icons/icon-64.png",
      "128": "icons/icon-128.png"
    }
  });

  // Separator
  browser.menus.create({
    id: "separator",
    type: "separator",
    contexts: ["message_list"]
  });

  // Email upload option (entire email as PDF)
  browser.menus.create({
    id: "upload-email-paperless",
    title: "E-Mail an Paperless-ngx senden",
    contexts: ["message_list"],
    icons: {
      "32": "icons/icon-32.png",
      "16": "icons/icon-16.png",
      "64": "icons/icon-64.png",
      "128": "icons/icon-128.png"
    }
  });
});

// Handle context menu clicks
browser.menus.onClicked.addListener(async (info, tab) => {
  // Message list context menu handlers
  if (info.menuItemId === "quick-upload-pdf-paperless") {
    await handleQuickPdfUpload(info);
  } else if (info.menuItemId === "advanced-upload-pdf-paperless") {
    await handleAdvancedPdfUpload(info);
  } else if (info.menuItemId === "upload-email-paperless") {
    await handleEmailUpload(info);
  }
});

async function handleQuickPdfUpload(info) {
  try {
    const messages = info.selectedMessages.messages;
    if (!messages || messages.length === 0) {
      showNotification("No messages selected", "error");
      return;
    }

    // Process each selected message for PDF attachments
    for (const message of messages) {
      await processQuickPdfUpload(message);
    }
  } catch (error) {
    console.error("Error handling quick PDF upload:", error);
    showNotification("Error processing attachments", "error");
  }
}

async function handleAdvancedPdfUpload(info) {
  try {
    const messages = info.selectedMessages.messages;
    if (!messages || messages.length === 0) {
      showNotification("No messages selected", "error");
      return;
    }

    // For now, just handle the first message (can be extended)
    const message = messages[0];

    // Get PDF attachments
    const attachments = await browser.messages.listAttachments(message.id);
    const pdfAttachments = attachments.filter(attachment =>
      attachment.contentType === "application/pdf" ||
      attachment.name.toLowerCase().endsWith('.pdf')
    );

    if (pdfAttachments.length === 0) {
      showNotification("No PDF attachments found in selected message", "info");
      return;
    }

    // Store current data for the dialog
    currentMessage = message;
    currentPdfAttachments = pdfAttachments;

    // Open the advanced upload dialog
    await openAdvancedUploadDialog(message, pdfAttachments);

  } catch (error) {
    console.error("Error handling advanced PDF upload:", error);
    showNotification("Error processing attachments", "error");
  }
}

async function processQuickPdfUpload(message) {
  try {
    const attachments = await browser.messages.listAttachments(message.id);
    const pdfAttachments = attachments.filter(attachment =>
      attachment.contentType === "application/pdf" ||
      attachment.name.toLowerCase().endsWith('.pdf')
    );

    if (pdfAttachments.length === 0) {
      showNotification("No PDF attachments found in selected messages", "info");
      return;
    }

    // If there's only one attachment, upload directly
    if (pdfAttachments.length === 1) {
      await uploadPdfToPaperless(message, pdfAttachments[0], { mode: 'quick' });
      return;
    }

    // If there are multiple attachments, show selection dialog
    await openAttachmentSelectionDialog(message, pdfAttachments);

  } catch (error) {
    console.error("Error processing PDF attachments:", error);
    showNotification(`Error processing attachments: ${error.message}`, "error");
  }
}

async function openAttachmentSelectionDialog(message, pdfAttachments) {
  try {
    // Store data for the dialog to access
    await browser.storage.local.set({
      quickUploadData: {
        message: {
          id: message.id,
          subject: message.subject,
          author: message.author,
          date: message.date
        },
        attachments: pdfAttachments.map(att => ({
          name: att.name,
          partName: att.partName,
          size: att.size
        }))
      }
    });

    // Open the selection dialog
    const dialogUrl = browser.runtime.getURL("select-attachments.html");
    browser.windows.create({
      url: dialogUrl,
      type: "popup",
      width: 500,
      height: 600
    });
  } catch (error) {
    console.error("Error opening attachment selection dialog:", error);
    showNotification("Error opening attachment selection dialog", "error");
  }
}

async function openAdvancedUploadDialog(message, pdfAttachments) {
  // Create a new window/tab for the upload dialog
  const dialogUrl = browser.runtime.getURL("upload-dialog.html");

  try {
    // Store data for the dialog to access
    await browser.storage.local.set({
      currentUploadData: {
        message: {
          id: message.id,
          subject: message.subject,
          author: message.author,
          date: message.date
        },
        attachments: pdfAttachments.map(att => ({
          name: att.name,
          partName: att.partName,
          size: att.size
        }))
      }
    });

    // Open the dialog
    browser.windows.create({
      url: dialogUrl,
      type: "popup",
      width: 550,
      height: 700
    });
  } catch (error) {
    console.error("Error opening dialog:", error);
    showNotification("Error opening upload dialog", "error");
  }
}

async function uploadPdfToPaperless(message, attachment, options = {}) {
  try {
    const config = await getPaperlessConfig();
    if (!config.url || !config.token) {
      showNotification("Please configure Paperless-ngx settings first", "error");
      return { success: false, error: "Paperless-ngx not configured" };
    }

    const uploadMode = options.mode || 'quick';
    showNotification(`Uploading ${attachment.name} to Paperless-ngx...`, "info");

    // Get attachment data
    const attachmentData = await browser.messages.getAttachmentFile(
      message.id,
      attachment.partName
    );

    // Prepare form data for upload
    const formData = new FormData();
    formData.append('document', attachmentData, attachment.name);

    // Prepare metadata based on mode
    let metadata = {};

    if (uploadMode === 'quick') {
      // Minimal metadata for quick upload
      metadata = {
        title: attachment.name.replace(/\.pdf$/i, ''), // Remove .pdf extension
      };
    } else if (uploadMode === 'advanced') {
      // Use provided options for advanced upload
      metadata = {
        title: options.title || attachment.name.replace(/\.pdf$/i, ''),
        correspondent: options.correspondent,
        document_type: options.document_type,
        tags: options.tags || [],
        created: options.created,
        source: options.source || 'Thunderbird Email',
      };
    }

    // Add metadata to form data (only if values exist)
    Object.keys(metadata).forEach(key => {
      if (metadata[key] !== undefined && metadata[key] !== null && metadata[key] !== '') {
        if (Array.isArray(metadata[key])) {
          metadata[key].forEach(item => formData.append(key, item));
        } else {
          formData.append(key, metadata[key]);
        }
      }
    });

    // Upload to Paperless-ngx
    const response = await fetch(`${config.url}/api/documents/post_document/`, {
      method: 'POST',
      headers: {
        'Authorization': `Token ${config.token}`
      },
      body: formData
    });

    if (response.ok) {
      const result = await response.json();
      showNotification(`✅ Successfully uploaded ${attachment.name} to Paperless-ngx`, "success");
      console.log("Upload successful:", result);

      // Return success data for dialog callback
      return { success: true, result };
    } else {
      const errorText = await response.text();
      throw new Error(`HTTP ${response.status}: ${errorText}`);
    }

  } catch (error) {
    console.error("Error uploading to Paperless-ngx:", error);
    showNotification(`❌ Failed to upload ${attachment.name}: ${error.message}`, "error");
    return { success: false, error: error.message };
  }
}

// Handle messages from the upload dialog
browser.runtime.onMessage.addListener(async (message, sender, sendResponse) => {
  if (message.action === "quickUploadFromDisplay") {
    await handleQuickUploadFromDisplay(message.messageId);
    sendResponse({ success: true });
    return true;
  }

  if (message.action === "advancedUploadFromDisplay") {
    await handleAdvancedUploadFromDisplay(message.messageId);
    sendResponse({ success: true });
    return true;
  }

  if (message.action === "quickUploadSelected") {
    try {
      const { messageData, selectedAttachments } = message;

      let successCount = 0;
      let errorCount = 0;

      // Upload each selected attachment
      for (const attachment of selectedAttachments) {
        try {
          const result = await uploadPdfToPaperless(
            messageData,
            attachment,
            { mode: 'quick' }
          );

          if (result.success) {
            successCount++;
          } else {
            errorCount++;
          }
        } catch (error) {
          errorCount++;
          console.error(`Error uploading ${attachment.name}:`, error);
        }
      }

      // Show summary notification
      if (successCount > 0 && errorCount === 0) {
        showNotification(`✅ Successfully uploaded ${successCount} document(s) to Paperless-ngx`, "success");
      } else if (successCount > 0) {
        showNotification(`⚠️ Uploaded ${successCount} document(s), ${errorCount} failed`, "info");
      } else {
        showNotification(`❌ Failed to upload all documents`, "error");
      }

      sendResponse({ success: true, successCount, errorCount });
    } catch (error) {
      console.error("Error in quickUploadSelected:", error);
      sendResponse({ success: false, error: error.message });
    }
    return true; // Keep the message channel open for async response
  }

  if (message.action === "uploadWithOptions") {
    console.log('📤 Background: Received uploadWithOptions message');

    (async () => {
      try {
        const { messageData, attachmentData, uploadOptions } = message;
        console.log('📤 Background: Processing upload for:', attachmentData.name);

        // Reconstruct message and attachment objects
        const messageObj = messageData;
        const attachmentObj = attachmentData;

        const result = await uploadPdfToPaperless(
          messageObj,
          attachmentObj,
          { mode: 'advanced', ...uploadOptions }
        );

        console.log('📋 Background: Upload result for', attachmentData.name, ':', result);
        console.log('📋 Background: About to send response:', JSON.stringify(result));

        // Ensure we always send a valid response
        if (result && typeof result === 'object' && result.hasOwnProperty('success')) {
          console.log('📋 Background: Sending valid result');
          sendResponse(result);
        } else {
          console.error('❌ Background: Invalid result, sending error response:', result);
          sendResponse({ success: false, error: "Invalid response from upload function" });
        }
      } catch (error) {
        console.error("❌ Background: Error in upload with options:", error);
        sendResponse({ success: false, error: error.message });
      }
    })();

    return true; // Keep the message channel open for async response
  }

  if (message.action === "getCorrespondents") {
    try {
      const config = await getPaperlessConfig();
      const response = await fetch(`${config.url}/api/correspondents/`, {
        headers: { 'Authorization': `Token ${config.token}` }
      });

      if (response.ok) {
        const data = await response.json();
        sendResponse({ success: true, correspondents: data.results });
      } else {
        sendResponse({ success: false, error: `HTTP ${response.status}` });
      }
    } catch (error) {
      sendResponse({ success: false, error: error.message });
    }
    return true;
  }

  if (message.action === "getDocumentTypes") {
    try {
      const config = await getPaperlessConfig();
      const response = await fetch(`${config.url}/api/document_types/`, {
        headers: { 'Authorization': `Token ${config.token}` }
      });

      if (response.ok) {
        const data = await response.json();
        sendResponse({ success: true, document_types: data.results });
      } else {
        sendResponse({ success: false, error: `HTTP ${response.status}` });
      }
    } catch (error) {
      sendResponse({ success: false, error: error.message });
    }
    return true;
  }

  if (message.action === "getTags") {
    try {
      const config = await getPaperlessConfig();
      const response = await fetch(`${config.url}/api/tags/`, {
        headers: { 'Authorization': `Token ${config.token}` }
      });

      if (response.ok) {
        const data = await response.json();
        sendResponse({ success: true, tags: data.results });
      } else {
        sendResponse({ success: false, error: `HTTP ${response.status}` });
      }
    } catch (error) {
      sendResponse({ success: false, error: error.message });
    }
    return true;
  }

  // Handle email upload to Paperless-ngx
  if (message.action === "uploadEmailToPaperless") {
    console.log('📧 Background: Received uploadEmailToPaperless message');

    (async () => {
      try {
        const { messageData, selectedAttachments, uploadOptions } = message;
        console.log('📧 Background: Processing email upload for:', messageData.subject);

        const result = await uploadEmailAsPdf(messageData, selectedAttachments, uploadOptions);

        console.log('📧 Background: Email upload result:', result);
        sendResponse(result);
      } catch (error) {
        console.error("❌ Background: Error in email upload:", error);
        sendResponse({ success: false, error: error.message });
      }
    })();

    return true; // Keep the message channel open for async response
  }
});

function extractCorrespondentFromEmail(emailString) {
  const match = emailString.match(/^(.+?)\s*<.+>$/);
  return match ? match[1].trim() : emailString.split('@')[0];
}

async function getPaperlessConfig() {
  const result = await browser.storage.sync.get(['paperlessUrl', 'paperlessToken', 'defaultTags']);
  return {
    url: result.paperlessUrl?.replace(/\/$/, ''),
    token: result.paperlessToken,
    defaultTags: result.defaultTags ? result.defaultTags.split(',').map(t => t.trim()) : []
  };
}

function showNotification(message, type = "info") {
  // const iconUrl = type === "error" ? "icons/error.png" :
  //   type === "success" ? "icons/success.png" : "icons/icon-32.png";
  const iconUrl = "icons/icon-32.png";

  browser.notifications.create({
    type: "basic",
    iconUrl: iconUrl,
    title: "📄 Paperless PDF Uploader",
    message: message
  });
}

// Handle quick upload from message display popup
async function handleQuickUploadFromDisplay(messageId) {
  try {
    const message = await browser.messages.get(messageId);
    await processQuickPdfUpload(message);
  } catch (error) {
    console.error("Error handling quick upload from display:", error);
    showNotification("Error processing quick upload", "error");
  }
}

// Handle advanced upload from message display popup
async function handleAdvancedUploadFromDisplay(messageId) {
  try {
    const message = await browser.messages.get(messageId);

    // Get PDF attachments
    const attachments = await browser.messages.listAttachments(message.id);
    const pdfAttachments = attachments.filter(attachment =>
      attachment.contentType === "application/pdf" ||
      attachment.name.toLowerCase().endsWith('.pdf')
    );

    if (pdfAttachments.length === 0) {
      showNotification("No PDF attachments found in displayed message", "info");
      return;
    }

    // Store current data for the dialog
    currentMessage = message;
    currentPdfAttachments = pdfAttachments;

    // Open the advanced upload dialog
    await openAdvancedUploadDialog(message, pdfAttachments);
  } catch (error) {
    console.error("Error handling advanced upload from display:", error);
    showNotification("Error processing advanced upload", "error");
  }
}

// ============================================================
// Email Upload Functions (Email as PDF to Paperless-ngx)
// ============================================================

/**
 * Handle email upload from context menu
 */
async function handleEmailUpload(info) {
  try {
    const messages = info.selectedMessages.messages;
    if (!messages || messages.length === 0) {
      showNotification("Keine Nachrichten ausgewählt", "error");
      return;
    }

    // For now, just handle the first message
    const message = messages[0];

    // Get full message details including body
    const fullMessage = await browser.messages.get(message.id);
    
    // Get all attachments (not just PDFs)
    const attachments = await browser.messages.listAttachments(message.id);

    // Get recipients from the message
    const recipients = fullMessage.recipients || [];

    // Open the email upload dialog
    await openEmailUploadDialog(fullMessage, attachments, recipients);

  } catch (error) {
    console.error("Error handling email upload:", error);
    showNotification("Fehler beim Verarbeiten der E-Mail", "error");
  }
}

/**
 * Open the email upload dialog
 */
async function openEmailUploadDialog(message, attachments, recipients) {
  const dialogUrl = browser.runtime.getURL("email-upload-dialog.html");

  try {
    // Store data for the dialog to access
    await browser.storage.local.set({
      emailUploadData: {
        message: {
          id: message.id,
          subject: message.subject,
          author: message.author,
          date: message.date,
          recipients: recipients
        },
        attachments: attachments.map(att => ({
          name: att.name,
          partName: att.partName,
          size: att.size,
          contentType: att.contentType
        }))
      }
    });

    // Open the dialog
    browser.windows.create({
      url: dialogUrl,
      type: "popup",
      width: 600,
      height: 750
    });
  } catch (error) {
    console.error("Error opening email upload dialog:", error);
    showNotification("Fehler beim Öffnen des Dialogs", "error");
  }
}

/**
 * Generate HTML content for email
 */
function generateEmailHtmlContent(emailData) {
  const {
    subject,
    author,
    date,
    recipients,
    body,
    isHtml,
    attachments
  } = emailData;

  // Format date
  const formattedDate = new Date(date).toLocaleDateString('de-DE', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });

  // Format recipients
  const recipientList = Array.isArray(recipients) ? recipients.join(', ') : recipients || '';

  // Get attachment icons
  function getAttachmentIcon(filename, contentType) {
    const ext = (filename || '').toLowerCase().split('.').pop();
    const extensionIcons = {
      'pdf': '📄', 'doc': '📝', 'docx': '📝', 'xls': '📊', 'xlsx': '📊',
      'ppt': '📽️', 'pptx': '📽️', 'jpg': '🖼️', 'jpeg': '🖼️', 'png': '🖼️',
      'gif': '🖼️', 'bmp': '🖼️', 'webp': '🖼️', 'svg': '🖼️', 'zip': '📦',
      'rar': '📦', '7z': '📦', 'tar': '📦', 'gz': '📦', 'txt': '📃',
      'csv': '📊', 'mp3': '🎵', 'wav': '🎵', 'mp4': '🎬', 'avi': '🎬',
      'mov': '🎬', 'eml': '✉️', 'msg': '✉️'
    };
    return extensionIcons[ext] || '📎';
  }

  // Escape HTML function
  function escapeHtmlStr(str) {
    if (!str) return '';
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // Format attachments with icons
  let attachmentsHtml = '';
  if (attachments && attachments.length > 0) {
    const attachmentItems = attachments.map(att => {
      const icon = getAttachmentIcon(att.name, att.contentType);
      return `<span style="margin-right: 12px; white-space: nowrap;">${icon} ${escapeHtmlStr(att.name)}</span>`;
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
    bodyContent = body;
  } else {
    bodyContent = `<pre style="white-space: pre-wrap; word-wrap: break-word; font-family: inherit; margin: 0;">${escapeHtmlStr(body)}</pre>`;
  }

  return `<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtmlStr(subject)}</title>
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
  </style>
</head>
<body>
  <div class="email-header">
    <div class="email-header-row">
      <span class="email-header-label">Datum:</span> ${escapeHtmlStr(formattedDate)}
    </div>
    <div class="email-header-row">
      <span class="email-header-label">Von:</span> ${escapeHtmlStr(author)}
    </div>
    <div class="email-header-row">
      <span class="email-header-label">An:</span> ${escapeHtmlStr(recipientList)}
    </div>
    <div class="email-subject">
      <span class="email-header-label">Betreff:</span> <strong>${escapeHtmlStr(subject)}</strong>
    </div>
    ${attachmentsHtml ? `<div class="email-attachments">${attachmentsHtml}</div>` : ''}
  </div>
  
  <hr class="email-divider">
  
  <div class="email-body">
    ${bodyContent}
  </div>
</body>
</html>`;
}

/**
 * Upload email as HTML/PDF to Paperless-ngx
 */
async function uploadEmailAsPdf(messageData, selectedAttachments, uploadOptions) {
  try {
    const config = await getPaperlessConfig();
    if (!config.url || !config.token) {
      showNotification("Bitte konfigurieren Sie zuerst die Paperless-ngx Einstellungen", "error");
      return { success: false, error: "Paperless-ngx nicht konfiguriert" };
    }

    showNotification("E-Mail wird an Paperless-ngx gesendet...", "info");

    // Get full message content including body
    const fullMessage = await browser.messages.getFull(messageData.id);
    
    // Extract body content
    const { body, isHtml } = extractMessageBody(fullMessage);

    // Get all attachments for the header display
    const allAttachments = await browser.messages.listAttachments(messageData.id);

    // Generate HTML content for the email
    const htmlContent = generateEmailHtmlContent({
      subject: messageData.subject,
      author: messageData.author,
      date: messageData.date,
      recipients: messageData.recipients,
      body: body,
      isHtml: isHtml,
      attachments: allAttachments
    });

    // Generate filename
    const dateStr = new Date(messageData.date).toISOString().split('T')[0];
    const sanitizedSubject = (messageData.subject || 'Email')
      .replace(/[<>:"/\\|?*\x00-\x1f]/g, '')
      .replace(/\s+/g, '_')
      .substring(0, 100);
    const filename = `${dateStr}_${sanitizedSubject}.html`;

    // Create HTML file blob
    const htmlBlob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const htmlFile = new File([htmlBlob], filename, { type: 'text/html' });

    // Prepare form data for upload
    const formData = new FormData();
    formData.append('document', htmlFile, filename);

    // Add metadata
    if (uploadOptions.title) {
      formData.append('title', uploadOptions.title);
    }
    if (uploadOptions.correspondent) {
      formData.append('correspondent', uploadOptions.correspondent);
    }
    if (uploadOptions.document_type) {
      formData.append('document_type', uploadOptions.document_type);
    }
    if (uploadOptions.tags && uploadOptions.tags.length > 0) {
      uploadOptions.tags.forEach(tag => formData.append('tags', tag));
    }

    // Upload email HTML to Paperless-ngx
    const response = await fetch(`${config.url}/api/documents/post_document/`, {
      method: 'POST',
      headers: {
        'Authorization': `Token ${config.token}`
      },
      body: formData
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP ${response.status}: ${errorText}`);
    }

    const emailResult = await response.json();
    console.log("Email upload successful:", emailResult);

    // Track uploaded document IDs for linking
    const uploadedDocumentIds = [];
    if (emailResult) {
      // The task UUID is returned, we'll need to track it for linking later
      uploadedDocumentIds.push({ type: 'email', taskId: emailResult });
    }

    // Upload selected attachments
    let attachmentSuccessCount = 0;
    let attachmentErrorCount = 0;

    for (const attachment of selectedAttachments) {
      try {
        const attachmentData = await browser.messages.getAttachmentFile(
          messageData.id,
          attachment.partName
        );

        const attFormData = new FormData();
        attFormData.append('document', attachmentData, attachment.name);

        // Use attachment filename without extension as title
        attFormData.append('title', attachment.name.replace(/\.[^/.]+$/, ''));
        
        if (uploadOptions.correspondent) {
          attFormData.append('correspondent', uploadOptions.correspondent);
        }
        if (uploadOptions.document_type) {
          attFormData.append('document_type', uploadOptions.document_type);
        }
        if (uploadOptions.tags && uploadOptions.tags.length > 0) {
          uploadOptions.tags.forEach(tag => attFormData.append('tags', tag));
        }

        const attResponse = await fetch(`${config.url}/api/documents/post_document/`, {
          method: 'POST',
          headers: {
            'Authorization': `Token ${config.token}`
          },
          body: attFormData
        });

        if (attResponse.ok) {
          const attResult = await attResponse.json();
          uploadedDocumentIds.push({ type: 'attachment', taskId: attResult, name: attachment.name });
          attachmentSuccessCount++;
        } else {
          attachmentErrorCount++;
          console.error(`Failed to upload attachment ${attachment.name}`);
        }
      } catch (attError) {
        attachmentErrorCount++;
        console.error(`Error uploading attachment ${attachment.name}:`, attError);
      }
    }

    // Add Paperless tag to email in Thunderbird if requested
    if (uploadOptions.addPaperlessTag) {
      await addPaperlessTagToEmail(messageData.id);
    }

    // Show success notification
    let successMessage = "✅ E-Mail erfolgreich an Paperless-ngx gesendet";
    if (attachmentSuccessCount > 0) {
      successMessage += ` (+ ${attachmentSuccessCount} Anhang/Anhänge)`;
    }
    if (attachmentErrorCount > 0) {
      successMessage += ` (${attachmentErrorCount} Fehler)`;
    }
    showNotification(successMessage, "success");

    return { 
      success: true, 
      emailResult, 
      attachmentSuccessCount, 
      attachmentErrorCount,
      uploadedDocumentIds 
    };

  } catch (error) {
    console.error("Error uploading email to Paperless-ngx:", error);
    showNotification(`❌ Fehler beim Hochladen: ${error.message}`, "error");
    return { success: false, error: error.message };
  }
}

/**
 * Extract message body from full message
 */
function extractMessageBody(fullMessage) {
  let body = '';
  let isHtml = false;

  function findBodyPart(part) {
    if (!part) return;

    // Check if this part has body content
    if (part.body) {
      if (part.contentType && part.contentType.includes('text/html')) {
        body = part.body;
        isHtml = true;
        return true;
      } else if (part.contentType && part.contentType.includes('text/plain')) {
        // Only use plain text if we haven't found HTML yet
        if (!isHtml) {
          body = part.body;
          isHtml = false;
        }
      }
    }

    // Recursively check parts
    if (part.parts) {
      for (const subPart of part.parts) {
        if (findBodyPart(subPart)) {
          return true;
        }
      }
    }

    return false;
  }

  findBodyPart(fullMessage);

  return { body, isHtml };
}

/**
 * Add "Paperless" tag to email in Thunderbird
 */
async function addPaperlessTagToEmail(messageId) {
  try {
    // Get current message to preserve existing tags
    const message = await browser.messages.get(messageId);
    const currentTags = message.tags || [];

    // Only add if not already present
    if (!currentTags.includes(PAPERLESS_TAG_KEY)) {
      const newTags = [...currentTags, PAPERLESS_TAG_KEY];
      
      // Update message tags
      await browser.messages.update(messageId, {
        tags: newTags
      });

      console.log("Added 'Paperless' tag to email");
    }
  } catch (error) {
    console.error("Error adding Paperless tag to email:", error);
    // Don't throw - this is a non-critical operation
  }
}

/**
 * Link related documents in Paperless-ngx
 * Note: This requires custom fields support in Paperless-ngx
 */
async function linkRelatedDocuments(documentIds) {
  // This function is a placeholder for future implementation
  // Paperless-ngx doesn't have built-in related documents feature
  // This could be implemented using custom fields or tags
  console.log("Document IDs for potential linking:", documentIds);
  return { success: true, message: "Linking not implemented yet" };
}