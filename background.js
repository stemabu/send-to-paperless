// Background script for Paperless-ngx PDF Uploader
console.log("Paperless PDF Uploader loaded!");

// Configuration constants for document processing
const DOCUMENT_PROCESSING_MAX_ATTEMPTS = 60;
const DOCUMENT_PROCESSING_DELAY_MS = 1000;

// Configuration constants for Thunderbird tag
const PAPERLESS_TAG_PREFERRED_KEY = "paperless"; // desired key when creating new
const PAPERLESS_TAG_LABEL = "Paperless";
const PAPERLESS_TAG_COLOR = "#007bff";

let currentPdfAttachments = [];
let currentMessage = null;

// Reusable TextDecoder for efficient decoding of ArrayBuffer/TypedArray
const UTF8_DECODER = new TextDecoder('utf-8');

// Helper function to move From header to beginning of EML content for libmagic compatibility
// This is a workaround for libmagic not recognizing message/rfc822 when From is not at the start
// See: Paperless-ngx mail.py lines 916-933
function ensureFromHeaderAtBeginning(emlContent) {
  console.log('📧 Processing EML for libmagic compatibility');
  console.log('📧 Input type:', typeof emlContent);
  console.log('📧 Input length:', emlContent ? emlContent.length : 0);
  
  let emlString;
  if (typeof emlContent === 'string') {
    // Already a string
    emlString = emlContent;
    console.log('📧 Input was already a string');
  } else if (emlContent instanceof ArrayBuffer || ArrayBuffer.isView(emlContent)) {
    // Thunderbird 140+ returns ArrayBuffer or Uint8Array from getRaw()
    emlString = UTF8_DECODER.decode(emlContent);
    console.log('📧 Decoded ArrayBuffer/TypedArray to string');
  } else {
    // Fallback: try to convert to string
    console.warn('📧 Unexpected emlContent type:', typeof emlContent);
    emlString = String(emlContent);
  }
  
  console.log('📧 String length after conversion:', emlString.length);
  console.log('📧 First 200 chars:', emlString.substring(0, 200));
  
  const lines = emlString.split(/\r?\n/);
  
  let fromIndex = -1;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].trim().toLowerCase().startsWith('from:')) {
      fromIndex = i;
      break;
    }
  }
  
  if (fromIndex > 0) {
    console.log('📧 Moving From header to beginning (was at line ' + (fromIndex + 1) + ')');
    const fromLine = lines.splice(fromIndex, 1)[0];
    lines.unshift(fromLine);
    emlString = lines.join('\n');
  } else if (fromIndex === 0) {
    console.log('📧 From header already at beginning');
  } else {
    console.log('📧 No From header found in EML');
  }
  
  console.log('📧 Output length:', emlString.length);
  console.log('📧 Output first 200 chars:', emlString.substring(0, 200));
  
  // Return string directly - Blob constructor can handle strings
  // Do NOT use TextEncoder here - not needed for Blob constructor
  return emlString;
}

// Create context menus for attachments
browser.runtime.onInstalled.addListener(async () => {
  // Remove all existing menus first to avoid conflicts
  await browser.menus.removeAll();

  // Message list context menus
  // E-Mail mit Anhängen hochladen (first option)
  browser.menus.create({
    id: "email-to-paperless",
    title: "E-Mail mit Anhängen hochladen",
    contexts: ["message_list"],
    icons: {
      "32": "icons/icon-32.png",
      "16": "icons/icon-16.png",
      "64": "icons/icon-64.png",
      "128": "icons/icon-128.png"
    }
  });

  // Nur Anhang hochladen (second option, was "Advanced Upload")
  browser.menus.create({
    id: "upload-attachments-only",
    title: "Nur Anhang hochladen",
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
  if (info.menuItemId === "email-to-paperless") {
    await handleEmailToPaperless(info);
  } else if (info.menuItemId === "upload-attachments-only") {
    await handleAdvancedPdfUpload(info);
  }
});

async function handleAdvancedPdfUpload(info) {
  try {
    const messages = info.selectedMessages.messages;
    if (!messages || messages.length === 0) {
      showNotification("Keine Nachricht ausgewählt", "error");
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
      showNotification("Keine PDF-Anhänge in der Nachricht gefunden", "info");
      return;
    }

    // Store current data for the dialog
    currentMessage = message;
    currentPdfAttachments = pdfAttachments;

    // Open the advanced upload dialog
    await openAdvancedUploadDialog(message, pdfAttachments);

  } catch (error) {
    console.error("Error handling advanced PDF upload:", error);
    showNotification("Fehler beim Verarbeiten der Anhänge", "error");
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

    // Open the selection dialog (centered, height 1000px)
    const dialogUrl = browser.runtime.getURL("select-attachments.html");
    await createCenteredWindow(dialogUrl, 550, 1000);
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

    // Open the dialog (centered, height 1000px)
    await createCenteredWindow(dialogUrl, 550, 1000);
  } catch (error) {
    console.error("Error opening dialog:", error);
    showNotification("Error opening upload dialog", "error");
  }
}

// Handle email to Paperless-ngx upload
async function handleEmailToPaperless(info) {
  try {
    const messages = info.selectedMessages.messages;
    if (!messages || messages.length === 0) {
      showNotification("Keine Nachricht ausgewählt", "error");
      return;
    }

    // For now, just handle the first message
    const message = messages[0];
    await openEmailUploadDialog(message);

  } catch (error) {
    console.error("Error handling email to Paperless:", error);
    showNotification("Fehler beim Verarbeiten der E-Mail", "error");
  }
}

// Open email upload dialog
async function openEmailUploadDialog(message) {
  try {
    // Get all attachments (not just PDFs) - for email upload, we allow uploading
    // any attachment type since the email itself is converted to PDF
    const attachments = await browser.messages.listAttachments(message.id);

    // Get email body (full message)
    const fullMessage = await browser.messages.getFull(message.id);
    const emailBody = extractEmailBody(fullMessage);

    // Store data for the dialog to access
    await browser.storage.local.set({
      emailUploadData: {
        message: {
          id: message.id,
          subject: message.subject,
          author: message.author,
          recipients: message.recipients || [],
          date: message.date
        },
        attachments: attachments.map(att => ({
          name: att.name,
          partName: att.partName,
          size: att.size,
          contentType: att.contentType
        })),
        emailBody: emailBody
      }
    });

    // Open the email upload dialog (centered, height 1000px)
    const dialogUrl = browser.runtime.getURL("email-upload-dialog.html");
    await createCenteredWindow(dialogUrl, 550, 1000);

  } catch (error) {
    console.error("Error opening email upload dialog:", error);
    showNotification("Fehler beim Öffnen des Dialogs", "error");
  }
}

// Extract email body from full message
// Returns HTML body if available (preferred for formatting), otherwise plain text
function extractEmailBody(fullMessage) {
  let htmlBody = '';
  let plainBody = '';
  
  // Recursive function to find body parts
  function findBody(part) {
    if (part.body) {
      if (part.contentType === 'text/html') {
        htmlBody = part.body;
      } else if (part.contentType === 'text/plain' || !part.contentType) {
        plainBody = part.body;
      }
    }
    
    if (part.parts) {
      for (const subPart of part.parts) {
        findBody(subPart);
      }
    }
  }
  
  findBody(fullMessage);
  
  // Return HTML if available (preferred for formatting), otherwise plain text
  // Also return isHtml flag to indicate content type
  if (htmlBody) {
    return { body: htmlBody, isHtml: true };
  }
  return { body: plainBody, isHtml: false };
}

// Get or create custom field by name
async function getOrCreateCustomField(config, fieldName, fieldType, selectOptions = null) {
  try {
    // First, try to find existing custom field
    const response = await fetch(`${config.url}/api/custom_fields/`, {
      headers: { 'Authorization': `Token ${config.token}` }
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const data = await response.json();
    const existingField = data.results.find(f => f.name === fieldName);

    if (existingField) {
      return existingField;
    }

    // Create new custom field if not found
    const createBody = {
      name: fieldName,
      data_type: fieldType
    };

    if (selectOptions) {
      createBody.extra_data = { select_options: selectOptions };
    }

    const createResponse = await fetch(`${config.url}/api/custom_fields/`, {
      method: 'POST',
      headers: {
        'Authorization': `Token ${config.token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(createBody)
    });

    if (!createResponse.ok) {
      const errorText = await createResponse.text();
      throw new Error(`Failed to create custom field: ${errorText}`);
    }

    return await createResponse.json();

  } catch (error) {
    console.error(`Error getting/creating custom field "${fieldName}":`, error);
    throw error;
  }
}

// Get or create document type by name
async function getOrCreateDocumentType(config, typeName) {
  try {
    // First, try to find existing document type
    const response = await fetch(`${config.url}/api/document_types/`, {
      headers: { 'Authorization': `Token ${config.token}` }
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const data = await response.json();
    const existingType = data.results.find(t => t.name === typeName);

    if (existingType) {
      console.log(`📄 Found existing document type "${typeName}" with ID: ${existingType.id}`);
      return existingType;
    }

    // Create new document type if not found
    console.log(`📄 Creating new document type "${typeName}"...`);
    const createResponse = await fetch(`${config.url}/api/document_types/`, {
      method: 'POST',
      headers: {
        'Authorization': `Token ${config.token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ name: typeName })
    });

    if (!createResponse.ok) {
      const errorText = await createResponse.text();
      throw new Error(`Failed to create document type: ${errorText}`);
    }

    const newType = await createResponse.json();
    console.log(`📄 Created document type "${typeName}" with ID: ${newType.id}`);
    return newType;

  } catch (error) {
    console.error(`Error getting/creating document type "${typeName}":`, error);
    throw error;
  }
}

// Update document custom fields
async function updateDocumentCustomFields(config, documentId, customFields) {
  try {
    const response = await fetch(`${config.url}/api/documents/${documentId}/`, {
      method: 'PATCH',
      headers: {
        'Authorization': `Token ${config.token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ custom_fields: customFields })
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP ${response.status}: ${errorText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('Error updating document custom fields:', error);
    throw error;
  }
}

// Helper function to find Paperless tag in a list of tags
// Returns the tag object if found, otherwise undefined
// First checks for exact key match with PAPERLESS_TAG_PREFERRED_KEY,
// then falls back to case-insensitive label matching with PAPERLESS_TAG_LABEL
function findPaperlessTag(tags) {
  const preferredKeyLower = PAPERLESS_TAG_PREFERRED_KEY.toLowerCase();
  const labelLower = PAPERLESS_TAG_LABEL.toLowerCase();
  
  // First: check for exact key match with preferred key
  const byKey = tags.find(tag => 
    tag.key?.toLowerCase() === preferredKeyLower
  );
  if (byKey) {
    return byKey;
  }
  
  // Second: case-insensitive label/tag match
  return tags.find(tag => 
    tag.label?.toLowerCase() === labelLower ||
    tag.tag?.toLowerCase() === labelLower
  );
}

/**
 * List all Thunderbird mail tags.
 * Prefers browser.messages.listTags; fallback to browser.messages.tags.list if present.
 * @returns {Promise<Array|null>} Array of tag objects or null if API unavailable
 */
async function listAllTags() {
  try {
    if (browser.messages.listTags) {
      console.log("🏷️ Paperless-Tag: Using browser.messages.listTags()");
      return await browser.messages.listTags();
    }
    if (browser.messages.tags?.list) {
      console.log("🏷️ Paperless-Tag: Fallback to browser.messages.tags.list()");
      return await browser.messages.tags.list();
    }
    console.warn("🏷️ Paperless-Tag: No listTags API available");
    return null;
  } catch (e) {
    console.error("🏷️ Paperless-Tag: Error listing tags:", e);
    return null;
  }
}

/**
 * Create a Thunderbird mail tag.
 * Prefers browser.messages.createTag; if browser.messages.tags.create exists, 
 * pass both tag and label fields, then verify creation.
 * @param {string} key - Tag key (lowercase identifier)
 * @param {string} label - Tag display label
 * @param {string} color - Tag color (hex format)
 * @returns {Promise<boolean>} True if tag was created successfully
 */
async function createMailTag(key, label, color) {
  try {
    if (browser.messages.createTag) {
      console.log("🏷️ Paperless-Tag: Creating tag via browser.messages.createTag()");
      await browser.messages.createTag(key, label, color);
      console.log("🏷️ Paperless-Tag: Tag created via createTag()");
      return true;
    }
    
    if (browser.messages.tags?.create) {
      console.log("🏷️ Paperless-Tag: Fallback to browser.messages.tags.create()");
      // Pass both 'tag' and 'label' fields for compatibility with different
      // Thunderbird API versions: older versions use 'tag', newer use 'label'
      await browser.messages.tags.create({
        key: key,
        tag: label,
        label: label,
        color: color
      });
      console.log("🏷️ Paperless-Tag: Tag created via tags.create()");
      
      // Verify creation by re-listing tags
      const verifyTags = await listAllTags();
      if (verifyTags) {
        const found = findPaperlessTag(verifyTags);
        if (found) {
          console.log("🏷️ Paperless-Tag: Tag creation verified");
          return true;
        } else {
          console.warn("🏷️ Paperless-Tag: Tag creation could not be verified");
          return false;
        }
      }
      // Could not re-list tags for verification - report failure for safety
      console.warn("🏷️ Paperless-Tag: Could not verify tag creation (unable to list tags)");
      return false;
    }
    
    console.warn("🏷️ Paperless-Tag: No createTag API available");
    return false;
  } catch (e) {
    console.error("🏷️ Paperless-Tag: Error creating tag:", e);
    return false;
  }
}

/**
 * Ensure the Paperless tag exists in Thunderbird, creating it if absent.
 * Returns the effective tag key to use for tagging messages.
 * - If a tag with key === PAPERLESS_TAG_PREFERRED_KEY exists, returns that key.
 * - Else if a tag with label === PAPERLESS_TAG_LABEL (case-insensitive) exists, returns its key.
 * - Else creates a new tag with PAPERLESS_TAG_PREFERRED_KEY and returns that key if successful.
 * @returns {Promise<string|null>} The effective tag key or null if tag could not be ensured
 */
async function ensurePaperlessTag() {
  try {
    let tags = await listAllTags();
    
    if (!tags) {
      console.warn("🏷️ Paperless-Tag: Tags API not available");
      return null;
    }
    
    // Check if tag already exists (by preferred key or by label)
    let existingTag = findPaperlessTag(tags);
    if (existingTag) {
      console.log("🏷️ Paperless-Tag: Tag already exists with key:", existingTag.key);
      return existingTag.key;
    }
    
    // Tag doesn't exist, create it with preferred key
    console.log("🏷️ Paperless-Tag: Tag does not exist, creating with key:", PAPERLESS_TAG_PREFERRED_KEY);
    const created = await createMailTag(PAPERLESS_TAG_PREFERRED_KEY, PAPERLESS_TAG_LABEL, PAPERLESS_TAG_COLOR);
    
    if (!created) {
      console.warn("🏷️ Paperless-Tag: Failed to create tag");
      return null;
    }
    
    // Re-list tags and verify registration
    tags = await listAllTags();
    if (!tags) {
      console.warn("🏷️ Paperless-Tag: Could not re-list tags after creation");
      return null;
    }
    
    existingTag = findPaperlessTag(tags);
    if (existingTag) {
      console.log("🏷️ Paperless-Tag: Tag creation verified, key:", existingTag.key);
      return existingTag.key;
    }
    
    console.warn("🏷️ Paperless-Tag: Tag not found after creation");
    return null;
  } catch (e) {
    console.error("🏷️ Paperless-Tag: Error ensuring tag:", e);
    return null;
  }
}

/**
 * Add the Paperless tag to an email in Thunderbird.
 * Preserves existing tags on the message.
 * Re-fetches the message to verify tag assignment and logs a warning if not present.
 * Never throws.
 * @param {number} messageId - Thunderbird message ID
 */
async function addPaperlessTagToEmail(messageId) {
  try {
    console.log("🏷️ Paperless-Tag: Starting tag assignment for message", messageId);
    
    // Ensure tag exists and get the effective key
    const effectiveKey = await ensurePaperlessTag();
    if (!effectiveKey) {
      console.warn("🏷️ Paperless-Tag: Could not ensure tag exists, skipping assignment");
      return;
    }
    
    console.log("🏷️ Paperless-Tag: Using effective key:", effectiveKey);
    
    // Get the message
    const msg = await browser.messages.get(messageId);
    if (!msg) {
      console.warn("🏷️ Paperless-Tag: Message not found:", messageId);
      return;
    }
    
    // Build new tag set preserving existing tags
    const existingTags = new Set(msg.tags || []);
    if (existingTags.has(effectiveKey)) {
      console.log("🏷️ Paperless-Tag: Tag already assigned to message, skipping");
      return;
    }
    
    existingTags.add(effectiveKey);
    const newTags = Array.from(existingTags);
    
    console.log("🏷️ Paperless-Tag: Updating message tags:", newTags);
    await browser.messages.update(messageId, { tags: newTags });
    
    // Re-fetch message to verify tag was assigned
    const verifyMsg = await browser.messages.get(messageId);
    if (verifyMsg && verifyMsg.tags && verifyMsg.tags.includes(effectiveKey)) {
      console.log("🏷️ Paperless-Tag: Tag successfully assigned and verified for message", messageId);
    } else {
      console.warn("🏷️ Paperless-Tag: Tag assignment could not be verified for message", messageId);
    }
  } catch (e) {
    console.error("🏷️ Paperless-Tag: Error adding tag to message", messageId, e);
    // Never throw - just log the error
  }
}

// Upload email PDF and attachments with custom fields
async function uploadEmailWithAttachments(messageData, emailPdfData, selectedAttachments, direction, correspondent, tags, documentDate) {
  console.log('📧 Starting uploadEmailWithAttachments');
  console.log('📧 Message data:', JSON.stringify(messageData));
  console.log('📧 Email PDF filename:', emailPdfData?.filename);
  console.log('📧 Selected attachments count:', selectedAttachments?.length);
  console.log('📧 Direction:', direction);
  console.log('📧 Correspondent:', correspondent);
  console.log('📧 Tags:', tags);
  console.log('📧 Document Date:', documentDate);
  
  // Get configuration
  console.log('📧 Getting Paperless configuration...');
  const config = await getPaperlessConfig();
  
  if (!config.url || !config.token) {
    const errorMsg = "Paperless-ngx ist nicht konfiguriert. Bitte Einstellungen prüfen.";
    console.error('📧 Configuration error:', errorMsg);
    throw new Error(errorMsg);
  }
  console.log('📧 Configuration OK, URL:', config.url);

  try {
    // Get or create custom fields
    console.log('📧 Getting/creating custom fields...');
    
    let relatedDocsField;
    try {
      relatedDocsField = await getOrCreateCustomField(
        config,
        'Dazugehörende Dokumente',
        'documentlink'
      );
      console.log('📧 Related docs field:', relatedDocsField);
    } catch (fieldError) {
      console.error('📧 Error getting/creating related docs field:', fieldError);
      throw new Error('Fehler beim Erstellen des Custom Fields "Dazugehörende Dokumente": ' + fieldError.message);
    }

    let directionField;
    try {
      directionField = await getOrCreateCustomField(
        config,
        'Richtung',
        'select',
        ['Eingang', 'Ausgang', 'Intern']  // Include all possible direction options
      );
      console.log('📧 Direction field:', directionField);
      console.log('📧 Direction field extra_data:', JSON.stringify(directionField.extra_data));
    } catch (fieldError) {
      console.error('📧 Error getting/creating direction field:', fieldError);
      throw new Error('Fehler beim Erstellen des Custom Fields "Richtung": ' + fieldError.message);
    }

    // Find the option ID for the selected direction
    let directionOptionId = null;
    if (directionField.extra_data && directionField.extra_data.select_options) {
      const options = directionField.extra_data.select_options;
      console.log('📧 Available direction options:', JSON.stringify(options));
      
      // Find the option that matches the selected direction (with trimming for safety)
      const directionTrimmed = direction.trim();
      const matchingOption = options.find(opt => opt.label && opt.label.trim() === directionTrimmed);
      if (matchingOption) {
        directionOptionId = matchingOption.id;
        console.log(`📧 Found direction option ID for "${direction}": ${directionOptionId}`);
      } else {
        console.error(`📧 Could not find option ID for direction: "${direction}"`);
        console.error(`📧 Available options:`, options.map(o => o.label).join(', '));
      }
    }

    // Prepare direction custom field value with the option ID (not the label!)
    // Note: Paperless-ngx select fields require a single string value, not an array
    const directionCustomField = directionOptionId ? {
      field: directionField.id,
      value: String(directionOptionId)  // Use the ID as a string, not an array!
    } : null;
    
    if (!directionCustomField) {
      console.warn('📧 ⚠️ Could not set direction custom field - option ID not found');
    }

    // Get or create document types for E-Mail and E-Mail-Anhang
    let emailDocumentType;
    let attachmentDocumentType;
    
    try {
      emailDocumentType = await getOrCreateDocumentType(config, 'E-Mail');
      console.log(`📧 Setting document type to: ${emailDocumentType.id} (E-Mail)`);
    } catch (typeError) {
      console.error('📧 Error getting/creating E-Mail document type:', typeError);
      // Continue without document type - not critical
    }
    
    try {
      attachmentDocumentType = await getOrCreateDocumentType(config, 'E-Mail-Anhang');
      console.log(`📧 Setting document type to: ${attachmentDocumentType.id} (E-Mail-Anhang)`);
    } catch (typeError) {
      console.error('📧 Error getting/creating E-Mail-Anhang document type:', typeError);
      // Continue without document type - not critical
    }

    // Upload email PDF
    console.log('📧 Starting email PDF upload...');
    
    // Convert base64 back to blob
    let emailPdfBlob;
    try {
      console.log('📧 Converting base64 to blob, data length:', emailPdfData?.blob?.length);
      emailPdfBlob = base64ToBlob(emailPdfData.blob, 'application/pdf');
      console.log('📧 Email PDF blob created, size:', emailPdfBlob.size);
    } catch (blobError) {
      console.error('📧 Error converting PDF to blob:', blobError);
      throw new Error('Fehler beim Konvertieren der PDF-Daten: ' + blobError.message);
    }
    
    const emailFormData = new FormData();
    emailFormData.append('document', emailPdfBlob, emailPdfData.filename);
    emailFormData.append('title', emailPdfData.filename.replace(/\.pdf$/i, ''));
    
    // Add document type "E-Mail" (using ID)
    if (emailDocumentType && emailDocumentType.id) {
      emailFormData.append('document_type', emailDocumentType.id);
    }
    
    // Add created date (email date)
    if (documentDate) {
      emailFormData.append('created', documentDate);
    }
    
    // Add correspondent if selected
    if (correspondent) {
      emailFormData.append('correspondent', correspondent);
    }
    
    // Add tags if selected
    if (tags && tags.length > 0) {
      tags.forEach(tagId => {
        emailFormData.append('tags', tagId);
      });
    }

    console.log('📧 Sending email PDF to Paperless...');
    const emailUploadResponse = await fetch(`${config.url}/api/documents/post_document/`, {
      method: 'POST',
      headers: {
        'Authorization': `Token ${config.token}`
      },
      body: emailFormData
    });

    console.log('📧 Email upload response status:', emailUploadResponse.status);
    
    if (!emailUploadResponse.ok) {
      const errorText = await emailUploadResponse.text();
      console.error('📧 Email upload failed:', errorText);
      throw new Error(`E-Mail-Upload fehlgeschlagen (HTTP ${emailUploadResponse.status}): ${errorText}`);
    }

    // Email PDF was accepted by Paperless-ngx - add tag to Thunderbird email
    console.log('🏷️ Paperless-Tag: E-Mail-Upload erfolgreich, füge Tag hinzu...');
    addPaperlessTagToEmail(messageData.id).catch(e =>
      console.warn("🏷️ Paperless-Tag: Fehler beim Taggen der E-Mail:", e)
    );

    // Get the task ID from the response
    const emailTaskId = await emailUploadResponse.text();
    console.log('📧 Email upload task ID:', emailTaskId);

    // Wait for document to be processed and get the document ID
    console.log('📧 Waiting for email document to be processed...');
    const emailDocId = await waitForDocumentId(config, emailTaskId.replace(/"/g, ''));
    console.log('📧 Email document ID:', emailDocId);

    if (!emailDocId) {
      console.warn('📧 ⚠️ Email document ID not found after waiting');
      console.warn('📧 Document was uploaded but Paperless is still processing it');
      
      // Return success with warning instead of throwing error
      return {
        success: true,
        emailDocId: null,
        attachmentDocIds: [],
        warning: 'Das Dokument wurde erfolgreich hochgeladen, aber Paperless-ngx verarbeitet es noch. Bitte pruefen Sie in einigen Minuten im Paperless-System.'
      };
    }

    // Upload selected attachments
    console.log('📧 Starting attachment uploads, count:', selectedAttachments?.length || 0);
    const attachmentDocIds = [];
    const attachmentErrors = [];
    
    for (const attachment of selectedAttachments || []) {
      console.log('📎 Processing attachment:', attachment.name, 'partName:', attachment.partName);
      
      // Get attachment file
      let attachmentFile;
      try {
        console.log('📎 Getting attachment file from Thunderbird...');
        attachmentFile = await browser.messages.getAttachmentFile(
          messageData.id,
          attachment.partName
        );
        console.log('📎 Attachment file retrieved, size:', attachmentFile?.size);
        
        if (!attachmentFile || attachmentFile.size === 0) {
          throw new Error('Anhang-Datei ist leer oder konnte nicht gelesen werden');
        }
      } catch (fileError) {
        console.error('📎 Error getting attachment file:', fileError);
        attachmentErrors.push(`${attachment.name}: Datei konnte nicht gelesen werden - ${fileError.message}`);
        continue;
      }

      const attachmentFormData = new FormData();
      attachmentFormData.append('document', attachmentFile, attachment.name);
      attachmentFormData.append('title', attachment.name.replace(/\.[^/.]+$/, ''));
      
      // Add document type "E-Mail-Anhang" (using ID)
      if (attachmentDocumentType && attachmentDocumentType.id) {
        attachmentFormData.append('document_type', attachmentDocumentType.id);
      }
      
      // Add created date (same as email date)
      if (documentDate) {
        attachmentFormData.append('created', documentDate);
      }
      
      // Add correspondent if selected
      if (correspondent) {
        attachmentFormData.append('correspondent', correspondent);
      }
      
      // Add tags if selected
      if (tags && tags.length > 0) {
        tags.forEach(tagId => {
          attachmentFormData.append('tags', tagId);
        });
      }

      try {
        console.log('📎 Uploading attachment to Paperless...');
        const attachmentResponse = await fetch(`${config.url}/api/documents/post_document/`, {
          method: 'POST',
          headers: {
            'Authorization': `Token ${config.token}`
          },
          body: attachmentFormData
        });

        console.log('📎 Attachment upload response status:', attachmentResponse.status);
        
        if (!attachmentResponse.ok) {
          const errorText = await attachmentResponse.text();
          console.error('📎 Attachment upload failed:', attachment.name, errorText);
          attachmentErrors.push(`${attachment.name}: Upload fehlgeschlagen (HTTP ${attachmentResponse.status})`);
          continue;
        }

        const attachmentTaskId = await attachmentResponse.text();
        console.log('📎 Attachment task ID:', attachmentTaskId);
        
        console.log('📎 Waiting for attachment document to be processed...');
        const attachmentDocId = await waitForDocumentId(config, attachmentTaskId.replace(/"/g, ''));
        console.log('📎 Attachment document ID:', attachmentDocId);

        if (attachmentDocId) {
          attachmentDocIds.push(attachmentDocId);
          console.log('📎 Attachment successfully uploaded:', attachment.name);
        } else {
          console.warn('📎 Attachment document ID not found:', attachment.name);
          attachmentErrors.push(`${attachment.name}: Dokument-ID nicht gefunden nach Upload`);
        }
      } catch (uploadError) {
        console.error('📎 Error uploading attachment:', attachment.name, uploadError);
        attachmentErrors.push(`${attachment.name}: ${uploadError.message}`);
      }
    }

    console.log('📧 Attachment upload complete. Success:', attachmentDocIds.length, 'Errors:', attachmentErrors.length);

    // Update custom fields for all documents
    console.log('📧 Updating custom fields...');

    // Update email document with links to attachments and direction
    try {
      const emailCustomFields = [];
      
      // Add direction if we found the option ID
      if (directionCustomField) {
        emailCustomFields.push(directionCustomField);
      }
      
      // Add related documents if any attachments were uploaded
      if (attachmentDocIds.length > 0) {
        emailCustomFields.push({
          field: relatedDocsField.id,
          value: attachmentDocIds
        });
      }
      
      if (emailCustomFields.length > 0) {
        console.log('📧 Setting email custom fields:', JSON.stringify(emailCustomFields));
        await updateDocumentCustomFields(config, emailDocId, emailCustomFields);
        console.log('📧 Email custom fields updated');
      } else {
        console.log('📧 No custom fields to set');
      }
    } catch (cfError) {
      console.error('📧 Error updating email custom fields:', cfError);
      // Don't throw - custom fields are not critical
    }

    // Update each attachment with link to email and direction
    for (const attachmentDocId of attachmentDocIds) {
      try {
        const attachmentCustomFields = [];
        
        // Add direction if available
        if (directionCustomField) {
          attachmentCustomFields.push(directionCustomField);
        }
        
        // Add link to email document
        attachmentCustomFields.push({
          field: relatedDocsField.id,
          value: [emailDocId]
        });
        
        if (attachmentCustomFields.length > 0) {
          console.log('📧 Setting attachment custom fields for doc', attachmentDocId);
          await updateDocumentCustomFields(config, attachmentDocId, attachmentCustomFields);
        }
      } catch (cfError) {
        console.error('📧 Error updating attachment custom fields:', cfError);
        // Don't throw - custom fields are not critical
      }
    }

    const totalDocs = 1 + attachmentDocIds.length;
    console.log('📧 Upload complete. Total documents:', totalDocs);
    
    let successMessage = `✅ ${totalDocs} Dokument(e) erfolgreich hochgeladen!`;
    if (attachmentErrors.length > 0) {
      successMessage += ` (${attachmentErrors.length} Anhang-Fehler)`;
    }
    showNotification(successMessage, "success");

    return {
      success: true,
      emailDocId: emailDocId,
      attachmentDocIds: attachmentDocIds,
      attachmentErrors: attachmentErrors.length > 0 ? attachmentErrors : undefined
    };

  } catch (error) {
    console.error("📧 ❌ Error uploading email with attachments:", error);
    console.error("📧 Error name:", error.name);
    console.error("📧 Error message:", error.message);
    console.error("📧 Error stack:", error.stack);
    
    // Return error object instead of throwing
    return {
      success: false,
      error: error.message || 'Unbekannter Fehler beim Hochladen',
      errorDetails: {
        name: error.name,
        message: error.message,
        stack: error.stack
      }
    };
  }
}

// Helper function for file size formatting
// Returns safe numeric strings only (no user input is used in the output)
function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// Escape HTML to prevent XSS
function escapeHtml(text) {
  if (!text) return '';
  const div = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };
  return String(text).replace(/[&<>"']/g, c => div[c]);
}

// Sanitize HTML content for safe rendering in Gotenberg
// Removes dangerous elements and attributes using regex-based approach
// (background scripts don't have DOM access for DOM-based sanitization)
//
// Security context (why regex sanitization is acceptable here):
// 1. Input source: Email HTML comes from Thunderbird's message parsing API, not arbitrary user input
// 2. Sandboxed execution: Gotenberg runs in a Docker container with Chromium in headless mode
// 3. Output is static: The result is a PDF file - scripts cannot execute in PDFs
// 4. Defense in depth: This sanitization is an additional layer, not the only security measure
//
// Note: CodeQL may flag regex-based sanitization as incomplete (js/incomplete-multi-character-sanitization).
// This is acknowledged - regex cannot catch all XSS variations, but the security context above makes this acceptable.
function sanitizeHtmlForGotenberg(html) {
  if (!html) return '';
  
  console.log('📧 Sanitizing HTML for Gotenberg, original length:', html.length);
  
  let sanitized = html;
  let previousLength;
  
  // Apply multiple passes to handle nested/repeated dangerous content
  // Continue until the content stabilizes (no more changes)
  do {
    previousLength = sanitized.length;
    
    // 1. Remove script tags and their content (handles various variations)
    // Use [\s\S] instead of . to match across newlines
    sanitized = sanitized.replace(/<script\b[^>]*>[\s\S]*?<\/script\s*>/gi, '');
    // Remove unclosed/malformed script tags
    sanitized = sanitized.replace(/<script\b[^>]*>/gi, '');
    sanitized = sanitized.replace(/<\/script\s*>/gi, '');
    
    // 2. Remove style tags (we'll add our own styles)
    sanitized = sanitized.replace(/<style\b[^>]*>[\s\S]*?<\/style\s*>/gi, '');
    sanitized = sanitized.replace(/<style\b[^>]*>/gi, '');
    sanitized = sanitized.replace(/<\/style\s*>/gi, '');
    
    // 3. Remove dangerous tags
    const dangerousTags = ['iframe', 'frame', 'frameset', 'object', 'embed', 'applet', 'form', 'input', 'button', 'select', 'textarea', 'link', 'meta', 'base', 'svg', 'math'];
    dangerousTags.forEach(tag => {
      // Remove opening tags with any content up to closing tag
      sanitized = sanitized.replace(new RegExp(`<${tag}\\b[^>]*>[\\s\\S]*?<\\/${tag}\\s*>`, 'gi'), '');
      // Remove self-closing and unclosed tags
      sanitized = sanitized.replace(new RegExp(`<${tag}\\b[^>]*\\/?>`, 'gi'), '');
      // Remove orphan closing tags
      sanitized = sanitized.replace(new RegExp(`<\\/${tag}\\s*>`, 'gi'), '');
    });
    
    // 4. Remove event handlers (on* attributes)
    // Handle quoted values with double quotes
    sanitized = sanitized.replace(/\s+on\w+\s*=\s*"[^"]*"/gi, '');
    // Handle quoted values with single quotes
    sanitized = sanitized.replace(/\s+on\w+\s*=\s*'[^']*'/gi, '');
    // Handle unquoted values (more aggressive pattern)
    sanitized = sanitized.replace(/\s+on\w+\s*=\s*[^\s>"']+/gi, '');
    
    // 5. Remove javascript: and vbscript: URLs in href and src attributes
    sanitized = sanitized.replace(/href\s*=\s*"[^"]*javascript:[^"]*"/gi, 'href="#"');
    sanitized = sanitized.replace(/href\s*=\s*'[^']*javascript:[^']*'/gi, "href='#'");
    sanitized = sanitized.replace(/href\s*=\s*"[^"]*vbscript:[^"]*"/gi, 'href="#"');
    sanitized = sanitized.replace(/href\s*=\s*'[^']*vbscript:[^']*'/gi, "href='#'");
    sanitized = sanitized.replace(/src\s*=\s*"[^"]*javascript:[^"]*"/gi, 'src=""');
    sanitized = sanitized.replace(/src\s*=\s*'[^']*javascript:[^']*'/gi, "src=''");
    sanitized = sanitized.replace(/src\s*=\s*"[^"]*vbscript:[^"]*"/gi, 'src=""');
    sanitized = sanitized.replace(/src\s*=\s*'[^']*vbscript:[^']*'/gi, "src=''");
    
    // 6. Remove data: URLs in src attributes (can contain embedded scripts)
    sanitized = sanitized.replace(/src\s*=\s*"[^"]*data:[^"]*"/gi, 'src=""');
    sanitized = sanitized.replace(/src\s*=\s*'[^']*data:[^']*'/gi, "src=''");
    
  } while (sanitized.length !== previousLength);
  
  console.log('📧 Sanitized HTML length:', sanitized.length);
  
  return sanitized;
}

// Create HTML template for email (for Gotenberg conversion)
// Based on Paperless-ngx email_msg_template.html but simplified with inline CSS
function createEmailHtml(messageData, emailBodyData, selectedAttachments) {
  const dateStr = new Date(messageData.date).toLocaleString('de-DE', {
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit'
  });
  
  const toRecipients = (messageData.recipients || []).join(', ');
  
  // Prepare content - HTML or plain text
  let contentHtml;
  if (emailBodyData.isHtml && emailBodyData.body) {
    // Sanitize HTML content to remove dangerous elements before rendering
    // This protects against malicious HTML/JavaScript in email bodies
    contentHtml = sanitizeHtmlForGotenberg(emailBodyData.body);
  } else {
    // Plain text - escape and preserve whitespace
    contentHtml = `<pre style="white-space: pre-wrap; word-wrap: break-word; font-family: inherit; margin: 0;">${escapeHtml(emailBodyData.body || '')}</pre>`;
  }
  
  // Build attachments section if any
  let attachmentsSection = '';
  if (selectedAttachments && selectedAttachments.length > 0) {
    const attachmentList = selectedAttachments.map(att => 
      `${escapeHtml(att.name)} (${formatFileSize(att.size)})`
    ).join(', ');
    attachmentsSection = `
      <div class="header-row">
        <span class="header-label">Anhänge:</span>
        <span class="header-value attachments-list">${attachmentList}</span>
      </div>
    `;
  }
  
  return `<!doctype html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
      margin: 20px;
      background: white;
      color: #1a1a1a;
      font-size: 14px;
      line-height: 1.5;
    }
    .container {
      max-width: 800px;
      margin: 0 auto;
    }
    .header {
      background: #e2e8f0;
      padding: 20px;
      margin-bottom: 20px;
      border-radius: 4px;
    }
    .header-row {
      margin: 8px 0;
      display: flex;
    }
    .header-label {
      color: #64748b;
      min-width: 80px;
      text-align: right;
      padding-right: 12px;
      font-weight: 500;
    }
    .header-value {
      flex: 1;
      word-break: break-word;
    }
    .subject {
      font-weight: bold;
    }
    .date {
      color: #64748b;
      float: right;
      font-size: 13px;
    }
    .separator {
      border-top: 1px solid #cbd5e1;
      margin: 20px 0;
    }
    .content {
      word-wrap: break-word;
      overflow-wrap: break-word;
      line-height: 1.6;
    }
    .attachments-list {
      color: #475569;
      font-size: 13px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="date">${escapeHtml(dateStr)}</div>
      <div class="header-row">
        <span class="header-label">Von:</span>
        <span class="header-value">${escapeHtml(messageData.author || '')}</span>
      </div>
      <div class="header-row">
        <span class="header-label">Betreff:</span>
        <span class="header-value subject">${escapeHtml(messageData.subject || 'Kein Betreff')}</span>
      </div>
      <div class="header-row">
        <span class="header-label">An:</span>
        <span class="header-value">${escapeHtml(toRecipients)}</span>
      </div>
      ${attachmentsSection}
    </div>
    <div class="separator"></div>
    <div class="content">
      ${contentHtml}
    </div>
  </div>
</body>
</html>`;
}

// Convert email to PDF via Gotenberg HTTP API
// Gotenberg provides HTML to PDF conversion via Chromium
async function convertEmailToPdfViaGotenberg(messageData, emailBodyData, selectedAttachments, gotenbergUrl) {
  console.log('📧 Converting email to PDF via Gotenberg:', gotenbergUrl);
  
  // Create HTML content
  const htmlContent = createEmailHtml(messageData, emailBodyData, selectedAttachments);
  console.log('📧 HTML content created, length:', htmlContent.length);
  
  // Create FormData for Gotenberg
  const formData = new FormData();
  const htmlFile = new File([htmlContent], 'index.html', { type: 'text/html' });
  formData.append('files', htmlFile);
  
  // POST to Gotenberg's Chromium HTML conversion endpoint
  // This endpoint is the standard Gotenberg API path for HTML to PDF conversion
  // Compatible with Gotenberg v7+ (https://gotenberg.dev/docs/routes#html-file-into-pdf-route)
  const gotenbergEndpoint = `${gotenbergUrl}/forms/chromium/convert/html`;
  console.log('📧 Calling Gotenberg endpoint:', gotenbergEndpoint);
  
  const response = await fetch(gotenbergEndpoint, {
    method: 'POST',
    body: formData
  });
  
  if (!response.ok) {
    const errorText = await response.text();
    console.error('📧 Gotenberg conversion failed:', response.status, errorText);
    throw new Error(`Gotenberg-Konvertierung fehlgeschlagen (HTTP ${response.status}): ${errorText}`);
  }
  
  const pdfBlob = await response.blob();
  console.log('📧 Gotenberg PDF received, size:', pdfBlob.size);
  
  return pdfBlob;
}

// Upload email as HTML file for Paperless Gotenberg conversion (better character encoding)
// Now uses direct Gotenberg API call instead of relying on Paperless internal Gotenberg
async function uploadEmailAsHtml(messageData, selectedAttachments, direction, correspondent, tags, documentDate) {
  console.log('📧 Starting HTML upload (Direct Gotenberg)');
  console.log('📧 Message data:', JSON.stringify(messageData));
  console.log('📧 Direction:', direction);
  console.log('📧 Correspondent:', correspondent);
  console.log('📧 Tags:', tags);
  console.log('📧 Document Date:', documentDate);
  console.log('📧 Selected attachments count:', selectedAttachments?.length);
  
  // Get Gotenberg URL from settings
  const gotenbergResult = await browser.storage.sync.get(['gotenbergUrl']);
  const gotenbergUrl = gotenbergResult.gotenbergUrl;
  
  if (!gotenbergUrl) {
    console.error('📧 Gotenberg URL not configured');
    return {
      success: false,
      error: 'Gotenberg URL nicht konfiguriert. Bitte in den Einstellungen angeben.'
    };
  }
  console.log('📧 Gotenberg URL:', gotenbergUrl);
  
  // Get Paperless configuration
  console.log('📧 Getting Paperless configuration...');
  const config = await getPaperlessConfig();
  
  if (!config.url || !config.token) {
    const errorMsg = "Paperless-ngx ist nicht konfiguriert. Bitte Einstellungen prüfen.";
    console.error('📧 Configuration error:', errorMsg);
    return {
      success: false,
      error: errorMsg
    };
  }
  console.log('📧 Configuration OK, URL:', config.url);

  try {
    // Get or create custom fields
    console.log('📧 Getting/creating custom fields...');
    
    let relatedDocsField;
    try {
      relatedDocsField = await getOrCreateCustomField(
        config,
        'Dazugehörende Dokumente',
        'documentlink'
      );
      console.log('📧 Related docs field:', relatedDocsField);
    } catch (fieldError) {
      console.error('📧 Error getting/creating related docs field:', fieldError);
      throw new Error('Fehler beim Erstellen des Custom Fields "Dazugehörende Dokumente": ' + fieldError.message);
    }

    let directionField;
    try {
      directionField = await getOrCreateCustomField(
        config,
        'Richtung',
        'select',
        ['Eingang', 'Ausgang', 'Intern']
      );
      console.log('📧 Direction field:', directionField);
    } catch (fieldError) {
      console.error('📧 Error getting/creating direction field:', fieldError);
      throw new Error('Fehler beim Erstellen des Custom Fields "Richtung": ' + fieldError.message);
    }

    // Find the option ID for the selected direction
    let directionOptionId = null;
    if (directionField.extra_data && directionField.extra_data.select_options) {
      const options = directionField.extra_data.select_options;
      const directionTrimmed = direction.trim();
      const matchingOption = options.find(opt => opt.label && opt.label.trim() === directionTrimmed);
      if (matchingOption) {
        directionOptionId = matchingOption.id;
        console.log(`📧 Found direction option ID for "${direction}": ${directionOptionId}`);
      }
    }

    // Get or create document types
    let emailDocumentType;
    let attachmentDocumentType;
    
    try {
      emailDocumentType = await getOrCreateDocumentType(config, 'E-Mail');
      console.log(`📧 Setting document type to: ${emailDocumentType.id} (E-Mail)`);
    } catch (typeError) {
      console.error('📧 Error getting/creating E-Mail document type:', typeError);
    }
    
    try {
      attachmentDocumentType = await getOrCreateDocumentType(config, 'E-Mail-Anhang');
      console.log(`📧 Setting document type to: ${attachmentDocumentType.id} (E-Mail-Anhang)`);
    } catch (typeError) {
      console.error('📧 Error getting/creating E-Mail-Anhang document type:', typeError);
    }

    // Get email body from Thunderbird
    console.log('📧 Getting email body from Thunderbird...');
    const fullMessage = await browser.messages.getFull(messageData.id);
    const emailBodyData = extractEmailBody(fullMessage);
    console.log('📧 Email body extracted, isHtml:', emailBodyData.isHtml, 'length:', emailBodyData.body?.length);

    // Convert email to PDF via Gotenberg
    console.log('📧 Converting email to PDF via Gotenberg...');
    let pdfBlob;
    try {
      pdfBlob = await convertEmailToPdfViaGotenberg(messageData, emailBodyData, selectedAttachments, gotenbergUrl);
    } catch (gotenbergError) {
      console.error('📧 Gotenberg conversion failed:', gotenbergError);
      return {
        success: false,
        error: `Gotenberg-Konvertierung fehlgeschlagen: ${gotenbergError.message}`
      };
    }

    // Create filename for the PDF
    const fileDateStr = documentDate || new Date().toISOString().split('T')[0];
    const safeSubject = (messageData.subject || 'Kein_Betreff')
      .replace(/[^a-zA-Z0-9äöüÄÖÜß\s-]/g, '')
      .replace(/\s+/g, '_')
      .substring(0, 50);
    const pdfFilename = `${fileDateStr}_${safeSubject}.pdf`;
    console.log('📧 PDF filename:', pdfFilename);

    // Create PDF file for FormData
    const pdfFile = new File([pdfBlob], pdfFilename, { type: 'application/pdf' });
    console.log('📧 PDF file size:', pdfFile.size);

    // Upload PDF to Paperless
    const pdfFormData = new FormData();
    pdfFormData.append('document', pdfFile);
    pdfFormData.append('title', safeSubject);
    
    if (emailDocumentType && emailDocumentType.id) {
      pdfFormData.append('document_type', emailDocumentType.id);
    }
    if (documentDate) {
      pdfFormData.append('created', documentDate);
    }
    if (correspondent) {
      pdfFormData.append('correspondent', correspondent);
    }
    if (tags && tags.length > 0) {
      tags.forEach(tagId => pdfFormData.append('tags', tagId));
    }

    console.log('📧 Uploading PDF to Paperless...');
    const pdfResponse = await fetch(`${config.url}/api/documents/post_document/`, {
      method: 'POST',
      headers: { 'Authorization': `Token ${config.token}` },
      body: pdfFormData
    });

    if (!pdfResponse.ok) {
      const errorText = await pdfResponse.text();
      console.error('📧 PDF upload failed:', errorText);
      throw new Error(`Upload failed (${pdfResponse.status}): ${errorText}`);
    }

    // Add tag to Thunderbird email
    console.log('🏷️ Paperless-Tag: PDF-Upload erfolgreich, füge Tag hinzu...');
    addPaperlessTagToEmail(messageData.id).catch(e =>
      console.warn("🏷️ Paperless-Tag: Fehler beim Taggen der E-Mail:", e)
    );

    // Wait for document processing
    const pdfTaskId = await pdfResponse.text();
    console.log('📧 PDF upload task ID:', pdfTaskId);
    console.log('📧 Waiting for Paperless to process...');
    const emailDocId = await waitForDocumentId(config, pdfTaskId.replace(/"/g, ''));
    console.log('📧 Email document ID:', emailDocId);

    if (!emailDocId) {
      console.warn('📧 ⚠️ Email document ID not found after waiting');
      return {
        success: true,
        warning: 'Dokument hochgeladen, wird noch verarbeitet. Bitte später im Paperless-System prüfen.'
      };
    }

    // Upload attachments
    console.log('📧 Starting attachment uploads, count:', selectedAttachments?.length || 0);
    const attachmentDocIds = [];
    const attachmentErrors = [];
    
    for (const attachment of selectedAttachments || []) {
      console.log('📎 Processing attachment:', attachment.name);
      
      try {
        const attachmentFile = await browser.messages.getAttachmentFile(
          messageData.id,
          attachment.partName
        );
        
        if (!attachmentFile || attachmentFile.size === 0) {
          throw new Error(`Anhang '${attachment.name}' ist leer oder konnte nicht gelesen werden`);
        }

        const attachmentFormData = new FormData();
        attachmentFormData.append('document', attachmentFile, attachment.name);
        attachmentFormData.append('title', attachment.name.replace(/\.[^/.]+$/, ''));
        
        if (attachmentDocumentType && attachmentDocumentType.id) {
          attachmentFormData.append('document_type', attachmentDocumentType.id);
        }
        if (documentDate) {
          attachmentFormData.append('created', documentDate);
        }
        if (correspondent) {
          attachmentFormData.append('correspondent', correspondent);
        }
        if (tags && tags.length > 0) {
          tags.forEach(tagId => attachmentFormData.append('tags', tagId));
        }

        const attachmentResponse = await fetch(`${config.url}/api/documents/post_document/`, {
          method: 'POST',
          headers: { 'Authorization': `Token ${config.token}` },
          body: attachmentFormData
        });

        if (!attachmentResponse.ok) {
          attachmentErrors.push(`${attachment.name}: Upload failed`);
          continue;
        }

        const attachmentTaskId = await attachmentResponse.text();
        const attachmentDocId = await waitForDocumentId(config, attachmentTaskId.replace(/"/g, ''));

        if (attachmentDocId) {
          attachmentDocIds.push(attachmentDocId);
          console.log('📎 Attachment successfully uploaded:', attachment.name);
        } else {
          attachmentErrors.push(`${attachment.name}: ID nicht gefunden`);
        }
      } catch (error) {
        console.error('📎 Error uploading attachment:', attachment.name, error);
        attachmentErrors.push(`${attachment.name}: ${error.message}`);
      }
    }

    // Update custom fields
    console.log('📧 Updating custom fields...');
    
    try {
      const emailCustomFields = [];
      if (directionOptionId) {
        emailCustomFields.push({ field: directionField.id, value: String(directionOptionId) });
      }
      if (attachmentDocIds.length > 0) {
        emailCustomFields.push({ field: relatedDocsField.id, value: attachmentDocIds });
      }
      if (emailCustomFields.length > 0) {
        await updateDocumentCustomFields(config, emailDocId, emailCustomFields);
        console.log('📧 Email custom fields updated');
      }
    } catch (error) {
      console.error('📧 Custom field update error:', error);
    }

    // Update attachment custom fields
    for (const attachmentDocId of attachmentDocIds) {
      try {
        const attachmentCustomFields = [];
        if (directionOptionId) {
          attachmentCustomFields.push({ field: directionField.id, value: String(directionOptionId) });
        }
        attachmentCustomFields.push({ field: relatedDocsField.id, value: [emailDocId] });
        await updateDocumentCustomFields(config, attachmentDocId, attachmentCustomFields);
      } catch (error) {
        console.error('📧 Attachment field update error:', error);
      }
    }

    const totalDocs = 1 + attachmentDocIds.length;
    console.log('📧 Gotenberg upload complete. Total documents:', totalDocs);
    showNotification(`✅ ${totalDocs} Dokument(e) hochgeladen (via Gotenberg)!`, "success");

    return {
      success: true,
      emailDocId: emailDocId,
      attachmentDocIds: attachmentDocIds,
      attachmentErrors: attachmentErrors.length > 0 ? attachmentErrors : undefined,
      strategy: 'gotenberg'
    };

  } catch (error) {
    console.error("📧 ❌ Gotenberg upload error:", error);
    return {
      success: false,
      error: error.message || 'Unbekannter Fehler'
    };
  }
}

// Upload email as .eml file for Paperless-ngx with libmagic compatibility
// Uses the From-header-first workaround for correct MIME type detection
async function uploadEmailAsEml(messageData, selectedAttachments, direction, correspondent, tags, documentDate) {
  console.log('📧 Starting EML upload (native format)');
  console.log('📧 Message data:', JSON.stringify(messageData));
  console.log('📧 Direction:', direction);
  console.log('📧 Correspondent:', correspondent);
  console.log('📧 Tags:', tags);
  console.log('📧 Document Date:', documentDate);
  console.log('📧 Selected attachments count:', selectedAttachments?.length);
  
  // Get configuration
  console.log('📧 Getting Paperless configuration...');
  const config = await getPaperlessConfig();
  
  if (!config.url || !config.token) {
    const errorMsg = "Paperless-ngx ist nicht konfiguriert. Bitte Einstellungen prüfen.";
    console.error('📧 Configuration error:', errorMsg);
    return {
      success: false,
      error: errorMsg
    };
  }
  console.log('📧 Configuration OK, URL:', config.url);

  try {
    // Get or create custom fields
    console.log('📧 Getting/creating custom fields...');
    
    let relatedDocsField;
    try {
      relatedDocsField = await getOrCreateCustomField(
        config,
        'Dazugehörende Dokumente',
        'documentlink'
      );
      console.log('📧 Related docs field:', relatedDocsField);
    } catch (fieldError) {
      console.error('📧 Error getting/creating related docs field:', fieldError);
      throw new Error('Fehler beim Erstellen des Custom Fields "Dazugehörende Dokumente": ' + fieldError.message);
    }

    let directionField;
    try {
      directionField = await getOrCreateCustomField(
        config,
        'Richtung',
        'select',
        ['Eingang', 'Ausgang', 'Intern']
      );
      console.log('📧 Direction field:', directionField);
    } catch (fieldError) {
      console.error('📧 Error getting/creating direction field:', fieldError);
      throw new Error('Fehler beim Erstellen des Custom Fields "Richtung": ' + fieldError.message);
    }

    // Find the option ID for the selected direction
    let directionOptionId = null;
    if (directionField.extra_data && directionField.extra_data.select_options) {
      const options = directionField.extra_data.select_options;
      const directionTrimmed = direction.trim();
      const matchingOption = options.find(opt => opt.label && opt.label.trim() === directionTrimmed);
      if (matchingOption) {
        directionOptionId = matchingOption.id;
        console.log(`📧 Found direction option ID for "${direction}": ${directionOptionId}`);
      }
    }

    // Get or create document types
    let emailDocumentType;
    let attachmentDocumentType;
    
    try {
      emailDocumentType = await getOrCreateDocumentType(config, 'E-Mail');
      console.log(`📧 Setting document type to: ${emailDocumentType.id} (E-Mail)`);
    } catch (typeError) {
      console.error('📧 Error getting/creating E-Mail document type:', typeError);
    }
    
    try {
      attachmentDocumentType = await getOrCreateDocumentType(config, 'E-Mail-Anhang');
      console.log(`📧 Setting document type to: ${attachmentDocumentType.id} (E-Mail-Anhang)`);
    } catch (typeError) {
      console.error('📧 Error getting/creating E-Mail-Anhang document type:', typeError);
    }

    // Get .eml file from Thunderbird - always fetch fresh without caching
    console.log('📧 Getting .eml file from Thunderbird for message ID:', messageData.id);
    let emlContent;
    try {
      // Force fresh fetch by calling getRaw directly each time
      const rawContent = await browser.messages.getRaw(messageData.id);
      
      if (!rawContent || rawContent.length === 0) {
        throw new Error('E-Mail-Inhalt ist leer');
      }
      
      console.log('📧 Raw EML size:', rawContent.length, 'bytes');
      console.log('📧 Raw EML type:', typeof rawContent);
      // Use 400 bytes buffer to avoid cutting multi-byte UTF-8 characters
      try {
        const previewChars = typeof rawContent === 'string' 
          ? rawContent.substring(0, 200)
          : UTF8_DECODER.decode(rawContent.slice(0, 400)).substring(0, 200);
        console.log('📧 Raw EML first 200 chars:', previewChars);
      } catch (previewError) {
        console.log('📧 Raw EML preview failed (encoding issue):', previewError.message);
      }
      
      // WORKAROUND for libmagic MIME-type detection:
      // libmagic often fails to recognize message/rfc822 when From: is not at the start.
      // Moving the From-header to the beginning ensures correct detection.
      // See: Paperless-ngx mail.py lines 916-933
      emlContent = ensureFromHeaderAtBeginning(rawContent);
      
      console.log('📧 Processed EML size:', emlContent.length, 'bytes');
      console.log('📧 Processed EML type:', typeof emlContent);
      console.log('📧 Processed EML first 200 chars:', emlContent.substring(0, 200));
      
    } catch (emlError) {
      console.error('📧 Error getting raw email:', emlError);
      throw new Error(`E-Mail konnte nicht geladen werden: ${emlError.message}`);
    }

    // Create filename
    const dateStr = documentDate || new Date().toISOString().split('T')[0];
    const safeSubject = (messageData.subject || 'Kein_Betreff')
      .replace(/[^a-zA-Z0-9äöüÄÖÜß\s-]/g, '')
      .replace(/\s+/g, '_')
      .substring(0, 50);
    const emlFilename = `${dateStr}_${safeSubject}.eml`;
    console.log('📧 EML filename:', emlFilename);

    // IMPORTANT: Do not specify MIME type in Blob/File constructor.
    // This allows Paperless to use libmagic for detection, which will now
    // correctly identify message/rfc822 thanks to the From-header workaround.
    // Note: Using File instead of Blob ensures correct content transmission in all browser environments
    const emlBlob = new Blob([emlContent]);
    const emlFile = new File([emlBlob], emlFilename);
    console.log('📧 EML file size:', emlFile.size);

    // Upload EML file
    const emlFormData = new FormData();
    emlFormData.append('document', emlFile);
    emlFormData.append('title', safeSubject);
    
    if (emailDocumentType && emailDocumentType.id) {
      emlFormData.append('document_type', emailDocumentType.id);
    }
    if (documentDate) {
      emlFormData.append('created', documentDate);
    }
    if (correspondent) {
      emlFormData.append('correspondent', correspondent);
    }
    if (tags && tags.length > 0) {
      tags.forEach(tagId => emlFormData.append('tags', tagId));
    }

    console.log('📧 Uploading EML to Paperless...');
    const emlResponse = await fetch(`${config.url}/api/documents/post_document/`, {
      method: 'POST',
      headers: { 'Authorization': `Token ${config.token}` },
      body: emlFormData
    });

    if (!emlResponse.ok) {
      const errorText = await emlResponse.text();
      console.error('📧 EML upload failed:', errorText);
      throw new Error(`Upload failed (${emlResponse.status}): ${errorText}`);
    }

    // Add tag to Thunderbird email
    console.log('🏷️ Paperless-Tag: EML-Upload erfolgreich, füge Tag hinzu...');
    addPaperlessTagToEmail(messageData.id).catch(e =>
      console.warn("🏷️ Paperless-Tag: Fehler beim Taggen der E-Mail:", e)
    );

    // Wait for document processing
    const emlTaskId = await emlResponse.text();
    console.log('📧 EML upload task ID:', emlTaskId);
    console.log('📧 Waiting for Paperless to process (MailDocumentParser)...');
    const emailDocId = await waitForDocumentId(config, emlTaskId.replace(/"/g, ''));
    console.log('📧 Email document ID:', emailDocId);

    if (!emailDocId) {
      console.warn('📧 ⚠️ Email document ID not found after waiting');
      return {
        success: true,
        warning: 'Dokument hochgeladen, wird noch verarbeitet. Bitte später im Paperless-System prüfen.'
      };
    }

    // Upload attachments
    console.log('📧 Starting attachment uploads, count:', selectedAttachments?.length || 0);
    const attachmentDocIds = [];
    const attachmentErrors = [];
    
    for (const attachment of selectedAttachments || []) {
      console.log('📎 Processing attachment:', attachment.name);
      
      try {
        const attachmentFile = await browser.messages.getAttachmentFile(
          messageData.id,
          attachment.partName
        );
        
        if (!attachmentFile || attachmentFile.size === 0) {
          throw new Error(`Anhang '${attachment.name}' ist leer oder konnte nicht gelesen werden`);
        }

        const attachmentFormData = new FormData();
        attachmentFormData.append('document', attachmentFile, attachment.name);
        attachmentFormData.append('title', attachment.name.replace(/\.[^/.]+$/, ''));
        
        if (attachmentDocumentType && attachmentDocumentType.id) {
          attachmentFormData.append('document_type', attachmentDocumentType.id);
        }
        if (documentDate) {
          attachmentFormData.append('created', documentDate);
        }
        if (correspondent) {
          attachmentFormData.append('correspondent', correspondent);
        }
        if (tags && tags.length > 0) {
          tags.forEach(tagId => attachmentFormData.append('tags', tagId));
        }

        const attachmentResponse = await fetch(`${config.url}/api/documents/post_document/`, {
          method: 'POST',
          headers: { 'Authorization': `Token ${config.token}` },
          body: attachmentFormData
        });

        if (!attachmentResponse.ok) {
          attachmentErrors.push(`${attachment.name}: Upload failed`);
          continue;
        }

        const attachmentTaskId = await attachmentResponse.text();
        const attachmentDocId = await waitForDocumentId(config, attachmentTaskId.replace(/"/g, ''));

        if (attachmentDocId) {
          attachmentDocIds.push(attachmentDocId);
          console.log('📎 Attachment successfully uploaded:', attachment.name);
        } else {
          attachmentErrors.push(`${attachment.name}: ID nicht gefunden`);
        }
      } catch (error) {
        console.error('📎 Error uploading attachment:', attachment.name, error);
        attachmentErrors.push(`${attachment.name}: ${error.message}`);
      }
    }

    // Update custom fields
    console.log('📧 Updating custom fields...');
    
    try {
      const emailCustomFields = [];
      if (directionOptionId) {
        emailCustomFields.push({ field: directionField.id, value: String(directionOptionId) });
      }
      if (attachmentDocIds.length > 0) {
        emailCustomFields.push({ field: relatedDocsField.id, value: attachmentDocIds });
      }
      if (emailCustomFields.length > 0) {
        await updateDocumentCustomFields(config, emailDocId, emailCustomFields);
        console.log('📧 Email custom fields updated');
      }
    } catch (error) {
      console.error('📧 Custom field update error:', error);
    }

    // Update attachment custom fields
    for (const attachmentDocId of attachmentDocIds) {
      try {
        const attachmentCustomFields = [];
        if (directionOptionId) {
          attachmentCustomFields.push({ field: directionField.id, value: String(directionOptionId) });
        }
        attachmentCustomFields.push({ field: relatedDocsField.id, value: [emailDocId] });
        await updateDocumentCustomFields(config, attachmentDocId, attachmentCustomFields);
      } catch (error) {
        console.error('📧 Attachment field update error:', error);
      }
    }

    const totalDocs = 1 + attachmentDocIds.length;
    console.log('📧 EML upload complete. Total documents:', totalDocs);
    showNotification(`✅ ${totalDocs} Dokument(e) hochgeladen (via MailDocumentParser)!`, "success");

    return {
      success: true,
      emailDocId: emailDocId,
      attachmentDocIds: attachmentDocIds,
      attachmentErrors: attachmentErrors.length > 0 ? attachmentErrors : undefined,
      strategy: 'eml'
    };

  } catch (error) {
    console.error("📧 ❌ EML upload error:", error);
    return {
      success: false,
      error: error.message || 'Unbekannter Fehler'
    };
  }
}

// Wait for document to be processed and return the document ID
// Polls the Paperless-ngx tasks API until the document is processed or timeout occurs
async function waitForDocumentId(config, taskId, maxAttempts = DOCUMENT_PROCESSING_MAX_ATTEMPTS, delayMs = DOCUMENT_PROCESSING_DELAY_MS) {
  console.log(`📋 Waiting for document ID, taskId: ${taskId}`);
  
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    try {
      const taskUrl = `${config.url}/api/tasks/?task_id=${taskId}`;
      
      // Check task status
      const taskResponse = await fetch(taskUrl, {
        headers: { 'Authorization': `Token ${config.token}` }
      });

      if (taskResponse.ok) {
        const taskData = await taskResponse.json();
        
        if (taskData.length > 0) {
          const task = taskData[0];
          
          if (task.status === 'SUCCESS' && task.related_document) {
            console.log(`📋 ✅ Task completed successfully!`);
            
            let docId = null;
            const relatedDoc = task.related_document;
            
            // Try to extract document ID - Paperless can return it in different formats:
            // 1. As a simple string/number: "465" or 465
            // 2. As a URL: "/api/documents/465/"
            
            if (typeof relatedDoc === 'number') {
              // Direct number
              docId = relatedDoc;
            } else if (typeof relatedDoc === 'string') {
              // Try to parse as URL first
              const urlMatch = relatedDoc.match(/\/api\/documents\/(\d+)\//);
              if (urlMatch) {
                docId = parseInt(urlMatch[1], 10);
              } else if (/^\d+$/.test(relatedDoc)) {
                // Try to parse as simple number string
                docId = parseInt(relatedDoc, 10);
              } else {
                console.error(`📋 ❌ Could not parse document ID from: "${relatedDoc}"`);
              }
            } else {
              console.error(`📋 ❌ Unexpected related_document type: ${typeof relatedDoc}`);
            }
            
            if (Number.isInteger(docId) && docId >= 0) {
              console.log(`📋 ✅ Document ID: ${docId}`);
              return docId;
            } else {
              console.error(`📋 ❌ Could not determine valid document ID`);
              return null;
            }
          } else if (task.status === 'FAILURE') {
            console.error("📋 ❌ Task failed:", task.result);
            return null;
          }
          // For PENDING/STARTED status, continue polling without logging
        }
      } else {
        console.error(`📋 ❌ Task API request failed with status ${taskResponse.status}`);
      }

      // Wait before next attempt (no logging)
      await new Promise(resolve => setTimeout(resolve, delayMs));
    } catch (error) {
      console.error("📋 ❌ Error checking task status:", error);
    }
  }

  console.warn(`📋 ⏱️ Timeout waiting for document ID after ${maxAttempts} attempts`);
  return null;
}

// Convert base64 to Blob
function base64ToBlob(base64, mimeType) {
  const byteCharacters = atob(base64);
  const byteNumbers = new Array(byteCharacters.length);
  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);
  return new Blob([byteArray], { type: mimeType });
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
  if (message.action === "emailUploadFromDisplay") {
    await handleEmailUploadFromDisplay(message.messageId);
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

  // Handle email upload with attachments
  if (message.action === "uploadEmailWithAttachments") {
    // For Thunderbird/Firefox, we need to return the Promise directly
    // Using sendResponse() with return true has timing issues
    const { messageData, emailPdf, selectedAttachments, direction, correspondent, tags, documentDate } = message;

    // Return the Promise directly - the caller will receive the resolved value
    return uploadEmailWithAttachments(
      messageData,
      emailPdf,
      selectedAttachments,
      direction,
      correspondent,
      tags,
      documentDate
    );
  }

  // Handle email upload as HTML file (Paperless Gotenberg conversion - better character encoding)
  if (message.action === "uploadEmailAsHtml") {
    const { messageData, selectedAttachments, direction, correspondent, tags, documentDate } = message;

    // Return the Promise directly - the caller will receive the resolved value
    return uploadEmailAsHtml(
      messageData,
      selectedAttachments,
      direction,
      correspondent,
      tags,
      documentDate
    );
  }

  // Handle email upload as EML file (native format with libmagic compatibility)
  if (message.action === "uploadEmailAsEml") {
    const { messageData, selectedAttachments, direction, correspondent, tags, documentDate } = message;

    // Return the Promise directly - the caller will receive the resolved value
    return uploadEmailAsEml(
      messageData,
      selectedAttachments,
      direction,
      correspondent,
      tags,
      documentDate
    );
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

// Handle email upload from message display popup
async function handleEmailUploadFromDisplay(messageId) {
  try {
    const message = await browser.messages.get(messageId);
    await openEmailUploadDialog(message);
  } catch (error) {
    console.error("Error handling email upload from display:", error);
    showNotification("Fehler beim Verarbeiten der E-Mail", "error");
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
      showNotification("Keine PDF-Anhänge in der Nachricht gefunden", "info");
      return;
    }

    // Store current data for the dialog
    currentMessage = message;
    currentPdfAttachments = pdfAttachments;

    // Open the advanced upload dialog
    await openAdvancedUploadDialog(message, pdfAttachments);
  } catch (error) {
    console.error("Error handling advanced upload from display:", error);
    showNotification("Fehler beim Verarbeiten des Anhangs", "error");
  }
}
