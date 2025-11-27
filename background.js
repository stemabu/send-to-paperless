// Background script for Paperless-ngx PDF Uploader
console.log("Paperless PDF Uploader loaded!");

// Configuration constants for document processing
const DOCUMENT_PROCESSING_MAX_ATTEMPTS = 60;
const DOCUMENT_PROCESSING_DELAY_MS = 1000;

// Configuration constants for Thunderbird tag
const PAPERLESS_TAG_KEY = 'paperless';
const PAPERLESS_TAG_LABEL = 'Paperless';
const PAPERLESS_TAG_COLOR = '#17a2b8';  // teal/cyan

let currentPdfAttachments = [];
let currentMessage = null;

// Create context menus for attachments
browser.runtime.onInstalled.addListener(async () => {
  // Remove all existing menus first to avoid conflicts
  await browser.menus.removeAll();

  // Message list context menus
  // E-Mail mit Anhängen hochladen (first option)
  browser.menus.create({
    id: "email-to-paperless",
    title: "E-Mail mit Anhaengen hochladen",
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
      showNotification("Keine Nachricht ausgewaehlt", "error");
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
      showNotification("Keine PDF-Anhaenge in der Nachricht gefunden", "info");
      return;
    }

    // Store current data for the dialog
    currentMessage = message;
    currentPdfAttachments = pdfAttachments;

    // Open the advanced upload dialog
    await openAdvancedUploadDialog(message, pdfAttachments);

  } catch (error) {
    console.error("Error handling advanced PDF upload:", error);
    showNotification("Fehler beim Verarbeiten der Anhaenge", "error");
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

    // Open the email upload dialog
    const dialogUrl = browser.runtime.getURL("email-upload-dialog.html");
    browser.windows.create({
      url: dialogUrl,
      type: "popup",
      width: 550,
      height: 700
    });

  } catch (error) {
    console.error("Error opening email upload dialog:", error);
    showNotification("Fehler beim Öffnen des Dialogs", "error");
  }
}

// Extract email body from full message
function extractEmailBody(fullMessage) {
  let body = '';
  
  // Recursive function to find the body part
  function findBody(part) {
    if (part.body) {
      // Prefer text/plain, but use text/html as fallback
      if (part.contentType === 'text/plain' || !part.contentType) {
        body = part.body;
        return true;
      }
      if (part.contentType === 'text/html' && !body) {
        body = part.body;
      }
    }
    
    if (part.parts) {
      for (const subPart of part.parts) {
        if (findBody(subPart)) {
          return true;
        }
      }
    }
    return false;
  }
  
  findBody(fullMessage);
  return body;
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
function findPaperlessTag(tags) {
  return tags.find(tag => 
    tag.tag.toLowerCase() === PAPERLESS_TAG_KEY || 
    tag.key.toLowerCase() === PAPERLESS_TAG_KEY
  );
}

// Add "Paperless" keyword to email in Thunderbird
// Temporarily disabled - user will implement this separately
async function addPaperlessTagToEmail(messageId) {
  console.log('🏷️ Tag function disabled, messageId:', messageId);
  return true;
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

    // Upload email PDF
    console.log('📧 Starting email PDF upload...');
    showNotification("E-Mail-PDF wird hochgeladen...", "info");
    
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
    
    // Add document type "E-Mail" if it doesn't exist as string but will be matched
    // Paperless-ngx accepts document_type as name string and will match or create
    // For now we just include it as a hint - actual type assignment happens later
    
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
      showNotification(`Anhang wird hochgeladen: ${attachment.name}...`, "info");
      
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
    showNotification("Verknüpfungen werden erstellt...", "info");

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

    // Add Paperless tag to email in Thunderbird
    // Disabled - will be implemented separately
    // console.log('📧 Adding Paperless tag to email in Thunderbird...');
    // await addPaperlessTagToEmail(messageData.id);

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

// Wait for document to be processed and return the document ID
// Polls the Paperless-ngx tasks API until the document is processed or timeout occurs
async function waitForDocumentId(config, taskId, maxAttempts = DOCUMENT_PROCESSING_MAX_ATTEMPTS, delayMs = DOCUMENT_PROCESSING_DELAY_MS) {
  console.log(`📋 Starting to wait for document ID, taskId: ${taskId}`);
  console.log(`📋 Max attempts: ${maxAttempts}, delay: ${delayMs}ms`);
  
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    try {
      console.log(`📋 Checking task status, attempt ${attempt + 1}/${maxAttempts}`);
      
      const taskUrl = `${config.url}/api/tasks/?task_id=${taskId}`;
      console.log(`📋 Task API URL: ${taskUrl}`);
      
      // Check task status
      const taskResponse = await fetch(taskUrl, {
        headers: { 'Authorization': `Token ${config.token}` }
      });

      console.log(`📋 Task response status: ${taskResponse.status}`);

      if (taskResponse.ok) {
        const taskData = await taskResponse.json();
        console.log(`📋 Task data received:`, JSON.stringify(taskData, null, 2));
        
        if (taskData.length > 0) {
          const task = taskData[0];
          console.log(`📋 Task status: ${task.status}`);
          console.log(`📋 Task related_document: ${task.related_document}`);
          
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
              console.log(`📋 Document ID is a number: ${docId}`);
            } else if (typeof relatedDoc === 'string') {
              // Try to parse as URL first
              const urlMatch = relatedDoc.match(/\/api\/documents\/(\d+)\//);
              if (urlMatch) {
                docId = parseInt(urlMatch[1], 10);
                console.log(`📋 Document ID extracted from URL: ${docId}`);
              } else if (/^\d+$/.test(relatedDoc)) {
                // Try to parse as simple number string
                docId = parseInt(relatedDoc, 10);
                console.log(`📋 Document ID parsed from string: ${docId}`);
              } else {
                console.error(`📋 ❌ Could not parse document ID from: "${relatedDoc}"`);
                console.error(`📋 Type: ${typeof relatedDoc}, Value: ${JSON.stringify(relatedDoc)}`);
              }
            } else {
              console.error(`📋 ❌ Unexpected related_document type: ${typeof relatedDoc}`);
              console.error(`📋 Value: ${JSON.stringify(relatedDoc)}`);
            }
            
            if (Number.isInteger(docId) && docId >= 0) {
              console.log(`📋 ✅ Final document ID: ${docId}`);
              return docId;
            } else {
              console.error(`📋 ❌ Could not determine valid document ID`);
              return null;
            }
          } else if (task.status === 'FAILURE') {
            console.error("📋 ❌ Task failed:", task.result);
            return null;
          } else {
            console.log(`📋 ⏳ Task still processing (status: ${task.status})...`);
          }
        } else {
          console.log(`📋 ⚠️ No task data returned (empty array)`);
        }
      } else {
        console.error(`📋 ❌ Task API request failed with status ${taskResponse.status}`);
      }

      // Wait before next attempt
      console.log(`📋 Waiting ${delayMs}ms before next attempt...`);
      await new Promise(resolve => setTimeout(resolve, delayMs));
    } catch (error) {
      console.error("📋 ❌ Error checking task status:", error);
      console.error("📋 Error stack:", error.stack);
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