// Email upload dialog - handles converting email to PDF and uploading to Paperless-ngx

let currentMessage = null;
let currentAttachments = [];
let emailBody = '';
let isHtmlBody = false;

// Extract email address from email string like "John Doe <john@example.com>"
function extractEmailAddress(emailString) {
  if (!emailString) return null;
  
  const match = emailString.match(/<(.+?)>/);
  if (match) {
    return match[1].trim().toLowerCase();
  }
  
  return emailString.trim().toLowerCase();
}

// Find correspondent match based on email addresses
async function findCorrespondentMatch() {
  try {
    const result = await browser.storage.sync.get('emailCorrespondentMapping');
    const mappings = result.emailCorrespondentMapping || [];
    
    if (mappings.length === 0) {
      return null;
    }
    
    const fromEmail = extractEmailAddress(currentMessage.author);
    const toEmails = (currentMessage.recipients || []).map(r => extractEmailAddress(r));
    
    console.log('📧 Checking for correspondent match...');
    console.log('📧 From:', fromEmail);
    console.log('📧 To:', toEmails);
    
    // Check FROM field first - if match found, it's incoming mail
    const fromMatch = mappings.find(m => m.email === fromEmail);
    if (fromMatch) {
      console.log('📧 Match found in FROM field:', fromMatch);
      return {
        correspondentId: fromMatch.correspondentId,
        correspondentName: fromMatch.correspondentName,
        direction: 'Eingang'
      };
    }
    
    // Check TO field - if match found, it's outgoing mail
    for (const toEmail of toEmails) {
      const toMatch = mappings.find(m => m.email === toEmail);
      if (toMatch) {
        console.log('📧 Match found in TO field:', toMatch);
        return {
          correspondentId: toMatch.correspondentId,
          correspondentName: toMatch.correspondentName,
          direction: 'Ausgang'
        };
      }
    }
    
    console.log('📧 No correspondent match found');
    return null;
  } catch (error) {
    console.error('Error finding correspondent match:', error);
    return null;
  }
}

// Get file icon (emoji) based on file extension - used in the dialog UI
function getFileIcon(filename) {
  const ext = filename.toLowerCase().split('.').pop();
  
  switch (ext) {
    case 'pdf':
      return '📄';
    case 'doc':
    case 'docx':
    case 'odt':
      return '📝';
    case 'xls':
    case 'xlsx':
    case 'ods':
    case 'csv':
      return '📊';
    case 'ppt':
    case 'pptx':
    case 'odp':
      return '📽️';
    case 'jpg':
    case 'jpeg':
    case 'png':
    case 'gif':
    case 'bmp':
    case 'svg':
    case 'webp':
    case 'ico':
    case 'tiff':
    case 'tif':
      return '🖼️';
    case 'zip':
    case 'rar':
    case '7z':
    case 'tar':
    case 'gz':
    case 'bz2':
    case 'xz':
      return '📦';
    case 'txt':
    case 'rtf':
    case 'md':
    case 'log':
      return '📃';
    case 'html':
    case 'htm':
    case 'xml':
    case 'json':
      return '🌐';
    case 'mp3':
    case 'wav':
    case 'ogg':
    case 'flac':
    case 'aac':
    case 'm4a':
      return '🎵';
    case 'mp4':
    case 'avi':
    case 'mov':
    case 'mkv':
    case 'webm':
    case 'wmv':
    case 'flv':
      return '🎬';
    case 'eml':
    case 'msg':
      return '✉️';
    case 'exe':
    case 'msi':
    case 'dmg':
    case 'app':
      return '⚙️';
    default:
      return '📎';
  }
}

document.addEventListener('DOMContentLoaded', async function () {
  await loadEmailData();
  await loadCorrespondents();
  await applyCorrespondentMatch();
  await loadTags();
  setupEventListeners();
});

async function loadEmailData() {
  console.log('📧 Loading email data...');
  
  try {
    const result = await browser.storage.local.get('emailUploadData');
    const uploadData = result.emailUploadData;
    
    console.log('📧 Upload data retrieved:', uploadData ? 'yes' : 'no');

    if (!uploadData) {
      console.error('📧 No email upload data found in storage');
      showError("Keine E-Mail-Daten gefunden. Bitte versuchen Sie es erneut.");
      return;
    }

    currentMessage = uploadData.message;
    currentAttachments = uploadData.attachments || [];
    
    // Handle both old string format and new object format from extractEmailBody
    if (uploadData.emailBody && typeof uploadData.emailBody === 'object') {
      emailBody = uploadData.emailBody.body || '';
      isHtmlBody = uploadData.emailBody.isHtml || false;
    } else {
      emailBody = uploadData.emailBody || '';
      // Auto-detect HTML content for backward compatibility using more robust pattern
      // Look for opening HTML tags with their closing brackets to avoid false positives
      isHtmlBody = /<html[\s>]|<body[\s>]|<div[\s>]|<p[\s>]|<table[\s>]/i.test(emailBody);
    }

    console.log('📧 - Email body length:', emailBody.length);
    console.log('📧 - Is HTML body:', isHtmlBody);

    // Debug: Show first 200 characters of email body for inspection
    if (emailBody) {
      console.log('📧 - Email body preview (first 200 chars):', emailBody.substring(0, 200));
      console.log('📧 - Email body char codes (first 50 chars):', 
        Array.from(emailBody.substring(0, 50)).map(c => c.charCodeAt(0)).join(','));
    }
    
    // Decode HTML entities if present (for text/plain emails with HTML entities)
    // Thunderbird decodes Quoted-Printable automatically, but NOT HTML entities
    if (emailBody && hasHtmlEntities(emailBody)) {
      console.log('📧 Detected HTML entities in email body, decoding...');
      const beforeLength = emailBody.length;
      emailBody = decodeHtmlEntities(emailBody);
      console.log('📧 After HTML entity decoding, length:', emailBody.length);
      console.log('📧 Decoded', (beforeLength - emailBody.length), 'characters');
      console.log('📧 Decoded preview (first 200 chars):', emailBody.substring(0, 200));
    }

    console.log('📧 Email loaded:');
    console.log('📧 - From:', currentMessage.author);
    console.log('📧 - Subject:', currentMessage.subject);
    console.log('📧 - Date:', currentMessage.date);
    console.log('📧 - Message ID:', currentMessage.id);
    console.log('📧 - Attachments:', currentAttachments.length);
    currentAttachments.forEach((att, i) => {
      console.log(`📧   [${i}] ${att.name} (${att.contentType}, ${att.size} bytes, partName: ${att.partName})`);
    });
    console.log('📧 - Email body length:', emailBody.length);
    console.log('📧 - Is HTML body:', isHtmlBody);

    // Populate email info
    document.getElementById('emailFrom').textContent = currentMessage.author;
    document.getElementById('emailSubject').textContent = currentMessage.subject;
    document.getElementById('emailDate').textContent = new Date(currentMessage.date).toLocaleDateString('de-DE', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });

    // Display Thunderbird tags if available
    if (currentMessage.tags && currentMessage.tags.length > 0) {
      try {
        // Get all available Thunderbird tags to map keys to labels
        let allTags = [];
        if (browser.messages?.listTags) {
          allTags = await browser.messages.listTags();
        } else if (browser.messages?.tags?.list) {
          allTags = await browser.messages.tags.list();
        }
        
        // Map tag keys to labels
        const tagLabels = currentMessage.tags.map(tagKey => {
          const tagInfo = allTags.find(t => t.key === tagKey);
          return tagInfo ? (tagInfo.label || tagInfo.tag || tagKey) : tagKey;
        });
        
        // Display tags
        if (tagLabels.length > 0) {
          document.getElementById('emailTags').textContent = tagLabels.join(', ');
          document.getElementById('emailTagsRow').style.display = 'block';
        }
      } catch (error) {
        console.error('Error loading Thunderbird tags:', error);
        // Don't show the row if there's an error
      }
    }

    // Populate attachments if any
    if (currentAttachments.length > 0) {
      console.log('📧 Showing attachment section');
      document.getElementById('attachmentSection').style.display = 'block';
      await populateAttachmentList();
    } else {
      console.log('📧 No attachments to display');
    }

    // Show main content
    document.getElementById('loadingSection').style.display = 'none';
    document.getElementById('mainContent').style.display = 'block';
    
    console.log('📧 Email data loaded successfully');

  } catch (error) {
    console.error('📧 Error loading email data:', error);
    console.error('📧 Error stack:', error.stack);
    showError('Fehler beim Laden der E-Mail-Daten: ' + error.message);
  }
}

// Apply correspondent match after correspondents are loaded
async function applyCorrespondentMatch() {
  const match = await findCorrespondentMatch();
  if (match) {
    console.log('📧 Applying correspondent suggestion:', match);
    
    const correspondentSelect = document.getElementById('correspondent');
    correspondentSelect.value = match.correspondentId;
    
    const directionSelect = document.getElementById('direction');
    directionSelect.value = match.direction;
    
    console.log('📧 Pre-selected correspondent:', match.correspondentName);
    console.log('📧 Pre-selected direction:', match.direction);
  }
}

async function loadCorrespondents() {
  try {
    const settings = await getPaperlessSettings();
    if (!settings.paperlessUrl || !settings.paperlessToken) {
      return;
    }

    const response = await makePaperlessRequest('/api/correspondents/?page_size=1000', {}, settings);

    if (response.ok) {
      const data = await response.json();
      const select = document.getElementById('correspondent');
      
      data.results.forEach(correspondent => {
        const option = document.createElement('option');
        option.value = correspondent.id;
        option.textContent = correspondent.name;
        select.appendChild(option);
      });
    }
  } catch (error) {
    console.error('Error loading correspondents:', error);
  }
}

async function loadTags() {
  try {
    const settings = await getPaperlessSettings();
    if (!settings.paperlessUrl || !settings.paperlessToken) {
      return;
    }

    const response = await makePaperlessRequest('/api/tags/?page_size=1000', {}, settings);

    if (response.ok) {
      const data = await response.json();
      const select = document.getElementById('tags');
      
      data.results.forEach(tag => {
        const option = document.createElement('option');
        option.value = tag.id;
        option.textContent = tag.name;
        select.appendChild(option);
      });
    }
  } catch (error) {
    console.error('Error loading tags:', error);
  }
}

async function populateAttachmentList() {
  const listContainer = document.getElementById('attachmentList');
  listContainer.textContent = ''; // Clear existing content

  for (const [index, attachment] of currentAttachments.entries()) {
    const item = document.createElement('div');
    item.className = 'attachment-item';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.className = 'attachment-checkbox';
    checkbox.id = `attachment-${index}`;
    checkbox.dataset.index = index;
    checkbox.checked = true; // Default: all selected

    const icon = document.createElement('span');
    icon.className = 'attachment-icon';
    icon.textContent = getFileIcon(attachment.name);

    const infoDiv = document.createElement('div');
    infoDiv.className = 'attachment-info';

    const nameDiv = document.createElement('div');
    nameDiv.className = 'attachment-name';
    nameDiv.textContent = attachment.name;

    const sizeDiv = document.createElement('div');
    sizeDiv.className = 'attachment-size';
    sizeDiv.textContent = await browser.messengerUtilities.formatFileSize(attachment.size);

    infoDiv.appendChild(nameDiv);
    infoDiv.appendChild(sizeDiv);
    
    item.appendChild(checkbox);
    item.appendChild(icon);
    item.appendChild(infoDiv);
    listContainer.appendChild(item);
  }
}

function setupEventListeners() {
  // Form submission
  document.getElementById('emailUploadForm').addEventListener('submit', handleUpload);

  // Cancel button
  document.getElementById('cancelBtn').addEventListener('click', () => {
    window.close();
  });

  // Select all button
  document.getElementById('selectAllBtn').addEventListener('click', () => {
    const checkboxes = document.querySelectorAll('.attachment-checkbox');
    checkboxes.forEach(cb => cb.checked = true);
  });

  // Select none button
  document.getElementById('selectNoneBtn').addEventListener('click', () => {
    const checkboxes = document.querySelectorAll('.attachment-checkbox');
    checkboxes.forEach(cb => cb.checked = false);
  });
}

// Get file type indicator for PDF (plain text, since jsPDF doesn't support emoji well)
function getFileTypeIndicator(filename) {
  const ext = filename.toLowerCase().split('.').pop();
  
  switch (ext) {
    case 'pdf':
      return '[PDF]';
    case 'doc':
    case 'docx':
      return '[DOC]';
    case 'xls':
    case 'xlsx':
    case 'ods':
    case 'csv':
      return '[XLS]';
    case 'ppt':
    case 'pptx':
    case 'odp':
      return '[PPT]';
    case 'jpg':
    case 'jpeg':
    case 'png':
    case 'gif':
    case 'bmp':
    case 'svg':
    case 'webp':
      return '[IMG]';
    case 'zip':
    case 'rar':
    case '7z':
    case 'tar':
    case 'gz':
      return '[ZIP]';
    case 'txt':
    case 'rtf':
    case 'md':
    case 'log':
      return '[TXT]';
    case 'html':
    case 'htm':
      return '[HTML]';
    case 'xml':
      return '[XML]';
    case 'json':
      return '[JSON]';
    case 'mp3':
    case 'wav':
    case 'ogg':
    case 'flac':
      return '[AUDIO]';
    case 'mp4':
    case 'avi':
    case 'mov':
    case 'mkv':
    case 'webm':
      return '[VIDEO]';
    case 'eml':
    case 'msg':
      return '[MAIL]';
    default:
      return '[FILE]';
  }
}

// Generate PDF from email (async to support HTML rendering)
async function generateEmailPdf() {
  console.log('📄 Starting PDF generation...');
  console.log('📄 Is HTML body:', isHtmlBody);
  
  // jsPDF is loaded from jspdf.umd.min.js as window.jspdf
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });

  console.log('📄 jsPDF initialized');

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 10;  // 1cm margin
  const contentWidth = pageWidth - (margin * 2);
  const lineHeight = 4;  // Base line height for font size 10
  const headerPadding = 3; // Padding inside header box
  const labelWidth = 22; // Width reserved for labels (Datum:, Von:, etc.)
  
  // Prepare header content
  const emailDate = new Date(currentMessage.date).toLocaleDateString('de-DE', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });

  // Get recipients
  const recipients = currentMessage.recipients ? currentMessage.recipients.join(', ') : '';

  // Prepare attachment list with file type indicators
  const selectedAttachments = getSelectedAttachments();
  let attachmentTextLines = [];
  if (selectedAttachments.length > 0) {
    // Build attachment text with type indicators
    const attachmentStrings = selectedAttachments.map(att => {
      const typeIndicator = getFileTypeIndicator(att.name);
      return `${typeIndicator} ${att.name}`;
    });
    // Join with comma and let splitTextToSize handle wrapping
    const attachmentText = attachmentStrings.join(', ');
    doc.setFontSize(9);
    attachmentTextLines = doc.splitTextToSize(attachmentText, contentWidth - labelWidth - headerPadding);
    doc.setFontSize(10);
  }

  // Get Thunderbird tag labels for PDF header
  let thunderbirdTagLabels = [];
  let thunderbirdTagLines = [];
  if (currentMessage.tags && currentMessage.tags.length > 0) {
    try {
      // Get all available Thunderbird tags to map keys to labels
      let allTags = [];
      if (browser.messages?.listTags) {
        allTags = await browser.messages.listTags();
      } else if (browser.messages?.tags?.list) {
        allTags = await browser.messages.tags.list();
      }
      
      // Map tag keys to labels
      thunderbirdTagLabels = currentMessage.tags.map(tagKey => {
        const tagInfo = allTags.find(t => t.key === tagKey);
        return tagInfo ? (tagInfo.label || tagInfo.tag || tagKey) : tagKey;
      });
      
      // Prepare tag text lines for PDF
      if (thunderbirdTagLabels.length > 0) {
        const tagText = thunderbirdTagLabels.join(', ');
        doc.setFontSize(10);
        thunderbirdTagLines = doc.splitTextToSize(tagText, contentWidth - labelWidth - headerPadding);
        console.log('📄 Thunderbird tags for PDF:', thunderbirdTagLabels.join(', '));
      }
    } catch (error) {
      console.error('Error loading Thunderbird tags for PDF:', error);
    }
  }

  // Pre-calculate all text lines for accurate header height
  doc.setFontSize(10);
  const subjectLines = doc.splitTextToSize(currentMessage.subject || '', contentWidth - labelWidth - headerPadding);
  const fromLines = doc.splitTextToSize(currentMessage.author || '', contentWidth - labelWidth - headerPadding);
  const toLines = recipients ? doc.splitTextToSize(recipients, contentWidth - labelWidth - headerPadding) : [];
  
  // Calculate exact header height based on content
  let headerContentHeight = 0;
  headerContentHeight += lineHeight; // Date line
  headerContentHeight += fromLines.length * lineHeight; // From lines
  if (toLines.length > 0) {
    headerContentHeight += toLines.length * lineHeight; // To lines
  }
  headerContentHeight += subjectLines.length * lineHeight; // Subject lines
  if (thunderbirdTagLines.length > 0) {
    headerContentHeight += thunderbirdTagLines.length * lineHeight; // Thunderbird tags lines
  }
  if (attachmentTextLines.length > 0) {
    headerContentHeight += attachmentTextLines.length * lineHeight; // Attachment lines
  }
  
  // Total header height = top padding + content + bottom padding
  const headerHeight = headerPadding + headerContentHeight + headerPadding;

  // Draw header background - exact fit to content
  doc.setFillColor(240, 240, 240);
  doc.rect(margin, margin, contentWidth, headerHeight, 'F');

  // Start writing content inside header
  let yPosition = margin + headerPadding;
  const textX = margin + headerPadding;
  const valueX = margin + headerPadding + labelWidth;

  // Date line
  doc.setFontSize(10);
  doc.setFont('helvetica', 'bold');
  doc.text('Datum:', textX, yPosition + 3);
  doc.setFont('helvetica', 'normal');
  doc.text(emailDate, valueX, yPosition + 3);
  yPosition += lineHeight;

  // From line
  doc.setFont('helvetica', 'bold');
  doc.text('Von:', textX, yPosition + 3);
  doc.setFont('helvetica', 'normal');
  fromLines.forEach((line, index) => {
    doc.text(line, valueX, yPosition + 3 + (index * lineHeight));
  });
  yPosition += fromLines.length * lineHeight;

  // To line (if recipients exist)
  if (toLines.length > 0) {
    doc.setFont('helvetica', 'bold');
    doc.text('An:', textX, yPosition + 3);
    doc.setFont('helvetica', 'normal');
    toLines.forEach((line, index) => {
      doc.text(line, valueX, yPosition + 3 + (index * lineHeight));
    });
    yPosition += toLines.length * lineHeight;
  }

  // Subject line (bold value)
  doc.setFont('helvetica', 'bold');
  doc.text('Betreff:', textX, yPosition + 3);
  subjectLines.forEach((line, index) => {
    doc.text(line, valueX, yPosition + 3 + (index * lineHeight));
  });
  yPosition += subjectLines.length * lineHeight;

  // Thunderbird tags line (if any)
  if (thunderbirdTagLines.length > 0) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('Schlagwörter:', textX, yPosition + 3);
    doc.setFont('helvetica', 'normal');
    thunderbirdTagLines.forEach((line, index) => {
      doc.text(line, valueX, yPosition + 3 + (index * lineHeight));
    });
    yPosition += thunderbirdTagLines.length * lineHeight;
  }

  // Attachments line (if any)
  if (attachmentTextLines.length > 0) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('Anhänge:', textX, yPosition + 3);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    attachmentTextLines.forEach((line, index) => {
      doc.text(line, valueX, yPosition + 3 + (index * lineHeight));
    });
    yPosition += attachmentTextLines.length * lineHeight;
    doc.setFontSize(10);
  }

  // Position for separator line - directly after header box
  yPosition = margin + headerHeight;
  
  // Horizontal line separator - directly at end of header
  doc.setDrawColor(180, 180, 180);
  doc.setLineWidth(0.5);
  doc.line(margin, yPosition, pageWidth - margin, yPosition);

  yPosition += 5; // Space after separator before body

  // Email body rendering
  if (isHtmlBody && emailBody.trim()) {
    console.log('📄 Rendering HTML body with html2canvas...');
    await renderHtmlBodyToPdf(doc, emailBody, margin, yPosition, contentWidth, pageHeight);
  } else {
    console.log('📄 Rendering plain text body...');
    renderPlainTextBody(doc, emailBody, margin, yPosition, contentWidth, pageHeight);
  }

  // Generate filename
  const dateStr = new Date(currentMessage.date).toISOString().split('T')[0];
  const safeSubject = (currentMessage.subject || 'Kein_Betreff')
    .replace(/[^a-zA-Z0-9äöüÄÖÜß\s-]/g, '')
    .replace(/\s+/g, '_')
    .substring(0, 50);
  const filename = `${dateStr}_${safeSubject}.pdf`;

  console.log('📄 PDF generated successfully');
  console.log('📄 Filename:', filename);
  
  const pdfBlob = doc.output('blob');
  console.log('📄 PDF blob size:', pdfBlob.size);

  return {
    blob: pdfBlob,
    filename: filename
  };
}

// Render HTML body to PDF using html2canvas with improved charset handling
async function renderHtmlBodyToPdf(doc, htmlContent, margin, startY, contentWidth, pageHeight) {
  console.log('📄 Rendering HTML body with html2canvas...');
  console.log('📄 HTML content length:', htmlContent.length);
  console.log('📄 HTML preview (first 200 chars):', htmlContent.substring(0, 200));
  
  // Conversion factor: 1 mm = ~3.78 pixels (at 96 DPI)
  const MM_TO_PIXELS = 3.78;
  
  // Create a temporary container for rendering HTML
  const container = document.createElement('div');
  container.style.cssText = `
    position: absolute;
    left: -9999px;
    top: -9999px;
    width: ${contentWidth * MM_TO_PIXELS}px;
    background: white;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    font-size: 12px;
    line-height: 1.5;
    color: #333;
    padding: 10px;
  `;
  
  // Sanitize and process HTML content
  const sanitizedHtml = sanitizeHtmlForPdf(htmlContent);
  
  // Set innerHTML with proper charset handling
  try {
    container.innerHTML = sanitizedHtml;
    console.log('📄 HTML parsed successfully');
  } catch (error) {
    console.error('📄 Error parsing HTML:', error);
    // Fallback: use DOMParser to safely extract text content
    // DOMParser always succeeds even with malformed HTML
    const parser = new DOMParser();
    const parsedDoc = parser.parseFromString(htmlContent, 'text/html');
    container.textContent = parsedDoc.body ? parsedDoc.body.textContent : '';
    console.log('📄 Fell back to plain text extraction');
  }
  
  // Add to document for rendering
  document.body.appendChild(container);
  
  try {
    // Use html2canvas to render HTML to canvas
    const canvas = await html2canvas(container, {
      scale: 2, // Higher quality
      useCORS: true,
      logging: false,
      backgroundColor: '#ffffff',
      windowWidth: contentWidth * MM_TO_PIXELS,
      removeContainer: false,
      // Add charset handling
      onclone: (clonedDoc) => {
        // Ensure proper charset in cloned document
        const meta = clonedDoc.createElement('meta');
        meta.setAttribute('charset', 'UTF-8');
        clonedDoc.head.insertBefore(meta, clonedDoc.head.firstChild);
      }
    });
    
    console.log('📄 HTML rendered to canvas:', canvas.width, 'x', canvas.height);
    
    // Convert canvas to image and add to PDF
    const imgData = canvas.toDataURL('image/jpeg', 0.95);
    
    // Calculate image dimensions
    const imgWidth = contentWidth;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;
    
    // Check if image fits on current page
    const availableHeight = pageHeight - startY - margin;
    
    if (imgHeight <= availableHeight) {
      // Image fits on one page
      doc.addImage(imgData, 'JPEG', margin, startY, imgWidth, imgHeight);
    } else {
      // Need to split across multiple pages
      let remainingHeight = imgHeight;
      let sourceY = 0;
      let currentY = startY;
      let isFirstPage = true;
      
      while (remainingHeight > 0) {
        const pageAvailableHeight = isFirstPage ? availableHeight : (pageHeight - margin * 2);
        const heightToDraw = Math.min(remainingHeight, pageAvailableHeight);
        
        // Calculate the portion of the original image to use
        const sourceHeight = (heightToDraw / imgHeight) * canvas.height;
        
        // Create a temporary canvas for this portion
        const tempCanvas = document.createElement('canvas');
        tempCanvas.width = canvas.width;
        tempCanvas.height = sourceHeight;
        const tempCtx = tempCanvas.getContext('2d');
        tempCtx.drawImage(canvas, 0, sourceY, canvas.width, sourceHeight, 0, 0, canvas.width, sourceHeight);
        
        const partImgData = tempCanvas.toDataURL('image/jpeg', 0.95);
        doc.addImage(partImgData, 'JPEG', margin, currentY, imgWidth, heightToDraw);
        
        remainingHeight -= heightToDraw;
        sourceY += sourceHeight;
        
        if (remainingHeight > 0) {
          doc.addPage();
          currentY = margin;
          isFirstPage = false;
        }
      }
    }
    
    console.log('📄 HTML body added to PDF');
  } catch (error) {
    console.error('📄 Error rendering HTML to PDF:', error);
    // Fallback to plain text rendering
    console.log('📄 Falling back to plain text rendering');
    const plainText = container.textContent || container.innerText || '';
    renderPlainTextBody(doc, plainText, margin, startY, contentWidth, pageHeight);
  } finally {
    // Clean up - use remove() for safer element removal
    if (container.parentNode) {
      container.remove();
    }
  }
}

// Check if a URL has a dangerous scheme (javascript:, vbscript:, data:)
// Also handles URL-encoded schemes
function hasDangerousUrlScheme(url) {
  if (!url) return false;
  
  // First try to decode the URL to handle encoded schemes
  let decodedUrl = url;
  try {
    decodedUrl = decodeURIComponent(url);
  } catch (e) {
    // If decoding fails, use original URL
  }
  
  // Check both original and decoded URLs
  const urlsToCheck = [url.toLowerCase(), decodedUrl.toLowerCase()];
  const dangerousSchemes = ['javascript:', 'vbscript:', 'data:'];
  
  return urlsToCheck.some(urlToCheck => 
    dangerousSchemes.some(scheme => urlToCheck.startsWith(scheme))
  );
}

// Sanitize and simplify HTML for reliable PDF rendering
// Converts complex HTML structures to simple, well-supported elements
function sanitizeHtmlForPdf(html) {
  console.log('📄 Sanitizing HTML for PDF...');
  
  // Create a temporary element to parse HTML
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = html;
  
  console.log('📄 Original HTML length:', html.length);
  
  // 1. Remove dangerous and non-rendering elements
  const elementsToRemove = tempDiv.querySelectorAll(
    'script, style, link, meta, head, iframe, frame, frameset, object, embed, applet, ' +
    'form, input, button, select, textarea, fieldset, legend'
  );
  console.log('📄 Removing', elementsToRemove.length, 'dangerous/non-rendering elements');
  elementsToRemove.forEach(el => el.remove());
  
  // 2. Expand <details> elements (convert to visible div)
  const detailsElements = tempDiv.querySelectorAll('details');
  console.log('📄 Expanding', detailsElements.length, 'details elements');
  detailsElements.forEach(details => {
    // Remove the 'open' behavior, just show content
    const div = document.createElement('div');
    div.className = 'expanded-details';
    div.innerHTML = details.innerHTML;
    details.replaceWith(div);
  });
  
  // 3. Simplify <summary> to bold text
  const summaryElements = tempDiv.querySelectorAll('summary');
  console.log('📄 Simplifying', summaryElements.length, 'summary elements');
  summaryElements.forEach(summary => {
    const strong = document.createElement('strong');
    strong.textContent = summary.textContent;
    summary.replaceWith(strong);
  });
  
  // 4. Convert complex semantic elements to simple divs/spans
  const semanticToSimple = {
    'article': 'div',
    'section': 'div',
    'aside': 'div',
    'nav': 'div',
    'header': 'div',
    'footer': 'div',
    'main': 'div',
    'figure': 'div',
    'figcaption': 'div',
    'mark': 'span',
    'time': 'span',
    'meter': 'span',
    'progress': 'span'
  };
  
  Object.keys(semanticToSimple).forEach(oldTag => {
    const elements = tempDiv.querySelectorAll(oldTag);
    console.log(`📄 Converting ${elements.length} <${oldTag}> to <${semanticToSimple[oldTag]}>`);
    elements.forEach(el => {
      const newEl = document.createElement(semanticToSimple[oldTag]);
      newEl.innerHTML = el.innerHTML;
      // Copy useful attributes
      if (el.className) newEl.className = el.className;
      if (el.id) newEl.id = el.id;
      el.replaceWith(newEl);
    });
  });
  
  // 5. Remove all event handlers and dangerous attributes
  const allElements = tempDiv.querySelectorAll('*');
  console.log('📄 Cleaning attributes from', allElements.length, 'elements');
  allElements.forEach(el => {
    // Remove all on* event attributes
    Array.from(el.attributes).forEach(attr => {
      if (attr.name.startsWith('on')) {
        el.removeAttribute(attr.name);
      }
    });
    
    // Remove dangerous URL schemes (javascript:, vbscript:, data:)
    // Also handles URL-encoded schemes
    if (el.href) {
      if (hasDangerousUrlScheme(el.href)) {
        el.removeAttribute('href');
      }
    }
    // Also check src attributes for dangerous schemes
    if (el.src) {
      if (hasDangerousUrlScheme(el.src)) {
        el.removeAttribute('src');
      }
    }
    
    // Remove Outlook-specific and vendor-specific attributes
    Array.from(el.attributes).forEach(attr => {
      if (attr.name.startsWith('mso-') || 
          attr.name.startsWith('v:') || 
          attr.name.startsWith('o:') ||
          attr.name.startsWith('xmlns:')) {
        el.removeAttribute(attr.name);
      }
    });
    
    // Simplify dir attribute (keep only ltr/rtl/auto)
    if (el.hasAttribute('dir')) {
      const dir = el.getAttribute('dir').toLowerCase();
      if (!['ltr', 'rtl', 'auto'].includes(dir)) {
        el.removeAttribute('dir');
      }
    }
  });
  
  // 6. Add comprehensive styling for reliable rendering
  const styleTag = document.createElement('style');
  styleTag.textContent = `
    /* Reset and base styles */
    * { box-sizing: border-box; }
    body, html { 
      margin: 0; 
      padding: 0; 
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
      font-size: 12px;
      line-height: 1.6;
      color: #333;
      background: white;
    }
    
    /* Text elements */
    p { margin: 0.5em 0; }
    h1 { font-size: 1.8em; margin: 0.8em 0 0.4em; font-weight: bold; }
    h2 { font-size: 1.5em; margin: 0.8em 0 0.4em; font-weight: bold; }
    h3 { font-size: 1.3em; margin: 0.8em 0 0.4em; font-weight: bold; }
    h4 { font-size: 1.1em; margin: 0.8em 0 0.4em; font-weight: bold; }
    h5, h6 { font-size: 1em; margin: 0.8em 0 0.4em; font-weight: bold; }
    
    /* Lists */
    ul, ol { margin: 0.5em 0; padding-left: 2em; }
    li { margin: 0.25em 0; }
    
    /* Tables */
    table { 
      border-collapse: collapse; 
      margin: 0.5em 0; 
      width: 100%;
      max-width: 100%;
    }
    td, th { 
      padding: 6px 8px; 
      border: 1px solid #ddd; 
      text-align: left;
      vertical-align: top;
    }
    th { 
      font-weight: bold; 
      background: #f5f5f5; 
    }
    
    /* Images */
    img { 
      max-width: 100%; 
      height: auto; 
      display: block;
      margin: 0.5em 0;
    }
    
    /* Links */
    a { 
      color: #0066cc; 
      text-decoration: underline; 
    }
    
    /* Quotes */
    blockquote { 
      margin: 0.5em 0; 
      padding-left: 1em; 
      border-left: 3px solid #ccc; 
      color: #666; 
      font-style: italic;
    }
    
    /* Code */
    pre, code { 
      font-family: 'Courier New', Courier, monospace; 
      background: #f5f5f5; 
      padding: 2px 4px; 
      border-radius: 3px;
      font-size: 0.9em;
    }
    pre { 
      padding: 8px; 
      overflow-x: auto; 
      white-space: pre-wrap;
      word-wrap: break-word;
    }
    
    /* Horizontal rule */
    hr { 
      border: none; 
      border-top: 1px solid #ddd; 
      margin: 1em 0; 
    }
    
    /* Expanded details styling */
    .expanded-details { 
      margin: 0.5em 0; 
      padding: 0.5em; 
      border: 1px solid #ddd; 
      border-radius: 3px;
      background: #fafafa;
    }
    
    /* Text formatting */
    strong, b { font-weight: bold; }
    em, i { font-style: italic; }
    u { text-decoration: underline; }
    s, strike, del { text-decoration: line-through; }
    
    /* Divs and spans */
    div { margin: 0; }
    span { display: inline; }
    
    /* Remove Outlook-specific styles */
    [class*="mso"] { display: block !important; }
    
    /* Ensure text is visible and readable */
    * {
      max-width: 100%;
      word-wrap: break-word;
      overflow-wrap: break-word;
    }
  `;
  
  // Insert style at the beginning
  tempDiv.insertBefore(styleTag, tempDiv.firstChild);
  
  const result = tempDiv.innerHTML;
  console.log('📄 Sanitized HTML length:', result.length);
  console.log('📄 HTML sanitization complete');
  
  return result;
}

// Render plain text body to PDF with improved character handling
function renderPlainTextBody(doc, text, margin, startY, contentWidth, pageHeight) {
  console.log('📄 Rendering plain text body...');
  console.log('📄 Text length:', text.length);
  console.log('📄 Text preview (first 100 chars):', text.substring(0, 100));
  
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  
  // Clean up the text - remove control characters except newlines (\x0A) and tabs (\x09)
  // This regex removes: NULL (\x00-\x08), vertical tab (\x0B), form feed (\x0C), 
  // other control chars (\x0E-\x1F), and DEL (\x7F)
  let bodyText = text.replace(/[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]/g, '');
  
  console.log('📄 After cleanup, length:', bodyText.length);
  console.log('📄 After cleanup, preview:', bodyText.substring(0, 100));
  
  // Split body into lines that fit the page width
  const bodyLines = doc.splitTextToSize(bodyText, contentWidth);
  const bodyLineHeight = 5;
  let yPosition = startY;

  console.log('📄 Total lines to render:', bodyLines.length);

  for (let i = 0; i < bodyLines.length; i++) {
    const line = bodyLines[i];
    
    // Check if we need a new page
    if (yPosition + bodyLineHeight > pageHeight - margin) {
      doc.addPage();
      yPosition = margin;
      console.log('📄 Added new page at line', i);
    }
    
    // Try to render the line, catch any errors
    try {
      doc.text(line, margin, yPosition);
    } catch (error) {
      console.error('📄 Error rendering line', i, ':', error);
      console.error('📄 Problematic line content:', line);
      // Try to render a placeholder instead with line number for debugging
      doc.text(`[Rendering error at line ${i + 1}]`, margin, yPosition);
    }
    
    yPosition += bodyLineHeight;
  }
  
  console.log('📄 Plain text rendering complete');
}

// Get selected attachments
function getSelectedAttachments() {
  const checkboxes = document.querySelectorAll('.attachment-checkbox:checked');
  return Array.from(checkboxes).map(cb => {
    const index = parseInt(cb.dataset.index, 10);
    return currentAttachments[index];
  });
}

async function handleUpload(event) {
  event.preventDefault();

  const uploadBtn = document.getElementById('uploadBtn');
  const originalHtml = uploadBtn.innerHTML;
  uploadBtn.disabled = true;
  uploadBtn.innerHTML = '⏳ Wird hochgeladen...';

  console.log('📤 Starting upload process...');

  try {
    clearMessages();

    const direction = document.getElementById('direction').value;
    // Checkbox: checked = local PDF, unchecked = Gotenberg (direct API call)
    const pdfStrategyCheckbox = document.getElementById('pdfStrategy');
    const pdfStrategy = pdfStrategyCheckbox.checked ? 'local' : 'gotenberg';
    const selectedAttachments = getSelectedAttachments();

    // Get correspondent
    const correspondentSelect = document.getElementById('correspondent');
    const correspondent = correspondentSelect.value ? parseInt(correspondentSelect.value) : null;

    // Get tags
    const tagsSelect = document.getElementById('tags');
    const selectedTags = Array.from(tagsSelect.selectedOptions).map(opt => parseInt(opt.value));

    // Get document date from email date
    const documentDate = currentMessage.date ? new Date(currentMessage.date).toISOString().split('T')[0] : null;

    console.log('📤 Upload parameters:');
    console.log('📤 - Direction:', direction);
    console.log('📤 - PDF Strategy:', pdfStrategy);
    console.log('📤 - Correspondent:', correspondent);
    console.log('📤 - Tags:', selectedTags);
    console.log('📤 - Document Date:', documentDate);
    console.log('📤 - Selected attachments:', selectedAttachments.length);
    selectedAttachments.forEach((att, i) => {
      console.log(`📤   [${i}] ${att.name} (${att.contentType}, partName: ${att.partName})`);
    });

    let result;

    if (pdfStrategy === 'gotenberg') {
      // Upload email via direct Gotenberg API call (HTML → PDF)
      console.log('📤 Using Gotenberg upload strategy (direct API call)...');
      
      result = await browser.runtime.sendMessage({
        action: 'uploadEmailAsHtml',
        messageData: currentMessage,
        selectedAttachments: selectedAttachments,
        direction: direction,
        correspondent: correspondent,
        tags: selectedTags,
        documentDate: documentDate
      });
    } else {
      // Generate PDF locally using html2canvas + jsPDF
      console.log('📤 Using local PDF generation strategy...');
      console.log('📤 Generating email PDF...');
      const { blob: pdfBlob, filename: pdfFilename } = await generateEmailPdf();
      console.log('📤 Generated PDF:', pdfFilename, 'size:', pdfBlob.size);

      // Convert blob to base64
      console.log('📤 Converting PDF to base64...');
      const pdfBase64 = await blobToBase64(pdfBlob);
      console.log('📤 PDF base64 length:', pdfBase64.length);

      // Send upload request to background script
      console.log('📤 Sending message to background script...');
      console.log('📤 Message data:', JSON.stringify(currentMessage));
      
      result = await browser.runtime.sendMessage({
        action: 'uploadEmailWithAttachments',
        messageData: currentMessage,
        emailPdf: {
          blob: pdfBase64,
          filename: pdfFilename
        },
        selectedAttachments: selectedAttachments,
        direction: direction,
        correspondent: correspondent,
        tags: selectedTags,
        documentDate: documentDate
      });
    }

    console.log('📤 Received result from background:', JSON.stringify(result));

    if (result && result.success) {
      let successMsg = pdfStrategy === 'gotenberg' 
        ? 'E-Mail via Gotenberg erfolgreich hochgeladen!'
        : 'E-Mail und Anhänge wurden erfolgreich an Paperless-ngx gesendet!';
      
      // Show warning if document is still processing
      if (result.warning) {
        console.warn('📤 Upload completed with warning:', result.warning);
        successMsg = result.warning;
      }
      
      // Show attachment errors if any
      if (result.attachmentErrors && result.attachmentErrors.length > 0) {
        console.warn('📤 Some attachments had errors:', result.attachmentErrors);
        successMsg += '\n\nHinweis: Einige Anhänge konnten nicht hochgeladen werden:\n• ' + 
          result.attachmentErrors.join('\n• ');
      }
      
      showSuccess(successMsg);
      closeWindowWithDelay(2000);
    } else {
      // Build a detailed error message
      let errorMsg = 'Fehler beim Upload';
      
      if (result && result.error) {
        errorMsg = result.error;
      } else if (!result) {
        errorMsg = 'Keine Antwort vom Hintergrundskript erhalten';
      }
      
      console.error('📤 Upload failed:', errorMsg);
      console.error('📤 Full result:', result);
      
      // Log additional error details if available
      if (result && result.errorDetails) {
        console.error('📤 Error details:', result.errorDetails);
      }
      
      showError('Fehler beim Upload: ' + errorMsg);
    }

  } catch (error) {
    console.error('📤 Upload exception:', error);
    console.error('📤 Error name:', error.name);
    console.error('📤 Error message:', error.message);
    console.error('📤 Error stack:', error.stack);
    
    showError('Fehler beim Upload: ' + (error.message || 'Unbekannter Fehler'));
  } finally {
    uploadBtn.disabled = false;
    uploadBtn.innerHTML = originalHtml;
  }
}

// Convert blob to base64 for sending via message
async function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = reader.result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
