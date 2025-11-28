// Email upload dialog - handles converting email to PDF and uploading to Paperless-ngx

let currentMessage = null;
let currentAttachments = [];
let emailBody = '';

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
    emailBody = uploadData.emailBody || '';

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

async function loadCorrespondents() {
  try {
    const settings = await getPaperlessSettings();
    if (!settings.paperlessUrl || !settings.paperlessToken) {
      return;
    }

    const response = await makePaperlessRequest('/api/correspondents/', {}, settings);

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

    const response = await makePaperlessRequest('/api/tags/', {}, settings);

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

// Generate PDF from email
function generateEmailPdf() {
  console.log('📄 Starting PDF generation...');
  
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

  // Email body
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  
  // Process email body - extract plain text if HTML
  let bodyText = emailBody;
  if (bodyText.includes('<html') || bodyText.includes('<body') || bodyText.includes('<div')) {
    bodyText = extractPlainTextFromHtml(bodyText);
  }

  // Split body into lines that fit the page width
  const bodyLines = doc.splitTextToSize(bodyText, contentWidth);
  const bodyLineHeight = 5;

  for (const line of bodyLines) {
    // Check if we need a new page
    if (yPosition + bodyLineHeight > pageHeight - margin) {
      doc.addPage();
      yPosition = margin;
    }
    
    doc.text(line, margin, yPosition);
    yPosition += bodyLineHeight;
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

// Extract plain text from HTML
function extractPlainTextFromHtml(html) {
  // Create a temporary element to parse HTML
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = html;
  
  // Remove script and style elements
  const scripts = tempDiv.querySelectorAll('script, style');
  scripts.forEach(el => el.remove());
  
  // Get text content
  let text = tempDiv.textContent || tempDiv.innerText || '';
  
  // Clean up whitespace
  text = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]+/g, ' ')
    .trim();
  
  return text;
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
    console.log('📤 - Correspondent:', correspondent);
    console.log('📤 - Tags:', selectedTags);
    console.log('📤 - Document Date:', documentDate);
    console.log('📤 - Selected attachments:', selectedAttachments.length);
    selectedAttachments.forEach((att, i) => {
      console.log(`📤   [${i}] ${att.name} (${att.contentType}, partName: ${att.partName})`);
    });

    // Generate PDF from email
    console.log('📤 Generating email PDF...');
    const { blob: pdfBlob, filename: pdfFilename } = generateEmailPdf();
    console.log('📤 Generated PDF:', pdfFilename, 'size:', pdfBlob.size);

    // Convert blob to base64
    console.log('📤 Converting PDF to base64...');
    const pdfBase64 = await blobToBase64(pdfBlob);
    console.log('📤 PDF base64 length:', pdfBase64.length);

    // Send upload request to background script
    console.log('📤 Sending message to background script...');
    console.log('📤 Message data:', JSON.stringify(currentMessage));
    
    const result = await browser.runtime.sendMessage({
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

    console.log('📤 Received result from background:', JSON.stringify(result));

    if (result && result.success) {
      let successMsg = 'E-Mail und Anhänge wurden erfolgreich an Paperless-ngx gesendet!';
      
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
