// Email upload dialog - handles converting email to PDF and uploading to Paperless-ngx

let currentMessage = null;
let currentAttachments = [];
let emailBody = '';

// Get file icon based on file extension
function getFileIcon(filename) {
  const ext = filename.toLowerCase().split('.').pop();
  
  switch (ext) {
    case 'pdf':
      return '📄';
    case 'doc':
    case 'docx':
      return '📝';
    case 'xls':
    case 'xlsx':
      return '📊';
    case 'jpg':
    case 'jpeg':
    case 'png':
    case 'gif':
    case 'bmp':
      return '🖼️';
    case 'zip':
    case 'rar':
    case '7z':
    case 'tar':
    case 'gz':
      return '📦';
    default:
      return '📎';
  }
}

document.addEventListener('DOMContentLoaded', async function () {
  await loadEmailData();
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
  const margin = 15;
  const contentWidth = pageWidth - (margin * 2);
  let yPosition = margin;

  // Header background (gray)
  doc.setFillColor(240, 240, 240);
  
  // Calculate header height (will be adjusted after adding content)
  let headerHeight = 50; // Initial estimate
  
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

  // Prepare attachment list
  const selectedAttachments = getSelectedAttachments();
  let attachmentText = '';
  if (selectedAttachments.length > 0) {
    attachmentText = selectedAttachments.map(att => `${getFileIcon(att.name)} ${att.name}`).join(', ');
  }

  // Calculate actual header height
  doc.setFontSize(10);
  const subjectLines = doc.splitTextToSize(currentMessage.subject || '', contentWidth - 50);
  const fromLines = doc.splitTextToSize(currentMessage.author || '', contentWidth - 30);
  const toLines = recipients ? doc.splitTextToSize(recipients, contentWidth - 30) : [];
  const attachmentLines = attachmentText ? doc.splitTextToSize(attachmentText, contentWidth - 50) : [];
  
  headerHeight = 20 + // Date line + padding
    (fromLines.length * 5) + // From lines
    (toLines.length > 0 ? toLines.length * 5 + 5 : 5) + // To lines or spacing
    (subjectLines.length * 6) + // Subject lines (slightly larger)
    (attachmentLines.length > 0 ? attachmentLines.length * 5 + 5 : 0) + // Attachment lines
    10; // Bottom padding

  // Draw header background
  doc.setFillColor(240, 240, 240);
  doc.rect(margin, margin, contentWidth, headerHeight, 'F');

  yPosition = margin + 5;

  // Date line
  doc.setFontSize(10);
  doc.setFont('helvetica', 'bold');
  doc.text('Datum:', margin + 5, yPosition + 4);
  doc.setFont('helvetica', 'normal');
  doc.text(emailDate, margin + 30, yPosition + 4);
  yPosition += 7;

  // From line
  doc.setFont('helvetica', 'bold');
  doc.text('Von:', margin + 5, yPosition + 4);
  doc.setFont('helvetica', 'normal');
  fromLines.forEach((line, index) => {
    doc.text(line, margin + 30, yPosition + 4 + (index * 5));
  });
  yPosition += 7 + ((fromLines.length - 1) * 5);

  // To line (if recipients exist)
  if (recipients) {
    doc.setFont('helvetica', 'bold');
    doc.text('An:', margin + 5, yPosition + 4);
    doc.setFont('helvetica', 'normal');
    toLines.forEach((line, index) => {
      doc.text(line, margin + 30, yPosition + 4 + (index * 5));
    });
    yPosition += 7 + ((toLines.length - 1) * 5);
  }

  // Subject line (bold)
  doc.setFont('helvetica', 'bold');
  doc.text('Betreff:', margin + 5, yPosition + 4);
  doc.setFont('helvetica', 'bold');
  subjectLines.forEach((line, index) => {
    doc.text(line, margin + 30, yPosition + 4 + (index * 6));
  });
  yPosition += 8 + ((subjectLines.length - 1) * 6);

  // Attachments line (if any)
  if (attachmentText) {
    doc.setFont('helvetica', 'bold');
    doc.text('Anhänge:', margin + 5, yPosition + 4);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    attachmentLines.forEach((line, index) => {
      doc.text(line, margin + 30, yPosition + 4 + (index * 5));
    });
    yPosition += 7 + ((attachmentLines.length - 1) * 5);
    doc.setFontSize(10);
  }

  // Horizontal line separator
  yPosition = margin + headerHeight + 2;
  doc.setDrawColor(180, 180, 180);
  doc.setLineWidth(0.5);
  doc.line(margin, yPosition, pageWidth - margin, yPosition);

  yPosition += 8;

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
  const lineHeight = 5;

  for (const line of bodyLines) {
    // Check if we need a new page
    if (yPosition + lineHeight > pageHeight - margin) {
      doc.addPage();
      yPosition = margin;
    }
    
    doc.text(line, margin, yPosition);
    yPosition += lineHeight;
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

    console.log('📤 Upload parameters:');
    console.log('📤 - Direction:', direction);
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
      direction: direction
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
