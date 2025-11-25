// email-to-pdf.js

const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const path = require('path');

async function convertEmailToPDF(emailData) {
    // Create a new PDF document
    const pdfDoc = await PDFDocument.create();
    
    // Create a header with the email metadata
    const page = pdfDoc.addPage();
    page.drawText(`Subject: ${emailData.subject}`, { x: 50, y: 700 });
    page.drawText(`From: ${emailData.from}`, { x: 50, y: 680 });
    page.drawText(`To: ${emailData.to}`, { x: 50, y: 660 });
    page.drawText(`Date: ${emailData.date}`, { x: 50, y: 640 });
    
    // Add email body content
    page.drawText(emailData.body, { x: 50, y: 600 });
    
    // Add an attachments list, if any
    if (emailData.attachments && emailData.attachments.length > 0) {
        page.drawText(`Attachments:`, { x: 50, y: 580 });
        emailData.attachments.forEach((attachment, index) => {
            page.drawText(`${index + 1}. ${attachment}`, { x: 50, y: 560 - (index * 20) });
        });
    }
    
    // Save the PDF document
    const pdfBytes = await pdfDoc.save();
    return pdfBytes;
}

module.exports = convertEmailToPDF;