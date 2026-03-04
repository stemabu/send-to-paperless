// Content script for messageDisplay - runs in every email view
// Reads the QNote note and sends it to the background script

function findQnoteElement() {
  // QNote injects its note into the outer Thunderbird window (parent frame),
  // not into the inner message display document.
  // Try multiple DOM levels and selectors.
  const selectors = [
    '.qnote-insidenote',
    '#qnote-insidenote',
    '[class*="qnote-inside"]',
    '[id*="qnote"]'
  ];

  const documents = [];
  // Current document (inner mail body)
  documents.push(document);
  // Parent document (Thunderbird 3-pane window or message display wrapper)
  try { if (window.parent && window.parent.document !== document) documents.push(window.parent.document); } catch(e) {}
  // Top-level document
  try { if (window.top && window.top.document !== document && window.top.document !== window.parent.document) documents.push(window.top.document); } catch(e) {}

  for (const doc of documents) {
    for (const sel of selectors) {
      try {
        const el = doc.querySelector(sel);
        if (el) return el;
      } catch(e) {}
    }
  }
  return null;
}

function tryReadQnote(attempts) {
  const el = findQnoteElement();
  if (el) {
    const noteText = el.innerText || el.textContent || null;
    if (noteText && noteText.trim()) {
      browser.runtime.sendMessage({
        action: 'qnoteNoteAvailable',
        noteText: noteText.trim()
      });
      return;
    }
  }

  // Retry up to 10 times with 300ms delay (total 3 seconds)
  // QNote may inject the note after the email content loads
  if (attempts < 10) {
    setTimeout(() => tryReadQnote(attempts + 1), 300);
  } else {
    // No note found after all attempts
    browser.runtime.sendMessage({
      action: 'qnoteNoteAvailable',
      noteText: null
    });
  }
}

// Start trying to read QNote
tryReadQnote(0);
