// Content script for messageDisplay - runs in every email view
// Reads the QNote note and sends it to the background script

// Prevent multiple executions if script is injected multiple times
if (window.__qnoteReaderExecuted) {
  console.log('⚠️ [QNote-Reader] Script already executed, skipping');
  // Still send a message in case the previous execution timed out
  // Small delay to ensure the background listener is ready
  setTimeout(() => {
    browser.runtime.sendMessage({
      action: 'qnoteNoteAvailable',
      noteText: window.__lastQnoteText || null
    }).catch(() => {});
  }, 100);
} else {
  window.__qnoteReaderExecuted = true;

console.log('🔍 [QNote-Reader] Script started');

function findQnoteElement() {
  console.log('🔍 [QNote-Reader] findQnoteElement() called');

  // QNote injects its note into the outer Thunderbird window (parent frame),
  // not into the inner message display document.
  // Try multiple DOM levels and selectors.
  const selectors = [
    '.qnote-text',        // Direct text container (more specific)
    '.qnote-insidenote',
    '#qnote-insidenote',
    '[class*="qnote-inside"]',
    '[id*="qnote"]'
  ];

  const documents = [];
  // Current document (inner mail body)
  documents.push(document);
  console.log('🔍 [QNote-Reader] Searching in document');
  // Parent document (Thunderbird 3-pane window or message display wrapper)
  try {
    if (window.parent && window.parent.document !== document) {
      documents.push(window.parent.document);
      console.log('🔍 [QNote-Reader] Searching in window.parent.document');
    }
  } catch(e) {
    console.log('⚠️ [QNote-Reader] Cannot access window.parent.document:', e.message);
  }
  // Top-level document
  try {
    if (window.top && window.top.document !== document && window.top.document !== window.parent.document) {
      documents.push(window.top.document);
      console.log('🔍 [QNote-Reader] Searching in window.top.document');
    }
  } catch(e) {
    console.log('⚠️ [QNote-Reader] Cannot access window.top.document:', e.message);
  }

  for (const doc of documents) {
    for (const sel of selectors) {
      try {
        const el = doc.querySelector(sel);
        if (el) {
          console.log(`✅ [QNote-Reader] FOUND element with selector "${sel}"!`);
          console.log('✅ [QNote-Reader] Element text:', el.innerText || el.textContent);
          return el;
        }
      } catch(e) {
        console.log(`⚠️ [QNote-Reader] Error with selector "${sel}":`, e.message);
      }
    }
  }

  console.log('❌ [QNote-Reader] No QNote element found in any document');
  return null;
}

function tryReadQnote(attempts) {
  console.log(`🔍 [QNote-Reader] tryReadQnote() attempt ${attempts + 1}/10`);

  const el = findQnoteElement();
  if (el) {
    const noteText = el.innerText || el.textContent || null;
    if (noteText && noteText.trim()) {
      console.log('✅ [QNote-Reader] Note text found:', noteText.trim());
      console.log('📤 [QNote-Reader] Sending message to background script');
      window.__lastQnoteText = noteText.trim();
      browser.runtime.sendMessage({
        action: 'qnoteNoteAvailable',
        noteText: noteText.trim()
      }).then(() => {
        console.log('✅ [QNote-Reader] Message sent successfully');
      }).catch(err => {
        console.error('❌ [QNote-Reader] Failed to send message:', err);
      });
      return;
    } else {
      console.log('⚠️ [QNote-Reader] Element found but text is empty');
    }
  }

  // Retry up to 10 times with 300ms delay (total 3 seconds)
  // QNote may inject the note after the email content loads
  if (attempts < 10) {
    console.log(`⏳ [QNote-Reader] Retrying in 300ms... (attempt ${attempts + 1}/10)`);
    setTimeout(() => tryReadQnote(attempts + 1), 300);
  } else {
    console.log('❌ [QNote-Reader] No note found after all attempts, sending null');
    browser.runtime.sendMessage({
      action: 'qnoteNoteAvailable',
      noteText: null
    }).then(() => {
      console.log('✅ [QNote-Reader] Null message sent successfully');
    }).catch(err => {
      console.error('❌ [QNote-Reader] Failed to send null message:', err);
    });
  }
}

// Start trying to read QNote
console.log('🚀 [QNote-Reader] Starting QNote search...');
tryReadQnote(0);
} // end else (not already executed)
