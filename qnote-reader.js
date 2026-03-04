// Content script for messageDisplay - runs in every email view
// Reads the QNote note and sends it to the background script

function tryReadQnote(attempts) {
  const el = document.querySelector('.qnote-insidenote');
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

  // Retry up to 5 times with 200ms delay (total 800ms)
  if (attempts < 5) {
    setTimeout(() => tryReadQnote(attempts + 1), 200);
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
