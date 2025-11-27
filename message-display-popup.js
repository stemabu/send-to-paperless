document.addEventListener('DOMContentLoaded', async function () {
  const uploadEmailBtn = document.getElementById('upload-email-btn');
  const uploadAttachmentsBtn = document.getElementById('upload-attachments-btn');
  const errorContainer = document.getElementById('error-container');

  uploadEmailBtn.addEventListener('click', async () => {
    await handleEmailUpload(errorContainer);
  });

  uploadAttachmentsBtn.addEventListener('click', async () => {
    await handleAttachmentsUpload(errorContainer);
  });
});

async function handleEmailUpload(errorContainer) {
  try {
    clearError(errorContainer);
    const currentTab = await getCurrentTab();

    if (!currentTab) {
      showError(errorContainer, 'Aktueller Tab konnte nicht ermittelt werden');
      return;
    }

    // Get the displayed messages (returns a MessageList)
    const messageList = await browser.messageDisplay.getDisplayedMessages(currentTab.id);

    // Get the first message from the MessageList
    let message = null;

    if (messageList && messageList.messages && messageList.messages.length > 0) {
      message = messageList.messages[0];
    }

    if (!message) {
      showError(errorContainer, 'Keine Nachricht wird angezeigt');
      return;
    }

    // Send message to background script for email upload
    await browser.runtime.sendMessage({
      action: 'emailUploadFromDisplay',
      messageId: message.id
    });

    // Close the popup
    window.close();
  } catch (error) {
    console.error('Error in email upload:', error);
    showError(errorContainer, 'Fehler beim Starten des E-Mail-Uploads');
  }
}

async function handleAttachmentsUpload(errorContainer) {
  try {
    clearError(errorContainer);
    const currentTab = await getCurrentTab();

    if (!currentTab) {
      showError(errorContainer, 'Aktueller Tab konnte nicht ermittelt werden');
      return;
    }

    // Get the displayed messages (returns a MessageList)
    const messageList = await browser.messageDisplay.getDisplayedMessages(currentTab.id);

    // Get the first message from the MessageList
    let message = null;

    if (messageList && messageList.messages && messageList.messages.length > 0) {
      message = messageList.messages[0];
    }

    if (!message) {
      showError(errorContainer, 'Keine Nachricht wird angezeigt');
      return;
    }

    // Send message to background script for attachments upload (was advanced upload)
    await browser.runtime.sendMessage({
      action: 'advancedUploadFromDisplay',
      messageId: message.id
    });

    // Close the popup
    window.close();
  } catch (error) {
    console.error('Error in attachments upload:', error);
    showError(errorContainer, 'Fehler beim Starten des Anhang-Uploads');
  }
}

async function getCurrentTab() {
  try {
    const tabs = await browser.tabs.query({ active: true, currentWindow: true });
    return tabs.length > 0 ? tabs[0] : null;
  } catch (error) {
    console.error('Error getting current tab:', error);
    return null;
  }
}

function showError(container, message) {
  const errorDiv = document.createElement('div');
  errorDiv.className = 'error-message';
  errorDiv.textContent = message;
  clearError(container);
  container.appendChild(errorDiv);
}

function clearError(container) {
  while (container.firstChild) {
    container.removeChild(container.firstChild);
  }
}
