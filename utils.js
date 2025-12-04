// Shared utility functions for Paperless-ngx PDF Uploader extension


/**
 * Display an error message in the message area
 * @param {string} message - Error message to display
 * @param {string} messageAreaId - ID of the message area element (default: 'messageArea')
 */
function showError(message, messageAreaId = 'messageArea') {
  const messageArea = document.getElementById(messageAreaId);
  if (!messageArea) {
    console.error('Message area element not found:', messageAreaId);
    return;
  }

  const errorDiv = document.createElement('div');
  errorDiv.className = 'error';
  errorDiv.textContent = message;
  messageArea.appendChild(errorDiv);
}

/**
 * Display a success message in the message area
 * @param {string} message - Success message to display
 * @param {string} messageAreaId - ID of the message area element (default: 'messageArea')
 */
function showSuccess(message, messageAreaId = 'messageArea') {
  const messageArea = document.getElementById(messageAreaId);
  if (!messageArea) {
    console.error('Message area element not found:', messageAreaId);
    return;
  }

  const successDiv = document.createElement('div');
  successDiv.className = 'success';
  successDiv.textContent = message;
  messageArea.appendChild(successDiv);
}

/**
 * Clear all messages from the message area
 * @param {string} messageAreaId - ID of the message area element (default: 'messageArea')
 */
function clearMessages(messageAreaId = 'messageArea') {
  const messageArea = document.getElementById(messageAreaId);
  if (!messageArea) {
    console.error('Message area element not found:', messageAreaId);
    return;
  }

  while (messageArea.firstChild) {
    messageArea.removeChild(messageArea.firstChild);
  }
}

/**
 * Get Paperless-ngx settings from browser storage
 * @returns {Promise<Object>} Settings object with paperlessUrl, paperlessToken, and defaultTags
 */
async function getPaperlessSettings() {
  return await browser.storage.sync.get(['paperlessUrl', 'paperlessToken', 'defaultTags']);
}

/**
 * Test connection to Paperless-ngx API
 * @param {string} url - Paperless-ngx URL
 * @param {string} token - Paperless-ngx API token
 * @returns {Promise<boolean>} True if connection successful, false otherwise
 */
async function testPaperlessConnection(url, token) {
  try {
    const response = await fetch(`${url.replace(/\/$/, '')}/api/documents/`, {
      method: 'GET',
      headers: {
        'Authorization': `Token ${token}`,
        'Content-Type': 'application/json'
      }
    });
    return response.ok;
  } catch (error) {
    console.error('Connection test failed:', error);
    return false;
  }
}

/**
 * Validate if a string is a valid URL
 * @param {string} string - String to validate
 * @returns {boolean} True if valid URL, false otherwise
 */
function isValidUrl(string) {
  try {
    new URL(string);
    return true;
  } catch (_) {
    return false;
  }
}

/**
 * Extract name from email string (removes email part in angle brackets)
 * @param {string} emailString - Email string like "John Doe <john@example.com>"
 * @returns {string} Extracted name or original string if no angle brackets
 */
function extractNameFromEmail(emailString) {
  if (!emailString) return '';

  // Check if email is in format "Name <email@domain.com>"
  const match = emailString.match(/^(.+?)\s*<.+>$/);
  if (match) {
    return match[1].trim();
  }

  // If no angle brackets, return the original string
  return emailString.trim();
}

/**
 * Set button to loading state
 * @param {HTMLButtonElement} button - Button element
 * @param {string} loadingText - Text to show while loading (default: 'Loading...')
 * @returns {string} Original button text
 */
function setButtonLoading(button, loadingText = 'Loading...') {
  const originalText = button.textContent;
  button.disabled = true;
  button.textContent = loadingText;
  return originalText;
}

/**
 * Reset button from loading state
 * @param {HTMLButtonElement} button - Button element
 * @param {string} originalText - Original button text to restore
 */
function resetButtonLoading(button, originalText) {
  button.disabled = false;
  button.textContent = originalText;
}

/**
 * Send message to parent window (for popup dialogs)
 * @param {string} action - Action name
 * @param {boolean} success - Whether the action was successful
 * @param {Object} data - Additional data to send (optional)
 */
function sendMessageToParent(action, success, data = {}) {
  if (window.opener) {
    window.opener.postMessage({
      action: action,
      success: success,
      ...data
    }, '*');
  }
}

/**
 * Close window with optional delay
 * @param {number} delay - Delay in milliseconds (default: 0)
 */
function closeWindowWithDelay(delay = 0) {
  if (delay > 0) {
    setTimeout(() => window.close(), delay);
  } else {
    window.close();
  }
}

/**
 * Create a centered popup window with position/size persistence
 * @param {string} url - URL for the popup window
 * @param {number} width - Default window width in pixels
 * @param {number} height - Default window height in pixels
 * @returns {Promise<Object>} The created window object
 */
async function createCenteredWindow(url, width, height) {
  try {
    // Try to load saved window position/size
    const stored = await browser.storage.local.get('dialogWindowPreferences');
    
    if (stored.dialogWindowPreferences) {
      console.log('🪟 Using saved window preferences:', stored.dialogWindowPreferences);
      return await browser.windows.create({
        url: url,
        type: "popup",
        width: stored.dialogWindowPreferences.width || width,
        height: stored.dialogWindowPreferences.height || height,
        left: stored.dialogWindowPreferences.left,
        top: stored.dialogWindowPreferences.top,
        allowScriptsToClose: true
      });
    }
    
    // No saved preferences - calculate centered position
    let currentWindow;
    try {
      currentWindow = await browser.windows.getCurrent();
    } catch (e) {
      console.warn('Could not get current window, trying fallback:', e);
      // Fallback: Get all windows and find the focused one
      const allWindows = await browser.windows.getAll();
      currentWindow = allWindows.find(w => w.focused) || allWindows[0];
    }

    let left = 0;
    let top = 0;
    
    if (currentWindow && currentWindow.left !== undefined && currentWindow.width) {
      // Center relative to current window
      left = Math.round(currentWindow.left + (currentWindow.width - width) / 2);
      top = Math.round(currentWindow.top + (currentWindow.height - height) / 2);
    } else {
      // Fallback: Center on screen (rough estimate)
      left = Math.round((screen.availWidth - width) / 2);
      top = Math.round((screen.availHeight - height) / 2);
    }
    
    // Ensure window is not outside screen bounds
    left = Math.max(0, Math.min(left, screen.availWidth - width));
    top = Math.max(0, Math.min(top, screen.availHeight - height));
    
    console.log(`🪟 Creating centered window at: left=${left}, top=${top}, size=${width}x${height}`);
    
    return await browser.windows.create({
      url: url,
      type: "popup",
      width: width,
      height: height,
      left: left,
      top: top,
      allowScriptsToClose: true
    });
  } catch (error) {
    console.error('Error creating centered window:', error);
    // Final fallback: create window without position
    return await browser.windows.create({
      url: url,
      type: "popup",
      width: width,
      height: height,
      allowScriptsToClose: true
    });
  }
}

/**
 * Make API request to Paperless-ngx
 * @param {string} endpoint - API endpoint (e.g., '/api/documents/')
 * @param {Object} options - Fetch options
 * @param {Object} settings - Paperless settings with url and token
 * @returns {Promise<Response>} Fetch response
 */
async function makePaperlessRequest(endpoint, options = {}, settings = null) {
  if (!settings) {
    settings = await getPaperlessSettings();
  }

  if (!settings.paperlessUrl || !settings.paperlessToken) {
    throw new Error('Paperless-ngx settings not configured');
  }

  const url = `${settings.paperlessUrl.replace(/\/$/, '')}${endpoint}`;
  const defaultOptions = {
    headers: {
      'Authorization': `Token ${settings.paperlessToken}`,
      'Content-Type': 'application/json'
    }
  };

  const mergedOptions = {
    ...defaultOptions,
    ...options,
    headers: {
      ...defaultOptions.headers,
      ...(options.headers || {})
    }
  };

  return await fetch(url, mergedOptions);
}

/**
 * Decode HTML entities in text
 * Uses a textarea element to safely decode HTML entities.
 * This is a safe approach because textarea.innerHTML only decodes entities
 * without executing any scripts or creating DOM elements.
 * @param {string} text - Text with HTML entities
 * @returns {string} Decoded text
 */
function decodeHtmlEntities(text) {
  if (!text) return text;
  
  const textarea = document.createElement('textarea');
  textarea.innerHTML = text;
  return textarea.value;
}

/**
 * Detect if text contains HTML entities
 * @param {string} text - Text to check
 * @returns {boolean} True if text contains HTML entities
 */
function hasHtmlEntities(text) {
  if (!text) return false;
  
  // Check for common HTML entities
  const entityPattern = /&(#\d+|#x[0-9A-F]+|[a-z]+);/i;
  return entityPattern.test(text);
}

// Export functions for use in other files
if (typeof window !== 'undefined') {
  window.showError = showError;
  window.showSuccess = showSuccess;
  window.clearMessages = clearMessages;
  window.getPaperlessSettings = getPaperlessSettings;
  window.testPaperlessConnection = testPaperlessConnection;
  window.isValidUrl = isValidUrl;
  window.extractNameFromEmail = extractNameFromEmail;
  window.setButtonLoading = setButtonLoading;
  window.resetButtonLoading = resetButtonLoading;
  window.sendMessageToParent = sendMessageToParent;
  window.closeWindowWithDelay = closeWindowWithDelay;
  window.createCenteredWindow = createCenteredWindow;
  window.makePaperlessRequest = makePaperlessRequest;
  window.decodeHtmlEntities = decodeHtmlEntities;
  window.hasHtmlEntities = hasHtmlEntities;
}
