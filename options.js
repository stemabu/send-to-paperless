document.addEventListener('DOMContentLoaded', loadSettings);
document.getElementById('settingsForm').addEventListener('submit', saveSettings);
document.getElementById('addMappingBtn').addEventListener('click', addMapping);

async function loadSettings() {
  const settings = await getPaperlessSettings();

  document.getElementById('paperlessUrl').value = settings.paperlessUrl || '';
  document.getElementById('paperlessToken').value = settings.paperlessToken || '';
  document.getElementById('defaultTags').value = settings.defaultTags || '';
  
  await loadCorrespondents();
  await displayMappings();
}


async function requestSitePermission(url) {
  // Normalize the origin to ensure it ends with /*
  const origin = url.replace(/\/?\*?$/, '/*');

  const hasPermission = await browser.permissions.contains({
    origins: [origin],
  });

  if (hasPermission) {
    // Permission already granted \u2014 safe to save and use the URL.
    return true;
  }

  const granted = await browser.permissions.request({
    origins: [origin],
  });

  // If not granted, it's not safe to save or use the URL,
  // since the user explicitly denied access.
  return granted;
}


async function saveSettings(event) {

  event.preventDefault();

  const paperlessUrl = document.getElementById('paperlessUrl').value.trim();
  const paperlessToken = document.getElementById('paperlessToken').value.trim();
  const defaultTags = document.getElementById('defaultTags').value.trim();

  // Validate URL format
  if (paperlessUrl && !isValidUrl(paperlessUrl)) {
    showStatus('Please enter a valid URL (including http:// or https://)', 'error');
    return;
  }

  // Request permission for the URL if it's provided
  if (paperlessUrl) {
    const permissionGranted = await requestSitePermission(paperlessUrl);
    if (!permissionGranted) {
      showStatus('Permission to access the specified URL was denied. Please allow access to save the settings.', 'error');
      return;
    }
  }

  try {
    await browser.storage.sync.set({
      paperlessUrl: paperlessUrl.replace(/\/$/, ''), // Remove trailing slash
      paperlessToken: paperlessToken,
      defaultTags: defaultTags
    });

    showStatus('Settings saved successfully!', 'success');

    // Test connection if both URL and token are provided
    if (paperlessUrl && paperlessToken) {
      setTimeout(testConnection, 1000);
    }

  } catch (error) {
    showStatus('Error saving settings: ' + error.message, 'error');
    console.error('Error saving settings:', error);
  }
}

async function testConnection() {
  const settings = await getPaperlessSettings();

  const success = await testPaperlessConnection(settings.paperlessUrl, settings.paperlessToken);

  if (success) {
    showStatus('Settings saved and connection test successful!', 'success');
  } else {
    showStatus('Settings saved but connection test failed', 'error');
  }
}

function showStatus(message, type) {
  const statusEl = document.getElementById('statusMessage');
  statusEl.textContent = message;
  statusEl.className = `status-message status-${type}`;
  statusEl.style.display = 'block';

  if (type === 'success') {
    setTimeout(() => {
      statusEl.style.display = 'none';
    }, 3000);
  }
}

// Load correspondents from Paperless-ngx API
async function loadCorrespondents() {
  const settings = await getPaperlessSettings();
  
  if (!settings.paperlessUrl || !settings.paperlessToken) {
    return;
  }
  
  try {
    const response = await fetch(`${settings.paperlessUrl}/api/correspondents/?page_size=1000`, {
      headers: {
        'Authorization': `Token ${settings.paperlessToken}`
      }
    });
    
    if (response.ok) {
      const data = await response.json();
      const select = document.getElementById('newMappingCorrespondent');
      
      while (select.options.length > 1) {
        select.remove(1);
      }
      
      data.results.forEach(correspondent => {
        const option = document.createElement('option');
        option.value = correspondent.id;
        option.textContent = correspondent.name;
        option.dataset.name = correspondent.name;
        select.appendChild(option);
      });
    }
  } catch (error) {
    console.error('Failed to load correspondents:', error);
  }
}

// Display existing mappings
async function displayMappings() {
  const result = await browser.storage.sync.get('emailCorrespondentMapping');
  const mappings = result.emailCorrespondentMapping || [];
  
  const container = document.getElementById('mappingList');
  container.innerHTML = '';
  
  if (mappings.length === 0) {
    container.innerHTML = '<div style="color: #6c757d; font-size: 13px; font-style: italic;">Noch keine Zuordnungen vorhanden</div>';
    return;
  }
  
  mappings.forEach((mapping, index) => {
    const item = document.createElement('div');
    item.className = 'mapping-item';
    item.innerHTML = `
      <div class="mapping-item-content">
        <span class="mapping-email">${escapeHtml(mapping.email)}</span>
        <span class="mapping-arrow">→</span>
        <span class="mapping-correspondent">${escapeHtml(mapping.correspondentName)}</span>
      </div>
      <button class="mapping-delete" data-index="${index}">🗑️ Löschen</button>
    `;
    
    item.querySelector('.mapping-delete').addEventListener('click', () => deleteMapping(index));
    container.appendChild(item);
  });
}

// Escape HTML to prevent XSS
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// Add new mapping
async function addMapping() {
  const emailInput = document.getElementById('newMappingEmail');
  const correspondentSelect = document.getElementById('newMappingCorrespondent');
  
  const email = emailInput.value.trim().toLowerCase();
  const correspondentId = correspondentSelect.value;
  const correspondentName = correspondentSelect.options[correspondentSelect.selectedIndex]?.dataset.name;
  
  if (!email || !correspondentId) {
    showStatus('Bitte E-Mail-Adresse und Korrespondent auswählen', 'error');
    return;
  }
  
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    showStatus('Bitte gültige E-Mail-Adresse eingeben', 'error');
    return;
  }
  
  try {
    const result = await browser.storage.sync.get('emailCorrespondentMapping');
    const mappings = result.emailCorrespondentMapping || [];
    
    if (mappings.some(m => m.email === email)) {
      showStatus('Diese E-Mail-Adresse ist bereits zugeordnet', 'error');
      return;
    }
    
    mappings.push({
      email: email,
      correspondentId: parseInt(correspondentId),
      correspondentName: correspondentName
    });
    
    await browser.storage.sync.set({ emailCorrespondentMapping: mappings });
    
    emailInput.value = '';
    correspondentSelect.selectedIndex = 0;
    await displayMappings();
    
    showStatus('Zuordnung hinzugefügt', 'success');
  } catch (error) {
    console.error('Error adding mapping:', error);
    showStatus('Fehler beim Hinzufügen der Zuordnung', 'error');
  }
}

// Delete mapping
async function deleteMapping(index) {
  try {
    const result = await browser.storage.sync.get('emailCorrespondentMapping');
    const mappings = result.emailCorrespondentMapping || [];
    
    mappings.splice(index, 1);
    
    await browser.storage.sync.set({ emailCorrespondentMapping: mappings });
    await displayMappings();
    
    showStatus('Zuordnung gelöscht', 'success');
  } catch (error) {
    console.error('Error deleting mapping:', error);
    showStatus('Fehler beim Löschen der Zuordnung', 'error');
  }
}