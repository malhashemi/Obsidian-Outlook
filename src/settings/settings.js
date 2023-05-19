// Load settings when the add-in is initialized

/* eslint-disable no-useless-escape */
/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */

const Office = window.Office;
Office.onReady(() => {
    loadSettings();
  });
  
  // Save settings to Office.js Settings API
  function saveSettings(event) {
    event.preventDefault();
    const settings = Office.context.roamingSettings;
    settings.set('vaultName', document.getElementById('vaultName').value);
    settings.set('attachmentFolder', document.getElementById('attachmentFolder').value);
    settings.set('defaultTags', document.getElementById('defaultTags').value);
    settings.set('defaultFolder', document.getElementById('defaultFolder').value);
    console.log(settings);
    settings.saveAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        showMessageBar('Settings saved successfully!');
      } else {
        console.error('Failed to save settings.', asyncResult.error);
        showMessageBar('Failed to save settings.');
      }
    });
  }
  
  // Load settings from Office.js Settings API
  function loadSettings() {
    const settings = Office.context.roamingSettings;
    console.log(settings);
    document.getElementById('vaultName').value = settings.get('vaultName') || '';
    document.getElementById('attachmentFolder').value = settings.get('attachmentFolder') || '';
    document.getElementById('defaultTags').value = settings.get('defaultTags') || '';
    document.getElementById('defaultFolder').value = settings.get('defaultFolder') || '';
  }
  
  // Attach the saveSettings function to the form submit event
  document.getElementById('settingsForm').addEventListener('submit', saveSettings);


  function showMessageBar(message) {
    const messageBarContainer = document.getElementById('messageBarContainer');
    const messageBarText = document.getElementById('messageBarText');
    messageBarText.textContent = message;
    messageBarContainer.style.display = 'block';
    setTimeout(() => {
      messageBarContainer.style.display = 'none';
    }, 5000); // Hide the message after 5 seconds
  }