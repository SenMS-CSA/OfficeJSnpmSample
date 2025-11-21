/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Import styles
import './taskpane.css';

// Import version checking functionality
import { 
  CURRENT_VERSION, 
  VERSION_CONFIG,
  compareVersions,
  isUpdateAvailable,
  fetchRemoteVersion,
  shouldCheckVersion,
  saveVersionCheckTimestamp,
  saveRemoteVersionInfo,
  getCachedRemoteVersion
} from '../version-config.js';

console.log("Taskpane.js loaded - waiting for Office.onReady");

Office.onReady((info) => {
  console.log("Office.onReady called", info);
  if (info.host === Office.HostType.Outlook) {
    console.log("Host is Outlook - initializing taskpane");
    // Clear local Office cache
    clearOfficeCache();
    
    // Check for add-in updates
    checkForUpdates();
    
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    
    // Add version information display
    addVersionDisplay();
    
    // Add button for checking recipient domains
    const checkDomainsBtn = document.createElement("button");
    checkDomainsBtn.textContent = "Check Recipient Domains";
    checkDomainsBtn.onclick = checkRecipientDomains;
    checkDomainsBtn.style.marginTop = "10px";
    document.getElementById("app-body").appendChild(checkDomainsBtn);
    
    // Add button for clearing cache manually
    const clearCacheBtn = document.createElement("button");
    clearCacheBtn.textContent = "Clear Office Cache";
    clearCacheBtn.onclick = clearOfficeCache;
    clearCacheBtn.style.marginTop = "10px";
    clearCacheBtn.style.marginLeft = "10px";
    document.getElementById("app-body").appendChild(clearCacheBtn);
    
    // Add button for manual version check
    const checkVersionBtn = document.createElement("button");
    checkVersionBtn.textContent = "Check for Updates";
    checkVersionBtn.onclick = () => checkForUpdates(true);
    checkVersionBtn.style.marginTop = "10px";
    checkVersionBtn.style.marginLeft = "10px";
    document.getElementById("app-body").appendChild(checkVersionBtn);
  }
});

/**
 * Clears Office add-in cache using available methods
 * Note: Direct file system access is not possible from Office add-ins due to security restrictions
 */
function clearOfficeCache() {
  try {
    console.log("Attempting to clear Office add-in cache...");
    
    // Method 1: Clear browser cache (localStorage, sessionStorage)
    if (typeof Storage !== "undefined") {
      localStorage.clear();
      sessionStorage.clear();
      console.log("Browser storage cleared successfully");
    }
    
    // Method 2: Clear any cached data in the add-in context
    if (window.caches) {
      caches.keys().then(function(names) {
        names.forEach(function(name) {
          caches.delete(name);
        });
        console.log("Service worker caches cleared");
      }).catch(function(error) {
        console.log("Cache clearing failed:", error);
      });
    }
    
    // Method 3: Force reload of add-in resources by appending timestamp
    const timestamp = new Date().getTime();
    console.log("Cache clearing timestamp:", timestamp);
    
    // Method 4: Log instructions for manual cache clearing
    console.log("=== Manual Office Cache Clearing Instructions ===");
    console.log("To manually clear Office cache, close Outlook and delete contents of:");
    console.log("Windows: %LOCALAPPDATA%\\Microsoft\\Office\\16.0\\Wef\\");
    console.log("Full path example: C:\\Users\\[Username]\\AppData\\Local\\Microsoft\\Office\\16.0\\Wef\\");
    console.log("Or run this PowerShell command as Administrator:");
    console.log("Remove-Item -Path \"$env:LOCALAPPDATA\\Microsoft\\Office\\16.0\\Wef\\*\" -Recurse -Force");
    
    // Method 5: Try to access Office settings if available
    if (Office && Office.context && Office.context.diagnostics) {
      console.log("Office context available for diagnostics");
      console.log("Office Host:", Office.context.diagnostics.host);
      console.log("Office Version:", Office.context.diagnostics.version);
    }
    
    // Show user notification
    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Cache clearing attempted. Check console for manual instructions if needed.",
      icon: "Icon.80x80",
      persistent: true,
    };
    
    if (Office.context && Office.context.mailbox && Office.context.mailbox.item) {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "CacheClearNotification",
        message
      );
    }
    
    console.log("Cache clearing process completed");
    
  } catch (error) {
    console.error("Error during cache clearing:", error);
    console.log("Manual cache clearing may be required. See console instructions above.");
  }
}

/**
 * Extracts domains from recipient email addresses
 * @param {Array} recipients - Array of recipient objects with emailAddress property
 * @returns {Array} Array of unique domains
 */
function extractDomainsFromRecipients(recipients) {
  const domains = new Set();
  
  if (recipients && recipients.length > 0) {
    recipients.forEach(recipient => {
      if (recipient.emailAddress) {
        const emailParts = recipient.emailAddress.split('@');
        if (emailParts.length === 2) {
          domains.add(emailParts[1].toLowerCase());
        }
      }
    });
  }
  
  return Array.from(domains);
}

/**
 * Gets all recipient domains from To, CC, and BCC fields
 * @param {Function} callback - Callback function to receive the domains
 */
function getAllRecipientDomains(callback) {
  const allDomains = new Set();
  let completedRequests = 0;
  const totalRequests = 3; // To, CC, BCC
  
  function processResults() {
    completedRequests++;
    if (completedRequests === totalRequests) {
      const domains = Array.from(allDomains);
      console.log("All recipient domains:", domains);
      if (callback) callback(domains);
    }
  }
  
  // Get TO recipients
  Office.context.mailbox.item.to.getAsync(function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const msgTo = asyncResult.value;
      console.log("Message being sent to:");
      for (let i = 0; i < msgTo.length; i++) {
        console.log(msgTo[i].displayName + " (" + msgTo[i].emailAddress + ")");
      }
      const toDomains = extractDomainsFromRecipients(msgTo);
      console.log("TO recipients domains:", toDomains);
      toDomains.forEach(domain => allDomains.add(domain));
    } else {
      console.error("Error getting TO recipients:", asyncResult.error);
    }
    processResults();
  });
  
  // Get CC recipients
  Office.context.mailbox.item.cc.getAsync(function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const ccDomains = extractDomainsFromRecipients(asyncResult.value);
      console.log("CC recipients domains:", ccDomains);
      ccDomains.forEach(domain => allDomains.add(domain));
    } else {
      console.error("Error getting CC recipients:", asyncResult.error);
    }
    processResults();
  });
  
  // Get BCC recipients
  Office.context.mailbox.item.bcc.getAsync(function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const bccDomains = extractDomainsFromRecipients(asyncResult.value);
      console.log("BCC recipients domains:", bccDomains);
      bccDomains.forEach(domain => allDomains.add(domain));
    } else {
      console.error("Error getting BCC recipients:", asyncResult.error);
    }
    processResults();
  });
}

/**
 * Function to check recipient domains manually from taskpane
 */
function checkRecipientDomains() {
  const resultDiv = document.getElementById("item-subject");
  resultDiv.innerHTML = "<b>Checking recipient domains...</b><br>";
  
  getAllRecipientDomains(function(domains) {
    resultDiv.innerHTML = "<b>Recipient Domains Found:</b><br>";
    if (domains.length > 0) {
      domains.forEach(domain => {
        resultDiv.appendChild(document.createTextNode("• " + domain));
        resultDiv.appendChild(document.createElement("br"));
      });
    } else {
      resultDiv.appendChild(document.createTextNode("No recipients found"));
    }
  });
}

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}

/**
 * Add version display to the UI
 */
function addVersionDisplay() {
  const versionContainer = document.createElement("div");
  versionContainer.id = "version-container";
  versionContainer.style.cssText = `
    margin-top: 15px;
    padding: 10px;
    border: 1px solid #e0e0e0;
    border-radius: 4px;
    background-color: #f9f9f9;
    font-size: 12px;
  `;
  
  versionContainer.innerHTML = `
    <div><strong>Add-in Version Info:</strong></div>
    <div id="current-version">Current: ${CURRENT_VERSION}</div>
    <div id="remote-version">Latest: <span style="color: #666;">Checking...</span></div>
    <div id="version-status" style="margin-top: 5px;"></div>
  `;
  
  document.getElementById("app-body").appendChild(versionContainer);
  
  // Update with cached version info if available
  updateVersionDisplay();
}

/**
 * Update version display with current information
 */
function updateVersionDisplay() {
  const cachedVersion = getCachedRemoteVersion();
  if (cachedVersion) {
    const remoteVersionElement = document.getElementById("remote-version");
    const statusElement = document.getElementById("version-status");
    
    if (remoteVersionElement) {
      remoteVersionElement.innerHTML = `Latest: ${cachedVersion.version}`;
    }
    
    if (statusElement) {
      const updateAvailable = isUpdateAvailable(CURRENT_VERSION, cachedVersion.version);
      if (updateAvailable) {
        statusElement.innerHTML = `<span style="color: #d73502; font-weight: bold;">⚠️ Update Available!</span>`;
      } else {
        statusElement.innerHTML = `<span style="color: #107c10;">✅ Up to date</span>`;
      }
    }
  }
}

/**
 * Check for add-in updates
 * @param {boolean} forceCheck - Force check even if recently checked
 */
async function checkForUpdates(forceCheck = false) {
  try {
    console.log('Checking for add-in updates...');
    console.log('Current version:', CURRENT_VERSION);
    
    // Check if we should perform the version check
    if (!forceCheck && !shouldCheckVersion()) {
      console.log('Version check skipped - checked recently');
      updateVersionDisplay();
      return;
    }
    
    // Update UI to show checking status
    const remoteVersionElement = document.getElementById("remote-version");
    if (remoteVersionElement) {
      remoteVersionElement.innerHTML = 'Latest: <span style="color: #666;">Checking...</span>';
    }
    
    // Fetch remote version information
    const remoteVersionInfo = await fetchRemoteVersion();
    console.log('Remote version info:', remoteVersionInfo);
    
    // Save the version info and timestamp
    saveRemoteVersionInfo(remoteVersionInfo);
    saveVersionCheckTimestamp();
    
    // Update the display
    updateVersionDisplay();
    
    // Check if update is available
    const updateAvailable = isUpdateAvailable(CURRENT_VERSION, remoteVersionInfo.version);
    console.log('Update available:', updateAvailable);
    
    if (updateAvailable) {
      showUpdateAlert(remoteVersionInfo);
    } else if (forceCheck) {
      // Show "up to date" message only for manual checks
      showUpToDateMessage();
    }
    
  } catch (error) {
    console.error('Error checking for updates:', error);
    
    const remoteVersionElement = document.getElementById("remote-version");
    if (remoteVersionElement) {
      remoteVersionElement.innerHTML = 'Latest: <span style="color: #d73502;">Check failed</span>';
    }
    
    if (forceCheck) {
      showErrorMessage('Failed to check for updates. Please try again later.');
    }
  }
}

/**
 * Show update alert to user
 * @param {Object} versionInfo - Remote version information
 */
function showUpdateAlert(versionInfo) {
  const updateMessage = `🔄 Add-in Update Available!

Current Version: ${CURRENT_VERSION}
Latest Version: ${versionInfo.version}

${versionInfo.changelog ? 'What\'s New:\n• ' + versionInfo.changelog.join('\n• ') : ''}

${versionInfo.securityUpdate ? '⚠️ This is a security update and is recommended.' : ''}

Would you like to download the update?`;

  if (confirm(updateMessage)) {
    // Open download link
    if (versionInfo.downloadUrl) {
      window.open(versionInfo.downloadUrl, '_blank');
    } else {
      alert('Please contact your administrator for the update.');
    }
  }
  
  // Also show Office notification
  try {
    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `Add-in update available: v${versionInfo.version}`,
      icon: "Icon.80x80",
      persistent: true,
    };
    
    if (Office.context && Office.context.mailbox && Office.context.mailbox.item) {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "UpdateAvailableNotification",
        message
      );
    }
  } catch (error) {
    console.warn('Could not show Office notification:', error);
  }
}

/**
 * Show "up to date" message
 */
function showUpToDateMessage() {
  alert(`✅ Add-in is up to date!\n\nCurrent Version: ${CURRENT_VERSION}`);
  
  try {
    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `Add-in is up to date (v${CURRENT_VERSION})`,
      icon: "Icon.80x80",
      persistent: false,
    };
    
    if (Office.context && Office.context.mailbox && Office.context.mailbox.item) {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "UpToDateNotification",
        message
      );
    }
  } catch (error) {
    console.warn('Could not show Office notification:', error);
  }
}

/**
 * Show error message
 * @param {string} errorMsg - Error message to display
 */
function showErrorMessage(errorMsg) {
  alert(`❌ ${errorMsg}`);
  
  try {
    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: errorMsg,
      icon: "Icon.80x80",
      persistent: true,
    };
    
    if (Office.context && Office.context.mailbox && Office.context.mailbox.item) {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "ErrorNotification",
        message
      );
    }
  } catch (error) {
    console.warn('Could not show Office notification:', error);
  }
}
