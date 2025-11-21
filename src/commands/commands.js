/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Import version checking functionality
import { 
  CURRENT_VERSION, 
  fetchRemoteVersion,
  isUpdateAvailable,
  shouldCheckVersion,
  saveVersionCheckTimestamp,
  saveRemoteVersionInfo
} from '../version-config.js';

Office.onReady(() => {
  // Clear Office cache on startup
  clearOfficeCache();
  
  // Check for updates in background
  checkForUpdatesBackground();
  
  // If needed, Office.js is ready to be called.
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
    
    // Method 3: Log instructions for manual cache clearing
    console.log("=== Manual Office Cache Clearing Instructions ===");
    console.log("To manually clear Office cache, close Outlook and delete contents of:");
    console.log("Windows: %LOCALAPPDATA%\\Microsoft\\Office\\16.0\\Wef\\");
    console.log("Full path example: C:\\Users\\[Username]\\AppData\\Local\\Microsoft\\Office\\16.0\\Wef\\");
    console.log("Or run this PowerShell command as Administrator:");
    console.log("Remove-Item -Path \"$env:LOCALAPPDATA\\Microsoft\\Office\\16.0\\Wef\\*\" -Recurse -Force");
    
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
  if (Office.context.mailbox.item.cc) {
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
  } else {
    processResults();
  }
  
  // Get BCC recipients
  if (Office.context.mailbox.item.bcc) {
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
  } else {
    processResults();
  }
}

/**
 * Shows recipient domains when the action button is clicked
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  console.log("Action button clicked - checking recipient domains");
  
  getAllRecipientDomains(function(domains) {
    let message;
    if (domains.length > 0) {
      message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `Recipient domains found: ${domains.join(', ')}`,
        icon: "Icon.80x80",
        persistent: true,
      };
      console.log("Recipient domains found:", domains);
    } else {
      message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "No recipients found or unable to access recipient information.",
        icon: "Icon.80x80",
        persistent: true,
      };
      console.log("No recipients found");
    }

    // Show a notification message.
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "RecipientDomainsNotification",
      message
    );
  });

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

/**
 * Background version check (silent, no UI alerts)
 */
async function checkForUpdatesBackground() {
  try {
    console.log('Background version check started...');
    console.log('Current version:', CURRENT_VERSION);
    
    // Check if we should perform the version check
    if (!shouldCheckVersion()) {
      console.log('Background version check skipped - checked recently');
      return;
    }
    
    // Fetch remote version information
    const remoteVersionInfo = await fetchRemoteVersion();
    console.log('Background check - Remote version info:', remoteVersionInfo);
    
    // Save the version info and timestamp
    saveRemoteVersionInfo(remoteVersionInfo);
    saveVersionCheckTimestamp();
    
    // Check if update is available (log only, no UI)
    const updateAvailable = isUpdateAvailable(CURRENT_VERSION, remoteVersionInfo.version);
    console.log('Background check - Update available:', updateAvailable);
    
    if (updateAvailable) {
      console.log(`🔄 Update available: ${CURRENT_VERSION} → ${remoteVersionInfo.version}`);
      console.log('User will be notified when they open the task pane.');
    }
    
  } catch (error) {
    console.error('Background version check error:', error);
  }
}
