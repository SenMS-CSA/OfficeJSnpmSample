/*
 * Version Configuration for Office Add-in
 * This file contains version information and update checking logic
 */

// Current add-in version (should match package.json)
export const CURRENT_VERSION = '0.0.1';

// Version check configuration
export const VERSION_CONFIG = {
  // Remote version endpoint (replace with your actual endpoint)
  remoteVersionUrl: 'https://localhost:3000/version.json',
  
  // Fallback version info for demo/development
  fallbackVersionInfo: {
    version: '0.0.2',
    releaseDate: '2025-09-30',
    updateRequired: false,
    securityUpdate: false,
    downloadUrl: 'https://example.com/download/latest',
    changelog: [
      'Enhanced recipient domain extraction with better error handling',
      'Added comprehensive cache clearing functionality', 
      'Improved version checking and update notifications',
      'Performance optimizations and bug fixes'
    ]
  },
  
  // Check interval (in milliseconds) - 24 hours
  checkInterval: 24 * 60 * 60 * 1000,
  
  // Local storage keys
  storageKeys: {
    lastCheck: 'addin_last_version_check',
    remoteVersion: 'addin_remote_version',
    updateDismissed: 'addin_update_dismissed'
  }
};

/**
 * Compare two version strings
 * @param {string} version1 - First version (e.g., "1.0.0")
 * @param {string} version2 - Second version (e.g., "1.0.1")
 * @returns {number} - Returns 1 if version1 > version2, -1 if version1 < version2, 0 if equal
 */
export function compareVersions(version1, version2) {
  const v1Parts = version1.split('.').map(Number);
  const v2Parts = version2.split('.').map(Number);
  
  const maxLength = Math.max(v1Parts.length, v2Parts.length);
  
  for (let i = 0; i < maxLength; i++) {
    const v1Part = v1Parts[i] || 0;
    const v2Part = v2Parts[i] || 0;
    
    if (v1Part > v2Part) return 1;
    if (v1Part < v2Part) return -1;
  }
  
  return 0;
}

/**
 * Check if an update is available
 * @param {string} currentVersion - Current add-in version
 * @param {string} remoteVersion - Remote/latest version
 * @returns {boolean} - True if update is available
 */
export function isUpdateAvailable(currentVersion, remoteVersion) {
  return compareVersions(remoteVersion, currentVersion) > 0;
}

/**
 * Fetch remote version information
 * @returns {Promise<Object>} - Promise resolving to version information
 */
export async function fetchRemoteVersion() {
  try {
    // Try to fetch from remote endpoint
    const response = await fetch(VERSION_CONFIG.remoteVersionUrl, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json'
      },
      // Add timeout
      signal: AbortSignal.timeout(5000)
    });
    
    if (response.ok) {
      const versionInfo = await response.json();
      console.log('Remote version fetched successfully:', versionInfo);
      return versionInfo;
    } else {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
  } catch (error) {
    console.warn('Failed to fetch remote version, using fallback:', error.message);
    
    // Return fallback version info for demo purposes
    return VERSION_CONFIG.fallbackVersionInfo;
  }
}

/**
 * Check if we should perform a version check
 * @returns {boolean} - True if check should be performed
 */
export function shouldCheckVersion() {
  try {
    const lastCheck = localStorage.getItem(VERSION_CONFIG.storageKeys.lastCheck);
    if (!lastCheck) return true;
    
    const timeSinceLastCheck = Date.now() - parseInt(lastCheck);
    return timeSinceLastCheck > VERSION_CONFIG.checkInterval;
  } catch (error) {
    console.warn('Error checking version check timing:', error);
    return true; // Default to checking if there's an error
  }
}

/**
 * Save version check timestamp
 */
export function saveVersionCheckTimestamp() {
  try {
    localStorage.setItem(VERSION_CONFIG.storageKeys.lastCheck, Date.now().toString());
  } catch (error) {
    console.warn('Error saving version check timestamp:', error);
  }
}

/**
 * Save remote version information
 * @param {Object} versionInfo - Version information to save
 */
export function saveRemoteVersionInfo(versionInfo) {
  try {
    localStorage.setItem(VERSION_CONFIG.storageKeys.remoteVersion, JSON.stringify(versionInfo));
  } catch (error) {
    console.warn('Error saving remote version info:', error);
  }
}

/**
 * Get cached remote version information
 * @returns {Object|null} - Cached version info or null
 */
export function getCachedRemoteVersion() {
  try {
    const cached = localStorage.getItem(VERSION_CONFIG.storageKeys.remoteVersion);
    return cached ? JSON.parse(cached) : null;
  } catch (error) {
    console.warn('Error getting cached remote version:', error);
    return null;
  }
}