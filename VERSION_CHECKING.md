# Add-in Version Checking System

This Office add-in now includes comprehensive version checking functionality that compares the local add-in version with a deployed/remote version and alerts users when updates are available.

## ✅ Features Implemented

### 1. **Automatic Version Checking**
- **On Startup**: Version check runs automatically when the add-in loads
- **Background Check**: Silent check in commands.js (no UI interruption)
- **Smart Timing**: Only checks every 24 hours to avoid excessive requests
- **Caching**: Stores version information locally for quick access

### 2. **Version Display in Task Pane**
- **Current Version**: Shows the installed add-in version (`0.0.1`)
- **Latest Version**: Displays the remote/deployed version (`0.0.2`)
- **Status Indicator**: Visual indicator showing update availability
- **Real-time Updates**: Version display updates as checks complete

### 3. **User Update Alerts**
- **Automatic Alerts**: Popup when update is detected
- **Detailed Information**: Shows current vs. latest version
- **Changelog**: Lists what's new in the update
- **Security Updates**: Special highlighting for security-related updates
- **Download Links**: Direct access to update downloads

### 4. **Manual Version Checking**
- **"Check for Updates" Button**: Users can manually trigger version checks
- **Force Check**: Bypasses timing restrictions for immediate checking
- **Up-to-Date Confirmation**: Shows confirmation when no updates available

## 📋 Version Information Display

### UI Elements Added:
```
Add-in Version Info:
Current: 0.0.1
Latest: 0.0.2
⚠️ Update Available!
```

### Status Indicators:
- ✅ **Up to date** (Green) - No updates available
- ⚠️ **Update Available!** (Red) - Update detected
- **Checking...** (Gray) - Version check in progress
- **Check failed** (Red) - Unable to reach version server

## 🔧 Configuration

### Version Configuration (`src/version-config.js`):
```javascript
// Current add-in version
CURRENT_VERSION = '0.0.1'

// Remote version endpoint
remoteVersionUrl: 'https://localhost:3000/version.json'

// Check interval (24 hours)
checkInterval: 24 * 60 * 60 * 1000
```

### Version Information Format (`public/version.json`):
```json
{
  "version": "0.0.2",
  "releaseDate": "2025-09-30", 
  "updateRequired": false,
  "securityUpdate": false,
  "downloadUrl": "https://example.com/download/latest",
  "changelog": [
    "Enhanced recipient domain extraction",
    "Added cache clearing functionality",
    "Improved version checking",
    "Performance optimizations"
  ]
}
```

## 🔄 Version Checking Flow

### 1. **Automatic Check (OnReady)**
```
Add-in starts → Check timing → Fetch remote version → Compare versions → Show alert if needed
```

### 2. **Manual Check (Button)**
```
User clicks "Check for Updates" → Force check → Fetch version → Show result (update or up-to-date)
```

### 3. **Background Check (Commands)**
```
Commands load → Silent version check → Log results → Cache for task pane display
```

## 📝 Version Comparison Logic

### Version String Comparison:
- Supports semantic versioning (e.g., `1.2.3`)
- Compares major.minor.patch numbers
- Handles different version string lengths
- Returns: `1` (newer), `-1` (older), `0` (same)

### Update Detection:
```javascript
// Example: 0.0.1 vs 0.0.2
isUpdateAvailable('0.0.1', '0.0.2') // Returns: true
```

## 🔔 User Alert Types

### 1. **Update Available Alert**
```
🔄 Add-in Update Available!

Current Version: 0.0.1
Latest Version: 0.0.2

What's New:
• Enhanced recipient domain extraction
• Added cache clearing functionality
• Improved version checking
• Performance optimizations

Would you like to download the update?
```

### 2. **Up-to-Date Confirmation**
```
✅ Add-in is up to date!

Current Version: 0.0.1
```

### 3. **Office Notifications**
- Integration with Outlook notification system
- Persistent notifications for updates
- Temporary notifications for confirmations

## 🏗️ Technical Implementation

### Files Created/Modified:

**New Files:**
- ✅ `src/version-config.js` - Version checking configuration and utilities
- ✅ `public/version.json` - Remote version information endpoint
- ✅ `src/mock-version-api.html` - Testing endpoint

**Modified Files:**
- ✅ `src/taskpane/taskpane.js` - Added UI elements and update alerts
- ✅ `src/commands/commands.js` - Added background version checking

### Key Functions:
- `checkForUpdates()` - Main version checking with UI
- `checkForUpdatesBackground()` - Silent background checking
- `compareVersions()` - Version string comparison
- `fetchRemoteVersion()` - Remote version retrieval
- `showUpdateAlert()` - User update notifications

## 🔒 Security & Privacy

### Data Handling:
- **Local Storage**: Caches version info and check timestamps
- **Network Requests**: Only to configured version endpoint
- **No Personal Data**: Version checks don't transmit user information

### Error Handling:
- **Network Failures**: Falls back to demo version info
- **Invalid Responses**: Graceful error handling with user feedback
- **Timeout Protection**: 5-second timeout on version requests

## 🧪 Testing the System

### Manual Testing:
1. **Open Task Pane**: See current version display
2. **Click "Check for Updates"**: Trigger manual version check
3. **View Console**: See detailed version checking logs
4. **Update version.json**: Change remote version to test alerts

### Demo Configuration:
- **Current Version**: `0.0.1` (in package.json)
- **Remote Version**: `0.0.2` (in version.json)
- **Result**: Update alert will be shown

### Production Setup:
1. Update `remoteVersionUrl` to your actual version API
2. Deploy version.json to your server
3. Update version numbers as needed
4. Test with real deployment scenarios

## 🔄 Deployment Integration

### CI/CD Integration:
```bash
# Update version in package.json
npm version patch

# Update remote version.json
# Deploy new add-in version
# Users will be automatically notified
```

The version checking system provides a professional update management experience for Office add-in users with comprehensive alerting, caching, and user-friendly interfaces.