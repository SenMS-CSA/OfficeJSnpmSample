# Office Add-in Cache Clearing

This add-in includes cache clearing functionality to help with development and troubleshooting.

## Automatic Cache Clearing

The add-in automatically attempts to clear cache when it starts up:
- **Triggered**: On `Office.onReady()` in both taskpane and commands
- **Methods**: Clears localStorage, sessionStorage, and service worker caches
- **Limitations**: Cannot directly access file system due to security restrictions

## Manual Cache Clearing Options

### 1. In-App Button (Task Pane)
- Open the task pane
- Click "Clear Office Cache" button
- Check browser console for results and instructions

### 2. NPM Scripts
```bash
# Interactive cache clearing (prompts if Outlook is running)
npm run clear-cache

# Force cache clearing (automatically closes/restarts Outlook)
npm run clear-cache-force
```

### 3. PowerShell Script
```powershell
# Run directly
.\clear-office-cache.ps1

# Force mode (auto-restart Outlook)
.\clear-office-cache.ps1 -Force
```

### 4. Manual File System Cleaning
Close Outlook completely, then delete contents of:
```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

## Cache Locations Cleared

### JavaScript Methods (Security Limited):
- `localStorage` - Browser local storage
- `sessionStorage` - Browser session storage  
- Service Worker caches
- Console instructions for manual clearing

### PowerShell Script (Full Access):
- `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*` - Main Office WEF cache
- `%TEMP%\OfficeAddins*` - Temporary Office add-in files
- `%LOCALAPPDATA%\Microsoft\Office\16.0\Addins\WebCache\*` - Edge WebView2 cache

## When to Clear Cache

### Development Scenarios:
- Add-in not reflecting code changes
- Manifest changes not taking effect
- Unexpected behavior after updates
- Debugging connectivity issues

### Production Scenarios:
- User reports outdated add-in behavior
- Add-in appears corrupted or unresponsive
- After Office updates

## Console Output

The cache clearing functions provide detailed console output:
- ✓ Success messages for completed operations
- ✗ Error messages with troubleshooting hints
- Manual instructions for file system access
- Office diagnostic information when available

## Best Practices

1. **Close Outlook First**: For most effective cache clearing
2. **Run as Administrator**: PowerShell script works best with elevated privileges
3. **Check Console**: Always review console output for errors or manual steps
4. **Restart Outlook**: After manual cache clearing
5. **Use Force Mode Carefully**: Only in development environments

## Troubleshooting

### If automatic clearing doesn't work:
1. Check browser console for error messages
2. Try the manual PowerShell script
3. Close Outlook and manually delete cache folders
4. Restart Outlook and test

### If PowerShell script fails:
1. Run PowerShell as Administrator
2. Check execution policy: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
3. Ensure Outlook is completely closed
4. Manually delete cache directories

## Security Notes

- JavaScript cache clearing is limited by browser security
- PowerShell script requires appropriate execution policies
- File system access requires proper permissions
- Always backup important data before clearing caches