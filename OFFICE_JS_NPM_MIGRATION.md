@"
# Outlook Add-in with Office.js from npm Package

This Outlook add-in demonstrates how to use Office.js from the npm package (`@microsoft/office-js`) instead of the CDN, with recipient domain extraction, version checking, and cache management features.

## Key Features

- ✅ **Office.js from npm Package** - Works offline, version-locked, no CDN dependency
- ✅ **Recipient Domain Extraction** - Extracts and displays domains from To, CC, BCC fields
- ✅ **Version Checking** - Automatic update notifications when new versions available
- ✅ **Cache Management** - Automatic Office cache clearing on startup
- ✅ **Webpack Build** - Modern build system with hot module replacement

## Quick Start

### Prerequisites
- Node.js (v14 or higher)
- npm (v6 or higher)
- Outlook (Desktop, Web, or Mac)

### Installation

1. **Clone the repository:**
   \`\`\`bash
   git clone https://github.com/SenMS-CSA/OfficeJSnpmSample.git
   cd OfficeJSnpmSample
   \`\`\`

2. **Install dependencies:**
   \`\`\`bash
   npm install
   \`\`\`

3. **Build and start:**
   \`\`\`bash
   npm start
   \`\`\`

   This will:
   - Copy Office.js files from npm package
   - Build the add-in
   - Start the dev server at https://localhost:3000
   - Sideload the add-in into Outlook

## Office.js npm Package Integration

### Why npm Package Instead of CDN?

| Benefit | Description |
|---------|-------------|
| **Offline Development** | No internet required after \`npm install\` |
| **Version Control** | Exact version locked in package.json |
| **Reproducible Builds** | Same Office.js version across all environments |
| **TypeScript Support** | Type definitions included in package |
| **Corporate Firewall Friendly** | No external CDN access needed |

### How It Works

#### 1. Pre-build Script (\`copy-office-js.js\`)

Automatically runs before every build:

\`\`\`bash
node copy-office-js.js && webpack --mode development
\`\`\`

**What it does:**
- Copies all 729 Office.js files from \`node_modules/@microsoft/office-js/dist/\`
- Creates \`public/office-js/\` directory with complete Office.js distribution
- Preserves directory structure for dynamic dependency loading

#### 2. Webpack Configuration

**Key Settings:**

\`\`\`javascript
// Exclude script tags from html-loader processing
{
  test: /\.html\$/,
  use: {
    loader: "html-loader",
    options: {
      sources: {
        list: [
          { tag: "img", attribute: "src", type: "src" },
          { tag: "link", attribute: "href", type: "src" }
          // Script tags excluded - office.js won't be bundled/hashed
        ]
      }
    }
  }
}

// Copy office-js directory to dist
new CopyWebpackPlugin({
  patterns: [
    { from: "public/office-js", to: "office-js" }
  ]
})

// Use inline loader to prevent source processing
new HtmlWebpackPlugin({
  template: "!!html-loader?{\"sources\":false}!./src/taskpane/taskpane.html"
})
\`\`\`

**Why these settings?**
- Office.js requires exact filename "office.js" (no webpack hashing)
- Office.js dynamically loads 728 dependencies from same directory
- Webpack must leave office.js references unchanged

#### 3. HTML Files

\`\`\`html
<!-- Uses npm package (not CDN) -->
<script type="text/javascript" src="./office-js/office.js"></script>
\`\`\`

#### 4. Build Output

\`\`\`
dist/
├── office-js/           (729 files, 83.4 MB from npm package)
│   ├── office.js        (Main library, 58 KB)
│   ├── office.debug.js
│   ├── MicrosoftAjax.js
│   ├── outlook-win32-16.00.js
│   └── [locales]/       (200+ language folders)
├── taskpane.html
├── taskpane.js          (62.5 KB with bundled CSS)
├── commands.html
├── commands.js
└── manifest.xml
\`\`\`

## Build Flow

\`\`\`
npm run build:dev
  ↓
1. copy-office-js.js
   - Copies node_modules/@microsoft/office-js/dist/ → public/office-js/
   - 729 files (83.4 MB)
  ↓
2. Webpack
   - Processes HTML templates (preserves office.js script tags)
   - Transpiles JavaScript (babel-loader)
   - Bundles CSS into JS (css-loader + style-loader)
  ↓
3. CopyWebpackPlugin
   - Copies public/office-js/ → dist/office-js/
   - Copies assets and manifest
  ↓
4. Output to dist/
\`\`\`

## Runtime Loading

\`\`\`
Outlook launches add-in
  ↓
Browser loads dist/taskpane.html
  ↓
Loads office.js from ./office-js/office.js
  ↓
Office.js validates filename === "office.js" ✓
  ↓
Office.js loads dependencies:
  - MicrosoftAjax.js
  - outlook-win32-16.00.js
  - en-us/outlook_strings.js
  ↓
Creates global Office object
  ↓
Loads polyfill.js & taskpane.js
  ↓
Office.onReady() fires → Add-in initialized
\`\`\`

## Available Scripts

| Command | Description |
|---------|-------------|
| \`npm install\` | Install dependencies including @microsoft/office-js |
| \`npm run build:dev\` | Build for development with source maps |
| \`npm run build\` | Build for production (minified) |
| \`npm start\` | Build, serve, and sideload into Outlook |
| \`npm stop\` | Stop debugging and unload add-in |
| \`npm run clear-cache\` | Clear Office cache (PowerShell) |
| \`npm run validate\` | Validate manifest.xml |

## Project Structure

\`\`\`
OLRecpTest/
├── src/
│   ├── commands/
│   │   ├── commands.html         (Ribbon commands page)
│   │   └── commands.js           (Domain extraction logic)
│   ├── taskpane/
│   │   ├── taskpane.html         (Task pane UI)
│   │   ├── taskpane.js           (Main application logic)
│   │   └── taskpane.css          (Styles)
│   └── version-config.js         (Version checking utilities)
├── public/
│   ├── office-js/                (Generated by copy-office-js.js)
│   └── version.json              (Remote version info)
├── dist/                         (Build output)
├── copy-office-js.js             (Pre-build script)
├── webpack.config.js             (Webpack configuration)
├── manifest.xml                  (Add-in manifest)
└── package.json
\`\`\`

## Key Files Explained

### copy-office-js.js
Pre-build script that copies Office.js from npm package to public folder.

### webpack.config.js
- Excludes script tags from html-loader
- Uses CopyWebpackPlugin for office-js directory
- CSS bundled into JavaScript via style-loader

### package.json
\`\`\`json
{
  "dependencies": {
    "@microsoft/office-js": "^1.1.110"
  },
  "scripts": {
    "build:dev": "node copy-office-js.js && webpack --mode development"
  }
}
\`\`\`

## .gitignore

\`\`\`
node_modules/
dist/
public/office-js/        # Generated from npm package (83.4 MB)
\`\`\`

**Why ignore public/office-js/?**
- Regenerated from npm package on every build
- Saves ~83 MB in repository
- Developers run \`npm install\` then \`npm run build:dev\` to recreate

## Troubleshooting

### Error: "Office Web Extension script library file name should be office.js"

**Cause:** Webpack bundled office.js with a content hash

**Solution:** Verify webpack.config.js excludes script tags from html-loader:
\`\`\`javascript
sources: {
  list: [
    { tag: "img", ... },
    { tag: "link", ... }
    // NO script tag entry
  ]
}
\`\`\`

### Error: "MicrosoftAjax.js is not loaded successfully"

**Cause:** Office.js dependencies not copied

**Solution:**
\`\`\`bash
# Verify files were copied
ls public/office-js/  # Should show 729 files
ls dist/office-js/    # Should show 729 files

# If missing, run:
node copy-office-js.js
npm run build:dev
\`\`\`

### CSS Not Loading

**Cause:** CSS not imported or loaders missing

**Solution:**
1. Verify import in taskpane.js: \`import './taskpane.css';\`
2. Verify webpack has css-loader and style-loader installed
3. Rebuild: \`npm run build:dev\`

## Migration from CDN to npm Package

See [OFFICE_JS_NPM_MIGRATION.md](OFFICE_JS_NPM_MIGRATION.md) for complete step-by-step migration guide.

**Quick Checklist:**
- [ ] Install: \`npm install @microsoft/office-js\`
- [ ] Create: \`copy-office-js.js\` script
- [ ] Update: \`package.json\` scripts
- [ ] Update: \`webpack.config.js\` (html-loader, CopyWebpackPlugin)
- [ ] Update: HTML files (CDN URL → \`./office-js/office.js\`)
- [ ] Update: \`.gitignore\` (add \`public/office-js/\`)
- [ ] Test: \`npm run build:dev && npm start\`

## Additional Documentation

- [PROJECT_STRUCTURE.md](PROJECT_STRUCTURE.md) - Complete project architecture
- [OFFICE_JS_NPM_MIGRATION.md](OFFICE_JS_NPM_MIGRATION.md) - Detailed migration guide
- [VERSION_CHECKING.md](VERSION_CHECKING.md) - Version management system
- [README_RECIPIENT_DOMAINS.md](README_RECIPIENT_DOMAINS.md) - Domain extraction feature

## Features

### Recipient Domain Extraction
- Extracts domains from To, CC, BCC recipients
- Shows notification with unique domains
- Detailed console logging

### Version Checking
- Compares local vs remote version
- Shows update notification
- Caches checks (24 hour interval)

### Cache Management
- Clears localStorage and sessionStorage on startup
- PowerShell script for full Office cache clearing

## Browser Support

- Microsoft Edge (Chromium)
- Internet Explorer 11
- Chrome, Firefox, Safari (via Outlook Web)

## License

MIT

## Resources

- [Office Add-ins Documentation](https://learn.microsoft.com/office/dev/add-ins/)
- [@microsoft/office-js npm package](https://www.npmjs.com/package/@microsoft/office-js)
- [Webpack Documentation](https://webpack.js.org/)
"@ > README.md