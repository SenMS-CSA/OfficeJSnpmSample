const fs = require('fs');
const path = require('path');

// Source directory containing all Office.js files
const sourceDir = path.join(__dirname, 'node_modules', '@microsoft', 'office-js', 'dist');

// Destination directory (only public/office-js is needed)
const destDir = path.join(__dirname, 'public', 'office-js');

// Create destination directory if it doesn't exist
if (!fs.existsSync(destDir)) {
  fs.mkdirSync(destDir, { recursive: true });
}

// Copy entire dist directory to preserve file structure for Office.js dependencies
function copyDirectory(src, dest) {
  const entries = fs.readdirSync(src, { withFileTypes: true });
  
  entries.forEach(entry => {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);
    
    if (entry.isDirectory()) {
      if (!fs.existsSync(destPath)) {
        fs.mkdirSync(destPath, { recursive: true });
      }
      copyDirectory(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  });
}

try {
  copyDirectory(sourceDir, destDir);
  console.log(`Successfully copied Office.js to ${destDir}`);
} catch (err) {
  console.error('Error copying Office.js files:', err);
  process.exit(1);
}
