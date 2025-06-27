const fs = require('fs');
const path = require('path');

// Copy manifest to dist folder so it's served by webpack
const manifestPath = path.join(__dirname, 'manifest.xml');
const distPath = path.join(__dirname, 'dist');
const distManifestPath = path.join(distPath, 'manifest.xml');

// Ensure dist folder exists
if (!fs.existsSync(distPath)) {
    fs.mkdirSync(distPath);
}

// Copy manifest to dist
fs.copyFileSync(manifestPath, distManifestPath);

console.log('Manifest copied to dist folder');
console.log('');
console.log('Now you can sideload using:');
console.log('https://localhost:3000/manifest.xml');
console.log('');
console.log('Or try the Office Developer Tools:');
console.log('npx office-addin-debugging start manifest.xml desktop');