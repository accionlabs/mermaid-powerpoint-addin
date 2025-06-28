#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// Configuration
const GITHUB_USERNAME = process.env.GITHUB_USERNAME || 'YOUR_USERNAME';
const REPO_NAME = process.env.REPO_NAME || 'REPO_NAME';

console.log('🚀 Preparing deployment files...');
console.log(`GitHub Username: ${GITHUB_USERNAME}`);
console.log(`Repository Name: ${REPO_NAME}`);

// Copy dist to docs for GitHub Pages
const distPath = path.join(__dirname, 'dist');
const docsPath = path.join(__dirname, 'docs');

// Remove existing docs folder if it exists
if (fs.existsSync(docsPath)) {
    fs.rmSync(docsPath, { recursive: true, force: true });
}

// Copy dist to docs
function copyRecursiveSync(src, dest) {
    const exists = fs.existsSync(src);
    const stats = exists && fs.statSync(src);
    const isDirectory = exists && stats.isDirectory();
    
    if (isDirectory) {
        fs.mkdirSync(dest, { recursive: true });
        fs.readdirSync(src).forEach((childItemName) => {
            copyRecursiveSync(path.join(src, childItemName), path.join(dest, childItemName));
        });
    } else {
        fs.copyFileSync(src, dest);
    }
}

copyRecursiveSync(distPath, docsPath);
console.log('✅ Copied dist folder to docs for GitHub Pages');

// Update manifest.xml in both dist and docs
const manifestPath = path.join(__dirname, 'dist', 'manifest.xml');
const docsManifestPath = path.join(__dirname, 'docs', 'manifest.xml');
let manifestContent = fs.readFileSync(manifestPath, 'utf8');

manifestContent = manifestContent.replace(/YOUR_USERNAME/g, GITHUB_USERNAME);
manifestContent = manifestContent.replace(/REPO_NAME/g, REPO_NAME);

fs.writeFileSync(manifestPath, manifestContent);
fs.writeFileSync(docsManifestPath, manifestContent);
console.log('✅ Updated manifest.xml in dist and docs');

// Update index.html (only exists in docs folder)
const docsIndexPath = path.join(__dirname, 'docs', 'index.html');
if (fs.existsSync(docsIndexPath)) {
  let indexContent = fs.readFileSync(docsIndexPath, 'utf8');
  
  indexContent = indexContent.replace(/YOUR_USERNAME/g, GITHUB_USERNAME);
  indexContent = indexContent.replace(/REPO_NAME/g, REPO_NAME);
  
  fs.writeFileSync(docsIndexPath, indexContent);
  console.log('✅ Updated index.html in docs');
} else {
  console.log('ℹ️ No index.html found - skipping (normal for build artifacts)');
}

// Update README.md
const readmePath = path.join(__dirname, 'README.md');
let readmeContent = fs.readFileSync(readmePath, 'utf8');

readmeContent = readmeContent.replace(/YOUR_USERNAME/g, GITHUB_USERNAME);
readmeContent = readmeContent.replace(/REPO_NAME/g, REPO_NAME);

fs.writeFileSync(readmePath, readmeContent);
console.log('✅ Updated README.md');

console.log('\n🎉 Deployment files ready!');
console.log('\n📋 Next steps:');
console.log('1. git add .');
console.log('2. git commit -m "Update URLs for deployment"');
console.log('3. git push origin main');
console.log('4. Go to GitHub → Settings → Pages → Source: "Deploy from a branch" → Branch: main → Folder: /docs');
console.log(`5. Your add-in will be available at: https://${GITHUB_USERNAME}.github.io/${REPO_NAME}/`);