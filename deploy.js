#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// Configuration
const GITHUB_USERNAME = process.env.GITHUB_USERNAME || 'YOUR_USERNAME';
const REPO_NAME = process.env.REPO_NAME || 'REPO_NAME';

console.log('🚀 Preparing deployment files...');
console.log(`GitHub Username: ${GITHUB_USERNAME}`);
console.log(`Repository Name: ${REPO_NAME}`);

// Update manifest-production.xml
const manifestPath = path.join(__dirname, 'dist', 'manifest-production.xml');
let manifestContent = fs.readFileSync(manifestPath, 'utf8');

manifestContent = manifestContent.replace(/YOUR_USERNAME/g, GITHUB_USERNAME);
manifestContent = manifestContent.replace(/REPO_NAME/g, REPO_NAME);

fs.writeFileSync(manifestPath, manifestContent);
console.log('✅ Updated manifest-production.xml');

// Update index.html
const indexPath = path.join(__dirname, 'dist', 'index.html');
let indexContent = fs.readFileSync(indexPath, 'utf8');

indexContent = indexContent.replace(/YOUR_USERNAME/g, GITHUB_USERNAME);
indexContent = indexContent.replace(/REPO_NAME/g, REPO_NAME);

fs.writeFileSync(indexPath, indexContent);
console.log('✅ Updated index.html');

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
console.log('4. Go to GitHub → Settings → Pages → Source: "Deploy from a branch" → Branch: main → Folder: /dist');
console.log(`5. Your add-in will be available at: https://${GITHUB_USERNAME}.github.io/${REPO_NAME}/`);