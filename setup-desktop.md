# Desktop PowerPoint Setup Guide

## Method 1: Automatic Sideload (Recommended)
```bash
# Make sure your server is running
npm start

# In another terminal, run:
npx office-addin-debugging start manifest.xml desktop
```

## Method 2: Manual Registry (Mac)
1. Create the folder: `~/Library/Containers/com.microsoft.PowerPoint/Data/Documents/wef`
2. Copy `manifest.xml` to that folder
3. Restart PowerPoint
4. Look for "Mermaid" in the ribbon

## Method 3: PowerPoint Add-ins Store
1. Open Desktop PowerPoint
2. Go to **Insert** → **Get Add-ins** 
3. Click **"Upload My Add-in"**
4. Browse to your `manifest.xml` file

## Method 4: Developer Tools
1. Open PowerPoint
2. Go to **File** → **Options** → **Advanced**
3. Scroll down to **"Developer"** section
4. Check **"Show Developer tab in the Ribbon"**
5. In the **Developer** tab, click **"COM Add-ins"**
6. Add your manifest file

## Troubleshooting
- Ensure the server is running at https://localhost:3000
- The manifest should be accessible at https://localhost:3000/manifest.xml
- Try restarting PowerPoint after adding the add-in
- Check if certificate is trusted in System Preferences → Security & Privacy

## Expected Result
You should see:
- **"Mermaid" group** in the Home ribbon
- **"Insert Mermaid" button** 
- **Task pane opens** when clicked
- **Better API support** than PowerPoint Online (direct insertion)