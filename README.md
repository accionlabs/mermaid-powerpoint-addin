# 📊 Mermaid PowerPoint Add-in

A powerful PowerPoint add-in that lets you create and edit [Mermaid](https://mermaid.js.org/) diagrams directly in your presentations!

## ✨ Features

- **🎨 Visual Editor**: Write Mermaid code with real-time preview
- **📝 Smart Editing**: Click on existing diagrams to edit their source code
- **💾 Persistent Storage**: Diagram source code is saved with your PowerPoint file
- **🎯 Perfect Integration**: High-quality PNG insertion with preserved aspect ratios
- **🔄 Live Updates**: Update diagrams and see changes immediately
- **🎨 Professional UI**: Clean, intuitive interface with icon-based controls
- **🔧 Debug Tools**: Built-in debugging for advanced users

## 🚀 Try it Live

**Live Demo**: [https://accionlabs.github.io/mermaid-powerpoint-addin/](https://accionlabs.github.io/mermaid-powerpoint-addin/)

## 📥 Installation

### Option 1: Direct Installation (Recommended)
1. Visit the [live demo page](https://accionlabs.github.io/mermaid-powerpoint-addin/)
2. Download the `manifest-production.xml` file
3. Open PowerPoint (Desktop version)
4. Go to **Insert** → **Office Add-ins** → **Upload My Add-in**
5. Select the downloaded manifest file and click **Upload**
6. The add-in will appear in your PowerPoint ribbon!

### Option 2: Development Installation
1. Clone this repository
2. Install dependencies:
   ```bash
   npm install
   ```
3. Start the development server:
   ```bash
   npm start
   ```
4. Sideload the add-in in PowerPoint:
   ```bash
   npm run sideload
   ```

## 🎯 How to Use

1. **Open the Add-in**: Click the "Insert Mermaid" button in the PowerPoint ribbon
2. **Create New Diagram**: 
   - Click "✨ New" to start fresh
   - Write your Mermaid code in the editor
   - Click "🔄 Update Preview" to see the diagram
   - Click "📊 Insert Diagram" to add it to your slide
3. **Edit Existing Diagram**:
   - Select a diagram you created with this add-in
   - Click "📝 Edit Selected" 
   - Modify the code and click "📝 Update Diagram"
4. **Debug Mode**: Check the "Show debug tools" checkbox for advanced debugging

## 📊 Supported Diagram Types

This add-in supports all Mermaid diagram types:

- **Flowcharts**: Decision trees, process flows
- **Sequence Diagrams**: System interactions
- **Class Diagrams**: Object-oriented designs
- **State Diagrams**: System states and transitions
- **Entity Relationship Diagrams**: Database schemas
- **User Journey Maps**: User experience flows
- **Gantt Charts**: Project timelines
- **Pie Charts**: Data visualization
- **Git Graphs**: Version control workflows

## 🛠️ Development

### Prerequisites
- Node.js (version 18 or higher)
- PowerPoint Desktop (Windows or Mac)

### Development Commands
- `npm start` - Start development server (https://localhost:3000)
- `npm run build` - Build for production  
- `npm run sideload` - Sideload add-in in PowerPoint
- `npm run validate` - Validate manifest

### Project Structure
```
src/
├── taskpane/
│   ├── components/
│   │   └── MermaidEditor.tsx    # Main UI component
│   ├── utils/
│   │   └── powerPointUtils.ts   # PowerPoint integration
│   ├── taskpane.tsx            # Entry point
│   └── taskpane.html           # HTML template
└── commands/
    ├── commands.ts             # Ribbon commands
    └── commands.html           # Commands HTML
```

## 🚀 Deployment to GitHub Pages

This project can be easily deployed to GitHub Pages using the `dist` folder.

### Quick Deployment Steps

1. **Build the project:**
   ```bash
   npm run build
   ```

2. **Update URLs for your repository:**
   ```bash
   # Set your GitHub username and repo name
   export GITHUB_USERNAME=yourusername
   export mermaid-powerpoint-addin=your-repo-name
   npm run deploy
   ```

3. **Push to GitHub:**
   ```bash
   git add .
   git commit -m "Deploy to GitHub Pages"
   git push origin main
   ```

4. **Enable GitHub Pages:**
   - Go to your repository on GitHub
   - Settings → Pages
   - Source: "Deploy from a branch"
   - Branch: `main`
   - Folder: `/dist`
   - Click "Save"

5. **Your add-in will be live at:**
   `https://yourusername.github.io/your-repo-name/`

### Manual URL Updates (Alternative)
If you prefer to update URLs manually:
1. Edit `dist/manifest-production.xml` - Replace `accionlabs` and `mermaid-powerpoint-addin`
2. Edit `dist/index.html` - Replace `accionlabs` and `mermaid-powerpoint-addin`
3. Edit `README.md` - Replace `accionlabs` and `mermaid-powerpoint-addin`

## 🏗️ Architecture

- **Frontend**: React + TypeScript
- **Diagram Rendering**: Mermaid.js
- **Office Integration**: Office.js API
- **Storage**: Custom XML Parts in PowerPoint files
- **Shape Tagging**: Advanced shape identification system
- **Build System**: Webpack

## 🔧 Technical Details

### Shape Identification
- Uses shape tags for reliable diagram identification
- Supports moving and resizing diagrams while maintaining editability
- Fallback to position-based matching for compatibility

### Storage System
- Diagram source code stored as Custom XML Parts
- Metadata includes creation/update timestamps
- Automatic cleanup of orphaned data

### Cross-Platform Support
- Works on PowerPoint for Windows and Mac
- Uses Office.js for maximum compatibility
- Graceful fallbacks for unsupported features

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [Mermaid.js](https://mermaid.js.org/) for the amazing diagramming library
- [Office.js](https://docs.microsoft.com/en-us/office/dev/add-ins/) for PowerPoint integration
- The open-source community for inspiration and tools

## 📞 Support

- 🐛 [Report Issues](https://github.com/accionlabs/mermaid-powerpoint-addin/issues)
- 💡 [Request Features](https://github.com/accionlabs/mermaid-powerpoint-addin/issues)
- 📖 [Mermaid Documentation](https://mermaid.js.org/)

---

Made with ❤️ for the PowerPoint community