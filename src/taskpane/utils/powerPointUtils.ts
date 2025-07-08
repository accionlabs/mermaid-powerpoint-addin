/* global Office, PowerPoint, Word */

// Check if we're running in Office context
const isOfficeContext = typeof Office !== 'undefined';

// Platform detection
export enum OfficePlatform {
  PowerPoint = 'PowerPoint',
  Word = 'Word',
  Unknown = 'Unknown'
}

export function detectOfficePlatform(): OfficePlatform {
  if (!isOfficeContext) return OfficePlatform.Unknown;
  
  const host = Office.context.host;
  switch (host) {
    case Office.HostType.PowerPoint:
      return OfficePlatform.PowerPoint;
    case Office.HostType.Word:
      return OfficePlatform.Word;
    default:
      return OfficePlatform.Unknown;
  }
}

export function checkWordApiSupport(): boolean {
  // Check for Word API 1.3 which has better picture insertion support
  if (Office.context.requirements.isSetSupported('WordApi', '1.3')) {
    console.log('Word API 1.3 supported');
    return true;
  }
  
  // Fallback to 1.1 for basic support
  if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
    console.log('Word API 1.1 supported (basic)');
    return true;
  }
  
  console.log('Word API not supported');
  return false;
}

export function checkPowerPointApiSupport(): boolean {
  return Office.context.requirements.isSetSupported('PowerPointApi', '1.1');
}

// Generate unique ID for diagrams
export function generateId(): string {
  return 'mermaid-' + Math.random().toString(36).substr(2, 9);
}

export interface DiagramData {
  id: string;
  code: string;
}

// Abstract diagram insertion interface
export interface DiagramInserter {
  insertDiagram(mermaidCode: string, svgContent: string): Promise<void>;
  updateDiagram(diagramId: string, mermaidCode: string, svgContent: string): Promise<void>;
  getSelectedDiagram(): Promise<DiagramData | null>;
  listStoredDiagrams(): Promise<string>;
  getSelectedShapeInfo(): Promise<string>;
  captureCursorPosition?(): Promise<void>; // Optional method for Word (legacy)
  exitInsertionMode?(): Promise<void>; // Optional method for Word
  insertAtCurrentPosition?(): Promise<void>; // Optional method for Word
}

// PowerPoint implementation
class PowerPointInserter implements DiagramInserter {
  async insertDiagram(mermaidCode: string, svgContent: string): Promise<void> {
    return insertDiagram(mermaidCode, svgContent);
  }
  
  async updateDiagram(diagramId: string, mermaidCode: string, svgContent: string): Promise<void> {
    return updateDiagram(diagramId, mermaidCode, svgContent);
  }
  
  async getSelectedDiagram(): Promise<DiagramData | null> {
    return getSelectedDiagram();
  }
  
  async listStoredDiagrams(): Promise<string> {
    return listAllStoredDiagrams();
  }
  
  async getSelectedShapeInfo(): Promise<string> {
    return getSelectedShapeInfo();
  }
}

// Word implementation
class WordInserter implements DiagramInserter {
  private insertionMode: boolean = false;
  private pendingInsertion: {mermaidCode: string, svgContent: string} | null = null;
  private selectionChangedHandler: any = null;
  private onInsertionComplete: (() => void) | null = null;
  
  // Set callback for when insertion completes
  setInsertionCompleteCallback(callback: () => void): void {
    this.onInsertionComplete = callback;
  }
  
  // Enter insertion mode - prepare for user to click where they want the diagram
  async enterInsertionMode(mermaidCode: string, svgContent: string): Promise<void> {
    console.log('WordInserter: Entering insertion mode');
    
    this.insertionMode = true;
    this.pendingInsertion = { mermaidCode, svgContent };
    
    // Set up Office document selection change event handler
    try {
      console.log('WordInserter: Setting up DocumentSelectionChanged event handler');
      
      // Use Office.context.document.addHandlerAsync to listen for selection changes
      const addHandlerResult = await new Promise<boolean>((resolve, reject) => {
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          this.handleDocumentSelectionChanged.bind(this),
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              console.log('WordInserter: DocumentSelectionChanged handler added successfully');
              resolve(true);
            } else {
              console.error('WordInserter: Failed to add DocumentSelectionChanged handler:', result.error);
              reject(new Error(result.error.message));
            }
          }
        );
      });
      
      console.log('WordInserter: Insertion mode active - waiting for user click');
      
    } catch (error) {
      console.error('WordInserter: Could not set up selection change handler:', error);
      throw new Error('Could not enter insertion mode');
    }
  }
  
  // Handle Office document selection changed event
  private async handleDocumentSelectionChanged(eventArgs: any): Promise<void> {
    if (!this.insertionMode || !this.pendingInsertion) {
      return;
    }
    
    console.log('WordInserter: Document selection changed during insertion mode - cursor positioned');
    console.log('WordInserter: Event args:', eventArgs);
    
    // Don't insert automatically - just log that the user has positioned their cursor
    // The actual insertion will happen when they click "Insert Here"
    console.log('WordInserter: User positioned cursor, waiting for "Insert Here" button click');
  }
  
  // Method to insert at current cursor position when user is ready
  async insertAtCurrentPosition(): Promise<void> {
    if (!this.insertionMode || !this.pendingInsertion) {
      throw new Error('Not in insertion mode');
    }
    
    console.log('WordInserter: Inserting diagram at current cursor position');
    
    try {
      // Perform insertion at current location
      await this.performActualInsertion(this.pendingInsertion.mermaidCode, this.pendingInsertion.svgContent);
      
      // Exit insertion mode
      await this.exitInsertionMode();
      
      // Notify UI that insertion is complete
      if (this.onInsertionComplete) {
        this.onInsertionComplete();
      }
      
      console.log('WordInserter: Diagram inserted successfully');
      
    } catch (error) {
      if (error instanceof Error && error.message.startsWith('INSERTION_COMPLETE:')) {
        console.log('WordInserter: Insertion completed successfully');
        // Exit insertion mode on successful completion
        await this.exitInsertionMode();
        
        // Notify UI that insertion is complete
        if (this.onInsertionComplete) {
          this.onInsertionComplete();
        }
        return;
      } else {
        console.error('WordInserter: Failed to insert:', error);
        await this.exitInsertionMode();
        throw error;
      }
    }
  }
  
  // Exit insertion mode and clean up
  async exitInsertionMode(): Promise<void> {
    console.log('WordInserter: Exiting insertion mode');
    
    this.insertionMode = false;
    this.pendingInsertion = null;
    
    // Clean up Office event handler
    try {
      console.log('WordInserter: Removing DocumentSelectionChanged event handler');
      
      await new Promise<void>((resolve, reject) => {
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              console.log('WordInserter: DocumentSelectionChanged handler removed successfully');
              resolve();
            } else {
              console.log('WordInserter: Could not remove DocumentSelectionChanged handler (non-critical):', result.error);
              resolve(); // Don't reject - this is non-critical
            }
          }
        );
      });
      
    } catch (error) {
      console.log('WordInserter: Could not remove selection handler (non-critical):', error);
    }
  }
  
  // Perform the actual insertion at current cursor position
  private async performActualInsertion(mermaidCode: string, svgContent: string): Promise<void> {
    await Word.run(async (context) => {
      // Get current selection (where user just clicked)
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();
      
      console.log('WordInserter: Inserting at user-selected position');
      
      // Generate optimal PNG as before
      const pageDimensions = await this.getDocumentPageDimensions();
      const svgDimensions = this.parseSvgViewBox(svgContent);
      let svgWidth = svgDimensions.width > 0 ? svgDimensions.width : 800;
      let svgHeight = svgDimensions.height > 0 ? svgDimensions.height : 600;
      
      const optimalSize = this.calculateOptimalDiagramSize(
        svgWidth, svgHeight, 
        pageDimensions.pageWidth, pageDimensions.pageHeight,
        pageDimensions.marginLeft, pageDimensions.marginRight,
        pageDimensions.marginTop, pageDimensions.marginBottom
      );
      
      const pngData = await this.convertSvgToPngForWord(svgContent, optimalSize.width, optimalSize.height);
      
      // Insert at current selection
      let picture;
      if (selection.text && selection.text.length > 0) {
        console.log('WordInserter: Text selected, inserting after selection');
        picture = selection.insertInlinePictureFromBase64(pngData.base64, Word.InsertLocation.after);
      } else {
        console.log('WordInserter: No text selected, inserting at cursor');
        picture = selection.insertInlinePictureFromBase64(pngData.base64, Word.InsertLocation.replace);
      }
      
      // Set optimal size
      picture.width = optimalSize.width;
      picture.height = optimalSize.height;
      
      // Add spacing and metadata as before
      picture.insertParagraph('', Word.InsertLocation.after);
      
      const diagramId = generateId();
      picture.altTextDescription = `MERMAID_ID:${diagramId}`;
      
      await this.storeDiagramMetadata(diagramId, mermaidCode);
      await context.sync();
      
      console.log('WordInserter: Diagram inserted successfully at user-chosen location');
      
      // Signal successful insertion to UI (will be caught as success message)
      throw new Error('INSERTION_COMPLETE: Diagram inserted successfully at your chosen location!');
    });
  }
  
  // Legacy capture method (now unused but keeping for compatibility)
  async captureCursorPosition(): Promise<void> {
    console.log('WordInserter: captureCursorPosition called (legacy method)');
  }

  async insertDiagram(mermaidCode: string, svgContent: string): Promise<void> {
    console.log('WordInserter: Starting two-step diagram insertion');
    console.log('WordInserter: SVG content length:', svgContent.length);
    
    if (!checkWordApiSupport()) {
      throw new Error('Word API 1.1 not supported. Please use Microsoft Word 2016 or newer.');
    }
    
    console.log('WordInserter: Word API support confirmed');
    
    // Validate SVG content
    if (!svgContent || svgContent.trim().length === 0) {
      throw new Error('SVG content is empty or invalid');
    }
    
    if (!svgContent.includes('<svg')) {
      throw new Error('Content does not appear to be valid SVG');
    }
    
    console.log('WordInserter: SVG validation passed');
    
    // Enter insertion mode instead of immediate insertion
    try {
      await this.enterInsertionMode(mermaidCode, svgContent);
      
      // Throw special success message to indicate mode change
      throw new Error('SUCCESS: Click in your document where you want the diagram to be inserted.');
      
    } catch (error) {
      if (error instanceof Error && error.message.startsWith('SUCCESS:')) {
        throw error; // Re-throw success messages
      }
      console.error('WordInserter: Failed to enter insertion mode:', error);
      throw new Error(`Failed to prepare insertion: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  
  async updateDiagram(diagramId: string, mermaidCode: string, svgContent: string): Promise<void> {
    console.log('WordInserter: Updating diagram with ID:', diagramId);
    
    await Word.run(async (context) => {
      try {
        // First, find and delete the old diagram image
        const inlinePictures = context.document.body.inlinePictures;
        inlinePictures.load(['items']);
        await context.sync();
        
        // Find the picture associated with this diagram ID by checking custom XML parts
        const customXmlParts = context.document.customXmlParts;
        customXmlParts.load('items');
        await context.sync();
        
        let foundDiagram = false;
        for (let i = 0; i < customXmlParts.items.length; i++) {
          const part = customXmlParts.items[i];
          try {
            const xml = part.getXml();
            await context.sync();
            
            if (xml.value.includes(`<Id>${diagramId}</Id>`)) {
              console.log('WordInserter: Found existing diagram metadata for update');
              foundDiagram = true;
              
              // Delete the old XML part
              part.delete();
              break;
            }
          } catch (error) {
            // Skip this part if we can't read it
            continue;
          }
        }
        
        if (!foundDiagram) {
          throw new Error(`Diagram with ID ${diagramId} not found. It may have been deleted or moved.`);
        }
        
        // Get document page dimensions for optimal sizing
        console.log('WordInserter: Getting document page dimensions for update');
        const pageDimensions = await this.getDocumentPageDimensions();
        
        // Parse SVG dimensions
        console.log('WordInserter: Parsing SVG dimensions for update');
        const svgDimensions = this.parseSvgViewBox(svgContent);
        let svgWidth = svgDimensions.width;
        let svgHeight = svgDimensions.height;
        
        // Fallback to default if parsing failed
        if (svgWidth <= 0 || svgHeight <= 0) {
          console.log('WordInserter: Could not parse SVG dimensions for update, using defaults');
          svgWidth = 800;
          svgHeight = 600;
        }
        
        // Calculate optimal diagram size for the document
        console.log('WordInserter: Calculating optimal diagram size for update');
        const optimalSize = this.calculateOptimalDiagramSize(
          svgWidth, 
          svgHeight, 
          pageDimensions.pageWidth, 
          pageDimensions.pageHeight,
          pageDimensions.marginLeft,
          pageDimensions.marginRight,
          pageDimensions.marginTop,
          pageDimensions.marginBottom
        );
        
        // Convert SVG to PNG at 300 DPI for the calculated optimal size
        console.log('WordInserter: Converting SVG to PNG at 300 DPI for update');
        const pngWithWhiteBackground = await this.convertSvgToPngForWord(svgContent, optimalSize.width, optimalSize.height);
        
        // Insert the new diagram at the current selection or end of document
        let picture;
        try {
          const selection = context.document.getSelection();
          picture = selection.insertInlinePictureFromBase64(
            pngWithWhiteBackground.base64, 
            Word.InsertLocation.replace
          );
        } catch (selectionError) {
          // Fallback to end of document if selection fails
          picture = context.document.body.insertInlinePictureFromBase64(
            pngWithWhiteBackground.base64, 
            Word.InsertLocation.end
          );
        }
        
        // Set picture dimensions to the calculated optimal size
        // The PNG was created at 300 DPI specifically for this display size
        console.log('WordInserter: Update - resizing picture to optimal calculated size:', optimalSize.width, 'x', optimalSize.height, 'points');
        console.log('WordInserter: Update - PNG was created at 300 DPI specifically for this size (', pngWithWhiteBackground.width, 'x', pngWithWhiteBackground.height, 'pixels)');
        picture.width = optimalSize.width;
        picture.height = optimalSize.height;
        
        // Store updated diagram metadata
        await this.storeDiagramMetadata(diagramId, mermaidCode);
        
        await context.sync();
        console.log('WordInserter: Diagram updated successfully');
        
      } catch (error) {
        console.error('WordInserter: Update failed:', error);
        throw new Error(`Failed to update diagram: ${error instanceof Error ? error.message : String(error)}`);
      }
    });
  }
  
  async getSelectedDiagram(): Promise<DiagramData | null> {
    console.log('WordInserter: Getting selected diagram for editing');
    
    return await Word.run(async (context) => {
      try {
        // Get the current selection to find selected images
        const selection = context.document.getSelection();
        const inlinePictures = selection.inlinePictures;
        inlinePictures.load(['items', 'altTextDescription']);
        await context.sync();
        
        console.log('WordInserter: Found', inlinePictures.items.length, 'selected pictures');
        
        if (inlinePictures.items.length === 0) {
          console.log('WordInserter: No picture selected, cannot edit');
          return null;
        }
        
        // Get the first selected picture and check its alt text for diagram ID
        const selectedPicture = inlinePictures.items[0];
        const altText = selectedPicture.altTextDescription;
        console.log('WordInserter: Selected picture alt text:', altText);
        
        if (!altText || !altText.startsWith('MERMAID_ID:')) {
          console.log('WordInserter: Selected image is not a Mermaid diagram (no ID in alt text)');
          return null;
        }
        
        // Extract the diagram ID from alt text
        const diagramId = altText.substring('MERMAID_ID:'.length);
        console.log('WordInserter: Found diagram ID in selected image:', diagramId);
        
        // Now find the metadata for this specific diagram ID
        const customXmlParts = context.document.customXmlParts;
        customXmlParts.load('items');
        await context.sync();
        
        console.log('WordInserter: Searching for metadata with ID:', diagramId);
        
        for (let i = 0; i < customXmlParts.items.length; i++) {
          const part = customXmlParts.items[i];
          try {
            const xml = part.getXml();
            await context.sync();
            
            if (xml.value.includes('<MermaidDiagram>') && xml.value.includes(`<Id>${diagramId}</Id>`)) {
              console.log('WordInserter: Found matching metadata for diagram:', diagramId);
              console.log('WordInserter: Full XML content:', xml.value);
              
              // Check what's actually in the CDATA section
              const cdataStartIndex = xml.value.indexOf('<![CDATA[');
              const cdataEndIndex = xml.value.indexOf(']]>');
              if (cdataStartIndex >= 0 && cdataEndIndex >= 0) {
                const cdataContent = xml.value.substring(cdataStartIndex + 9, cdataEndIndex);
                console.log('WordInserter: Raw CDATA content:', JSON.stringify(cdataContent));
              }
              
              // Updated regex to handle newlines and special characters in CDATA
              const codeMatch = xml.value.match(/<Code><!\[CDATA\[([\s\S]*?)\]\]><\/Code>/);
              console.log('WordInserter: Code regex result:', codeMatch);
              
              if (codeMatch) {
                console.log('WordInserter: Successfully extracted diagram code:', codeMatch[1].substring(0, 50) + '...');
                return {
                  id: diagramId,
                  code: codeMatch[1]
                };
              } else {
                console.log('WordInserter: Failed to parse diagram code from metadata');
                // Try alternative regex patterns
                const altCodeMatch1 = xml.value.match(/<Code>([^<]+)<\/Code>/);
                const altCodeMatch2 = xml.value.match(/<!\[CDATA\[([^\]]+)\]\]>/);
                console.log('WordInserter: Alternative regex 1:', altCodeMatch1);
                console.log('WordInserter: Alternative regex 2:', altCodeMatch2);
              }
            }
          } catch (error) {
            console.log(`WordInserter: Error reading XML part ${i + 1}:`, error);
            continue;
          }
        }
        
        console.log('WordInserter: No metadata found for diagram ID:', diagramId);
        return null;
        
      } catch (error) {
        console.error('WordInserter: Failed to get selected diagram:', error);
        return null;
      }
    });
  }
  
  async listStoredDiagrams(): Promise<string> {
    console.log('WordInserter: Listing stored diagrams');
    
    return await Word.run(async (context) => {
      try {
        const customXmlParts = context.document.customXmlParts;
        customXmlParts.load('items');
        await context.sync();
        
        const diagrams: any[] = [];
        
        for (let i = 0; i < customXmlParts.items.length; i++) {
          const part = customXmlParts.items[i];
          try {
            const xml = part.getXml();
            await context.sync();
            
            if (xml.value.includes('<MermaidDiagram>')) {
              const idMatch = xml.value.match(/<Id>([^<]+)<\/Id>/);
              const createdMatch = xml.value.match(/<CreatedAt>([^<]+)<\/CreatedAt>/);
              const codeMatch = xml.value.match(/<Code><!\[CDATA\[([^\]]+)\]\]><\/Code>/);
              
              if (idMatch && createdMatch && codeMatch) {
                diagrams.push({
                  id: idMatch[1],
                  created: createdMatch[1],
                  codePreview: codeMatch[1].substring(0, 50) + '...'
                });
              }
            }
          } catch (error) {
            // Skip this part if we can't read it
            continue;
          }
        }
        
        if (diagrams.length === 0) {
          return 'No Mermaid diagrams found in this document.';
        }
        
        return `Found ${diagrams.length} Mermaid diagram(s):\n\n` +
          diagrams.map((d, index) => 
            `${index + 1}. ID: ${d.id}\n   Created: ${d.created}\n   Code: ${d.codePreview}`
          ).join('\n\n');
          
      } catch (error) {
        return `Error listing diagrams: ${error instanceof Error ? error.message : String(error)}`;
      }
    });
  }
  
  async getSelectedShapeInfo(): Promise<string> {
    return await Word.run(async (context) => {
      try {
        const selection = context.document.getSelection();
        const inlinePictures = selection.inlinePictures;
        inlinePictures.load(['items', 'width', 'height']);
        await context.sync();
        
        if (inlinePictures.items.length === 0) {
          return 'No picture selected in Word.';
        }
        
        const picture = inlinePictures.items[0];
        return `Selected Word Picture Info:
- Width: ${picture.width}
- Height: ${picture.height}
- Type: Inline Picture
- Context: Word Document

Note: Word uses inline pictures instead of shapes like PowerPoint.`;
        
      } catch (error) {
        return `Error getting picture info: ${error instanceof Error ? error.message : String(error)}`;
      }
    });
  }
  
  private async storeDiagramMetadata(diagramId: string, mermaidCode: string): Promise<void> {
    console.log('WordInserter: Starting metadata storage for diagram:', diagramId);
    
    await Word.run(async (context) => {
      try {
        const customXmlParts = context.document.customXmlParts;
        customXmlParts.load('items');
        await context.sync();
        
        console.log('WordInserter: Loaded custom XML parts, current count:', customXmlParts.items.length);
        
        const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<MermaidDiagram>
  <Id>${diagramId}</Id>
  <Code><![CDATA[${mermaidCode}]]></Code>
  <CreatedAt>${new Date().toISOString()}</CreatedAt>
</MermaidDiagram>`;
        
        console.log('WordInserter: Adding XML content (full):', xmlContent);
        
        const newPart = customXmlParts.add(xmlContent);
        await context.sync();
        
        // Verify the part was added
        customXmlParts.load('items');
        await context.sync();
        console.log('WordInserter: Metadata stored successfully. New XML parts count:', customXmlParts.items.length);
        
      } catch (error) {
        console.error('WordInserter: Failed to store metadata:', error);
        throw error;
      }
    });
  }
  
  private async convertSvgToPngForWord(svgString: string, targetDisplayWidth: number, targetDisplayHeight: number): Promise<{base64: string, width: number, height: number}> {
    return new Promise((resolve, reject) => {
      try {
        console.log('WordInserter: Starting SVG to PNG conversion for Word');
        console.log('WordInserter: Target display dimensions:', targetDisplayWidth, 'x', targetDisplayHeight, 'points');
        
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        const img = new Image();
        
        img.onload = () => {
          try {
            // Parse viewBox from SVG string for accurate dimensions
            const viewBoxDimensions = this.parseSvgViewBox(svgString);
            console.log('WordInserter: SVG viewBox dimensions:', viewBoxDimensions);
            console.log('WordInserter: Natural image dimensions:', img.naturalWidth, 'x', img.naturalHeight);
            
            // ALWAYS prefer viewBox dimensions over natural dimensions for accuracy
            let svgWidth, svgHeight;
            if (viewBoxDimensions.width > 0 && viewBoxDimensions.height > 0) {
              svgWidth = viewBoxDimensions.width;
              svgHeight = viewBoxDimensions.height;
              console.log('WordInserter: Using viewBox dimensions:', svgWidth, 'x', svgHeight);
            } else {
              svgWidth = img.naturalWidth || 800;
              svgHeight = img.naturalHeight || 600;
              console.log('WordInserter: Fallback to natural/default dimensions:', svgWidth, 'x', svgHeight);
            }
            
            // Calculate 300 DPI pixel dimensions for target display size
            // 300 DPI = 300 pixels per inch = 4.167 pixels per point (72 points per inch)
            const pixelsPerPoint = 300 / 72; // 4.167 pixels per point at 300 DPI
            const canvasWidth = Math.round(targetDisplayWidth * pixelsPerPoint);
            const canvasHeight = Math.round(targetDisplayHeight * pixelsPerPoint);
            
            console.log('WordInserter: Creating 300 DPI canvas:', canvasWidth, 'x', canvasHeight, 'pixels');
            
            // Set canvas to 300 DPI resolution for target display size
            canvas.width = canvasWidth;
            canvas.height = canvasHeight;
            
            // Enable high-quality rendering
            ctx!.imageSmoothingEnabled = true;
            ctx!.imageSmoothingQuality = 'high';
            
            // Fill with white background (Word expects non-transparent images)
            ctx!.fillStyle = 'white';
            ctx!.fillRect(0, 0, canvasWidth, canvasHeight);
            
            // Draw the SVG scaled to fill the canvas exactly
            ctx!.drawImage(img, 0, 0, canvasWidth, canvasHeight);
            
            // Convert to PNG with maximum quality
            const base64 = canvas.toDataURL('image/png', 1.0).split(',')[1];
            
            console.log('WordInserter: 300 DPI PNG created:', canvasWidth, 'x', canvasHeight, 'pixels for', targetDisplayWidth, 'x', targetDisplayHeight, 'points display');
            
            resolve({
              base64,
              width: canvasWidth,
              height: canvasHeight
            });
          } catch (canvasError) {
            console.error('WordInserter: Canvas processing failed:', canvasError);
            reject(new Error(`Canvas processing failed: ${canvasError}`));
          }
        };
        
        img.onerror = () => {
          console.error('WordInserter: Failed to load SVG');
          reject(new Error('Failed to load SVG image'));
        };
        
        // Create data URL for the SVG
        const svgDataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgString)}`;
        img.src = svgDataUrl;
        
      } catch (error) {
        console.error('WordInserter: SVG conversion setup failed:', error);
        reject(new Error(`SVG conversion setup failed: ${error}`));
      }
    });
  }
  
  private async getDocumentPageDimensions(): Promise<{pageWidth: number, pageHeight: number, marginLeft: number, marginRight: number, marginTop: number, marginBottom: number}> {
    return await Word.run(async (context) => {
      try {
        // Try to get page setup information from the document
        // Use default values since Word API doesn't expose section page dimensions in the same way
        console.log('WordInserter: Using default Letter size with 1" margins');
        console.log('WordInserter: Page dimensions (points): 612 x 792');
        console.log('WordInserter: Margins (points) - L: 72, R: 72, T: 72, B: 72');
        
        // Default to US Letter size (8.5" x 11" = 612 x 792 points) with 1" margins
        // This covers the majority of documents and provides good sizing
        return {
          pageWidth: 612,
          pageHeight: 792,
          marginLeft: 72,
          marginRight: 72,
          marginTop: 72,
          marginBottom: 72
        };
      } catch (error) {
        console.log('WordInserter: Could not get page dimensions, using defaults:', error);
        // Fallback to US Letter size with 1" margins
        return {
          pageWidth: 612,
          pageHeight: 792,
          marginLeft: 72,
          marginRight: 72,
          marginTop: 72,
          marginBottom: 72
        };
      }
    });
  }
  
  private calculateOptimalDiagramSize(svgWidth: number, svgHeight: number, pageWidth: number, pageHeight: number, marginLeft: number, marginRight: number, marginTop: number, marginBottom: number): {width: number, height: number} {
    // Calculate available space (subtract margins)
    const availableWidth = pageWidth - marginLeft - marginRight;
    const availableHeight = pageHeight - marginTop - marginBottom;
    
    console.log('WordInserter: Available space:', availableWidth, 'x', availableHeight, 'points');
    console.log('WordInserter: SVG aspect ratio:', svgWidth, 'x', svgHeight);
    
    // Calculate aspect ratio
    const svgAspectRatio = svgWidth / svgHeight;
    const availableAspectRatio = availableWidth / availableHeight;
    
    let targetWidth, targetHeight;
    
    if (svgAspectRatio > availableAspectRatio) {
      // SVG is wider relative to available space - fit to width
      targetWidth = availableWidth;
      targetHeight = targetWidth / svgAspectRatio;
      console.log('WordInserter: Fitting to width');
    } else {
      // SVG is taller relative to available space - fit to height
      targetHeight = availableHeight;
      targetWidth = targetHeight * svgAspectRatio;
      console.log('WordInserter: Fitting to height');
    }
    
    // Ensure minimum size (at least 2 inches wide or tall)
    const minSize = 144; // 2 inches = 144 points
    if (targetWidth < minSize && targetHeight < minSize) {
      if (svgAspectRatio > 1) {
        targetWidth = minSize;
        targetHeight = minSize / svgAspectRatio;
      } else {
        targetHeight = minSize;
        targetWidth = minSize * svgAspectRatio;
      }
      console.log('WordInserter: Applied minimum size constraint');
    }
    
    console.log('WordInserter: Calculated optimal size:', targetWidth, 'x', targetHeight, 'points');
    
    return {
      width: Math.round(targetWidth),
      height: Math.round(targetHeight)
    };
  }
  
  private parseSvgViewBox(svgString: string): {width: number, height: number} {
    try {
      // Try to extract viewBox from SVG tag
      const viewBoxMatch = svgString.match(/viewBox\s*=\s*["']([^"']+)["']/);
      if (viewBoxMatch) {
        const viewBoxValues = viewBoxMatch[1].split(/\s+/);
        if (viewBoxValues.length === 4) {
          const width = parseFloat(viewBoxValues[2]);
          const height = parseFloat(viewBoxValues[3]);
          
          if (!isNaN(width) && !isNaN(height) && width > 0 && height > 0) {
            console.log('WordInserter: Parsed viewBox successfully:', width, 'x', height);
            return { width, height };
          }
        }
      }
      
      // Fallback: try to extract width/height attributes
      const widthMatch = svgString.match(/width\s*=\s*["']?([^"'\s>]+)/);
      const heightMatch = svgString.match(/height\s*=\s*["']?([^"'\s>]+)/);
      
      if (widthMatch && heightMatch) {
        const width = parseFloat(widthMatch[1]);
        const height = parseFloat(heightMatch[1]);
        
        if (!isNaN(width) && !isNaN(height) && width > 0 && height > 0) {
          console.log('WordInserter: Parsed width/height attributes:', width, 'x', height);
          return { width, height };
        }
      }
      
      console.log('WordInserter: Could not parse SVG dimensions, using fallback');
      return { width: 0, height: 0 }; // Signal to use natural dimensions
      
    } catch (error) {
      console.error('WordInserter: Error parsing SVG dimensions:', error);
      return { width: 0, height: 0 };
    }
  }
}

// Factory function to get the appropriate inserter
export function createDiagramInserter(): DiagramInserter {
  const platform = detectOfficePlatform();
  
  switch (platform) {
    case OfficePlatform.PowerPoint:
      return new PowerPointInserter();
    case OfficePlatform.Word:
      return new WordInserter();
    default:
      throw new Error(`Unsupported platform: ${platform}`);
  }
}

// Helper function to capture cursor position (Word only) - Legacy
export async function captureCursorPosition(inserter: DiagramInserter): Promise<void> {
  const platform = detectOfficePlatform();
  if (platform === OfficePlatform.Word && inserter.captureCursorPosition) {
    console.log('Capturing cursor position before task pane interaction (legacy)');
    await inserter.captureCursorPosition();
  } else {
    console.log('Cursor position capture not needed for this platform');
  }
}

// Helper function to exit insertion mode (Word only)
export async function exitInsertionMode(inserter: DiagramInserter): Promise<void> {
  const platform = detectOfficePlatform();
  if (platform === OfficePlatform.Word && inserter.exitInsertionMode) {
    console.log('Exiting insertion mode');
    await inserter.exitInsertionMode();
  } else {
    console.log('Insertion mode exit not needed for this platform');
  }
}

// Helper function to insert at current position (Word only)
export async function insertAtCurrentPosition(inserter: DiagramInserter): Promise<void> {
  const platform = detectOfficePlatform();
  if (platform === OfficePlatform.Word && inserter.insertAtCurrentPosition) {
    console.log('Inserting at current position');
    await inserter.insertAtCurrentPosition();
  } else {
    throw new Error('Insert at current position not available for this platform');
  }
}

export interface MermaidSettings {
  fontFamily: string;
  fontSize: number;
  primaryColor: string;
  primaryTextColor: string;
  primaryBorderColor: string;
  lineColor: string;
  backgroundColor: string;
  secondaryColor: string;
  tertiaryColor: string;
  theme: 'default' | 'dark' | 'forest' | 'base' | 'custom';
}

export const defaultSettings: MermaidSettings = {
  fontFamily: 'Arial, sans-serif',
  fontSize: 16,
  primaryColor: '#0078d4',
  primaryTextColor: '#000000',
  primaryBorderColor: '#0078d4',
  lineColor: '#000000',
  backgroundColor: '#ffffff',
  secondaryColor: '#e6f3ff',
  tertiaryColor: '#b3d9ff',
  theme: 'default'
};

// Convert SVG to base64 PNG with transparent background and correct dimensions
export const svgToPng = (svgString: string): Promise<{base64: string, width: number, height: number}> => {
  return new Promise((resolve, reject) => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const img = new Image();
    
    // Set crossOrigin to anonymous to avoid CORS issues
    img.crossOrigin = 'anonymous';
    
    img.onload = () => {
      // Get the natural dimensions from the SVG
      const { naturalWidth, naturalHeight } = img;
      
      // Use a scale factor for high resolution while preserving aspect ratio
      const scaleFactor = 2; // 2x resolution for good quality without being too large
      
      // Set canvas to the scaled natural dimensions
      canvas.width = naturalWidth * scaleFactor;
      canvas.height = naturalHeight * scaleFactor;
      
      // IMPORTANT: Don't fill background - leave transparent!
      // This creates PNG with transparent background instead of white
      
      // Enable high-quality rendering
      ctx!.imageSmoothingEnabled = true;
      ctx!.imageSmoothingQuality = 'high';
      
      // Draw the SVG at the scaled size (on transparent background)
      ctx!.drawImage(img, 0, 0, canvas.width, canvas.height);
      
      try {
        // Get base64 data without data URL prefix - use highest quality
        const pngDataUrl = canvas.toDataURL('image/png', 1.0);
        const base64Data = pngDataUrl.split(',')[1];
        
        // Return both the base64 data and the original natural dimensions
        resolve({
          base64: base64Data,
          width: naturalWidth,
          height: naturalHeight
        });
      } catch (error) {
        console.error('Canvas toDataURL failed:', error);
        reject(new Error('Failed to convert diagram to image'));
      }
    };
    
    img.onerror = (err) => {
      console.error('Failed to load SVG image:', err);
      reject(new Error('Failed to load diagram image'));
    };
    
    // Use data URL directly instead of blob URL to avoid CORS
    const svgDataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgString)}`;
    img.src = svgDataUrl;
  });
};

// Generate unique ID for diagrams
const generateDiagramId = (): string => {
  return 'mermaid_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
};

// Convert base64 to blob
const base64ToBlob = async (base64: string, mimeType: string): Promise<Blob> => {
  const response = await fetch(`data:${mimeType};base64,${base64}`);
  return response.blob();
};

// Store diagram data in custom XML part with shape association
const storeDiagramData = async (diagramId: string, mermaidCode: string): Promise<void> => {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const customXmlParts = presentation.customXmlParts;
    
    // Try to get the most recently added shape for association
    const slides = context.presentation.getSelectedSlides();
    slides.load('items');
    await context.sync();
    
    let shapeInfo = '';
    if (slides.items.length > 0) {
      const slide = slides.items[0];
      const shapes = slide.shapes;
      shapes.load('items');
      await context.sync();
      
      if (shapes.items.length > 0) {
        // Get the last shape (most recently added)
        const lastShape = shapes.items[shapes.items.length - 1];
        lastShape.load(['left', 'top', 'width', 'height']);
        await context.sync();
        
        shapeInfo = `${lastShape.left},${lastShape.top},${lastShape.width},${lastShape.height}`;
      }
    }
    
    const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>
      <MermaidDiagram>
        <Id>${diagramId}</Id>
        <Code><![CDATA[${mermaidCode}]]></Code>
        <ShapeInfo>${shapeInfo}</ShapeInfo>
        <CreatedAt>${new Date().toISOString()}</CreatedAt>
      </MermaidDiagram>
    `;
    
    customXmlParts.add(xmlContent);
    await context.sync();
    console.log('Diagram data stored with ID:', diagramId, 'and shape info:', shapeInfo);
  });
};

// Store diagram data after insertion and tag the shape
const storeDiagramDataAfterInsertion = async (diagramId: string, mermaidCode: string): Promise<string> => {
  let debugLog = 'STORAGE DEBUG LOG:\n';
  debugLog += `Starting storage after insertion for diagram ID: ${diagramId}\n`;
  
  // Wait a short moment for the insertion to complete
  await new Promise(resolve => setTimeout(resolve, 1000));
  debugLog += 'Waited 1 second for insertion to complete\n';
  
  try {
    return PowerPoint.run(async (context) => {
      debugLog += 'PowerPoint.run started for storage\n';
      const presentation = context.presentation;
      const customXmlParts = presentation.customXmlParts;
      
      // Get the most recently added shape and tag it with our diagram ID
      const slides = context.presentation.getSelectedSlides();
      slides.load('items');
      await context.sync();
      debugLog += `Loaded slides, count: ${slides.items.length}\n`;
      
      let shapeTagged = false;
      if (slides.items.length > 0) {
        const slide = slides.items[0];
        const shapes = slide.shapes;
        shapes.load('items');
        await context.sync();
        debugLog += `Loaded shapes on slide, count: ${shapes.items.length}\n`;
        
        if (shapes.items.length > 0) {
          // Get the last shape (most recently added) and tag it
          const lastShape = shapes.items[shapes.items.length - 1];
          lastShape.load(['left', 'top', 'width', 'height']);
          
          try {
            // Tag the shape with our diagram ID using shape tags
            lastShape.tags.add('mermaid_diagram_id', diagramId);
            await context.sync();
            debugLog += `✅ Shape tagged with diagram ID: ${diagramId}\n`;
            shapeTagged = true;
          } catch (tagError) {
            debugLog += `⚠️ Shape tagging failed: ${tagError}\n`;
            debugLog += `Will fall back to position-based matching\n`;
            
            // Load shape info as fallback
            await context.sync();
            const shapeInfo = `${lastShape.left},${lastShape.top},${lastShape.width},${lastShape.height}`;
            debugLog += `Captured shape info as fallback: ${shapeInfo}\n`;
          }
        } else {
          debugLog += 'No shapes found on slide!\n';
        }
      } else {
        debugLog += 'No slides selected!\n';
      }
      
      // Store the diagram metadata
      const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<MermaidDiagram>
  <Id>${diagramId}</Id>
  <Code><![CDATA[${mermaidCode}]]></Code>
  <ShapeTagged>${shapeTagged}</ShapeTagged>
  <CreatedAt>${new Date().toISOString()}</CreatedAt>
</MermaidDiagram>`;
      
      debugLog += `About to add XML content (${xmlContent.length} chars)\n`;
      customXmlParts.add(xmlContent);
      await context.sync();
      debugLog += `✅ XML part added successfully!\n`;
      debugLog += `Final result: ID=${diagramId}, Tagged=${shapeTagged}\n`;
      
      console.log(debugLog);
      return debugLog;
    });
  } catch (error) {
    debugLog += `❌ Storage failed: ${error}\n`;
    console.error(debugLog);
    throw new Error(debugLog);
  }
};

// Retrieve diagram data from custom XML part
const getDiagramData = async (diagramId: string): Promise<string | null> => {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const customXmlParts = presentation.customXmlParts;
    customXmlParts.load('items');
    
    await context.sync();
    
    for (let i = 0; i < customXmlParts.items.length; i++) {
      const xmlPart = customXmlParts.items[i];
      xmlPart.load(['xml']);
      await context.sync();
      
      const xmlDoc = new DOMParser().parseFromString((xmlPart as any).xml, 'text/xml');
      const idElement = xmlDoc.querySelector('Id');
      
      if (idElement && idElement.textContent === diagramId) {
        const codeElement = xmlDoc.querySelector('Code');
        return codeElement ? codeElement.textContent : null;
      }
    }
    
    return null;
  });
};

// Update diagram data in custom XML part
const updateDiagramData = async (diagramId: string, mermaidCode: string): Promise<void> => {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const customXmlParts = presentation.customXmlParts;
    customXmlParts.load('items');
    
    await context.sync();
    
    for (let i = 0; i < customXmlParts.items.length; i++) {
      const xmlPart = customXmlParts.items[i];
      
      try {
        const xmlContent = xmlPart.getXml();
        await context.sync();
        
        if (xmlContent && xmlContent.value) {
          const xmlDoc = new DOMParser().parseFromString(xmlContent.value, 'text/xml');
          const idElement = xmlDoc.querySelector('Id');
          
          if (idElement && idElement.textContent === diagramId) {
            // Update the XML part with new code and timestamp
            const updatedXmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<MermaidDiagram>
  <Id>${diagramId}</Id>
  <Code><![CDATA[${mermaidCode}]]></Code>
  <ShapeTagged>true</ShapeTagged>
  <UpdatedAt>${new Date().toISOString()}</UpdatedAt>
</MermaidDiagram>`;
            
            xmlPart.delete();
            customXmlParts.add(updatedXmlContent);
            await context.sync();
            console.log('Diagram data updated for ID:', diagramId);
            return;
          }
        }
      } catch (error) {
        continue;
      }
    }
    
    throw new Error('Diagram data not found for update');
  });
};


// Download SVG file for manual insertion into PowerPoint
export const downloadSvg = (svgContent: string, diagramId?: string): void => {
  try {
    const filename = `mermaid-diagram-${diagramId || Date.now()}.svg`;
    
    // Method 1: Try using data URL approach (more compatible)
    const dataStr = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgContent);
    const link = document.createElement('a');
    link.setAttribute('href', dataStr);
    link.setAttribute('download', filename);
    link.style.display = 'none';
    
    // Append to body, click, and remove
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    console.log('SVG file download triggered');
  } catch (error) {
    console.error('Failed to download SVG:', error);
    // Fallback: open in new window for manual save
    const filename = `mermaid-diagram-${diagramId || Date.now()}.svg`;
    try {
      const newWindow = window.open('', '_blank');
      if (newWindow) {
        newWindow.document.write(`
          <html>
            <head><title>Save SVG - ${filename}</title></head>
            <body style="font-family: Arial; padding: 20px;">
              <h3>SVG Content (Right-click and Save As...)</h3>
              <p>Right-click the content below and select "Save As..." to save the SVG file:</p>
              <textarea style="width: 100%; height: 400px; font-family: monospace;">${svgContent}</textarea>
              <hr>
              <div style="border: 1px solid #ccc; padding: 10px; margin: 10px 0;">
                ${svgContent}
              </div>
            </body>
          </html>
        `);
        throw new Error('SUCCESS: SVG opened in new window. Right-click the content and "Save As..." to download.');
      } else {
        throw new Error('Could not download SVG file. Please allow popups or check browser download permissions.');
      }
    } catch (fallbackError) {
      if (fallbackError instanceof Error && fallbackError.message.startsWith('SUCCESS:')) {
        throw fallbackError;
      }
      throw new Error('Failed to download SVG file and fallback method failed');
    }
  }
};

// Download high-resolution PNG file
export const downloadPng = async (svgContent: string, diagramId?: string): Promise<void> => {
  try {
    const filename = `mermaid-diagram-${diagramId || Date.now()}.png`;
    const pngResult = await svgToPng(svgContent);
    const pngDataUrl = `data:image/png;base64,${pngResult.base64}`;
    
    // Create download link with more explicit attributes
    const link = document.createElement('a');
    link.setAttribute('href', pngDataUrl);
    link.setAttribute('download', filename);
    link.style.display = 'none';
    
    // Trigger download
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    console.log('PNG file download triggered');
  } catch (error) {
    console.error('Failed to download PNG:', error);
    // Fallback: open PNG in new window for manual save
    const fallbackFilename = `mermaid-diagram-${diagramId || Date.now()}.png`;
    try {
      const pngResult = await svgToPng(svgContent);
      const pngDataUrl = `data:image/png;base64,${pngResult.base64}`;
      
      const newWindow = window.open('', '_blank');
      if (newWindow) {
        newWindow.document.write(`
          <html>
            <head><title>Save PNG - ${fallbackFilename}</title></head>
            <body style="font-family: Arial; padding: 20px; text-align: center;">
              <h3>PNG Image (Right-click and Save As...)</h3>
              <p>Right-click the image below and select "Save Image As..." to download:</p>
              <img src="${pngDataUrl}" style="max-width: 100%; border: 1px solid #ccc;" alt="Mermaid Diagram">
              <br><br>
              <a href="${pngDataUrl}" download="${fallbackFilename}" style="
                background: #0078d4; 
                color: white; 
                padding: 10px 20px; 
                text-decoration: none; 
                border-radius: 4px;
                display: inline-block;
                margin: 10px;
              ">📥 Click to Download PNG</a>
            </body>
          </html>
        `);
        throw new Error('SUCCESS: PNG opened in new window. Right-click the image and "Save Image As..." to download.');
      } else {
        throw new Error('Could not download PNG file. Please allow popups or check browser download permissions.');
      }
    } catch (fallbackError) {
      if (fallbackError instanceof Error && fallbackError.message.startsWith('SUCCESS:')) {
        throw fallbackError;
      }
      throw new Error('Failed to download PNG file and fallback method failed');
    }
  }
};

// Copy PNG image to clipboard for pasting into PowerPoint
export const copyPngToClipboard = async (svgContent: string): Promise<void> => {
  try {
    // Convert SVG to high-resolution PNG with transparent background
    const pngResult = await svgToPng(svgContent);
    const response = await fetch(`data:image/png;base64,${pngResult.base64}`);
    const blob = await response.blob();
    
    // Copy PNG to clipboard
    await navigator.clipboard.write([
      new ClipboardItem({ 'image/png': blob })
    ]);
    console.log('High-resolution transparent PNG copied to clipboard');
  } catch (pngError) {
    console.log('PNG clipboard failed:', pngError);
    throw new Error('Failed to copy image to clipboard. Browser clipboard permissions may be required.');
  }
};

// Insert new mermaid diagram into PowerPoint slide
export const insertDiagram = async (mermaidCode: string, svgContent: string): Promise<void> => {
  if (!isOfficeContext) {
    console.log('Demo mode: Would insert diagram with code:', mermaidCode);
    return;
  }
  
  const diagramId = generateDiagramId();
  console.log('Starting diagram insertion with ID:', diagramId);
  
  // Primary method: PowerPoint API PNG insertion
  try {
    console.log('Attempting PNG insertion via Office.context API...');
    
    const pngResult = await svgToPng(svgContent);
    console.log('SVG to PNG conversion successful, dimensions:', pngResult.width, 'x', pngResult.height);
    
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load('items');
      await context.sync();
      
      if (slides.items.length === 0) {
        throw new Error('No slide selected. Please select a slide first.');
      }
      
      const slide = slides.items[0];
      const shapes = slide.shapes;
      
      // Use Office.context method for image insertion with correct aspect ratio
      await new Promise<void>((resolve, reject) => {
        if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
          Office.context.document.setSelectedDataAsync(
            pngResult.base64,  // Just the base64 string, not the full data URL
            { 
              coercionType: Office.CoercionType.Image,
              imageLeft: 50,
              imageTop: 50,
              imageWidth: pngResult.width,
              imageHeight: pngResult.height
            },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('PNG insertion via Office.context successful!');
                resolve();
              } else {
                console.error('Office.context insertion failed:', result.error);
                reject(new Error(`Office API insertion failed: ${result.error?.message || 'Unknown error'}`));
              }
            }
          );
        } else {
          reject(new Error('Office context not available'));
        }
      });
      
      // Note: Office.context.document.setSelectedDataAsync doesn't return a shape object
      // We cannot directly tag the inserted image with this method
      // The diagram ID will be stored in custom XML parts for retrieval
      
      console.log('PNG insertion via Office.context successful!');
    });
    
    // Store the diagram data for editing - do this after insertion so we can get accurate shape info
    try {
      const storageDebugLog = await storeDiagramDataAfterInsertion(diagramId, mermaidCode);
      console.log('✅ Diagram data stored successfully');
      console.log('Storage debug log:', storageDebugLog);
    } catch (storageError) {
      console.error('❌ Failed to store diagram data:', storageError);
      // Don't fail the entire insertion just because storage failed
      // But we should surface this error to help debug
      throw new Error(`Image inserted successfully, but metadata storage failed: ${storageError}`);
    }
    
  } catch (apiError) {
    console.error('PowerPoint API insertion failed:', apiError);
    
    // Fallback: Try Office.context method
    try {
      console.log('Trying Office.context insertion method...');
      
      const pngResult = await svgToPng(svgContent);
      
      await new Promise<void>((resolve, reject) => {
        if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
          Office.context.document.setSelectedDataAsync(
            `data:image/png;base64,${pngResult.base64}`,
            { 
              coercionType: Office.CoercionType.Image,
              imageLeft: 50,
              imageTop: 50,
              imageWidth: pngResult.width,
              imageHeight: pngResult.height
            },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('PNG insertion via Office.context successful!');
                resolve();
              } else {
                console.error('Office.context insertion failed:', result.error);
                reject(new Error(`Office API insertion failed: ${result.error?.message || 'Unknown error'}`));
              }
            }
          );
        } else {
          reject(new Error('Office context not available'));
        }
      });
      
      // Store the diagram data
      await storeDiagramData(diagramId, mermaidCode);
      console.log('Diagram inserted via Office.context and data stored');
      
    } catch (officeError) {
      console.error('Office.context insertion also failed:', officeError);
      
      // Final fallback: PNG clipboard with instructions
      try {
        console.log('Trying PNG clipboard as final fallback...');
        await copyPngToClipboard(svgContent);
        
        // Note: For clipboard method, we can't easily associate the shape since 
        // the user will manually paste it. We still store the diagram data 
        // but editing may not work reliably for clipboard-inserted images.
        
        throw new Error('SUCCESS: Could not insert directly. High-resolution PNG copied to clipboard! Go to PowerPoint and paste with Ctrl+V (Cmd+V on Mac).');
        
      } catch (clipboardError) {
        // Check if it's actually a success message
        if (clipboardError instanceof Error && clipboardError.message.startsWith('SUCCESS:')) {
          throw clipboardError;
        }
        console.error('All insertion methods failed:', clipboardError);
        throw new Error('Failed to insert diagram. Please check PowerPoint is running and a slide is selected, or try refreshing the page.');
      }
    }
  }
};

// Helper function to generate the full popup HTML
const getFullPopupHtml = (svgContent: string, diagramId: string): string => {
  const imageBase64Promise = svgToPng(svgContent);
  return `
    <html>
      <head>
        <title>Mermaid Diagram - Copy & Download</title>
        <style>
          body { 
            margin: 0; 
            padding: 20px; 
            font-family: Arial, sans-serif; 
            background: #f5f5f5;
          }
          .container {
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
          }
          .format-section {
            margin: 30px 0;
            padding: 20px;
            border: 1px solid #ddd;
            border-radius: 8px;
            background: #fafafa;
          }
          .format-title {
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 10px;
            color: #333;
          }
          .diagram-container {
            text-align: center;
            margin: 20px 0;
            padding: 20px;
            border: 2px dashed #ddd;
            background: white;
            border-radius: 4px;
          }
          .buttons {
            text-align: center;
            margin: 15px 0;
          }
          button {
            background: #0078d4;
            color: white;
            border: none;
            padding: 12px 20px;
            margin: 0 10px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
          }
          button:hover { background: #106ebe; }
          .primary-btn { background: #0078d4; }
          .secondary-btn { background: #6c757d; }
          .clipboard-btn { background: #28a745; }
          .instructions {
            background: #e6f3ff;
            padding: 15px;
            border-radius: 4px;
            margin: 20px 0;
            border-left: 4px solid #0078d4;
          }
          .recommended {
            background: #d4edda;
            border-left: 4px solid #28a745;
          }
          .note {
            font-size: 12px;
            color: #666;
            margin-top: 10px;
            font-style: italic;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Your Mermaid Diagram</h2>
          
          <div class="instructions recommended">
            <strong>🎯 Best Quality:</strong> Try the "Copy SVG to Clipboard" button first for vector graphics. Use PNG as fallback.
          </div>
          
          <!-- SVG Section (Primary for Desktop PowerPoint) -->
          <div class="format-section">
            <div class="format-title">🎯 SVG Vector (Best Quality - Try This First!)</div>
            <div class="diagram-container">
              ${svgContent}
            </div>
            <div class="buttons">
              <button class="clipboard-btn" onclick="copySVGToClipboard()">📋 Copy SVG to Clipboard</button>
              <button class="primary-btn" onclick="downloadSVG()">Download SVG</button>
            </div>
            <div class="note">SVG provides infinite scaling quality. Copy to clipboard, then paste in PowerPoint with Ctrl+V</div>
          </div>
          
          <!-- PNG Section (Fallback) -->
          <div class="format-section">
            <div class="format-title">🖼️ PNG Image (Fallback Option)</div>
            <div class="diagram-container" id="png-container">
              <p>Loading PNG...</p>
            </div>
            <div class="buttons">
              <button class="secondary-btn" onclick="copyPNG()" id="copy-png-btn" disabled>Copy PNG to Clipboard</button>
              <button class="primary-btn" onclick="downloadPNG()" id="download-png-btn" disabled>Download PNG</button>
            </div>
            <div class="note">High-resolution PNG (4x quality) for universal compatibility</div>
          </div>
          
          <div class="instructions">
            <strong>💡 Instructions:</strong>
            <ol>
              <li><strong>Try SVG first:</strong> Click "Copy SVG to Clipboard", go to PowerPoint, and paste</li>
              <li><strong>If SVG doesn't work:</strong> Use the PNG option or right-click to copy image</li>
              <li><strong>Vector quality:</strong> SVG maintains crisp quality at any size</li>
            </ol>
          </div>
        </div>
        
        <script>
          let pngDataUrl = '';
          
          // Generate PNG on page load
          async function generatePNG() {
            try {
              // This would need to be injected from the parent context
              // For now, show placeholder
              document.getElementById('png-container').innerHTML = '<p>PNG generation in progress...</p>';
              
              // Enable buttons once PNG is ready
              setTimeout(() => {
                document.getElementById('png-container').innerHTML = '<p>PNG ready (implementation needed)</p>';
                document.getElementById('copy-png-btn').disabled = false;
                document.getElementById('download-png-btn').disabled = false;
              }, 1000);
            } catch (err) {
              document.getElementById('png-container').innerHTML = '<p>PNG generation failed</p>';
            }
          }
          
          async function copySVGToClipboard() {
            try {
              const svgContent = \`${svgContent.replace(/`/g, '\\`')}\`;
              await navigator.clipboard.write([
                new ClipboardItem({
                  'text/html': new Blob([\`<div>\${svgContent}</div>\`], { type: 'text/html' }),
                  'image/svg+xml': new Blob([svgContent], { type: 'image/svg+xml' }),
                  'text/plain': new Blob([svgContent], { type: 'text/plain' })
                })
              ]);
              alert('✅ SVG copied to clipboard as vector! Go to PowerPoint and paste with Ctrl+V (Cmd+V on Mac)');
            } catch (err) {
              console.error('SVG clipboard failed:', err);
              alert('❌ SVG clipboard failed. Try the PNG option instead.');
            }
          }
          
          async function copyPNG() {
            try {
              if (!pngDataUrl) {
                alert('PNG not ready yet. Please wait...');
                return;
              }
              const response = await fetch(pngDataUrl);
              const blob = await response.blob();
              await navigator.clipboard.write([
                new ClipboardItem({ 'image/png': blob })
              ]);
              alert('PNG copied to clipboard! Paste into PowerPoint.');
            } catch (err) {
              alert('Copy failed. Right-click the image and select "Copy image".');
            }
          }
          
          function downloadSVG() {
            const svg = \`${svgContent.replace(/`/g, '\\`')}\`;
            const blob = new Blob([svg], { type: 'image/svg+xml' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = 'mermaid-diagram-${diagramId}.svg';
            link.click();
            URL.revokeObjectURL(url);
          }
          
          function downloadPNG() {
            if (!pngDataUrl) {
              alert('PNG not ready yet.');
              return;
            }
            const link = document.createElement('a');
            link.href = pngDataUrl;
            link.download = 'mermaid-diagram-${diagramId}.png';
            link.click();
          }
          
          // Initialize
          generatePNG();
        </script>
      </body>
    </html>
  `;
};

// Update existing mermaid diagram
export const updateDiagram = async (diagramId: string, mermaidCode: string, svgContent: string): Promise<void> => {
  if (!isOfficeContext) {
    console.log('Demo mode: Would update diagram', diagramId, 'with code:', mermaidCode);
    return;
  }
  
  console.log('Starting diagram update for ID:', diagramId);
  const pngResult = await svgToPng(svgContent);
  
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    
    // Find the shape with the matching mermaid diagram tag
    let targetShape = null;
    let targetLeft = 50;
    let targetTop = 50;
    let targetWidth = 600;
    let targetHeight = 450;
    
    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      const shapes = slide.shapes;
      shapes.load('items');
      await context.sync();
      
      for (let j = 0; j < shapes.items.length; j++) {
        const shape = shapes.items[j];
        
        // Check if this shape has our mermaid diagram tag
        try {
          shape.tags.load('items');
          await context.sync();
          
          for (let k = 0; k < shape.tags.items.length; k++) {
            const tag = shape.tags.items[k];
            tag.load(['key', 'value']);
            await context.sync();
            
            if (tag.key.toLowerCase() === 'mermaid_diagram_id' && tag.value === diagramId) {
              console.log('Found diagram shape to update via tag:', diagramId);
              targetShape = shape;
              
              // Get current position and size to maintain them
              shape.load(['left', 'top', 'width', 'height']);
              await context.sync();
              
              targetLeft = shape.left;
              targetTop = shape.top;
              targetWidth = shape.width;
              targetHeight = shape.height;
              
              console.log(`Preserving position: ${targetLeft},${targetTop} size: ${targetWidth}x${targetHeight}`);
              break;
            }
          }
          
          if (targetShape) break;
        } catch (tagError) {
          console.log('Error reading shape tags:', tagError);
          continue;
        }
      }
      
      if (targetShape) break;
    }
    
    if (!targetShape) {
      throw new Error('Diagram shape not found - no shape with matching mermaid_diagram_id tag');
    }
    
    // Delete the old shape
    targetShape.delete();
    await context.sync();
    console.log('Old shape deleted');
    
    // Insert new image at the same position using Office.context
    await new Promise<void>((resolve, reject) => {
      if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
        Office.context.document.setSelectedDataAsync(
          pngResult.base64,
          { 
            coercionType: Office.CoercionType.Image,
            imageLeft: targetLeft,
            imageTop: targetTop,
            imageWidth: targetWidth,
            imageHeight: targetHeight
          },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              console.log('New diagram inserted successfully at preserved position');
              resolve();
            } else {
              reject(new Error(`Failed to insert updated image: ${result.error?.message}`));
            }
          }
        );
      } else {
        reject(new Error('Office context not available'));
      }
    });
    
    // Wait a moment for insertion to complete, then re-tag the new shape
    await new Promise(resolve => setTimeout(resolve, 500));
    
    try {
      // Get the most recently added shape (should be our new image) and tag it
      const slides = context.presentation.getSelectedSlides();
      slides.load('items');
      await context.sync();
      
      if (slides.items.length > 0) {
        const slide = slides.items[0];
        const shapes = slide.shapes;
        shapes.load('items');
        await context.sync();
        
        if (shapes.items.length > 0) {
          const newShape = shapes.items[shapes.items.length - 1];
          newShape.tags.add('mermaid_diagram_id', diagramId);
          await context.sync();
          console.log('New shape re-tagged with diagram ID');
        }
      }
    } catch (retagError) {
      console.warn('Failed to re-tag new shape:', retagError);
    }
    
    // Update stored diagram data
    await updateDiagramData(diagramId, mermaidCode);
    console.log('Diagram data updated in storage');
  });
};

// Debug function to list all stored diagrams
export const listAllStoredDiagrams = async (): Promise<string> => {
  if (!isOfficeContext) {
    return 'Demo mode: Cannot list diagrams';
  }
  
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const customXmlParts = presentation.customXmlParts;
    customXmlParts.load('items');
    await context.sync();
    
    let debugOutput = '=== STORED DIAGRAMS DEBUG ===\n';
    debugOutput += `Total XML parts found: ${customXmlParts.items.length}\n\n`;
    
    let mermaidDiagramCount = 0;
    
    for (let i = 0; i < customXmlParts.items.length; i++) {
      const xmlPart = customXmlParts.items[i];
      
      try {
        // Use getXml() method to retrieve the XML content
        const xmlContent = xmlPart.getXml();
        await context.sync();
        
        debugOutput += `XML Part ${i + 1}: `;
        
        if (xmlContent && xmlContent.value && typeof xmlContent.value === 'string') {
          const xmlString = xmlContent.value;
          debugOutput += `${xmlString.substring(0, 100)}...\n`;
          
          const xmlDoc = new DOMParser().parseFromString(xmlString, 'text/xml');
          
          // Check for parsing errors
          const parserError = xmlDoc.querySelector('parsererror');
          if (parserError) {
            debugOutput += `  Parser error: ${parserError.textContent}\n`;
            continue;
          }
          
          // Look for MermaidDiagram root element
          const mermaidRoot = xmlDoc.querySelector('MermaidDiagram');
          if (mermaidRoot) {
            const idElement = xmlDoc.querySelector('Id');
            const codeElement = xmlDoc.querySelector('Code');
            const shapeInfoElement = xmlDoc.querySelector('ShapeInfo');
            const createdAtElement = xmlDoc.querySelector('CreatedAt');
            
            mermaidDiagramCount++;
            debugOutput += `✅ Mermaid Diagram ${mermaidDiagramCount} found:\n`;
            debugOutput += `  ID: ${idElement?.textContent || 'MISSING'}\n`;
            debugOutput += `  Code length: ${codeElement?.textContent?.length || 0} chars\n`;
            debugOutput += `  Shape info: ${shapeInfoElement?.textContent || 'MISSING'}\n`;
            debugOutput += `  Created: ${createdAtElement?.textContent || 'Unknown'}\n`;
            debugOutput += `  ---\n`;
          } else if (xmlString.includes('MermaidDiagram') || xmlString.includes('mermaid')) {
            debugOutput += `  Contains 'MermaidDiagram' but not parsed correctly\n`;
            debugOutput += `  Root element: ${xmlDoc.documentElement?.tagName || 'NONE'}\n`;
          }
        } else {
          debugOutput += `No XML content found (value: ${xmlContent?.value})\n`;
        }
      } catch (error) {
        debugOutput += `XML part ${i + 1}: Parse error - ${error}\n`;
      }
    }
    
    if (mermaidDiagramCount === 0) {
      debugOutput += '\n❌ No mermaid diagrams found in storage!\n';
      debugOutput += 'Possible issues:\n';
      debugOutput += '1. XML format is incorrect\n';
      debugOutput += '2. XML parsing is failing\n';
      debugOutput += '3. Root element name is wrong\n';
    } else {
      debugOutput += `\n✅ Found ${mermaidDiagramCount} mermaid diagrams total!\n`;
    }
    
    debugOutput += '=== END STORED DIAGRAMS ===';
    return debugOutput;
  });
};

// Get selected shape info for debugging
export const getSelectedShapeInfo = async (): Promise<string> => {
  if (!isOfficeContext) {
    return 'Demo mode: Cannot get shape info';
  }
  
  return PowerPoint.run(async (context) => {
    let output = '=== SELECTED SHAPE DEBUG ===\n';
    
    // Try to get the selected shape
    let selectedShape = null;
    let selectionMethod = '';
    
    try {
      const selection = context.presentation.getSelectedShapes();
      selection.load('items');
      await context.sync();
      
      output += `PowerPoint API selected shapes count: ${selection.items.length}\n`;
      
      if (selection.items.length === 1) {
        selectedShape = selection.items[0];
        selectionMethod = 'PowerPoint API Selection';
        output += `Single shape selected via PowerPoint API\n`;
      } else if (selection.items.length > 1) {
        output += 'Multiple shapes selected - need exactly one\n';
      } else {
        output += 'No shapes selected via PowerPoint API\n';
      }
    } catch (selectionError) {
      output += `PowerPoint API selection failed: ${selectionError}\n`;
    }
    
    // Fallback to most recent shape
    if (!selectedShape) {
      try {
        const slides = context.presentation.getSelectedSlides();
        slides.load('items');
        await context.sync();
        
        if (slides.items.length > 0) {
          const slide = slides.items[0];
          const shapes = slide.shapes;
          shapes.load('items');
          await context.sync();
          
          if (shapes.items.length > 0) {
            selectedShape = shapes.items[shapes.items.length - 1];
            selectionMethod = 'Most Recent Shape (Fallback)';
            output += `Using most recent shape as fallback\n`;
          }
        }
      } catch (fallbackError) {
        output += `Fallback shape detection failed: ${fallbackError}\n`;
      }
    }
    
    if (!selectedShape) {
      output += 'No shape available for analysis\n';
      output += '=== END SELECTED SHAPE DEBUG ===';
      return output;
    }
    
    // Get basic shape info
    try {
      selectedShape.load(['left', 'top', 'width', 'height']);
      await context.sync();
      
      output += `\nShape info (${selectionMethod}):\n`;
      output += `  Left: ${selectedShape.left}\n`;
      output += `  Top: ${selectedShape.top}\n`;
      output += `  Width: ${selectedShape.width}\n`;
      output += `  Height: ${selectedShape.height}\n`;
    } catch (error) {
      output += `Failed to load shape dimensions: ${error}\n`;
    }
    
    // Check shape tags for mermaid diagram ID
    let foundMermaidTag = false;
    try {
      selectedShape.tags.load('items');
      await context.sync();
      
      output += `\nShape tags (${selectedShape.tags.items.length} total):\n`;
      
      for (let i = 0; i < selectedShape.tags.items.length; i++) {
        const tag = selectedShape.tags.items[i];
        tag.load(['key', 'value']);
        await context.sync();
        
        output += `  Tag ${i + 1}: ${tag.key} = ${tag.value}\n`;
        
        if (tag.key.toLowerCase() === 'mermaid_diagram_id') {
          output += `    ✅ MERMAID DIAGRAM TAG FOUND!\n`;
          output += `    Diagram ID: ${tag.value}\n`;
          foundMermaidTag = true;
        }
      }
      
      if (!foundMermaidTag) {
        output += `  ❌ No 'mermaid_diagram_id' tag found\n`;
        output += `  This shape was not created by the mermaid add-in\n`;
      }
    } catch (tagError) {
      output += `Failed to read shape tags: ${tagError}\n`;
      output += `Shape tagging may not be supported\n`;
    }
    
    // If we found a mermaid tag, verify the diagram exists in storage
    if (foundMermaidTag) {
      try {
        const diagramData = await getSelectedDiagram();
        if (diagramData) {
          output += `\n✅ Diagram data found in storage:\n`;
          output += `  ID: ${diagramData.id}\n`;
          output += `  Code length: ${diagramData.code.length} characters\n`;
          output += `  This shape can be edited!\n`;
        } else {
          output += `\n❌ Diagram data not found in storage\n`;
          output += `  The tag exists but metadata is missing\n`;
        }
      } catch (error) {
        output += `\n❌ Error checking diagram data: ${error}\n`;
      }
    }
    
    output += '\n=== END SELECTED SHAPE DEBUG ===';
    return output;
  });
};

// Get selected diagram data for editing
export const getSelectedDiagram = async (): Promise<DiagramData | null> => {
  if (!isOfficeContext) {
    return null;
  }
  
  return PowerPoint.run(async (context) => {
    // Try to get the selected shape and check if it has a mermaid diagram tag
    let selectedShape = null;
    
    try {
      const selection = context.presentation.getSelectedShapes();
      selection.load('items');
      await context.sync();
      
      if (selection.items.length === 1) {
        selectedShape = selection.items[0];
        console.log('Single shape selected via PowerPoint API');
      } else if (selection.items.length > 1) {
        console.log('Multiple shapes selected, need exactly one');
        return null;
      } else {
        console.log('No shapes selected via PowerPoint API');
      }
    } catch (selectionError) {
      console.log('PowerPoint selection API failed:', selectionError);
    }
    
    // If no selection, try fallback to most recent shape
    if (!selectedShape) {
      try {
        const slides = context.presentation.getSelectedSlides();
        slides.load('items');
        await context.sync();
        
        if (slides.items.length > 0) {
          const slide = slides.items[0];
          const shapes = slide.shapes;
          shapes.load('items');
          await context.sync();
          
          if (shapes.items.length > 0) {
            selectedShape = shapes.items[shapes.items.length - 1];
            console.log('Using most recent shape as fallback');
          }
        }
      } catch (fallbackError) {
        console.log('Fallback shape detection failed:', fallbackError);
        return null;
      }
    }
    
    if (!selectedShape) {
      console.log('No shape available for checking');
      return null;
    }
    
    // Check if the selected shape has a mermaid diagram tag
    let diagramId = '';
    try {
      selectedShape.tags.load('items');
      await context.sync();
      
      // Look for our mermaid diagram ID tag
      for (let i = 0; i < selectedShape.tags.items.length; i++) {
        const tag = selectedShape.tags.items[i];
        tag.load(['key', 'value']);
        await context.sync();
        
        if (tag.key.toLowerCase() === 'mermaid_diagram_id') {
          diagramId = tag.value;
          console.log('Found mermaid diagram tag:', diagramId);
          break;
        }
      }
    } catch (tagError) {
      console.log('Failed to read shape tags:', tagError);
    }
    
    // If we found a diagram ID via tag, look up the code
    if (diagramId) {
      const presentation = context.presentation;
      const customXmlParts = presentation.customXmlParts;
      customXmlParts.load('items');
      await context.sync();
      
      for (let i = 0; i < customXmlParts.items.length; i++) {
        const xmlPart = customXmlParts.items[i];
        
        try {
          const xmlContent = xmlPart.getXml();
          await context.sync();
          
          if (xmlContent && xmlContent.value) {
            const xmlDoc = new DOMParser().parseFromString(xmlContent.value, 'text/xml');
            const idElement = xmlDoc.querySelector('Id');
            const codeElement = xmlDoc.querySelector('Code');
            
            if (idElement && codeElement && idElement.textContent === diagramId) {
              console.log('Found matching diagram via tag:', diagramId);
              return {
                id: diagramId,
                code: codeElement.textContent || ''
              };
            }
          }
        } catch (error) {
          continue;
        }
      }
    }
    
    console.log('No mermaid diagram tag found on selected shape');
    return null;
  });
};

// Debug function to check Office context
export const checkOfficeContext = (): string => {
  let debug = '=== OFFICE CONTEXT DEBUG ===\n';
  debug += `typeof Office: ${typeof Office}\n`;
  debug += `typeof PowerPoint: ${typeof PowerPoint}\n`;
  debug += `isOfficeContext: ${isOfficeContext}\n`;
  
  if (typeof Office !== 'undefined') {
    debug += `Office.context exists: ${!!Office.context}\n`;
    if (Office.context) {
      debug += `Office.context.document exists: ${!!Office.context.document}\n`;
      debug += `Office.context.host exists: ${!!Office.context.host}\n`;
      if (Office.context.host) {
        debug += `Host: ${JSON.stringify(Office.context.host)}\n`;
      }
    }
    debug += `Office.CoercionType exists: ${!!Office.CoercionType}\n`;
    debug += `Office.AsyncResultStatus exists: ${!!Office.AsyncResultStatus}\n`;
  } else {
    debug += 'Office API not available!\n';
  }
  
  if (typeof PowerPoint !== 'undefined') {
    debug += `PowerPoint.run exists: ${!!PowerPoint.run}\n`;
  } else {
    debug += 'PowerPoint API not available (this is normal for Office.js apps)\n';
  }
  
  debug += '=== END OFFICE CONTEXT DEBUG ===';
  return debug;
};

// Debug function to test storage directly
export const testDiagramStorage = async (): Promise<string> => {
  if (!isOfficeContext) {
    return `OFFICE CONTEXT ISSUE:\n${checkOfficeContext()}`;
  }
  
  const testId = 'test_' + Date.now();
  const testCode = 'graph TD\n  A[Test] --> B[Storage]';
  
  try {
    const debugLog = await storeDiagramDataAfterInsertion(testId, testCode);
    return `TEST STORAGE RESULT:\n${debugLog}\n\nNow check "List Stored" to see if it appears!`;
  } catch (error) {
    return `TEST STORAGE FAILED:\n${error}`;
  }
};

// Helper function to check if two shape info strings approximately match
const shapesApproximatelyMatch = (info1: string, info2: string): boolean => {
  console.log('Comparing shape info:', info1, 'vs', info2);
  
  const parts1 = info1.split(',').map(parseFloat);
  const parts2 = info2.split(',').map(parseFloat);
  
  if (parts1.length !== 4 || parts2.length !== 4) {
    console.log('Invalid shape info format');
    return false;
  }
  
  const [left1, top1, width1, height1] = parts1;
  const [left2, top2, width2, height2] = parts2;
  
  // Position matching: strict tolerance (within 5 pixels)
  const positionTolerance = 5;
  const leftMatch = Math.abs(left1 - left2) <= positionTolerance;
  const topMatch = Math.abs(top1 - top2) <= positionTolerance;
  
  if (!leftMatch || !topMatch) {
    console.log(`Position mismatch: left diff=${Math.abs(left1 - left2)}, top diff=${Math.abs(top1 - top2)}`);
    return false;
  }
  
  // Size matching: flexible tolerance for PowerPoint auto-resizing
  // Allow up to 50% difference in size (PowerPoint often resizes images significantly)
  const sizeTolerancePercent = 0.5; // 50%
  
  const widthRatio = Math.min(width1, width2) / Math.max(width1, width2);
  const heightRatio = Math.min(height1, height2) / Math.max(height1, height2);
  
  const widthMatch = widthRatio >= (1 - sizeTolerancePercent);
  const heightMatch = heightRatio >= (1 - sizeTolerancePercent);
  
  if (!widthMatch || !heightMatch) {
    console.log(`Size mismatch: width ratio=${widthRatio.toFixed(3)}, height ratio=${heightRatio.toFixed(3)}`);
    console.log(`Required ratio >= ${(1 - sizeTolerancePercent).toFixed(3)}`);
    return false;
  }
  
  console.log('Shape info match found!');
  console.log(`Position match: left diff=${Math.abs(left1 - left2)}, top diff=${Math.abs(top1 - top2)}`);
  console.log(`Size match: width ratio=${widthRatio.toFixed(3)}, height ratio=${heightRatio.toFixed(3)}`);
  return true;
};

// Save settings to Custom XML Parts (platform-agnostic)
export const saveSettings = async (settings: MermaidSettings): Promise<void> => {
  if (!isOfficeContext) {
    console.log('Demo mode: Would save settings:', settings);
    return;
  }

  const platform = detectOfficePlatform();
  console.log('Saving settings for platform:', platform);

  if (platform === OfficePlatform.PowerPoint) {
    return PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const customXmlParts = presentation.customXmlParts;
    
    // First, remove any existing settings XML parts
    customXmlParts.load('items');
    await context.sync();
    
    // Remove existing settings
    for (let i = customXmlParts.items.length - 1; i >= 0; i--) {
      const xmlPart = customXmlParts.items[i];
      try {
        const xmlContent = xmlPart.getXml();
        await context.sync();
        
        if (xmlContent && xmlContent.value && xmlContent.value.includes('<MermaidSettings>')) {
          xmlPart.delete();
          console.log('Removed existing settings XML part');
        }
      } catch (error) {
        // Ignore errors when checking individual parts
        continue;
      }
    }
    
    await context.sync();
    
    // Create new settings XML
    const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<MermaidSettings>
  <FontFamily><![CDATA[${settings.fontFamily}]]></FontFamily>
  <FontSize>${settings.fontSize}</FontSize>
  <PrimaryColor>${settings.primaryColor}</PrimaryColor>
  <PrimaryTextColor>${settings.primaryTextColor}</PrimaryTextColor>
  <PrimaryBorderColor>${settings.primaryBorderColor}</PrimaryBorderColor>
  <LineColor>${settings.lineColor}</LineColor>
  <BackgroundColor>${settings.backgroundColor}</BackgroundColor>
  <SecondaryColor>${settings.secondaryColor}</SecondaryColor>
  <TertiaryColor>${settings.tertiaryColor}</TertiaryColor>
  <Theme>${settings.theme}</Theme>
  <UpdatedAt>${new Date().toISOString()}</UpdatedAt>
</MermaidSettings>`;

    customXmlParts.add(xmlContent);
    await context.sync();
    console.log('Settings saved successfully');
    });
  } else if (platform === OfficePlatform.Word) {
    return Word.run(async (context) => {
      const customXmlParts = context.document.customXmlParts;
      
      // First, remove any existing settings XML parts
      customXmlParts.load('items');
      await context.sync();
      
      // Remove existing settings
      for (let i = customXmlParts.items.length - 1; i >= 0; i--) {
        const xmlPart = customXmlParts.items[i];
        try {
          const xmlContent = xmlPart.getXml();
          await context.sync();
          
          if (xmlContent && xmlContent.value && xmlContent.value.includes('<MermaidSettings>')) {
            xmlPart.delete();
            console.log('Removed existing settings XML part');
          }
        } catch (error) {
          // Ignore errors when checking individual parts
          continue;
        }
      }
      
      await context.sync();
      
      // Create new settings XML (same format as PowerPoint)
      const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<MermaidSettings>
  <FontFamily><![CDATA[${settings.fontFamily}]]></FontFamily>
  <FontSize>${settings.fontSize}</FontSize>
  <PrimaryColor>${settings.primaryColor}</PrimaryColor>
  <PrimaryTextColor>${settings.primaryTextColor}</PrimaryTextColor>
  <PrimaryBorderColor>${settings.primaryBorderColor}</PrimaryBorderColor>
  <LineColor>${settings.lineColor}</LineColor>
  <BackgroundColor>${settings.backgroundColor}</BackgroundColor>
  <SecondaryColor>${settings.secondaryColor}</SecondaryColor>
  <TertiaryColor>${settings.tertiaryColor}</TertiaryColor>
  <Theme>${settings.theme}</Theme>
  <UpdatedAt>${new Date().toISOString()}</UpdatedAt>
</MermaidSettings>`;

      customXmlParts.add(xmlContent);
      await context.sync();
      console.log('Settings saved successfully in Word');
    });
  } else {
    throw new Error(`Unsupported platform for settings: ${platform}`);
  }
};

// Load settings from Custom XML Parts (platform-agnostic)
export const loadSettings = async (): Promise<MermaidSettings> => {
  if (!isOfficeContext) {
    console.log('Demo mode: Using default settings');
    return defaultSettings;
  }

  const platform = detectOfficePlatform();
  console.log('Loading settings for platform:', platform);

  if (platform === OfficePlatform.PowerPoint) {
    return PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      const customXmlParts = presentation.customXmlParts;
    customXmlParts.load('items');
    await context.sync();

    // Look for settings XML part
    for (let i = 0; i < customXmlParts.items.length; i++) {
      const xmlPart = customXmlParts.items[i];
      
      try {
        const xmlContent = xmlPart.getXml();
        await context.sync();
        
        if (xmlContent && xmlContent.value && xmlContent.value.includes('<MermaidSettings>')) {
          const xmlDoc = new DOMParser().parseFromString(xmlContent.value, 'text/xml');
          
          // Check for parsing errors
          const parserError = xmlDoc.querySelector('parsererror');
          if (parserError) {
            console.log('Settings XML parse error:', parserError.textContent);
            continue;
          }
          
          // Extract settings values
          const fontFamilyElement = xmlDoc.querySelector('FontFamily');
          const fontSizeElement = xmlDoc.querySelector('FontSize');
          const primaryColorElement = xmlDoc.querySelector('PrimaryColor');
          const primaryTextColorElement = xmlDoc.querySelector('PrimaryTextColor');
          const primaryBorderColorElement = xmlDoc.querySelector('PrimaryBorderColor');
          const lineColorElement = xmlDoc.querySelector('LineColor');
          const backgroundColorElement = xmlDoc.querySelector('BackgroundColor');
          const secondaryColorElement = xmlDoc.querySelector('SecondaryColor');
          const tertiaryColorElement = xmlDoc.querySelector('TertiaryColor');
          const themeElement = xmlDoc.querySelector('Theme');
          
          if (fontFamilyElement && fontSizeElement && primaryColorElement) {
            const loadedSettings: MermaidSettings = {
              fontFamily: fontFamilyElement.textContent || defaultSettings.fontFamily,
              fontSize: parseInt(fontSizeElement.textContent || '16') || defaultSettings.fontSize,
              primaryColor: primaryColorElement.textContent || defaultSettings.primaryColor,
              primaryTextColor: primaryTextColorElement?.textContent || defaultSettings.primaryTextColor,
              primaryBorderColor: primaryBorderColorElement?.textContent || defaultSettings.primaryBorderColor,
              lineColor: lineColorElement?.textContent || defaultSettings.lineColor,
              backgroundColor: backgroundColorElement?.textContent || defaultSettings.backgroundColor,
              secondaryColor: secondaryColorElement?.textContent || defaultSettings.secondaryColor,
              tertiaryColor: tertiaryColorElement?.textContent || defaultSettings.tertiaryColor,
              theme: (themeElement?.textContent as MermaidSettings['theme']) || defaultSettings.theme
            };
            
            console.log('Settings loaded successfully:', loadedSettings);
            return loadedSettings;
          }
        }
      } catch (error) {
        console.log('Error reading settings XML part:', error);
        continue;
      }
    }
    
    console.log('No settings found, using defaults');
    return defaultSettings;
    });
  } else if (platform === OfficePlatform.Word) {
    return Word.run(async (context) => {
      const customXmlParts = context.document.customXmlParts;
      customXmlParts.load('items');
      await context.sync();

      // Look for settings XML part (same logic as PowerPoint)
      for (let i = 0; i < customXmlParts.items.length; i++) {
        const xmlPart = customXmlParts.items[i];
        
        try {
          const xmlContent = xmlPart.getXml();
          await context.sync();
          
          if (xmlContent && xmlContent.value && xmlContent.value.includes('<MermaidSettings>')) {
            const xmlDoc = new DOMParser().parseFromString(xmlContent.value, 'text/xml');
            
            // Check for parsing errors
            const parserError = xmlDoc.querySelector('parsererror');
            if (parserError) {
              console.log('Settings XML parse error:', parserError.textContent);
              continue;
            }
            
            // Extract settings values
            const fontFamilyElement = xmlDoc.querySelector('FontFamily');
            const fontSizeElement = xmlDoc.querySelector('FontSize');
            const primaryColorElement = xmlDoc.querySelector('PrimaryColor');
            const primaryTextColorElement = xmlDoc.querySelector('PrimaryTextColor');
            const primaryBorderColorElement = xmlDoc.querySelector('PrimaryBorderColor');
            const lineColorElement = xmlDoc.querySelector('LineColor');
            const backgroundColorElement = xmlDoc.querySelector('BackgroundColor');
            const secondaryColorElement = xmlDoc.querySelector('SecondaryColor');
            const tertiaryColorElement = xmlDoc.querySelector('TertiaryColor');
            const themeElement = xmlDoc.querySelector('Theme');
            
            if (fontFamilyElement && fontSizeElement && primaryColorElement) {
              const loadedSettings: MermaidSettings = {
                fontFamily: fontFamilyElement.textContent || defaultSettings.fontFamily,
                fontSize: parseInt(fontSizeElement.textContent || '16') || defaultSettings.fontSize,
                primaryColor: primaryColorElement.textContent || defaultSettings.primaryColor,
                primaryTextColor: primaryTextColorElement?.textContent || defaultSettings.primaryTextColor,
                primaryBorderColor: primaryBorderColorElement?.textContent || defaultSettings.primaryBorderColor,
                lineColor: lineColorElement?.textContent || defaultSettings.lineColor,
                backgroundColor: backgroundColorElement?.textContent || defaultSettings.backgroundColor,
                secondaryColor: secondaryColorElement?.textContent || defaultSettings.secondaryColor,
                tertiaryColor: tertiaryColorElement?.textContent || defaultSettings.tertiaryColor,
                theme: (themeElement?.textContent as MermaidSettings['theme']) || defaultSettings.theme
              };
              
              console.log('Settings loaded successfully from Word:', loadedSettings);
              return loadedSettings;
            }
          }
        } catch (error) {
          console.log('Error reading settings XML part:', error);
          continue;
        }
      }
      
      console.log('No settings found in Word, using defaults');
      return defaultSettings;
    });
  } else {
    console.log('Unsupported platform, using default settings');
    return defaultSettings;
  }
};