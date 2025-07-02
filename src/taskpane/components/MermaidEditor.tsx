import * as React from 'react';
import { useState, useEffect } from 'react';
import mermaid from 'mermaid';
import { insertDiagram, updateDiagram, getSelectedDiagram, listAllStoredDiagrams, getSelectedShapeInfo, testDiagramStorage, checkOfficeContext, loadSettings, saveSettings, MermaidSettings, defaultSettings } from '../utils/powerPointUtils';
import Settings from './Settings';

/* global Office */

const defaultMermaidCode = `graph TD
    A[Start] --> B{Decision}
    B -->|Yes| C[Action 1]
    B -->|No| D[Action 2]
    C --> E[End]
    D --> E`;

const MermaidEditor: React.FC = () => {
  const [mermaidCode, setMermaidCode] = useState(defaultMermaidCode);
  const [svgContent, setSvgContent] = useState('');
  const [error, setError] = useState('');
  const [successMessage, setSuccessMessage] = useState('');
  const [showSvgCode, setShowSvgCode] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [selectedDiagramId, setSelectedDiagramId] = useState<string | null>(null);
  const [debugInfo, setDebugInfo] = useState('');
  const [showDebugMode, setShowDebugMode] = useState(false);
  const [settings, setSettings] = useState<MermaidSettings>(defaultSettings);
  const [showSettings, setShowSettings] = useState(false);

  useEffect(() => {
    initializeMermaidAndSettings();
  }, []);

  useEffect(() => {
    // Apply settings whenever they change
    applyMermaidSettings();
  }, [settings]);

  const initializeMermaidAndSettings = async () => {
    try {
      // Load saved settings
      const loadedSettings = await loadSettings();
      setSettings(loadedSettings);
      
      // Initialize mermaid with loaded settings
      applyMermaidSettings(loadedSettings);
      
      // Check if there's a selected diagram to edit
      checkSelectedDiagram();
      
      // Initial render
      setTimeout(() => renderMermaid(), 100);
    } catch (error) {
      console.error('Failed to load settings:', error);
      // Use default settings if loading fails
      applyMermaidSettings(defaultSettings);
      checkSelectedDiagram();
      setTimeout(() => renderMermaid(), 100);
    }
  };

  const applyMermaidSettings = (settingsToApply: MermaidSettings = settings) => {
    const mermaidConfig = {
      startOnLoad: false,
      theme: settingsToApply.theme === 'custom' ? 'base' : settingsToApply.theme,
      securityLevel: 'loose' as const,
      fontFamily: settingsToApply.fontFamily
    };

    // Add theme variables for custom theme
    if (settingsToApply.theme === 'custom') {
      (mermaidConfig as any).themeVariables = {
        primaryColor: settingsToApply.primaryColor,
        primaryTextColor: settingsToApply.primaryTextColor,
        primaryBorderColor: settingsToApply.primaryBorderColor,
        lineColor: settingsToApply.lineColor,
        secondaryColor: settingsToApply.secondaryColor,
        tertiaryColor: settingsToApply.tertiaryColor,
        fontFamily: settingsToApply.fontFamily,
        fontSize: `${settingsToApply.fontSize}px`,
        
        // Timeline-specific colors (use primary color for all timeline sections)
        cScale0: settingsToApply.primaryColor,
        cScale1: settingsToApply.secondaryColor,
        cScale2: settingsToApply.tertiaryColor,
        cScale3: settingsToApply.primaryColor,
        cScale4: settingsToApply.secondaryColor,
        cScale5: settingsToApply.tertiaryColor,
        cScale6: settingsToApply.primaryColor,
        cScale7: settingsToApply.secondaryColor,
        cScale8: settingsToApply.tertiaryColor,
        cScale9: settingsToApply.primaryColor,
        cScale10: settingsToApply.secondaryColor,
        cScale11: settingsToApply.tertiaryColor,
        
        // Timeline text colors
        cScaleLabel0: settingsToApply.primaryTextColor,
        cScaleLabel1: settingsToApply.primaryTextColor,
        cScaleLabel2: settingsToApply.primaryTextColor,
        cScaleLabel3: settingsToApply.primaryTextColor,
        cScaleLabel4: settingsToApply.primaryTextColor,
        cScaleLabel5: settingsToApply.primaryTextColor,
        cScaleLabel6: settingsToApply.primaryTextColor,
        cScaleLabel7: settingsToApply.primaryTextColor,
        cScaleLabel8: settingsToApply.primaryTextColor,
        cScaleLabel9: settingsToApply.primaryTextColor,
        cScaleLabel10: settingsToApply.primaryTextColor,
        cScaleLabel11: settingsToApply.primaryTextColor
      };
    }

    mermaid.initialize(mermaidConfig);
  };

  // Removed auto-refresh on code change - now only manual refresh

  const checkSelectedDiagram = async () => {
    try {
      const diagramData = await getSelectedDiagram();
      if (diagramData) {
        setMermaidCode(diagramData.code);
        setSelectedDiagramId(diagramData.id);
        setIsEditing(true);
      }
    } catch (error) {
      // No diagram selected, continue with new diagram flow
      console.log('No diagram selected for editing');
    }
  };

  const renderMermaid = async () => {
    try {
      // Clear previous error but don't show new errors immediately
      setError('');
      
      // Basic validation - don't render if code looks incomplete
      const trimmedCode = mermaidCode.trim();
      if (trimmedCode.length < 5) {
        // Code is too short, probably still typing
        return;
      }
      
      // Create a unique ID for this render
      const renderID = 'mermaid-preview-' + Date.now();
      
      // Use mermaid.render with proper error handling
      const { svg } = await mermaid.render(renderID, mermaidCode);
      setSvgContent(svg);
    } catch (err) {
      // Only show error if it looks like user is done typing (debounced)
      const errorMessage = err instanceof Error ? err.message : 'Invalid mermaid syntax';
      setError(`Syntax error: ${errorMessage}`);
      setSvgContent('');
    }
  };

  const generateSvgForInsertion = async (): Promise<string> => {
    // Basic validation - check if code looks valid
    const trimmedCode = mermaidCode.trim();
    if (trimmedCode.length < 5) {
      throw new Error('Code is too short - please enter a complete Mermaid diagram');
    }
    
    // Create a unique ID for this render
    const renderID = 'mermaid-insertion-' + Date.now();
    
    // Use mermaid.render to generate SVG
    const { svg } = await mermaid.render(renderID, mermaidCode);
    return svg;
  };

  const handleInsert = async () => {
    let svgToUse = svgContent;
    
    // Auto-generate preview if it doesn't exist
    if (!svgToUse) {
      setError('');
      setSuccessMessage('Generating preview and inserting diagram...');
      
      try {
        // Generate the SVG directly for insertion
        svgToUse = await generateSvgForInsertion();
        
        // Also update the preview state for the UI
        setSvgContent(svgToUse);
      } catch (err) {
        const errorMessage = err instanceof Error ? err.message : 'Failed to generate diagram';
        setError(`Failed to generate diagram: ${errorMessage}`);
        setSuccessMessage('');
        return;
      }
    }

    try {
      setError('');
      if (!svgToUse) {
        setSuccessMessage('');
      }
      
      console.log('Starting diagram insertion...');
      
      if (isEditing && selectedDiagramId) {
        console.log('Updating existing diagram:', selectedDiagramId);
        await updateDiagram(selectedDiagramId, mermaidCode, svgToUse);
        setSuccessMessage('Diagram updated successfully!');
        // Keep the current state - don't reset after update
        setError('');
      } else {
        console.log('Inserting new diagram...');
        await insertDiagram(mermaidCode, svgToUse);
        setSuccessMessage('Diagram inserted successfully!');
        // Only reset after inserting a new diagram
        setIsEditing(false);
        setSelectedDiagramId(null);
        setMermaidCode(defaultMermaidCode);
        setSvgContent('');
        setError('');
      }
      
      // Clear success message after 5 seconds
      setTimeout(() => {
        setSuccessMessage('');
      }, 5000);
      
    } catch (err) {
      console.error('Insert diagram error:', err);
      const errorMessage = err instanceof Error ? err.message : 'Failed to insert diagram';
      
      if (errorMessage.startsWith('SUCCESS:')) {
        // This is actually a success message with clipboard instructions
        setError('');
        setSuccessMessage(errorMessage.replace('SUCCESS: ', ''));
        
        // Don't clear the preview or reset form for clipboard fallback
        // User needs to manually paste in PowerPoint and might want to try direct insertion again
        setTimeout(() => {
          setSuccessMessage('');
        }, 10000);
      } else {
        setError(`Insertion failed: ${errorMessage}`);
        setSuccessMessage('');
        // Don't clear preview on error - user might want to try again
      }
    }
  };

  const handleNewDiagram = () => {
    setIsEditing(false);
    setSelectedDiagramId(null);
    setMermaidCode(defaultMermaidCode);
    setError('');
    setSuccessMessage('');
    setSvgContent(''); // Clear preview
  };

  const handleCheckSelectedDiagram = async () => {
    setError('');
    setSuccessMessage('');
    try {
      const diagramData = await getSelectedDiagram();
      if (diagramData) {
        setMermaidCode(diagramData.code);
        setSelectedDiagramId(diagramData.id);
        setIsEditing(true);
        setSuccessMessage('Diagram loaded for editing! Modify the code and click "Update Preview", then "Update Diagram".');
        setTimeout(() => setSuccessMessage(''), 8000);
        // Auto-refresh preview when loading for editing
        setTimeout(() => renderMermaid(), 100);
      } else {
        setError('No editable mermaid diagram selected. Please select a diagram that was inserted using this add-in. If you just inserted a diagram, wait a moment and try again (the metadata may still be processing).');
      }
    } catch (error) {
      setError('Failed to load selected diagram. Only diagrams with stored metadata can be edited. For clipboard-inserted images, please create a new diagram.');
    }
  };

  const handleShowSvgCode = () => {
    if (!svgContent) {
      setError('No valid diagram to show');
      return;
    }
    
    setShowSvgCode(!showSvgCode);
    setError('');
    if (!showSvgCode) {
      setSuccessMessage('SVG code displayed below. Copy it and save as .svg file, then insert into PowerPoint!');
      setTimeout(() => setSuccessMessage(''), 8000);
    }
  };

  const copySvgCode = async () => {
    if (!svgContent) {
      setError('No SVG content to copy');
      return;
    }
    
    try {
      await navigator.clipboard.writeText(svgContent);
      setSuccessMessage('SVG code copied to clipboard! Paste into a text editor and save as .svg file');
      setTimeout(() => setSuccessMessage(''), 8000);
    } catch (err) {
      setError('Failed to copy SVG code to clipboard');
    }
  };

  const handleManualRefresh = () => {
    setError('');
    renderMermaid();
  };

  const handleDebugDiagrams = async () => {
    setError('');
    setSuccessMessage('');
    setDebugInfo('');
    try {
      const debugOutput = await listAllStoredDiagrams();
      setDebugInfo(debugOutput);
      setSuccessMessage('Debug info displayed below');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (error) {
      setError('Failed to list stored diagrams');
    }
  };

  const handleDebugSelectedShape = async () => {
    setError('');
    setSuccessMessage('');
    setDebugInfo('');
    try {
      const debugOutput = await getSelectedShapeInfo();
      setDebugInfo(debugOutput);
      setSuccessMessage('Debug info displayed below');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (error) {
      setError('Failed to get selected shape info');
    }
  };

  const handleTestStorage = async () => {
    setError('');
    setSuccessMessage('');
    setDebugInfo('');
    try {
      const debugOutput = await testDiagramStorage();
      setDebugInfo(debugOutput);
      setSuccessMessage('Storage test completed - check debug output');
      setTimeout(() => setSuccessMessage(''), 3000);
    } catch (error) {
      setError('Failed to test storage');
    }
  };

  const handleCheckOfficeContext = () => {
    setError('');
    setSuccessMessage('');
    setDebugInfo('');
    const debugOutput = checkOfficeContext();
    setDebugInfo(debugOutput);
    setSuccessMessage('Office context check completed');
    setTimeout(() => setSuccessMessage(''), 3000);
  };

  const handleSettingsChange = async (newSettings: MermaidSettings) => {
    try {
      setSettings(newSettings);
      await saveSettings(newSettings);
      setSuccessMessage('Settings saved successfully! New diagrams will use these settings.');
      setTimeout(() => setSuccessMessage(''), 5000);
      
      // Re-render current preview with new settings
      setTimeout(() => renderMermaid(), 100);
    } catch (error) {
      setError('Failed to save settings. Please try again.');
      console.error('Settings save error:', error);
    }
  };

  const handleOpenSettings = () => {
    setShowSettings(true);
    setError('');
    setSuccessMessage('');
  };

  const handleCloseSettings = () => {
    setShowSettings(false);
  };

  return (
    <div style={{ padding: '20px', display: 'flex', flexDirection: 'column', height: '100vh' }}>
      {/* Header Section */}
      <div style={{ marginBottom: '15px' }}>
        <h2 style={{ margin: '0 0 15px 0', fontSize: '18px', textAlign: 'center' }}>
          {isEditing ? 'Edit Mermaid Diagram' : 'Insert Mermaid Diagram'}
        </h2>
        
        {/* Main Action Buttons Row */}
        <div style={{ display: 'flex', gap: '8px', width: '100%', marginBottom: '15px' }}>
          <button 
            onClick={handleNewDiagram}
            title="Create New Diagram"
            style={{
              padding: '12px 16px',
              backgroundColor: '#6c757d',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              flex: '1',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}>
            ✨
          </button>
          
          <button 
            onClick={handleCheckSelectedDiagram}
            title="Edit Selected Diagram"
            style={{
              padding: '12px 16px',
              backgroundColor: '#28a745',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              flex: '1',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}>
            📝
          </button>

          <button 
            onClick={handleOpenSettings}
            title="Diagram Settings"
            style={{
              padding: '12px 16px',
              backgroundColor: '#17a2b8',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              flex: '1',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}>
            ⚙️
          </button>

          <button 
            onClick={() => setShowDebugMode(!showDebugMode)}
            title="Toggle Debug Tools"
            style={{
              padding: '12px 16px',
              backgroundColor: showDebugMode ? '#ffc107' : '#6c757d',
              color: showDebugMode ? '#000' : 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              flex: '1',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}>
            🔧
          </button>
        </div>

        {/* Debug Tools (hidden by default) */}
        {showDebugMode && (
          <div style={{ 
            border: '1px solid #ddd', 
            borderRadius: '4px', 
            padding: '10px', 
            backgroundColor: '#f8f9fa',
            marginBottom: '10px'
          }}>
            <div style={{ fontSize: '12px', fontWeight: 'bold', color: '#666', marginBottom: '8px' }}>
              Debug Tools:
            </div>
            <div style={{ display: 'flex', gap: '5px', flexWrap: 'wrap' }}>
              <button 
                onClick={handleDebugDiagrams}
                title="List Stored Diagrams"
                style={{
                  padding: '8px 10px',
                  backgroundColor: '#6c757d',
                  color: 'white',
                  border: 'none',
                  borderRadius: '3px',
                  cursor: 'pointer',
                  fontSize: '12px',
                  fontWeight: 'normal',
                  flex: '1 0 45%',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '4px'
                }}>
                🔍 List
              </button>
              <button 
                onClick={handleDebugSelectedShape}
                title="Check Selected Shape"
                style={{
                  padding: '8px 10px',
                  backgroundColor: '#17a2b8',
                  color: 'white',
                  border: 'none',
                  borderRadius: '3px',
                  cursor: 'pointer',
                  fontSize: '12px',
                  fontWeight: 'normal',
                  flex: '1 0 45%',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '4px'
                }}>
                🎯 Check
              </button>
              <button 
                onClick={handleTestStorage}
                title="Test Storage System"
                style={{
                  padding: '8px 10px',
                  backgroundColor: '#dc3545',
                  color: 'white',
                  border: 'none',
                  borderRadius: '3px',
                  cursor: 'pointer',
                  fontSize: '12px',
                  fontWeight: 'normal',
                  flex: '1 0 45%',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '4px'
                }}>
                🧪 Test
              </button>
              <button 
                onClick={handleCheckOfficeContext}
                title="Check Office Context"
                style={{
                  padding: '8px 10px',
                  backgroundColor: '#ffc107',
                  color: '#000',
                  border: 'none',
                  borderRadius: '3px',
                  cursor: 'pointer',
                  fontSize: '12px',
                  fontWeight: 'normal',
                  flex: '1 0 45%',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '4px'
                }}>
                🏢 Office
              </button>
            </div>
            <div style={{
              fontSize: '10px',
              color: '#666',
              marginTop: '8px',
              fontStyle: 'italic'
            }}>
              Note: Only API-inserted diagrams can be edited (not clipboard PNGs)
            </div>
          </div>
        )}
      </div>

      {/* Code Editor Section */}
      <div style={{ marginBottom: '15px' }}>
        <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold' }}>
          Mermaid Code:
        </label>
        <textarea
          value={mermaidCode}
          onChange={(e) => setMermaidCode(e.target.value)}
          style={{
            width: '100%',
            height: '180px',
            fontFamily: 'Consolas, Monaco, monospace',
            fontSize: '12px',
            padding: '10px',
            border: '1px solid #ccc',
            borderRadius: '3px',
            resize: 'vertical',
            marginBottom: '8px'
          }}
          placeholder="Enter your mermaid diagram code here..."
        />
        <div style={{ display: 'flex', gap: '8px', width: '100%' }}>
          <button
            onClick={handleManualRefresh}
            title="Update Preview"
            style={{
              padding: '12px 16px',
              backgroundColor: '#0078d4',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              flex: '1',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}
          >
            🔄
          </button>

          <button
            onClick={handleInsert}
            title={isEditing ? 'Update Diagram (auto-generates preview if needed)' : 'Insert Diagram (auto-generates preview if needed)'}
            style={{
              padding: '12px 16px',
              backgroundColor: '#0078d4',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: 'bold',
              flex: '1',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}>
            {isEditing ? '📝' : '📊'}
          </button>

          <button
            onClick={handleShowSvgCode}
            disabled={!svgContent}
            title={showSvgCode ? 'Hide SVG Code' : 'Show SVG Code'}
            style={{
              padding: '12px 16px',
              backgroundColor: svgContent ? '#28a745' : '#ccc',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: svgContent ? 'pointer' : 'not-allowed',
              fontSize: '16px',
              fontWeight: 'bold',
              flex: '1',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}>
            📄
          </button>
        </div>
      </div>

      {/* Error/Success Messages */}
      {error && (
        <div style={{
          padding: '10px',
          backgroundColor: '#ffebee',
          border: '1px solid #f44336',
          borderRadius: '3px',
          color: '#c62828',
          marginBottom: '10px',
          fontSize: '12px'
        }}>
          {error}
        </div>
      )}

      {successMessage && (
        <div style={{
          padding: '10px',
          backgroundColor: '#d4edda',
          border: '1px solid #28a745',
          borderRadius: '3px',
          color: '#155724',
          marginBottom: '10px',
          fontSize: '12px',
          fontWeight: 'bold'
        }}>
          ✅ {successMessage}
        </div>
      )}


      {/* Preview Section */}
      <div style={{ marginBottom: '10px', flex: 1 }}>
        <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold' }}>
          Preview:
        </label>
        <div style={{
          border: '1px solid #ccc',
          borderRadius: '3px',
          padding: '15px',
          backgroundColor: '#f9f9f9',
          minHeight: '200px',
          overflow: 'auto',
          flex: 1
        }}>
          {svgContent ? (
            <div dangerouslySetInnerHTML={{ __html: svgContent }} />
          ) : (
            <div style={{ color: '#666', fontStyle: 'italic' }}>
              Preview will appear here after you click "Update Preview"...
            </div>
          )}
        </div>
      </div>

      {/* SVG Code Section */}
      {showSvgCode && svgContent && (
        <div style={{
          border: '1px solid #ddd',
          borderRadius: '3px',
          padding: '10px',
          backgroundColor: '#f9f9f9',
          marginBottom: '10px'
        }}>
          <div style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: '10px'
          }}>
            <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#333' }}>
              SVG Code (Copy & Save as .svg file):
            </span>
            <button
              onClick={copySvgCode}
              title="Copy SVG Code"
              style={{
                padding: '6px 12px',
                backgroundColor: '#0078d4',
                color: 'white',
                border: 'none',
                borderRadius: '3px',
                cursor: 'pointer',
                fontSize: '14px',
                display: 'flex',
                alignItems: 'center',
                gap: '4px'
              }}>
              📋 Copy
            </button>
          </div>
          <textarea
            value={svgContent}
            readOnly
            style={{
              width: '100%',
              height: '120px',
              fontFamily: 'Consolas, Monaco, monospace',
              fontSize: '10px',
              padding: '5px',
              border: '1px solid #ccc',
              borderRadius: '3px',
              resize: 'vertical',
              backgroundColor: 'white'
            }}
          />
          <div style={{
            fontSize: '10px',
            color: '#666',
            marginTop: '5px',
            fontStyle: 'italic'
          }}>
            💡 Copy this code, paste into a text editor, save as "diagram.svg", then Insert → Pictures → This Device in PowerPoint
          </div>
        </div>
      )}

      {/* Debug Information */}
      {showDebugMode && debugInfo && (
        <div style={{
          border: '1px solid #ffc107',
          borderRadius: '3px',
          padding: '10px',
          backgroundColor: '#fff3cd',
          marginBottom: '10px'
        }}>
          <div style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: '10px'
          }}>
            <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#856404' }}>
              🔍 Debug Information:
            </span>
            <button
              onClick={() => setDebugInfo('')}
              title="Close Debug Info"
              style={{
                padding: '6px 10px',
                backgroundColor: '#ffc107',
                color: '#000',
                border: 'none',
                borderRadius: '3px',
                cursor: 'pointer',
                fontSize: '14px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>
              ✕
            </button>
          </div>
          <textarea
            value={debugInfo}
            readOnly
            style={{
              width: '100%',
              height: '150px',
              fontFamily: 'Consolas, Monaco, monospace',
              fontSize: '10px',
              padding: '5px',
              border: '1px solid #ffc107',
              borderRadius: '3px',
              resize: 'vertical',
              backgroundColor: 'white'
            }}
          />
        </div>
      )}

      {/* Settings Modal */}
      {showSettings && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0, 0, 0, 0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000
        }}>
          <div style={{
            backgroundColor: 'white',
            borderRadius: '8px',
            maxWidth: '600px',
            maxHeight: '80vh',
            width: '90%',
            overflowY: 'auto',
            boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)'
          }}>
            <Settings
              settings={settings}
              onSettingsChange={handleSettingsChange}
              onClose={handleCloseSettings}
            />
          </div>
        </div>
      )}
    </div>
  );
};

export default MermaidEditor;