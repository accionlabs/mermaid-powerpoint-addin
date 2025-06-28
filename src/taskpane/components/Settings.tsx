import * as React from 'react';
import { useState, useEffect } from 'react';

/* global Office */

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

interface SettingsProps {
  settings: MermaidSettings;
  onSettingsChange: (settings: MermaidSettings) => void;
  onClose: () => void;
}

const Settings: React.FC<SettingsProps> = ({ settings, onSettingsChange, onClose }) => {
  const [localSettings, setLocalSettings] = useState<MermaidSettings>(settings);
  const [previewSvg, setPreviewSvg] = useState('');

  const fontOptions = [
    'Arial, sans-serif',
    'Helvetica, Arial, sans-serif', 
    'Times New Roman, serif',
    'Georgia, serif',
    'Courier New, monospace',
    'Verdana, sans-serif',
    'Trebuchet MS, sans-serif',
    'Impact, sans-serif'
  ];

  const themePresets = [
    { name: 'Default', value: 'default' as const },
    { name: 'Dark', value: 'dark' as const },
    { name: 'Forest', value: 'forest' as const },
    { name: 'Custom', value: 'custom' as const }
  ];

  useEffect(() => {
    generatePreview();
  }, [localSettings]);

  const generatePreview = async () => {
    try {
      const { default: mermaid } = await import('mermaid');
      
      // Configure mermaid with current settings
      mermaid.initialize({
        startOnLoad: false,
        theme: localSettings.theme === 'custom' ? 'base' : localSettings.theme,
        themeVariables: localSettings.theme === 'custom' ? {
          primaryColor: localSettings.primaryColor,
          primaryTextColor: localSettings.primaryTextColor,
          primaryBorderColor: localSettings.primaryBorderColor,
          lineColor: localSettings.lineColor,
          secondaryColor: localSettings.secondaryColor,
          tertiaryColor: localSettings.tertiaryColor,
          fontFamily: localSettings.fontFamily,
          fontSize: `${localSettings.fontSize}px`
        } : {},
        securityLevel: 'loose',
        fontFamily: localSettings.fontFamily
      });

      const previewCode = `graph TD
    A[Settings Preview] --> B{Theme: ${localSettings.theme}}
    B --> C[Font: ${localSettings.fontFamily}]
    C --> D[Size: ${localSettings.fontSize}px]`;

      const { svg } = await mermaid.render('settings-preview', previewCode);
      setPreviewSvg(svg);
    } catch (error) {
      console.error('Preview generation failed:', error);
      setPreviewSvg('');
    }
  };

  const handleSettingChange = (key: keyof MermaidSettings, value: any) => {
    const newSettings = { ...localSettings, [key]: value };
    setLocalSettings(newSettings);
  };

  const handleSave = () => {
    onSettingsChange(localSettings);
    onClose();
  };

  const handleReset = () => {
    setLocalSettings(defaultSettings);
  };

  const handlePresetChange = (preset: MermaidSettings['theme']) => {
    let newSettings = { ...localSettings, theme: preset };
    
    // Apply preset-specific defaults
    if (preset === 'dark') {
      newSettings = {
        ...newSettings,
        primaryColor: '#64b5f6',
        primaryTextColor: '#ffffff',
        primaryBorderColor: '#64b5f6',
        lineColor: '#ffffff',
        backgroundColor: '#1e1e1e',
        secondaryColor: '#424242',
        tertiaryColor: '#616161'
      };
    } else if (preset === 'forest') {
      newSettings = {
        ...newSettings,
        primaryColor: '#4caf50',
        primaryTextColor: '#2e7d32',
        primaryBorderColor: '#4caf50',
        lineColor: '#2e7d32',
        backgroundColor: '#f1f8e9',
        secondaryColor: '#c8e6c9',
        tertiaryColor: '#a5d6a7'
      };
    } else if (preset === 'default') {
      newSettings = {
        ...newSettings,
        primaryColor: '#0078d4',
        primaryTextColor: '#000000',
        primaryBorderColor: '#0078d4',
        lineColor: '#000000',
        backgroundColor: '#ffffff',
        secondaryColor: '#e6f3ff',
        tertiaryColor: '#b3d9ff'
      };
    }
    
    setLocalSettings(newSettings);
  };

  return (
    <div style={{ padding: '20px', maxHeight: '600px', overflowY: 'auto' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
        <h2 style={{ margin: 0, color: '#323130' }}>⚙️ Diagram Settings</h2>
        <button 
          onClick={onClose}
          style={{
            background: 'none',
            border: 'none',
            fontSize: '20px',
            cursor: 'pointer',
            color: '#605e5c'
          }}
        >
          ✕
        </button>
      </div>

      {/* Preview Section */}
      <div style={{ marginBottom: '25px' }}>
        <h3 style={{ color: '#323130', marginBottom: '10px' }}>🔍 Live Preview</h3>
        <div 
          style={{ 
            border: '1px solid #edebe9', 
            borderRadius: '4px', 
            padding: '15px',
            backgroundColor: localSettings.backgroundColor,
            minHeight: '150px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
          }}
        >
          {previewSvg ? (
            <div dangerouslySetInnerHTML={{ __html: previewSvg }} />
          ) : (
            <span style={{ color: '#605e5c' }}>Generating preview...</span>
          )}
        </div>
      </div>

      {/* Theme Presets */}
      <div style={{ marginBottom: '25px' }}>
        <h3 style={{ color: '#323130', marginBottom: '10px' }}>🎨 Theme Presets</h3>
        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
          {themePresets.map(preset => (
            <button
              key={preset.value}
              onClick={() => handlePresetChange(preset.value)}
              style={{
                padding: '8px 16px',
                border: `2px solid ${localSettings.theme === preset.value ? '#0078d4' : '#edebe9'}`,
                borderRadius: '4px',
                background: localSettings.theme === preset.value ? '#f3f2f1' : 'white',
                cursor: 'pointer',
                color: '#323130'
              }}
            >
              {preset.name}
            </button>
          ))}
        </div>
      </div>

      {/* Font Settings */}
      <div style={{ marginBottom: '25px' }}>
        <h3 style={{ color: '#323130', marginBottom: '15px' }}>📝 Font Settings</h3>
        
        <div style={{ marginBottom: '15px' }}>
          <label style={{ display: 'block', marginBottom: '5px', color: '#323130', fontWeight: '600' }}>
            Font Family:
          </label>
          <select
            value={localSettings.fontFamily}
            onChange={(e) => handleSettingChange('fontFamily', e.target.value)}
            style={{
              width: '100%',
              padding: '8px',
              border: '1px solid #edebe9',
              borderRadius: '4px',
              fontSize: '14px'
            }}
          >
            {fontOptions.map(font => (
              <option key={font} value={font} style={{ fontFamily: font }}>
                {font.split(',')[0]}
              </option>
            ))}
          </select>
        </div>

        <div style={{ marginBottom: '15px' }}>
          <label style={{ display: 'block', marginBottom: '5px', color: '#323130', fontWeight: '600' }}>
            Font Size: {localSettings.fontSize}px
          </label>
          <input
            type="range"
            min="12"
            max="24"
            value={localSettings.fontSize}
            onChange={(e) => handleSettingChange('fontSize', parseInt(e.target.value))}
            style={{ width: '100%' }}
          />
        </div>
      </div>

      {/* Color Settings - Only show for custom theme */}
      {localSettings.theme === 'custom' && (
        <div style={{ marginBottom: '25px' }}>
          <h3 style={{ color: '#323130', marginBottom: '15px' }}>🎨 Custom Colors</h3>
          
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '15px' }}>
            <div>
              <label style={{ display: 'block', marginBottom: '5px', color: '#323130', fontSize: '12px' }}>
                Primary Color:
              </label>
              <input
                type="color"
                value={localSettings.primaryColor}
                onChange={(e) => handleSettingChange('primaryColor', e.target.value)}
                style={{ width: '100%', height: '35px', border: 'none', borderRadius: '4px' }}
              />
            </div>
            
            <div>
              <label style={{ display: 'block', marginBottom: '5px', color: '#323130', fontSize: '12px' }}>
                Text Color:
              </label>
              <input
                type="color"
                value={localSettings.primaryTextColor}
                onChange={(e) => handleSettingChange('primaryTextColor', e.target.value)}
                style={{ width: '100%', height: '35px', border: 'none', borderRadius: '4px' }}
              />
            </div>
            
            <div>
              <label style={{ display: 'block', marginBottom: '5px', color: '#323130', fontSize: '12px' }}>
                Border Color:
              </label>
              <input
                type="color"
                value={localSettings.primaryBorderColor}
                onChange={(e) => handleSettingChange('primaryBorderColor', e.target.value)}
                style={{ width: '100%', height: '35px', border: 'none', borderRadius: '4px' }}
              />
            </div>
            
            <div>
              <label style={{ display: 'block', marginBottom: '5px', color: '#323130', fontSize: '12px' }}>
                Line Color:
              </label>
              <input
                type="color"
                value={localSettings.lineColor}
                onChange={(e) => handleSettingChange('lineColor', e.target.value)}
                style={{ width: '100%', height: '35px', border: 'none', borderRadius: '4px' }}
              />
            </div>

            <div>
              <label style={{ display: 'block', marginBottom: '5px', color: '#323130', fontSize: '12px' }}>
                Secondary Shapes:
              </label>
              <input
                type="color"
                value={localSettings.secondaryColor}
                onChange={(e) => handleSettingChange('secondaryColor', e.target.value)}
                style={{ width: '100%', height: '35px', border: 'none', borderRadius: '4px' }}
              />
            </div>
            
            <div>
              <label style={{ display: 'block', marginBottom: '5px', color: '#323130', fontSize: '12px' }}>
                Tertiary Shapes:
              </label>
              <input
                type="color"
                value={localSettings.tertiaryColor}
                onChange={(e) => handleSettingChange('tertiaryColor', e.target.value)}
                style={{ width: '100%', height: '35px', border: 'none', borderRadius: '4px' }}
              />
            </div>
          </div>

          <div style={{ marginTop: '15px' }}>
            <label style={{ display: 'block', marginBottom: '5px', color: '#323130', fontSize: '12px' }}>
              Background Color:
            </label>
            <input
              type="color"
              value={localSettings.backgroundColor}
              onChange={(e) => handleSettingChange('backgroundColor', e.target.value)}
              style={{ width: '100%', height: '35px', border: 'none', borderRadius: '4px' }}
            />
          </div>
        </div>
      )}

      {/* Action Buttons */}
      <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end', paddingTop: '20px', borderTop: '1px solid #edebe9' }}>
        <button
          onClick={handleReset}
          style={{
            padding: '10px 20px',
            border: '1px solid #edebe9',
            borderRadius: '4px',
            background: 'white',
            color: '#323130',
            cursor: 'pointer'
          }}
        >
          🔄 Reset
        </button>
        <button
          onClick={onClose}
          style={{
            padding: '10px 20px',
            border: '1px solid #edebe9',
            borderRadius: '4px',
            background: 'white',
            color: '#323130',
            cursor: 'pointer'
          }}
        >
          Cancel
        </button>
        <button
          onClick={handleSave}
          style={{
            padding: '10px 20px',
            border: 'none',
            borderRadius: '4px',
            background: '#0078d4',
            color: 'white',
            cursor: 'pointer',
            fontWeight: '600'
          }}
        >
          💾 Save Settings
        </button>
      </div>
    </div>
  );
};

export default Settings;