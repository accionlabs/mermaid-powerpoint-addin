import * as React from 'react';
import { createRoot } from 'react-dom/client';
import MermaidEditor from './components/MermaidEditor';

/* global Office */

// Check if Office is available, if not render anyway for browser testing
if (typeof Office !== 'undefined') {
  Office.onReady((info) => {
    const container = document.getElementById('container');
    const root = createRoot(container!);
    root.render(<MermaidEditor />);
  });
} else {
  // Fallback for browser testing without Office.js
  document.addEventListener('DOMContentLoaded', () => {
    const container = document.getElementById('container');
    const root = createRoot(container!);
    root.render(<MermaidEditor />);
  });
}