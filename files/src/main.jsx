import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App.jsx';
import './index.css';
import { installStoragePolyfill } from './lib/storage.js';

// ClientMaster.jsx persists data through `window.storage`. Install the
// localStorage-backed polyfill BEFORE the app mounts so the initial data load
// (which runs in a useEffect on first render) finds the API ready.
installStoragePolyfill();

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
