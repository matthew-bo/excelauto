import React, { useState } from 'react';
import { useOffice } from '../services/OfficeProvider';
import { useStore } from '../services/StoreProvider';
import { useAI } from '../services/AIProvider';
import { PromptInput } from './components/PromptInput';
import { ResponseDisplay } from './components/ResponseDisplay';
import { LoadingSpinner } from './components/LoadingSpinner';
import { ErrorBoundary } from './components/ErrorBoundary';
import { OperationExecutor } from './components/OperationExecutor';
import { SettingsPanel } from './components/SettingsPanel';
import '../styles/App.css';

export const App: React.FC = () => {
  const { isReady, error: officeError } = useOffice();
  const { isLoading, error: storeError } = useStore();
  const { isProcessing } = useAI();
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);

  // Combine all errors
  const error = officeError || storeError;

  if (!isReady) {
    return (
      <div className="app-container">
        <div className="loading-container">
          <LoadingSpinner />
          <p>Initializing SheetSense...</p>
          {officeError && (
            <p className="error-text">Error: {officeError}</p>
          )}
        </div>
      </div>
    );
  }

  return (
    <ErrorBoundary>
      <div className="app-container">
        <header className="app-header">
          <div className="header-content">
            <div className="header-left">
              <h1>SheetSense</h1>
              <p>AI-powered Excel assistant</p>
            </div>
            <button
              className="settings-button"
              onClick={() => setIsSettingsOpen(true)}
              aria-label="Open settings"
            >
              ⚙️
            </button>
          </div>
        </header>
        
        <main className="app-main">
          {error && (
            <div className="error-message">
              <strong>Error:</strong> {error}
            </div>
          )}
          
          <PromptInput />
          
          {(isLoading || isProcessing) && (
            <div className="processing-container">
              <LoadingSpinner />
              <p>Processing your request...</p>
            </div>
          )}
          
          <ResponseDisplay />
          <OperationExecutor />
        </main>

        <SettingsPanel 
          isOpen={isSettingsOpen} 
          onClose={() => setIsSettingsOpen(false)} 
        />
      </div>
    </ErrorBoundary>
  );
}; 