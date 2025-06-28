import React from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './ui/App';
import { OfficeProvider } from './services/OfficeProvider';
import { AIProvider } from './services/AIProvider';
import { StoreProvider } from './services/StoreProvider';

// Initialize Office.js
Office.onReady((info) => {
  console.log('Office.js is ready', info);
  
  // Render the React app
  const container = document.getElementById('root');
  if (container) {
    const root = createRoot(container);
    root.render(
      <React.StrictMode>
        <OfficeProvider>
          <StoreProvider>
            <AIProvider>
              <App />
            </AIProvider>
          </StoreProvider>
        </OfficeProvider>
      </React.StrictMode>
    );
  }
});

// Handle Office.js errors - only after Office is ready
Office.onReady(() => {
  try {
    // Check if we're in an Office environment with document context
    if (Office.context && Office.context.document && Office.context.document.addHandlerAsync) {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        (eventArgs: any) => {
          console.log('Selection changed', eventArgs);
        }
      );
    } else {
      console.log('Running outside Office environment - document handlers not available');
    }
  } catch (error) {
    console.error('Failed to add document handler:', error);
  }
}); 