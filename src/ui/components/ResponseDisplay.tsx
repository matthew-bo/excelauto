import React from 'react';
import { useStore } from '../../services/StoreProvider';
import { useAI } from '../../services/AIProvider';

export const ResponseDisplay: React.FC = () => {
  const { promptHistory, clearHistory } = useStore();
  const { lastResponse } = useAI();

  if (promptHistory.length === 0) {
    return (
      <div className="response-display">
        <div className="welcome-message">
          <h3>Welcome to SheetSense!</h3>
          <p>Try asking me to:</p>
          <ul>
            <li>"Explain this formula"</li>
            <li>"Add a profit margin column"</li>
            <li>"Create a summary table"</li>
            <li>"Format this data as currency"</li>
            <li>"Copy A1:B5 to C1"</li>
            <li>"Move column B to column D"</li>
            <li>"Sort by column A descending"</li>
            <li>"Filter to show only values &gt; 100"</li>
            <li>"Create a line chart from this data"</li>
            <li>"Remove duplicate values"</li>
          </ul>
          <div className="welcome-tips">
            <p><strong>Tips:</strong></p>
            <ul>
              <li>Select cells in Excel before asking questions</li>
              <li>Be specific about what you want to do</li>
              <li>Use natural language - no need for technical terms</li>
              <li>Specify ranges like "A1:B10" for precise operations</li>
              <li>Use "Copy from X to Y" for data movement</li>
            </ul>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="response-display">
      {lastResponse && (
        <div className="last-response">
          <h4>Last Response</h4>
          <div className={`response-item ${lastResponse.success ? 'success' : 'error'}`}>
            <div className="response-header">
              <span className="response-status">
                {lastResponse.success ? '✓ Success' : '✗ Error'}
              </span>
              <span className="response-time">
                {new Date(lastResponse.timestamp).toLocaleTimeString()}
              </span>
            </div>
            {lastResponse.success ? (
              <div className="response-content">
                <p className="response-description">{lastResponse.data?.description}</p>
                {lastResponse.data?.excelOperations && lastResponse.data.excelOperations.length > 0 && (
                  <div className="operations-preview">
                    <h4>Operations:</h4>
                    <ul>
                      {lastResponse.data.excelOperations.map((op, index) => (
                        <li key={index}>{op.description}</li>
                      ))}
                    </ul>
                  </div>
                )}
                {lastResponse.data?.suggestions && lastResponse.data.suggestions.length > 0 && (
                  <div className="suggestions">
                    <h4>Suggestions:</h4>
                    <ul>
                      {lastResponse.data.suggestions.map((suggestion, index) => (
                        <li key={index}>{suggestion}</li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            ) : (
              <div className="response-content">
                <p className="error-text">{lastResponse.error}</p>
              </div>
            )}
          </div>
        </div>
      )}
      
      <div className="history-section">
        <div className="history-header">
          <h4>Recent Prompts ({promptHistory.length})</h4>
          {promptHistory.length > 0 && (
            <button 
              onClick={clearHistory}
              className="clear-history-button"
              aria-label="Clear history"
            >
              Clear
            </button>
          )}
        </div>
        <div className="prompt-history">
          {promptHistory.map((prompt, index) => (
            <div key={index} className="prompt-item">
              <span className="prompt-text">{prompt}</span>
              <span className="prompt-number">#{promptHistory.length - index}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}; 