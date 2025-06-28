import React, { useEffect } from 'react';
import { useAI } from '../../services/AIProvider';
import { useStore } from '../../services/StoreProvider';

export const OperationExecutor: React.FC = () => {
  const { lastResponse, executeOperations, isProcessing } = useAI();
  const { setError } = useStore();

  // Auto-execute operations when AI responds successfully
  useEffect(() => {
    if (lastResponse?.success && lastResponse.data?.excelOperations?.length && !isProcessing) {
      // Auto-execute after a short delay to let user see what will happen
      const timer = setTimeout(() => {
        handleExecute();
      }, 1000);
      
      return () => clearTimeout(timer);
    }
    return undefined; // Explicit return for when conditions aren't met
  }, [lastResponse, isProcessing]);

  if (!lastResponse?.success || !lastResponse.data?.excelOperations?.length) {
    return null;
  }

  const handleExecute = async () => {
    try {
      await executeOperations();
    } catch (error) {
      setError(`Failed to execute operations: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  };

  const handleSkip = () => {
    // Skip execution - just clear the response
    // This could be enhanced to store skipped operations
  };

  return (
    <div className="operation-executor">
      <div className="operation-header">
        <h4>Ready to Execute</h4>
        <p>{lastResponse.data.description}</p>
      </div>
      
      <div className="operations-list">
        {lastResponse.data.excelOperations.map((operation, index) => (
          <div key={index} className="operation-item">
            <div className="operation-icon">
              {operation.type === 'formula' && '∑'}
              {operation.type === 'format' && '🎨'}
              {operation.type === 'insert' && '➕'}
              {operation.type === 'delete' && '🗑️'}
              {operation.type === 'modify' && '✏️'}
            </div>
            <div className="operation-details">
              <div className="operation-description">{operation.description}</div>
              <div className="operation-target">Target: {operation.target}</div>
              {operation.value && (
                <div className="operation-value">Value: {operation.value}</div>
              )}
            </div>
          </div>
        ))}
      </div>
      
      <div className="operation-actions">
        <button
          onClick={handleExecute}
          disabled={isProcessing}
          className="execute-button"
          aria-label="Execute operations"
        >
          {isProcessing ? 'Executing...' : 'Execute Operations'}
        </button>
        <button
          onClick={handleSkip}
          disabled={isProcessing}
          className="skip-button"
          aria-label="Skip operations"
        >
          Skip
        </button>
      </div>
      
      {lastResponse.data.suggestions && lastResponse.data.suggestions.length > 0 && (
        <div className="suggestions">
          <h5>Suggestions:</h5>
          <ul>
            {lastResponse.data.suggestions.map((suggestion, index) => (
              <li key={index}>{suggestion}</li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}; 