import React, { useState, useCallback, useEffect } from 'react';
import { useAI } from '../../services/AIProvider';
import { useStore } from '../../services/StoreProvider';

export const PromptInput: React.FC = () => {
  const [prompt, setPrompt] = useState('');
  const [isValid, setIsValid] = useState(false);
  const { sendPrompt, isProcessing, cancelRequest } = useAI();
  const { addPromptToHistory, setError } = useStore();

  // Validate prompt
  useEffect(() => {
    setIsValid(prompt.trim().length >= 3);
  }, [prompt]);

  const handleSubmit = useCallback(async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isValid || isProcessing) return;

    try {
      addPromptToHistory(prompt);
      const response = await sendPrompt(prompt);
      
      if (response.success) {
        setPrompt('');
      } else {
        setError(response.error || 'Failed to process request');
      }
    } catch (error) {
      setError(`Unexpected error: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }, [prompt, isValid, isProcessing, sendPrompt, addPromptToHistory, setError]);

  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    // Submit on Enter
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit(e as any);
    }
    
    // Cancel on Escape
    if (e.key === 'Escape' && isProcessing) {
      e.preventDefault();
      cancelRequest();
    }
  }, [handleSubmit, isProcessing, cancelRequest]);

  const handleCancel = useCallback(() => {
    cancelRequest();
  }, [cancelRequest]);

  return (
    <div className="prompt-input-container">
      <form onSubmit={handleSubmit}>
        <div className="input-group">
          <input
            type="text"
            value={prompt}
            onChange={(e) => setPrompt(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder="Ask me anything about your spreadsheet... (min 3 characters)"
            disabled={isProcessing}
            className={`prompt-input ${!isValid && prompt.length > 0 ? 'invalid' : ''}`}
            maxLength={500}
            aria-label="Enter your prompt"
          />
          {isProcessing ? (
            <button
              type="button"
              onClick={handleCancel}
              className="cancel-button"
              aria-label="Cancel request"
            >
              Cancel
            </button>
          ) : (
            <button
              type="submit"
              disabled={!isValid}
              className="submit-button"
              aria-label="Send prompt"
            >
              Send
            </button>
          )}
        </div>
        {prompt.length > 0 && !isValid && (
          <div className="validation-message">
            Prompt must be at least 3 characters long
          </div>
        )}
        <div className="input-hints">
          <small>
            Press Enter to send • Press Escape to cancel • Minimum 3 characters
          </small>
        </div>
      </form>
    </div>
  );
}; 