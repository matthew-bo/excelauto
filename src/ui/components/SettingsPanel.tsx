import React, { useState, useEffect } from 'react';
import { useStore } from '../../services/StoreProvider';
import { aiService } from '../../services/AIService';

interface SettingsPanelProps {
  isOpen: boolean;
  onClose: () => void;
}

export const SettingsPanel: React.FC<SettingsPanelProps> = ({ isOpen, onClose }) => {
  const { userPreferences, updatePreferences } = useStore();
  const [apiKey, setApiKey] = useState('');
  const [showApiKey, setShowApiKey] = useState(false);
  const [isValidating, setIsValidating] = useState(false);
  const [validationMessage, setValidationMessage] = useState('');

  useEffect(() => {
    // Load current API key from localStorage (masked)
    const currentKey = localStorage.getItem('openai_api_key');
    if (currentKey) {
      setApiKey(currentKey.substring(0, 7) + '...' + currentKey.substring(currentKey.length - 4));
    }
  }, []);

  const handleApiKeyChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const value = e.target.value;
    setApiKey(value);
    setValidationMessage('');
  };

  const handleSaveApiKey = async () => {
    if (!apiKey.trim()) {
      setValidationMessage('API key is required');
      return;
    }

    if (!apiKey.startsWith('sk-')) {
      setValidationMessage('Invalid API key format. Should start with "sk-"');
      return;
    }

    setIsValidating(true);
    setValidationMessage('');

    try {
      // Test the API key with a simple request
      const testResponse = await fetch('https://api.openai.com/v1/models', {
        headers: {
          'Authorization': `Bearer ${apiKey}`,
        },
      });

      if (testResponse.ok) {
        // Save the API key
        aiService.setApiKey(apiKey);
        setValidationMessage('API key saved successfully!');
        setValidationMessage('');
        setTimeout(() => setValidationMessage(''), 3000);
      } else {
        setValidationMessage('Invalid API key. Please check and try again.');
      }
    } catch (error) {
      setValidationMessage('Failed to validate API key. Please check your internet connection.');
    } finally {
      setIsValidating(false);
    }
  };

  const handleClearApiKey = () => {
    localStorage.removeItem('openai_api_key');
    setApiKey('');
    setValidationMessage('API key cleared');
    setTimeout(() => setValidationMessage(''), 3000);
  };

  const handleThemeChange = (theme: 'light' | 'dark') => {
    updatePreferences({ theme });
  };

  const handleMaxHistoryChange = (maxHistoryItems: number) => {
    updatePreferences({ maxHistoryItems });
  };

  if (!isOpen) return null;

  return (
    <div className="settings-overlay" onClick={onClose}>
      <div className="settings-panel" onClick={(e) => e.stopPropagation()}>
        <div className="settings-header">
          <h2>Settings</h2>
          <button className="close-button" onClick={onClose}>
            ×
          </button>
        </div>

        <div className="settings-content">
          {/* API Key Section */}
          <section className="settings-section">
            <h3>OpenAI API Configuration</h3>
            <p className="settings-description">
              Configure your OpenAI API key to enable AI-powered features. 
              Get your API key from the{' '}
              <a 
                href="https://platform.openai.com/api-keys" 
                target="_blank" 
                rel="noopener noreferrer"
                className="link"
              >
                OpenAI Platform
              </a>
            </p>

            <div className="api-key-input-group">
              <div className="input-wrapper">
                <input
                  type={showApiKey ? 'text' : 'password'}
                  value={apiKey}
                  onChange={handleApiKeyChange}
                  placeholder="sk-your-api-key-here"
                  className="api-key-input"
                  disabled={isValidating}
                />
                <button
                  type="button"
                  onClick={() => setShowApiKey(!showApiKey)}
                  className="toggle-visibility"
                  aria-label={showApiKey ? 'Hide API key' : 'Show API key'}
                >
                  {showApiKey ? '👁️' : '👁️‍🗨️'}
                </button>
              </div>

              <div className="api-key-actions">
                <button
                  onClick={handleSaveApiKey}
                  disabled={isValidating || !apiKey.trim()}
                  className="save-button"
                >
                  {isValidating ? 'Validating...' : 'Save API Key'}
                </button>
                <button
                  onClick={handleClearApiKey}
                  className="clear-button"
                  disabled={isValidating}
                >
                  Clear
                </button>
              </div>

              {validationMessage && (
                <div className={`validation-message ${validationMessage.includes('successfully') ? 'success' : 'error'}`}>
                  {validationMessage}
                </div>
              )}
            </div>
          </section>

          {/* Preferences Section */}
          <section className="settings-section">
            <h3>User Preferences</h3>

            <div className="preference-group">
              <label htmlFor="theme-select">Theme:</label>
              <select
                id="theme-select"
                value={userPreferences.theme}
                onChange={(e) => handleThemeChange(e.target.value as 'light' | 'dark')}
                className="preference-select"
              >
                <option value="light">Light</option>
                <option value="dark">Dark</option>
              </select>
            </div>

            <div className="preference-group">
              <label htmlFor="history-limit">Max History Items:</label>
              <select
                id="history-limit"
                value={userPreferences.maxHistoryItems}
                onChange={(e) => handleMaxHistoryChange(Number(e.target.value))}
                className="preference-select"
              >
                <option value={10}>10</option>
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
              </select>
            </div>
          </section>

          {/* Privacy Section */}
          <section className="settings-section">
            <h3>Privacy & Data</h3>
            <div className="privacy-info">
              <p>
                <strong>Data Usage:</strong> Your prompts and Excel context (selected ranges, worksheet names) 
                are sent to OpenAI for processing. No sensitive workbook data is permanently stored.
              </p>
              <p>
                <strong>API Key Storage:</strong> Your API key is stored locally in your browser's localStorage 
                and is only sent to OpenAI for API requests.
              </p>
              <p>
                <strong>No Tracking:</strong> We do not track or store your usage patterns or personal data.
              </p>
            </div>
          </section>

          {/* About Section */}
          <section className="settings-section">
            <h3>About SheetSense</h3>
            <div className="about-info">
              <p><strong>Version:</strong> 1.0.0</p>
              <p><strong>License:</strong> MIT</p>
              <p>
                <strong>Support:</strong>{' '}
                <a 
                  href="https://github.com/your-repo/sheetsense/issues" 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="link"
                >
                  GitHub Issues
                </a>
              </p>
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}; 