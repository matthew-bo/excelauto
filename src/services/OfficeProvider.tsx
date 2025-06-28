import React, { createContext, useContext, useEffect, useState, useCallback } from 'react';

interface OfficeContextType {
  isReady: boolean;
  document: Office.Document | null;
  selection: any; // Using any for now since Office.Range type is not available
  error: string | null;
  refreshSelection: () => void;
}

const OfficeContext = createContext<OfficeContextType>({
  isReady: false,
  document: null,
  selection: null,
  error: null,
  refreshSelection: () => {},
});

export const useOffice = () => useContext(OfficeContext);

interface OfficeProviderProps {
  children: React.ReactNode;
}

export const OfficeProvider: React.FC<OfficeProviderProps> = ({ children }) => {
  const [isReady, setIsReady] = useState(false);
  const [document, setDocument] = useState<Office.Document | null>(null);
  const [selection, setSelection] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);

  const refreshSelection = useCallback(() => {
    if (!document) return;

    try {
      document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setSelection(result.value);
          setError(null);
        } else {
          setError(`Failed to get selection: ${result.error?.message || 'Unknown error'}`);
        }
      });
    } catch (err) {
      setError(`Error refreshing selection: ${err instanceof Error ? err.message : 'Unknown error'}`);
    }
  }, [document]);

  useEffect(() => {
    const initializeOffice = () => {
      try {
        if (typeof Office === 'undefined') {
          setError('Office.js is not available');
          return;
        }

        // Check if Office is already ready
        if (Office.context) {
          setIsReady(true);
          setDocument(Office.context.document);
          
          refreshSelection();
        } else {
          // Wait for Office to be ready
          Office.onReady((info) => {
            console.log('Office.js is ready', info);
            setIsReady(true);
            setDocument(Office.context.document);
            
            refreshSelection();
          });
        }
      } catch (err) {
        setError(`Failed to initialize Office.js: ${err instanceof Error ? err.message : 'Unknown error'}`);
      }
    };

    initializeOffice();
  }, [refreshSelection]);

  const value: OfficeContextType = {
    isReady,
    document,
    selection,
    error,
    refreshSelection,
  };

  return (
    <OfficeContext.Provider value={value}>
      {children}
    </OfficeContext.Provider>
  );
}; 