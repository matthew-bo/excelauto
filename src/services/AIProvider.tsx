import React, { createContext, useContext, useState, useRef, useCallback } from 'react';
import { aiService, AIRequest, AIResponse } from './AIService';
import { excelOperationsService } from './ExcelOperationsService';

interface AIContextType {
  sendPrompt: (prompt: string) => Promise<AIResponse>;
  isProcessing: boolean;
  lastResponse: AIResponse | null;
  cancelRequest: () => void;
  executeOperations: () => Promise<void>;
}

const AIContext = createContext<AIContextType | null>(null);

export const useAI = () => {
  const context = useContext(AIContext);
  if (!context) {
    throw new Error('useAI must be used within AIProvider');
  }
  return context;
};

interface AIProviderProps {
  children: React.ReactNode;
  maxRetries?: number;
  retryDelay?: number;
}

export const AIProvider: React.FC<AIProviderProps> = ({ 
  children, 
  maxRetries = 3, 
  retryDelay = 1000 
}) => {
  const [isProcessing, setIsProcessing] = useState(false);
  const [lastResponse, setLastResponse] = useState<AIResponse | null>(null);
  const [pendingOperations, setPendingOperations] = useState<AIResponse | null>(null);
  const abortControllerRef = useRef<AbortController | null>(null);

  const cancelRequest = useCallback(() => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
      abortControllerRef.current = null;
    }
    setIsProcessing(false);
  }, []);

  const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

  const sendPrompt = useCallback(async (prompt: string): Promise<AIResponse> => {
    // Cancel any ongoing request
    cancelRequest();
    
    // Create new abort controller
    abortControllerRef.current = new AbortController();
    setIsProcessing(true);

    let lastError: string | undefined;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        // Check if request was cancelled
        if (abortControllerRef.current?.signal.aborted) {
          throw new Error('Request cancelled');
        }

        // Get Excel context
        const excelContext = await excelOperationsService.getContext();
        
        // Create AI request
        const aiRequest: AIRequest = {
          prompt,
          context: {
            ...(excelContext.selectedRange && { selectedRange: excelContext.selectedRange }),
            ...(excelContext.worksheetName && { worksheetName: excelContext.worksheetName }),
          },
        };

        // Process with AI service
        const response = await aiService.processPrompt(aiRequest);
        
        // Check if request was cancelled during processing
        if (abortControllerRef.current?.signal.aborted) {
          throw new Error('Request cancelled');
        }

        setLastResponse(response);
        setPendingOperations(response);
        setIsProcessing(false);
        abortControllerRef.current = null;
        return response;

      } catch (error) {
        lastError = error instanceof Error ? error.message : 'Unknown error';
        
        // Don't retry if request was cancelled
        if (error instanceof Error && error.message === 'Request cancelled') {
          break;
        }

        // Don't retry on last attempt
        if (attempt < maxRetries) {
          console.warn(`AI request failed (attempt ${attempt}/${maxRetries}):`, error);
          await delay(retryDelay * attempt); // Exponential backoff
          continue;
        }
      }
    }

    const errorResponse: AIResponse = {
      success: false,
      error: lastError || 'Request failed after all retries',
      timestamp: Date.now(),
    };

    setLastResponse(errorResponse);
    setIsProcessing(false);
    abortControllerRef.current = null;
    return errorResponse;
  }, [maxRetries, retryDelay, cancelRequest]);

  const executeOperations = useCallback(async (): Promise<void> => {
    if (!pendingOperations?.success || !pendingOperations.data?.excelOperations) {
      return;
    }

    try {
      setIsProcessing(true);
      
      const results = await excelOperationsService.executeOperations(
        pendingOperations.data.excelOperations
      );

      // Check for any failed operations
      const failedOperations = results.filter(result => !result.success);
      
      if (failedOperations.length > 0) {
        console.warn('Some operations failed:', failedOperations);
      }

      // Clear pending operations after execution
      setPendingOperations(null);
      
    } catch (error) {
      console.error('Failed to execute operations:', error);
    } finally {
      setIsProcessing(false);
    }
  }, [pendingOperations]);

  const value: AIContextType = {
    sendPrompt,
    isProcessing,
    lastResponse,
    cancelRequest,
    executeOperations,
  };

  return (
    <AIContext.Provider value={value}>
      {children}
    </AIContext.Provider>
  );
}; 