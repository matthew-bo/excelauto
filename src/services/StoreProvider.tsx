import React, { createContext, useContext, useEffect } from 'react';
import { create } from 'zustand';
import { persist } from 'zustand/middleware';

interface AppState {
  isLoading: boolean;
  error: string | null;
  promptHistory: string[];
  userPreferences: {
    theme: 'light' | 'dark';
    maxHistoryItems: number;
  };
  setIsLoading: (loading: boolean) => void;
  setError: (error: string | null) => void;
  addPromptToHistory: (prompt: string) => void;
  clearHistory: () => void;
  updatePreferences: (preferences: Partial<AppState['userPreferences']>) => void;
}

const useAppStore = create<AppState>()(
  persist(
    (set) => ({
      isLoading: false,
      error: null,
      promptHistory: [],
      userPreferences: {
        theme: 'light',
        maxHistoryItems: 10,
      },
      setIsLoading: (loading) => set({ isLoading: loading }),
      setError: (error) => set({ error }),
      addPromptToHistory: (prompt) => 
        set((state) => ({ 
          promptHistory: [...state.promptHistory, prompt].slice(-state.userPreferences.maxHistoryItems)
        })),
      clearHistory: () => set({ promptHistory: [] }),
      updatePreferences: (preferences) => 
        set((state) => ({
          userPreferences: { ...state.userPreferences, ...preferences }
        })),
    }),
    {
      name: 'sheetsense-storage',
      partialize: (state) => ({
        promptHistory: state.promptHistory,
        userPreferences: state.userPreferences,
      }),
    }
  )
);

const StoreContext = createContext<AppState | null>(null);

export const useStore = () => {
  const context = useContext(StoreContext);
  if (!context) {
    throw new Error('useStore must be used within StoreProvider');
  }
  return context;
};

interface StoreProviderProps {
  children: React.ReactNode;
}

export const StoreProvider: React.FC<StoreProviderProps> = ({ children }) => {
  const store = useAppStore();
  
  // Clear error after 5 seconds
  useEffect(() => {
    if (store.error) {
      const timer = setTimeout(() => {
        store.setError(null);
      }, 5000);
      return () => clearTimeout(timer);
    }
    return undefined;
  }, [store.error, store.setError]);
  
  return (
    <StoreContext.Provider value={store}>
      {children}
    </StoreContext.Provider>
  );
}; 