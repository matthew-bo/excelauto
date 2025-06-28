# SheetSense Architecture Documentation

## System Overview

SheetSense is an Excel Add-in that provides AI-powered natural language interaction with spreadsheets. The architecture follows a modular, service-oriented design with clear separation of concerns.

## Component Architecture

### Core Services

#### 1. AI Service (`AIService.ts`)
- **Purpose**: Handles communication with AI providers (OpenAI, Claude, etc.)
- **Responsibilities**:
  - Process natural language prompts
  - Convert user requests to structured Excel operations
  - Handle API authentication and error handling
  - Provide fallback mock responses for development

#### 2. Excel Operations Service (`ExcelOperationsService.ts`)
- **Purpose**: Manages all Excel workbook interactions
- **Responsibilities**:
  - Execute Excel operations (formulas, formatting, data manipulation)
  - Get current workbook context (selected ranges, worksheet info)
  - Handle Office.js API interactions
  - Provide operation result feedback

#### 3. State Management (`StoreProvider.tsx`)
- **Purpose**: Global application state management using Zustand
- **Responsibilities**:
  - Store user preferences and prompt history
  - Manage loading states and error handling
  - Persist data across sessions
  - Provide reactive state updates

### UI Components

#### 1. App (`App.tsx`)
- **Purpose**: Main application container
- **Responsibilities**:
  - Coordinate between services
  - Handle initialization states
  - Manage error boundaries
  - Provide loading indicators

#### 2. Prompt Input (`PromptInput.tsx`)
- **Purpose**: User input interface
- **Responsibilities**:
  - Capture and validate user prompts
  - Handle keyboard shortcuts (Enter/Escape)
  - Provide real-time validation feedback
  - Manage input state and submission

#### 3. Response Display (`ResponseDisplay.tsx`)
- **Purpose**: Show AI responses and suggestions
- **Responsibilities**:
  - Display AI-generated content
  - Show operation previews
  - Present user suggestions
  - Handle response formatting

#### 4. Operation Executor (`OperationExecutor.tsx`)
- **Purpose**: Execute and manage Excel operations
- **Responsibilities**:
  - Preview pending operations
  - Execute operations with confirmation
  - Show operation results
  - Handle operation failures

## Data Flow

```
User Input → PromptInput → AIProvider → AIService → AI API
                                    ↓
ResponseDisplay ← AIProvider ← AIResponse ← ExcelOperations
                                    ↓
OperationExecutor → ExcelOperationsService → Office.js → Excel
```

## State Management

### Zustand Store Structure
```typescript
interface AppState {
  isLoading: boolean;
  error: string | null;
  promptHistory: string[];
  userPreferences: {
    theme: 'light' | 'dark';
    maxHistoryItems: number;
  };
}
```

### Context Providers
- **OfficeProvider**: Manages Office.js initialization and context
- **AIProvider**: Handles AI service interactions and response management
- **StoreProvider**: Global state management and persistence

## Security Considerations

### API Key Management
- API keys are stored in localStorage (needs improvement)
- Fallback to environment variables
- Mock responses for development without API keys

### Data Privacy
- User prompts are sent to AI services
- Excel data context is included in requests
- No sensitive data is logged or stored permanently

## Performance Optimizations

### Memory Management
- Prompt history is limited by user preferences
- Operation results are cleared after execution
- React components use proper cleanup in useEffect

### Caching Strategy
- Zustand persistence for user preferences
- No aggressive caching to ensure data freshness
- Operation results are not cached

## Error Handling

### Graceful Degradation
- Mock AI responses when API is unavailable
- Fallback UI states for Office.js errors
- User-friendly error messages

### Retry Logic
- AI requests retry up to 3 times with exponential backoff
- Office.js operations have built-in retry mechanisms
- Network failures are handled gracefully

## Testing Strategy

### Unit Tests
- AI service mock responses
- Excel operations validation
- Component rendering tests

### Integration Tests
- End-to-end Office.js interactions
- AI service integration
- State management flows

## Deployment Architecture

### Development
- Local development server on port 3000
- Localtunnel for Office Add-in testing
- Hot reloading for rapid development

### Production
- Static file hosting required
- HTTPS required for Office Add-in
- CDN for asset delivery

## Future Enhancements

### Planned Features
- Offline mode with local AI models
- Collaborative features
- Advanced formula parsing
- Custom AI model support

### Scalability Considerations
- Microservice architecture for AI processing
- Database for prompt history and analytics
- Multi-tenant support for enterprise deployment 