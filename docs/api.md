# SheetSense API Documentation

## Overview

This document describes the internal APIs and interfaces used within the SheetSense Excel Add-in.

## Core Interfaces

### AI Service API

#### `AIRequest`
```typescript
interface AIRequest {
  prompt: string;
  context?: {
    selectedRange?: string;
    worksheetName?: string;
    workbookData?: any;
  };
}
```

#### `AIResponse`
```typescript
interface AIResponse {
  success: boolean;
  data?: {
    action: 'explain' | 'create' | 'modify' | 'format' | 'analyze';
    description: string;
    excelOperations: ExcelOperation[];
    suggestions: string[];
  };
  error?: string;
  timestamp: number;
}
```

#### `ExcelOperation`
```typescript
interface ExcelOperation {
  type: 'formula' | 'format' | 'insert' | 'delete' | 'modify';
  target: string; // Cell range or description
  value?: string; // Formula, format, or data
  description: string;
}
```

### Excel Operations API

#### `ExcelContext`
```typescript
interface ExcelContext {
  selectedRange?: string;
  worksheetName?: string;
  activeCell?: string;
}
```

#### `OperationResult`
```typescript
interface OperationResult {
  success: boolean;
  message: string;
  error?: string;
}
```

## Service Methods

### AIService

#### `processPrompt(request: AIRequest): Promise<AIResponse>`
Processes a natural language prompt and returns structured Excel operations.

**Parameters:**
- `request`: AIRequest object containing prompt and context

**Returns:**
- Promise resolving to AIResponse with operations or error

**Example:**
```typescript
const response = await aiService.processPrompt({
  prompt: "Add a sum formula to the selected range",
  context: {
    selectedRange: "A1:A10",
    worksheetName: "Sheet1"
  }
});
```

#### `setApiKey(apiKey: string): void`
Sets the API key for AI service authentication.

### ExcelOperationsService

#### `getContext(): Promise<ExcelContext>`
Retrieves current Excel workbook context.

**Returns:**
- Promise resolving to ExcelContext with current selection and worksheet info

**Example:**
```typescript
const context = await excelOperationsService.getContext();
console.log(context.selectedRange); // "A1:B5"
```

#### `executeOperations(operations: ExcelOperation[]): Promise<OperationResult[]>`
Executes a list of Excel operations.

**Parameters:**
- `operations`: Array of ExcelOperation objects to execute

**Returns:**
- Promise resolving to array of OperationResult objects

**Example:**
```typescript
const results = await excelOperationsService.executeOperations([
  {
    type: 'formula',
    target: 'A1',
    value: '=SUM(B1:B10)',
    description: 'Add sum formula to A1'
  }
]);
```

## React Hooks

### `useAI()`
Provides access to AI service functionality.

**Returns:**
```typescript
{
  sendPrompt: (prompt: string) => Promise<AIResponse>;
  isProcessing: boolean;
  lastResponse: AIResponse | null;
  cancelRequest: () => void;
  executeOperations: () => Promise<void>;
}
```

**Example:**
```typescript
const { sendPrompt, isProcessing, lastResponse } = useAI();

const handleSubmit = async () => {
  const response = await sendPrompt("Add a column with profit margin");
  if (response.success) {
    console.log(response.data?.excelOperations);
  }
};
```

### `useStore()`
Provides access to global application state.

**Returns:**
```typescript
{
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
  updatePreferences: (preferences: Partial<UserPreferences>) => void;
}
```

### `useOffice()`
Provides access to Office.js functionality.

**Returns:**
```typescript
{
  isReady: boolean;
  error: string | null;
}
```

## Error Handling

### Error Types

#### AI Service Errors
- `API_KEY_MISSING`: No API key configured
- `API_REQUEST_FAILED`: Network or API error
- `INVALID_RESPONSE`: Malformed response from AI service
- `RATE_LIMIT_EXCEEDED`: API rate limit reached

#### Excel Operation Errors
- `OFFICE_NOT_READY`: Office.js not initialized
- `INVALID_RANGE`: Invalid cell range specified
- `OPERATION_FAILED`: Excel operation failed
- `PERMISSION_DENIED`: Insufficient permissions

### Error Response Format
```typescript
{
  success: false,
  error: "Human-readable error message",
  timestamp: 1234567890
}
```

## Configuration

### Environment Variables
- `OPENAI_API_KEY`: OpenAI API key for AI service
- `NODE_ENV`: Environment (development/production)

### Local Storage Keys
- `openai_api_key`: Stored API key
- `sheetsense-storage`: Zustand persisted state

## Rate Limiting

### AI Service
- Maximum 3 retries with exponential backoff
- 1 second base delay, doubling on each retry
- Request cancellation support

### Excel Operations
- No built-in rate limiting
- Office.js handles internal throttling
- Operations are executed sequentially

## Security Considerations

### API Key Storage
- API keys are stored in localStorage (client-side)
- No server-side storage of sensitive data
- Keys are included in HTTP headers for API requests

### Data Privacy
- User prompts are sent to external AI services
- Excel context data is included in requests
- No permanent storage of sensitive workbook data

## Testing

### Mock Responses
The AI service provides mock responses when no API key is configured:

```typescript
// Mock response for explain requests
{
  action: 'explain',
  description: 'I\'ll explain the selected formula or data structure.',
  excelOperations: [],
  suggestions: [
    'Select a cell with a formula to get a detailed explanation',
    'Ask "What does this range contain?" to understand your data'
  ]
}
```

### Test Utilities
- `AIService.test.ts`: Unit tests for AI service
- Mock Office.js environment for testing
- Jest configuration for TypeScript and React testing 