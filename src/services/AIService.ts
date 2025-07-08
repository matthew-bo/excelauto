export interface AIRequest {
  prompt: string;
  context?: {
    selectedRange?: string;
    worksheetName?: string;
    workbookData?: any;
  };
}

export interface AIResponse {
  success: boolean;
  data?: {
    action: 'explain' | 'create' | 'modify' | 'format' | 'analyze' | 'transform' | 'clean';
    description: string;
    excelOperations: ExcelOperation[];
    suggestions: string[];
  };
  error?: string;
  timestamp: number;
}

// Discriminated union for ExcelOperation
export type ExcelOperation =
  | { type: 'formula'; target: string; value: string; description: string; range?: string; options?: any }
  | { type: 'format'; target: string; value: string; description: string; range?: string; options?: any }
  | { type: 'insert'; target: string; description: string; value?: string; range?: string; options?: any }
  | { type: 'delete'; target: string; description: string; value?: string; range?: string; options?: any }
  | { type: 'modify'; target: string; value: string; description: string; range?: string; options?: any }
  | { type: 'copy'; target: string; description: string; value?: string; range?: string; options?: any }
  | { type: 'move'; target: string; description: string; value?: string; range?: string; options?: any }
  | { type: 'filter'; target: string; description: string; value?: string; range?: string; options?: any }
  | { type: 'sort'; target: string; description: string; value?: string; range?: string; options?: any }
  | { type: 'chart'; target: string; description: string; value?: string; range?: string; options?: any }
  | { type: 'table'; target: string; description: string; value?: string; range?: string; options?: any }
  | { type: 'clean'; target: string; description: string; value?: string; range?: string; options?: any };

// Schema validation utilities
export class SchemaValidator {
  static isValidAction(action: any): action is 'explain' | 'create' | 'modify' | 'format' | 'analyze' | 'transform' | 'clean' {
    const validActions = ['explain', 'create', 'modify', 'format', 'analyze', 'transform', 'clean'];
    return typeof action === 'string' && validActions.includes(action);
  }

  static isValidOperationType(type: any): type is ExcelOperation['type'] {
    const validTypes = ['formula', 'format', 'insert', 'delete', 'modify', 'copy', 'move', 'sort', 'filter', 'chart', 'table'];
    return typeof type === 'string' && validTypes.includes(type);
  }

  static isValidCellReference(ref: any): boolean {
    if (typeof ref !== 'string') return false;
    // Basic Excel cell reference validation (A1, B2, etc.)
    const cellRefRegex = /^[A-Z]+\d+$/;
    return cellRefRegex.test(ref.toUpperCase());
  }

  static isValidRangeReference(ref: any): boolean {
    if (typeof ref !== 'string') return false;
    // Basic Excel range validation (A1:B10, etc.)
    const rangeRefRegex = /^[A-Z]+\d+:[A-Z]+\d+$/;
    return rangeRefRegex.test(ref.toUpperCase());
  }

  static validateExcelOperation(operation: any): operation is ExcelOperation {
    if (!operation || typeof operation !== 'object') return false;
    if (!this.isValidOperationType(operation.type)) return false;
    if (!operation.target || !this.isValidCellReference(operation.target)) return false;
    if (!operation.description || typeof operation.description !== 'string') return false;

    // Type-specific validation
    switch (operation.type) {
      case 'formula':
      case 'format':
      case 'modify':
        return operation.value && typeof operation.value === 'string';
      
      case 'insert':
      case 'delete':
      case 'table':
        return true; // Only target and description required
      
      case 'copy':
      case 'move':
        return true; // range is optional
      
      case 'sort':
        return operation.options && 
               typeof operation.options === 'object' &&
               operation.options.sortBy &&
               typeof operation.options.sortBy === 'string';
      
      case 'filter':
        return operation.options && 
               typeof operation.options === 'object' &&
               operation.options.filterCriteria;
      
      case 'chart':
        return operation.options && 
               typeof operation.options === 'object' &&
               operation.options.chartType &&
               typeof operation.options.chartType === 'string';
      
      default:
        return false;
    }
  }

  static validateAIResponse(response: any): response is AIResponse {
    if (!response || typeof response !== 'object') return false;
    if (typeof response.success !== 'boolean') return false;
    if (typeof response.timestamp !== 'number') return false;
    
    if (response.success) {
      if (!response.data || typeof response.data !== 'object') return false;
      if (!this.isValidAction(response.data.action)) return false;
      if (!response.data.description || typeof response.data.description !== 'string') return false;
      if (!Array.isArray(response.data.excelOperations)) return false;
      if (!Array.isArray(response.data.suggestions)) return false;
      
      // Validate each operation
      for (const operation of response.data.excelOperations) {
        if (!this.validateExcelOperation(operation)) return false;
      }
      
      // Validate suggestions
      for (const suggestion of response.data.suggestions) {
        if (typeof suggestion !== 'string') return false;
      }
    } else {
      if (response.error && typeof response.error !== 'string') return false;
    }
    
    return true;
  }
}

export class AIService {
  private apiKey: string | null = null;

  constructor() {
    // Initialize with any stored API key
    this.apiKey = localStorage.getItem('openai_api_key');
  }

  setApiKey(apiKey: string): void {
    this.apiKey = apiKey;
    localStorage.setItem('openai_api_key', apiKey);
  }

  async processPrompt(request: AIRequest): Promise<AIResponse> {
    try {
      // Validate input
      if (!request.prompt || typeof request.prompt !== 'string') {
        return {
          success: false,
          error: 'Invalid prompt: must be a non-empty string',
          timestamp: Date.now(),
        };
      }

      // Check for API key in various sources
      const apiKey = this.apiKey || 
                    (typeof process !== 'undefined' && process.env ? process.env['OPENAI_API_KEY'] : null) ||
                    (typeof window !== 'undefined' && (window as any).OPENAI_API_KEY);
      
      if (apiKey) {
        return await this.openAIResponse(request, apiKey);
      }
      // Fallback to mock if no API key
      return await this.mockAIResponse(request);
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error',
        timestamp: Date.now(),
      };
    }
  }

  private async openAIResponse(request: AIRequest, apiKey: string): Promise<AIResponse> {
    try {
      const systemPrompt = `You are an expert Excel assistant. Given a user prompt and context, return a JSON object with the following structure: { action: (explain|create|modify|format|analyze|transform|clean), description: string, excelOperations: [{ type: (formula|format|insert|delete|modify|copy|move|filter|sort|chart|table), target: string, value?: string, description: string }], suggestions: string[] }. Only return valid JSON.`;
      const userPrompt = `Prompt: ${request.prompt}\nContext: ${JSON.stringify(request.context || {})}`;
      
      const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${apiKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          model: 'gpt-3.5-turbo',
          messages: [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: userPrompt },
          ],
          temperature: 0.2,
          max_tokens: 512,
        }),
      });

      if (!response.ok) {
        throw new Error(`OpenAI API error: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      const content = data.choices[0].message.content;
      
      let parsed;
      try {
        parsed = JSON.parse(content);
      } catch (e) {
        return {
          success: false,
          error: 'Failed to parse AI response as JSON',
          timestamp: Date.now(),
        };
      }
      
      // Validate the parsed response
      if (!SchemaValidator.validateAIResponse(parsed)) {
        return {
          success: false,
          error: 'AI response failed schema validation',
          timestamp: Date.now(),
        };
      }
      
      return {
        success: true,
        data: parsed.data!,
        timestamp: Date.now(),
      };
    } catch (error) {
      // Fallback to mock on error
      return await this.mockAIResponse(request);
    }
  }

  private async mockAIResponse(request: AIRequest): Promise<AIResponse> {
    // Simulate API delay
    await new Promise(resolve => setTimeout(resolve, 1000 + Math.random() * 2000));

    const prompt = request.prompt.toLowerCase();

    // --- Enhanced parsing for complex operations ---
    // 1. Explicit range/cell references (e.g., "A1:B10", "column B", "row 5")
    const rangeRegex = /([a-z]+\d+:[a-z]+\d+|[a-z]+\d+)/i;
    const columnRegex = /column\s+([a-z]+)/i;
    const rowRegex = /row\s+(\d+)/i;
    const hasRange = rangeRegex.test(prompt);
    const hasColumn = columnRegex.test(prompt);
    const hasRow = rowRegex.test(prompt);

    // 2. Bulk/advanced operations - order matters for precision
    // Most specific patterns first
    if (prompt.includes('remove') && (prompt.includes('duplicates') || prompt.includes('duplicate'))) {
      return this.handleCleanRequest(request);
    } else if (prompt.includes('filter') || prompt.includes('show only') || prompt.includes('show values')) {
      return this.handleFilterRequest(request);
    } else if (prompt.includes('explain') || prompt.includes('what') || prompt.includes('how')) {
      return this.handleExplainRequest(request);
    } else if (prompt.includes('delete') && (hasRow || hasColumn || hasRange)) {
      return this.handleCleanRequest(request);
    } else if (prompt.includes('copy') && hasRange) {
      return this.handleCopyRequest(request);
    } else if (prompt.includes('move') && (hasRange || hasColumn)) {
      return this.handleMoveRequest(request);
    } else if (prompt.includes('sort') && (hasColumn || hasRange)) {
      return this.handleSortRequest(request);
    } else if (prompt.includes('chart') || prompt.includes('graph') || prompt.includes('visualize')) {
      return this.handleChartRequest(request);
    } else if (prompt.includes('table') || prompt.includes('pivot')) {
      return this.handleTableRequest(request);
    } else if (prompt.includes('add') && (prompt.includes('column') || prompt.includes('row'))) {
      return this.handleCreateRequest(request);
    } else if (prompt.includes('format') || prompt.includes('style')) {
      return this.handleFormatRequest(request);
    } else if (prompt.includes('sum') || prompt.includes('average') || prompt.includes('calculate') || prompt.includes('formula')) {
      return this.handleCalculationRequest(request);
    } else if (prompt.includes('add') && (prompt.includes('cell') || prompt.includes('number') || prompt.includes('value'))) {
      return this.handleAddValueRequest(request);
    } else if (prompt.includes('add') || prompt.includes('create') || prompt.includes('insert')) {
      return this.handleCreateRequest(request);
    } else {
      // Fallback: try to infer intent from context
      if (hasRange || hasColumn || hasRow) {
        // If a range/column/row is mentioned but not a clear action, treat as analyze
        return this.handleGenericRequest(request);
      }
      return this.handleGenericRequest(request);
    }
  }

  private handleExplainRequest(_request: AIRequest): AIResponse {
    return {
      success: true,
      data: {
        action: 'explain',
        description: 'I\'ll explain the selected formula or data structure.',
        excelOperations: [],
        suggestions: [
          'Select a cell with a formula to get a detailed explanation',
          'Ask "What does this range contain?" to understand your data',
          'Use "Explain this chart" to understand visualizations'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleAddValueRequest(request: AIRequest): AIResponse {
    const prompt = request.prompt.toLowerCase();
    const operations: ExcelOperation[] = [];
    // Extract cell reference and value from prompt
    const cellMatch = prompt.match(/cell\s+([A-Z]+\d+)/i);
    const numberMatch = prompt.match(/number\s+(\d+)/i) || prompt.match(/(\d+)/);
    if (cellMatch && numberMatch && cellMatch[1] && numberMatch[1]) {
      const cellAddress = cellMatch[1].toUpperCase();
      const value = numberMatch[1];
      operations.push({
        type: 'modify',
        target: cellAddress,
        value: value,
        description: `Add the number ${value} to cell ${cellAddress}`,
      });
    } else {
      operations.push({
        type: 'modify',
        target: 'Selected cell',
        value: '32',
        description: 'Add a number to the selected cell',
      });
    }
    return {
      success: true,
      data: {
        action: 'modify',
        description: 'I\'ll add the value to the specified cell.',
        excelOperations: operations,
        suggestions: [
          'Select a cell before adding values',
          'Use "Add formula to cell A1" for calculations',
          'Try "Format cell A1 as currency" for formatting'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleCreateRequest(request: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    if (request.prompt.includes('column')) {
      operations.push({
        type: 'insert',
        target: 'Next available column',
        description: 'Add a new column',
      });
    } else if (request.prompt.includes('row')) {
      operations.push({
        type: 'insert',
        target: 'Next available row',
        description: 'Add a new row',
      });
    } else if (request.prompt.includes('chart') || request.prompt.includes('graph')) {
      operations.push({
        type: 'chart',
        target: 'Chart object',
        options: { chartType: 'ColumnClustered' },
        description: 'Create a chart based on selected data',
      });
    }
    return {
      success: true,
      data: {
        action: 'create',
        description: 'I\'ll create the requested elements in your spreadsheet.',
        excelOperations: operations,
        suggestions: [
          'Select the data range before creating charts',
          'Specify column headers for better organization',
          'Use "Add a column for [purpose]" for specific needs'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleFormatRequest(request: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    if (request.prompt.includes('currency')) {
      operations.push({
        type: 'format',
        target: 'Selected range',
        value: 'Currency format',
        description: 'Apply currency formatting to selected cells',
      });
    } else if (request.prompt.includes('percentage')) {
      operations.push({
        type: 'format',
        target: 'Selected range',
        value: 'Percentage format',
        description: 'Apply percentage formatting to selected cells',
      });
    } else if (request.prompt.includes('date')) {
      operations.push({
        type: 'format',
        target: 'Selected range',
        value: 'Date format',
        description: 'Apply date formatting to selected cells',
      });
    }
    return {
      success: true,
      data: {
        action: 'format',
        description: 'I\'ll apply the requested formatting to your data.',
        excelOperations: operations,
        suggestions: [
          'Select the cells you want to format first',
          'Use "Format as table" for professional appearance',
          'Try "Auto-fit columns" for better readability'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleCalculationRequest(request: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    if (request.prompt.includes('sum')) {
      operations.push({
        type: 'formula',
        target: 'Next available cell',
        value: '=SUM(selected_range)',
        description: 'Calculate sum of selected range',
      });
    } else if (request.prompt.includes('average')) {
      operations.push({
        type: 'formula',
        target: 'Next available cell',
        value: '=AVERAGE(selected_range)',
        description: 'Calculate average of selected range',
      });
    } else if (request.prompt.includes('count')) {
      operations.push({
        type: 'formula',
        target: 'Next available cell',
        value: '=COUNT(selected_range)',
        description: 'Count items in selected range',
      });
    }
    return {
      success: true,
      data: {
        action: 'analyze',
        description: 'I\'ll perform the requested calculation.',
        excelOperations: operations,
        suggestions: [
          'Select the range before calculating',
          'Use "Sum column B" for totals',
          'Try "Average of A1:A10" for averages'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleCopyRequest(request: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    const prompt = request.prompt.toLowerCase();
    // Example: "Copy A1:B5 to C1"
    const rangeMatch = prompt.match(/copy\s+([a-z]+\d+:[a-z]+\d+)/i);
    const targetMatch = prompt.match(/to\s+([a-z]+\d+)/i);
    if (rangeMatch && targetMatch && rangeMatch[1] && targetMatch[1]) {
      operations.push({
        type: 'copy',
        target: targetMatch[1].toUpperCase(),
        range: rangeMatch[1].toUpperCase(),
        description: `Copy ${rangeMatch[1].toUpperCase()} to ${targetMatch[1].toUpperCase()}`,
      });
    } else {
      operations.push({
        type: 'copy',
        target: 'Next available cell',
        description: 'Copy selected data to next available cell',
      });
    }
    return {
      success: true,
      data: {
        action: 'transform',
        description: 'I\'ll copy the data as requested.',
        excelOperations: operations,
        suggestions: [
          'Specify source and target ranges for copying',
          'Use "Copy A1:B5 to C1" for explicit copy',
          'Try "Copy selected data" for quick copy'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleMoveRequest(request: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    const prompt = request.prompt.toLowerCase();
    // Example: "Move column B to column D"
    const rangeMatch = prompt.match(/move\s+([a-z]+\d+:[a-z]+\d+)/i);
    const targetMatch = prompt.match(/to\s+([a-z]+\d+)/i);
    if (rangeMatch && targetMatch && rangeMatch[1] && targetMatch[1]) {
      operations.push({
        type: 'move',
        target: targetMatch[1].toUpperCase(),
        range: rangeMatch[1].toUpperCase(),
        description: `Move ${rangeMatch[1].toUpperCase()} to ${targetMatch[1].toUpperCase()}`,
      });
    } else {
      operations.push({
        type: 'move',
        target: 'New location',
        description: 'Move selected data to new location',
      });
    }
    return {
      success: true,
      data: {
        action: 'transform',
        description: 'I\'ll move the data as requested.',
        excelOperations: operations,
        suggestions: [
          'Specify source and target for moving',
          'Use "Move column B to column D" for explicit move',
          'Try "Move selected data" for quick move'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleSortRequest(request: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    const prompt = request.prompt.toLowerCase();
    // Extract sort criteria
    const columnMatch = prompt.match(/sort\s+(?:by\s+)?(?:column\s+)?([A-Z]+)/i);
    const orderMatch = prompt.match(/(ascending|descending|a-z|z-a)/i);
    const sortBy = columnMatch && columnMatch[1] ? columnMatch[1] : 'A';
    operations.push({
      type: 'sort',
      target: 'Selected range',
      options: { sortBy },
      description: `Sort data by column ${sortBy} in ${orderMatch ? orderMatch[1] : 'ascending'} order`,
    });
    return {
      success: true,
      data: {
        action: 'transform',
        description: 'I\'ll sort the data as requested.',
        excelOperations: operations,
        suggestions: [
          'Select the data range before sorting',
          'Use "Sort by column B descending" for specific sorting',
          'Try "Sort alphabetically" for text data'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleFilterRequest(request: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    const prompt = request.prompt.toLowerCase();
    // Extract filter criteria
    const criteriaMatch = prompt.match(/show\s+(?:only\s+)?(.+?)(?:\s+in\s+|\s+where\s+|\s+with\s+)/i);
    const columnMatch = prompt.match(/(?:in\s+|where\s+|with\s+)(?:column\s+)?([A-Z]+)/i);
    const criteria = criteriaMatch ? criteriaMatch[1] : 'matching criteria';
    const column = columnMatch ? columnMatch[1] : 'A';
    operations.push({
      type: 'filter',
      target: 'Selected range',
      options: { filterCriteria: { column, value: criteria } },
      description: `Filter data to show ${criteria} in column ${column}`,
    });
    return {
      success: true,
      data: {
        action: 'transform',
        description: 'I\'ll filter the data as requested.',
        excelOperations: operations,
        suggestions: [
          'Select the data range before filtering',
          'Use "Show only values > 100" for numeric filters',
          'Try "Filter by date" for date-based filtering'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleChartRequest(request: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    const prompt = request.prompt.toLowerCase();
    let chartType = 'ColumnClustered';
    if (prompt.includes('line')) chartType = 'Line';
    else if (prompt.includes('pie')) chartType = 'Pie';
    else if (prompt.includes('bar')) chartType = 'BarClustered';
    else if (prompt.includes('scatter')) chartType = 'XYScatter';
    operations.push({
      type: 'chart',
      target: 'Chart object',
      options: { chartType },
      description: `Create a ${chartType} chart from selected data`,
    });
    return {
      success: true,
      data: {
        action: 'create',
        description: 'I\'ll create a chart from your data.',
        excelOperations: operations,
        suggestions: [
          'Select the data range before creating charts',
          'Use "Create line chart" for trend visualization',
          'Try "Make pie chart" for proportion data'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleTableRequest(_: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    operations.push({
      type: 'table',
      target: 'Selected range',
      description: 'Convert selected data to a formatted table',
    });
    return {
      success: true,
      data: {
        action: 'create',
        description: 'I\'ll convert your data to a formatted table.',
        excelOperations: operations,
        suggestions: [
          'Select the data range before creating tables',
          'Use "Create pivot table" for data analysis',
          'Try "Format as table" for better appearance'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleCleanRequest(request: AIRequest): AIResponse {
    const operations: ExcelOperation[] = [];
    const prompt = request.prompt.toLowerCase();
    if (prompt.includes('duplicate')) {
      operations.push({
        type: 'delete',
        target: 'Selected range',
        description: 'Remove duplicate values from selected data',
      });
    } else if (prompt.includes('empty') || prompt.includes('blank')) {
      operations.push({
        type: 'delete',
        target: 'Selected range',
        description: 'Remove empty cells from selected data',
      });
    } else {
      operations.push({
        type: 'delete',
        target: 'Selected range',
        description: 'Clean up selected data',
      });
    }
    return {
      success: true,
      data: {
        action: 'clean',
        description: 'I\'ll clean up your data as requested.',
        excelOperations: operations,
        suggestions: [
          'Select the data range before cleaning',
          'Use "Remove duplicates" for duplicate elimination',
          'Try "Delete empty rows" for data cleanup'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleGenericRequest(_: AIRequest): AIResponse {
    return {
      success: true,
      data: {
        action: 'analyze',
        description: 'I understand your request. Let me help you with that.',
        excelOperations: [],
        suggestions: [
          'Be more specific about what you want to do',
          'Try selecting cells before asking questions',
          'Use phrases like "Add a column for..." or "Format this as..."'
        ],
      },
      timestamp: Date.now(),
    };
  }
}

// Export singleton instance
export const aiService = new AIService(); 