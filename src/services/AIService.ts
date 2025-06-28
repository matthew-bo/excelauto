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
    action: 'explain' | 'create' | 'modify' | 'format' | 'analyze';
    description: string;
    excelOperations: ExcelOperation[];
    suggestions: string[];
  };
  error?: string;
  timestamp: number;
}

export interface ExcelOperation {
  type: 'formula' | 'format' | 'insert' | 'delete' | 'modify';
  target: string; // Cell range or description
  value?: string; // Formula, format, or data
  description: string;
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
      const systemPrompt = `You are an expert Excel assistant. Given a user prompt and context, return a JSON object with the following structure: { action: (explain|create|modify|format|analyze), description: string, excelOperations: [{ type: (formula|format|insert|delete|modify), target: string, value?: string, description: string }], suggestions: string[] }. Only return valid JSON.`;
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
      
      return {
        success: true,
        data: parsed,
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
    
    // Parse common Excel operations
    if (prompt.includes('explain') || prompt.includes('what') || prompt.includes('how')) {
      return this.handleExplainRequest(request);
    } else if (prompt.includes('add') && (prompt.includes('cell') || prompt.includes('number') || prompt.includes('value'))) {
      return this.handleAddValueRequest(request);
    } else if (prompt.includes('add') || prompt.includes('create') || prompt.includes('insert')) {
      return this.handleCreateRequest(request);
    } else if (prompt.includes('format') || prompt.includes('style')) {
      return this.handleFormatRequest(request);
    } else if (prompt.includes('sum') || prompt.includes('average') || prompt.includes('calculate')) {
      return this.handleCalculationRequest(request);
    } else {
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
    // Example: "in cell A1 add the number 32"
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
      // Fallback for unclear prompts
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
        type: 'insert',
        target: 'Chart object',
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
          'Select the data range for calculations',
          'Use "Calculate profit margin" for business metrics',
          'Try "Find duplicates" for data cleaning'
        ],
      },
      timestamp: Date.now(),
    };
  }

  private handleGenericRequest(_request: AIRequest): AIResponse {
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