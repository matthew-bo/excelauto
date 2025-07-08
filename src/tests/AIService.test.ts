import { AIService, AIRequest } from '../services/AIService';

describe('AIService', () => {
  let aiService: AIService;

  beforeEach(() => {
    aiService = new AIService();
  });

  describe('processPrompt', () => {
    it('should handle explain requests', async () => {
      const request: AIRequest = {
        prompt: 'Explain this formula',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('explain');
      expect(response.data?.description).toContain('explain');
      expect(response.data?.suggestions).toBeDefined();
    });

    it('should handle create requests', async () => {
      const request: AIRequest = {
        prompt: 'Add a new column',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('create');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('insert');
    });

    it('should handle format requests', async () => {
      const request: AIRequest = {
        prompt: 'Format as currency',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('format');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('format');
    });

    it('should handle calculation requests', async () => {
      const request: AIRequest = {
        prompt: 'Calculate the sum',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('analyze');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('formula');
    });

    it('should handle generic requests', async () => {
      const request: AIRequest = {
        prompt: 'Hello world',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('analyze');
      expect(response.data?.suggestions).toBeDefined();
    });

    it('should include context in request', async () => {
      const request: AIRequest = {
        prompt: 'Format this range',
        context: {
          selectedRange: 'A1:B10',
          worksheetName: 'Sheet1',
        },
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('format');
    });

    it('should handle copy requests', async () => {
      const request: AIRequest = {
        prompt: 'Copy A1:B5 to C1',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('transform');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('copy');
    });

    it('should handle move requests', async () => {
      const request: AIRequest = {
        prompt: 'Move column B to column D',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('transform');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('move');
    });

    it('should handle sort requests', async () => {
      const request: AIRequest = {
        prompt: 'Sort by column A descending',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('transform');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('sort');
    });

    it('should handle filter requests', async () => {
      const request: AIRequest = {
        prompt: 'Filter to show only values > 100',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('transform');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('filter');
    });

    it('should handle chart requests', async () => {
      const request: AIRequest = {
        prompt: 'Create a line chart from this data',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('create');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('chart');
    });

    it('should handle table requests', async () => {
      const request: AIRequest = {
        prompt: 'Convert this to a table',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('create');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('table');
    });

    it('should handle clean requests', async () => {
      const request: AIRequest = {
        prompt: 'Remove duplicate values',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('clean');
      expect(response.data?.excelOperations).toHaveLength(1);
      expect(response.data?.excelOperations?.[0]?.type).toBe('delete');
    });

    it('should handle multi-cell range copy', async () => {
      const request: AIRequest = {
        prompt: 'Copy A1:B10 to C1',
      };
      const response = await aiService.processPrompt(request);
      expect(response.success).toBe(true);
      expect(response.data?.excelOperations?.[0]?.type).toBe('copy');
    });

    it('should handle ambiguous prompt with range', async () => {
      const request: AIRequest = {
        prompt: 'A1:B10',
      };
      const response = await aiService.processPrompt(request);
      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('analyze');
    });

    it('should handle delete row operation', async () => {
      const request: AIRequest = {
        prompt: 'Delete row 5',
      };
      const response = await aiService.processPrompt(request);
      expect(response.success).toBe(true);
      expect(response.data?.excelOperations?.[0]?.type).toBe('delete');
    });

    it('should handle remove duplicates in range', async () => {
      const request: AIRequest = {
        prompt: 'Remove duplicates in A1:B10',
      };
      const response = await aiService.processPrompt(request);
      expect(response.success).toBe(true);
      expect(response.data?.excelOperations?.[0]?.type).toBe('delete');
    });

    it('should handle filter with show only', async () => {
      const request: AIRequest = {
        prompt: 'Show only values greater than 100 in column B',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(true);
      expect(response.data?.excelOperations?.[0]?.type).toBe('filter');
    });

    it('should handle ambiguous prompt with column', async () => {
      const request: AIRequest = {
        prompt: 'column C',
      };
      const response = await aiService.processPrompt(request);
      expect(response.success).toBe(true);
      expect(response.data?.action).toBe('analyze');
    });

    it('should validate operation types are discriminated unions', async () => {
      const request: AIRequest = {
        prompt: 'Set A1 to 100',
      };
      const response = await aiService.processPrompt(request);
      
      expect(response.success).toBe(true);
      expect(response.data?.excelOperations).toBeDefined();
      
      if (response.data?.excelOperations && response.data.excelOperations.length > 0) {
        const operation = response.data.excelOperations[0]!; // Assert non-null
        
        // Test type guards for discriminated unions
        if (operation.type === 'modify') {
          expect(operation.target).toBeDefined();
          expect(operation.value).toBeDefined();
          expect(operation.description).toBeDefined();
        } else if (operation.type === 'insert') {
          expect(operation.target).toBeDefined();
          expect(operation.description).toBeDefined();
        } else if (operation.type === 'delete') {
          expect(operation.target).toBeDefined();
          expect(operation.description).toBeDefined();
        } else if (operation.type === 'copy') {
          expect(operation.target).toBeDefined();
          expect(operation.description).toBeDefined();
          // range is optional
        } else if (operation.type === 'move') {
          expect(operation.target).toBeDefined();
          expect(operation.description).toBeDefined();
          // range is optional
        } else if (operation.type === 'format') {
          expect(operation.target).toBeDefined();
          expect(operation.value).toBeDefined();
          expect(operation.description).toBeDefined();
        } else if (operation.type === 'formula') {
          expect(operation.target).toBeDefined();
          expect(operation.value).toBeDefined();
          expect(operation.description).toBeDefined();
        } else if (operation.type === 'sort') {
          expect(operation.target).toBeDefined();
          expect(operation.options).toBeDefined();
          expect(operation.options.sortBy).toBeDefined();
          expect(operation.description).toBeDefined();
        } else if (operation.type === 'filter') {
          expect(operation.target).toBeDefined();
          expect(operation.options).toBeDefined();
          expect(operation.options.filterCriteria).toBeDefined();
          expect(operation.description).toBeDefined();
        } else if (operation.type === 'chart') {
          expect(operation.target).toBeDefined();
          expect(operation.options).toBeDefined();
          expect(operation.options.chartType).toBeDefined();
          expect(operation.description).toBeDefined();
        } else if (operation.type === 'table') {
          expect(operation.target).toBeDefined();
          expect(operation.description).toBeDefined();
        }
      }
    });
  });

  describe('error handling', () => {
    it('should handle API errors gracefully', async () => {
      // Mock a failed request by overriding the method
      const originalMockResponse = (aiService as any).mockAIResponse;
      (aiService as any).mockAIResponse = () => {
        throw new Error('API Error');
      };

      const request: AIRequest = {
        prompt: 'Test prompt',
      };

      const response = await aiService.processPrompt(request);

      expect(response.success).toBe(false);
      expect(response.error).toBe('API Error');

      // Restore original method
      (aiService as any).mockAIResponse = originalMockResponse;
    });
  });
}); 