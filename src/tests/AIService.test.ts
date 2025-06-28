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