import { ExcelOperationsService } from '../services/ExcelOperationsService';
import { ExcelOperation } from '../services/AIService';

// Mock Office.js
global.Office = {
  onReady: jest.fn(),
  context: {
    document: {}
  }
} as any;

// Mock Excel
global.Excel = {
  run: jest.fn(),
} as any;

describe('ExcelOperationsService', () => {
  let service: ExcelOperationsService;

  beforeEach(() => {
    service = new ExcelOperationsService();
    jest.clearAllMocks();
  });

  describe('executeOperations', () => {
    it('should handle modify operations correctly', async () => {
      const mockRange = {
        load: jest.fn(),
        values: [],
        address: 'A1',
        rowCount: 1,
        columnCount: 1
      };

      const mockWorksheet = {
        getRange: jest.fn().mockReturnValue(mockRange)
      };

      const mockContext = {
        workbook: {
          worksheets: {
            getActiveWorksheet: jest.fn().mockReturnValue(mockWorksheet)
          }
        },
        sync: jest.fn()
      };

      (Excel.run as jest.Mock).mockImplementation(async (callback) => {
        await callback(mockContext);
      });

      const operations: ExcelOperation[] = [
        {
          type: 'modify',
          target: 'A1',
          value: '32',
          description: 'Add the number 32 to cell A1'
        }
      ];

      const results = await service.executeOperations(operations);

      expect(results).toHaveLength(1);
      expect(results[0]?.success).toBe(true);
      expect(results[0]?.message).toContain('Modified cell A1 with value: 32');
    });

    it('should handle invalid cell references gracefully', async () => {
      const operations: ExcelOperation[] = [
        {
          type: 'modify',
          target: 'INVALID_CELL',
          value: '32',
          description: 'Add the number 32 to invalid cell'
        }
      ];

      // Mock Excel.run to throw an error
      (Excel.run as jest.Mock).mockRejectedValue(new Error('Invalid cell reference'));

      const results = await service.executeOperations(operations);

      expect(results).toHaveLength(1);
      expect(results[0]?.success).toBe(false);
      expect(results[0]?.error).toContain('Invalid cell reference');
    });
  });

  describe('getContext', () => {
    it('should return Excel context when Office.js is ready', async () => {
      const mockWorksheet = {
        load: jest.fn(),
        name: 'Sheet1'
      };

      const mockRange = {
        load: jest.fn(),
        address: 'A1:B5',
        rowCount: 5,
        columnCount: 2
      };

      const mockContext = {
        workbook: {
          worksheets: {
            getActiveWorksheet: jest.fn().mockReturnValue(mockWorksheet)
          },
          getSelectedRange: jest.fn().mockReturnValue(mockRange)
        },
        sync: jest.fn()
      };

      (Excel.run as jest.Mock).mockImplementation(async (callback) => {
        await callback(mockContext);
      });

      const context = await service.getContext();

      expect(context.worksheetName).toBe('Sheet1');
      expect(context.selectedRange).toBe('A1:B5');
    });
  });
}); 