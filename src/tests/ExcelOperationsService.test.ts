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
    
    // Mock the isOfficeReady method to return true
    jest.spyOn(service as any, 'isOfficeReady').mockReturnValue(true);
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

    it('should handle formula operations correctly', async () => {
      const mockCell = {
        load: jest.fn(),
        formulas: []
      };

      const mockRange = {
        load: jest.fn(),
        values: [],
        address: 'A1',
        rowCount: 1,
        columnCount: 1,
        formulas: []
      };

      const mockWorksheet = {
        getRange: jest.fn().mockReturnValue(mockRange),
        getCell: jest.fn().mockReturnValue(mockCell)
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

      const operations: ExcelOperation[] = [
        {
          type: 'formula',
          target: 'A1',
          value: '=SUM(B1:B10)',
          description: 'Add sum formula to cell A1'
        }
      ];

      const results = await service.executeOperations(operations);

      expect(results).toHaveLength(1);
      expect(results[0]?.success).toBe(true);
      expect(results[0]?.message).toContain('Formula inserted: =SUM(B1:B10)');
    });

    it('should handle format operations correctly', async () => {
      const mockRange = {
        load: jest.fn(),
        values: [],
        address: 'A1',
        rowCount: 1,
        columnCount: 1,
        format: {
          fill: {},
          font: {}
        }
      };

      const mockWorksheet = {
        getRange: jest.fn().mockReturnValue(mockRange)
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

      const operations: ExcelOperation[] = [
        {
          type: 'format',
          target: 'A1',
          value: 'currency',
          description: 'Format cell A1 as currency'
        }
      ];

      const results = await service.executeOperations(operations);

      expect(results).toHaveLength(1);
      expect(results[0]?.success).toBe(true);
      expect(results[0]?.message).toContain('Applied currency formatting');
    });

    it('should handle copy operations correctly', async () => {
      const mockSourceRange = {
        load: jest.fn(),
        values: [['Data']],
        address: 'A1',
        rowCount: 1,
        columnCount: 1
      };

      const mockTargetRange = {
        load: jest.fn(),
        values: [],
        address: 'B1',
        rowCount: 1,
        columnCount: 1
      };

      const mockWorksheet = {
        getRange: jest.fn()
          .mockReturnValueOnce(mockSourceRange)
          .mockReturnValueOnce(mockTargetRange)
      };

      const mockContext = {
        workbook: {
          worksheets: {
            getActiveWorksheet: jest.fn().mockReturnValue(mockWorksheet)
          },
          getSelectedRange: jest.fn().mockReturnValue(mockSourceRange)
        },
        sync: jest.fn()
      };

      (Excel.run as jest.Mock).mockImplementation(async (callback) => {
        await callback(mockContext);
      });

      const operations: ExcelOperation[] = [
        {
          type: 'copy',
          target: 'B1',
          range: 'A1',
          description: 'Copy data from A1 to B1'
        }
      ];

      const results = await service.executeOperations(operations);

      expect(results).toHaveLength(1);
      expect(results[0]?.success).toBe(true);
      expect(results[0]?.message).toContain('Copied data to B1');
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

    it('should handle multiple operations in sequence', async () => {
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
          },
          getSelectedRange: jest.fn().mockReturnValue(mockRange)
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
          value: '100',
          description: 'Set A1 to 100'
        },
        {
          type: 'modify',
          target: 'B1',
          value: '200',
          description: 'Set B1 to 200'
        }
      ];

      const results = await service.executeOperations(operations);

      expect(results).toHaveLength(2);
      expect(results[0]?.success).toBe(true);
      expect(results[1]?.success).toBe(true);
      expect(results[0]?.message).toContain('A1');
      expect(results[1]?.message).toContain('B1');
    });
  });

  describe('getContext', () => {
    it('should return Excel context when Office.js is ready', async () => {
      const mockCell = {
        address: 'A1'
      };

      const mockWorksheet = {
        load: jest.fn(),
        name: 'Sheet1'
      };

      const mockRange = {
        load: jest.fn(),
        address: 'A1:B5',
        rowCount: 5,
        columnCount: 2,
        getCell: jest.fn().mockReturnValue(mockCell)
      };

      const mockContext = {
        workbook: {
          worksheets: {
            getActiveWorksheet: jest.fn().mockReturnValue(mockWorksheet)
          },
          getSelectedRange: jest.fn().mockReturnValue(mockRange)
        },
        sync: jest.fn().mockImplementation(async () => {
          // Simulate the sync operation - ensure worksheet name is available
          return Promise.resolve();
        })
      };

      (Excel.run as jest.Mock).mockImplementation(async (callback) => {
        // Simulate the Excel.run behavior by calling the callback and returning its result
        return await callback(mockContext);
      });

      const context = await service.getContext();

      expect(context).toBeDefined();
      expect(context.worksheetName).toBe('Sheet1');
      expect(context.selectedRange).toBe('A1:B5');
    });
  });
}); 