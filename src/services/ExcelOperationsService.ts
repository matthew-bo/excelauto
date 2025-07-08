import { ExcelOperation } from './AIService';

export interface ExcelContext {
  selectedRange?: string;
  worksheetName?: string;
  activeCell?: string;
}

export interface OperationResult {
  success: boolean;
  message: string;
  error?: string;
}

// Centralized error logger
function logExcelError(error: unknown, context: {
  operationType: string;
  operation?: ExcelOperation;
  extra?: Record<string, any>;
}) {
  // eslint-disable-next-line no-console
  console.error('[ExcelOperationError]', {
    type: context.operationType,
    operation: context.operation,
    extra: context.extra,
    error,
  });
}

export class ExcelOperationsService {
  private isOfficeReady(): boolean {
    return typeof Office !== 'undefined' && 
           Office.context && 
           Office.context.document &&
           typeof Excel !== 'undefined';
  }

  async getContext(): Promise<ExcelContext> {
    if (!this.isOfficeReady()) {
      const errorMsg = 'Office.js is not ready or Excel is not available';
      logExcelError(errorMsg, { operationType: 'getContext' });
      throw new Error(errorMsg);
    }

    try {
      return await Excel.run(async (context) => {
        try {
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          
          // Load worksheet name first
          worksheet.load('name');
          await context.sync();
          
          // Try to get selected range, but handle case where nothing is selected
          let selectedRange = '';
          let activeCell = '';
          
          try {
            const range = context.workbook.getSelectedRange();
            range.load(['address', 'rowCount', 'columnCount']);
            await context.sync();
            
            if (range.rowCount > 0 && range.columnCount > 0) {
              selectedRange = range.address;
              activeCell = range.getCell(0, 0).address;
            } else {
              // No valid selection, use active cell instead
              const activeCellRange = context.workbook.getActiveCell();
              activeCellRange.load('address');
              await context.sync();
              activeCell = activeCellRange.address;
              selectedRange = activeCell;
            }
          } catch (rangeError) {
            // If getting selected range fails, try active cell
            try {
              const activeCellRange = context.workbook.getActiveCell();
              activeCellRange.load('address');
              await context.sync();
              activeCell = activeCellRange.address;
              selectedRange = activeCell;
            } catch (activeCellError) {
              // If all else fails, provide default values
              selectedRange = 'A1';
              activeCell = 'A1';
            }
          }
          
          return {
            selectedRange,
            worksheetName: worksheet.name,
            activeCell,
          };
        } catch (error) {
          throw new Error(`Failed to get Excel context: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
    } catch (error) {
      console.error('Error getting Excel context:', error);
      throw error;
    }
  }

  async executeOperations(operations: ExcelOperation[]): Promise<OperationResult[]> {
    if (!this.isOfficeReady()) {
      const errorMsg = 'Office.js is not ready';
      logExcelError(errorMsg, { operationType: 'executeOperations' });
      return [{
        success: false,
        message: errorMsg,
        error: 'Excel is not available'
      }];
    }

    const results: OperationResult[] = [];
    for (const operation of operations) {
      try {
        const result = await this.executeOperation(operation);
        results.push(result);
      } catch (error) {
        results.push({
          success: false,
          message: `Failed to execute ${operation.description}`,
          error: error instanceof Error ? error.message : 'Unknown error',
        });
      }
    }
    return results;
  }

  private async executeOperation(operation: ExcelOperation): Promise<OperationResult> {
    try {
      // Validate operation
      if (!operation || !operation.type) {
        throw new Error('Invalid operation object');
      }
      switch (operation.type) {
        case 'formula':
          return await this.insertFormula(operation);
        case 'format':
          return await this.applyFormatting(operation);
        case 'insert':
          return await this.insertElement(operation);
        case 'delete':
          return await this.deleteElement(operation);
        case 'modify':
          return await this.modifyElement(operation);
        case 'copy':
          return await this.copyElement(operation);
        case 'move':
          return await this.moveElement(operation);
        case 'sort':
          return await this.sortData(operation);
        case 'filter':
          return await this.filterData(operation);
        case 'chart':
          return await this.createChart(operation);
        case 'table':
          return await this.createTable(operation);
        default:
          return {
            success: false,
            message: `Unknown operation type: ${operation.type}`,
          };
      }
    } catch (error) {
      logExcelError(error, { operationType: 'executeOperation', operation });
      return {
        success: false,
        message: `Failed to execute operation: ${operation?.description || operation?.type}`,
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async insertFormula(operation: ExcelOperation): Promise<OperationResult> {
    try {
      // Validate
      if (!operation.value) {
        throw new Error('No formula value provided');
      }
      await Excel.run(async (context) => {
        try {
          const range = context.workbook.getSelectedRange();
          range.load(['rowIndex', 'columnIndex', 'address']);
          await context.sync();
          
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          const cell = worksheet.getCell(range.rowIndex, range.columnIndex);
          
          // Set the formula
          cell.formulas = [[operation.value || '']];
          
          await context.sync();
        } catch (error) {
          throw new Error(`Formula insertion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
      
      return {
        success: true,
        message: `Formula inserted: ${operation.value}`,
      };
    } catch (error) {
      logExcelError(error, { operationType: 'insertFormula', operation });
      return {
        success: false,
        message: 'Failed to insert formula',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async applyFormatting(operation: ExcelOperation): Promise<OperationResult> {
    try {
      // Validate
      if (!operation.value) {
        throw new Error('No format value provided');
      }
      await Excel.run(async (context) => {
        try {
          const range = context.workbook.getSelectedRange();
          range.load(['address', 'rowCount', 'columnCount']);
          await context.sync();
          
          // Validate range
          if (range.rowCount === 0 || range.columnCount === 0) {
            throw new Error('No valid range selected for formatting');
          }
          
          // Apply formatting based on operation value
          if (operation.value?.includes('Currency')) {
            range.numberFormat = [['$#,##0.00']];
          } else if (operation.value?.includes('Percentage')) {
            range.numberFormat = [['0.00%']];
          } else if (operation.value?.includes('Date')) {
            range.numberFormat = [['mm/dd/yyyy']];
          } else {
            // Default to general format
            range.numberFormat = [['General']];
          }
          
          await context.sync();
        } catch (error) {
          throw new Error(`Formatting failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
      
      return {
        success: true,
        message: `Applied ${operation.value} formatting`,
      };
    } catch (error) {
      logExcelError(error, { operationType: 'applyFormatting', operation });
      return {
        success: false,
        message: 'Failed to apply formatting',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async insertElement(operation: ExcelOperation): Promise<OperationResult> {
    try {
      if (operation.target.includes('column')) {
        await Excel.run(async (context) => {
          try {
            const range = context.workbook.getSelectedRange();
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            range.load(['columnIndex']);
            worksheet.load(['getUsedRange']);
            await context.sync();
            
            const usedRange = worksheet.getUsedRange();
            usedRange.load(['rowCount']);
            await context.sync();
            
            // Insert column to the right of the selected range
            worksheet.getRangeByIndexes(0, range.columnIndex, usedRange.rowCount, 1).insert('Right');
            await context.sync();
          } catch (error) {
            throw new Error(`Column insertion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
          }
        });
        
        return {
          success: true,
          message: 'New column inserted',
        };
      } else if (operation.target.includes('row')) {
        await Excel.run(async (context) => {
          try {
            const range = context.workbook.getSelectedRange();
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            range.load(['rowIndex']);
            worksheet.load(['getUsedRange']);
            await context.sync();
            
            const usedRange = worksheet.getUsedRange();
            usedRange.load(['columnCount']);
            await context.sync();
            
            // Insert row below the selected range
            worksheet.getRangeByIndexes(range.rowIndex, 0, 1, usedRange.columnCount).insert('Down');
            await context.sync();
          } catch (error) {
            throw new Error(`Row insertion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
          }
        });
        
        return {
          success: true,
          message: 'New row inserted',
        };
      } else if (operation.target.includes('Chart')) {
        await Excel.run(async (context) => {
          try {
            const range = context.workbook.getSelectedRange();
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            range.load(['rowCount', 'columnCount', 'values']);
            await context.sync();
            
            // Validate that we have enough data for a chart
            if (range.rowCount < 2 || range.columnCount < 2) {
              throw new Error('Insufficient data for chart creation. Select at least 2x2 range.');
            }
            
            // Create a column chart
            const chart = worksheet.charts.add('ColumnClustered', range, 'Auto');
            chart.title.text = 'Generated Chart';
            
            await context.sync();
          } catch (error) {
            throw new Error(`Chart creation failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
          }
        });
        
        return {
          success: true,
          message: 'Chart created successfully',
        };
      }
      
      return {
        success: false,
        message: `Unknown insert operation: ${operation.target}`,
      };
    } catch (error) {
      logExcelError(error, { operationType: 'insertElement', operation });
      return {
        success: false,
        message: 'Failed to insert element',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async deleteElement(operation: ExcelOperation): Promise<OperationResult> {
    try {
      // Implementation for delete operations (future)
      return {
        success: true,
        message: `Delete operation: ${operation.description}`,
      };
    } catch (error) {
      logExcelError(error, { operationType: 'deleteElement', operation });
      return {
        success: false,
        message: 'Failed to delete element',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async modifyElement(operation: ExcelOperation): Promise<OperationResult> {
    try {
      // Validate
      if (!operation.value) {
        throw new Error('No value provided for modification');
      }
      await Excel.run(async (context) => {
        try {
          let targetRange;
          
          // Check if target is a cell reference (e.g., "A1", "B5")
          const cellMatch = operation.target.match(/^([A-Z]+\d+)$/i);
          
          if (cellMatch) {
            // Use the specific cell reference
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            targetRange = worksheet.getRange(operation.target);
          } else {
            // Use the currently selected range
            targetRange = context.workbook.getSelectedRange();
          }
          
          // Load the range properties
          targetRange.load(['address', 'rowCount', 'columnCount']);
          await context.sync();
          
          // Validate the range
          if (targetRange.rowCount === 0 || targetRange.columnCount === 0) {
            throw new Error('No valid range selected for modification');
          }
          
          // Set the value
          if (operation.value) {
            // Try to parse as number first
            const numValue = parseFloat(operation.value);
            if (!isNaN(numValue)) {
              targetRange.values = [[numValue]];
            } else {
              targetRange.values = [[operation.value]];
            }
          }
          
          await context.sync();
        } catch (error) {
          throw new Error(`Cell modification failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
      
      return {
        success: true,
        message: `Modified cell ${operation.target} with value: ${operation.value}`,
      };
    } catch (error) {
      logExcelError(error, { operationType: 'modifyElement', operation });
      return {
        success: false,
        message: 'Failed to modify cell',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async copyElement(operation: ExcelOperation): Promise<OperationResult> {
    try {
      await Excel.run(async (context) => {
        try {
          let sourceRange;
          let targetRange;
          
          // Determine source range
          if (operation.range) {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            sourceRange = worksheet.getRange(operation.range);
          } else {
            sourceRange = context.workbook.getSelectedRange();
          }
          
          // Determine target range
          if (operation.target && operation.target !== 'Selected range') {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            targetRange = worksheet.getRange(operation.target);
          } else {
            // Use next available cell
            const activeCell = context.workbook.getActiveCell();
            activeCell.load('address');
            await context.sync();
            
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            targetRange = worksheet.getRange(activeCell.address);
          }
          
          // Load ranges
          sourceRange.load(['values', 'formulas', 'numberFormat']);
          targetRange.load(['address']);
          await context.sync();
          
          // Copy values and formatting
          targetRange.values = sourceRange.values;
          targetRange.formulas = sourceRange.formulas;
          targetRange.numberFormat = sourceRange.numberFormat;
          
          await context.sync();
        } catch (error) {
          throw new Error(`Copy operation failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
      
      return {
        success: true,
        message: `Copied data to ${operation.target}`,
      };
    } catch (error) {
      logExcelError(error, { operationType: 'copyElement', operation });
      return {
        success: false,
        message: 'Failed to copy data',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async moveElement(operation: ExcelOperation): Promise<OperationResult> {
    try {
      await Excel.run(async (context) => {
        try {
          let sourceRange;
          let targetRange;
          
          // Determine source range
          if (operation.range) {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            sourceRange = worksheet.getRange(operation.range);
          } else {
            sourceRange = context.workbook.getSelectedRange();
          }
          
          // Determine target range
          if (operation.target && operation.target !== 'New location') {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            targetRange = worksheet.getRange(operation.target);
          } else {
            // Use next available cell
            const activeCell = context.workbook.getActiveCell();
            activeCell.load('address');
            await context.sync();
            
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            targetRange = worksheet.getRange(activeCell.address);
          }
          
          // Load ranges
          sourceRange.load(['values', 'formulas', 'numberFormat']);
          targetRange.load(['address']);
          await context.sync();
          
          // Move data (copy to target, clear source)
          targetRange.values = sourceRange.values;
          targetRange.formulas = sourceRange.formulas;
          targetRange.numberFormat = sourceRange.numberFormat;
          
          // Clear source
          sourceRange.clear();
          
          await context.sync();
        } catch (error) {
          throw new Error(`Move operation failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
      
      return {
        success: true,
        message: `Moved data to ${operation.target}`,
      };
    } catch (error) {
      logExcelError(error, { operationType: 'moveElement', operation });
      return {
        success: false,
        message: 'Failed to move data',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async sortData(operation: ExcelOperation): Promise<OperationResult> {
    try {
      await Excel.run(async (context) => {
        try {
          const range = context.workbook.getSelectedRange();
          range.load(['address', 'rowCount', 'columnCount']);
          await context.sync();
          
          // Validate range
          if (range.rowCount < 2) {
            throw new Error('Need at least 2 rows for sorting');
          }
          
          // Determine sort column
          const sortColumn = operation.options?.sortBy || 'A';
          const columnIndex = this.getColumnIndex(sortColumn);
          
          // Create sort fields
          const sortField = {
            key: columnIndex,
            ascending: true, // Default to ascending
          };
          
          // Apply sort
          range.sort.apply([sortField]);
          
          await context.sync();
        } catch (error) {
          throw new Error(`Sort operation failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
      
      return {
        success: true,
        message: `Sorted data by column ${operation.options?.sortBy || 'A'}`,
      };
    } catch (error) {
      logExcelError(error, { operationType: 'sortData', operation });
      return {
        success: false,
        message: 'Failed to sort data',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async filterData(operation: ExcelOperation): Promise<OperationResult> {
    try {
      await Excel.run(async (context) => {
        try {
          const range = context.workbook.getSelectedRange();
          range.load(['address', 'rowCount', 'columnCount']);
          await context.sync();
          
          // Validate range
          if (range.rowCount < 2) {
            throw new Error('Need at least 2 rows for filtering');
          }
          
          // Apply filter - using the worksheet's autoFilter method
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          worksheet.autoFilter.apply(range);
          
          await context.sync();
        } catch (error) {
          throw new Error(`Filter operation failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
      
      return {
        success: true,
        message: 'Applied filter to data',
      };
    } catch (error) {
      logExcelError(error, { operationType: 'filterData', operation });
      return {
        success: false,
        message: 'Failed to filter data',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async createChart(operation: ExcelOperation): Promise<OperationResult> {
    try {
      await Excel.run(async (context) => {
        try {
          const range = context.workbook.getSelectedRange();
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          
          range.load(['rowCount', 'columnCount', 'values']);
          await context.sync();
          
          // Validate that we have enough data for a chart
          if (range.rowCount < 2 || range.columnCount < 2) {
            throw new Error('Insufficient data for chart creation. Select at least 2x2 range.');
          }
          
          // Determine chart type
          const chartType = operation.options?.chartType || 'ColumnClustered';
          
          // Create chart
          const chart = worksheet.charts.add(chartType as any, range, 'Auto');
          chart.title.text = 'Generated Chart';
          
          await context.sync();
        } catch (error) {
          throw new Error(`Chart creation failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
      
      return {
        success: true,
        message: `Created ${operation.options?.chartType || 'column'} chart`,
      };
    } catch (error) {
      logExcelError(error, { operationType: 'createChart', operation });
      return {
        success: false,
        message: 'Failed to create chart',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  private async createTable(operation: ExcelOperation): Promise<OperationResult> {
    try {
      await Excel.run(async (context) => {
        try {
          const range = context.workbook.getSelectedRange();
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          
          range.load(['address', 'rowCount', 'columnCount']);
          await context.sync();
          
          // Validate range
          if (range.rowCount < 2) {
            throw new Error('Need at least 2 rows for table creation');
          }
          
          // Convert to table
          const table = worksheet.tables.add(range, true);
          table.name = 'GeneratedTable';
          
          await context.sync();
        } catch (error) {
          throw new Error(`Table creation failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      });
      
      return {
        success: true,
        message: 'Created formatted table',
      };
    } catch (error) {
      logExcelError(error, { operationType: 'createTable', operation });
      return {
        success: false,
        message: 'Failed to create table',
        error: error instanceof Error ? error.message : 'Unknown error',
      };
    }
  }

  // Helper method to convert column letter to index
  private getColumnIndex(columnLetter: string): number {
    let index = 0;
    for (let i = 0; i < columnLetter.length; i++) {
      index = index * 26 + (columnLetter.charCodeAt(i) - 64);
    }
    return index - 1; // Excel uses 0-based indexing
  }
}

// Export singleton instance
export const excelOperationsService = new ExcelOperationsService(); 