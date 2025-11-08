/**
 * CashFlow Automator - Spreadsheet Integration System
 * Handles Google Sheets synchronization, data updates, and validation
 * @version 2.1.0
 */

/* ==================== SPREADSHEET SYNC FUNCTIONS ==================== */

/**
 * Updates Google Sheets with extracted financial data
 * @param {Array} rows - Array of extracted data objects
 */
function updateSpreadsheet(rows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    Logger.log('Sheet not found: ' + CONFIG.SHEET_NAME);
    return;
  }
  
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const columnIndex = {};
  headers.forEach((header, index) => {
    columnIndex[header] = index;
  });
  
  const rowMap = new Map();
  dataRows.forEach((row, rowIndex) => {
    const key = normalizeDate(row[columnIndex['Date']]) + '|' + row[columnIndex['Shift']] + '|' + row[columnIndex['Branch']];
    rowMap.set(key, rowIndex + 2);
  });
  
  const updates = [];
  
  rows.forEach(rowData => {
    const key = normalizeDate(rowData.closureDate) + '|' + rowData.shift + '|' + rowData.branch;
    const sheetRow = rowMap.get(key);
    
    if (!sheetRow) {
      Logger.log('Row not found for: ' + key);
      return;
    }
    
    const fieldMappings = [
      ['openingCash', 'Opening Cash'],
      ['cashSales', 'Cash Sales'],
      ['totalSales', 'Total Sales'],
      ['cardSales', 'Card Payments'],
      ['digitalPayments', 'Digital Payments'],
      ['closingCash', 'Closing Cash'],
      ['cashWithdrawal', 'Cash Withdrawal']
    ];
    
    fieldMappings.forEach(([dataField, sheetColumn]) => {
      if (rowData[dataField] && columnIndex[sheetColumn] !== undefined) {
        const colIndex = columnIndex[sheetColumn];
        const existingValue = dataRows[sheetRow - 2][colIndex];
        
        if (!existingValue || existingValue === 0 || existingValue === '') {
          updates.push({
            row: sheetRow,
            column: colIndex + 1,
            value: convertArgentineNumber(rowData[dataField])
          });
        }
      }
    });
  });
  
  // Apply updates to spreadsheet
  if (updates.length > 0) {
    applySpreadsheetUpdates(sheet, updates);
    
    // Highlight updated rows
    const updatedRows = [...new Set(updates.map(u => u.row))];
    highlightUpdatedRows(sheet, updatedRows, headers.length);
    
    Logger.log('Applied ' + updates.length + ' updates to spreadsheet');
  } else {
    Logger.log('No updates needed - all data already present');
  }
}

/**
 * Applies updates to the spreadsheet in batch
 * @param {Sheet} sheet - Google Sheet
 * @param {Array} updates - Array of update objects
 */
function applySpreadsheetUpdates(sheet, updates) {
  // Group updates by row for efficiency
  const updatesByRow = new Map();
  
  updates.forEach(({row, column, value}) => {
    if (!updatesByRow.has(row)) {
      updatesByRow.set(row, []);
    }
    updatesByRow.get(row).push({column, value});
  });
  
  // Apply updates row by row
  for (const [row, rowUpdates] of updatesByRow) {
    rowUpdates.forEach(({column, value}) => {
      sheet.getRange(row, column).setValue(value);
    });
  }
}

/**
 * Highlights updated rows in the spreadsheet
 * @param {Sheet} sheet - Google Sheet
 * @param {Array} updatedRows - Array of modified row numbers
 * @param {number} numColumns - Number of columns to highlight
 */
function highlightUpdatedRows(sheet, updatedRows, numColumns) {
  updatedRows.forEach(row => {
    sheet.getRange(row, 1, 1, numColumns).setBackground('#E8F5E8');
  });
  Logger.log('Highlighted ' + updatedRows.length + ' updated rows');
}

/**
 * Finds or creates row for financial data
 * @param {Object} data - Financial data object
 * @returns {number} Row number where data should be placed
 */
function findOrCreateRow(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return -1;
  
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const columnIndex = {};
  headers.forEach((header, index) => {
    columnIndex[header] = index;
  });
  
  const normalizedDate = normalizeDate(data.closureDate);
  const searchKey = normalizedDate + '|' + data.shift + '|' + data.branch;
  
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const rowKey = normalizeDate(row[columnIndex['Date']]) + '|' + row[columnIndex['Shift']] + '|' + row[columnIndex['Branch']];
    
    if (rowKey === searchKey) {
      return i + 2; // +2 because data starts at row 2 (headers + 1-indexing)
    }
  }
  
  // Row not found - would need to create new row
  Logger.log('Row not found for: ' + searchKey);
  return -1;
}

/**
 * Validates financial data before updating spreadsheet
 * @param {Object} data - Financial data to validate
 * @returns {Object} Validation result {isValid: boolean, errors: Array}
 */
function validateFinancialData(data) {
  const errors = [];
  
  // Required fields validation
  if (!data.closureDate) {
    errors.push('Missing closure date');
  }
  
  if (!data.branch) {
    errors.push('Missing branch information');
  }
  
  if (!data.shift) {
    errors.push('Missing shift information');
  }
  
  // Financial data validation
  if (data.totalSales) {
    const sales = convertArgentineNumber(data.totalSales);
    if (sales < 0) {
      errors.push('Total sales cannot be negative');
    }
  }
  
  if (data.openingCash && data.closingCash) {
    const opening = convertArgentineNumber(data.openingCash);
    const closing = convertArgentineNumber(data.closingCash);
    
    if (opening < 0 || closing < 0) {
      errors.push('Cash amounts cannot be negative');
    }
  }
  
  // Data consistency checks
  if (data.cashSales && data.cardSales && data.totalSales) {
    const cash = convertArgentineNumber(data.cashSales) || 0;
    const cards = convertArgentineNumber(data.cardSales) || 0;
    const digital = convertArgentineNumber(data.digitalPayments) || 0;
    const totalSales = convertArgentineNumber(data.totalSales) || 0;
    
    const paymentMethodsSum = cash + cards + digital;
    const difference = Math.abs(paymentMethodsSum - totalSales);
    
    // Allow small rounding differences
    if (difference > 1) {
      errors.push('Payment methods sum (' + paymentMethodsSum + ') doesn\'t match total sales (' + totalSales + ')');
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * Gets spreadsheet statistics and metrics
 * @returns {Object} Spreadsheet statistics
 */
function getSpreadsheetStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    return { error: 'Sheet not found' };
  }
  
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const columnIndex = {};
  headers.forEach((header, index) => {
    columnIndex[header] = index;
  });
  
  const stats = {
    totalRows: dataRows.length,
    totalBranches: new Set(dataRows.map(row => row[columnIndex['Branch']])).size,
    dateRange: {
      min: null,
      max: null
    },
    completedRows: 0,
    incompleteRows: 0
  };
  
  // Calculate date range and completion stats
  const dates = [];
  dataRows.forEach(row => {
    const date = row[columnIndex['Date']];
    if (date) {
      dates.push(new Date(date));
    }
    
    // Check if row is complete (has all financial data)
    const hasFinancialData = row[columnIndex['Total Sales']] && 
                            row[columnIndex['Cash Sales']] && 
                            row[columnIndex['Closing Cash']];
    
    if (hasFinancialData) {
      stats.completedRows++;
    } else {
      stats.incompleteRows++;
    }
  });
  
  if (dates.length > 0) {
    stats.dateRange.min = new Date(Math.min(...dates.map(d => d.getTime())));
    stats.dateRange.max = new Date(Math.max(...dates.map(d => d.getTime())));
  }
  
  stats.completionRate = stats.totalRows > 0 ? (stats.completedRows / stats.totalRows * 100).toFixed(1) + '%' : '0%';
  
  return stats;
}

/**
 * Logs spreadsheet statistics for monitoring
 */
function logSpreadsheetStats() {
  const stats = getSpreadsheetStats();
  
  Logger.log('=== SPREADSHEET STATISTICS ===');
  Logger.log('Total rows: ' + stats.totalRows);
  Logger.log('Branches: ' + stats.totalBranches);
  
  if (stats.dateRange.min && stats.dateRange.max) {
    Logger.log('Date range: ' + stats.dateRange.min.toLocaleDateString() + ' to ' + stats.dateRange.max.toLocaleDateString());
  }
  
  Logger.log('Completed rows: ' + stats.completedRows);
  Logger.log('Incomplete rows: ' + stats.incompleteRows);
  Logger.log('Completion rate: ' + stats.completionRate);
}

/**
 * Creates backup of current spreadsheet state
 * @returns {string} Backup creation timestamp
 */
function createSpreadsheetBackup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    Logger.log('Sheet not found for backup');
    return null;
  }
  
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const backupName = CONFIG.SHEET_NAME + '_backup_' + timestamp;
  
  try {
    // Create backup by copying the sheet
    sheet.copyTo(spreadsheet).setName(backupName);
    Logger.log('Backup created: ' + backupName);
    return timestamp;
  } catch (error) {
    Logger.log('Backup failed: ' + error.message);
    return null;
  }
}

/**
 * Exports processing data to CSV for external analysis
 * @returns {string} CSV data as string
 */
function exportDataToCSV() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return '';
  
  const data = sheet.getDataRange().getValues();
  const csvContent = data.map(row => 
    row.map(cell => {
      // Handle cells that might contain commas or quotes
      if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"'))) {
        return '"' + cell.replace(/"/g, '""') + '"';
      }
      return cell;
    }).join(',')
  ).join('\n');
  
  return csvContent;
}

/**
 * Checks for data anomalies in the spreadsheet
 * @returns {Array} Array of detected anomalies
 */
function findDataAnomalies() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return [];
  
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  const columnIndex = {};
  headers.forEach((header, index) => {
    columnIndex[header] = index;
  });
  
  const anomalies = [];
  
  dataRows.forEach((row, index) => {
    const rowNumber = index + 2;
    
    // Check for negative values where they shouldn't exist
    const positiveFields = ['Total Sales', 'Cash Sales', 'Opening Cash', 'Closing Cash'];
    positiveFields.forEach(field => {
      if (columnIndex[field] !== undefined && row[columnIndex[field]] < 0) {
        anomalies.push({
          type: 'NEGATIVE_VALUE',
          field: field,
          row: rowNumber,
          value: row[columnIndex[field]],
          message: 'Negative value in ' + field
        });
      }
    });
    
    // Check for unusually high or low values
    if (columnIndex['Total Sales'] !== undefined && row[columnIndex['Total Sales']] > 0) {
      const sales = row[columnIndex['Total Sales']];
      
      // Example threshold - adjust based on business logic
      if (sales > 100000) {
        anomalies.push({
          type: 'UNUSUALLY_HIGH',
          field: 'Total Sales',
          row: rowNumber,
          value: sales,
          message: 'Unusually high sales amount'
        });
      }
      
      if (sales < 100) {
        anomalies.push({
          type: 'UNUSUALLY_LOW',
          field: 'Total Sales',
          row: rowNumber,
          value: sales,
          message: 'Unusually low sales amount'
        });
      }
    }
    
    // Check data consistency between related fields
    if (columnIndex['Opening Cash'] !== undefined && columnIndex['Closing Cash'] !== undefined) {
      const opening = row[columnIndex['Opening Cash']] || 0;
      const closing = row[columnIndex['Closing Cash']] || 0;
      
      if (closing > 0 && opening > 0 && closing < opening * 0.5) {
        anomalies.push({
          type: 'CASH_DISCREPANCY',
          field: 'Closing Cash',
          row: rowNumber,
          value: closing,
          message: 'Closing cash is less than 50% of opening cash'
        });
      }
    }
  });
  
  return anomalies;
}

// Export functions for testing and external use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    updateSpreadsheet,
    applySpreadsheetUpdates,
    highlightUpdatedRows,
    findOrCreateRow,
    validateFinancialData,
    getSpreadsheetStats,
    logSpreadsheetStats,
    createSpreadsheetBackup,
    exportDataToCSV,
    findDataAnomalies
  };
}
