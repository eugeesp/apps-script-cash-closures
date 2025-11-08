/**
 * CashFlow Automator - User Interface Controls
 * Handles custom menus, dialogs, status monitoring, and user interactions
 * @version 2.1.0
 */

/* ==================== MENU SYSTEM ==================== */

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 PDF Processor')
    .addItem('🚀 Start Full Processing', 'startProcessing')
    .addItem('📊 View System Status', 'viewStatus')
    .addSeparator()
    .addItem('⏸️ Pause Processing', 'pauseProcessing')
    .addItem('▶️ Resume Processing', 'resumeProcessing')
    .addSeparator()
    .addItem('🔄 Process Specific Date', 'showDateDialog')
    .addItem('📋 Browse Available Dates', 'showAvailableFolders')
    .addItem('📝 Process From Cell', 'processFromCell')
    .addSeparator()
    .addItem('📧 Process Emails', 'showEmailDialog')
    .addItem('🔄 Reprocess Date', 'showReprocessDialog')
    .addToUi();
}

/**
 * Shows dialog for processing specific date
 */
function showDateDialog() {
  const ui = SpreadsheetApp.getUi();
  
  // Get available dates for suggestions
  let dateSuggestions = '';
  try {
    const root = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER).next();
    const subfolders = root.getFolders();
    const dates = [];
    
    while (subfolders.hasNext()) {
      const name = subfolders.next().getName();
      if (/^\d{4}-\d{2}-\d{2}$/.test(name)) dates.push(name);
    }
    
    if (dates.length > 0) {
      dates.sort().reverse();
      dateSuggestions = '\n\nRecent dates available:\n• ' + dates.slice(0, 5).join('\n• ');
    }
  } catch (e) {
    // Continue without suggestions if folder access fails
  }
  
  const response = ui.prompt(
    'Process Specific Date',
    'Enter date to process (format: yyyy-mm-dd):\n\nExample: 2025-07-22' + dateSuggestions,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const date = response.getResponseText().trim();
    
    // Validate date format
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      ui.alert('Invalid Format', 'Please use: yyyy-mm-dd (example: 2025-07-22)', ui.ButtonSet.OK);
      return;
    }
    
    // Confirm processing
    const confirmation = ui.alert(
      'Confirm Processing',
      'Process all PDFs for date ' + date + '?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmation === ui.Button.YES) {
      ui.alert('Processing Started', 'Starting processing for ' + date + '.\nCheck logs for progress.', ui.ButtonSet.OK);
      processDateFolder(date);
    }
  }
}

/**
 * Shows dialog for email processing
 */
function showEmailDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Process Emails',
    'Enter date range to process emails (format: yyyy/mm/dd):\n\nExample: 2025/07/01 to 2025/07/31\n\nLeave empty for default range:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText().trim();
    
    if (input === '') {
      // Use default range
      procesarMisFechas();
    } else {
      // Parse date range
      const dates = input.split(' to ');
      if (dates.length === 2) {
        const startDate = dates[0].trim();
        const endDate = dates[1].trim();
        
        if (/^\d{4}\/\d{2}\/\d{2}$/.test(startDate) && /^\d{4}\/\d{2}\/\d{2}$/.test(endDate)) {
          ui.alert('Email Processing Started', 
                  'Processing emails from ' + startDate + ' to ' + endDate + '.\nCheck logs for progress.', 
                  ui.ButtonSet.OK);
          procesarCorreos(startDate, endDate);
        } else {
          ui.alert('Invalid Format', 'Please use: yyyy/mm/dd to yyyy/mm/dd', ui.ButtonSet.OK);
        }
      } else {
        ui.alert('Invalid Format', 'Please use: start_date to end_date', ui.ButtonSet.OK);
      }
    }
  }
}

/**
 * Shows dialog for reprocessing specific date
 */
function showReprocessDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Reprocess Date',
    'Enter date to reprocess (format: yyyy-mm-dd):\n\nThis will remove existing entries and reprocess all files for this date.',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const date = response.getResponseText().trim();
    
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      ui.alert('Invalid Format', 'Please use: yyyy-mm-dd (example: 2025-07-22)', ui.ButtonSet.OK);
      return;
    }
    
    const confirmation = ui.alert(
      'Confirm Reprocessing',
      'Reprocess all files for date ' + date + '?\n\nThis will remove existing processing entries and start fresh.',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmation === ui.Button.YES) {
      ui.alert('Reprocessing Started', 'Starting reprocessing for ' + date + '.\nCheck logs for progress.', ui.ButtonSet.OK);
      reprocesarCorreosFecha(date);
    }
  }
}

/* ==================== PROCESSING CONTROLS ==================== */

/**
 * Pauses the processing system
 */
function pauseProcessing() {
  PropertiesService.getScriptProperties().setProperty('processing_active', 'false');
  clearTriggers();
  Logger.log('Processing paused');
  SpreadsheetApp.getUi().alert('Processing Paused', 'Processing has been paused.', ui.ButtonSet.OK);
}

/**
 * Resumes the processing system
 */
function resumeProcessing() {
  PropertiesService.getScriptProperties().setProperty('processing_active', 'true');
  Logger.log('Resuming processing...');
  SpreadsheetApp.getUi().alert('Processing Resumed', 'Processing has been resumed.', ui.ButtonSet.OK);
  processNextBatch();
}

/**
 * Processes files based on cell input
 */
function processFromCell() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Find cell with "PROCESS_DATE:" in column A
  const range = sheet.getRange('A:A');
  const values = range.getValues();
  let dateRow = -1;
  
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().includes('PROCESS_DATE:')) {
      dateRow = i + 1;
      break;
    }
  }
  
  if (dateRow === -1) {
    SpreadsheetApp.getUi().alert('Configuration Needed', 
                                'Add "PROCESS_DATE: 2025-07-22" in column A to trigger processing.', 
                                ui.ButtonSet.OK);
    return;
  }
  
  // Get date from column B in same row
  const dateCell = sheet.getRange(dateRow, 2);
  const date = dateCell.getValue();
  
  if (!date) {
    SpreadsheetApp.getUi().alert('Date Required', 
                                'Cell B' + dateRow + ' is empty. Enter date to process.', 
                                ui.ButtonSet.OK);
    return;
  }
  
  let formattedDate;
  if (date instanceof Date) {
    formattedDate = Utilities.formatDate(date, 'GMT-3', 'yyyy-MM-dd');
  } else {
    formattedDate = date.toString().trim();
  }
  
  Logger.log('Processing date from cell: ' + formattedDate);
  processDateFolder(formattedDate);
  
  // Mark as processed
  dateCell.setValue('Processed: ' + new Date().toLocaleString('es-AR'));
  SpreadsheetApp.getUi().alert('Processing Complete', 
                              'Date ' + formattedDate + ' processed successfully.\nCell updated with timestamp.', 
                              ui.ButtonSet.OK);
}

/* ==================== STATUS MONITORING ==================== */

/**
 * Shows current system status
 */
function viewStatus() {
  const props = PropertiesService.getScriptProperties();
  const active = props.getProperty('processing_active') === 'true';
  const batch = props.getProperty('current_batch') || '1';
  const processed = props.getProperty('files_processed') || '0';
  const startTime = props.getProperty('start_time');
  
  let statusMessage = '=== SYSTEM STATUS ===\n';
  statusMessage += 'Status: ' + (active ? 'ACTIVE 🟢' : 'PAUSED 🔴') + '\n';
  statusMessage += 'Current batch: ' + batch + '\n';
  statusMessage += 'Files processed: ' + processed + '\n';
  
  if (startTime) {
    const startDate = new Date(startTime);
    const elapsedMinutes = Math.round((new Date() - startDate) / (1000 * 60));
    statusMessage += 'Started: ' + startDate.toLocaleString('es-AR') + '\n';
    statusMessage += 'Elapsed time: ' + elapsedMinutes + ' minutes\n';
    
    if (processed > 0 && elapsedMinutes > 0) {
      const speed = (processed / elapsedMinutes).toFixed(1);
      statusMessage += 'Processing speed: ' + speed + ' files/minute\n';
    }
  }
  
  // Count pending files
  try {
    const folder = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER).next();
    const pdfs = folder.getFilesByType(MimeType.PDF);
    let pending = 0;
    
    while (pdfs.hasNext()) {
      pdfs.next();
      pending++;
    }
    
    statusMessage += 'Pending files in root: ' + pending + '\n';
    
    if (pending > 0 && active) {
      const batchesRemaining = Math.ceil(pending / CONFIG.BATCH_SIZE);
      const estimatedMinutes = batchesRemaining * (CONFIG.DELAY_SECONDS / 60);
      statusMessage += 'Batches remaining: ~' + batchesRemaining + '\n';
      statusMessage += 'Estimated time: ~' + Math.round(estimatedMinutes) + ' minutes\n';
    }
  } catch (e) {
    statusMessage += 'Error checking pending files: ' + e.message + '\n';
  }
  
  // Check scheduled triggers
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processNextBatch');
  statusMessage += 'Scheduled triggers: ' + triggers.length;
  
  // Show in alert dialog
  SpreadsheetApp.getUi().alert('System Status', statusMessage, ui.ButtonSet.OK);
  
  // Also log to console
  Logger.log(statusMessage);
}

/**
 * Shows detailed folder statistics
 */
function showFolderStats() {
  const stats = getFolderStats();
  
  if (stats.error) {
    SpreadsheetApp.getUi().alert('Error', 'Could not get folder stats: ' + stats.error, ui.ButtonSet.OK);
    return;
  }
  
  let statsMessage = '=== FOLDER STATISTICS ===\n';
  statsMessage += 'Total files: ' + stats.totalFiles + '\n';
  statsMessage += 'PDF files: ' + stats.totalPDFs + '\n';
  statsMessage += 'Date folders: ' + stats.dateFolders + '\n';
  statsMessage += 'Total size: ' + stats.totalSizeMB + ' MB\n';
  
  if (stats.dateFolders > 0) {
    statsMessage += '\nFiles by date folder:\n';
    Object.keys(stats.filesByDate).sort().reverse().slice(0, 10).forEach(date => {
      const folderStats = stats.filesByDate[date];
      statsMessage += '• ' + date + ': ' + folderStats.pdfFiles + ' PDFs\n';
    });
    
    if (Object.keys(stats.filesByDate).length > 10) {
      statsMessage += '• ... and ' + (Object.keys(stats.filesByDate).length - 10) + ' more\n';
    }
  }
  
  SpreadsheetApp.getUi().alert('Folder Statistics', statsMessage, ui.ButtonSet.OK);
}

/**
 * Shows spreadsheet statistics
 */
function showSpreadsheetStats() {
  const stats = getSpreadsheetStats();
  
  if (stats.error) {
    SpreadsheetApp.getUi().alert('Error', 'Could not get spreadsheet stats: ' + stats.error, ui.ButtonSet.OK);
    return;
  }
  
  let statsMessage = '=== SPREADSHEET STATISTICS ===\n';
  statsMessage += 'Total rows: ' + stats.totalRows + '\n';
  statsMessage += 'Branches: ' + stats.totalBranches + '\n';
  
  if (stats.dateRange.min && stats.dateRange.max) {
    statsMessage += 'Date range: ' + stats.dateRange.min.toLocaleDateString() + ' to ' + stats.dateRange.max.toLocaleDateString() + '\n';
  }
  
  statsMessage += 'Completed rows: ' + stats.completedRows + '\n';
  statsMessage += 'Incomplete rows: ' + stats.incompleteRows + '\n';
  statsMessage += 'Completion rate: ' + stats.completionRate + '\n';
  
  SpreadsheetApp.getUi().alert('Spreadsheet Statistics', statsMessage, ui.ButtonSet.OK);
}

/* ==================== QUICK ACTIONS ==================== */

/**
 * Processes today's files
 */
function processToday() {
  const today = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd');
  SpreadsheetApp.getUi().alert('Processing Today', 'Processing files for today: ' + today, ui.ButtonSet.OK);
  processDateFolder(today);
}

/**
 * Processes yesterday's files
 */
function processYesterday() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayDate = Utilities.formatDate(yesterday, 'GMT-3', 'yyyy-MM-dd');
  SpreadsheetApp.getUi().alert('Processing Yesterday', 'Processing files for yesterday: ' + yesterdayDate, ui.ButtonSet.OK);
  processDateFolder(yesterdayDate);
}

/**
 * Shows available dates for processing
 */
function showAvailableDates() {
  try {
    const root = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER).next();
    const subfolders = root.getFolders();
    const options = [];
    
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      const name = folder.getName();
      if (/^\d{4}-\d{2}-\d{2}$/.test(name)) {
        options.push(name);
      }
    }
    
    options.sort().reverse();
    
    if (options.length === 0) {
      SpreadsheetApp.getUi().alert('No Dates', 'No date folders available.', ui.ButtonSet.OK);
      return;
    }
    
    let datesMessage = 'Available date folders:\n\n';
    options.slice(0, 15).forEach((date, i) => {
      datesMessage += (i + 1) + '. ' + date + '\n';
    });
    
    if (options.length > 15) {
      datesMessage += '... and ' + (options.length - 15) + ' more\n';
    }
    
    datesMessage += '\nTo process, use: processDateFolder("2025-07-22")';
    datesMessage += '\nMost recent: processDateFolder("' + options[0] + '")';
    
    SpreadsheetApp.getUi().alert('Available Dates', datesMessage, ui.ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Error: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Runs data diagnostics
 */
function runDiagnostics() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Run Diagnostics',
    'Enter date range for diagnostics (format: yyyy/mm/dd):\n\nExample: 2025/07/01 to 2025/07/31',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText().trim();
    const dates = input.split(' to ');
    
    if (dates.length === 2) {
      const startDate = dates[0].trim();
      const endDate = dates[1].trim();
      
      if (/^\d{4}\/\d{2}\/\d{2}$/.test(startDate) && /^\d{4}\/\d{2}\/\d{2}$/.test(endDate)) {
        ui.alert('Diagnostics Started', 
                'Running diagnostics from ' + startDate + ' to ' + endDate + '.\nCheck logs for results.', 
                ui.ButtonSet.OK);
        diagnosticarCorreosProblematicos(startDate, endDate);
      } else {
        ui.alert('Invalid Format', 'Please use: yyyy/mm/dd to yyyy/mm/dd', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Invalid Format', 'Please use: start_date to end_date', ui.ButtonSet.OK);
    }
  }
}

/* ==================== UTILITY FUNCTIONS ==================== */

/**
 * Shows system information
 */
function showSystemInfo() {
  const infoMessage = '=== SYSTEM INFORMATION ===\n' +
                     'CashFlow Automator v2.1.0\n' +
                     'Main Folder: ' + CONFIG.MAIN_FOLDER + '\n' +
                     'Batch Size: ' + CONFIG.BATCH_SIZE + ' files\n' +
                     'Delay Between Batches: ' + CONFIG.DELAY_SECONDS + ' seconds\n' +
                     'Max Retries: ' + CONFIG.MAX_RETRIES + '\n' +
                     'Email Processing: ' + (CONFIG.EMAIL_BATCH_SIZE > 0 ? 'Enabled' : 'Disabled') + '\n' +
                     'Duplicate Detection: Enabled\n' +
                     '\nBuilt with Google Apps Script';
  
  SpreadsheetApp.getUi().alert('System Information', infoMessage, ui.ButtonSet.OK);
}

/**
 * Tests data extraction with sample data
 */
function testDataExtraction() {
  const testText = 
    'Company Name: Sample Business - Main Branch\n' +
    'Closure date: 15/07/2025 14:30:00\n' +
    'Opening cash: $1,500.00\n' +
    'Total sales: $45,230.75\n' +
    'Cash: $25,100.00\n' +
    'Cards: $18,130.75\n' +
    'Digital: $2,000.00\n' +
    'Closing cash: $3,200.50\n' +
    'Withdrawal at Closure: -$23,400.50';
  
  const result = extractPDFData(testText, 'test_file.pdf');
  
  let testMessage = '=== DATA EXTRACTION TEST ===\n';
  testMessage += 'Branch: ' + result.branch + '\n';
  testMessage += 'Date: ' + result.closureDate + '\n';
  testMessage += 'Time: ' + result.closureTime + '\n';
  testMessage += 'Shift: ' + result.shift + '\n';
  testMessage += 'Opening Cash: ' + result.openingCash + '\n';
  testMessage += 'Total Sales: ' + result.totalSales + '\n';
  testMessage += 'Cash Sales: ' + result.cashSales + '\n';
  testMessage += 'Card Sales: ' + result.cardSales + '\n';
  testMessage += 'Digital Payments: ' + result.digitalPayments + '\n';
  testMessage += 'Closing Cash: ' + result.closingCash + '\n';
  testMessage += 'Withdrawal: ' + result.cashWithdrawal;
  
  SpreadsheetApp.getUi().alert('Data Extraction Test', testMessage, ui.ButtonSet.OK);
  Logger.log('Data extraction test completed successfully');
}

// Export functions for testing and external use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    onOpen,
    showDateDialog,
    showEmailDialog,
    showReprocessDialog,
    pauseProcessing,
    resumeProcessing,
    processFromCell,
    viewStatus,
    showFolderStats,
    showSpreadsheetStats,
    processToday,
    processYesterday,
    showAvailableDates,
    runDiagnostics,
    showSystemInfo,
    testDataExtraction
  };
}
