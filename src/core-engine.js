/**
 * CashFlow Automator - Core Processing Engine
 * Main system for PDF processing, data extraction, and batch management
 * @version 2.1.0
 */

/* ==================== SYSTEM CONFIGURATION ==================== */
const CONFIG = {
  SHEET_NAME: 'Financial_Reports_2025',
  MAIN_FOLDER: 'PDF_PROCESSING_MAIN',
  BATCH_SIZE: 18,
  DELAY_SECONDS: 30,
  MAX_RETRIES: 3,
  DESTINATION_FOLDER_ID: "your_drive_folder_id_here",
  INDEX_FILE_NAME: "processing_index.doc",
  EMAIL_BATCH_SIZE: 8,
  SHIFT_CUTOFF_HOUR: 16,
  EMAIL_SUBJECT_REGEX: /business\s+(.*?)\s+-\s+Daily Closure Report\s+-\s+(\d{2}\/\d{2}\/\d{4})\s+-\s+(\d{2}:\d{2}:\d{2})/,
  MAX_EXECUTION_TIME: 5 * 60 * 1000
};

/* ==================== MAIN PROCESSING FUNCTIONS ==================== */

/**
 * Starts the automated processing system
 * Initializes properties and begins batch processing
 */
function startProcessing() {
  clearTriggers();
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'processing_active': 'true',
    'current_batch': '1',
    'files_processed': '0',
    'failed_attempts': '0',
    'start_time': new Date().toISOString()
  });
  
  const now = new Date().toLocaleString('es-AR');
  Logger.log('=== PROCESSING STARTED ===');
  Logger.log('Date/Time: ' + now);
  Logger.log('Folder: ' + CONFIG.MAIN_FOLDER);
  Logger.log('Batch size: ' + CONFIG.BATCH_SIZE + ' files, ' + CONFIG.DELAY_SECONDS + 's delay');
  
  processNextBatch();
}

/**
 * Processes the next batch of files
 * Handles batch management and progress tracking
 */
function processNextBatch() {
  const props = PropertiesService.getScriptProperties();
  
  if (props.getProperty('processing_active') !== 'true') {
    Logger.log('Processing paused');
    return;
  }
  
  const currentBatch = parseInt(props.getProperty('current_batch') || '1');
  const totalProcessed = parseInt(props.getProperty('files_processed') || '0');
  
  Logger.log('=== BATCH ' + currentBatch + ' ===');
  Logger.log('Total processed: ' + totalProcessed + ' files');
  
  try {
    // Get pending files
    const rootFolder = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER).next();
    const pdfs = rootFolder.getFilesByType(MimeType.PDF);
    const pendingFiles = [];
    
    while (pdfs.hasNext()) {
      pendingFiles.push(pdfs.next());
    }
    
    Logger.log('Pending files in root folder: ' + pendingFiles.length);
    
    if (pendingFiles.length === 0) {
      Logger.log('=== PROCESSING COMPLETED ===');
      Logger.log('No more files to process');
      finishProcessing('COMPLETED');
      return;
    }
    
    // Process current batch
    const batchFiles = pendingFiles.slice(0, CONFIG.BATCH_SIZE);
    Logger.log('Processing ' + batchFiles.length + ' files in this batch...');
    
    const results = processFiles(batchFiles);
    const successful = results.filter(r => !r.error).length;
    const failed = results.length - successful;
    
    // Update progress
    props.setProperties({
      'current_batch': (currentBatch + 1).toString(),
      'files_processed': (totalProcessed + successful).toString(),
      'failed_attempts': '0'
    });
    
    Logger.log('=== BATCH ' + currentBatch + ' RESULTS ===');
    Logger.log('Successful: ' + successful);
    Logger.log('Failed: ' + failed);
    Logger.log('Total accumulated: ' + (totalProcessed + successful));
    Logger.log('Remaining: ' + (pendingFiles.length - batchFiles.length));
    
    // Continue if there are more files
    if (pendingFiles.length > batchFiles.length) {
      Logger.log('Next batch in ' + CONFIG.DELAY_SECONDS + ' seconds...');
      scheduleNextBatch();
    } else {
      Logger.log('All batches completed');
      finishProcessing('COMPLETED');
    }
  } catch (error) {
    Logger.log('Batch ' + currentBatch + ' error: ' + error.message);
    handleError();
  }
}

/**
 * Processes individual PDF files and extracts financial data
 * @param {Array} files - Array of PDF files to process
 * @returns {Array} Processing results with extracted data or errors
 */
function processFiles(files) {
  const rows = [];
  const filesByDate = new Map();
  const foundDates = new Set();
  
  Logger.log('Processing ' + files.length + ' files...');
  
  files.forEach((pdf, index) => {
    let tempDocId = null;
    
    try {
      // Progress logging
      if (index % 5 === 0 || index === files.length - 1) {
        Logger.log('Progress: ' + (index + 1) + '/' + files.length + ' - ' + pdf.getName());
      }
      
      // Convert PDF to text
      const docMetadata = Drive.Files.copy(
        { title: pdf.getName() + ' (temp)', mimeType: MimeType.GOOGLE_DOCS },
        pdf.getId()
      );
      tempDocId = docMetadata.id;
      Utilities.sleep(2000);
      
      const exportUrl = 'https://docs.google.com/document/d/' + tempDocId + '/export?format=txt';
      const response = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
      });
      
      if (response.getResponseCode() !== 200) throw new Error('HTTP ' + response.getResponseCode());
      
      const textContent = response.getContentText();
      if (!textContent.trim()) throw new Error('PDF without extractable text');
      
      // Extract data
      const extractedData = extractPDFData(textContent, pdf.getName());
      
      if (extractedData.closureDate) {
        rows.push(extractedData);
        foundDates.add(extractedData.closureDate);
        
        // Group by date
        const normalizedDate = normalizeDate(extractedData.closureDate);
        if (normalizedDate) {
          if (!filesByDate.has(normalizedDate)) {
            filesByDate.set(normalizedDate, []);
          }
          filesByDate.get(normalizedDate).push(pdf);
        }
        
        // Log important data
        if (index < 3 || index === files.length - 1) {
          Logger.log('Processed: ' + pdf.getName() + ' → ' + extractedData.closureDate + ' (' + extractedData.branch + ') ' + extractedData.shift);
        }
      } else {
        throw new Error('Could not extract date');
      }
    } catch (error) {
      Logger.log('Failed: ' + pdf.getName() + ' → Error: ' + error.message);
      rows.push({ file: pdf.getName(), error: error.message });
    } finally {
      if (tempDocId) {
        try {
          DriveApp.getFileById(tempDocId).setTrashed(true);
        } catch (_) {}
      }
    }
  });
  
  // Show processed dates summary
  if (foundDates.size > 0) {
    const datesArray = Array.from(foundDates).sort();
    Logger.log('Dates processed in this batch: ' + datesArray.join(', '));
  }
  
  // Update sheet and organize files
  const successfulRows = rows.filter(r => !r.error);
  if (successfulRows.length > 0) {
    Logger.log('Updating ' + successfulRows.length + ' rows in spreadsheet...');
    updateSpreadsheet(successfulRows);
    
    Logger.log('Organizing files into ' + filesByDate.size + ' date folders...');
    const rootFolder = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER).next();
    organizeFiles(rootFolder, filesByDate);
  }
  
  return rows;
}

/* ==================== DATA EXTRACTION FUNCTIONS ==================== */

/**
 * Extracts financial data from PDF text content
 * @param {string} text - Text content extracted from PDF
 * @param {string} filename - Original filename for reference
 * @returns {Object} Structured financial data
 */
function extractPDFData(text, filename) {
  const companyLine = text.split('\n')
    .find(line => line.toLowerCase().includes('company name:') && line.toLowerCase().includes('sample business')) || '';
  
  const branch = (companyLine.match(/SAMPLE BUSINESS\s*[-\s]*(.+)/i) || [])[1]?.trim() || '';
  
  const dateMatch = text.match(/Closure date:\s*(\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2}:\d{2})/);
  const closureDate = dateMatch ? dateMatch[1] : '';
  const closureTime = dateMatch ? dateMatch[2] : '';
  const shift = closureTime && parseInt(closureTime.split(':')[0], 10) < 16 ? 'Morning' : 'Evening';
  
  return {
    file: filename,
    closureDate,
    closureTime,
    shift,
    branch,
    openingCash: extractAmount(text, /Opening cash:/i),
    totalSales: extractAmount(text, /Total sales:/i),
    cashSales: extractAmount(text, /(?:^|\n)\s*Cash:/i),
    cardSales: extractAmount(text, /Cards:/i),
    digitalPayments: extractAmount(text, /Digital:/i),
    closingCash: extractAmount(text, /Closing cash:/i),
    cashWithdrawal: extractCashWithdrawal(text)
  };
}

/**
 * Extracts monetary amounts from text using regex patterns
 * @param {string} text - Text to search
 * @param {RegExp} labelRegex - Pattern to identify amount labels
 * @returns {string} Extracted amount or empty string
 */
function extractAmount(text, labelRegex) {
  const pattern = new RegExp(labelRegex.source + '\\s*\\$?\\s*([0-9]{1,3}(?:\\.[0-9]{3})*,[0-9]{2})', 'i');
  const match = text.match(pattern);
  return match ? match[1] : '';
}

/**
 * Extracts cash withdrawal amounts from closure reports
 * @param {string} text - Text to search
 * @returns {string} Extracted withdrawal amount or empty string
 */
function extractCashWithdrawal(text) {
  const lines = text.split('\n');
  const idx = lines.findIndex(line => /Withdrawal\s+at\s+Closure\s*-?/i.test(line));
  
  if (idx === -1) return '';
  
  for (let i = idx; i < Math.min(idx + 3, lines.length); i++) {
    const match = lines[i].match(/-?\$?\s*([0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2})/);
    if (match) return match[1];
  }
  
  return '';
}

/* ==================== UTILITY FUNCTIONS ==================== */

/**
 * Normalizes date formats to consistent ISO format
 * @param {*} value - Date value to normalize
 * @returns {string} Normalized date in YYYY-MM-DD format
 */
function normalizeDate(value) {
  if (!value) return '';
  
  if (value instanceof Date) {
    return Utilities.formatDate(value, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  
  const match = value.toString().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return match ? match[3] + '-' + match[2] + '-' + match[1] : value.toString().trim();
}

/**
 * Converts Argentine number format to standard numeric format
 * @param {string} str - Number string in Argentine format
 * @returns {string} Converted number or empty string
 */
function convertArgentineNumber(str) {
  if (!str) return '';
  
  let cleaned = str.replace(/[^0-9.,-]/g, '');
  
  if (cleaned.includes('.') && cleaned.includes(',')) {
    cleaned = cleaned.replace(/\./g, '').replace(',', '.');
  } else if (cleaned.includes(',')) {
    cleaned = cleaned.replace(',', '.');
  } else if ((cleaned.match(/\./g) || []).length > 1) {
    cleaned = cleaned.replace(/\./g, '');
  }
  
  const number = parseFloat(cleaned);
  return isNaN(number) ? '' : Math.abs(number);
}

/**
 * Schedules the next processing batch
 */
function scheduleNextBatch() {
  clearTriggers();
  ScriptApp.newTrigger('processNextBatch')
    .timeBased()
    .after(CONFIG.DELAY_SECONDS * 1000)
    .create();
  Logger.log('Next batch in ' + CONFIG.DELAY_SECONDS + 's');
}

/**
 * Clears all existing processing triggers
 */
function clearTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'processNextBatch')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Handles processing errors with retry logic
 */
function handleError() {
  const props = PropertiesService.getScriptProperties();
  const attempts = parseInt(props.getProperty('failed_attempts') || '0') + 1;
  
  if (attempts >= CONFIG.MAX_RETRIES) {
    finishProcessing('MULTIPLE_ERRORS');
  } else {
    props.setProperty('failed_attempts', attempts.toString());
    Logger.log('Retry ' + attempts + '/' + CONFIG.MAX_RETRIES);
    scheduleNextBatch();
  }
}

/**
 * Finalizes processing and cleans up resources
 * @param {string} reason - Reason for finishing
 */
function finishProcessing(reason) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('processing_active', 'false');
  clearTriggers();
  
  const processed = props.getProperty('files_processed') || '0';
  Logger.log('Finished: ' + reason + ' - ' + processed + ' files processed');
}

// Export functions for testing and external use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    startProcessing,
    processNextBatch,
    processFiles,
    extractPDFData,
    extractAmount,
    extractCashWithdrawal,
    normalizeDate,
    convertArgentineNumber,
    CONFIG
  };
}
