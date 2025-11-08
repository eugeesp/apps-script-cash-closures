/**
 * CASHFLOW AUTOMATOR - Core Processing Engine
 * Main module for PDF processing and data extraction
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

/* ==================== MAIN PROCESSING ENGINE ==================== */
function processFiles(files) {
  const rows = [];
  const filesByDate = new Map();
  const foundDates = new Set();
  
  Logger.log(`Processing ${files.length} files...`);
  
  files.forEach((pdf, index) => {
    let tempDocId = null;
    
    try {
      // Progress tracking
      if (index % 5 === 0 || index === files.length - 1) {
        Logger.log(`Processing ${index + 1}/${files.length}: ${pdf.getName()}`);
      }
      
      // Convert PDF to editable text
      const docMetadata = Drive.Files.copy(
        { title: pdf.getName() + ' (temp)', mimeType: MimeType.GOOGLE_DOCS },
        pdf.getId()
      );
      tempDocId = docMetadata.id;
      Utilities.sleep(2000);
      
      const exportUrl = `https://docs.google.com/document/d/${tempDocId}/export?format=txt`;
      const response = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
      });
      
      if (response.getResponseCode() !== 200) {
        throw new Error(`Export failed: HTTP ${response.getResponseCode()}`);
      }
      
      const textContent = response.getContentText();
      if (!textContent.trim()) throw new Error('PDF contains no extractable text');
      
      // Extract structured data from text
      const extractedData = extractPDFData(textContent, pdf.getName());
      
      if (extractedData.closureDate) {
        rows.push(extractedData);
        foundDates.add(extractedData.closureDate);
        
        // Organize files by date for later processing
        const normalizedDate = normalizeDate(extractedData.closureDate);
        if (normalizedDate) {
          if (!filesByDate.has(normalizedDate)) {
            filesByDate.set(normalizedDate, []);
          }
          filesByDate.get(normalizedDate).push(pdf);
        }
        
        // Sample logging for monitoring
        if (index < 3 || index === files.length - 1) {
          Logger.log(`Processed: ${pdf.getName()} → ${extractedData.closureDate} (${extractedData.branch})`);
        }
      } else {
        throw new Error('No closure date found in document');
      }
    } catch (error) {
      Logger.log(`Failed: ${pdf.getName()} - ${error.message}`);
      rows.push({ file: pdf.getName(), error: error.message });
    } finally {
      // Cleanup temporary conversion document
      if (tempDocId) {
        try {
          DriveApp.getFileById(tempDocId).setTrashed(true);
        } catch (cleanupError) {
          // Silent cleanup - non-critical
        }
      }
    }
  });
  
  // Update spreadsheet with successful extractions
  const successfulRows = rows.filter(r => !r.error);
  if (successfulRows.length > 0) {
    Logger.log(`Updating spreadsheet with ${successfulRows.length} records...`);
    updateSpreadsheet(successfulRows);
    
    Logger.log(`Organizing files into date-based folders...`);
    const rootFolder = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER).next();
    organizeFiles(rootFolder, filesByDate);
  }
  
  return rows;
}

/* ==================== DATA EXTRACTION LOGIC ==================== */
function extractPDFData(text, filename) {
  // Locate company information line
  const companyLine = text.split('\n')
    .find(line => line.toLowerCase().includes('company name:') && 
                  line.toLowerCase().includes('sample business')) || '';
  
  const branch = (companyLine.match(/SAMPLE BUSINESS\s*[-\s]*(.+)/i) || [])[1]?.trim() || '';
  
  // Extract date and time patterns
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

/* ==================== BATCH PROCESSING SYSTEM ==================== */
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
  Logger.log(`=== PROCESSING STARTED ===`);
  Logger.log(`Started: ${now}`);
  Logger.log(`Source: ${CONFIG.MAIN_FOLDER}`);
  Logger.log(`Batch size: ${CONFIG.BATCH_SIZE} files`);
  
  processNextBatch();
}

function processNextBatch() {
  const props = PropertiesService.getScriptProperties();
  
  if (props.getProperty('processing_active') !== 'true') {
    Logger.log('Processing paused');
    return;
  }
  
  const currentBatch = parseInt(props.getProperty('current_batch') || '1');
  const totalProcessed = parseInt(props.getProperty('files_processed') || '0');
  
  Logger.log(`Batch ${currentBatch} - Total processed: ${totalProcessed}`);
  
  try {
    const rootFolder = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER).next();
    const pdfs = rootFolder.getFilesByType(MimeType.PDF);
    const pendingFiles = [];
    
    while (pdfs.hasNext()) {
      pendingFiles.push(pdfs.next());
    }
    
    Logger.log(`Pending files: ${pendingFiles.length}`);
    
    if (pendingFiles.length === 0) {
      Logger.log('=== PROCESSING COMPLETED ===');
      Logger.log('All files processed successfully');
      finishProcessing('COMPLETED');
      return;
    }
    
    const batchFiles = pendingFiles.slice(0, CONFIG.BATCH_SIZE);
    Logger.log(`Processing ${batchFiles.length} files...`);
    
    const results = processFiles(batchFiles);
    const successful = results.filter(r => !r.error).length;
    const failed = results.length - successful;
    
    props.setProperties({
      'current_batch': (currentBatch + 1).toString(),
      'files_processed': (totalProcessed + successful).toString(),
      'failed_attempts': '0'
    });
    
    Logger.log(`Batch ${currentBatch} results: ${successful} successful, ${failed} failed`);
    Logger.log(`Total: ${totalProcessed + successful} files`);
    Logger.log(`Remaining: ${pendingFiles.length - batchFiles.length}`);
    
    if (pendingFiles.length > batchFiles.length) {
      Logger.log(`Next batch in ${CONFIG.DELAY_SECONDS} seconds...`);
      scheduleNextBatch();
    } else {
      Logger.log('All batches completed');
      finishProcessing('COMPLETED');
    }
  } catch (error) {
    Logger.log(`Batch ${currentBatch} error: ${error.message}`);
    handleError();
  }
}



