/**
 * CashFlow Automator - Email Processing System
 * Handles Gmail integration, attachment processing, and file organization
 * @version 2.1.0
 */

/* ==================== EMAIL PROCESSING FUNCTIONS ==================== */

/**
 * Main function to process emails within a date range
 * @param {string} startDate - Start date (YYYY/MM/DD)
 * @param {string} endDate - End date (YYYY/MM/DD)
 * @param {boolean} forceReprocess - Whether to force reprocessing
 * @returns {Object} Processing results and statistics
 */
function processEmails(startDate, endDate, forceReprocess = false) {
  const startTime = Date.now();
  Logger.log('Processing emails from ' + startDate + ' to ' + endDate + (forceReprocess ? ' (FORCED)' : ''));
  
  const destinationFolder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
  
  // Initialize processing index
  const { indexFile, processedEmails } = initializeIndex(destinationFolder);
  Logger.log('Emails already in index: ' + processedEmails.size);
  
  // Build cache of existing files
  const existingFilesCache = buildExistingFilesCache(destinationFolder);
  Logger.log('Files in cache: ' + existingFilesCache.size);
  
  // Prepare search query
  const adjustedEndDate = adjustEndDate(endDate);
  const query = 'subject:"Daily Closure Report" has:attachment filename:pdf after:' + startDate + ' before:' + adjustedEndDate;
  Logger.log('Search query: ' + query);
  
  const threads = GmailApp.search(query);
  Logger.log('Threads found: ' + threads.length);
  
  if (threads.length === 0) {
    Logger.log('No emails found in this date range');
    return createEmptySummary();
  }
  
  // Process email threads
  const result = processEmailThreads(threads, processedEmails, existingFilesCache, destinationFolder, indexFile, forceReprocess, startTime);
  
  // Final summary
  showProcessingSummary(result);
  return result;
}

/**
 * Processes specific date emails (convenience function)
 * @returns {Object} Processing results
 */
function processMyDates() {
  // Configure date range for processing
  const startDate = "2025/07/05";
  const endDate = "2025/07/05";
  
  return processEmails(startDate, endDate);
}

/**
 * Reprocesses emails for a specific date
 * @param {string} date - Date to reprocess (YYYY/MM/DD)
 * @param {boolean} removeFromIndex - Whether to remove from index first
 * @returns {Object} Reprocessing results
 */
function reprocessDateEmails(date, removeFromIndex = true) {
  Logger.log('=== REPROCESSING DATE: ' + date + ' ===');
  
  if (removeFromIndex) {
    const destinationFolder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
    const { indexFile } = initializeIndex(destinationFolder);
    
    // Remove existing entries for this date
    let content = indexFile.getBlob().getDataAsString();
    const originalLineCount = content.split('\n').length - 1;
    
    const formattedDate = date.replace(/\//g, '-');
    const filteredLines = content.split('\n')
      .filter(line => !line.includes(formattedDate))
      .filter(line => line.trim() !== '');
    
    // Rewrite index file
    indexFile.setContent(filteredLines.join('\n') + '\n');
    const removedEntries = originalLineCount - filteredLines.length;
    Logger.log('Removed ' + removedEntries + ' entries from index');
  }
  
  // Process normally
  return processEmails(date, date, true);
}

/**
 * Diagnoses email processing issues
 * @param {string} startDate - Start date
 * @param {string} endDate - End date
 * @returns {Array} Problematic emails found
 */
function diagnoseEmailIssues(startDate, endDate) {
  Logger.log('=== EMAIL DIAGNOSIS ===');
  
  const destinationFolder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
  const { processedEmails } = initializeIndex(destinationFolder);
  const existingFilesCache = buildExistingFilesCache(destinationFolder);
  
  const adjustedEndDate = adjustEndDate(endDate);
  const query = 'subject:"Daily Closure Report" has:attachment filename:pdf after:' + startDate + ' before:' + adjustedEndDate;
  const threads = GmailApp.search(query);
  
  const problematicEmails = [];
  
  threads.forEach((thread, threadIndex) => {
    thread.getMessages().forEach(message => {
      const emailId = generateEmailId(message);
      const subject = message.getSubject().trim();
      const match = subject.match(CONFIG.EMAIL_SUBJECT_REGEX);
      
      if (!match) {
        problematicEmails.push({
          type: 'INVALID_FORMAT',
          emailId: emailId,
          subject: subject,
          date: message.getDate()
        });
        return;
      }
      
      const filename = generateFilename(match, 0, 1);
      const inIndex = processedEmails.has(emailId);
      const fileExists = existingFilesCache.has(filename);
      
      if (inIndex && !fileExists) {
        problematicEmails.push({
          type: 'INDEXED_BUT_MISSING_FILE',
          emailId: emailId,
          subject: subject,
          filename: filename,
          date: message.getDate()
        });
      }
    });
  });
  
  Logger.log('=== DIAGNOSIS COMPLETE ===');
  Logger.log('Emails analyzed: ' + threads.reduce((sum, h) => sum + h.getMessageCount(), 0));
  Logger.log('Problematic emails: ' + problematicEmails.length);
  
  problematicEmails.forEach((issue, i) => {
    Logger.log((i + 1) + '. ' + issue.type + ':');
    Logger.log('   Subject: ' + issue.subject);
    Logger.log('   Date: ' + issue.date.toLocaleDateString());
    if (issue.filename) {
      Logger.log('   Expected file: ' + issue.filename);
    }
  });
  
  return problematicEmails;
}

/* ==================== EMAIL PROCESSING UTILITIES ==================== */

/**
 * Initializes or loads the processing index
 * @param {Folder} destinationFolder - Destination folder
 * @returns {Object} Index file and processed emails set
 */
function initializeIndex(destinationFolder) {
  let indexFile = null;
  const files = destinationFolder.getFilesByName(CONFIG.INDEX_FILE_NAME);
  
  if (files.hasNext()) {
    indexFile = files.next();
  } else {
    indexFile = destinationFolder.createFile(CONFIG.INDEX_FILE_NAME, "", MimeType.PLAIN_TEXT);
  }
  
  const indexContent = indexFile.getBlob().getDataAsString();
  const processedEmails = new Set(
    indexContent.split('\n')
      .map(line => line.trim())
      .filter(line => line !== "")
  );
  
  return { indexFile, processedEmails };
}

/**
 * Adjusts end date for Gmail search (exclusive boundary)
 * @param {string} endDate - Original end date
 * @returns {string} Adjusted end date
 */
function adjustEndDate(endDate) {
  const endDateObj = new Date(endDate);
  endDateObj.setDate(endDateObj.getDate() + 1);
  return endDateObj.toISOString().split('T')[0].replace(/-/g, '/');
}

/**
 * Generates unique ID for email tracking
 * @param {GmailMessage} message - Gmail message
 * @returns {string} Unique email ID
 */
function generateEmailId(message) {
  const receiveDate = message.getDate();
  const subject = message.getSubject().trim();
  return receiveDate.getTime() + '_' + subject.replace(/[^\w\s-]/g, '').substring(0, 50);
}

/**
 * Generates standardized filename for attachments
 * @param {Array} match - Regex match results
 * @param {number} index - Attachment index
 * @param {number} totalAttachments - Total attachments
 * @returns {string} Generated filename
 */
function generateFilename(match, index, totalAttachments) {
  const businessName = match[1].trim().replace(/\s+/g, "_");
  const [day, month, year] = match[2].split("/");
  const hour = parseInt(match[3].split(":")[0], 10);
  const shift = hour < CONFIG.SHIFT_CUTOFF_HOUR ? "MORNING" : "EVENING";
  const formattedDate = year + '-' + month + '-' + day;
  const suffix = totalAttachments > 1 ? '_A' + (index + 1) : '';
  
  return businessName + '_' + formattedDate + '_' + shift + suffix + '.pdf';
}

/**
 * Processes email threads and extracts attachments
 * @param {Array} threads - Gmail threads to process
 * @param {Set} processedEmails - Set of already processed emails
 * @param {Set} existingFilesCache - Cache of existing files
 * @param {Folder} destinationFolder - Destination folder
 * @param {File} indexFile - Index file
 * @param {boolean} forceReprocess - Whether to force reprocessing
 * @param {number} startTime - Processing start time
 * @returns {Object} Processing results and statistics
 */
function processEmailThreads(threads, processedEmails, existingFilesCache, destinationFolder, indexFile, forceReprocess, startTime) {
  const newlyProcessed = [];
  let createdFiles = [];
  let stats = {
    emailsFound: 0,
    emailsAlreadyProcessed: 0,
    emailsNewlyProcessed: 0,
    filesCreated: 0,
    filesAlreadyExist: 0,
    errors: 0
  };
  
  try {
    for (let threadIndex = 0; threadIndex < threads.length; threadIndex++) {
      // Execution time check
      if (Date.now() - startTime > CONFIG.MAX_EXECUTION_TIME) {
        Logger.log('Time limit reached, saving progress...');
        break;
      }
      
      const thread = threads[threadIndex];
      Logger.log('Processing thread ' + (threadIndex + 1) + '/' + threads.length);
      
      thread.getMessages().forEach(message => {
        stats.emailsFound++;
        const emailId = generateEmailId(message);
        const subject = message.getSubject().trim();
        
        // Skip if already processed (unless forced)
        if (!forceReprocess && processedEmails.has(emailId)) {
          stats.emailsAlreadyProcessed++;
          return;
        }
        
        Logger.log((forceReprocess ? 'Reprocessing' : 'Processing') + ': ' + subject + ' - ' + message.getDate().toLocaleDateString());
        
        const match = subject.match(CONFIG.EMAIL_SUBJECT_REGEX);
        if (!match) {
          Logger.log('Invalid format: ' + subject);
          stats.errors++;
          return;
        }
        
        try {
          const attachments = message.getAttachments();
          const pdfAttachments = attachments.filter(file => file.getContentType() === MimeType.PDF);
          
          if (pdfAttachments.length === 0) {
            Logger.log('No PDF attachments found');
            return;
          }
          
          let messageFiles = [];
          let anyFileCreated = false;
          
          pdfAttachments.forEach((attachment, index) => {
            const newFilename = generateFilename(match, index, pdfAttachments.length);
            
            if (!forceReprocess && existingFilesCache.has(newFilename)) {
              Logger.log('File already exists: ' + newFilename);
              stats.filesAlreadyExist++;
            } else {
              destinationFolder.createFile(attachment.copyBlob().setName(newFilename));
              messageFiles.push(newFilename);
              anyFileCreated = true;
              Logger.log((forceReprocess ? 'Recreated' : 'Created') + ': ' + newFilename);
            }
          });
          
          // Only mark as processed if files were created
          if (anyFileCreated) {
            newlyProcessed.push(emailId);
            createdFiles.push(...messageFiles);
            stats.emailsNewlyProcessed++;
            stats.filesCreated += messageFiles.length;
            
            // Batch index updates
            if (newlyProcessed.length >= CONFIG.EMAIL_BATCH_SIZE) {
              writeBatchToIndex(indexFile, newlyProcessed, processedEmails);
              newlyProcessed.length = 0;
            }
          } else if (forceReprocess) {
            // In forced mode, mark as processed even if no files created
            newlyProcessed.push(emailId);
            stats.emailsNewlyProcessed++;
          }
        } catch (error) {
          Logger.log('Error processing: ' + error.message);
          stats.errors++;
        }
      });
    }
    
    // Final batch write
    if (newlyProcessed.length > 0) {
      writeBatchToIndex(indexFile, newlyProcessed, processedEmails);
    }
  } catch (error) {
    Logger.log('PROCESSING ERROR: ' + error.message);
    // Save progress on error
    if (newlyProcessed.length > 0) {
      try {
        writeBatchToIndex(indexFile, newlyProcessed, processedEmails);
        Logger.log('Progress saved before error');
      } catch (recoveryError) {
        Logger.log('Save error: ' + recoveryError.message);
      }
    }
    throw error;
  }
  
  return { stats, createdFiles };
}

/**
 * Writes batch of processed emails to index
 * @param {File} indexFile - Index file
 * @param {Array} newlyProcessed - Newly processed email IDs
 * @param {Set} processedEmails - Processed emails set
 */
function writeBatchToIndex(indexFile, newlyProcessed, processedEmails) {
  const currentContent = indexFile.getBlob().getDataAsString();
  const newContent = currentContent + newlyProcessed.join("\n") + "\n";
  indexFile.setContent(newContent);
  newlyProcessed.forEach(id => processedEmails.add(id));
  Logger.log('Batch saved: ' + newlyProcessed.length + ' emails');
}

/**
 * Shows processing summary
 * @param {Object} result - Processing results
 */
function showProcessingSummary(result) {
  const { stats } = result;
  
  Logger.log('=== PROCESSING SUMMARY ===');
  Logger.log('Emails found: ' + stats.emailsFound);
  Logger.log('Already processed: ' + stats.emailsAlreadyProcessed);
  Logger.log('Newly processed: ' + stats.emailsNewlyProcessed);
  Logger.log('Files created: ' + stats.filesCreated);
  Logger.log('Files already existed: ' + stats.filesAlreadyExist);
  Logger.log('Errors: ' + stats.errors);
}

/**
 * Creates empty summary for no results
 * @returns {Object} Empty results structure
 */
function createEmptySummary() {
  return {
    stats: {
      emailsFound: 0,
      emailsAlreadyProcessed: 0,
      emailsNewlyProcessed: 0,
      filesCreated: 0,
      filesAlreadyExist: 0,
      errors: 0
    },
    createdFiles: []
  };
}

/* ==================== FILE CACHE SYSTEM ==================== */

/**
 * Builds cache of existing files for duplicate detection
 * @param {Folder} destinationFolder - Destination folder
 * @returns {Set} Set of existing filenames
 */
function buildExistingFilesCache(destinationFolder) {
  const cache = new Set();
  
  // Root folder files
  const rootFiles = destinationFolder.getFilesByType(MimeType.PDF);
  while (rootFiles.hasNext()) {
    cache.add(rootFiles.next().getName());
  }
  
  // Subfolder files (date-based organization)
  const subfolders = destinationFolder.getFolders();
  let scannedSubfolders = 0;
  
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const subfolderName = subfolder.getName();
    
    // Only scan date-formatted folders (yyyy-mm-dd)
    if (/^\d{4}-\d{2}-\d{2}$/.test(subfolderName)) {
      const subfolderFiles = subfolder.getFilesByType(MimeType.PDF);
      let filesInSubfolder = 0;
      
      while (subfolderFiles.hasNext()) {
        cache.add(subfolderFiles.next().getName());
        filesInSubfolder++;
      }
      
      scannedSubfolders++;
      if (filesInSubfolder > 0) {
        Logger.log(subfolderName + ': ' + filesInSubfolder + ' files');
      }
    }
  }
  
  Logger.log('Subfolders scanned: ' + scannedSubfolders);
  return cache;
}

// Export functions for testing and external use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    processEmails,
    processMyDates,
    reprocessDateEmails,
    diagnoseEmailIssues,
    initializeIndex,
    generateEmailId,
    generateFilename,
    buildExistingFilesCache
  };
}
