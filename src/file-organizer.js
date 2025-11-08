/**
 * CashFlow Automator - File Organization System
 * Handles Drive file organization, duplicate management, and folder structure
 * @version 2.1.0
 */

/* ==================== FILE ORGANIZATION FUNCTIONS ==================== */

/**
 * Organizes files into date-based folder structure
 * @param {Folder} rootFolder - Root folder to organize from
 * @param {Map} filesByDate - Map of dates to file arrays
 */
function organizeFiles(rootFolder, filesByDate) {
  let totalMoved = 0;
  
  filesByDate.forEach((files, dateISO) => {
    try {
      let dateFolder = null;
      const existingFolders = rootFolder.getFoldersByName(dateISO);
      
      if (existingFolders.hasNext()) {
        dateFolder = existingFolders.next();
        Logger.log('Using existing folder: ' + dateISO);
      } else {
        dateFolder = rootFolder.createFolder(dateISO);
        Logger.log('Created new folder: ' + dateISO);
      }
      
      // Move files to date folder
      files.forEach(file => {
        try {
          file.moveTo(dateFolder);
          totalMoved++;
        } catch (moveError) {
          Logger.log('Error moving file ' + file.getName() + ': ' + moveError.message);
        }
      });
      
      Logger.log('Moved ' + files.length + ' files to ' + dateISO);
    } catch (error) {
      Logger.log('Error organizing ' + dateISO + ': ' + error.message);
    }
  });
  
  Logger.log('Total files organized: ' + totalMoved);
}

/**
 * Processes a specific date folder (manual operation)
 * @param {string} dateISO - Date in YYYY-MM-DD format
 */
function processDateFolder(dateISO) {
  if (!dateISO) {
    Logger.log('Usage: processDateFolder("2025-07-22")');
    return;
  }
  
  const startTime = new Date();
  Logger.log('=== MANUAL FOLDER PROCESSING ===');
  Logger.log('Target date: ' + dateISO);
  Logger.log('Started: ' + startTime.toLocaleString('es-AR'));
  
  try {
    const rootFolder = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER).next();
    const dateFolder = rootFolder.getFoldersByName(dateISO).next();
    const files = [];
    const pdfs = dateFolder.getFilesByType(MimeType.PDF);
    
    while (pdfs.hasNext()) files.push(pdfs.next());
    
    Logger.log('Files found: ' + files.length);
    
    if (files.length === 0) {
      Logger.log('Folder empty - no PDFs to process');
      return;
    }
    
    // Show sample file names
    if (files.length <= 3) {
      files.forEach(pdf => Logger.log(' • ' + pdf.getName()));
    } else {
      Logger.log(' • ' + files[0].getName());
      Logger.log(' • ' + files[1].getName());
      Logger.log(' • ... (and ' + (files.length - 2) + ' more)');
    }
    
    const results = processFiles(files);
    const successful = results.filter(r => !r.error).length;
    const failed = results.length - successful;
    const totalSeconds = Math.round((new Date() - startTime) / 1000);
    
    Logger.log('=== PROCESSING ' + dateISO + ' COMPLETED ===');
    Logger.log('Successful: ' + successful);
    Logger.log('Failed: ' + failed);
    Logger.log('Total time: ' + totalSeconds + ' seconds');
    
    if (failed > 0) {
      Logger.log('Files with errors:');
      results.filter(r => r.error).forEach(r => {
        Logger.log(' • ' + r.file + ': ' + r.error);
      });
    }
  } catch (error) {
    Logger.log('Error accessing folder ' + dateISO + ': ' + error.message);
  }
}

/**
 * Shows available date folders for processing
 */
function showAvailableFolders() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const root = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER).next();
    const subfolders = root.getFolders();
    const dates = [];
    
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      const name = folder.getName();
      
      if (/^\d{4}-\d{2}-\d{2}$/.test(name)) {
        // Count PDF files in folder
        const pdfs = folder.getFilesByType(MimeType.PDF);
        let count = 0;
        while (pdfs.hasNext()) {
          pdfs.next();
          count++;
        }
        dates.push(name + ' (' + count + ' PDFs)');
      }
    }
    
    if (dates.length === 0) {
      ui.alert('No Date Folders', 'No subfolders with date format found.', ui.ButtonSet.OK);
      return;
    }
    
    dates.sort().reverse();
    const list = dates.join('\n• ');
    
    ui.alert(
      'Available Date Folders',
      'Date folders found:\n\n• ' + list + '\n\nUse "Process Specific Date" to select one.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Access Error', 'Could not access folders: ' + error.message, ui.ButtonSet.OK);
  }
}

/* ==================== DUPLICATE MANAGEMENT ==================== */

/**
 * Recursively removes duplicate files by name
 */
function removeDuplicateFilesRecursive() {
  const root = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER);
  if (!root.hasNext()) {
    Logger.log('Root folder not found: ' + CONFIG.MAIN_FOLDER);
    return;
  }
  
  const rootFolder = root.next();
  Logger.log('Scanning file hierarchy from: ' + rootFolder.getName());
  
  const allFiles = getAllFilesRecursive(rootFolder);
  Logger.log('Total files found: ' + allFiles.length);
  
  const fileMap = new Map();
  let totalRemoved = 0;
  
  allFiles.forEach(file => {
    const fileName = file.getName();
    if (!fileMap.has(fileName)) {
      fileMap.set(fileName, [file]);
    } else {
      fileMap.get(fileName).push(file);
    }
  });
  
  for (const [fileName, fileList] of fileMap.entries()) {
    if (fileList.length > 1) {
      Logger.log('Removing ' + (fileList.length - 1) + ' copies of "' + fileName + '"');
      for (let i = 1; i < fileList.length; i++) {
        try {
          fileList[i].setTrashed(true);
          Logger.log(' • Removed: ' + fileList[i].getName());
          totalRemoved++;
        } catch (error) {
          Logger.log('Error removing "' + fileName + '": ' + error.message);
        }
      }
    }
  }
  
  Logger.log('Total files removed: ' + totalRemoved);
}

/**
 * Counts duplicate files without removing them
 */
function countDuplicateFilesRecursive() {
  const root = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER);
  if (!root.hasNext()) {
    Logger.log('Root folder not found: ' + CONFIG.MAIN_FOLDER);
    return;
  }
  
  const rootFolder = root.next();
  Logger.log('Scanning file hierarchy from: ' + rootFolder.getName());
  
  const allFiles = getAllFilesRecursive(rootFolder);
  Logger.log('Total files found: ' + allFiles.length);
  
  const fileMap = new Map();
  let duplicateNames = 0;
  let totalDuplicateFiles = 0;
  
  allFiles.forEach(file => {
    const fileName = file.getName();
    fileMap.set(fileName, (fileMap.get(fileName) || 0) + 1);
  });
  
  for (const [fileName, count] of fileMap.entries()) {
    if (count > 1) {
      Logger.log('"' + fileName + '" → ' + count + ' copies');
      duplicateNames++;
      totalDuplicateFiles += (count - 1);
    }
  }
  
  Logger.log('Duplicate file names detected: ' + duplicateNames);
  Logger.log('Total files that would be removed (keeping one each): ' + totalDuplicateFiles);
}

/**
 * Recursively gets all files from a folder and its subfolders
 * @param {Folder} folder - Starting folder
 * @returns {Array} All files found
 */
function getAllFilesRecursive(folder) {
  const files = [];
  
  // Files in current folder
  const fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    files.push(fileIterator.next());
  }
  
  // Recursively process subfolders
  const subfolderIterator = folder.getFolders();
  while (subfolderIterator.hasNext()) {
    const subfolder = subfolderIterator.next();
    files.push(...getAllFilesRecursive(subfolder));
  }
  
  return files;
}

/* ==================== FOLDER MANAGEMENT ==================== */

/**
 * Creates the main folder structure if it doesn't exist
 */
function initializeFolderStructure() {
  try {
    const rootFolders = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER);
    if (!rootFolders.hasNext()) {
      const rootFolder = DriveApp.createFolder(CONFIG.MAIN_FOLDER);
      Logger.log('Created main folder: ' + CONFIG.MAIN_FOLDER);
      return rootFolder;
    } else {
      Logger.log('Main folder already exists: ' + CONFIG.MAIN_FOLDER);
      return rootFolders.next();
    }
  } catch (error) {
    Logger.log('Error initializing folder structure: ' + error.message);
    return null;
  }
}

/**
 * Gets folder statistics (file counts, sizes, etc.)
 * @returns {Object} Folder statistics
 */
function getFolderStats() {
  try {
    const root = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER);
    if (!root.hasNext()) {
      return { error: 'Main folder not found' };
    }
  
    const rootFolder = root.next();
    const stats = {
      totalFiles: 0,
      totalPDFs: 0,
      dateFolders: 0,
      filesByDate: {},
      totalSize: 0
    };
    
    // Count files in root
    const rootFiles = rootFolder.getFiles();
    while (rootFiles.hasNext()) {
      const file = rootFiles.next();
      stats.totalFiles++;
      stats.totalSize += file.getSize();
      if (file.getMimeType() === MimeType.PDF) {
        stats.totalPDFs++;
      }
    }
    
    // Count files in date folders
    const subfolders = rootFolder.getFolders();
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      const folderName = folder.getName();
      
      if (/^\d{4}-\d{2}-\d{2}$/.test(folderName)) {
        stats.dateFolders++;
        
        let folderFileCount = 0;
        let folderPDFCount = 0;
        let folderSize = 0;
        
        const folderFiles = folder.getFiles();
        while (folderFiles.hasNext()) {
          const file = folderFiles.next();
          folderFileCount++;
          folderSize += file.getSize();
          if (file.getMimeType() === MimeType.PDF) {
            folderPDFCount++;
          }
        }
        
        stats.filesByDate[folderName] = {
          totalFiles: folderFileCount,
          pdfFiles: folderPDFCount,
          totalSize: folderSize
        };
        
        stats.totalFiles += folderFileCount;
        stats.totalPDFs += folderPDFCount;
        stats.totalSize += folderSize;
      }
    }
    
    // Convert size to MB
    stats.totalSizeMB = (stats.totalSize / (1024 * 1024)).toFixed(2);
    
    return stats;
  } catch (error) {
    Logger.log('Error getting folder stats: ' + error.message);
    return { error: error.message };
  }
}

/**
 * Logs folder statistics for monitoring
 */
function logFolderStats() {
  const stats = getFolderStats();
  
  if (stats.error) {
    Logger.log('Error getting folder stats: ' + stats.error);
    return;
  }
  
  Logger.log('=== FOLDER STATISTICS ===');
  Logger.log('Total files: ' + stats.totalFiles);
  Logger.log('PDF files: ' + stats.totalPDFs);
  Logger.log('Date folders: ' + stats.dateFolders);
  Logger.log('Total size: ' + stats.totalSizeMB + ' MB');
  
  if (stats.dateFolders > 0) {
    Logger.log('Files by date folder:');
    Object.keys(stats.filesByDate).sort().forEach(date => {
      const folderStats = stats.filesByDate[date];
      Logger.log(' • ' + date + ': ' + folderStats.pdfFiles + ' PDFs, ' + 
                folderStats.totalFiles + ' total files');
    });
  }
}

/**
 * Cleans up empty date folders
 * @returns {number} Number of folders removed
 */
function cleanupEmptyFolders() {
  try {
    const root = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER);
    if (!root.hasNext()) return 0;
    
    const rootFolder = root.next();
    const subfolders = rootFolder.getFolders();
    let removedCount = 0;
    
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      const folderName = folder.getName();
      
      // Only process date-formatted folders
      if (/^\d{4}-\d{2}-\d{2}$/.test(folderName)) {
        const files = folder.getFiles();
        const hasFiles = files.hasNext();
        
        if (!hasFiles) {
          try {
            folder.setTrashed(true);
            Logger.log('Removed empty folder: ' + folderName);
            removedCount++;
          } catch (error) {
            Logger.log('Error removing folder ' + folderName + ': ' + error.message);
          }
        }
      }
    }
    
    Logger.log('Removed ' + removedCount + ' empty folders');
    return removedCount;
  } catch (error) {
    Logger.log('Error cleaning up folders: ' + error.message);
    return 0;
  }
}

/**
 * Moves all PDF files from root to appropriate date folders
 * @returns {number} Number of files moved
 */
function organizeAllFiles() {
  try {
    const root = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER);
    if (!root.hasNext()) {
      Logger.log('Main folder not found');
      return 0;
    }
    
    const rootFolder = root.next();
    const pdfs = rootFolder.getFilesByType(MimeType.PDF);
    const filesByDate = new Map();
    let processedCount = 0;
    
    // Group files by date from their names
    while (pdfs.hasNext()) {
      const file = pdfs.next();
      const fileName = file.getName();
      
      // Extract date from filename (assuming format contains YYYY-MM-DD)
      const dateMatch = fileName.match(/(\d{4}-\d{2}-\d{2})/);
      if (dateMatch) {
        const date = dateMatch[1];
        if (!filesByDate.has(date)) {
          filesByDate.set(date, []);
        }
        filesByDate.get(date).push(file);
        processedCount++;
      } else {
        Logger.log('Could not extract date from: ' + fileName);
      }
    }
    
    if (filesByDate.size > 0) {
      organizeFiles(rootFolder, filesByDate);
    }
    
    return processedCount;
  } catch (error) {
    Logger.log('Error organizing all files: ' + error.message);
    return 0;
  }
}

// Export functions for testing and external use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    organizeFiles,
    processDateFolder,
    showAvailableFolders,
    removeDuplicateFilesRecursive,
    countDuplicateFilesRecursive,
    getAllFilesRecursive,
    initializeFolderStructure,
    getFolderStats,
    logFolderStats,
    cleanupEmptyFolders,
    organizeAllFiles
  };
}
