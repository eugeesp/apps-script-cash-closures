/**
 * CashFlow Automator - Core Processing Engine
 * Main system for PDF processing, data extraction, and batch management
 * @version 2.1.0
 */

/* ==================== SYSTEM CONFIGURATION ==================== */
const CONFIG = {
  SHEET_NAME: 'Control 2025',
  MAIN_FOLDER: 'CIERRES DE CAJA',
  BATCH_SIZE: 18,
  DELAY_SECONDS: 30,
  MAX_RETRIES: 3,
  DESTINATION_FOLDER_ID: "your_drive_folder_id_here",
  INDEX_FILE_NAME: "index.doc",
  EMAIL_BATCH_SIZE: 8,
  SHIFT_CUTOFF_HOUR: 16,
  EMAIL_SUBJECT_REGEX: /comercio\s+(.*?)\s+-\s+Reporte de Cierre de Caja\s+-\s+(\d{2}\/\d{2}\/\d{4})\s+-\s+(\d{2}:\d{2}:\d{2})/,
  MAX_EXECUTION_TIME: 5 * 60 * 1000
};

/* ==================== MAIN PROCESSING FUNCTIONS ==================== */

/**
 * Starts the automated processing system
 * Initializes properties and begins batch processing
 */
function iniciarProcesamiento() {
  limpiarTriggers();
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'procesamiento_activo': 'true',
    'tanda_actual': '1',
    'archivos_procesados': '0',
    'intentos_fallidos': '0',
    'inicio': new Date().toISOString()
  });
  
  const ahora = new Date().toLocaleString('es-AR');
  Logger.log('=== PROCESSING STARTED ===');
  Logger.log('Date/Time: ' + ahora);
  Logger.log('Folder: ' + CONFIG.CARPETA_PRINCIPAL);
  Logger.log('Batch size: ' + CONFIG.TANDA_SIZE + ' files, ' + CONFIG.DELAY_SEGUNDOS + 's delay');
  
  procesarSiguienteTanda();
}

/**
 * Processes the next batch of files
 * Handles batch management and progress tracking
 */
function procesarSiguienteTanda() {
  const props = PropertiesService.getScriptProperties();
  
  if (props.getProperty('procesamiento_activo') !== 'true') {
    Logger.log('Processing paused');
    return;
  }
  
  const tandaActual = parseInt(props.getProperty('tanda_actual') || '1');
  const totalProcesados = parseInt(props.getProperty('archivos_procesados') || '0');
  
  Logger.log('=== BATCH ' + tandaActual + ' ===');
  Logger.log('Total processed: ' + totalProcesados + ' files');
  
  try {
    // Get pending files
    const carpetaRaiz = DriveApp.getFoldersByName(CONFIG.CARPETA_PRINCIPAL).next();
    const pdfs = carpetaRaiz.getFilesByType(MimeType.PDF);
    const archivosPendientes = [];
    
    while (pdfs.hasNext()) {
      archivosPendientes.push(pdfs.next());
    }
    
    Logger.log('Pending files in root folder: ' + archivosPendientes.length);
    
    if (archivosPendientes.length === 0) {
      Logger.log('=== PROCESSING COMPLETED ===');
      Logger.log('No more files to process');
      finalizarProcesamiento('COMPLETED');
      return;
    }
    
    // Process current batch
    const tandaArchivos = archivosPendientes.slice(0, CONFIG.TANDA_SIZE);
    Logger.log('Processing ' + tandaArchivos.length + ' files in this batch...');
    
    const resultados = procesarArchivos(tandaArchivos);
    const exitosos = resultados.filter(r => !r.error).length;
    const fallidos = resultados.length - exitosos;
    
    // Update progress
    props.setProperties({
      'tanda_actual': (tandaActual + 1).toString(),
      'archivos_procesados': (totalProcesados + exitosos).toString(),
      'intentos_fallidos': '0'
    });
    
    Logger.log('=== BATCH ' + tandaActual + ' RESULTS ===');
    Logger.log('Successful: ' + exitosos);
    Logger.log('Failed: ' + fallidos);
    Logger.log('Total accumulated: ' + (totalProcesados + exitosos));
    Logger.log('Remaining: ' + (archivosPendientes.length - tandaArchivos.length));
    
    // Continue if there are more files
    if (archivosPendientes.length > tandaArchivos.length) {
      Logger.log('Next batch in ' + CONFIG.DELAY_SEGUNDOS + ' seconds...');
      programarSiguienteTanda();
    } else {
      Logger.log('All batches completed');
      finalizarProcesamiento('COMPLETED');
    }
  } catch (error) {
    Logger.log('Batch ' + tandaActual + ' error: ' + error.message);
    manejarError();
  }
}

/**
 * Processes individual PDF files and extracts financial data
 * @param {Array} archivos - Array of PDF files to process
 * @returns {Array} Processing results with extracted data or errors
 */
function procesarArchivos(archivos) {
  const rows = [];
  const archivosPorFecha = new Map();
  const fechasEncontradas = new Set();
  
  Logger.log('Processing ' + archivos.length + ' files...');
  
  archivos.forEach((pdf, i) => {
    let docId = null;
    
    try {
      // Progress logging
      if (i % 5 === 0 || i === archivos.length - 1) {
        Logger.log('Progress: ' + (i + 1) + '/' + archivos.length + ' - ' + pdf.getName());
      }
      
      // Convert PDF to text
      const meta = Drive.Files.copy(
        { title: pdf.getName() + ' (tmp)', mimeType: MimeType.GOOGLE_DOCS },
        pdf.getId()
      );
      docId = meta.id;
      Utilities.sleep(2000);
      
      const url = 'https://docs.google.com/document/d/' + docId + '/export?format=txt';
      const resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
      });
      
      if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
      
      const txt = resp.getContentText();
      if (!txt.trim()) throw new Error('PDF without extractable text');
      
      // Extract data
      const datos = extraerDatosPDF(txt, pdf.getName());
      
      if (datos.fechaCierre) {
        rows.push(datos);
        fechasEncontradas.add(datos.fechaCierre);
        
        // Group by date
        const fechaISO = normalizarFecha(datos.fechaCierre);
        if (fechaISO) {
          if (!archivosPorFecha.has(fechaISO)) {
            archivosPorFecha.set(fechaISO, []);
          }
          archivosPorFecha.get(fechaISO).push(pdf);
        }
        
        // Log important data
        if (i < 3 || i === archivos.length - 1) {
          Logger.log('Processed: ' + pdf.getName() + ' → ' + datos.fechaCierre + ' (' + datos.sucursal + ') ' + datos.turno);
        }
      } else {
        throw new Error('Could not extract date');
      }
    } catch (e) {
      Logger.log('Failed: ' + pdf.getName() + ' → Error: ' + e.message);
      rows.push({ archivo: pdf.getName(), error: e.message });
    } finally {
      if (docId) {
        try {
          DriveApp.getFileById(docId).setTrashed(true);
        } catch (_) {}
      }
    }
  });
  
  // Show processed dates summary
  if (fechasEncontradas.size > 0) {
    const fechasArray = Array.from(fechasEncontradas).sort();
    Logger.log('Dates processed in this batch: ' + fechasArray.join(', '));
  }
  
  // Update sheet and organize files
  const rowsExitosos = rows.filter(r => !r.error);
  if (rowsExitosos.length > 0) {
    Logger.log('Updating ' + rowsExitosos.length + ' rows in spreadsheet...');
    actualizarHoja(rowsExitosos);
    
    Logger.log('Organizing files into ' + archivosPorFecha.size + ' date folders...');
    const carpetaRaiz = DriveApp.getFoldersByName(CONFIG.CARPETA_PRINCIPAL).next();
    organizarArchivos(carpetaRaiz, archivosPorFecha);
  }
  
  return rows;
}

/* ==================== DATA EXTRACTION FUNCTIONS ==================== */

/**
 * Extracts financial data from PDF text content
 * @param {string} texto - Text content extracted from PDF
 * @param {string} archivo - Original filename for reference
 * @returns {Object} Structured financial data
 */
function extraerDatosPDF(texto, archivo) {
  const lineaRazon = texto.split('\n')
    .find(l => l.toLowerCase().includes('razon social:') && l.toLowerCase().includes('cafe de barrio')) || '';
  
  const sucursal = (lineaRazon.match(/CAFE DE BARRIO\s*[-\s]*(.+)/i) || [])[1]?.trim() || '';
  
  const fh = texto.match(/Fecha de cierre:\s*(\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2}:\d{2})/);
  const fechaCierre = fh ? fh[1] : '';
  const horaCierre = fh ? fh[2] : '';
  const turno = horaCierre && parseInt(horaCierre.split(':')[0], 10) < 16 ? 'Mañana' : 'Tarde';
  
  return {
    archivo,
    fechaCierre,
    horaCierre,
    turno,
    sucursal,
    efectivoApertura: capturarImporte(texto, /Efectivo en caja apertura:/i),
    totalVentas: capturarImporte(texto, /Total de Ventas:/i),
    efectivo: capturarImporte(texto, /(?:^|\n)\s*Efectivo:/i),
    tarjetas: capturarImporte(texto, /Tarjetas:/i),
    qr: capturarImporte(texto, /QR:/i),
    efectivoCierre: capturarImporte(texto, /Efectivo en (?:caja cierre|cierre de caja):/i),
    retiroCierre: capturarRetiroCierre(texto)
  };
}

/**
 * Extracts monetary amounts from text using regex patterns
 * @param {string} texto - Text to search
 * @param {RegExp} labelRegex - Pattern to identify amount labels
 * @returns {string} Extracted amount or empty string
 */
function capturarImporte(texto, labelRegex) {
  const pattern = new RegExp(labelRegex.source + '\\s*\\$?\\s*([0-9]{1,3}(?:\\.[0-9]{3})*,[0-9]{2})', 'i');
  const m = texto.match(pattern);
  return m ? m[1] : '';
}

/**
 * Extracts cash withdrawal amounts from closure reports
 * @param {string} texto - Text to search
 * @returns {string} Extracted withdrawal amount or empty string
 */
function capturarRetiroCierre(texto) {
  const lineas = texto.split('\n');
  const idx = lineas.findIndex(l => /Retiro\s+por\s+Cierre\s*-?/i.test(l));
  
  if (idx === -1) return '';
  
  for (let i = idx; i < Math.min(idx + 3, lineas.length); i++) {
    const m = lineas[i].match(/-?\$?\s*([0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2})/);
    if (m) return m[1];
  }
  
  return '';
}

/* ==================== UTILITY FUNCTIONS ==================== */

/**
 * Normalizes date formats to consistent ISO format
 * @param {*} valor - Date value to normalize
 * @returns {string} Normalized date in YYYY-MM-DD format
 */
function normalizarFecha(valor) {
  if (!valor) return '';
  
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  
  const m = valor.toString().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return m ? m[3] + '-' + m[2] + '-' + m[1] : valor.toString().trim();
}

/**
 * Converts Argentine number format to standard numeric format
 * @param {string} str - Number string in Argentine format
 * @returns {string} Converted number or empty string
 */
function convertirNumeroArgentino(str) {
  if (!str) return '';
  
  let s = str.replace(/[^0-9.,-]/g, '');
  
  if (s.includes('.') && s.includes(',')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  } else if ((s.match(/\./g) || []).length > 1) {
    s = s.replace(/\./g, '');
  }
  
  const n = parseFloat(s);
  return isNaN(n) ? '' : Math.abs(n);
}

/**
 * Schedules the next processing batch
 */
function programarSiguienteTanda() {
  limpiarTriggers();
  ScriptApp.newTrigger('procesarSiguienteTanda')
    .timeBased()
    .after(CONFIG.DELAY_SEGUNDOS * 1000)
    .create();
  Logger.log('Next batch in ' + CONFIG.DELAY_SEGUNDOS + 's');
}

/**
 * Clears all existing processing triggers
 */
function limpiarTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'procesarSiguienteTanda')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/**
 * Handles processing errors with retry logic
 */
function manejarError() {
  const props = PropertiesService.getScriptProperties();
  const intentos = parseInt(props.getProperty('intentos_fallidos') || '0') + 1;
  
  if (intentos >= CONFIG.MAX_INTENTOS) {
    finalizarProcesamiento('ERROR_MULTIPLE');
  } else {
    props.setProperty('intentos_fallidos', intentos.toString());
    Logger.log('Retry ' + intentos + '/' + CONFIG.MAX_INTENTOS);
    programarSiguienteTanda();
  }
}

/**
 * Finalizes processing and cleans up resources
 * @param {string} motivo - Reason for finishing
 */
function finalizarProcesamiento(motivo) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('procesamiento_activo', 'false');
  limpiarTriggers();
  
  const procesados = props.getProperty('archivos_procesados') || '0';
  Logger.log('Finished: ' + motivo + ' - ' + procesados + ' files processed');
}

// Export functions for testing and external use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    iniciarProcesamiento,
    procesarSiguienteTanda,
    procesarArchivos,
    extraerDatosPDF,
    capturarImporte,
    capturarRetiroCierre,
    normalizarFecha,
    convertirNumeroArgentino,
    CONFIG
  };
}
