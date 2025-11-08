/**
 * CashFlow Automator - Email Processing System
 * Handles Gmail integration, attachment processing, and file organization
 * @version 2.1.0
 */

/* ==================== EMAIL PROCESSING FUNCTIONS ==================== */

/**
 * Main function to process emails within a date range
 * @param {string} fechaDesde - Start date (YYYY/MM/DD)
 * @param {string} fechaHasta - End date (YYYY/MM/DD)
 * @param {boolean} forzarReproceso - Whether to force reprocessing
 * @returns {Object} Processing results and statistics
 */
function procesarCorreos(fechaDesde, fechaHasta, forzarReproceso = false) {
  const tiempoInicio = Date.now();
  Logger.log('Processing emails from ' + fechaDesde + ' to ' + fechaHasta + (forzarReproceso ? ' (FORCED)' : ''));
  
  const carpetaDestino = DriveApp.getFolderById(CONFIG.CARPETA_DESTINO_ID);
  
  // Initialize processing index
  const { archivoIndex, correosYaProcesados } = inicializarIndice(carpetaDestino);
  Logger.log('Emails already in index: ' + correosYaProcesados.size);
  
  // Build cache of existing files
  const archivosExistentesCache = construirCacheArchivosExistentes(carpetaDestino);
  Logger.log('Files in cache: ' + archivosExistentesCache.size);
  
  // Prepare search query
  const fechaHastaAjustada = ajustarFechaHasta(fechaHasta);
  const query = 'subject:"Reporte de Cierre de Caja" has:attachment filename:pdf after:' + fechaDesde + ' before:' + fechaHastaAjustada;
  Logger.log('Search query: ' + query);
  
  const hilos = GmailApp.search(query);
  Logger.log('Threads found: ' + hilos.length);
  
  if (hilos.length === 0) {
    Logger.log('No emails found in this date range');
    return crearResumenVacio();
  }
  
  // Process email threads
  const resultado = procesarHilos(hilos, correosYaProcesados, archivosExistentesCache, carpetaDestino, archivoIndex, forzarReproceso, tiempoInicio);
  
  // Final summary
  mostrarResumen(resultado);
  return resultado;
}

/**
 * Processes specific date emails (convenience function)
 * @returns {Object} Processing results
 */
function procesarMisFechas() {
  // Configure date range for processing
  const fechaDesde = "2025/07/05";
  const fechaHasta = "2025/07/05";
  
  return procesarCorreos(fechaDesde, fechaHasta);
}

/**
 * Reprocesses emails for a specific date
 * @param {string} fecha - Date to reprocess (YYYY/MM/DD)
 * @param {boolean} eliminarDelIndice - Whether to remove from index first
 * @returns {Object} Reprocessing results
 */
function reprocesarCorreosFecha(fecha, eliminarDelIndice = true) {
  Logger.log('=== REPROCESSING DATE: ' + fecha + ' ===');
  
  if (eliminarDelIndice) {
    const carpetaDestino = DriveApp.getFolderById(CONFIG.CARPETA_DESTINO_ID);
    const { archivoIndex } = inicializarIndice(carpetaDestino);
    
    // Remove existing entries for this date
    let contenido = archivoIndex.getBlob().getDataAsString();
    const lineasOriginales = contenido.split('\n').length - 1;
    
    const fechaFormateada = fecha.replace(/\//g, '-');
    const lineasFiltradas = contenido.split('\n')
      .filter(linea => !linea.includes(fechaFormateada))
      .filter(linea => linea.trim() !== '');
    
    // Rewrite index file
    archivoIndex.setContent(lineasFiltradas.join('\n') + '\n');
    const lineasEliminadas = lineasOriginales - lineasFiltradas.length;
    Logger.log('Removed ' + lineasEliminadas + ' entries from index');
  }
  
  // Process normally
  return procesarCorreos(fecha, fecha, true);
}

/**
 * Diagnoses email processing issues
 * @param {string} fechaDesde - Start date
 * @param {string} fechaHasta - End date
 * @returns {Array} Problematic emails found
 */
function diagnosticarCorreosProblematicos(fechaDesde, fechaHasta) {
  Logger.log('=== EMAIL DIAGNOSIS ===');
  
  const carpetaDestino = DriveApp.getFolderById(CONFIG.CARPETA_DESTINO_ID);
  const { correosYaProcesados } = inicializarIndice(carpetaDestino);
  const archivosExistentesCache = construirCacheArchivosExistentes(carpetaDestino);
  
  const fechaHastaAjustada = ajustarFechaHasta(fechaHasta);
  const query = 'subject:"Reporte de Cierre de Caja" has:attachment filename:pdf after:' + fechaDesde + ' before:' + fechaHastaAjustada;
  const hilos = GmailApp.search(query);
  
  const problematicos = [];
  
  hilos.forEach((hilo, hiloIndex) => {
    hilo.getMessages().forEach(mensaje => {
      const correoId = generarIdCorreo(mensaje);
      const asunto = mensaje.getSubject().trim();
      const match = asunto.match(CONFIG.REGEX_ASUNTO);
      
      if (!match) {
        problematicos.push({
          tipo: 'INVALID_FORMAT',
          correoId: correoId,
          asunto: asunto,
          fecha: mensaje.getDate()
        });
        return;
      }
      
      const nombreArchivo = generarNombreArchivo(match, 0, 1);
      const enIndice = correosYaProcesados.has(correoId);
      const archivoExiste = archivosExistentesCache.has(nombreArchivo);
      
      if (enIndice && !archivoExiste) {
        problematicos.push({
          tipo: 'INDEXED_BUT_MISSING_FILE',
          correoId: correoId,
          asunto: asunto,
          nombreArchivo: nombreArchivo,
          fecha: mensaje.getDate()
        });
      }
    });
  });
  
  Logger.log('=== DIAGNOSIS COMPLETE ===');
  Logger.log('Emails analyzed: ' + hilos.reduce((sum, h) => sum + h.getMessageCount(), 0));
  Logger.log('Problematic emails: ' + problematicos.length);
  
  problematicos.forEach((problema, i) => {
    Logger.log((i + 1) + '. ' + problema.tipo + ':');
    Logger.log('   Subject: ' + problema.asunto);
    Logger.log('   Date: ' + problema.fecha.toLocaleDateString());
    if (problema.nombreArchivo) {
      Logger.log('   Expected file: ' + problema.nombreArchivo);
    }
  });
  
  return problematicos;
}

/* ==================== EMAIL PROCESSING UTILITIES ==================== */

/**
 * Initializes or loads the processing index
 * @param {Folder} carpetaDestino - Destination folder
 * @returns {Object} Index file and processed emails set
 */
function inicializarIndice(carpetaDestino) {
  let archivoIndex = null;
  const archivos = carpetaDestino.getFilesByName(CONFIG.NOMBRE_INDEX);
  
  if (archivos.hasNext()) {
    archivoIndex = archivos.next();
  } else {
    archivoIndex = carpetaDestino.createFile(CONFIG.NOMBRE_INDEX, "", MimeType.PLAIN_TEXT);
  }
  
  const contenidoIndex = archivoIndex.getBlob().getDataAsString();
  const correosYaProcesados = new Set(
    contenidoIndex.split('\n')
      .map(l => l.trim())
      .filter(l => l !== "")
  );
  
  return { archivoIndex, correosYaProcesados };
}

/**
 * Adjusts end date for Gmail search (exclusive boundary)
 * @param {string} fechaHasta - Original end date
 * @returns {string} Adjusted end date
 */
function ajustarFechaHasta(fechaHasta) {
  const fechaHastaObj = new Date(fechaHasta);
  fechaHastaObj.setDate(fechaHastaObj.getDate() + 1);
  return fechaHastaObj.toISOString().split('T')[0].replace(/-/g, '/');
}

/**
 * Generates unique ID for email tracking
 * @param {GmailMessage} mensaje - Gmail message
 * @returns {string} Unique email ID
 */
function generarIdCorreo(mensaje) {
  const fechaRecepcion = mensaje.getDate();
  const asunto = mensaje.getSubject().trim();
  return fechaRecepcion.getTime() + '_' + asunto.replace(/[^\w\s-]/g, '').substring(0, 50);
}

/**
 * Generates standardized filename for attachments
 * @param {Array} match - Regex match results
 * @param {number} index - Attachment index
 * @param {number} totalAdjuntos - Total attachments
 * @returns {string} Generated filename
 */
function generarNombreArchivo(match, index, totalAdjuntos) {
  const razon = match[1].trim().replace(/\s+/g, "_");
  const [dia, mes, anio] = match[2].split("/");
  const hora = parseInt(match[3].split(":")[0], 10);
  const turno = hora < CONFIG.HORA_CORTE_TURNO ? "MANANA" : "TARDE";
  const fechaFormateada = anio + '-' + mes + '-' + dia;
  const sufijo = totalAdjuntos > 1 ? '_A' + (index + 1) : '';
  
  return razon + '_' + fechaFormateada + '_' + turno + sufijo + '.pdf';
}

/**
 * Processes email threads and extracts attachments
 * @param {Array} hilos - Gmail threads to process
 * @param {Set} correosYaProcesados - Set of already processed emails
 * @param {Set} archivosExistentesCache - Cache of existing files
 * @param {Folder} carpetaDestino - Destination folder
 * @param {File} archivoIndex - Index file
 * @param {boolean} forzarReproceso - Whether to force reprocessing
 * @param {number} tiempoInicio - Processing start time
 * @returns {Object} Processing results and statistics
 */
function procesarHilos(hilos, correosYaProcesados, archivosExistentesCache, carpetaDestino, archivoIndex, forzarReproceso, tiempoInicio) {
  const nuevosProcesados = [];
  let archivosCreados = [];
  let estadisticas = {
    correosEncontrados: 0,
    correosYaProcesados: 0,
    correosNuevosProcesados: 0,
    archivosCreados: 0,
    archivosYaExistentes: 0,
    errores: 0
  };
  
  try {
    for (let hiloIndex = 0; hiloIndex < hilos.length; hiloIndex++) {
      // Execution time check
      if (Date.now() - tiempoInicio > CONFIG.MAX_TIEMPO_EJECUCION) {
        Logger.log('Time limit reached, saving progress...');
        break;
      }
      
      const hilo = hilos[hiloIndex];
      Logger.log('Processing thread ' + (hiloIndex + 1) + '/' + hilos.length);
      
      hilo.getMessages().forEach(mensaje => {
        estadisticas.correosEncontrados++;
        const correoId = generarIdCorreo(mensaje);
        const asunto = mensaje.getSubject().trim();
        
        // Skip if already processed (unless forced)
        if (!forzarReproceso && correosYaProcesados.has(correoId)) {
          estadisticas.correosYaProcesados++;
          return;
        }
        
        Logger.log((forzarReproceso ? 'Reprocessing' : 'Processing') + ': ' + asunto + ' - ' + mensaje.getDate().toLocaleDateString());
        
        const match = asunto.match(CONFIG.REGEX_ASUNTO);
        if (!match) {
          Logger.log('Invalid format: ' + asunto);
          estadisticas.errores++;
          return;
        }
        
        try {
          const adjuntos = mensaje.getAttachments();
          const pdfAdjuntos = adjuntos.filter(archivo => archivo.getContentType() === MimeType.PDF);
          
          if (pdfAdjuntos.length === 0) {
            Logger.log('No PDF attachments found');
            return;
          }
          
          let archivosDelMensaje = [];
          let algunArchivoCreado = false;
          
          pdfAdjuntos.forEach((archivo, index) => {
            const nuevoNombre = generarNombreArchivo(match, index, pdfAdjuntos.length);
            
            if (!forzarReproceso && archivosExistentesCache.has(nuevoNombre)) {
              Logger.log('File already exists: ' + nuevoNombre);
              estadisticas.archivosYaExistentes++;
            } else {
              carpetaDestino.createFile(archivo.copyBlob().setName(nuevoNombre));
              archivosDelMensaje.push(nuevoNombre);
              algunArchivoCreado = true;
              Logger.log((forzarReproceso ? 'Recreated' : 'Created') + ': ' + nuevoNombre);
            }
          });
          
          // Only mark as processed if files were created
          if (algunArchivoCreado) {
            nuevosProcesados.push(correoId);
            archivosCreados.push(...archivosDelMensaje);
            estadisticas.correosNuevosProcesados++;
            estadisticas.archivosCreados += archivosDelMensaje.length;
            
            // Batch index updates
            if (nuevosProcesados.length >= CONFIG.LOTE_SIZE) {
              escribirLoteAlIndice(archivoIndex, nuevosProcesados, correosYaProcesados);
              nuevosProcesados.length = 0;
            }
          } else if (forzarReproceso) {
            // In forced mode, mark as processed even if no files created
            nuevosProcesados.push(correoId);
            estadisticas.correosNuevosProcesados++;
          }
        } catch (error) {
          Logger.log('Error processing: ' + error.message);
          estadisticas.errores++;
        }
      });
    }
    
    // Final batch write
    if (nuevosProcesados.length > 0) {
      escribirLoteAlIndice(archivoIndex, nuevosProcesados, correosYaProcesados);
    }
  } catch (error) {
    Logger.log('PROCESSING ERROR: ' + error.message);
    // Save progress on error
    if (nuevosProcesados.length > 0) {
      try {
        escribirLoteAlIndice(archivoIndex, nuevosProcesados, correosYaProcesados);
        Logger.log('Progress saved before error');
      } catch (recoveryError) {
        Logger.log('Save error: ' + recoveryError.message);
      }
    }
    throw error;
  }
  
  return { estadisticas, archivosCreados };
}

/**
 * Writes batch of processed emails to index
 * @param {File} archivoIndex - Index file
 * @param {Array} nuevosProcesados - Newly processed email IDs
 * @param {Set} correosYaProcesados - Processed emails set
 */
function escribirLoteAlIndice(archivoIndex, nuevosProcesados, correosYaProcesados) {
  const contenidoActual = archivoIndex.getBlob().getDataAsString();
  const nuevoContenido = contenidoActual + nuevosProcesados.join("\n") + "\n";
  archivoIndex.setContent(nuevoContenido);
  nuevosProcesados.forEach(id => correosYaProcesados.add(id));
  Logger.log('Batch saved: ' + nuevosProcesados.length + ' emails');
}

/**
 * Shows processing summary
 * @param {Object} resultado - Processing results
 */
function mostrarResumen(resultado) {
  const { estadisticas } = resultado;
  
  Logger.log('=== PROCESSING SUMMARY ===');
  Logger.log('Emails found: ' + estadisticas.correosEncontrados);
  Logger.log('Already processed: ' + estadisticas.correosYaProcesados);
  Logger.log('Newly processed: ' + estadisticas.correosNuevosProcesados);
  Logger.log('Files created: ' + estadisticas.archivosCreados);
  Logger.log('Files already existed: ' + estadisticas.archivosYaExistentes);
  Logger.log('Errors: ' + estadisticas.errores);
}

/**
 * Creates empty summary for no results
 * @returns {Object} Empty results structure
 */
function crearResumenVacio() {
  return {
    estadisticas: {
      correosEncontrados: 0,
      correosYaProcesados: 0,
      correosNuevosProcesados: 0,
      archivosCreados: 0,
      archivosYaExistentes: 0,
      errores: 0
    },
    archivosCreados: []
  };
}

/* ==================== FILE CACHE SYSTEM ==================== */

/**
 * Builds cache of existing files for duplicate detection
 * @param {Folder} carpetaDestino - Destination folder
 * @returns {Set} Set of existing filenames
 */
function construirCacheArchivosExistentes(carpetaDestino) {
  const cache = new Set();
  
  // Root folder files
  const archivosRaiz = carpetaDestino.getFilesByType(MimeType.PDF);
  while (archivosRaiz.hasNext()) {
    cache.add(archivosRaiz.next().getName());
  }
  
  // Subfolder files (date-based organization)
  const subcarpetas = carpetaDestino.getFolders();
  let subcarpetasEscaneadas = 0;
  
  while (subcarpetas.hasNext()) {
    const subcarpeta = subcarpetas.next();
    const nombreSubcarpeta = subcarpeta.getName();
    
    // Only scan date-formatted folders (yyyy-mm-dd)
    if (/^\d{4}-\d{2}-\d{2}$/.test(nombreSubcarpeta)) {
      const archivosSubcarpeta = subcarpeta.getFilesByType(MimeType.PDF);
      let archivosEnSubcarpeta = 0;
      
      while (archivosSubcarpeta.hasNext()) {
        cache.add(archivosSubcarpeta.next().getName());
        archivosEnSubcarpeta++;
      }
      
      subcarpetasEscaneadas++;
      if (archivosEnSubcarpeta > 0) {
        Logger.log(nombreSubcarpeta + ': ' + archivosEnSubcarpeta + ' files');
      }
    }
  }
  
  Logger.log('Subfolders scanned: ' + subcarpetasEscaneadas);
  return cache;
}

// Export functions for testing and external use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    procesarCorreos,
    procesarMisFechas,
    reprocesarCorreosFecha,
    diagnosticarCorreosProblematicos,
    inicializarIndice,
    generarIdCorreo,
    generarNombreArchivo,
    construirCacheArchivosExistentes
  };
}
