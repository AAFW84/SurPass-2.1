// logs.js - Sistema de Logging Unificado para SurPass

// ==========================================
// CONFIGURACIÓN Y CONSTANTES
// ==========================================
// Las constantes SHEET_LOGS y SHEET_ERRORES se importan desde constantes.js

// ==========================================
// SISTEMA DE LOGGING UNIFICADO
// ==========================================
/**
 * Sistema de Logging Unificado para SurPass
 * Centraliza todo el manejo de logs del sistema
 */
const Logger = {
  /**
   * Registra un evento en la hoja de logs.
   * @param {string} nivel - INFO, WARNING, ERROR, DEBUG.
   * @param {string} mensaje - Mensaje del log.
   * @param {Object} detalles - Objeto con datos adicionales.
   * @param {string} fuente - Quién o qué generó el log.
   * @param {boolean} [noEscribirEnHoja=false] - Si es true, solo imprime en consola.
   */
  log: function(nivel, mensaje, detalles = {}, fuente = 'Sistema', noEscribirEnHoja = false) {
    try {
      const timestamp = new Date();
      const logEntry = { 
        nivel, 
        mensaje, 
        detalles, 
        fuente, 
        timestamp: timestamp.toISOString(),
        usuario: Session.getActiveUser().getEmail() || 'N/A'
      };
      
      // Imprimir siempre en la consola de Apps Script
      console.log(`[${nivel}] ${fuente}: ${mensaje}`, detalles ? JSON.stringify(detalles, null, 2) : '');
      
      // Si no se debe escribir en hoja, terminar aquí
      if (noEscribirEnHoja) return;
      
      const hojaLogs = SpreadsheetApp.getActive().getSheetByName(SHEET_LOGS);
      if (!hojaLogs) {
        console.error('❌ Hoja de logs no encontrada: ' + SHEET_LOGS);
        return;
      }
      
      // Escribir en la hoja de logs
      hojaLogs.appendRow([
        timestamp,
        nivel,
        mensaje,
        JSON.stringify(detalles),
        fuente,
        logEntry.usuario
      ]);
      
      // Si es un error crítico, también registrar en la hoja de errores
      if (nivel === 'ERROR') {
        this.registrarEnHojaErrores(timestamp, mensaje, detalles, fuente);
      }
      
    } catch (e) {
      console.error('❌ Error CRÍTICO en el sistema de logging:', e);
    }
  },
  
  /**
   * Registra errores en la hoja específica de errores
   * @private
   */
  registrarEnHojaErrores: function(timestamp, mensaje, detalles, fuente) {
    try {
      const hojaErrores = SpreadsheetApp.getActive().getSheetByName(SHEET_ERRORES);
      if (hojaErrores) {
        hojaErrores.appendRow([
          timestamp,
          'ERR_' + Date.now(), // Código único de error
          mensaje,
          JSON.stringify(detalles),
          fuente
        ]);
      }
    } catch (e) {
      console.error('❌ Error al registrar en hoja de errores:', e);
    }
  },
  
  // Métodos de conveniencia
  info: function(mensaje, detalles, fuente) {
    this.log('INFO', mensaje, detalles, fuente);
  },
  
  warn: function(mensaje, detalles, fuente) {
    this.log('WARNING', mensaje, detalles, fuente);
  },
  
  error: function(mensaje, detalles, fuente) {
    this.log('ERROR', mensaje, detalles, fuente);
  },
  
  debug: function(mensaje, detalles, fuente) {
    this.log('DEBUG', mensaje, detalles, fuente);
  }
};

// ==========================================
// FUNCIONES DE COMPATIBILIDAD (WRAPPER)
// ==========================================
/**
 * Registra un evento en la hoja de logs (compatibilidad con código antiguo)
 * @param {string} nivel - Nivel del log (INFO, WARNING, ERROR).
 * @param {string} mensaje - Mensaje del log.
 * @param {Object} detalles - Objeto con detalles adicionales.
 * @param {string} fuente - Fuente del log (e.g., 'Sistema', 'Panel Admin').
 */
function registrarLog(nivel, mensaje, detalles = {}, fuente = 'Sistema') {
  Logger.log(nivel, mensaje, detalles, fuente);
}

/**
 * Registra un error en la hoja de errores (compatibilidad con código antiguo)
 * @param {string} codigo - Código del error.
 * @param {string} mensaje - Mensaje del error.
 * @param {Object} datos - Datos adicionales del error.
 */
function registrarErrorLog(codigo, mensaje, datos) {
  Logger.error(mensaje, { codigo, ...datos }, 'Sistema');
}

/**
 * Registra eventos de información en el sistema de logs
 * @param {string} mensaje - Mensaje del evento.
 * @param {string} funcion - Función que genera el evento.
 * @param {Object} detalles - Detalles adicionales.
 */
function registrarEventoInfo(mensaje, funcion, detalles = null) {
  Logger.info(mensaje, { funcion, ...detalles }, 'Panel Admin');
}

/**
 * Registra eventos de error en el sistema de logs
 * @param {string} mensaje - Mensaje del error.
 * @param {string} funcion - Función que genera el error.
 * @param {string} stack - Stack trace del error.
 */
function registrarEventoError(mensaje, funcion, stack = null) {
  Logger.error(mensaje, { funcion, stack }, 'Panel Admin');
}

// ==========================================
// FUNCIONES DE CONSULTA Y RECUPERACIÓN
// ==========================================

/**
 * Obtiene los logs más recientes del sistema
 * @param {number} limite - Número máximo de logs a obtener
 * @return {Object} Resultado con los logs y estadísticas
 */
function obtenerLogsRecientes(limite = 50) {
  try {
    const hojaLogs = SpreadsheetApp.getActive().getSheetByName(SHEET_LOGS);
    if (!hojaLogs || hojaLogs.getLastRow() <= 1) {
      return {
        success: true,
        logs: [],
        stats: { total: 0, errores: 0, warnings: 0, info: 0 }
      };
    }
    
    const datos = hojaLogs.getDataRange().getValues();
    const encabezados = datos[0];
    const logs = [];
    
    // Procesar desde el final (logs más recientes primero)
    for (let i = Math.min(datos.length - 1, limite); i >= 1; i--) {
      const fila = datos[datos.length - logs.length - 1];
      if (!fila || fila.every(cell => cell === '')) continue;
      
      logs.push({
        id: `log_${i}_${Date.now()}`,
        timestamp: fila[0],
        fecha: fila[0],
        nivel: fila[1] || 'INFO',
        mensaje: fila[2] || '',
        detalles: fila[3] || '',
        usuario: fila[4] || 'Sistema',
        ip: fila[5] || 'N/A'
      });
      
      if (logs.length >= limite) break;
    }
    
    // Calcular estadísticas
    const stats = {
      total: logs.length,
      errores: logs.filter(l => l.nivel === 'ERROR').length,
      warnings: logs.filter(l => l.nivel === 'WARNING' || l.nivel === 'WARN').length,
      info: logs.filter(l => l.nivel === 'INFO').length
    };
    
    return {
      success: true,
      logs: logs.reverse(), // Mostrar en orden cronológico
      stats
    };
    
  } catch (error) {
    Logger.error('Error obteniendo logs recientes', { error: error.message }, 'obtenerLogsRecientes');
    return {
      success: false,
      message: 'Error al obtener logs: ' + error.message,
      logs: [],
      stats: { total: 0, errores: 1, warnings: 0, info: 0 }
    };
  }
}

/**
 * Carga logs del sistema con manejo robusto de múltiples fuentes
 * @param {number} limite - Número máximo de logs a obtener
 * @return {Object} Logs del sistema con estadísticas
 */
function loadLogs(limite = 50) {
  console.log('📋 [ADMIN-INTERFACE] === INICIO loadLogs ===');
  console.log('📋 [ADMIN-INTERFACE] Parámetros:', { limite });
  console.log('📋 [ADMIN-INTERFACE] Timestamp:', new Date().toISOString());
  
  Logger.info('Iniciando carga de logs del sistema', { limite }, 'loadLogs');
  
  try {
    let logsObtenidos = [];
    let fuenteUsada = 'ninguna';
    
    // Intentar obtener logs usando la función principal
    try {
      const resultLogs = obtenerLogsRecientes(limite);
      if (resultLogs && resultLogs.success && Array.isArray(resultLogs.logs)) {
        logsObtenidos = resultLogs.logs;
        fuenteUsada = 'obtenerLogsRecientes';
        console.log(`✅ [ADMIN-INTERFACE] ${logsObtenidos.length} logs obtenidos de obtenerLogsRecientes`);
      }
    } catch (error) {
      console.error('❌ [ADMIN-INTERFACE] Error usando obtenerLogsRecientes:', error);
    }
    
    // Si no se obtuvieron logs, intentar lectura directa
    if (logsObtenidos.length === 0) {
      console.log('📄 [ADMIN-INTERFACE] Intentando lectura directa de Google Sheets...');
      try {
        const ss = SpreadsheetApp.getActive();
        const hojaLogs = ss.getSheetByName(SHEET_LOGS);
        
        if (hojaLogs && hojaLogs.getLastRow() > 1) {
          const datos = hojaLogs.getDataRange().getValues();
          const encabezados = datos[0].map(h => String(h).toLowerCase().trim());
          
          // Mapear índices de columnas
          const getColIndex = (possibleNames) => {
            for (const name of possibleNames) {
              const idx = encabezados.findIndex(h => h.includes(name));
              if (idx >= 0) return idx;
            }
            return -1;
          };
          
          const indices = {
            timestamp: getColIndex(['timestamp', 'fecha', 'fechahora']),
            nivel: getColIndex(['nivel', 'level', 'tipo']),
            mensaje: getColIndex(['mensaje', 'message', 'detalle']),
            detalles: getColIndex(['detalles', 'details', 'data']),
            usuario: getColIndex(['usuario', 'user', 'autor', 'fuente']),
            ip: getColIndex(['ip', 'direccionip'])
          };
          
          // Procesar filas de logs
          logsObtenidos = [];
          for (let i = Math.max(1, datos.length - limite); i < datos.length; i++) {
            const fila = datos[i];
            if (!fila || fila.every(cell => cell === '')) continue;
            
            const getValue = (key) => {
              const idx = indices[key];
              return (idx >= 0 && idx < fila.length) ? fila[idx] : '';
            };
            
            const logEntry = {
              id: `log_${i}_${Date.now()}`,
              timestamp: getValue('timestamp') || new Date().toISOString(),
              fecha: getValue('timestamp') || new Date().toISOString(),
              nivel: (getValue('nivel') || 'INFO').toString().toUpperCase(),
              mensaje: getValue('mensaje') || 'Mensaje no especificado',
              detalles: getValue('detalles') || '',
              usuario: getValue('usuario') || 'Sistema',
              ip: getValue('ip') || 'N/A'
            };
            
            logsObtenidos.push(logEntry);
          }
          
          fuenteUsada = 'Google Sheets directo';
          console.log(`✅ [ADMIN-INTERFACE] ${logsObtenidos.length} logs obtenidos por lectura directa`);
        } else {
          console.warn('⚠️ [ADMIN-INTERFACE] Hoja de logs vacía o no encontrada');
          // Crear un log de ejemplo si no hay datos
          logsObtenidos = [{
            id: 'log_demo_1',
            timestamp: new Date().toISOString(),
            fecha: new Date().toISOString(),
            nivel: 'INFO',
            mensaje: 'Sistema de logs inicializado. No se encontraron registros previos.',
            usuario: 'Sistema',
            ip: 'N/A'
          }];
          fuenteUsada = 'Datos de inicialización';
        }
      } catch (sheetsError) {
        console.error('❌ [ADMIN-INTERFACE] Error accediendo directamente a Google Sheets:', sheetsError);
        Logger.error('Error accediendo directamente a Google Sheets para logs', 
                    { error: sheetsError.message }, 'loadLogs');
        
        return {
          success: false,
          message: 'Error al cargar los logs: ' + sheetsError.message,
          logs: [{
            id: 'log_error_1',
            timestamp: new Date().toISOString(),
            nivel: 'ERROR',
            mensaje: 'Error al cargar los logs: ' + sheetsError.message,
            usuario: 'Sistema',
            ip: 'N/A'
          }],
          stats: { total: 0, errores: 1, warnings: 0, info: 0 },
          error: sheetsError.message
        };
      }
    }
    
    // Procesar logs obtenidos
    const result = processLogsResult(logsObtenidos, fuenteUsada);
    
    console.log(`✅ [ADMIN-INTERFACE] ${result.stats.total} logs cargados desde: ${fuenteUsada}`);
    Logger.info(`Logs cargados exitosamente: ${result.stats.total} registros`, {
      fuente: fuenteUsada,
      limite: limite,
      stats: result.stats
    }, 'loadLogs');
    
    return result;
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error crítico cargando logs:', error);
    Logger.error(`Error crítico cargando logs: ${error.message}`, 
                { stack: error.stack }, 'loadLogs');
    
    return {
      success: false,
      message: 'Error crítico cargando logs del sistema: ' + (error.message || 'Error desconocido'),
      logs: [{
        id: 'log_error_critical',
        fecha: new Date().toISOString(),
        nivel: 'ERROR',
        mensaje: 'Error crítico en el sistema de logs: ' + (error.message || 'Error desconocido'),
        usuario: 'Sistema',
        detalles: error.stack || 'Sin detalles de stack trace',
        ip: 'N/A'
      }],
      stats: { total: 1, errores: 1, warnings: 0, info: 0 },
      error: error.message || 'Error desconocido'
    };
  }
}

/**
 * Procesa el resultado de logs para el frontend
 * @param {Array} logs - Array de logs
 * @param {string} fuente - Fuente de donde se obtuvieron los logs
 * @return {Object} Resultado procesado con estadísticas
 */
function processLogsResult(logs, fuente = 'desconocida') {
  if (!Array.isArray(logs)) {
    console.warn('⚠️ [ADMIN-INTERFACE] processLogsResult recibió datos inválidos:', logs);
    return {
      success: false,
      message: 'Datos de logs inválidos',
      logs: [],
      stats: { total: 0, errores: 0, warnings: 0, info: 0 },
      fuente: fuente
    };
  }
  
  // Calcular estadísticas
  const stats = {
    total: logs.length,
    errores: logs.filter(log => String(log.nivel || log.tipo || '').toUpperCase() === 'ERROR').length,
    warnings: logs.filter(log => {
      const nivel = String(log.nivel || log.tipo || '').toUpperCase();
      return nivel === 'WARNING' || nivel === 'WARN' || nivel === 'ADVERTENCIA';
    }).length,
    info: logs.filter(log => String(log.nivel || log.tipo || '').toUpperCase() === 'INFO').length,
    debug: logs.filter(log => String(log.nivel || log.tipo || '').toUpperCase() === 'DEBUG').length
  };
  
  // Formatear logs para el frontend
  const logsFormateados = logs.map((log, index) => {
    // Manejar fecha de forma robusta
    let fechaFinal;
    try {
      if (log.fecha instanceof Date) {
        fechaFinal = log.fecha;
      } else if (log.timestamp) {
        fechaFinal = new Date(log.timestamp);
        if (isNaN(fechaFinal.getTime())) {
          fechaFinal = new Date();
        }
      } else {
        fechaFinal = new Date();
      }
    } catch (error) {
      fechaFinal = new Date();
    }
    
    return {
      id: log.id || `log_${index}_${Date.now()}`,
      fecha: fechaFinal.toLocaleString('es-CO', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      }),
      timestamp: fechaFinal.toISOString(),
      nivel: String(log.nivel || log.tipo || log.level || 'INFO').toUpperCase(),
      mensaje: log.mensaje || log.message || 'Sin mensaje',
      usuario: log.usuario || log.user || log.fuente || 'Sistema',
      detalles: log.detalles || log.datos || log.details || log.data || '',
      ip: log.ip || 'N/A'
    };
  });
  
  console.log(`✅ [ADMIN-INTERFACE] Logs procesados: ${stats.total} registros (${stats.errores} errores, ${stats.warnings} warnings, ${stats.info} info, ${stats.debug} debug)`);
  
  return {
    success: true,
    message: `${stats.total} logs procesados desde ${fuente}`,
    logs: logsFormateados,
    stats: stats,
    fuente: fuente
  };
}

// ==========================================
// FUNCIONES DE MANTENIMIENTO
// ==========================================

/**
 * Limpia los logs del sistema según los criterios especificados
 * @param {Object} opciones - Opciones de limpieza
 * @param {number} [opciones.dias=30] - Número de días de antigüedad máxima a conservar
 * @param {boolean} [opciones.soloErrores=false] - Si es true, solo limpia los errores
 * @param {string} [opciones.usuario] - Si se especifica, solo limpia los logs de este usuario
 * @return {Object} Resultado de la operación
 */
function limpiarLogs(opciones = {}) {
  const { dias = 30, soloErrores = false, usuario } = opciones;
  const fechaLimite = new Date();
  fechaLimite.setDate(fechaLimite.getDate() - dias);
  
  try {
    const ss = SpreadsheetApp.getActive();
    const hojaLogs = ss.getSheetByName(SHEET_LOGS);
    
    if (!hojaLogs) {
      return {
        success: false,
        message: 'No se encontró la hoja de logs',
        detalles: { hojaBuscada: SHEET_LOGS }
      };
    }
    
    const datos = hojaLogs.getDataRange().getValues();
    if (datos.length <= 1) {
      return {
        success: true,
        message: 'No hay logs para limpiar',
        registrosEliminados: 0
      };
    }
    
    const encabezados = datos[0].map(h => String(h).toLowerCase().trim());
    const indiceFecha = encabezados.findIndex(h => h.includes('fecha') || h.includes('timestamp'));
    const indiceNivel = encabezados.findIndex(h => h.includes('nivel') || h.includes('tipo'));
    const indiceUsuario = encabezados.findIndex(h => h.includes('usuario') || h.includes('user'));
    
    if (indiceFecha === -1) {
      return {
        success: false,
        message: 'No se pudo identificar la columna de fecha en los logs',
        detalles: { encabezadosEncontrados: encabezados }
      };
    }
    
    // Filtrar filas a eliminar
    const filasAEliminar = [];
    for (let i = datos.length - 1; i >= 1; i--) {
      const fila = datos[i];
      const fechaStr = fila[indiceFecha];
      let fechaLog;
      
      try {
        fechaLog = new Date(fechaStr);
        if (isNaN(fechaLog.getTime())) {
          filasAEliminar.push(i + 1);
          continue;
        }
      } catch (e) {
        filasAEliminar.push(i + 1);
        continue;
      }
      
      // Verificar si cumple con los criterios de eliminación
      const esAntiguo = fechaLog < fechaLimite;
      const esError = indiceNivel !== -1 && 
                     String(fila[indiceNivel] || '').toUpperCase().includes('ERROR');
      const esDelUsuario = usuario && indiceUsuario !== -1 && 
                          String(fila[indiceUsuario] || '').toLowerCase() === usuario.toLowerCase();
      
      const debeEliminar = esAntiguo && 
                          (soloErrores ? esError : true) && 
                          (usuario ? esDelUsuario : true);
      
      if (debeEliminar) {
        filasAEliminar.push(i + 1);
      }
    }
    
    // Eliminar filas (empezar por las de mayor índice)
    filasAEliminar.sort((a, b) => b - a);
    
    let eliminados = 0;
    filasAEliminar.forEach(fila => {
      try {
        hojaLogs.deleteRow(fila);
        eliminados++;
      } catch (e) {
        console.error(`Error eliminando fila ${fila}:`, e);
      }
    });
    
    // Registrar la acción
    Logger.info(
      `Limpieza de logs completada: ${eliminados} registros eliminados`,
      { 
        totalRegistros: datos.length - 1, 
        eliminados,
        criterios: { dias, soloErrores, usuario }
      },
      'limpiarLogs'
    );
    
    return {
      success: true,
      message: `Se eliminaron ${eliminados} registros de logs`,
      registrosEliminados: eliminados,
      totalRegistros: datos.length - 1
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error en limpiarLogs:', error);
    Logger.error('Error al limpiar logs: ' + error.message, 
                { stack: error.stack }, 'limpiarLogs');
    
    return {
      success: false,
      message: 'Error al limpiar logs: ' + error.message,
      error: error.toString()
    };
  }
}

/**
 * Exporta los logs a un archivo CSV o JSON
 * @param {Object} opciones - Opciones de exportación
 * @param {number} [opciones.limite=1000] - Número máximo de registros a exportar
 * @param {string} [opciones.formato='csv'] - Formato de exportación (csv o json)
 * @param {string} [opciones.filtroNivel] - Filtrar por nivel (ERROR, WARNING, INFO, etc.)
 * @param {string} [opciones.usuario] - Filtrar por usuario
 * @param {string} [opciones.fechaDesde] - Fecha desde (formato YYYY-MM-DD)
 * @param {string} [opciones.fechaHasta] - Fecha hasta (formato YYYY-MM-DD)
 * @return {Object} Resultado con el contenido del archivo y metadatos
 */
function exportarLogs(opciones = {}) {
  const {
    limite = 1000,
    formato = 'csv',
    filtroNivel,
    usuario,
    fechaDesde,
    fechaHasta
  } = opciones;
  
  try {
    // Obtener logs con un límite mayor para aplicar filtros
    const resultadoLogs = loadLogs(limite * 2);
    
    if (!resultadoLogs.success) {
      throw new Error(resultadoLogs.message || 'Error al cargar logs para exportación');
    }
    
    let logs = resultadoLogs.logs || [];
    
    // Aplicar filtros si se especificaron
    if (filtroNivel || usuario || fechaDesde || fechaHasta) {
      logs = logs.filter(log => {
        // Filtrar por nivel
        if (filtroNivel && !String(log.nivel || '').toUpperCase().includes(filtroNivel.toUpperCase())) {
          return false;
        }
        
        // Filtrar por usuario
        if (usuario && !String(log.usuario || '').toLowerCase().includes(usuario.toLowerCase())) {
          return false;
        }
        
        // Filtrar por rango de fechas
        if (fechaDesde || fechaHasta) {
          try {
            const fechaLog = new Date(log.timestamp || log.fecha);
            const desde = fechaDesde ? new Date(fechaDesde) : null;
            const hasta = fechaHasta ? new Date(fechaHasta) : null;
            
            if (desde && fechaLog < desde) return false;
            if (hasta) {
              const hastaFinDia = new Date(hasta);
              hastaFinDia.setHours(23, 59, 59, 999);
              if (fechaLog > hastaFinDia) return false;
            }
          } catch (e) {
            console.warn('Error al filtrar por fecha:', e);
            return false;
          }
        }
        
        return true;
      }).slice(0, limite);
    } else {
      logs = logs.slice(0, limite);
    }
    
    // Generar el contenido según el formato solicitado
    let contenido = '';
    let nombreArchivo = `logs_surpass_${new Date().toISOString().split('T')[0]}`;
    
    if (formato.toLowerCase() === 'json') {
      // Formato JSON
      contenido = JSON.stringify(logs, null, 2);
      nombreArchivo += '.json';
    } else {
      // Formato CSV por defecto
      if (logs.length === 0) {
        contenido = 'No hay registros que coincidan con los criterios de búsqueda';
      } else {
        // Obtener todos los campos únicos de los logs
        const campos = ['fecha', 'nivel', 'mensaje', 'usuario', 'detalles', 'ip'];
        
        // Crear fila de encabezados
        contenido = campos.join(';') + '\n';
        
        // Agregar filas de datos
        logs.forEach(log => {
          const fila = campos.map(campo => {
            let valor = log[campo] !== undefined ? log[campo] : '';
            
            // Convertir a string y escapar comillas
            if (valor === null || valor === undefined) valor = '';
            let valorStr = String(valor);
            
            // Si el valor contiene punto y coma, comillas o saltos de línea, encerrar entre comillas
            if (valorStr.includes(';') || valorStr.includes('"') || valorStr.includes('\n')) {
              valorStr = `"${valorStr.replace(/"/g, '""')}"`;
            }
            
            return valorStr;
          });
          
          contenido += fila.join(';') + '\n';
        });
      }
      
      nombreArchivo += '.csv';
    }
    
    // Registrar la acción de exportación
    Logger.info(
      `Exportación de logs completada: ${logs.length} registros en formato ${formato.toUpperCase()}`,
      {
        totalRegistros: logs.length,
        formato,
        filtros: { filtroNivel, usuario, fechaDesde, fechaHasta }
      },
      'exportarLogs'
    );
    
    return {
      success: true,
      message: `Exportación completada: ${logs.length} registros`,
      nombreArchivo,
      formato,
      totalRegistros: logs.length,
      contenido,
      urlDescarga: `data:text/${formato.toLowerCase()};charset=utf-8,${encodeURIComponent(contenido)}`
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error en exportarLogs:', error);
    Logger.error('Error al exportar logs: ' + error.message, 
                { stack: error.stack }, 'exportarLogs');
    return {
      success: false,
      message: 'Error al exportar logs: ' + error.message,
      error: error.toString()
    };
  }
}

// ==========================================
// FUNCIONES DE ANÁLISIS Y ESTADÍSTICAS
// ==========================================

/**
 * Obtiene estadísticas detalladas de los logs
 * @param {Object} opciones - Opciones para el análisis
 * @param {number} [opciones.dias=7] - Número de días a analizar
 * @return {Object} Estadísticas detalladas de los logs
 */
function obtenerEstadisticasLogs(opciones = {}) {
  const { dias = 7 } = opciones;
  const fechaLimite = new Date();
  fechaLimite.setDate(fechaLimite.getDate() - dias);
  
  try {
    const hojaLogs = SpreadsheetApp.getActive().getSheetByName(SHEET_LOGS);
    if (!hojaLogs || hojaLogs.getLastRow() <= 1) {
      return {
        success: false,
        message: 'No hay datos de logs para analizar',
        estadisticas: {}
      };
    }
    
    const datos = hojaLogs.getDataRange().getValues();
    const logs = [];
    
    // Procesar logs dentro del rango de fechas
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!fila || fila.every(cell => cell === '')) continue;
      
      const fechaLog = new Date(fila[0]);
      if (fechaLog >= fechaLimite) {
        logs.push({
          fecha: fechaLog,
          nivel: String(fila[1] || 'INFO').toUpperCase(),
          mensaje: fila[2] || '',
          fuente: fila[4] || 'Sistema',
          usuario: fila[5] || 'N/A'
        });
      }
    }
    
    // Calcular estadísticas
    const estadisticas = {
      periodo: {
        desde: fechaLimite.toISOString(),
        hasta: new Date().toISOString(),
        dias: dias
      },
      totales: {
        total: logs.length,
        errores: logs.filter(l => l.nivel === 'ERROR').length,
        warnings: logs.filter(l => l.nivel === 'WARNING' || l.nivel === 'WARN').length,
        info: logs.filter(l => l.nivel === 'INFO').length,
        debug: logs.filter(l => l.nivel === 'DEBUG').length
      },
      porFuente: {},
      porUsuario: {},
      porDia: {},
      errorMasFrecuente: null,
      horasPico: []
    };
    
    // Agrupar por fuente
    logs.forEach(log => {
      if (!estadisticas.porFuente[log.fuente]) {
        estadisticas.porFuente[log.fuente] = 0;
      }
      estadisticas.porFuente[log.fuente]++;
      
      // Agrupar por usuario
      if (!estadisticas.porUsuario[log.usuario]) {
        estadisticas.porUsuario[log.usuario] = 0;
      }
      estadisticas.porUsuario[log.usuario]++;
      
      // Agrupar por día
      const dia = log.fecha.toISOString().split('T')[0];
      if (!estadisticas.porDia[dia]) {
        estadisticas.porDia[dia] = {
          total: 0,
          errores: 0,
          warnings: 0,
          info: 0
        };
      }
      estadisticas.porDia[dia].total++;
      if (log.nivel === 'ERROR') estadisticas.porDia[dia].errores++;
      if (log.nivel === 'WARNING' || log.nivel === 'WARN') estadisticas.porDia[dia].warnings++;
      if (log.nivel === 'INFO') estadisticas.porDia[dia].info++;
    });
    
    // Encontrar el error más frecuente
    const errores = logs.filter(l => l.nivel === 'ERROR');
    if (errores.length > 0) {
      const conteoErrores = {};
      errores.forEach(error => {
        const clave = error.mensaje.substring(0, 100); // Usar los primeros 100 caracteres como clave
        conteoErrores[clave] = (conteoErrores[clave] || 0) + 1;
      });
      
      const errorMasComun = Object.entries(conteoErrores)
        .sort((a, b) => b[1] - a[1])[0];
      
      estadisticas.errorMasFrecuente = {
        mensaje: errorMasComun[0],
        cantidad: errorMasComun[1]
      };
    }
    
    // Calcular horas pico
    const porHora = {};
    logs.forEach(log => {
      const hora = log.fecha.getHours();
      porHora[hora] = (porHora[hora] || 0) + 1;
    });
    
    estadisticas.horasPico = Object.entries(porHora)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([hora, cantidad]) => ({
        hora: parseInt(hora),
        cantidad
      }));
    
    Logger.info('Estadísticas de logs generadas', { dias, totalLogs: logs.length }, 'obtenerEstadisticasLogs');
    
    return {
      success: true,
      message: `Estadísticas calculadas para ${logs.length} logs en los últimos ${dias} días`,
      estadisticas
    };
    
  } catch (error) {
    Logger.error('Error obteniendo estadísticas de logs', { error: error.message }, 'obtenerEstadisticasLogs');
    return {
      success: false,
      message: 'Error al obtener estadísticas: ' + error.message,
      estadisticas: {}
    };
  }
}

/**
 * Busca logs según criterios específicos
 * @param {Object} criterios - Criterios de búsqueda
 * @param {string} [criterios.texto] - Texto a buscar en el mensaje
 * @param {string} [criterios.nivel] - Nivel específico de log
 * @param {string} [criterios.fuente] - Fuente específica
 * @param {string} [criterios.usuario] - Usuario específico
 * @param {Date} [criterios.fechaDesde] - Fecha desde
 * @param {Date} [criterios.fechaHasta] - Fecha hasta
 * @param {number} [criterios.limite=100] - Límite de resultados
 * @return {Object} Resultado de la búsqueda
 */
function buscarLogs(criterios = {}) {
  const {
    texto,
    nivel,
    fuente,
    usuario,
    fechaDesde,
    fechaHasta,
    limite = 100
  } = criterios;
  
  try {
    const hojaLogs = SpreadsheetApp.getActive().getSheetByName(SHEET_LOGS);
    if (!hojaLogs || hojaLogs.getLastRow() <= 1) {
      return {
        success: true,
        message: 'No hay logs disponibles para buscar',
        logs: [],
        totalEncontrados: 0
      };
    }
    
    const datos = hojaLogs.getDataRange().getValues();
    const logsEncontrados = [];
    
    // Buscar desde el final (logs más recientes primero)
    for (let i = datos.length - 1; i >= 1 && logsEncontrados.length < limite; i--) {
      const fila = datos[i];
      if (!fila || fila.every(cell => cell === '')) continue;
      
      const log = {
        id: `log_${i}_${Date.now()}`,
        fecha: fila[0],
        nivel: String(fila[1] || 'INFO').toUpperCase(),
        mensaje: fila[2] || '',
        detalles: fila[3] || '',
        fuente: fila[4] || 'Sistema',
        usuario: fila[5] || 'N/A'
      };
      
      // Aplicar filtros
      let cumpleFiltros = true;
      
      // Filtrar por texto
      if (texto && cumpleFiltros) {
        const textoLower = texto.toLowerCase();
        cumpleFiltros = log.mensaje.toLowerCase().includes(textoLower) ||
                        log.detalles.toLowerCase().includes(textoLower);
      }
      
      // Filtrar por nivel
      if (nivel && cumpleFiltros) {
        cumpleFiltros = log.nivel === nivel.toUpperCase();
      }
      
      // Filtrar por fuente
      if (fuente && cumpleFiltros) {
        cumpleFiltros = log.fuente.toLowerCase().includes(fuente.toLowerCase());
      }
      
      // Filtrar por usuario
      if (usuario && cumpleFiltros) {
        cumpleFiltros = log.usuario.toLowerCase().includes(usuario.toLowerCase());
      }
      
      // Filtrar por rango de fechas
      if ((fechaDesde || fechaHasta) && cumpleFiltros) {
        try {
          const fechaLog = new Date(log.fecha);
          if (fechaDesde && fechaLog < new Date(fechaDesde)) cumpleFiltros = false;
          if (fechaHasta && fechaLog > new Date(fechaHasta)) cumpleFiltros = false;
        } catch (e) {
          cumpleFiltros = false;
        }
      }
      
      if (cumpleFiltros) {
        logsEncontrados.push(log);
      }
    }
    
    Logger.info('Búsqueda de logs completada', {
      criterios,
      encontrados: logsEncontrados.length
    }, 'buscarLogs');
    
    return {
      success: true,
      message: `Se encontraron ${logsEncontrados.length} logs que coinciden con los criterios`,
      logs: logsEncontrados,
      totalEncontrados: logsEncontrados.length,
      criteriosAplicados: criterios
    };
    
  } catch (error) {
    Logger.error('Error buscando logs', { error: error.message, criterios }, 'buscarLogs');
    return {
      success: false,
      message: 'Error al buscar logs: ' + error.message,
      logs: [],
      totalEncontrados: 0
    };
  }
}

// ==========================================
// FUNCIONES DE MONITOREO Y ALERTAS
// ==========================================

/**
 * Verifica si hay errores críticos recientes y envía alertas si es necesario
 * @param {Object} opciones - Opciones de monitoreo
 * @param {number} [opciones.minutos=60] - Ventana de tiempo en minutos para verificar
 * @param {number} [opciones.umbralErrores=10] - Número de errores para considerar crítico
 * @param {boolean} [opciones.enviarEmail=false] - Si debe enviar email de alerta
 * @return {Object} Resultado del monitoreo
 */
function monitorearErroresCriticos(opciones = {}) {
  const { 
    minutos = 60, 
    umbralErrores = 10,
    enviarEmail = false 
  } = opciones;
  
  const fechaLimite = new Date();
  fechaLimite.setMinutes(fechaLimite.getMinutes() - minutos);
  
  try {
    // Buscar errores recientes
    const resultado = buscarLogs({
      nivel: 'ERROR',
      fechaDesde: fechaLimite.toISOString(),
      limite: 1000
    });
    
    if (!resultado.success) {
      throw new Error(resultado.message);
    }
    
    const erroresRecientes = resultado.logs;
    const hayAlerta = erroresRecientes.length >= umbralErrores;
    
    const resumen = {
      periodo: `Últimos ${minutos} minutos`,
      erroresEncontrados: erroresRecientes.length,
      umbralConfigrado: umbralErrores,
      alertaActivada: hayAlerta,
      timestamp: new Date().toISOString()
    };
    
    if (hayAlerta) {
      // Agrupar errores por tipo
      const erroresPorTipo = {};
      erroresRecientes.forEach(error => {
        const tipo = error.mensaje.substring(0, 50);
        erroresPorTipo[tipo] = (erroresPorTipo[tipo] || 0) + 1;
      });
      
      resumen.erroresPorTipo = erroresPorTipo;
      resumen.errorMasFrecuente = Object.entries(erroresPorTipo)
        .sort((a, b) => b[1] - a[1])[0];
      
      // Registrar la alerta
      Logger.warn('⚠️ ALERTA: Umbral de errores críticos alcanzado', resumen, 'monitorearErroresCriticos');
      
      // Enviar email si está configurado
      if (enviarEmail) {
        try {
          const email = Session.getActiveUser().getEmail();
          const asunto = `[ALERTA SurPass] ${erroresRecientes.length} errores detectados`;
          const mensaje = `
            Se han detectado ${erroresRecientes.length} errores en los últimos ${minutos} minutos.
            
            Resumen:
            - Total de errores: ${erroresRecientes.length}
            - Error más frecuente: ${resumen.errorMasFrecuente[0]} (${resumen.errorMasFrecuente[1]} ocurrencias)
            - Período monitoreado: ${resumen.periodo}
            
            Por favor, revise el panel de administración para más detalles.
          `;
          
          MailApp.sendEmail(email, asunto, mensaje);
          resumen.emailEnviado = true;
        } catch (emailError) {
          Logger.error('Error enviando email de alerta', { error: emailError.message }, 'monitorearErroresCriticos');
          resumen.emailEnviado = false;
          resumen.errorEmail = emailError.message;
        }
      }
    } else {
      Logger.info('Sistema funcionando normalmente', resumen, 'monitorearErroresCriticos');
    }
    
    return {
      success: true,
      message: hayAlerta ? 
        `⚠️ ALERTA: Se encontraron ${erroresRecientes.length} errores` : 
        '✅ Sistema funcionando normalmente',
      resumen,
      errores: hayAlerta ? erroresRecientes.slice(0, 10) : [] // Retornar solo los primeros 10 si hay alerta
    };
    
  } catch (error) {
    Logger.error('Error monitoreando errores críticos', { error: error.message }, 'monitorearErroresCriticos');
    return {
      success: false,
      message: 'Error al monitorear el sistema: ' + error.message,
      resumen: {},
      errores: []
    };
  }
}

// ==========================================
// FUNCIONES DE INICIALIZACIÓN Y VALIDACIÓN
// ==========================================

/**
 * Inicializa o valida la estructura de la hoja de logs
 * @return {Object} Resultado de la inicialización
 */
function inicializarSistemaLogs() {
  try {
    const ss = SpreadsheetApp.getActive();
    let hojaLogs = ss.getSheetByName(SHEET_LOGS);
    let hojaErrores = ss.getSheetByName(SHEET_ERRORES);
    let creadas = [];
    
    // Crear hoja de logs si no existe
    if (!hojaLogs) {
      hojaLogs = ss.insertSheet(SHEET_LOGS);
      hojaLogs.appendRow([
        'Timestamp',
        'Nivel',
        'Mensaje',
        'Detalles',
        'Fuente',
        'Usuario'
      ]);
      
      // Formatear encabezados
      const encabezados = hojaLogs.getRange(1, 1, 1, 6);
      encabezados.setFontWeight('bold');
      encabezados.setBackground('#4285F4');
      encabezados.setFontColor('#FFFFFF');
      
      creadas.push(SHEET_LOGS);
      Logger.info('Hoja de logs creada', {}, 'inicializarSistemaLogs');
    }
    
    // Crear hoja de errores si no existe
    if (!hojaErrores) {
      hojaErrores = ss.insertSheet(SHEET_ERRORES);
      hojaErrores.appendRow([
        'Timestamp',
        'Código',
        'Mensaje',
        'Detalles',
        'Fuente'
      ]);
      
      // Formatear encabezados
      const encabezados = hojaErrores.getRange(1, 1, 1, 5);
      encabezados.setFontWeight('bold');
      encabezados.setBackground('#EA4335');
      encabezados.setFontColor('#FFFFFF');
      
      creadas.push(SHEET_ERRORES);
      Logger.info('Hoja de errores creada', {}, 'inicializarSistemaLogs');
    }
    
    // Ajustar anchos de columna
    if (hojaLogs) {
      hojaLogs.setColumnWidth(1, 150); // Timestamp
      hojaLogs.setColumnWidth(2, 80);  // Nivel
      hojaLogs.setColumnWidth(3, 300); // Mensaje
      hojaLogs.setColumnWidth(4, 200); // Detalles
      hojaLogs.setColumnWidth(5, 100); // Fuente
      hojaLogs.setColumnWidth(6, 150); // Usuario
    }
    
    if (hojaErrores) {
      hojaErrores.setColumnWidth(1, 150); // Timestamp
      hojaErrores.setColumnWidth(2, 100); // Código
      hojaErrores.setColumnWidth(3, 300); // Mensaje
      hojaErrores.setColumnWidth(4, 200); // Detalles
      hojaErrores.setColumnWidth(5, 100); // Fuente
    }
    
    const mensaje = creadas.length > 0 ? 
      `Sistema de logs inicializado. Hojas creadas: ${creadas.join(', ')}` :
      'Sistema de logs validado correctamente. Todas las hojas existen.';
    
    Logger.info(mensaje, { hojasCreadas: creadas }, 'inicializarSistemaLogs');
    
    return {
      success: true,
      message: mensaje,
      hojasCreadas: creadas,
      estado: {
        hojaLogs: !!hojaLogs,
        hojaErrores: !!hojaErrores
      }
    };
    
  } catch (error) {
    console.error('❌ Error inicializando sistema de logs:', error);
    return {
      success: false,
      message: 'Error al inicializar sistema de logs: ' + error.message,
      error: error.toString()
    };
  }
}

// ==========================================
// EXPORTAR FUNCIONES PARA TESTING
// ==========================================
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    Logger,
    registrarLog,
    registrarErrorLog,
    registrarEventoInfo,
    registrarEventoError,
    obtenerLogsRecientes,
    loadLogs,
    processLogsResult,
    limpiarLogs,
    exportarLogs,
    obtenerEstadisticasLogs,
    buscarLogs,
    monitorearErroresCriticos,
    inicializarSistemaLogs
  };
}