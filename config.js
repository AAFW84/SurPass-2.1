/**
 * Diagnóstico integral del sistema SurPass.
 * @returns {Object} Resultado del diagnóstico.
 */
function diagnosticarSistemaCompleto() {
  const diagnostico = {
    timestamp: new Date().toISOString(),
    tests: {},
    resumen: {},
    recomendaciones: [],
    success: false
  };

  // Test 1: Hojas necesarias
  try {
    const ss = SpreadsheetApp.getActive();
    const hojasNecesarias = [SHEET_REGISTRO, SHEET_BD, SHEET_LOGS, SHEET_CONFIG];
    const hojasEncontradas = hojasNecesarias.filter(nombre => ss.getSheetByName(nombre));
    diagnostico.tests.hojas = {
      success: hojasEncontradas.length === hojasNecesarias.length,
      encontradas: hojasEncontradas,
      faltantes: hojasNecesarias.filter(h => !hojasEncontradas.includes(h)),
      mensaje: `${hojasEncontradas.length}/${hojasNecesarias.length} hojas encontradas`
    };
    if (hojasEncontradas.length < hojasNecesarias.length) {
      diagnostico.recomendaciones.push('Ejecutar inicializarConfiguracionSistema()');
    }
  } catch (error) {
    diagnostico.tests.hojas = { success: false, mensaje: error.message };
  }

  // Test 2: Configuración
  try {
    const configResult = obtenerConfiguracion();
    diagnostico.tests.configuracion = {
      success: configResult.success,
      totalConfigs: configResult.success ? Object.keys(configResult.config).length : 0,
      mensaje: configResult.message || (configResult.success ? 'Configuración accesible' : 'Error en configuración')
    };
    if (!configResult.success) {
      diagnostico.recomendaciones.push('Ejecutar inicializarConfiguracionCompleta()');
    }
  } catch (error) {
    diagnostico.tests.configuracion = { success: false, mensaje: error.message };
  }

  // Test 3: Cache de usuarios
  try {
    const totalUsuarios = Object.keys(cacheUsuarios || {}).length;
    diagnostico.tests.usuarios = {
      success: totalUsuarios > 0,
      totalUsuarios,
      mensaje: totalUsuarios > 0 ? `${totalUsuarios} usuarios en cache` : 'Cache vacío'
    };
    if (totalUsuarios === 0) {
      diagnostico.recomendaciones.push('Ejecutar inicializarCacheUsuarios()');
    }
  } catch (error) {
    diagnostico.tests.usuarios = { success: false, mensaje: error.message };
  }

  // Test 4: Funciones críticas
  try {
    const funcionesCriticas = [
      'buscarUsuario', 'obtenerEstadisticas', 'registrarEnHoja',
      'inicializarCacheUsuarios', 'getColumnasRegistro'
    ];
    const funcionesDisponibles = funcionesCriticas.filter(fn =>
      typeof this[fn] === 'function' || (typeof global !== 'undefined' && typeof global[fn] === 'function')
    );
    diagnostico.tests.funciones = {
      success: funcionesDisponibles.length === funcionesCriticas.length,
      disponibles: funcionesDisponibles,
      faltantes: funcionesCriticas.filter(f => !funcionesDisponibles.includes(f)),
      mensaje: `${funcionesDisponibles.length}/${funcionesCriticas.length} funciones disponibles`
    };
    if (funcionesDisponibles.length < funcionesCriticas.length) {
      diagnostico.recomendaciones.push('Revisar implementación de funciones críticas');
    }
  } catch (error) {
    diagnostico.tests.funciones = { success: false, mensaje: error.message };
  }

  // Resumen global
  const testsRealizados = Object.keys(diagnostico.tests).length;
  const testsExitosos = Object.values(diagnostico.tests).filter(t => t.success).length;
  diagnostico.resumen = {
    testsRealizados,
    testsExitosos,
    porcentajeExito: Math.round((testsExitosos / testsRealizados) * 100),
    sistemaOperativo: testsExitosos === testsRealizados
  };
  diagnostico.success = testsExitosos === testsRealizados;

  return diagnostico;
}
/**
 * Inicializa la configuración completa del sistema, asegurando valores críticos.
 */
function inicializarConfiguracionCompleta() {
  try {
    console.log('Inicializando configuración completa del sistema...');
    // Primero, intentar obtener configuración existente
    const configResult = obtenerConfiguracion();
    let config = {};
    if (configResult.success) {
      config = configResult.config;
    }
    // Definir configuraciones críticas que deben existir
    const configuracionesCriticas = {
      EMPRESA_NOMBRE: config.EMPRESA_NOMBRE || 'SurPass',
      HORARIO_APERTURA: config.HORARIO_APERTURA || '07:00',
      HORARIO_CIERRE: config.HORARIO_CIERRE || '18:00',
      MODO_24_HORAS: config.MODO_24_HORAS || 'NO',
      DIAS_LABORABLES: config.DIAS_LABORABLES || 'LUN,MAR,MIÉ,JUE,VIE',
      ZONA_HORARIA: config.ZONA_HORARIA || 'America/Panama',
      SONIDOS_ACTIVADOS: config.SONIDOS_ACTIVADOS || 'SI',
      UI_COLOR_PRIMARIO: config.UI_COLOR_PRIMARIO || '#007bff',
      UI_COLOR_SECUNDARIO: config.UI_COLOR_SECUNDARIO || '#28a745',
      CACHE_DURACION_ESTADISTICAS: config.CACHE_DURACION_ESTADISTICAS || '180',
      CACHE_DURACION_COLUMNAS: config.CACHE_DURACION_COLUMNAS || '1800',
      CACHE_DURACION_CONFIG: config.CACHE_DURACION_CONFIG || '900',
      NOTIFICACIONES_EMAIL: config.NOTIFICACIONES_EMAIL || 'admin@surpass.com',
      NOTIFICAR_ACCESOS_DENEGADOS: config.NOTIFICAR_ACCESOS_DENEGADOS || 'SI',
      QR_CAMPO_VISION: config.QR_CAMPO_VISION || '85',
      QR_TIEMPO_ESCANEO: config.QR_TIEMPO_ESCANEO || '250',
      QR_MODO_CAMARA: config.QR_MODO_CAMARA || 'environment',
      QR_SONIDO_EXITO: config.QR_SONIDO_EXITO || 'SI',
      INTENTOS_MAX_LOGIN: config.INTENTOS_MAX_LOGIN || '5',
      TIEMPO_BLOQUEO_LOGIN: config.TIEMPO_BLOQUEO_LOGIN || '10',
      AUDIT_LOG_ACTIVIDAD: config.AUDIT_LOG_ACTIVIDAD || 'SI'
    };
    // Actualizar configuraciones faltantes
    const updateResult = actualizarConfiguracionMultiple(configuracionesCriticas);
    if (!updateResult.success) {
      console.error('Error actualizando configuraciones críticas:', updateResult.message);
      return { success: false, message: updateResult.message };
    }
    // Limpiar cache de configuración para forzar recarga
    if (SurPassCache) {
      SurPassCache.configuracion = null;
      SurPassCache.tiempoConfiguracion = null;
    }
    console.log('Configuración completa inicializada correctamente');
    return {
      success: true,
      message: 'Configuración inicializada correctamente',
      configuracionesActualizadas: Object.keys(configuracionesCriticas).length,
      lastSaved: updateResult.lastSaved
    };
  } catch (error) {
    console.error('Error inicializando configuración completa:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Valida y repara la configuración de horarios.
 */
function validarYRepararHorarios() {
  try {
    console.log('Validando configuración de horarios...');
    const configResult = obtenerConfiguracion();
    if (!configResult.success) {
      console.error('No se pudo obtener configuración para validar horarios');
      return { success: false, message: 'Error obteniendo configuración' };
    }
    const config = configResult.config;
    let reparacionesNecesarias = {};
    if (!config.HORARIO_APERTURA || config.HORARIO_APERTURA.trim() === '') {
      reparacionesNecesarias.HORARIO_APERTURA = '07:00';
    }
    if (!config.HORARIO_CIERRE || config.HORARIO_CIERRE.trim() === '') {
      reparacionesNecesarias.HORARIO_CIERRE = '18:00';
    }
    if (!config.MODO_24_HORAS || !['SI', 'NO', 'true', 'false'].includes(config.MODO_24_HORAS)) {
      reparacionesNecesarias.MODO_24_HORAS = 'NO';
    }
    if (!config.DIAS_LABORABLES || config.DIAS_LABORABLES.trim() === '') {
      reparacionesNecesarias.DIAS_LABORABLES = 'LUN,MAR,MIÉ,JUE,VIE';
    }
    if (Object.keys(reparacionesNecesarias).length > 0) {
      console.log('Aplicando reparaciones de horarios:', reparacionesNecesarias);
      const updateResult = actualizarConfiguracionMultiple(reparacionesNecesarias);
      if (!updateResult.success) {
        return { success: false, message: 'Error aplicando reparaciones: ' + updateResult.message };
      }
      return {
        success: true,
        message: 'Configuración de horarios reparada',
        reparaciones: reparacionesNecesarias,
        lastSaved: updateResult.lastSaved
      };
    }
    return {
      success: true,
      message: 'Configuración de horarios correcta',
      reparaciones: {}
    };
  } catch (error) {
    console.error('Error validando horarios:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Obtiene el horario laboral con valores por defecto y normalización.
 */
function obtenerHorarioLaboralSeguro() {
  try {
    const result = obtenerConfiguracion();
    let horario = {
      apertura: '07:00',
      cierre: '18:00',
      modo24h: false,
      diasLaborables: 'LUN,MAR,MIÉ,JUE,VIE'
    };
    if (result.success && result.config) {
      const config = result.config;
      if (config.HORARIO_APERTURA && config.HORARIO_APERTURA.trim() !== '') {
        horario.apertura = normalizarHora(config.HORARIO_APERTURA);
      }
      if (config.HORARIO_CIERRE && config.HORARIO_CIERRE.trim() !== '') {
        horario.cierre = normalizarHora(config.HORARIO_CIERRE);
      }
      if (config.MODO_24_HORAS) {
        horario.modo24h = normalizarModo24h(config.MODO_24_HORAS);
      }
      if (config.DIAS_LABORABLES && config.DIAS_LABORABLES.trim() !== '') {
        horario.diasLaborables = normalizarDiasLaborables(config.DIAS_LABORABLES);
      }
    } else {
      console.warn('No se pudo obtener configuración, usando valores por defecto para horario');
    }
    return horario;
  } catch (error) {
    console.error('Error obteniendo horario laboral:', error);
    return {
      apertura: '07:00',
      cierre: '18:00',
      modo24h: false,
      diasLaborables: 'LUN,MAR,MIÉ,JUE,VIE'
    };
  }
}

/**
 * Inicialización completa del sistema (debe ejecutarse al inicio).
 */
function ejecutarInicializacionCompleta() {
  try {
    console.log('=== INICIANDO CONFIGURACIÓN COMPLETA DEL SISTEMA ===');
    const resultados = {
      estructura: null,
      configuracion: null,
      horarios: null,
      cache: null,
      errores: []
    };
    // 1. Inicializar estructura de hojas
    console.log('1. Inicializando estructura de hojas...');
    try {
      resultados.estructura = inicializarConfiguracionSistema();
      if (!resultados.estructura.success) {
        resultados.errores.push('Error en estructura: ' + resultados.estructura.mensaje);
      }
    } catch (error) {
      resultados.errores.push('Error crítico en estructura: ' + error.message);
      resultados.estructura = { success: false, mensaje: error.message };
    }
    // 2. Inicializar configuraciones críticas
    console.log('2. Inicializando configuraciones críticas...');
    try {
      resultados.configuracion = inicializarConfiguracionCompleta();
      if (!resultados.configuracion.success) {
        resultados.errores.push('Error en configuración: ' + resultados.configuracion.message);
      }
    } catch (error) {
      resultados.errores.push('Error crítico en configuración: ' + error.message);
      resultados.configuracion = { success: false, message: error.message };
    }
    // 3. Validar y reparar horarios
    console.log('3. Validando configuración de horarios...');
    try {
      resultados.horarios = validarYRepararHorarios();
      if (!resultados.horarios.success) {
        resultados.errores.push('Error en horarios: ' + resultados.horarios.message);
      }
    } catch (error) {
      resultados.errores.push('Error crítico en horarios: ' + error.message);
      resultados.horarios = { success: false, message: error.message };
    }
    // 4. Inicializar cache de usuarios
    console.log('4. Inicializando cache de usuarios...');
    try {
      if (typeof inicializarCacheUsuarios === 'function') {
        inicializarCacheUsuarios();
        const totalUsuarios = Object.keys(cacheUsuarios || {}).length;
        resultados.cache = { 
          success: true, 
          message: `Cache inicializado con ${totalUsuarios} usuarios` 
        };
      } else {
        resultados.cache = { 
          success: false, 
          message: 'Función inicializarCacheUsuarios no encontrada' 
        };
        resultados.errores.push('Cache de usuarios no inicializado');
      }
    } catch (error) {
      resultados.errores.push('Error en cache: ' + error.message);
      resultados.cache = { success: false, message: error.message };
    }
    // Resumen final
    const exitoso = resultados.errores.length === 0;
    console.log(exitoso ? '✅ INICIALIZACIÓN COMPLETA EXITOSA' : '⚠️ INICIALIZACIÓN CON ERRORES');
    if (resultados.errores.length > 0) {
      console.warn('Errores encontrados:', resultados.errores);
    }
    return {
      success: exitoso,
      message: exitoso ? 'Sistema inicializado completamente' : 'Sistema inicializado con errores',
      resultados: resultados,
      erroresEncontrados: resultados.errores.length,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('ERROR CRÍTICO EN INICIALIZACIÓN:', error);
    return {
      success: false,
      message: 'Error crítico durante la inicialización: ' + error.message,
      error: error.stack,
      timestamp: new Date().toISOString()
    };
  }
}
/**
 * Verifica y repara los valores críticos en la hoja de configuración.
 * Si faltan claves esenciales, las agrega con valores por defecto.
 * Devuelve un resumen de la reparación.
 */
function repararConfiguracion() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
    if (!hoja) {
      return { success: false, message: 'Hoja de configuración no encontrada.' };
    }
    const datos = hoja.getDataRange().getValues();
    const clavesExistentes = new Set();
    for (let i = 1; i < datos.length; i++) {
      const key = String(datos[i][0]).trim();
      if (key) clavesExistentes.add(key);
    }
    // Claves críticas y sus valores por defecto
    const criticos = {
      HORARIO_APERTURA: '07:00',
      HORARIO_CIERRE: '18:00',
      MODO_24_HORAS: 'false',
      DIAS_LABORABLES: 'LUN,MAR,MIÉ,JUE,VIE'
    };
    let reparadas = [];
    Object.keys(criticos).forEach(clave => {
      if (!clavesExistentes.has(clave)) {
        hoja.appendRow([clave, criticos[clave], 'sistema']);
        reparadas.push(clave);
      }
    });
    if (reparadas.length > 0) {
      return { success: true, message: 'Configuración reparada', reparadas };
    } else {
      return { success: true, message: 'Configuración ya estaba completa', reparadas: [] };
    }
  } catch (e) {
    return { success: false, message: 'Error reparando configuración: ' + e.message };
  }
}
// config.js - Funciones para la gestión de configuración del sistema SurPass V6.2

// =================================================================
// SISTEMA DE GESTIÓN DE CONFIGURACIÓN CENTRALIZADO
// Este archivo gestiona la lectura, escritura y diagnóstico de la
// configuración del sistema, que se almacena en una hoja de Google
// Sheets para fácil acceso y modificación.
// =================================================================

/**
 * Función auxiliar para determinar la categoría de una clave de configuración.
 * Centraliza la lógica de categorización para mantener la consistencia.
 * @param {string} key - La clave de la configuración.
 * @returns {string} La categoría a la que pertenece la clave.
 */
function determinarCategoria(key) {
  // Normalizar y proteger ante valores no string para evitar TypeError
  const k = (typeof key === 'string') ? key : (key == null ? '' : String(key));
  if (!k) return 'GENERAL';
  if (k.startsWith('EMPRESA_')) return 'EMPRESA';
  if (k === 'HORARIO_APERTURA' || k === 'HORARIO_CIERRE' || k === 'MODO_24_HORAS' || k === 'DIAS_LABORABLES' || k === 'ZONA_HORARIA' || k === 'FORMATO_FECHA_SISTEMA') return 'EMPRESA';
  if (k.startsWith('UI_')) return 'INTERFAZ';
  if (k.startsWith('QR_')) return 'CAMARA_QR';
  if (k.includes('CONTRASEÑA') || k.includes('SESION') || k.includes('BLOQUEO') || k.includes('AUDIT') || k.includes('2FA') || k === 'INTENTOS_MAX_LOGIN' || k === 'TIEMPO_BLOQUEO_LOGIN') return 'SEGURIDAD';
  if (k.startsWith('NOTIF') || k.includes('EMAIL') || k === 'SONIDOS_ACTIVADOS' || k.startsWith('WEBHOOK_')) return 'NOTIFICACIONES';
  if (k.startsWith('CACHE_')) return 'CACHE';
  if (k.startsWith('REPORTE_') || k.startsWith('ESTADISTICAS_')) return 'REPORTES';
  if (k.startsWith('EVACUACION_')) return 'EVACUACION';
  if (k.includes('BACKUP') || k.includes('ARCHIVO') || k.includes('LIMPIAR') || k.includes('OPTIMIZAR') || k === 'DIAS_RETENER_LOGS') return 'ARCHIVO';
  if (k.startsWith('TEMA_') || k.includes('IDIOMA') || k.includes('MOSTRAR') || k.includes('DASHBOARD') || k.includes('SONIDO_') || k.includes('VIBRACIÓN') || k === 'FORMATO_HORA_12H' || k === 'REGISTRO_ACTIVIDAD_ADMINS' || k === 'LIMPIAR_CACHE_STARTUP') return 'PERSONALIZACION';
  if (k.startsWith('INTEGRACION_')) return 'INTEGRACION';
  
  return 'GENERAL';
}


/**
 * Obtiene toda la configuración del sistema desde la hoja "Configuración".
 * Asume que la hoja tiene una estructura de dos columnas: Clave | Valor.
 *
 * @returns {Object} Un objeto con la configuración actual.
 * Retorna { success: true, config: Object } en caso de éxito,
 * o { success: false, message: string } en caso de error.
 */
function obtenerConfiguracion() {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
    if (!hoja) {
      // Si la hoja no existe, crea una básica con configuración completa
      const ss = SpreadsheetApp.getActive();
      const newSheet = ss.insertSheet(SHEET_CONFIG);
      newSheet.appendRow(['Clave', 'Valor', 'Categoría']);
      newSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
      newSheet.setColumnWidth(1, 250); // Ancho para la clave
      newSheet.setColumnWidth(2, 300); // Ancho para el valor
      newSheet.setColumnWidth(3, 200); // Ancho para la categoría
      
      // Añadir configuración completa por defecto
      const configDefault = obtenerConfiguracionPorDefecto();
      const configEntries = [];
      
      Object.keys(configDefault).forEach(key => {
        const categoria = determinarCategoria(key); // Usar la nueva función
        configEntries.push([key, configDefault[key], categoria]);
      });
      
      if (configEntries.length > 0) {
        newSheet.getRange(2, 1, configEntries.length, 3).setValues(configEntries);
      }
      
      console.log(`✅ Hoja de configuración "${SHEET_CONFIG}" creada con ${configEntries.length} configuraciones organizadas por categoría.`);
      
      // Después de crearla y poblarla, la recargamos para obtener los valores
      return obtenerConfiguracion(); 
    }

    const datos = hoja.getDataRange().getValues();
    const config = {};

        // La primera fila son encabezados, configuración empieza desde la segunda fila (índice 1)
    // Columna 0 (A): Clave, Columna 1 (B): Valor, Columna 2 (C): Categoría (opcional)
    for (let i = 1; i < datos.length; i++) {
      const key = String(datos[i][0]).trim();
      if (key) { // Asegurarse de que la clave no esté vacía
        let value = datos[i][1];
        // Convertir todos los valores a texto para asegurar compatibilidad
        config[key] = String(value);
      }
    }
    
    console.log(`✅ Configuración obtenida: ${Object.keys(config).length} parámetros cargados`);
    return { success: true, config: config };
    
  } catch (error) {
    console.error("❌ Error al obtener configuración:", error);
    if (typeof registrarLog === 'function') {
      registrarLog('ERROR', 'Error al obtener configuración', { error: error.message, sheet: SHEET_CONFIG });
    }
    return { success: false, message: error.message };
  }
}


/**
 * Actualiza múltiples configuraciones en la hoja "Configuración".
 * Si una clave no existe en la hoja, la añade como una nueva fila.
 *
 * @param {Object} newConfig - Un objeto con las nuevas configuraciones (clave: valor).
 * @returns {Object} Resultado de la operación, incluyendo la fecha/hora del último guardado.
 */
function actualizarConfiguracionMultiple(newConfig) {
  try {
    // Validación defensiva de entrada
    if (!newConfig || typeof newConfig !== 'object') {
      const msg = 'Parámetro newConfig inválido: se esperaba un objeto no nulo.';
      console.error('❌', msg, 'Valor recibido:', newConfig);
      if (typeof registrarLog === 'function') {
        registrarLog('ERROR', 'actualizarConfiguracionMultiple: parámetro inválido', { recibido: String(newConfig) });
      }
      return { success: false, message: msg };
    }

    const keysToUpdate = Object.keys(newConfig);
    console.log('💾 Actualizando configuración múltiple:', keysToUpdate);
    
    let hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
    if (!hoja) {
      console.error(`❌ Hoja de configuración "${SHEET_CONFIG}" no encontrada.`);
      return { success: false, message: 'Hoja de configuración no encontrada' };
    }

    const datos = hoja.getDataRange().getValues();
    const keysInSheet = {}; // Mapa para buscar rápidamente las claves existentes y sus números de fila

    // Construir el mapa de claves existentes para optimizar la búsqueda
    for (let i = 1; i < datos.length; i++) {
      const key = String(datos[i][0]).trim();
      if (key) {
        keysInSheet[key] = i + 1; // Guarda el número de fila (1-indexed)
      }
    }

    const updates = []; // Array para almacenar las actualizaciones a realizar
    const newRows = [];   // Array para almacenar las nuevas filas a añadir

    let updatedCount = 0;
    let appendedCount = 0;

    for (const key in newConfig) {
      if (newConfig.hasOwnProperty(key)) {
        // Normalizar la clave; omitir claves vacías o nulas
        const keyStr = (typeof key === 'string') ? key.trim() : String(key).trim();
        if (!keyStr) {
          continue;
        }
        let valueToSave = newConfig[key];

        // Convertir valores booleanos a 'SI' o 'NO' para el almacenamiento
        if (typeof valueToSave === 'boolean') {
          valueToSave = valueToSave ? 'SI' : 'NO';
        }
        // Normalizar null/undefined a cadena vacía para evitar errores en setValue/setValues
        if (valueToSave == null) {
          valueToSave = '';
        }

        const rowNum = keysInSheet[keyStr];
        if (rowNum) {
          // Si la clave ya existe, prepara la actualización
          updates.push({ row: rowNum, col: 2, value: valueToSave }); // Col 2 es la columna de valores
          updatedCount++;
        } else {
          // Si la clave no existe, prepara una nueva fila
          const categoria = determinarCategoria(keyStr); // Usar la nueva función
          newRows.push([keyStr, valueToSave, categoria]);
          appendedCount++;
        }
      }
    }

    // Realizar todas las actualizaciones de celdas existentes
    updates.forEach(item => {
      hoja.getRange(item.row, item.col).setValue(item.value);
    });

    // Añadir todas las nuevas filas
    if (newRows.length > 0) {
      const startRow = hoja.getLastRow() + 1;
      const numRows = newRows.length;
      const numCols = newRows[0].length;
      hoja.getRange(startRow, 1, numRows, numCols).setValues(newRows);
    }

    const lastSaved = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    console.log(`✅ Configuración actualizada. Actualizadas: ${updatedCount}, Añadidas: ${appendedCount}. Último guardado: ${lastSaved}`);

    if (typeof registrarLog === 'function') {
      registrarLog('INFO', 'Configuración del sistema actualizada', { 
        updated: updatedCount, 
        appended: appendedCount, 
        configKeys: Object.keys(newConfig),
        lastSaved: lastSaved
      });
    }

    return { 
      success: true, 
      message: 'Configuración guardada correctamente', 
      lastSaved: lastSaved 
    };
    
  } catch (error) {
    console.error("❌ Error al actualizar configuración:", error);
    if (typeof registrarLog === 'function') {
      registrarLog('ERROR', 'Error al actualizar configuración', { 
        error: error.message, 
        config: newConfig, 
        sheet: SHEET_CONFIG 
      });
    }
    return { success: false, message: error.message };
  }
}


/**
 * Obtiene configuración por defecto completa del sistema SurPass V6.2
 * @returns {Object} Configuración por defecto con 94 parámetros organizados por categorías
 */
function obtenerConfiguracionPorDefecto() {
  return {
    // === CONFIGURACIONES DE EMPRESA (11 parámetros) ===
    EMPRESA_NOMBRE: 'SurPass',
    EMPRESA_LOGO: '',
    EMPRESA_DIRECCION: '',
    EMPRESA_TELEFONO: '',
    EMPRESA_EMAIL: 'contacto@surpass.com',
    HORARIO_APERTURA: '', // Sin valor por defecto: depender de hoja
    HORARIO_CIERRE: '',   // Sin valor por defecto: depender de hoja
    MODO_24_HORAS: '',    // Sin valor por defecto: depender de hoja
    DIAS_LABORABLES: '',  // Sin valor por defecto: depender de hoja
    ZONA_HORARIA: 'America/Panama',
    FORMATO_FECHA_SISTEMA: 'dd/MM/yyyy',

    // === CONFIGURACIONES DE INTERFAZ DE USUARIO (10 parámetros) ===
    UI_TITULO_PRINCIPAL: '70px',
    UI_SUBTITULO_PRINCIPAL: '35px',
    UI_COLOR_PRIMARIO: '#007bff',
    UI_COLOR_SECUNDARIO: '#28a745',
    UI_COLOR_PELIGRO: '#dc3545',
    UI_COLOR_ADVERTENCIA: '#ffc107',
    UI_RADIUS_BORDES: '15px',
    UI_SOMBRA_TARJETAS: '0 8px 32px rgba(0,0,0,0.15)',
    UI_TAMAÑO_MODAL_TITULO: '49px',
    UI_TAMAÑO_ICONO: '38.5px',

    // === CONFIGURACIONES DE CÁMARA QR (8 parámetros) ===
    QR_CAMPO_VISION: '85',
    QR_TIEMPO_ESCANEO: '250',
    QR_MODO_CAMARA: 'environment',
    QR_AUTOCLOSE_DELAY: '3000',
    QR_SONIDO_EXITO: 'SI',
    QR_SONIDO_ERROR: 'SI',
    QR_MOSTRAR_PUNTOS_REFERENCIA: 'SI',
    QR_RESOLUCION_PREFERIDA: 'HD',

    // === CONFIGURACIONES DE SEGURIDAD (11 parámetros) ===
    INTENTOS_MAX_LOGIN: '5',
    TIEMPO_BLOQUEO_LOGIN: '10',
    SESION_DURACION_MAXIMA: '480',
    BLOQUEO_IP_INTENTOS: '10',
    BLOQUEO_IP_DURACION: '60',
    AUDIT_LOG_ACTIVIDAD: 'SI',
    REQUIERE_2FA_ADMIN: 'NO',
    CONTRASEÑA_MIN_LONGITUD: '6',
    CONTRASEÑA_REQUIERE_NUMEROS: 'SI',
    CONTRASEÑA_REQUIERE_MAYUSCULAS: 'NO',
    CONTRASEÑA_REQUIERE_SIMBOLOS: 'NO',

    // === CONFIGURACIONES DE NOTIFICACIONES (12 parámetros) ===
    NOTIFICACIONES_EMAIL: 'admin@surpass.com',
    NOTIFICAR_ACCESOS_DENEGADOS: 'SI',
    NOTIFICACIONES_EMAIL_ERRORES_CRITICOS: 'SI',
    SONIDOS_ACTIVADOS: 'SI',
    NOTIF_WEBHOOK_URL: '',
    NOTIF_SLACK_TOKEN: '',
    NOTIF_TEAMS_URL: '',
    NOTIF_TELEGRAM_BOT: '',
    NOTIF_TELEGRAM_CHAT: '',
    NOTIF_EMAIL_SMTP_HOST: 'smtp.gmail.com',
    NOTIF_EMAIL_SMTP_PORT: '587',
    WEBHOOK_ENTRADA_URL: '',
    WEBHOOK_SALIDA_URL: '',
    WEBHOOK_ERROR_URL: '',

    // === CONFIGURACIONES DE CACHE Y RENDIMIENTO (4 parámetros) ===
    CACHE_DURACION_ESTADISTICAS: '180',
    CACHE_DURACION_COLUMNAS: '1800',
    CACHE_DURACION_CONFIG: '900',
    CACHE_DURACION_USUARIOS: '600',

    // === CONFIGURACIONES DE REPORTES (7 parámetros) ===
    REPORTE_FORMATO_FECHA: 'dd/MM/yyyy',
    REPORTE_INCLUIR_LOGO: 'SI',
    REPORTE_MAX_REGISTROS: '10000',
    REPORTE_AUTOEXPORT_PDF: 'NO',
    REPORTE_TIMEZONE: 'America/Panama',
    ESTADISTICAS_ACTUALIZAR_AUTO: 'SI',
    ESTADISTICAS_INTERVALO: '300',

    // === CONFIGURACIONES DE EVACUACIÓN (7 parámetros) ===
    EVACUACION_SONIDO_ALARMA: 'SI',
    EVACUACION_DURACION_ALARMA: '120',
    EVACUACION_AUTO_REGISTRO: 'SI',
    EVACUACION_NOTIFICAR_RESPONSABLES: 'SI',
    EVACUACION_MOSTRAR_CONTADOR: 'SI',
    EVACUACION_EXPORTAR_LISTA: 'SI',
    EVACUACION_REINICIO_AUTOMATICO: 'NO',

    // === CONFIGURACIONES DE ARCHIVO Y LIMPIEZA (6 parámetros) ===
    BACKUP_AUTOMATICO: 'NO',
    FRECUENCIA_BACKUP: 'SEMANAL',
    LIMPIAR_LOGS_AUTOMATICO: 'SI',
    DIAS_RETENER_LOGS: '90',
    ARCHIVO_AUTO_REGISTROS: 'NO',
    ARCHIVO_DIAS_ANTIGÜEDAD: '365',

    // === CONFIGURACIONES DE PERSONALIZACIÓN (10 parámetros) ===
    TEMA_OSCURO_DISPONIBLE: 'NO',
    IDIOMA_INTERFAZ: 'es',
    FORMATO_HORA_12H: 'NO',
    MOSTRAR_AYUDA_CONTEXTUAL: 'SI',
    MOSTRAR_TIPS_STARTUP: 'SI',
    SONIDO_NOTIFICACIONES: 'SI',
    DASHBOARD_AUTO_REFRESH: 'SI',
    DASHBOARD_REFRESH_INTERVALO: '30',
    REGISTRO_ACTIVIDAD_ADMINS: 'SI',
    LIMPIAR_CACHE_STARTUP: 'NO',

    // === CONFIGURACIONES DE INTEGRACIÓN (3 parámetros) ===
    INTEGRACION_API_ACTIVA: 'NO',
    INTEGRACION_API_URL: '',
    INTEGRACION_API_TOKEN: '',
  };
}


/**
 * Función auxiliar para obtener configuración por categoría
 * @param {string} categoria - La categoría a filtrar
 * @returns {Object} Configuraciones de la categoría específica
 */
function obtenerConfiguracionPorCategoria(categoria) {
  try {
    const result = obtenerConfiguracion();
    if (!result.success) {
      return result;
    }
    
    const config = result.config;
    const configCategoria = {};
    
    Object.keys(config).forEach(key => {
      const keyCategoria = determinarCategoria(key); // Usar la nueva función
      if (keyCategoria === categoria) {
        configCategoria[key] = config[key];
      }
    });
    
    return {
      success: true,
      config: configCategoria,
      categoria: categoria,
      total: Object.keys(configCategoria).length
    };
    
  } catch (error) {
    console.error(`❌ Error obteniendo configuración por categoría ${categoria}:`, error);
    return { success: false, message: error.message };
  }
}


/**
 * Función de diagnóstico del sistema de configuración
 * @returns {Object} Resultado del diagnóstico completo
 */
function diagnosticarSistemaConfiguracion() {
  console.log('🔧 [CONFIG] Ejecutando diagnóstico del sistema de configuración...');
  
  try {
    const inicio = Date.now();
    const diagnostico = {
      timestamp: new Date().toISOString(),
      tests: {},
      resumen: {}
    };
    
    // Test 1: Verificar acceso a la hoja de configuración
    try {
      const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
      diagnostico.tests.hojaAcceso = {
        success: !!hoja,
        mensaje: hoja ? 'Hoja de configuración accesible' : 'Hoja de configuración no encontrada',
        filas: hoja ? hoja.getLastRow() : 0,
        columnas: hoja ? hoja.getLastColumn() : 0
      };
    } catch (error) {
      diagnostico.tests.hojaAcceso = {
        success: false,
        mensaje: 'Error accediendo a la hoja: ' + error.message
      };
    }
    
    // Test 2: Cargar configuración
    try {
      const configResult = obtenerConfiguracion();
      diagnostico.tests.cargarConfig = {
        success: configResult.success,
        mensaje: configResult.message || (configResult.success ? 'Configuración cargada correctamente' : 'Error cargando configuración'),
        totalConfiguraciones: configResult.success ? Object.keys(configResult.config).length : 0
      };
    } catch (error) {
      diagnostico.tests.cargarConfig = {
        success: false,
        mensaje: 'Error cargando configuración: ' + error.message
      };
    }
    
    // Test 3: Verificar configuración por defecto
    try {
      const configDefault = obtenerConfiguracionPorDefecto();
      diagnostico.tests.configDefecto = {
        success: true,
        mensaje: 'Configuración por defecto disponible',
        totalDefecto: Object.keys(configDefault).length
      };
    } catch (error) {
      diagnostico.tests.configDefecto = {
        success: false,
        mensaje: 'Error obteniendo configuración por defecto: ' + error.message
      };
    }
    
    // Test 4: Test de escritura (solo si las anteriores pasaron)
    if (diagnostico.tests.hojaAcceso.success && diagnostico.tests.cargarConfig.success) {
      try {
        const testConfig = { 'TEST_DIAGNOSTICO': 'EJECUTADO_' + Date.now() };
        const updateResult = actualizarConfiguracionMultiple(testConfig);
        diagnostico.tests.escritura = {
          success: updateResult.success,
          mensaje: updateResult.message || (updateResult.success ? 'Escritura funcionando' : 'Error en escritura'),
          lastSaved: updateResult.lastSaved
        };
      } catch (error) {
        diagnostico.tests.escritura = {
          success: false,
          mensaje: 'Error en test de escritura: ' + error.message
        };
      }
    }
    
    // Calcular resumen
    const testsRealizados = Object.keys(diagnostico.tests).length;
    const testsExitosos = Object.values(diagnostico.tests).filter(t => t.success).length;
    const tiempoTotal = Date.now() - inicio;
    
    diagnostico.resumen = {
      testsRealizados: testsRealizados,
      testsExitosos: testsExitosos,
      porcentajeExito: Math.round((testsExitosos / testsRealizados) * 100),
      tiempoEjecucion: tiempoTotal + 'ms',
      sistemaOperativo: testsExitosos === testsRealizados
    };
    
    const icono = diagnostico.resumen.sistemaOperativo ? '✅' : '❌';
    console.log(`${icono} [CONFIG] Diagnóstico completado: ${testsExitosos}/${testsRealizados} tests pasados (${diagnostico.resumen.porcentajeExito}%)`);
    
    return diagnostico;
    
  } catch (error) {
    console.error('❌ [CONFIG] Error crítico en diagnóstico:', error);
    return {
      success: false,
      mensaje: 'Error crítico ejecutando diagnóstico: ' + error.message,
      error: error.stack
    };
  }
}


/**
 * Función para obtener resumen completo del sistema de configuración
 * @returns {Object} Estadísticas y resumen del sistema
 */
function obtenerEstadisticasConfiguracion() {
  try {
    console.log('📊 [CONFIG] Generando estadísticas del sistema...');
    
    const configResult = obtenerConfiguracion();
    if (!configResult.success) {
      return { success: false, message: 'Error obteniendo configuración para estadísticas' };
    }
    
    const config = configResult.config;
    const categorias = {};
    const tiposValores = { booleanos: 0, numericos: 0, textos: 0, vacios: 0 };
    
    // Analizar configuraciones por categoría y tipo
    Object.keys(config).forEach(key => {
      const valor = config[key];
      
      // Determinar categoría
      const categoria = determinarCategoria(key); // Usar la nueva función
      
      // Contar por categoría
      if (!categorias[categoria]) categorias[categoria] = 0;
      categorias[categoria]++;
      
      // Analizar tipo de valor
      if (valor === '' || valor === null || valor === undefined) {
        tiposValores.vacios++;
      } else if (typeof valor === 'boolean' || valor === 'SI' || valor === 'NO') {
        tiposValores.booleanos++;
      } else if (!isNaN(valor) && valor !== '') {
        tiposValores.numericos++;
      } else {
        tiposValores.textos++;
      }
    });
    
    const estadisticas = {
      success: true,
      timestamp: new Date().toISOString(),
      totalConfiguraciones: Object.keys(config).length,
      categorias: {
        total: Object.keys(categorias).length,
        detalle: categorias
      },
      tiposValores: tiposValores,
      cobertura: {
        configuradas: Object.keys(config).length - tiposValores.vacios,
        vacias: tiposValores.vacios,
        porcentaje: Math.round(((Object.keys(config).length - tiposValores.vacios) / Object.keys(config).length) * 100)
      }
    };
    
    console.log(`📈 [CONFIG] Estadísticas: ${estadisticas.totalConfiguraciones} configs en ${estadisticas.categorias.total} categorías (${estadisticas.cobertura.porcentaje}% configurado)`);
    
    return estadisticas;
    
  } catch (error) {
    console.error('❌ [CONFIG] Error generando estadísticas:', error);
    return { success: false, message: error.message };
  }
}


/**
 * Función de verificación para confirmar el total de configuraciones
 * @returns {Object} Desglose por categoría y total
 */
function verificarTotalConfiguraciones() {
  try {
    const config = obtenerConfiguracionPorDefecto();
    const categorias = {
      EMPRESA: 0,
      INTERFAZ: 0,
      CAMARA_QR: 0,
      SEGURIDAD: 0,
      NOTIFICACIONES: 0,
      CACHE: 0,
      REPORTES: 0,
      EVACUACION: 0,
      ARCHIVO: 0,
      PERSONALIZACION: 0,
      INTEGRACION: 0
    };
    
    Object.keys(config).forEach(key => {
      const categoria = determinarCategoria(key); // Usar la nueva función
      if (categorias[categoria] !== undefined) {
        categorias[categoria]++;
      }
    });
    
    const total = Object.values(categorias).reduce((sum, count) => sum + count, 0);
    
    console.log('📊 Desglose de configuraciones por categoría:');
    Object.keys(categorias).forEach(cat => {
      console.log(`  ${cat}: ${categorias[cat]} configuraciones`);
    });
    console.log(`📋 TOTAL: ${total} configuraciones`);
    
    return {
      success: true,
      total: total,
      categorias: categorias,
      configuraciones: Object.keys(config)
    };
    
  } catch (error) {
    console.error('❌ Error verificando configuraciones:', error);
    return { success: false, error: error.message };
  }
}

console.log('✅ Config.js SurPass V6.2 cargado correctamente - Sistema de configuración ÚNICO y COMPLETO con 94 parámetros');

// =============================================================
// INICIALIZACIÓN DE SEGURIDAD
// =============================================================
/**
 * Inicializa la clave maestra de administrador en PropertiesService si aún no existe.
 * @returns {Object} Resultado de la operación.
 */
function inicializarClaveMaestra() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const existingKey = scriptProperties.getProperty('ADMIN_KEY');

    if (!existingKey) {
      const newKey = generarClaveAleatoria();
      scriptProperties.setProperty('ADMIN_KEY', newKey);
      console.log('🔐 Clave maestra ADMIN_KEY inicializada en PropertiesService.');
      if (typeof registrarLog === 'function') {
        try {
          registrarLog('INFO', 'Clave maestra ADMIN_KEY inicializada por defecto.', { clave: '********' }, 'Sistema de Seguridad');
        } catch (eLog) {
          console.warn('⚠️ Error registrando inicialización de clave maestra:', eLog);
        }
      }
      return { success: true, message: `Clave maestra inicializada.`, newKey: newKey };
    } else {
      console.log('🔐 Clave maestra ADMIN_KEY ya existe en PropertiesService.');
      return { success: true, message: 'Clave maestra ya configurada.' };
    }
  } catch (error) {
    console.error('❌ Error en inicializarClaveMaestra:', error);
    return { success: false, message: error.message };
  }
}

function generarClaveAleatoria(length = 12) {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*';
  let clave = '';
  for (let i = 0; i < length; i++) {
    clave += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return clave;
}

function obtenerClaveMaestraInicial() {
    const scriptProperties = PropertiesService.getScriptProperties();
    const existingKey = scriptProperties.getProperty('ADMIN_KEY');
    return existingKey;
}

/**
 * Inicializa una sal criptográfica segura en PropertiesService.
 * EJECUTAR ESTA FUNCIÓN UNA SOLA VEZ DESDE EL EDITOR DE APPS SCRIPT.
 */
function _inicializarSalDeSeguridad() {
  const scriptProperties = PropertiesService.getScriptProperties();
  if (!scriptProperties.getProperty('SURPASS_SECRET_SALT')) {
    const salt = Utilities.getUuid() + '-' + Utilities.getUuid();
    scriptProperties.setProperty('SURPASS_SECRET_SALT', salt);
    console.log('✅ Sal de seguridad inicializada correctamente. Guarde este valor en un lugar seguro y no lo cambie.');
  } else {
    console.log('ℹ️ La sal de seguridad ya existe. No se ha realizado ningún cambio.');
  }
}
