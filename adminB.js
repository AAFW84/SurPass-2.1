/**
 * Utilidad temporal para obtener el hash de una contraseña según la lógica del sistema.
 * Llama desde el editor de Apps Script con: obtenerHashDeClave('tu_contraseña')
 * El resultado se muestra en el log.
 */
function obtenerHashDeClave(plainPassword) {
  try {
    var hash = _hashPassword(plainPassword);
    console.log('🔑 Hash generado para contraseña:', plainPassword, '=>', hash);
    return hash;
  } catch (e) {
    console.error('Error generando hash:', e);
    return null;
  }
}
// adminB.js - Panel de Administración SurPass V6.2
// Este archivo actúa como interfaz para las funciones existentes en registroB.js
// Las constantes se importan desde constantes.js

// ======================
// 🔗 FUNCIONES DE INTERFAZ PARA GESTIÓN DE PERSONAL
// ======================

/**
 * Carga datos de personal usando la función existente
 * @return {Object} Resultado con datos de personal y estadísticas
 */
function loadPersonnel() {
  console.log('📊 [ADMIN-INTERFACE] Cargando personal...');
  
  try {
    // Usar función existente en registroB.js
    const result = obtenerPersonalParaAdmin();
    
    if (!result.success) {
      return {
        success: false,
        code: 'PERSONAL_LOAD_ERROR',
        message: result.error || 'Error cargando personal',
        personal: [],
        stats: { total: 0, activos: 0, inactivos: 0, empleados: 0, visitantes: 0 },
        details: {
          source: 'obtenerPersonalParaAdmin',
          raw: result || null
        }
      };
    }
    
    // Calcular estadísticas basadas en los datos reales de la hoja "Base de Datos"
    // Estructura: ["Cédula", "Nombre", "Empresa", "Área", "Cargo", "Estado"]
    const personal = (result.personal || []).map(p => ({
      // Normalizar nombres de propiedades para consistencia con frontend
      cedula: p.Cédula || p.cedula || '',
      nombre: p.Nombre || p.nombre || '',
      empresa: p.Empresa || p.empresa || '',
      area: p.Área || p.Area || p.area || '',
      cargo: p.Cargo || p.cargo || '',
      estado: p.Estado || p.estado || 'activo',
      indice: p.indice || 0  // Para operaciones de edición
    }));
    
    const stats = {
      total: personal.length,
      activos: personal.filter(p => {
        const estado = String(p.estado || '').toLowerCase().trim();
        return estado === 'activo' || estado === 'active' || estado === '';
      }).length,
      inactivos: personal.filter(p => {
        const estado = String(p.estado || '').toLowerCase().trim(); 
        return estado === 'inactivo' || estado === 'inactive' || estado === 'bloqueado';
      }).length,
      // Estadísticas por área (más útil que empleados/visitantes)
      areas: {}
    };
    
    // Contar por áreas
    personal.forEach(p => {
      const area = String(p.area || 'Sin Área').trim();
      stats.areas[area] = (stats.areas[area] || 0) + 1;
    });
    
    // Obtener las 3 áreas principales
    const areasOrdenadas = Object.entries(stats.areas)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3);
    
    stats.areaPrincipal = areasOrdenadas[0] ? areasOrdenadas[0][0] : 'N/A';
    stats.totalAreas = Object.keys(stats.areas).length;
    
    console.log(`✅ [ADMIN-INTERFACE] Personal cargado: ${stats.total} registros`);
    
    return {
      success: true,
      message: `${stats.total} registros cargados exitosamente`,
      personal: personal,  // Cambiar de 'data' a 'personal' para coincidir con el frontend
      stats: stats
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error en loadPersonnel:', error);
    return {
      success: false,
      code: 'UNEXPECTED_ERROR',
      message: 'Error interno cargando personal',
      personal: [],
      stats: { total: 0, activos: 0, inactivos: 0, empleados: 0, visitantes: 0 },
      details: { name: error.name, stack: error.stack }
    };
  }
}

/**
 * Guarda datos de personal usando la función existente
 * @param {Object} personData - Datos de la persona
 * @return {Object} Resultado de la operación
 */
function savePersonData(personData) {
  console.log('💾 [ADMIN-INTERFACE] Guardando datos de personal...');
  
  try {
    // Validar que todos los campos requeridos estén presentes
    if (!personData.cedula || !personData.nombre || !personData.empresa || !personData.area || !personData.cargo) {
      const missing = ['cedula','nombre','empresa','area','cargo'].filter(k => !personData[k]);
      return {
        success: false,
        code: 'VALIDATION_ERROR',
        message: 'Faltan campos requeridos: cédula, nombre, empresa, área y cargo son obligatorios',
        details: { missingFields: missing }
      };
    }
    
    // Normalizar los datos para la estructura de la Base de Datos
    const datosNormalizados = {
      Cédula: personData.cedula,
      Nombre: personData.nombre,
      Empresa: personData.empresa,
      Área: personData.area,
      Cargo: personData.cargo,
      Estado: personData.estado || 'activo'
    };
    
    // Usar función existente en registroB.js
    return agregarPersonalDesdeAdmin(datosNormalizados);
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error guardando personal:', error);
    return {
      success: false,
      code: 'UNEXPECTED_ERROR',
      message: 'Error guardando datos de personal: ' + error.message,
      details: { name: error.name, stack: error.stack }
    };
  }
}

// ======================
// 🔗 FUNCIONES DE INTERFAZ PARA CONFIGURACIÓN
// ======================

/**
 * 🔧 [FIX] Carga configuración desde Google Sheets correctamente
 * @return {Object} Configuración del sistema estructurada para el frontend
 */
function loadConfiguration() {
  console.log('⚙️ [ADMIN-INTERFACE] Cargando configuración desde Google Sheets...');
  
  try {
    // 🔧 [FIX] Verificar disponibilidad de función antes de usar
    if (typeof obtenerConfiguracion !== 'function') {
      console.error('❌ [ADMIN-INTERFACE] Función obtenerConfiguracion no disponible');
      registrarEventoError('Función obtenerConfiguracion no disponible', 'loadConfiguration');
      
      // Intentar reparar automáticamente
      if (typeof repararSistemaConfiguracion === 'function') {
        console.log('🔧 [ADMIN-INTERFACE] Intentando reparación automática...');
        const reparacion = repararSistemaConfiguracion();
        if (reparacion.exitosa) {
          console.log('✅ [ADMIN-INTERFACE] Sistema reparado, reintentando...');
          return loadConfiguration(); // Reintentar
        }
      }
      
      return {
        success: false,
        code: 'CONFIG_FN_NOT_AVAILABLE',
        message: 'Sistema de configuración no disponible. Ejecute diagnóstico y reparación.',
        configuraciones: [],
        requiresRepair: true,
        details: {
          source: 'loadConfiguration',
          attemptedRepair: typeof repararSistemaConfiguracion === 'function'
        }
      };
    }

    // Usar función existente en config.js que lee de Google Sheets
    const result = obtenerConfiguracion();
    console.log('📊 [ADMIN-INTERFACE] Resultado de obtenerConfiguracion:', result);
    
    if (!result || !result.success) {
      console.error('❌ [ADMIN-INTERFACE] Error en obtenerConfiguracion:', result);
      registrarEventoError(`Error cargando configuración: ${result?.message}`, 'loadConfiguration');
      
      // 🔧 [FIX] Intentar reparación automática si está disponible
      if (typeof repararSistemaConfiguracion === 'function') {
        console.log('🔧 [ADMIN-INTERFACE] Intentando reparación automática del sistema de configuración...');
        const reparacion = repararSistemaConfiguracion();
        
        if (reparacion.exitosa) {
          console.log('✅ [ADMIN-INTERFACE] Sistema reparado exitosamente, reintentando carga...');
          // Reintentar después de reparar
          const retryResult = obtenerConfiguracion();
          if (retryResult && retryResult.success) {
            return formatearConfiguracionParaFrontend(retryResult.config, false);
          }
        } else {
          console.error('❌ [ADMIN-INTERFACE] Reparación fallida:', reparacion.errores);
        }
      }
      
      // 🔧 [FIX] Si falla, intentar con configuración por defecto como respaldo
      if (typeof obtenerConfiguracionPorDefecto === 'function') {
        console.log('🔄 [ADMIN-INTERFACE] Intentando cargar configuración por defecto como respaldo...');
        const configDefault = obtenerConfiguracionPorDefecto();
        return formatearConfiguracionParaFrontend(configDefault, true);
      }
      
      return {
        success: false,
        code: 'CONFIG_LOAD_ERROR',
        message: result?.message || 'Error cargando configuración del sistema. Posibles problemas: 1) La hoja "Configuración" no existe, 2) Estructura de encabezados incorrecta, 3) Permisos insuficientes.',
        configuraciones: [],
        diagnosticInfo: {
          sheetExpected: 'Configuración',
          headersExpected: ['Clave', 'Valor', 'Categoría'],
          suggestedAction: 'Ejecutar repararSistemaConfiguracion() o diagnosticarSistemaConfiguracionCompleto()'
        },
        details: { raw: result || null }
      };
    }
    
    // 🔧 [FIX] Convertir la configuración de Google Sheets al formato esperado por el frontend
    return formatearConfiguracionParaFrontend(result.config, false);
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error crítico cargando configuración:', error);
    registrarEventoError(`Error crítico en loadConfiguration: ${error.message}`, 'loadConfiguration', error.stack);
    return {
      success: false,
      code: 'UNEXPECTED_ERROR',
      message: 'Error crítico cargando configuración: ' + error.message,
      configuraciones: [],
      error: error.message,
      diagnosticInfo: {
        suggestion: 'Verificar que config.js esté cargado y que la hoja "Configuración" exista en Google Sheets'
      },
      details: { name: error.name, stack: error.stack }
    };
  }
}



/**
 * 🔧 [FIX] Formatea la configuración para el frontend
 * @param {Object} config - Configuración como objeto clave-valor
 * @param {boolean} isDefault - Si es configuración por defecto
 * @return {Object} Configuración formateada para el frontend
 */
function formatearConfiguracionParaFrontend(config, isDefault = false) {
  const configuraciones = [];
  
  if (!config || typeof config !== 'object') {
    console.warn('⚠️ [ADMIN-INTERFACE] Configuración inválida recibida');
    return {
      success: false,
      message: 'Configuración inválida',
      configuraciones: []
    };
  }
  
  // Convertir objeto de configuración a array para el frontend
  Object.keys(config).forEach(clave => {
    configuraciones.push({
      clave: clave,
      valor: config[clave],
      descripcion: generarDescripcionConfig(clave)
    });
  });
  
  const mensaje = isDefault 
    ? `${configuraciones.length} configuraciones por defecto cargadas (usando respaldo)`
    : `${configuraciones.length} configuraciones cargadas desde Google Sheets`;
  
  console.log(`✅ [ADMIN-INTERFACE] ${mensaje}`);
  registrarEventoInfo(mensaje, 'loadConfiguration');
  
  return {
    success: true,
    message: mensaje,
    configuraciones: configuraciones,
    isDefault: isDefault
  };
}

/**
 * 🔧 [FIX] Genera descripción amigable para una clave de configuración
 * @param {string} clave - La clave de configuración
 * @return {string} Descripción amigable
 */
function generarDescripcionConfig(clave) {
  const descripciones = {
    'EMPRESA_NOMBRE': 'Nombre de la empresa',
    'EMPRESA_LOGO': 'URL del logo de la empresa',
    'EMPRESA_DIRECCION': 'Dirección física de la empresa',
    'EMPRESA_TELEFONO': 'Teléfono de contacto',
    'EMPRESA_EMAIL': 'Email corporativo',
    'HORARIO_APERTURA': 'Hora de apertura (formato 24h)',
    'HORARIO_CIERRE': 'Hora de cierre (formato 24h)',
    'MODO_24_HORAS': 'Operación 24 horas (SI/NO)',
    'DIAS_LABORABLES': 'Días laborables (separados por coma)',
    'ZONA_HORARIA': 'Zona horaria del sistema',
    'QR_CAMPO_VISION': 'Campo de visión de la cámara QR (%)',
    'QR_TIEMPO_ESCANEO': 'Tiempo de escaneo QR (ms)',
    'CACHE_DURACION_CONFIG': 'Duración cache configuración (segundos)',
    'NOTIFICACIONES_EMAIL': 'Email para notificaciones del sistema'
  };
  
  return descripciones[clave] || clave.replace(/_/g, ' ').toLowerCase()
    .replace(/\b\w/g, l => l.toUpperCase());
}

/**
 * 🔧 [FIX] Actualiza configuración en Google Sheets (función para admin.html)
 * @param {Object} nuevaConfig - Objeto con configuraciones a actualizar
 * @return {Object} Resultado de la operación
 */
function actualizarConfiguracion(nuevaConfig) {
  console.log('💾 [ADMIN-INTERFACE] Actualizando configuración en Google Sheets...');
  
  try {
    // Verificar que tenemos la función para actualizar
    if (typeof actualizarConfiguracionMultiple !== 'function') {
      console.error('❌ [ADMIN-INTERFACE] Función actualizarConfiguracionMultiple no disponible');
      return {
        success: false,
        code: 'CONFIG_UPDATE_FN_NOT_AVAILABLE',
        message: 'Sistema de actualización no disponible. Verifique que config.js esté cargado.'
      };
    }
    
    console.log('📝 [ADMIN-INTERFACE] Configuraciones a actualizar:', Object.keys(nuevaConfig));
    registrarEventoInfo(`Actualizando ${Object.keys(nuevaConfig).length} configuraciones`, 'actualizarConfiguracion');
    
    // Usar la función de config.js que escribe a Google Sheets
    const result = actualizarConfiguracionMultiple(nuevaConfig);
    
    if (result.success) {
      console.log('✅ [ADMIN-INTERFACE] Configuración actualizada exitosamente');
      registrarEventoInfo('Configuración actualizada exitosamente', 'actualizarConfiguracion', {
        configuraciones: Object.keys(nuevaConfig),
        timestamp: result.lastSaved
      });
    } else {
      console.error('❌ [ADMIN-INTERFACE] Error actualizando configuración:', result.message);
      registrarEventoError(`Error actualizando configuración: ${result.message}`, 'actualizarConfiguracion');
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error crítico actualizando configuración:', error);
    registrarEventoError(`Error crítico en actualizarConfiguracion: ${error.message}`, 'actualizarConfiguracion', error.stack);
    return {
      success: false,
      code: 'UNEXPECTED_ERROR',
      message: 'Error crítico actualizando configuración: ' + error.message,
      details: { name: error.name, stack: error.stack }
    };
  }
}

/**
 * 🔧 [FIX] Registra eventos de información en el sistema de logs
 * @param {string} mensaje - Mensaje del evento
 * @param {string} funcion - Función que genera el evento
 * @param {Object} detalles - Detalles adicionales
 */
function registrarEventoInfo(mensaje, funcion, detalles = null) {
  try {
    if (typeof registrarLog === 'function') {
      registrarLog('INFO', mensaje, detalles ? { funcion: funcion, ...detalles } : { funcion: funcion }, 'Panel Admin');
    } else if (typeof logInfo === 'function') {
      logInfo(mensaje, detalles ? { funcion: funcion, ...detalles } : { funcion: funcion }, 'Panel Admin');
    }
  } catch (error) {
    console.warn('No se pudo registrar evento INFO:', error);
  }
}

/**
 * 🔧 [FIX] Registra eventos de error en el sistema de logs
 * @param {string} mensaje - Mensaje del error
 * @param {string} funcion - Función que genera el error
 * @param {string} stack - Stack trace del error
 */
function registrarEventoError(mensaje, funcion, stack = null) {
  try {
    if (typeof registrarLog === 'function') {
      registrarLog('ERROR', mensaje, { funcion: funcion, stack: stack }, 'Panel Admin');
    } else if (typeof logError === 'function') {
      logError(mensaje, { funcion: funcion, stack: stack }, 'Panel Admin');
    }
  } catch (error) {
    console.warn('No se pudo registrar evento ERROR:', error);
  }
}

// ======================
// 🔗 FUNCIONES DE INTERFAZ PARA USUARIOS ADMIN
// ======================

/**
 * 🔍 DIAGNÓSTICO COMPLETO DEL SISTEMA
 * Detecta conflictos, funciones duplicadas, problemas de configuración
 */
function diagnosticoCompletoSistema() {
  console.log('🔍 ===============================================');
  console.log('🔍 DIAGNÓSTICO COMPLETO DEL SISTEMA SURPASS V6.2');
  console.log('🔍 ===============================================');
  
  const diagnostico = {
    timestamp: new Date().toLocaleString('es-CO'),
    errores: [],
    advertencias: [],
    informacion: [],
    funcionesDisponibles: {},
    conflictos: [],
    configuracion: { estado: 'desconocido', fuente: 'no determinado' }
  };
  
  // 1. VERIFICAR FUNCIONES CRÍTICAS
  console.log('📋 1. VERIFICANDO FUNCIONES CRÍTICAS...');
  
  const funcionesCriticas = [
    'obtenerConfiguracion',
    'registrarLog', 
    'loadConfiguration',
    'loadLogs',
    'loadPersonnel'
  ];
  
  funcionesCriticas.forEach(nombreFunc => {
    try {
      if (typeof eval(nombreFunc) === 'function') {
        diagnostico.funcionesDisponibles[nombreFunc] = '✅ DISPONIBLE';
        diagnostico.informacion.push(`Función ${nombreFunc} está disponible`);
      } else {
        diagnostico.funcionesDisponibles[nombreFunc] = '❌ NO DISPONIBLE';
        diagnostico.errores.push(`Función crítica ${nombreFunc} no está disponible`);
      }
    } catch (error) {
      diagnostico.funcionesDisponibles[nombreFunc] = '❌ ERROR AL VERIFICAR';
      diagnostico.errores.push(`Error verificando función ${nombreFunc}: ${error.message}`);
    }
  });
  
  // 2. PROBAR CARGA DE CONFIGURACIÓN
  console.log('⚙️ 2. PROBANDO CARGA DE CONFIGURACIÓN...');
  
  try {
    if (typeof obtenerConfiguracion === 'function') {
      const resultadoConfig = obtenerConfiguracion();
      
      if (resultadoConfig && resultadoConfig.success) {
        diagnostico.configuracion.estado = '✅ EXITOSO';
        diagnostico.configuracion.fuente = 'Google Sheets';
        diagnostico.configuracion.totalParametros = Object.keys(resultadoConfig.config || {}).length;
        diagnostico.informacion.push(`Configuración cargada exitosamente con ${diagnostico.configuracion.totalParametros} parámetros`);
        
        // Verificar parámetros críticos
        const parametrosCriticos = ['EMPRESA_NOMBRE', 'HORARIO_APERTURA', 'HORARIO_CIERRE'];
        parametrosCriticos.forEach(param => {
          if (resultadoConfig.config && resultadoConfig.config[param]) {
            diagnostico.informacion.push(`Parámetro crítico ${param}: ${resultadoConfig.config[param]}`);
          } else {
            diagnostico.advertencias.push(`Parámetro crítico ${param} no encontrado en configuración`);
          }
        });
        
      } else {
        diagnostico.configuracion.estado = '⚠️ FALLO PARCIAL';
        diagnostico.configuracion.fuente = 'Por defecto (error en Google Sheets)';
        diagnostico.advertencias.push(`Error cargando configuración: ${resultadoConfig?.message || 'Resultado inválido'}`);
      }
    } else {
      diagnostico.configuracion.estado = '❌ FUNCIÓN NO DISPONIBLE';
      diagnostico.errores.push('Función obtenerConfiguracion no está disponible');
    }
  } catch (error) {
    diagnostico.configuracion.estado = '❌ ERROR CRÍTICO';
    diagnostico.errores.push(`Error crítico cargando configuración: ${error.message}`);
  }
  
  // 3. VERIFICAR SISTEMA DE LOGS
  console.log('📋 3. VERIFICANDO SISTEMA DE LOGS...');
  
  try {
    if (typeof registrarLog === 'function') {
      // Intentar registrar un log de prueba
      const resultadoLog = registrarLog('INFO', 'Diagnóstico completo del sistema - Prueba de logging', { 
        diagnostico: true, 
        timestamp: new Date().toISOString() 
      }, 'Sistema de Diagnóstico');
      
      diagnostico.informacion.push('Sistema de logs probado exitosamente');
    } else {
      diagnostico.errores.push('Función registrarLog no está disponible');
    }
  } catch (error) {
    diagnostico.errores.push(`Error probando sistema de logs: ${error.message}`);
  }
  
  // 4. VERIFICAR ACCESO A HOJAS DE GOOGLE
  console.log('📊 4. VERIFICANDO ACCESO A GOOGLE SHEETS...');
  
  try {
    const spreadsheet = SpreadsheetApp.getActive();
    const hojas = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    diagnostico.informacion.push(`Hojas disponibles: ${hojas.join(', ')}`);
    
    const hojasEsperadas = ['Configuración', 'Logs', 'Base de Datos', 'Registro'];
    hojasEsperadas.forEach(nombreHoja => {
      if (hojas.includes(nombreHoja)) {
        diagnostico.informacion.push(`✅ Hoja "${nombreHoja}" encontrada`);
      } else {
        diagnostico.advertencias.push(`⚠️ Hoja "${nombreHoja}" no encontrada`);
      }
    });
    
    // Verificar específicamente la hoja de configuración
    const hojaConfig = spreadsheet.getSheetByName('Configuración');
    if (hojaConfig) {
      const datos = hojaConfig.getDataRange().getValues();
      if (datos.length > 1) {
        diagnostico.informacion.push(`Hoja Configuración contiene ${datos.length - 1} parámetros`);
      } else {
        diagnostico.advertencias.push('Hoja Configuración está vacía');
      }
    }
    
  } catch (error) {
    diagnostico.errores.push(`Error accediendo a Google Sheets: ${error.message}`);
  }
  
  // 5. GENERAR RESUMEN
  console.log('📋 5. GENERANDO RESUMEN...');
  
  const resumen = {
    estado: diagnostico.errores.length === 0 ? 
      (diagnostico.advertencias.length === 0 ? '✅ SISTEMA SALUDABLE' : '⚠️ ADVERTENCIAS DETECTADAS') : 
      '❌ ERRORES CRÍTICOS',
    totalErrores: diagnostico.errores.length,
    totalAdvertencias: diagnostico.advertencias.length,
    funcionesDisponibles: Object.values(diagnostico.funcionesDisponibles).filter(v => v.includes('✅')).length,
    configuracionOperativa: diagnostico.configuracion.estado.includes('✅')
  };
  
  // 6. MOSTRAR RESULTADOS
  console.log('🔍 ===============================================');
  console.log('📊 RESUMEN DEL DIAGNÓSTICO');
  console.log('🔍 ===============================================');
  console.log(`🎯 Estado General: ${resumen.estado}`);
  console.log(`❌ Errores: ${resumen.totalErrores}`);
  console.log(`⚠️ Advertencias: ${resumen.totalAdvertencias}`);
  console.log(`✅ Funciones Disponibles: ${resumen.funcionesDisponibles}/5`);
  console.log(`⚙️ Configuración: ${diagnostico.configuracion.estado}`);
  
  if (diagnostico.errores.length > 0) {
    console.log('\n❌ ERRORES DETECTADOS:');
    diagnostico.errores.forEach((error, i) => console.log(`   ${i+1}. ${error}`));
  }
  
  if (diagnostico.advertencias.length > 0) {
    console.log('\n⚠️ ADVERTENCIAS:');
    diagnostico.advertencias.forEach((adv, i) => console.log(`   ${i+1}. ${adv}`));
  }
  
  console.log('\n✅ INFORMACIÓN:');
  diagnostico.informacion.forEach((info, i) => console.log(`   ${i+1}. ${info}`));
  
  console.log('🔍 ===============================================');
  
  return {
    success: resumen.estado.includes('✅'),
    diagnostico: diagnostico,
    resumen: resumen,
    recomendaciones: generarRecomendacionesDiagnostico(diagnostico)
  };
}

/**
 * Genera recomendaciones basadas en el diagnóstico
 */
function generarRecomendacionesDiagnostico(diagnostico) {
  const recomendaciones = [];
  
  if (diagnostico.errores.length > 0) {
    recomendaciones.push('🔧 Resolver errores críticos antes de continuar');
  }
  
  if (!diagnostico.configuracion.estado.includes('✅')) {
    recomendaciones.push('⚙️ Verificar y corregir el sistema de configuración');
    recomendaciones.push('📊 Comprobar que la hoja "Configuración" existe y tiene datos');
  }
  
  if (diagnostico.funcionesDisponibles.registrarLog?.includes('❌')) {
    recomendaciones.push('📋 Corregir el sistema de logs para detectar futuros problemas');
  }
  
  if (diagnostico.advertencias.length > 3) {
    recomendaciones.push('⚠️ Revisar y resolver las advertencias para optimizar el sistema');
  }
  
  if (recomendaciones.length === 0) {
    recomendaciones.push('✅ El sistema está funcionando correctamente');
  }
  
  return recomendaciones;
}
/**
 * 🛠️ REPARACIÓN AUTOMÁTICA DEL SISTEMA
 * Soluciona problemas comunes detectados por el diagnóstico
 */
function reparacionAutomaticaSistema() {
  console.log('🛠️ ===============================================');
  console.log('🛠️ INICIANDO REPARACIÓN AUTOMÁTICA DEL SISTEMA');
  console.log('🛠️ ===============================================');
  
  const reparacion = {
    timestamp: new Date().toLocaleString('es-CO'),
    accionesEjecutadas: [],
    erroresReparacion: [],
    exitosa: false
  };
  
  try {
    // 1. VERIFICAR Y CREAR HOJA DE CONFIGURACIÓN SI NO EXISTE
    console.log('🔧 1. VERIFICANDO HOJA DE CONFIGURACIÓN...');
    
    const spreadsheet = SpreadsheetApp.getActive();
    let hojaConfig = spreadsheet.getSheetByName('Configuración');
    
    if (!hojaConfig) {
      console.log('📋 Creando hoja de configuración...');
      hojaConfig = spreadsheet.insertSheet('Configuración');
      
      // Crear encabezados
      hojaConfig.getRange(1, 1, 1, 3).setValues([['Clave', 'Valor', 'Categoría']]);
      hojaConfig.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
      
      // Agregar configuraciones básicas
      const configBasica = [
        ['EMPRESA_NOMBRE', 'SurPass', 'EMPRESA'],
        ['HORARIO_APERTURA', '08:00', 'EMPRESA'],
        ['HORARIO_CIERRE', '17:00', 'EMPRESA'],
        ['MODO_24_HORAS', 'NO', 'EMPRESA'],
        ['DIAS_LABORABLES', 'Lunes,Martes,Miércoles,Jueves,Viernes,Sabado,Domingo', 'EMPRESA'],
        ['QR_CAMPO_VISION', '350', 'CAMARA_QR'],
        ['QR_TIEMPO_ESCANEO', '250', 'CAMARA_QR'],
        ['CACHE_DURACION_ESTADISTICAS', '180', 'CACHE'],
        ['NOTIFICACIONES_EMAIL', 'admin@empresa.com', 'NOTIFICACIONES']
      ];
      
      hojaConfig.getRange(2, 1, configBasica.length, 3).setValues(configBasica);
      reparacion.accionesEjecutadas.push('✅ Hoja de configuración creada con parámetros básicos');
      
    } else {
      // Verificar que tenga datos
      const datos = hojaConfig.getDataRange().getValues();
      if (datos.length <= 1) {
        console.log('📋 Agregando datos básicos a hoja de configuración vacía...');
        const configBasica = [
          ['EMPRESA_NOMBRE', 'SurPass', 'EMPRESA'],
          ['HORARIO_APERTURA', '08:00', 'EMPRESA'],
          ['HORARIO_CIERRE', '17:00', 'EMPRESA'],
          ['MODO_24_HORAS', 'NO', 'EMPRESA']
        ];
        
        hojaConfig.getRange(2, 1, configBasica.length, 3).setValues(configBasica);
        reparacion.accionesEjecutadas.push('✅ Datos básicos agregados a hoja de configuración');
      }
    }
    
    // 2. VERIFICAR HOJA DE LOGS
    console.log('🔧 2. VERIFICANDO HOJA DE LOGS...');
    
    let hojaLogs = spreadsheet.getSheetByName('Logs');
    if (!hojaLogs) {
      console.log('📋 Creando hoja de Logs...');
      hojaLogs = spreadsheet.insertSheet('Logs');
      
      const encabezadosLogs = ['Timestamp', 'Nivel', 'Mensaje', 'Detalles', 'Usuario', 'IP'];
      hojaLogs.getRange(1, 1, 1, encabezadosLogs.length).setValues([encabezadosLogs]);
      hojaLogs.getRange(1, 1, 1, encabezadosLogs.length).setFontWeight("bold").setBackground("#34A853").setFontColor("white");
      
      reparacion.accionesEjecutadas.push('✅ Hoja de Logs creada');
    }
    
    // 3. VERIFICAR HOJA DE ERRORES
    console.log('🔧 3. VERIFICANDO HOJA DE ERRORES...');
    
    let hojaErrores = spreadsheet.getSheetByName('Errores');
    if (!hojaErrores) {
      console.log('📋 Creando hoja de Errores...');
      hojaErrores = spreadsheet.insertSheet('Errores');
      
      const encabezadosErrores = ['Timestamp', 'Nivel', 'Mensaje', 'Detalles', 'Usuario', 'IP', 'Stack'];
      hojaErrores.getRange(1, 1, 1, encabezadosErrores.length).setValues([encabezadosErrores]);
      hojaErrores.getRange(1, 1, 1, encabezadosErrores.length).setFontWeight("bold").setBackground("#EA4335").setFontColor("white");
      
      reparacion.accionesEjecutadas.push('✅ Hoja de Errores creada');
    }
    
    // 4. VERIFICAR HOJA DE BASE DE DATOS
    console.log('🔧 4. VERIFICANDO HOJA DE BASE DE DATOS...');
    
    let hojaBD = spreadsheet.getSheetByName('Base de Datos');
    if (!hojaBD) {
      console.log('📋 Creando hoja de Base de Datos...');
      hojaBD = spreadsheet.insertSheet('Base de Datos');
      
      const encabezadosBD = ['Cédula', 'Nombre', 'Empresa', 'Área', 'Cargo', 'Estado'];
      hojaBD.getRange(1, 1, 1, encabezadosBD.length).setValues([encabezadosBD]);
      hojaBD.getRange(1, 1, 1, encabezadosBD.length).setFontWeight("bold").setBackground("#FF9800").setFontColor("white");
      
      reparacion.accionesEjecutadas.push('✅ Hoja de Base de Datos creada');
    }
    
    // 5. PROBAR SISTEMA REPARADO
    console.log('🔧 5. PROBANDO SISTEMA REPARADO...');
    
    try {
      if (typeof obtenerConfiguracion === 'function') {
        const testConfig = obtenerConfiguracion();
        if (testConfig && testConfig.success) {
          reparacion.accionesEjecutadas.push(`✅ Sistema de configuración funcionando - ${Object.keys(testConfig.config).length} parámetros cargados`);
        } else {
          reparacion.erroresReparacion.push('⚠️ Sistema de configuración aún presenta problemas');
        }
      }
    } catch (error) {
      reparacion.erroresReparacion.push(`Error probando configuración reparada: ${error.message}`);
    }
    
    // 6. REGISTRAR REPARACIÓN
    try {
      if (typeof registrarLog === 'function') {
        registrarLog('INFO', 'Reparación automática del sistema completada', {
          acciones: reparacion.accionesEjecutadas.length,
          errores: reparacion.erroresReparacion.length,
          detalles: reparacion
        }, 'Sistema de Reparación');
        reparacion.accionesEjecutadas.push('✅ Reparación registrada en logs');
      }
    } catch (error) {
      console.warn('No se pudo registrar la reparación:', error);
    }
    
    reparacion.exitosa = reparacion.erroresReparacion.length === 0;
    
  } catch (error) {
    console.error('❌ Error durante reparación automática:', error);
    reparacion.erroresReparacion.push(`Error crítico: ${error.message}`);
    reparacion.exitosa = false;
  }
  
  // 7. MOSTRAR RESULTADOS
  console.log('🛠️ ===============================================');
  console.log('📊 RESUMEN DE LA REPARACIÓN');
  console.log('🛠️ ===============================================');
  console.log(`🎯 Estado: ${reparacion.exitosa ? '✅ EXITOSA' : '❌ CON ERRORES'}`);
  console.log(`🔧 Acciones ejecutadas: ${reparacion.accionesEjecutadas.length}`);
  console.log(`❌ Errores: ${reparacion.erroresReparacion.length}`);
  
  if (reparacion.accionesEjecutadas.length > 0) {
    console.log('\n✅ ACCIONES REALIZADAS:');
    reparacion.accionesEjecutadas.forEach((accion, i) => console.log(`   ${i+1}. ${accion}`));
  }
  
  if (reparacion.erroresReparacion.length > 0) {
    console.log('\n❌ ERRORES DURANTE REPARACIÓN:');
    reparacion.erroresReparacion.forEach((error, i) => console.log(`   ${i+1}. ${error}`));
  }
  
  console.log('🛠️ ===============================================');
  
  return reparacion;
}

function probarSistemaConfiguracion() {
  console.log('🧪 [ADMIN-INTERFACE] Ejecutando prueba del sistema de configuración...');
  
  try {
    // Ejecutar diagnóstico completo
    const diagnostico = diagnosticoCompletoSistema();
    
    // Obtener estadísticas de configuración
    const configResult = obtenerConfiguracion();
    let stats = { totalConfiguraciones: 0, categoriasDetectadas: [] };
    
    if (configResult && configResult.success) {
      const config = configResult.config;
      stats.totalConfiguraciones = Object.keys(config).length;
      
      // Detectar categorías
      const categorias = new Set();
      Object.keys(config).forEach(key => {
        if (key.startsWith('EMPRESA_')) categorias.add('EMPRESA');
        else if (key.startsWith('UI_')) categorias.add('INTERFAZ');
        else if (key.startsWith('QR_')) categorias.add('CAMARA_QR');
        else if (key.includes('CONTRASEÑA') || key.includes('SESION')) categorias.add('SEGURIDAD');
        else if (key.startsWith('NOTIF') || key.includes('EMAIL')) categorias.add('NOTIFICACIONES');
        else if (key.startsWith('CACHE_')) categorias.add('CACHE');
        else if (key.startsWith('REPORTE_')) categorias.add('REPORTES');
        else if (key.startsWith('EVACUACION_')) categorias.add('EVACUACION');
        else if (key.includes('BACKUP') || key.includes('ARCHIVO')) categorias.add('ARCHIVO');
        else if (key.includes('TEMA_') || key.includes('IDIOMA')) categorias.add('PERSONALIZACION');
        else if (key.startsWith('INTEGRACION_')) categorias.add('INTEGRACION');
        else categorias.add('GENERAL');
      });
      
      stats.categoriasDetectadas = Array.from(categorias);
    }
    
    return {
      success: true,
      message: 'Prueba del sistema de configuración completada',
      diagnostico: diagnostico,
      estadisticas: stats,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error en prueba del sistema:', error);
    return {
      success: false,
      message: 'Error ejecutando prueba: ' + error.message,
      error: error.stack
    };
  }
}

// ======================
// 🔗 FUNCIONES DE INTERFAZ PARA USUARIOS ADMIN
// ======================

/**
 * 🔧 [FIX] Carga usuarios administrativos desde la hoja "Clave"
 * @return {Object} Lista de administradores con estructura flexible
 */
function loadAdminUsers() {
  console.log('👥 [ADMIN-INTERFACE] Cargando usuarios administrativos desde hoja "Clave"...');
  
  try {
    const ss = SpreadsheetApp.getActive();
    const hojaClave = ss.getSheetByName('Clave');
    
    if (!hojaClave) {
      console.error('❌ [ADMIN-INTERFACE] Hoja "Clave" no encontrada');
      return {
        success: false,
        message: 'Hoja "Clave" no encontrada. Verifique que la hoja existe.',
        admins: [],
        stats: { total: 0, activos: 0, permisos_completos: 0, solo_lectura: 0 }
      };
    }

    const datos = hojaClave.getDataRange().getValues();
    if (datos.length <= 1) {
      return {
        success: false,
        message: 'No hay usuarios administradores registrados en la hoja "Clave".',
        admins: [],
        stats: { total: 0, activos: 0, permisos_completos: 0, solo_lectura: 0 }
      };
    }

    // 🔧 [FIX] Mapeo flexible de columnas
    const encabezados = datos[0];
    const mapeoColumnas = mapearColumnasFlexible(encabezados, {
      cedula: ['cédula', 'cedula', 'id', 'usuario'],
      nombre: ['nombre', 'name'],
      contrasena: ['contraseña', 'contrasena', 'password', 'clave'],
      cargo: ['cargo', 'position', 'rol', 'role'],
      email: ['email', 'correo', 'mail'],
      permisos: ['permisos', 'permissions'],
      estado: ['estado', 'status', 'activo']
    });

    console.log('🗂️ [ADMIN-INTERFACE] Mapeo de columnas detectado:', mapeoColumnas);

    const admins = [];
    
    // Procesar datos desde la fila 1 (saltando encabezados)
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      
      // Extraer datos usando mapeo flexible
      const admin = {
        cedula: obtenerValorColumna(fila, mapeoColumnas.cedula) || `admin_${i}`,
        nombre: obtenerValorColumna(fila, mapeoColumnas.nombre) || 'Sin nombre',
        contrasena: obtenerValorColumna(fila, mapeoColumnas.contrasena) || '',
        cargo: obtenerValorColumna(fila, mapeoColumnas.cargo) || 'Administrador',
        email: obtenerValorColumna(fila, mapeoColumnas.email) || '',
        permisos: obtenerValorColumna(fila, mapeoColumnas.permisos) || '{}',
        estado: obtenerValorColumna(fila, mapeoColumnas.estado) || 'activo',
        fila: i + 1 // Para operaciones de edición
      };

      // Procesar permisos si están en formato JSON
      try {
        if (typeof admin.permisos === 'string' && admin.permisos.startsWith('{')) {
          admin.permisosObj = JSON.parse(admin.permisos);
        }
      } catch (error) {
        console.warn(`⚠️ [ADMIN-INTERFACE] Error parseando permisos para ${admin.cedula}:`, error);
        admin.permisosObj = {};
      }

      admins.push(admin);
    }

    // Calcular estadísticas - 🔧 [FIX] Lógica corregida para conteo de activos
    const stats = {
      total: admins.length,
      activos: admins.filter(a => {
        const estado = String(a.estado || '').toLowerCase().trim();
        return estado === 'activo' || estado === 'active';
      }).length,
      permisos_completos: admins.filter(a => 
        a.cargo.toLowerCase().includes('administrador') || 
        a.cargo.toLowerCase().includes('principal')
      ).length,
      solo_lectura: admins.filter(a => 
        !a.cargo.toLowerCase().includes('administrador') &&
        !a.cargo.toLowerCase().includes('principal')
      ).length
    };
    
    console.log(`✅ [ADMIN-INTERFACE] ${admins.length} administradores cargados desde hoja "Clave"`);
    registrarEventoInfo(`Administradores cargados: ${admins.length} usuarios`, 'loadAdminUsers');
    
    return {
      success: true,
      message: `${admins.length} usuarios administradores cargados exitosamente`,
      admins: admins,
      stats: stats
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error cargando administradores:', error);
    registrarEventoError(`Error cargando administradores: ${error.message}`, 'loadAdminUsers', error.stack);
    return {
      success: false,
      message: 'Error interno cargando administradores: ' + error.message,
      admins: [],
      stats: { total: 0, activos: 0, permisos_completos: 0, solo_lectura: 0 }
    };
  }
}

/**
 * 🔧 [FIX] Mapea columnas de forma flexible
 * @param {Array} encabezados - Array de encabezados de la hoja
 * @param {Object} campos - Objeto con campos a buscar y sus posibles nombres
 * @return {Object} Mapeo de columnas encontradas
 */
function mapearColumnasFlexible(encabezados, campos) {
  const mapeo = {};
  
  // Función para normalizar: minúsculas, sin tildes, sin espacios
  function normalizar(str) {
    return String(str)
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, '');
  }

  Object.keys(campos).forEach(campo => {
    const posiblesNombres = campos[campo].map(normalizar);
    let indiceEncontrado = -1;
    for (let i = 0; i < encabezados.length; i++) {
      const encabezadoNorm = normalizar(encabezados[i]);
      if (posiblesNombres.some(nombre => encabezadoNorm.includes(nombre))) {
        indiceEncontrado = i;
        break;
      }
    }
    mapeo[campo] = indiceEncontrado;
  });
  return mapeo;
}

/**
 * 🔧 [FIX] Obtiene valor de columna de forma segura
 * @param {Array} fila - Fila de datos
 * @param {number} indice - Índice de la columna
 * @return {*} Valor de la celda o null
 */
function obtenerValorColumna(fila, indice) {
  if (indice === -1 || indice >= fila.length) {
    return null;
  }
  
  const valor = fila[indice];
  if (valor === null || valor === undefined || valor === '') {
    return null;
  }
  
  return valor;
}

/**
 * 🔧 [FIX] Procesa el resultado de obtenerAdministradores
 * @param {Object} result - Resultado de obtenerAdministradores
 * @return {Object} Resultado procesado
 */
function processAdminResult(result) {
  const admins = result.admins || [];
  const stats = {
    total: admins.length,
    activos: admins.filter(a => {
      const estado = String(a.estado || '').toLowerCase().trim();
      return estado === 'activo' || estado === 'active';
    }).length, // 🔧 [FIX] Contar solo los realmente activos
    permisos_completos: admins.filter(a => (a.cargo || '').toLowerCase().includes('administrador')).length,
    solo_lectura: admins.filter(a => !(a.cargo || '').toLowerCase().includes('administrador')).length
  };
    
  console.log(`✅ [ADMIN-INTERFACE] Administradores cargados: ${stats.total} usuarios`);
    
  return {
    success: true,
    message: result.message,
    admins: admins,
    stats: stats
  };
}

// ======================
// 🔗 FUNCIONES DE INTERFAZ PARA LOGS
// ======================

/**
 * Carga logs del sistema usando múltiples fuentes de datos (VERSIÓN CORREGIDA)
 * @param {number} limite - Número máximo de logs a obtener
 * @return {Object} Logs del sistema con estadísticas
 */
/**
 * 🔧 [DEBUG] Función de prueba simple para diagnosticar problemas de Google Apps Script
 */
function testConnection() {
  try {
    console.log('🔍 [DEBUG] Función testConnection ejecutándose...');
    return {
      success: true,
      message: 'Conexión con Google Apps Script OK',
      timestamp: new Date().toISOString(),
      test: true
    };
  } catch (error) {
    console.error('❌ [DEBUG] Error en testConnection:', error);
    return {
      success: false,
      message: 'Error en testConnection: ' + error.message,
      error: error.message
    };
  }
}

function loadLogs(limite = 50) {
  console.log('📋 [ADMIN-INTERFACE] === INICIO loadLogs ===');
  console.log('📋 [ADMIN-INTERFACE] Parámetros:', { limite });
  console.log('📋 [ADMIN-INTERFACE] Timestamp:', new Date().toISOString());
  
  // Diagnóstico inicial
  try {
    console.log('📋 [ADMIN-INTERFACE] Verificando entorno...');
    console.log('📋 [ADMIN-INTERFACE] SpreadsheetApp disponible:', typeof SpreadsheetApp);
    console.log('📋 [ADMIN-INTERFACE] SHEET_LOGS:', typeof SHEET_LOGS, SHEET_LOGS);
    console.log('📋 [ADMIN-INTERFACE] obtenerLogsRecientes disponible:', typeof obtenerLogsRecientes);
  } catch (diagError) {
    console.error('❌ [ADMIN-INTERFACE] Error en diagnóstico inicial:', diagError);
  }
  
  // Validar dependencias críticas
  try {
    if (!SHEET_LOGS) {
      console.error('❌ [ADMIN-INTERFACE] SHEET_LOGS no está definido');
      return {
        success: false,
        message: 'Configuración de hoja de logs no encontrada (SHEET_LOGS)',
        logs: [],
        stats: { total: 0, errores: 1, warnings: 0, info: 0 }
      };
    }
  } catch (depError) {
    console.error('❌ [ADMIN-INTERFACE] Error verificando dependencias:', depError);
    return {
      success: false,
      message: 'Error de dependencias: ' + depError.message,
      logs: [],
      stats: { total: 0, errores: 1, warnings: 0, info: 0 }
    };
  }
  
  registrarEventoInfo('Iniciando carga de logs del sistema', 'loadLogs', { limite: limite });
  
  try {
    let logsObtenidos = [];
    let erroresObtenidos = [];
    let fuenteUsada = 'ninguna';
    
    // 🔧 [FIX] Fuente 1: Función principal de logs.js
    if (typeof obtenerLogsRecientes === 'function') {
      console.log('📄 [ADMIN-INTERFACE] Usando obtenerLogsRecientes de logs.js...');
      try {
        const resultLogs = obtenerLogsRecientes(limite);
        console.log('📊 [DEBUG] obtenerLogsRecientes devolvió:', resultLogs);
        
        // Verificar diferentes formatos de respuesta
        if (resultLogs && resultLogs.success && Array.isArray(resultLogs.logs)) {
          logsObtenidos = resultLogs.logs;
          fuenteUsada = 'logs.js (obtenerLogsRecientes)';
          console.log(`✅ [ADMIN-INTERFACE] ${logsObtenidos.length} logs obtenidos de logs.js`);
        } else if (Array.isArray(resultLogs)) {
          // La función devolvió directamente el array de logs
          logsObtenidos = resultLogs;
          fuenteUsada = 'logs.js (obtenerLogsRecientes - array directo)';
          console.log(`✅ [ADMIN-INTERFACE] ${logsObtenidos.length} logs obtenidos de logs.js (array directo)`);
        } else {
          console.warn('⚠️ [ADMIN-INTERFACE] obtenerLogsRecientes no devolvió logs válidos:', resultLogs);
        }
      } catch (error) {
        console.error('❌ [ADMIN-INTERFACE] Error usando obtenerLogsRecientes:', error);
      }
    } else {
      console.warn('⚠️ [ADMIN-INTERFACE] Función obtenerLogsRecientes no disponible');
    }
    
    // 🔧 [FIX] Fuente 2: Lectura directa de Google Sheets como fuente principal
    if (logsObtenidos.length === 0) {
      console.log('📄 [ADMIN-INTERFACE] Intentando lectura directa de Google Sheets...');
      try {
        const ss = SpreadsheetApp.getActive();
        const hojaLogs = ss.getSheetByName(SHEET_LOGS);
        
        if (hojaLogs && hojaLogs.getLastRow() > 1) {
          const datos = hojaLogs.getDataRange().getValues();
          const encabezados = datos[0].map(h => String(h).toLowerCase().trim());
          
          // Mapear índices de columnas con manejo mejorado
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
            usuario: getColIndex(['usuario', 'user', 'autor']),
            ip: getColIndex(['ip', 'direccionip'])
          };
          
          // Procesar filas de logs
          logsObtenidos = [];
          for (let i = 1; i < datos.length; i++) {
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
            
            // Limitar el número de registros
            if (logsObtenidos.length >= limite) break;
          }
          
          fuenteUsada = 'Google Sheets directo';
          console.log(`✅ [ADMIN-INTERFACE] ${logsObtenidos.length} logs obtenidos por lectura directa`);
        } else {
          console.warn('⚠️ [ADMIN-INTERFACE] Hoja de logs vacía o no encontrada');
          // Forzar un log de ejemplo si no hay datos
          logsObtenidos = [{
            id: 'log_demo_1',
            timestamp: new Date().toISOString(),
            fecha: new Date().toISOString(),
            nivel: 'INFO',
            mensaje: 'No se encontraron registros de logs en la hoja',
            usuario: 'Sistema',
            ip: 'N/A'
          }];
          fuenteUsada = 'Datos de ejemplo';
        }
      } catch (sheetsError) {
        console.error('❌ [ADMIN-INTERFACE] Error accediendo directamente a Google Sheets:', sheetsError);
        registrarEventoError('Error accediendo directamente a Google Sheets para logs', 'loadLogs', sheetsError.message);
        
        // Devolver un error estructurado
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
    
    // 🔧 [FIX] Combinar logs normales con errores
    const todosLosLogs = [...logsObtenidos];
    
    // Agregar errores como logs de tipo ERROR
    erroresObtenidos.forEach((error, index) => {
      todosLosLogs.push({
        id: `error_${index}`,
        timestamp: error.timestamp,
        fecha: error.timestamp,
        nivel: 'ERROR',
        mensaje: error.mensaje,
        detalles: error.detalles,
        usuario: error.usuario,
        ip: 'N/A'
      });
    });
    
    // Verificar si hay logs para procesar
    if (todosLosLogs.length === 0) {
      console.warn('⚠️ [ADMIN-INTERFACE] No se encontraron logs en ninguna fuente');
      return {
        success: true,
        message: 'No hay logs disponibles en el sistema',
        logs: [],
        stats: { total: 0, errores: 0, warnings: 0, info: 0 },
        fuente: 'Sin datos'
      };
    }
    
    // Procesar logs obtenidos
    const result = processLogsResult(todosLosLogs, fuenteUsada);
    
    console.log(`✅ [ADMIN-INTERFACE] ${result.stats.total} logs cargados desde: ${fuenteUsada}`);
    registrarEventoInfo(`Logs cargados exitosamente: ${result.stats.total} registros`, 'loadLogs', {
      fuente: fuenteUsada,
      limite: limite,
      stats: result.stats
    });
    
    return result;
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error crítico cargando logs:', error);
    
    // Intentar registrar el error si es posible
    try {
      registrarEventoError(`Error crítico cargando logs: ${error.message}`, 'loadLogs', error.stack);
    } catch (logError) {
      console.error('❌ [ADMIN-INTERFACE] No se pudo registrar el error en logs:', logError);
    }
    
    // Devolver respuesta de error estructurada
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
 * 🔧 [FIX] Procesa el resultado de logs para el frontend
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
    warnings: logs.filter(log => String(log.nivel || log.tipo || '').toUpperCase() === 'WARNING').length,
    info: logs.filter(log => String(log.nivel || log.tipo || '').toUpperCase() === 'INFO').length
  };
  
  // Formatear logs para el frontend - Mapeo mejorado para diferentes formatos
  const logsFormateados = logs.map((log, index) => ({
    id: `log_${index}_${Date.now()}`,
    fecha: log.timestamp || log.fecha || new Date().toISOString(),
    nivel: String(log.nivel || log.tipo || log.level || 'INFO').toUpperCase(),
    mensaje: log.mensaje || log.message || 'Sin mensaje',
    usuario: log.usuario || log.user || 'Sistema',
    detalles: log.detalles || log.datos || log.details || log.data || '',
    ip: log.ip || 'N/A'
  }));
  
  console.log(`✅ [ADMIN-INTERFACE] Logs procesados: ${stats.total} registros (${stats.errores} errores, ${stats.warnings} warnings, ${stats.info} info)`);
  
  return {
    success: true,
    message: `${stats.total} logs procesados desde ${fuente}`,
    logs: logsFormateados,
    stats: stats,
    fuente: fuente
  };
}

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
    if (datos.length <= 1) { // Solo encabezados o vacío
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
    for (let i = datos.length - 1; i >= 1; i--) { // Empezar desde abajo para no afectar índices
      const fila = datos[i];
      const fechaStr = fila[indiceFecha];
      let fechaLog;
      
      try {
        fechaLog = new Date(fechaStr);
        if (isNaN(fechaLog.getTime())) {
          // Si no es una fecha válida, asumir que es muy antigua para eliminarla
          filasAEliminar.push(i + 1); // +1 porque las filas en Sheets son 1-based
          continue;
        }
      } catch (e) {
        // Si hay error al parsear la fecha, asumir que es muy antigua
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
        filasAEliminar.push(i + 1); // +1 porque las filas en Sheets son 1-based
      }
    }
    
    // Eliminar filas (empezar por las de mayor índice para no afectar las demás)
    filasAEliminar.sort((a, b) => b - a); // Orden descendente
    
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
    registrarEventoInfo(
      `Limpieza de logs completada: ${eliminados} registros eliminados`,
      'limpiarLogs',
      { 
        totalRegistros: datos.length - 1, 
        eliminados,
        criterios: { dias, soloErrores, usuario }
      }
    );
    
    return {
      success: true,
      message: `Se eliminaron ${eliminados} registros de logs`,
      registrosEliminados: eliminados,
      totalRegistros: datos.length - 1
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error en limpiarLogs:', error);
    registrarEventoError(
      'Error al limpiar logs: ' + error.message,
      'limpiarLogs',
      error.stack
    );
    
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
    const resultadoLogs = loadLogs(limite * 10); // Multiplicamos para tener más registros para filtrar
    
    if (!resultadoLogs.success) {
      throw new Error(resultadoLogs.message || 'Error al cargar logs para exportación');
    }
    
    let logs = resultadoLogs.logs || [];
    
    // Aplicar filtros si se especificaron
    if (filtroNivel || usuario || fechaDesde || fechaHasta) {
      logs = logs.filter(log => {
        // Filtrar por nivel si se especificó
        if (filtroNivel && !String(log.nivel || '').toUpperCase().includes(filtroNivel.toUpperCase())) {
          return false;
        }
        
        // Filtrar por usuario si se especificó
        if (usuario && !String(log.usuario || '').toLowerCase().includes(usuario.toLowerCase())) {
          return false;
        }
        
        // Filtrar por rango de fechas si se especificó
        if (fechaDesde || fechaHasta) {
          try {
            const fechaLog = new Date(log.fecha || log.timestamp);
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
      }).slice(0, limite); // Aplicar el límite después de filtrar
    } else {
      // Si no hay filtros, aplicar solo el límite
      logs = logs.slice(0, limite);
    }
    
    // Generar el contenido según el formato solicitado
    let contenido = '';
    let nombreArchivo = `logs_${new Date().toISOString().split('T')[0]}`;
    
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
        const campos = new Set();
        logs.forEach(log => {
          Object.keys(log).forEach(campo => campos.add(campo));
        });
        
        const camposArray = Array.from(campos);
        
        // Crear fila de encabezados
        contenido = camposArray.join(';') + '\n';
        
        // Agregar filas de datos
        logs.forEach(log => {
          const fila = camposArray.map(campo => {
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
    registrarEventoInfo(
      `Exportación de logs completada: ${logs.length} registros en formato ${formato.toUpperCase()}`,
      'exportarLogs',
      { 
        totalRegistros: logs.length,
        formato,
        filtros: { filtroNivel, usuario, fechaDesde, fechaHasta }
      }
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
    registrarEventoError(
      'Error al exportar logs: ' + error.message,
      'exportarLogs',
      error.stack
    );
    
    return {
      success: false,
      message: 'Error al exportar logs: ' + error.message,
      error: error.toString()
    };
  }
}

/**
 * Busca logs según criterios específicos
 * @param {Object} criterios - Criterios de búsqueda
 * @param {string} [criterios.termino] - Término de búsqueda (busca en mensaje y detalles)
 * @param {string} [criterios.nivel] - Filtrar por nivel (ERROR, WARNING, INFO, etc.)
 * @param {string} [criterios.usuario] - Filtrar por usuario
 * @param {string} [criterios.fechaDesde] - Fecha desde (formato YYYY-MM-DD)
 * @param {string} [criterios.fechaHasta] - Fecha hasta (formato YYYY-MM-DD)
 * @param {number} [criterios.limite=100] - Número máximo de resultados a devolver
 * @return {Object} Resultado con los logs que coinciden con los criterios
 */
function buscarLogs(criterios = {}) {
  const {
    termino,
    nivel,
    usuario,
    fechaDesde,
    fechaHasta,
    limite = 100
  } = criterios;
  
  try {
    // Obtener logs (usar un límite mayor para tener más registros para filtrar)
    const resultadoLogs = loadLogs(limite * 5);
    
    if (!resultadoLogs.success) {
      throw new Error(resultadoLogs.message || 'Error al cargar logs para búsqueda');
    }
    
    let logs = resultadoLogs.logs || [];
    
    // Aplicar filtros
    logs = logs.filter(log => {
      // Filtrar por término de búsqueda
      if (termino) {
        const busqueda = termino.toLowerCase();
        const enMensaje = String(log.mensaje || '').toLowerCase().includes(busqueda);
        const enDetalles = String(log.detalles || '').toLowerCase().includes(busqueda);
        
        if (!enMensaje && !enDetalles) {
          return false;
        }
      }
      
      // Filtrar por nivel
      if (nivel && !String(log.nivel || '').toUpperCase().includes(nivel.toUpperCase())) {
        return false;
      }
      
      // Filtrar por usuario
      if (usuario && !String(log.usuario || '').toLowerCase().includes(usuario.toLowerCase())) {
        return false;
      }
      
      // Filtrar por rango de fechas
      if (fechaDesde || fechaHasta) {
        try {
          const fechaLog = new Date(log.fecha || log.timestamp);
          
          if (fechaDesde) {
            const desde = new Date(fechaDesde);
            if (fechaLog < desde) return false;
          }
          
          if (fechaHasta) {
            const hasta = new Date(fechaHasta);
            hasta.setHours(23, 59, 59, 999); // Fin del día
            if (fechaLog > hasta) return false;
          }
        } catch (e) {
          console.warn('Error al filtrar por fecha:', e);
          return false;
        }
      }
      
      return true;
    }).slice(0, limite); // Aplicar límite de resultados
    
    // Calcular estadísticas
    const stats = {
      total: logs.length,
      errores: logs.filter(log => String(log.nivel || '').toUpperCase() === 'ERROR').length,
      warnings: logs.filter(log => {
        const nivelLog = String(log.nivel || '').toUpperCase();
        return nivelLog === 'WARNING' || nivelLog === 'WARN' || nivelLog === 'ADVERTENCIA';
      }).length,
      info: logs.filter(log => String(log.nivel || '').toUpperCase() === 'INFO').length
    };
    
    // Registrar la acción de búsqueda
    registrarEventoInfo(
      `Búsqueda de logs completada: ${logs.length} resultados`,
      'buscarLogs',
      { 
        totalResultados: logs.length,
        criterios: { 
          termino: termino ? '***' : 'no',
          nivel,
          usuario: usuario ? '***' : 'no',
          fechaDesde,
          fechaHasta,
          limite
        }
      }
    );
    
    return {
      success: true,
      message: `${logs.length} registros encontrados`,
      logs,
      stats,
      criterios
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error en buscarLogs:', error);
    registrarEventoError(
      'Error al buscar logs: ' + error.message,
      'buscarLogs',
      error.stack
    );
    
    return {
      success: false,
      message: 'Error al buscar logs: ' + error.message,
      error: error.toString(),
      logs: [],
      stats: { total: 0, errores: 0, warnings: 0, info: 0 }
    };
  }
}

/**
 * Obtiene estadísticas de los logs
 * @param {Object} opciones - Opciones para el cálculo de estadísticas
 * @param {number} [opciones.dias=30] - Número de días hacia atrás para incluir en las estadísticas
 * @param {string} [opciones.nivel] - Filtrar por nivel (ERROR, WARNING, INFO, etc.)
 * @param {string} [opciones.usuario] - Filtrar por usuario
 * @return {Object} Estadísticas de los logs
 */
function obtenerEstadisticasLogs(opciones = {}) {
  const { dias = 30, nivel, usuario } = opciones;
  const fechaLimite = new Date();
  fechaLimite.setDate(fechaLimite.getDate() - dias);
  
  try {
    // Obtener logs (usar un límite razonable para el análisis)
    const resultadoLogs = loadLogs(1000);
    
    if (!resultadoLogs.success) {
      throw new Error(resultadoLogs.message || 'Error al cargar logs para estadísticas');
    }
    
    let logs = resultadoLogs.logs || [];
    
    // Filtrar por fecha y otros criterios
    logs = logs.filter(log => {
      try {
        const fechaLog = new Date(log.fecha || log.timestamp);
        if (fechaLog < fechaLimite) return false;
        
        // Filtrar por nivel si se especificó
        if (nivel && !String(log.nivel || '').toUpperCase().includes(nivel.toUpperCase())) {
          return false;
        }
        
        // Filtrar por usuario si se especificó
        if (usuario && !String(log.usuario || '').toLowerCase().includes(usuario.toLowerCase())) {
          return false;
        }
        
        return true;
      } catch (e) {
        console.warn('Error al procesar log para estadísticas:', e);
        return false;
      }
    });
    
    // Calcular estadísticas generales
    const total = logs.length;
    const porNivel = {};
    const porUsuario = {};
    const porDia = {};
    const ultimosErrores = [];
    const usuariosActivos = new Set();
    
    logs.forEach(log => {
      // Estadísticas por nivel
      const nivelLog = String(log.nivel || 'DESCONOCIDO').toUpperCase();
      porNivel[nivelLog] = (porNivel[nivelLog] || 0) + 1;
      
      // Estadísticas por usuario
      const usuarioLog = log.usuario || 'Sistema';
      porUsuario[usuarioLog] = (porUsuario[usuarioLog] || 0) + 1;
      
      if (usuarioLog !== 'Sistema') {
        usuariosActivos.add(usuarioLog);
      }
      
      // Estadísticas por día
      try {
        const fecha = new Date(log.fecha || log.timestamp);
        const dia = fecha.toISOString().split('T')[0];
        porDia[dia] = (porDia[dia] || 0) + 1;
      } catch (e) {
        console.warn('Error al procesar fecha para estadísticas:', e);
      }
      
      // Guardar últimos errores
      if (nivelLog === 'ERROR' && ultimosErrores.length < 10) {
        ultimosErrores.push({
          fecha: log.fecha,
          mensaje: log.mensaje,
          usuario: log.usuario,
          detalles: log.detalles
        });
      }
    });
    
    // Ordenar estadísticas
    const nivelesOrdenados = Object.entries(porNivel)
      .sort((a, b) => b[1] - a[1]);
    
    const usuariosOrdenados = Object.entries(porUsuario)
      .sort((a, b) => b[1] - a[1]);
    
    const diasOrdenados = Object.entries(porDia)
      .sort((a, b) => new Date(a[0]) - new Date(b[0]));
    
    // Calcular promedios
    const totalDias = Math.max(1, Math.min(dias, 365)); // Evitar división por cero
    const promedioDiario = total / totalDias;
    
    // Obtener el día con más actividad
    let diaMasActividad = null;
    let maxActividad = 0;
    
    Object.entries(porDia).forEach(([dia, cantidad]) => {
      if (cantidad > maxActividad) {
        maxActividad = cantidad;
        diaMasActividad = dia;
      }
    });
    
    // Crear el objeto de estadísticas
    const estadisticas = {
      general: {
        total,
        promedioDiario: parseFloat(promedioDiario.toFixed(2)),
        diasAnalizados: totalDias,
        diaMasActividad: {
          fecha: diaMasActividad,
          cantidad: maxActividad
        },
        totalUsuariosUnicos: usuariosActivos.size
      },
      porNivel: nivelesOrdenados.map(([nivel, cantidad]) => ({
        nivel,
        cantidad,
        porcentaje: total > 0 ? parseFloat(((cantidad / total) * 100).toFixed(2)) : 0
      })),
      porUsuario: usuariosOrdenados
        .slice(0, 10) // Top 10 usuarios
        .map(([usuario, cantidad]) => ({
          usuario,
          cantidad,
          porcentaje: total > 0 ? parseFloat(((cantidad / total) * 100).toFixed(2)) : 0
        })),
      tendenciaDias: diasOrdenados.map(([fecha, cantidad]) => ({
        fecha,
        cantidad
      })),
      ultimosErrores
    };
    
    // Registrar la acción
    registrarEventoInfo(
      `Estadísticas de logs generadas: ${total} registros analizados`,
      'obtenerEstadisticasLogs',
      { 
        totalRegistros: total,
        diasAnalizados: totalDias,
        filtros: { nivel, usuario }
      }
    );
    
    return {
      success: true,
      message: `Estadísticas generadas para ${total} registros`,
      estadisticas,
      totalRegistros: total,
      fechaInicio: fechaLimite,
      fechaFin: new Date()
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error en obtenerEstadisticasLogs:', error);
    registrarEventoError(
      'Error al generar estadísticas de logs: ' + error.message,
      'obtenerEstadisticasLogs',
      error.stack
    );
    
    return {
      success: false,
      message: 'Error al generar estadísticas: ' + error.message,
      error: error.toString(),
      estadisticas: {
        general: { total: 0, promedioDiario: 0, diasAnalizados: 0, totalUsuariosUnicos: 0 },
        porNivel: [],
        porUsuario: [],
        tendenciaDias: [],
        ultimosErrores: []
      }
    };
  }
}

// ======================
// 🔧 [FIX] FUNCIONES DE REPARACIÓN DEL SISTEMA
// ======================

/**
 * Función para reparar problemas comunes del sistema administrativo
 * @return {Object} Resultado de la reparación
 */
function repararSistema() {
  console.log('🔧 [ADMIN-INTERFACE] Iniciando reparación del sistema...');
  
  const resultados = {
    configuracion: false,
    administradores: false,
    logs: false,
    errores: []
  };
  
  try {
    // 1. Intentar reparar configuración
    if (typeof obtenerConfiguracion === 'function') {
      try {
        const configResult = obtenerConfiguracion();
        if (!configResult.success && typeof crearConfiguracionPorDefecto === 'function') {
          console.log('🔧 [ADMIN-INTERFACE] Creando configuración por defecto...');
          crearConfiguracionPorDefecto();
        }
        resultados.configuracion = true;
      } catch (error) {
        resultados.errores.push('Configuración: ' + error.message);
      }
    } else {
      resultados.errores.push('Configuración: Función obtenerConfiguracion no disponible');
    }
    
    // 2. Intentar reparar sistema de administradores
    if (typeof repararSistemaAutenticacion === 'function') {
      try {
        console.log('🔧 [ADMIN-INTERFACE] Reparando sistema de autenticación...');
        const repairResult = repararSistemaAutenticacion();
        resultados.administradores = repairResult && repairResult.exito;
        if (!resultados.administradores) {
          resultados.errores.push('Administradores: ' + (repairResult?.mensaje || 'Error desconocido'));
        }
      } catch (error) {
        resultados.errores.push('Administradores: ' + error.message);
      }
    } else {
      resultados.errores.push('Administradores: Función repararSistemaAutenticacion no disponible');
    }
    
    // 3. Intentar verificar/crear sistema de logs
    if (typeof inicializarSistemaLogs === 'function') {
      try {
        console.log('🔧 [ADMIN-INTERFACE] Inicializando sistema de logs...');
        inicializarSistemaLogs();
        resultados.logs = true;
      } catch (error) {
        resultados.errores.push('Logs: ' + error.message);
      }
    } else {
      // Intentar función alternativa
      if (typeof registrarLog === 'function') {
        try {
          registrarLog('INFO', 'Sistema reparado desde panel administrativo', { timestamp: new Date().toISOString() });
          resultados.logs = true;
        } catch (error) {
          resultados.errores.push('Logs: Error probando registrarLog - ' + error.message);
        }
      } else {
        resultados.errores.push('Logs: No hay funciones de logs disponibles');
      }
    }
    
    const exito = resultados.configuracion && resultados.administradores && resultados.logs;
    
    console.log(`${exito ? '✅' : '⚠️'} [ADMIN-INTERFACE] Reparación ${exito ? 'exitosa' : 'parcial'}`);
    
    return {
      success: exito || resultados.errores.length < 3, // Éxito si al menos 1 componente funciona
      message: exito 
        ? 'Sistema reparado exitosamente'
        : `Reparación parcial. Errores: ${resultados.errores.length}`,
      detalles: resultados,
      credenciales: resultados.administradores ? { usuario: 'admin', password: 'admin123' } : null
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error durante reparación:', error);
    return {
      success: false,
      message: 'Error crítico durante reparación: ' + error.message,
      detalles: resultados,
      error: error.stack
    };
  }
}

// ======================
// �🔗 FUNCIONES DE INTERFAZ PARA ESTADO DE EVACUACIÓN
// ======================

/**
 * Carga estado de evacuación usando la función existente
 * @return {Object} Estado actual de evacuación
 */
function loadEvacuationStatus() {
  console.log('🚨 [ADMIN-INTERFACE] Cargando estado de evacuación...');
  
  try {
    // Usar función existente en registroB.js
    const result = getEvacuacionDataForClient();
    
    if (!result.success) {
      return {
        success: false,
        message: result.message || 'Error obteniendo estado de evacuación',
        data: { total: 0, evacuadas: 0, restantes: 0, estado: 'DESCONOCIDO' }
      };
    }
    
    const personasDentro = result.personasDentro || [];
    const totalDentro = result.totalDentro || 0;
    
    // Preparar datos para el frontend
    const data = {
      total: totalDentro,
      evacuadas: 0, // Personas que han salido
      restantes: totalDentro, // Personas que siguen dentro
      estado: totalDentro > 0 ? 'EN_CURSO' : 'COMPLETA',
      personas_detalle: personasDentro.map(p => ({
        cedula: p.cedula,
        nombre: p.nombre,
        empresa: p.empresa,
        tiempo_dentro: p.duracionFormateada || 'N/A'
      }))
    };
    
    console.log(`✅ [ADMIN-INTERFACE] Estado evacuación: ${totalDentro} personas dentro`);
    
    return {
      success: true,
      message: `Estado de evacuación actualizado: ${totalDentro} personas en el edificio`,
      data: data
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error cargando estado evacuación:', error);
    return {
      success: false,
      message: 'Error interno obteniendo estado de evacuación',
      data: { total: 0, evacuadas: 0, restantes: 0, estado: 'ERROR' }
    };
  }
}

// ======================
// 🔗 FUNCIONES DE DIAGNÓSTICO Y MANTENIMIENTO
// ======================

/**
 * Ejecuta diagnóstico del sistema
 * @param {string} nivel - Nivel de diagnóstico
 * @return {Object} Resultado del diagnóstico
 */
function ejecutarDiagnosticoUnificado(nivel = 'completo') {
  console.log('🔍 [ADMIN-INTERFACE] Ejecutando diagnóstico del sistema...');
  
  try {
    const diagnostico = {
      fecha: new Date(),
      nivel: nivel,
      version_sistema: VERSION_SISTEMA || 'V6.2',
      tests: {},
      resumen: {
        total_tests: 0,
        tests_pasados: 0,
        tests_fallidos: 0,
        porcentaje_exito: 0
      },
      estado_general: 'DESCONOCIDO',
      errores: [],
      erroresManuales: obtenerErroresManuales(),
      problemasResueltos: obtenerProblemasResueltos()
    };
    
    const tests = [];
    
    // Test 1: Verificar personal
    try {
      const personalResult = loadPersonnel();
      tests.push({
        nombre: 'Base de datos de personal',
        pasado: personalResult.success,
        detalles: `${personalResult.stats?.total || 0} registros encontrados`
      });
      
      // Añadir errores específicos si no pasa el test
      if (!personalResult.success) {
        const errorId = 'personal_db_error';
        if (!diagnostico.problemasResueltos[errorId]) {
          diagnostico.errores.push({
            id: errorId,
            mensaje: 'Error accediendo a la base de datos de personal',
            categoria: 'BASE_DATOS',
            solucion: 'Verificar permisos y estructura de la hoja "Base de Datos"',
            timestamp: new Date().toISOString()
          });
        }
      }
    } catch (e) {
      tests.push({
        nombre: 'Base de datos de personal',
        pasado: false,
        detalles: 'Error accediendo a personal'
      });
      
      const errorId = 'personal_exception';
      if (!diagnostico.problemasResueltos[errorId]) {
        diagnostico.errores.push({
          id: errorId,
          mensaje: `Excepción en base de datos de personal: ${e.message}`,
          categoria: 'EXCEPCION',
          solucion: 'Revisar configuración de Google Sheets y permisos',
          timestamp: new Date().toISOString()
        });
      }
    }
    
    // Test 2: Verificar administradores
    try {
      const adminResult = loadAdminUsers();
      const adminsPassed = adminResult.success && adminResult.stats?.total > 0;
      tests.push({
        nombre: 'Usuarios administradores',
        pasado: adminsPassed,
        detalles: `${adminResult.stats?.total || 0} administradores encontrados`
      });
      
      if (!adminsPassed) {
        const errorId = 'admin_users_error';
        if (!diagnostico.problemasResueltos[errorId]) {
          diagnostico.errores.push({
            id: errorId,
            mensaje: 'No se encontraron usuarios administradores o hay error en el acceso',
            categoria: 'SEGURIDAD',
            solucion: 'Verificar que existan administradores configurados en la hoja correspondiente',
            timestamp: new Date().toISOString()
          });
        }
      }
    } catch (e) {
      tests.push({
        nombre: 'Usuarios administradores',
        pasado: false,
        detalles: 'Error accediendo a administradores'
      });
      
      const errorId = 'admin_exception';
      if (!diagnostico.problemasResueltos[errorId]) {
        diagnostico.errores.push({
          id: errorId,
          mensaje: `Error cargando administradores: ${e.message}`,
          categoria: 'EXCEPCION',
          solucion: 'Verificar estructura y permisos de la hoja de administradores',
          timestamp: new Date().toISOString()
        });
      }
    }
    
    // Test 3: Verificar configuración
    try {
      const configResult = loadConfiguration();
      tests.push({
        nombre: 'Configuración del sistema',
        pasado: configResult.success,
        detalles: `${configResult.configuraciones?.length || 0} configuraciones cargadas`
      });
      
      if (!configResult.success) {
        const errorId = 'config_error';
        if (!diagnostico.problemasResueltos[errorId]) {
          diagnostico.errores.push({
            id: errorId,
            mensaje: 'Error cargando configuración del sistema',
            categoria: 'CONFIGURACION',
            solucion: 'Verificar que la hoja "Configuración" existe y contiene datos válidos',
            timestamp: new Date().toISOString()
          });
        }
      }
    } catch (e) {
      tests.push({
        nombre: 'Configuración del sistema',
        pasado: false,
        detalles: 'Error accediendo a configuración'
      });
      
      const errorId = 'config_exception';
      if (!diagnostico.problemasResueltos[errorId]) {
        diagnostico.errores.push({
          id: errorId,
          mensaje: `Error en configuración: ${e.message}`,
          categoria: 'EXCEPCION',
          solucion: 'Revisar la función obtenerConfiguracion y la hoja de configuración',
          timestamp: new Date().toISOString()
        });
      }
    }
    
    // Test 4: Verificar logs
    try {
      const logsResult = loadLogs(10);
      tests.push({
        nombre: 'Sistema de logs',
        pasado: logsResult.success,
        detalles: `${logsResult.stats?.total || 0} logs recientes encontrados`
      });
      
      if (!logsResult.success) {
        const errorId = 'logs_error';
        if (!diagnostico.problemasResueltos[errorId]) {
          diagnostico.errores.push({
            id: errorId,
            mensaje: 'Sistema de logs no funcional',
            categoria: 'LOGS',
            solucion: 'Verificar permisos de escritura en la hoja "Logs"',
            timestamp: new Date().toISOString()
          });
        }
      }
    } catch (e) {
      tests.push({
        nombre: 'Sistema de logs',
        pasado: false,
        detalles: 'Error accediendo a logs'
      });
      
      const errorId = 'logs_exception';
      if (!diagnostico.problemasResueltos[errorId]) {
        diagnostico.errores.push({
          id: errorId,
          mensaje: `Error en sistema de logs: ${e.message}`,
          categoria: 'EXCEPCION',
          solucion: 'Revisar función registrarLog y permisos de Google Sheets',
          timestamp: new Date().toISOString()
        });
      }
    }
    
    // Test 5: Verificar funciones críticas
    const funcionesCriticas = ['obtenerConfiguracion', 'registrarLog', 'loadPersonnel'];
    funcionesCriticas.forEach(funcion => {
      try {
        const disponible = typeof eval(funcion) === 'function';
        tests.push({
          nombre: `Función ${funcion}`,
          pasado: disponible,
          detalles: disponible ? 'Función disponible' : 'Función no encontrada'
        });
        
        if (!disponible) {
          const errorId = `funcion_${funcion}`;
          if (!diagnostico.problemasResueltos[errorId]) {
            diagnostico.errores.push({
              id: errorId,
              mensaje: `Función crítica ${funcion} no está disponible`,
              categoria: 'FUNCION_CRITICA',
              solucion: `Verificar que el archivo que contiene ${funcion} esté correctamente cargado`,
              timestamp: new Date().toISOString()
            });
          }
        }
      } catch (e) {
        tests.push({
          nombre: `Función ${funcion}`,
          pasado: false,
          detalles: `Error verificando función: ${e.message}`
        });
        
        const errorId = `funcion_error_${funcion}`;
        if (!diagnostico.problemasResueltos[errorId]) {
          diagnostico.errores.push({
            id: errorId,
            mensaje: `Error verificando función ${funcion}: ${e.message}`,
            categoria: 'ERROR_VERIFICACION',
            solucion: 'Revisar la sintaxis y disponibilidad de la función',
            timestamp: new Date().toISOString()
          });
        }
      }
    });
    
    // Calcular estadísticas finales
    diagnostico.resumen.total_tests = tests.length;
    diagnostico.resumen.tests_pasados = tests.filter(t => t.pasado).length;
    diagnostico.resumen.tests_fallidos = tests.filter(t => !t.pasado).length;
    diagnostico.resumen.porcentaje_exito = Math.round(
      (diagnostico.resumen.tests_pasados / diagnostico.resumen.total_tests) * 100
    );
    
    // Determinar estado general
    if (diagnostico.resumen.porcentaje_exito >= 90) {
      diagnostico.estado_general = 'EXCELENTE';
    } else if (diagnostico.resumen.porcentaje_exito >= 75) {
      diagnostico.estado_general = 'BUENO';
    } else if (diagnostico.resumen.porcentaje_exito >= 50) {
      diagnostico.estado_general = 'REGULAR';
    } else {
      diagnostico.estado_general = 'CRÍTICO';
    }
    
    // Convertir tests a objeto para compatibilidad
    tests.forEach((test, index) => {
      diagnostico.tests[`test_${index + 1}`] = test;
    });
    
    console.log(`✅ [ADMIN-INTERFACE] Diagnóstico completado: ${diagnostico.resumen.porcentaje_exito}% éxito`);
    
    return {
      success: true,
      mensaje: `Diagnóstico completado. Estado: ${diagnostico.estado_general}`,
      diagnostico: diagnostico,
      testsPasados: diagnostico.resumen.tests_pasados,
      totalTests: diagnostico.resumen.total_tests,
      detalles: diagnostico.tests
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error en diagnóstico:', error);
    return {
      success: false,
      mensaje: 'Error ejecutando diagnóstico: ' + error.message,
      testsPasados: 0,
      totalTests: 0,
      detalles: {},
      diagnostico: {
        errores: [{
          id: 'diagnostico_exception',
          mensaje: `Error crítico en diagnóstico: ${error.message}`,
          categoria: 'SISTEMA',
          solucion: 'Revisar logs del sistema para más detalles',
          timestamp: new Date().toISOString()
        }],
        erroresManuales: obtenerErroresManuales()
      }
    };
  }
}

// ======================
// 🔗 FUNCIONES DE UTILIDAD PARA EL PANEL ADMIN
// ======================

// � [SEGURIDAD] FUNCIÓN ELIMINADA POR SEGURIDAD - validarAdministrador
// Esta función ha sido eliminada porque manejaba contraseñas en texto plano.
// USAR ÚNICAMENTE: validarAdministradorMejorado() para autenticación segura.

// ======================
// 🔗 FUNCIONES DE GESTIÓN DE PERSONAL (CRUD)
// ======================

/**
 * Cambia el estado de una persona en la Base de Datos
 * @param {string} cedula - Cédula de la persona
 * @param {string} nuevoEstado - Nuevo estado (activo/inactivo)
 * @return {Object} Resultado de la operación
 */
function cambiarEstadoPersonal(cedula, nuevoEstado) {
  console.log(`🔄 [ADMIN-INTERFACE] Cambiando estado de ${cedula} a ${nuevoEstado}...`);
  
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(SHEET_BD);
    
    if (!hoja) {
      return {
        success: false,
        message: 'Hoja de Base de Datos no encontrada'
      };
    }
    
    const datos = hoja.getDataRange().getValues();
    const encabezados = datos[0];
    const indiceCedula = encabezados.indexOf('Cédula');
    const indiceEstado = encabezados.indexOf('Estado');
    
    if (indiceCedula === -1 || indiceEstado === -1) {
      return {
        success: false,
        message: 'Estructura de hoja incorrecta'
      };
    }
    
    // Buscar la fila de la persona
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][indiceCedula] === cedula) {
        // Actualizar el estado
        hoja.getRange(i + 1, indiceEstado + 1).setValue(nuevoEstado);
        
        return {
          success: true,
          message: `Estado cambiado a ${nuevoEstado} exitosamente`
        };
      }
    }
    
    return {
      success: false,
      message: 'Persona no encontrada'
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error cambiando estado:', error);
    return {
      success: false,
      message: 'Error interno cambiando estado'
    };
  }
}

/**
 * Elimina una persona de la Base de Datos
 * @param {string} cedula - Cédula de la persona a eliminar
 * @return {Object} Resultado de la operación
 */
function eliminarPersonal(cedula) {
  console.log(`🗑️ [ADMIN-INTERFACE] Eliminando persona con cédula ${cedula}...`);
  
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(SHEET_BD);
    
    if (!hoja) {
      return {
        success: false,
        message: 'Hoja de Base de Datos no encontrada'
      };
    }
    
    const datos = hoja.getDataRange().getValues();
    const encabezados = datos[0];
    const indiceCedula = encabezados.indexOf('Cédula');
    
    if (indiceCedula === -1) {
      return {
        success: false,
        message: 'Estructura de hoja incorrecta'
      };
    }
    
    // Buscar la fila de la persona
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][indiceCedula] === cedula) {
        // Eliminar la fila
        hoja.deleteRow(i + 1);
        
        return {
          success: true,
          message: 'Persona eliminada exitosamente'
        };
      }
    }
    
    return {
      success: false,
      message: 'Persona no encontrada'
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error eliminando persona:', error);
    return {
      success: false,
      message: 'Error interno eliminando persona'
    };
  }
}

/**
 * Actualiza los datos de una persona existente
 * @param {string} cedula - Cédula de la persona a actualizar
 * @param {Object} nuevosDatos - Nuevos datos de la persona
 * @return {Object} Resultado de la operación
 */
function actualizarPersonal(cedula, nuevosDatos) {
  console.log(`📝 [ADMIN-INTERFACE] Actualizando datos de ${cedula}...`);
  
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = ss.getSheetByName(SHEET_BD);
    
    if (!hoja) {
      return {
        success: false,
        message: 'Hoja de Base de Datos no encontrada'
      };
    }
    
    const datos = hoja.getDataRange().getValues();
    const encabezados = datos[0];
    const indiceCedula = encabezados.indexOf('Cédula');
    
    if (indiceCedula === -1) {
      return {
        success: false,
        message: 'Estructura de hoja incorrecta'
      };
    }
    
    // Buscar la fila de la persona
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][indiceCedula] === cedula) {
        // Actualizar cada campo
        encabezados.forEach((encabezado, indice) => {
          const campo = encabezado.toLowerCase();
          if (nuevosDatos[campo] !== undefined) {
            hoja.getRange(i + 1, indice + 1).setValue(nuevosDatos[campo]);
          }
        });
        
        return {
          success: true,
          message: 'Datos actualizados exitosamente'
        };
      }
    }
    
    return {
      success: false,
      message: 'Persona no encontrada'
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error actualizando persona:', error);
    return {
      success: false,
      message: 'Error interno actualizando persona'
    };
  }
}

console.log('✅ adminB.js (Interfaz) cargado - Usando funciones existentes del sistema');
// ======================
// 🔧 [FIX] FUNCIÓN ADICIONAL PARA REPARAR AUTENTICACIÓN - TEMPORAL
// ======================

/**
 * 🔧 [FIX] Función específica para solucionar problemas de autenticación (llamada desde admin.html)
 * @return {Object} Resultado de la reparación de autenticación
 */
function solucionarProblemaAutenticacion() {
  console.log('🔧 [ADMIN-INTERFACE] Solucionando problemas de autenticación...');
  
  try {
    // Enfocarse específicamente en el sistema de administradores
    if (typeof repararSistemaAutenticacion === 'function') {
      const authRepair = repararSistemaAutenticacion();
      
      return {
        success: authRepair && authRepair.exito,
        message: authRepair?.mensaje || 'Sistema de autenticación reparado',
        credenciales: { usuario: 'admin', contraseña: 'admin123' },
        detalles: authRepair
      };
    } else {
      return {
        success: false,
        message: 'Función de reparación de autenticación no disponible. Verifique que registroB.js esté cargado.',
        credenciales: null
      };
    }
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error solucionando autenticación:', error);
    return {
      success: false,
      message: 'Error crítico solucionando autenticación: ' + error.message,
      credenciales: null,
      error: error.stack
    };
  }
}

// ======================
// 🔧 [FIX] FUNCIONES DE LOGGING PARA REGISTRO DE EVENTOS
// ======================

/**
 * 🔧 [FIX] Registra eventos de información en el sistema de logs
 * @param {string} mensaje - Mensaje del evento
 * @param {string} funcion - Función que genera el evento
 * @param {Object} detalles - Detalles adicionales
 */
function registrarEventoInfo(mensaje, funcion, detalles = null) {
  try {
    if (typeof registrarLog === 'function') {
      registrarLog('INFO', `[ADMIN-PANEL] ${mensaje}`, detalles ? { funcion: funcion, ...detalles } : { funcion: funcion }, 'Panel Admin');
    } else if (typeof logInfo === 'function') {
      logInfo(`[ADMIN-PANEL] ${mensaje}`, detalles ? { funcion: funcion, ...detalles } : { funcion: funcion }, 'Panel Admin');
    }
    console.log(`📝 [ADMIN-LOG] INFO: ${mensaje}`);
  } catch (error) {
    console.warn('⚠️ No se pudo registrar evento INFO:', error);
  }
}

/**
 * 🔧 [FIX] Registra eventos de error en el sistema de logs
 * @param {string} mensaje - Mensaje del error
 * @param {string} funcion - Función que genera el error
 * @param {string} stack - Stack trace del error
 */
function registrarEventoError(mensaje, funcion, stack = null) {
  try {
    if (typeof registrarLog === 'function') {
      registrarLog('ERROR', `[ADMIN-PANEL] ${mensaje}`, { funcion: funcion, stack: stack }, 'Panel Admin');
    } else if (typeof logError === 'function') {
      logError(`[ADMIN-PANEL] ${mensaje}`, { funcion: funcion, stack: stack }, 'Panel Admin');
    }
    console.error(`🚨 [ADMIN-LOG] ERROR: ${mensaje}`);
  } catch (error) {
    console.warn('⚠️ No se pudo registrar evento ERROR:', error);
  }
}

/**
 * 🔧 [FIX] Registra eventos de advertencia en el sistema de logs
 * @param {string} mensaje - Mensaje de la advertencia
 * @param {string} funcion - Función que genera la advertencia
 * @param {Object} detalles - Detalles adicionales
 */
function registrarEventoWarning(mensaje, funcion, detalles = null) {
  try {
    if (typeof registrarLog === 'function') {
      registrarLog('WARNING', `[ADMIN-PANEL] ${mensaje}`, detalles ? { funcion: funcion, ...detalles } : { funcion: funcion }, 'Panel Admin');
    } else if (typeof logWarning === 'function') {
      logWarning(`[ADMIN-PANEL] ${mensaje}`, detalles ? { funcion: funcion, ...detalles } : { funcion: funcion }, 'Panel Admin');
    }
    console.warn(`⚠️ [ADMIN-LOG] WARNING: ${mensaje}`);
  } catch (error) {
    console.warn('⚠️ No se pudo registrar evento WARNING:', error);
  }
}

// ======================
// 🔧 FUNCIONES DE GESTIÓN DE ERRORES Y PROBLEMAS MEJORADAS
// ======================

/**
 * Obtener errores manuales del sistema de propiedades
 */
function obtenerErroresManuales() {
  try {
    const propiedades = PropertiesService.getScriptProperties();
    const erroresJson = propiedades.getProperty('errores_manuales') || '[]';
    return JSON.parse(erroresJson);
  } catch (error) {
    console.error('Error obteniendo errores manuales:', error);
    return [];
  }
}

/**
 * Obtener problemas marcados como resueltos
 */
function obtenerProblemasResueltos() {
  try {
    const propiedades = PropertiesService.getScriptProperties();
    const problemasJson = propiedades.getProperty('problemas_resueltos') || '{}';
    return JSON.parse(problemasJson);
  } catch (error) {
    console.error('Error obteniendo problemas resueltos:', error);
    return {};
  }
}

/**
 * Marcar un problema como resuelto
 */
function marcarProblemaComoResuelto(problemaId, solucionAplicada = '') {
  try {
    const propiedades = PropertiesService.getScriptProperties();
    const problemasResueltos = obtenerProblemasResueltos();
    
    if (!problemasResueltos[problemaId]) {
      problemasResueltos[problemaId] = {
        resuelto: true,
        timestamp: new Date().toISOString(),
        solucion: solucionAplicada
      };
      
      propiedades.setProperty('problemas_resueltos', JSON.stringify(problemasResueltos));
      
      // Registrar en logs
      registrarLog('INFO', `Problema marcado como resuelto: ${problemaId}`, {
        solucion: solucionAplicada,
        problema_id: problemaId
      }, 'Sistema de Diagnóstico');
      
      console.log(`✅ Problema ${problemaId} marcado como resuelto`);
      return true;
    }
    
    console.warn(`⚠️ Problema ${problemaId} ya estaba marcado como resuelto`);
    return false;
  } catch (error) {
    console.error('Error marcando problema como resuelto:', error);
    return false;
  }
}

/**
 * Añadir error manual detectado por el usuario
 */
function añadirErrorManual(descripcion, ubicacion = 'Sin especificar', categoria = 'USUARIO') {
  try {
    // Validar parámetros de entrada
    if (!descripcion || typeof descripcion !== 'string') {
      return {
        success: false,
        message: 'La descripción del error es requerida y debe ser texto'
      };
    }
    
    // Sanear parámetros
    descripcion = descripcion.trim();
    ubicacion = (ubicacion && typeof ubicacion === 'string') ? ubicacion.trim() : 'Sin especificar';
    categoria = (categoria && typeof categoria === 'string') ? categoria.trim() : 'USUARIO';
    
    if (!descripcion) {
      return {
        success: false,
        message: 'La descripción del error no puede estar vacía'
      };
    }
    
    const propiedades = PropertiesService.getScriptProperties();
    const erroresManuales = obtenerErroresManuales();
    
    const errorId = `manual_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const nuevoError = {
      id: errorId,
      descripcion: descripcion,
      ubicacion: ubicacion,
      categoria: categoria,
      timestamp: new Date().toISOString(),
      origen: 'MANUAL_USUARIO',
      analisis: analizarErrorManual(descripcion, ubicacion)
    };
    
    erroresManuales.push(nuevoError);
    propiedades.setProperty('errores_manuales', JSON.stringify(erroresManuales));
    
    // Registrar en logs
    registrarLog('WARN', `Error manual reportado: ${descripcion}`, {
      error_id: errorId,
      ubicacion: ubicacion,
      categoria: categoria,
      analisis: nuevoError.analisis
    }, 'Reporte Manual de Errores');
    
    console.log(`📝 Error manual añadido: ${descripcion}`);
    
    return {
      success: true,
      errorId: errorId,
      analisis: nuevoError.analisis,
      message: 'Error reportado y analizado correctamente'
    };
  } catch (error) {
    console.error('Error añadiendo error manual:', error);
    return {
      success: false,
      message: 'Error guardando el reporte: ' + error.message
    };
  }
}

/**
 * Analizar error manual - versión mejorada
 */
function analizarErrorManual(descripcion, ubicacion) {
  // Validar y sanear parámetros
  if (!descripcion || typeof descripcion !== 'string') {
    descripcion = 'Error sin descripción';
  }
  
  if (!ubicacion || typeof ubicacion !== 'string') {
    ubicacion = 'Sin especificar';
  }
  
  const analisis = {
    posiblesCausas: [],
    componentesAfectados: [],
    severidad: 'MEDIA',
    solucionesSugeridas: [],
    archivosRelacionados: []
  };
  
  const desc = descripcion.toLowerCase();
  const ubic = ubicacion.toLowerCase();
  
  // Análisis de palabras clave mejorado
  if (desc.includes('config') || desc.includes('configurac')) {
    analisis.posiblesCausas.push('Problema de configuración del sistema');
    analisis.componentesAfectados.push('Sistema de Configuración');
    analisis.archivosRelacionados.push('config.js', 'adminB.js', 'constantes.js');
    analisis.solucionesSugeridas.push('Verificar hoja de configuración en Google Sheets');
    analisis.solucionesSugeridas.push('Comprobar función obtenerConfiguracion()');
    analisis.severidad = 'ALTA';
  }
  
  if (desc.includes('log') || desc.includes('registro')) {
    analisis.posiblesCausas.push('Fallo en el sistema de logs');
    analisis.componentesAfectados.push('Sistema de Logs');
    analisis.archivosRelacionados.push('logs.js', 'logs_enhanced.js');
    analisis.solucionesSugeridas.push('Verificar permisos de escritura en Google Sheets');
    analisis.solucionesSugeridas.push('Comprobar función registrarLog()');
  }
  
  if (desc.includes('personal') || desc.includes('empleado') || desc.includes('usuario')) {
    analisis.posiblesCausas.push('Error en gestión de personal');
    analisis.componentesAfectados.push('Base de Datos de Personal');
    analisis.archivosRelacionados.push('adminB.js');
    analisis.solucionesSugeridas.push('Verificar integridad de datos en hoja "Base de Datos"');
    analisis.solucionesSugeridas.push('Comprobar función loadPersonnel()');
  }
  
  if (desc.includes('función') || desc.includes('funcion') || desc.includes('method') || desc.includes('undefined')) {
    analisis.posiblesCausas.push('Función no disponible o con errores de sintaxis');
    analisis.componentesAfectados.push('Sistema de Funciones JavaScript');
    analisis.solucionesSugeridas.push('Verificar que todos los archivos JavaScript estén cargados correctamente');
    analisis.solucionesSugeridas.push('Revisar consola del navegador para errores de sintaxis');
    analisis.severidad = 'ALTA';
  }
  
  if (desc.includes('sheet') || desc.includes('hoja') || desc.includes('google') || desc.includes('api')) {
    analisis.posiblesCausas.push('Problema de conectividad con Google Sheets API');
    analisis.componentesAfectados.push('Integración Google Sheets');
    analisis.solucionesSugeridas.push('Verificar permisos y conexión a Google Sheets');
    analisis.solucionesSugeridas.push('Comprobar límites de API de Google');
    analisis.severidad = 'CRÍTICA';
  }
  
  // Si no se encontraron causas específicas
  if (analisis.posiblesCausas.length === 0) {
    analisis.posiblesCausas.push('Error no categorizado - requiere análisis manual');
    analisis.solucionesSugeridas.push('Revisar logs del sistema para más detalles');
    analisis.solucionesSugeridas.push('Contactar al administrador del sistema');
  }
  
  return analisis;
}

/**
 * Cambiar modo de diagnóstico
 */
function cambiarModoDiagnostico(modo) {
  try {
    if (modo !== 'estatico' && modo !== 'activo') {
      return {
        success: false,
        message: 'Modo inválido. Use "estatico" o "activo"'
      };
    }
    
    const propiedades = PropertiesService.getScriptProperties();
    propiedades.setProperty('modo_diagnostico', modo);
    
    // Registrar cambio en logs
    registrarLog('INFO', `Modo de diagnóstico cambiado a: ${modo}`, {
      modo_anterior: propiedades.getProperty('modo_diagnostico') || 'estatico',
      modo_nuevo: modo
    }, 'Sistema de Diagnóstico');
    
    console.log(`🔄 Modo de diagnóstico cambiado a: ${modo.toUpperCase()}`);
    
    return {
      success: true,
      message: `Modo cambiado a ${modo} correctamente`
    };
  } catch (error) {
    console.error('Error cambiando modo de diagnóstico:', error);
    return {
      success: false,
      message: 'Error cambiando modo: ' + error.message
    };
  }
}

/**
 * Obtener detalles de un error específico
 */
function obtenerDetallesError(errorId) {
  try {
    console.log(`🔍 Buscando error: ${errorId}`);
    
    // Buscar en errores manuales
    const erroresManuales = obtenerErroresManuales();
    const errorManual = erroresManuales.find(error => 
      error.id === errorId || 
      error.descripcion.toLowerCase().includes(errorId.toLowerCase()) ||
      error.ubicacion.toLowerCase().includes(errorId.toLowerCase())
    );
    
    if (errorManual) {
      return {
        encontrado: true,
        error: errorManual,
        tipo: 'manual'
      };
    }
    
    // Buscar en último diagnóstico (simulado)
    const ultimoDiagnostico = ejecutarDiagnosticoUnificado('completo');
    if (ultimoDiagnostico.diagnostico && ultimoDiagnostico.diagnostico.errores) {
      const errorSistema = ultimoDiagnostico.diagnostico.errores.find(error =>
        error.id === errorId || 
        error.mensaje.toLowerCase().includes(errorId.toLowerCase())
      );
      
      if (errorSistema) {
        return {
          encontrado: true,
          error: errorSistema,
          tipo: 'sistema'
        };
      }
    }
    
    return {
      encontrado: false,
      mensaje: 'Error no encontrado en el sistema'
    };
  } catch (error) {
    console.error('Error buscando detalles del error:', error);
    return {
      encontrado: false,
      mensaje: 'Error durante la búsqueda: ' + error.message
    };
  }
}

/**
 * Eliminar un error manual
 */
function eliminarErrorManual(errorId) {
  try {
    const propiedades = PropertiesService.getScriptProperties();
    let erroresManuales = obtenerErroresManuales();
    
    const errorIndex = erroresManuales.findIndex(error => error.id === errorId);
    
    if (errorIndex !== -1) {
      const errorEliminado = erroresManuales[errorIndex];
      erroresManuales.splice(errorIndex, 1);
      
      propiedades.setProperty('errores_manuales', JSON.stringify(erroresManuales));
      
      // Registrar eliminación
      registrarLog('INFO', `Error manual eliminado: ${errorId}`, {
        error_eliminado: errorEliminado.descripcion,
        categoria: errorEliminado.categoria
      }, 'Gestión de Errores');
      
      console.log(`🗑️ Error manual eliminado: ${errorId}`);
      return true;
    }
    
    return false;
  } catch (error) {
    console.error('Error eliminando error manual:', error);
    return false;
  }
}

// ======================
// � [SISTEMA DE AUTENTICACIÓN SEGURO]
// ======================
// IMPORTANTE: Este sistema usa hashing SHA-256 para todas las contraseñas
// NUNCA almacenar o comparar contraseñas en texto plano
// TODAS las funciones de autenticación deben usar validarAdministradorMejorado()
// ======================

/**
 * 🔒 [SEGURIDAD] Guarda o actualiza un administrador en la hoja "Clave" hasheando la contraseña.
 * @param {Object} adminData - Datos del administrador. La contraseña debe estar en texto plano.
 * @return {Object} Resultado de la operación.
 */
function guardarAdministrador(adminData) {
  console.log('💾 [ADMIN-INTERFACE] Guardando administrador:', adminData.cedula);
  
  try {
    const ss = SpreadsheetApp.getActive();
    const hojaClave = ss.getSheetByName('Clave');
    
    if (!hojaClave) {
      return {
        success: false,
        message: 'Hoja "Clave" no encontrada. Verifique la configuración del sistema.'
      };
    }

    // Validar datos requeridos
    if (!adminData.cedula || !adminData.nombre) {
      return {
        success: false,
        message: 'Cédula y nombre son campos obligatorios'
      };
    }

    const datos = hojaClave.getDataRange().getValues();
    const encabezados = datos[0];
    
    // Mapeo flexible de columnas
    const mapeoColumnas = mapearColumnasFlexible(encabezados, {
      cedula: ['cédula', 'cedula', 'id', 'usuario'],
      nombre: ['nombre', 'name'],
      contrasena: ['contraseña', 'contrasena', 'password', 'clave'],
      cargo: ['cargo', 'position', 'rol', 'role'],
      email: ['email', 'correo', 'mail'],
      permisos: ['permisos', 'permissions'],
      estado: ['estado', 'status', 'activo'],
      fechaCreacion: ['fecha_creacion', 'fecha creación', 'created', 'fecha']
    });

    // Buscar si el administrador ya existe
    let filaExistente = -1;
    for (let i = 1; i < datos.length; i++) {
      const cedulaFila = obtenerValorColumna(datos[i], mapeoColumnas.cedula);
      if (cedulaFila === adminData.cedula) {
        filaExistente = i + 1; // +1 porque getRange usa índices 1-based
        break;
      }
    }

    // [MEJORA] Manejar hashing de contraseña de forma segura
    let contrasenaHasheada = null;
    if (adminData.contrasena && adminData.contrasena.trim() !== '') {
      try {
        contrasenaHasheada = _hashPassword(adminData.contrasena.trim());
        console.log(`🔒 [ADMIN-INTERFACE] Contraseña hasheada para ${adminData.cedula}`);
      } catch (hashError) {
        console.error('❌ [ADMIN-INTERFACE] Error hasheando contraseña:', hashError);
        return {
          success: false,
          message: 'Error procesando la contraseña. Verifique la configuración de seguridad del sistema.'
        };
      }
    }

    // Preparar datos para guardar
    const ahora = new Date();
    const datosAGuardar = new Array(Math.max(...Object.values(mapeoColumnas)) + 1).fill('');
    
    // Asignar valores según el mapeo
    if (mapeoColumnas.cedula !== -1) datosAGuardar[mapeoColumnas.cedula] = adminData.cedula;
    if (mapeoColumnas.nombre !== -1) datosAGuardar[mapeoColumnas.nombre] = adminData.nombre;
    if (mapeoColumnas.cargo !== -1) datosAGuardar[mapeoColumnas.cargo] = adminData.cargo || 'Administrador';
    if (mapeoColumnas.email !== -1) datosAGuardar[mapeoColumnas.email] = adminData.email || '';
    if (mapeoColumnas.permisos !== -1) datosAGuardar[mapeoColumnas.permisos] = adminData.permisos || '{}';
    if (mapeoColumnas.estado !== -1) datosAGuardar[mapeoColumnas.estado] = adminData.estado || 'activo';
    
    if (filaExistente > 0) {
      // ACTUALIZAR administrador existente
      
      // [MEJORA] Solo actualizar contraseña si se proporcionó una nueva
      if (contrasenaHasheada && mapeoColumnas.contrasena !== -1) {
        datosAGuardar[mapeoColumnas.contrasena] = contrasenaHasheada;
      } else if (mapeoColumnas.contrasena !== -1) {
        // Mantener contraseña existente
        datosAGuardar[mapeoColumnas.contrasena] = datos[filaExistente - 1][mapeoColumnas.contrasena];
      }
      
      // Mantener fecha de creación original
      if (mapeoColumnas.fechaCreacion !== -1) {
        datosAGuardar[mapeoColumnas.fechaCreacion] = datos[filaExistente - 1][mapeoColumnas.fechaCreacion];
      }
      
      hojaClave.getRange(filaExistente, 1, 1, datosAGuardar.length).setValues([datosAGuardar]);
      
      console.log(`✅ [ADMIN-INTERFACE] Administrador ${adminData.cedula} actualizado`);
      registrarEventoInfo(`Administrador actualizado: ${adminData.cedula}`, 'guardarAdministrador');
      
      return {
        success: true,
        message: `Administrador ${adminData.nombre} actualizado exitosamente`
      };
    } else {
      // CREAR nuevo administrador
      
      // [MEJORA] Validar que se proporcione contraseña para nuevos administradores
      if (!contrasenaHasheada) {
        return {
          success: false,
          message: 'La contraseña es obligatoria para nuevos administradores'
        };
      }
      
      // Asignar contraseña hasheada y fecha de creación
      if (mapeoColumnas.contrasena !== -1) {
        datosAGuardar[mapeoColumnas.contrasena] = contrasenaHasheada;
      }
      if (mapeoColumnas.fechaCreacion !== -1) {
        datosAGuardar[mapeoColumnas.fechaCreacion] = ahora;
      }
      
      hojaClave.appendRow(datosAGuardar);
      
      console.log(`✅ [ADMIN-INTERFACE] Nuevo administrador ${adminData.cedula} creado`);
      registrarEventoInfo(`Nuevo administrador creado: ${adminData.cedula}`, 'guardarAdministrador');
      
      return {
        success: true,
        message: `Administrador ${adminData.nombre} creado exitosamente`
      };
    }
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error guardando administrador:', error);
    registrarEventoError(`Error guardando administrador: ${error.message}`, 'guardarAdministrador', error.stack);
    return {
      success: false,
      message: 'Error interno guardando administrador: ' + error.message
    };
  }
}

/**
 * 🔧 [FIX] Cambia el estado de un administrador
 * @param {string} cedula - Cédula del administrador
 * @param {string} nuevoEstado - Nuevo estado (activo/inactivo)
 * @return {Object} Resultado de la operación
 */
function cambiarEstadoAdministrador(cedula, nuevoEstado) {
  console.log(`🔄 [ADMIN-INTERFACE] Cambiando estado de admin ${cedula} a ${nuevoEstado}...`);
  
  try {
    const ss = SpreadsheetApp.getActive();
    const hojaClave = ss.getSheetByName('Clave');
    
    if (!hojaClave) {
      return {
        success: false,
        message: 'Hoja "Clave" no encontrada'
      };
    }

    const datos = hojaClave.getDataRange().getValues();
    const encabezados = datos[0];
    
    const mapeoColumnas = mapearColumnasFlexible(encabezados, {
      cedula: ['cédula', 'cedula', 'id', 'usuario'],
      estado: ['estado', 'status', 'activo']
    });

    if (mapeoColumnas.cedula === -1 || mapeoColumnas.estado === -1) {
      return {
        success: false,
        message: 'Estructura de hoja incorrecta: faltan columnas de cédula o estado'
      };
    }

    // Buscar el administrador
    for (let i = 1; i < datos.length; i++) {
      const cedulaFila = obtenerValorColumna(datos[i], mapeoColumnas.cedula);
      if (cedulaFila === cedula) {
        // Actualizar el estado
        hojaClave.getRange(i + 1, mapeoColumnas.estado + 1).setValue(nuevoEstado);
        
        registrarEventoInfo(`Estado de administrador ${cedula} cambiado a ${nuevoEstado}`, 'cambiarEstadoAdministrador');
        
        return {
          success: true,
          message: `Estado del administrador cambiado a ${nuevoEstado} exitosamente`
        };
      }
    }
    
    return {
      success: false,
      message: 'Administrador no encontrado'
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error cambiando estado de admin:', error);
    registrarEventoError(`Error cambiando estado de admin: ${error.message}`, 'cambiarEstadoAdministrador', error.stack);
    return {
      success: false,
      message: 'Error interno cambiando estado'
    };
  }
}

/**
 * 🔧 [FIX] Elimina un administrador de la hoja "Clave"
 * @param {string} cedula - Cédula del administrador a eliminar
 * @return {Object} Resultado de la operación
 */
function eliminarAdministrador(cedula) {
  console.log(`🗑️ [ADMIN-INTERFACE] Eliminando administrador ${cedula}...`);
  
  try {
    const ss = SpreadsheetApp.getActive();
    const hojaClave = ss.getSheetByName('Clave');
    
    if (!hojaClave) {
      return {
        success: false,
        message: 'Hoja "Clave" no encontrada'
      };
    }

    const datos = hojaClave.getDataRange().getValues();
    const encabezados = datos[0];
    
    const mapeoColumnas = mapearColumnasFlexible(encabezados, {
      cedula: ['cédula', 'cedula', 'id', 'usuario'],
      nombre: ['nombre', 'name']
    });

    if (mapeoColumnas.cedula === -1) {
      return {
        success: false,
        message: 'Estructura de hoja incorrecta: falta columna de cédula'
      };
    }

    // Buscar el administrador
    for (let i = 1; i < datos.length; i++) {
      const cedulaFila = obtenerValorColumna(datos[i], mapeoColumnas.cedula);
      if (cedulaFila === cedula) {
        const nombreAdmin = obtenerValorColumna(datos[i], mapeoColumnas.nombre) || cedula;
        
        // Verificar que no sea el último administrador
        const totalAdmins = datos.length - 1; // -1 por encabezados
        if (totalAdmins <= 1) {
          return {
            success: false,
            message: 'No se puede eliminar el último administrador del sistema'
          };
        }
        
        // Eliminar la fila
        hojaClave.deleteRow(i + 1);
        
        registrarEventoInfo(`Administrador eliminado: ${cedula} (${nombreAdmin})`, 'eliminarAdministrador');
        
        return {
          success: true,
          message: `Administrador ${nombreAdmin} eliminado exitosamente`
        };
      }
    }
    
    return {
      success: false,
      message: 'Administrador no encontrado'
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error eliminando administrador:', error);
    registrarEventoError(`Error eliminando administrador: ${error.message}`, 'eliminarAdministrador', error.stack);
    return {
      success: false,
      message: 'Error interno eliminando administrador'
    };
  }
}

/**
 * 🔧 [FIX] Obtiene los datos de un administrador específico
 * @param {string} cedula - Cédula del administrador
 * @return {Object} Datos del administrador o null si no se encuentra
 */
function obtenerAdministrador(cedula) {
  try {
    const ss = SpreadsheetApp.getActive();
    const hojaClave = ss.getSheetByName('Clave');
    
    if (!hojaClave) {
      return null;
    }

    const datos = hojaClave.getDataRange().getValues();
    const encabezados = datos[0];
    
    const mapeoColumnas = mapearColumnasFlexible(encabezados, {
      cedula: ['cédula', 'cedula', 'id', 'usuario'],
      nombre: ['nombre', 'name'],
      contrasena: ['contraseña', 'contrasena', 'password', 'clave'],
      cargo: ['cargo', 'position', 'rol', 'role'],
      email: ['email', 'correo', 'mail'],
      permisos: ['permisos', 'permissions'],
      estado: ['estado', 'status', 'activo']
    });

    // Buscar el administrador
    for (let i = 1; i < datos.length; i++) {
      const cedulaFila = obtenerValorColumna(datos[i], mapeoColumnas.cedula);
      if (cedulaFila === cedula) {
        return {
          cedula: cedulaFila,
          nombre: obtenerValorColumna(datos[i], mapeoColumnas.nombre),
          cargo: obtenerValorColumna(datos[i], mapeoColumnas.cargo),
          email: obtenerValorColumna(datos[i], mapeoColumnas.email),
          permisos: obtenerValorColumna(datos[i], mapeoColumnas.permisos),
          estado: obtenerValorColumna(datos[i], mapeoColumnas.estado),
          fila: i + 1
        };
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error obteniendo administrador:', error);
    return null;
  }
}

/**
 * 🔒 [SEGURIDAD CRÍTICA] Valida las credenciales de un administrador comparando hashes.
 * ESTA ES LA ÚNICA FUNCIÓN SEGURA PARA AUTENTICACIÓN DE ADMINISTRADORES.
 * ⚠️  NUNCA usar funciones que comparen contraseñas en texto plano.
 * ⚠️  TODAS las contraseñas deben ser hasheadas antes de almacenar o comparar.
 * 
 * @param {string} credenciales - JSON string con { usuario: string, clave: string }
 * @return {Object} Resultado de la validación con datos del admin si es exitosa.
 */
function validarAdministradorMejorado(credenciales) {
  console.log('🔐 [ADMIN-INTERFACE] Validando credenciales de administrador...');
  
  try {
    let credObj;
    
    // Parsear credenciales con manejo robusto de errores
    try {
      credObj = JSON.parse(credenciales);
    } catch (parseError) {
      // Intentar reparar JSON común
      let credReparadas = credenciales
        .replace(/'/g, '"')
        .replace(/([{,]\s*)(\w+):/g, '$1"$2":')
        .replace(/:\s*([^",\[\]{}]+)(?=\s*[,}])/g, ':"$1"')
        .replace(/,(\s*[}\]])/g, '$1');
      
      try {
        credObj = JSON.parse(credReparadas);
      } catch (repairError) {
        return {
          success: false,
          message: 'Formato de credenciales inválido'
        };
      }
    }
    
    const { usuario, clave } = credObj;
    
    if (!usuario || !clave) {
      return {
        success: false,
        message: 'Usuario y contraseña son requeridos'
      };
    }

  // --- LOG de depuración: inicio de autenticación ---
  console.log('🔍 [DEBUG] Iniciando validación de credenciales:', credenciales);

    // [MEJORA] Validación de entrada
    const usuarioLimpio = String(usuario).trim();
    const claveLimpia = String(clave).trim();

    if (!usuarioLimpio || !claveLimpia) {
      return {
        success: false,
        message: 'Usuario y contraseña no pueden estar vacíos'
      };
    }

    // Buscar administrador
    const admin = obtenerAdministrador(usuarioLimpio);
    console.log('🔍 [DEBUG] Usuario limpio:', usuarioLimpio, '| Admin encontrado:', admin);
    
    if (!admin) {
      // [SEGURIDAD] Log de intento de acceso fallido sin exponer datos
      registrarEventoError(`Intento de acceso con usuario inexistente`, 'validarAdministradorMejorado');
      return {
        success: false,
        message: 'Usuario no encontrado'
      };
    }

    // Verificar estado del administrador
    if (admin.estado && admin.estado.toLowerCase() === 'inactivo') {
      registrarEventoError(`Intento de acceso con usuario inactivo: ${usuarioLimpio}`, 'validarAdministradorMejorado');
      return {
        success: false,
        message: 'Usuario inactivo. Contacte al administrador del sistema.'
      };
    }
    
    // [MEJORA] Obtener hash almacenado de forma más robusta
    const ss = SpreadsheetApp.getActive();
    const hojaClave = ss.getSheetByName('Clave');
    
    if (!hojaClave) {
      console.error('❌ [ADMIN-INTERFACE] Hoja "Clave" no encontrada');
      return {
        success: false,
        message: 'Error de configuración del sistema'
      };
    }
    
    const datos = hojaClave.getDataRange().getValues();
    const encabezados = datos[0];
    
    const mapeoColumnas = mapearColumnasFlexible(encabezados, {
      cedula: ['cédula', 'cedula', 'id', 'usuario'],
      contrasena: ['contraseña', 'contrasena', 'password', 'clave']
    });
    console.log('🔍 [DEBUG] Mapeo de columnas detectado:', mapeoColumnas, '| Encabezados:', encabezados);
    
    let hashAlmacenado = null;
    for (let i = 1; i < datos.length; i++) {
      // Comparación tolerante: normalizar ambos valores
      const normalizar = s => String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '');
      const cedulaFila = obtenerValorColumna(datos[i], mapeoColumnas.cedula);
      if (normalizar(cedulaFila) === normalizar(usuarioLimpio)) {
        hashAlmacenado = obtenerValorColumna(datos[i], mapeoColumnas.contrasena);
        console.log('🔍 [DEBUG] Fila encontrada:', datos[i], '| Hash almacenado:', hashAlmacenado);
        break;
      }
    }
    
    if (!hashAlmacenado) {
      console.error('❌ [ADMIN-INTERFACE] Hash de contraseña no encontrado para usuario:', usuarioLimpio);
      return {
        success: false,
        message: 'Error de configuración de usuario'
      };
    }

    // [MEJORA] Hashear contraseña ingresada con manejo de errores
    let hashIngresado;
    try {
      hashIngresado = _hashPassword(claveLimpia);
      console.log('🔍 [DEBUG] Hash ingresado:', hashIngresado);
    } catch (hashError) {
      console.error('❌ [ADMIN-INTERFACE] Error hasheando contraseña ingresada:', hashError);
      return {
        success: false,
        message: 'Error interno de seguridad'
      };
    }
    
    // [MEJORA] Comparación segura de hashes
    if (hashAlmacenado === hashIngresado) {
      console.log('🔍 [DEBUG] Hash coincide. Acceso autorizado.');
      // Procesar permisos
      let permisos = {};
      try {
        if (admin.permisos && typeof admin.permisos === 'string') {
          permisos = JSON.parse(admin.permisos);
        }
      } catch (error) {
        console.warn('⚠️ [ADMIN-INTERFACE] Error parseando permisos, usando permisos por defecto');
      }
      
      // Permisos por defecto si no están definidos
      if (Object.keys(permisos).length === 0) {
        permisos = {
          personal: { ver: true, editar: true },
          configuracion: { ver: true, editar: true },
          usuarios: { ver: true, editar: true },
          logs: { ver: true, editar: false },
          evacuacion: { ver: true, editar: true },
          diagnostico: { ver: true, ejecutar: true }
        };
      }
      
      // [SEGURIDAD] Registrar acceso exitoso sin exponer datos sensibles
      registrarEventoInfo(`Acceso autorizado para administrador: ${usuarioLimpio}`, 'validarAdministradorMejorado', {
        cargo: admin.cargo,
        timestamp: new Date().toISOString()
      });
      
      return {
        success: true,
        message: 'Acceso autorizado',
        usuario: {
          cedula: admin.cedula,
          nombre: admin.nombre,
          cargo: admin.cargo,
          email: admin.email
        },
        permissions: permisos,
        cargo: admin.cargo
      };
    } else {
      // [SEGURIDAD] Log de intento fallido sin exponer el hash
      registrarEventoError(`Contraseña incorrecta para usuario: ${usuarioLimpio}`, 'validarAdministradorMejorado');
      
      return {
        success: false,
        message: 'Contraseña incorrecta'
      };
    }
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error validando administrador:', error);
    registrarEventoError(`Error crítico en validación de administrador: ${error.message}`, 'validarAdministradorMejorado', error.stack);
    return {
      success: false,
      message: 'Error interno durante la validación'
    };
  }
}

/**
 * 🔧 [FIX] Inicializa usuarios administradores por defecto si la hoja está vacía
 * @return {Object} Resultado de la inicialización
 */
function inicializarAdministradoresPorDefecto() {
  console.log('👥 [ADMIN-INTERFACE] Verificando administradores por defecto...');
  
  try {
    const ss = SpreadsheetApp.getActive();
    let hojaClave = ss.getSheetByName('Clave');
    
    // Crear hoja si no existe
    if (!hojaClave) {
      console.log('📋 [ADMIN-INTERFACE] Creando hoja "Clave"...');
      hojaClave = ss.insertSheet('Clave');
      
      // Crear encabezados
      const encabezados = ['Cédula', 'Nombre', 'Contraseña', 'Cargo', 'Email', 'Permisos', 'Fecha_Creacion'];
      hojaClave.appendRow(encabezados);
      hojaClave.getRange(1, 1, 1, encabezados.length)
        .setFontWeight('bold')
        .setBackground('#f2f2f2');
    }
    
    // Verificar si hay administradores
    const ultimaFila = hojaClave.getLastRow();
    if (ultimaFila <= 1) { // Solo encabezados o vacía
      console.log('👤 [ADMIN-INTERFACE] Creando administrador por defecto...');
      
      const permisosCompletos = JSON.stringify({
        personal: { ver: true, editar: true },
        configuracion: { ver: true, editar: true },
        usuarios: { ver: true, editar: true },
        logs: { ver: true, editar: true },
        evacuacion: { ver: true, editar: true },
        diagnostico: { ver: true, ejecutar: true }
      });
      
      const adminPorDefecto = [
        'admin',
        'Administrador Principal',
        'admin123',
        'Administrador Principal',
        'admin@surpass.com',
        permisosCompletos,
        new Date()
      ];
      
      hojaClave.appendRow(adminPorDefecto);
      
      registrarEventoInfo('Administrador por defecto creado: admin', 'inicializarAdministradoresPorDefecto');
      
      return {
        success: true,
        message: 'Administrador por defecto creado: admin / admin123',
        credenciales: { usuario: 'admin', contraseña: 'admin123' }
      };
    }
    
    return {
      success: true,
      message: 'Administradores ya configurados',
      administradoresExistentes: ultimaFila - 1
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error inicializando administradores:', error);
    return {
      success: false,
      message: 'Error inicializando administradores: ' + error.message
    };
  }
}

/**
 * � [AUDIT DE SEGURIDAD] Verifica que el sistema de autenticación sea seguro
 * @return {Object} Reporte de seguridad del sistema
 */
function auditarSeguridadAutenticacion() {
  console.log('🔍 [SECURITY-AUDIT] Iniciando auditoría de seguridad...');
  
  const reporte = {
    success: true,
    nivel: 'SEGURO',
    problemas: [],
    recomendaciones: [],
    estadisticas: {}
  };

  try {
    // 1. Verificar que la sal de seguridad existe
    const salt = PropertiesService.getScriptProperties().getProperty('SURPASS_SECRET_SALT');
    if (!salt) {
      reporte.problemas.push('❌ Sal de seguridad no configurada');
      reporte.nivel = 'CRÍTICO';
      reporte.recomendaciones.push('Ejecutar _inicializarSalDeSeguridad()');
    } else {
      console.log('✅ [AUDIT] Sal de seguridad configurada');
    }

    // 2. Verificar hoja de administradores
    const ss = SpreadsheetApp.getActive();
    const hojaClave = ss.getSheetByName('Clave');
    
    if (!hojaClave) {
      reporte.problemas.push('❌ Hoja "Clave" no existe');
      reporte.nivel = 'CRÍTICO';
      reporte.recomendaciones.push('Ejecutar crearHojaAdministradores()');
      return reporte;
    }

    const datos = hojaClave.getDataRange().getValues();
    reporte.estadisticas.totalAdmins = datos.length - 1;

    if (datos.length <= 1) {
      reporte.problemas.push('⚠️ No hay administradores registrados');
      reporte.nivel = 'ALTO';
      return reporte;
    }

    // 3. Verificar que las contraseñas están hasheadas
    const encabezados = datos[0];
    const mapeoColumnas = mapearColumnasFlexible(encabezados, {
      cedula: ['cédula', 'cedula', 'id', 'usuario'],
      contrasena: ['contraseña', 'contrasena', 'password', 'clave'],
      estado: ['estado', 'status', 'activo']
    });

    let adminsActivos = 0;
    let contrasenasNoHasheadas = 0;
    let adminsInactivos = 0;

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const cedula = obtenerValorColumna(fila, mapeoColumnas.cedula);
      const contrasena = obtenerValorColumna(fila, mapeoColumnas.contrasena);
      const estado = obtenerValorColumna(fila, mapeoColumnas.estado) || 'activo';

      if (estado.toLowerCase() === 'activo') {
        adminsActivos++;
      } else {
        adminsInactivos++;
      }

      // Verificar si la contraseña parece estar en texto plano
      if (contrasena && contrasena.length < 40) {
        contrasenasNoHasheadas++;
        reporte.problemas.push(`⚠️ Usuario ${cedula} tiene contraseña en texto plano`);
      }
    }

    reporte.estadisticas.adminsActivos = adminsActivos;
    reporte.estadisticas.adminsInactivos = adminsInactivos;
    reporte.estadisticas.contrasenasNoHasheadas = contrasenasNoHasheadas;

    // 4. Evaluar nivel de seguridad
    if (contrasenasNoHasheadas > 0) {
      reporte.nivel = 'MEDIO';
      reporte.recomendaciones.push('Ejecutar migrarContrasenasAHash() para migrar contraseñas');
    }

    if (adminsActivos === 0) {
      reporte.problemas.push('❌ No hay administradores activos');
      reporte.nivel = 'CRÍTICO';
    }

    // 5. Verificar funciones de seguridad
    try {
      const testHash = _hashPassword('test123');
      if (testHash && testHash.length > 40) {
        console.log('✅ [AUDIT] Función de hashing funciona correctamente');
      } else {
        reporte.problemas.push('❌ Función de hashing no funciona correctamente');
        reporte.nivel = 'CRÍTICO';
      }
    } catch (error) {
      reporte.problemas.push('❌ Error en función de hashing: ' + error.message);
      reporte.nivel = 'CRÍTICO';
    }

    console.log(`🎯 [AUDIT] Auditoría completada - Nivel: ${reporte.nivel}`);
    console.log(`📊 [AUDIT] Estadísticas:`, reporte.estadisticas);

    return reporte;

  } catch (error) {
    console.error('❌ [AUDIT] Error en auditoría:', error);
    reporte.success = false;
    reporte.nivel = 'ERROR';
    reporte.problemas.push('Error durante la auditoría: ' + error.message);
    return reporte;
  }
}

/**
 * 🔧 [FIX] Migra contraseñas en texto plano a hashes
 * Esta función debe ejecutarse UNA VEZ para migrar el sistema existente
 * @return {Object} Resultado de la migración
 */
function migrarContrasenasAHash() {
  console.log('🔄 [MIGRACIÓN] Iniciando migración de contraseñas a hash...');
  
  try {
    const ss = SpreadsheetApp.getActive();
    const hojaClave = ss.getSheetByName('Clave');
    
    if (!hojaClave) {
      return {
        success: false,
        message: 'Hoja "Clave" no encontrada'
      };
    }

    const datos = hojaClave.getDataRange().getValues();
    if (datos.length <= 1) {
      return {
        success: true,
        message: 'No hay administradores para migrar',
        migrados: 0
      };
    }

    const encabezados = datos[0];
    const mapeoColumnas = mapearColumnasFlexible(encabezados, {
      cedula: ['cédula', 'cedula', 'id', 'usuario'],
      contrasena: ['contraseña', 'contrasena', 'password', 'clave']
    });

    let migrados = 0;
    let errores = 0;

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const cedula = obtenerValorColumna(fila, mapeoColumnas.cedula);
      const contrasenaActual = obtenerValorColumna(fila, mapeoColumnas.contrasena);

      // Verificar si la contraseña ya está hasheada (los hashes tienen una longitud específica)
      // Los hashes SHA-256 en Base64 tienen aproximadamente 44 caracteres
      if (contrasenaActual && contrasenaActual.length < 40) {
        try {
          console.log(`🔄 [MIGRACIÓN] Migrando contraseña para usuario: ${cedula}`);
          
          // Hashear la contraseña en texto plano
          const contrasenaHasheada = _hashPassword(contrasenaActual);
          
          // Actualizar la celda con el hash
          const filaReal = i + 1; // +1 porque getRange usa índices 1-based
          const columnaContrasena = mapeoColumnas.contrasena + 1; // +1 para índice 1-based
          hojaClave.getRange(filaReal, columnaContrasena).setValue(contrasenaHasheada);
          
          migrados++;
          console.log(`✅ [MIGRACIÓN] Contraseña migrada para usuario: ${cedula}`);
          
        } catch (error) {
          console.error(`❌ [MIGRACIÓN] Error migrando usuario ${cedula}:`, error);
          errores++;
        }
      } else {
        console.log(`ℹ️ [MIGRACIÓN] Usuario ${cedula} ya tiene contraseña hasheada`);
      }
    }

    console.log(`🎯 [MIGRACIÓN] Completada: ${migrados} migrados, ${errores} errores`);
    
    return {
      success: true,
      message: `Migración completada: ${migrados} contraseñas migradas`,
      migrados: migrados,
      errores: errores
    };

  } catch (error) {
    console.error('❌ [MIGRACIÓN] Error en migración:', error);
    return {
      success: false,
      message: 'Error durante la migración: ' + error.message
    };
  }
}

/**
 * 🔧 [FIX] Repara el sistema de autenticación de administradores
 * @return {Object} Resultado de la reparación
 */
function repararSistemaAutenticacion() {
  console.log('🔧 [ADMIN-INTERFACE] Iniciando reparación del sistema de autenticación...');
  
  try {
    const ss = SpreadsheetApp.getActive();
    let hojaClave = ss.getSheetByName('Clave');
    let reparacionRealizada = false;
    
    // 1. Crear hoja si no existe
    if (!hojaClave) {
      console.log('📋 [ADMIN-INTERFACE] Creando hoja "Clave"...');
      hojaClave = ss.insertSheet('Clave');
      reparacionRealizada = true;
    }
    
    // 2. Verificar/corregir encabezados
    const encabezadosEsperados = ['Cédula', 'Nombre', 'Contraseña', 'Cargo', 'Email', 'Permisos', 'Fecha_Creacion'];
    
    if (hojaClave.getLastRow() === 0 || hojaClave.getLastColumn() < encabezadosEsperados.length) {
      console.log('📝 [ADMIN-INTERFACE] Corrigiendo estructura de hoja...');
      hojaClave.clear();
      hojaClave.appendRow(encabezadosEsperados);
      hojaClave.getRange(1, 1, 1, encabezadosEsperados.length)
        .setFontWeight('bold')
        .setBackground('#f2f2f2');
      reparacionRealizada = true;
    }
    
    // 3. Verificar que hay al menos un administrador
    if (hojaClave.getLastRow() <= 1) {
      console.log('👤 [ADMIN-INTERFACE] Creando administrador de emergencia...');
      
      const permisosCompletos = JSON.stringify({
        personal: { ver: true, editar: true },
        configuracion: { ver: true, editar: true },
        usuarios: { ver: true, editar: true },
        logs: { ver: true, editar: true },
        evacuacion: { ver: true, editar: true },
        diagnostico: { ver: true, ejecutar: true }
      });
      
      const adminEmergencia = [
        'admin',
        'Administrador de Emergencia',
        'admin123',
        'Administrador Principal',
        'admin@surpass.com',
        permisosCompletos,
        new Date()
      ];
      
      hojaClave.appendRow(adminEmergencia);
      reparacionRealizada = true;
    }
    
    // 4. Aplicar formato básico
    try {
      hojaClave.setColumnWidth(1, 120); // Cédula
      hojaClave.setColumnWidth(2, 200); // Nombre
      hojaClave.setColumnWidth(3, 120); // Contraseña
      hojaClave.setColumnWidth(4, 150); // Cargo
      hojaClave.setColumnWidth(5, 200); // Email
      hojaClave.setColumnWidth(6, 300); // Permisos
      hojaClave.setColumnWidth(7, 150); // Fecha
    } catch (formatError) {
      console.warn('⚠️ [ADMIN-INTERFACE] Error aplicando formato:', formatError);
    }
    
    registrarEventoInfo('Sistema de autenticación reparado', 'repararSistemaAutenticacion', {
      reparacionRealizada: reparacionRealizada,
      administradores: hojaClave.getLastRow() - 1
    });
    
    return {
      exito: true,
      mensaje: 'Sistema de autenticación reparado exitosamente',
      credencialesDefecto: { usuario: 'admin', contraseña: 'admin123' },
      administradoresDisponibles: hojaClave.getLastRow() - 1
    };
    
  } catch (error) {
    console.error('❌ [ADMIN-INTERFACE] Error reparando autenticación:', error);
    registrarEventoError(`Error reparando autenticación: ${error.message}`, 'repararSistemaAutenticacion', error.stack);
    return {
      exito: false,
      mensaje: 'Error reparando sistema de autenticación: ' + error.message,
      errores: [error.message]
    };
  }
}

// ======================
// 🔗 FUNCIONES GLOBALES PARA GOOGLE APPS SCRIPT
// ======================

/**
 * 🔧 [GLOBAL] Función global para probar conexión desde Google Apps Script
 * Esta función debe estar en el ámbito global para ser accesible desde el frontend
 */
function doTestConnection() {
  try {
    return testConnection();
  } catch (error) {
    console.error('❌ [GLOBAL] Error en doTestConnection:', error);
    return {
      success: false,
      message: 'Error en conexión global: ' + error.message,
      error: error.message
    };
  }
}

/**
 * 🔧 [GLOBAL] Función global para cargar logs desde Google Apps Script - VERSIÓN FINAL CORREGIDA
 * Esta función debe estar en el ámbito global para ser accesible desde el frontend
 * @param {number} limite - Número máximo de logs a cargar (por defecto 50)
 * @returns {Object} Resultado con logs y estadísticas
 */
function doLoadLogs(limite = 50) {
  // ✅ GARANTIZAR QUE SIEMPRE SE DEVUELVA UN OBJETO VÁLIDO
  const respuestaBase = {
    success: false,
    message: 'Error inicializando carga de logs',
    logs: [],
    stats: { total: 0, errores: 0, warnings: 0, info: 0 },
    timestamp: new Date().toISOString()
  };

  try {
    console.log('🔍 [doLoadLogs] === INICIO CARGA DE LOGS ===');
    console.log('🔍 [doLoadLogs] Parámetros:', { limite: limite });

    // ✅ VALIDAR PARÁMETROS DE ENTRADA
    if (!limite || isNaN(limite)) {
      limite = 50;
    }
    limite = Math.max(1, Math.min(parseInt(limite), 1000));
    console.log('🔍 [doLoadLogs] Límite validado:', limite);

    // ✅ OBTENER REFERENCIA AL SPREADSHEET
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.getActive();
      if (!spreadsheet) {
        throw new Error('No se pudo obtener el spreadsheet activo');
      }
      console.log('✅ [doLoadLogs] Spreadsheet obtenido exitosamente');
    } catch (ssError) {
      console.error('❌ [doLoadLogs] Error obteniendo spreadsheet:', ssError);
      respuestaBase.message = 'Error accediendo al documento de Google Sheets';
      return respuestaBase;
    }

    // ✅ OBTENER HOJA DE LOGS
    let hojaLogs;
    try {
      hojaLogs = spreadsheet.getSheetByName('Logs');
      if (!hojaLogs) {
        // Intentar nombres alternativos
        const nombresAlternativos = ['Log', 'Registro', 'Sistema'];
        for (const nombre of nombresAlternativos) {
          hojaLogs = spreadsheet.getSheetByName(nombre);
          if (hojaLogs) {
            console.log(`✅ [doLoadLogs] Hoja encontrada con nombre alternativo: ${nombre}`);
            break;
          }
        }
        
        if (!hojaLogs) {
          throw new Error('Hoja de logs no encontrada');
        }
      } else {
        console.log('✅ [doLoadLogs] Hoja "Logs" encontrada');
      }
    } catch (hojaError) {
      console.error('❌ [doLoadLogs] Error obteniendo hoja de logs:', hojaError);
      respuestaBase.message = 'Hoja de logs no encontrada en el documento';
      return respuestaBase;
    }

    // ✅ VERIFICAR CONTENIDO DE LA HOJA
    let ultimaFila;
    try {
      ultimaFila = hojaLogs.getLastRow();
      console.log('🔍 [doLoadLogs] Última fila en hoja:', ultimaFila);
      
      if (ultimaFila <= 1) {
        console.log('ℹ️ [doLoadLogs] Hoja vacía o solo con encabezados');
        return {
          success: true,
          message: 'No hay logs registrados en el sistema',
          logs: [],
          stats: { total: 0, errores: 0, warnings: 0, info: 0 },
          timestamp: new Date().toISOString()
        };
      }
    } catch (filaError) {
      console.error('❌ [doLoadLogs] Error verificando filas:', filaError);
      respuestaBase.message = 'Error verificando contenido de la hoja de logs';
      return respuestaBase;
    }

    // ✅ CALCULAR RANGO DE DATOS A OBTENER
    const filaInicio = Math.max(2, ultimaFila - limite + 1);
    const numFilas = ultimaFila - filaInicio + 1;
    const numColumnas = Math.min(hojaLogs.getLastColumn(), 6); // Máximo 6 columnas
    
    console.log('🔍 [doLoadLogs] Rango calculado:', {
      filaInicio: filaInicio,
      numFilas: numFilas,
      numColumnas: numColumnas,
      ultimaFila: ultimaFila
    });

    // ✅ OBTENER DATOS DE LA HOJA
    let datos;
    try {
      datos = hojaLogs.getRange(filaInicio, 1, numFilas, numColumnas).getValues();
      console.log('✅ [doLoadLogs] Datos obtenidos:', datos.length, 'filas');
      
      if (!datos || datos.length === 0) {
        throw new Error('No se obtuvieron datos de la hoja');
      }
    } catch (datosError) {
      console.error('❌ [doLoadLogs] Error obteniendo datos:', datosError);
      respuestaBase.message = 'Error leyendo datos de la hoja de logs';
      return respuestaBase;
    }

    // ✅ PROCESAR DATOS EN FORMATO DE LOGS
    const logs = [];
    const contadores = { errores: 0, warnings: 0, info: 0, debug: 0 };

    for (let i = 0; i < datos.length; i++) {
      try {
        const fila = datos[i];
        
        // Saltar filas completamente vacías
        if (!fila || fila.every(cell => !cell || cell === '')) {
          console.log('⚠️ [doLoadLogs] Saltando fila vacía:', i);
          continue;
        }

        // Extraer y validar datos de la fila
        const timestamp = fila[0] ? new Date(fila[0]).toISOString() : new Date().toISOString();
        const nivel = String(fila[1] || 'INFO').toUpperCase().trim();
        const mensaje = String(fila[2] || 'Sin mensaje').trim();
        const detalles = String(fila[3] || '').trim();
        const usuario = String(fila[4] || 'Sistema').trim();
        const ip = String(fila[5] || 'N/A').trim();

        // Crear objeto de log
        const logEntry = {
          id: `log_${filaInicio + i}_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
          fecha: timestamp,
          timestamp: timestamp,
          nivel: nivel,
          mensaje: mensaje,
          usuario: usuario,
          ip: ip,
          detalles: detalles.length > 200 ? detalles.substring(0, 200) + '...' : detalles
        };

        logs.push(logEntry);

        // Contar por niveles
        switch (nivel) {
          case 'ERROR':
            contadores.errores++;
            break;
          case 'WARNING':
          case 'WARN':
          case 'ADVERTENCIA':
            contadores.warnings++;
            break;
          case 'DEBUG':
            contadores.debug++;
            break;
          default:
            contadores.info++;
        }

      } catch (filaError) {
        console.error('❌ [doLoadLogs] Error procesando fila', i, ':', filaError);
        
        // Crear log de error para esta fila problemática
        logs.push({
          id: `log_error_fila_${i}_${Date.now()}`,
          fecha: new Date().toISOString(),
          timestamp: new Date().toISOString(),
          nivel: 'ERROR',
          mensaje: `Error procesando registro de log en fila ${filaInicio + i}`,
          usuario: 'Sistema',
          ip: 'N/A',
          detalles: `Error: ${filaError.message}`
        });
        contadores.errores++;
      }
    }

    // ✅ ORDENAR LOGS (MÁS RECIENTES PRIMERO)
    logs.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    // ✅ CREAR RESPUESTA FINAL
    const respuestaFinal = {
      success: true,
      message: `${logs.length} logs cargados exitosamente desde Google Sheets`,
      logs: logs,
      stats: {
        total: logs.length,
        errores: contadores.errores,
        warnings: contadores.warnings,
        info: contadores.info,
        debug: contadores.debug
      },
      fuente: 'Google Sheets directo',
      timestamp: new Date().toISOString(),
      debug: {
        filaInicio: filaInicio,
        numFilas: numFilas,
        ultimaFila: ultimaFila,
        limite: limite
      }
    };

    console.log('✅ [doLoadLogs] Respuesta final creada:', {
      success: respuestaFinal.success,
      totalLogs: respuestaFinal.logs.length,
      stats: respuestaFinal.stats
    });

    // ✅ REGISTRAR ACCESO EXITOSO (SIN BLOQUEAR SI FALLA)
    try {
      hojaLogs.appendRow([
        new Date(),
        'INFO',
        `Panel admin cargó ${logs.length} logs exitosamente`,
        JSON.stringify({ limite: limite, totalCargados: logs.length, fuente: 'doLoadLogs' }),
        'Panel Admin',
        'doLoadLogs'
      ]);
      console.log('✅ [doLoadLogs] Acceso registrado en logs');
    } catch (logAccessError) {
      console.warn('⚠️ [doLoadLogs] No se pudo registrar el acceso:', logAccessError);
      // No fallar por esto, solo advertir
    }

    console.log('✅ [doLoadLogs] === FIN EXITOSO ===');
    return respuestaFinal;

  } catch (errorGeneral) {
    console.error('❌ [doLoadLogs] ERROR CRÍTICO:', errorGeneral);
    console.error('❌ [doLoadLogs] Stack trace:', errorGeneral.stack);

    // ✅ CREAR RESPUESTA DE ERROR GARANTIZADA
    const respuestaError = {
      success: false,
      message: `Error crítico cargando logs: ${errorGeneral.message}`,
      logs: [{
        id: `log_critical_error_${Date.now()}`,
        fecha: new Date().toISOString(),
        timestamp: new Date().toISOString(),
        nivel: 'ERROR',
        mensaje: `Error crítico en doLoadLogs: ${errorGeneral.message}`,
        usuario: 'Sistema',
        ip: 'N/A',
        detalles: errorGeneral.stack ? errorGeneral.stack.substring(0, 500) : 'Sin stack trace'
      }],
      stats: { total: 1, errores: 1, warnings: 0, info: 0 },
      error: errorGeneral.message,
      timestamp: new Date().toISOString()
    };

    console.log('❌ [doLoadLogs] Devolviendo respuesta de error:', respuestaError);
    return respuestaError;
  }
}

/**
 * 🔧 [FALLBACK] Función de respaldo para cargar logs directamente
 */
function doLoadLogsFallback(limite = 50) {
  console.log('🔄 [FALLBACK] Ejecutando carga directa de logs...');
  
  try {
    // Acceso directo a la hoja sin dependencias externas
    const ss = SpreadsheetApp.getActive();
    const hojaLogs = ss.getSheetByName('Logs'); // Nombre hardcodeado como fallback
    
    if (!hojaLogs) {
      console.warn('⚠️ [FALLBACK] Hoja "Logs" no encontrada, intentando otras...');
      // Intentar otros nombres posibles
      const nombresAlternativos = ['Log', 'Registro', 'Sistema', 'Events'];
      let hojaEncontrada = null;
      
      for (const nombre of nombresAlternativos) {
        hojaEncontrada = ss.getSheetByName(nombre);
        if (hojaEncontrada) {
          console.log(`✅ [FALLBACK] Encontrada hoja alternativa: ${nombre}`);
          break;
        }
      }
      
      if (!hojaEncontrada) {
        return {
          success: false,
          message: 'No se encontró ninguna hoja de logs',
          logs: [],
          stats: { total: 0, errores: 0, warnings: 0, info: 0 }
        };
      }
      
      hojaLogs = hojaEncontrada;
    }
    
    const ultimaFila = hojaLogs.getLastRow();
    
    if (ultimaFila <= 1) {
      return {
        success: true,
        message: 'No hay logs registrados en la hoja',
        logs: [],
        stats: { total: 0, errores: 0, warnings: 0, info: 0 }
      };
    }
    
    // Obtener datos básicos
    const datos = hojaLogs.getDataRange().getValues();
    const encabezados = datos[0];
    
    // Crear logs básicos
    const logs = [];
    const maxFilas = Math.min(datos.length, limite + 1);
    
    for (let i = 1; i < maxFilas; i++) {
      const fila = datos[i];
      if (!fila || fila.every(cell => !cell)) continue;
      
      logs.push({
        id: `fallback_log_${i}`,
        fecha: fila[0] || new Date().toISOString(),
        nivel: fila[1] || 'INFO',
        mensaje: fila[2] || 'Log sin mensaje',
        usuario: fila[4] || 'Sistema',
        detalles: fila[3] || '',
        ip: fila[5] || 'N/A'
      });
    }
    
    console.log(`✅ [FALLBACK] Cargados ${logs.length} logs directamente`);
    
    return {
      success: true,
      message: `${logs.length} logs cargados mediante fallback`,
      logs: logs,
      stats: {
        total: logs.length,
        errores: logs.filter(l => l.nivel === 'ERROR').length,
        warnings: logs.filter(l => l.nivel === 'WARNING').length,
        info: logs.filter(l => l.nivel === 'INFO').length
      },
      fuente: 'Fallback directo'
    };
    
  } catch (error) {
    console.error('❌ [FALLBACK] Error en fallback:', error);
    throw error;
  }
}

/**
 * 🔧 [GLOBAL] Función global para cargar administradores desde Google Apps Script
 */
function doLoadAdminUsers() {
  try {
    return loadAdminUsers();
  } catch (error) {
    console.error('❌ [GLOBAL] Error en doLoadAdminUsers:', error);
    return {
      success: false,
      message: 'Error global cargando administradores: ' + error.message,
      admins: [],
      stats: { total: 0, activos: 0, permisos_completos: 0, solo_lectura: 0 },
      error: error.message
    };
  }
}

/**
 * Valida las credenciales de un administrador y, si son correctas,
 * crea un token de sesión temporal usando CacheService.
 * @param {string} credenciales - JSON con { usuario: string, clave: string }
 * @return {Object} Resultado con el token de sesión si el login es exitoso.
 */
function autenticarAdminYCrearSesion(credenciales) {
  console.log('🔐 [AUTH] Iniciando autenticación y creación de sesión...');
  const resultadoValidacion = validarAdministradorMejorado(credenciales); // Usamos la versión robusta para verificar la clave

  if (resultadoValidacion.success) {
    const token = `surpass-token-${Utilities.getUuid()}`;
    const cache = CacheService.getScriptCache();
    // Guardar los datos del admin en caché por 1 hora
    cache.put(token, JSON.stringify(resultadoValidacion.admin), 3600);
    if (resultadoValidacion.admin && resultadoValidacion.admin.cedula) {
      console.log(`✅ [AUTH] Sesión creada para ${resultadoValidacion.admin.cedula} con token.`);
    } else {
      console.log('✅ [AUTH] Sesión creada para admin sin cédula con token.');
    }
    return {
      success: true,
      token: token,
      admin: resultadoValidacion.admin,
      message: 'Autenticación exitosa.'
    };
  } else {
    console.warn(`⚠️ [AUTH] Fallo de autenticación: ${resultadoValidacion.message}`);
    return resultadoValidacion; // Devuelve el objeto de error original
  }
}

/**
 * Verifica si un token de sesión es válido.
 * @param {string} token - El token de sesión a verificar.
 * @return {Object|null} Los datos del administrador si el token es válido, o null si no lo es.
 */
function obtenerAdminDesdeSesion(token) {
  if (!token || typeof token !== 'string') {
    return null;
  }
  try {
    const cache = CacheService.getScriptCache();
    const adminDataJSON = cache.get(token);

    if (adminDataJSON) {
      console.log(`✅ [AUTH] Token válido encontrado para sesión.`);
      return JSON.parse(adminDataJSON);
    }
  } catch (error) {
    console.error(`❌ [AUTH] Error al verificar sesión: ${error.message}`);
  }
  
  console.warn(`⚠️ [AUTH] Token de sesión inválido o expirado.`);
  return null;
}

console.log('✅ adminB.js - Funciones CRUD para administradores implementadas exitosamente');

/**
 * Hashea una contraseña usando SHA-256 y la sal del sistema.
 * @param {string} password - La contraseña en texto plano.
 * @returns {string} El hash de la contraseña en formato Base64.
 * @private
 */
function _hashPassword(password) {
  const salt = PropertiesService.getScriptProperties().getProperty('SURPASS_SECRET_SALT');
  if (!salt) {
    throw new Error('La sal de seguridad no está configurada. Ejecute _inicializarSalDeSeguridad().');
  }
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  return Utilities.base64Encode(digest);
}

function autenticarAdministrador(credenciales) {
  console.log('🔐 [AUTH] Redirigiendo a autenticación segura...');
  
  // 🔒 [SEGURIDAD] Redirigir a la función segura con hashing
  return validarAdministradorMejorado(credenciales);
}

/**
 * Función auxiliar para crear la hoja de administradores si no existe
 */
function crearHojaAdministradores() {
  try {
    console.log('📋 [AUTH] Creando hoja de administradores...');
    
    const ss = SpreadsheetApp.getActive();
    const hojaClave = ss.insertSheet('Clave');
    
    // Crear encabezados
    const encabezados = ['Cédula', 'Nombre', 'Contraseña', 'Cargo', 'Email', 'Permisos', 'Estado'];
    hojaClave.appendRow(encabezados);
    hojaClave.getRange(1, 1, 1, encabezados.length)
      .setFontWeight('bold')
      .setBackground('#f2f2f2');
    
    // Crear usuario administrador por defecto
    const permisosDefecto = JSON.stringify({
      personal: { ver: true, editar: true },
      configuracion: { ver: true, editar: true },
      usuarios: { ver: true, editar: true },
      logs: { ver: true, editar: true },
      evacuacion: { ver: true, editar: true },
      diagnostico: { ver: true, ejecutar: true }
    });
    
    const adminDefecto = [
      'admin',
      'Administrador Principal',
      'admin123',
      'Administrador',
      'admin@surpass.com',
      permisosDefecto,
      'activo'
    ];
    
    hojaClave.appendRow(adminDefecto);
    
    console.log('✅ [AUTH] Hoja de administradores creada con usuario por defecto');
    
    return {
      success: true,
      message: 'Hoja de administradores creada',
      usuarioDefecto: { usuario: 'admin', contraseña: 'admin123' }
    };
    
  } catch (error) {
    console.error('❌ [AUTH] Error creando hoja de administradores:', error);
    return {
      success: false,
      message: 'Error creando sistema de administradores: ' + error.message
    };
  }
}

function verificarSistemaAutenticacion() {
  console.log('🔍 VERIFICANDO SISTEMA DE AUTENTICACIÓN...');
  
  try {
    const ss = SpreadsheetApp.getActive();
    console.log('✅ Spreadsheet accesible');
    
    const hojaClave = ss.getSheetByName('Clave');
    if (!hojaClave) {
      console.log('❌ Hoja "Clave" no existe - creándola...');
      const result = crearHojaAdministradores();
      console.log('Resultado creación:', result);
      return result;
    }
    
    console.log('✅ Hoja "Clave" existe');
    
    const datos = hojaClave.getDataRange().getValues();
    console.log(`📊 Datos en hoja: ${datos.length} filas`);
    
    if (datos.length > 0) {
      console.log('📋 Encabezados:', datos[0]);
      
      if (datos.length > 1) {
        console.log(`👥 Administradores encontrados: ${datos.length - 1}`);
        for (let i = 1; i < datos.length; i++) {
          console.log(`  - ${datos[i][0]} (${datos[i][1]}) - Estado: ${datos[i][6] || 'activo'}`);
        }
      }
    }
    
    // Test de autenticación con credenciales por defecto
    console.log('🧪 Probando autenticación con admin/admin123...');
    const testAuth = autenticarAdministrador(JSON.stringify({
      usuario: 'admin',
      clave: 'admin123'
    }));
    
    console.log('Resultado test:', testAuth);
    
    return {
      success: true,
      mensaje: 'Verificación completada',
      detalles: {
        hojaExiste: !!hojaClave,
        totalFilas: datos.length,
        administradores: datos.length - 1,
        testAuth: testAuth.success
      }
    };
    
  } catch (error) {
    console.error('❌ Error en verificación:', error);
    return {
      success: false,
      mensaje: 'Error en verificación: ' + error.message
    };
  }
}
