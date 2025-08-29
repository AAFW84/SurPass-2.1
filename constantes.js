// constantes.js - Definición centralizada de constantes y funciones de inicialización del sistema

// =================================================================
// SISTEMA DE CACHÉS PARA OPTIMIZACIÓN DE RENDIMIENTO
// =================================================================

/**
 * Sistema de caché en memoria para optimizar el rendimiento del sistema.
 * Reduce las consultas repetitivas a Google Sheets.
 * DECLARADO AQUÍ COMO ÚNICA FUENTE DE VERDAD.
 */
var SurPassCache = {
  // Cache de estructuras de columnas
  columnasRegistro: null,
  columnasPersonal: null,
  tiempoColumnas: null,
  
  // Cache de estadísticas
  ultimasEstadisticas: null,
  tiempoEstadisticas: null,
  
  // Cache de configuración
  configuracion: null,
  tiempoConfiguracion: null,
  
  // Configuración del cache (duraciones en milisegundos)
  get DURACION_CACHE_ESTADISTICAS() {
    const config = this.obtenerConfiguracionCache();
    return parseInt(config.CACHE_DURACION_ESTADISTICAS || '180') * 1000;
  },
  
  get DURACION_CACHE_COLUMNAS() {
    const config = this.obtenerConfiguracionCache();
    return parseInt(config.CACHE_DURACION_COLUMNAS || '1800') * 1000;
  },
  
  get DURACION_CACHE_CONFIG() {
    const config = this.obtenerConfiguracionCache();
    return parseInt(config.CACHE_DURACION_CONFIG || '900') * 1000;
  },
  
  get DURACION_CACHE_USUARIOS() {
    const config = this.obtenerConfiguracionCache();
    return parseInt(config.CACHE_DURACION_USUARIOS || '600') * 1000;
  },
  
  obtenerConfiguracionCache: function() {
    try {
      if (typeof obtenerConfiguracion === 'function') {
        const result = obtenerConfiguracion();
        if (result.success && result.config) {
          return result.config;
        }
      }
    } catch (error) {
      console.warn('⚠️ [CONFIG-INTEGRATION] Error obteniendo configuración de cache:', error);
    }
    return {};
  },
  
  obtenerColumnasRegistroCache: function() {
    const ahora = new Date().getTime();
    if (!this.columnasRegistro || !this.tiempoColumnas || (ahora - this.tiempoColumnas) > this.DURACION_CACHE_COLUMNAS) {
      console.log('🔄 Recargando cache de columnas de registro...');
      this.columnasRegistro = getColumnasRegistro();
      this.tiempoColumnas = ahora;
    }
    return this.columnasRegistro;
  },
  
  invalidarEstadisticas: function() {
    this.ultimasEstadisticas = null;
    this.tiempoEstadisticas = null;
  },
  
  limpiarTodo: function() {
    this.columnasRegistro = null;
    this.columnasPersonal = null;
    this.tiempoColumnas = null;
    this.ultimasEstadisticas = null;
    this.tiempoEstadisticas = null;
    this.configuracion = null;
    this.tiempoConfiguracion = null;
    console.log('🧹 Cache del sistema limpiado');
  }
};

// --- NOMBRES DE HOJAS ---
const SHEET_REGISTRO = 'Registro';
const SHEET_HISTORIAL = 'Historial Archivados';
const SHEET_ERRORES = 'Errores'; // Hoja para errores específicos, además de Logs
const SHEET_LOGS = 'Logs';
const SHEET_BD = 'Base de Datos'; // Hoja única para todos los datos de personal y usuarios
// SHEET_PERSONAL eliminado - ahora todo se maneja en SHEET_BD (Base de Datos)
const SHEET_CONFIG = 'Configuración';
const SHEET_REPORTES = 'Reportes';
const SHEET_EVACUACIONES = 'Evacuaciones'; // Hoja específica para registros de eventos de evacuación/simulacros
const SHEET_CLAVES = 'Clave'; // Para gestión de usuarios y contraseñas de admin (si aplica), corregido para coincidir con el nombre de la hoja

// --- VERSIÓN Y FECHA DEL SISTEMA ---
const VERSION_SISTEMA = 'V6.2';
const FECHA_ACTUALIZACION = '2 de agosto de 2025';

// --- CACHE DE USUARIOS (Backend) ---
let cacheUsuarios = {}; // Variable global para el cache de usuarios cargado desde SHEET_BD

// --- MAPEO DINÁMICO DE COLUMNAS ---
// Estas variables globales almacenarán los índices de las columnas una vez detectados.
let COLUMNAS_REGISTRO_DINAMICAS = {};
let COLUMNAS_PERSONAL_DINAMICAS = {};
let COLUMNAS_CLAVES_DINAMICO = {}; // Nuevo mapeo para hoja de claves
let COLUMNAS_LOGS_DINAMICAS = {};
let COLUMNAS_EVACUACIONES_DINAMICAS = {};

// También conservamos la versión alternativa del nuevo archivo por compatibilidad
let COLUMNAS_REGISTRO_DINAMICO = {};
let COLUMNAS_PERSONAL_DINAMICO = {};
let COLUMNAS_LOGS_DINAMICO = {};

/**
 * Detecta la estructura de columnas de una hoja específica leyendo sus encabezados.
 *
 * @param {string} nombreHoja - El nombre de la hoja a analizar.
 * @returns {Object} Un objeto donde las claves son los nombres normalizados de los encabezados
 *                   y los valores son sus índices de columna (0-indexed).
 *                   Retorna un objeto vacío si la hoja no se encuentra o hay un error.
 */
function detectarEstructuraColumnas(nombreHoja) {
  try {
    if (!nombreHoja || typeof nombreHoja !== 'string' || nombreHoja.trim() === '') {
      console.error(`❌ detectarEstructuraColumnas: Nombre de hoja inválido proporcionado: "${nombreHoja}".`);
      return {};
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
    if (!sheet) {
      console.warn(`⚠️ detectarEstructuraColumnas: Hoja "${nombreHoja}" no encontrada.`);
      return {};
    }

    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
        console.warn(`⚠️ detectarEstructuraColumnas: Hoja "${nombreHoja}" está vacía o no tiene encabezados.`);
        return {};
    }

    const encabezados = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const mapeoColumnas = {};
    
    encabezados.forEach((encabezado, indice) => {
      const encabezadoNormalizado = String(encabezado || '').trim().toLowerCase();
      if (encabezadoNormalizado) { // Asegurarse de que el encabezado no esté vacío
        mapeoColumnas[encabezadoNormalizado] = indice;
      }
    });

    console.log(`📊 Estructura detectada para "${nombreHoja}":`, mapeoColumnas);
    return mapeoColumnas;
  } catch (error) {
    console.error(`❌ Error detectando estructura de "${nombreHoja}": ${error.message}`, error.stack);
    return {};
  }
}

/**
 * Inicializa las estructuras de mapeo de columnas globales (`COLUMNAS_X_DINAMICAS`)
 * leyendo los encabezados de las hojas principales. Utiliza valores por defecto
 * si no se encuentran las columnas esperadas.
 */
function inicializarEstructurasColumnas() {
  try {
    console.log('🔍 Inicializando estructuras de columnas dinámicamente...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('No se pudo acceder a la hoja de cálculo activa para inicializar columnas.');
    }
    
    // Detectar estructura de hoja Registro
    const mapeoRegistro = detectarEstructuraColumnas(SHEET_REGISTRO);
    COLUMNAS_REGISTRO_DINAMICAS = {
      FECHA: mapeoRegistro['fecha'] ?? mapeoRegistro['date'] ?? 0,
      CEDULA: mapeoRegistro['cédula'] ?? mapeoRegistro['cedula'] ?? mapeoRegistro['id'] ?? 1,
      NOMBRE: mapeoRegistro['nombre'] ?? mapeoRegistro['name'] ?? 2,
      ESTADO_ACCESO: mapeoRegistro['estado acceso'] ?? mapeoRegistro['estado del acceso'] ?? mapeoRegistro['estado'] ?? mapeoRegistro['access'] ?? 3,
      ENTRADA: mapeoRegistro['entrada'] ?? mapeoRegistro['hora entrada'] ?? mapeoRegistro['entry'] ?? 4,
      SALIDA: mapeoRegistro['salida'] ?? mapeoRegistro['hora salida'] ?? mapeoRegistro['exit'] ?? 5,
      DURACION: mapeoRegistro['duración'] ?? mapeoRegistro['duracion'] ?? mapeoRegistro['duration'] ?? 6,
      EMPRESA: mapeoRegistro['empresa'] ?? mapeoRegistro['company'] ?? 7,
      COMENTARIO: mapeoRegistro['comentario'] ?? mapeoRegistro['comment'] ?? 8
    };

    // Detectar estructura de hoja Personal (o BD si se usa como personal)
    // Solo usa SHEET_BD (Base de Datos) para todos los datos de personal
    let hojaPersonalUsada = SHEET_BD;
    let mapeoPersonal = detectarEstructuraColumnas(SHEET_BD);
    if (Object.keys(mapeoPersonal).length === 0) {
        console.log(`❌ Hoja "${SHEET_BD}" no encontrada o vacía. Creando estructura básica...`);
        // Aquí se podría crear la hoja si no existe
    } else {
        console.log(`✅ Usando hoja "${SHEET_BD}" para mapeo de personal.`);
    }

    COLUMNAS_PERSONAL_DINAMICAS = {
      CEDULA: mapeoPersonal['cédula'] ?? mapeoPersonal['cedula'] ?? mapeoPersonal['id'] ?? 0,
      NOMBRE: mapeoPersonal['nombre'] ?? mapeoPersonal['name'] ?? 1,
      EMPRESA: mapeoPersonal['empresa'] ?? mapeoPersonal['company'] ?? 2,
      AREA: mapeoPersonal['área'] ?? mapeoPersonal['area'] ?? 3,
      CARGO: mapeoPersonal['cargo'] ?? mapeoPersonal['position'] ?? 4,
      ESTADO: mapeoPersonal['estado'] ?? mapeoPersonal['status'] ?? 5
    };

    // Detectar estructura de hoja Logs
    const mapeoLogs = detectarEstructuraColumnas(SHEET_LOGS);
    COLUMNAS_LOGS_DINAMICAS = {
      TIMESTAMP: mapeoLogs['timestamp'] ?? mapeoLogs['fecha'] ?? mapeoLogs['date'] ?? 0,
      NIVEL: mapeoLogs['nivel'] ?? mapeoLogs['level'] ?? 1,
      MENSAJE: mapeoLogs['mensaje'] ?? mapeoLogs['message'] ?? 2,
      DETALLES: mapeoLogs['detalles'] ?? mapeoLogs['details'] ?? 3,
      USUARIO: mapeoLogs['usuario'] ?? mapeoLogs['user'] ?? 4,
      IP: mapeoLogs['ip'] ?? mapeoLogs['dirección ip'] ?? mapeoLogs['ip address'] ?? 5
    };

    // Detectar estructura de hoja Evacuaciones
    const mapeoEvacuaciones = detectarEstructuraColumnas(SHEET_EVACUACIONES);
    COLUMNAS_EVACUACIONES_DINAMICAS = {
      FECHA: mapeoEvacuaciones['fecha'] ?? mapeoEvacuaciones['date'] ?? 0,
      HORA: mapeoEvacuaciones['hora'] ?? mapeoEvacuaciones['time'] ?? 1,
      TIPO: mapeoEvacuaciones['tipo'] ?? mapeoEvacuaciones['type'] ?? 2,
      PERSONAS_EVACUADAS: mapeoEvacuaciones['personas evacuadas'] ?? mapeoEvacuaciones['evacuated'] ?? 3,
      TOTAL_PERSONAS: mapeoEvacuaciones['total personas'] ?? mapeoEvacuaciones['total'] ?? 4,
      ESTADO: mapeoEvacuaciones['estado'] ?? mapeoEvacuaciones['status'] ?? 5,
      RESPONSABLE: mapeoEvacuaciones['responsable'] ?? mapeoEvacuaciones['responsible'] ?? 6
    };

    // Detectar estructura de hoja Claves (nueva funcionalidad)
    const mapeoClaves = detectarEstructuraColumnas(SHEET_CLAVES);
    COLUMNAS_CLAVES_DINAMICO = {
      CEDULA: mapeoClaves['cédula'] ?? mapeoClaves['cedula'] ?? mapeoClaves['id'] ?? 0,
      NOMBRE: mapeoClaves['nombre'] ?? mapeoClaves['name'] ?? 1,
      CONTRASENA: mapeoClaves['contraseña'] ?? mapeoClaves['password'] ?? 2,
      CARGO: mapeoClaves['cargo'] ?? mapeoClaves['position'] ?? 3,
      EMAIL: mapeoClaves['email'] ?? mapeoClaves['correo'] ?? 4,
      PERMISOS: mapeoClaves['permisos'] ?? mapeoClaves['permissions'] ?? 5,
      FECHA_CREACION: mapeoClaves['fecha_creacion'] ?? mapeoClaves['fecha creación'] ?? mapeoClaves['created'] ?? 6
    };

    // Sincronizar con variables alternativas para compatibilidad
    COLUMNAS_REGISTRO_DINAMICO = COLUMNAS_REGISTRO_DINAMICAS;
    COLUMNAS_PERSONAL_DINAMICO = COLUMNAS_PERSONAL_DINAMICAS;
    COLUMNAS_LOGS_DINAMICO = COLUMNAS_LOGS_DINAMICAS;

    console.log('✅ Estructuras de columnas detectadas y cargadas correctamente.');
    console.log('  -> REGISTRO:', JSON.stringify(COLUMNAS_REGISTRO_DINAMICAS));
    console.log('  -> PERSONAL:', JSON.stringify(COLUMNAS_PERSONAL_DINAMICAS));
    console.log('  -> CLAVES:', JSON.stringify(COLUMNAS_CLAVES_DINAMICO));
    console.log('  -> LOGS:', JSON.stringify(COLUMNAS_LOGS_DINAMICAS));
    console.log('  -> EVACUACIONES:', JSON.stringify(COLUMNAS_EVACUACIONES_DINAMICAS));

    return true;

  } catch (error) {
    console.error('❌ Error inicializando estructuras de columnas dinámicas:', error.message, error.stack);
    // En caso de error crítico, usar valores por defecto para evitar que la aplicación falle.
    usarEstructurasPorDefecto();
    return false;
  }
}

/**
 * Establece las estructuras de mapeo de columnas a valores por defecto (índices fijos).
 * Esto se usa como fallback si la detección dinámica falla.
 */
function usarEstructurasPorDefecto() {
  console.log('⚠️ Usando estructura de columnas por defecto (índices fijos) debido a un error en la detección.');
  
  COLUMNAS_REGISTRO_DINAMICAS = {
    FECHA: 0, CEDULA: 1, NOMBRE: 2, ESTADO_ACCESO: 3,
    ENTRADA: 4, SALIDA: 5, DURACION: 6, EMPRESA: 7, COMENTARIO: 8
  };
  
  COLUMNAS_PERSONAL_DINAMICAS = {
    CEDULA: 0, NOMBRE: 1, EMPRESA: 2, AREA: 3, CARGO: 4, ESTADO: 5
  };
  
  COLUMNAS_CLAVES_DINAMICO = {
    CEDULA: 0, NOMBRE: 1, CONTRASENA: 2, CARGO: 3, EMAIL: 4, PERMISOS: 5, FECHA_CREACION: 6
  };
  
  COLUMNAS_LOGS_DINAMICAS = {
    TIMESTAMP: 0, NIVEL: 1, MENSAJE: 2, DETALLES: 3, USUARIO: 4, IP: 5
  };
  
  COLUMNAS_EVACUACIONES_DINAMICAS = {
    FECHA: 0, HORA: 1, TIPO: 2, PERSONAS_EVACUADAS: 3, TOTAL_PERSONAS: 4, ESTADO: 5, RESPONSABLE: 6
  };

  // Sincronizar variables alternativas
  COLUMNAS_REGISTRO_DINAMICO = COLUMNAS_REGISTRO_DINAMICAS;
  COLUMNAS_PERSONAL_DINAMICO = COLUMNAS_PERSONAL_DINAMICAS;
  COLUMNAS_LOGS_DINAMICO = COLUMNAS_LOGS_DINAMICAS;
}

// --- GETTERS PARA LAS ESTRUCTURAS DE COLUMNAS (Siempre usar estos) ---
// Estas funciones aseguran que las estructuras estén inicializadas antes de ser usadas.

function getColumnasRegistro() {
  if (Object.keys(COLUMNAS_REGISTRO_DINAMICAS).length === 0) {
    inicializarEstructurasColumnas(); // Inicializa si aún no se han cargado
  }
  return COLUMNAS_REGISTRO_DINAMICAS;
}

function getColumnasPersonal() {
  if (Object.keys(COLUMNAS_PERSONAL_DINAMICAS).length === 0) {
    inicializarEstructurasColumnas();
  }
  return COLUMNAS_PERSONAL_DINAMICAS;
}

function getColumnasLogs() {
  if (Object.keys(COLUMNAS_LOGS_DINAMICAS).length === 0) {
    inicializarEstructurasColumnas();
  }
  return COLUMNAS_LOGS_DINAMICAS;
}

function getColumnasEvacuaciones() {
  if (Object.keys(COLUMNAS_EVACUACIONES_DINAMICAS).length === 0) {
    inicializarEstructurasColumnas();
  }
  return COLUMNAS_EVACUACIONES_DINAMICAS;
}

// --- GETTERS ADICIONALES PARA COMPATIBILIDAD CON NUEVA NOMENCLATURA ---
function getColumnasClaves() {
  if (Object.keys(COLUMNAS_CLAVES_DINAMICO).length === 0) {
    inicializarEstructurasColumnas();
  }
  return COLUMNAS_CLAVES_DINAMICO;
}

// --- CONSTANTES DE COLUMNAS ESTÁTICAS (Mantenidas por compatibilidad, preferir los getters dinámicos) ---
// ...eliminadas constantes estáticas obsoletas, usar solo getters dinámicos...

// --- ESTADOS DE ACCESO ---
const ESTADOS_ACCESO = {
    PERMITIDO: 'Acceso Permitido',
    TEMPORAL: 'Acceso Temporal',
    DENEGADO: 'Acceso Denegado'
};

// --- TIPOS DE EVACUACIÓN ---
const TIPOS_EVACUACION = {
    SIMULACRO: 'SIMULACRO', // Usar mayúsculas para consistencia con valores internos
    REAL: 'REAL'
};

// --- NIVELES DE LOGGING ---
const NIVELES_LOG = {
    INFO: 'INFO',
    ADVERTENCIA: 'ADVERTENCIA',
    WARNING: 'WARNING', // Mantener 'WARNING' para compatibilidad si hay llamadas existentes
    ERROR: 'ERROR',
    DEBUG: 'DEBUG'
};

/**
 * Verifica y crea la estructura necesaria de hojas y sus encabezados.
 * Si una hoja no existe, la crea con los encabezados predefinidos.
 * Si existe, verifica y corrige los encabezados si no coinciden.
 * También configura el formato condicional para la columna 'Estado del Acceso' en la hoja 'Registro'.
 *
 * @returns {Object} Resultado de la operación: { success: boolean, message: string, hojasCreadas: Array, hojasExistentes: Array }.
 */
function verificarYCrearEstructuraHojas() {
  console.log("🔍 Iniciando verificación de estructura de hojas...");
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("No se pudo acceder a la hoja de cálculo activa para verificar estructura.");
    }
    
    // Definición de estructura de hojas y sus encabezados
    const estructuraHojas = {
      [SHEET_REGISTRO]: getEncabezadosPredefinidos(SHEET_REGISTRO),
      [SHEET_BD]: getEncabezadosPredefinidos(SHEET_BD), // Hoja única para personal
      [SHEET_LOGS]: getEncabezadosPredefinidos(SHEET_LOGS),
      [SHEET_EVACUACIONES]: getEncabezadosPredefinidos(SHEET_EVACUACIONES),
      [SHEET_CONFIG]: getEncabezadosPredefinidos(SHEET_CONFIG), // Asegura que 'Configuración' exista
      [SHEET_ERRORES]: getEncabezadosPredefinidos(SHEET_ERRORES), // Si se usa para errores detallados
      [SHEET_HISTORIAL]: getEncabezadosPredefinidos(SHEET_HISTORIAL) // Para registros archivados
    };
    
    let hojasCreadas = [];
    let hojasExistentes = [];
    
    // Iterar sobre las hojas que deben existir y verificar/crear
    for (const nombreHoja in estructuraHojas) {
      const encabezadosEsperados = estructuraHojas[nombreHoja];
      let sheet = ss.getSheetByName(nombreHoja);
      
      if (!sheet) {
        // Crear hoja si no existe
        sheet = ss.insertSheet(nombreHoja);
        hojasCreadas.push(nombreHoja);
        
        // Añadir encabezados
        if (encabezadosEsperados && encabezadosEsperados.length > 0) {
          sheet.getRange(1, 1, 1, encabezadosEsperados.length).setValues([encabezadosEsperados]);
          // Formato básico de encabezados
          sheet.getRange(1, 1, 1, encabezadosEsperados.length)
            .setFontWeight("bold")
            .setBackground("#4285F4") // Azul Google
            .setFontColor("white");
          // Ajustar ancho de columnas para mejor visualización
          encabezadosEsperados.forEach((_, i) => sheet.setColumnWidth(i + 1, 150));
        }
        console.log(`✅ Hoja "${nombreHoja}" creada con encabezados.`);
      } else {
        hojasExistentes.push(nombreHoja);
        // Si la hoja ya existe, verificar sus encabezados
        const lastColumn = sheet.getLastColumn();
        if (lastColumn === 0 && encabezadosEsperados.length > 0) {
            // Hoja existe pero está vacía, añadir encabezados
            sheet.getRange(1, 1, 1, encabezadosEsperados.length).setValues([encabezadosEsperados]);
            sheet.getRange(1, 1, 1, encabezadosEsperados.length)
              .setFontWeight("bold")
              .setBackground("#4285F4")
              .setFontColor("white");
            console.log(`⚠️ Hoja "${nombreHoja}" estaba vacía, se le añadieron encabezados.`);
        } else if (encabezadosEsperados && encabezadosEsperados.length > 0) {
            const encabezadosActuales = sheet.getRange(1, 1, 1, Math.min(lastColumn, encabezadosEsperados.length)).getValues()[0];
            let needsUpdate = false;
            for (let i = 0; i < encabezadosEsperados.length; i++) {
                if (encabezadosActuales[i] !== encabezadosEsperados[i]) {
                    needsUpdate = true;
                    break;
                }
            }
            if (needsUpdate || lastColumn < encabezadosEsperados.length) {
                // Actualizar o expandir encabezados si no coinciden o faltan
                sheet.getRange(1, 1, 1, encabezadosEsperados.length).setValues([encabezadosEsperados]);
                sheet.getRange(1, 1, 1, encabezadosEsperados.length)
                    .setFontWeight("bold")
                    .setBackground("#4285F4")
                    .setFontColor("white");
                console.log(`🔄 Hoja "${nombreHoja}": Encabezados actualizados/corregidos.`);
            } else {
                console.log(`✅ Hoja "${nombreHoja}" verificada, estructura de encabezados correcta.`);
            }
        }
      }
    }
    
    // Configurar formato condicional para la hoja de Registro
    const formatoResult = configurarFormatoCondicionalEstadoAcceso();
    if (!formatoResult.success) {
        console.warn(`⚠️ Fallo al configurar formato condicional: ${formatoResult.message}`);
    }

    // Asegurarse de que la hoja de Base de Datos tenga algunos datos de prueba si está vacía
    const BD_sheet = ss.getSheetByName(SHEET_BD);
    if (BD_sheet && BD_sheet.getLastRow() <= 1) { // Solo encabezados o vacía
        crearDatosPrueba(); // Llama a la función de abajo
    }
    
    return {
      success: true,
      message: "Estructura de hojas verificada y configurada correctamente.",
      hojasCreadas: hojasCreadas,
      hojasExistentes: hojasExistentes
    };
  } catch (error) {
    console.error(`❌ Error crítico en verificarYCrearEstructuraHojas: ${error.message}`, error.stack);
    if (typeof registrarLog === 'function') {
        registrarLog('ERROR', 'Error crítico al verificar/crear estructura de hojas', { error: error.message, stack: error.stack });
    }
    return {
      success: false,
      message: `Error verificando estructura de hojas: ${error.message}`
    };
  }
}

/**
 * Función auxiliar para obtener los encabezados predefinidos para cada hoja.
 * Esto ayuda a mantener DRY (Don't Repeat Yourself) en la definición de encabezados.
 *
 * @param {string} sheetName - El nombre de la hoja para la que se necesitan los encabezados.
 * @returns {Array<string>} Un array de cadenas con los nombres de los encabezados.
 */
function getEncabezadosPredefinidos(sheetName) {
  switch (sheetName) {
    case SHEET_REGISTRO:
      return ["Fecha", "Cédula", "Nombre", "Estado del Acceso", "Entrada", "Salida", "Duración", "Empresa", "Comentario"];
    case SHEET_BD: // Hoja única para todos los datos de personal
      return ["Cédula", "Nombre", "Empresa", "Área", "Cargo", "Estado"];
    case SHEET_LOGS:
      return ["Timestamp", "Nivel", "Mensaje", "Detalles", "Usuario", "IP"];
    case SHEET_EVACUACIONES:
      return ["Fecha", "Hora", "Tipo", "Personas Evacuadas", "Total Personas", "Estado", "Responsable"];
    case SHEET_CONFIG:
        return ["Clave", "Valor", "Categoría"];
    case SHEET_ERRORES:
        return ["Timestamp", "Código Error", "Mensaje Error", "Componente", "Detalles"];
    case SHEET_HISTORIAL: // Asume la misma estructura que Registro para archivado
        return ["Fecha", "Cédula", "Nombre", "Estado del Acceso", "Entrada", "Salida", "Duración", "Empresa", "Comentario"];
    case SHEET_CLAVES:
        return ["Cédula", "Nombre", "Contraseña", "Cargo", "Email", "Permisos", "Fecha_Creacion"];
    default:
      console.warn(`⚠️ No se encontraron encabezados predefinidos para la hoja: "${sheetName}".`);
      return [];
  }
}

/**
 * Configura automáticamente el formato condicional para la columna 'Estado del Acceso'
 * en la hoja de 'Registro'. Colorea las celdas según los valores de estado predefinidos.
 * Esta es la implementación ÚNICA para esta funcionalidad.
 *
 * @returns {Object} Resultado de la operación: { success: boolean, message: string }.
 */
function configurarFormatoCondicionalEstadoAcceso() {
  console.log("🎨 Configurando formato condicional para 'Estado del Acceso' en hoja Registro...");
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_REGISTRO);
    
    if (!sheet) {
      const msg = `Hoja de Registro "${SHEET_REGISTRO}" no encontrada para formato condicional.`;
      console.error(`❌ ${msg}`);
      return { success: false, message: msg };
    }
    
    // Identificar la columna de Estado del Acceso usando el mapeo dinámico
    const COLS = getColumnasRegistro();
    const columnaEstadoAccesoIndex = COLS.ESTADO_ACCESO; // 0-indexed
    if (columnaEstadoAccesoIndex === undefined || columnaEstadoAccesoIndex < 0) {
        const msg = `Columna "Estado del Acceso" no encontrada en hoja "${SHEET_REGISTRO}". No se puede aplicar formato condicional.`;
        console.warn(`⚠️ ${msg}`);
        return { success: false, message: msg };
    }
    const columnaEstadoAccesoSheetIndex = columnaEstadoAccesoIndex + 1; // 1-indexed para SpreadsheetApp

    // Obtener el rango de toda la columna (excluyendo encabezado)
    // Se extiende hasta la fila 1000 por si hay pocas filas, para reglas futuras.
    const ultimaFilaConDatos = sheet.getLastRow();
    const rangoAplicacion = sheet.getRange(2, columnaEstadoAccesoSheetIndex, Math.max(1000, ultimaFilaConDatos - 1), 1);
    
    // Obtener y filtrar reglas existentes para esta columna
    const reglasExistentes = sheet.getConditionalFormatRules();
    const nuevasReglas = reglasExistentes.filter(regla => {
      const rangosDeRegla = regla.getRanges();
      // Si la regla afecta la columna de Estado del Acceso, la descartamos para añadir una nueva.
      return !rangosDeRegla.some(rango => rango.getColumn() === columnaEstadoAccesoSheetIndex);
    });
    
    // Regla 1: Acceso Permitido (Verde)
    nuevasReglas.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(ESTADOS_ACCESO.PERMITIDO)
      .setBackground("#E6FFE6") // Verde claro
      .setFontColor("#006400") // Verde oscuro
      .setRanges([rangoAplicacion])
      .build());
    
    // Regla 2: Acceso Temporal (Naranja)
    nuevasReglas.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(ESTADOS_ACCESO.TEMPORAL)
      .setBackground("#FFF2E6") // Naranja claro
      .setFontColor("#E65C00") // Naranja oscuro
      .setRanges([rangoAplicacion])
      .build());
    
    // Regla 3: Acceso Denegado (Rojo)
    nuevasReglas.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(ESTADOS_ACCESO.DENEGADO)
      .setBackground("#FFE6E6") // Rojo claro
      .setFontColor("#CC0000") // Rojo oscuro
      .setRanges([rangoAplicacion])
      .build());
    
    // Aplicar todas las reglas (existentes que no afectaban y las nuevas)
    sheet.setConditionalFormatRules(nuevasReglas);
    
    console.log(`✅ Formato condicional configurado correctamente para la columna ${sheet.getRange(1, columnaEstadoAccesoSheetIndex).getValue()}.`);
    return { success: true, message: "Formato condicional configurado correctamente." };
  } catch (error) {
    console.error(`❌ Error al configurar formato condicional: ${error.message}`, error.stack);
    if (typeof registrarLog === 'function') {
        registrarLog('ERROR', 'Error al configurar formato condicional', { error: error.message, sheet: SHEET_REGISTRO });
    }
    return { success: false, message: `Error al configurar formato condicional: ${error.message}` };
  }
}

/**
 * Función para diagnosticar la estructura actual de las hojas.
 * Útil para debugging cuando hay problemas de mapeo de columnas.
 *
 * @returns {Object} Un objeto con el diagnóstico de cada hoja y recomendaciones.
 */
function diagnosticarEstructuraHojas() {
  console.log("🔍 Iniciando diagnóstico completo de estructura de hojas...");
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("No se pudo acceder a la hoja de cálculo activa para el diagnóstico.");
    }
    
    const hojasAAnalizar = [
      SHEET_REGISTRO, SHEET_BD, SHEET_LOGS, 
      SHEET_EVACUACIONES, SHEET_CONFIG, SHEET_ERRORES, SHEET_HISTORIAL, SHEET_CLAVES
    ];
    const diagnostico = {};
    let recomendaciones = [];
    
    hojasAAnalizar.forEach(nombreHoja => {
      const sheet = ss.getSheetByName(nombreHoja);
      const encabezadosEsperados = getEncabezadosPredefinidos(nombreHoja);
      
      if (!sheet) {
        diagnostico[nombreHoja] = {
          existe: false,
          mensaje: "⚠️ HOJA NO ENCONTRADA",
          recomendacion: `Crear hoja "${nombreHoja}" con los encabezados esperados.`
        };
        recomendaciones.push(`Falta la hoja "${nombreHoja}".`);
        return;
      }
      
      const numColumnasActual = sheet.getLastColumn();
      const encabezadosActuales = sheet.getRange(1, 1, 1, Math.max(1, numColumnasActual)).getValues()[0];
      
      diagnostico[nombreHoja] = {
        existe: true,
        columnasDetectadas: numColumnasActual,
        encabezadosActuales: encabezadosActuales,
        encabezadosEsperados: encabezadosEsperados
      };

      // Comparar encabezados
      let problemasEncabezados = false;
      if (encabezadosEsperados.length > 0) {
        if (numColumnasActual < encabezadosEsperados.length) {
            problemasEncabezados = true;
            diagnostico[nombreHoja].mensaje = `❌ Faltan columnas. Esperadas: ${encabezadosEsperados.length}, Actuales: ${numColumnasActual}.`;
            recomendaciones.push(`La hoja "${nombreHoja}" tiene menos columnas de las esperadas. Se recomienda ejecutar 'inicializarConfiguracionSistema()'.`);
        } else {
            for (let i = 0; i < encabezadosEsperados.length; i++) {
                if (encabezadosActuales[i] !== encabezadosEsperados[i]) {
                    problemasEncabezados = true;
                    diagnostico[nombreHoja].mensaje = `❌ Encabezados no coinciden en la posición ${i} (Esperado: "${encabezadosEsperados[i]}", Actual: "${encabezadosActuales[i]}").`;
                    recomendaciones.push(`La hoja "${nombreHoja}" tiene encabezados incorrectos. Se recomienda ejecutar 'inicializarConfiguracionSistema()'.`);
                    break;
                }
            }
        }
      }

      if (!problemasEncabezados) {
        diagnostico[nombreHoja].mensaje = "✅ Estructura de encabezados correcta.";
      }
    });
    
    console.log("📊 Diagnóstico completo de estructura:", JSON.stringify(diagnostico, null, 2));
    if (recomendaciones.length > 0) {
      console.warn("⚠️ RECOMENDACIONES DETECTADAS:", recomendaciones);
    } else {
      console.log("✅ No se encontraron problemas en la estructura de hojas.");
    }
    
    return {
      success: true,
      diagnostico: diagnostico,
      recomendaciones: recomendaciones
    };
  } catch (error) {
    console.error(`❌ Error en diagnosticarEstructuraHojas: ${error.message}`, error.stack);
    if (typeof registrarLog === 'function') {
        registrarLog('ERROR', 'Error al diagnosticar estructura de hojas', { error: error.message, stack: error.stack });
    }
    return {
      success: false,
      message: `Error al diagnosticar estructura de hojas: ${error.message}`
    };
  }
}

/**
 * Inicializa todas las configuraciones y estructuras necesarias para SurPass.
 * Esta función es el punto de entrada para asegurar que el entorno de Google Sheet
 * esté configurado correctamente para la aplicación.
 *
 * @returns {Object} Resultado de la inicialización.
 */
function inicializarConfiguracionSistema() {
    console.log(`🚀 Iniciando configuración del sistema SurPass ${VERSION_SISTEMA}...`);
    try {
        // Paso 1: Verificar y crear estructura de hojas y sus encabezados.
        // Esto también configura el formato condicional para la hoja de Registro.
        console.log('🔄 PASO 1: Verificando y creando estructura de hojas...');
        const resultadoEstructura = verificarYCrearEstructuraHojas();
        if (!resultadoEstructura.success) {
            throw new Error(`Fallo en la verificación de estructura: ${resultadoEstructura.message}`);
        }
        
        // Paso 2: Inicializar las estructuras de columnas dinámicas
        // Esto se hace después de asegurar que las hojas y encabezados existen.
        console.log('🔄 PASO 2: Inicializando mapeo de columnas dinámicas...');
        const resultadoColumnas = inicializarEstructurasColumnas();
        if (!resultadoColumnas) { // inicializarEstructurasColumnas devuelve true/false
             throw new Error('Fallo al inicializar las estructuras de columnas dinámicas.');
        }
        
        // Paso 3: Inicializar el caché de usuarios desde la base de datos
        // La función inicializarCacheUsuarios debe estar en registroB.js y ser global.
        console.log('🔄 PASO 3: Inicializando caché de usuarios...');
        let resultadoCache = { success: true, mensaje: 'Cache de usuarios no inicializado (función no encontrada).' };
        if (typeof inicializarCacheUsuarios === 'function') {
            try {
                inicializarCacheUsuarios();
                resultadoCache = { success: true, mensaje: 'Cache de usuarios inicializado correctamente.' };
            } catch (errorCache) {
                console.error(`⚠️ Error inicializando cache de usuarios: ${errorCache.message}`);
                resultadoCache = { success: false, mensaje: `Error: ${errorCache.message}` };
            }
        } else {
             console.warn('⚠️ La función `inicializarCacheUsuarios` no está definida o no es global.');
        }
        
        console.log(`✅ Configuración completa del sistema ${VERSION_SISTEMA} finalizada.`);
        return { 
            success: true,
            mensaje: `Sistema SurPass ${VERSION_SISTEMA} inicializado correctamente.`,
            detalles: {
                estructuraHojas: resultadoEstructura,
                columnasDinamicas: resultadoColumnas,
                cacheUsuarios: resultadoCache
            }
        };
        
    } catch (error) {
        console.error(`❌ Error fatal en inicializarConfiguracionSistema: ${error.message}`, error.stack);
        if (typeof registrarLog === 'function') {
            registrarLog('ERROR', 'Error fatal en inicialización del sistema', { error: error.message, stack: error.stack });
        }
        return {
            success: false,
            mensaje: `Error fatal al inicializar el sistema: ${error.message}`
        };
    }
}

/**
// ...eliminada función de datos de prueba, no necesaria en producción...
/**
 * Centraliza la validación de duplicados en un array de objetos por campo clave
 * @param {Array<Object>} lista - Lista de objetos a validar
 * @param {string} campo - Campo a verificar duplicados (ej: 'cedula')
 * @returns {Array} Array con los valores duplicados encontrados
 */
function obtenerDuplicadosPorCampo(lista, campo) {
  const valores = lista.map(obj => obj[campo]).filter(Boolean);
  const set = new Set();
  const duplicados = new Set();
  valores.forEach(val => {
    if (set.has(val)) {
      duplicados.add(val);
    } else {
      set.add(val);
    }
  });
  return Array.from(duplicados);
}

/**
 * Normaliza una cédula eliminando espacios y caracteres no alfanuméricos,
 * y la convierte a minúsculas. Útil para búsquedas y comparaciones.
 *
 * @param {string} cedula - La cédula a normalizar.
 * @returns {string} La cédula normalizada.
 */
function normalizarCedula(cedula) {
    if (!cedula) return '';
    return String(cedula).trim().toLowerCase().replace(/[^\w\-]/g, '');
}

/**
 * Mapea los encabezados de una hoja a índices de columna para un acceso más fácil.
 *
 * @param {Array<string>} headers - Un array de cadenas con los nombres de los encabezados de la hoja.
 * @returns {Object} Un objeto con propiedades como 'cedula', 'nombre', 'empresa', etc.,
 *                   con sus respectivos índices de columna (0-indexed).
 */
function mapearIndicesColumnas(headers) {
    const indices = {
        cedula: -1,
        nombre: -1,
        empresa: -1,
        estado: -1,
        entrada: -1,
        salida: -1,
        fecha: -1,
        comentario: -1,
        duracion: -1
    };
    
    headers.forEach((header, index) => {
        const headerNorm = String(header).toLowerCase().trim();
        if (headerNorm.includes('cédula') || headerNorm.includes('cedula') || headerNorm === 'id') indices.cedula = index;
        if (headerNorm.includes('nombre') || headerNorm === 'name') indices.nombre = index;
        if (headerNorm.includes('empresa') || headerNorm === 'company') indices.empresa = index;
        if (headerNorm.includes('estado') || headerNorm === 'status') indices.estado = index;
        if (headerNorm.includes('entrada') || headerNorm === 'entry' || headerNorm.includes('hora entrada')) indices.entrada = index;
        if (headerNorm.includes('salida') || headerNorm === 'exit' || headerNorm.includes('hora salida')) indices.salida = index;
        if (headerNorm.includes('fecha') || headerNorm === 'date') indices.fecha = index;
        if (headerNorm.includes('comentario') || headerNorm === 'comment') indices.comentario = index;
        if (headerNorm.includes('duración') || headerNorm.includes('duracion') || headerNorm === 'duration') indices.duracion = index;
    });
    
    return indices;
}

/**
 * Función auxiliar para obtener o crear una hoja
 * @param {string} nombreHoja - Nombre de la hoja
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} La hoja
 */
function obtenerOCrearHoja(nombreHoja) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(nombreHoja);
    
    if (!hoja) {
      console.log(`📝 Creando hoja: ${nombreHoja}`);
      hoja = ss.insertSheet(nombreHoja);
      
      // Añadir encabezados básicos según el tipo de hoja
      if (nombreHoja.includes('Log') || nombreHoja.includes('Stats')) {
        hoja.appendRow(['Timestamp', 'Tipo', 'Datos', 'Usuario']);
      }
    }
    
    return hoja;
    
  } catch (error) {
    console.error(`Error obteniendo/creando hoja ${nombreHoja}:`, error);
    throw error;
  }
}

/**
 * Verifica si la base de datos está en modo unificado
 * @returns {boolean} True si está unificada
 */
function verificarSiDBEstaUnificada() {
  try {
    console.log('🔍 [ADMIN] Verificando estado de la base de datos...');
    
    const ss = SpreadsheetApp.getActive();
    const hojaPersonal = ss.getSheetByName(SHEET_BD);
    const hojaBD = ss.getSheetByName(SHEET_BD);
    
    // Considerar unificada si ambas hojas existen y tienen estructura similar
    if (hojaPersonal && hojaBD) {
      const datosPersonal = hojaPersonal.getDataRange().getValues();
      const datosBD = hojaBD.getDataRange().getValues();
      
      // Verificar si tienen estructura similar (mismo número de columnas)
      const unificada = datosPersonal.length > 1 && datosBD.length > 1 && 
                       datosPersonal[0].length === datosBD[0].length;
      
      console.log(`📊 [ADMIN] Base de datos ${unificada ? 'UNIFICADA' : 'TRADICIONAL'}`);
      return unificada;
    }
    
    // Si solo existe una, considerar como funcional pero no unificada
    const existeUna = !!(hojaPersonal || hojaBD);
    console.log(`📊 [ADMIN] Base de datos ${existeUna ? 'FUNCIONAL' : 'NO CONFIGURADA'}`);
    return existeUna;
    
  } catch (error) {
    console.error('❌ [ADMIN] Error verificando estado de DB:', error);
    return false;
  }
}

/**
 * Ejecuta diagnóstico completo del sistema para admin
 * @param {string} nivel - Nivel de diagnóstico
 * @returns {Object} Resultado del diagnóstico
 */
function ejecutarDiagnosticoUnificado(nivel = 'completo') {
  try {
    console.log(`🔧 [ADMIN] Ejecutando diagnóstico ${nivel}...`);
    
    const inicioTiempo = Date.now();
    const detalles = {};
    let testsPasados = 0;
    let totalTests = 0;
    
    // Test 1: Conectividad básica
    totalTests++;
    const inicioTest1 = Date.now();
    try {
      const ss = SpreadsheetApp.getActive();
      if (ss) {
        testsPasados++;
        detalles.conectividad = {
          status: true,
          mensaje: 'Conexión con Google Sheets OK',
          tiempo: Date.now() - inicioTest1
        };
      }
    } catch (error) {
      detalles.conectividad = {
        status: false,
        mensaje: 'Error de conexión: ' + error.message,
        tiempo: Date.now() - inicioTest1
      };
    }
    
    // Test 2: Estructura de hojas
    totalTests++;
    const inicioTest2 = Date.now();
    try {
      const ss = SpreadsheetApp.getActive();
      const hojasNecesarias = [SHEET_REGISTRO, SHEET_BD, SHEET_LOGS];
      const hojasEncontradas = [];
      
      hojasNecesarias.forEach(nombreHoja => {
        if (ss.getSheetByName(nombreHoja)) {
          hojasEncontradas.push(nombreHoja);
        }
      });
      
      if (hojasEncontradas.length >= 2) {
        testsPasados++;
        detalles.estructura = {
          status: true,
          mensaje: `${hojasEncontradas.length}/${hojasNecesarias.length} hojas encontradas`,
          tiempo: Date.now() - inicioTest2
        };
      } else {
        detalles.estructura = {
          status: false,
          mensaje: `Solo ${hojasEncontradas.length}/${hojasNecesarias.length} hojas encontradas`,
          tiempo: Date.now() - inicioTest2
        };
      }
    } catch (error) {
      detalles.estructura = {
        status: false,
        mensaje: 'Error verificando estructura: ' + error.message,
        tiempo: Date.now() - inicioTest2
      };
    }
    
    // Test 3: Datos de personal (solo si existe la función)
    if (typeof obtenerTodoElPersonal === 'function') {
      totalTests++;
      const inicioTest3 = Date.now();
      try {
        const personal = obtenerTodoElPersonal();
        if (personal.length > 0) {
          testsPasados++;
          detalles.personal = {
            status: true,
            mensaje: `${personal.length} registros de personal encontrados`,
            tiempo: Date.now() - inicioTest3
          };
        } else {
          detalles.personal = {
            status: false,
            mensaje: 'No se encontraron registros de personal',
            tiempo: Date.now() - inicioTest3
          };
        }
      } catch (error) {
        detalles.personal = {
          status: false,
          mensaje: 'Error verificando personal: ' + error.message,
          tiempo: Date.now() - inicioTest3
        };
      }
    }
    
    // Test 4: Sistema de logs (solo en diagnóstico completo)
    if (nivel === 'completo' && typeof registrarLog === 'function') {
      totalTests++;
      const inicioTest4 = Date.now();
      try {
        registrarLog('INFO', 'Test de diagnóstico ejecutado desde admin', { nivel: nivel });
        testsPasados++;
        detalles.logs = {
          status: true,
          mensaje: 'Sistema de logs operativo',
          tiempo: Date.now() - inicioTest4
        };
      } catch (error) {
        detalles.logs = {
          status: false,
          mensaje: 'Error en sistema de logs: ' + error.message,
          tiempo: Date.now() - inicioTest4
        };
      }
    }
    
    const tiempoTotal = Date.now() - inicioTiempo;
    const exito = testsPasados === totalTests;
    
    console.log(`${exito ? '✅' : '❌'} [ADMIN] Diagnóstico completado: ${testsPasados}/${totalTests} tests pasados en ${tiempoTotal}ms`);
    
    return {
      success: exito,
      mensaje: exito ? 'Sistema operativo correctamente' : 'Se encontraron problemas en el sistema',
      testsPasados: testsPasados,
      totalTests: totalTests,
      tiempoTotal: tiempoTotal,
      detalles: detalles
    };
    
  } catch (error) {
    console.error('❌ [ADMIN] Error en diagnóstico:', error);
    return {
      success: false,
      mensaje: 'Error ejecutando diagnóstico: ' + error.message,
      testsPasados: 0,
      totalTests: 1,
      tiempoTotal: 0,
      detalles: {}
    };
  }
}

/**
 * Función de inicialización completa del sistema SurPass V6.2
 * Asegura que todas las hojas y estructuras estén configuradas correctamente.
 * Esta es la función principal de inicialización que debe ser llamada al inicio.
 * 
 * NOTA: Esta función combina la funcionalidad de inicializarConfiguracionSistema
 * con la creación automática de usuarios administradores por defecto.
 * 
 * @return {Object} Resultado de la inicialización
 */
function inicializarSistemaCompleto() {
  try {
    console.log(`🚀 [SISTEMA] Iniciando inicialización completa de SurPass ${VERSION_SISTEMA}...`);
    
    const ss = SpreadsheetApp.getActive();
    if (!ss) {
      throw new Error('No se pudo acceder a la hoja de cálculo');
    }
    
    const resultados = {
      hojas: {},
      administradores: {},
      personal: {},
      configuracion: {}
    };
    
    // 1. Asegurar que la hoja "Clave" existe con estructura correcta
    let sheetClave = ss.getSheetByName('Clave');
    if (!sheetClave) {
      console.log('📝 [SISTEMA] Creando hoja "Clave"...');
      sheetClave = ss.insertSheet('Clave');
      sheetClave.appendRow(['Cédula', 'Nombre', 'Contraseña', 'Cargo', 'Email', 'Permisos', 'Fecha_Creacion']);
      sheetClave.getRange(1, 1, 1, 7).setBackground('#f2f2f2').setFontWeight('bold');
      
      // Crear usuario administrador por defecto
      const adminDefecto = [
        'admin',
        'Administrador Principal', 
        'admin123',
        'Administrador',
        'admin@surpass.com',
        JSON.stringify({
          personal: { ver: true, editar: true },
          configuracion: { ver: true, editar: true },
          usuarios: { ver: true, editar: true },
          logs: { ver: true, editar: true },
          evacuacion: { ver: true, editar: true },
          diagnostico: { ver: true, ejecutar: true }
        }),
        new Date()
      ];
      sheetClave.appendRow(adminDefecto);
      
      resultados.administradores = {
        creada: true,
        usuarioDefecto: 'admin / admin123',
        mensaje: 'Hoja creada con usuario administrador por defecto'
      };
    } else {
      // Verificar estructura existente
      const encabezados = sheetClave.getRange(1, 1, 1, sheetClave.getLastColumn()).getValues()[0];
      const estructuraCorrecta = encabezados.length >= 6 && 
                                encabezados.includes('Cédula') && 
                                encabezados.includes('Contraseña') && 
                                encabezados.includes('Permisos');
      
      if (!estructuraCorrecta) {
        console.log('🔄 [SISTEMA] Actualizando estructura de hoja "Clave"...');
        sheetClave.getRange(1, 1, 1, 7).setValues([['Cédula', 'Nombre', 'Contraseña', 'Cargo', 'Email', 'Permisos', 'Fecha_Creacion']]);
        sheetClave.getRange(1, 1, 1, 7).setBackground('#f2f2f2').setFontWeight('bold');
      }
      
      resultados.administradores = {
        creada: false,
        existente: true,
        registros: sheetClave.getLastRow() - 1,
        mensaje: 'Hoja existente verificada'
      };
    }
    
    // 2. Ejecutar inicialización de constantes si existe
    if (typeof inicializarConfiguracionSistema === 'function') {
      console.log('🔧 [SISTEMA] Ejecutando inicialización de configuración...');
      const configResult = inicializarConfiguracionSistema();
      resultados.configuracion = configResult;
    }
    
    // 3. Inicializar cache de usuarios
    if (typeof inicializarCacheUsuarios === 'function') {
      console.log('💾 [SISTEMA] Inicializando cache de usuarios...');
      inicializarCacheUsuarios();
      resultados.personal = {
        cache: 'inicializado',
        registros: Object.keys(cacheUsuarios || {}).length
      };
    }
    
    console.log(`✅ [SISTEMA] Inicialización completa exitosa - SurPass ${VERSION_SISTEMA}`);
    
    return {
      success: true,
      message: `Sistema SurPass ${VERSION_SISTEMA} inicializado correctamente`,
      resultados: resultados,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ [SISTEMA] Error en inicialización:', error);
    return {
      success: false,
      message: 'Error en inicialización: ' + error.message,
      error: error.stack,
      timestamp: new Date().toISOString()
    };
  }
}

// Mensaje de depuración al cargar el script
console.log(`✅ SurPass ${VERSION_SISTEMA} - Constantes y funciones de inicialización del sistema cargadas correctamente.`);
