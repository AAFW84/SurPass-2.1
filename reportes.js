// Incluir utilidades compartidas
// <!--? include('utils.js'); ?-->
// reportes.js - Funciones de backend para reportes

/**
 * Genera reportes estáticos (función de backend)
 * @param {string} tipo - Tipo de reporte a generar
 * @param {Object} parametros - Parámetros del reporte
 * @return {Object} Resultado de la operación
 */
function generarReporte(tipo, parametros) {
  try {
    console.log(`📊 Generando reporte tipo: ${tipo}`);
    
    const hojaReportes = SpreadsheetApp.getActive().getSheetByName(SHEET_REPORTES);
    const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
    
    if (!hojaRegistro) {
      return {exito: false, mensaje: 'Hoja de registro no encontrada'};
    }
    
    const datos = hojaRegistro.getDataRange().getValues();
    
    // Ejemplo: reporte de entradas del día
    if (tipo === 'entradasDia') {
      const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const entradas = datos.filter((fila, idx) => {
        return idx > 0 && fila[2] && 
               Utilities.formatDate(new Date(fila[2]), Session.getScriptTimeZone(), 'yyyy-MM-dd') === hoy;
      });
      
      if (hojaReportes) {
        hojaReportes.clearContents();
        hojaReportes.appendRow(['Cédula', 'Nombre', 'Hora Entrada', 'Hora Salida', 'Usuario', 'Comentario']);
        entradas.forEach(fila => hojaReportes.appendRow(fila));
      }
      
      return {exito: true, mensaje: 'Reporte generado', registros: entradas.length};
    }
    
    return {exito: false, mensaje: 'Tipo de reporte no soportado'};
    
  } catch (error) {
    console.error('Error generando reporte:', error);
    return {exito: false, mensaje: 'Error interno generando reporte: ' + error.message};
  }
}

/**
 * Función para calcular estadísticas de evacuación
 * @param {Object} simulacro - Datos del simulacro
 * @return {Object} Estadísticas calculadas
 */
function calcularEstadisticasEvacuacion(simulacro) {
  try {
    const {
      personasOriginales = [],
      personasEvacuadas = [],
      personasRestantes = []
    } = simulacro;
    
    const total = personasOriginales.length;
    const evacuadas = personasEvacuadas.length;
    const restantes = personasRestantes.length;
    
    const porcentajeEvacuado = total > 0 ? (evacuadas / total) * 100 : 0;
    const porcentajeRestante = total > 0 ? (restantes / total) * 100 : 0;
    
    // Análisis por empresa
    const empresas = {};
    personasOriginales.forEach(persona => {
      const empresa = persona.empresa || 'Sin empresa';
      if (!empresas[empresa]) {
        empresas[empresa] = { total: 0, evacuadas: 0, restantes: 0 };
      }
      empresas[empresa].total++;
    });
    
    personasEvacuadas.forEach(cedula => {
      const persona = personasOriginales.find(p => p.cedula === cedula);
      if (persona) {
        const empresa = persona.empresa || 'Sin empresa';
        empresas[empresa].evacuadas++;
      }
    });
    
    personasRestantes.forEach(persona => {
      const empresa = persona.empresa || 'Sin empresa';
      empresas[empresa].restantes++;
    });
    
    return {
      general: {
        total,
        evacuadas,
        restantes,
        porcentajeEvacuado: Math.round(porcentajeEvacuado * 100) / 100,
        porcentajeRestante: Math.round(porcentajeRestante * 100) / 100,
        completado: restantes === 0
      },
      porEmpresa: empresas,
      resumen: {
        empresaConMasPersonas: Object.keys(empresas).reduce((a, b) => 
          empresas[a].total > empresas[b].total ? a : b, 'N/A'),
        empresaConMejorEvacuacion: Object.keys(empresas).reduce((a, b) => 
          (empresas[a].evacuadas / empresas[a].total) > (empresas[b].evacuadas / empresas[b].total) ? a : b, 'N/A')
      }
    };
    
  } catch (error) {
    console.error('Error calculando estadísticas:', error);
    return {
      general: { total: 0, evacuadas: 0, restantes: 0, porcentajeEvacuado: 0, porcentajeRestante: 0, completado: false },
      porEmpresa: {},
      resumen: { empresaConMasPersonas: 'N/A', empresaConMejorEvacuacion: 'N/A' }
    };
  }
}

console.log('✅ Funciones de reportes y estadísticas de evacuación cargadas correctamente');

/**
 * Obtiene el historial de registros por cédula o general.
 * @param {string} cedula - Cédula a buscar (opcional, si vacío trae todos)
 * @return {Array} Historial de registros formateados
 */
function getHistorialPorCedula(cedula) {
  try {
    const cedulaLimpia = cedula ? cedula.trim() : '';
    const esBusquedaGeneral = !cedulaLimpia;
    const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
    if (!hojaRegistro) {
      return [];
    }
    const datos = hojaRegistro.getDataRange().getValues();
    if (datos.length <= 1) {
      return [];
    }
    const datosHistorial = [];
    const COLS = getColumnasRegistro();
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const cedulaRegistro = String(fila[COLS.CEDULA] || '').trim();
      const coincide = esBusquedaGeneral || (cedulaRegistro === cedulaLimpia);
      if (coincide && cedulaRegistro) {
        let fechaRegistro = fila[COLS.FECHA];
        let fechaFormateada = '';
        if (fechaRegistro instanceof Date) {
          fechaFormateada = fechaRegistro.toISOString().split('T')[0];
        } else if (typeof fechaRegistro === 'string') {
          if (fechaRegistro.match(/^\d{4}-\d{2}-\d{2}$/)) {
            fechaFormateada = fechaRegistro;
          } else {
            try {
              const fecha = new Date(fechaRegistro);
              if (!isNaN(fecha.getTime())) {
                fechaFormateada = fecha.toISOString().split('T')[0];
              } else {
                fechaFormateada = fechaRegistro;
              }
            } catch (e) {
              fechaFormateada = fechaRegistro;
            }
          }
        } else {
          fechaFormateada = new Date().toISOString().split('T')[0];
        }
        const cedulaValue = String(fila[COLS.CEDULA] || 'N/A');
        const nombreValue = String(fila[COLS.NOMBRE] || 'N/A');
        const empresaValue = String(fila[COLS.EMPRESA] || 'N/A');
        const estadoValue = String(fila[COLS.ESTADO_ACCESO] || 'N/A');
        let entradaFormateada = 'N/A';
        if (fila[COLS.ENTRADA]) {
          entradaFormateada = (typeof formatearHora === 'function')
            ? formatearHora(fila[COLS.ENTRADA])
            : fila[COLS.ENTRADA];
        }
        let salidaFormateada = 'N/A';
        if (fila[COLS.SALIDA]) {
          salidaFormateada = (typeof formatearHora === 'function')
            ? formatearHora(fila[COLS.SALIDA])
            : fila[COLS.SALIDA];
        }
        let duracionFormateada = 'N/A';
        if (fila[COLS.DURACION]) {
          const duracion = fila[COLS.DURACION];
          if (typeof duracion === 'string') {
            if (duracion.match(/^\d{1,2}:\d{2}:\d{2}$/)) {
              duracionFormateada = duracion;
            } else {
              const timeMatch = duracion.match(/(\d{1,2}:\d{2}:\d{2})/);
              if (timeMatch) {
                duracionFormateada = timeMatch[1];
              } else {
                duracionFormateada = duracion;
              }
            }
          } else if (duracion instanceof Date) {
            const hours = duracion.getHours();
            const minutes = duracion.getMinutes();
            const seconds = duracion.getSeconds();
            duracionFormateada = `${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
          } else {
            duracionFormateada = String(duracion);
          }
        }
        datosHistorial.push([
          fechaFormateada,
          cedulaValue,
          nombreValue,
          estadoValue,
          entradaFormateada,
          salidaFormateada,
          duracionFormateada,
          empresaValue
        ]);
      }
    }
    datosHistorial.sort((a, b) => {
      const fechaA = a[0] + ' ' + (a[4] || '00:00:00');
      const fechaB = b[0] + ' ' + (b[4] || '00:00:00');
      return fechaB.localeCompare(fechaA);
    });
    return datosHistorial;
  } catch (error) {
    console.error('❌ Error obteniendo historial:', error);
    registrarLog('ERROR', 'Error obteniendo historial', {
      cedula: cedula,
      error: error.message
    });
    return [];
  }
}

/**
 * Obtiene el historial de registros por cédula o general, con paginación.
 * @param {string} cedula - Cédula a buscar (opcional, si vacío trae todos)
 * @param {number} pagina - Número de página (empezando en 1)
 * @param {number} tamanoPagina - Tamaño de página (registros por página)
 * @return {Object} { total, pagina, tamanoPagina, registros }
 */
function getHistorialPorCedulaPaginado(cedula, pagina, tamanoPagina) {
  try {
    const cedulaLimpia = cedula ? cedula.trim() : '';
    const esBusquedaGeneral = !cedulaLimpia;
    pagina = Math.max(1, parseInt(pagina) || 1);
    tamanoPagina = Math.max(10, parseInt(tamanoPagina) || 50);

    const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
    if (!hojaRegistro) {
      return { total: 0, pagina, tamanoPagina, registros: [] };
    }
    const datos = hojaRegistro.getDataRange().getValues();
    if (datos.length <= 1) {
      return { total: 0, pagina, tamanoPagina, registros: [] };
    }
    const datosHistorial = [];
    const COLS = getColumnasRegistro();
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const cedulaRegistro = String(fila[COLS.CEDULA] || '').trim();
      const coincide = esBusquedaGeneral || (cedulaRegistro === cedulaLimpia);
      if (coincide && cedulaRegistro) {
        let fechaRegistro = fila[COLS.FECHA];
        let fechaFormateada = '';
        if (fechaRegistro instanceof Date) {
          fechaFormateada = fechaRegistro.toISOString().split('T')[0];
        } else if (typeof fechaRegistro === 'string') {
          if (fechaRegistro.match(/^\d{4}-\d{2}-\d{2}$/)) {
            fechaFormateada = fechaRegistro;
          } else {
            try {
              const fecha = new Date(fechaRegistro);
              if (!isNaN(fecha.getTime())) {
                fechaFormateada = fecha.toISOString().split('T')[0];
              } else {
                fechaFormateada = fechaRegistro;
              }
            } catch (e) {
              fechaFormateada = fechaRegistro;
            }
          }
        } else {
          fechaFormateada = new Date().toISOString().split('T')[0];
        }
        const cedulaValue = String(fila[COLS.CEDULA] || 'N/A');
        const nombreValue = String(fila[COLS.NOMBRE] || 'N/A');
        const empresaValue = String(fila[COLS.EMPRESA] || 'N/A');
        const estadoValue = String(fila[COLS.ESTADO_ACCESO] || 'N/A');
        let entradaFormateada = 'N/A';
        if (fila[COLS.ENTRADA]) {
          entradaFormateada = (typeof formatearHora === 'function')
            ? formatearHora(fila[COLS.ENTRADA])
            : fila[COLS.ENTRADA];
        }
        let salidaFormateada = 'N/A';
        if (fila[COLS.SALIDA]) {
          salidaFormateada = (typeof formatearHora === 'function')
            ? formatearHora(fila[COLS.SALIDA])
            : fila[COLS.SALIDA];
        }
        let duracionFormateada = 'N/A';
        if (fila[COLS.DURACION]) {
          const duracion = fila[COLS.DURACION];
          if (typeof duracion === 'string') {
            if (duracion.match(/^\d{1,2}:\d{2}:\d{2}$/)) {
              duracionFormateada = duracion;
            } else {
              const timeMatch = duracion.match(/(\d{1,2}:\d{2}:\d{2})/);
              if (timeMatch) {
                duracionFormateada = timeMatch[1];
              } else {
                duracionFormateada = duracion;
              }
            }
          } else if (duracion instanceof Date) {
            const hours = duracion.getHours();
            const minutes = duracion.getMinutes();
            const seconds = duracion.getSeconds();
            duracionFormateada = `${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
          } else {
            duracionFormateada = String(duracion);
          }
        }
        datosHistorial.push([
          fechaFormateada,
          cedulaValue,
          nombreValue,
          estadoValue,
          entradaFormateada,
          salidaFormateada,
          duracionFormateada,
          empresaValue
        ]);
      }
    }
    datosHistorial.sort((a, b) => {
      const fechaA = a[0] + ' ' + (a[4] || '00:00:00');
      const fechaB = b[0] + ' ' + (b[4] || '00:00:00');
      return fechaB.localeCompare(fechaA);
    });
    const total = datosHistorial.length;
    const inicio = (pagina - 1) * tamanoPagina;
    const fin = inicio + tamanoPagina;
    const registros = datosHistorial.slice(inicio, fin);
    return { total, pagina, tamanoPagina, registros };
  } catch (error) {
    console.error('❌ Error obteniendo historial paginado:', error);
    registrarLog('ERROR', 'Error obteniendo historial paginado', {
      cedula: cedula,
      error: error.message
    });
    return { total: 0, pagina: 1, tamanoPagina: 50, registros: [] };
  }
}