/**
 * Devuelve todos los registros de personal para el panel de administración.
 * Estructura: [{Cédula, Nombre, Empresa, Área, Cargo, Estado, indice}]
 * @return {Object} {success, personal: Array, error}
 */
function obtenerPersonalParaAdmin() {
	try {
		var ss = SpreadsheetApp.getActive();
		var hoja = ss.getSheetByName('Base de Datos');
		if (!hoja) {
			return { success: false, personal: [], error: 'Hoja "Base de Datos" no encontrada' };
		}
		var datos = hoja.getDataRange().getValues();
		if (datos.length <= 1) {
			return { success: true, personal: [] };
		}
		var encabezados = datos[0];
		var personal = [];
		for (var i = 1; i < datos.length; i++) {
			var fila = datos[i];
			var persona = {};
			for (var j = 0; j < encabezados.length; j++) {
				var key = encabezados[j];
				persona[key] = fila[j];
			}
			persona.indice = i + 1; // Para edición directa si se requiere
			personal.push(persona);
		}
		return { success: true, personal: personal };
	} catch (e) {
		return { success: false, personal: [], error: e.message };
	}
}
function getEvacuacionDataForClient() {
    try {
        console.log('🚨 Iniciando getEvacuacionDataForClient mejorado...');
        
        const ss = SpreadsheetApp.getActive();
        const hoja = ss.getSheetByName('Registro');
        
        if (!hoja) {
            return {
                success: true,
                personasDentro: [],
                totalDentro: 0,
                message: 'Hoja de Registro no encontrada'
            };
        }
        
        const lastRow = hoja.getLastRow();
        console.log(`📊 Procesando ${lastRow} filas`);
        
        if (lastRow <= 1) {
            return {
                success: true,
                personasDentro: [],
                totalDentro: 0,
                message: 'No hay registros en la hoja'
            };
        }
        
        // Obtener todos los datos
        const datos = hoja.getRange(1, 1, lastRow, hoja.getLastColumn()).getValues();
        const encabezados = datos[0];
        
        console.log('📋 Encabezados:', encabezados);
        
        // Mapear columnas con los nombres exactos de la hoja
        const colCedula = 1;    // Columna B = índice 1
        const colNombre = 2;    // Columna C = índice 2  
        const colEmpresa = 7;   // Columna H = índice 7
        const colEntrada = 4;   // Columna E = índice 4
        const colSalida = 5;    // Columna F = índice 5
        
        console.log('🗺️ Mapeo de columnas:', {colCedula, colNombre, colEmpresa, colEntrada, colSalida});
        
        // Procesar datos con lógica mejorada
        const personasPorCedula = {};
        
        // Recorrer todos los registros
        for (let i = 1; i < datos.length; i++) {
            const fila = datos[i];
            const cedula = fila[colCedula] ? String(fila[colCedula]).trim() : '';
            
            if (!cedula) continue;
            
            const entrada = fila[colEntrada];
            const salida = fila[colSalida];
            
            // Verificar si tiene entrada válida
            const tieneEntrada = entrada && 
                                entrada !== null && 
                                entrada !== undefined && 
                                String(entrada).trim() !== '' &&
                                entrada.toString() !== 'Invalid Date';
            
            // Verificar si tiene salida válida
            const tieneSalida = salida && 
                               salida !== null && 
                               salida !== undefined && 
                               String(salida).trim() !== '' &&
                               salida.toString() !== 'Invalid Date';
            
            console.log(`Fila ${i+1} - Cédula: ${cedula}, Entrada: ${tieneEntrada ? 'SI' : 'NO'}, Salida: ${tieneSalida ? 'SI' : 'NO'}`);
            
            if (tieneEntrada && !tieneSalida) {
                // Persona con entrada pero sin salida = está dentro
                personasPorCedula[cedula] = {
                    cedula: cedula,
                    nombre: fila[colNombre] ? String(fila[colNombre]).trim() : 'Sin nombre',
                    empresa: fila[colEmpresa] ? String(fila[colEmpresa]).trim() : 'Sin empresa',
                    horaEntrada: typeof entrada === 'object' && entrada.getHours ? 
                        entrada.toLocaleTimeString('es-PA', {hour: '2-digit', minute: '2-digit'}) : 
                        String(entrada)
                };
                console.log(`✅ Persona DENTRO: ${cedula} - ${personasPorCedula[cedula].nombre}`);
            } else if (tieneSalida) {
                // Si tiene salida, eliminar de personas dentro (puede haber entrado y salido varias veces)
                delete personasPorCedula[cedula];
                console.log(`🚪 Persona SALIÓ: ${cedula}`);
            }
        }
        
        const personasDentro = Object.values(personasPorCedula);
        
        console.log(`🏢 RESULTADO FINAL: ${personasDentro.length} personas dentro del edificio`);
        personasDentro.forEach((p, i) => {
            console.log(`  ${i+1}. ${p.nombre} (${p.cedula}) - ${p.empresa}`);
        });
        
        return {
            success: true,
            personasDentro: personasDentro,
            totalDentro: personasDentro.length,
            message: `${personasDentro.length} personas encontradas dentro del edificio`,
            debug: {
                filasProcesadas: datos.length - 1,
                personasEncontradas: personasDentro.length,
                mapeoColumnas: {colCedula, colNombre, colEmpresa, colEntrada, colSalida}
            }
        };
        
    } catch (error) {
        console.error('❌ Error en getEvacuacionDataForClient:', error);
        return {
            success: false,
            personasDentro: [],
            totalDentro: 0,
            message: `Error: ${error.message}`,
            debug: { error: error.message, stack: error.stack }
        };
    }
}

/**
 * Registra salidas masivas para evacuación real o simulacro.
 * @param {Array<string>} cedulas - Lista de cédulas a evacuar
 * @param {string} tipo - "REAL" o "SIMULACRO"
 * @returns {Object} Resultado de la operación
 */
function registrarSalidasMasivas(cedulas, tipo) {
	try {
		if (!Array.isArray(cedulas) || cedulas.length === 0) {
			return { success: false, message: 'No se proporcionaron cédulas.' };
		}
		tipo = (tipo || '').toUpperCase();
		const now = new Date();
		const horaSalida = formatearHora(now);
		const ss = SpreadsheetApp.getActive();
		const COLS = getColumnasRegistro();
		let procesados = 0, errores = [];

		if (tipo === 'REAL') {
			// Evacuación real: completar la última entrada activa de cada cédula
			const hojaRegistro = ss.getSheetByName(SHEET_REGISTRO);
			if (!hojaRegistro) throw new Error('Hoja de registro no encontrada');
			const datos = hojaRegistro.getDataRange().getValues();
			for (let cedula of cedulas) {
				let encontrada = false;
				for (let i = datos.length - 1; i >= 1; i--) {
					const fila = datos[i];
					const cedulaRegistro = String(fila[COLS.CEDULA]).trim();
					const entrada = fila[COLS.ENTRADA];
					const salida = fila[COLS.SALIDA];
					if (cedulaRegistro === cedula && entrada && !salida) {
						hojaRegistro.getRange(i + 1, COLS.SALIDA + 1).setValue(horaSalida);
						// Calcular duración
						let duracion = '';
						try {
							duracion = calcularDuracionSimple(entrada, horaSalida);
							hojaRegistro.getRange(i + 1, COLS.DURACION + 1).setValue(duracion);
						} catch (e) {}
						encontrada = true;
						procesados++;
						break;
					}
				}
				if (!encontrada) {
					errores.push(`No se encontró entrada activa para ${cedula}`);
				}
			}
			registrarLog('INFO', 'Evacuación REAL ejecutada', { cedulas, hora: horaSalida, procesados, errores });
			return {
				success: true,
				message: `Evacuación REAL registrada para ${procesados} personas${errores.length ? '. Errores: ' + errores.join('; ') : ''}`,
				procesados,
				errores
			};
		} else if (tipo === 'SIMULACRO') {
			// Simulacro: crear hoja nueva y registrar salidas
			const fechaStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
			const nombreHoja = `Simulacro_${fechaStr}`;
			let hojaSimulacro = ss.getSheetByName(nombreHoja);
			if (!hojaSimulacro) {
				hojaSimulacro = ss.insertSheet(nombreHoja);
				// Encabezados básicos
				hojaSimulacro.appendRow(['Cédula', 'Nombre', 'Empresa', 'Hora Salida', 'Tipo Evento']);
			}
			inicializarCacheUsuarios();
			for (let cedula of cedulas) {
				let persona = buscarUsuario(cedula);
				hojaSimulacro.appendRow([
					cedula,
					persona && persona.nombre ? persona.nombre : '',
					persona && persona.empresa ? persona.empresa : '',
					horaSalida,
					'SIMULACRO'
				]);
				procesados++;
			}
			registrarLog('INFO', 'Simulacro de evacuación ejecutado', { cedulas, hoja: nombreHoja, hora: horaSalida });
			return {
				success: true,
				message: `Simulacro registrado para ${procesados} personas en hoja ${nombreHoja}`,
				hoja: nombreHoja,
				procesados
			};
		} else {
			return { success: false, message: 'Tipo de evento no válido. Debe ser REAL o SIMULACRO.' };
		}
	} catch (e) {
		console.error('Error en registrarSalidasMasivas:', e);
		registrarLog('ERROR', 'Error en registrarSalidasMasivas', { error: e.message, cedulas, tipo });
		return { success: false, message: 'Error al registrar salidas masivas: ' + e.message };
	}
}

 /**
 * Obtiene estadísticas en tiempo real para el panel flotante - VERSIÓN OPTIMIZADA
 * @return {Object} Estadísticas completas
 */
function obtenerEstadisticas() {
    try {
        const COLS = getColumnasRegistro();
        const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
        
        // OPTIMIZACIÓN 1: Solo obtener datos si hay filas
        const lastRow = hojaRegistro.getLastRow();
        if (lastRow <= 1) {
            return {
                entradas: 0,
                salidas: 0,
                salidasValidas: 0,
                personasDentro: 0,
                salidasSinEntrada: 0,
                recentRecords: [],
                timestamp: new Date().toISOString()
            };
        }
        
        // OPTIMIZACIÓN 2: Obtener solo las columnas necesarias y con límite de filas
        const maxRows = Math.min(lastRow - 1, 300); // Limitar a 300 registros más recientes
        const startRow = Math.max(2, lastRow - maxRows + 1);
        const numCols = Math.max(COLS.CEDULA, COLS.NOMBRE, COLS.EMPRESA, COLS.ENTRADA, COLS.SALIDA, COLS.ESTADO_ACCESO, COLS.COMENTARIO) + 1;
        
        const datos = hojaRegistro.getRange(startRow, 1, maxRows, numCols).getValues();
        
        let entradas = 0, salidas = 0, personasDentro = 0, salidasSinEntrada = 0;
        const cedulasDentro = {};
        const entradasPorCedula = {};
        const registrosRecientes = [];
        
		// OPTIMIZACIÓN 3: Procesar registros con menos operaciones de string
		for (let i = 0; i < datos.length; i++) {
			const cedula = String(datos[i][COLS.CEDULA] || '').trim();
			if (!cedula) continue;

			const nombre = datos[i][COLS.NOMBRE] || 'Sin nombre';
			const entrada = datos[i][COLS.ENTRADA];
			const salida = datos[i][COLS.SALIDA];
			const estadoAcceso = String(datos[i][COLS.ESTADO_ACCESO] || '').trim();
			const comentario = String(datos[i][COLS.COMENTARIO] || '').trim();

			// Registrar las entradas 
			if (entrada) {
				entradas++;
				entradasPorCedula[cedula] = true;

				if (!salida) {
					cedulasDentro[cedula] = {
						cedula: cedula,
						nombre: nombre,
						empresa: datos[i][COLS.EMPRESA] || 'N/A',
						horaEntrada: entrada,
						estadoAcceso: estadoAcceso
					};
				}
			}

			// Contar salidas
			if (salida) {
				salidas++;

				// OPTIMIZACIÓN 4: Verificación más rápida para salidas sin entrada
				if (!entrada && !entradasPorCedula[cedula]) {
					const esJustificada = comentario.includes('JUSTIFICADO') || 
										comentario.includes('VISITANTE') || 
										comentario.includes('EVACUACIÓN') ||
										estadoAcceso === 'Acceso Temporal';

					if (!esJustificada) {
						salidasSinEntrada++;
					}
				}

				// Eliminar de personas dentro
				delete cedulasDentro[cedula];
			}

			// OPTIMIZACIÓN 5: Solo recopilar los últimos 10 registros
			if (i >= Math.max(0, datos.length - 10)) {
				const tipo = salida ? 'salida' : (entrada ? 'entrada' : '');
				if (tipo) {
					const hora = salida || entrada;
					let duracion = '';
					// Solo calcular duración si hay entrada y salida
					if (entrada && salida) {
						try {
							duracion = calcularDuracionSimple(entrada, salida);
						} catch (e) {
							duracion = '';
						}
					}
					registrosRecientes.unshift({
						cedula: cedula.length > 8 ? cedula.substring(0, 8) : cedula,
						nombre: nombre,
						accion: tipo,
						estadoAcceso: estadoAcceso,
						hora: hora ? (typeof hora === 'string' ? hora : new Date(hora).toLocaleTimeString('es-PA', { 
							hour: '2-digit', 
							minute: '2-digit' 
						})) : 'N/A',
						duracion: duracion,
						timestamp: new Date().toISOString()
					});
				}
			}
		}
        
        personasDentro = Object.keys(cedulasDentro).length;
        
        return {
            entradas: entradas,
            salidas: salidas,
            salidasValidas: salidas - salidasSinEntrada,
            personasDentro: personasDentro,
            salidasSinEntrada: salidasSinEntrada,
            recentRecords: registrosRecientes.slice(0, 10),
            timestamp: new Date().toISOString(),
            optimizado: true,
            procesados: maxRows
        };
        
    } catch (e) {
        console.error('Error obteniendo estadísticas:', e);
        return {
            entradas: 0,
            salidas: 0,
            salidasValidas: 0,
            personasDentro: 0,
            salidasSinEntrada: 0,
            recentRecords: [],
            timestamp: new Date().toISOString(),
            error: e.message
        };
    }
}

// Función auxiliar simplificada para calcular duración
function calcularDuracionSimple(horaEntrada, horaSalida) {
    try {
        if (!horaEntrada || !horaSalida) return '0:00:00';
        
        let entradaHoras = 0, entradaMinutos = 0, entradaSegundos = 0;
        let salidaHoras = 0, salidaMinutos = 0, salidaSegundos = 0;
        
        // Parsear hora de entrada
        if (typeof horaEntrada === 'string' && horaEntrada.includes(':')) {
            const partes = horaEntrada.split(':');
            entradaHoras = parseInt(partes[0]) || 0;
            entradaMinutos = parseInt(partes[1]) || 0;
            entradaSegundos = parseInt(partes[2]) || 0;
        } else if (horaEntrada instanceof Date) {
            entradaHoras = horaEntrada.getHours();
            entradaMinutos = horaEntrada.getMinutes();
            entradaSegundos = horaEntrada.getSeconds();
        }
        
        // Parsear hora de salida
        if (typeof horaSalida === 'string' && horaSalida.includes(':')) {
            const partes = horaSalida.split(':');
            salidaHoras = parseInt(partes[0]) || 0;
            salidaMinutos = parseInt(partes[1]) || 0;
            salidaSegundos = parseInt(partes[2]) || 0;
        } else if (horaSalida instanceof Date) {
            salidaHoras = horaSalida.getHours();
            salidaMinutos = horaSalida.getMinutes();
            salidaSegundos = horaSalida.getSeconds();
        }
        
        // Calcular diferencia total en segundos
        let totalSegundosEntrada = entradaHoras * 3600 + entradaMinutos * 60 + entradaSegundos;
        let totalSegundosSalida = salidaHoras * 3600 + salidaMinutos * 60 + salidaSegundos;
        
        // Si la salida es menor que la entrada, asumimos que es el día siguiente
        if (totalSegundosSalida < totalSegundosEntrada) {
            totalSegundosSalida += 24 * 3600;
        }
        
        const diferencia = totalSegundosSalida - totalSegundosEntrada;
        
        const horas = Math.floor(diferencia / 3600);
        const minutos = Math.floor((diferencia % 3600) / 60);
        const segundos = diferencia % 60;
        
        return `${horas}:${String(minutos).padStart(2, '0')}:${String(segundos).padStart(2, '0')}`;
        
    } catch (e) {
        console.error('Error en calcularDuracionSimple:', e);
        return '0:00:00';
    }
}

/**
 * Busca un usuario en el cache por cédula o nombre.
 * @param {string} cedulaONombre - Cédula o nombre a buscar
 * @returns {Object|null} Usuario encontrado o null
 */
function buscarUsuario(cedulaONombre) {
	if (!cedulaONombre) return null;
	const busqueda = String(cedulaONombre).trim();
	// Función interna para buscar en el cache
	function buscarEnCache() {
		// Búsqueda exacta por cédula
		if (cacheUsuarios[busqueda]) {
			return cacheUsuarios[busqueda];
		}
		// Búsqueda por cédula normalizada (solo números)
		const cedulaNormalizada = busqueda.replace(/[^\d]/g, '');
		for (let cedula in cacheUsuarios) {
			const usuario = cacheUsuarios[cedula];
			if (usuario.cedulaNormalizada === cedulaNormalizada) {
				return usuario;
			}
		}
		// Búsqueda por texto en campo busqueda
		const query = busqueda.toLowerCase();
		for (let cedula in cacheUsuarios) {
			const usuario = cacheUsuarios[cedula];
			if (usuario.busqueda && usuario.busqueda.includes(query)) {
				return usuario;
			}
		}
		return null;
	}
	// Primer intento en cache actual
	let usuario = buscarEnCache();
	if (usuario) return usuario;
	// Si no se encuentra, forzar recarga del cache y reintentar una vez
	if (typeof inicializarCacheUsuarios === 'function') {
		inicializarCacheUsuarios();
		usuario = buscarEnCache();
		if (usuario) return usuario;
	}
	return null;
}
/**
 * Exporta la base de datos de personal a un archivo CSV en Drive y retorna el enlace.
 * @returns {Object} Resultado de la exportación
 */
function exportarBaseDeDatos() {
		try {
			console.log('📤 [ADMIN] Exportando base de datos...');
      
			let hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_BD);
			if (!hoja) {
				hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_BD);
			}

			if (!hoja) {
				return {
					success: false,
					message: 'Hoja de personal no encontrada'
				};
			}

			const data = hoja.getDataRange().getValues();
			const csvContent = data.map(row => row.join(',')).join('\n');
      
			const fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
			const nombreArchivo = `SurPass_Personal_${fechaActual}.csv`;
      
			const blob = Utilities.newBlob(csvContent, 'text/csv', nombreArchivo);
			const file = DriveApp.createFile(blob);
			file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

			if (typeof registrarLog === 'function') {
				registrarLog('INFO', 'Base de datos exportada desde admin', {
					registros: data.length - 1,
					archivo: nombreArchivo,
					usuario: 'Administrador'
				});
			}

			console.log(`✅ [ADMIN] Base de datos exportada: ${nombreArchivo} (${data.length - 1} registros)`);
      
			return {
				success: true,
				message: `Base de Datos exportada: ${data.length - 1} registros`,
				fileUrl: file.getUrl(),
				fileName: nombreArchivo
			};
      
		} catch (error) {
			console.error('❌ [ADMIN] Error exportando base de datos:', error);
			if (typeof registrarLog === 'function') {
				registrarLog('ERROR', 'Error exportando base de datos desde admin', { 
					error: error.message,
					stack: error.stack
				}, 'Panel Admin');
			}
			return {
				success: false,
				code: 'EXPORT_FAILED',
				message: 'Error al exportar la Base de Datos: ' + error.message,
				details: { originalError: error.message, stack: error.stack }
			};
		}
	}

function handleHTMLFormSubmit(cedula, tipoAcceso) {
    console.log('DEBUG handleHTMLFormSubmit - parámetros recibidos:');
    console.log('  cedula:', cedula, 'tipo:', typeof cedula);
    console.log('  tipoAcceso:', tipoAcceso, 'tipo:', typeof tipoAcceso);
    
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);

    try {
        if (!cedula || !tipoAcceso) {
            console.error('ERROR: Parámetros faltantes en handleHTMLFormSubmit');
            return { success: false, message: "Cédula o tipo de acceso no proporcionados.", color: 'red' };
        }

        inicializarCacheUsuarios();
        const persona = buscarUsuario(cedula.trim());
        const entradaActiva = verificarEntradaActiva(cedula.trim());

        // USAR LAS FUNCIONES ESPECIALIZADAS EN LUGAR DE LÓGICA DIRECTA
        if (tipoAcceso === 'entrada') {
            return procesarEntrada(persona, entradaActiva, cedula.trim());
        } else if (tipoAcceso === 'salida') {
            return procesarSalida(persona, entradaActiva, cedula.trim());
        } else {
            throw new Error("Tipo de acceso inválido");
        }
        
    } catch (e) {
        console.error('❌ Error en handleHTMLFormSubmit:', e);
        return { success: false, message: `Error: ${e.message}`, color: 'red' };
    } finally {
        lock.releaseLock();
    }
}

	/**
	 * Maneja el submit del formulario HTML para entrada o salida, permitiendo agregar un comentario adicional al registro.
	 * @param {string} cedula - Cédula de la persona
	 * @param {string} tipoAcceso - 'entrada' o 'salida'
	 * @param {string} [comentarioOpcional] - Comentario adicional a agregar
	 * @returns {Object} Resultado del proceso
	 */
	function handleHTMLFormSubmitWithComment(cedula, tipoAcceso, comentarioOpcional) {
		try {
			console.log(`🔄 Procesando con comentario: ${tipoAcceso} para ${cedula}`);
			// Usar la función principal pero con comentario adicional
			var resultado = handleHTMLFormSubmit(cedula, tipoAcceso);
			// Si el registro fue exitoso y hay comentario, agregarlo
			if (resultado.success && comentarioOpcional && comentarioOpcional.trim()) {
				try {
					var hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
					var datos = hojaRegistro.getDataRange().getValues();
					var COLS = getColumnasRegistro();
					var cedulaLimpia = String(cedula).trim();
					// Buscar la entrada más reciente de esta persona (desde el final hacia atrás)
					var filaEncontrada = -1;
					for (var i = datos.length - 1; i >= 1; i--) {
						var cedulaRegistro = String(datos[i][COLS.CEDULA]).trim();
						if (cedulaRegistro === cedulaLimpia) {
							filaEncontrada = i + 1; // +1 porque getRange usa índices basados en 1
							break;
						}
					}
					if (filaEncontrada > 1) {
						var comentarioActual = hojaRegistro.getRange(filaEncontrada, COLS.COMENTARIO + 1).getValue() || '';
						var comentarioCompleto = comentarioActual ?
							comentarioActual + ' | COMENTARIO: ' + comentarioOpcional.trim() :
							'COMENTARIO: ' + comentarioOpcional.trim();
						hojaRegistro.getRange(filaEncontrada, COLS.COMENTARIO + 1).setValue(comentarioCompleto);
						// También actualizar en historial (buscar el registro correspondiente)
						var hojaHistorial = SpreadsheetApp.getActive().getSheetByName(SHEET_HISTORIAL);
						if (hojaHistorial) {
							var datosHistorial = hojaHistorial.getDataRange().getValues();
							for (var j = datosHistorial.length - 1; j >= 1; j--) {
								var cedulaHistorial = String(datosHistorial[j][COLS.CEDULA]).trim();
								var fechaHistorial = datosHistorial[j][COLS.FECHA];
								var fechaRegistro = datos[filaEncontrada - 1][COLS.FECHA];
								if (cedulaHistorial === cedulaLimpia && fechaHistorial === fechaRegistro) {
									hojaHistorial.getRange(j + 1, COLS.COMENTARIO + 1).setValue(comentarioCompleto);
									break;
								}
							}
						}
						console.log('📝 Comentario agregado al registro de ' + cedulaLimpia + ' en fila ' + filaEncontrada + ': ' + comentarioOpcional.substring(0, 50) + '...');
					} else {
						console.warn('⚠️ No se encontró registro para agregar comentario: ' + cedulaLimpia);
					}
				} catch (comentarioError) {
					console.warn('⚠️ Error agregando comentario, pero registro completado:', comentarioError);
				}
			}
			return resultado;
		} catch (e) {
			console.error('❌ Error en handleHTMLFormSubmitWithComment:', e);
			return {
				success: false,
				message: 'Error interno procesando registro con comentario',
				color: 'red',
				tipoFlujo: 'error_interno'
			};
		}
	}

function procesarEntrada(persona, entradaActiva, cedula) {
    console.log('DEBUG procesarEntrada - cedula recibida:', cedula, 'tipo:', typeof cedula);
    const now = new Date();
    const horaFormateada = formatearHora(now);
    let resultado = {
        success: false,
        message: '',
        color: 'red',
        icon: '',
        tipoFlujo: '',
        comentarioObligatorio: false,
        modalTipo: '',
        registroPendiente: false,
        registroPendienteExistente: false
    };

    if (persona) { // Persona registrada
        if (entradaActiva && entradaActiva.found) {
            // ESCENARIO: Entrada duplicada
            resultado.tipoFlujo = 'entrada_duplicada';
            resultado.message = `⚠️ ENTRADA DUPLICADA\nYa existe una entrada activa para ${persona.nombre} desde las ${entradaActiva.horaEntrada}.`;
            resultado.color = 'orange';
            resultado.icon = 'fas fa-exclamation-triangle';
            resultado.registroPendienteExistente = true;
        } else {
            // ESCENARIO 1: Entrada normal
            registrarEnHoja(cedula, persona.nombre, persona.empresa, now, '', 'entrada', 'web', '');
            resultado.success = true;
            resultado.tipoFlujo = 'entrada_registrada';
            resultado.message = `✅ ENTRADA REGISTRADA\n\n👤 ${persona.nombre}\n🏢 ${persona.empresa}\n⏰ ${horaFormateada}`;
            resultado.color = 'green';
            resultado.icon = 'fas fa-sign-in-alt';
        }
    } else { // Persona no registrada
        // ESCENARIO 4: Entrada de persona no registrada - ABRIR MODAL DE REGISTRO
        resultado.tipoFlujo = 'entrada_persona_no_registrada';
        resultado.message = `🚫 PERSONA NO IDENTIFICADA\n\nComplete la información para un registro temporal.`;
        resultado.color = 'red';
        resultado.icon = 'fas fa-user-slash';
        resultado.comentarioObligatorio = true;
        resultado.modalTipo = 'registro';  // ← ESTO ABRE EL MODAL DE REGISTRO
        resultado.registroPendiente = true;
    }

    // Loguear la acción
    if (typeof registrarLog === 'function') {
        registrarLog('INFO', `Proceso de entrada finalizado para ${cedula}`, { tipoFlujo: resultado.tipoFlujo });
    }
    return resultado;
}

function procesarSalida(persona, entradaActiva, cedula) {
    console.log('DEBUG procesarSalida - cedula recibida:', cedula, 'tipo:', typeof cedula);
    const now = new Date();
    const horaFormateada = formatearHora(now);
    let resultado = {
        success: false,
        message: '',
        color: 'red',
        icon: '',
        tipoFlujo: '',
        comentarioObligatorio: false,
        modalTipo: '',
        registroPendiente: false
    };

    if (entradaActiva && entradaActiva.found) {
        // ESCENARIO 2 o 5: Salida con entrada previa
        completarSalida(entradaActiva.fila, now);
        resultado.success = true;
        resultado.tipoFlujo = persona ? 'salida_normal' : 'salida_persona_no_registrada_con_entrada';
        resultado.message = `✅ SALIDA REGISTRADA\n\n👤 ${entradaActiva.nombre}\n🏢 ${entradaActiva.empresa}\n⏰ ${horaFormateada}`;
        resultado.color = 'green';
        resultado.icon = 'fas fa-sign-out-alt';
    } else {
        if (persona) {
            // ESCENARIO 3: Persona registrada sin entrada previa - MODAL DE JUSTIFICACIÓN
            resultado.tipoFlujo = 'salida_sin_entrada_previa';
            resultado.message = `⚠️ SALIDA SIN ENTRADA PREVIA\n\nNo se encontró una entrada para ${persona.nombre}. Se requiere una justificación.`;
            resultado.color = 'red';
            resultado.icon = 'fas fa-exclamation-triangle';
            resultado.comentarioObligatorio = true;
            resultado.modalTipo = 'justificacion';  // ← MODAL DE JUSTIFICACIÓN
        } else {
            // ESCENARIO 6: Persona no registrada sin entrada previa - MODAL DE REGISTRO
            resultado.tipoFlujo = 'salida_sin_entrada_previa_persona_no_registrada';
            resultado.message = `🚫 PERSONA NO IDENTIFICADA\n\nNo hay registro de entrada para esta cédula. Complete la información para registrar la salida.`;
            resultado.color = 'red';
            resultado.icon = 'fas fa-user-times';
            resultado.comentarioObligatorio = true;
            resultado.modalTipo = 'registro';  // ← MODAL DE REGISTRO
            resultado.registroPendiente = true;
        }
    }

    // Loguear la acción
    if (typeof registrarLog === 'function') {
        registrarLog('INFO', `Proceso de salida finalizado para ${cedula}`, { tipoFlujo: resultado.tipoFlujo });
    }
    return resultado;
}

/**
 * Registra un evento de entrada o salida en la hoja de registro y en el historial
 */
function registrarEnHoja(cedula, nombre, empresa, horaEntrada, horaSalida, tipo, origen, comentario) {
	try {
		// Validación robusta de parámetros
		if (!cedula || typeof cedula !== 'string' || cedula.trim() === '') {
			console.error(`❌ registrarEnHoja: Cédula inválida: ${cedula}`);
			return { success: false, error: 'Cédula no proporcionada o inválida' };
		}
		// Estructura de columnas solo una vez
		const COLS = getColumnasRegistro();
		const maxColumna = Math.max(...Object.values(COLS)) + 1;
		const fila = new Array(maxColumna).fill('');
		// Fecha actual ISO (solo fecha)
		const fechaActual = new Date().toISOString().split('T')[0];
		fila[COLS.FECHA] = fechaActual;
		fila[COLS.CEDULA] = cedula.trim();
		fila[COLS.NOMBRE] = nombre || 'N/A';
		fila[COLS.EMPRESA] = empresa || 'N/A';
		// Estado de acceso
		let esAccesoTemporal = false;
		if (comentario) {
			const comentarioMayus = comentario.toUpperCase();
			esAccesoTemporal = /VISITANTE|JUSTIFICADO|SIN ENTRADA PREVIA/.test(comentarioMayus);
		}
		if (tipo === 'entrada') {
			fila[COLS.ESTADO_ACCESO] = esAccesoTemporal ? ESTADOS_ACCESO.TEMPORAL : ESTADOS_ACCESO.PERMITIDO;
			fila[COLS.ENTRADA] = horaEntrada ? formatearHora(horaEntrada) : '';
			fila[COLS.SALIDA] = '';
		} else if (tipo === 'salida') {
			if (!horaEntrada || esAccesoTemporal) {
				fila[COLS.ESTADO_ACCESO] = ESTADOS_ACCESO.TEMPORAL;
			} else {
				fila[COLS.ESTADO_ACCESO] = 'Salida Registrada';
			}
			fila[COLS.ENTRADA] = '';
			fila[COLS.SALIDA] = horaSalida ? formatearHora(horaSalida) : '';
			// Calcular duración solo si hay entrada y salida válidas
			if (horaEntrada && horaSalida) {
				try {
					const [anio, mes, dia] = fechaActual.split('-');
					const fechaFormateada = `${dia}/${mes}/${anio}`;
					const entradaFormateada = formatearHora(horaEntrada);
					const salidaFormateada = formatearHora(horaSalida);
					const entradaStr = `${fechaFormateada} ${entradaFormateada}`;
					const salidaStr = `${fechaFormateada} ${salidaFormateada}`;
					fila[COLS.DURACION] = calcularDuracion(entradaStr, salidaStr);
				} catch (e) {
					console.warn('Error calculando duración:', e);
					fila[COLS.DURACION] = '';
				}
			}
		} else {
			fila[COLS.ESTADO_ACCESO] = '';
			fila[COLS.ENTRADA] = '';
			fila[COLS.SALIDA] = '';
		}
		fila[COLS.COMENTARIO] = comentario || '';
		// Obtener hojas solo una vez
		const ss = SpreadsheetApp.getActive();
		const hojaRegistro = ss.getSheetByName(SHEET_REGISTRO);
		const hojaHistorial = ss.getSheetByName(SHEET_HISTORIAL);
		if (!hojaRegistro || !hojaHistorial) {
			const msg = 'No se encontró la hoja de registro o historial';
			console.error(msg);
			return { success: false, error: msg };
		}
		// Evitar duplicados: solo registrar si la última fila no es idéntica
		const ultimaFila = hojaRegistro.getLastRow() > 1 ? hojaRegistro.getRange(hojaRegistro.getLastRow(), 1, 1, hojaRegistro.getLastColumn()).getValues()[0] : null;
		let esDuplicado = false;
		if (ultimaFila) {
			esDuplicado = fila.every((val, idx) => String(val) === String(ultimaFila[idx]));
		}
		if (!esDuplicado) {
			hojaRegistro.appendRow(fila);
			hojaHistorial.appendRow(fila);
		} else {
			console.warn('Intento de registro duplicado evitado para:', cedula, nombre, tipo);
		}
		const filaIndex = hojaRegistro.getLastRow();
		console.log(`📝 Registro añadido optimizado: ${tipo} para ${nombre} (${cedula}) en fila ${filaIndex}`);
		return {
			success: true,
			filaIndex: filaIndex,
			cedula: cedula,
			nombre: nombre,
			tipo: tipo,
			optimizado: true,
			duplicado: esDuplicado
		};
	} catch (e) {
		console.error('Error registrando en hoja:', e);
		if (typeof registrarErrorLog === 'function') {
			registrarErrorLog('ERROR_REGISTRAR_HOJA', e.message, { cedula, nombre, tipo });
		}
		return { success: false, error: e.message };
	}
}

/**
 * Analiza la estructura y calidad de la hoja de base de datos (Personal)
 */
function analizarEstructuraBaseDatos() {
	try {
		console.log('=== ANÁLISIS DE ESTRUCTURA DE BASE DE DATOS ===');
		const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_BD);
		const datos = hoja.getDataRange().getValues();
		if (datos.length === 0) {
			console.log('❌ La hoja de base de datos está vacía');
			return {
				success: false,
				message: 'La hoja de base de datos está vacía'
			};
		}
		const encabezados = datos[0];
		console.log('🔍 ENCABEZADOS DE LA BASE DE DATOS:');
		encabezados.forEach((encabezado, indice) => {
			console.log(`   Columna ${indice}: "${encabezado}"`);
		});
		console.log('\n📊 ESTRUCTURA ESPERADA POR EL SISTEMA:');
		console.log('   Columna 0: Cédula (usado como identificador principal)');
		console.log('   Columna 1: Nombre (usado para mostrar información)');
		console.log('   Columna 2: Empresa (usado para mostrar información)');
		console.log('   Columnas adicionales: Se mapean según los encabezados');
		if (datos.length > 1) {
			console.log('\n📝 MUESTRA DE DATOS (primeras 3 filas):');
			for (let i = 1; i <= Math.min(3, datos.length - 1); i++) {
				const fila = datos[i];
				console.log(`   Fila ${i}:`);
				fila.forEach((valor, indice) => {
					console.log(`     Columna ${indice} (${encabezados[indice] || 'Sin encabezado'}): "${valor}"`);
				});
				console.log('   ---');
			}
		}
		console.log('\n🔧 NORMALIZACIÓN DE CAMPOS:');
		console.log('   El sistema busca los campos "Nombre" o "nombre" para el nombre');
		console.log('   El sistema busca los campos "Empresa" o "empresa" para la empresa');
		console.log('   La cédula siempre se toma de la columna 0');
		// Análisis de datos
		let totalRegistros = datos.length - 1;
		let registrosConCedula = 0;
		let registrosConNombre = 0;
		let registrosConEmpresa = 0;
		for (let i = 1; i < datos.length; i++) {
			const fila = datos[i];
			if (fila[0] && String(fila[0]).trim()) registrosConCedula++;
			if (fila[1] && String(fila[1]).trim()) registrosConNombre++;
			if (fila[2] && String(fila[2]).trim()) registrosConEmpresa++;
		}
		console.log('\n📈 ESTADÍSTICAS DE CALIDAD DE DATOS:');
		console.log(`   Total de registros: ${totalRegistros}`);
		console.log(`   Registros con cédula: ${registrosConCedula} (${((registrosConCedula/totalRegistros)*100).toFixed(1)}%)`);
		console.log(`   Registros con nombre: ${registrosConNombre} (${((registrosConNombre/totalRegistros)*100).toFixed(1)}%)`);
		console.log(`   Registros con empresa: ${registrosConEmpresa} (${((registrosConEmpresa/totalRegistros)*100).toFixed(1)}%)`);
		return {
			success: true,
			encabezados: encabezados,
			totalRegistros: totalRegistros,
			estructura: {
				cedula: { columna: 0, nombre: encabezados[0], llenos: registrosConCedula },
				nombre: { columna: 1, nombre: encabezados[1], llenos: registrosConNombre },
				empresa: { columna: 2, nombre: encabezados[2], llenos: registrosConEmpresa }
			},
			muestra: datos.slice(1, 4).map(fila => ({
				cedula: fila[0],
				nombre: fila[1],
				empresa: fila[2],
				otros: fila.slice(3)
			}))
		};
	} catch (e) {
		console.error('❌ Error analizando estructura:', e);
		return {
			success: false,
			message: 'Error analizando la estructura: ' + e.message
		};
	}
}
/**
 * Busca la entrada previa de una persona para mostrar información en tarjeta
 */
function buscarEntradaPrevia(cedula) {
	try {
		// Validar parámetro obligatorio
		if (!cedula || typeof cedula !== 'string' || cedula.trim() === '') {
			console.warn(`⚠️ buscarEntradaPrevia: Cédula inválida: ${cedula}`);
			return { found: false, error: 'Cédula no proporcionada o inválida' };
		}
		const cedulaLimpia = cedula.trim();
		console.log(`🔍 Buscando entrada previa para cédula: ${cedulaLimpia}`);
		const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REGISTRO);
		if (!sheet) {
			console.error(`❌ Hoja ${SHEET_REGISTRO} no encontrada`);
			return { found: false, error: 'Hoja de registro no encontrada' };
		}
		const data = sheet.getDataRange().getValues();
		if (data.length <= 1) {
			return { found: false, message: 'No hay registros' };
		}
		// Obtener estructura de columnas dinámicas
		const COLS = getColumnasRegistro();
		// Buscar la última entrada sin salida correspondiente
		for (let i = data.length - 1; i >= 1; i--) {
			const fila = data[i];
			const cedulaRegistro = String(fila[COLS.CEDULA]).trim();
			const estadoAcceso = String(fila[COLS.ESTADO_ACCESO]).trim();
			const entrada = fila[COLS.ENTRADA];
			const salida = fila[COLS.SALIDA];
			if (cedulaRegistro === cedulaLimpia && 
					(estadoAcceso === ESTADOS_ACCESO.PERMITIDO || estadoAcceso === ESTADOS_ACCESO.TEMPORAL) && 
					entrada && !salida) {
				console.log(`✅ Entrada previa encontrada: ${entrada}`);
				return {
					found: true,
					horaEntrada: entrada,
					timestamp: new Date(`${fila[COLS.FECHA]}T${entrada}`),
					fecha: fila[COLS.FECHA],
					nombre: fila[COLS.NOMBRE],
					empresa: fila[COLS.EMPRESA]
				};
			}
		}
		console.log(`⚠️ No se encontró entrada previa para ${cedulaLimpia}`);
		return { found: false, message: 'Sin entrada previa registrada' };
	} catch (error) {
		console.error(`❌ Error buscando entrada previa: ${error.message}`);
		return { found: false, error: error.message };
	}
}
/**
 * Obtiene información completa de una persona desde la hoja Personal
 */
function obtenerInformacionPersona(cedula) {
	try {
		// Validar parámetro obligatorio
		if (!cedula || typeof cedula !== 'string' || cedula.trim() === '') {
			console.warn(`⚠️ obtenerInformacionPersona: Cédula inválida: ${cedula}`);
			return { found: false, error: 'Cédula no proporcionada o inválida' };
		}
		const cedulaLimpia = cedula.trim();
		console.log(`👤 Obteniendo información de persona: ${cedulaLimpia}`);
		const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BD);
		if (!sheet) {
			console.warn(`⚠️ Hoja ${SHEET_BD} no encontrada`);
			return { found: false, message: 'Hoja Base de Datos no encontrada' };
		}
		const data = sheet.getDataRange().getValues();
		if (data.length <= 1) {
			return { found: false, message: 'No hay datos de personal' };
		}
		// Obtener estructura de columnas dinámicas
		const COLS = getColumnasPersonal();
		// Buscar la persona
		for (let i = 1; i < data.length; i++) {
			const fila = data[i];
			const cedulaPersona = String(fila[COLS.CEDULA]).trim();
			if (cedulaPersona === cedulaLimpia) {
				return {
					found: true,
					cedula: cedulaPersona,
					nombre: fila[COLS.NOMBRE],
					empresa: fila[COLS.EMPRESA],
					area: fila[COLS.AREA],
					cargo: fila[COLS.CARGO],
					estado: fila[COLS.ESTADO]
				};
			}
		}
		return { found: false, message: 'Persona no encontrada en hoja Personal' };
	} catch (error) {
		console.error(`❌ Error obteniendo información de persona: ${error.message}`);
		return { found: false, error: error.message };
	}
}
/**
 * Fallback: obtener información de la BD original
 */
function obtenerInformacionPersonaDB(cedula) {
	try {
		if (cacheUsuarios[cedula]) {
			return {
				found: true,
				cedula: cedula,
				nombre: cacheUsuarios[cedula].nombre,
				empresa: cacheUsuarios[cedula].empresa,
				area: cacheUsuarios[cedula].area || 'N/A',
				cargo: cacheUsuarios[cedula].cargo || 'N/A',
				estado: 'Activo'
			};
		}
		return { found: false, message: 'Persona no encontrada en BD' };
	} catch (error) {
		return { found: false, error: error.message };
	}
}
/**
 * Función de diagnóstico para verificar la estructura de las hojas
 * Útil para debugging cuando hay problemas de mapeo de columnas
 */
function diagnosticarEstructuraHojas() {
	try {
		console.log('🔍 === DIAGNÓSTICO DE ESTRUCTURA DE HOJAS ===');
		// Reinicializar estructuras
		inicializarEstructurasColumnas();
		const hojas = [SHEET_REGISTRO, SHEET_BD, SHEET_LOGS];
		hojas.forEach(nombreHoja => {
			try {
				const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
				if (!sheet) {
					console.warn(`⚠️ Hoja "${nombreHoja}" no encontrada`);
					return;
				}
				const encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
				console.log(`\n📊 HOJA: ${nombreHoja}`);
				console.log(`📋 ENCABEZADOS ACTUALES:`);
				encabezados.forEach((encabezado, indice) => {
					console.log(`   Columna ${indice}: "${encabezado}"`);
				});
				// Mostrar mapeo detectado si es una hoja principal
				if (nombreHoja === SHEET_REGISTRO) {
					console.log(`🔗 MAPEO DETECTADO (REGISTRO):`, getColumnasRegistro());
				} else if (nombreHoja === SHEET_BD) {
					console.log(`🔗 MAPEO DETECTADO (PERSONAL):`, getColumnasPersonal());
				} else if (nombreHoja === SHEET_LOGS) {
					console.log(`🔗 MAPEO DETECTADO (LOGS):`, getColumnasLogs());
				}
			} catch (error) {
				console.error(`❌ Error procesando hoja ${nombreHoja}:`, error);
			}
		});
		console.log('\n✅ Diagnóstico completado');
		return true;
	} catch (error) {
		console.error('❌ Error en diagnóstico:', error);
		return false;
	}
}
/**
 * Función para limpiar cache y optimizar rendimiento
 * Se puede llamar manualmente cuando sea necesario
 */
function limpiarCacheRendimiento() {
	try {
		SurPassCache.limpiarTodo();
		// Forzar garbage collection si está disponible
		if (typeof Utilities !== 'undefined' && Utilities.sleep) {
			Utilities.sleep(100);
		}
		console.log('✅ Cache de rendimiento limpiado exitosamente');
		return { success: true, message: 'Cache limpiado correctamente' };
	} catch (error) {
		console.error('❌ Error limpiando cache:', error);
		return { success: false, error: error.message };
	}
}
/**
 * Función de diagnóstico de rendimiento
 * Muestra información sobre el estado del cache y rendimiento
 */
function diagnosticoRendimiento() {
	try {
		const ahora = new Date().getTime();
		const diagnostico = {
			timestamp: new Date().toISOString(),
			cache: {
				columnasRegistro: {
					cacheado: !!SurPassCache.columnasRegistro,
					tiempoCache: SurPassCache.tiempoColumnas,
					edadCache: SurPassCache.tiempoColumnas ? (ahora - SurPassCache.tiempoColumnas) : null
				},
				estadisticas: {
					cacheado: !!SurPassCache.ultimasEstadisticas,
					tiempoCache: SurPassCache.tiempoEstadisticas,
					edadCache: SurPassCache.tiempoEstadisticas ? (ahora - SurPassCache.tiempoEstadisticas) : null
				}
			},
			hojas: {},
			recomendaciones: []
		};
		// Verificar estado de las hojas principales
		const ss = SpreadsheetApp.getActiveSpreadsheet();
		const hojasImportantes = [SHEET_REGISTRO, SHEET_BD, SHEET_HISTORIAL];
		hojasImportantes.forEach(nombreHoja => {
			try {
				const hoja = ss.getSheetByName(nombreHoja);
				if (hoja) {
					diagnostico.hojas[nombreHoja] = {
						filas: hoja.getLastRow(),
						columnas: hoja.getLastColumn(),
						existe: true
					};
				} else {
					diagnostico.hojas[nombreHoja] = { existe: false };
				}
			} catch (e) {
				diagnostico.hojas[nombreHoja] = { existe: false, error: e.message };
			}
		});
		// Generar recomendaciones
		if (diagnostico.hojas[SHEET_REGISTRO]?.filas > 1000) {
			diagnostico.recomendaciones.push('Considere archivar registros antiguos para mejorar el rendimiento');
		}
		if (!SurPassCache.columnasRegistro) {
			diagnostico.recomendaciones.push('El cache de columnas no está inicializado');
		}
		console.log('📊 Diagnóstico de rendimiento:', diagnostico);
		return diagnostico;
	} catch (error) {
		console.error('❌ Error en diagnóstico de rendimiento:', error);
		return { error: error.message };
	}
}
/**
 * Función para pre-cargar datos críticos y mejorar rendimiento
 * Se puede ejecutar al inicio del día o periódicamente
 */
function precargarDatosCriticos() {
	try {
		console.log('🚀 Precargando datos críticos para optimizar rendimiento...');
		// Precargar estructura de columnas
		SurPassCache.obtenerColumnasRegistroCache();
		// Precargar estadísticas si no están en cache
		const ahora = new Date().getTime();
		if (!SurPassCache.ultimasEstadisticas || !SurPassCache.tiempoEstadisticas ||
				(ahora - SurPassCache.tiempoEstadisticas) > SurPassCache.DURACION_CACHE_ESTADISTICAS) {
			const estadisticas = obtenerEstadisticas();
			if (estadisticas.optimizado) {
				SurPassCache.ultimasEstadisticas = estadisticas;
				SurPassCache.tiempoEstadisticas = ahora;
			}
		}
		console.log('✅ Datos críticos precargados exitosamente');
		return { success: true, message: 'Datos críticos precargados' };
	} catch (error) {
		console.error('❌ Error precargando datos críticos:', error);
		return { success: false, error: error.message };
	}
}
/**
 * Aplica formato condicional a la hoja de registro para colorear entradas temporales
 */
function aplicarFormatoCondicional() {
	try {
		console.log('🎨 Aplicando formato condicional para entradas temporales...');
		const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
		const COLS = getColumnasRegistro();
		// Obtener el rango donde aplicar el formato
		const lastRow = hojaRegistro.getLastRow();
		if (lastRow <= 1) return true; // No hay datos para aplicar formato
		const rango = hojaRegistro.getRange(2, COLS.ESTADO_ACCESO + 1, lastRow - 1, 1);
		// Eliminar reglas existentes para esta columna
		const reglas = hojaRegistro.getConditionalFormatRules();
		const nuevasReglas = [];
		// Mantener reglas que no afectan a la columna de estado
		for (let i = 0; i < reglas.length; i++) {
			const regla = reglas[i];
			const rangosRegla = regla.getRanges();
			let mantener = true;
			for (let j = 0; j < rangosRegla.length; j++) {
				const rangoRegla = rangosRegla[j];
				if (rangoRegla.getColumn() === COLS.ESTADO_ACCESO + 1) {
					mantener = false;
					break;
				}
			}
			if (mantener) {
				nuevasReglas.push(regla);
			}
		}
		// Agregar nueva regla para estado TEMPORAL (naranja)
		const reglaTemporales = SpreadsheetApp.newConditionalFormatRule()
			.whenTextEqualTo(ESTADOS_ACCESO.TEMPORAL)
			.setBackground('#FFA500') // Naranja
			.setFontColor('#FFFFFF')
			.setRanges([rango])
			.build();
		nuevasReglas.push(reglaTemporales);
		// Agregar regla para estado PERMITIDO (verde)
		const reglaPermitidos = SpreadsheetApp.newConditionalFormatRule()
			.whenTextEqualTo(ESTADOS_ACCESO.PERMITIDO)
			.setBackground('#00A651') // Verde
			.setFontColor('#FFFFFF')
			.setRanges([rango])
			.build();
		nuevasReglas.push(reglaPermitidos);
		// Aplicar las reglas
		hojaRegistro.setConditionalFormatRules(nuevasReglas);
		console.log('✅ Formato condicional aplicado correctamente');
		return true;
	} catch (e) {
		console.error('❌ Error aplicando formato condicional:', e);
		return false;
	}
}
function sincronizarEntradasActivas() {
  try {
    const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
    if (!hojaRegistro) {
      console.error(`Hoja ${SHEET_REGISTRO} no encontrada`);
      return [];
    }

    const lastRow = hojaRegistro.getLastRow();
    if (lastRow <= 1) {
      return [];
    }

    // Lectura optimizada: últimos 300 registros
    const maxRows = Math.min(lastRow - 1, 300);
    const startRow = Math.max(2, lastRow - maxRows + 1);
    const datos = hojaRegistro.getRange(startRow, 1, maxRows, hojaRegistro.getLastColumn()).getValues();
    const encabezados = hojaRegistro.getRange(1, 1, 1, hojaRegistro.getLastColumn()).getValues()[0];

    // Mapeo flexible de columnas específico para esta función
    const mapearIndiceColumna = (posiblesNombres) => {
      for (const nombre of posiblesNombres) {
        const indice = encabezados.findIndex(h => 
          String(h).toLowerCase().trim().includes(nombre.toLowerCase())
        );
        if (indice >= 0) return indice;
      }
      return -1;
    };

    const colCedula = mapearIndiceColumna(['cédula', 'cedula']);
    const colNombre = mapearIndiceColumna(['nombre', 'name']);
    const colEmpresa = mapearIndiceColumna(['empresa', 'company']);
    const colHoraEntrada = mapearIndiceColumna(['hora entrada', 'entrada', 'entry']);
    const colHoraSalida = mapearIndiceColumna(['hora salida', 'salida', 'exit']);
    const colTipoAcceso = mapearIndiceColumna(['tipo acceso', 'estado acceso', 'tipo', 'estado']);

    // Validar que se encontraron las columnas críticas
    if (colCedula < 0) {
      console.error('No se encontró columna de Cédula. Encabezados:', encabezados);
      return [];
    }

    const entradasActivas = [];

    for (let i = 0; i < datos.length; i++) {
      const fila = datos[i];
      
      // Validar que la fila tenga datos en la columna de cédula
      if (!fila[colCedula]) continue;

      const cedula = String(fila[colCedula]).trim();
      if (!cedula) continue;

      const tieneEntrada = fila[colHoraEntrada] && String(fila[colHoraEntrada]).trim() !== '';
      const tieneSalida = colHoraSalida >= 0 && fila[colHoraSalida] && String(fila[colHoraSalida]).trim() !== '';

      // Solo incluir entradas sin salida
      if (tieneEntrada && !tieneSalida) {
        entradasActivas.push({
          fila: startRow + i,
          cedula: cedula,
          nombre: colNombre >= 0 ? (fila[colNombre] || '') : '',
          empresa: colEmpresa >= 0 ? (fila[colEmpresa] || '') : '',
          horaEntrada: fila[colHoraEntrada]
        });
      }
    }

    console.log(`Entradas activas optimizadas: ${entradasActivas.length} (procesados ${maxRows} registros)`);
    return entradasActivas;

  } catch (e) {
    console.error('Error en sincronizarEntradasActivas:', e.message);
    return [];
  }
}

/**
 * Función auxiliar para actualizar entradas activas después de una operación
 * Esta función debe llamarse después de cada entrada o salida para mantener
 * la sincronización con el frontend.
 * 
 * @return {Array} Lista actualizada de entradas activas
 */
function actualizarEntradasActivas() {
	return sincronizarEntradasActivas();
}
// Normaliza una hora a formato HH:MM (acepta 6:00 a.m., 18:00, etc.)
function normalizarHora(valor) {
	if (!valor) return '';
	let str = String(valor).trim().toLowerCase();
	// Reemplazar puntos, espacios y am/pm
	str = str.replace(/\./g, ':').replace(/\s+/g, '');
	// Convertir 6:00am o 6:00a.m. a 06:00
	let match = str.match(/^(\d{1,2}):(\d{2})(?::\d{2})?(am|pm|a\.m\.|p\.m\.|a\.m|p\.m)?$/);
	if (match) {
		let h = parseInt(match[1], 10);
		let m = match[2];
		let suf = match[3] || '';
		if (suf.includes('p') && h < 12) h += 12;
		if (suf.includes('a') && h === 12) h = 0;
		return `${h.toString().padStart(2, '0')}:${m}`;
	}
	// Si ya está en formato HH:MM
	if (/^\d{2}:\d{2}$/.test(str)) return str;
	// Si es solo hora
	if (/^\d{1,2}$/.test(str)) return str.padStart(2, '0') + ':00';
	return valor;
}

// Normaliza el valor de MODO_24_HORAS a booleano
function normalizarModo24h(valor) {
	if (typeof valor === 'boolean') return valor;
	if (!valor) return false;
	const str = String(valor).trim().toLowerCase();
	return str === 'true' || str === 'si' || str === '1';
}

// Normaliza los días laborables a abreviaturas estándar (LUN,MAR,MIÉ,JUE,VIE)
function normalizarDiasLaborables(valor) {
	if (!valor) return '';
	let dias = String(valor).toUpperCase()
		.replace(/Á/g, 'A').replace(/É/g, 'E').replace(/Í/g, 'I').replace(/Ó/g, 'O').replace(/Ú/g, 'U')
		.replace(/\s+/g, '')
		.split(',');
	const mapa = {
		'LUNES': 'LUN', 'LUN': 'LUN',
		'MARTES': 'MAR', 'MAR': 'MAR',
		'MIERCOLES': 'MIE', 'MIÉRCOLES': 'MIE', 'MIE': 'MIE',
		'JUEVES': 'JUE', 'JUE': 'JUE',
		'VIERNES': 'VIE', 'VIE': 'VIE',
		'SABADO': 'SAB', 'SÁBADO': 'SAB', 'SAB': 'SAB',
		'DOMINGO': 'DOM', 'DOM': 'DOM'
	};
	return dias.map(d => mapa[d] || d).filter(Boolean).join(',');
}
// Devuelve el horario laboral desde la configuración

function obtenerHorarioLaboral() {
	const result = obtenerConfiguracion();
	if (!result.success) {
		console.warn('No se pudo obtener la configuración:', result.message);
		return { apertura: '', cierre: '', modo24h: false };
	}
	const config = result.config;
	const aperturaNorm = normalizarHora(config.HORARIO_APERTURA);
	const cierreNorm = normalizarHora(config.HORARIO_CIERRE);
	const modo24hNorm = normalizarModo24h(config.MODO_24_HORAS);
	// Registrar en hoja de logs
	try {
		registrarErrorLog('DEBUG_HORARIO', 'Crudo', {
			HORARIO_APERTURA: config.HORARIO_APERTURA,
			HORARIO_CIERRE: config.HORARIO_CIERRE,
			MODO_24_HORAS: config.MODO_24_HORAS
		});
		registrarErrorLog('DEBUG_HORARIO', 'Normalizado', {
			apertura: aperturaNorm,
			cierre: cierreNorm,
			modo24h: modo24hNorm
		});
	} catch (e) {
		// Ignorar errores de log para no interrumpir flujo
	}
	return {
		apertura: aperturaNorm,
		cierre: cierreNorm,
		modo24h: modo24hNorm
	};
}

// Devuelve los días laborables desde la configuración
function obtenerDiasLaborables() {
	const result = obtenerConfiguracion();
	if (!result.success) {
		console.warn('No se pudo obtener la configuración:', result.message);
		return '';
	}
	const config = result.config;
	const diasNorm = normalizarDiasLaborables(config.DIAS_LABORABLES);
	try {
		registrarErrorLog('DEBUG_DIAS', 'Crudo', { DIAS_LABORABLES: config.DIAS_LABORABLES });
		registrarErrorLog('DEBUG_DIAS', 'Normalizado', { diasNorm });
	} catch (e) {
		// Ignorar errores de log para no interrumpir flujo
	}
	return diasNorm;
}
function inicializarCacheUsuarios() {
	try {
		// Si el cache ya está inicializado, no volver a cargarlo
		if (cacheUsuarios && Object.keys(cacheUsuarios).length > 0) {
			return; // Ya está cargado, salir para mejorar rendimiento
		}

		// Verificar y crear estructura de hojas si es necesario
		verificarYCrearEstructuraHojas();
		// Inicializar estructuras de columnas dinámicas
		inicializarEstructurasColumnas();

		const hoja = SpreadsheetApp.getActive().getSheetByName(SHEET_BD);
		const datos = hoja.getDataRange().getValues();
		const encabezados = datos[0];
		cacheUsuarios = {};

		for (let i = 1; i < datos.length; i++) {
			const fila = datos[i];
			const cedula = String(fila[0]).trim();
			if (!cedula) continue;

			let usuario = {};
			for (let idx = 0; idx < encabezados.length; idx++) {
				usuario[encabezados[idx]] = fila[idx];
			}

			usuario.cedula = cedula;
			usuario.nombre = usuario.Nombre || usuario.nombre || '';
			usuario.empresa = usuario.Empresa || usuario.empresa || '';
			usuario.busqueda = `${cedula} ${usuario.nombre} ${usuario.empresa}`.toLowerCase();
			usuario.cedulaNormalizada = cedula.replace(/[^\d]/g, '');

			cacheUsuarios[cedula] = usuario;
		}
		console.log(`Cache inicializado con ${Object.keys(cacheUsuarios).length} usuarios`);
	} catch (e) {
		console.error('Error inicializando cache:', e);
		registrarErrorLog('ERROR_CACHE', e.message, {});
	}
}
function obtenerOpcionesMenu() {
	try {
		const hojaConfig = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
		const datos = hojaConfig.getDataRange().getValues();

		const opciones = {
			mostrarEstadisticas: true,
			usarJsQR: false,
			sonidosActivados: true
		};

		// Leer configuraciones existentes
		for (let i = 1; i < datos.length; i++) {
			const key = datos[i][0];
			const value = datos[i][1];
			if (key === 'mostrarEstadisticas') {
				opciones.mostrarEstadisticas = Boolean(value);
			} else if (key === 'usarJsQR') {
				opciones.usarJsQR = Boolean(value);
			} else if (key === 'sonidosActivados') {
				opciones.sonidosActivados = Boolean(value);
			}
		}

		console.log('⚙️ Opciones de menú cargadas:', opciones);
		return opciones;

	} catch (e) {
		console.error('Error obteniendo opciones del menú:', e);
		registrarLog('ERROR', 'Error obteniendo opciones del menú', {
			error: e.message,
			stack: e.stack
		});
		return {
			mostrarEstadisticas: true,
			usarJsQR: false,
			sonidosActivados: true
		};
	}
}

/**
 * Actualiza o agrega una configuración individual en la hoja de configuración.
 * @param {string} key - Clave de la configuración
 * @param {any} value - Valor a establecer
 */
function actualizarConfiguracion(key, value) {
		try {
			const hojaConfig = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
			const datos = hojaConfig.getDataRange().getValues();
			// Buscar si ya existe la configuración
			for (let i = 1; i < datos.length; i++) {
				if (datos[i][0] === key) {
					hojaConfig.getRange(i + 1, 2).setValue(value);
					return;
				}
			}
			// Si no existe, agregarla
			hojaConfig.appendRow([key, value]);
		} catch (e) {
			console.error('Error actualizando configuración:', e);
			throw e;
		}
}

function obtenerTodoElPersonal() {
	try {
		inicializarCacheUsuarios();
		const personal = Object.values(cacheUsuarios);
		console.log(`📋 Enviando ${personal.length} registros de personal al frontend`);
		return personal;
	} catch (e) {
		console.error('Error obteniendo personal:', e);
		registrarErrorLog('ERROR_OBTENER_PERSONAL', e.message, {});
		return [];
	}
}
// Permite incluir archivos .html en plantillas (para CSS, JS, etc.)
function include(filename) {
	if (!filename || typeof filename !== 'string' || filename.trim() === '') {
		throw new Error('El parámetro filename es inválido o está vacío en include().');
	}
	try {
		return HtmlService.createHtmlOutputFromFile(filename).getContent();
	} catch (e) {
		throw new Error(`No se pudo incluir el archivo HTML "${filename}.html": ${e.message}`);
	}
}
// Funciones web básicas para Google Apps Script
function doGet(e) {
  // Cargar la interfaz de administración si el parámetro 'page=admin' está presente.
  if (e.parameter && e.parameter.page === 'admin') {
    return HtmlService.createTemplateFromFile('admin.html').evaluate();
  }
  
  // De lo contrario, cargar la interfaz de registro por defecto.
  return HtmlService.createTemplateFromFile('registro.html').evaluate();
}

function doPost(e) {
	// Aquí puedes manejar peticiones POST si lo necesitas
	return ContentService.createTextOutput('Método POST recibido');
}

/**
 * Registra un evento o error en la hoja de logs
 * @param {string} tipo - Nivel o tipo de log (INFO, ERROR, DEBUG, etc)
 * @param {string} mensaje - Mensaje descriptivo
 * @param {Object} datos - Datos adicionales (opcional)
 */
function registrarLog(tipo, mensaje, datos) {
	try {
		const hojaLogs = SpreadsheetApp.getActive().getSheetByName(SHEET_LOGS);
		if (!hojaLogs) throw new Error('Hoja de logs no encontrada');
		hojaLogs.appendRow([
			new Date(),
			tipo || 'INFO',
			mensaje || '',
			datos ? JSON.stringify(datos) : ''
		]);
	} catch (e) {
		// Fallback: loguear en consola si falla el registro en hoja
		console.error('Error en registrarLog:', e, { tipo, mensaje, datos });
	}
}

/**
 * Alterna el estado de los sonidos
 * @param {boolean} estado
 * @return {Object} Resultado
 */
function toggleSonidos(estado) {
	try {
		actualizarConfiguracion('sonidosActivados', estado);
		console.log(`🔊 Sonidos ${estado ? 'activados' : 'desactivados'}`);
		return {
			success: true,
			sonidosActivados: estado,
			message: `Sonidos ${estado ? 'activados' : 'desactivados'}`
		};
	} catch (e) {
		console.error('Error toggling sonidos:', e);
		registrarLog('ERROR', 'Error toggling sonidos', {
			error: e.message,
			stack: e.stack,
			estado
		});
		return {
			success: false,
			code: 'TOGGLE_SONIDOS_FAILED',
			message: 'Error al cambiar configuración de sonidos: ' + e.message,
			details: { error: e.message, stack: e.stack, estado }
		};
	}
}

/**
   * Alterna la visibilidad del panel de estadísticas
   * @param {boolean} estado
   * @return {Object} Resultado
   */
  function toggleEstadisticas(estado) {
    try {
      actualizarConfiguracion('mostrarEstadisticas', estado);
      console.log(`📊 Panel de estadísticas ${estado ? 'activado' : 'desactivado'}`);
      return {
        success: true,
        mostrarEstadisticas: estado,
        message: `Panel de estadísticas ${estado ? 'activado' : 'desactivado'}`
      };
    } catch (e) {
      console.error('Error toggling estadísticas:', e);
      registrarLog('ERROR', 'Error toggling estadísticas', {
        error: e.message,
        stack: e.stack,
        estado
      });
      return {
        success: false,
        code: 'TOGGLE_ESTADISTICAS_FAILED',
        message: 'Error al cambiar visibilidad de estadísticas: ' + e.message,
        details: { error: e.message, stack: e.stack, estado }
      };
    }
  }

/**
 * Muestra información general del sistema
 * @return {string} HTML con información del sistema
 */
function mostrarInformacionSistema() {
	try {
		console.log('ℹ️ Generando información del sistema...');
		const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
		const hojaHistorial = SpreadsheetApp.getActive().getSheetByName(SHEET_HISTORIAL);
		const hojaBD = SpreadsheetApp.getActive().getSheetByName(SHEET_BD);
		const hojaLogs = SpreadsheetApp.getActive().getSheetByName(SHEET_LOGS);
		const registrosActuales = Math.max(0, hojaRegistro.getLastRow() - 1);
		const totalHistorial = Math.max(0, hojaHistorial.getLastRow() - 1);
		const totalPersonal = Math.max(0, hojaBD.getLastRow() - 1);
		const totalLogs = Math.max(0, hojaLogs.getLastRow() - 1);
		// Obtener estadísticas actuales
		const stats = obtenerEstadisticas();
		const ahora = new Date();
		let html = `
			<div style="padding: 20px; font-family: 'Inter', sans-serif;">
				<h2 style="color: #1e293b; margin-bottom: 30px;">
					<i class="fas fa-info-circle"></i> Información del Sistema SurPass
				</h2>
				<!-- Estadísticas Generales -->
				<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px;">
					<div style="background: linear-gradient(135deg, #3b82f6, #1d4ed8); color: white; padding: 20px; border-radius: 12px; text-align: center;">
						<i class="fas fa-users" style="font-size: 2rem; margin-bottom: 10px;"></i>
						<h3 style="margin: 0; font-size: 2rem;">${totalPersonal}</h3>
						<p style="margin: 5px 0 0 0; opacity: 0.9;">Personal Registrado</p>
					</div>
					<div style="background: linear-gradient(135deg, #10b981, #059669); color: white; padding: 20px; border-radius: 12px; text-align: center;">
						<i class="fas fa-sign-in-alt" style="font-size: 2rem; margin-bottom: 10px;"></i>
						<h3 style="margin: 0; font-size: 2rem;">${stats.entradas}</h3>
						<p style="margin: 5px 0 0 0; opacity: 0.9;">Entradas Hoy</p>
					</div>
					<div style="background: linear-gradient(135deg, #f59e0b, #d97706); color: white; padding: 20px; border-radius: 12px; text-align: center;">
						<i class="fas fa-sign-out-alt" style="font-size: 2rem; margin-bottom: 10px;"></i>
						<h3 style="margin: 0; font-size: 2rem;">${stats.salidas}</h3>
						<p style="margin: 5px 0 0 0; opacity: 0.9;">Salidas Hoy</p>
					</div>
					<div style="background: linear-gradient(135deg, #8b5cf6, #7c3aed); color: white; padding: 20px; border-radius: 12px; text-align: center;">
						<i class="fas fa-building" style="font-size: 2rem; margin-bottom: 10px;"></i>
						<h3 style="margin: 0; font-size: 2rem;">${stats.personasDentro}</h3>
						<p style="margin: 5px 0 0 0; opacity: 0.9;">Personas Dentro</p>
					</div>
				</div>
				<!-- Información Técnica -->
				<div style="background: white; border: 1px solid #e2e8f0; border-radius: 12px; padding: 20px; margin-bottom: 20px;">
					<h3 style="color: #1e293b; margin-bottom: 15px;">
						<i class="fas fa-database"></i> Estado de la Base de Datos
					</h3>
					<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px;">
						<div>
							<strong>Registros Turno Actual:</strong> ${registrosActuales}
						</div>
						<div>
							<strong>Total Historial:</strong> ${totalHistorial}
						</div>
						<div>
							<strong>Logs del Sistema:</strong> ${totalLogs}
						</div>
						<div>
							<strong>Última Actualización:</strong> ${ahora.toLocaleString('es-PA')}
						</div>
					</div>
				</div>
				<!-- Alertas y Advertencias -->
				${stats.salidasSinEntrada > 0 ? `
				<div style="background: #fef2f2; border: 1px solid #fecaca; border-radius: 12px; padding: 20px; margin-bottom: 20px;">
					<h3 style="color: #dc2626; margin-bottom: 10px;">
						<i class="fas fa-exclamation-triangle"></i> Advertencias del Sistema
					</h3>
					<p style="color: #991b1b; margin: 0;">
						<strong>Salidas sin Entrada:</strong> ${stats.salidasSinEntrada} registro(s) requieren revisión.<br>Estas salidas no tienen un registro de entrada correspondiente.
					</p>
				</div>
				` : ''}
				<!-- Información del Sistema -->
				<div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 20px;">
					<h3 style="color: #1e293b; margin-bottom: 15px;">
						<i class="fas fa-cog"></i> Información del Sistema
					</h3>
					<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 15px; color: #64748b;">
						<div>
							<strong>Versión:</strong> SurPass v3.0
						</div>
						<div>
							<strong>Plataforma:</strong> Google Apps Script
						</div>
						<div>
							<strong>Última Sincronización:</strong> ${ahora.toLocaleTimeString('es-PA')}
						</div>
						<div>
							<strong>Estado del Servidor:</strong> <span style="color: #10b981;">✅ En línea</span>
						</div>
					</div>
				</div>
			</div>
		`;
		console.log('ℹ️ Información del sistema generada');
		return html;
	} catch (e) {
		console.error('Error generando información del sistema:', e);
		registrarLog('ERROR', 'Error generando información del sistema', {
			error: e.message,
			stack: e.stack
		});
		return `
			<div style="padding: 20px; text-align: center; color: #ef4444;">
				<h3>Error cargando información</h3>
				<p>No se pudo cargar la información del sistema.</p>
			</div>
		`;
	}
}

/**
 * Genera los controles de paginación para la vista previa
 * @param {number} paginaActual - Página actual
 * @param {number} totalPaginas - Total de páginas
 * @param {string} cedula - Cédula para filtrar, o vacío para todos
 * @param {number} registrosPorPagina - Registros por página
 * @return {string} HTML con los controles de paginación
 */
function generarControlPaginacion(paginaActual, totalPaginas, cedula, registrosPorPagina) {
	// Si solo hay una página, no mostrar controles
	if (totalPaginas <= 1) {
		return '';
	}
	let html = `
		<div style="display: flex; justify-content: center; margin-top: 20px; user-select: none;">
			<div class="pagination" style="display: flex; align-items: center; background: white; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden;">
	`;
	// Botón anterior
	const prevDisabled = paginaActual <= 1;
	html += `
		<button 
			onclick="${prevDisabled ? '' : `cambiarPagina(${paginaActual - 1}, '${cedula || ''}', ${registrosPorPagina})`}" 
			style="
				padding: 10px 15px; 
				border: none; 
				background: ${prevDisabled ? '#f1f5f9' : 'white'}; 
				color: ${prevDisabled ? '#94a3b8' : '#1e40af'};
				cursor: ${prevDisabled ? 'default' : 'pointer'};
				font-size: 14px;
				display: flex;
				align-items: center;
			"
			${prevDisabled ? 'disabled' : ''}
		>
			<i class="fas fa-chevron-left" style="margin-right: 5px;"></i> Anterior
		</button>
	`;
	// Indicador de página
	html += `
		<div style="
			padding: 10px 15px;
			border-left: 1px solid #e2e8f0;
			border-right: 1px solid #e2e8f0;
			color: #334155;
			font-size: 14px;
		">
			Página ${paginaActual} de ${totalPaginas}
		</div>
	`;
	// Botón siguiente
	const nextDisabled = paginaActual >= totalPaginas;
	html += `
		<button 
			onclick="${nextDisabled ? '' : `cambiarPagina(${paginaActual + 1}, '${cedula || ''}', ${registrosPorPagina})`}" 
			style="
				padding: 10px 15px; 
				border: none; 
				background: ${nextDisabled ? '#f1f5f9' : 'white'}; 
				color: ${nextDisabled ? '#94a3b8' : '#1e40af'};
				cursor: ${nextDisabled ? 'default' : 'pointer'};
				font-size: 14px;
				display: flex;
				align-items: center;
			"
			${nextDisabled ? 'disabled' : ''}
		>
			Siguiente <i class="fas fa-chevron-right" style="margin-left: 5px;"></i>
		</button>
	`;
	html += `
			</div>
		</div>
	`;
	return html;
}

/**
 * Finaliza el turno, archiva los registros actuales en el historial y limpia la hoja de registro
 */
function finalizarTurno() {
	try {
		console.log('🔚 Finalizando turno...');
		const COLS = getColumnasRegistro();
		const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
		const hojaHistorial = SpreadsheetApp.getActive().getSheetByName(SHEET_HISTORIAL);
		const datos = hojaRegistro.getDataRange().getValues();
		if (datos.length <= 1) {
			return { success: false, message: 'No hay registros para finalizar el turno' };
		}
		for (let i = 1; i < datos.length; i++) {
			const fila = [...datos[i]];
			fila[COLS.COMENTARIO] = (fila[COLS.COMENTARIO] || '') + ' [FIN TURNO]';
			hojaHistorial.appendRow(fila);
		}
		if (hojaRegistro.getLastRow() > 1) {
			hojaRegistro.getRange(2, 1, hojaRegistro.getLastRow() - 1, hojaRegistro.getLastColumn()).clearContent();
		}
		registrarLog('INFO', 'Turno finalizado correctamente', {
			registrosProcesados: datos.length - 1,
			fecha: new Date().toISOString()
		});
		return {
			success: true,
			message: `Turno finalizado correctamente.\n${datos.length - 1} registros han sido archivados en el historial.`
		};
	} catch (e) {
		console.error('Error finalizando turno:', e);
		registrarLog('ERROR', 'Error finalizando turno', { error: e.message });
		return { success: false, message: 'Error al finalizar el turno: ' + e.message };
	}
}

/**
 * Limpia todos los registros actuales de la hoja de registro
 */
function limpiarRegistros() {
	try {
		console.log('🧹 Limpiando registros...');
		const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
		if (hojaRegistro.getLastRow() <= 1) {
			return { exito: false, mensaje: 'No hay registros para limpiar' };
		}
		hojaRegistro.getRange(2, 1, hojaRegistro.getLastRow() - 1, hojaRegistro.getLastColumn()).clearContent();
		registrarErrorLog('LIMPIEZA_REGISTROS', 'Registros limpiados manualmente', { fecha: new Date().toISOString() });
		return { exito: true, mensaje: 'Registros limpiados exitosamente' };
	} catch (e) {
		console.error('Error limpiando registros:', e);
		registrarErrorLog('ERROR_LIMPIAR_REGISTROS', e.message, {});
		return { exito: false, mensaje: 'Error limpiando registros: ' + e.message };
	}
}

/**
 * Muestra una vista previa paginada de los registros del día actual, filtrando por cédula si se indica
 */
function mostrarVistaPreviaRondas(cedula, pagina = 1, registrosPorPagina = 100) {
	try {
		console.log(`👁️ Generando vista previa para: ${cedula || 'TODOS LOS REGISTROS'}, página ${pagina}, ${registrosPorPagina} por página`);
		const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
		if (!hojaRegistro) {
			return `<div style="padding: 20px; text-align: center; color: #ef4444;">
				<h3>Error</h3>
				<p>No se encontró la hoja de Registro.</p>
			</div>`;
		}
		const datos = hojaRegistro.getDataRange().getValues();
		if (datos.length <= 1) {
			return `<div style="padding: 20px; text-align: center; color: #64748b;">
				<h3>Sin registros</h3>
				<p>No hay registros del día actual.</p>
			</div>`;
		}
		const encabezados = datos[0];
		pagina = parseInt(pagina) || 1;
		registrosPorPagina = parseInt(registrosPorPagina) || 100;
		if (pagina < 1) pagina = 1;
		if (registrosPorPagina < 10) registrosPorPagina = 10;
		if (registrosPorPagina > 500) registrosPorPagina = 500;
		let html = `
			<div style="padding: 20px; font-family: 'Inter', sans-serif;">
				<h2 style="color: #1e293b; margin-bottom: 20px;">
					<i class="fas fa-history"></i> Registros del Día Actual
				</h2>
				<p style="color: #64748b; margin-bottom: 30px;">
					${cedula ? `Mostrando registros de hoy para: <strong>${cedula}</strong>` : 'Mostrando todos los registros de hoy'}
				</p>
		`;
		const registrosPersona = [];
		const mostrarTodos = !cedula || cedula.trim() === '';
		for (let i = 1; i < datos.length; i++) {
			const cumpleFiltro = mostrarTodos || String(datos[i][1]).trim() === String(cedula).trim();
			if (cumpleFiltro) {
				const registro = {};
				encabezados.forEach((col, idx) => {
					registro[col] = datos[i][idx];
				});
				registrosPersona.push(registro);
			}
		}
		if (registrosPersona.length === 0) {
			html += `
				<div style="text-align: center; padding: 40px; color: #64748b;">
					<i class="fas fa-search" style="font-size: 3rem; margin-bottom: 20px; opacity: 0.5;"></i>
					<h3>No se encontraron registros</h3>
					<p>No hay registros de hoy ${cedula ? 'para esta cédula' : ''}.</p>
				</div>
			`;
		} else {
			registrosPersona.sort((a, b) => {
				const horaA = a[encabezados[4]] || a[encabezados[5]] || '';
				const horaB = b[encabezados[4]] || b[encabezados[5]] || '';
				return horaB.localeCompare(horaA);
			});
			const totalRegistros = registrosPersona.length;
			const totalPaginas = Math.ceil(totalRegistros / registrosPorPagina);
			if (pagina > totalPaginas) pagina = totalPaginas;
			const inicio = (pagina - 1) * registrosPorPagina;
			const fin = Math.min(inicio + registrosPorPagina, totalRegistros);
			const registrosPagina = registrosPersona.slice(inicio, fin);
			html += `
				<div style="margin-bottom: 20px;">
					<strong>Total de registros de hoy: ${totalRegistros}</strong>
					| Página ${pagina} de ${totalPaginas}
				</div>
				<div style="overflow-x: auto;">
					<table style="width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden;">
						<thead>
							<tr style="background: #3b82f6; color: white;">
								<th style="padding: 12px; text-align: left;">Fecha</th>
								<th style="padding: 12px; text-align: left;">Cédula</th>
								<th style="padding: 12px; text-align: left;">Nombre</th>
								<th style="padding: 12px; text-align: left;">Estado del Acceso</th>
								<th style="padding: 12px; text-align: left;">Entrada</th>
								<th style="padding: 12px; text-align: left;">Salida</th>
								<th style="padding: 12px; text-align: left;">Duración</th>
								<th style="padding: 12px; text-align: left;">Empresa</th>
							</tr>
						</thead>
						<tbody>
			`;
			registrosPagina.forEach((registro, index) => {
				const estadoAcceso = registro[encabezados[3]] || '';
				let colorEstado = '#64748b';
				if (estadoAcceso.includes('Permitido')) colorEstado = '#10b981';
				else if (estadoAcceso.includes('Temporal')) colorEstado = '#f59e0b';
				else if (estadoAcceso.includes('Denegado')) colorEstado = '#ef4444';
				html += `
					<tr style="border-bottom: 1px solid #e2e8f0; ${index % 2 === 0 ? 'background: #f8fafc;' : ''}">
						<td style="padding: 12px;">${registro[encabezados[0]] || 'N/A'}</td>
						<td style="padding: 12px; font-family: monospace;">${registro[encabezados[1]] || 'N/A'}</td>
						<td style="padding: 12px; font-weight: 500;">${registro[encabezados[2]] || 'N/A'}</td>
						<td style="padding: 12px; color: ${colorEstado}; font-weight: 600;">${estadoAcceso || 'N/A'}</td>
						<td style="padding: 12px; font-family: monospace;">${registro[encabezados[4]] || '-'}</td>
						<td style="padding: 12px; font-family: monospace;">${registro[encabezados[5]] || '-'}</td>
						<td style="padding: 12px; font-family: monospace; font-weight: 600;">${registro[encabezados[6]] || '-'}</td>
						<td style="padding: 12px;">${registro[encabezados[7]] || 'N/A'}</td>
					</tr>
				`;
			});
			html += `
						</tbody>
					</table>
				</div>
			`;
			if (totalPaginas > 1) {
				html += `
					<div style="text-align: center; margin-top: 20px;">
						<div style="display: inline-flex; gap: 10px; align-items: center;">
				`;
				if (pagina > 1) {
					html += `<button onclick="cambiarPagina(${pagina - 1}, '${cedula}', ${registrosPorPagina})" style="padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer;">Anterior</button>`;
				}
				html += `<span style="margin: 0 15px;">Página ${pagina} de ${totalPaginas}</span>`;
				if (pagina < totalPaginas) {
					html += `<button onclick="cambiarPagina(${pagina + 1}, '${cedula}', ${registrosPorPagina})" style="padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 4px; cursor: pointer;">Siguiente</button>`;
				}
				html += `
						</div>
					</div>
				`;
			}
		}
		html += `
			<script>
				function cambiarPagina(pagina, cedula, registrosPorPagina) {
					document.getElementById('contenidoVistaPrevia').innerHTML = '<div style="text-align: center; padding: 40px;"><div style="display: inline-block; width: 40px; height: 40px; border: 4px solid #f3f3f3; border-top: 4px solid #3498db; border-radius: 50%; animation: spin 1s linear infinite;"></div><p style="margin-top: 15px;">Cargando...</p></div>';
					google.script.run
						.withSuccessHandler(function(html) {
							document.getElementById('contenidoVistaPrevia').innerHTML = html;
						})
						.withFailureHandler(function(error) {
							document.getElementById('contenidoVistaPrevia').innerHTML = '<div style="padding: 20px; text-align: center; color: #ef4444;"><h3>Error</h3><p>No se pudo cargar la página: ' + error + '</p></div>';
						})
						.mostrarVistaPreviaRondas(cedula, pagina, registrosPorPagina);
				}
			</script>
		</div>`;
		console.log(`👁️ Vista previa generada: ${registrosPersona.length} registros del día actual`);
		return html;
	} catch (e) {
		console.error('Error generando vista previa:', e);
		return `
			<div style="padding: 20px; text-align: center; color: #ef4444;">
				<h3>Error cargando registros</h3>
				<p>No se pudo cargar los registros del día: ${e.message}</p>
			</div>
		`;
	}
}

function completarSalida(fila, horaSalida) {
    try {
        if (!fila || fila <= 1) {
            console.error('Número de fila inválido:', fila);
            return { success: false, error: 'Número de fila inválido' };
        }
        
        const hojaRegistro = SpreadsheetApp.getActive().getSheetByName('Registro');
        const hojaHistorial = SpreadsheetApp.getActive().getSheetByName('Historial');
        
        if (!hojaRegistro) {
            return { success: false, error: 'Hoja de registro no encontrada' };
        }
        
        // Obtener datos de la fila
        const datosRegistro = hojaRegistro.getRange(fila, 1, 1, 9).getValues()[0];
        
        const fechaOriginal = datosRegistro[0];
        const cedulaOriginal = datosRegistro[1];
        const nombreOriginal = datosRegistro[2];
        const horaEntradaOriginal = datosRegistro[4]; // Columna E
        
        // Formatear hora de salida
        const horaSalidaFormateada = formatearHora(horaSalida);
        
        // Actualizar hora de salida
        hojaRegistro.getRange(fila, 6).setValue(horaSalidaFormateada); // Columna F
        
        // Calcular duración correctamente
        let duracionCalculada = '0:00:00';
        
        if (horaEntradaOriginal) {
            // Extraer solo la parte de hora si es necesario
            let horaEntradaStr = horaEntradaOriginal;
            let horaSalidaStr = horaSalidaFormateada;
            
            if (typeof horaEntradaOriginal === 'string') {
                horaEntradaStr = horaEntradaOriginal.trim();
            }
            
            duracionCalculada = calcularDuracionSimple(horaEntradaStr, horaSalidaStr);
        }
        
        // Actualizar duración
        hojaRegistro.getRange(fila, 7).setValue(duracionCalculada); // Columna G
        
        // Preparar datos para historial
        const datosHistorial = [...datosRegistro];
        datosHistorial[5] = horaSalidaFormateada;
        datosHistorial[6] = duracionCalculada;
        
        // Agregar al historial
        if (hojaHistorial) {
            hojaHistorial.appendRow(datosHistorial);
        }
        
        console.log(`Salida completada - Entrada: ${horaEntradaOriginal}, Salida: ${horaSalidaFormateada}, Duración: ${duracionCalculada}`);
        
        return { 
            success: true, 
            duracion: duracionCalculada
        };
        
    } catch (error) {
        console.error('Error en completarSalida:', error);
        return { success: false, error: error.message };
    }
}

// Función mejorada para calcular duración
function calcularDuracionMejorada(horaEntrada, horaSalida) {
    try {
        const hoy = new Date();
        const fechaBase = Utilities.formatDate(hoy, 'America/Panama', 'yyyy-MM-dd');
        
        // Parsear hora de entrada
        let entradaDate;
        if (horaEntrada instanceof Date) {
            entradaDate = horaEntrada;
        } else {
            const entradaStr = String(horaEntrada).trim();
            if (entradaStr.includes(':')) {
                const [h, m, s] = entradaStr.split(':').map(n => parseInt(n) || 0);
                entradaDate = new Date(`${fechaBase}T${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`);
            } else {
                return '0:00:00';
            }
        }
        
        // Parsear hora de salida
        let salidaDate;
        if (horaSalida instanceof Date) {
            salidaDate = horaSalida;
        } else {
            const salidaStr = String(horaSalida).trim();
            if (salidaStr.includes(':')) {
                const [h, m, s] = salidaStr.split(':').map(n => parseInt(n) || 0);
                salidaDate = new Date(`${fechaBase}T${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`);
            } else {
                return '0:00:00';
            }
        }
        
        // Calcular diferencia
        let diffMs = salidaDate - entradaDate;
        
        // Si la diferencia es negativa, asumimos que la salida es al día siguiente
        if (diffMs < 0) {
            salidaDate.setDate(salidaDate.getDate() + 1);
            diffMs = salidaDate - entradaDate;
        }
        
        // Convertir a horas, minutos, segundos
        const horas = Math.floor(diffMs / (1000 * 60 * 60));
        const minutos = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
        const segundos = Math.floor((diffMs % (1000 * 60)) / 1000);
        
        return `${horas}:${String(minutos).padStart(2, '0')}:${String(segundos).padStart(2, '0')}`;
        
    } catch (e) {
        console.error('Error en calcularDuracionMejorada:', e);
        return '0:00:00';
    }
}

/**
 * Maneja comentario de persona registrada (justificación de salida sin entrada)
 */
function manejarComentarioPersonaRegistrada(cedula, comentario) {
		try {
			console.log(`💬 Procesando comentario justificación para: ${cedula}`);
			inicializarCacheUsuarios();
			const persona = buscarUsuario(cedula);
			const now = new Date();
			if (persona) {
				// Obtener la estructura actual de columnas de Registro
				const COLS = getColumnasRegistro();
				const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
				// Registrar salida con justificación (ESCENARIO 3 completado)
				const fila = registrarEnHoja(cedula, persona.nombre, persona.empresa, '', now, 'salida', 'web', `JUSTIFICADO: ${comentario}`);
				// Actualizar el estado de acceso a Temporal
				if (fila && fila.filaIndex) {
					hojaRegistro.getRange(fila.filaIndex, COLS.ESTADO_ACCESO + 1).setValue(ESTADOS_ACCESO.TEMPORAL);
					// Aplicar formato condicional para resaltar el estado temporal
					aplicarFormatoCondicional();
				}
				registrarErrorLog('SALIDA_JUSTIFICADA', 'Salida sin entrada con justificación', {
					cedula: cedula,
					nombre: persona.nombre,
					comentario: comentario.substring(0, 100),
					estadoAcceso: ESTADOS_ACCESO.TEMPORAL
				});
				console.log(`✅ ESCENARIO 3 COMPLETADO: Salida justificada para ${persona.nombre} con estado TEMPORAL`);
				return {
					success: true,
					message: 'Salida registrada con justificación exitosamente',
					nombre: persona.nombre,
					empresa: persona.empresa
				};
			} else {
				return {
					success: false,
					message: 'Persona no encontrada en la base de datos'
				};
			}
		} catch (e) {
			console.error('Error procesando comentario de persona registrada:', e);
			registrarErrorLog('ERROR_COMENTARIO_REGISTRADO', e.message, { cedula, comentario });
			return {
				success: false,
				message: 'Error procesando el comentario'
			};
		}
}

/**
 * Procesa registro de persona no registrada (visitante temporal)
 */
function procesarRegistroPersonaNoRegistrada(datos) {
		try {
			// Validar que el objeto datos existe y tiene las propiedades mínimas
			if (!datos || typeof datos !== 'object') {
				console.error(`❌ procesarRegistroPersonaNoRegistrada: Datos inválidos: ${datos}`);
				return {
					success: false,
					color: 'error',
					message: 'Error: Datos de visitante no proporcionados'
				};
			}
			// Validar propiedades obligatorias
			if (!datos.cedula || !datos.nombre || !datos.tipoAcceso) {
				console.error(`❌ procesarRegistroPersonaNoRegistrada: Datos incompletos:`, datos);
				return {
					success: false,
					color: 'error',
					message: 'Error: Datos de visitante incompletos (cédula, nombre y tipo de acceso son obligatorios)'
				};
			}
			console.log(`👤 Procesando visitante:`, datos);
			const now = new Date();
			const comentario = `VISITANTE: ${datos.nombre || 'N/A'} | ${datos.empresa || 'N/A'} | MOTIVO: ${datos.motivo || 'No especificado'}`;
			// Obtener la estructura actual de columnas de Registro
			const COLS = getColumnasRegistro();
			const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
			if (datos.tipoAcceso === 'entrada') {
				// ESCENARIO 4 completado: Entrada de visitante
				const fila = registrarEnHoja(datos.cedula, datos.nombre, datos.empresa || 'Visitante', now, '', 'entrada', 'web', comentario);
				// Actualizar el estado de acceso a Temporal
				if (fila && fila.filaIndex) {
					hojaRegistro.getRange(fila.filaIndex, COLS.ESTADO_ACCESO + 1).setValue(ESTADOS_ACCESO.TEMPORAL);
					// Aplicar formato condicional para resaltar el estado temporal
					aplicarFormatoCondicional();
				}
				console.log(`✅ ESCENARIO 4 COMPLETADO: Entrada de visitante ${datos.nombre} con estado TEMPORAL`);
			} else {
				// ESCENARIO 6 completado: Salida de visitante sin entrada previa
				const fila = registrarEnHoja(datos.cedula, datos.nombre, datos.empresa || 'Visitante', '', now, 'salida', 'web', comentario);
				// Actualizar el estado de acceso a Temporal
				if (fila && fila.filaIndex) {
					hojaRegistro.getRange(fila.filaIndex, COLS.ESTADO_ACCESO + 1).setValue(ESTADOS_ACCESO.TEMPORAL);
					// Aplicar formato condicional para resaltar el estado temporal
					aplicarFormatoCondicional();
				}
				console.log(`✅ ESCENARIO 6 COMPLETADO: Salida de visitante ${datos.nombre} con estado TEMPORAL`);
			}
			registrarLog('INFO', `${datos.tipoAcceso.toUpperCase()} de visitante`, {
				cedula: datos.cedula,
				nombre: datos.nombre,
				empresa: datos.empresa,
				tipoAcceso: datos.tipoAcceso,
				estadoAcceso: ESTADOS_ACCESO.TEMPORAL
			});
			return {
				success: true,
				message: `${datos.tipoAcceso.charAt(0).toUpperCase() + datos.tipoAcceso.slice(1)} de visitante registrada correctamente`
			};
		} catch (e) {
			console.error('Error registrando visitante:', e);
			registrarLog('ERROR', 'Error registro visitante', { error: e.message, datos: datos });
			return {
				success: false,
				message: 'Error registrando visitante'
			};
		}
}

/**
 * ⚡ [RENDIMIENTO] Verifica entrada activa con lectura optimizada por rango de fechas
 */
function verificarEntradaActiva(cedula) {
  try {
    // Validación robusta de parámetro obligatorio
    if (!cedula || typeof cedula !== 'string') {
      console.warn(`Parámetro cedula inválido en verificarEntradaActiva: ${cedula}`);
      return { found: false, error: 'Cédula no proporcionada o inválida' };
    }

    const cedulaLimpia = cedula.trim();
    
    if (cedulaLimpia === '') {
      return { found: false, error: 'Cédula no puede estar vacía' };
    }

    console.log(`Verificando entrada activa para: ${cedulaLimpia}`);
    
    const hojaRegistro = SpreadsheetApp.getActive().getSheetByName(SHEET_REGISTRO);
    if (!hojaRegistro) {
      console.error(`Hoja ${SHEET_REGISTRO} no encontrada`);
      return { found: false, error: 'Hoja de registro no encontrada' };
    }
    
    const COLS = getColumnasRegistro();
    const lastRow = hojaRegistro.getLastRow();
    
    if (lastRow <= 1) {
      return { found: false, message: 'No hay registros' };
    }
    
    // Solo leer los últimos 100 registros del día actual
    const maxFilasABuscar = Math.min(100, lastRow - 1);
    const filaInicio = Math.max(2, lastRow - maxFilasABuscar + 1);
    
    // Solo leer las columnas necesarias
    const numColumnas = Math.max(COLS.CEDULA, COLS.NOMBRE, COLS.EMPRESA, COLS.ENTRADA, COLS.SALIDA, COLS.FECHA) + 1;
    
    const rangoOptimizado = hojaRegistro.getRange(
      filaInicio, 
      1, 
      maxFilasABuscar, 
      numColumnas
    );
    
    const datos = rangoOptimizado.getValues();
    const fechaHoy = new Date().toISOString().split('T')[0];
    
    // Buscar desde el final hacia atrás
    for (let i = datos.length - 1; i >= 0; i--) {
      const fila = datos[i];
      
      // Validar que la fila tenga datos
      if (!fila || !fila[COLS.CEDULA]) {
        continue;
      }
      
      const cedulaRegistro = String(fila[COLS.CEDULA]).trim();
      const fechaRegistro = fila[COLS.FECHA];
      const entrada = fila[COLS.ENTRADA];
      const salida = fila[COLS.SALIDA];
      
      // Solo buscar en registros del día actual
      let fechaRegistroStr;
      try {
        fechaRegistroStr = fechaRegistro instanceof Date ? 
          fechaRegistro.toISOString().split('T')[0] : 
          String(fechaRegistro).split('T')[0]; // Por si viene como string ISO
      } catch (dateError) {
        continue; // Saltar filas con fechas inválidas
      }
      
      if (fechaRegistroStr !== fechaHoy) {
        continue; // Saltar registros de otros días
      }
      
      if (cedulaRegistro === cedulaLimpia && entrada && !salida) {
        let horaEntradaFormateada = 'N/A';
        let timestampEntrada = null;
        
        try {
          if (entrada instanceof Date) {
            horaEntradaFormateada = entrada.toLocaleTimeString('es-PA', {
              hour12: false,
              hour: '2-digit',
              minute: '2-digit',
              second: '2-digit'
            });
            timestampEntrada = entrada;
          } else if (typeof entrada === 'string' && entrada.trim() !== '') {
            horaEntradaFormateada = entrada.trim();
            timestampEntrada = new Date(`${fechaHoy}T${entrada.trim()}`);
            if (isNaN(timestampEntrada.getTime())) {
              timestampEntrada = new Date();
            }
          }
        } catch (timeError) {
          console.warn('Error procesando hora de entrada:', timeError);
          horaEntradaFormateada = 'Hora no disponible';
          timestampEntrada = new Date();
        }

        console.log(`Entrada activa encontrada para ${cedulaLimpia} a las ${horaEntradaFormateada}`);
        
        return {
          found: true,
          fila: filaInicio + i,
          cedula: cedulaRegistro,
          nombre: fila[COLS.NOMBRE] || '',
          empresa: fila[COLS.EMPRESA] || '',
          horaEntrada: horaEntradaFormateada,
          timestamp: timestampEntrada
        };
      }
    }
    
    console.log(`No hay entrada activa para ${cedulaLimpia}`);
    return { found: false, message: 'Sin entrada activa registrada' };
    
  } catch (error) {
    console.error('Error verificando entrada activa:', error);
    return { found: false, error: error.message };
  }
}