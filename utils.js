// utils.js

function formatearHora(hora) {
    try {
        if (!hora) return '';
        
        // Si es un objeto Date
        if (hora instanceof Date) {
            const horas = hora.getHours();
            const minutos = hora.getMinutes();
            const segundos = hora.getSeconds();
            return `${String(horas).padStart(2, '0')}:${String(minutos).padStart(2, '0')}:${String(segundos).padStart(2, '0')}`;
        }
        
        // Si es string
        if (typeof hora === 'string') {
            const horaLimpia = hora.trim();
            
            // Ya tiene formato correcto HH:MM:SS
            if (/^\d{1,2}:\d{2}:\d{2}$/.test(horaLimpia)) {
                const [h, m, s] = horaLimpia.split(':');
                return `${h.padStart(2, '0')}:${m}:${s}`;
            }
            
            // Formato HH:MM
            if (/^\d{1,2}:\d{2}$/.test(horaLimpia)) {
                const [h, m] = horaLimpia.split(':');
                return `${h.padStart(2, '0')}:${m}:00`;
            }
            
            // Retornar como está si no coincide con ningún formato
            return horaLimpia;
        }
        
        return String(hora);
        
    } catch (e) {
        console.error('Error formateando hora:', e);
        return '';
    }
}

function calcularDuracion(entradaStr, salidaStr) {
    try {
        if (!entradaStr || !salidaStr) return '0:00:00';
        
        // Función para parsear fechas en formato DD/MM/YYYY HH:mm:ss
        function parseFechaHora(str) {
            if (!str) return null;
            
            const strLimpio = String(str).trim();
            
            // Si tiene fecha y hora separadas por espacio
            if (strLimpio.includes(' ')) {
                const [fecha, hora] = strLimpio.split(' ');
                
                if (fecha && fecha.includes('/')) {
                    const [d, m, y] = fecha.split('/');
                    const anio = y.length === 2 ? '20' + y : y;
                    
                    // Asegurar que la hora tenga formato completo
                    let horaCompleta = hora;
                    if (hora && hora.split(':').length === 2) {
                        horaCompleta = hora + ':00';
                    }
                    
                    const fechaISO = `${anio}-${m.padStart(2,'0')}-${d.padStart(2,'0')}T${horaCompleta}`;
                    const parsed = new Date(fechaISO);
                    
                    if (!isNaN(parsed.getTime())) {
                        return parsed;
                    }
                }
            }
            
            return null;
        }
        
        const entrada = parseFechaHora(entradaStr);
        const salida = parseFechaHora(salidaStr);
        
        if (!entrada || !salida) {
            console.warn('No se pudieron parsear las fechas:', { entradaStr, salidaStr });
            return '0:00:00';
        }
        
        let diff = salida - entrada;
        
        // Si la diferencia es negativa, asumimos que cruza medianoche
        if (diff < 0) {
            diff += 24 * 60 * 60 * 1000;
        }
        
        const segundosTotales = Math.floor(diff / 1000);
        const horas = Math.floor(segundosTotales / 3600);
        const minutos = Math.floor((segundosTotales % 3600) / 60);
        const segundos = segundosTotales % 60;
        
        return `${horas}:${minutos.toString().padStart(2,'0')}:${segundos.toString().padStart(2,'0')}`;
        
    } catch (e) {
        console.error('Error en calcularDuracion:', e);
        return '0:00:00';
    }
}
