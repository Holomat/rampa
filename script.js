const { useState, useRef, useEffect } = React;

const MSCHFEventsBillboard = () => {
    const [selectedEvent, setSelectedEvent] = useState(null);
    const [showBrutalistInfo, setShowBrutalistInfo] = useState(true);
    const [brutalistEventData, setBrutalistEventData] = useState(null);
    const [currentEventIndex, setCurrentEventIndex] = useState(0);
    
    // Estados para swipe
    const [startX, setStartX] = useState(0);
    const [startY, setStartY] = useState(0);
    const [isSwipping, setIsSwipping] = useState(false);
    const overlayRef = useRef(null);
    
    // Estados para carrusel
    const [carouselOffset, setCarouselOffset] = useState(0);
    const [isDraggingCarousel, setIsDraggingCarousel] = useState(false);

    // Estados para etiqueta con physics
    const [tagPosition, setTagPosition] = useState({ x: 0, y: 0 });
    const [tagWidth, setTagWidth] = useState(60);
    const [isDraggingTag, setIsDraggingTag] = useState(false);
    const [eventPositions, setEventPositions] = useState([]);
    const [isTagVisible, setIsTagVisible] = useState(true);
    const tagsContainerRef = useRef(null);

    // Estados para eventos dinámicos desde Excel
    const [events, setEvents] = useState([]);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState(null);
    const [lastUpdated, setLastUpdated] = useState(null);

    // ⚠️ CONFIGURACIÓN - REEMPLAZA CON TUS DATOS
    const EXCEL_CONFIG = {
        GITHUB_USER: 'tu-usuario',        // ⚠️ CAMBIAR POR TU USUARIO GITHUB
        GITHUB_REPO: 'tu-repositorio',    // ⚠️ CAMBIAR POR TU REPOSITORIO
        EXCEL_FILE: 'eventos.xlsx',       // Nombre del archivo Excel
        SHEET_NAME: 'Rampa'               // Nombre de la hoja (pestaña) en Excel
    };

    // FUNCIÓN PARA CARGAR EVENTOS DESDE EXCEL EN GITHUB
    const loadEventsFromExcel = async () => {
        try {
            setLoading(true);
            
            // URL RAW del archivo Excel en GitHub
            const excelUrl = `https://raw.githubusercontent.com/${EXCEL_CONFIG.GITHUB_USER}/${EXCEL_CONFIG.GITHUB_REPO}/main/${EXCEL_CONFIG.EXCEL_FILE}`;
            
            console.log('🔄 Cargando Excel desde:', excelUrl);
            
            // Descargar el archivo Excel
            const response = await fetch(excelUrl);
            if (!response.ok) {
                throw new Error(`Error HTTP ${response.status}: ¿Existe el archivo ${EXCEL_CONFIG.EXCEL_FILE} en tu repositorio?`);
            }
            
            // Convertir a ArrayBuffer para XLSX
            const arrayBuffer = await response.arrayBuffer();
            
            // Verificar que XLSX esté disponible
            if (typeof XLSX === 'undefined') {
                throw new Error('Librería XLSX no cargada. Verifica que esté incluida en index.html');
            }
            
            // Leer el workbook
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            
            // Obtener la hoja específica
            const sheetName = EXCEL_CONFIG.SHEET_NAME;
            if (!workbook.Sheets[sheetName]) {
                // Si no encuentra la hoja, usar la primera disponible
                const availableSheets = Object.keys(workbook.Sheets);
                if (availableSheets.length === 0) {
                    throw new Error('El archivo Excel no tiene hojas válidas');
                }
                
                console.warn(`Hoja "${sheetName}" no encontrada. Usando: "${availableSheets[0]}"`);
                const firstSheet = availableSheets[0];
                var worksheet = workbook.Sheets[firstSheet];
            } else {
                var worksheet = workbook.Sheets[sheetName];
            }
            
            // Convertir hoja a JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonData.length < 2) {
                throw new Error('El Excel debe tener al menos 2 filas (header + datos)');
            }
            
            // Primera fila son los headers
            const headers = jsonData[0];
            const rows = jsonData.slice(1);
            
            console.log('📊 Headers detectados:', headers);
            console.log('📊 Filas de datos:', rows.length);
            
            // Mapear columnas (flexible - busca por nombre de columna)
            const getColumnIndex = (searchTerms) => {
                for (let i = 0; i < headers.length; i++) {
                    const header = headers[i]?.toString().toLowerCase().trim();
                    if (searchTerms.some(term => header.includes(term))) {
                        return i;
                    }
                }
                return -1;
            };
            
            // Encontrar índices de columnas
            const columnIndices = {
                nombre: getColumnIndex(['nombre', 'name', 'evento', 'event']),
                fecha: getColumnIndex(['fecha', 'date', 'día']),
                hora: getColumnIndex(['hora', 'time', 'horario']),
                venue: getColumnIndex(['venue', 'lugar', 'ubicación', 'location']),
                dia: getColumnIndex(['dia', 'día', 'day']),
                categoria: getColumnIndex(['categoria', 'categoría', 'category', 'tipo'])
            };
            
            console.log('🗂️ Columnas mapeadas:', columnIndices);
            
            // Verificar que al menos tengamos nombre
            if (columnIndices.nombre === -1) {
                throw new Error('No se encontró columna de "Nombre" en el Excel. Verifica los headers.');
            }
            
            // Procesar filas de datos
            const processedEvents = rows
                .filter(row => row && row.length > 0 && row[columnIndices.nombre]) // Filtrar filas vacías
                .map((row, index) => {
                    // Función auxiliar para obtener valor de celda
                    const getValue = (colIndex, defaultValue = '') => {
                        if (colIndex === -1) return defaultValue;
                        const value = row[colIndex];
                        if (value === null || value === undefined || value === '') return defaultValue;
                        return value.toString().trim();
                    };
                    
                    // Procesar fecha
                    let fecha = getValue(columnIndices.fecha, '01.01.25');
                    // Si la fecha viene como número de Excel, convertirla
                    if (!isNaN(fecha) && fecha !== '') {
                        const excelDate = new Date((fecha - 25569) * 86400 * 1000);
                        const day = String(excelDate.getDate()).padStart(2, '0');
                        const month = String(excelDate.getMonth() + 1).padStart(2, '0');
                        const year = String(excelDate.getFullYear()).slice(-2);
                        fecha = `${day}.${month}.${year}`;
                    }
                    
                    // Procesar hora
                    let hora = getValue(columnIndices.hora, '20:00');
                    // Si la hora viene como decimal de Excel, convertirla
                    if (!isNaN(hora) && hora !== '' && hora < 1) {
                        const totalMinutes = Math.round(hora * 24 * 60);
                        const hours = Math.floor(totalMinutes / 60);
                        const minutes = totalMinutes % 60;
                        hora = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
                    }
                    
                    return {
                        id: index + 1,
                        name: getValue(columnIndices.nombre, 'Sin nombre'),
                        date: fecha,
                        time: hora,
                        venue: getValue(columnIndices.venue, 'Venue TBD'),
                        day: getValue(columnIndices.dia, 'VIERNES').toUpperCase(),
                        category: getValue(columnIndices.categoria, 'EVENTO').toUpperCase()
                    };
                })
                .filter(event => event.name !== 'Sin nombre' && event.name !== ''); // Filtrar eventos sin nombre
            
            if (processedEvents.length === 0) {
                throw new Error('No se encontraron eventos válidos en el Excel. Verifica que haya datos en las filas.');
            }
            
            console.log(`✅ ${processedEvents.length} eventos cargados exitosamente:`, processedEvents);
            
            setEvents(processedEvents);
            setError(null);
            setLastUpdated(new Date());
            
            // Auto-seleccionar primer evento
            if (!selectedEvent && processedEvents.length > 0) {
                setSelectedEvent(processedEvents[0].id);
                setBrutalistEventData(processedEvents[0]);
                setCurrentEventIndex(0);
            }
            
        } catch (err) {
            console.error('❌ Error cargando eventos desde Excel:', err);
            setError(`${err.message}`);
            
            // Fallback a eventos por defecto
            const fallbackEvents = [
                { 
                    id: 1, 
                    name: "Error al cargar Excel", 
                    date: "01.01.25", 
                    time: "20:00", 
                    venue: "Revisa la configuración", 
                    category: "ERROR", 
                    day: "VIERNES" 
                }
            ];
            setEvents(fallbackEvents);
        } finally {
            setLoading(false);
        }
    };

    // DETECTAR CAMBIOS EN EL ARCHIVO (opcional)
    const checkForUpdates = async () => {
        try {
            // Usar GitHub API para verificar última modificación
            const apiUrl = `https://api.github.com/repos/${EXCEL_CONFIG.GITHUB_USER}/${EXCEL_CONFIG.GITHUB_REPO}/contents/${EXCEL_CONFIG.EXCEL_FILE}`;
            
            const response = await fetch(apiUrl);
            if (response.ok) {
                const fileInfo = await response.json();
                
                // Si el SHA cambió, recargar
                if (window.lastKnownSha && fileInfo.sha !== window.lastKnownSha) {
                    console.log('📄 Excel actualizado, recargando eventos...');
                    loadEventsFromExcel();
                }
                window.lastKnownSha = fileInfo.sha;
            }
        } catch (err) {
            // Si falla la verificación, continuar normalmente
            console.log('⚠️ No se pudo verificar actualizaciones:', err.message);
        }
    };

    // CARGAR EVENTOS AL MONTAR EL COMPONENTE
    useEffect(() => {
        loadEventsFromExcel();
        
        // Verificar actualizaciones cada 3 minutos
        const interval = setInterval(checkForUpdates, 3 * 60 * 1000);
        return () => clearInterval(interval);
    }, []);

    // Calcular posiciones de eventos con mejor centrado vertical
    useEffect(() => {
        if (tagsContainerRef.current && events.length > 0) {
            const container = tagsContainerRef.current;
            const eventElements = container.querySelectorAll('.event-text');
            const positions = Array.from(eventElements).map((element, index) => {
                const rect = element.getBoundingClientRect();
                const containerRect = container.getBoundingClientRect();
                
                const textCenterY = rect.top - containerRect.top + (rect.height / 2) - 7;
                
                return {
                    id: events[index].id,
                    x: rect.left - containerRect.left,
                    y: textCenterY,
                    width: rect.width,
                    height: rect.height
                };
            });
            setEventPositions(positions);
        }
    }, [events]);

    // Inicializar primer evento
    useEffect(() => {
        if (events.length > 0 && eventPositions.length > 0) {
            const firstEvent = events[0];
            setSelectedEvent(firstEvent.id);
            setBrutalistEventData(firstEvent);
            setCurrentEventIndex(0);
            setCarouselOffset(0);
            
            const firstPosition = eventPositions[0];
            const firstWidth = getEventWidth(0);
            if (firstPosition) {
                setTagPosition({ x: firstPosition.x, y: firstPosition.y });
                setTagWidth(firstWidth);
                setIsTagVisible(true);
            }
        }
    }, [eventPositions]);

    // Obtener ancho del evento
    const getEventWidth = (eventIndex) => {
        if (tagsContainerRef.current) {
            const eventElement = tagsContainerRef.current.querySelectorAll('.event-text')[eventIndex];
            if (eventElement) {
                return eventElement.offsetWidth;
            }
        }
        return 60;
    };

    // Detectar cambio de línea con mayor precisión para mobile
    const isLineBreak = (fromIndex, toIndex) => {
        if (!eventPositions[fromIndex] || !eventPositions[toIndex]) return false;
        
        const fromY = eventPositions[fromIndex].y;
        const toY = eventPositions[toIndex].y;
        const yDifference = Math.abs(toY - fromY);
        
        return yDifference > 8;
    };

    const isLastInLine = (eventIndex) => {
        if (eventIndex >= events.length - 1) return true;
        
        const currentY = eventPositions[eventIndex]?.y;
        const nextY = eventPositions[eventIndex + 1]?.y;
        
        if (!currentY || !nextY) return false;
        
        return Math.abs(nextY - currentY) > 8;
    };

    const isFirstInLine = (eventIndex) => {
        if (eventIndex === 0) return true;
        
        const currentY = eventPositions[eventIndex]?.y;
        const prevY = eventPositions[eventIndex - 1]?.y;
        
        if (!currentY || !prevY) return false;
        
        return Math.abs(currentY - prevY) > 8;
    };

    const moveTagToEvent = (eventIndex) => {
        if (eventPositions[eventIndex]) {
            const position = eventPositions[eventIndex];
            const width = getEventWidth(eventIndex);
            
            const hasLineBreak = isLineBreak(currentEventIndex, eventIndex);
            const isFromLastInLine = isLastInLine(currentEventIndex);
            const isToFirstInLine = isFirstInLine(eventIndex);
            
            const tagElement = document.querySelector('.sliding-tag');
            
            if (hasLineBreak && isFromLastInLine && isToFirstInLine) {
                if (tagElement) {
                    tagElement.className = 'sliding-tag fade-out';
                }
                setIsTagVisible(false);
                
                setTimeout(() => {
                    setTagPosition({ x: position.x, y: position.y });
                    setTagWidth(width);
                    
                    setTimeout(() => {
                        if (tagElement) {
                            tagElement.className = 'sliding-tag fade-in';
                        }
                        setIsTagVisible(true);
                        setIsDraggingTag(false);
                    }, 50);
                }, 300);
                
            } else {
                if (tagElement) {
                    tagElement.className = 'sliding-tag smooth-horizontal';
                }
                
                setTagPosition({ x: position.x, y: position.y });
                setTagWidth(width);
                setIsTagVisible(true);
                setIsDraggingTag(false);
            }
        }
    };

    const moveCarousel = (index) => {
        if (events.length === 0) return;
        const offset = -index * (100 / events.length);
        setCarouselOffset(offset);
        setIsDraggingCarousel(false);
    };

    const animateEventTransition = (newEvent, newIndex) => {
        moveTagToEvent(newIndex);
        moveCarousel(newIndex);
        setSelectedEvent(newEvent.id);
        setBrutalistEventData(newEvent);
    };

    // Funciones de swipe
    const handleTouchStart = (e) => {
        setStartX(e.touches[0].clientX);
        setStartY(e.touches[0].clientY);
        setIsSwipping(true);
        setIsDraggingTag(true);
        setIsDraggingCarousel(true);
    };

    const handleTouchMove = (e) => {
        if (!isSwipping || events.length === 0) return;
        
        const currentX = e.touches[0].clientX;
        const currentY = e.touches[0].clientY;
        const diffX = startX - currentX;
        const diffY = startY - currentY;
        
        if (Math.abs(diffX) > Math.abs(diffY) && Math.abs(diffX) > 5) {
            e.preventDefault();
            
            const deadZone = 15;
            const effectiveDiff = Math.abs(diffX) < deadZone ? 0 : diffX - (deadZone * Math.sign(diffX));
            
            const rawProgress = effectiveDiff / 140;
            const physicsProgress = Math.tanh(rawProgress) * 0.8;
            
            if (eventPositions.length > 0) {
                const direction = diffX > 0 ? 1 : -1;
                const currentPos = eventPositions[currentEventIndex];
                let targetIndex = currentEventIndex + direction;
                
                if (targetIndex >= events.length) targetIndex = 0;
                if (targetIndex < 0) targetIndex = events.length - 1;
                
                const targetPos = eventPositions[targetIndex];
                const currentWidth = getEventWidth(currentEventIndex);
                const targetWidth = getEventWidth(targetIndex);
                
                if (currentPos && targetPos) {
                    const progressAbs = Math.min(Math.abs(physicsProgress), 1);
                    const smoothProgress = progressAbs * progressAbs * (3 - 2 * progressAbs);
                    
                    const interpolatedPos = {
                        x: currentPos.x + (targetPos.x - currentPos.x) * smoothProgress,
                        y: currentPos.y + (targetPos.y - currentPos.y) * smoothProgress
                    };
                    
                    const interpolatedWidth = currentWidth + (targetWidth - currentWidth) * smoothProgress;
                    
                    setTagPosition(interpolatedPos);
                    setTagWidth(interpolatedWidth);
                }
            }
            
            const baseOffset = events.length > 0 ? -currentEventIndex * (100 / events.length) : 0;
            const carouselEffectiveDiff = Math.abs(diffX) < deadZone ? 0 : diffX - (deadZone * Math.sign(diffX));
            const carouselPhysicsProgress = Math.tanh(carouselEffectiveDiff / 140) * 0.8;
            const dragOffset = carouselPhysicsProgress * (100 / events.length);
            setCarouselOffset(baseOffset - dragOffset);
        }
    };

    const handleTouchEnd = (e) => {
        if (!isSwipping) return;
        
        const endX = e.changedTouches[0].clientX;
        const diffX = startX - endX;
        
        setIsSwipping(false);
        setIsDraggingTag(false);
        setIsDraggingCarousel(false);
        
        if (Math.abs(diffX) > 45) {
            if (diffX > 0) {
                handleArrowDown();
            } else {
                handleArrowUp();
            }
        } else {
            moveTagToEvent(currentEventIndex);
            moveCarousel(currentEventIndex);
        }
    };

    const handleMouseStart = (e) => {
        setStartX(e.clientX);
        setStartY(e.clientY);
        setIsSwipping(true);
        setIsDraggingTag(true);
        setIsDraggingCarousel(true);
    };

    const handleMouseMove = (e) => {
        if (!isSwipping || events.length === 0) return;
        
        const currentX = e.clientX;
        const diffX = startX - currentX;
        const diffY = startY - e.clientY;
        
        if (Math.abs(diffX) > Math.abs(diffY) && Math.abs(diffX) > 5) {
            e.preventDefault();
            
            const deadZone = 15;
            const effectiveDiff = Math.abs(diffX) < deadZone ? 0 : diffX - (deadZone * Math.sign(diffX));
            
            const rawProgress = effectiveDiff / 140;
            const physicsProgress = Math.tanh(rawProgress) * 0.8;
            
            if (eventPositions.length > 0) {
                const direction = diffX > 0 ? 1 : -1;
                const currentPos = eventPositions[currentEventIndex];
                let targetIndex = currentEventIndex + direction;
                
                if (targetIndex >= events.length) targetIndex = 0;
                if (targetIndex < 0) targetIndex = events.length - 1;
                
                const targetPos = eventPositions[targetIndex];
                const currentWidth = getEventWidth(currentEventIndex);
                const targetWidth = getEventWidth(targetIndex);
                
                if (currentPos && targetPos) {
                    const progressAbs = Math.min(Math.abs(physicsProgress), 1);
                    const smoothProgress = progressAbs * progressAbs * (3 - 2 * progressAbs);
                    
                    const interpolatedPos = {
                        x: currentPos.x + (targetPos.x - currentPos.x) * smoothProgress,
                        y: currentPos.y + (targetPos.y - currentPos.y) * smoothProgress
                    };
                    
                    const interpolatedWidth = currentWidth + (targetWidth - currentWidth) * smoothProgress;
                    
                    setTagPosition(interpolatedPos);
                    setTagWidth(interpolatedWidth);
                }
            }
            
            const baseOffset = events.length > 0 ? -currentEventIndex * (100 / events.length) : 0;
            const carouselEffectiveDiff = Math.abs(diffX) < deadZone ? 0 : diffX - (deadZone * Math.sign(diffX));
            const carouselPhysicsProgress = Math.tanh(carouselEffectiveDiff / 140) * 0.8;
            const dragOffset = carouselPhysicsProgress * (100 / events.length);
            setCarouselOffset(baseOffset - dragOffset);
        }
    };

    const handleMouseEnd = (e) => {
        if (!isSwipping) return;
        
        const endX = e.clientX;
        const diffX = startX - endX;
        
        setIsSwipping(false);
        setIsDraggingTag(false);
        setIsDraggingCarousel(false);
        
        if (Math.abs(diffX) > 45) {
            if (diffX > 0) {
                handleArrowDown();
            } else {
                handleArrowUp();
            }
        } else {
            moveTagToEvent(currentEventIndex);
            moveCarousel(currentEventIndex);
        }
    };

    const handleArrowDown = () => {
        if (events.length === 0) return;
        const nextIndex = (currentEventIndex + 1) % events.length;
        setCurrentEventIndex(nextIndex);
        const nextEvent = events[nextIndex];
        animateEventTransition(nextEvent, nextIndex);
    };

    const handleArrowUp = () => {
        if (events.length === 0) return;
        const prevIndex = currentEventIndex === 0 ? events.length - 1 : currentEventIndex - 1;
        setCurrentEventIndex(prevIndex);
        const prevEvent = events[prevIndex];
        animateEventTransition(prevEvent, prevIndex);
    };

    const handleEventTagClick = (event) => {
        const eventIndex = events.findIndex(e => e.id === event.id);
        setCurrentEventIndex(eventIndex);
        animateEventTransition(event, eventIndex);
        setShowBrutalistInfo(true);
    };

    // MOSTRAR ESTADO DE CARGA
    if (loading) {
        return (
            <div className="min-h-screen bg-black flex items-center justify-center">
                <div className="text-center" style={{ fontFamily: 'Space Mono, monospace', color: '#B8B0A3' }}>
                    <div className="text-2xl mb-4">📊 Cargando eventos...</div>
                    <div className="animate-pulse text-sm">Leyendo Excel desde GitHub</div>
                    <div className="text-xs mt-2 opacity-60">
                        {EXCEL_CONFIG.GITHUB_USER}/{EXCEL_CONFIG.GITHUB_REPO}/{EXCEL_CONFIG.EXCEL_FILE}
                    </div>
                </div>
            </div>
        );
    }

    // MOSTRAR ERROR SI HAY PROBLEMAS
    if (error && events.length === 1 && events[0].name === "Error al cargar Excel") {
        return (
            <div className="min-h-screen bg-black flex items-center justify-center">
                <div className="text-center" style={{ fontFamily: 'Space Mono, monospace', color: '#B8B0A3' }}>
                    <div className="text-xl mb-4">❌ Error cargando Excel</div>
                    <div className="text-sm mb-4 max-w-md">{error}</div>
                    <button 
                        onClick={loadEventsFromExcel}
                        className="px-4 py-2 bg-gray-800 hover:bg-gray-700 rounded text-sm mb-4"
                    >
                        🔄 Reintentar
                    </button>
                    <div className="text-xs opacity-60 max-w-md">
                        <p>Verifica que:</p>
                        <p>• El archivo eventos.xlsx existe en tu repositorio</p>
                        <p>• La configuración en script.js es correcta</p>
                        <p>• El repositorio es público o GitHub Pages está habilitado</p>
                    </div>
                </div>
            </div>
        );
    }

    return (
        <div className="min-h-screen bg-black overflow-hidden" style={{ fontFamily: 'Space Mono, monospace', padding: '10px' }}>
            {/* Gradiente de rayos horizontales en el 1/4 inferior */}
            <div style={{
                position: 'fixed',
                bottom: 0,
                left: 0,
                width: '100%',
                height: '25vh',
                background: 'linear-gradient(to top, #1a1a2e 0%, #16213e 30%, #0f3460 60%, rgba(15, 52, 96, 0.8) 80%, rgba(15, 52, 96, 0.3) 90%, transparent 100%)',
                zIndex: 1,
                pointerEvents: 'none'
            }} />
            
            {/* Logo RAMPA */}
            <div className="logo-section">
                <svg className="rampa-logo" width="452" height="154" viewBox="0 0 452 154" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M67.495 54.9663C98.6065 38.1977 131.04 37.9868 139.937 54.4946C148.835 71.0025 130.827 97.9781 99.7157 114.747C68.6044 131.515 36.1708 131.726 27.2734 115.218C18.3763 98.7104 36.3839 71.7348 67.495 54.9663ZM126.693 61.1655C119.111 47.0971 93.2836 46.299 69.0077 59.3833C44.7322 72.4675 31.1999 94.479 38.7821 108.547C46.3648 122.616 72.1918 123.414 96.4677 110.33C120.743 97.2454 134.275 75.2339 126.693 61.1655Z" fill="currentColor"/>
                    <path d="M394.951 114.713C373.073 137.843 345.914 145.245 334.29 131.244C322.666 117.243 330.979 87.1425 352.858 64.0122C374.736 40.882 401.895 33.4808 413.519 47.4813C425.143 61.4822 416.83 91.583 394.951 114.713ZM343.722 121.849C353.628 133.781 375.498 128.823 392.57 110.774C409.641 92.7261 415.45 68.4226 405.544 56.4907C395.637 44.5588 373.768 49.5167 356.696 67.565C339.624 85.6133 333.816 109.917 343.722 121.849Z" fill="currentColor"/>
                    <path d="M94.6431 111.003H85.2631L88.4131 96.3027C88.5064 95.836 88.5531 95.4627 88.5531 95.1827C88.5531 94.3894 88.2964 93.7594 87.7831 93.2927C87.2698 92.7794 86.4764 92.5227 85.4031 92.5227H75.2531L71.3331 111.003H62.1631L72.5231 62.0027H89.7431C92.0298 62.0027 94.0364 62.3293 95.7631 62.9827C97.5364 63.5893 99.0064 64.4293 100.173 65.5027C101.386 66.5294 102.273 67.766 102.833 69.2127C103.44 70.6594 103.743 72.176 103.743 73.7627C103.743 77.1694 102.95 79.946 101.363 82.0927C99.8231 84.1927 97.4431 85.7794 94.2231 86.8527V88.1127C95.5764 88.2527 96.6031 88.7427 97.3031 89.5827C98.0498 90.376 98.4231 91.4027 98.4231 92.6627C98.4231 93.1294 98.3764 93.596 98.2831 94.0627L94.6431 111.003ZM77.1431 83.7027H85.5431C88.0164 83.7027 90.0464 83.0494 91.6331 81.7427C93.2198 80.436 94.0131 78.6627 94.0131 76.4227C94.0131 74.6494 93.5231 73.2727 92.5431 72.2927C91.6098 71.3127 90.1631 70.8227 88.2031 70.8227H79.8731L77.1431 83.7027Z" fill="currentColor"/>
                    <path d="M155.527 75.2973C181.002 69.1841 204.228 75.2448 207.405 88.8338C210.581 102.423 192.503 118.395 167.028 124.508C141.553 130.622 118.327 124.562 115.15 110.973C111.974 97.3834 130.052 81.4106 155.527 75.2973ZM196.65 91.1002C193.943 79.5191 175.635 73.9979 155.757 78.7679C135.879 83.5381 121.958 96.7929 124.665 108.374C127.371 119.955 145.68 125.477 165.559 120.706C185.436 115.936 199.356 102.681 196.65 91.1002Z" fill="currentColor"/>
                    <path d="M305.308 66.2098C328.74 77.928 342.615 97.5155 336.299 109.96C329.983 122.404 305.868 122.993 282.436 111.275C259.004 99.5567 245.129 79.9697 251.444 67.5252C257.76 55.0807 281.876 54.4915 305.308 66.2098ZM326.609 104.772C331.992 94.1668 321.534 78.1574 303.25 69.0138C284.967 59.8702 265.781 61.0548 260.399 71.6601C255.016 82.2656 265.475 98.2758 283.758 107.419C302.042 116.563 321.227 115.378 326.609 104.772Z" fill="currentColor"/>
                    <path d="M252.669 76.3667C258.782 101.842 252.722 125.068 239.133 128.244C225.543 131.42 209.572 113.343 203.458 87.8675C197.345 62.3923 203.405 39.1662 216.994 35.9898C230.583 32.8139 246.556 50.8912 252.669 76.3667ZM236.866 117.489C248.447 114.783 253.969 96.4745 249.199 76.5964C244.428 56.7183 231.174 42.7978 219.593 45.5042C208.011 48.2109 202.49 66.5199 207.26 86.3981C212.03 106.276 225.285 120.196 236.866 117.489Z" fill="currentColor"/>
                    <path d="M376.616 101.351L362.125 97.4176L354.909 105.537L346.065 103.136L380.323 65.4031L395.629 69.558L381.758 112.825L373.54 110.595L376.616 101.351ZM368.139 90.584L379.179 93.581L386.505 71.5157L385.376 71.2092L368.139 90.584Z" fill="currentColor"/>
                    <path d="M165.461 107.154L155.132 108.316L153.024 115.346L146.719 116.055L156.97 83.203L167.881 81.9753L172.164 113.192L166.306 113.851L165.461 107.154ZM156.868 102.415L164.738 101.529L162.713 85.5456L161.908 85.6362L156.868 102.415Z" fill="currentColor"/>
                    <path d="M287.72 87.7026L279.04 96.1966L274.132 92.9306L297.129 70.3967L306.494 76.6295C307.593 77.3608 308.463 78.1922 309.105 79.1235C309.771 80.0715 310.209 81.0655 310.418 82.1054C310.626 83.1453 310.611 84.1978 310.37 85.263C310.155 86.3449 309.715 87.3853 309.05 88.3843C308.119 89.7829 307.034 90.8987 305.795 91.7318C304.556 92.5648 303.225 93.1026 301.802 93.345C300.38 93.5875 298.907 93.5263 297.384 93.1614C295.861 92.7965 294.35 92.1155 292.852 91.1182L287.72 87.7026ZM291.872 83.6555L296.443 86.6971C297.717 87.5448 299.027 87.8765 300.375 87.6924C301.723 87.5083 302.796 86.8168 303.593 85.618C304.208 84.6939 304.374 83.7952 304.091 82.9217C303.848 82.04 303.253 81.2833 302.304 80.6517L297.921 77.7347L291.872 83.6555Z" fill="currentColor"/>
                    <path d="M212.485 72.3721L223.6 68.5005L229.151 95.8042L229.967 102.159L231.213 101.725L230.393 95.0222L228.69 66.7274L240.117 62.7472L245.043 101.8L239.121 103.863L235.889 78.1978L235.404 69.2811L234.158 69.7153L236.005 104.948L226.708 108.187L219.667 74.763L218.42 75.1972L220.095 83.3499L223.332 109.363L217.463 111.407L212.485 72.3721Z" fill="currentColor"/>
                </svg>
            </div>

            {/* Indicadores de estado */}
            {lastUpdated && (
                <div style={{
                    position: 'fixed',
                    top: '10px',
                    right: '10px',
                    background: 'rgba(0, 255, 0, 0.1)',
                    color: '#4ade80',
                    padding: '5px 10px',
                    borderRadius: '5px',
                    fontSize: '10px',
                    zIndex: 999
                }}>
                    ✅ Actualizado: {lastUpdated.toLocaleTimeString()}
                </div>
            )}

            {error && !loading && (
                <div style={{
                    position: 'fixed',
                    top: '30px',
                    right: '10px',
                    background: 'rgba(255, 255, 0, 0.1)',
                    color: '#fbbf24',
                    padding: '5px 10px',
                    borderRadius: '5px',
                    fontSize: '10px',
                    zIndex: 999,
                    maxWidth: '200px'
                }}>
                    ⚠ Usando datos de prueba
                </div>
            )}

            {/* Sección de Eventos */}
            <div className="contributors-section">
                <div className="contributors-tags" ref={tagsContainerRef}>
                    {/* Etiqueta deslizante con physics */}
                    {selectedEvent && eventPositions.length > 0 && isTagVisible && (
                        <div 
                            className={`sliding-tag ${isDraggingTag ? 'dragging' : ''}`}
                            style={{
                                left: `${tagPosition.x}px`,
                                top: `${tagPosition.y}px`,
                                width: `${tagWidth}px`
                            }}
                        />
                    )}

                    {events.map((event, index) => (
                        <div key={event.id} className="event-wrapper">
                            <span 
                                className={`event-text ${selectedEvent === event.id ? 'has-tag' : ''}`}
                                onClick={() => handleEventTagClick(event)}
                                title={`${event.date} - ${event.time} - ${event.venue}`}
                            >
                                {event.name}
                            </span>
                        </div>
                    ))}
                </div>
            </div>

            {/* Overlay principal */}
            <div 
                className={`brutalist-overlay ${showBrutalistInfo ? 'show' : ''}`}
                ref={overlayRef}
                onTouchStart={handleTouchStart}
                onTouchMove={handleTouchMove}
                onTouchEnd={handleTouchEnd}
                onMouseDown={handleMouseStart}
                onMouseMove={handleMouseMove}
                onMouseUp={handleMouseEnd}
                onMouseLeave={handleMouseEnd}
                style={{ userSelect: 'none', touchAction: 'pan-y' }}
            >
                {brutalistEventData && (
                    <div className="event-content">
                        <div 
                            className={`event-carousel ${isDraggingCarousel ? 'dragging' : ''}`}
                            style={{
                                transform: `translateX(${carouselOffset}%)`,
                                width: `${events.length * 100}%`
                            }}
                        >
                            {events.map((event, index) => (
                                <div key={event.id} className="event-slide" style={{ width: `${100 / events.length}%` }}>
                                    <div className="date-line">
                                        <div className="day-circle">
                                            <div className="brutalist-text day day-text">{event.day}</div>
                                        </div>
                                        <div className="brutalist-text date">{event.date.replace(/\./g, '-')}</div>
                                        <div className="day-circle">
                                            <div className="brutalist-text day day-text">{event.day}</div>
                                        </div>
                                    </div>
                                    
                                    <div className="brutalist-text eventname">{event.name}</div>
                                    
                                    {event.venue && event.venue !== event.name && (
                                        <div className="brutalist-text time-label venue-text">{event.venue}</div>
                                    )}
                                    
                                    <div className="brutalist-text time-label">{event.time}H</div>
                                </div>
                            ))}
                        </div>
                    </div>
                )}
            </div>

            {/* Barra de scroll inferior */}
            <div style={{
                position: 'fixed',
                bottom: 0,
                left: 0,
                width: '100vw',
                height: '19px',
                background: '#FD7B1E',
                zIndex: 100,
                overflow: 'hidden',
                display: 'flex',
                alignItems: 'center'
            }}>
                <div style={{
                    whiteSpace: 'nowrap',
                    color: 'black',
                    background: '#FD7B1E',
                    fontSize: '10px',
                    lineHeight: '1.0',
                    animation: 'scroll-left 30s linear infinite',
                    paddingLeft: '100%',
                    paddingTop: '2px',
                    paddingBottom: '2px',
                    fontWeight: 300,
                    letterSpacing: '0.01em',
                    fontFamily: 'Inter, sans-serif',
                    filter: 'blur(0.3px)',
                    minWidth: '20px',
                    boxSizing: 'border-box',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    transform: 'translateY(0px)'
                }}>
                    Publicá tu evento en Rampa → Escribinos a hola@rampa.world
                </div>
            </div>
        </div>
    );
};

ReactDOM.render(<MSCHFEventsBillboard />, document.getElementById('root'));