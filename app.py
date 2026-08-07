from flask import Flask, request, jsonify, render_template_string
import requests

app = Flask(__name__)

API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

HTML_INTERFACE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte Minuto a Minuto - IDT</title>
    
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; padding: 20px; margin: 0; }
        .card { max-width: 820px; margin: 0 auto; background: #ffffff; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
        .header { text-align: center; margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px; }
        .header h2 { color: #1a252f; margin: 0 0 5px 0; font-size: 22px; }
        .header p { color: #7f8c8d; font-size: 13px; margin: 0; }
        
        .main-container { display: flex; gap: 20px; }
        .presets-sidebar { width: 180px; border-right: 1px solid #eee; padding-right: 15px; display: flex; flex-direction: column; gap: 6px; }
        .presets-sidebar button { background: #f8f9fa; border: 1px solid #dde2e5; color: #495057; text-align: left; padding: 8px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; transition: all 0.2s; }
        .presets-sidebar button:hover { background: #e9ecef; color: #212529; }
        
        .form-content { flex: 1; }
        .form-group { margin-bottom: 15px; }
        .form-row { display: flex; gap: 10px; }
        .form-row .form-group { flex: 1; }
        
        label { display: block; font-weight: 600; margin-bottom: 5px; color: #34495e; font-size: 12px; }
        select, input { width: 100%; padding: 9px; border: 1px solid #dcdfe6; border-radius: 6px; box-sizing: border-box; font-size: 13px; color: #2c3e50; }
        
        .select2-container .select2-selection--single { height: 40px !important; border: 1px solid #dcdfe6 !important; border-radius: 6px !important; display: flex !important; align-items: center !important; }
        .select2-container--default .select2-selection--single .select2-selection__rendered { color: #2c3e50 !important; font-size: 13px !important; }
        .select2-container--default .select2-selection--single .select2-selection__arrow { height: 38px !important; }
        
        .btn-submit { width: 100%; background: #28a745; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; font-size: 14px; cursor: pointer; transition: background 0.2s; margin-top: 10px; }
        .btn-submit:hover { background: #218838; }
        .btn-submit:disabled { background: #a5d6a7; cursor: not-allowed; }
        .status-msg { font-size: 13px; color: #e67e22; margin-top: 10px; text-align: center; font-weight: bold; }
        .retry-btn { font-size: 11px; color: #007bff; text-decoration: underline; cursor: pointer; margin-left: 8px; }
    </style>
</head>
<body>
    <div class="card">
        <div class="header">
            <h2>📊 Reporte Minuto a Minuto</h2>
            <p>IDT Tecnologías - Generador de Histórico GPS Avanzado</p>
        </div>
        
        <div class="main-container">
            <div class="presets-sidebar">
                <label>Atajos de Fecha:</label>
                <button type="button" onclick="setRange('hoy')">Hoy</button>
                <button type="button" onclick="setRange('ayer')">Ayer</button>
                <button type="button" onclick="setRange('esta_semana')">Esta Semana</button>
                <button type="button" onclick="setRange('semana_anterior')">Semana anterior</button>
                <button type="button" onclick="setRange('ultimos_7_dias')">Últimos 7 días</button>
                <button type="button" onclick="setRange('este_mes')">Este mes</button>
                <button type="button" onclick="setRange('mes_anterior')">Mes anterior</button>
            </div>

            <div class="form-content">
                <div class="form-group">
                    <label>Buscar / Seleccionar Unidad: <span class="retry-btn" onclick="cargarUnidades()">🔄 Reintentar carga</span></label>
                    <select id="unit_select" style="width: 100%;">
                        <option value="">⏳ Cargando catálogo de unidades...</option>
                    </select>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label>Fecha Inicial:</label>
                        <input type="date" id="fecha_inicio" required>
                    </div>
                    <div class="form-group">
                        <label>Fecha Final:</label>
                        <input type="date" id="fecha_fin" required>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label>Hora Inicio:</label>
                        <input type="time" id="hora_inicio" value="00:00" required>
                    </div>
                    <div class="form-group">
                        <label>Hora Fin:</label>
                        <input type="time" id="hora_fin" value="23:59" required>
                    </div>
                </div>

                <div class="form-group">
                    <label>Límite de Velocidad Permitido (km/h):</label>
                    <input type="number" id="limite_velocidad" value="80" min="1" max="150" required>
                </div>

                <button type="button" class="btn-submit" id="btn_submit" onclick="generarReporte()" disabled>📥 Generar y Descargar Excel</button>
                <div id="status_msg" class="status-msg"></div>
            </div>
        </div>
    </div>

    <script>
        let unidadesCache = [];

        async function fetchLocalProxy(endpoint) {
            try {
                const res = await fetch(`/api_proxy?endpoint=${encodeURIComponent(endpoint)}`);
                if (res.ok) return await res.json();
            } catch (e) { console.error(e); }
            return null;
        }

        async function cargarUnidades() {
            const select = $('#unit_select');
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');

            status.innerText = "⏳ Cargando catálogo de unidades...";
            btn.disabled = true;

            const data = await fetchLocalProxy("/unit/list.json");
            unidadesCache = data?.data?.units || data?.units || [];

            select.empty();
            if (unidadesCache.length === 0) {
                select.append(new Option('No se encontraron unidades', ''));
                status.innerText = "Error de conexión con la API.";
                return;
            }

            select.append(new Option('🔍 Escribe para buscar unidad...', ''));
            unidadesCache.forEach(u => {
                const text = `${u.label || ''} ${u.number || ''} (ID: ${u.unit_id})`.trim();
                select.append(new Option(text, u.unit_id));
            });

            select.select2({ placeholder: "🔍 Escribe para buscar unidad...", allowClear: true, width: '100%' });
            btn.disabled = false;
            status.innerText = "";
        }

        function haversineKm(lat1, lon1, lat2, lon2) {
            if (!lat1 || !lon1 || !lat2 || !lon2) return 0;
            const R = 6371;
            const dLat = (lat2 - lat1) * Math.PI / 180;
            const dLon = (lon2 - lon1) * Math.PI / 180;
            const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
                      Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
                      Math.sin(dLon/2) * Math.sin(dLon/2);
            return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)));
        }

        function extraerPuntosExhaustivo(routeData) {
            let points = [];
            if (!routeData) return points;
            const units = routeData.data?.units || routeData.units || (Array.isArray(routeData.data) ? routeData.data : []);

            const addPt = (pt) => {
                if (!pt) return;
                const timeStr = pt.gmt || pt.time || pt.datetime || pt.t || '';
                let ts = 0;
                if (typeof timeStr === 'number') {
                    ts = timeStr > 10000000000 ? timeStr : timeStr * 1000;
                } else if (timeStr) {
                    ts = new Date(String(timeStr).replace(' ', 'T')).getTime();
                }
                
                const lat = parseFloat(pt.lat || pt.latitude);
                const lng = parseFloat(pt.lng || pt.longitude);
                const speed = parseFloat(pt.speed || pt.s || 0);

                if (!isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0) {
                    points.push({
                        timestamp: ts,
                        lat: lat,
                        lng: lng,
                        speed: speed,
                        address: pt.address || pt.addr || '',
                        acc: pt.params?.acc !== undefined ? parseInt(pt.params.acc) : (speed > 0 ? 1 : 0)
                    });
                }
            };

            if (Array.isArray(units)) {
                units.forEach(u => {
                    if (Array.isArray(u.points)) u.points.forEach(addPt);
                    if (Array.isArray(u.routes)) {
                        u.routes.forEach(r => {
                            if (Array.isArray(r.points)) r.points.forEach(addPt);
                            if (r.start) addPt(r.start);
                            if (r.end) addPt(r.end);
                        });
                    }
                    if (Array.isArray(u.tracks)) {
                        u.tracks.forEach(t => {
                            if (Array.isArray(t.points)) t.points.forEach(addPt);
                        });
                    }
                });
            }
            return points;
        }

        async function generarReporte() {
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');
            const unitId = parseInt($('#unit_select').val());
            const unitText = $('#unit_select option:selected').text();
            const fechaInicio = document.getElementById('fecha_inicio').value;
            const fechaFin = document.getElementById('fecha_fin').value;
            const horaInicio = document.getElementById('hora_inicio').value;
            const horaFin = document.getElementById('hora_fin').value;
            const limiteVel = parseInt(document.getElementById('limite_velocidad').value) || 80;

            if (!unitId) { alert("Por favor selecciona una unidad."); return; }

            btn.disabled = true;
            status.innerText = "⏳ 1/2 Obteniendo alertas...";

            try {
                const startDt = new Date(`${fechaInicio}T${horaInicio}:00`);
                const endDt = new Date(`${fechaFin}T${horaFin}:59`);
                const startMsTotal = startDt.getTime();
                const endMsTotal = endDt.getTime();

                const fromSec = Math.floor(startMsTotal / 1000);
                const tillSec = Math.floor(endMsTotal / 1000);

                const alertsData = await fetchLocalProxy(`/alert/list.json?unit_id=${unitId}&from=${fromSec}&till=${tillSec}`);
                const alertas = alertsData?.data?.alerts || alertsData?.alerts || [];

                status.innerText = "⏳ 2/2 Obteniendo recorrido GPS...";

                let routeData = await fetchLocalProxy(`/route/list.json?unit_id=${unitId}&from=${fromSec}&till=${tillSec}&include_points=1&include[]=points`);
                let rawPoints = extraerPuntosExhaustivo(routeData);

                if (rawPoints.length === 0) {
                    alert(`La API de IDT no devolvió registros de movimiento para ${unitText} en las fechas seleccionadas.`);
                    status.innerText = "Sin datos en el rango.";
                    btn.disabled = false;
                    return;
                }

                status.innerText = "⚡ Generando archivo Excel...";

                rawPoints.sort((a, b) => a.timestamp - b.timestamp);

                let currentIndex = 0;
                let lastKnown = rawPoints[0];
                let rows = [];

                let maxVelocidad = 0;
                let minutosMovimiento = 0;
                let minutosRalenti = 0;
                let recorridoTotalKm = 0;

                for (let timeMs = startMsTotal; timeMs <= endMsTotal; timeMs += 60000) {
                    while (currentIndex < rawPoints.length && rawPoints[currentIndex].timestamp <= timeMs) {
                        lastKnown = rawPoints[currentIndex];
                        currentIndex++;
                    }

                    const fechaObj = new Date(timeMs);
                    const fechaFormatted = fechaObj.getFullYear() + '-' +
                        String(fechaObj.getMonth() + 1).padStart(2, '0') + '-' +
                        String(fechaObj.getDate()).padStart(2, '0') + ' ' +
                        String(fechaObj.getHours()).padStart(2, '0') + ':' +
                        String(fechaObj.getMinutes()).padStart(2, '0') + ':00';

                    const speedInt = Math.round(lastKnown.speed);
                    if (speedInt > maxVelocidad) maxVelocidad = speedInt;

                    const tieneAlerta = alertas.some(a => {
                        const aTime = new Date((a.gmt || a.time || '').replace('Z', '')).getTime();
                        return Math.abs(aTime - timeMs) <= 60000;
                    });

                    let evento = "Apagado";
                    let detalle = "-";

                    if (speedInt > limiteVel || tieneAlerta) {
                        evento = "🚨 Exceso de velocidad";
                        detalle = `Velocidad: ${speedInt} km/h (Límite: ${limiteVel} km/h)`;
                        minutosMovimiento++;
                    } else if (speedInt > 0) {
                        evento = "En movimiento";
                        minutosMovimiento++;
                    } else if (lastKnown.acc === 1) {
                        evento = "Ralentí / Motor ON";
                        minutosRalenti++;
                    }

                    let direccion = lastKnown.address || "Sonora, Mexico";
                    let ciudad = "";
                    if (direccion.includes(',')) {
                        let partes = direccion.split(',');
                        ciudad = partes[partes.length - 2]?.trim() || "";
                    }

                    const mapsUrl = `https://www.google.com/maps?q=${lastKnown.lat},${lastKnown.lng}`;

                    if (rows.length > 0) {
                        const prevLat = rows[rows.length - 1][9];
                        const prevLng = rows[rows.length - 1][8];
                        if (prevLat !== lastKnown.lat || prevLng !== lastKnown.lng) {
                            recorridoTotalKm += haversineKm(prevLat, prevLng, lastKnown.lat, lastKnown.lng);
                        }
                    }

                    rows.push([
                        unitText,
                        fechaFormatted,
                        direccion,
                        ciudad,
                        speedInt,
                        evento,
                        detalle,
                        { f: `HYPERLINK("${mapsUrl}", "Ver en Mapa")` },
                        lastKnown.lng,
                        lastKnown.lat
                    ]);
                }

                const fmtHM = (mins) => `${Math.floor(mins / 60)}h ${mins % 60}m`;
                const horasTrabajadasMins = minutosMovimiento + minutosRalenti;
                const velocidadPromedio = minutosMovimiento > 0 ? Math.round(recorridoTotalKm / (minutosMovimiento / 60)) : 0;

                let aoa = [
                    ["", "", "Histórico Minuto a Minuto con Excesos de Velocidad"],
                    [],
                    ["", "", unitText],
                    [],
                    ["Recorrido Aprox:", `${recorridoTotalKm.toFixed(2)} km`, "Tiempo en Movimiento:", fmtHM(minutosMovimiento), "Fecha Inicial:", `${fechaInicio} ${horaInicio}`],
                    ["Velocidad Máxima:", `${maxVelocidad} km/h`, "Tiempo Muerto:", fmtHM(minutosRalenti), "Fecha Final:", `${fechaFin} ${horaFin}`],
                    ["Velocidad Promedio:", `${velocidadPromedio} km/h`, "Horas Trabajadas:", fmtHM(horasTrabajadasMins), "Consumo Combustible:", "N/A"],
                    ["Costo Combustible:", "N/A"],
                    [],
                    ["Vehículo", "Fecha", "Dirección", "Ciudad", "Velocidad (Km/h)", "Evento", "Detalle", "Mapa", "Longitud", "Latitud"]
                ].concat(rows);

                const ws = XLSX.utils.aoa_to_sheet(aoa);
                ws['!cols'] = [
                    { wch: 32 }, { wch: 20 }, { wch: 45 }, { wch: 20 }, 
                    { wch: 15 }, { wch: 24 }, { wch: 35 }, { wch: 18 }, 
                    { wch: 15 }, { wch: 15 }
                ];

                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Reporte");
                
                const safeName = unitText.replace(/[^a-zA-Z0-9]/g, '_');
                XLSX.writeFile(wb, `Reporte_Minuto_a_Minuto_${safeName}_${fechaInicio}.xlsx`);

                status.innerText = "¡Reporte generado y descargado con éxito!";
            } catch (err) {
                console.error(err);
                alert("Error durante la generación del reporte.");
                status.innerText = "Error de procesamiento.";
            }

            btn.disabled = false;
        }

        function setRange(type) {
            const now = new Date();
            let start = new Date();
            let end = new Date();

            const formatDate = (d) => d.toISOString().substring(0, 10);

            if (type === 'hoy') { start = now; end = now; }
            else if (type === 'ayer') { start.setDate(now.getDate() - 1); end.setDate(now.getDate() - 1); }
            else if (type === 'esta_semana') { const day = now.getDay() || 7; start.setDate(now.getDate() - day + 1); end = now; }
            else if (type === 'semana_anterior') { const day = now.getDay() || 7; start.setDate(now.getDate() - day - 6); end.setDate(now.getDate() - day); }
            else if (type === 'ultimos_7_dias') { start.setDate(now.getDate() - 6); end = now; }
            else if (type === 'este_mes') { start = new Date(now.getFullYear(), now.getMonth(), 1); end = now; }
            else if (type === 'mes_anterior') { start = new Date(now.getFullYear(), now.getMonth() - 1, 1); end = new Date(now.getFullYear(), now.getMonth(), 0); }

            document.getElementById('fecha_inicio').value = formatDate(start);
            document.getElementById('fecha_fin').value = formatDate(end);
        }

        setRange('hoy');
        cargarUnidades();
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_INTERFACE)

@app.route('/api_proxy')
def api_proxy():
    endpoint = request.args.get('endpoint', '')
    if not endpoint:
        return jsonify({'error': 'No endpoint specified'}), 400
    
    url = f"{BASE_URL}{endpoint}"
    if 'key=' not in url:
        sep = '&' if '?' in url else '?'
        url = f"{url}{sep}key={API_KEY}"

    try:
        res = requests.get(url, timeout=25)
        return jsonify(res.json()), res.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)