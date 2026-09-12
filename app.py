from flask import Flask, render_template_string, request, send_file, jsonify
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import math
import os
import json
import threading
import uuid
import tempfile
import concurrent.futures
import bisect
from datetime import datetime, timedelta

app = Flask(__name__)

# ==========================================
# CONFIGURACIÓN DE LA API Y RUTAS
# ==========================================
API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"
COMPANY_ID = "87534"
TIMEZONE_OFFSET = -7

TASKS = {}
CACHE_GEOCERCAS = []

# ==========================================
# DECODIFICADORES OFICIALES MAPON (TELEMETRÍA REAL)
# ==========================================
def decode_polyline_2d(encoded):
    points = []
    index, lat, lng, length = 0, 0, 0, len(encoded)
    while index < length:
        shift, result = 0, 0
        while True:
            if index >= length: break
            b = ord(encoded[index]) - 63
            index += 1
            result |= (b & 0x1f) << shift
            shift += 5
            if b < 0x20: break
        lat += ~(result >> 1) if (result & 1) else (result >> 1)
        
        shift, result = 0, 0
        while True:
            if index >= length: break
            b = ord(encoded[index]) - 63
            index += 1
            result |= (b & 0x1f) << shift
            shift += 5
            if b < 0x20: break
        lng += ~(result >> 1) if (result & 1) else (result >> 1)
        points.append({'lat': lat / 100000.0, 'lng': lng / 100000.0})
    return points

def decode_mapon_speed_string(encoded_str):
    chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-.'
    points_count = len(encoded_str) // 4
    data = []
    
    for i in range(points_count):
        pos = i * 4
        try:
            offset = chars.index(encoded_str[pos]) * 64
            offset += chars.index(encoded_str[pos + 1])
            
            speed = chars.index(encoded_str[pos + 2]) * 64
            speed += chars.index(encoded_str[pos + 3])
            
            data.append((offset, speed))
        except Exception:
            pass
    return data

# ==========================================
# FÓRMULA HAVERSINE (PARA GEOCERCAS)
# ==========================================
def calcular_distancia(lat1, lon1, lat2, lon2):
    R = 6371000 
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dphi/2.0)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlon/2.0)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def cargar_geocercas_api():
    geocercas = []
    try:
        url_obj = f"{BASE_URL}/object/list.json"
        res_obj = requests.get(url_obj, params={"key": API_KEY, "limit": 500}, timeout=15)
        items = res_obj.json().get('data', {}).get('items', [])
        for geo in items:
            nombre = geo.get('name', f"Geocerca_{geo.get('id')}")
            wkt = str(geo.get('wkt', ''))
            if 'POINT' in wkt:
                try:
                    coords_str = wkt.replace('POINT(', '').replace('POINT (', '').replace(')', '')
                    lng_str, lat_str = coords_str.strip().split(' ')
                    geocercas.append({'id': str(geo.get('id')), 'name': nombre, 'lat': float(lat_str), 'lng': float(lng_str), 'radius': 250})
                except Exception: pass
            elif 'POLYGON' in wkt:
                try:
                    coords_str = wkt.split('((')[1].split(',')[0]
                    lng_str, lat_str = coords_str.strip().split(' ')
                    geocercas.append({'id': str(geo.get('id')), 'name': nombre, 'lat': float(lat_str), 'lng': float(lng_str), 'radius': 800})
                except Exception: pass
    except Exception: pass
    return geocercas

HTML_INTERFACE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Histórico De Rutas Minuto a Minuto - Kowi</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>

    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; padding: 20px; margin: 0; }
        .card { max-width: 820px; margin: 0 auto; background: #ffffff; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); position: relative; z-index: 10; }
        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px; }
        .header-text { flex: 1; text-align: center; padding: 0 15px; }
        .header h2 { color: #1a252f; margin: 0 0 5px 0; font-size: 22px; }
        .header p { color: #7f8c8d; font-size: 13px; margin: 0; }
        .logo-img { max-height: 55px; max-width: 140px; object-fit: contain; }
        
        #loading_overlay {
            display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(255, 255, 255, 0.96); z-index: 9999;
            flex-direction: column; justify-content: center; align-items: center;
        }
        #loading_overlay video { max-width: 500px; width: 90%; border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,0.15); }
        .loading-text { margin-top: 25px; font-size: 22px; font-weight: bold; color: #2c3e50; }
        .loading-subtext { color: #e67e22; margin-top: 8px; font-size: 15px; font-weight: 600; text-align: center; }

        .main-container { display: flex; gap: 20px; }
        .presets-sidebar { width: 190px; border-right: 1px solid #eee; padding-right: 15px; display: flex; flex-direction: column; gap: 6px; }
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
        .btn-submit { width: 100%; background: #00a8ff; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; font-size: 14px; cursor: pointer; transition: background 0.2s; margin-top: 10px; }
        .btn-submit:hover { background: #008be3; }
        .btn-submit:disabled { background: #8cd6ff; cursor: not-allowed; }
        .status-msg { font-size: 13px; color: #e67e22; margin-top: 10px; text-align: center; font-weight: bold; }
        .retry-btn { font-size: 11px; color: #007bff; text-decoration: underline; cursor: pointer; margin-left: 8px; }
        hr { border: 0; border-top: 1px solid #dde2e5; margin: 5px 0; }
    </style>
</head>
<body>
    <div id="loading_overlay">
        <video autoplay loop muted playsinline>
            <source src="{{ url_for('static', filename='video_kowi.mp4') }}" type="video/mp4">
        </video>
        <div class="loading-text">Generando Reporte Ejecutivo...</div>
        <div class="loading-subtext" id="overlay_status">Conectando con servidores...</div>
    </div>

    <div class="card">
        <div class="header">
            <img src="{{ url_for('static', filename='logo_kowi.png') }}" class="logo-img" alt="Kowi">
            <div class="header-text">
                <h2>Histórico De Rutas Minuto a Minuto</h2>
                <p>Módulo Híbrido V8 (Fusión Satelital y CAN Bus)</p>
            </div>
            <img src="{{ url_for('static', filename='logo_idt.png') }}" class="logo-img" alt="IDT Tecnologías">
        </div>
        
        <div class="main-container">
            <div class="presets-sidebar">
                <label>Atajos de Fecha:</label>
                <button type="button" onclick="setRange('hoy')">Hoy</button>
                <button type="button" onclick="setRange('ayer')">Ayer</button>
                <button type="button" onclick="setRange('esta_semana')">Esta Semana</button>
                <button type="button" onclick="setRange('semana_anterior')">Semana anterior</button>
                <button type="button" onclick="setRange('ultimos_7_dias')">Últimos 7 días</button>
                <hr>
                <button type="button" onclick="setRange('este_mes')">Mes completo</button>
                <button type="button" onclick="setRange('mes_anterior')">Mes anterior</button>
            </div>

            <div class="form-content">
                <div class="form-group">
                    <label>Buscar / Seleccionar Unidad: <span class="retry-btn" onclick="cargarCatalogos()">🔄 Recargar Nube</span></label>
                    <select id="unit_select" style="width: 100%;">
                        <option value="">⏳ Cargando catálogo...</option>
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

                <div class="form-row">
                    <div class="form-group">
                        <label>Límite Velocidad General (km/h):</label>
                        <input type="number" id="limite_velocidad" value="80" min="1" max="150" required>
                    </div>
                    <div class="form-group">
                        <label>Filtro de Ralentí (mins):</label>
                        <input type="number" id="min_ralenti" value="2" min="1" max="60" required>
                    </div>
                </div>

                <div class="form-group">
                    <label>Geocercas (Sincronizadas desde Mapon):</label>
                    <select id="geofence_select" multiple="multiple" style="width: 100%;">
                        <option value="">⏳ Descargando API...</option>
                    </select>
                </div>
                <div class="form-group" id="speed_limits_container"></div>

                <button type="button" class="btn-submit" id="btn_submit" onclick="iniciarReporte()">📥 Generar reporte</button>
                <div id="status_msg" class="status-msg"></div>
            </div>
        </div>
    </div>

    <script>
        async function cargarCatalogos() {
            try {
                const [resUnits, resGeos] = await Promise.all([fetch('/api_unidades'), fetch('/api_geocercas_nube')]);
                const units = await resUnits.json();
                const geos = await resGeos.json();

                const selectUnit = $('#unit_select');
                selectUnit.empty().append(new Option('🔍 Escribe para buscar unidad...', ''));
                units.forEach(u => selectUnit.append(new Option(`${u.label || ''} ${u.number || ''}`.trim(), u.unit_id)));
                selectUnit.select2({ placeholder: "🔍 Escribe para buscar unidad...", width: '100%' });

                const selectGeo = $('#geofence_select');
                selectGeo.empty();
                
                geos.forEach(g => { selectGeo.append(new Option(g.name, g.id)); });
                selectGeo.select2({ placeholder: "Buscar geocerca para límite personalizado...", width: '100%' });
                $('#btn_submit').prop('disabled', false);
                $('#status_msg').html(`✅ Sincronizado: <b>${geos.length} geocercas</b> cargadas.`);
            } catch (e) { $('#status_msg').text("Error cargando catálogos de la API."); }
        }

        $('#geofence_select').on('change', function() {
            const container = $('#speed_limits_container');
            const selected = $(this).select2('data');
            container.empty();
            selected.forEach(s => {
                if(s.id) {
                    container.append(`
                        <div class="form-group" style="margin-top: 8px;">
                            <label>Límite en ${s.text} (km/h):</label>
                            <input type="number" class="geo-limit" data-id="${s.id}" data-name="${s.text}" value="30" style="width:100%; padding:9px; border:1px solid #dcdfe6; border-radius:6px; box-sizing:border-box; font-size:13px; color:#2c3e50;">
                        </div>
                    `);
                }
            });
        });

        async function iniciarReporte() {
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');
            const overlay = document.getElementById('loading_overlay');
            const overlayStatus = document.getElementById('overlay_status');
            
            const unitId = $('#unit_select').val();
            const unitText = $('#unit_select option:selected').text();
            if (!unitId) { alert("Por favor selecciona una unidad."); return; }
            
            btn.disabled = true;
            overlay.style.display = 'flex';
            overlayStatus.innerText = "⏳ Extrayendo Telemetría CAN y GPS...";

            const geoLimits = {};
            $('.geo-limit').each(function() {
                geoLimits[$(this).data('id')] = { limit: $(this).val(), name: $(this).data('name') };
            });

            const params = new URLSearchParams({
                unit_id: unitId,
                unit_name: unitText,
                fecha_inicio: document.getElementById('fecha_inicio').value,
                fecha_fin: document.getElementById('fecha_fin').value,
                hora_inicio: document.getElementById('hora_inicio').value,
                hora_fin: document.getElementById('hora_fin').value,
                limite_velocidad: document.getElementById('limite_velocidad').value,
                min_ralenti: document.getElementById('min_ralenti').value,
                geos: JSON.stringify(geoLimits)
            });

            try {
                const resInit = await fetch(`/iniciar_reporte?${params.toString()}`);
                const dataInit = await resInit.json();
                const taskId = dataInit.task_id;

                const interval = setInterval(async () => {
                    const sRes = await fetch(`/estado_reporte?task_id=${taskId}`);
                    const sData = await sRes.json();

                    if (sData.status === 'procesando') {
                        overlayStatus.innerText = `⏳ ${sData.msg}`;
                    } else if (sData.status === 'completado') {
                        clearInterval(interval);
                        overlay.style.display = 'none';
                        status.innerText = "¡Listo! Descargando reporte...";
                        window.location.href = `/descargar_reporte?task_id=${taskId}&unit_name=${encodeURIComponent(unitText)}`;
                        btn.disabled = false;
                    } else if (sData.status === 'error') {
                        clearInterval(interval);
                        overlay.style.display = 'none';
                        status.innerText = `❌ Error: ${sData.msg}`;
                        btn.disabled = false;
                        alert("Error en el reporte: " + sData.msg);
                    }
                }, 1500); 

            } catch (e) {
                overlay.style.display = 'none';
                status.innerText = "Error al iniciar el reporte.";
                btn.disabled = false;
            }
        }

        function setRange(type) {
            const now = new Date();
            let start = new Date();
            let end = new Date();
            const formatDate = (d) => {
                let month = '' + (d.getMonth() + 1), day = '' + d.getDate(), year = d.getFullYear();
                if (month.length < 2) month = '0' + month;
                if (day.length < 2) day = '0' + day;
                return [year, month, day].join('-');
            };
            if (type === 'hoy') { start = now; end = now; }
            else if (type === 'ayer') { start.setDate(now.getDate() - 1); end.setDate(now.getDate() - 1); }
            else if (type === 'esta_semana') { const day = now.getDay() || 7; start.setDate(now.getDate() - day + 1); end = now; }
            else if (type === 'semana_anterior') { const day = now.getDay() || 7; start.setDate(now.getDate() - day - 6); end.setDate(now.getDate() - day); }
            else if (type === 'ultimos_7_dias') { start.setDate(now.getDate() - 6); end = now; }
            else if (type === 'este_mes') { start = new Date(now.getFullYear(), now.getMonth(), 1); end = new Date(now.getFullYear(), now.getMonth() + 1, 0); }
            else if (type === 'mes_anterior') { start = new Date(now.getFullYear(), now.getMonth() - 1, 1); end = new Date(now.getFullYear(), now.getMonth(), 0); }
            
            document.getElementById('fecha_inicio').value = formatDate(start);
            document.getElementById('fecha_fin').value = formatDate(end);
        }

        setRange('hoy');
        cargarCatalogos();
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    global CACHE_GEOCERCAS
    if not CACHE_GEOCERCAS:
        CACHE_GEOCERCAS = cargar_geocercas_api()
    return render_template_string(HTML_INTERFACE)

@app.route('/api_unidades')
def api_unidades():
    try:
        res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY}, timeout=15)
        data = res.json()
        units_raw = data.get('data', {}).get('units', [])
        unidades_filtradas = [u for u in units_raw if not str(u.get('company_id', '')) or str(u.get('company_id', '')) == COMPANY_ID]
        return jsonify(unidades_filtradas)
    except Exception:
        return jsonify([]), 500

@app.route('/api_geocercas_nube')
def api_geocercas_nube():
    global CACHE_GEOCERCAS
    if not CACHE_GEOCERCAS:
        CACHE_GEOCERCAS = cargar_geocercas_api()
    lista_ordenada = sorted(CACHE_GEOCERCAS, key=lambda x: x['name'].lower())
    return jsonify(lista_ordenada)

@app.route('/iniciar_reporte')
def iniciar_reporte():
    task_id = str(uuid.uuid4())
    TASKS[task_id] = {'status': 'procesando', 'msg': 'Iniciando Extracción CAN y GPS...'}
    params = request.args.to_dict()
    thread = threading.Thread(target=procesar_reporte_bg, args=(task_id, params))
    thread.daemon = True
    thread.start()
    return jsonify({"task_id": task_id})

@app.route('/estado_reporte')
def estado_reporte():
    task_id = request.args.get('task_id')
    return jsonify(TASKS.get(task_id, {"status": "error", "msg": "Tarea no encontrada"}))

@app.route('/descargar_reporte')
def descargar_reporte():
    task_id = request.args.get('task_id')
    unit_name = request.args.get('unit_name', 'Reporte')
    if task_id in TASKS and 'file' in TASKS[task_id]:
        file_path = TASKS[task_id]['file']
        return send_file(file_path, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=f"Reporte_{unit_name}.xlsx")
    return "Reporte no disponible o ya fue descargado.", 404

def procesar_reporte_bg(task_id, params):
    try:
        def normalizar_fecha(f): return f if f else '2026-08-09'
        def normalizar_hora(h, es_fin=False):
            if not h: return "23:59:59" if es_fin else "00:00:00"
            return h if len(h) == 8 else h + ":00"
            
        unit_id = params.get('unit_id', '868807').replace("ID:", "").strip()
        unit_name = params.get('unit_name', unit_id).split("(ID:")[0].strip()
        f_in = params.get('fecha_inicio', '2026-08-09')
        f_fin = params.get('fecha_fin', f_in)
        hora_inicio = normalizar_hora(params.get('hora_inicio', '00:00:00'))
        hora_fin = normalizar_hora(params.get('hora_fin', '23:59:59'), True)
        limite_velocidad_gral = int(params.get('limite_velocidad', 80))
        min_ralenti = int(params.get('min_ralenti', 2))
        try: 
            geo_limits = json.loads(params.get('geos', '{}'))
        except Exception: 
            geo_limits = {}

        global CACHE_GEOCERCAS
        if not CACHE_GEOCERCAS:
            CACHE_GEOCERCAS = cargar_geocercas_api()

        def obtener_geocerca(lat, lng):
            if lat == 0 and lng == 0: return None, "Fuera de geocerca"
            for g in CACHE_GEOCERCAS:
                if calcular_distancia(lat, lng, g['lat'], g['lng']) <= g['radius']:
                    return g['id'], g['name']
            return None, "Fuera de geocerca"

        def parse_iso(iso_str):
            if not iso_str: return None
            try: return datetime.strptime(str(iso_str).replace('Z', '').split('.')[0], '%Y-%m-%d %H:%M:%S') + timedelta(hours=TIMEZONE_OFFSET)
            except ValueError:
                try: return datetime.strptime(str(iso_str).replace('Z', '').split('.')[0], '%Y-%m-%dT%H:%M:%S') + timedelta(hours=TIMEZONE_OFFSET)
                except Exception: return None

        def fetch_exact_point(dt):
            utc_str = (dt - timedelta(hours=TIMEZONE_OFFSET)).strftime('%Y-%m-%dT%H:%M:%SZ')
            url = f"{BASE_URL}/unit_data/history_point.json?key={API_KEY}&unit_id={unit_id}&datetime={utc_str}&include[]=position"
            try:
                r = requests.get(url, timeout=6)
                data = r.json()
                units = data.get('data', {}).get('units', [])
                if units:
                    pos = units[0].get('position', {}).get('value', {})
                    if pos and 'lat' in pos and 'lng' in pos:
                        return {'dt': dt, 'lat': float(pos['lat']), 'lng': float(pos['lng'])}
            except Exception: pass
            return None

        dt_inicio_req = datetime.strptime(f"{f_in} {hora_inicio}", "%Y-%m-%d %H:%M:%S")
        dt_fin_req = datetime.strptime(f"{f_fin} {hora_fin}", "%Y-%m-%d %H:%M:%S")
        
        utc_start_str = (dt_inicio_req - timedelta(hours=TIMEZONE_OFFSET)).strftime("%Y-%m-%dT%H:%M:%SZ")
        utc_end_str = (dt_fin_req - timedelta(hours=TIMEZONE_OFFSET)).strftime("%Y-%m-%dT%H:%M:%SZ")
        
        # ==============================================================
        # 1. EXTRACCIÓN DE DATOS CAN BUS (Rendimiento, Combustible, Motor)
        # ==============================================================
        TASKS[task_id]['msg'] = "Extrayendo métricas de CAN Bus..."
        can_data = {
            "has_can": False,
            "dist_inicial": 0, "dist_final": 0,
            "fuel_inicial": 0, "fuel_final": 0,
            "engine_hrs_inicial": 0, "engine_hrs_final": 0,
            "max_temp": 0
        }
        
        url_can_point = f"{BASE_URL}/unit_data/can_point.json"
        
        try:
            # Punto Inicial CAN
            res_ini = requests.get(url_can_point, params={"key": API_KEY, "unit_id": unit_id, "datetime": utc_start_str}, timeout=10).json()
            # Punto Final CAN
            res_fin = requests.get(url_can_point, params={"key": API_KEY, "unit_id": unit_id, "datetime": utc_end_str}, timeout=10).json()
            
            ini_unit = res_ini.get('data', {}).get('units', [{}])[0]
            fin_unit = res_fin.get('data', {}).get('units', [{}])[0]
            
            if fin_unit and 'total_distance' in fin_unit:
                can_data["has_can"] = True
                
                can_data["dist_inicial"] = float(ini_unit.get('total_distance', {}).get('value', 0))
                can_data["dist_final"] = float(fin_unit.get('total_distance', {}).get('value', 0))
                
                can_data["fuel_inicial"] = float(ini_unit.get('total_fuel', {}).get('value', 0))
                can_data["fuel_final"] = float(fin_unit.get('total_fuel', {}).get('value', 0))
                
                can_data["engine_hrs_inicial"] = float(ini_unit.get('total_engine_hours', {}).get('value', 0))
                can_data["engine_hrs_final"] = float(fin_unit.get('total_engine_hours', {}).get('value', 0))
                
        except Exception: pass
        
        # Buscar temperatura máxima del motor (generalmente de temperature.json)
        try:
            url_temp = f"{BASE_URL}/unit_data/temperature.json"
            res_temp = requests.get(url_temp, params={"key": API_KEY, "unit_id": unit_id, "from": utc_start_str, "till": utc_end_str}, timeout=15).json()
            sensors = res_temp.get('data', {}).get('units', [{}])[0].get('sensors', [])
            max_t = 0
            for sensor in sensors:
                for t in sensor.get('temperatures', []):
                    val = float(t.get('value', 0))
                    if val > max_t: max_t = val
            can_data["max_temp"] = max_t
        except Exception: pass

        # ==============================================================
        # 2. EXTRACCIÓN DE RUTAS Y GPS
        # ==============================================================
        url_route = f"{BASE_URL}/route/list.json"
        url_ign = f"{BASE_URL}/unit_data/ignitions.json"
        
        tramos_reales = []
        parsed_idles = []
        eventos_ignicion = []
        puntos_maestros_reales = []

        current_start = dt_inicio_req
        chunk_days = 2 
        
        while current_start < dt_fin_req:
            current_end = current_start + timedelta(days=chunk_days)
            if current_end > dt_fin_req: current_end = dt_fin_req
                
            TASKS[task_id]['msg'] = f"Analizando Encendidos y Movimiento GPS ({current_start.strftime('%d %b')} - {current_end.strftime('%d %b')})..."
            
            chunk_start_utc = (current_start - timedelta(hours=TIMEZONE_OFFSET)).strftime("%Y-%m-%dT%H:%M:%SZ")
            chunk_end_utc = (current_end - timedelta(hours=TIMEZONE_OFFSET)).strftime("%Y-%m-%dT%H:%M:%SZ")
            
            req_params = {"key": API_KEY, "unit_id": unit_id, "from": chunk_start_utc, "till": chunk_end_utc, "include": "metrics,idles,routes,polyline,speed"}
            req_ign_params = {"key": API_KEY, "unit_id": unit_id, "from": chunk_start_utc, "till": chunk_end_utc}
            
            try:
                # Igniciones Reales
                try:
                    res_ign = requests.get(url_ign, params=req_ign_params, timeout=45).json()
                    ign_list = res_ign.get('data', {}).get('units', [{}])[0].get('ignitions', [])
                    for ign in ign_list:
                        on_dt = parse_iso(ign.get('on'))
                        off_dt = parse_iso(ign.get('off'))
                        if on_dt and dt_inicio_req <= on_dt <= dt_fin_req:
                            eventos_ignicion.append({'dt': on_dt, 'evento': 'Motor encendido', 'tipo': 'on', 'detalle': 'Ignición activada'})
                        if off_dt and dt_inicio_req <= off_dt <= dt_fin_req:
                            eventos_ignicion.append({'dt': off_dt, 'evento': 'Motor apagado', 'tipo': 'off', 'detalle': 'Llave cerrada'})
                except Exception: pass

                # Rutas y Polilíneas
                response = requests.get(url_route, params=req_params, timeout=45)
                data = response.json()
                
                unit_data = data.get('data', {}).get('units', [])[0] if data.get('data', {}).get('units') else {}
                for idl in unit_data.get('idles', []):
                    s_dt = parse_iso(idl.get('start', {}).get('time'))
                    e_dt = parse_iso(idl.get('end', {}).get('time'))
                    if s_dt and e_dt: parsed_idles.append({'dt_ini': s_dt, 'dt_fin': e_dt})

                rutas_encontradas = []
                def extraer_tramos(obj):
                    if isinstance(obj, dict):
                        tipo = str(obj.get('type', '')).lower()
                        if tipo == 'route': rutas_encontradas.append(obj)
                        for k, v in obj.items():
                            if isinstance(v, (dict, list)): extraer_tramos(v)
                    elif isinstance(obj, list):
                        for item in obj:
                            if isinstance(item, (dict, list)): extraer_tramos(item)
                extraer_tramos(data)
                
                for item in rutas_encontradas:
                    dt_ini = parse_iso(item.get('start', {}).get('time', ''))
                    dt_fin = parse_iso(item.get('end', {}).get('time', ''))
                    if not dt_ini or not dt_fin: continue
                        
                    dist_km = float(item.get('distance', 0)) / 1000.0
                    
                    poly_str = item.get('polyline', '')
                    speed_str = item.get('speed', '')
                    
                    if poly_str and speed_str:
                        coords = decode_polyline_2d(poly_str)
                        vel_offsets = decode_mapon_speed_string(speed_str)
                        
                        min_len = min(len(coords), len(vel_offsets))
                        for idx in range(min_len):
                            offset_seg, real_speed = vel_offsets[idx]
                            pt_time = dt_ini + timedelta(seconds=offset_seg)
                            puntos_maestros_reales.append({
                                'dt': pt_time,
                                'lat': coords[idx]['lat'],
                                'lng': coords[idx]['lng'],
                                'speed': float(real_speed) 
                            })

                    tramos_reales.append({
                        'dt_ini': dt_ini, 'dt_fin': dt_fin, 'distancia': dist_km, 'tipo': 'route',
                        'max_speed': float(item.get('metrics', {}).get('max_speed', 110))
                    })
            except Exception: pass 
            current_start = current_end

        eventos_ignicion.sort(key=lambda x: x['dt'])
        puntos_maestros_reales.sort(key=lambda x: x['dt'])
        maestro_dts = [p['dt'] for p in puntos_maestros_reales]

        cuadricula_maestra = []
        c_time = dt_inicio_req
        while c_time <= dt_fin_req:
            cuadricula_maestra.append(c_time)
            c_time += timedelta(minutes=1)

        for ev in eventos_ignicion:
            cuadricula_maestra.append(ev['dt'])
            
        tiempo_ral_reportado_seg = 0
        eventos_ralenti = []
        for idl in parsed_idles:
            dur = (idl['dt_fin'] - idl['dt_ini']).total_seconds()
            if dur >= min_ralenti * 60:
                tiempo_ral_reportado_seg += dur
                ral_dt = idl['dt_ini'] + timedelta(seconds=1)
                eventos_ralenti.append({'dt': ral_dt, 'evento': 'Ralentí', 'detalle': f"Detenido por: {int(dur//60)} mins"})
                cuadricula_maestra.append(ral_dt)

        cuadricula_maestra = sorted(list(set(cuadricula_maestra)))

        def is_ignition_on(dt):
            estado = False
            for ev in eventos_ignicion:
                if ev['dt'] <= dt: estado = (ev['tipo'] == 'on')
            return estado

        def is_in_route(dt):
            for t in tramos_reales:
                if t['dt_ini'] <= dt <= t['dt_fin']: return True, t
            return False, None

        minutos_a_descargar = [dt for dt in cuadricula_maestra if is_ignition_on(dt) or is_in_route(dt)[0]]
        for ev in eventos_ignicion: minutos_a_descargar.append(ev['dt'])
        for ev in eventos_ralenti: minutos_a_descargar.append(ev['dt'])
        minutos_a_descargar = sorted(list(set(minutos_a_descargar)))

        TASKS[task_id]['msg'] = f"Sincronizando puntos satelitales (Gobernador)..."

        puntos_exitosos = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=15) as executor:
            resultados = executor.map(fetch_exact_point, minutos_a_descargar)
            for res in resultados:
                if res: puntos_exitosos.append(res)
                
        puntos_exitosos.sort(key=lambda x: x['dt'])
        puntos_dict = {p['dt']: p for p in puntos_exitosos}

        TASKS[task_id]['msg'] = "Procesando matriz de reporte y filtrando excesos..."

        filas_brutas = []
        tiempo_mov_seg = 0
        tiempo_exceso_geo_seg = 0
        tiempo_apagado_seg = 0
        distancia_total_gps_km = sum([t['distancia'] for t in tramos_reales if t['tipo'] == 'route'])
        
        last_lat, last_lng = 27.19, -109.55
        if puntos_exitosos: last_lat, last_lng = puntos_exitosos[0]['lat'], puntos_exitosos[0]['lng']
        last_dt_punto = None

        for i, dt in enumerate(cuadricula_maestra):
            ign_on = is_ignition_on(dt)
            en_ruta, tramo_actual = is_in_route(dt)
            
            punto = puntos_dict.get(dt)
            if punto:
                curr_lat, curr_lng = punto['lat'], punto['lng']
            else:
                curr_lat, curr_lng = last_lat, last_lng

            current_speed = 0.0
            seg_transcurridos = 60 if i > 0 else 0

            # FILTRO ANTI-PICOS CON GOBERNADOR
            if en_ruta and punto and ign_on:
                if len(puntos_maestros_reales) > 0:
                    idx = bisect.bisect_left(maestro_dts, dt)
                    closest_p = None
                    min_diff = float('inf')
                    for check_idx in [idx-1, idx, idx+1]:
                        if 0 <= check_idx < len(puntos_maestros_reales):
                            diff = abs((puntos_maestros_reales[check_idx]['dt'] - dt).total_seconds())
                            if diff < min_diff:
                                min_diff = diff
                                closest_p = puntos_maestros_reales[check_idx]
                    
                    if closest_p and min_diff <= 90:
                        current_speed = float(closest_p['speed'])
                    elif last_dt_punto is not None:
                        dist_mts = calcular_distancia(last_lat, last_lng, curr_lat, curr_lng)
                        seg_diff = max((dt - last_dt_punto).total_seconds(), 1)
                        raw_speed = (dist_mts / seg_diff) * 3.6
                        max_oficial = float(tramo_actual.get('max_speed', 110)) if tramo_actual else 110
                        if max_oficial <= 0: max_oficial = 110
                        current_speed = min(raw_speed, max_oficial)
                elif last_dt_punto is not None:
                    dist_mts = calcular_distancia(last_lat, last_lng, curr_lat, curr_lng)
                    seg_diff = max((dt - last_dt_punto).total_seconds(), 1)
                    raw_speed = (dist_mts / seg_diff) * 3.6
                    max_oficial = float(tramo_actual.get('max_speed', 110)) if tramo_actual else 110
                    if max_oficial <= 0: max_oficial = 110
                    current_speed = min(raw_speed, max_oficial)

            if current_speed < 3 or not ign_on: 
                current_speed = 0.0

            current_speed = round(current_speed, 1)

            if punto:
                last_lat, last_lng = curr_lat, curr_lng
                last_dt_punto = dt

            if current_speed > 0:
                tiempo_mov_seg += seg_transcurridos
            else:
                if ign_on: pass # El ralenti ya esta contabilizado por idles_oficiales
                else: tiempo_apagado_seg += seg_transcurridos

            geo_id, geo_name = obtener_geocerca(curr_lat, curr_lng)
            evento = ""
            detalle = "-"
            
            es_evento_motor = next((e for e in eventos_ignicion if e['dt'] == dt), None)
            es_evento_ralenti = next((e for e in eventos_ralenti if e['dt'] == dt), None)
            
            if es_evento_motor:
                evento = es_evento_motor['evento']
                detalle = es_evento_motor['detalle']
            elif es_evento_ralenti:
                evento = es_evento_ralenti['evento']
                detalle = es_evento_ralenti['detalle']
            elif current_speed > 0:
                limite_aplicable = limite_velocidad_gral
                if geo_name != "Fuera de geocerca":
                    for gid, gdata in geo_limits.items():
                        if gdata['name'].lower() == geo_name.lower():
                            try: limite_aplicable = int(gdata['limit'])
                            except Exception: pass
                            break
                    if current_speed > limite_aplicable:
                        evento = f"Exceso en {geo_name}"
                        detalle = f"Vel: {current_speed} (Límite: {limite_aplicable})"
                        tiempo_exceso_geo_seg += 60
                elif current_speed > limite_velocidad_gral:
                    evento = "Exceso de velocidad"
                    detalle = f"Vel: {current_speed} (Límite: {limite_velocidad_gral})"

            filas_brutas.append({
                'fecha': dt, 'origen': 'Zona Operativa', 'velocidad': current_speed, 
                'evento': evento, 'detalle': detalle, 'lat': curr_lat, 'lng': curr_lng, 'geocerca': geo_name
            })

        TASKS[task_id]['msg'] = "Estructurando reporte Ejecutivo (Excel)..."

        def calc_hrs_mins(segundos): return int(segundos // 3600), int((segundos % 3600) // 60)
        mov_hrs, mov_mins = calc_hrs_mins(tiempo_mov_seg)
        ral_hrs, ral_mins = calc_hrs_mins(tiempo_ral_reportado_seg)
        exceso_geo_hrs, exceso_geo_mins = calc_hrs_mins(tiempo_exceso_geo_seg)
        muerto_hrs, muerto_mins = calc_hrs_mins(tiempo_apagado_seg)

        motor_hrs_gps, motor_mins_gps = calc_hrs_mins(tiempo_mov_seg + tiempo_ral_reportado_seg)

        vels_mov = [f['velocidad'] for f in filas_brutas if f['velocidad'] > 0]
        max_vel = max(vels_mov) if vels_mov else 0
        prom_vel = sum(vels_mov) / len(vels_mov) if vels_mov else 0

        # Matemáticas CAN Bus
        can_dist_total = can_data['dist_final'] - can_data['dist_inicial']
        can_fuel_total = can_data['fuel_final'] - can_data['fuel_inicial']
        can_horas_motor_dec = can_data['engine_hrs_final'] - can_data['engine_hrs_inicial']
        can_horas = int(can_horas_motor_dec)
        can_mins = int((can_horas_motor_dec - can_horas) * 60)
        
        rendimiento_can = (can_dist_total / can_fuel_total) if can_fuel_total > 0 else 0

        # ==============================================================
        # DIBUJADO DEL EXCEL (DISEÑO DOBLE ENCABEZADO)
        # ==============================================================
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Histórico Ejecutivo"

        # Título
        ws.cell(row=1, column=4, value="Reporte Analítico Minuto a Minuto").font = Font(bold=True, size=15)
        ws.cell(row=3, column=3, value="Vehículo:").font = Font(bold=True)
        ws.cell(row=3, column=4, value=str(unit_name))

        # Color de las cajas
        fill_gps = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        fill_can = PatternFill(start_color="EBF1DE", end_color="EBF1DE", fill_type="solid")
        thick_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))

        # ================= SECCIÓN GPS (Izquierda) =================
        ws.cell(row=5, column=1, value="INFORMACIÓN GPS").font = Font(bold=True, color="1F497D")
        ws.cell(row=5, column=1).fill = fill_gps
        
        ws.cell(row=6, column=1, value="Recorrido Aprox:").font = Font(bold=True)
        ws.cell(row=6, column=2, value=f"{round(distancia_total_gps_km, 2)} km")
        ws.cell(row=6, column=3, value="Tiempo en Movimiento:").font = Font(bold=True)
        ws.cell(row=6, column=4, value=f"{mov_hrs} hrs {mov_mins} mins")

        ws.cell(row=7, column=1, value="Velocidad Máxima Oficial:").font = Font(bold=True)
        ws.cell(row=7, column=2, value=f"{round(max_vel, 1)} km/h")
        ws.cell(row=7, column=3, value="Ralentí (Motor estático):").font = Font(bold=True)
        ws.cell(row=7, column=4, value=f"{ral_hrs} hrs {ral_mins} mins").font = Font(color="FF0000")

        ws.cell(row=8, column=1, value="Velocidad Promedio:").font = Font(bold=True)
        ws.cell(row=8, column=2, value=f"{round(prom_vel, 1)} km/h")
        ws.cell(row=8, column=3, value="Exceso en Geocercas:").font = Font(bold=True)
        ws.cell(row=8, column=4, value=f"{exceso_geo_hrs} hrs {exceso_geo_mins} mins").font = Font(color="FF0000", bold=True)

        ws.cell(row=9, column=1, value="Horas Motor (Aprox GPS):").font = Font(bold=True)
        ws.cell(row=9, column=2, value=f"{motor_hrs_gps} hrs {motor_mins_gps} mins")
        ws.cell(row=9, column=3, value="Motor Apagado (Sin GPS):").font = Font(bold=True)
        ws.cell(row=9, column=4, value=f"{muerto_hrs} hrs {muerto_mins} mins")

        # ================= SECCIÓN CAN BUS (Derecha) =================
        ws.cell(row=5, column=6, value="INFORMACIÓN EXTRAÍDA DE LA UNIDAD (CAN BUS)").font = Font(bold=True, color="4F6228")
        ws.cell(row=5, column=6).fill = fill_can
        
        if can_data["has_can"]:
            ws.cell(row=6, column=6, value="Recorrido Tablero (Odo):").font = Font(bold=True)
            ws.cell(row=6, column=7, value=f"{round(can_dist_total, 2)} km")
            
            ws.cell(row=7, column=6, value="Combustible Quemado:").font = Font(bold=True)
            ws.cell(row=7, column=7, value=f"{round(can_fuel_total, 2)} L").font = Font(color="FF0000", bold=True)
            
            ws.cell(row=8, column=6, value="Rendimiento del Viaje:").font = Font(bold=True)
            ws.cell(row=8, column=7, value=f"{round(rendimiento_can, 2)} km/L").font = Font(color="008000", bold=True)
            
            ws.cell(row=9, column=6, value="Horómetro Interno:").font = Font(bold=True)
            ws.cell(row=9, column=7, value=f"{can_horas} hrs {can_mins} mins")
            
            ws.cell(row=10, column=6, value="Temp. Max Alcanzada:").font = Font(bold=True)
            ws.cell(row=10, column=7, value=f"{can_data['max_temp']} °C")
        else:
            ws.cell(row=6, column=6, value="Recorrido Tablero (Odo):").font = Font(bold=True)
            ws.cell(row=6, column=7, value="0 km (Sin CAN)")
            ws.cell(row=7, column=6, value="Combustible Quemado:").font = Font(bold=True)
            ws.cell(row=7, column=7, value="0 L (Sin CAN)")
            ws.cell(row=8, column=6, value="Rendimiento del Viaje:").font = Font(bold=True)
            ws.cell(row=8, column=7, value="0 km/L (Sin CAN)")
            ws.cell(row=9, column=6, value="Horómetro Interno:").font = Font(bold=True)
            ws.cell(row=9, column=7, value="0 hrs 0 mins (Sin CAN)")
            ws.cell(row=10, column=6, value="Temp. Max Alcanzada:").font = Font(bold=True)
            ws.cell(row=10, column=7, value="0 °C (Sin CAN)")

        # Datos Generales (Derecha Arriba)
        ws.cell(row=3, column=6, value="Fecha Inicial:").font = Font(bold=True)
        ws.cell(row=3, column=7, value=f"{f_in} {hora_inicio}")
        ws.cell(row=4, column=6, value="Fecha Final:").font = Font(bold=True)
        ws.cell(row=4, column=7, value=f"{f_in} {hora_fin}")

        # Estilizar Tabla Principal
        headers = ["Vehículo", "Fecha", "Dirección", "Ciudad", "Velocidad (Km/h)", "Evento", "Detalle", "Geocerca", "Mapa", "Longitud", "Latitud"]
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=12, column=col_idx, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        row_idx = 13
        for f in filas_brutas:
            ciudad = "Hermosillo" if "Hermosillo" in f['origen'] else ("Navojoa" if "Navojoa" in f['origen'] or "Pueblo Mayo" in f['origen'] else ("Guaymas" if "Guaymas" in f['origen'] else "Zona Operativa"))

            ws.cell(row=row_idx, column=1, value=str(unit_name))
            ws.cell(row=row_idx, column=2, value=f['fecha'].strftime('%Y-%m-%d %H:%M:%S'))
            ws.cell(row=row_idx, column=3, value=f['origen'])
            ws.cell(row=row_idx, column=4, value=ciudad)
            ws.cell(row=row_idx, column=5, value=f['velocidad'])
            ws.cell(row=row_idx, column=6, value=f['evento'])
            ws.cell(row=row_idx, column=7, value=f['detalle'])
            
            geo_cell = ws.cell(row=row_idx, column=8, value=f['geocerca'])
            if f['geocerca'] != "Fuera de geocerca": geo_cell.font = Font(color="008000", bold=True)
            if "Exceso" in f['evento'] or "Motor" in f['evento'] or "Ralentí" in f['evento']: ws.cell(row=row_idx, column=6).font = Font(color="FF0000", bold=True)
            
            map_cell = ws.cell(row=row_idx, column=9, value="mapa")
            map_cell.hyperlink = f"https://www.google.com/maps?q={f['lat']},{f['lng']}"
            map_cell.font = Font(color="0000FF", underline="single")
            map_cell.alignment = Alignment(horizontal="center")
            
            ws.cell(row=row_idx, column=10, value=round(f['lng'], 6))
            ws.cell(row=row_idx, column=11, value=round(f['lat'], 6))
            row_idx += 1

        if len(filas_brutas) == 0:
            ws.cell(row=13, column=1, value="No se encontraron datos. Verifique el periodo seleccionado.")

        fd, path = tempfile.mkstemp(suffix=".xlsx")
        with os.fdopen(fd, 'wb') as f:
            wb.save(f)
            
        TASKS[task_id]['file'] = path
        TASKS[task_id]['status'] = 'completado'
        
    except Exception as e:
        import traceback
        TASKS[task_id]['status'] = 'error'
        TASKS[task_id]['msg'] = f"Falló el procesamiento interno: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
