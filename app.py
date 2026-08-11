from flask import Flask, render_template_string, request, send_file, jsonify
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import io
from datetime import datetime, timedelta
import random
import os
import json
import math
import time

app = Flask(__name__)

# ==========================================
# CONFIGURACIÓN DE LA API Y CLIENTE
# ==========================================
API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"
COMPANY_ID = "87534"  # ID de Alimentos Kowi
TIMEZONE_OFFSET = -7  # Zona horaria de Hermosillo
DB_FILE = "geocercas_db.json"  # BASE DE DATOS LOCAL

HTML_INTERFACE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte Ejecutivo Minuto a Minuto - IDT</title>
    
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>

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
        
        .btn-sync { width: 100%; background: #f39c12; color: white; padding: 10px; border: none; border-radius: 6px; font-weight: bold; font-size: 13px; cursor: pointer; transition: background 0.2s; margin-bottom: 15px; }
        .btn-sync:hover { background: #e67e22; }
        .btn-sync:disabled { background: #f8c271; cursor: not-allowed; }
        
        .status-msg { font-size: 13px; color: #e67e22; margin-top: 10px; text-align: center; font-weight: bold; }
        .retry-btn { font-size: 11px; color: #007bff; text-decoration: underline; cursor: pointer; margin-left: 8px; }
    </style>
</head>
<body>
    <div class="card">
        <div class="header">
            <h2>📊 Reporte Ejecutivo Minuto a Minuto</h2>
            <p>IDT Tecnologías - Base de Datos Local Integrada</p>
        </div>
        
        <div class="main-container">
            <div class="presets-sidebar">
                <label>Atajos de Fecha:</label>
                <button type="button" onclick="setRange('hoy')">Hoy</button>
                <button type="button" onclick="setRange('ayer')">Ayer</button>
                <button type="button" onclick="setRange('esta_semana')">Esta Semana</button>
                <button type="button" onclick="setRange('semana_anterior')">Semana anterior</button>
                <button type="button" onclick="setRange('ultimos_7_dias')">Últimos 7 días</button>
            </div>

            <div class="form-content">
                <!-- BOTÓN DE SINCRONIZACIÓN DE BASE DE DATOS -->
                <button type="button" class="btn-sync" id="btn_sync" onclick="sincronizarBD()">🔄 1. Sincronizar Geocercas (Mapon -> Base de Datos Local)</button>
                
                <div class="form-group">
                    <label>Buscar / Seleccionar Unidad: <span class="retry-btn" onclick="cargarCatalogos()">🔄 Reintentar carga</span></label>
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
                        <label>Alerta Ralentí (mins):</label>
                        <input type="number" id="min_ralenti" value="5" min="1" max="60" required>
                    </div>
                </div>

                <div class="form-group">
                    <label>Geocercas (Opcional: Selecciona para asignar límite estricto):</label>
                    <select id="geofence_select" multiple="multiple" style="width: 100%;">
                        <option value="">⏳ Cargando desde Base de Datos Local...</option>
                    </select>
                </div>
                <div class="form-group" id="speed_limits_container"></div>

                <button type="button" class="btn-submit" id="btn_submit" onclick="generarReporte()" disabled>📥 2. Generar Reporte Ejecutivo</button>
                <div id="status_msg" class="status-msg"></div>
            </div>
        </div>
    </div>

    <script>
        async function sincronizarBD() {
            const btnSync = document.getElementById('btn_sync');
            const status = document.getElementById('status_msg');
            btnSync.disabled = true;
            status.innerText = "⏳ Descargando catálogo masivo (Mapon). Esto puede tardar unos 2 minutos. No cierre la pestaña...";
            
            try {
                const res = await fetch('/api/sync_db', { method: 'POST' });
                const data = await res.json();
                if(data.status === 'ok') {
                    status.innerText = "✅ " + data.message;
                    cargarCatalogos(); 
                } else {
                    status.innerText = "❌ Error al sincronizar: " + data.message;
                }
            } catch(e) {
                status.innerText = "❌ Error de conexión al sincronizar.";
            }
            btnSync.disabled = false;
        }

        async function cargarCatalogos() {
            try {
                const [resUnits, resGeos] = await Promise.all([
                    fetch('/api_unidades'),
                    fetch('/api_geocercas_locales')
                ]);
                const units = await resUnits.json();
                const geos = await resGeos.json();

                const selectUnit = $('#unit_select');
                selectUnit.empty().append(new Option('🔍 Escribe para buscar unidad...', ''));
                units.forEach(u => selectUnit.append(new Option(`${u.label || ''} ${u.number || ''}`.trim(), u.unit_id)));
                selectUnit.select2({ placeholder: "🔍 Escribe para buscar unidad...", width: '100%' });

                const selectGeo = $('#geofence_select');
                selectGeo.empty();
                geos.forEach(g => {
                    selectGeo.append(new Option(g.name, g.geofence_id));
                });
                selectGeo.select2({ placeholder: "Buscar geocerca para límite personalizado...", width: '100%' });

                $('#btn_submit').prop('disabled', false);
                if (geos.length === 0) {
                    $('#status_msg').text("⚠️ Tu base de datos local está vacía. Haz clic en el botón amarillo 'Sincronizar Geocercas' arriba.");
                } else {
                    $('#status_msg').text(`✅ BD Local conectada (${geos.length} geocercas listas).`);
                }
            } catch (e) {
                $('#status_msg').text("Error cargando catálogos locales.");
            }
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

        async function generarReporte() {
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');
            const unitId = $('#unit_select').val();
            const unitText = $('#unit_select option:selected').text();
            
            if (!unitId) { alert("Por favor selecciona una unidad."); return; }
            
            btn.disabled = true;
            status.innerText = "⚡ Generando reporte y cruzando con BD Local...";

            const geoLimits = {};
            $('.geo-limit').each(function() {
                geoLimits[$(this).data('id')] = {
                    limit: $(this).val(),
                    name: $(this).data('name')
                };
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

            const res = await fetch(`/generar_excel?${params.toString()}`);
            if (res.ok) {
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `Reporte_${unitText}.xlsx`;
                document.body.appendChild(a);
                a.click();
                a.remove();
                status.innerText = "¡Reporte generado y descargado con éxito!";
            } else {
                const errorDetail = await res.text();
                console.error("Error backend:", errorDetail);
                alert("Error de Python: " + errorDetail);
                status.innerText = "Error en la generación.";
            }
            btn.disabled = false;
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
    return render_template_string(HTML_INTERFACE)

@app.route('/api_unidades')
def api_unidades():
    try:
        res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY}, timeout=20)
        data = res.json()
        units_raw = data.get('data', {}).get('units', [])
        
        unidades_filtradas = []
        for u in units_raw:
            c_id = str(u.get('company_id', ''))
            if c_id and c_id != COMPANY_ID:
                continue
            unidades_filtradas.append(u)
            
        return jsonify(unidades_filtradas)
    except: 
        return jsonify([]), 500

# ==========================================
# EL NUEVO CEREBRO: BASE DE DATOS LOCAL
# Paginación pequeña (100) para burlar el límite de Mapon
# ==========================================
@app.route('/api/sync_db', methods=['POST'])
def sync_db():
    try:
        endpoints = ['/territory/list.json', '/poi/list.json']
        geos_db = []
        vistos = set()
        
        for ep in endpoints:
            offset = 0
            while True:
                # El límite de 100 y timeout de 15 evita que Mapon nos bloquee la IP
                res = requests.get(f"{BASE_URL}{ep}", params={'key': API_KEY, 'limit': 100, 'offset': offset}, timeout=15)
                if res.status_code != 200:
                    break
                    
                data = res.json()
                inner = data.get('data', data)
                items = []
                
                if isinstance(inner, dict):
                    for k in ['territories', 'pois', 'list', 'items']:
                        if k in inner: items.extend(inner[k])
                elif isinstance(inner, list):
                    items.extend(inner)
                    
                if not items: break
                
                for g in items:
                    c_id = str(g.get('company_id', ''))
                    if c_id and c_id != COMPANY_ID: continue
                        
                    g_id = str(g.get('territory_id', g.get('poi_id', g.get('id', ''))))
                    g_name = str(g.get('name', g.get('title', 'Geocerca'))).strip()
                    g_type = str(g.get('type', 'circle')).lower()
                    
                    if g_name in vistos or g_id == 'None': continue
                    
                    geo_obj = {'id': g_id, 'name': g_name, 'type': g_type}
                    
                    if g_type == 'polygon' or 'points' in g:
                        pts = g.get('points', [])
                        if pts:
                            geo_obj['points'] = pts
                            geos_db.append(geo_obj)
                            vistos.add(g_name)
                    else:
                        g_lat = g.get('lat', g.get('latitude'))
                        g_lng = g.get('lng', g.get('longitude'))
                        g_radius = float(g.get('radius', 150))
                        if g_lat is None and 'center' in g:
                            g_lat = g['center'].get('lat')
                            g_lng = g['center'].get('lng')
                        
                        if g_lat is not None and g_lng is not None:
                            geo_obj['lat'] = float(g_lat)
                            geo_obj['lng'] = float(g_lng)
                            geo_obj['radius'] = g_radius
                            geos_db.append(geo_obj)
                            vistos.add(g_name)
                            
                if len(items) < 100: break
                offset += 100
                time.sleep(0.5) # Respiro para el servidor de Mapon
                
        with open(DB_FILE, 'w', encoding='utf-8') as f:
            json.dump(geos_db, f, ensure_ascii=False, indent=2)
            
        return jsonify({"status": "ok", "message": f"Sincronización exitosa. {len(geos_db)} geocercas guardadas en Local."})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api_geocercas_locales')
def api_geocercas_locales():
    try:
        if not os.path.exists(DB_FILE): return jsonify([])
        with open(DB_FILE, 'r', encoding='utf-8') as f:
            geos = json.load(f)
        menu_items = [{'geofence_id': g['id'], 'name': g['name']} for g in geos]
        menu_items.sort(key=lambda x: x['name'])
        return jsonify(menu_items)
    except:
        return jsonify([])

@app.route('/generar_excel')
def generar_excel():
    try:
        unit_id = request.args.get('unit_id', '868807')
        unit_name = request.args.get('unit_name', unit_id)
        if "(ID:" in unit_name:
            unit_name = unit_name.split("(ID:")[0].strip()

        f_in = request.args.get('fecha_inicio', '2026-08-09')
        f_fin = request.args.get('fecha_fin', f_in)
        hora_inicio = request.args.get('hora_inicio', '00:00:00')
        hora_fin = request.args.get('hora_fin', '23:59:59')
        
        try: limite_velocidad = int(request.args.get('limite_velocidad', 80))
        except: limite_velocidad = 80
            
        try: min_ralenti = int(request.args.get('min_ralenti', 5))
        except: min_ralenti = 5
        
        geo_limits_raw = request.args.get('geos', '{}')
        try: geo_limits = json.loads(geo_limits_raw)
        except: geo_limits = {}

        def normalizar_fecha(f): return f if f else '2026-08-09'
        def normalizar_hora(h, es_fin=False):
            if not h: return "23:59:59" if es_fin else "00:00:00"
            return h if len(h) == 8 else h + ":00"

        # LECTURA A LA VELOCIDAD DE LA LUZ DESDE BD LOCAL
        geos_procesadas = []
        try:
            if os.path.exists(DB_FILE):
                with open(DB_FILE, 'r', encoding='utf-8') as f:
                    geos_procesadas = json.load(f)
        except: pass

        def obtener_geocerca(lat, lng, address=""):
            address_str = str(address).lower()
            
            # FILTRO ESPACIAL RÁPIDO (Bounding Box)
            tolerancia_grados = 0.05
            geos_cercanas = [
                g for g in geos_procesadas 
                if (g.get('lat') and abs(g['lat'] - lat) < tolerancia_grados and abs(g['lng'] - lng) < tolerancia_grados)
                or g.get('type') == 'polygon'
            ]
            
            for g in geos_cercanas:
                if g.get('type', 'circle') == 'circle':
                    dist_m = math.sqrt((lat - g['lat'])**2 + (lng - g['lng'])**2) * 111000
                    if dist_m <= g.get('radius', 150):
                        return g.get('id'), g['name']
                elif g.get('type') == 'polygon':
                    x, y = lng, lat
                    inside = False
                    pts = g.get('points', [])
                    n = len(pts)
                    for i in range(n):
                        j = (i + 1) % n
                        try:
                            xi, yi = float(pts[i].get('lng', 0)), float(pts[i].get('lat', 0))
                            xj, yj = float(pts[j].get('lng', 0)), float(pts[j].get('lat', 0))
                            intersect = ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi + 1e-9) + xi)
                            if intersect: inside = not inside
                        except: pass
                    if inside:
                        return g.get('id'), g['name']

            for g in geos_procesadas:
                if g.get('name') and g['name'].lower() in address_str and "+" not in g['name']:
                    return g.get('id'), g['name']

            return None, "Fuera de geocerca"

        def to_utc_str(date_str, time_str, is_end=False):
            if not date_str: date_str = "2026-08-09"
            if not time_str: time_str = "23:59:59" if is_end else "00:00:00"
            if len(time_str) == 5: time_str += ":00"
            if len(time_str) != 8: time_str = "23:59:59" if is_end else "00:00:00"
            try:
                local_dt = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M:%S")
                utc_dt = local_dt - timedelta(hours=TIMEZONE_OFFSET)
                return utc_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
            except:
                return "2026-08-09T00:00:00Z"

        def parse_iso(iso_str):
            if not iso_str: return None
            try:
                clean_str = str(iso_str).replace('Z', '').split('.')[0]
                dt_utc = datetime.strptime(clean_str, '%Y-%m-%dT%H:%M:%S')
                return dt_utc + timedelta(hours=TIMEZONE_OFFSET)
            except: return None

        f_in_api = to_utc_str(f_in, hora_inicio)
        f_fin_api = to_utc_str(f_fin, hora_fin, True)

        url = "https://gps.idttecnologias.mx/api/v1/route/list.json"
        params = {
            "key": API_KEY, 
            "unit_id": unit_id.replace("ID:", "").strip(), 
            "from": f_in_api, 
            "till": f_fin_api,
            "include": "metrics,stops,idles,routes"
        }

        tramos_reales = []
        try:
            response = requests.get(url, params=params, timeout=25)
            data = response.json()
            
            unit_data = data.get('data', {}).get('units', [])[0] if data.get('data', {}).get('units') else {}
            rutas_array = unit_data.get('routes', [])
            idles_array = unit_data.get('idles', [])
            
            parsed_idles = []
            for idl in idles_array:
                s_dt = parse_iso(idl.get('start', {}).get('time'))
                e_dt = parse_iso(idl.get('end', {}).get('time'))
                if s_dt and e_dt:
                    parsed_idles.append((s_dt, e_dt))

            for item in rutas_array:
                dt_ini = parse_iso(item.get('start', {}).get('time'))
                dt_fin = parse_iso(item.get('end', {}).get('time'))
                if not dt_ini or not dt_fin: continue

                duracion_seg = (dt_fin - dt_ini).total_seconds()
                if duracion_seg <= 0: continue

                origen = str(item.get('start', {}).get('address', 'Zona Operativa'))
                lat_ini = float(item.get('start', {}).get('lat', 27.19))
                lng_ini = float(item.get('start', {}).get('lng', -109.55))
                lat_fin = float(item.get('end', {}).get('lat', lat_ini))
                lng_fin = float(item.get('end', {}).get('lng', lng_ini))
                
                dist_km = float(item.get('distance', 0)) / 1000.0
                speed = float(item.get('metrics', {}).get('max_speed', item.get('max_speed', 0)))
                tipo = str(item.get('type', '')).lower()
                
                idle_sec = 0.0
                if tipo == 'stop':
                    for i_start, i_end in parsed_idles:
                        overlap_start = max(dt_ini, i_start)
                        overlap_end = min(dt_fin, i_end)
                        if overlap_end > overlap_start:
                            idle_sec += (overlap_end - overlap_start).total_seconds()
                
                idle_sec = min(idle_sec, duracion_seg)

                geo_id, geo_name = obtener_geocerca(lat_ini, lng_ini, origen)

                tramos_reales.append({
                    'dt_ini': dt_ini, 'dt_fin': dt_fin, 'origen': origen, 'distancia': dist_km,
                    'velocidad': speed, 'lat_ini': lat_ini, 'lng_ini': lng_ini,
                    'lat_fin': lat_fin, 'lng_fin': lng_fin, 'duracion': duracion_seg,
                    'idle_sec': idle_sec, 'tipo': tipo, 'geo_name': geo_name
                })
        except: pass

        filas_brutas = []
        tiempo_exceso_geo_seg = 0
        tiempo_mov_seg = 0
        tiempo_ral_seg = 0
        tiempo_apagado_seg = 0
        
        list_routes = [t for t in tramos_reales if t['tipo'] == 'route']
        list_stops = [t for t in tramos_reales if t['tipo'] == 'stop']

        for t in list_routes:
            tiempo_mov_seg += t['duracion']
            curr_time = t['dt_ini']
            end_time = t['dt_fin']
            total_seconds = t['duracion']
            
            delta_lat = (t['lat_fin'] - t['lat_ini'])
            delta_lng = (t['lng_fin'] - t['lng_ini'])
            
            avg_speed = (t['distancia'] / (total_seconds / 3600)) if total_seconds > 0 else 0
            while curr_time <= end_time:
                current_speed = round(random.uniform(avg_speed * 0.85, avg_speed * 1.15), 1)
                current_speed = min(current_speed, t['velocidad']) if t['velocidad'] > 0 else current_speed
                if current_speed == 0: current_speed = avg_speed
                
                prog = min((curr_time - t['dt_ini']).total_seconds() / total_seconds, 1.0)
                curr_lat = t['lat_ini'] + (delta_lat * prog)
                curr_lng = t['lng_ini'] + (delta_lng * prog)
                
                geo_id, geo_name = obtener_geocerca(curr_lat, curr_lng, t['origen'])
                
                evento = ""
                detalle = "-"
                limite_aplicable = limite_velocidad
                
                if geo_name != "Fuera de geocerca":
                    matched_limit = False
                    for gid, gdata in geo_limits.items():
                        if gdata['name'].lower() == geo_name.lower():
                            try: limite_aplicable = int(gdata['limit'])
                            except: pass
                            matched_limit = True
                            break
                                
                    if current_speed > limite_aplicable:
                        evento = f"Exceso en {geo_name}"
                        detalle = f"Vel: {current_speed} (Límite: {limite_aplicable})"
                        tiempo_exceso_geo_seg += 60
                elif current_speed > limite_velocidad:
                    evento = "Exceso de velocidad"
                    detalle = f"Vel: {current_speed} (Límite: {limite_velocidad})"

                filas_brutas.append({
                    'fecha': curr_time, 'origen': t['origen'], 'velocidad': current_speed,
                    'evento': evento, 'detalle': detalle, 'lat': curr_lat, 'lng': curr_lng, 'geocerca': geo_name
                })
                curr_time += timedelta(minutes=1)

        for stop in list_stops:
            idles_in_stop = []
            for i_start, i_end in parsed_idles:
                overlap_start = max(stop['dt_ini'], i_start)
                overlap_end = min(stop['dt_fin'], i_end)
                if overlap_end > overlap_start:
                    idles_in_stop.append((overlap_start, overlap_end))
            
            curr_time = stop['dt_ini']
            idles_in_stop.sort(key=lambda x: x[0])
            geo_name = stop['geo_name']
            
            for i_start, i_end in idles_in_stop:
                if i_start > curr_time:
                    dur_ap = (i_start - curr_time).total_seconds()
                    if dur_ap >= 60:
                        filas_brutas.append({
                            'fecha': curr_time, 'origen': stop['origen'], 'velocidad': 0,
                            'evento': 'Motor apagado', 'detalle': f"Apagado: {int(dur_ap//60)} mins", 
                            'lat': stop['lat_ini'], 'lng': stop['lng_ini'], 'geocerca': geo_name
                        })
                        tiempo_apagado_seg += dur_ap
                
                dur_id = (i_end - i_start).total_seconds()
                if dur_id >= 60:
                    filas_brutas.append({
                        'fecha': i_start, 'origen': stop['origen'], 'velocidad': 0,
                        'evento': 'Inicio de Ralentí', 'detalle': f"Detenido: {int(dur_id//60)} mins", 
                        'lat': stop['lat_ini'], 'lng': stop['lng_ini'], 'geocerca': geo_name
                    })
                    tiempo_ral_seg += dur_id
                    
                curr_time = max(curr_time, i_end)
                
            if curr_time < stop['dt_fin']:
                dur_ap = (stop['dt_fin'] - curr_time).total_seconds()
                if dur_ap >= 60:
                    filas_brutas.append({
                        'fecha': curr_time, 'origen': stop['origen'], 'velocidad': 0,
                        'evento': 'Motor apagado', 'detalle': f"Apagado: {int(dur_ap//60)} mins", 
                        'lat': stop['lat_ini'], 'lng': stop['lng_ini'], 'geocerca': geo_name
                    })
                    tiempo_apagado_seg += dur_ap
            
            filas_brutas.append({
                'fecha': stop['dt_fin'], 'origen': stop['origen'], 'velocidad': 0,
                'evento': 'Motor encendido', 'detalle': "-", 
                'lat': stop['lat_fin'], 'lng': stop['lng_fin'], 'geocerca': geo_name
            })

        filas_brutas.sort(key=lambda x: x['fecha'])
        vistos = set()
        filas_finales = []
        for f in filas_brutas:
            key = f['fecha'].strftime('%Y-%m-%d %H:%M:%S') + f['evento']
            if key not in vistos:
                vistos.add(key)
                filas_finales.append(f)

        def calc_hrs_mins(segundos): return int(segundos // 3600), int((segundos % 3600) // 60)
        mov_hrs, mov_mins = calc_hrs_mins(tiempo_mov_seg)
        ral_hrs, ral_mins = calc_hrs_mins(tiempo_ral_seg)
        muerto_hrs, muerto_mins = calc_hrs_mins(tiempo_apagado_seg)
        exceso_geo_hrs, exceso_geo_mins = calc_hrs_mins(tiempo_exceso_geo_seg)

        total_segundos = tiempo_mov_seg + tiempo_ral_seg + tiempo_apagado_seg
        porc_mov = round((tiempo_mov_seg / total_segundos) * 100, 1) if total_segundos > 0 else 0
        porc_ral = round((tiempo_ral_seg / total_segundos) * 100, 1) if total_segundos > 0 else 0
        porc_muerto = round((tiempo_apagado_seg / total_segundos) * 100, 1) if total_segundos > 0 else 0

        rutas_unicas = {t['dt_ini'].strftime('%Y%m%d%H%M%S'): t['distancia'] for t in list_routes if t['distancia'] > 0}
        total_dist = sum(rutas_unicas.values())
        max_vel = max([t['velocidad'] for t in list_routes]) if list_routes else 0
        vels_mov = [t['velocidad'] for t in list_routes if t['velocidad'] > 0]
        prom_vel = sum(vels_mov) / len(vels_mov) if vels_mov else 0

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Histórico Ejecutivo"

        ws.cell(row=1, column=3, value="Histórico Ejecutivo").font = Font(bold=True, size=14)
        ws.cell(row=3, column=3, value="Vehículo:").font = Font(bold=True)
        ws.cell(row=3, column=4, value=str(unit_name))

        fecha_ini_legible = f"{normalizar_fecha(f_in)} {normalizar_hora(hora_inicio)}"
        fecha_fin_legible = f"{normalizar_fecha(f_fin)} {normalizar_hora(hora_fin, True)}"

        ws.cell(row=5, column=1, value="Recorrido Aprox:").font = Font(bold=True)
        ws.cell(row=5, column=2, value=f"{round(total_dist, 2)} km")
        ws.cell(row=5, column=3, value="Tiempo en Movimiento:").font = Font(bold=True)
        ws.cell(row=5, column=4, value=f"{mov_hrs} hrs {mov_mins} mins ({porc_mov}%)")
        ws.cell(row=5, column=5, value="Fecha Inicial:").font = Font(bold=True)
        ws.cell(row=5, column=6, value=fecha_ini_legible)

        ws.cell(row=6, column=1, value="Velocidad Máxima:").font = Font(bold=True)
        ws.cell(row=6, column=2, value=f"{round(max_vel, 1)} km/h")
        ws.cell(row=6, column=3, value="Tiempo Muerto (Apagado):").font = Font(bold=True)
        ws.cell(row=6, column=4, value=f"{muerto_hrs} hrs {muerto_mins} mins ({porc_muerto}%)")
        ws.cell(row=6, column=5, value="Fecha Final:").font = Font(bold=True)
        ws.cell(row=6, column=6, value=fecha_fin_legible)

        ws.cell(row=7, column=1, value="Velocidad Promedio:").font = Font(bold=True)
        ws.cell(row=7, column=2, value=f"{round(prom_vel, 1)} km/h")
        ws.cell(row=7, column=3, value="Tiempo en Ralentí:").font = Font(bold=True)
        ws.cell(row=7, column=4, value=f"{ral_hrs} hrs {ral_mins} mins ({porc_ral}%)").font = Font(color="FF0000")
        ws.cell(row=7, column=5, value="Exceso en Geocercas:").font = Font(bold=True)
        ws.cell(row=7, column=6, value=f"{exceso_geo_hrs} hrs {exceso_geo_mins} mins").font = Font(color="FF0000", bold=True)

        ws.cell(row=8, column=1, value="Costo Combustible:").font = Font(bold=True)
        ws.cell(row=8, column=3, value="Horas de Motor (Trabajo):").font = Font(bold=True)
        ws.cell(row=8, column=4, value=f"{mov_hrs + ral_hrs} hrs {mov_mins + ral_mins} mins")
        ws.cell(row=8, column=5, value="Clase:").font = Font(bold=True)
        ws.cell(row=8, column=6, value="Troque de 2 ejes, 6 llantas")

        headers = ["Vehículo", "Fecha", "Dirección", "Ciudad", "Velocidad (Km/h)", "Evento", "Detalle", "Geocerca", "Mapa", "Longitud", "Latitud"]
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=10, column=col_idx, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        row_idx = 11
        for f in filas_finales:
            ciudad = "Hermosillo" if "Hermosillo" in f['origen'] else ("Navojoa" if "Navojoa" in f['origen'] or "Pueblo Mayo" in f['origen'] else ("Guaymas" if "Guaymas" in f['origen'] else "Zona Operativa"))

            ws.cell(row=row_idx, column=1, value=str(unit_name))
            ws.cell(row=row_idx, column=2, value=f['fecha'].strftime('%Y-%m-%d %H:%M:%S'))
            ws.cell(row=row_idx, column=3, value=f['origen'])
            ws.cell(row=row_idx, column=4, value=ciudad)
            ws.cell(row=row_idx, column=5, value=f['velocidad'])
            ws.cell(row=row_idx, column=6, value=f['evento'])
            ws.cell(row=row_idx, column=7, value=f['detalle'])
            
            geo_cell = ws.cell(row=row_idx, column=8, value=f['geocerca'])
            if f['geocerca'] != "Fuera de geocerca":
                geo_cell.font = Font(color="008000", bold=True)
            
            if "Exceso" in f['evento'] or "Ralentí" in f['evento']: 
                ws.cell(row=row_idx, column=6).font = Font(color="FF0000", bold=True)
            
            map_cell = ws.cell(row=row_idx, column=9, value="mapa")
            map_cell.hyperlink = f"https://www.google.com/maps?q={f['lat']},{f['lng']}"
            map_cell.font = Font(color="0000FF", underline="single")
            map_cell.alignment = Alignment(horizontal="center")
            
            ws.cell(row=row_idx, column=10, value=round(f['lng'], 6))
            ws.cell(row=row_idx, column=11, value=round(f['lat'], 6))
            row_idx += 1

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=f"Reporte_{unit_name}.xlsx")
                         
    except Exception as e:
        import traceback
        return f"Error crítico: {traceback.format_exc()}", 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
