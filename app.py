from flask import Flask, request, jsonify, render_template_string, send_file
import requests
import openpyxl
import io
import math
from datetime import datetime, timedelta

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
        async function cargarUnidades() {
            const select = $('#unit_select');
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');

            status.innerText = "⏳ Cargando catálogo de unidades...";
            btn.disabled = true;

            try {
                const res = await fetch('/api_unidades');
                if (!res.ok) throw new Error();
                const units = await res.json();

                select.empty();
                if (!units || units.length === 0) {
                    select.append(new Option('No se encontraron unidades', ''));
                    status.innerText = "Error al consultar IDT.";
                    return;
                }

                select.append(new Option('🔍 Escribe para buscar unidad...', ''));
                units.forEach(u => {
                    const text = `${u.label || ''} ${u.number || ''} (ID: ${u.unit_id})`.trim();
                    select.append(new Option(text, u.unit_id));
                });

                select.select2({ placeholder: "🔍 Escribe para buscar unidad...", allowClear: true, width: '100%' });
                btn.disabled = false;
                status.innerText = "";
            } catch (e) {
                status.innerText = "Error de conexión con la API de IDT.";
            }
        }

        async function generarReporte() {
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');
            const unitId = $('#unit_select').val();
            const unitText = $('#unit_select option:selected').text();
            const fechaInicio = document.getElementById('fecha_inicio').value;
            const fechaFin = document.getElementById('fecha_fin').value;
            const horaInicio = document.getElementById('hora_inicio').value;
            const horaFin = document.getElementById('hora_fin').value;
            const limiteVel = document.getElementById('limite_velocidad').value || 80;

            if (!unitId) { alert("Por favor selecciona una unidad."); return; }

            btn.disabled = true;
            status.innerText = "⏳ Calculando telemetría de velocidades y trayectos...";

            try {
                const params = new URLSearchParams({
                    unit_id: unitId,
                    unit_name: unitText,
                    fecha_inicio: fechaInicio,
                    fecha_fin: fechaFin,
                    hora_inicio: horaInicio,
                    hora_fin: horaFin,
                    limite_velocidad: limiteVel
                });

                const res = await fetch(`/generar_excel?${params.toString()}`);
                if (!res.ok) {
                    const errData = await res.json();
                    throw new Error(errData.error || "Error al procesar los datos.");
                }

                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                const safeName = unitText.replace(/[^a-zA-Z0-9]/g, '_');
                a.download = `Reporte_Minuto_a_Minuto_${safeName}_${fechaInicio}.xlsx`;
                document.body.appendChild(a);
                a.click();
                a.remove();

                status.innerText = "¡Reporte generado y descargado con éxito!";
            } catch (err) {
                console.error(err);
                alert(err.message || "Error al generar el archivo Excel.");
                status.innerText = "Error en la generación.";
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

def haversine(lat1, lon1, lat2, lon2):
    if not lat1 or not lon1 or not lat2 or not lon2:
        return 0.0
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)) if (1 - a) > 0 else 0
    return R * c

def parse_point_timestamp(pt):
    if not isinstance(pt, dict):
        return None
    time_val = pt.get('gmt') or pt.get('time') or pt.get('datetime') or pt.get('t')
    if not time_val:
        return None
    
    if isinstance(time_val, (int, float)):
        return time_val - (7 * 3600) if time_val > 1000000000 else time_val
    
    clean_str = str(time_val).replace('T', ' ').replace('Z', '')[:19]
    try:
        dt = datetime.strptime(clean_str, "%Y-%m-%d %H:%M:%S")
        if 'gmt' in pt and pt['gmt']:
            dt = dt - timedelta(hours=7)
        return dt.timestamp()
    except Exception:
        return None

def extraer_puntos_python(route_data):
    points = []
    if not route_data or not isinstance(route_data, dict):
        return points
    
    data = route_data.get('data', {})
    units = data.get('units', []) if isinstance(data, dict) else (data if isinstance(data, list) else route_data.get('units', []))

    if not isinstance(units, list):
        return points

    def add_pt(pt):
        if not pt or not isinstance(pt, dict):
            return
        
        ts = parse_point_timestamp(pt)
        if ts is None:
            return

        try:
            lat = float(pt.get('lat') or pt.get('latitude') or 0)
            lng = float(pt.get('lng') or pt.get('longitude') or 0)
            speed = float(pt.get('speed') or pt.get('s') or pt.get('spd') or 0)
        except (ValueError, TypeError):
            return

        if lat != 0 and lng != 0:
            params = pt.get('params', {})
            acc = 1 if speed > 0 else 0
            if isinstance(params, dict):
                if 'acc' in params:
                    try: acc = int(params['acc'])
                    except: pass
                elif 'din1' in params:
                    try: acc = int(params['din1'])
                    except: pass
                elif 'ignition' in params:
                    try: acc = int(params['ignition'])
                    except: pass

            points.append({
                'timestamp': ts,
                'lat': lat,
                'lng': lng,
                'speed': speed,
                'address': pt.get('address') or pt.get('addr') or '',
                'acc': acc
            })

    for u in units:
        if not isinstance(u, dict): continue
        if isinstance(u.get('points'), list):
            for pt in u['points']: add_pt(pt)
        if isinstance(u.get('routes'), list):
            for r in u['routes']:
                if isinstance(r, dict):
                    if isinstance(r.get('points'), list):
                        for pt in r['points']: add_pt(pt)
                    if isinstance(r.get('decoded_route'), list):
                        for pt in r['decoded_route']: add_pt(pt)
                    if r.get('start'): add_pt(r['start'])
                    if r.get('end'): add_pt(r['end'])
        if isinstance(u.get('tracks'), list):
            for t in u['tracks']:
                if isinstance(t, dict) and isinstance(t.get('points'), list):
                    for pt in t['points']: add_pt(pt)

    return points

@app.route('/')
def index():
    return render_template_string(HTML_INTERFACE)

@app.route('/api_unidades')
def api_unidades():
    try:
        res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY}, timeout=15)
        if res.ok:
            data = res.json()
            units = data.get('data', {}).get('units', []) or data.get('units', [])
            return jsonify(units)
    except Exception as e:
        print("Error al cargar unidades:", e)
    return jsonify([]), 500

@app.route('/generar_excel')
def generar_excel():
    unit_id = request.args.get('unit_id')
    unit_name = request.args.get('unit_name', f'Unidad {unit_id}')
    fecha_inicio = request.args.get('fecha_inicio')
    fecha_fin = request.args.get('fecha_fin')
    hora_inicio = request.args.get('hora_inicio', '00:00')
    hora_fin = request.args.get('hora_fin', '23:59')
    limite_vel = int(request.args.get('limite_velocidad', 80))

    if not unit_id or not fecha_inicio or not fecha_fin:
        return jsonify({'error': 'Faltan parámetros requeridos'}), 400

    try:
        start_dt = datetime.strptime(f"{fecha_inicio} {hora_inicio}", "%Y-%m-%d %H:%M")
        end_dt = datetime.strptime(f"{fecha_fin} {hora_fin}", "%Y-%m-%d %H:%M")
    except Exception as e:
        return jsonify({'error': f'Formato de fecha u hora inválido: {str(e)}'}), 400

    start_utc = start_dt + timedelta(hours=7)
    end_utc = end_dt + timedelta(hours=7)

    str_from = f"{fecha_inicio} {hora_inicio}:00"
    str_till = f"{fecha_fin} {hora_fin}:59"
    iso_from = start_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    iso_till = end_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    sec_from = int(start_utc.timestamp())
    sec_till = int(end_utc.timestamp())

    # 1. Alertas IDT
    alertas = []
    try:
        r_alerts = requests.get(f"{BASE_URL}/alert/list.json", params={
            'key': API_KEY, 'unit_id': unit_id, 'from': sec_from, 'till': sec_till
        }, timeout=15)
        if r_alerts.ok:
            data_a = r_alerts.json()
            alertas = data_a.get('data', {}).get('alerts', []) or data_a.get('alerts', []) or []
    except Exception as e:
        print("Aviso de alertas:", e)

    # 2. Puntos GPS
    raw_points = []
    url_1 = f"{BASE_URL}/route/list.json?key={API_KEY}&unit_id={unit_id}&from={str_from}&till={str_till}&include[]=points&include[]=decoded_route&include[]=stops"
    try:
        r1 = requests.get(url_1, timeout=25)
        if r1.ok: raw_points = extraer_puntos_python(r1.json())
    except Exception as e: pass

    if not raw_points:
        url_2 = f"{BASE_URL}/route/list.json?key={API_KEY}&unit_id={unit_id}&from={iso_from}&till={iso_till}&include[]=points&include[]=decoded_route&include[]=stops"
        try:
            r2 = requests.get(url_2, timeout=25)
            if r2.ok: raw_points = extraer_puntos_python(r2.json())
        except Exception as e: pass

    if not raw_points:
        url_3 = f"{BASE_URL}/route/list.json?key={API_KEY}&unit_id={unit_id}&from={sec_from}&till={sec_till}&include[]=points&include[]=decoded_route&include[]=stops"
        try:
            r3 = requests.get(url_3, timeout=25)
            if r3.ok: raw_points = extraer_puntos_python(r3.json())
        except Exception as e: pass

    if not raw_points:
        return jsonify({
            'error': f'No se obtuvieron posiciones GPS para la unidad {unit_name} en el rango {fecha_inicio} {hora_inicio} a {fecha_fin} {hora_fin}.'
        }), 404

    raw_points.sort(key=lambda x: x['timestamp'])

    # Construcción de tramos para interpolación lineal entre puntos
    legs = []
    if len(raw_points) == 1:
        legs.append({'pA': raw_points[0], 'pB': raw_points[0], 'dt_sec': 1, 'dist_km': 0, 'speed_leg': 0.0})
    else:
        for i in range(len(raw_points) - 1):
            pA = raw_points[i]
            pB = raw_points[i+1]
            dt_sec = pB['timestamp'] - pA['timestamp']
            dist_km = haversine(pA['lat'], pA['lng'], pB['lat'], pB['lng'])
            speed_leg = (dist_km / (dt_sec / 3600.0)) if dt_sec > 0 else 0.0
            legs.append({
                'pA': pA,
                'pB': pB,
                'dt_sec': dt_sec,
                'dist_km': dist_km,
                'speed_leg': speed_leg
            })

    rows = []
    max_velocidad = 0
    minutos_movimiento = 0
    minutos_ralenti = 0
    recorrido_total_km = 0.0

    curr_ts = start_dt.timestamp()
    end_ts = end_dt.timestamp()
    prev_lat, prev_lng = None, None

    while curr_ts <= end_ts:
        active_leg = None
        if curr_ts <= raw_points[0]['timestamp']:
            active_leg = legs[0]
            alpha = 0.0
        elif curr_ts >= raw_points[-1]['timestamp']:
            active_leg = legs[-1]
            alpha = 1.0
        else:
            for leg in legs:
                if leg['pA']['timestamp'] <= curr_ts <= leg['pB']['timestamp']:
                    active_leg = leg
                    dt_sec = leg['dt_sec']
                    alpha = (curr_ts - leg['pA']['timestamp']) / dt_sec if dt_sec > 0 else 0.0
                    break

        pA = active_leg['pA']
        pB = active_leg['pB']
        v_leg = active_leg.get('speed_leg', 0.0)

        # Velocidad directa reportada por el sensor o interpolada por la distancia del tramo
        speed_raw = pA.get('speed', 0.0) + alpha * (pB.get('speed', 0.0) - pA.get('speed', 0.0)) if pB.get('speed') is not None else pA.get('speed', 0.0)

        if v_leg > 5.0:
            cur_lat = pA['lat'] + alpha * (pB['lat'] - pA['lat'])
            cur_lng = pA['lng'] + alpha * (pB['lng'] - pA['lng'])
            speed_final = int(round(max(speed_raw, v_leg)))
        else:
            cur_lat = pA['lat']
            cur_lng = pA['lng']
            speed_final = int(round(speed_raw)) if speed_raw > 5 else 0

        acc_state = pA.get('acc', 0)

        if speed_final > max_velocidad:
            max_velocidad = speed_final

        dt_item = datetime.fromtimestamp(curr_ts)
        fecha_formatted = dt_item.strftime("%Y-%m-%d %H:%M:00")

        # Verificación de Alertas del Servidor IDT
        alerta_match = None
        for a in alertas:
            a_str = a.get('gmt') or a.get('time') or ''
            a_ts = parse_point_timestamp({'time': a_str, 'gmt': a.get('gmt')})
            if a_ts and abs(a_ts - curr_ts) <= 60:
                alerta_match = a
                break

        if alerta_match is not None:
            evento = "🚨 Exceso de velocidad"
            msg = alerta_match.get('msg') or alerta_match.get('title') or 'Notificación de exceso de velocidad'
            v_val = alerta_match.get('value') or speed_final
            detalle = f"Notificación IDT: {msg} ({v_val} km/h)"
            minutos_movimiento += 1
        elif speed_final > 0:
            evento = "En movimiento"
            detalle = "-"
            minutos_movimiento += 1
        elif acc_state == 1:
            evento = "Ralentí / Motor ON"
            detalle = "-"
            minutos_ralenti += 1
        else:
            evento = "Apagado"
            detalle = "-"

        direccion = pA.get('address', 'Sonora, Mexico')
        ciudad = ""
        if ',' in direccion:
            partes = direccion.split(',')
            if len(partes) >= 2:
                ciudad = partes[-2].strip()

        maps_url = f"https://www.google.com/maps?q={cur_lat},{cur_lng}"

        if prev_lat is not None and (prev_lat != cur_lat or prev_lng != cur_lng):
            recorrido_total_km += haversine(prev_lat, prev_lng, cur_lat, cur_lng)

        prev_lat, prev_lng = cur_lat, cur_lng

        rows.append([
            unit_name,
            fecha_formatted,
            direccion,
            ciudad,
            speed_final,
            evento,
            detalle,
            maps_url,
            cur_lng,
            cur_lat
        ])

        curr_ts += 60.0

    fmt_mov = f"{minutos_movimiento // 60}h {minutos_movimiento % 60}m"
    fmt_ral = f"{minutos_ralenti // 60}h {minutos_ralenti % 60}m"
    min_trab = minutos_movimiento + minutos_ralenti
    fmt_trab = f"{min_trab // 60}h {min_trab % 60}m"
    vel_promedio = int(round(recorrido_total_km / (minutos_movimiento / 60.0))) if minutos_movimiento > 0 else 0

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte"

    ws.append(["", "", "Histórico Minuto a Minuto con Excesos de Velocidad"])
    ws.append([])
    ws.append(["", "", unit_name])
    ws.append([])
    ws.append(["Recorrido Aprox:", f"{recorrido_total_km:.2f} km", "Tiempo en Movimiento:", fmt_mov, "Fecha Inicial:", f"{fecha_inicio} {hora_inicio}"])
    ws.append(["Velocidad Máxima:", f"{max_velocidad} km/h", "Tiempo Muerto:", fmt_ral, "Fecha Final:", f"{fecha_fin} {hora_fin}"])
    ws.append(["Velocidad Promedio:", f"{vel_promedio} km/h", "Horas Trabajadas:", fmt_trab, "Consumo Combustible:", "N/A"])
    ws.append(["Costo Combustible:", "N/A"])
    ws.append([])
    ws.append(["Vehículo", "Fecha", "Dirección", "Ciudad", "Velocidad (Km/h)", "Evento", "Detalle", "Mapa", "Longitud", "Latitud"])

    for r in rows:
        row_copy = list(r)
        row_copy[7] = f'=HYPERLINK("{r[7]}", "Ver en Mapa")'
        ws.append(row_copy)

    col_widths = [32, 20, 45, 20, 15, 24, 35, 18, 15, 15]
    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    safe_name = "".join([c if c.isalnum() else "_" for c in unit_name])
    filename = f"Reporte_Minuto_a_Minuto_{safe_name}_{fecha_inicio}.xlsx"

    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
