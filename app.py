from flask import Flask, render_template_string, request, send_file, jsonify
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import io
from datetime import datetime, timedelta
import random

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
                const units = await res.json();
                select.empty();
                select.append(new Option('🔍 Escribe para buscar unidad...', ''));
                units.forEach(u => select.append(new Option(`${u.label || ''} ${u.number || ''} (ID: ${u.unit_id})`.trim(), u.unit_id)));
                select.select2({ placeholder: "🔍 Escribe para buscar unidad...", width: '100%' });
                btn.disabled = false;
                status.innerText = "";
            } catch (e) { status.innerText = "Error de conexión."; }
        }

        async function generarReporte() {
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');
            const unitId = $('#unit_select').val();
            const unitText = $('#unit_select option:selected').text();
            
            if (!unitId) { alert("Por favor selecciona una unidad."); return; }
            
            btn.disabled = true;
            status.innerText = "⏳ Generando reporte...";

            const params = new URLSearchParams({
                unit_id: unitId,
                unit_name: unitText,
                fecha_inicio: document.getElementById('fecha_inicio').value,
                fecha_fin: document.getElementById('fecha_fin').value,
                hora_inicio: document.getElementById('hora_inicio').value,
                hora_fin: document.getElementById('hora_fin').value
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
                alert("Error al generar el reporte.");
                status.innerText = "Error en la generación.";
            }
            btn.disabled = false;
        }

        function setRange(type) {
            const now = new Date();
            let start = new Date();
            let end = new Date();

            const formatDate = (d) => {
                let month = '' + (d.getMonth() + 1),
                    day = '' + d.getDate(),
                    year = d.getFullYear();
                if (month.length < 2) month = '0' + month;
                if (day.length < 2) day = '0' + day;
                return [year, month, day].join('-');
            };

            if (type === 'hoy') {
                start = now;
                end = now;
            } else if (type === 'ayer') {
                start.setDate(now.getDate() - 1);
                end.setDate(now.getDate() - 1);
            } else if (type === 'esta_semana') {
                const day = now.getDay() || 7;
                start.setDate(now.getDate() - day + 1);
                end = now;
            } else if (type === 'semana_anterior') {
                const day = now.getDay() || 7;
                start.setDate(now.getDate() - day - 6);
                end.setDate(now.getDate() - day);
            } else if (type === 'ultimos_7_dias') {
                start.setDate(now.getDate() - 6);
                end = now;
            } else if (type === 'este_mes') {
                start = new Date(now.getFullYear(), now.getMonth(), 1);
                end = now;
            } else if (type === 'mes_anterior') {
                start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
                end = new Date(now.getFullYear(), now.getMonth(), 0);
            }

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

@app.route('/api_unidades')
def api_unidades():
    try:
        res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY}, timeout=15)
        return jsonify(res.json().get('data', {}).get('units', []))
    except: return jsonify([]), 500

import random

@app.route('/generar_excel')
def generar_excel():
    try:
        from flask import request, send_file
        import requests
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill
        import io
        import os
        from datetime import datetime
        import re

        # 1. Parámetros
        unit_id = request.args.get('unit_id', '868807')
        if "ID:" in unit_id:
            unit_id = unit_id.split("ID:")[1].replace(")", "").strip()

        f_in = request.args.get('fecha_inicio', '2026-08-09')
        f_fin = request.args.get('fecha_fin', f_in)
        hora_inicio = request.args.get('hora_inicio', '00:00:00')
        hora_fin = request.args.get('hora_fin', '23:59:59')

        def normalizar_fecha(fecha_str):
            try:
                if "/" in fecha_str:
                    return datetime.strptime(fecha_str, '%d/%m/%Y').strftime('%Y-%m-%d')
                elif "-" in fecha_str and len(fecha_str) == 10:
                    return fecha_str
            except:
                pass
            return '2026-08-09'

        def normalizar_hora(hora_str, es_fin=False):
            try:
                if "AM" in hora_str.upper() or "PM" in hora_str.upper():
                    return datetime.strptime(hora_str.strip(), '%I:%M %p').strftime('%H:%M:%00')
                if len(hora_str) == 5:
                    return f"{hora_str}:00"
                if len(hora_str) == 8:
                    return hora_str
            except:
                pass
            return "23:59:59" if es_fin else "00:00:00"

        # ISO 8601 Estricto para Mapon
        f_in_api = f"{normalizar_fecha(f_in)}T{normalizar_hora(hora_inicio)}Z"
        f_fin_api = f"{normalizar_fecha(f_fin)}T{normalizar_hora(hora_fin, True)}Z"

        api_key = os.environ.get('MAPON_API_KEY')

        # 3. Petición a Mapon 
        url = "https://gps.idttecnologias.mx/api/v1/route/list.json"
        
        params = {
            "key": api_key,
            "unit_id": unit_id,
            "from": f_in_api,
            "till": f_fin_api
        }

        tramos_reales = []
        api_debug_info = ""
        data = None

        try:
            response = requests.get(url, params=params, timeout=15)
            data = response.json()

            rutas_encontradas = []
            def extraer_tramos(obj):
                if isinstance(obj, dict):
                    if 'start_time' in obj or 'start' in obj or 'distance' in obj:
                        rutas_encontradas.append(obj)
                    else:
                        for k, v in obj.items():
                            extraer_tramos(v)
                elif isinstance(obj, list):
                    for item in obj:
                        extraer_tramos(item)

            extraer_tramos(data)

            # Extractor inteligente de hora (soluciona el error de "3:59::00")
            def extraer_hhmmss(time_str):
                if not time_str: return "00:00:00"
                match = re.search(r'(\d{2}:\d{2}:\d{2})', str(time_str))
                if match: return match.group(1)
                match = re.search(r'(\d{2}:\d{2})', str(time_str))
                if match: return match.group(1) + ":00"
                return "00:00:00"

            # Procesamos y limpiamos los datos de Mapon
            for item in rutas_encontradas:
                # Hora
                h_str = item.get('start', {}).get('time', item.get('start_time', '00:00'))
                h_ini = extraer_hhmmss(h_str)
                
                # Duración (para calcular tiempos)
                duracion_seg = float(item.get('duration', item.get('time', 0))) # Tiempo en segundos
                
                # Dirección y Coordenadas
                origen = str(item.get('start', {}).get('address', item.get('start_address', 'Zona Operativa')))
                lat = float(item.get('start', {}).get('lat', item.get('start_lat', 27.19289)))
                lng = float(item.get('start', {}).get('lng', item.get('start_lng', -109.55168)))
                
                # Distancia (Mapon la manda en metros, hay que pasarla a kilómetros)
                dist_raw = item.get('distance', 0)
                dist_km = float(dist_raw) / 1000.0 if dist_raw else 0.0
                
                # Velocidad
                speed_raw = item.get('metrics', {}).get('max_speed', item.get('max_speed', 0))
                speed = float(speed_raw) if speed_raw is not None else 0.0
                
                tramos_reales.append({
                    'hora': h_ini,
                    'origen': origen,
                    'distancia': dist_km,
                    'velocidad': speed,
                    'lat': lat,
                    'lng': lng,
                    'duracion': duracion_seg
                })
        except Exception as api_err:
            api_debug_info = f"Error Parseando JSON: {str(api_err)}"

        # 4. Cálculos exactos
        total_dist = sum([t['distancia'] for t in tramos_reales])
        max_vel = max([t['velocidad'] for t in tramos_reales]) if tramos_reales else 0
        tramos_mov = [t for t in tramos_reales if t['velocidad'] > 0]
        prom_vel = sum([t['velocidad'] for t in tramos_mov]) / len(tramos_mov) if tramos_mov else 0

        tiempo_movimiento_seg = sum([t['duracion'] for t in tramos_reales if t['velocidad'] > 0])
        tiempo_muerto_seg = sum([t['duracion'] for t in tramos_reales if t['velocidad'] == 0])

        mov_hrs = int(tiempo_movimiento_seg // 3600)
        mov_mins = int((tiempo_movimiento_seg % 3600) // 60)
        muerto_hrs = int(tiempo_muerto_seg // 3600)
        muerto_mins = int((tiempo_muerto_seg % 3600) // 60)

        total_segundos = tiempo_movimiento_seg + tiempo_muerto_seg
        porc_mov = round((tiempo_movimiento_seg / total_segundos) * 100, 1) if total_segundos > 0 else 0
        porc_muerto = round((tiempo_muerto_seg / total_segundos) * 100, 1) if total_segundos > 0 else 0

        # 5. Construcción del Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Histórico"

        ws.cell(row=1, column=3, value="Histórico").font = Font(bold=True, size=14)
        ws.cell(row=3, column=3, value="Unidad").font = Font(bold=True)
        ws.cell(row=3, column=4, value=str(unit_id))

        # Texto legible para las cabeceras de fechas
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
        ws.cell(row=6, column=3, value="Tiempo Muerto").font = Font(bold=True)
        ws.cell(row=6, column=4, value=f"{muerto_hrs} hrs {muerto_mins} mins ({porc_muerto}%)")
        ws.cell(row=6, column=5, value="Fecha Final:").font = Font(bold=True)
        ws.cell(row=6, column=6, value=fecha_fin_legible)

        ws.cell(row=7, column=1, value="Velocidad Promedio:").font = Font(bold=True)
        ws.cell(row=7, column=2, value=f"{round(prom_vel, 1)} km/h")
        ws.cell(row=7, column=3, value="Horas Trabajadas:").font = Font(bold=True)
        ws.cell(row=7, column=4, value="24.0 hrs")
        ws.cell(row=7, column=5, value="Consumo Combustible:").font = Font(bold=True)
        ws.cell(row=7, column=6, value="A calcular")

        ws.cell(row=8, column=1, value="Costo Combustible:").font = Font(bold=True)
        ws.cell(row=8, column=3, value="Costo Viaje:").font = Font(bold=True)
        ws.cell(row=8, column=5, value="Clase:").font = Font(bold=True)
        ws.cell(row=8, column=6, value="Troque de 2 ejes, 6 llantas")

        headers = [
            "Vehículo", "Fecha", "Dirección", "Ciudad", 
            "Velocidad (Km/h)", "Evento", "Detalle", "Mapa", "Longitud", "Latitud"
        ]
        
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=10, column=col_idx, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        row_idx = 11

        for t in tramos_reales:
            fecha_str = f"{normalizar_fecha(f_in)} {t['hora']}"
            ciudad = "Navojoa" if "Navojoa" in t['origen'] or "Pueblo Mayo" in t['origen'] else ("Guaymas" if "Guaymas" in t['origen'] else "Zona Operativa")
            
            speed = t['velocidad']
            dist = round(t['distancia'], 2)
            
            if speed > 0:
                evento = "Exceso de velocidad" if speed > 80 else "Motor encendido"
                detalle = f"Distancia: {dist} km"
            else:
                evento = "Motor apagado"
                # Convertimos segundos de detención a formato legible
                mins_detenido = int(t['duracion'] // 60)
                hrs_detenido = mins_detenido // 60
                mins_restantes = mins_detenido % 60
                if hrs_detenido > 0:
                    detalle = f"Detenido: {hrs_detenido} hrs {mins_restantes} mins"
                else:
                    detalle = f"Detenido: {mins_detenido} mins"

            ws.cell(row=row_idx, column=1, value=str(unit_id))
            ws.cell(row=row_idx, column=2, value=fecha_str)
            ws.cell(row=row_idx, column=3, value=t['origen'])
            ws.cell(row=row_idx, column=4, value=ciudad)
            ws.cell(row=row_idx, column=5, value=speed)
            ws.cell(row=row_idx, column=6, value=evento)
            ws.cell(row=row_idx, column=7, value=detalle)
            
            map_cell = ws.cell(row=row_idx, column=8, value="mapa")
            map_cell.hyperlink = f"https://www.google.com/maps?q={t['lat']},{t['lng']}"
            map_cell.font = Font(color="0000FF", underline="single")
            map_cell.alignment = Alignment(horizontal="center")
            
            ws.cell(row=row_idx, column=9, value=t['lng'])
            ws.cell(row=row_idx, column=10, value=t['lat'])
            row_idx += 1

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="Historico_Oficial.xlsx")
                         
    except Exception as e:
        import traceback
        return f"Error crítico: {traceback.format_exc()}", 500
