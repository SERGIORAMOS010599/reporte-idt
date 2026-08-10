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

app = Flask(__name__)

API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

# --- HTML CON TUS DISEÑOS ORIGINALES + NUEVOS SELECTORES ---
HTML_INTERFACE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Reporte Minuto a Minuto - IDT</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; padding: 20px; }
        .card { max-width: 820px; margin: 0 auto; background: #ffffff; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
        .header { text-align: center; margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px; }
        .form-content { flex: 1; }
        .form-group { margin-bottom: 15px; }
        .form-row { display: flex; gap: 10px; }
        .form-row .form-group { flex: 1; }
        label { display: block; font-weight: 600; margin-bottom: 5px; color: #34495e; font-size: 12px; }
        select, input { width: 100%; padding: 9px; border: 1px solid #dcdfe6; border-radius: 6px; }
        .btn-submit { width: 100%; background: #28a745; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; margin-top: 10px; }
    </style>
</head>
<body>
    <div class="card">
        <div class="header"><h2>📊 Reporte Minuto a Minuto</h2></div>
        <div class="form-content">
            <div class="form-group">
                <label>Unidad:</label>
                <select id="unit_select" style="width: 100%;"></select>
            </div>
            <div class="form-row">
                <div class="form-group"><label>Fecha Inicial:</label><input type="date" id="fecha_inicio"></div>
                <div class="form-group"><label>Fecha Final:</label><input type="date" id="fecha_fin"></div>
            </div>
            <div class="form-group">
                <label>Geocercas para monitoreo:</label>
                <select id="geofence_select" multiple="multiple" style="width: 100%;"></select>
            </div>
            <div id="speed_limits_container"></div>
            <button type="button" class="btn-submit" onclick="generarReporte()">📥 Generar Excel</button>
        </div>
    </div>
    <script>
        $(document).ready(async function() {
            const units = await (await fetch('/api_unidades')).json();
            $('#unit_select').select2({ data: units.map(u => ({id: u.unit_id, text: `${u.label} (ID: ${u.unit_id})`})) });
            const geos = await (await fetch('/api_geocercas')).json();
            $('#geofence_select').select2({ data: geos.map(g => ({id: g.geofence_id, text: g.name})), placeholder: "Seleccionar..." });
            
            $('#geofence_select').on('change', function() {
                const container = $('#speed_limits_container');
                container.empty();
                $(this).select2('data').forEach(s => {
                    container.append(`<div class="form-group"><label>Límite en ${s.text} (km/h):</label><input type="number" class="geo-limit" data-id="${s.id}" value="60"></div>`);
                });
            });
        });
        async function generarReporte() {
            const geoLimits = {}; $('.geo-limit').each(function() { geoLimits[$(this).data('id')] = $(this).val(); });
            const p = new URLSearchParams({ unit_id: $('#unit_select').val(), fecha_inicio: $('#fecha_inicio').val(), fecha_fin: $('#fecha_fin').val(), geos: JSON.stringify(geoLimits) });
            window.location.href = '/generar_excel?' + p.toString();
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index(): return render_template_string(HTML_INTERFACE)

@app.route('/api_unidades')
def api_unidades():
    res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY})
    return jsonify(res.json().get('data', {}).get('units', []))

@app.route('/api_geocercas')
def api_geocercas():
    res = requests.get(f"{BASE_URL}/geofence/list.json", params={'key': API_KEY})
    return jsonify(res.json().get('data', {}).get('geofences', []))

@app.route('/generar_excel')
def generar_excel():
    # Lógica de extracción (usando tu versión probada)
    # ... [Insertar aquí la lógica de extracción de tramos_reales que ya tienes] ...
    # Al procesar las filas:
    
    # 1. Obtener geocercas para comparar
    geo_limits = json.loads(request.args.get('geos', '{}'))
    geos_data = requests.get(f"{BASE_URL}/geofence/list.json", params={'key': API_KEY}).json().get('data', {}).get('geofences', [])
    
    def get_geo(lat, lng):
        for g in geos_data:
            # Cálculo simple de distancia a centro
            dist = math.sqrt((lat-float(g['lat']))**2 + (lng-float(g['lng']))**2) * 111
            if dist < float(g.get('radius', 500))/1000: return g
        return None

    # 2. Dentro del bucle de filas (donde asignas eventos):
    # geo = get_geo(current_lat, current_lng)
    # geo_name = geo['name'] if geo else "Fuera"
    # if geo and geo['geofence_id'] in geo_limits and current_speed > int(geo_limits[geo['geofence_id']]):
    #     evento = f"Exceso en {geo_name}"
    
    # 3. Añadir columna al Excel: ws.cell(row=row_idx, column=11, value=geo_name)
    return "Reporte generado con geocercas"
