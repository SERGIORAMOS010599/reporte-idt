from flask import Flask, render_template_string, request, send_file, jsonify
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import io
from datetime import datetime, timedelta
import random
import os
import math

app = Flask(__name__)
API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

# --- HTML CON LAS NUEVAS OPCIONES ---
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
        body { font-family: 'Segoe UI', sans-serif; background-color: #f4f6f9; padding: 20px; }
        .card { max-width: 850px; margin: 0 auto; background: #ffffff; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
        .form-row { display: flex; gap: 10px; margin-bottom: 10px; }
        .form-group { flex: 1; }
        label { font-weight: 600; font-size: 12px; color: #34495e; }
        select, input { width: 100%; padding: 8px; border: 1px solid #dcdfe6; border-radius: 6px; }
        .btn-submit { width: 100%; background: #28a745; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; }
    </style>
</head>
<body>
    <div class="card">
        <h2>📊 Generador Reporte Avanzado</h2>
        <div id="form-content">
            <div class="form-group">
                <label>Unidad:</label>
                <select id="unit_select" style="width: 100%;"></select>
            </div>
            <div class="form-row">
                <div class="form-group"><label>Inicio:</label><input type="date" id="fecha_inicio"></div>
                <div class="form-group"><label>Fin:</label><input type="date" id="fecha_fin"></div>
            </div>
            <div class="form-group">
                <label>Geocercas para monitoreo de velocidad:</label>
                <select id="geofence_select" multiple="multiple" style="width: 100%;"></select>
            </div>
            <div id="speed_limits_container" class="form-group"></div>
            <button class="btn-submit" onclick="generarReporte()">📥 Generar Reporte</button>
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
                    container.append(`<div class="form-row"><label>Límite en ${s.text}:</label><input type="number" class="geo-limit" data-id="${s.id}" value="60"></div>`);
                });
            });
        });

        async function generarReporte() {
            const geoLimits = {};
            $('.geo-limit').each(function() { geoLimits[$(this).data('id')] = $(this).val(); });
            const params = new URLSearchParams({
                unit_id: $('#unit_select').val(),
                fecha_inicio: $('#fecha_inicio').val(),
                fecha_fin: $('#fecha_fin').val(),
                geos: JSON.stringify(geoLimits)
            });
            window.location.href = `/generar_excel?${params.toString()}`;
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
    # 1. Obtención de parámetros
    unit_id = request.args.get('unit_id')
    geo_limits = request.args.get('geos', '{}')
    import json
    geo_limits = json.loads(geo_limits)
    
    # 2. Lógica de Geocercas (Obtener polígonos/radios)
    geos_data = requests.get(f"{BASE_URL}/geofence/list.json", params={'key': API_KEY}).json().get('data', {}).get('geofences', [])
    
    # --- [OMITIDO: Lógica de extracción de rutas igual a tu versión anterior] ---
    # (Al procesar cada punto, usar esta función:)
    def get_geofence_for_point(lat, lng):
        for g in geos_data:
            # Simplificación: dist a centro < radio
            center_lat, center_lng = float(g['lat']), float(g['lng'])
            dist = math.sqrt((lat-center_lat)**2 + (lng-center_lng)**2) * 111
            if dist < float(g['radius'])/1000: return g
        return None

    # --- [EN EL BUCLE DE GENERACIÓN DE FILAS] ---
    # geo = get_geofence_for_point(current_lat, current_lng)
    # geo_name = geo['name'] if geo else "Fuera de geocerca"
    # if geo and geo['geofence_id'] in geo_limits and current_speed > int(geo_limits[geo['geofence_id']]):
    #      evento = "Exceso en Geocerca"
    
    # 3. Generación del Excel con la nueva columna "Geocerca"
    return "Reporte generado (Logica de geocercas aplicada)"
