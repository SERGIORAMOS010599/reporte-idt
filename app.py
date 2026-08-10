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
        .btn-submit { width: 100%; background: #28a745; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; font-size: 14px; cursor: pointer; margin-top: 10px; }
        .btn-submit:hover { background: #218838; }
        .status-msg { font-size: 13px; color: #e67e22; margin-top: 10px; text-align: center; font-weight: bold; }
        .retry-btn { font-size: 11px; color: #007bff; text-decoration: underline; cursor: pointer; margin-left: 8px; }
    </style>
</head>
<body>
    <div class="card">
        <div class="header"><h2>📊 Reporte Minuto a Minuto</h2><p>IDT Tecnologías - Generador de Histórico GPS Avanzado</p></div>
        <div class="main-container">
            <div class="presets-sidebar">
                <label>Atajos de Fecha:</label>
                <button type="button" onclick="setRange('hoy')">Hoy</button>
                <button type="button" onclick="setRange('ayer')">Ayer</button>
            </div>
            <div class="form-content">
                <div class="form-group">
                    <label>Unidad:</label>
                    <select id="unit_select" style="width: 100%;"></select>
                </div>
                <div class="form-row">
                    <div class="form-group"><label>Fecha Inicial:</label><input type="date" id="fecha_inicio"></div>
                    <div class="form-group"><label>Fecha Final:</label><input type="date" id="fecha_fin"></div>
                </div>
                <div class="form-row">
                    <div class="form-group"><label>Límite Velocidad:</label><input type="number" id="limite_velocidad" value="80"></div>
                    <div class="form-group"><label>Alerta Ralentí:</label><input type="number" id="min_ralenti" value="5"></div>
                </div>
                <!-- NUEVO DISEÑO AGREGADO AQUÍ -->
                <div class="form-group">
                    <label>Monitoreo Geocercas:</label>
                    <select id="geofence_select" multiple="multiple" style="width: 100%;"></select>
                </div>
                <div id="speed_limits_container"></div>
                
                <button type="button" class="btn-submit" onclick="generarReporte()">📥 Generar y Descargar Excel</button>
                <div id="status_msg" class="status-msg"></div>
            </div>
        </div>
    </div>
    <script>
        $(document).ready(async function() {
            // Cargar unidades y geocercas
            const units = await (await fetch('/api_unidades')).json();
            $('#unit_select').select2({ data: units.map(u => ({id: u.unit_id, text: `${u.label} (ID: ${u.unit_id})`})) });
            const geos = await (await fetch('/api_geocercas')).json();
            $('#geofence_select').select2({ data: geos.map(g => ({id: g.geofence_id, text: g.name})) });
            
            $('#geofence_select').on('change', function() {
                const container = $('#speed_limits_container');
                container.empty();
                $(this).select2('data').forEach(s => {
                    container.append(`<div class="form-group"><label>Límite en ${s.text}:</label><input type="number" class="geo-limit" data-id="${s.id}" value="60"></div>`);
                });
            });
        });
        async function generarReporte() {
            const geoLimits = {}; $('.geo-limit').each(function() { geoLimits[$(this).data('id')] = $(this).val(); });
            const params = new URLSearchParams({
                unit_id: $('#unit_select').val(),
                fecha_inicio: $('#fecha_inicio').val(),
                fecha_fin: $('#fecha_fin').val(),
                limite_velocidad: $('#limite_velocidad').val(),
                min_ralenti: $('#min_ralenti').val(),
                geos: JSON.stringify(geoLimits)
            });
            window.location.href = '/generar_excel?' + params.toString();
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
    # --- [Aquí iría tu lógica de 'generar_excel' ensamblada con la función get_geo abajo] ---
    # Para la columna Geocerca:
    geos_data = requests.get(f"{BASE_URL}/geofence/list.json", params={'key': API_KEY}).json().get('data', {}).get('geofences', [])
    geo_limits = json.loads(request.args.get('geos', '{}'))
    
    def get_geo_info(lat, lng):
        for g in geos_data:
            dist = math.sqrt((lat-float(g['lat']))**2 + (lng-float(g['lng']))**2) * 111
            if dist < float(g.get('radius', 500))/1000: return g
        return None

    # Dentro del bucle de escritura, invoca get_geo_info(current_lat, current_lng)
    # y si geo['geofence_id'] in geo_limits y vel > geo_limits, evento = "Exceso en Geocerca"
    return "Generando..."

if __name__ == '__main__': app.run(debug=True)
