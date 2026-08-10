from flask import Flask, render_template_string, request, send_file, jsonify
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import io
from datetime import datetime, timedelta
import random
import os

app = Flask(__name__)

# Asegúrate de que esta API_KEY sea la correcta
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
        .form-content { flex: 1; }
        .form-group { margin-bottom: 15px; }
        .form-row { display: flex; gap: 10px; }
        .btn-submit { width: 100%; background: #28a745; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; }
    </style>
</head>
<body>
    <div class="card">
        <div class="header"><h2>📊 Reporte Minuto a Minuto</h2></div>
        <div class="form-content">
            <div class="form-group"><label>Unidad:</label><select id="unit_select" style="width: 100%;"></select></div>
            <div class="form-row">
                <div class="form-group"><label>Inicio:</label><input type="date" id="fecha_inicio"></div>
                <div class="form-group"><label>Fin:</label><input type="date" id="fecha_fin"></div>
            </div>
            <button type="button" class="btn-submit" onclick="generarReporte()">📥 Descargar Excel</button>
        </div>
    </div>
    <script>
        $(document).ready(async function() {
            const units = await (await fetch('/api_unidades')).json();
            $('#unit_select').select2({ data: units.map(u => ({id: u.unit_id, text: `${u.label} (ID: ${u.unit_id})`})) });
        });
        function generarReporte() {
            const p = new URLSearchParams({ unit_id: $('#unit_select').val(), fecha_inicio: $('#fecha_inicio').val(), fecha_fin: $('#fecha_fin').val() });
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
    try:
        res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY}, timeout=15)
        return jsonify(res.json().get('data', {}).get('units', []))
    except: return jsonify([]), 500

@app.route('/generar_excel')
def generar_excel():
    # Aquí iría tu función completa de generación
    return "Reporte funcionando"

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
