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
        .btn-submit { width: 100%; background: #28a745; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; font-size: 14px; cursor: pointer; transition: background 0.2s; margin-top: 10px; }
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
                    <label>Buscar / Seleccionar Unidad:</label>
                    <select id="unit_select" style="width: 100%;"></select>
                </div>
                <div class="form-row">
                    <div class="form-group"><label>Fecha Inicial:</label><input type="date" id="fecha_inicio" required></div>
                    <div class="form-group"><label>Fecha Final:</label><input type="date" id="fecha_fin" required></div>
                </div>
                <button type="button" class="btn-submit" id="btn_submit" onclick="generarReporte()">📥 Generar y Descargar Excel</button>
                <div id="status_msg" class="status-msg"></div>
            </div>
        </div>
    </div>
    <script>
        async function cargarUnidades() {
            const res = await fetch('/api_unidades');
            const units = await res.json();
            const select = $('#unit_select');
            select.empty();
            units.forEach(u => select.append(new Option(`${u.label} (ID: ${u.unit_id})`, u.unit_id)));
            select.select2();
        }
        async function generarReporte() {
            const params = new URLSearchParams({
                unit_id: $('#unit_select').val(),
                fecha_inicio: $('#fecha_inicio').val(),
                fecha_fin: $('#fecha_fin').val()
            });
            window.location.href = `/generar_excel?${params.toString()}`;
        }
        function setRange(t) {
            const d = new Date().toISOString().split('T')[0];
            document.getElementById('fecha_inicio').value = d;
            document.getElementById('fecha_fin').value = d;
        }
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

@app.route('/generar_excel')
def generar_excel():
    try:
        unit_id = request.args.get('unit_id')
        f_in = request.args.get('fecha_inicio')
        f_fin = request.args.get('fecha_fin')
        
        url = f"{BASE_URL}/route/list.json"
        params = {
            'key': API_KEY,
            'unit_id': unit_id,
            'from': f"{f_in} 00:00:00",
            'till': f"{f_fin} 23:59:59",
            'include[]': ['points']
        }
        
        res = requests.get(url, params=params, timeout=45)
        data = res.json()
        
        # Lógica de diagnóstico: si la API no trae puntos, generamos un Excel con el error
        points = data.get('data', {}).get('units', [{}])[0].get('points', [])
        
        wb = openpyxl.Workbook()
        ws = wb.active
        if not points:
            ws.append(["Error: No hay datos, respuesta API:"])
            ws.append([str(data)])
        else:
            ws.append(["Fecha", "Latitud", "Longitud", "Velocidad"])
            for p in points:
                ws.append([p.get('time'), p.get('lat'), p.get('lng'), p.get('speed', 0)])
        
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                         as_attachment=True, download_name="Reporte_IDT.xlsx")
    except Exception as e:
        return f"Error: {str(e)}"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
