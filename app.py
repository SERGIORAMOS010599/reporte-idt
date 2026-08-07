from flask import Flask, request, jsonify, render_template_string, send_file
import requests
import openpyxl
import io
import math
from datetime import datetime, timedelta

app = Flask(__name__)

API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

# --- HTML INTERFACE ---
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
        body { font-family: sans-serif; background-color: #f4f6f9; padding: 20px; }
        .card { max-width: 600px; margin: 0 auto; background: #fff; padding: 25px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        .btn-submit { width: 100%; background: #28a745; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; }
        .status-msg { margin-top: 10px; text-align: center; font-weight: bold; color: #d9534f; }
    </style>
</head>
<body>
    <div class="card">
        <h2>📊 Reporte Minuto a Minuto</h2>
        <select id="unit_select" style="width: 100%;"></select>
        <div style="margin-top: 15px;">
            <label>Fecha:</label>
            <input type="date" id="fecha_inicio" style="width: 45%;">
            <input type="date" id="fecha_fin" style="width: 45%;">
        </div>
        <button type="button" class="btn-submit" id="btn_submit" onclick="generarReporte()">Generar Excel</button>
        <div id="status_msg" class="status-msg"></div>
    </div>
    <script>
        async function cargarUnidades() {
            const res = await fetch('/api_unidades');
            const units = await res.json();
            const select = $('#unit_select');
            units.forEach(u => select.append(new Option(`${u.label} (ID: ${u.unit_id})`, u.unit_id)));
            select.select2();
        }
        async function generarReporte() {
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');
            btn.disabled = true; status.innerText = "⏳ Generando...";
            const params = new URLSearchParams({
                unit_id: $('#unit_select').val(),
                fecha_inicio: $('#fecha_inicio').val(),
                fecha_fin: $('#fecha_fin').val()
            });
            const res = await fetch('/generar_excel?' + params.toString());
            if (res.ok) {
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = "Reporte.xlsx";
                a.click();
                status.innerText = "¡Listo!";
            } else {
                status.innerText = "Error en la generación.";
            }
            btn.disabled = false;
        }
        cargarUnidades();
    </script>
</body>
</html>
"""

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat/2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon/2)**2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))

@app.route('/')
def index(): return render_template_string(HTML_INTERFACE)

@app.route('/api_unidades')
def api_unidades():
    res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY})
    return jsonify(res.json().get('data', {}).get('units', []))

@app.route('/generar_excel')
def generar_excel():
    unit_id = request.args.get('unit_id')
    f_in = request.args.get('fecha_inicio')
    f_fin = request.args.get('fecha_fin')
    
    url = f"{BASE_URL}/route/list.json?key={API_KEY}&unit_id={unit_id}&from={f_in}%2000:00:00&till={f_fin}%2023:59:59&include[]=points&include[]=summary"
    res = requests.get(url, timeout=45)
    data = res.json()
    
    points = []
    u = data.get('data', {}).get('units', [])
    if u: points = u[0].get('points', [])
    
    if not points: return jsonify({'error': 'No hay datos'}), 404

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Vehículo", "Fecha", "Velocidad", "Evento", "Lat", "Lng"])
    for p in points:
        ws.append([unit_id, p.get('time'), p.get('speed', 0), "Movimiento" if p.get('speed', 0) > 0 else "Parado", p.get('lat'), p.get('lng')])
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(buffer, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="Reporte.xlsx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
