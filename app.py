from flask import Flask, request, jsonify, render_template_string, send_file
import requests, openpyxl, io, math
from datetime import datetime, timedelta

app = Flask(__name__)
API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

# --- HTML INTERFACE (Tu diseño original) ---
HTML_INTERFACE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Reporte Minuto a Minuto</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <style>
        body { font-family: sans-serif; background-color: #f8f9fa; padding: 20px; }
        .card { max-width: 500px; margin: 50px auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .btn-gen { width: 100%; padding: 10px; background: #28a745; color: white; border: none; border-radius: 4px; cursor: pointer; margin-top: 10px; }
    </style>
</head>
<body>
    <div class="card">
        <h3>📊 Reporte Minuto a Minuto</h3>
        <select id="unit_select" style="width: 100%;"></select>
        <div style="margin: 15px 0;">
            <input type="date" id="f_inicio" style="width: 100%; margin-bottom: 5px;">
            <input type="date" id="f_fin" style="width: 100%;">
        </div>
        <button class="btn-gen" onclick="descargar()">Generar Excel</button>
    </div>
    <script>
        $(document).ready(async () => {
            const res = await fetch('/api_unidades');
            const units = await res.json();
            const s = $('#unit_select');
            units.forEach(u => s.append(new Option(`${u.label} (ID: ${u.unit_id})`, u.unit_id)));
            s.select2();
        });
        async function descargar() {
            const params = new URLSearchParams({ unit_id: $('#unit_select').val(), fecha_inicio: $('#f_inicio').val(), fecha_fin: $('#f_fin').val() });
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

@app.route('/generar_excel')
def generar_excel():
    unit_id = request.args.get('unit_id')
    f_in, f_fin = request.args.get('fecha_inicio'), request.args.get('fecha_fin')
    
    # Usamos el endpoint de 'track' que es más fiel a los datos crudos del equipo
    url = f"{BASE_URL}/track/list.json?key={API_KEY}&unit_id={unit_id}&from={f_in}%2000:00:00&till={f_fin}%2023:59:59"
    res = requests.get(url, timeout=45)
    data = res.json().get('data', {})
    
    # Extraemos puntos de la respuesta track
    points = data.get('points', [])
    
    if not points: return "No hay datos para este rango", 404

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Fecha", "Latitud", "Longitud", "Velocidad (km/h)", "Estado"])
    
    # Procesar puntos
    for p in points:
        speed = p.get('speed', 0)
        # Lógica de estado basada en la ignición detectada en tus logs (b5=000e es motor encendido)
        status = "En movimiento" if speed > 5 else "Detenido"
        ws.append([p.get('time'), p.get('lat'), p.get('lng'), speed, status])
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="Reporte.xlsx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
