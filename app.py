from flask import Flask, request, jsonify, render_template_string, send_file
import requests
import openpyxl
import io
import math
from datetime import datetime, timedelta

app = Flask(__name__)

API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

# Diseño original solicitado
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
        #status_msg { margin-top: 10px; text-align: center; color: #d9534f; font-weight: bold; }
    </style>
</head>
<body>
    <div class="card">
        <h3>📊 Reporte Minuto a Minuto</h3>
        <select id="unit_select" style="width: 100%;"></select>
        <div style="margin: 15px 0;">
            <label>Fecha Inicial:</label>
            <input type="date" id="f_inicio" style="width: 100%;">
            <label>Fecha Final:</label>
            <input type="date" id="f_fin" style="width: 100%;">
        </div>
        <button class="btn-gen" id="btn_submit" onclick="descargar()">Generar Excel</button>
        <div id="status_msg"></div>
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
            const btn = $('#btn_submit');
            btn.prop('disabled', true).text('Generando...');
            const params = new URLSearchParams({ 
                unit_id: $('#unit_select').val(), 
                fecha_inicio: $('#f_inicio').val(), 
                fecha_fin: $('#f_fin').val() 
            });
            try {
                const res = await fetch('/generar_excel?' + params.toString());
                if(res.ok) {
                    const blob = await res.blob();
                    const a = document.createElement('a');
                    a.href = window.URL.createObjectURL(blob);
                    a.download = "Reporte.xlsx";
                    a.click();
                    $('#status_msg').text("¡Éxito!");
                } else {
                    const err = await res.json();
                    alert(err.error);
                }
            } catch(e) { alert("Error de conexión"); }
            btn.prop('disabled', false).text('Generar Excel');
        }
    </script>
</body>
</html>
"""

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat, dlon = math.radians(lat2-lat1), math.radians(lon2-lon1)
    a = math.sin(dlat/2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon/2)**2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))

@app.route('/')
def index(): return render_template_string(HTML_INTERFACE)

@app.route('/api_unidades')
def api_unidades():
    try:
        res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY}, timeout=10)
        return jsonify(res.json().get('data', {}).get('units', []))
    except: return jsonify([])

@app.route('/generar_excel')
def generar_excel():
    unit_id = request.args.get('unit_id')
    f_in, f_fin = request.args.get('fecha_inicio'), request.args.get('fecha_fin')
    
    # URL directa a la API para evitar errores de codificación
    url = f"{BASE_URL}/route/list.json?key={API_KEY}&unit_id={unit_id}&from={f_in}%2000:00:00&till={f_fin}%2023:59:59&include[]=points&include[]=summary"
    
    try:
        res = requests.get(url, timeout=45)
        data = res.json().get('data', {})
        units = data.get('units', [])
        if not units or 'points' not in units[0]:
            return jsonify({'error': 'No se obtuvieron posiciones GPS para la unidad en este rango.'}), 404
        
        points = units[0]['points']
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Vehículo", "Fecha", "Velocidad", "Evento", "Lat", "Lng"])
        
        for p in points:
            speed = p.get('speed', 0)
            ws.append([unit_id, p.get('time'), speed, "Movimiento" if speed > 0 else "Parado", p.get('lat'), p.get('lng')])
        
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="Reporte.xlsx")
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
